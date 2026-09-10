"""Read-only independent verifier for construction project consolidation.

Usage: python -X utf8 scripts/verify-consolidation.py manifest.json [--out report.json]
Exit 0: all checks passed; exit 1: invalid input or a failed check.
Only --out may be written; source/target files are never modified.
"""

import argparse
import copy
import datetime as dt
import hashlib
import json
from pathlib import Path
import sys
import unicodedata


def read_json(path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def dates(start, end):
    current, last = dt.date.fromisoformat(start), dt.date.fromisoformat(end)
    if current > last:
        raise ValueError("start_date is after end_date")
    while current <= last:
        yield current.isoformat()
        current += dt.timedelta(days=1)


def clean_daily(record):
    result = copy.deepcopy(record)
    if result.get("_weekend_auto") == "off":
        result.pop("_weekend_auto")
    return result


def flags(project, record):
    if record is not None:
        return record.get("day") is not False, record.get("night") is True
    night = project.get("default_shift") == "night"
    return not night, night


def texts(project, record):
    result = {}
    for shift in ("day", "night"):
        for field in ("our_person", "safety_person", "partner_person"):
            key = f"{shift}_{field}"
            result[key] = (record or {}).get(key) or project.get(field) or ""
    for key, value in (record or {}).items():
        if isinstance(value, str) and value and key != "_weekend_auto":
            result[key] = value
    return {k: v for k, v in result.items() if isinstance(v, str) and v}


def equivalent_spacing(value):
    return " ".join(unicodedata.normalize("NFKC", value).split())


class Verifier:
    def __init__(self):
        self.errors = []
        self.checks = 0
        self.days = 0
        self.users = []

    def check(self, ok, context, message):
        self.checks += 1
        if not ok:
            self.errors.append({"context": context, "error": message})
        return ok

    def dataset(self, data, context):
        projects = data["projects"]
        ids = [p["id"] for p in projects]
        self.check(len(set(ids)) == len(ids), context, "duplicate project IDs")
        for p in projects:
            list(dates(p["start_date"], p["end_date"]))
        for key, value in data["daily"].items():
            pid, date = key.rsplit("/", 1)
            dt.date.fromisoformat(date)
            self.check(pid in ids and isinstance(value, dict), context, f"invalid daily {key}")
        return {p["id"]: p for p in projects}

    def verify_user(self, entry, root):
        user = entry["user"]
        before = read_json(root / entry["before"])
        after = read_json(root / entry["after"])
        old = self.dataset(before, user + ":before")
        new = self.dataset(after, user + ":after")
        groups = entry.get("groups", [])
        grouped = [pid for group in groups for pid in group["ids"]]
        removed = {pid for group in groups for pid in group["ids"] if pid != group["primary_id"]}
        self.check(len(grouped) == len(set(grouped)), user, "project belongs to multiple groups")
        self.check(set(new) == set(old) - removed, user, "unexpected new/missing IDs")
        expected_count = len(old) - sum(len(g["ids"]) - 1 for g in groups)
        self.check(len(new) == expected_count, user, "project count does not match declared consolidation")
        unrelated = set(old) - set(grouped)
        for pid in unrelated:
            self.check(old[pid] == new.get(pid), user + ":" + pid, "unrelated project changed")
        old_daily = {k: v for k, v in before["daily"].items() if k.rsplit("/", 1)[0] in unrelated}
        new_daily = {k: v for k, v in after["daily"].items() if k.rsplit("/", 1)[0] in unrelated}
        self.check(old_daily == new_daily, user, "unrelated daily changed")
        for group in groups:
            self.verify_group(user, group, old, new, before["daily"], after["daily"])
        self.users.append({"user": user, "before_projects": len(old), "after_projects": len(new), "groups": len(groups), "removed_ids": len(removed), "before_daily": len(before["daily"]), "after_daily": len(after["daily"])})

    def verify_group(self, user, group, old, new, before_daily, after_daily):
        ids, primary = group["ids"], group["primary_id"]
        context = f"{user}:{primary}"
        if not self.check(len(ids) >= 2 and primary in ids and all(pid in old for pid in ids) and primary in new, context, "invalid group IDs"):
            return
        self.check(bool(group.get("reason", "").strip()), context, "missing consolidation reason")
        expected = copy.deepcopy(old[primary])
        expected["start_date"] = min(old[pid]["start_date"] for pid in ids)
        expected["end_date"] = max(old[pid]["end_date"] for pid in ids)
        merged = new[primary]
        self.check(merged == expected, context, "merged metadata differs from primary or date union")
        normalized = {pid: copy.deepcopy(old[pid]) for pid in ids}
        normalization_keys = set()
        for declaration in group.get("normalizations", []):
            pid, field = declaration.get("id"), declaration.get("field")
            source, target = declaration.get("from"), declaration.get("to")
            valid = (
                pid in ids and field == "partner"
                and isinstance(source, str) and isinstance(target, str)
                and old[pid].get(field) == source
                and equivalent_spacing(source) == equivalent_spacing(target)
                and (pid, field) not in normalization_keys
            )
            self.check(valid, context, "invalid or non-equivalent partner normalization")
            if valid:
                normalized[pid][field] = target
                normalization_keys.add((pid, field))
        for pid, source_project in normalized.items():
            for field in set(source_project) | set(merged):
                if field in ("id", "start_date", "end_date"):
                    continue
                source_value, output_value = source_project.get(field), merged.get(field)
                # Legacy omission and explicit empty string carry no metadata.
                if source_value in (None, "") and output_value in (None, ""):
                    continue
                self.check(source_value == output_value, context + ":" + pid, f"nonprimary metadata differs without allowed normalization: {field}")
        resolutions = group.get("resolutions", {})
        overrides = group.get("daily_overrides", {})
        losses = group.get("documented_loss", [])
        for loss in losses:
            self.check(all(key in loss for key in ("date", "field", "old_value", "reason")) and bool(loss.get("reason", "").strip()), context, "incomplete documented loss")
        relevant = {k.rsplit("/", 1)[1] for k in before_daily if k.rsplit("/", 1)[0] in ids}
        range_dates = set(dates(expected["start_date"], expected["end_date"]))
        all_dates = relevant | range_dates
        for date in sorted(all_dates):
            self.days += 1
            here = context + ":" + date
            active_sources = [pid for pid in ids if old[pid]["start_date"] <= date <= old[pid]["end_date"]]
            records = {pid: before_daily[f"{pid}/{date}"] for pid in ids if f"{pid}/{date}" in before_daily}
            actual = after_daily.get(f"{primary}/{date}")
            for pid in ids:
                if pid != primary:
                    self.check(f"{pid}/{date}" not in after_daily, here, "removed ID daily remains")
            selected = resolutions.get(date)
            if selected is not None:
                self.check(selected in ids and (selected in active_sources or selected in records), here, "resolution references unavailable source")
            baseline = selected or (primary if primary in records else next(iter(records), None)) or (primary if primary in active_sources else next(iter(active_sources), primary))
            explicit_override = date in overrides
            if explicit_override:
                self.check(isinstance(overrides[date], dict) and actual == overrides[date], here, "daily override does not exactly match output")
            elif records:
                distinct = {json.dumps(clean_daily(v), sort_keys=True, ensure_ascii=False) for v in records.values()}
                self.check(len(distinct) <= 1 or selected is not None, here, "conflicting explicit daily needs declared resolution/override")
                record = records.get(baseline)
                if record is not None:
                    self.check(actual is not None and clean_daily(actual) == clean_daily(record), here, "selected explicit daily was not preserved")
                    if actual is not None and "_weekend_auto" in actual:
                        self.check(actual.get("_weekend_auto") == record.get("_weekend_auto"), here, "undeclared automatic weekend marker added")
            if not active_sources and date in range_dates:
                self.check(actual is not None and actual.get("day") is False and actual.get("night") is False, here, "gap day must be explicitly off")
            elif active_sources:
                source_flags = {flags(old[pid], records.get(pid)) for pid in active_sources}
                self.check(len(source_flags) <= 1 or selected is not None or explicit_override, here, "different effective flags require declared resolution/override")
                expected_flags = flags(merged, overrides[date]) if explicit_override else flags(old[baseline], records.get(baseline))
                self.check(flags(merged, actual) == expected_flags, here, "effective day/night flags changed")
            if actual is not None and not records and not explicit_override and active_sources:
                self.check(not texts(merged, actual) or all(k in texts(merged, None) and v == texts(merged, None)[k] for k, v in texts(merged, actual).items()), here, "new undeclared daily text appeared")
            output_text = texts(merged, actual)
            for pid in set(active_sources) | set(records):
                for field, value in texts(old[pid], records.get(pid)).items():
                    retained = value in output_text.get(field, "")
                    documented = any(loss.get("date") == date and loss.get("field") == field and loss.get("old_value") == value and bool(loss.get("reason", "").strip()) for loss in losses)
                    self.check(retained or documented, here, f"nonempty source text lost: {pid} {field} {value!r}")
        output_dates = {k.rsplit("/", 1)[1] for k in after_daily if k.rsplit("/", 1)[0] == primary}
        self.check(output_dates <= all_dates, context, "unexpected daily dates added")
        self.check(set(resolutions) <= all_dates and set(overrides) <= all_dates, context, "manifest contains out-of-range resolution/override")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("manifest", type=Path)
    parser.add_argument("--out", type=Path)
    args = parser.parse_args()
    checker = Verifier()
    protected = {args.manifest.resolve()}
    manifest = {}
    hashes = {}
    try:
        manifest = read_json(args.manifest)
        root = args.manifest.resolve().parent
        for entry in manifest["users"]:
            for field in ("before", "after"):
                protected.add((root / entry[field]).resolve())
        if args.out and args.out.resolve() in protected:
            raise ValueError("--out must not overwrite manifest or before/after files")
        hashes = {str(path): hashlib.sha256(path.read_bytes()).hexdigest() for path in protected}
        checker.check(bool(manifest.get("source_commit")), "manifest", "source_commit missing")
        names = [u["user"] for u in manifest["users"]]
        checker.check(len(names) == len(set(names)), "manifest", "duplicate user entries")
        for entry in manifest["users"]:
            checker.verify_user(entry, root)
        checker.check(all(hashlib.sha256(Path(path).read_bytes()).hexdigest() == digest for path, digest in hashes.items()), "files", "input file changed during read-only verification")
    except Exception as exc:
        checker.errors.append({"context": "execution", "error": f"{type(exc).__name__}: {exc}"})
    report = {"status": "FAIL" if checker.errors else "PASS", "source_commit": manifest.get("source_commit"), "checks": checker.checks, "group_dates_checked": checker.days, "users": checker.users, "failures": checker.errors, "input_sha256": hashes}
    output = json.dumps(report, ensure_ascii=False, indent=2)
    if args.out and args.out.resolve() not in protected:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(output + "\n", encoding="utf-8")
    print(output)
    return 1 if checker.errors else 0


if __name__ == "__main__":
    sys.exit(main())
