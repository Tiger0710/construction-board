/* Explicit, loss-aware consolidation. This helper never writes to a server. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ProjectMerge = factory();
}(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';
  var FIELDS = ['client', 'title', 'our_person', 'safety_person', 'partner', 'partner_person'];
  var TEXT_FIELDS = ['day_work', 'day_priority', 'day_priority_detail', 'day_our_person',
    'day_safety_person', 'day_partner_person', 'night_work', 'night_priority',
    'night_priority_detail', 'night_our_person', 'night_safety_person', 'night_partner_person'];
  var own = function (obj, key) { return Object.prototype.hasOwnProperty.call(obj, key); };
  function clone(value) {
    if (Array.isArray(value)) return value.map(clone);
    if (value && typeof value === 'object') {
      var result = {};
      Object.keys(value).forEach(function (key) {
        Object.defineProperty(result, key, { value: clone(value[key]), enumerable: true, writable: true, configurable: true });
      });
      return result;
    }
    return value;
  }
  function stable(value) {
    if (Array.isArray(value)) return '[' + value.map(stable).join(',') + ']';
    if (value && typeof value === 'object') return '{' + Object.keys(value).sort().map(function (k) {
      return JSON.stringify(k) + ':' + stable(value[k]);
    }).join(',') + '}';
    return JSON.stringify(value);
  }
  function signature(project) {
    return JSON.stringify(FIELDS.map(function (field) { return String(project[field] == null ? '' : project[field]).trim(); }));
  }
  function validDate(value) {
    if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
    var time = Date.parse(value + 'T00:00:00Z');
    return Number.isFinite(time) && new Date(time).toISOString().slice(0, 10) === value;
  }
  function validProject(p) {
    return p && typeof p.id === 'string' && p.id.length > 0 && p.id.indexOf('/') === -1 &&
      validDate(p.start_date) && validDate(p.end_date) && p.start_date <= p.end_date;
  }
  function nextDate(ds) { return new Date(Date.parse(ds + 'T00:00:00Z') + 86400000).toISOString().slice(0, 10); }
  function order(a, b) { return a.start_date.localeCompare(b.start_date) || a.id.localeCompare(b.id); }
  function describe(p) {
    var info = { id: p.id, start_date: p.start_date, end_date: p.end_date };
    FIELDS.forEach(function (f) { info[f] = p[f] || ''; });
    return info;
  }
  function findCandidates(projects) {
    var buckets = new Map(), results = [], counts = new Map();
    (projects || []).forEach(function (p) { if (p) counts.set(p.id, (counts.get(p.id) || 0) + 1); });
    (projects || []).forEach(function (p) {
      if (!validProject(p) || counts.get(p.id) !== 1) return;
      var key = signature(p);
      if (!buckets.has(key)) buckets.set(key, []);
      buckets.get(key).push(p);
    });
    function add(group, end) {
      if (group.length < 2) return;
      results.push({ ids: group.map(function (p) { return p.id; }), projects: group.map(describe),
        title: group[0].title, start_date: group[0].start_date, end_date: end });
    }
    buckets.forEach(function (items) {
      items.sort(order);
      var group = [], end = '';
      items.forEach(function (p) {
        if (group.length && Date.parse(p.start_date) - Date.parse(end) > 86400000) {
          add(group, end); group = []; end = '';
        }
        group.push(p);
        if (p.end_date > end) end = p.end_date;
      });
      add(group, end);
    });
    return results;
  }
  function effective(value, implicit) {
    if (implicit) return { day: true, night: false };
    if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('日別データの形式が不正です');
    // Match the UI/API: absent day means active; only literal true enables night.
    var result = clone(value);
    result.day = value.day !== false;
    result.night = value.night === true;
    TEXT_FIELDS.forEach(function (field) {
      if (!own(result, field) || result[field] === '') delete result[field];
    });
    return result;
  }
  function metadata(p) {
    var result = clone(p);
    ['id', 'start_date', 'end_date'].concat(FIELDS).forEach(function (f) { delete result[f]; });
    return stable(result);
  }
  function preview(data, ids) {
    if (!data || !Array.isArray(data.projects) || !data.daily || typeof data.daily !== 'object' || Array.isArray(data.daily))
      throw new Error('工事データの形式が不正です');
    if (!Array.isArray(ids) || ids.length < 2 || new Set(ids).size !== ids.length)
      throw new Error('異なる工事を2件以上選択してください');
    var selected = ids.map(function (id) {
      var matches = data.projects.filter(function (p) { return p.id === id; });
      if (matches.length !== 1 || !validProject(matches[0])) throw new Error('対象の工事IDまたは期間が不正です: ' + id);
      return matches[0];
    });
    selected.forEach(function (p) {
      if (signature(p) !== signature(selected[0])) throw new Error('客先・工事件名・担当者・協力会社が異なる工事は統合できません');
      if (metadata(p) !== metadata(selected[0])) throw new Error('工事の追加情報が異なります。統合前に内容を確認してください');
    });
    var project = clone(selected[0]);
    project.start_date = selected.reduce(function (v, p) { return p.start_date < v ? p.start_date : v; }, project.start_date);
    project.end_date = selected.reduce(function (v, p) { return p.end_date > v ? p.end_date : v; }, project.end_date);
    var dates = new Set(), selectedIds = new Set(ids);
    Object.keys(data.daily).forEach(function (key) {
      var slash = key.indexOf('/');
      if (!selectedIds.has(key.slice(0, slash))) return;
      var date = key.slice(slash + 1);
      if (!validDate(date)) throw new Error('対象工事の日付キーが不正です: ' + key);
      dates.add(date);
    });
    // Gaps must remain off instead of acquiring the merged project's implicit day shift.
    var sorted = selected.slice().sort(order), coveredEnd = sorted[0].end_date;
    sorted.slice(1).forEach(function (p) {
      if (Date.parse(p.start_date) - Date.parse(coveredEnd) > 86400000) {
        for (var ds = nextDate(coveredEnd); ds < p.start_date; ds = nextDate(ds)) dates.add(ds);
      }
      if (p.end_date > coveredEnd) coveredEnd = p.end_date;
    });
    var daily = {}, conflicts = [];
    Array.from(dates).sort().forEach(function (date) {
      var choices = [];
      selected.forEach(function (p) {
        var key = p.id + '/' + date, explicit = own(data.daily, key);
        if (!explicit && !(p.start_date <= date && date <= p.end_date)) return;
        var value = explicit ? clone(data.daily[key]) : { day: true, night: false };
        choices.push({ projectId: p.id, start_date: p.start_date, end_date: p.end_date,
          value: value, implicit: !explicit });
      });
      if (!choices.length) { daily[date] = { day: false, night: false }; return; }
      var values = new Set(choices.map(function (choice) { return stable(effective(choice.value, choice.implicit)); }));
      if (values.size > 1) conflicts.push({ date: date, choices: choices });
      else {
        // Prefer an explicit record so empty-but-present metadata is retained.
        var explicitChoices = choices.filter(function (c) { return !c.implicit; });
        var chosen = explicitChoices.length ? explicitChoices[0] : choices[0];
        var result = clone(chosen.value);
        explicitChoices.forEach(function (c) { Object.keys(c.value).forEach(function (k) {
          if (!own(result, k)) result[k] = clone(c.value[k]);
        }); });
        daily[date] = result;
      }
    });
    return { ids: ids.slice(), primaryId: ids[0], project: project, daily: daily, conflicts: conflicts };
  }
  function merge(data, ids, resolutions) {
    var plan = preview(data, ids), result = clone(data), selectedIds = new Set(ids);
    resolutions = resolutions || {};
    plan.conflicts.forEach(function (conflict) {
      var chosen = own(resolutions, conflict.date) && conflict.choices.find(function (choice) {
        return choice.projectId === resolutions[conflict.date];
      });
      if (!chosen) throw new Error('日別内容の選択が必要です: ' + conflict.date);
      plan.daily[conflict.date] = clone(chosen.value);
    });
    result.projects = result.projects.filter(function (p) { return !selectedIds.has(p.id) || p.id === plan.primaryId; })
      .map(function (p) { return p.id === plan.primaryId ? clone(plan.project) : p; });
    Object.keys(result.daily).forEach(function (key) {
      if (selectedIds.has(key.slice(0, key.indexOf('/')))) delete result.daily[key];
    });
    Object.keys(plan.daily).forEach(function (date) { result.daily[plan.primaryId + '/' + date] = clone(plan.daily[date]); });
    return result;
  }
  return { findCandidates: findCandidates, preview: preview, merge: merge };
}));
