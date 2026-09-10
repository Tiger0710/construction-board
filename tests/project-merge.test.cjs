const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const Merge = require('../static/project-merge.js');
const p = (id, start = '2026-08-30', end = '2026-09-02', extra = {}) => ({
  id, client: '客先', title: '工事', our_person: '森', safety_person: '神邊', partner: '', partner_person: '',
  start_date: start, end_date: end, ...extra
});
const data = (daily = {}, projects = [p('a'), p('b')]) => ({ projects, daily, _sha: 'unchanged', extra: { flag: 1 } });

test('UMD exports in browser without Node', () => {
  const sandbox = {};
  vm.runInNewContext(fs.readFileSync(require.resolve('../static/project-merge.js'), 'utf8'), sandbox);
  assert.equal(typeof sandbox.ProjectMerge.merge, 'function');
});
test('candidate groups match trimmed identity and transitive adjacent periods', () => {
  const projects = [p('c', '2026-09-02', '2026-09-03'), p('a', '2026-08-30', '2026-08-31'),
    p('b', '2026-09-01', '2026-09-01', { title: ' 工事 ' }),
    p('far', '2026-10-01', '2026-10-02'), p('different', undefined, undefined, { our_person: '別人' })];
  const before = structuredClone(projects);
  assert.deepEqual(Merge.findCandidates(projects).map(g => g.ids), [['a', 'b', 'c']]);
  assert.deepEqual(projects, before);
});
test('invalid dates and duplicate IDs cannot become candidates', () => {
  assert.deepEqual(Merge.findCandidates([p('a'), p('a'), p('b', '2026-02-30')]), []);
});
test('primary follows explicit selection order and preserves unrelated and top-level data', () => {
  const source = data({ 'a/2026-08-30': { day: true, day_work: '基礎' },
    'b/2026-09-03': { day: false, night: true, night_work: '舗装', night_our_person: '猪股' },
    'abc/2026-08-30': { custom: ['unchanged'] } }, [p('a', '2026-08-30', '2026-08-31'),
    p('b', '2026-09-01', '2026-09-03'), p('abc')]);
  const before = structuredClone(source);
  const result = Merge.merge(source, ['b', 'a']);
  assert.equal(result.projects[0].id, 'b');
  assert.equal(result.projects[0].start_date, '2026-08-30');
  assert.equal(result.projects[0].end_date, '2026-09-03');
  assert.deepEqual(result.daily['b/2026-08-30'], source.daily['a/2026-08-30']);
  assert.deepEqual(result.daily['abc/2026-08-30'], source.daily['abc/2026-08-30']);
  assert.deepEqual(result.extra, source.extra);
  assert.equal(result._sha, 'unchanged');
  assert.deepEqual(source, before);
  result.extra.flag = 2;
  assert.equal(source.extra.flag, 1);
});
test('identical daily data with different key ordering collapse without conflict', () => {
  const source = data({ 'a/2026-09-01': { day: true, night: false, day_work: '作業' },
    'b/2026-09-01': { day_work: '作業', night: false, day: true } });
  assert.equal(Merge.preview(source, ['a', 'b']).conflicts.length, 0);
  assert.equal(Object.keys(Merge.merge(source, ['a', 'b']).daily).length, 1);
});
test('explicit default equals implicit day; empty metadata is retained', () => {
  const source = data({ 'b/2026-09-01': { day: true, night: false, day_work: '', night_work: '' } });
  assert.equal(Merge.preview(source, ['a', 'b']).conflicts.length, 0);
  assert.deepEqual(Merge.merge(source, ['a', 'b']).daily['a/2026-09-01'], source.daily['b/2026-09-01']);
});
test('explicit off vs missing implicit on requires date choice', () => {
  const source = data({ 'a/2026-09-01': { day: false, night: false } });
  const plan = Merge.preview(source, ['a', 'b']);
  assert.equal(plan.conflicts.length, 1);
  assert.equal(plan.conflicts[0].choices[1].implicit, true);
  assert.throws(() => Merge.merge(source, ['a', 'b']), /選択/);
  assert.throws(() => Merge.merge(source, ['a', 'b'], { '2026-09-01': 'other' }), /選択/);
  assert.deepEqual(Merge.merge(source, ['a', 'b'], { '2026-09-01': 'b' }).daily['a/2026-09-01'], { day: true, night: false });
  assert.deepEqual(Merge.merge(source, ['a', 'b'], { '2026-09-01': 'a' }).daily['a/2026-09-01'], { day: false, night: false });
});
test('work, people, priority, nights and unknown fields never disappear silently', () => {
  for (const field of ['day_work', 'day_our_person', 'day_safety_person', 'day_partner_person',
    'day_priority', 'day_priority_detail', 'night_work', 'night_our_person', 'night_priority', 'custom']) {
    const source = data({ 'a/2026-09-01': { day: true, night: true, [field]: 'A' },
      'b/2026-09-01': { day: true, night: true, [field]: 'B' } });
    assert.equal(Merge.preview(source, ['a', 'b']).conflicts.length, 1, field);
    assert.deepEqual(Merge.merge(source, ['a', 'b'], { '2026-09-01': 'b' }).daily['a/2026-09-01'], source.daily['b/2026-09-01']);
  }
  assert.equal(Merge.preview(data({ 'a/2026-09-01': { day: true, night: true } }), ['a', 'b']).conflicts.length, 1);
});
test('out-of-period daily records are preserved and conflicts checked', () => {
  const source = data({ 'a/2026-07-01': { day: false, night: true, custom: { nested: 1 } },
    'b/2026-10-01': { day: true, day_work: '後工程' } });
  const result = Merge.merge(source, ['a', 'b']);
  assert.deepEqual(result.daily['a/2026-07-01'], source.daily['a/2026-07-01']);
  assert.deepEqual(result.daily['a/2026-10-01'], source.daily['b/2026-10-01']);
  source.daily['b/2026-07-01'] = { day: true };
  assert.throws(() => Merge.merge(source, ['a', 'b']), /選択/);
});
test('noncontiguous selected periods do not invent work in gap days', () => {
  const result = Merge.merge(data({}, [p('a', '2026-08-30', '2026-08-31'), p('b', '2026-09-03', '2026-09-04')]), ['a', 'b']);
  assert.deepEqual(result.daily, { 'a/2026-09-01': { day: false, night: false }, 'a/2026-09-02': { day: false, night: false } });
});
test('invalid selections, metadata differences and malformed daily data reject loss', () => {
  for (const ids of [[], ['a'], ['a', 'a'], ['a', 'missing']]) assert.throws(() => Merge.preview(data(), ids));
  assert.throws(() => Merge.preview(data({}, [p('a'), p('b', undefined, undefined, { partner: '別会社' })]), ['a', 'b']));
  assert.throws(() => Merge.preview(data({}, [p('a', undefined, undefined, { memo: 'A' }), p('b')]), ['a', 'b']), /追加情報/);
  assert.throws(() => Merge.preview(data({ 'a/not-date': {} }), ['a', 'b']), /日付キー/);
  assert.throws(() => Merge.preview(data({ 'a/2026-09-01': null }), ['a', 'b']), /形式/);
  assert.throws(() => Merge.preview(data({}, [p('a'), p('b', '2026-02-30')]), ['a', 'b']), /期間/);
});
test('all implicit overlapping schedules need no redundant daily rows', () => {
  assert.deepEqual(Merge.merge(data(), ['a', 'b']).daily, {});
});
test('explicit record without day flag retains the UI implicit active day', () => {
  const source = data({ 'a/2026-09-01': { day_work: '作業' },
    'b/2026-09-01': { day: true, night: false, day_work: '作業' } });
  assert.equal(Merge.preview(source, ['a', 'b']).conflicts.length, 0);
  const merged = Merge.merge(source, ['a', 'b']).daily['a/2026-09-01'];
  assert.equal(merged.day !== false, true);
  assert.equal(merged.day_work, '作業');
  assert.equal(Merge.preview(data({ 'a/2026-09-01': {} }), ['a', 'b']).conflicts.length, 0);
});
test('candidate groups separate defaults while absent and explicit day are equivalent', () => {
  const projects = [p('legacy'), p('day', undefined, undefined, { default_shift: 'day' }),
    p('night1', undefined, undefined, { default_shift: 'night' }), p('night2', undefined, undefined, { default_shift: 'night' })];
  assert.deepEqual(Merge.findCandidates(projects).map(g => g.ids.slice().sort()), [['day', 'legacy'], ['night1', 'night2']]);
  assert.deepEqual(Merge.findCandidates(projects).map(g => g.projects[0].default_shift), ['day', 'night']);
  assert.throws(() => Merge.preview(data({}, projects), ['legacy', 'night1']), /昼夜設定が異なる/);
});
test('absent and explicit day merge while preserving primary metadata shape', () => {
  const source = data({}, [p('a'), p('b', undefined, undefined, { default_shift: 'day' })]);
  const before = structuredClone(source);
  const legacyPrimary = Merge.merge(source, ['a', 'b']).projects[0];
  assert.equal(Object.hasOwn(legacyPrimary, 'default_shift'), false);
  assert.equal(Merge.merge(source, ['b', 'a']).projects[0].default_shift, 'day');
  assert.deepEqual(source, before);
});
test('night default implicit day conflicts with explicit off or legacy daytime', () => {
  const projects = [p('a', undefined, undefined, { default_shift: 'night' }),
    p('b', undefined, undefined, { default_shift: 'night' })];
  for (const explicit of [{ day: false, night: false }, { day: true, night: false }, {}, { day_work: '昼間の作業' }]) {
    const source = data({ 'a/2026-09-01': explicit }, projects);
    const plan = Merge.preview(source, ['a', 'b']);
    assert.equal(plan.conflicts.length, 1);
    assert.deepEqual(plan.conflicts[0].choices[1].value, { day: false, night: true });
    assert.equal(plan.conflicts[0].choices[1].implicit, true);
    assert.throws(() => Merge.merge(source, ['a', 'b']), /選択/);
    const implicitChosen = Merge.merge(source, ['a', 'b'], { '2026-09-01': 'b' });
    assert.deepEqual(implicitChosen.daily['a/2026-09-01'], { day: false, night: true });
    assert.equal(implicitChosen.projects[0].default_shift, 'night');
    assert.deepEqual(Merge.merge(source, ['a', 'b'], { '2026-09-01': 'a' }).daily['a/2026-09-01'], explicit);
  }
});
test('explicit matching night default collapses without changing entered records', () => {
  const projects = [p('a', undefined, undefined, { default_shift: 'night' }),
    p('b', undefined, undefined, { default_shift: 'night' })];
  const source = data({ 'a/2026-09-01': { day: false, night: true, night_work: '' } }, projects);
  assert.equal(Merge.preview(source, ['a', 'b']).conflicts.length, 0);
  assert.deepEqual(Merge.merge(source, ['a', 'b']).daily['a/2026-09-01'], source.daily['a/2026-09-01']);
  assert.deepEqual(Merge.merge(data({}, projects), ['a', 'b']).daily, {});
});
test('night default keeps gaps off and explicit legacy day records unchanged', () => {
  const source = data({ 'a/2026-08-31': { day_work: '既存の昼作業' } }, [
    p('a', '2026-08-30', '2026-08-31', { default_shift: 'night' }),
    p('b', '2026-09-03', '2026-09-04', { default_shift: 'night' })]);
  const merged = Merge.merge(source, ['a', 'b']);
  assert.deepEqual(merged.daily['a/2026-08-31'], { day_work: '既存の昼作業' });
  assert.deepEqual(merged.daily['a/2026-09-01'], { day: false, night: false });
  assert.deepEqual(merged.daily['a/2026-09-02'], { day: false, night: false });
  assert.equal(merged.projects[0].default_shift, 'night');
});
test('weekend candidate compatibility treats absent/off alike and separates work', () => {
  const projects = [p('legacy'), p('off', undefined, undefined, { weekend_policy: 'off' }),
    p('work1', undefined, undefined, { weekend_policy: 'work' }),
    p('work2', undefined, undefined, { weekend_policy: 'work' })];
  assert.deepEqual(Merge.findCandidates(projects).map(g => g.ids.slice().sort()), [['legacy', 'off'], ['work1', 'work2']]);
  assert.deepEqual(Merge.findCandidates(projects).map(g => g.projects[0].weekend_policy), ['off', 'work']);
  assert.throws(() => Merge.preview(data({}, projects), ['legacy', 'work1']), /土日の稼働設定が異なる/);
  assert.equal(Object.hasOwn(Merge.merge(data({}, projects), ['legacy', 'off']).projects.find(p => p.id === 'legacy'), 'weekend_policy'), false);
  assert.equal(Merge.merge(data({}, projects), ['off', 'legacy']).projects.find(p => p.id === 'off').weekend_policy, 'off');
});
test('identical automatic/explicit weekend off has no conflict and becomes explicit', () => {
  const source = data({
    'a/2026-08-30': { day: false, night: false, _weekend_auto: 'off' },
    'b/2026-08-30': { day: false, night: false },
    'other/2026-08-30': { day: false, night: false, _weekend_auto: 'off' }
  }, [p('a'), p('b'), p('other')]);
  const before = structuredClone(source);
  const plan = Merge.preview(source, ['a', 'b']);
  assert.equal(plan.conflicts.length, 0);
  assert.deepEqual(plan.daily['2026-08-30'], { day: false, night: false });
  const merged = Merge.merge(source, ['a', 'b']);
  assert.deepEqual(merged.daily['a/2026-08-30'], { day: false, night: false });
  assert.deepEqual(merged.daily['other/2026-08-30'], source.daily['other/2026-08-30']);
  assert.deepEqual(source, before);
});
test('selected conflict and out-of-period auto markers are stripped without losing details', () => {
  const source = data({
    'a/2026-08-30': { day: false, night: false, _weekend_auto: 'off', day_work: '指定休み', custom: '保持' },
    'a/2026-07-05': { day: false, night: false, _weekend_auto: 'off', day_priority_detail: '保持' }
  });
  assert.equal(Merge.preview(source, ['a', 'b']).conflicts.length, 1);
  const merged = Merge.merge(source, ['a', 'b'], { '2026-08-30': 'a' });
  assert.deepEqual(merged.daily['a/2026-08-30'], { day: false, night: false, day_work: '指定休み', custom: '保持' });
  assert.deepEqual(merged.daily['a/2026-07-05'], { day: false, night: false, day_priority_detail: '保持' });
  assert.equal(source.daily['a/2026-08-30']._weekend_auto, 'off');
});
test('work-policy night projects retain implicit weekends', () => {
  const projects = [p('a', undefined, undefined, { default_shift: 'night', weekend_policy: 'work' }),
    p('b', undefined, undefined, { default_shift: 'night', weekend_policy: 'work' })];
  assert.deepEqual(Merge.merge(data({}, projects), ['a', 'b']).daily, {});
  assert.equal(Merge.merge(data({}, projects), ['a', 'b']).projects[0].weekend_policy, 'work');
  const source = data({ 'a/2026-08-30': { day: false, night: false, _weekend_auto: 'off' } }, projects);
  const plan = Merge.preview(source, ['a', 'b']);
  assert.deepEqual(plan.conflicts[0].choices[1].value, { day: false, night: true });
});
