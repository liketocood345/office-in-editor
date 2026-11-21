import { readFileSync, writeFileSync, renameSync } from 'node:fs';
import path from 'node:path';

const args = process.argv.slice(2);
if (args.length !== 1 && (args.length !== 2 || args[0] !== '--suite')) {
  throw new Error('Usage: node analyze.mjs <run-directory> | --suite <sweep-directory>');
}
const suite = args.length === 2;
const root = path.resolve(args.at(-1));
const text = (directory, name) => readFileSync(path.join(directory, name), 'utf8').replace(/^\uFEFF/, '');
const json = (directory, name) => JSON.parse(text(directory, name));

function distribution(values) {
  if (!values.length) return null;
  const sorted = [...values].sort((a, b) => a - b);
  const rank = percentile => sorted[Math.ceil(sorted.length * percentile) - 1];
  return { count: sorted.length, p50: rank(0.5), p95: rank(0.95), max: sorted.at(-1) };
}

function saveMetrics(directory, report) {
  const output = path.join(directory, 'metrics.json');
  writeFileSync(output + '.partial', JSON.stringify(report, null, 2) + '\n', 'utf8');
  renameSync(output + '.partial', output);
}

function loadRun(directory) {
  const environment = json(directory, 'environment.json');
  const events = text(directory, 'events.jsonl').split(/\r?\n/).filter(Boolean).map(line => JSON.parse(line));
  const issues = [];
  if (events.some((event, index) => event.seq !== index + 1 || event.run !== environment.run ||
      event.app !== environment.app || !Number.isFinite(event.tMs) || (index && event.tMs < events[index - 1].tMs))) {
    issues.push('journal_identity_or_order');
  }
  return { environment, result: json(directory, 'result.json'), finished: json(directory, 'finished.json'), events, issues };
}

// D1: a parked Office window raised from outside must end up visible in the tab, never focused off-screen.
function analyzeRaise(directory) {
  const { environment, result, events, issues } = loadRun(directory);
  const baseline = environment.scenario === 'raise-baseline';
  const starts = events.filter(event => event.stage === 'sweep_round_started');
  const rounds = starts.map((start, index) => {
    const next = starts[index + 1]?.seq ?? Infinity;
    const interval = events.filter(event => event.seq > start.seq && event.seq < next);
    const find = stage => interval.find(event => event.stage === stage);
    const sent = find('smoke_parked_raise_sent'), outcome = find('smoke_parked_raise_result');
    const adopted = find('office_foreground_adopted');
    // fitVisibleWindow reports alignment before the adoption event itself.
    const alignment = adopted && sent && interval.find(event => event.stage === 'native_edit_alignment' &&
      event.seq > sent.seq && event.seq < adopted.seq);
    const released = adopted && interval.find(event => event.stage === 'office_foreground_released' && event.seq > adopted.seq);
    const painted = released && interval.find(event => event.stage === 'preview_painted' && event.parked &&
      event.corr.parkId === released.corr.parkId);
    return {
      round: start.round, returnMethod: start.returnMethod, direct: sent?.direct ?? null,
      raised: outcome?.raised === 1, recovered: outcome?.recovered === 1, offscreenAtDeadline: outcome?.offscreenAtDeadline === 1,
      offscreenObserved: interval.some(event => event.stage === 'screen_office_foreground_observed'),
      docked: adopted?.docked === 1, watchdog: adopted?.watchdog === 1, adoptMs: sent && adopted ? adopted.tMs - sent.tMs : null,
      alignmentErrorPx: alignment?.alignmentErrorPx ?? null, returnPaintMs: released && painted ? painted.tMs - released.tMs : null,
      missing: Object.entries(baseline ? { sent, outcome } : { sent, outcome, adopted, alignment, released, painted })
        .filter(([, event]) => !event).map(([name]) => name),
    };
  });
  const count = predicate => rounds.filter(predicate).length;
  const values = key => rounds.flatMap(round => round[key] === null ? [] : [round[key]]);
  const summary = {
    rounds: rounds.length, raised: count(round => round.raised), recovered: count(round => round.recovered),
    offscreenAtDeadline: count(round => round.offscreenAtDeadline), offscreenObserved: count(round => round.offscreenObserved),
    adoptMs: distribution(values('adoptMs')), returnPaintMs: distribution(values('returnPaintMs')),
    alignmentErrorPx: distribution(values('alignmentErrorPx')),
  };
  if (result.sweep?.passed !== environment.cycles || rounds.length !== environment.cycles) issues.push('incomplete_sweep');
  if (rounds.some(round => round.missing.length)) issues.push('missing_transition');
  if (summary.raised !== rounds.length) issues.push('raise_not_triggered');
  if (!environment.build?.native || !environment.build?.extension) issues.push('build_fingerprint');
  if (!baseline) {
    if (summary.recovered !== rounds.length) issues.push('not_recovered');
    if (summary.offscreenAtDeadline || summary.offscreenObserved) issues.push('offscreen_focus');
    if (!summary.alignmentErrorPx || summary.alignmentErrorPx.max > 0) issues.push('alignment');
    if (summary.returnPaintMs && summary.returnPaintMs.p95 > 500) issues.push('return_latency_over_500ms');
  }
  return { environment, scenario: environment.scenario, passed: !issues.length, issues,
    defectReproduced: summary.offscreenAtDeadline > 0, ...summary, roundsDetail: rounds };
}

function analyzeRun(directory) {
  const { environment, result, finished, events, issues } = loadRun(directory);
  if (environment.scenario === 'raise' || environment.scenario === 'raise-baseline') return analyzeRaise(directory);
  const starts = events.filter(event => event.stage === 'sweep_round_started');
  const rounds = starts.map((start, index) => {
    const next = starts[index + 1]?.seq ?? Infinity;
    const interval = events.filter(event => event.seq > start.seq && event.seq < next);
    const click = interval.find(event => event.stage === 'preview_user_click' && event.trusted === true);
    const activation = click?.corr.activationId;
    const forActivation = stage => interval.find(event => event.stage === stage && event.corr.activationId === activation);
    const keyboard = forActivation('office_keyboard_focus_confirmed');
    const office = forActivation('native_foreground_confirmed');
    const released = forActivation('office_foreground_released');
    const park = released?.corr.parkId;
    const forPark = stage => interval.find(event => event.stage === stage && event.corr.parkId === park);
    const parked = forPark('native_parked');
    const startedCapture = forPark('wgc_started_while_parked');
    const sampled = forPark('wgc_parked_frame_sampled');
    const painted = forPark('preview_painted');
    const alignment = forActivation('native_edit_alignment');
    const prewarm = forActivation('capture_prewarm_mouse');
    const missing = Object.entries({ click, keyboard, office, released, parked, startedCapture, sampled, painted, alignment })
      .filter(([, event]) => !event).map(([name]) => name);
    if (!missing.length && !(click.seq < keyboard.seq && keyboard.seq <= office.seq && office.seq < released.seq &&
        released.seq < parked.seq && parked.seq < startedCapture.seq && startedCapture.seq < sampled.seq &&
        sampled.seq < painted.seq && painted.parked && Number.isInteger(painted.corr.frameSeq))) {
      missing.push('transition_order');
    }
    return {
      round: start.round, returnMethod: start.returnMethod, missing,
      activationMs: click && keyboard ? keyboard.tMs - click.tMs : null,
      returnPaintMs: released && painted ? painted.tMs - released.tMs : null,
      returnSampleMs: released && sampled ? sampled.tMs - released.tMs : null,
      alignmentErrorPx: alignment?.alignmentErrorPx ?? null,
      mouseLeadMs: released && prewarm && prewarm.seq < released.seq ? released.tMs - prewarm.tMs : null,
    };
  });
  const idle = events.find(event => event.stage === 'sweep_idle');
  const firstPaint = events.find(event => event.stage === 'preview_painted');
  const idleSamples = events.filter(event => event.stage === 'capture_stats_1s' && idle && firstPaint &&
    event.tMs - event.intervalMs >= firstPaint.tMs && event.seq < idle.seq);
  const idleMs = idleSamples.reduce((sum, event) => sum + event.intervalMs, 0);
  const rate = (value, duration) => duration > 0 ? Number((value / duration * 1000).toFixed(2)) : null;
  const savePassed = ['sourceExists', 'readable', 'hashChanged', 'markerPresent'].every(key => finished.sourceSave?.[key]);
  const offscreenFocusCount = events.filter(event => event.stage === 'screen_office_foreground_observed').length;
  const alignmentErrorPx = distribution(events.filter(event => event.stage === 'native_edit_alignment').map(event => event.alignmentErrorPx));
  if (result.sweep?.passed !== environment.cycles || rounds.length !== environment.cycles) issues.push('incomplete_sweep');
  if (rounds.some(round => round.missing.length)) issues.push('missing_transition');
  if (!savePassed) issues.push('source_save');
  if (!idleMs) issues.push('idle_telemetry');
  if (!alignmentErrorPx || alignmentErrorPx.max > 0) issues.push('alignment');
  if (offscreenFocusCount) issues.push('offscreen_focus');
  if (!environment.build?.native || !environment.build?.extension) issues.push('build_fingerprint');
  const completed = rounds.filter(round => !round.missing.length);
  const verified = events.find(event => event.stage === 'office_window_verified');
  const calibration = events.find(event => event.stage === 'native_clock_calibrated');
  const clockSamples = events.filter(event => event.stage === 'native_clock' && Number.isFinite(event.roundTripMs));
  const clock = calibration ? {
    samples: calibration.samples, roundTripMs: calibration.roundTripMs, hostOffsetMs: calibration.hostOffsetMs,
    afterWindowVerified: Boolean(verified && clockSamples.length && clockSamples[0].seq > verified.seq),
    roundTripSpreadMs: distribution(clockSamples.map(event => event.roundTripMs)),
  } : clockSamples.length ? { samples: clockSamples.length, legacyStartupSample: true,
    roundTripMs: clockSamples[0].roundTripMs, hostOffsetMs: clockSamples[0].hostOffsetMs } : null;
  const idleFrames = idleSamples.reduce((sum, event) => sum + event.framesSent, 0);
  return {
    environment, nativeEnvironment: events.find(event => event.stage === 'native_environment') ?? null,
    passed: !issues.length, issues, rounds, savedOriginal: savePassed,
    activationMs: distribution(completed.map(round => round.activationMs)),
    returnPaintMs: distribution(completed.map(round => round.returnPaintMs)),
    returnSampleMs: distribution(completed.map(round => round.returnSampleMs)),
    mouseLeadMs: distribution(completed.flatMap(round => round.mouseLeadMs === null ? [] : [round.mouseLeadMs])),
    alignmentErrorPx, offscreenFocusCount, clock,
    idle: { hostCpuPct: rate((idle?.hostCpuMs ?? 0) / 10, idle?.intervalMs),
      helperCpuPct: rate(idleSamples.reduce((sum, event) => sum + event.helperCpuMs, 0) / 10, idleMs),
      helperSampleMs: idleMs, framesSent: idleFrames, sentFps: rate(idleFrames, idleMs) },
  };
}

function analyzeSuite(directory) {
  const manifest = json(directory, 'manifest.json');
  const cases = manifest.cases.map(item => {
    if (!/^[a-f0-9]{32}$/.test(item.run ?? '')) return { ...item, passed: false };
    try {
      const report = json(path.join(directory, '..', '..', item.run), 'metrics.json');
      const { passed, issues, activationMs, returnPaintMs, idle, alignmentErrorPx, offscreenFocusCount } = report;
      if (report.scenario?.startsWith('raise')) {
        return { ...item, passed, issues, count: report.rounds, build: report.environment.build, returnPaintMs, alignmentErrorPx,
          raised: report.raised, recovered: report.recovered, offscreenAtDeadline: report.offscreenAtDeadline,
          defectReproduced: report.defectReproduced };
      }
      return { ...item, passed, issues, count: report.rounds.length, build: report.environment.build,
        activationMs, returnPaintMs, idle, alignmentErrorPx, offscreenFocusCount };
    } catch {
      return { ...item, passed: false, issues: ['missing_metrics'] };
    }
  });
  const groups = Map.groupBy(cases, row => [row.app, row.policy, row.scenario ?? 'click'].join('/'));
  const repeats = [];
  for (const [key, rows] of groups) {
    // Baseline raise runs reproduce the defect on purpose and have no return path to time.
    if (key.endsWith('/raise-baseline')) continue;
    const baseline = rows.find(row => row.repeat === 1);
    for (const row of rows.filter(item => item.repeat > 1)) {
      const sameBuild = Boolean(baseline?.build && row.build &&
        baseline.build.native === row.build.native && baseline.build.extension === row.build.extension &&
        baseline.build.mcp === row.build.mcp && baseline.build.addin === row.build.addin);
      const relative = (a, b) => Number.isFinite(a) && Number.isFinite(b) ? Number((100 * Math.abs(b - a) / Math.max(1, a)).toFixed(1)) : null;
      const activationP95Pct = relative(baseline?.activationMs?.p95, row.activationMs?.p95);
      const returnP95Pct = relative(baseline?.returnPaintMs?.p95, row.returnPaintMs?.p95);
      // 2026-11-21 user decision: latency up to 500 ms is not an optimization target; the P95 deviation is informational.
      const passed = sameBuild && [row.activationMs?.p95 ?? 0, row.returnPaintMs?.p95].every(value => Number.isFinite(value) && value <= 500);
      repeats.push({ key, repeat: row.repeat, sameBuild, activationP95Pct, returnP95Pct, passed });
    }
  }
  return { sweep: manifest.sweep, passed: cases.every(row => row.passed) && repeats.every(row => row.passed),
    cases, repeats, totalRounds: cases.reduce((sum, row) => sum + (row.count ?? 0), 0) };
}

const report = suite ? analyzeSuite(root) : analyzeRun(root);
saveMetrics(root, report);
console.log(JSON.stringify(suite ? { passed: report.passed, totalRounds: report.totalRounds,
  failedCases: report.cases.filter(row => !row.passed).map(({ app, policy, scenario, repeat, issues }) => ({ app, policy, scenario, repeat, issues })),
  repeatFailures: report.repeats.filter(row => !row.passed) } : report.scenario?.startsWith('raise') ?
  { passed: report.passed, scenario: report.scenario, issues: report.issues, rounds: report.rounds, raised: report.raised,
    recovered: report.recovered, offscreenAtDeadline: report.offscreenAtDeadline, defectReproduced: report.defectReproduced,
    returnPaintMs: report.returnPaintMs } :
  { passed: report.passed, issues: report.issues, rounds: report.rounds.length,
    activationMs: report.activationMs, returnPaintMs: report.returnPaintMs, idle: report.idle }));
if (!report.passed) process.exitCode = 2;

