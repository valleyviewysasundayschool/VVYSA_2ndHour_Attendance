// LOAD TEST — paste into the browser console while a classroom QR link is open
// (e.g. https://yourname.github.io/2nd-hour-attendance/?class=Sunday%20School).
// It sends fake VISITOR check-ins named "ZZ Load Test 001…", so no real member is touched.
// Afterwards run:  2nd Hour Tools ▸ 🧹 Delete Load-Test Visitors  in the Google Sheet.
(async function () {
  const N = 200;            // number of simulated people
  const SPREAD_SEC = 10;    // how long the crowd takes to arrive (lower = harsher test)
  const date = state.currentDate, cls = state.currentClass;
  if (!date) { console.error('Open a ?class= link first so a date is loaded.'); return; }
  const times = [];
  let ok = 0, failed = 0;
  const t0 = performance.now();
  await Promise.all(Array.from({ length: N }, (_, i) => new Promise(r => setTimeout(r, Math.random() * SPREAD_SEC * 1000)).then(async () => {
    const s = performance.now();
    try {
      const res = await apiRetry('selfCheckInVisitor', {
        className: cls, dateStr: date,
        visitorData: { name: 'ZZ Load Test ' + String(i + 1).padStart(3, '0') }
      }, { tries: 6 });
      if (res && res.ok) ok++; else failed++;
    } catch (e) { failed++; }
    times.push(performance.now() - s);
  })));
  times.sort((a, b) => a - b);
  const pct = p => Math.round(times[Math.min(times.length - 1, Math.floor(times.length * p))]) + ' ms';
  console.log(`Done in ${Math.round((performance.now() - t0) / 1000)}s — saved ${ok}/${N}, failed ${failed}.`);
  console.log(`Latency (incl. retries): median ${pct(0.5)}, 95th ${pct(0.95)}, slowest ${pct(1)}`);
  console.log('Now count the "ZZ Load Test" rows on the Visitors tab — it should equal ' + ok + '.');
})();
