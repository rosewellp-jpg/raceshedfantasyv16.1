
const WORKBOOK_PATH = './Race_Shed_Fantasy_LeagueV1.xlsx';

const DRIVER_NAMES = {
  "1":"Ross Chastain","2":"Austin Cindric","3":"Austin Dillon","4":"Noah Gragson","5":"Kyle Larson",
  "6":"Brad Keselowski","7":"Justin Haley","8":"Kyle Busch","9":"Chase Elliott","10":"Ty Dillon",
  "11":"Denny Hamlin","12":"Ryan Blaney","16":"AJ Allmendinger","17":"Chris Buescher","19":"Chase Briscoe",
  "20":"Christopher Bell","21":"Josh Berry","22":"Joey Logano","23":"Bubba Wallace","24":"William Byron",
  "34":"Todd Gilliland","35":"Riley Herbst","38":"Zane Smith","41":"Cole Custer","42":"John H. Nemechek",
  "43":"Erik Jones","45":"Tyler Reddick","47":"Ricky Stenhouse Jr.","48":"Alex Bowman","51":"Cody Ware",
  "54":"Ty Gibbs","60":"Ryan Preece","71":"Michael McDowell","77":"Carson Hocevar","88":"Shane van Gisbergen",
  "99":"Daniel Suárez"
};
const TRACK_TYPES = {
  1:"Superspeedway",2:"Intermediate",3:"Road Course",4:"Short Track",5:"Intermediate",6:"Short Track",
  7:"Short Track",8:"Road Course",9:"Short Track",10:"Superspeedway",11:"Intermediate",12:"Intermediate",
  13:"All-Star / Intermediate",14:"Intermediate",15:"Intermediate",16:"Road Course",17:"Road Course",
  18:"Short Track",19:"Street Course",20:"Road Course",21:"Intermediate",22:"Superspeedway",
  23:"Short Track",24:"Intermediate",25:"Short Track",26:"Road Course",27:"Intermediate",
  28:"Short Track",29:"Road Course",30:"Intermediate",31:"Short Track",32:"Superspeedway",
  33:"Road Course",34:"Short Track",35:"Intermediate",36:"Championship / Short Track"
};
const COLOR_MAP = {
  "5":["#f6c544","#111"],"20":["#ffe16a","#111"],"23":["#58a8ff","#08111a"],"11":["#e10600","#fff"],
  "12":["#d9dde4","#08111a"],"9":["#25d75f","#08111a"],"24":["#ffd94a","#08111a"],"45":["#66e0d8","#08111a"],
  "17":["#2f8bff","#fff"],"19":["#f35b4f","#fff"],"22":["#ffec8a","#111"],"1":["#ff6b6b","#fff"]
};

function el(id){ return document.getElementById(id); }
function asNum(v){ const n = Number(v); return Number.isFinite(n) ? n : null; }
function normCar(v){ return String(v ?? '').replace(/\.0$/,'').trim(); }
function colorFor(car){ return COLOR_MAP[car] || ['#dfe6f1','#08111a']; }

async function loadWorkbook(){
  const res = await fetch(WORKBOOK_PATH, { cache: 'no-store' });
  if(!res.ok) throw new Error(`Could not load workbook: ${res.status}`);
  const buf = await res.arrayBuffer();
  return XLSX.read(buf, { type: 'array' });
}
function sheetToRows(wb, name){
  const ws = wb.Sheets[name];
  if(!ws) return [];
  return XLSX.utils.sheet_to_json(ws, { defval: null });
}
function latestRaceWithResults(results){
  const nums = [...new Set(results.filter(r => asNum(r.finish)).map(r => asNum(r.raceNumber)))].filter(Boolean);
  return nums.length ? Math.max(...nums) : 0;
}
function raceScore(player, raceNo, picks, pointsByRaceCar){
  const row = picks.find(p => p.player === player && p.raceNumber === raceNo);
  if(!row) return 0;
  return [row.pick1,row.pick2,row.pick3].reduce((sum, car) => sum + (pointsByRaceCar[`${raceNo}-${normCar(car)}`] || 0), 0);
}
function badge(car){
  if(!car) return '';
  const carStr = normCar(car);
  const name = DRIVER_NAMES[carStr] || 'Driver TBD';
  const colors = colorFor(carStr);
  return `<span class="badge"><span class="num" style="background:${colors[0]};color:${colors[1]}">#${carStr}</span><span>${name}</span></span>`;
}
function getFeaturedRaceNumber(schedule, results){
  const latest = latestRaceWithResults(results);
  if(latest > 0) {
    const next = latest + 1;
    if(schedule.some(r => r.raceNumber === next)) return next;
  }
  return 5;
}

function renderDashboard(data){
  const { players, schedule, picks, results, featuredRaceNo, featuredPicks, featuredPot, pointsByRaceCar, standings, latestCompleted, driverPoints, mostPicked, trendingDrivers } = data;
  const featuredRace = schedule.find(r => r.raceNumber === featuredRaceNo) || schedule[0] || {};
  const completedRaceNos = [...new Set(results.map(r => r.raceNumber))].sort((a,b)=>a-b);

  el('workbookStatus').textContent = 'Connected';
  el('featuredRaceName').textContent = featuredRace.raceName || 'Race Shed Fantasy';
  el('featuredMeta').textContent = [featuredRace.date, featuredRace.track].filter(Boolean).join(' • ');
  el('heroPills').innerHTML = [
    `<span class="pill">$5 / race</span>`,
    `<span class="pill">${players.length} players in league</span>`,
    `<span class="pill">${featuredPicks.length} entries submitted</span>`,
    `<span class="pill">${completedRaceNos.length} completed league races</span>`
  ].join('');
  el('picksSubmittedCount').textContent = `${featuredPicks.length} / ${players.length}`;
  el('featuredPotValue').textContent = `$${featuredPot}`;
  el('featuredPotPlayers').textContent = `${featuredPicks.length} players submitted for this race`;

  el('trackTitle').textContent = featuredRace.raceName || 'Track details';
  el('trackBadge').textContent = (featuredRace.track || 'AUTO').toString().split(' ')[0].slice(0,4).toUpperCase();
  el('trackGrid').innerHTML = [
    ['Track', featuredRace.track || '-'],
    ['Date', featuredRace.date || '-'],
    ['Track type', featuredRace.trackType || '-'],
    ['Featured pot', `$${featuredPot}`],
    ['Submitted entries', `${featuredPicks.length}`],
    ['Workbook', 'Race_Shed_Fantasy_V7.xlsx']
  ].map(([label, value]) => `<div class="track-item"><div class="label">${label}</div><div class="value">${value}</div></div>`).join('');
  el('trackNote').textContent = 'V16 adds driver season points, most-picked driver excitement, trending drivers, and a stronger logo presentation.';

  el('weeklyTrackBox').innerHTML = `
    <div class="weekly-track-item featured">
      <div class="label">Race</div>
      <div class="value">${featuredRace.raceName || '-'}</div>
      <div class="small-muted">${[featuredRace.date, featuredRace.track].filter(Boolean).join(' • ')}</div>
    </div>
    <div class="weekly-track-item">
      <div class="label">Track type</div>
      <div class="value">${featuredRace.trackType || '-'}</div>
    </div>
    <div class="weekly-track-item">
      <div class="label">Current pot</div>
      <div class="value">$${featuredPot}</div>
    </div>
    <div class="weekly-track-item">
      <div class="label">Picks in</div>
      <div class="value">${featuredPicks.length} / ${players.length}</div>
    </div>
  `;

  el('trendingDriversBox').innerHTML = trendingDrivers.length ? `<div class="trending-list">${
    trendingDrivers.map(d => {
      const colors = colorFor(d.car);
      return `<div class="trend-item">
        <div>
          <div class="driver-title"><span class="big-badge" style="min-width:38px;height:38px;font-size:.9rem;background:${colors[0]};color:${colors[1]}">#${d.car}</span> ${d.driver}</div>
          <div class="driver-sub">Last race: ${d.lastFinishLabel}</div>
        </div>
        <div class="trend-up">+${d.lastPts}</div>
      </div>`;
    }).join('')
  }</div>` : `<div class="empty">Trending drivers appear after results are loaded.</div>`;

  el('statsGrid').innerHTML = [
    {label:'League leader', value: standings[0]?.player || '-'},
    {label:'Leader points', value: String(standings[0]?.seasonPoints || 0)},
    {label:'Featured race pot', value: `$${featuredPot}`},
    {label:'Most picked driver', value: mostPicked ? `#${mostPicked.car} ${mostPicked.driver}` : '-'}
  ].map(s => `<div class="panel stat"><div class="label">${s.label}</div><div class="value">${s.value}</div></div>`).join('');

  const picker = el('racePicker');
  picker.innerHTML = '';
  schedule.forEach(r => {
    const opt = document.createElement('option');
    opt.value = r.raceNumber;
    opt.textContent = `${r.raceNumber}. ${r.raceName}`;
    if(r.raceNumber === featuredRaceNo) opt.selected = true;
    picker.appendChild(opt);
  });

  function renderRace(raceNo){
    const n = Number(raceNo);
    const raceRows = picks.filter(p => p.raceNumber === n && [p.pick1,p.pick2,p.pick3].some(Boolean));
    el('picksSummary').textContent = raceRows.length ? `${raceRows.length} players currently on the board for this race.` : 'No picks entered yet for this race.';
    el('picksBoard').innerHTML = raceRows.length ? `<table>
      <thead><tr><th>Player</th><th>Picks</th><th>Total</th></tr></thead>
      <tbody>${raceRows.map(r => {
        const total = raceScore(r.player, n, picks, pointsByRaceCar);
        return `<tr>
          <td class="player-cell">${r.player}</td>
          <td><div class="pick-badges">${badge(r.pick1)}${badge(r.pick2)}${badge(r.pick3)}</div></td>
          <td>${total || '-'}</td>
        </tr>`;
      }).join('')}</tbody>
    </table>` : `<div class="empty">Awaiting picks.</div>`;

    const raceResults = results.filter(r => r.raceNumber === n).sort((a,b)=>a.finish-b.finish).slice(0,12);
    el('resultsTitle').textContent = raceResults.length ? `${raceResults[0].raceName} results` : 'Race results';
    el('raceResults').innerHTML = raceResults.length ? `<table>
      <thead><tr><th>Fin</th><th>Car</th><th>Driver</th><th>Pts</th></tr></thead>
      <tbody>${raceResults.map(r => `<tr><td>${r.finish}</td><td>#${r.carNumber}</td><td>${DRIVER_NAMES[r.carNumber] || '-'}</td><td>${r.finishPts}</td></tr>`).join('')}</tbody>
    </table>` : `<div class="empty">No race results loaded yet.</div>`;
  }
  picker.addEventListener('change', e => renderRace(e.target.value));
  renderRace(featuredRaceNo);

  el('standingsBoard').innerHTML = `<table>
    <thead><tr><th>Rank</th><th>Player</th><th>Season Pts</th><th>Wins</th><th>Best Week</th></tr></thead>
    <tbody>${standings.map((s,i)=> `<tr><td class="rank">${i+1}</td><td class="player-cell">${s.player}</td><td>${s.seasonPoints}</td><td>${s.wins}</td><td>${s.bestWeek}</td></tr>`).join('')}</tbody>
  </table>`;

  el('latestScoring').innerHTML = latestCompleted ? (() => {
    const latestDriverPts = [];
    const seen = new Set();
    results.filter(r => r.raceNumber === latestCompleted).sort((a,b)=> a.finish - b.finish).forEach(r => {
      if(seen.has(r.carNumber)) return;
      seen.add(r.carNumber);
      latestDriverPts.push({ car:r.carNumber, driver:DRIVER_NAMES[r.carNumber] || 'Driver TBD', pts:r.finishPts });
    });
    return latestDriverPts.length ? `<div class="latest-list">${
      latestDriverPts.slice(0,12).map(r => `<div class="latest-item"><div class="driver-title">#${r.car} ${r.driver}</div><div class="pts">${r.pts} pts</div></div>`).join('')
    }</div>` : `<div class="empty">Latest scoring will appear once race points are loaded.</div>`;
  })() : `<div class="empty">Latest scoring will appear once race points are loaded.</div>`;

  el('driverPointsSource').textContent = 'Workbook result totals';
  el('driverPointsBoard').innerHTML = driverPoints.length ? `<div class="driver-points-list">${
    driverPoints.slice(0,10).map(d => {
      const colors = colorFor(d.car);
      return `<div class="driver-points-item">
        <div>
          <div class="driver-title"><span class="big-badge" style="min-width:38px;height:38px;font-size:.9rem;background:${colors[0]};color:${colors[1]}">#${d.car}</span> ${d.driver}</div>
          <div class="driver-sub">${d.wins} wins • avg finish ${d.avgFinish}</div>
        </div>
        <div class="pts">${d.points} pts</div>
      </div>`;
    }).join('')
  }</div>` : `<div class="empty">Driver season points will appear as results accumulate.</div>`;

  if(mostPicked){
    const colors = colorFor(mostPicked.car);
    el('mostPickedMini').textContent = `#${mostPicked.car} ${mostPicked.driver}`;
    el('mostPickedDriverBox').innerHTML = `
      <div class="most-picked-box">
        <div class="most-picked-head">
          <div class="big-badge" style="background:${colors[0]};color:${colors[1]}">#${mostPicked.car}</div>
          <div>
            <div class="most-picked-name">${mostPicked.driver}</div>
            <div class="small-muted">Most selected driver across the season so far</div>
          </div>
        </div>
        <div class="most-picked-stats">
          <div class="most-picked-stat"><div class="label">Times picked</div><div class="value">${mostPicked.timesPicked}</div></div>
          <div class="most-picked-stat"><div class="label">Season points</div><div class="value">${mostPicked.points}</div></div>
          <div class="most-picked-stat"><div class="label">Avg finish</div><div class="value">${mostPicked.avgFinish}</div></div>
        </div>
      </div>
    `;
  } else {
    el('mostPickedDriverBox').innerHTML = `<div class="empty">Most-picked driver will appear once picks are in.</div>`;
  }

  if(latestCompleted){
    const podium = players.map(player => ({ player, score: raceScore(player, latestCompleted, picks, pointsByRaceCar) }))
      .sort((a,b)=> b.score - a.score || a.player.localeCompare(b.player))
      .slice(0,3);
    el('podiumBoard').innerHTML = `<div class="podium">${
      podium.map((p,i) => `<div class="podium-card ${i===0?'first':''}"><div class="spot">${['1st','2nd','3rd'][i]}</div><div class="name">${p.player}</div><div class="score">${p.score} pts</div></div>`).join('')
    }</div>`;
  } else {
    el('podiumBoard').innerHTML = '';
  }

  el('scheduleBoard').innerHTML = `<div class="schedule-list">${
    schedule.map(r => {
      const cls = completedRaceNos.includes(r.raceNumber) ? 'completed' : 'pending';
      return `<div class="schedule-item ${cls}">
        <div class="schedule-num">${r.raceNumber}</div>
        <div><div class="schedule-name">${r.raceName}</div><div class="schedule-meta">${[r.date, r.track].filter(Boolean).join(' • ')}</div></div>
        <div class="schedule-type">${r.trackType || '-'}</div>
      </div>`;
    }).join('')
  }</div>`;
}

async function init(){
  const wb = await loadWorkbook();

  const scheduleRows = sheetToRows(wb, 'Schedule').map(r => ({
    raceNumber: asNum(r['Race #']),
    raceName: r['Race'],
    date: r['Date'],
    track: r['Track'],
    tv: r['TV'],
    trackType: r['Track Type'] || TRACK_TYPES[asNum(r['Race #'])] || ''
  })).filter(r => r.raceNumber);

  const picksRows = sheetToRows(wb, 'WeeklyPicks').map(r => ({
    raceNumber: asNum(r['Race #']),
    raceName: r['Race'],
    player: r['Player'],
    pick1: r['Pick 1'],
    pick2: r['Pick 2'],
    pick3: r['Pick 3'],
    weeklyTotal: asNum(r['Weekly Total']) || 0
  })).filter(r => r.raceNumber && r.player);

  const resultsRows = sheetToRows(wb, 'RaceResults').map(r => ({
    raceNumber: asNum(r['Race #']),
    raceName: r['Race'],
    finish: asNum(r['Finish']),
    carNumber: normCar(r['Car #']),
    finishPts: asNum(r['Finish Pts']) || 0
  })).filter(r => r.raceNumber && r.finish && r.carNumber);

  const playerRows = sheetToRows(wb, 'Players').map(r => String(r['Player'] || '').trim()).filter(Boolean);

  const featuredRaceNo = getFeaturedRaceNumber(scheduleRows, resultsRows);
  const featuredPicks = picksRows.filter(p => p.raceNumber === featuredRaceNo && [p.pick1,p.pick2,p.pick3].some(Boolean));
  const featuredPot = featuredPicks.length * 5;

  const pointsByRaceCar = {};
  const driverTotals = {};
  resultsRows.forEach(r => {
    pointsByRaceCar[`${r.raceNumber}-${r.carNumber}`] = r.finishPts;
    if(!driverTotals[r.carNumber]) driverTotals[r.carNumber] = { car:r.carNumber, driver:DRIVER_NAMES[r.carNumber] || 'Driver TBD', points:0, finishes:[], wins:0, lastPts:0, lastFinishLabel:'-' };
    driverTotals[r.carNumber].points += r.finishPts;
    driverTotals[r.carNumber].finishes.push(r.finish);
    if(r.finish === 1) driverTotals[r.carNumber].wins += 1;
  });

  const latestCompleted = latestRaceWithResults(resultsRows);
  resultsRows.filter(r => r.raceNumber === latestCompleted).forEach(r => {
    if(driverTotals[r.carNumber]){
      driverTotals[r.carNumber].lastPts = r.finishPts;
      driverTotals[r.carNumber].lastFinishLabel = `${r.finish}${r.finish===1?'st':r.finish===2?'nd':r.finish===3?'rd':'th'}`;
    }
  });

  const standings = playerRows.map(player => {
    const playerRowsFiltered = picksRows.filter(p => p.player === player);
    let seasonPoints = 0, bestWeek = 0, wins = 0;
    playerRowsFiltered.forEach(row => {
      const total = raceScore(player, row.raceNumber, picksRows, pointsByRaceCar);
      seasonPoints += total;
      if(total > bestWeek) bestWeek = total;
    });
    return { player, seasonPoints, bestWeek, wins };
  });

  const completedRaceNos = [...new Set(resultsRows.map(r => r.raceNumber))];
  completedRaceNos.forEach(raceNo => {
    const scores = playerRows.map(player => ({ player, score: raceScore(player, raceNo, picksRows, pointsByRaceCar) }));
    const top = Math.max(0, ...scores.map(s => s.score));
    scores.filter(s => s.score === top && top > 0).forEach(s => {
      const found = standings.find(x => x.player === s.player);
      if(found) found.wins += 1;
    });
  });
  standings.sort((a,b)=> b.seasonPoints - a.seasonPoints || b.wins - a.wins || a.player.localeCompare(b.player));

  const pickCounts = {};
  picksRows.forEach(r => {
    [r.pick1, r.pick2, r.pick3].forEach(car => {
      const c = normCar(car);
      if(!c) return;
      pickCounts[c] = (pickCounts[c] || 0) + 1;
    });
  });

  const driverPoints = Object.values(driverTotals).map(d => ({
    ...d,
    avgFinish: d.finishes.length ? (d.finishes.reduce((a,b)=>a+b,0)/d.finishes.length).toFixed(1) : '-'
  })).sort((a,b)=> b.points - a.points || a.driver.localeCompare(b.driver));

  const mostPickedCar = Object.entries(pickCounts).sort((a,b)=> b[1]-a[1] || Number(a[0])-Number(b[0]))[0];
  const mostPicked = mostPickedCar ? {
    car: mostPickedCar[0],
    driver: DRIVER_NAMES[mostPickedCar[0]] || 'Driver TBD',
    timesPicked: mostPickedCar[1],
    points: driverTotals[mostPickedCar[0]] ? driverTotals[mostPickedCar[0]].points : 0,
    avgFinish: driverTotals[mostPickedCar[0]] && driverTotals[mostPickedCar[0]].finishes.length
      ? (driverTotals[mostPickedCar[0]].finishes.reduce((a,b)=>a+b,0)/driverTotals[mostPickedCar[0]].finishes.length).toFixed(1)
      : '-'
  } : null;

  const trendingDrivers = Object.values(driverTotals)
    .filter(d => d.lastPts > 0)
    .sort((a,b)=> b.lastPts - a.lastPts || a.driver.localeCompare(b.driver))
    .slice(0,5)
    .map(d => ({
      car: d.car,
      driver: d.driver,
      lastPts: d.lastPts,
      lastFinishLabel: d.lastFinishLabel
    }));

  renderDashboard({
    players: playerRows,
    schedule: scheduleRows,
    picks: picksRows,
    results: resultsRows,
    featuredRaceNo,
    featuredPicks,
    featuredPot,
    pointsByRaceCar,
    standings,
    latestCompleted,
    driverPoints,
    mostPicked,
    trendingDrivers
  });
}

init().catch(err => {
  el('workbookStatus').textContent = 'Load failed';
  document.body.innerHTML += `<div class="shell"><div class="panel" style="padding:24px;margin-top:20px"><strong>Workbook load failed.</strong><div style="margin-top:8px;color:#9aa9bc">${String(err.message || err)}</div></div></div>`;
});
