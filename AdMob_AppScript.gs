// ============================================================
// XGROWTH CONSOLE — Google Apps Script
// Reads from Google Sheet (dummy or real AdMob export)
// and serves JSON to the XGrowth Console UI
// ============================================================

// ── CONFIGURATION ───────────────────────────────────────────

var SPREADSHEET_ID = '1nnn3vTdekcOD3u3GcBgA2EfXmCPXWqlWJjpVhH7yrBw';

var CLIENT = {
  id:      'A',
  name:    'Acme Apps',
  since:   'Jan 2023',
  model:   'Ads + IAP',
};

// ── ENTRY POINT ─────────────────────────────────────────────

function doGet(e) {
  try {
    var data = buildAllData();
    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── MAIN BUILD FUNCTION ──────────────────────────────────────

function buildAllData() {
  var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  var netRows = readSheet(ss, 'network_report');
  var medRows = readSheet(ss, 'mediation_report');

  var byApp   = aggregateByApp(netRows);
  var byFmt   = aggregateByFormat(medRows);
  var daily   = aggregateDaily(netRows);

  return {
    clients:  buildClients(byApp, daily),
    apps:     buildApps(byApp, byFmt, daily),
    summary:  buildSummary(byApp),
    weekKpis: buildWeekKpis(byApp, daily),
  };
}

// ── READ A SHEET TAB INTO ARRAY OF OBJECTS ───────────────────

function readSheet(ss, sheetName) {
  var sheet  = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  var headers = values[0].map(function(h) { return String(h).trim(); });
  var rows    = [];
  for (var i = 1; i < values.length; i++) {
    var row = {};
    headers.forEach(function(h, j) { row[h] = values[i][j]; });
    // Skip completely empty rows
    if (!row['app_id'] || row['app_id'] === '') continue;
    rows.push(row);
  }
  return rows;
}

// ── AGGREGATE network_report BY APP ─────────────────────────

function aggregateByApp(rows) {
  var byApp = {};
  rows.forEach(function(r) {
    var id   = String(r['app_id']).trim();
    var name = String(r['app_name']).trim();
    if (!byApp[id]) {
      byApp[id] = {
        id:               id,
        name:             name,
        earn:             0,   // sum of estimated_earnings (micros)
        impressions:      0,
        clicks:           0,
        adRequests:       0,
        matchedRequests:  0,
        rpmSum:           0,   // to average eCPM
        ctrSum:           0,
        matchRateSum:     0,
        rowCount:         0,
      };
    }
    var a = byApp[id];
    a.earn            += num(r['estimated_earnings']);
    a.impressions     += num(r['impressions']);
    a.clicks          += num(r['clicks']);
    a.adRequests      += num(r['ad_requests']);
    a.matchedRequests += num(r['matched_requests']);
    a.rpmSum          += num(r['impression_rpm']);
    a.ctrSum          += num(r['impression_ctr']);
    a.matchRateSum    += num(r['match_rate']);
    a.rowCount++;
  });
  return byApp;
}

// ── AGGREGATE mediation_report BY APP + FORMAT ───────────────

function aggregateByFormat(rows) {
  // Returns: { app_id: { FORMAT: { earn, impressions, rpm, ctr, matchRate } } }
  var byFmt = {};
  rows.forEach(function(r) {
    var id  = String(r['app_id']).trim();
    var fmt = String(r['format']).trim();
    if (!byFmt[id]) byFmt[id] = {};
    if (!byFmt[id][fmt]) byFmt[id][fmt] = { earn:0, impressions:0, rpmSum:0, ctrSum:0, matchRateSum:0, rowCount:0 };
    var f = byFmt[id][fmt];
    f.earn         += num(r['estimated_earnings']);
    f.impressions  += num(r['impressions']);
    f.rpmSum       += num(r['impression_rpm']);
    f.ctrSum       += num(r['impression_ctr']);
    f.matchRateSum += num(r['match_rate']);
    f.rowCount++;
  });
  return byFmt;
}

// ── AGGREGATE DAILY TOTALS (for spark lines) ─────────────────

function aggregateDaily(rows) {
  // Returns: { app_id: { date: earn_in_dollars } }
  // Also builds overall daily total across all apps
  var byAppDate = {};
  var allDates  = {};

  rows.forEach(function(r) {
    var id   = String(r['app_id']).trim();
    var date = String(r['date']).trim();
    var earn = num(r['estimated_earnings']) / 1e6;

    if (!byAppDate[id]) byAppDate[id] = {};
    byAppDate[id][date] = (byAppDate[id][date] || 0) + earn;
    allDates[date]      = (allDates[date] || 0) + earn;
  });

  // Sort dates and convert to arrays
  var sortedDates = Object.keys(allDates).sort();

  var result = { _allDates: sortedDates, _allEarnings: sortedDates.map(function(d){return allDates[d];}) };
  Object.keys(byAppDate).forEach(function(id) {
    result[id] = sortedDates.map(function(d) { return byAppDate[id][d] || 0; });
  });
  return result;
}

// ── BUILD: clients array ─────────────────────────────────────

function buildClients(byApp, daily) {
  var totalEarn = 0;
  Object.keys(byApp).forEach(function(id) { totalEarn += byApp[id].earn; });

  return [{
    id:           CLIENT.id,
    name:         CLIENT.name,
    initials:     initials(CLIENT.name),
    since:        CLIENT.since,
    model:        CLIENT.model,
    apps:         Object.keys(byApp).length,
    revenue:      '$' + dollars(totalEarn),
    delta:        '▲ live',
    dailyRevenue: daily._allEarnings || [],
  }];
}

// ── BUILD: apps array ────────────────────────────────────────

function buildApps(byApp, byFmt, daily) {
  return Object.keys(byApp).map(function(id) {
    var a    = byApp[id];
    var n    = a.rowCount || 1;
    var earn = a.earn / 1e6;
    var ecpm = (a.rpmSum / n) / 1e6;
    var ctr  = (a.ctrSum / n) * 100;
    var fill = (a.matchRateSum / n) * 100;

    // Format breakdown for Ad Units tab
    var fmtData  = byFmt[id] || {};
    var formats  = ['BANNER','INTERSTITIAL','REWARDED','REWARDED_INTERSTITIAL','NATIVE','APP_OPEN'];
    var adUnits  = formats.filter(function(f){return fmtData[f];}).map(function(fmt) {
      var f  = fmtData[fmt];
      var fn = f.rowCount || 1;
      return {
        format:      fmt,
        revenue:     '$' + dollars(f.earn),
        impressions: formatNum(f.impressions),
        ecpm:        '$' + (((f.rpmSum/fn)/1e6)).toFixed(2),
        ctr:         ((f.ctrSum/fn)*100).toFixed(2) + '%',
        fillRate:    ((f.matchRateSum/fn)*100).toFixed(1) + '%',
      };
    });

    return {
      id:           id,
      clientId:     CLIENT.id,
      name:         a.name,
      pkg:          id,
      icon:         initials(a.name),
      dailyRevenue: daily[id] || [],
      adUnits:      adUnits,
      metrics: [
        { label:'Ad Revenue',   val:'$' + dollars(a.earn),          delta:'▲ live' },
        { label:'Impressions',  val:formatNum(a.impressions),        delta:'▲ live' },
        { label:'eCPM',         val:'$' + ecpm.toFixed(2),          delta:'▲ live' },
        { label:'CTR',          val:ctr.toFixed(2) + '%',           delta:'▲ live' },
        { label:'Fill Rate',    val:fill.toFixed(1) + '%',          delta:'▲ live' },
        { label:'Ad Requests',  val:formatNum(a.adRequests),        delta:'▲ live' },
      ],
    };
  });
}

// ── BUILD: summary block (shown in workspace) ────────────────

function buildSummary(byApp) {
  var totalEarn = 0, totalImpr = 0, totalReq = 0, totalMatched = 0;
  var ecpmSum   = 0, fillSum   = 0, n = 0;

  Object.keys(byApp).forEach(function(id) {
    var a = byApp[id];
    totalEarn    += a.earn;
    totalImpr    += a.impressions;
    totalReq     += a.adRequests;
    totalMatched += a.matchedRequests;
    ecpmSum      += a.rpmSum / (a.rowCount || 1);
    fillSum      += a.matchRateSum / (a.rowCount || 1);
    n++;
  });

  var avgEcpm = n ? (ecpmSum / n) / 1e6 : 0;
  var avgFill = n ? (fillSum / n) * 100  : 0;
  var overall = totalReq ? (totalMatched / totalReq) * 100 : 0;

  var result = {};
  result[CLIENT.id] = [
    { label:'Total Revenue',    val:'$' + dollars(totalEarn),    delta:'▲ live' },
    { label:'Total Impressions',val:formatNum(totalImpr),         delta:'▲ live' },
    { label:'Avg eCPM',         val:'$' + avgEcpm.toFixed(2),    delta:'▲ live' },
    { label:'Avg Fill Rate',    val:avgFill.toFixed(1) + '%',    delta:'▲ live' },
    { label:'Overall Match Rate',val:overall.toFixed(1) + '%',   delta:'▲ live' },
    { label:'Active Apps',      val:String(n),                    delta:''       },
  ];
  return result;
}

// ── BUILD: weekKpis for nav sidebar ─────────────────────────

function buildWeekKpis(byApp, daily) {
  var totalEarn = 0, totalImpr = 0, ecpmSum = 0, fillSum = 0, n = 0;
  Object.keys(byApp).forEach(function(id) {
    var a = byApp[id];
    totalEarn += a.earn;
    totalImpr += a.impressions;
    ecpmSum   += a.rpmSum / (a.rowCount || 1);
    fillSum   += a.matchRateSum / (a.rowCount || 1);
    n++;
  });

  var avgEcpm    = n ? (ecpmSum/n)/1e6 : 0;
  var avgFill    = n ? (fillSum/n)*100  : 0;
  var revSpark   = makeSpark(daily._allEarnings || [], 64, 22);
  var imprSpark  = makeSpark((daily._allEarnings||[]).map(function(v){return v*1000;}), 64, 22);

  return [
    { label:'Revenue',     val:'$' + dollars(totalEarn),    delta:'▲ live', spark:revSpark  },
    { label:'Impressions', val:formatNum(totalImpr),         delta:'▲ live', spark:imprSpark },
    { label:'Avg eCPM',    val:'$' + avgEcpm.toFixed(2),    delta:'▲ live', spark:revSpark  },
    { label:'Fill Rate',   val:avgFill.toFixed(1) + '%',    delta:'▲ live', spark:imprSpark },
  ];
}

// ── HELPERS ──────────────────────────────────────────────────

function num(v) {
  var n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function dollars(micros) {
  var v = micros / 1e6;
  if (v >= 1e6) return (v/1e6).toFixed(2) + 'M';
  if (v >= 1e3) return (v/1e3).toFixed(1) + 'K';
  return v.toFixed(2);
}

function formatNum(n) {
  if (n >= 1e9) return (n/1e9).toFixed(1) + 'B';
  if (n >= 1e6) return (n/1e6).toFixed(1) + 'M';
  if (n >= 1e3) return (n/1e3).toFixed(1) + 'K';
  return String(Math.round(n));
}

function initials(name) {
  return (name || '').split(' ').map(function(w){return w[0]||'';}).join('').slice(0,2).toUpperCase();
}

function makeSpark(vals, w, h) {
  if (!vals || !vals.length) return '';
  var mx = Math.max.apply(null, vals) || 1;
  return vals.map(function(v, i) {
    return (i / (vals.length - 1) * w).toFixed(1) + ',' + (h - (v/mx) * (h-3)).toFixed(1);
  }).join(' ');
}

// ── TEST — run this inside Apps Script to verify ─────────────

function testSetup() {
  try {
    var data = buildAllData();
    Logger.log('✅ Clients:  ' + data.clients.length);
    Logger.log('✅ Apps:     ' + data.apps.length);
    Logger.log('✅ Revenue:  ' + (data.clients[0] ? data.clients[0].revenue : '—'));
    Logger.log('✅ KPIs:     ' + data.weekKpis.length);
    Logger.log('✅ App names: ' + data.apps.map(function(a){return a.name;}).join(', '));
    data.apps.forEach(function(a) {
      Logger.log('   ' + a.name + ' → ' + a.metrics[0].val + ' revenue, ' + a.adUnits.length + ' ad formats');
    });
  } catch(e) {
    Logger.log('❌ Error: ' + e.message);
  }
}

