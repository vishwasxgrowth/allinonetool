// ============================================================
// XGROWTH CONSOLE — Google Apps Script
// Reads real AdMob export format from Google Sheet
// Client column (B) drives client grouping dynamically
// ============================================================

var SPREADSHEET_ID = '1nnn3vTdekcOD3u3GcBgA2EfXmCPXWqlWJjpVhH7yrBw';

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

function buildAllData() {
  var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  var netRows = readSheet(ss, 'network_report');
  var medRows = readSheet(ss, 'mediation_report');

  var clients  = getClients(netRows);
  var byApp    = aggregateByApp(netRows);
  var byFmt    = aggregateByFormat(medRows);
  var daily    = aggregateDaily(netRows);

  return {
    clients:  buildClients(clients, byApp, daily),
    apps:     buildApps(byApp, byFmt, daily),
    summary:  buildSummary(byApp),
    weekKpis: buildWeekKpis(byApp, daily),
  };
}

// ── READ SHEET ───────────────────────────────────────────────

function readSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0].map(function(h) { return String(h).trim(); });
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var row = {};
    headers.forEach(function(h, j) { row[h] = values[i][j]; });
    var appName = String(row['App'] || '').trim();
    if (!appName || appName === '' || appName === '\u2014') continue;
    rows.push(row);
  }
  return rows;
}

// ── GET UNIQUE CLIENTS ───────────────────────────────────────

function getClients(rows) {
  var seen = {};
  var clients = [];
  rows.forEach(function(r) {
    var name = String(r['Client'] || 'Unknown').trim();
    if (!seen[name]) {
      seen[name] = true;
      clients.push({ id: slugify(name), name: name });
    }
  });
  return clients;
}

// ── AGGREGATE BY APP ─────────────────────────────────────────

function aggregateByApp(rows) {
  var byApp = {};
  rows.forEach(function(r) {
    var appName    = String(r['App']).trim();
    var clientName = String(r['Client'] || 'Unknown').trim();
    var key        = slugify(clientName) + '::' + appName;
    if (!byApp[key]) {
      byApp[key] = {
        id: slugify(appName), name: appName,
        clientId: slugify(clientName), clientName: clientName,
        earn: 0, impressions: 0, clicks: 0, requests: 0, matchedRequests: 0,
        ecpmSum: 0, ctrSum: 0, matchRateSum: 0, rowCount: 0,
      };
    }
    var a = byApp[key];
    a.earn            += num(r['Estimated earnings (USD)']);
    a.impressions     += num(r['Impressions']);
    a.clicks          += num(r['Clicks']);
    a.requests        += num(r['Requests']);
    a.matchedRequests += num(r['Matched requests']);
    a.ecpmSum         += num(r['Observed eCPM (USD)']);
    a.ctrSum          += pct(r['CTR']);
    a.matchRateSum    += pct(r['Match rate']);
    a.rowCount        += 1;
  });
  return byApp;
}

// ── AGGREGATE BY FORMAT ──────────────────────────────────────

function aggregateByFormat(rows) {
  var byApp = {};
  rows.forEach(function(r) {
    var appName    = String(r['App']).trim();
    var clientName = String(r['Client'] || 'Unknown').trim();
    var key        = slugify(clientName) + '::' + appName;
    var fmt        = String(r['Format'] || '').trim().toUpperCase();
    if (!fmt || fmt === '\u2014') return;
    if (!byApp[key]) byApp[key] = {};
    if (!byApp[key][fmt]) byApp[key][fmt] = { earn: 0, impressions: 0 };
    byApp[key][fmt].earn        += num(r['Estimated earnings (USD)']);
    byApp[key][fmt].impressions += num(r['Impressions']);
  });
  return byApp;
}

// ── AGGREGATE DAILY ──────────────────────────────────────────

function aggregateDaily(rows) {
  var byApp = {};
  rows.forEach(function(r) {
    var appName    = String(r['App']).trim();
    var clientName = String(r['Client'] || 'Unknown').trim();
    var key        = slugify(clientName) + '::' + appName;
    var dateRaw    = r['Date'];
    var dateStr    = (dateRaw instanceof Date)
      ? Utilities.formatDate(dateRaw, 'UTC', 'yyyy-MM-dd')
      : String(dateRaw).trim().substring(0, 10);
    if (!byApp[key]) byApp[key] = {};
    byApp[key][dateStr] = (byApp[key][dateStr] || 0) + num(r['Estimated earnings (USD)']);
  });
  var result = {};
  Object.keys(byApp).forEach(function(key) {
    var dates = Object.keys(byApp[key]).sort();
    result[key] = dates.map(function(d) { return { date: d, earn: round2(byApp[key][d]) }; });
  });
  return result;
}

// ── BUILD CLIENTS ────────────────────────────────────────────

function buildClients(clients, byApp, daily) {
  return clients.map(function(c) {
    // Find all apps for this client
    var appKeys = Object.keys(byApp).filter(function(k) {
      return byApp[k].clientId === c.id;
    });

    var totalEarn = 0, totalImpr = 0;
    appKeys.forEach(function(k) {
      totalEarn += byApp[k].earn;
      totalImpr += byApp[k].impressions;
    });

    // Daily revenue across all apps for this client
    var allDates = {};
    appKeys.forEach(function(k) {
      (daily[k] || []).forEach(function(d) { allDates[d.date] = true; });
    });
    var sortedDates = Object.keys(allDates).sort();
    var dailyEarn = sortedDates.map(function(date) {
      var sum = 0;
      appKeys.forEach(function(k) {
        (daily[k] || []).forEach(function(d) { if (d.date === date) sum += d.earn; });
      });
      return round2(sum);
    });

    return {
      id:           c.id,
      name:         c.name,
      since:        '',
      model:        'Ads',
      revenue:      round2(totalEarn),
      impressions:  totalImpr,
      apps:         appKeys.length,
      dailyRevenue: dailyEarn,
    };
  });
}

// ── BUILD APPS ───────────────────────────────────────────────

function buildApps(byApp, byFmt, daily) {
  var apps = [];
  Object.keys(byApp).forEach(function(key) {
    var a    = byApp[key];
    var ecpm = a.rowCount > 0 ? a.ecpmSum / a.rowCount : 0;
    var ctr  = a.rowCount > 0 ? a.ctrSum  / a.rowCount : 0;
    var mr   = a.rowCount > 0 ? a.matchRateSum / a.rowCount : 0;

    var fmtData = byFmt[key] || {};
    var formats = {};
    Object.keys(fmtData).forEach(function(fmt) {
      formats[fmt] = { earn: round2(fmtData[fmt].earn), impressions: fmtData[fmt].impressions };
    });

    apps.push({
      id:           a.id,
      clientId:     a.clientId,
      name:         a.name,
      revenue:      round2(a.earn),
      impressions:  a.impressions,
      clicks:       a.clicks,
      requests:     a.requests,
      ecpm:         round2(ecpm),
      ctr:          round2(ctr * 100),
      matchRate:    round2(mr * 100),
      formats:      formats,
      dailyRevenue: (daily[key] || []).map(function(d) { return d.earn; }),
    });
  });

  apps.sort(function(x, y) { return y.revenue - x.revenue; });
  return apps;
}

// ── BUILD SUMMARY ────────────────────────────────────────────

function buildSummary(byApp) {
  var totalEarn = 0, totalImpr = 0, totalClicks = 0, ecpmVals = [], ctrVals = [];
  Object.keys(byApp).forEach(function(k) {
    var a = byApp[k];
    totalEarn += a.earn; totalImpr += a.impressions; totalClicks += a.clicks;
    if (a.rowCount > 0) { ecpmVals.push(a.ecpmSum / a.rowCount); ctrVals.push(a.ctrSum / a.rowCount); }
  });
  var avg = function(arr) { return arr.length ? arr.reduce(function(s,v){return s+v;},0)/arr.length : 0; };
  return {
    revenue: round2(totalEarn), impressions: totalImpr, clicks: totalClicks,
    ecpm: round2(avg(ecpmVals)), ctr: round2(avg(ctrVals) * 100),
    apps: Object.keys(byApp).length,
  };
}

// ── BUILD WEEK KPIs ──────────────────────────────────────────

function buildWeekKpis(byApp, daily) {
  var allRevByDate = {};
  Object.keys(byApp).forEach(function(k) {
    (daily[k] || []).forEach(function(d) {
      allRevByDate[d.date] = (allRevByDate[d.date] || 0) + d.earn;
    });
  });
  var sorted  = Object.keys(allRevByDate).sort();
  var n       = sorted.length;
  var recent  = sorted.slice(Math.max(0, n-7)).reduce(function(s,d){return s+allRevByDate[d];},0);
  var prev    = sorted.slice(Math.max(0,n-14),Math.max(0,n-7)).reduce(function(s,d){return s+allRevByDate[d];},0);
  var trend   = prev > 0 ? round2((recent-prev)/prev*100) : 0;
  var totalImpr = 0;
  Object.keys(byApp).forEach(function(k) { totalImpr += byApp[k].impressions; });
  var ecpmVals = Object.keys(byApp).map(function(k) {
    var a = byApp[k]; return a.rowCount > 0 ? a.ecpmSum/a.rowCount : 0;
  }).filter(function(v){return v>0;});
  var avgEcpm = ecpmVals.length ? ecpmVals.reduce(function(s,v){return s+v;},0)/ecpmVals.length : 0;
  return [
    { label: 'Revenue (7d)', value: '$' + round2(recent).toFixed(2), trend: trend },
    { label: 'Impressions',  value: fmtNum(totalImpr), trend: 0 },
    { label: 'Avg eCPM',     value: '$' + round2(avgEcpm).toFixed(2), trend: 0 },
    { label: 'Active Apps',  value: String(Object.keys(byApp).length), trend: 0 },
  ];
}

// ── HELPERS ──────────────────────────────────────────────────

function num(v) {
  if (v === null || v === undefined || v === '') return 0;
  var n = parseFloat(String(v).replace(/[,%]/g, ''));
  return isNaN(n) ? 0 : n;
}
function pct(v) {
  if (v === null || v === undefined || v === '') return 0;
  var s = String(v).trim();
  if (s.indexOf('%') !== -1) return parseFloat(s.replace('%','')) / 100;
  var n = parseFloat(s);
  return isNaN(n) ? 0 : (n > 1.1 ? n/100 : n);
}
function round2(v) { return Math.round(v * 100) / 100; }
function slugify(s) { return String(s).toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/^-|-$/g,''); }
function fmtNum(n) {
  if (n >= 1e9) return (n/1e9).toFixed(1)+'B';
  if (n >= 1e6) return (n/1e6).toFixed(1)+'M';
  if (n >= 1e3) return (n/1e3).toFixed(1)+'K';
  return String(n);
}

// ── TEST ─────────────────────────────────────────────────────

function testSetup() {
  var data = buildAllData();
  Logger.log('Clients: ' + data.clients.length);
  data.clients.forEach(function(c) {
    Logger.log('  ' + c.name + ' — ' + c.apps + ' apps, $' + c.revenue);
  });
  Logger.log('Total apps: ' + data.apps.length);
  Logger.log('Revenue: $' + data.summary.revenue);
}
