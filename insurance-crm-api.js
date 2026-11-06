// ============================================================
// INSURANCE CRM v2 — Google Apps Script API
// ============================================================
// SETUP: Extensions → Apps Script → paste → Save → Deploy → New Deployment
//   Type: Web App | Execute as: Me | Access: Anyone
// IMPORTANT: Every time you edit this file, create a NEW deployment version.
// ============================================================

var SPREADSHEET_ID = '1UpsrWs0JB9lMfbAenROJg67xFuqbEomxSha4H9vSaIY';

var SHEETS = {
  VTC:   'Visitor & Super Visa',
  EDGE:  'Edge Benefits',
  LIFE:  'Life & CI Insurance',
  HD:    'H&D Group Benefits',
  LEADS: 'Lead Pipeline'
};

var DATA_ROW = 5; // rows 1-4 are headers

// Module-level row cache — reset each request, shared within one execution
var _rowsCache = null;
function getAllPolicyRowsCached() {
  if (!_rowsCache) _rowsCache = getAllPolicyRows();
  return _rowsCache;
}

// ── YEAR COLORS (same for both API and dashboard) ──────────
var YEAR_COLORS = {
  2019: '#9e9e9e', 2020: '#9b59b6', 2021: '#7b68ee',
  2022: '#c8920a', 2023: '#2d8c4e', 2024: '#4a90d9',
  2025: '#c0392b', 2026: '#f07030', 2027: '#1a4a30'
};

// ── ROUTER ─────────────────────────────────────────────────
function doGet(e) {
  var p = e.parameter;
  var action = p.action || '';
  var result;
  try {
    switch (action) {
      case 'fullData':  result = getFullData();                         break;
      case 'search':    result = searchClients(p.q || '');             break;
      case 'renewals':  result = getRenewals(parseInt(p.days) || 90);  break;
      case 'leads':     result = getLeads();                           break;
      case 'addPolicy':    result = addPolicy(p);                      break;
      case 'addLead':      result = addLead(p);                       break;
      case 'updatePolicy': result = updatePolicy(p);                  break;
      case 'addNote':      result = addNote(p);                       break;
      case 'saveNotes':    result = saveNotes(p);                     break;
      case 'getPolicy':      result = getPolicy(p.segment, p.policyNum);  break;
      case 'getAllPolicies': result = getAllPolicies();                    break;
      case 'getUpcoming':   result = getUpcomingStarts(parseInt(p.days)||90); break;
      case 'getState':      result = getCrmState();                       break;
      case 'saveState':     result = saveCrmState(p);                    break;
      case 'getLeadConversion': result = getLeadConversion();            break;
      case 'combo':
        _rowsCache = null; // fresh read for this request
        _rowsCache = getAllPolicyRows();
        result = {
          fullData:   getFullData(),
          allPolicies: getAllPolicies(),
          renewals:   getRenewals(730),
          upcoming:   getUpcomingStarts(90)
        };
        _rowsCache = null;
        break;
      default:              result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── FULL DATA — one call loads everything ──────────────────
function getFullData() {
  var rows = getAllPolicyRowsCached();
  var today = new Date(); today.setHours(0, 0, 0, 0);
  var curYear  = today.getFullYear();
  var curMonth = today.getMonth();
  var dom      = today.getDate();
  var daysInMonth = new Date(curYear, curMonth + 1, 0).getDate();
  var doy      = getDayOfYear(today);

  // Aggregate buckets
  var ytdComm = 0, lastYrYTDComm = 0, ytdPolicies = 0;
  var thisMonthComm = 0, thisMonthPolicies = 0;
  var thisMonthCommPrevYear = 0, thisMonthPolsPrevYear = 0;
  var allTimeComm = 0;
  var refComm = 0, refCount = 0, refYtdComm = 0;
  var claimsCount = 0, chargebackTotal = 0;
  var renewals30Count = 0;

  var monthlyProductMap = {}; // { year: { monthIdx: { product: {count,comm} } } }
  var monthlyCompanyMap = {}; // { year: { monthIdx: { company: {count,comm} } } }
  var monthlySourceMap  = {}; // { year: { monthIdx: { source:  {count,comm} } } }
  var monthlySegMap     = {}; // { year: { monthIdx: { vtc:{count,comm}, edge:…, life:…, hd:… } } }

  var segMap = {
    vtc:  { ytdComm:0, totalComm:0, active:0, count:0, chargebackTotal:0 },
    edge: { ytdComm:0, totalComm:0, active:0, count:0, chargebackTotal:0 },
    life: { ytdComm:0, totalComm:0, active:0, count:0, chargebackTotal:0 },
    hd:   { ytdComm:0, totalComm:0, active:0, count:0, chargebackTotal:0 }
  };

  var monthlyByYear = {}; // { 2026: [0..11], ... }
  var productMap    = {}; // { "Super Visa": { comm, count, yearComm: {2026: x} } }
  var clientMap     = {}; // { "C-001": { comm, policies: 0 } }
  var sourceMap     = {}; // { "Instagram": { comm, count, clients: {} } }
  var companyMap    = {}; // { "Insure Canada": { comm, count } }
  var saleChannelMap = {}; // { "Rattan / IBR": { comm, count, clients: {} } }
  var provinceMap   = {}; // { "ON": { comm, count } }
  var paymentMap    = {}; // { "Annual": { comm, count } }
  var icComm = 0, icCount = 0, icYtdComm = 0; // Insure Canada
  var birthdayAlerts = []; // people with DOB in next 30 days

  rows.forEach(function(row) {
    var d = parseDate(row.date);
    if (!d) return;
    var ry   = d.getFullYear();
    var rm   = d.getMonth();
    var comm = toNum(row.comm);
    var seg  = row.segment;
    var isActive = (row.status || '').toLowerCase().indexOf('active') !== -1;

    // All-time
    allTimeComm += comm;

    // YTD current year
    if (ry === curYear) {
      ytdComm += comm;
      ytdPolicies++;
      if (rm === curMonth) { thisMonthComm += comm; thisMonthPolicies++; }
    }

    // Last year same period (YoY) + same calendar month last year
    if (ry === curYear - 1) {
      var lastDoy = getDayOfYear(new Date(curYear - 1, curMonth, dom));
      if (getDayOfYear(d) <= lastDoy) lastYrYTDComm += comm;
      if (rm === curMonth) { thisMonthCommPrevYear += comm; thisMonthPolsPrevYear++; }
    }

    // Referral flag: yellow row (Rattan website) OR explicit leadSource
    var isRef = detectReferral(row.saleChannel);
    if (isRef) { refComm += comm; refCount++; if(ry===curYear) refYtdComm += comm; }

    // Claims + chargebacks (VTC only)
    var chargeback = toNum(row.chargeback);
    if (row.hasClaim) { claimsCount++; }
    chargebackTotal += chargeback;

    // Renewals within 30 days (avoids extra sheet read)
    if (row.segment === 'vtc' && row.endDate) {
      var ed30 = parseDate(row.endDate);
      if (ed30) {
        var dl30 = Math.floor((ed30 - today) / 86400000);
        if (dl30 >= 0 && dl30 <= 30) renewals30Count++;
      }
    }

    // Segment
    if (segMap[seg]) {
      segMap[seg].totalComm += comm;
      segMap[seg].count++;
      segMap[seg].chargebackTotal += chargeback;
      if (ry === curYear) segMap[seg].ytdComm += comm;
      if (isActive) segMap[seg].active++;
    }

    // Monthly by year
    if (!monthlyByYear[ry]) monthlyByYear[ry] = new Array(12).fill(0);
    monthlyByYear[ry][rm] = round2(monthlyByYear[ry][rm] + comm);

    // Monthly segment breakdown
    if (!monthlySegMap[ry]) monthlySegMap[ry] = {};
    if (!monthlySegMap[ry][rm]) monthlySegMap[ry][rm] = {
      vtc:{count:0,comm:0}, edge:{count:0,comm:0}, life:{count:0,comm:0}, hd:{count:0,comm:0}
    };
    if (monthlySegMap[ry][rm][seg]) {
      monthlySegMap[ry][rm][seg].count++;
      monthlySegMap[ry][rm][seg].comm += comm;
    }

    // Product
    var prod = row.product || 'Unknown';
    if (!productMap[prod]) productMap[prod] = { comm:0, count:0, yearComm:{} };
    productMap[prod].comm += comm;
    productMap[prod].count++;
    if (!productMap[prod].yearComm[ry]) productMap[prod].yearComm[ry] = 0;
    productMap[prod].yearComm[ry] = round2(productMap[prod].yearComm[ry] + comm);
    // Monthly products breakdown (prod now defined)
    if (!monthlyProductMap[ry]) monthlyProductMap[ry] = {};
    if (!monthlyProductMap[ry][rm]) monthlyProductMap[ry][rm] = {};
    var mp = monthlyProductMap[ry][rm];
    if (!mp[prod]) mp[prod] = { count:0, comm:0 };
    mp[prod].count++;
    mp[prod].comm += comm;

    // Client
    var cid = row.clientId || ('_' + row.clientName);
    if (!clientMap[cid]) clientMap[cid] = { comm:0, policies:0, name: row.clientName||'', phone: row.phone||'' };
    clientMap[cid].comm += comm;
    clientMap[cid].policies++;

    // Lead source
    var src = (row.leadSource || 'Unknown').trim() || 'Unknown';
    if (!sourceMap[src]) sourceMap[src] = { comm:0, count:0, clients:{} };
    sourceMap[src].comm += comm;
    sourceMap[src].count++;
    if (row.clientId) sourceMap[src].clients[row.clientId] = true;
    // Monthly lead source breakdown (src now defined)
    if (!monthlySourceMap[ry]) monthlySourceMap[ry] = {};
    if (!monthlySourceMap[ry][rm]) monthlySourceMap[ry][rm] = {};
    var ms = monthlySourceMap[ry][rm];
    if (!ms[src]) ms[src] = { count:0, comm:0 };
    ms[src].count++;
    ms[src].comm += comm;

    // Company
    var co = (row.company || 'Unknown').trim() || 'Unknown';
    if (!companyMap[co]) companyMap[co] = { comm:0, count:0 };
    companyMap[co].comm += comm;
    companyMap[co].count++;
    // Monthly company breakdown (co now defined)
    if (!monthlyCompanyMap[ry]) monthlyCompanyMap[ry] = {};
    if (!monthlyCompanyMap[ry][rm]) monthlyCompanyMap[ry][rm] = {};
    var mco = monthlyCompanyMap[ry][rm];
    if (!mco[co]) mco[co] = { count:0, comm:0 };
    mco[co].count++;
    mco[co].comm += comm;

    // Sale Channel breakdown
    var sc2 = (row.saleChannel || 'Own Website').trim() || 'Own Website';
    if (!saleChannelMap[sc2]) saleChannelMap[sc2] = { comm:0, count:0, clients:{} };
    saleChannelMap[sc2].comm += comm;
    saleChannelMap[sc2].count++;
    if (row.clientId) saleChannelMap[sc2].clients[row.clientId] = true;

    // Insure Canada KPI
    var scLow = sc2.toLowerCase();
    var isIC = scLow.indexOf('insure canada') !== -1 || scLow.indexOf('insurecanada') !== -1;
    if (isIC) {
      icComm += comm; icCount++;
      if (ry === curYear) icYtdComm += comm;
    }

    // Province breakdown (VTC only)
    if (row.province) {
      var prov = (row.province || '').trim();
      if (prov) {
        if (!provinceMap[prov]) provinceMap[prov] = { comm:0, count:0 };
        provinceMap[prov].comm += comm;
        provinceMap[prov].count++;
      }
    }

    // Payment type breakdown (VTC only)
    if (row.paymentType) {
      var pt2 = (row.paymentType || '').trim() || 'Annual';
      if (!paymentMap[pt2]) paymentMap[pt2] = { comm:0, count:0 };
      paymentMap[pt2].comm += comm;
      paymentMap[pt2].count++;
    }

    // Birthday alerts — insured DOB within 30 days
    if (row.dob && row.segment === 'vtc') {
      var dobStr = (row.dob || '').toString();
      var dobD = parseDate(dobStr);
      if (dobD) {
        // Compare month+day to today
        var thisYearBday = new Date(curYear, dobD.getMonth(), dobD.getDate());
        var daysUntil = Math.floor((thisYearBday - today) / 86400000);
        if (daysUntil < 0) {
          // Already passed this year, check next year
          thisYearBday = new Date(curYear + 1, dobD.getMonth(), dobD.getDate());
          daysUntil = Math.floor((thisYearBday - today) / 86400000);
        }
        if (daysUntil >= 0 && daysUntil <= 30) {
          birthdayAlerts.push({
            name: row.insuredName || row.clientName,
            clientId: row.clientId,
            phone: row.phone,
            dob: row.dob,
            daysUntil: daysUntil,
            age: curYear - dobD.getFullYear() + (daysUntil === 0 ? 0 : 1)
          });
        }
      }
    }
  });

  // ── Build output objects ───────────────────────────────

  // Segments
  Object.keys(segMap).forEach(function(k) {
    var s = segMap[k];
    s.totalComm = round2(s.totalComm);
    s.ytdComm   = round2(s.ytdComm);
    s.avgComm   = s.count > 0 ? round2(s.totalComm / s.count) : 0;
  });

  // Products — sorted by total commission desc
  var products = Object.keys(productMap).map(function(p) {
    var pm = productMap[p];
    return {
      product:   p,
      count:     pm.count,
      totalComm: round2(pm.comm),
      avgComm:   pm.count > 0 ? round2(pm.comm / pm.count) : 0,
      yearComm:  pm.yearComm
    };
  }).sort(function(a, b) { return b.totalComm - a.totalComm; });

  // Client loyalty
  var buckets = { 1:{clients:0,totalComm:0,list:[]}, 2:{clients:0,totalComm:0,list:[]},
                  3:{clients:0,totalComm:0,list:[]}, 4:{clients:0,totalComm:0,list:[]},
                  '5+':{clients:0,totalComm:0,list:[]} };
  Object.values(clientMap).forEach(function(c) {
    var k = c.policies >= 5 ? '5+' : String(c.policies);
    if (buckets[k]) {
      buckets[k].clients++;
      buckets[k].totalComm += c.comm;
      buckets[k].list.push({ name: c.name, phone: c.phone, policies: c.policies, comm: round2(c.comm) });
    }
  });
  var totalClients = Object.keys(clientMap).length;
  Object.keys(buckets).forEach(function(k) {
    var b = buckets[k];
    b.totalComm   = round2(b.totalComm);
    b.avgLifetime = b.clients > 0 ? round2(b.totalComm / b.clients) : 0;
    b.pct         = totalClients > 0 ? round2(b.clients / totalClients * 100) : 0;
    b.list.sort(function(a,b){ return b.comm - a.comm; });
  });

  // Lead sources
  var atc = allTimeComm || 1;
  var leadSources = Object.keys(sourceMap).map(function(s) {
    var d = sourceMap[s];
    return {
      source:        s,
      count:         d.count,
      uniqueClients: Object.keys(d.clients).length,
      totalComm:     round2(d.comm),
      avgComm:       d.count > 0 ? round2(d.comm / d.count) : 0,
      pct:           round2(d.comm / atc * 100)
    };
  }).sort(function(a, b) { return b.totalComm - a.totalComm; });

  // Companies
  var companies = Object.keys(companyMap).map(function(c) {
    return {
      company:   c,
      count:     companyMap[c].count,
      totalComm: round2(companyMap[c].comm),
      avgComm:   companyMap[c].count > 0 ? round2(companyMap[c].comm / companyMap[c].count) : 0
    };
  }).sort(function(a, b) { return b.totalComm - a.totalComm; });

  // Monthly segment — convert to clean arrays
  var monthlyBySegment = {};
  Object.keys(monthlySegMap).forEach(function(y) {
    monthlyBySegment[Number(y)] = {};
    Object.keys(monthlySegMap[y]).forEach(function(m) {
      var s = monthlySegMap[y][m];
      monthlyBySegment[Number(y)][Number(m)] = {
        vtc:  { count:s.vtc.count,  comm:round2(s.vtc.comm),  avgComm:s.vtc.count>0?round2(s.vtc.comm/s.vtc.count):0 },
        edge: { count:s.edge.count, comm:round2(s.edge.comm), avgComm:s.edge.count>0?round2(s.edge.comm/s.edge.count):0 },
        life: { count:s.life.count, comm:round2(s.life.comm), avgComm:s.life.count>0?round2(s.life.comm/s.life.count):0 },
        hd:   { count:s.hd.count,   comm:round2(s.hd.comm),   avgComm:s.hd.count>0?round2(s.hd.comm/s.hd.count):0 }
      };
    });
  });

  // Finalize segment chargebacks
  Object.keys(segMap).forEach(function(k) {
    segMap[k].chargebackTotal = round2(segMap[k].chargebackTotal);
  });

  // Projections
  var projected      = doy > 0 && ytdComm > 0 ? round2(ytdComm / doy * 365)            : 0;
  var monthProj      = dom > 0 && thisMonthComm > 0 ? round2(thisMonthComm / dom * daysInMonth) : 0;
  var avgCommPerPol  = ytdPolicies > 0 ? round2(ytdComm / ytdPolicies)                  : 0;
  var avgCommPerPolMonth = thisMonthPolicies > 0 ? round2(thisMonthComm / thisMonthPolicies) : 0;
  var yoyPct         = lastYrYTDComm > 0 ? round2((ytdComm - lastYrYTDComm) / lastYrYTDComm * 100) : null;

  // Month-over-same-month YoY %
  var monthVsPrevYearPct   = thisMonthCommPrevYear > 0 ? round2((thisMonthComm - thisMonthCommPrevYear) / thisMonthCommPrevYear * 100) : null;
  var monthPolsPrevYearPct = thisMonthPolsPrevYear > 0 ? round2((thisMonthPolicies - thisMonthPolsPrevYear) / thisMonthPolsPrevYear * 100) : null;

  // Monthly products — convert to sorted arrays
  var monthlyProducts = {};
  Object.keys(monthlyProductMap).forEach(function(y) {
    monthlyProducts[Number(y)] = {};
    Object.keys(monthlyProductMap[y]).forEach(function(m) {
      var prods = Object.keys(monthlyProductMap[y][m]).map(function(p) {
        return { product:p, count:monthlyProductMap[y][m][p].count, comm:round2(monthlyProductMap[y][m][p].comm) };
      }).sort(function(a,b){return b.comm-a.comm;});
      monthlyProducts[Number(y)][Number(m)] = prods;
    });
  });

  // Monthly company — convert to sorted arrays
  var monthlyByCompany = {};
  Object.keys(monthlyCompanyMap).forEach(function(y) {
    monthlyByCompany[Number(y)] = {};
    Object.keys(monthlyCompanyMap[y]).forEach(function(m) {
      var cos = Object.keys(monthlyCompanyMap[y][m]).map(function(c) {
        var d = monthlyCompanyMap[y][m][c];
        return { company:c, count:d.count, comm:round2(d.comm), avgComm: d.count>0?round2(d.comm/d.count):0 };
      }).sort(function(a,b){return b.comm-a.comm;});
      monthlyByCompany[Number(y)][Number(m)] = cos;
    });
  });

  // Monthly source — convert to sorted arrays
  var monthlyBySource = {};
  Object.keys(monthlySourceMap).forEach(function(y) {
    monthlyBySource[Number(y)] = {};
    Object.keys(monthlySourceMap[y]).forEach(function(m) {
      var srcs = Object.keys(monthlySourceMap[y][m]).map(function(s) {
        var d = monthlySourceMap[y][m][s];
        return { source:s, count:d.count, comm:round2(d.comm), avgComm: d.count>0?round2(d.comm/d.count):0 };
      }).sort(function(a,b){return b.comm-a.comm;});
      monthlyBySource[Number(y)][Number(m)] = srcs;
    });
  });

  // Years present in data
  var years = Object.keys(monthlyByYear).map(Number).sort();

  // Sale Channel — sorted by comm
  var saleChannels = Object.keys(saleChannelMap).map(function(s) {
    var d = saleChannelMap[s];
    return { channel:s, count:d.count, uniqueClients:Object.keys(d.clients).length,
             totalComm:round2(d.comm), avgComm:d.count>0?round2(d.comm/d.count):0,
             pct:round2(d.comm/(allTimeComm||1)*100) };
  }).sort(function(a,b){return b.totalComm-a.totalComm;});

  // Province — sorted by count
  var provinces = Object.keys(provinceMap).map(function(p) {
    var d = provinceMap[p];
    return { province:p, count:d.count, totalComm:round2(d.comm),
             avgComm:d.count>0?round2(d.comm/d.count):0,
             pct:round2(d.count/rows.length*100) };
  }).sort(function(a,b){return b.count-a.count;});

  // Payment types — sorted by count
  var paymentTypes = Object.keys(paymentMap).map(function(p) {
    var d = paymentMap[p];
    return { type:p, count:d.count, totalComm:round2(d.comm),
             avgComm:d.count>0?round2(d.comm/d.count):0,
             pct:round2(d.count/rows.length*100) };
  }).sort(function(a,b){return b.count-a.count;});

  // Deduplicate birthday alerts by clientId (same person, multiple policies)
  var bdaySeen = {};
  birthdayAlerts = birthdayAlerts.filter(function(b) {
    var key = b.clientId || b.name;
    if (bdaySeen[key]) return false;
    bdaySeen[key] = true;
    return true;
  });
  birthdayAlerts.sort(function(a,b){return a.daysUntil-b.daysUntil;});

  return {
    summary: {
      ytdComm:                round2(ytdComm),
      projected:              projected,
      yoyPct:                 yoyPct,
      ytdPolicies:            ytdPolicies,
      thisMonthComm:          round2(thisMonthComm),
      monthProjected:         monthProj,
      thisMonthPolicies:      thisMonthPolicies,
      avgCommPerPolicy:       avgCommPerPol,
      avgCommPerPolicyMonth:  avgCommPerPolMonth,
      thisMonthCommPrevYear:  round2(thisMonthCommPrevYear),
      thisMonthPolsPrevYear:  thisMonthPolsPrevYear,
      monthVsPrevYearPct:     monthVsPrevYearPct,
      monthPolsPrevYearPct:   monthPolsPrevYearPct,
      allTimeComm:            round2(allTimeComm),
      allTimePolicies:        rows.length,
      totalClients:           totalClients,
      referralCount:          refCount,
      referralComm:           round2(refComm),
      referralYtdComm:        round2(refYtdComm),
      referralPct:            rows.length > 0 ? round2(refCount / rows.length * 100) : 0,
      claimsCount:            claimsCount,
      chargebackTotal:        round2(chargebackTotal),
      renewals30:             renewals30Count,
      icCount:                icCount,
      icComm:                 round2(icComm),
      icYtdComm:              round2(icYtdComm),
      icPct:                  rows.length > 0 ? round2(icCount / rows.length * 100) : 0,
      dom:                    dom,
      daysInMonth:            daysInMonth,
      curMonth:               curMonth,
      curYear:                curYear
    },
    segments:          segMap,
    monthly:           monthlyByYear,
    monthlyProducts:   monthlyProducts,
    monthlyByCompany:  monthlyByCompany,
    monthlyBySource:   monthlyBySource,
    monthlyBySegment:  monthlyBySegment,
    years:             years,
    products:          products,
    leadSources:       leadSources,
    companies:         companies,
    clientLoyalty:     buckets,
    saleChannels:      saleChannels,
    provinces:         provinces,
    paymentTypes:      paymentTypes,
    birthdayAlerts:    birthdayAlerts
  };
}

// ── ALL POLICY ROWS — normalized from all 4 sheets ─────────
function getAllPolicyRows() {
  var rows = [];

  // VTC: AppDate(0), ClientID(1), Holder(2), Insured(3), DOB(4), Prov(5), Co(6), Pol#(7),
  //      Product(8), Start(9), End(10), PreEx(11), Payment(12), Premium(13), CommRate(14),
  //      AnnComm(15), MonComm(16), UpfComm(17), Months(18), Chargeback(19), TotalComm(20),
  //      Status(21), Claim(22), LeadSrc(23), RefBy(24), SaleChannel(25), Phone(26), Email(27), Notes(28)
  getSheetData(SHEETS.VTC).forEach(function(r) {
    if (!r[2]) return;
    rows.push({
      segment: 'vtc', date: r[0], clientId: r[1], clientName: r[2],
      insuredName: r[3], dob: r[4] ? fmtD(parseDate(r[4])) : '', province: norm(r[5]),
      company: r[6], policyNum: r[7],
      product: norm(r[8]), preExisting: r[11]||'No', paymentType: r[12]||'',
      premium: toNum(r[13]), commRate: toNum(r[14]),
      annComm: toNum(r[15]), monComm: toNum(r[16]), upfComm: toNum(r[17]),
      months: toNum(r[18]), chargeback: toNum(r[19]),
      comm: toNum(r[20]), status: r[21], claimStatus: normalizeClaimStatus(r[22]),
      leadSource: r[23], referredBy: r[24],
      saleChannel: (r[25]||'').toString().trim(),
      phone: r[26], email: r[27], notes: r[28],
      claimOpenDate:  r[29] ? fmtD(parseDate(r[29])) : '',
      claimCloseDate: r[30] ? fmtD(parseDate(r[30])) : '',
      startDate: r[9] ? fmtD(parseDate(r[9])) : '',
      endDate: r[10] ? fmtD(parseDate(r[10])) : ''
    });
  });

  // EDGE: Date(0), ClientID(1), Name(2), Co(3), Pol#(4), Product(5),
  //       MonPrem(6), MonComm(7), AnnComm(8), Status(9), Chargeback(10), NetAnnComm(11),
  //       LeadSrc(12), RefBy(13), Phone(14), Email(15), Notes(16)
  getSheetData(SHEETS.EDGE).forEach(function(r) {
    if (!r[2]) return;
    rows.push({
      segment: 'edge', date: r[0], clientId: r[1], clientName: r[2],
      insuredName: r[2], company: r[3], policyNum: r[4],
      product: norm(r[5]), premium: toNum(r[6]) * 12,
      comm: toNum(r[11]) || toNum(r[8]), status: r[9],
      leadSource: r[12], referredBy: r[13], phone: r[14], email: r[15], notes: r[16]
    });
  });

  // LIFE: Date(0), ClientID(1), Name(2), Co(3), Pol#(4), Product(5),
  //       Premium(6), CommRate(7), CommEarned(8), Chargeback(9), Status(10),
  //       NetComm(11), LeadSrc(12), RefBy(13), Phone(14), Email(15), Notes(16)
  getSheetData(SHEETS.LIFE).forEach(function(r) {
    if (!r[2]) return;
    rows.push({
      segment: 'life', date: r[0], clientId: r[1], clientName: r[2],
      insuredName: r[2], company: r[3], policyNum: r[4],
      product: norm(r[5]), premium: toNum(r[6]),
      comm: toNum(r[11]) || toNum(r[8]), status: r[10],
      leadSource: r[12], referredBy: r[13], phone: r[14], email: r[15], notes: r[16]
    });
  });

  // H&D: Date(0), ClientID(1), Holder(2), Co(3), Pol#(4), Product(5), PlanType(6),
  //      GroupSize(7), MonPrem(8), AnnPrem(9), CommRate(10), AnnComm(11),
  //      Status(12), LeadSrc(13), RefBy(14), Phone(15), Notes(16)
  getSheetData(SHEETS.HD).forEach(function(r) {
    if (!r[2]) return;
    rows.push({
      segment: 'hd', date: r[0], clientId: r[1], clientName: r[2],
      insuredName: r[2], company: r[3], policyNum: r[4],
      product: norm(r[5]), premium: toNum(r[9]),
      comm: toNum(r[11]), status: r[12],
      leadSource: r[13], referredBy: r[14], phone: r[15], notes: r[16]
    });
  });

  return rows;
}

// ── SEARCH ─────────────────────────────────────────────────
function searchClients(query) {
  if (!query || query.trim().length < 2) return [];
  var q = query.trim().toLowerCase();
  var results = [];

  function addIfMatch(rows, segment, nameCol, checkFn) {
    rows.forEach(function(r) {
      if (!r[nameCol]) return;
      if (checkFn(r)) {
        results.push(buildSearchResult(r, segment));
      }
    });
  }

  getSheetData(SHEETS.VTC).forEach(function(r) {
    if (!r[2]) return;
    if ((r[2]+'').toLowerCase().indexOf(q)!==-1 ||
        (r[3]+'').toLowerCase().indexOf(q)!==-1 ||
        (r[1]+'').toLowerCase()===q ||
        (r[26]+'').indexOf(q)!==-1) {
      results.push({
        segment:'Visitor & Super Visa', clientId:r[1], policyHolder:r[2], insuredName:r[3],
        dob:fmtD(parseDate(r[4])), province:r[5], company:r[6], policyNum:r[7], product:norm(r[8]),
        startDate:fmtD(parseDate(r[9])), endDate:fmtD(parseDate(r[10])), preExisting:r[11],
        paymentType:r[12], premium:toNum(r[13]), commRate:toNum(r[14]), totalComm:toNum(r[20]),
        status:r[21], leadSource:r[23], referredBy:r[24],
        saleChannel:(r[25]||'').toString().trim(), phone:r[26], email:r[27], notes:r[28],
        isReferral: detectReferral((r[25]||'').toString().trim())
      });
    }
  });

  getSheetData(SHEETS.EDGE).forEach(function(r) {
    if (!r[2]) return;
    if ((r[2]+'').toLowerCase().indexOf(q)!==-1 || (r[1]+'').toLowerCase()===q) {
      results.push({
        segment:'Edge Benefits', clientId:r[1], policyHolder:r[2], insuredName:r[2],
        company:r[3], policyNum:r[4], product:norm(r[5]),
        monthlyPremium:toNum(r[6]), annualComm:toNum(r[8]), netAnnualComm:toNum(r[11]),
        status:r[9], leadSource:r[12], referredBy:r[13], phone:r[14], email:r[15], notes:r[16]
      });
    }
  });

  getSheetData(SHEETS.LIFE).forEach(function(r) {
    if (!r[2]) return;
    if ((r[2]+'').toLowerCase().indexOf(q)!==-1 || (r[1]+'').toLowerCase()===q) {
      results.push({
        segment:'Life & CI', clientId:r[1], policyHolder:r[2], insuredName:r[2],
        company:r[3], policyNum:r[4], product:norm(r[5]),
        premium:toNum(r[6]), commRate:toNum(r[7]), netComm:toNum(r[11]),
        status:r[10], leadSource:r[12], referredBy:r[13], phone:r[14], email:r[15], notes:r[16]
      });
    }
  });

  getSheetData(SHEETS.HD).forEach(function(r) {
    if (!r[2]) return;
    if ((r[2]+'').toLowerCase().indexOf(q)!==-1 || (r[1]+'').toLowerCase()===q) {
      results.push({
        segment:'H&D / Group Benefits', clientId:r[1], policyHolder:r[2], insuredName:r[2],
        company:r[3], policyNum:r[4], product:norm(r[5]),
        planType:r[6], groupSize:r[7], annualPremium:toNum(r[9]), annualComm:toNum(r[11]),
        status:r[12], leadSource:r[13], referredBy:r[14], phone:r[15], notes:r[16]
      });
    }
  });

  return results;
}

function buildSearchResult(r, seg) { return { segment: seg, raw: r }; }

// ── RENEWALS ───────────────────────────────────────────────
function getRenewals(days) {
  var today  = new Date(); today.setHours(0,0,0,0);
  var out = [];
  // Use cached rows if available (avoids 2nd sheet read during combo)
  var src = _rowsCache
    ? _rowsCache.filter(function(r){return r.segment==='vtc';})
    : null;
  if (src) {
    src.forEach(function(r) {
      if (!r.endDate) return;
      var ed = parseDate(r.endDate);
      if (!ed) return;
      var dl = Math.floor((ed - today) / 86400000);
      if (dl <= days) {
        out.push({
          segment:'VTC / Super Visa', clientId:r.clientId, policyHolder:r.clientName,
          insuredName:r.insuredName, policyNum:r.policyNum, product:r.product,
          endDate:r.endDate, daysLeft:dl, premium:r.premium, totalComm:r.comm,
          status:r.status, phone:r.phone, email:r.email, notes:r.notes
        });
      }
    });
  } else {
    getSheetData(SHEETS.VTC).forEach(function(r) {
      if (!r[2]) return;
      var ed = parseDate(r[10]);
      if (!ed) return;
      var dl = Math.floor((ed - today) / 86400000);
      if (dl <= days) {
        out.push({
          segment:'VTC / Super Visa', clientId:r[1], policyHolder:r[2], insuredName:r[3],
          policyNum:r[7], product:norm(r[8]), endDate:fmtD(ed), daysLeft:dl,
          premium:toNum(r[13]), totalComm:toNum(r[20]), status:r[21],
          phone:r[26], email:r[27], notes:r[28]
        });
      }
    });
  }
  out.sort(function(a,b){return a.daysLeft-b.daysLeft;});
  return out;
}

// ── ALL POLICIES — for table view ─────────────────────────
function getAllPolicies() {
  var rows = getAllPolicyRowsCached();
  var today = new Date(); today.setHours(0,0,0,0);
  return rows.map(function(r) {
    var startD = r.startDate ? parseDate(r.startDate) : null;
    var endD   = r.endDate   ? parseDate(r.endDate)   : null;
    var daysToStart = startD ? Math.floor((startD - today) / 86400000) : null;
    var daysToEnd   = endD   ? Math.floor((endD   - today) / 86400000) : null;
    var isRef = detectReferral(r.saleChannel);
    return {
      segment:    r.segment,
      clientId:   r.clientId,
      clientName: r.clientName,
      policyNum:  r.policyNum,
      product:    r.product,
      company:    r.company,
      premium:    r.premium,
      comm:       r.comm,
      annComm:    r.annComm,
      monComm:    r.monComm,
      upfComm:    r.upfComm,
      months:     r.months,
      commRate:   r.commRate,
      chargeback:     r.chargeback,
      claimStatus:    r.claimStatus,
      claimOpenDate:  r.claimOpenDate  || '',
      claimCloseDate: r.claimCloseDate || '',
      status:     r.status,
      startDate:  r.startDate || '',
      endDate:    r.endDate   || '',
      daysToStart: daysToStart,
      daysToEnd:   daysToEnd,
      phone:      r.phone,
      email:      r.email,
      insuredName: r.insuredName,
      dob:        r.dob,
      province:   r.province || '',
      preExisting: r.preExisting,
      paymentType: r.paymentType,
      leadSource:  r.leadSource,
      referredBy:  r.referredBy,
      notes:       r.notes,
      saleChannel: r.saleChannel || '',
      isReferral:  isRef
    };
  }).sort(function(a,b){
    // sort by date desc (most recent first)
    if(a.startDate && b.startDate) return a.startDate < b.startDate ? 1 : -1;
    return 0;
  });
}

// ── UPCOMING STARTS — policies starting within N days ──────
function getUpcomingStarts(days) {
  var today  = new Date(); today.setHours(0,0,0,0);
  var out = [];
  var src = _rowsCache
    ? _rowsCache.filter(function(r){return r.segment==='vtc';})
    : null;
  if (src) {
    src.forEach(function(r) {
      if (!r.startDate) return;
      var sd = parseDate(r.startDate);
      if (!sd) return;
      var dl = Math.floor((sd - today) / 86400000);
      if (dl >= 0 && dl <= days) {
        out.push({
          segment:'VTC / Super Visa', clientId:r.clientId, policyHolder:r.clientName,
          insuredName:r.insuredName, policyNum:r.policyNum, product:r.product,
          startDate:r.startDate, daysToStart:dl, premium:r.premium, totalComm:r.comm,
          status:r.status, phone:r.phone, email:r.email
        });
      }
    });
  } else {
    getSheetData(SHEETS.VTC).forEach(function(r) {
      if (!r[2]) return;
      var sd = parseDate(r[9]);
      if (!sd) return;
      var dl = Math.floor((sd - today) / 86400000);
      if (dl >= 0 && dl <= days) {
        out.push({
          segment:'VTC / Super Visa', clientId:r[1], policyHolder:r[2], insuredName:r[3],
          policyNum:r[7], product:norm(r[8]), startDate:fmtD(sd), daysToStart:dl,
          premium:toNum(r[13]), totalComm:toNum(r[20]), status:r[21],
          phone:r[26], email:r[27]
        });
      }
    });
  }
  out.sort(function(a,b){return a.daysToStart-b.daysToStart;});
  return out;
}

// ── LEADS ──────────────────────────────────────────────────
function getLeads() {
  return getSheetData(SHEETS.LEADS).filter(function(r){return r[2];}).map(function(r){
    return {
      dateAdded:r[0]?fmtD(parseDate(r[0])):'', clientId:r[1], prospectName:r[2],
      phone:r[3], email:r[4], leadSource:r[5], referredBy:r[6], productInterest:r[7],
      estPremium:toNum(r[8]), estComm:toNum(r[9]), stage:r[10],
      lastContact:r[11]?fmtD(parseDate(r[11])):'',
      followUpDate:r[12]?fmtD(parseDate(r[12])):'',
      convertedPolicy:r[13], notes:r[14]
    };
  });
}

// ── CRM STATE — persist review/follow-up flags ────────────
// Requires a sheet named "CRM_State" with columns: Key | Value | UpdatedAt
function getCrmState() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName('CRM_State');
  if (!sh) return { error: 'CRM_State sheet not found. Create a sheet named "CRM_State" with columns: Key | Value | UpdatedAt' };
  var data = sh.getDataRange().getValues();
  var state = {};
  data.forEach(function(r) { if (r[0]) state[r[0]] = r[1]; });
  return { state: state };
}

function saveCrmState(p) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName('CRM_State');
  if (!sh) return { error: 'CRM_State sheet not found.' };
  var key = (p.key || '').toString().trim();
  var val = (p.value || '').toString().trim();
  if (!key) return { error: 'key required' };
  var tz = Session.getScriptTimeZone();
  var now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  var data = sh.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      sh.getRange(i + 1, 2).setValue(val);
      sh.getRange(i + 1, 3).setValue(now);
      return { ok: true };
    }
  }
  sh.appendRow([key, val, now]);
  return { ok: true };
}

// ── LEAD CONVERSION ────────────────────────────────────────
function getLeadConversion() {
  var leads = getLeads();
  var total = leads.length;
  var converted = leads.filter(function(l){ return l.convertedPolicy && l.convertedPolicy.toString().trim(); }).length;
  var byStage = {};
  leads.forEach(function(l){
    var s = l.stage || 'Unknown';
    if (!byStage[s]) byStage[s] = 0;
    byStage[s]++;
  });
  var stageArr = Object.keys(byStage).map(function(s){ return {stage:s, count:byStage[s]}; })
    .sort(function(a,b){return b.count-a.count;});
  return { total:total, converted:converted, conversionRate:total>0?round2(converted/total*100):0, byStage:stageArr };
}

// ── ADD POLICY ─────────────────────────────────────────────
function addPolicy(p) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var seg = p.segment || 'vtc';

  if (seg === 'vtc') {
    var payType   = (p.paymentType||'Annual').toString().trim().toLowerCase();
    var commRate  = toNum(p.commRate) || 0;
    var numMonths = parseInt(p.months) || 0;
    var annualPrem, monComm, upfComm, annComm, totalComm;

    if (payType === 'monthly') {
      // Monthly: user enters monthly premium + number of months
      var monPrem  = toNum(p.monthlyPremium) || toNum(p.premium);
      annualPrem   = round2(monPrem * 12);          // annualise for storage
      monComm      = round2(monPrem * commRate);     // per-month earned
      upfComm      = round2(2 * monComm);            // first + last upfront
      annComm      = 0;
      totalComm    = numMonths > 0 ? round2(numMonths * monComm) : upfComm;
    } else {
      // Annual: straightforward
      annualPrem   = toNum(p.premium);
      annComm      = round2(annualPrem * commRate);
      monComm      = 0;
      upfComm      = 0;
      totalComm    = annComm;
    }

    ss.getSheetByName(SHEETS.VTC).appendRow([
      today, p.clientId, p.policyHolder, p.insuredName, p.dob,
      p.province, p.company, p.policyNum, p.product,
      p.startDate, p.endDate, p.preExisting||'No', p.paymentType,
      annualPrem, commRate, annComm, monComm, upfComm, numMonths || '', 0, totalComm,
      p.status||'Pending', 'N', p.leadSource, p.referredBy, p.saleChannel||'',
      p.phone, p.email, '['+today+'] '+(p.notes||'')
    ]);
  } else if (seg === 'edge') {
    ss.getSheetByName(SHEETS.EDGE).appendRow([
      today, p.clientId, p.clientName, p.company, p.policyNum, p.product,
      toNum(p.monthlyPremium), '', '', p.status||'Pending',
      '', '', p.leadSource, p.referredBy, p.phone, p.email,
      '['+today+'] '+p.notes
    ]);
  } else if (seg === 'life') {
    ss.getSheetByName(SHEETS.LIFE).appendRow([
      today, p.clientId, p.applicantName, p.company, p.policyNum, p.product,
      toNum(p.premium), toNum(p.commRate)||0.55, '', '',
      p.status||'Pending', '', p.leadSource, p.referredBy, p.phone, p.email,
      '['+today+'] '+p.notes
    ]);
  } else if (seg === 'hd') {
    ss.getSheetByName(SHEETS.HD).appendRow([
      today, p.clientId, p.policyHolder, p.company, p.policyNum, p.product,
      p.planType, toNum(p.groupSize)||1,
      toNum(p.monthlyPremium), toNum(p.annualPremium),
      toNum(p.commRate)||0.10, '', p.status||'Pending',
      p.leadSource, p.referredBy, p.phone,
      '['+today+'] '+p.notes
    ]);
  }
  return { success: true };
}

// ── POLICY CONFIG (for update/note functions) ──────────────
var POLICY_CFG = {
  vtc: {
    sheet: 'Visitor & Super Visa', polNumCol: 7, notesCol: 28,
    fields: {
      startDate:   [9,  'Start Date'],
      endDate:     [10, 'End Date'],
      preExisting: [11, 'Pre-Existing Condition'],
      paymentType: [12, 'Payment Type'],
      premium:     [13, 'Annual Premium'],
      commRate:    [14, 'Commission Rate'],
      status:         [21, 'Status'],
      claimStatus:    [22, 'Claim Status'],
      claimOpenDate:  [29, 'Claim Open Date'],
      claimCloseDate: [30, 'Claim Close Date'],
      chargeback:     [19, 'Chargeback'],
      phone:       [26, 'Phone'],
      email:       [27, 'Email'],
      notes:       [28, 'Notes']
    }
  },
  edge: {
    sheet: 'Edge Benefits', polNumCol: 4, notesCol: 16,
    fields: {
      product:        [5,  'Product'],
      monthlyPremium: [6,  'Monthly Premium'],
      status:         [9,  'Status'],
      phone:          [14, 'Phone'],
      email:          [15, 'Email']
    }
  },
  life: {
    sheet: 'Life & CI Insurance', polNumCol: 4, notesCol: 16,
    fields: {
      product:  [5,  'Product'],
      premium:  [6,  'Premium'],
      commRate: [7,  'Commission Rate'],
      status:   [10, 'Status'],
      phone:    [14, 'Phone'],
      email:    [15, 'Email']
    }
  },
  hd: {
    sheet: 'H&D Group Benefits', polNumCol: 4, notesCol: 16,
    fields: {
      product:        [5,  'Product'],
      planType:       [6,  'Plan Type'],
      groupSize:      [7,  'Group Size'],
      monthlyPremium: [8,  'Monthly Premium'],
      annualPremium:  [9,  'Annual Premium'],
      commRate:       [10, 'Commission Rate'],
      status:         [12, 'Status'],
      phone:          [15, 'Phone']
    }
  }
};

// ── UPDATE POLICY ──────────────────────────────────────────
function updatePolicy(p) {
  var seg = (p.segment || 'vtc').toLowerCase();
  var polNum = (p.policyNum || '').toString().trim();
  if (!polNum) return { error: 'No policy number provided' };
  var cfg = POLICY_CFG[seg];
  if (!cfg) return { error: 'Unknown segment: ' + seg };

  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(cfg.sheet);
  if (!sheet) return { error: 'Sheet not found: ' + cfg.sheet };

  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;
  for (var i = DATA_ROW - 1; i < data.length; i++) {
    if ((data[i][cfg.polNumCol] + '').trim() === polNum) { rowIdx = i; break; }
  }
  if (rowIdx === -1) return { error: 'Policy ' + polNum + ' not found in sheet' };

  var tz      = Session.getScriptTimeZone();
  var now     = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  var changes = [];

  Object.keys(cfg.fields).forEach(function(field) {
    if (p[field] === undefined || p[field] === null) return;
    var col   = cfg.fields[field][0];
    var label = cfg.fields[field][1];
    var rawOld = data[rowIdx][col];
    // Format Google Sheets Date objects as YYYY-MM-DD for clean display
    var oldVal = (rawOld instanceof Date)
      ? Utilities.formatDate(rawOld, tz, 'yyyy-MM-dd')
      : (rawOld + '').trim();
    var writeVal = p[field];

    // Date fields: empty string → write blank (not epoch); skip if effectively blank→blank
    var dateFields = ['startDate','endDate','claimOpenDate','claimCloseDate'];
    if (dateFields.indexOf(field) !== -1) {
      if (!writeVal || writeVal === '') {
        // Only write blank if something was there before; skip silently if already blank
        if (!oldVal || oldVal === '') return;
        writeVal = '';
      }
      // Old value might be a raw number (sheet stores as serial) — skip those silently
      if (oldVal && !isNaN(oldVal) && oldVal.indexOf('-') === -1) oldVal = '';
    }

    // notes: allow empty string to clear notes
    if (field === 'notes' && p[field] === '') { writeVal = ''; }
    else if (p[field] === '') return;
    var newVal = (writeVal + '').trim();
    // Treat blank and 0 as equivalent for numeric fields to avoid spurious changelog entries
    var numericFields = ['premium','commRate','chargeback','monthlyPremium','annualPremium'];
    if (numericFields.indexOf(field) !== -1) {
      if ((oldVal === '' || oldVal === '0' || parseFloat(oldVal) === 0) &&
          (newVal === '' || newVal === '0' || parseFloat(newVal) === 0)) return;
    }
    if (oldVal !== newVal) {
      sheet.getRange(rowIdx + 1, col + 1).setValue(writeVal);
      // Don't log notes changes in changelog — notes are managed directly
      if (field !== 'notes') {
        // Compact label for claim date fields
        var logLabel = field === 'claimOpenDate'  ? 'Claim opened'
                     : field === 'claimCloseDate' ? 'Claim closed'
                     : field === 'claimStatus'    ? 'Claim'
                     : label;
        var logOld = oldVal || 'blank';
        var logNew = newVal || 'blank';
        changes.push(logLabel + ': ' + logOld + ' → ' + logNew);
      }
    }
  });

  // Build note entry
  var parts = [];
  if (changes.length > 0) {
    parts.push('[' + now + '] Policy updated — ' + changes.join('; '));
  }
  if (p.note && p.note.toString().trim()) {
    parts.push('[' + now + '] ' + p.note.toString().trim());
  }
  if (parts.length > 0) {
    var existing = (data[rowIdx][cfg.notesCol] + '').trim();
    var updated  = parts.join('\n') + (existing ? '\n' + existing : '');
    sheet.getRange(rowIdx + 1, cfg.notesCol + 1).setValue(updated);
  }

  return { success: true, changes: changes, changeCount: changes.length };
}

// ── ADD NOTE ───────────────────────────────────────────────
function addNote(p) {
  var seg    = (p.segment || 'vtc').toLowerCase();
  var polNum = (p.policyNum || '').toString().trim();
  var note   = (p.note || '').toString().trim();
  if (!polNum || !note) return { error: 'Missing policyNum or note text' };

  var cfg = POLICY_CFG[seg];
  if (!cfg) return { error: 'Unknown segment' };

  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(cfg.sheet);
  if (!sheet) return { error: 'Sheet not found' };

  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;
  for (var i = DATA_ROW - 1; i < data.length; i++) {
    if ((data[i][cfg.polNumCol] + '').trim() === polNum) { rowIdx = i; break; }
  }
  if (rowIdx === -1) return { error: 'Policy not found' };

  // Use custom date if provided (e.g. backdating a note), otherwise now
  var tz = Session.getScriptTimeZone();
  var ts = (p.noteDate || '').toString().trim();
  if (!ts) ts = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  var existing = (data[rowIdx][cfg.notesCol] + '').trim();
  var newNotes = '[' + ts + '] ' + note + (existing ? '\n' + existing : '');
  sheet.getRange(rowIdx + 1, cfg.notesCol + 1).setValue(newNotes);
  return { success: true };
}

// Directly overwrite the notes cell — no changelog, no side effects.
// Used by delete/edit note operations.
function saveNotes(p) {
  var seg    = (p.segment || 'vtc').toLowerCase();
  var polNum = (p.policyNum || '').toString().trim();
  var cfg    = POLICY_CFG[seg];
  if (!cfg) return { error: 'Unknown segment' };
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(cfg.sheet);
  if (!sheet) return { error: 'Sheet not found' };
  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;
  for (var i = DATA_ROW - 1; i < data.length; i++) {
    if ((data[i][cfg.polNumCol] + '').trim() === polNum) { rowIdx = i; break; }
  }
  if (rowIdx === -1) return { error: 'Policy not found' };
  sheet.getRange(rowIdx + 1, cfg.notesCol + 1).setValue((p.notes || '').toString().trim());
  return { success: true };
}

// ── GET SINGLE POLICY ──────────────────────────────────────
function getPolicy(segment, policyNum) {
  var seg    = (segment || 'vtc').toLowerCase();
  var polNum = (policyNum || '').toString().trim();
  var cfg    = POLICY_CFG[seg];
  if (!cfg) return { error: 'Unknown segment' };

  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(cfg.sheet);
  if (!sheet) return { error: 'Sheet not found' };

  var data = sheet.getDataRange().getValues();
  for (var i = DATA_ROW - 1; i < data.length; i++) {
    if ((data[i][cfg.polNumCol] + '').trim() === polNum) {
      var row = data[i];
      var obj = { segment: seg, policyNum: polNum };
      Object.keys(cfg.fields).forEach(function(field) {
        var col = cfg.fields[field][0];
        obj[field] = (row[col] + '').trim();
      });
      obj.notes = (row[cfg.notesCol] + '').trim();
      return obj;
    }
  }
  return { error: 'Not found' };
}

// ── ADD LEAD ───────────────────────────────────────────────
function addLead(p) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  ss.getSheetByName(SHEETS.LEADS).appendRow([
    today, p.clientId, p.prospectName, p.phone, p.email,
    p.leadSource, p.referredBy, p.productInterest,
    toNum(p.estPremium), toNum(p.estComm),
    p.stage||'New Lead', today, p.followUpDate, '',
    '['+today+'] '+p.notes
  ]);
  return { success: true };
}

// ── HELPERS ────────────────────────────────────────────────
function getSheetData(name) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var ws = ss.getSheetByName(name);
  if (!ws) return [];
  var last = ws.getLastRow();
  if (last < DATA_ROW) return [];
  return ws.getRange(DATA_ROW, 1, last - DATA_ROW + 1, ws.getLastColumn()).getValues();
}

// Rattan/IBR detection: check dedicated Sale Channel column (col 28).
// Lead Source is untouched — it stays as "Personal", "Referral", etc.
function normalizeClaimStatus(v) {
  v = (v + '').trim();
  if (v.toUpperCase() === 'Y') return 'Open';   // backward compat
  if (v.toUpperCase() === 'N' || v === '') return 'None';
  if (['Open', 'Closed', 'Denied'].indexOf(v) !== -1) return v;
  return 'None';
}

function detectReferral(saleChannel) {
  var sc = (saleChannel || '').toLowerCase().trim().replace(/\s+/g, '');
  // Match "Rattan/IBR", "rattan / ibr", "Rattan/ibr", etc.
  return sc.indexOf('rattan') !== -1 || sc === 'ibr';
}

function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  // Parse yyyy-MM-dd strings as local time (not UTC) to avoid timezone offset bugs
  var m = (v + '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  var d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function fmtD(d) {
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function toNum(v) {
  var n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}

function round2(n) { return Math.round(n * 100) / 100; }

function norm(v) {
  var s = (v || '').toString().trim();
  return (!s || s === '-') ? 'Unknown' : s;
}

function getDayOfYear(d) {
  var start = new Date(d.getFullYear(), 0, 0);
  return Math.floor((d - start) / 86400000);
}
