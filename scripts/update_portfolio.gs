/**
 * 台股投資紀錄每日更新（Google Apps Script）
 * - 讀取 Sheet: Trades
 * - 取得台股價格（TWSE MIS）
 * - 計算持倉/未實現/已實現/股利/總報酬
 * - 更新 GitHub: data/portfolio.json
 */

function dailyUpdateTWPortfolio() {
  const cfg = getConfig_();
  const trades = readTrades_('Trades');

  const symbols = [...new Set(trades.map(t => t.symbol))].filter(Boolean);
  const priceMap = fetchTWPrices_(symbols);

  const result = buildPortfolio_(trades, priceMap);
  pushJsonToGitHub_(cfg, 'data/portfolio.json', result);
}

function getConfig_() {
  const p = PropertiesService.getScriptProperties();
  return {
    token: p.getProperty('GITHUB_TOKEN'),
    owner: p.getProperty('GITHUB_OWNER'),
    repo: p.getProperty('GITHUB_REPO'),
    branch: p.getProperty('GITHUB_BRANCH') || 'main'
  };
}

function readTrades_(sheetName) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh) throw new Error(`找不到工作表：${sheetName}`);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => String(h).trim());
  const idx = Object.fromEntries(headers.map((h, i) => [h, i]));

  const col = {
    date: pickCol_(idx, ['Date', '日期']),
    symbol: pickCol_(idx, ['Symbol', '股票代號', '代號']),
    market: pickCol_(idx, ['Market', '市場']),
    side: pickCol_(idx, ['Side', '買賣別', '買賣', '類型']),
    shares: pickCol_(idx, ['Shares', '股數']),
    price: pickCol_(idx, ['Price', '成交價', '價格', '每股金額']),
    amount: pickCol_(idx, ['Amount', '金額', '總金額']),
    feeTax: pickCol_(idx, ['FeeTax', '費用', '手續費+交易稅', '手續費及交易稅']),
    fee: pickCol_(idx, ['Fee', '手續費']),
    tax: pickCol_(idx, ['Tax', '交易稅', '稅']),
    note: pickCol_(idx, ['Note', '備註'])
  };

  if (col.symbol < 0) throw new Error('找不到股票代號欄位（Symbol/股票代號/代號）');
  if (col.side < 0) throw new Error('找不到買賣別欄位（Side/買賣別/買賣/類型）');

  return values.slice(1)
    .filter(r => String(r[col.symbol] || '').trim())
    .map(r => {
      const side = normalizeSide_(col.side >= 0 ? String(r[col.side] || '').trim() : '');
      const shares = col.shares >= 0 ? Number(r[col.shares] || 0) : 0;
      const price = col.price >= 0 ? Number(r[col.price] || 0) : 0;
      const amount = col.amount >= 0 ? Number(r[col.amount] || 0) : 0;
      const feeTax = col.feeTax >= 0 ? Number(r[col.feeTax] || 0) : ((col.fee >= 0 ? Number(r[col.fee] || 0) : 0) + (col.tax >= 0 ? Number(r[col.tax] || 0) : 0));
      return {
        date: col.date >= 0 ? formatDate_(r[col.date]) : '',
        symbol: String(r[col.symbol] || '').trim(),
        market: col.market >= 0 ? String(r[col.market] || 'TW').trim() : 'TW',
        side,
        shares,
        price,
        amount,
        feeTax,
        note: col.note >= 0 ? String(r[col.note] || '').trim() : ''
      };
    })
    .sort((a, b) => (a.date > b.date ? 1 : -1));
}

function fetchTWPrices_(symbols) {
  const out = {};
  symbols.forEach(symbol => {
    let quote = fetchFromTWSE_(symbol);

    // 1) 先用 TWSE MIS；2) 抓不到再用 Yahoo Finance；3) 再抓不到標示無即時
    if (!(quote && quote.lastPrice > 0)) {
      const yahooQuote = fetchFromYahoo_(symbol);
      if (yahooQuote && yahooQuote.lastPrice > 0) {
        quote = yahooQuote;
      }
    }

    if (!(quote && quote.lastPrice > 0)) {
      quote = {
        lastPrice: 0,
        name: symbol,
        priceSource: '無即時',
        lastPriceAt: ''
      };
    }

    out[symbol] = quote;
  });
  return out;
}

function fetchFromTWSE_(symbol) {
  const urls = [
    `https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch=tse_${symbol}.tw&json=1&delay=0`,
    `https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch=otc_${symbol}.tw&json=1&delay=0`
  ];

  for (const url of urls) {
    try {
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
      if (res.getResponseCode() !== 200) continue;
      const data = JSON.parse(res.getContentText());
      const arr = data.msgArray || [];
      if (!arr.length) continue;

      const row = arr[0];
      const name = row.n || symbol;
      const z = Number(row.z || 0);
      const y = Number(row.y || 0);
      const pz = Number(row.pz || 0); // 最近揭示成交價（有些時段 z 會是 '-'）
      let lastPrice = 0;
      let priceSource = '無即時';

      if (z > 0) {
        lastPrice = z;
        priceSource = 'TWSE當盤';
      } else if (pz > 0) {
        lastPrice = pz;
        priceSource = 'TWSE最近成交';
      } else if (y > 0) {
        lastPrice = y;
        priceSource = 'TWSE昨收';
      }

      let lastPriceAt = '';
      if (row.tlong) {
        try {
          lastPriceAt = Utilities.formatDate(new Date(Number(row.tlong)), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
        } catch (e) {}
      }

      if (lastPrice > 0) {
        return { lastPrice, name, priceSource, lastPriceAt };
      }
    } catch (e) {}
  }

  return null;
}

function fetchFromYahoo_(symbol) {
  // Yahoo 台股：上市用 .TW，上櫃用 .TWO
  const ySymbols = [`${symbol}.TW`, `${symbol}.TWO`];

  for (const ys of ySymbols) {
    const url = `https://query1.finance.yahoo.com/v7/finance/quote?symbols=${encodeURIComponent(ys)}`;
    try {
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
      if (res.getResponseCode() !== 200) continue;

      const data = JSON.parse(res.getContentText());
      const result = (((data || {}).quoteResponse || {}).result || [])[0];
      if (!result) continue;

      const lastPrice = Number(result.regularMarketPrice || 0);
      if (!(lastPrice > 0)) continue;

      const name = result.shortName || result.longName || symbol;
      const ts = Number(result.regularMarketTime || 0);
      const lastPriceAt = ts > 0
        ? Utilities.formatDate(new Date(ts * 1000), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss')
        : '';

      return {
        lastPrice,
        name,
        priceSource: 'Yahoo',
        lastPriceAt
      };
    } catch (e) {}
  }

  return null;
}


function buildPortfolio_(trades, priceMap) {
  const positions = {}; // symbol -> {shares, totalCost, realizedPnl, dividendCash, dividendStockShares}

  trades.forEach(t => {
    if (!positions[t.symbol]) {
      positions[t.symbol] = { shares: 0, totalCost: 0, realizedPnl: 0, dividendCash: 0, dividendStockShares: 0 };
    }
    const p = positions[t.symbol];

    if (t.side === 'BUY') {
      p.totalCost += t.shares * t.price + t.feeTax;
      p.shares += t.shares;
    } else if (t.side === 'SELL') {
      const avg = p.shares > 0 ? p.totalCost / p.shares : 0;
      const proceeds = t.shares * t.price - t.feeTax;
      const costOut = avg * t.shares;
      p.realizedPnl += proceeds - costOut;
      p.totalCost -= costOut;
      p.totalCost = Math.max(0, p.totalCost);
      p.shares -= t.shares;
      p.shares = Math.max(0, p.shares);
    } else if (t.side === 'DIVIDEND_CASH') {
      const cash = t.amount > 0 ? t.amount : (t.shares * t.price);
      p.dividendCash += Math.max(0, cash - t.feeTax);
    } else if (t.side === 'DIVIDEND_STOCK') {
      const bonusShares = t.shares > 0 ? t.shares : (t.amount > 0 ? t.amount : 0);
      p.dividendStockShares += bonusShares;
      p.shares += bonusShares;
      // 股票股利成本視為0，不增加 totalCost
    }
  });

  const posArr = Object.entries(positions)
    .filter(([, p]) => p.shares > 0 || p.realizedPnl !== 0 || p.dividendCash !== 0)
    .map(([symbol, p]) => {
      const quote = priceMap[symbol] || { lastPrice: 0, name: symbol };
      const avgCost = p.shares > 0 ? p.totalCost / p.shares : 0;
      const marketValue = p.shares * quote.lastPrice;
      const unrealizedPnl = marketValue - p.totalCost;
      const totalPnl = unrealizedPnl + p.realizedPnl + p.dividendCash;
      const totalPnlPct = p.totalCost > 0 ? totalPnl / p.totalCost : 0;
      return {
        symbol,
        name: quote.name,
        shares: p.shares,
        avgCost: round2_(avgCost),
        lastPrice: round2_(quote.lastPrice),
        priceSource: quote.priceSource || 'NO_DATA',
        lastPriceAt: quote.lastPriceAt || '',
        marketValue: round2_(marketValue),
        unrealizedPnl: round2_(unrealizedPnl),
        realizedPnl: round2_(p.realizedPnl),
        dividendCash: round2_(p.dividendCash),
        dividendStockShares: round4_(p.dividendStockShares),
        totalPnl: round2_(totalPnl),
        totalPnlPct: round4_(totalPnlPct)
      };
    });

  const totalCost = posArr.reduce((s, x) => s + x.avgCost * x.shares, 0);
  const marketValue = posArr.reduce((s, x) => s + x.marketValue, 0);
  const unrealizedPnl = posArr.reduce((s, x) => s + x.unrealizedPnl, 0);
  const realizedPnl = posArr.reduce((s, x) => s + x.realizedPnl, 0);
  const dividendCash = posArr.reduce((s, x) => s + x.dividendCash, 0);
  const totalPnl = unrealizedPnl + realizedPnl + dividendCash;

  return {
    updatedAt: Utilities.formatDate(new Date(), 'Asia/Taipei', "yyyy-MM-dd'T'HH:mm:ssXXX"),
    market: 'TW',
    currency: 'TWD',
    summary: {
      totalCost: round2_(totalCost),
      marketValue: round2_(marketValue),
      unrealizedPnl: round2_(unrealizedPnl),
      realizedPnl: round2_(realizedPnl),
      dividendCash: round2_(dividendCash),
      totalPnl: round2_(totalPnl),
      totalPnlPct: totalCost > 0 ? round4_(totalPnl / totalCost) : 0
    },
    positions: posArr,
    trades
  };
}

function pushJsonToGitHub_(cfg, path, obj) {
  if (!cfg.token || !cfg.owner || !cfg.repo) {
    throw new Error('請先設定 Script Properties: GITHUB_TOKEN / GITHUB_OWNER / GITHUB_REPO');
  }

  const api = `https://api.github.com/repos/${cfg.owner}/${cfg.repo}/contents/${path}`;

  let sha = null;
  const getRes = UrlFetchApp.fetch(`${api}?ref=${cfg.branch}`, {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      Authorization: `token ${cfg.token}`,
      Accept: 'application/vnd.github+json'
    }
  });
  if (getRes.getResponseCode() === 200) {
    sha = JSON.parse(getRes.getContentText()).sha;
  }

  const payload = {
    message: `chore: update portfolio data (${new Date().toISOString()})`,
    content: Utilities.base64Encode(JSON.stringify(obj, null, 2), Utilities.Charset.UTF_8),
    branch: cfg.branch
  };
  if (sha) payload.sha = sha;

  const putRes = UrlFetchApp.fetch(api, {
    method: 'put',
    muteHttpExceptions: true,
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    headers: {
      Authorization: `token ${cfg.token}`,
      Accept: 'application/vnd.github+json'
    }
  });

  if (putRes.getResponseCode() < 200 || putRes.getResponseCode() >= 300) {
    throw new Error(`GitHub 更新失敗: ${putRes.getResponseCode()} ${putRes.getContentText()}`);
  }
}

function formatDate_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd');
  }

  const s = String(v).trim();
  const m = s.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
  if (m) {
    const yyyy = m[1];
    const mm = String(Number(m[2])).padStart(2, '0');
    const dd = String(Number(m[3])).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  }
  return s.slice(0, 10);
}

function normalizeSide_(v) {
  const s = String(v || '').trim().toUpperCase();
  if (['買', '買進', 'BUY', 'B'].includes(s)) return 'BUY';
  if (['賣', '賣出', 'SELL', 'S'].includes(s)) return 'SELL';
  if (['現金股利', '股利', '股息', 'DIVIDEND', 'DIVIDEND_CASH', 'CASH_DIVIDEND'].includes(s)) return 'DIVIDEND_CASH';
  if (['股票股利', '配股', 'DIVIDEND_STOCK', 'STOCK_DIVIDEND'].includes(s)) return 'DIVIDEND_STOCK';
  return s;
}

function pickCol_(idx, names) {
  for (const n of names) {
    if (Object.prototype.hasOwnProperty.call(idx, n)) return idx[n];
  }
  return -1;
}

function round2_(n) { return Math.round((Number(n) || 0) * 100) / 100; }
function round4_(n) { return Math.round((Number(n) || 0) * 10000) / 10000; }
