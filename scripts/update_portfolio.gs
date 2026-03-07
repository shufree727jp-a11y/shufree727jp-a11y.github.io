/**
 * 台股投資紀錄每日更新（Google Apps Script）
 * - 讀取 Sheet: Trades
 * - 取得台股價格（TWSE MIS）
 * - 計算持倉/損益
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

  // 支援英文與中文欄位名稱
  const col = {
    date: pickCol_(idx, ['Date', '日期']),
    symbol: pickCol_(idx, ['Symbol', '股票代號', '代號']),
    market: pickCol_(idx, ['Market', '市場']),
    side: pickCol_(idx, ['Side', '買賣別', '買賣']),
    shares: pickCol_(idx, ['Shares', '股數']),
    price: pickCol_(idx, ['Price', '成交價', '價格']),
    fee: pickCol_(idx, ['Fee', '手續費']),
    tax: pickCol_(idx, ['Tax', '交易稅', '稅']),
    note: pickCol_(idx, ['Note', '備註'])
  };

  if (col.symbol < 0) {
    throw new Error('找不到股票代號欄位（Symbol/股票代號/代號）');
  }

  return values.slice(1)
    .filter(r => String(r[col.symbol] || '').trim())
    .map(r => ({
      date: col.date >= 0 ? formatDate_(r[col.date]) : '',
      symbol: String(r[col.symbol] || '').trim(),
      market: col.market >= 0 ? String(r[col.market] || 'TW').trim() : 'TW',
      side: normalizeSide_(col.side >= 0 ? String(r[col.side] || '').trim() : ''),
      shares: col.shares >= 0 ? Number(r[col.shares] || 0) : 0,
      price: col.price >= 0 ? Number(r[col.price] || 0) : 0,
      fee: col.fee >= 0 ? Number(r[col.fee] || 0) : 0,
      tax: col.tax >= 0 ? Number(r[col.tax] || 0) : 0,
      note: col.note >= 0 ? String(r[col.note] || '').trim() : ''
    }))
    .sort((a, b) => (a.date > b.date ? 1 : -1));
}

function fetchTWPrices_(symbols) {
  const out = {};
  symbols.forEach(symbol => {
    // 優先上市，再嘗試上櫃
    const urls = [
      `https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch=tse_${symbol}.tw&json=1&delay=0`,
      `https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch=otc_${symbol}.tw&json=1&delay=0`
    ];

    let price = 0;
    let name = symbol;

    for (const url of urls) {
      try {
        const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        if (res.getResponseCode() !== 200) continue;

        const data = JSON.parse(res.getContentText());
        const arr = data.msgArray || [];
        if (!arr.length) continue;

        const row = arr[0];
        name = row.n || symbol;
        // z: 當盤成交價，若無則用 y(昨收)
        const z = Number(row.z || 0);
        const y = Number(row.y || 0);
        price = z > 0 ? z : y;

        if (price > 0) break;
      } catch (e) {
        // ignore and try next
      }
    }

    out[symbol] = { lastPrice: price, name };
  });

  return out;
}

function buildPortfolio_(trades, priceMap) {
  const positions = {}; // symbol -> {shares, totalCost}

  trades.forEach(t => {
    if (!positions[t.symbol]) positions[t.symbol] = { shares: 0, totalCost: 0 };
    const p = positions[t.symbol];

    if (t.side === 'BUY') {
      p.totalCost += t.shares * t.price + t.fee + t.tax;
      p.shares += t.shares;
    } else if (t.side === 'SELL') {
      const avg = p.shares > 0 ? p.totalCost / p.shares : 0;
      p.totalCost -= avg * t.shares;
      p.totalCost = Math.max(0, p.totalCost);
      p.shares -= t.shares;
      p.shares = Math.max(0, p.shares);
    }
  });

  const posArr = Object.entries(positions)
    .filter(([, p]) => p.shares > 0)
    .map(([symbol, p]) => {
      const quote = priceMap[symbol] || { lastPrice: 0, name: symbol };
      const avgCost = p.shares > 0 ? p.totalCost / p.shares : 0;
      const marketValue = p.shares * quote.lastPrice;
      const unrealizedPnl = marketValue - p.totalCost;
      const unrealizedPnlPct = p.totalCost > 0 ? unrealizedPnl / p.totalCost : 0;
      return {
        symbol,
        name: quote.name,
        shares: p.shares,
        avgCost: round2_(avgCost),
        lastPrice: round2_(quote.lastPrice),
        marketValue: round2_(marketValue),
        unrealizedPnl: round2_(unrealizedPnl),
        unrealizedPnlPct: round4_(unrealizedPnlPct)
      };
    });

  const totalCost = posArr.reduce((s, x) => s + x.avgCost * x.shares, 0);
  const marketValue = posArr.reduce((s, x) => s + x.marketValue, 0);
  const unrealizedPnl = marketValue - totalCost;

  return {
    updatedAt: Utilities.formatDate(new Date(), 'Asia/Taipei', "yyyy-MM-dd'T'HH:mm:ssXXX"),
    market: 'TW',
    currency: 'TWD',
    summary: {
      totalCost: round2_(totalCost),
      marketValue: round2_(marketValue),
      unrealizedPnl: round2_(unrealizedPnl),
      unrealizedPnlPct: totalCost > 0 ? round4_(unrealizedPnl / totalCost) : 0
    },
    positions: posArr,
    trades: trades
  };
}

function pushJsonToGitHub_(cfg, path, obj) {
  if (!cfg.token || !cfg.owner || !cfg.repo) {
    throw new Error('請先設定 Script Properties: GITHUB_TOKEN / GITHUB_OWNER / GITHUB_REPO');
  }

  const api = `https://api.github.com/repos/${cfg.owner}/${cfg.repo}/contents/${path}`;

  // 先取 sha（若檔案已存在）
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
    const data = JSON.parse(getRes.getContentText());
    sha = data.sha;
  }

  const content = Utilities.base64Encode(
    JSON.stringify(obj, null, 2),
    Utilities.Charset.UTF_8
  );

  const payload = {
    message: `chore: update portfolio data (${new Date().toISOString()})`,
    content,
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
  return String(v).slice(0, 10);
}

function normalizeSide_(v) {
  const s = String(v || '').trim().toUpperCase();
  if (s === '買' || s === '買進' || s === 'BUY' || s === 'B') return 'BUY';
  if (s === '賣' || s === '賣出' || s === 'SELL' || s === 'S') return 'SELL';
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
