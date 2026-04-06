// ================================================================
// 실시간 코인 트래커 Pro — 업비트
// ================================================================

const APP_NAME = '실시간 코인 트래커 Pro';
const UPBIT_API = 'https://api.upbit.com/v1/ticker';
const BITHUMB_API = 'https://api.bithumb.com/public/ticker';
const EXCHANGE_API = 'https://api.exchangerate-api.com/v4/latest/USD';

const SHEET_PORTFOLIO  = '📊 포트폴리오';
const SHEET_HISTORY    = '📋 거래내역';
const SHEET_ALERT      = '⚙️ 알림설정';
const SHEET_DASHBOARD  = '📈 대시보드';

// ================================================================
// 1. 메뉴 생성
// ================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 ' + APP_NAME)
    .addItem('📊 지금 시세 업데이트', 'updatePrices')
    .addItem('⏰ 자동 업데이트 시작 (1분)', 'setupTrigger')
    .addItem('⏹️ 자동 업데이트 중지', 'removeTrigger')
    .addSeparator()
    .addItem('📋 거래 입력 (매수)', 'openBuyDialog')
    .addItem('📋 거래 입력 (매도)', 'openSellDialog')
    .addSeparator()
    .addItem('💱 환율 업데이트', 'updateExchangeRate')
    .addItem('📈 대시보드 새로고침', 'refreshDashboard')
    .addItem('💰 수익 시뮬레이터', 'openSimulator')
    .addSeparator()
    .addItem('🔧 시트 초기 설정', 'initializeSheets')
    .addItem('🔍 API 연결 테스트', 'testAPI')
    .addItem('❓ 사용 가이드', 'showGuide')
    .addToUi();
}

// ================================================================
// 2. 시트 초기화
// ================================================================
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  setupPortfolioSheet(ss);
  setupHistorySheet(ss);
  setupAlertSheet(ss);
  setupDashboardSheet(ss);
  
  SpreadsheetApp.getUi().alert(
    '✅ 초기 설정 완료!\n\n' +
    '1. 📊 포트폴리오 시트에 코인 정보를 입력하세요\n' +
    '2. ⚙️ 알림설정 시트에 이메일과 목표가를 입력하세요\n' +
    '3. 메뉴 → 자동 업데이트 시작을 클릭하세요'
  );
}

function setupPortfolioSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_PORTFOLIO);
  if (!sheet) sheet = ss.insertSheet(SHEET_PORTFOLIO, 0);
  sheet.clear();

  // 상단 타이틀
  sheet.getRange('A1:L1').merge()
    .setValue('🚀 실시간 코인 트래커 Pro — 업비트')
    .setBackground('#0d0f1a')
    .setFontColor('#F0B90B')
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 50);

  // 요약 행
  const summaryLabels = [['총 평가금 (KRW)', '', '총 투자금 (KRW)', '', '총 수익금', '', '수익률', '', '현금 비율', '', '마지막 업데이트', '']];
  sheet.getRange('A2:L2').setValues(summaryLabels)
    .setBackground('#151929')
    .setFontColor('#8b93bb')
    .setFontSize(9)
    .setFontWeight('bold');

  sheet.getRange('A3:L3').setBackground('#151929').setFontSize(13).setFontWeight('bold');
  sheet.setRowHeight(3, 40);

  // 헤더
  const headers = [
    '거래소', '코인명', '심볼',
    '현재가 (KRW)', '전일대비 (%)', '현재가 (USD)',
    '보유수량', '매수평균가', '평가금 (KRW)',
    '수익금 (KRW)', '수익률 (%)', '메모'
  ];
  const headerRange = sheet.getRange(4, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setBackground('#1e2236')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontSize(11);
  sheet.setRowHeight(4, 35);

  // 예시 데이터
  const sampleData = [
    ['업비트', '비트코인', 'BTC', '', '', '', 0.05, 130000000, '', '', '', '장기 보유'],
    ['업비트', '이더리움', 'ETH', '', '', '', 1.2, 5000000, '', '', '', ''],
    ['업비트', '리플', 'XRP', '', '', '', 5000, 700, '', '', '', ''],
    ['빗썸', '도지코인', 'DOGE', '', '', '', 27000, 200, '', '', '', ''],
  ];
  sheet.getRange(5, 1, sampleData.length, headers.length).setValues(sampleData);

  // 열 너비
  const colWidths = [80, 100, 70, 130, 100, 120, 100, 130, 130, 120, 90, 150];
  colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // 입력칸 노란 하이라이트
  sheet.getRange('A5:C100').setBackground('#fffde7');
  sheet.getRange('G5:H100').setBackground('#fffde7');
  sheet.getRange('L5:L100').setBackground('#fffde7');

  // 자동계산칸 잠금 표시
  sheet.getRange('D5:F100').setBackground('#f5f5f5').setFontColor('#9e9e9e');
  sheet.getRange('I5:K100').setBackground('#f5f5f5').setFontColor('#9e9e9e');

  sheet.setFrozenRows(4);
}

function setupHistorySheet(ss) {
  let sheet = ss.getSheetByName(SHEET_HISTORY);
  if (!sheet) sheet = ss.insertSheet(SHEET_HISTORY);
  sheet.clear();

  sheet.getRange('A1:J1').merge()
    .setValue('📋 거래내역 — 자동 기록됩니다 (수정하지 마세요)')
    .setBackground('#0d0f1a')
    .setFontColor('#F0B90B')
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(1, 40);

  const headers = ['날짜·시간', '거래소', '코인명', '심볼', '구분', '수량', '거래가격', '거래금액', '수수료 (0.05%)', '비고'];
  sheet.getRange(2, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#1e2236')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontSize(11);
  sheet.setRowHeight(2, 35);

  const colWidths = [150, 80, 100, 70, 60, 100, 130, 130, 120, 150];
  colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.setFrozenRows(2);
}

function setupAlertSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_ALERT);
  if (!sheet) sheet = ss.insertSheet(SHEET_ALERT);
  sheet.clear();

  sheet.getRange('A1:H1').merge()
    .setValue('⚙️ 알림설정 — 목표가·손절가 도달 시 이메일 발송')
    .setBackground('#0d0f1a')
    .setFontColor('#F0B90B')
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(1, 40);

  // 이메일 설정
  sheet.getRange('A2').setValue('📧 알림 받을 이메일')
    .setFontWeight('bold').setFontColor('#1e2236');
  sheet.getRange('B2').setValue('')
    .setBackground('#fffde7')
    .setFontColor('#333333');
  sheet.getRange('C2').setValue('← 여기에 이메일 주소를 입력하세요')
    .setFontColor('#9e9e9e').setFontStyle('italic');

  // 주간 요약 설정
  sheet.getRange('A3').setValue('📅 주간 요약 이메일')
    .setFontWeight('bold').setFontColor('#1e2236');
  sheet.getRange('B3').setValue('ON')
    .setBackground('#fffde7');
  sheet.getRange('C3').setValue('← ON 또는 OFF 입력')
    .setFontColor('#9e9e9e').setFontStyle('italic');

  sheet.getRange('A4').setValue('📅 월간 요약 이메일')
    .setFontWeight('bold').setFontColor('#1e2236');
  sheet.getRange('B4').setValue('ON')
    .setBackground('#fffde7');

  sheet.setRowHeight(2, 30);
  sheet.setRowHeight(3, 30);
  sheet.setRowHeight(4, 30);

  const headers = ['거래소', '코인명', '심볼', '목표가 (KRW)', '손절가 (KRW)', '알림', '마지막 알림', '상태'];
  sheet.getRange(6, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#1e2236')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontSize(11);
  sheet.setRowHeight(6, 35);

  const sampleAlerts = [
    ['업비트', '비트코인', 'BTC', 180000000, 120000000, 'ON', '', '대기중'],
    ['업비트', '이더리움', 'ETH', 7000000, 4000000, 'ON', '', '대기중'],
  ];
  sheet.getRange(7, 1, sampleAlerts.length, headers.length).setValues(sampleAlerts);
  sheet.getRange('A7:F100').setBackground('#fffde7');

  const colWidths = [80, 100, 70, 130, 130, 60, 150, 80];
  colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.setFrozenRows(6);
}

function setupDashboardSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_DASHBOARD);
  if (!sheet) sheet = ss.insertSheet(SHEET_DASHBOARD);
  sheet.clear();

  sheet.getRange('A1:J1').merge()
    .setValue('📈 대시보드 — 포트폴리오 종합 현황')
    .setBackground('#0d0f1a')
    .setFontColor('#F0B90B')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(1, 45);

  // 요약 카드 영역
  const cardLabels = [
    ['총 평가금 (KRW)', '총 평가금 (USD)', '총 수익금', '수익률', '현금 비율'],
    ['', '', '', '', '']
  ];
  sheet.getRange('A3:E4').setValues(cardLabels)
    .setBackground('#151929').setFontColor('#8b93bb')
    .setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(3, 30);
  sheet.setRowHeight(4, 45);

  // 거래소별 분리
  sheet.getRange('A6').setValue('🏦 거래소별 현황')
    .setFontWeight('bold').setFontSize(12).setFontColor('#0d0f1a');
  const exchHeaders = [['거래소', '코인수', '평가금 (KRW)', '수익금', '수익률']];
  sheet.getRange('A7:E7').setValues(exchHeaders)
    .setBackground('#1e2236').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');

  // 시뮬레이터 영역
  sheet.getRange('G3').setValue('💰 수익 실현 시뮬레이터')
    .setFontWeight('bold').setFontSize(12).setFontColor('#0d0f1a');
  sheet.getRange('G4').setValue('코인 심볼').setFontWeight('bold');
  sheet.getRange('H4').setValue('').setBackground('#fffde7');
  sheet.getRange('G5').setValue('매도 수량').setFontWeight('bold');
  sheet.getRange('H5').setValue('').setBackground('#fffde7');
  sheet.getRange('G6').setValue('매도 가격').setFontWeight('bold');
  sheet.getRange('H6').setValue('').setBackground('#fffde7');
  sheet.getRange('G7').setValue('예상 수익금').setFontWeight('bold');
  sheet.getRange('H7').setValue('시뮬레이터 실행 후 표시')
    .setFontColor('#9e9e9e').setFontStyle('italic');
  sheet.getRange('G8').setValue('예상 수익률').setFontWeight('bold');
  sheet.getRange('H8').setValue('')
    .setFontColor('#9e9e9e').setFontStyle('italic');

  const colWidths = [120, 120, 130, 120, 100, 30, 130, 150, 130];
  colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}

// ================================================================
// 3. 시세 업데이트 (메인 함수)
// ================================================================
function updatePrices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PORTFOLIO);
  if (!sheet) { initializeSheets(); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 5) return;

  const dataRange = sheet.getRange(5, 1, lastRow - 4, 12);
  const rows = dataRange.getValues();

  // 업비트 심볼 수집
  const upbitSymbols = [];
  const bithumbSymbols = [];

  rows.forEach(row => {
    const exchange = String(row[0]).trim();
    const symbol = String(row[2]).trim().toUpperCase();
    if (!symbol) return;
    if (exchange === '빗썸') bithumbSymbols.push(symbol);
    else upbitSymbols.push(symbol);
  });

  // 업비트 시세
  const upbitPrices = upbitSymbols.length > 0 ? fetchUpbitPrices(upbitSymbols) : {};
  // 빗썸 시세
  const bithumbPrices = bithumbSymbols.length > 0 ? fetchBithumbPrices(bithumbSymbols) : {};
  // 환율
  const usdRate = getUSDRate();

  // 행별 업데이트
  rows.forEach((row, i) => {
    const exchange = String(row[0]).trim();
    const symbol = String(row[2]).trim().toUpperCase();
    if (!symbol) return;

    const priceData = exchange === '빗썸' ? bithumbPrices[symbol] : upbitPrices[symbol];
    if (!priceData) return;

    const currentPrice = priceData.currentPrice;
    const changeRate = priceData.changeRate;
    const qty = Number(row[6]) || 0;
    const avgPrice = Number(row[7]) || 0;
    const evaluation = currentPrice * qty;
    const investment = avgPrice * qty;
    const profit = evaluation - investment;
    const profitRate = investment > 0 ? ((profit / investment) * 100) : 0;
    const priceUSD = usdRate > 0 ? currentPrice / usdRate : 0;

    const rowNum = i + 5;

    // 현재가 KRW
    sheet.getRange(rowNum, 4).setValue(currentPrice)
      .setNumberFormat('#,##0');

    // 전일대비
    const changeCell = sheet.getRange(rowNum, 5);
    changeCell.setValue(changeRate).setNumberFormat('0.00"%"');
    changeCell.setFontColor(changeRate >= 0 ? '#FF5757' : '#4EA8DE');

    // 현재가 USD
    sheet.getRange(rowNum, 6).setValue(priceUSD)
      .setNumberFormat('$#,##0.00');

    // 평가금
    sheet.getRange(rowNum, 9).setValue(evaluation)
      .setNumberFormat('#,##0');

    // 수익금
    sheet.getRange(rowNum, 10).setValue(profit)
      .setNumberFormat('#,##0');

    // 수익률
    const profitCell = sheet.getRange(rowNum, 11);
    profitCell.setValue(profitRate).setNumberFormat('0.00"%"');
    profitCell.setFontColor(profitRate >= 0 ? '#FF5757' : '#4EA8DE');
  });

  // 요약 업데이트
  updateSummary(sheet, lastRow);

  // 알림 체크
  checkAlerts(upbitPrices, bithumbPrices);

  // 업데이트 시간
  sheet.getRange('K2').setValue(new Date())
    .setNumberFormat('yyyy-MM-dd HH:mm:ss')
    .setFontColor('#8b93bb').setFontSize(10);
}

// ================================================================
// 4. 요약 업데이트
// ================================================================
function updateSummary(sheet, lastRow) {
  if (lastRow < 5) return;

  let totalEval = 0, totalInvest = 0;
  const rows = sheet.getRange(5, 1, lastRow - 4, 11).getValues();

  rows.forEach(row => {
    if (!String(row[2]).trim()) return;
    totalEval   += Number(row[8]) || 0;
    totalInvest += (Number(row[6]) || 0) * (Number(row[7]) || 0);
  });

  const totalProfit = totalEval - totalInvest;
  const totalRate = totalInvest > 0 ? ((totalProfit / totalInvest) * 100) : 0;
  const cashRatio = totalEval > 0 ? ((totalInvest / totalEval) * 100).toFixed(1) + '%' : '-';

  // A3: 총 평가금
  sheet.getRange('A3').setValue(totalEval).setNumberFormat('#,##0')
    .setFontColor('#ffffff').setFontSize(15).setFontWeight('bold');

  // C3: 총 투자금
  sheet.getRange('C3').setValue(totalInvest).setNumberFormat('#,##0')
    .setFontColor('#ffffff').setFontSize(15).setFontWeight('bold');

  // E3: 총 수익금
  const profitCell = sheet.getRange('E3');
  profitCell.setValue(totalProfit).setNumberFormat('#,##0')
    .setFontSize(15).setFontWeight('bold');
  profitCell.setFontColor(totalProfit >= 0 ? '#FF5757' : '#4EA8DE');

  // G3: 수익률
  const rateCell = sheet.getRange('G3');
  rateCell.setValue(totalRate).setNumberFormat('0.00"%"')
    .setFontSize(15).setFontWeight('bold');
  rateCell.setFontColor(totalRate >= 0 ? '#FF5757' : '#4EA8DE');

  // I3: 현금 비율
  sheet.getRange('I3').setValue(cashRatio)
    .setFontColor('#F0B90B').setFontSize(15).setFontWeight('bold');
}

// ================================================================
// 5. 업비트 API
// ================================================================
function fetchUpbitPrices(symbols) {
  try {
    const markets = symbols.map(s => 'KRW-' + s).join(',');
    const url = UPBIT_API + '?markets=' + markets;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return {};

    const data = JSON.parse(res.getContentText());
    const result = {};
    data.forEach(item => {
      const symbol = item.market.replace('KRW-', '');
      result[symbol] = {
        currentPrice: item.trade_price,
        changeRate: item.signed_change_rate * 100
      };
    });
    return result;
  } catch (e) {
    console.error('업비트 API 오류:', e);
    return {};
  }
}

// ================================================================
// 6. 빗썸 API
// ================================================================
function fetchBithumbPrices(symbols) {
  const result = {};
  symbols.forEach(symbol => {
    try {
      const url = BITHUMB_API + '/' + symbol + '_KRW';
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) return;
      const data = JSON.parse(res.getContentText());
      if (data.status !== '0000') return;
      const d = data.data;
      result[symbol] = {
        currentPrice: Number(d.closing_price),
        changeRate: Number(d.fluctate_rate_24H)
      };
    } catch (e) {
      console.error('빗썸 API 오류 (' + symbol + '):', e);
    }
  });
  return result;
}

// ================================================================
// 7. 환율
// ================================================================
function getUSDRate() {
  try {
    const res = UrlFetchApp.fetch(EXCHANGE_API, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return 0;
    const data = JSON.parse(res.getContentText());
    return data.rates.KRW || 0;
  } catch (e) {
    return 0;
  }
}

function updateExchangeRate() {
  const rate = getUSDRate();
  if (rate > 0) {
    SpreadsheetApp.getUi().alert('💱 현재 환율: 1 USD = ' + rate.toLocaleString('ko-KR') + ' KRW');
  } else {
    SpreadsheetApp.getUi().alert('❌ 환율 조회 실패. 잠시 후 다시 시도해주세요.');
  }
}

// ================================================================
// 8. 거래 입력 (매수)
// ================================================================
function openBuyDialog() {
  const ui = SpreadsheetApp.getUi();
  const exchange = ui.prompt('매수 입력 (1/4)', '거래소를 입력하세요\n(업비트 또는 빗썸)', ui.ButtonSet.OK_CANCEL);
  if (exchange.getSelectedButton() !== ui.Button.OK) return;

  const symbol = ui.prompt('매수 입력 (2/4)', '코인 심볼을 입력하세요\n(예: BTC, ETH, XRP)', ui.ButtonSet.OK_CANCEL);
  if (symbol.getSelectedButton() !== ui.Button.OK) return;

  const qty = ui.prompt('매수 입력 (3/4)', '매수 수량을 입력하세요', ui.ButtonSet.OK_CANCEL);
  if (qty.getSelectedButton() !== ui.Button.OK) return;

  const price = ui.prompt('매수 입력 (4/4)', '매수 가격 (KRW)을 입력하세요', ui.ButtonSet.OK_CANCEL);
  if (price.getSelectedButton() !== ui.Button.OK) return;

  const qtyNum = Number(qty.getResponseText());
  const priceNum = Number(price.getResponseText().replace(/,/g, ''));
  const fee = qtyNum * priceNum * 0.0005;

  recordTrade(
    exchange.getResponseText(),
    symbol.getResponseText().toUpperCase(),
    '매수', qtyNum, priceNum, fee, ''
  );

  updatePortfolioAfterTrade(
    exchange.getResponseText(),
    symbol.getResponseText().toUpperCase(),
    '매수', qtyNum, priceNum
  );

  ui.alert('✅ 매수 거래가 기록되었습니다!\n\n' +
    '거래소: ' + exchange.getResponseText() + '\n' +
    '코인: ' + symbol.getResponseText().toUpperCase() + '\n' +
    '수량: ' + qtyNum.toLocaleString() + '\n' +
    '금액: ₩' + (qtyNum * priceNum).toLocaleString('ko-KR') + '\n' +
    '수수료: ₩' + Math.round(fee).toLocaleString('ko-KR'));
}

// ================================================================
// 9. 거래 입력 (매도)
// ================================================================
function openSellDialog() {
  const ui = SpreadsheetApp.getUi();
  const exchange = ui.prompt('매도 입력 (1/4)', '거래소를 입력하세요\n(업비트 또는 빗썸)', ui.ButtonSet.OK_CANCEL);
  if (exchange.getSelectedButton() !== ui.Button.OK) return;

  const symbol = ui.prompt('매도 입력 (2/4)', '코인 심볼을 입력하세요\n(예: BTC, ETH, XRP)', ui.ButtonSet.OK_CANCEL);
  if (symbol.getSelectedButton() !== ui.Button.OK) return;

  const qty = ui.prompt('매도 입력 (3/4)', '매도 수량을 입력하세요', ui.ButtonSet.OK_CANCEL);
  if (qty.getSelectedButton() !== ui.Button.OK) return;

  const price = ui.prompt('매도 입력 (4/4)', '매도 가격 (KRW)을 입력하세요', ui.ButtonSet.OK_CANCEL);
  if (price.getSelectedButton() !== ui.Button.OK) return;

  const qtyNum = Number(qty.getResponseText());
  const priceNum = Number(price.getResponseText().replace(/,/g, ''));
  const fee = qtyNum * priceNum * 0.0005;

  recordTrade(
    exchange.getResponseText(),
    symbol.getResponseText().toUpperCase(),
    '매도', qtyNum, priceNum, fee, ''
  );

  updatePortfolioAfterTrade(
    exchange.getResponseText(),
    symbol.getResponseText().toUpperCase(),
    '매도', qtyNum, priceNum
  );

  ui.alert('✅ 매도 거래가 기록되었습니다!\n\n' +
    '거래소: ' + exchange.getResponseText() + '\n' +
    '코인: ' + symbol.getResponseText().toUpperCase() + '\n' +
    '수량: ' + qtyNum.toLocaleString() + '\n' +
    '금액: ₩' + (qtyNum * priceNum).toLocaleString('ko-KR') + '\n' +
    '수수료: ₩' + Math.round(fee).toLocaleString('ko-KR'));
}

// ================================================================
// 10. 거래내역 기록
// ================================================================
function recordTrade(exchange, symbol, type, qty, price, fee, memo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_HISTORY);
  if (!sheet) return;

  const lastRow = Math.max(sheet.getLastRow(), 2);
  const newRow = lastRow + 1;
  const coinName = getCoinName(symbol);

  sheet.getRange(newRow, 1, 1, 10).setValues([[
    new Date(), exchange, coinName, symbol,
    type, qty, price, qty * price, Math.round(fee), memo
  ]]);

  sheet.getRange(newRow, 5).setFontColor(type === '매수' ? '#FF5757' : '#4EA8DE')
    .setFontWeight('bold');
  sheet.getRange(newRow, 1).setNumberFormat('yyyy-MM-dd HH:mm:ss');
  sheet.getRange(newRow, 7, 1, 3).setNumberFormat('#,##0');
}

function getCoinName(symbol) {
  const names = {
    'BTC': '비트코인', 'ETH': '이더리움', 'XRP': '리플',
    'DOGE': '도지코인', 'SOL': '솔라나', 'ADA': '에이다',
    'MATIC': '폴리곤', 'DOT': '폴카닷', 'AVAX': '아발란체',
    'LINK': '체인링크', 'UNI': '유니스왑', 'ATOM': '코스모스'
  };
  return names[symbol] || symbol;
}

// ================================================================
// 11. 포트폴리오 매수/매도 반영
// ================================================================
function updatePortfolioAfterTrade(exchange, symbol, type, qty, price) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PORTFOLIO);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 5) return;

  const rows = sheet.getRange(5, 1, lastRow - 4, 9).getValues();
  let found = false;

  rows.forEach((row, i) => {
    if (String(row[0]).trim() !== exchange) return;
    if (String(row[2]).trim().toUpperCase() !== symbol) return;

    found = true;
    const rowNum = i + 5;
    let currentQty = Number(row[6]) || 0;
    let currentAvg = Number(row[7]) || 0;

    if (type === '매수') {
      const totalCost = currentAvg * currentQty + price * qty;
      currentQty += qty;
      currentAvg = currentQty > 0 ? totalCost / currentQty : 0;
    } else {
      currentQty = Math.max(0, currentQty - qty);
      if (currentQty === 0) currentAvg = 0;
    }

    sheet.getRange(rowNum, 7).setValue(currentQty);
    sheet.getRange(rowNum, 8).setValue(Math.round(currentAvg)).setNumberFormat('#,##0');
  });

  if (!found && type === '매수') {
    const newRow = lastRow + 1;
    const coinName = getCoinName(symbol);
    sheet.getRange(newRow, 1, 1, 8).setValues([[
      exchange, coinName, symbol, '', '', '', qty, price
    ]]);
    sheet.getRange(newRow, 1, 1, 3).setBackground('#fffde7');
    sheet.getRange(newRow, 7, 1, 2).setBackground('#fffde7');
  }
}

// ================================================================
// 12. 알림 체크 (이메일)
// ================================================================
function checkAlerts(upbitPrices, bithumbPrices) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const alertSheet = ss.getSheetByName(SHEET_ALERT);
  if (!alertSheet) return;

  const email = alertSheet.getRange('B2').getValue();
  if (!email) return;

  const lastRow = alertSheet.getLastRow();
  if (lastRow < 7) return;

  const rows = alertSheet.getRange(7, 1, lastRow - 6, 8).getValues();

  rows.forEach((row, i) => {
    const exchange  = String(row[0]).trim();
    const coinName  = String(row[1]).trim();
    const symbol    = String(row[2]).trim().toUpperCase();
    const targetHigh = Number(row[3]) || 0;
    const targetLow  = Number(row[4]) || 0;
    const alertOn    = String(row[5]).trim().toUpperCase() === 'ON';
    if (!symbol || !alertOn) return;

    const priceData = exchange === '빗썸' ? bithumbPrices[symbol] : upbitPrices[symbol];
    if (!priceData) return;

    const currentPrice = priceData.currentPrice;
    const rowNum = i + 7;
    let alertMsg = '';

    if (targetHigh > 0 && currentPrice >= targetHigh) {
      alertMsg = '🎯 목표가 도달!';
    } else if (targetLow > 0 && currentPrice <= targetLow) {
      alertMsg = '⚠️ 손절가 도달!';
    }

    if (!alertMsg) return;

    const subject = '[코인 트래커 Pro] ' + alertMsg + ' ' + coinName + ' (' + symbol + ')';
    const body =
      alertMsg + '\n\n' +
      '━━━━━━━━━━━━━━━━━━━━\n' +
      '코인: ' + coinName + ' (' + symbol + ')\n' +
      '거래소: ' + exchange + '\n' +
      '현재가: ₩' + currentPrice.toLocaleString('ko-KR') + '\n' +
      (targetHigh > 0 ? '목표가: ₩' + targetHigh.toLocaleString('ko-KR') + '\n' : '') +
      (targetLow > 0 ? '손절가: ₩' + targetLow.toLocaleString('ko-KR') + '\n' : '') +
      '알림시간: ' + new Date().toLocaleString('ko-KR') + '\n' +
      '━━━━━━━━━━━━━━━━━━━━\n\n' +
      '실시간 코인 트래커 Pro — 업비트';

    try {
      MailApp.sendEmail(email, subject, body);
      alertSheet.getRange(rowNum, 7).setValue(new Date())
        .setNumberFormat('yyyy-MM-dd HH:mm:ss');
      alertSheet.getRange(rowNum, 8).setValue(alertMsg)
        .setFontColor(alertMsg.includes('목표가') ? '#FF5757' : '#4EA8DE');
    } catch (e) {
      console.error('이메일 발송 오류:', e);
    }
  });
}

// ================================================================
// 13. 수익 시뮬레이터
// ================================================================
function openSimulator() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const portfolio = ss.getSheetByName(SHEET_PORTFOLIO);

  const symbol = ui.prompt('💰 수익 실현 시뮬레이터 (1/2)',
    '시뮬레이션할 코인 심볼을 입력하세요\n(예: BTC, ETH)', ui.ButtonSet.OK_CANCEL);
  if (symbol.getSelectedButton() !== ui.Button.OK) return;

  const sellQty = ui.prompt('💰 수익 실현 시뮬레이터 (2/2)',
    '매도할 수량을 입력하세요', ui.ButtonSet.OK_CANCEL);
  if (sellQty.getSelectedButton() !== ui.Button.OK) return;

  const sym = symbol.getResponseText().toUpperCase().trim();
  const qty = Number(sellQty.getResponseText());

  const lastRow = portfolio.getLastRow();
  if (lastRow < 5) return;

  const rows = portfolio.getRange(5, 1, lastRow - 4, 11).getValues();
  let found = false;

  rows.forEach(row => {
    if (String(row[2]).trim().toUpperCase() !== sym) return;
    found = true;

    const currentPrice = Number(row[3]) || 0;
    const avgPrice = Number(row[7]) || 0;
    const sellAmount = currentPrice * qty;
    const costAmount = avgPrice * qty;
    const profit = sellAmount - costAmount;
    const profitRate = costAmount > 0 ? ((profit / costAmount) * 100) : 0;
    const fee = sellAmount * 0.0005;

    ui.alert(
      '💰 수익 실현 시뮬레이션 결과\n\n' +
      '━━━━━━━━━━━━━━━━━━━━\n' +
      '코인: ' + getCoinName(sym) + ' (' + sym + ')\n' +
      '현재가: ₩' + currentPrice.toLocaleString('ko-KR') + '\n' +
      '매수평균가: ₩' + avgPrice.toLocaleString('ko-KR') + '\n' +
      '매도수량: ' + qty.toLocaleString() + '\n' +
      '━━━━━━━━━━━━━━━━━━━━\n' +
      '예상 매도금액: ₩' + Math.round(sellAmount).toLocaleString('ko-KR') + '\n' +
      '예상 수수료: ₩' + Math.round(fee).toLocaleString('ko-KR') + '\n' +
      '예상 수익금: ₩' + Math.round(profit).toLocaleString('ko-KR') + '\n' +
      '예상 수익률: ' + profitRate.toFixed(2) + '%\n' +
      '━━━━━━━━━━━━━━━━━━━━'
    );
  });

  if (!found) {
    ui.alert('❌ 포트폴리오에서 ' + sym + ' 을 찾을 수 없습니다.');
  }
}

// ================================================================
// 14. 대시보드 새로고침
// ================================================================
function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const portfolio = ss.getSheetByName(SHEET_PORTFOLIO);
  const dashboard = ss.getSheetByName(SHEET_DASHBOARD);
  if (!portfolio || !dashboard) return;

  const lastRow = portfolio.getLastRow();
  if (lastRow < 5) return;

  const rows = portfolio.getRange(5, 1, lastRow - 4, 11).getValues();
  const usdRate = getUSDRate();

  let totalEval = 0, totalInvest = 0;
  const exchangeMap = {};

  rows.forEach(row => {
    const exchange = String(row[0]).trim();
    if (!String(row[2]).trim()) return;
    const eval_ = Number(row[8]) || 0;
    const invest = (Number(row[6]) || 0) * (Number(row[7]) || 0);
    totalEval += eval_;
    totalInvest += invest;

    if (!exchangeMap[exchange]) exchangeMap[exchange] = { eval: 0, invest: 0, count: 0 };
    exchangeMap[exchange].eval += eval_;
    exchangeMap[exchange].invest += invest;
    exchangeMap[exchange].count++;
  });

  const totalProfit = totalEval - totalInvest;
  const totalRate = totalInvest > 0 ? ((totalProfit / totalInvest) * 100) : 0;
  const totalEvalUSD = usdRate > 0 ? totalEval / usdRate : 0;
  const cashRatio = totalEval > 0 ? ((totalInvest / totalEval) * 100).toFixed(1) + '%' : '-';

  // 요약 카드
  const summaryValues = [[totalEval, totalEvalUSD, totalProfit, totalRate, cashRatio]];
  dashboard.getRange('A4:E4').setValues(summaryValues);
  dashboard.getRange('A4').setNumberFormat('#,##0').setFontWeight('bold').setFontSize(13);
  dashboard.getRange('B4').setNumberFormat('$#,##0').setFontWeight('bold').setFontSize(13);
  dashboard.getRange('C4').setNumberFormat('#,##0').setFontWeight('bold').setFontSize(13)
    .setFontColor(totalProfit >= 0 ? '#FF5757' : '#4EA8DE');
  dashboard.getRange('D4').setNumberFormat('0.00"%"').setFontWeight('bold').setFontSize(13)
    .setFontColor(totalRate >= 0 ? '#FF5757' : '#4EA8DE');
  dashboard.getRange('E4').setFontWeight('bold').setFontSize(13).setFontColor('#F0B90B');

  // 거래소별 현황
  let startRow = 8;
  Object.keys(exchangeMap).forEach(exchange => {
    const d = exchangeMap[exchange];
    const profit = d.eval - d.invest;
    const rate = d.invest > 0 ? ((profit / d.invest) * 100) : 0;
    dashboard.getRange(startRow, 1, 1, 5).setValues([
      [exchange, d.count, d.eval, profit, rate]
    ]);
    dashboard.getRange(startRow, 3).setNumberFormat('#,##0');
    dashboard.getRange(startRow, 4).setNumberFormat('#,##0')
      .setFontColor(profit >= 0 ? '#FF5757' : '#4EA8DE');
    dashboard.getRange(startRow, 5).setNumberFormat('0.00"%"')
      .setFontColor(rate >= 0 ? '#FF5757' : '#4EA8DE');
    startRow++;
  });

  SpreadsheetApp.getUi().alert('✅ 대시보드가 업데이트되었습니다!');
}

// ================================================================
// 15. 주간·월간 이메일 자동 발송
// ================================================================
function sendWeeklySummary() {
  sendSummaryEmail('주간');
}

function sendMonthlySummary() {
  sendSummaryEmail('월간');
}

function sendSummaryEmail(type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const alertSheet = ss.getSheetByName(SHEET_ALERT);
  if (!alertSheet) return;

  const email = alertSheet.getRange('B2').getValue();
  const setting = alertSheet.getRange(type === '주간' ? 'B3' : 'B4').getValue();
  if (!email || String(setting).toUpperCase() !== 'ON') return;

  const portfolio = ss.getSheetByName(SHEET_PORTFOLIO);
  const lastRow = portfolio.getLastRow();
  if (lastRow < 5) return;

  const rows = portfolio.getRange(5, 1, lastRow - 4, 11).getValues();
  let totalEval = 0, totalInvest = 0;
  let coinDetails = '';

  rows.forEach(row => {
    if (!String(row[2]).trim()) return;
    const eval_ = Number(row[8]) || 0;
    const invest = (Number(row[6]) || 0) * (Number(row[7]) || 0);
    totalEval += eval_;
    totalInvest += invest;

    const profit = eval_ - invest;
    const rate = invest > 0 ? ((profit / invest) * 100).toFixed(2) : '0.00';
    coinDetails += String(row[1]) + '(' + String(row[2]) + '): ' +
      '₩' + Math.round(eval_).toLocaleString('ko-KR') +
      ' (' + (profit >= 0 ? '+' : '') + rate + '%)\n';
  });

  const totalProfit = totalEval - totalInvest;
  const totalRate = totalInvest > 0 ? ((totalProfit / totalInvest) * 100).toFixed(2) : '0.00';

  const subject = '[코인 트래커 Pro] ' + type + ' 포트폴리오 요약 — ' +
    new Date().toLocaleDateString('ko-KR');
  const body =
    '안녕하세요! ' + type + ' 포트폴리오 요약입니다.\n\n' +
    '━━━━━━━━━━━━━━━━━━━━\n' +
    '📊 총 평가금: ₩' + Math.round(totalEval).toLocaleString('ko-KR') + '\n' +
    '💵 총 투자금: ₩' + Math.round(totalInvest).toLocaleString('ko-KR') + '\n' +
    '💰 총 수익금: ₩' + Math.round(totalProfit).toLocaleString('ko-KR') + '\n' +
    '📈 총 수익률: ' + (totalProfit >= 0 ? '+' : '') + totalRate + '%\n' +
    '━━━━━━━━━━━━━━━━━━━━\n\n' +
    '📋 코인별 현황:\n' + coinDetails + '\n' +
    '━━━━━━━━━━━━━━━━━━━━\n' +
    '기준시간: ' + new Date().toLocaleString('ko-KR') + '\n\n' +
    '실시간 코인 트래커 Pro — 업비트';

  try {
    MailApp.sendEmail(email, subject, body);
  } catch (e) {
    console.error(type + ' 이메일 발송 오류:', e);
  }
}

// ================================================================
// 16. 트리거 설정
// ================================================================
function setupTrigger() {
  removeTrigger();

  // 1분마다 시세 업데이트
  ScriptApp.newTrigger('updatePrices')
    .timeBased().everyMinutes(1).create();

  // 매주 월요일 오전 9시 주간 요약
  ScriptApp.newTrigger('sendWeeklySummary')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9).create();

  // 매월 1일 오전 9시 월간 요약
  ScriptApp.newTrigger('sendMonthlySummary')
    .timeBased().onMonthDay(1).atHour(9).create();

  SpreadsheetApp.getUi().alert(
    '✅ 자동 업데이트 설정 완료!\n\n' +
    '• 시세: 1분마다 자동 업데이트\n' +
    '• 주간 요약: 매주 월요일 오전 9시\n' +
    '• 월간 요약: 매월 1일 오전 9시\n\n' +
    '중지하려면 메뉴 → 자동 업데이트 중지'
  );

  updatePrices();
}

function removeTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
}

// ================================================================
// 17. API 테스트
// ================================================================
function testAPI() {
  const ui = SpreadsheetApp.getUi();
  try {
    const upbit = fetchUpbitPrices(['BTC']);
    const bithumb = fetchBithumbPrices(['BTC']);
    const usd = getUSDRate();

    ui.alert(
      '✅ API 연결 테스트 성공!\n\n' +
      '━━━━━━━━━━━━━━━━━━━━\n' +
      '업비트 BTC: ₩' + (upbit['BTC'] ? upbit['BTC'].currentPrice.toLocaleString('ko-KR') : '조회 실패') + '\n' +
      '빗썸 BTC: ₩' + (bithumb['BTC'] ? bithumb['BTC'].currentPrice.toLocaleString('ko-KR') : '조회 실패') + '\n' +
      '환율: 1 USD = ₩' + (usd > 0 ? usd.toLocaleString('ko-KR') : '조회 실패') + '\n' +
      '━━━━━━━━━━━━━━━━━━━━'
    );
  } catch (e) {
    ui.alert('❌ API 연결 실패\n\n오류: ' + e.toString());
  }
}

// ================================================================
// 18. 사용 가이드
// ================================================================
function showGuide() {
  SpreadsheetApp.getUi().alert(
    '📖 실시간 코인 트래커 Pro — 사용 가이드\n\n' +
    '━━━━━━━━━━━━━━━━━━━━\n' +
    '1️⃣ 처음 시작\n' +
    '   메뉴 → 시트 초기 설정 클릭\n\n' +
    '2️⃣ 코인 정보 입력\n' +
    '   📊 포트폴리오 시트의 노란칸에 입력\n' +
    '   거래소 / 코인명 / 심볼 / 보유수량 / 매수평균가\n\n' +
    '3️⃣ 이메일 알림 설정\n' +
    '   ⚙️ 알림설정 시트에 이메일 주소 입력\n' +
    '   목표가·손절가 설정 후 알림 ON\n\n' +
    '4️⃣ 자동 업데이트 시작\n' +
    '   메뉴 → 자동 업데이트 시작 클릭\n' +
    '   1분마다 자동으로 시세 반영\n\n' +
    '5️⃣ 거래 입력\n' +
    '   메뉴 → 거래 입력 (매수/매도)\n' +
    '   자동으로 평균가와 수량 업데이트\n\n' +
    '━━━━━━━━━━━━━━━━━━━━\n' +
    '🟡 노란칸: 직접 입력\n' +
    '⬜ 회색칸: 자동 계산 (수정 불가)\n' +
    '🔴 빨간 숫자: 상승\n' +
    '🔵 파란 숫자: 하락\n' +
    '━━━━━━━━━━━━━━━━━━━━'
  );
}
