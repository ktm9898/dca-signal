/**
 * Google Apps Script - DCA Signal (정량 분할 적립식 퀀트 시그널 시스템)
 * Repository: ktm9898/dca-signal
 * 
 * [배포 및 자동화 설정 가이드]
 * 1. 스프레드시트 [확장 프로그램] -> [Apps Script]에서 본 코드로 덮어쓰기
 * 2. 좌측 톱니바퀴 [프로젝트 설정] -> 하단 [스크립트 속성(Script Properties)] 클릭
 *    - 속성 추가:
 *      * 속성: GITHUB_TOKEN
 *      * 값: 본인의 GitHub Personal Access Token (repo, workflow 권한 포함)
 * 3. 상단 함수 선택에서 `setupSheets` 선택 후 [실행] (시트 탭 자동 구성)
 * 4. 상단 함수 선택에서 `setupDailyTrigger` 선택 후 [실행]
 *    -> 매일 아침 06:45~07:00 KST(미국 장 마감 후) 자동으로 구글 시트 시세 갱신 및
 *       GitHub Actions(update_data.yml)를 원격으로 깨워 최신 ETF 데이터를 정밀하게 빌드합니다!
 * 5. 우측 상단 [배포] -> [배포 관리] -> 연필(수정) 아이콘 클릭 후
 *    - 버전: "새 버전" 선택
 *    - [배포] 클릭 (기존 웹앱 URL 유지)
 */

const TARGET_TICKERS = ["TQQQ", "SOXL", "UPRO", "QLD", "QQQ", "SPY", "122630.KS"];

/**
 * 📊 시트 자동 생성 및 탭 통합 (DCA_Active_Strategy 별도 탭 없이 DCA_Strategy_Slots 하나로 통합)
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 구버전 개별 활성탭이 남아있다면 삭제하여 탭 간소화
  const oldActiveSheet = ss.getSheetByName("DCA_Active_Strategy");
  if (oldActiveSheet) {
    try { ss.deleteSheet(oldActiveSheet); } catch (e) {}
  }

  // 1. 전략 슬롯 시트 표준화 및 생성
  getStandardizedSlotsSheet(ss);

  // 2. 최신 시세 캐시 시트 (DCA_Latest_Quotes)
  let quoteSheet = ss.getSheetByName("DCA_Latest_Quotes");
  if (!quoteSheet) {
    quoteSheet = ss.insertSheet("DCA_Latest_Quotes");
    quoteSheet.getRange("A1:G1").setValues([[
      "Ticker", "ClosePrice", "ChangePct", "High52", "Low52", "Date", "UpdatedAt"
    ]]);
    quoteSheet.getRange("A1:G1").setFontWeight("bold").setBackground("#fef3c7");
  }

  // 3. 완료 사이클 이력 시트 (DCA_Completed_Cycles)
  let cyclesSheet = ss.getSheetByName("DCA_Completed_Cycles");
  if (!cyclesSheet) {
    cyclesSheet = ss.insertSheet("DCA_Completed_Cycles");
    cyclesSheet.getRange("A1:H1").setValues([[
      "CycleNo", "Ticker", "StartDate", "EndDate", "DurationDays", "Invested", "Profit", "ReturnPct"
    ]]);
    cyclesSheet.getRange("A1:H1").setFontWeight("bold").setBackground("#dcfce7");
  }

  // 4. 실행 및 트리거 로그 시트 (Execution_Logs)
  let logSheet = ss.getSheetByName("Execution_Logs");
  if (!logSheet) {
    logSheet = ss.insertSheet("Execution_Logs");
    logSheet.getRange("A1:C1").setValues([["Timestamp", "Status", "Message"]]);
    logSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#e0e7ff");
  }
}

/**
 * 🛠️ DCA_Strategy_Slots 시트의 헤더 및 데이터 스키마 자동 보정/마이그레이션
 * 구버전(12컬럼: SlotID, Name, Memo, Ticker, SeedCapital...)이 있더라도
 * 최신 15컬럼 표준(SlotID, IsActive, Name, Memo, Ticker, StartDate, SeedCapital...)으로 완벽 자동 변환합니다.
 */
function getStandardizedSlotsSheet(ss) {
  const standardHeaders = [
    "SlotID", "IsActive", "Name", "Memo", "Ticker", "StartDate", 
    "SeedCapital", "Portions", "PhaseSplitRound", "LateMode", 
    "LateThresholdPct", "TargetProfitPct", "CompoundMode", "CurrentCycleNo", "UpdatedAt"
  ];

  let slotsSheet = ss.getSheetByName("DCA_Strategy_Slots");
  if (!slotsSheet) {
    slotsSheet = ss.insertSheet("DCA_Strategy_Slots");
    slotsSheet.getRange(1, 1, 1, standardHeaders.length).setValues([standardHeaders]);
    slotsSheet.getRange(1, 1, 1, standardHeaders.length).setFontWeight("bold").setBackground("#dbeafe");
    return slotsSheet;
  }

  const lastRow = slotsSheet.getLastRow();
  const lastCol = Math.max(slotsSheet.getLastColumn(), 1);

  if (lastRow === 0) {
    slotsSheet.getRange(1, 1, 1, standardHeaders.length).setValues([standardHeaders]);
    slotsSheet.getRange(1, 1, 1, standardHeaders.length).setFontWeight("bold").setBackground("#dbeafe");
    return slotsSheet;
  }

  const headerRow = slotsSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim().toLowerCase());
  const hasIsActive = headerRow.includes("isactive");
  const hasStartDate = headerRow.includes("startdate");

  // 만약 구버전 헤더(IsActive 컬럼 누락 등)라면 자동 마이그레이션 실행
  if (!hasIsActive || !hasStartDate || lastCol < standardHeaders.length) {
    const allData = slotsSheet.getDataRange().getValues();
    const oldHeaders = allData[0].map(h => String(h).trim().toLowerCase());
    const newRows = [];

    for (let i = 1; i < allData.length; i++) {
      const r = allData[i];
      if (!r[0]) continue;

      const getVal = (name, fallbackIdx, defVal) => {
        const idx = oldHeaders.indexOf(name.toLowerCase());
        if (idx >= 0 && r[idx] !== undefined && r[idx] !== "") return r[idx];
        if (fallbackIdx >= 0 && fallbackIdx < r.length && r[fallbackIdx] !== undefined && r[fallbackIdx] !== "") return r[fallbackIdx];
        return defVal;
      };

      const slotId = Number(r[0]) || i;
      // 기존 구버전 컬럼 위치:
      // [0] SlotID, [1] Name, [2] Memo, [3] Ticker, [4] SeedCapital, [5] Portions,
      // [6] PhaseSplitRound, [7] LateMode, [8] LateThresholdPct, [9] TargetProfitPct, [10] CompoundMode, [11] UpdatedAt
      const name = String(getVal("name", 1, `슬롯 ${slotId}`));
      const memo = String(getVal("memo", 2, ""));
      const ticker = String(getVal("ticker", 3, "TQQQ"));
      const seedCapital = Number(getVal("seedcapital", 4, 10000000)) || 10000000;
      const portions = Number(getVal("portions", 5, 40)) || 40;
      const phaseSplitRound = Number(getVal("phasesplitround", 6, 20)) || 20;
      const lateMode = String(getVal("latemode", 7, "cond"));
      const lateThresholdPct = Number(getVal("latethresholdpct", 8, -5)) || 0;
      const targetProfitPct = Number(getVal("targetprofitpct", 9, 10)) || 10;
      const compoundMode = String(getVal("compoundmode", 10, "simple"));
      const updatedAt = formatDateVal(getVal("updatedat", 11, Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm:ss")));
      const startDate = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd");
      const isActive = (slotId === 1); // 1번 슬롯을 기본 활성으로 부여

      newRows.push([
        slotId, isActive, name, memo, ticker, startDate,
        seedCapital, portions, phaseSplitRound, lateMode,
        lateThresholdPct, targetProfitPct, compoundMode, 1, updatedAt
      ]);
    }

    slotsSheet.clear();
    slotsSheet.clearFormats();
    slotsSheet.getRange(1, 1, 1, standardHeaders.length).setValues([standardHeaders]);
    slotsSheet.getRange(1, 1, 1, standardHeaders.length).setFontWeight("bold").setBackground("#dbeafe");
    if (newRows.length > 0) {
      slotsSheet.getRange(2, 1, newRows.length, standardHeaders.length).setValues(newRows);
    }
  }

  // 12번째 열(TargetProfitPct, L열)의 셀 서식을 명시적으로 숫자로 지정
  slotsSheet.getRange("L:L").setNumberFormat("0.##");

  const totalRows = slotsSheet.getLastRow();
  if (totalRows > 1) {
    const profitCol = slotsSheet.getRange(2, 12, totalRows - 1, 1);
    const pVals = profitCol.getValues();
    let changed = false;
    for (let i = 0; i < pVals.length; i++) {
      // 기존 날짜 서식으로 인해 Date 객체로 변환되어 저장되어 있던 셀은 숫자로 재설정
      if (pVals[i][0] instanceof Date) {
        pVals[i][0] = 10;
        changed = true;
      }
    }
    if (changed) {
      profitCol.setValues(pVals);
    }
  }

  return slotsSheet;
}

/**
 * ⏰ 매일 아침 자동 실행 통합 일일 작업
 * 1) 구글 시트의 DCA_Latest_Quotes 최신 시세 갱신
 * 2) GitHub Actions update_data.yml 원격 트리거 실행 (ETF 가격 수집 및 data_bundle.js 자동 빌드)
 */
function runDailyAutomation() {
  const nowStr = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm:ss");
  Logger.log(`[runDailyAutomation] 일일 자동화 시작: ${nowStr}`);

  // 1. 구글 시트 내부 최신 시세 갱신
  try {
    updateDailyQuotes();
  } catch (e) {
    Logger.log(`[runDailyAutomation] updateDailyQuotes 실패: ${e.toString()}`);
  }

  // 2. GitHub Actions 원격 트리거 실행
  try {
    triggerGitHubDataUpdate();
  } catch (e) {
    Logger.log(`[runDailyAutomation] triggerGitHubDataUpdate 실패: ${e.toString()}`);
  }
}

/**
 * ⏰ 매일 아침 06:45~07:00 KST 자동 실행 시간 트리거 등록 함수
 */
function setupDailyTrigger() {
  // 기존 중복 트리거 삭제
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    const fn = triggers[i].getHandlerFunction();
    if (fn === "updateDailyQuotes" || fn === "runDailyAutomation" || fn === "triggerGitHubDataUpdate") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 매일 오전 6시 45분~7시 사이에 실행되는 일일 타이머 등록
  ScriptApp.newTrigger("runDailyAutomation")
    .timeBased()
    .atHour(6)
    .nearMinute(45)
    .everyDays(1)
    .inTimezone("Asia/Seoul")
    .create();

  Logger.log("매일 아침 시세 갱신 및 GitHub Actions 자동 업데이트 트리거 등록 완료!");
  runDailyAutomation(); // 즉시 1회 테스트 실행
}

/**
 * 🚀 GitHub Actions 'update_data.yml' 워크플로우 원격 트리거 함수
 * - Google Apps Script 정기 트리거 또는 웹앱 요청 시 실행
 * - GitHub REST API (dispatches)를 호출하여 GitHub Actions를 정밀하게 실행
 */
function triggerGitHubDataUpdate() {
  const githubToken = PropertiesService.getScriptProperties().getProperty("GITHUB_TOKEN");
  if (!githubToken) {
    const msg = "GITHUB_TOKEN이 스크립트 속성(Script Properties)에 설정되지 않았습니다. [프로젝트 설정] -> [스크립트 속성]에 GITHUB_TOKEN을 추가해주세요.";
    Logger.log("[triggerGitHubDataUpdate] " + msg);
    logTriggerResult("FAILED", msg);
    return { success: false, message: msg };
  }

  const url = "https://api.github.com/repos/ktm9898/dca-signal/actions/workflows/update_data.yml/dispatches";
  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "Accept": "application/vnd.github+json",
      "Authorization": "Bearer " + githubToken.trim(),
      "User-Agent": "GoogleAppsScript"
    },
    payload: JSON.stringify({ ref: "main" }),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    if (code === 204 || code === 200) {
      Logger.log("[triggerGitHubDataUpdate] GitHub Actions update_data.yml 원격 트리거 성공 (HTTP " + code + ")");
      logTriggerResult("SUCCESS", "GitHub Actions update_data.yml 원격 트리거 완료 (정상 " + code + ")");
      return { success: true, message: "GitHub Actions 데이터 업데이트가 성공적으로 시작되었습니다." };
    } else {
      const errBody = response.getContentText();
      Logger.log(`[triggerGitHubDataUpdate] 트리거 실패 (HTTP ${code}): ${errBody}`);
      logTriggerResult("FAILED", `HTTP ${code}: ${errBody}`);
      return { success: false, message: `HTTP ${code}: ${errBody}` };
    }
  } catch (err) {
    Logger.log("[triggerGitHubDataUpdate] 예외 발생: " + err.toString());
    logTriggerResult("ERROR", err.toString());
    return { success: false, message: err.toString() };
  }
}

/**
 * 트리거 실행 결과를 Execution_Logs 시트에 기록
 */
function logTriggerResult(status, message) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    let logSheet = ss.getSheetByName("Execution_Logs");
    if (!logSheet) {
      logSheet = ss.insertSheet("Execution_Logs");
      logSheet.getRange("A1:C1").setValues([["Timestamp", "Status", "Message"]]);
      logSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#e0e7ff");
    }
    const timestamp = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm:ss");
    logSheet.appendRow([timestamp, status, message]);
  } catch (e) {
    Logger.log("logTriggerResult error: " + e.toString());
  }
}

/**
 * 주요 7개 ETF의 최신 종가를 야후 파이낸스에서 가져와 구글 시트에 업데이트 (아침 7시 30분 트리거 전용)
 */
function updateDailyQuotes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let quoteSheet = ss.getSheetByName("DCA_Latest_Quotes");
  if (!quoteSheet) {
    setupSheets();
    quoteSheet = ss.getSheetByName("DCA_Latest_Quotes");
  }

  const nowStr = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm:ss");
  const rows = [];

  for (const ticker of TARGET_TICKERS) {
    try {
      const q = fetchYahooQuote(ticker);
      if (q) {
        rows.push([ticker, q.close, q.changePct, q.high52, q.low52, q.date, nowStr]);
      }
    } catch (err) {
      Logger.log(`[${ticker}] fetch error: ${err}`);
    }
  }

  if (rows.length > 0) {
    quoteSheet.getRange(2, 1, quoteSheet.getLastRow() > 1 ? quoteSheet.getLastRow() - 1 : 1, 7).clearContent();
    quoteSheet.getRange(2, 1, rows.length, 7).setValues(rows);
    Logger.log(`최신 시세 ${rows.length}개 종목 업데이트 완료 (${nowStr})`);
  }
}

/**
 * Yahoo Finance API로부터 종가 및 통계 수신 (배치 트리거 전용)
 */
function fetchYahooQuote(ticker) {
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(ticker)}?range=1y&interval=1d`;
  const resp = UrlFetchApp.fetch(url, {
    headers: { "User-Agent": "Mozilla/5.0" },
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() === 200) {
    const json = JSON.parse(resp.getContentText());
    const chart = json.chart && json.chart.result && json.chart.result[0];
    if (chart && chart.timestamp && chart.indicators && chart.indicators.quote) {
      const ts = chart.timestamp;
      const quotes = chart.indicators.quote[0];
      const closes = [];
      const highs = [];
      const lows = [];
      let lastDate = "";

      for (let i = 0; i < ts.length; i++) {
        if (quotes.close[i] != null) {
          closes.push(quotes.close[i]);
          highs.push(quotes.high[i] || quotes.close[i]);
          lows.push(quotes.low[i] || quotes.close[i]);
          lastDate = Utilities.formatDate(new Date(ts[i] * 1000), "GMT", "yyyy-MM-dd");
        }
      }

      if (closes.length > 0) {
        const lastClose = closes[closes.length - 1];
        const prevClose = closes.length > 1 ? closes[closes.length - 2] : lastClose;
        const chgPct = Number((((lastClose - prevClose) / prevClose) * 100).toFixed(2));
        const high52 = Number(Math.max(...highs).toFixed(2));
        const low52 = Number(Math.min(...lows).toFixed(2));

        return {
          ticker: ticker,
          close: Number(lastClose.toFixed(2)),
          prevClose: Number(prevClose.toFixed(2)),
          changePct: chgPct,
          high52: high52,
          low52: low52,
          date: lastDate
        };
      }
    }
  }
  return null;
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "all";
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // 1. 전략 슬롯 목록 및 활성 전략 로드 (자동 스키마 검증)
    const slotsSheet = getStandardizedSlotsSheet(ss);
    const sData = slotsSheet.getDataRange().getValues();
    const headers = sData[0].map(h => String(h).trim().toLowerCase());

    const getCol = (r, name, fallbackIdx, defVal) => {
      const idx = headers.indexOf(name.toLowerCase());
      if (idx >= 0 && r[idx] !== undefined && r[idx] !== "") return r[idx];
      if (fallbackIdx >= 0 && fallbackIdx < r.length && r[fallbackIdx] !== undefined && r[fallbackIdx] !== "") return r[fallbackIdx];
      return defVal;
    };

    const slots = [];
    let activeStrat = null;

    for (let i = 1; i < sData.length; i++) {
      const r = sData[i];
      if (!r[0]) continue;
      
      const rawActive = getCol(r, "isactive", 1, false);
      const isActive = (rawActive === true || String(rawActive).toUpperCase() === "TRUE" || String(rawActive).toUpperCase() === "Y" || rawActive === 1);
      const slotId = Number(r[0]);

      const slotObj = {
        slotId: slotId,
        isActive: isActive,
        name: String(getCol(r, "name", 2, `슬롯 ${slotId}`)),
        memo: String(getCol(r, "memo", 3, "")),
        ticker: String(getCol(r, "ticker", 4, "")),
        startDate: formatDateVal(getCol(r, "startdate", 5, Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd"))).substring(0, 10),
        seedCapital: Number(getCol(r, "seedcapital", 6, 0)) || 0,
        portions: Number(getCol(r, "portions", 7, 0)) || 0,
        phaseSplitRound: Number(getCol(r, "phasesplitround", 8, 0)) || 0,
        lateMode: String(getCol(r, "latemode", 9, "cond")),
        lateThresholdPct: Number(getCol(r, "latethresholdpct", 10, 0)) || 0,
        targetProfitPct: Number(getCol(r, "targetprofitpct", 11, 10)) || 10,
        compoundMode: String(getCol(r, "compoundmode", 12, "simple")),
        currentCycleNo: Number(getCol(r, "currentcycleno", 13, 1)) || 1,
        updatedAt: formatDateVal(getCol(r, "updatedat", 14, ""))
      };

      slots.push(slotObj);
      if (isActive && !activeStrat) {
        activeStrat = slotObj;
      }
    }

    // 만약 활성 슬롯이 명시되지 않았지만 슬롯이 존재한다면 첫 번째 슬롯을 활성 전략으로 간주
    if (!activeStrat && slots.length > 0) {
      activeStrat = slots[0];
    }

    if (action === "get_slots") {
      return respondJSON({ success: true, slots: slots });
    }

    if (action === "get_active_strategy") {
      return respondJSON({ success: true, activeStrategy: activeStrat });
    }

    // 최신 시세 정보 (DCA_Latest_Quotes 시트에서 조회)
    let quoteObj = null;
    let quoteSheet = ss.getSheetByName("DCA_Latest_Quotes");
    if (quoteSheet && activeStrat && activeStrat.ticker) {
      const qData = quoteSheet.getDataRange().getValues();
      for (let i = 1; i < qData.length; i++) {
        if (String(qData[i][0]).toUpperCase() === activeStrat.ticker.toUpperCase()) {
          quoteObj = {
            ticker: qData[i][0],
            close: Number(qData[i][1]),
            changePct: Number(qData[i][2]),
            high52: Number(qData[i][3]),
            low52: Number(qData[i][4]),
            date: String(qData[i][5])
          };
          break;
        }
      }
    }

    // 완료된 사이클 이력
    let cyclesSheet = ss.getSheetByName("DCA_Completed_Cycles");
    let completedCycles = [];
    if (cyclesSheet) {
      const cData = cyclesSheet.getDataRange().getValues();
      for (let i = 1; i < cData.length; i++) {
        const r = cData[i];
        if (!r[0]) continue;
        completedCycles.push({
          cycleNo: Number(r[0]),
          ticker: String(r[1]),
          startDate: formatDateVal(r[2]).substring(0, 10),
          endDate: formatDateVal(r[3]).substring(0, 10),
          durationDays: Number(r[4]) || 0,
          invested: Number(r[5]) || 0,
          profit: Number(r[6]) || 0,
          retPct: Number(r[7]) || 0
        });
      }
    }

    return respondJSON({
      success: true,
      activeStrategy: activeStrat,
      latestQuote: quoteObj,
      slots: slots,
      completedCycles: completedCycles
    });

  } catch (err) {
    return respondJSON({ success: false, error: err.toString() });
  }
}

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let body = {};

  try {
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
  } catch (ex) {
    return respondJSON({ success: false, error: "Invalid JSON: " + ex });
  }

  const action = body.action || (e && e.parameter && e.parameter.action) || "";

  try {
    const slotsSheet = getStandardizedSlotsSheet(ss);

    // 1. 슬롯 저장 (save_slot)
    if (action === "save_slot") {
      const slotId = Number(body.slotId) || 1;
      const p = body.strategy || {};
      const makeActive = body.isActive === true;
      const nowStr = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm:ss");

      const sData = slotsSheet.getDataRange().getValues();
      let targetRow = -1;

      for (let i = 1; i < sData.length; i++) {
        if (Number(sData[i][0]) === slotId) {
          targetRow = i + 1;
        }
        // 만약 이번 슬롯을 활성화한다면 다른 슬롯들은 IsActive = false로 변경
        if (makeActive && Number(sData[i][0]) !== slotId) {
          slotsSheet.getRange(i + 1, 2).setValue(false);
        }
      }

      const rowValues = [
        slotId,
        makeActive,
        body.name || `슬롯 ${slotId}`,
        body.memo || "",
        p.ticker || "",
        p.startDate || Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd"),
        Number(p.seedCapital) || 0,
        Number(p.portions) || 0,
        Number(p.phaseSplitRound) || 0,
        p.lateMode || "cond",
        p.lateThresholdPct != null ? Number(p.lateThresholdPct) : 0,
        Number(p.targetProfitPct) || 10,
        p.compoundMode || "simple",
        Number(p.currentCycleNo) || 1,
        nowStr
      ];

      if (targetRow > 0) {
        slotsSheet.getRange(targetRow, 1, 1, 15).setValues([rowValues]);
        slotsSheet.getRange(targetRow, 12).setNumberFormat("0.##");
      } else {
        slotsSheet.appendRow(rowValues);
        slotsSheet.getRange(slotsSheet.getLastRow(), 12).setNumberFormat("0.##");
      }

      return respondJSON({ success: true, message: `슬롯 ${slotId}에 전략이 저장되었습니다.` });
    }

    // 2. 실전 활성 전략 지정 (set_active_strategy 또는 set_active_slot)
    if (action === "set_active_strategy" || action === "set_active_slot") {
      const targetSlotId = Number(body.slotId) || 0;
      const p = body.strategy || {};
      const nowStr = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm:ss");

      const sData = slotsSheet.getDataRange().getValues();
      let found = false;

      for (let i = 1; i < sData.length; i++) {
        const curSlotId = Number(sData[i][0]);
        const shouldActive = (targetSlotId > 0) ? (curSlotId === targetSlotId) : (sData[i][1] === true);

        slotsSheet.getRange(i + 1, 2).setValue(shouldActive);

        if (shouldActive) {
          found = true;
          // 활성 슬롯의 투자개시일이나 사이클 번호 업데이트가 전달된 경우 반영
          if (p.startDate) slotsSheet.getRange(i + 1, 6).setValue(p.startDate);
          if (p.currentCycleNo) slotsSheet.getRange(i + 1, 14).setValue(p.currentCycleNo);
          slotsSheet.getRange(i + 1, 15).setValue(nowStr);
        }
      }

      // 만약 시트에 아무 슬롯도 없는 상태에서 set_active_strategy가 들어왔다면 1번 슬롯으로 신규 생성
      if (!found && targetSlotId === 0 && p.ticker) {
        slotsSheet.appendRow([
          1, true, `${p.ticker} 실전 전략`, "", p.ticker,
          p.startDate || Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd"),
          Number(p.seedCapital) || 0, Number(p.portions) || 0, Number(p.phaseSplitRound) || 0,
          p.lateMode || "cond", Number(p.lateThresholdPct) || 0, Number(p.targetProfitPct) || 0,
          p.compoundMode || "simple", Number(p.currentCycleNo) || 1, nowStr
        ]);
      }

      // 만약 완료된 사이클 정보가 동봉되어 있다면 추가 기록
      if (p.completedCycles && p.completedCycles.length > 0) {
        let cyclesSheet = ss.getSheetByName("DCA_Completed_Cycles");
        if (cyclesSheet) {
          const lastCycle = p.completedCycles[0];
          cyclesSheet.appendRow([
            lastCycle.cycleNo || 1,
            p.ticker || "",
            lastCycle.startDate || "",
            lastCycle.endDate || "",
            lastCycle.durationDays || 0,
            lastCycle.invested || 0,
            lastCycle.profit || 0,
            lastCycle.retPct || 0
          ]);
        }
      }

      return respondJSON({ success: true, message: "실전 활성 전략이 동기화되었습니다." });
    }

    // 3. 전략 슬롯 삭제 (초기화)
    if (action === "delete_slot" || action === "clear_slot") {
      const slotId = Number(body.slotId) || 0;
      if (slotId > 0) {
        const sData = slotsSheet.getDataRange().getValues();
        for (let i = 1; i < sData.length; i++) {
          if (Number(sData[i][0]) === slotId) {
            slotsSheet.deleteRow(i + 1);
            break;
          }
        }
      }
      return respondJSON({ success: true, message: `슬롯 ${slotId}이 삭제되었습니다.` });
    }

    // 4. GitHub Actions 데이터 업데이트 원격 트리거
    if (action === "trigger_data_update" || action === "trigger_github_update") {
      const res = triggerGitHubDataUpdate();
      return respondJSON(res);
    }

    return respondJSON({ success: false, message: "Unknown action: " + action });

  } catch (err) {
    return respondJSON({ success: false, error: err.toString() });
  }
}

function respondJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatDateVal(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, "GMT+9", "yyyy-MM-dd HH:mm:ss");
  }
  return String(val || "");
}
