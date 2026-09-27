/**
 * Google Apps Script - DCA Signal (정량 분할 적립식 퀀트 시그널 시스템)
 * Repository: ktm9898/dca-signal
 * 
 * [배포 및 트리거 설정]
 * 1. 스프레드시트 [확장 프로그램] -> [Apps Script]에서 본 코드로 덮어쓰기
 * 2. 상단 함수 선택에서 `setupSheets` 선택 후 [실행] (시트 탭 3개 생성)
 * 3. 상단 함수 선택에서 `setupDailyTrigger` 선택 후 [실행]
 *    -> 매일 아침 06:30 KST(미국 장 마감 후) 자동으로 최신 시세를 갱신하는 시간 트리거가 자동 등록됩니다!
 * 4. 우측 상단 [배포] -> [배포 관리] -> 연필(수정) 아이콘 클릭 후
 *    - 버전: "새 버전" 선택
 *    - [배포] 클릭 (기존 웹앱 URL 유지)
 */

const TARGET_TICKERS = ["TQQQ", "SOXL", "UPRO", "QLD", "QQQ", "SPY", "122630.KS"];

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 활성 실전 전략 (현재 대시보드에서 가동 중인 전략)
  let activeSheet = ss.getSheetByName("DCA_Active_Strategy");
  if (!activeSheet) {
    activeSheet = ss.insertSheet("DCA_Active_Strategy");
    activeSheet.getRange("A1:K1").setValues([[
      "Ticker", "StartDate", "SeedCapital", "Portions", "PhaseSplitRound",
      "LateMode", "LateThresholdPct", "TargetProfitPct", "CompoundMode",
      "CurrentCycleNo", "UpdatedAt"
    ]]);
    activeSheet.getRange("A1:K1").setFontWeight("bold").setBackground("#e0f2fe");
  }

  // 2. 전략 슬롯 (1~10번 보관함)
  let slotsSheet = ss.getSheetByName("DCA_Strategy_Slots");
  if (!slotsSheet) {
    slotsSheet = ss.insertSheet("DCA_Strategy_Slots");
    slotsSheet.getRange("A1:L1").setValues([[
      "SlotID", "Name", "Memo", "Ticker", "SeedCapital", "Portions", 
      "PhaseSplitRound", "LateMode", "LateThresholdPct", "TargetProfitPct", 
      "CompoundMode", "UpdatedAt"
    ]]);
    slotsSheet.getRange("A1:L1").setFontWeight("bold").setBackground("#dbeafe");
    // 사용자가 직접 저장한 전략만 보관 (초기 예시 행 없음)
  }

  // 3. 최신 시세 캐시 시트 (DCA_Latest_Quotes)
  let quoteSheet = ss.getSheetByName("DCA_Latest_Quotes");
  if (!quoteSheet) {
    quoteSheet = ss.insertSheet("DCA_Latest_Quotes");
    quoteSheet.getRange("A1:G1").setValues([[
      "Ticker", "ClosePrice", "ChangePct", "High52", "Low52", "Date", "UpdatedAt"
    ]]);
    quoteSheet.getRange("A1:G1").setFontWeight("bold").setBackground("#fef3c7");
  }

  // 4. 완료 사이클 이력 시트 (DCA_Completed_Cycles)
  let cyclesSheet = ss.getSheetByName("DCA_Completed_Cycles");
  if (!cyclesSheet) {
    cyclesSheet = ss.insertSheet("DCA_Completed_Cycles");
    cyclesSheet.getRange("A1:H1").setValues([[
      "CycleNo", "Ticker", "StartDate", "EndDate", "DurationDays", "Invested", "Profit", "ReturnPct"
    ]]);
    cyclesSheet.getRange("A1:H1").setFontWeight("bold").setBackground("#dcfce7");
  }
}

/**
 * ⏰ 매일 아침 07:00~08:00 KST 자동 실행 시간 트리거 등록 함수
 */
function setupDailyTrigger() {
  // 기존 중복 트리거 삭제
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "updateDailyQuotes") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 매일 오전 7시~8시 사이에 실행되는 일일 타이머 등록
  ScriptApp.newTrigger("updateDailyQuotes")
    .timeBased()
    .atHour(7)
    .nearMinute(30)
    .everyDays(1)
    .inTimezone("Asia/Seoul")
    .create();

  Logger.log("매일 아침 07:30 시세 자동 업데이트 트리거 등록 완료!");
  updateDailyQuotes(); // 즉시 1회 실행
}

/**
 * 주요 7개 ETF의 최신 시세를 야후 파이낸스에서 가져와 구글 시트에 업데이트
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

  for (let i = 0; i < TARGET_TICKERS.length; i++) {
    const ticker = TARGET_TICKERS[i];
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
 * Yahoo Finance API로부터 0.2초 초고속 시세 조회
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
    // 1. 대시보드 전체 데이터 일괄 로드 (활성 전략 + 시트에 저장된 종가 + 슬롯 + 이력)
    if (action === "get_active_strategy" || action === "all") {
      let activeSheet = ss.getSheetByName("DCA_Active_Strategy");
      if (!activeSheet) {
        setupSheets();
        activeSheet = ss.getSheetByName("DCA_Active_Strategy");
      }
      const data = activeSheet.getDataRange().getValues();
      let activeStrat = null;
      if (data.length > 1) {
        const row = data[1];
        activeStrat = {
          ticker: String(row[0]),
          startDate: formatDateVal(row[1]).substring(0, 10),
          seedCapital: Number(row[2]) || 10000000,
          portions: Number(row[3]) || 40,
          phaseSplitRound: Number(row[4]) || 20,
          lateMode: String(row[5]) || "cond",
          lateThresholdPct: Number(row[6]) || -5,
          targetProfitPct: Number(row[7]) || 10,
          compoundMode: String(row[8]) || "simple",
          currentCycleNo: Number(row[9]) || 1,
          updatedAt: formatDateVal(row[10])
        };
      }

      // 시트(DCA_Latest_Quotes)에 캐시된 최신 종가 읽기 (온디맨드 실시간 조회 불필요)
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

      // 슬롯 목록
      let slotsSheet = ss.getSheetByName("DCA_Strategy_Slots");
      let slots = [];
      if (slotsSheet) {
        const sData = slotsSheet.getDataRange().getValues();
        for (let i = 1; i < sData.length; i++) {
          const r = sData[i];
          if (!r[0]) continue;
          slots.push({
            slotId: Number(r[0]),
            name: String(r[1]),
            memo: String(r[2]),
            ticker: String(r[3]),
            seedCapital: Number(r[4]) || 10000000,
            portions: Number(r[5]) || 40,
            phaseSplitRound: Number(r[6]) || 20,
            lateMode: String(r[7]) || "cond",
            lateThresholdPct: Number(r[8]) || -5,
            targetProfitPct: Number(r[9]) || 10,
            compoundMode: String(r[10]) || "simple",
            updatedAt: formatDateVal(r[11])
          });
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
    }

    if (action === "get_slots") {
      let slotsSheet = ss.getSheetByName("DCA_Strategy_Slots");
      if (!slotsSheet) return respondJSON({ success: true, slots: [] });
      const sData = slotsSheet.getDataRange().getValues();
      let slots = [];
      for (let i = 1; i < sData.length; i++) {
        const r = sData[i];
        if (!r[0]) continue;
        slots.push({
          slotId: Number(r[0]),
          name: String(r[1]),
          memo: String(r[2]),
          ticker: String(r[3]),
          seedCapital: Number(r[4]) || 10000000,
          portions: Number(r[5]) || 40,
          phaseSplitRound: Number(r[6]) || 20,
          lateMode: String(r[7]) || "cond",
          lateThresholdPct: Number(r[8]) || -5,
          targetProfitPct: Number(r[9]) || 10,
          compoundMode: String(r[10]) || "simple",
          updatedAt: formatDateVal(r[11])
        });
      }
      return respondJSON({ success: true, slots: slots });
    }

    return respondJSON({ success: false, message: "Unknown action: " + action });

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
    // 1. 실전 활성 전략 및 투자개시일 설정/동기화
    if (action === "set_active_strategy") {
      let activeSheet = ss.getSheetByName("DCA_Active_Strategy");
      if (!activeSheet) {
        setupSheets();
        activeSheet = ss.getSheetByName("DCA_Active_Strategy");
      }

      const p = body.strategy || {};
      const nowStr = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm:ss");

      activeSheet.getRange(2, 1, 1, 11).setValues([[
        p.ticker || "",
        p.startDate || Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd"),
        Number(p.seedCapital) || 0,
        Number(p.portions) || 0,
        Number(p.phaseSplitRound) || 0,
        p.lateMode || "cond",
        p.lateThresholdPct != null ? Number(p.lateThresholdPct) : 0,
        Number(p.targetProfitPct) || 0,
        p.compoundMode || "simple",
        Number(p.currentCycleNo) || 1,
        nowStr
      ]]);

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

      return respondJSON({ success: true, message: "실전 전략 및 개시일이 구글 시트에 저장되었습니다." });
    }

    // 2. 전략 슬롯 저장
    if (action === "save_slot") {
      let slotsSheet = ss.getSheetByName("DCA_Strategy_Slots");
      if (!slotsSheet) {
        setupSheets();
        slotsSheet = ss.getSheetByName("DCA_Strategy_Slots");
      }

      const slotId = Number(body.slotId) || 1;
      const p = body.strategy || {};
      const nowStr = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm:ss");

      const sData = slotsSheet.getDataRange().getValues();
      let targetRow = -1;
      for (let i = 1; i < sData.length; i++) {
        if (Number(sData[i][0]) === slotId) {
          targetRow = i + 1;
          break;
        }
      }

      const rowValues = [
        slotId,
        body.name || `슬롯 ${slotId}`,
        body.memo || "",
        p.ticker || "",
        Number(p.seedCapital) || 0,
        Number(p.portions) || 0,
        Number(p.phaseSplitRound) || 0,
        p.lateMode || "cond",
        p.lateThresholdPct != null ? Number(p.lateThresholdPct) : 0,
        Number(p.targetProfitPct) || 0,
        p.compoundMode || "simple",
        nowStr
      ];

      if (targetRow > 0) {
        slotsSheet.getRange(targetRow, 1, 1, 12).setValues([rowValues]);
      } else {
        slotsSheet.appendRow(rowValues);
      }

      return respondJSON({ success: true, message: `슬롯 ${slotId}에 전략이 저장되었습니다.` });
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
