/**
 * Google Apps Script - DCA Signal (정량 분할 적립식 퀀트 시그널 시스템)
 * Repository: ktm9898/dca-signal
 * 
 * [배포 및 트리거 설정]
 * 1. 스프레드시트 [확장 프로그램] -> [Apps Script]에서 본 코드로 덮어쓰기
 * 2. 상단 함수 선택에서 `setupSheets` 선택 후 [실행] (시트 탭 3개 자동 구성)
 * 3. 상단 함수 선택에서 `setupDailyTrigger` 선택 후 [실행]
 *    -> 매일 아침 07:30 KST(미국 장 마감 후) 자동으로 최신 시세를 갱신하는 시간 트리거가 자동 등록됩니다!
 * 4. 우측 상단 [배포] -> [배포 관리] -> 연필(수정) 아이콘 클릭 후
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

  // 1. 전략 슬롯 시트 (1~10번 보관함 - IsActive 컬럼으로 활성 실전 전략 표시)
  let slotsSheet = ss.getSheetByName("DCA_Strategy_Slots");
  if (!slotsSheet) {
    slotsSheet = ss.insertSheet("DCA_Strategy_Slots");
    slotsSheet.getRange("A1:O1").setValues([[
      "SlotID", "IsActive", "Name", "Memo", "Ticker", "StartDate", 
      "SeedCapital", "Portions", "PhaseSplitRound", "LateMode", 
      "LateThresholdPct", "TargetProfitPct", "CompoundMode", "CurrentCycleNo", "UpdatedAt"
    ]]);
    slotsSheet.getRange("A1:O1").setFontWeight("bold").setBackground("#dbeafe");
    // 사용자가 직접 저장한 전략만 보관 (초기 예시 행 없음)
  }

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
    // 1. 전략 슬롯 목록 및 활성 전략 로드
    let slotsSheet = ss.getSheetByName("DCA_Strategy_Slots");
    if (!slotsSheet) {
      setupSheets();
      slotsSheet = ss.getSheetByName("DCA_Strategy_Slots");
    }

    const sData = slotsSheet.getDataRange().getValues();
    const slots = [];
    let activeStrat = null;

    for (let i = 1; i < sData.length; i++) {
      const r = sData[i];
      if (!r[0]) continue;
      const isActive = (r[1] === true || String(r[1]).toUpperCase() === "TRUE" || String(r[1]).toUpperCase() === "Y" || r[1] === 1);
      const slotObj = {
        slotId: Number(r[0]),
        isActive: isActive,
        name: String(r[2] || `슬롯 ${r[0]}`),
        memo: String(r[3] || ""),
        ticker: String(r[4] || ""),
        startDate: formatDateVal(r[5]).substring(0, 10),
        seedCapital: Number(r[6]) || 0,
        portions: Number(r[7]) || 0,
        phaseSplitRound: Number(r[8]) || 0,
        lateMode: String(r[9] || "cond"),
        lateThresholdPct: Number(r[10]) || 0,
        targetProfitPct: Number(r[11]) || 0,
        compoundMode: String(r[12] || "simple"),
        currentCycleNo: Number(r[13]) || 1,
        updatedAt: formatDateVal(r[14])
      };

      slots.push(slotObj);
      if (isActive && !activeStrat) {
        activeStrat = slotObj;
      }
    }

    // 만약 활성 슬롯이 없는데 슬롯이 존재한다면 첫 번째 슬롯을 활성 전략으로 간주
    if (!activeStrat && slots.length > 0) {
      activeStrat = slots[0];
    }

    if (action === "get_slots") {
      return respondJSON({ success: true, slots: slots });
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
    let slotsSheet = ss.getSheetByName("DCA_Strategy_Slots");
    if (!slotsSheet) {
      setupSheets();
      slotsSheet = ss.getSheetByName("DCA_Strategy_Slots");
    }

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
        Number(p.targetProfitPct) || 0,
        p.compoundMode || "simple",
        Number(p.currentCycleNo) || 1,
        nowStr
      ];

      if (targetRow > 0) {
        slotsSheet.getRange(targetRow, 1, 1, 15).setValues([rowValues]);
      } else {
        slotsSheet.appendRow(rowValues);
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
