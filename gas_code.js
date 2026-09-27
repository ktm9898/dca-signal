/**
 * Google Apps Script - DCA Signal (정량 분할 적립식 퀀트 시그널 시스템)
 * Repository: ktm9898/dca-signal
 * 
 * [배포 방법]
 * 1. 구글 스프레드시트 생성 -> 상단 메뉴 [확장 프로그램] -> [Apps Script] 클릭
 * 2. 기존 코드를 모두 지우고 본 파일 내용을 전체 복사하여 붙여넣기
 * 3. 상단 함수 선택에서 `setupSheets` 선택 후 [실행] 클릭 (초기 시트 탭 3개 자동 생성)
 * 4. 우측 상단 [배포] -> [새 배포] 클릭
 *    - 유형: "웹 앱" 선택
 *    - 설명: "dca-signal api v1"
 *    - 다음 사용자 권한으로 실행: "나 (사용자 계정)"
 *    - 액세스 권한이 있는 사용자: "모든 사용자 (Anyone)" 선택 (중요!)
 * 5. 배포 후 나오는 [웹 앱 URL] (https://script.google.com/macros/s/.../exec)을 복사하여
 *    프론트엔드(index.html, backtest.html)의 GAS_WEBAPP_URL 변수에 붙여넣으시면 됩니다.
 */

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 활성 실전 전략 (현재 대시보드에서 가동 중인 전략)
  let activeSheet = ss.getSheetByName("DCA_Active_Strategy");
  if (!activeSheet) {
    activeSheet = ss.insertSheet("DCA_Active_Strategy");
    activeSheet.getRange("A1:M1").setValues([[
      "Ticker", "StartDate", "SeedCapital", "Portions", "PhaseSplitRound",
      "LateMode", "LateThresholdPct", "TargetProfitPct", "CompoundMode",
      "CurrentRound", "HoldQty", "AvgPrice", "UpdatedAt"
    ]]);
    activeSheet.getRange("A1:M1").setFontWeight("bold").setBackground("#e0f2fe");

    // 기본 초기 실전 전략 1행 삽입
    activeSheet.appendRow([
      "TQQQ", 
      Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd"),
      10000, 40, 20,
      "cond", -5.0, 10.0, "simple",
      1, 0, 0,
      Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm:ss")
    ]);
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

    // 추천 기본 슬롯 등록
    slotsSheet.appendRow([
      1, "TQQQ 정석 40분할", "20회차 후 -5% 이하만 매수, 목표 10%", 
      "TQQQ", 10000, 40, 20, "cond", -5.0, 10.0, "simple",
      Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm:ss")
    ]);
    slotsSheet.appendRow([
      2, "SOXL 공격형 50분할", "25회차 후 -7% 이하 매수, 목표 12%", 
      "SOXL", 10000, 50, 25, "cond", -7.0, 12.0, "simple",
      Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm:ss")
    ]);
  }

  // 3. 매매 일지 & 사이클 기록 (DCA_Trade_Journal)
  let journalSheet = ss.getSheetByName("DCA_Trade_Journal");
  if (!journalSheet) {
    journalSheet = ss.insertSheet("DCA_Trade_Journal");
    journalSheet.getRange("A1:J1").setValues([[
      "Date", "CycleNo", "Round", "Action", "Price", "Qty", "Amount", "AvgPrice", "Profit", "Memo"
    ]]);
    journalSheet.getRange("A1:J1").setFontWeight("bold").setBackground("#dcfce7");
  }
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "all";
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
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
          startDate: formatDateVal(row[1]),
          seedCapital: Number(row[2]) || 10000,
          portions: Number(row[3]) || 40,
          phaseSplitRound: Number(row[4]) || 20,
          lateMode: String(row[5]) || "cond",
          lateThresholdPct: Number(row[6]) || -5,
          targetProfitPct: Number(row[7]) || 10,
          compoundMode: String(row[8]) || "simple",
          currentRound: Number(row[9]) || 1,
          holdQty: Number(row[10]) || 0,
          avgPrice: Number(row[11]) || 0,
          updatedAt: formatDateVal(row[12])
        };
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
            seedCapital: Number(r[4]) || 10000,
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

      // 매매 일지 최근 50건
      let journalSheet = ss.getSheetByName("DCA_Trade_Journal");
      let journal = [];
      if (journalSheet) {
        const jData = journalSheet.getDataRange().getValues();
        const startIdx = Math.max(1, jData.length - 50);
        for (let i = jData.length - 1; i >= startIdx; i--) {
          const r = jData[i];
          if (!r[0]) continue;
          journal.push({
            date: formatDateVal(r[0]),
            cycleNo: Number(r[1]) || 1,
            round: Number(r[2]) || 0,
            action: String(r[3]),
            price: Number(r[4]) || 0,
            qty: Number(r[5]) || 0,
            amount: Number(r[6]) || 0,
            avgPrice: Number(r[7]) || 0,
            profit: Number(r[8]) || 0,
            memo: String(r[9] || "")
          });
        }
      }

      return respondJSON({
        success: true,
        activeStrategy: activeStrat,
        slots: slots,
        journal: journal
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
          seedCapital: Number(r[4]) || 10000,
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
    return respondJSON({ success: false, error: "Invalid JSON post body: " + ex });
  }

  const action = body.action || (e && e.parameter && e.parameter.action) || "";

  try {
    // 1. 실전 활성 전략 변경
    if (action === "set_active_strategy") {
      let activeSheet = ss.getSheetByName("DCA_Active_Strategy");
      if (!activeSheet) {
        setupSheets();
        activeSheet = ss.getSheetByName("DCA_Active_Strategy");
      }

      const p = body.strategy || {};
      const nowStr = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm:ss");

      // 2번째 행 덮어쓰기
      activeSheet.getRange(2, 1, 1, 13).setValues([[
        p.ticker || "TQQQ",
        p.startDate || Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd"),
        p.seedCapital || 10000,
        p.portions || 40,
        p.phaseSplitRound || 20,
        p.lateMode || "cond",
        p.lateThresholdPct != null ? p.lateThresholdPct : -5,
        p.targetProfitPct || 10,
        p.compoundMode || "simple",
        p.currentRound != null ? p.currentRound : 1,
        p.holdQty != null ? p.holdQty : 0,
        p.avgPrice != null ? p.avgPrice : 0,
        nowStr
      ]]);

      return respondJSON({ success: true, message: "실전 전략이 저장되었습니다." });
    }

    // 2. 전략 슬롯 저장/수정
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
        body.name || `전략 슬롯 ${slotId}`,
        body.memo || "",
        p.ticker || "TQQQ",
        p.seedCapital || 10000,
        p.portions || 40,
        p.phaseSplitRound || 20,
        p.lateMode || "cond",
        p.lateThresholdPct != null ? p.lateThresholdPct : -5,
        p.targetProfitPct || 10,
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

    // 3. 실전 매수/익절 진행 상태 업데이트 (회차, 수량, 평단가 갱신 및 일지 기록)
    if (action === "log_trade_action") {
      let activeSheet = ss.getSheetByName("DCA_Active_Strategy");
      let journalSheet = ss.getSheetByName("DCA_Trade_Journal");
      if (!activeSheet || !journalSheet) {
        setupSheets();
        activeSheet = ss.getSheetByName("DCA_Active_Strategy");
        journalSheet = ss.getSheetByName("DCA_Trade_Journal");
      }

      const trade = body.trade || {};
      const nowStr = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm:ss");
      const todayStr = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd");

      // 일지 추가
      journalSheet.appendRow([
        todayStr,
        trade.cycleNo || 1,
        trade.round || 1,
        trade.action || "BUY",
        trade.price || 0,
        trade.qty || 0,
        trade.amount || 0,
        trade.avgPrice || 0,
        trade.profit || 0,
        trade.memo || ""
      ]);

      // 활성 상태 갱신
      if (trade.currentRound != null) activeSheet.getRange(2, 10).setValue(trade.currentRound);
      if (trade.holdQty != null) activeSheet.getRange(2, 11).setValue(trade.holdQty);
      if (trade.avgPrice != null) activeSheet.getRange(2, 12).setValue(trade.avgPrice);
      activeSheet.getRange(2, 13).setValue(nowStr);

      return respondJSON({ success: true, message: "매매 기록 및 상태가 동기화되었습니다." });
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
