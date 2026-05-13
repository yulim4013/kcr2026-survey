// ============================================================
// KCR 2026 만족도 설문 — Apps Script 백엔드
//
// [배포 방법]
// 1. 구글 스프레드시트 새로 만들기 (설문 전용)
// 2. 확장 프로그램 → Apps Script → 이 코드 전체 붙여넣기
// 3. 저장 → 배포 → 새 배포 → 웹 앱 선택
// 4. 설정:
//    - 설명: KCR2026 만족도 설문
//    - 다음 사용자로 실행: 나(본인)
//    - 액세스 권한: 모든 사용자
// 5. 배포 → 승인 → URL 복사
// 6. 복사한 URL을 survey.html의 GAS_URL 변수에 붙여넣기
// ============================================================

const PM_EMAIL      = "info@kcr2026.com";
const SHEET_NAME    = "만족도설문";
const ALERT_EVERY   = 50;  // 이메일 알림 주기 (N건당 1회)
const DASH_EVERY    = 50;  // 대시보드 갱신 주기 (N건당 1회)
const HEADERS    = [
  "타임스탬프", "응답자구분", "언어",
  "소속직종", "참석횟수", "인지경로", "AI친숙도", "통역경험",
  "AI번역인지", "이용수단", "APP편리성",
  "용어정확성", "번역자연스러움", "실시간성", "가독성", "음성번역품질", "음성vs자막도움",
  "스크린번역만족", "효과적방식", "AI지속찬성", "AI개선의견",
  "프로그램구성", "연자전문성", "등록절차", "APP공지", "전시홀", "시설편의성",
  "식사품질", "다과서비스", "FB개선의견",
  "재참석의향", "전반적만족도", "추천의향NPS", "기타건의",
  "이름(추첨)", "연락처(추첨)", "경품유형", "기념품수령"
];

// ── 기존 시트 컬럼 마이그레이션 (누락 컬럼 추가, 에디터에서 수동 실행) ──
function addMissingColumns() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('시트 없음'); return; }

  const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const missing  = HEADERS.filter(h => !existing.includes(h));

  missing.forEach(h => {
    const c = sheet.getLastColumn() + 1;
    sheet.getRange(1, c).setValue(h)
      .setBackground("#185FA5").setFontColor("#FFFFFF").setFontWeight("bold");
    Logger.log('추가: ' + h + ' (열 ' + c + ')');
  });

  const msg = missing.length > 0
    ? '추가된 컬럼: ' + missing.join(', ')
    : '누락 컬럼 없음 - 이미 최신 상태';
  Logger.log(msg);
}

// ── 에디터에서 직접 실행: 대시보드+차트 즉시 갱신 ──────────────
// 사용법: 에디터 상단 함수 선택에서 runDashboard 선택 후 실행
function runDashboard() {
  updateSurveyDashboard_();
  Logger.log('대시보드/차트 갱신 완료');
}

// ── GET 요청: 연결 테스트 ─────────────────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "KCR2026 만족도 설문 API 정상 작동 중" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── POST 요청: 설문 접수 / 기념품 수령 처리 ──────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // 기념품 수령 처리 (해외 전용)
    if (data.type === 'prize_claim') {
      return ContentService
        .createTextOutput(JSON.stringify(claimPrize_(data.ts, data.name)))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const result = saveSurvey_(data);  // { count, ts }

    if (result.count === 1 || result.count % ALERT_EVERY === 0) {
      sendAdminAlert_(data, result.count);
    }

    if (result.count === 1 || result.count % DASH_EVERY === 0) {
      updateSurveyDashboard_();
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", count: result.count, ts: result.ts }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 시트에 저장 (LockService로 동시 쓰기 보호) ───────────────
// 반환값: 저장 후 총 응답 수 (헤더 제외)
function saveSurvey_(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);  // 최대 10초 대기 후 락 획득

  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
        .setBackground("#185FA5").setFontColor("#FFFFFF").setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    const row = [
      data.timestamp    || "",
      data.respondent   || "",
      data.lang         || "",
      data.q1           || "",
      data.q2           || "",
      data.q3           || "",
      data.q4           || "",
      data.q5           || "",
      data.q6           || "",
      data.q7           || "",
      data.q8           || "",
      data.q9_1         || "",
      data.q9_2         || "",
      data.q9_3         || "",
      data.q9_4         || "",
      data.q9_5         || "",
      data.q10          || "",
      data.q11          || "",
      data.q12          || "",
      data.q13          || "",
      data.q14          || "",
      data.q15          || "",
      data.q16          || "",
      data.q17          || "",
      data.q18          || "",
      data.q19          || "",
      data.q20          || "",
      data.q21          || "",
      data.q22          || "",
      data.q23          || "",
      data.q24          || "",
      data.q25          || "",
      data.q26          || "",
      data.q27          || "",
      data.prize_name   || "",
      data.prize_phone  || "",
      data.prize_type   || ""
    ];

    const lastRow = sheet.getLastRow() + 1;
    sheet.getRange(lastRow, 1, 1, row.length).setValues([row]);
    return { count: lastRow - 1, ts: data.timestamp };
  } finally {
    lock.releaseLock();
  }
}

// ── 기념품 수령 처리 (해외 전용) ─────────────────────────────
function claimPrize_(ts, name) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { status: 'error', message: 'Sheet not found' };

    const allData     = sheet.getDataRange().getValues();
    const prizeColIdx = HEADERS.indexOf("기념품수령") + 1;  // 1-based

    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][0]) === String(ts)) {
        if (allData[i][prizeColIdx - 1] === '수령완료') {
          return { status: 'already_claimed' };
        }
        sheet.getRange(i + 1, prizeColIdx).setValue('수령완료');
        return { status: 'ok' };
      }
    }
    return { status: 'not_found' };
  } finally {
    lock.releaseLock();
  }
}

// ── 관리자 알림 이메일 (N건마다 1회) ─────────────────────────
function sendAdminAlert_(data, count) {
  const type    = data.respondent === 'domestic' ? '내국인' : '외국인';
  const subject = `[KCR2026 설문] 누적 ${count}건 도달`;
  const body =
    `설문 응답이 ${count}건에 도달했습니다.\n\n` +
    `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n` +
    `최근 응답자 유형 : ${type}\n` +
    `소속             : ${data.q1 || ""}\n` +
    `전반적 만족도    : ${data.q25 || "-"}\n` +
    `추천 의향(NPS)   : ${data.q26 || "-"}\n` +
    `재참석 의향      : ${data.q24 || ""}\n` +
    `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n` +
    `전체 응답 현황: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;

  GmailApp.sendEmail(PM_EMAIL, subject, body, { name: "KCR2026 만족도 설문" });
}

// ── 대시보드 자동 갱신 ────────────────────────────────────────
function updateSurveyDashboard_() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(SHEET_NAME);
  if (!src) return;

  const data = src.getDataRange().getValues();
  if (data.length < 2) return;

  const headers = data[0];
  const rows    = data.slice(1).filter(r => r[0]);
  const col     = name => headers.indexOf(name);

  let dash = ss.getSheetByName("대시보드");
  if (!dash) dash = ss.insertSheet("대시보드");
  dash.clearContents();
  dash.clearFormats();

  const BLUE  = "#185FA5";
  const LBLUE = "#E6F1FB";
  const GRAY  = "#F9F9F9";

  let r = 1;

  function secTitle(text) {
    dash.getRange(r, 1).setValue(text)
      .setFontSize(12).setFontWeight("bold").setFontColor(BLUE);
    r++;
  }
  function distTable(hdrs, vals) {
    dash.getRange(r, 1, 1, hdrs.length).setValues([hdrs])
      .setBackground(LBLUE).setFontWeight("bold").setHorizontalAlignment("center");
    dash.getRange(r + 1, 1, 1, vals.length).setValues([vals])
      .setHorizontalAlignment("center");
    r += 3;
  }

  // 1. 제목
  dash.getRange(r, 1).setValue("KCR 2026 만족도 설문 대시보드")
    .setFontSize(15).setFontWeight("bold").setFontColor(BLUE);
  r++;
  dash.getRange(r, 1)
    .setValue("최종 업데이트: " + Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm"))
    .setFontSize(10).setFontColor("#8E8E93");
  r += 2;

  // 2. 응답 요약
  secTitle("응답 요약");
  const total    = rows.length;
  const domestic = rows.filter(x => x[col("응답자구분")] === "domestic").length;
  const foreign  = rows.filter(x => x[col("응답자구분")] === "foreign").length;
  const langKo   = rows.filter(x => x[col("언어")] === "ko").length;
  const langEn   = rows.filter(x => x[col("언어")] === "en").length;
  distTable(["총 응답", "내국인", "외국인", "국문(ko)", "영문(en)"],
            [total,    domestic, foreign,  langKo,     langEn]);

  // 3. 소속별 분포
  secTitle("소속별 분포");
  const affs = ["대학교수","전문의","전임의","전공의","간호사","연구원","학생","기업관계자","기타"];
  distTable(affs, affs.map(a => rows.filter(x => x[col("소속직종")] === a).length));

  // 4. 참석횟수(Q2) 분포
  secTitle("KCR 참석 횟수 분포");
  const q2opts = ["처음","2-3회","4-5회","6회 이상"];
  distTable(q2opts, q2opts.map(v => rows.filter(x => x[col("참석횟수")] === v).length));

  // 5. 인지경로(Q3) 분포
  secTitle("행사 인지 경로 분포");
  const q3opts = ["학회홈페이지","이메일","동료권유","기타"];
  distTable(q3opts, q3opts.map(v => rows.filter(x => x[col("인지경로")] === v).length));

  // 6. 통역경험(Q5) 분포
  secTitle("인간 통역 서비스 이용 경험");
  distTable(["예","아니오"],
            [rows.filter(x => x[col("통역경험")] === "예").length,
             rows.filter(x => x[col("통역경험")] === "아니오").length]);

  // 7. 이용수단(Q7) 분포 (복수선택 — 쉼표 구분 문자열 포함 여부로 카운트)
  secTitle("AI 번역 이용 수단 분포 (복수선택)");
  const q7opts = ["스크린자막","APP텍스트","APP음성","기타"];
  distTable(q7opts, q7opts.map(v => rows.filter(x => String(x[col("이용수단")]).includes(v)).length));

  // 8. AI 번역 인지여부(Q6) 분포
  secTitle("AI 번역 자막 인지 여부");
  const q6opts = ["예","아니오","몰랐음"];
  distTable(q6opts, q6opts.map(v => rows.filter(x => x[col("AI번역인지")] === v).length));

  // 9. 스크린번역만족(Q11) 분포
  secTitle("스크린 번역 도입 만족도 분포");
  const q11opts = ["매우만족","만족","보통","불만족","매우불만족"];
  distTable(q11opts, q11opts.map(v => rows.filter(x => x[col("스크린번역만족")] === v).length));

  // 10. 효과적방식(Q12) 분포
  secTitle("가장 효과적인 AI 번역 방식");
  const q12opts = ["스크린자막","전용APP","음성번역","병행사용"];
  distTable(q12opts, q12opts.map(v => rows.filter(x => x[col("효과적방식")] === v).length));

  // 11. 재참석의향(Q24) 분포
  secTitle("내년 재참석 의향 분포");
  const q24opts = ["예","아니오","미정"];
  distTable(q24opts, q24opts.map(v => rows.filter(x => x[col("재참석의향")] === v).length));

  // 12. 전반적만족도(Q25) 분포
  secTitle("전반적 만족도 분포");
  const q25scores = [5,4,3,2,1];
  distTable(["5점(매우만족)","4점","3점","2점","1점(매우불만족)"],
            q25scores.map(v => rows.filter(x => parseInt(x[col("전반적만족도")]) === v).length));

  // 13. 추천의향(Q26) 분포 + 추천자/중립/비추천 분류
  secTitle("추천 의향 분포 (5점 스케일)");
  const q26scores = [5,4,3,2,1];
  const q26counts = q26scores.map(v => rows.filter(x => parseInt(x[col("추천의향NPS")]) === v).length);
  distTable(["5점","4점","3점","2점","1점"], q26counts);
  const promoters  = rows.filter(x => parseInt(x[col("추천의향NPS")]) === 5).length;
  const passives   = rows.filter(x => parseInt(x[col("추천의향NPS")]) === 4).length;
  const detractors = rows.filter(x => [1,2,3].includes(parseInt(x[col("추천의향NPS")]))).length;
  dash.getRange(r, 1).setValue(`추천자(5점): ${promoters}명  /  중립(4점): ${passives}명  /  비추천(1-3점): ${detractors}명`)
    .setFontColor("#8E8E93").setFontSize(10);
  r += 2;

  // 14. 항목별 평균 점수
  secTitle("항목별 평균 점수 (5점 만점)");
  const likertDefs = [
    ["AI친숙도",        "AI 기술 친숙도"],
    ["APP편리성",       "APP 설치·접속 편리성"],
    ["용어정확성",      "의학 용어 정확성"],
    ["번역자연스러움",  "번역 자연스러움"],
    ["실시간성",        "실시간성(딜레이)"],
    ["가독성",          "APP 텍스트 가독성"],
    ["음성번역품질",    "음성 번역 품질"],
    ["음성vs자막도움",  "음성번역 vs 자막 도움도"],
    ["AI지속찬성",      "AI 번역 지속 도입 찬성"],
    ["프로그램구성",    "학술 프로그램 구성·주제"],
    ["연자전문성",      "연자·좌장 전문성"],
    ["등록절차",        "현장 등록·키오스크 절차"],
    ["APP공지",         "APP 프로그램 공지"],
    ["전시홀",          "전시 홀 규모·유익성"],
    ["시설편의성",      "행사장 시설 편의성"],
    ["식사품질",        "식사 품질·양"],
    ["다과서비스",      "커피 브레이크 다과·음료"],
    ["전반적만족도",    "전반적 만족도"],
    ["추천의향NPS",     "추천 의향"]
  ];
  dash.getRange(r, 1, 1, 2).setValues([["항목", "평균"]])
    .setBackground(LBLUE).setFontWeight("bold");
  r++;
  likertDefs.forEach((pair, i) => {
    const ci  = col(pair[0]);
    const vs  = rows.map(x => parseFloat(x[ci])).filter(v => !isNaN(v) && v > 0);
    const avg = vs.length > 0 ? (vs.reduce((a,b) => a+b,0) / vs.length).toFixed(2) : "-";
    dash.getRange(r, 1, 1, 2).setValues([[pair[1], avg]])
      .setBackground(i % 2 === 0 ? GRAY : "#FFFFFF");
    r++;
  });
  r++;

  // 15. 최근 응답 10건
  secTitle(`최근 응답 (${Math.min(10, total)}건)`);
  const listH = ["접수일시","유형","소속","AI친숙도","전반적만족도","추천의향","재참석"];
  dash.getRange(r, 1, 1, listH.length).setValues([listH])
    .setBackground(LBLUE).setFontWeight("bold");
  r++;
  const recent = rows.slice(-10).reverse();
  if (recent.length > 0) {
    const listData = recent.map(x => [
      x[col("타임스탬프")],
      x[col("응답자구분")] === "domestic" ? "내국인" : "외국인",
      x[col("소속직종")], x[col("AI친숙도")],
      x[col("전반적만족도")], x[col("추천의향NPS")], x[col("재참석의향")]
    ]);
    dash.getRange(r, 1, listData.length, listH.length).setValues(listData);
    r += listData.length;
  }
  r++;

  // 16. 경품 응모자 목록 (내국인 추첨)
  const prizeRows = rows.filter(x => x[col("이름(추첨)")] && String(x[col("이름(추첨)")]).trim());
  secTitle(`경품 응모자 (${prizeRows.length}명)`);
  const prizeH = ["이름","연락처","경품유형","소속","접수일시"];
  dash.getRange(r, 1, 1, prizeH.length).setValues([prizeH])
    .setBackground(LBLUE).setFontWeight("bold");
  r++;
  if (prizeRows.length > 0) {
    const prizeData = prizeRows.map(x => [
      x[col("이름(추첨)")], x[col("연락처(추첨)")], x[col("경품유형")],
      x[col("소속직종")], x[col("타임스탬프")]
    ]);
    dash.getRange(r, 1, prizeData.length, prizeH.length).setValues(prizeData);
    r += prizeData.length;
  }
  r++;

  // 17. 기념품 수령 현황 (해외 전용)
  const foreignRows   = rows.filter(x => x[col("응답자구분")] === "foreign");
  const claimedRows   = foreignRows.filter(x => x[col("기념품수령")] === "수령완료");
  const unclaimedRows = foreignRows.filter(x => x[col("기념품수령")] !== "수령완료");
  secTitle(`기념품 수령 현황 (해외 ${foreignRows.length}명)`);
  distTable(["해외 총 응답","수령 완료","미수령"],
            [foreignRows.length, claimedRows.length, unclaimedRows.length]);
  if (foreignRows.length > 0) {
    const claimH = ["이름","수령여부","소속","접수일시"];
    dash.getRange(r, 1, 1, claimH.length).setValues([claimH])
      .setBackground(LBLUE).setFontWeight("bold");
    r++;
    const claimData = foreignRows.map(x => [
      x[col("이름(추첨)")],
      x[col("기념품수령")] === "수령완료" ? "수령완료" : "미수령",
      x[col("소속직종")],
      x[col("타임스탬프")]
    ]);
    dash.getRange(r, 1, claimData.length, claimH.length).setValues(claimData);
    r += claimData.length;
  }

  // 열 폭 조정
  for (let c = 1; c <= 10; c++) dash.setColumnWidth(c, 120);
  dash.autoResizeColumns(1, 10);

  // ── 자유의견 시트 ──────────────────────────────────────────
  let opinSheet = ss.getSheetByName("자유의견");
  if (!opinSheet) opinSheet = ss.insertSheet("자유의견");
  opinSheet.clearContents();
  opinSheet.clearFormats();

  opinSheet.getRange(1, 1).setValue("KCR 2026 - 자유의견 모음")
    .setFontSize(14).setFontWeight("bold").setFontColor(BLUE);
  opinSheet.getRange(2, 1)
    .setValue("최종 업데이트: " + Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm"))
    .setFontSize(10).setFontColor("#8E8E93");

  const opH = ["접수일시","유형","소속","Q14 AI번역 의견","Q23 F&B 개선의견","Q27 기타 건의사항"];
  opinSheet.getRange(4, 1, 1, opH.length).setValues([opH])
    .setBackground(BLUE).setFontColor("#FFFFFF").setFontWeight("bold");

  const opRows = rows.filter(x =>
    (x[col("AI개선의견")] && String(x[col("AI개선의견")]).trim()) ||
    (x[col("FB개선의견")] && String(x[col("FB개선의견")]).trim()) ||
    (x[col("기타건의")]   && String(x[col("기타건의")]).trim())
  );
  if (opRows.length > 0) {
    const opData = opRows.map(x => [
      x[col("타임스탬프")],
      x[col("응답자구분")] === "domestic" ? "내국인" : "외국인",
      x[col("소속직종")],
      x[col("AI개선의견")],
      x[col("FB개선의견")],
      x[col("기타건의")]
    ]);
    opinSheet.getRange(5, 1, opData.length, opH.length).setValues(opData);
  }
  for (let c = 1; c <= opH.length; c++) opinSheet.setColumnWidth(c, 200);
  opinSheet.setFrozenRows(4);

  // 차트 시트 생성
  buildChartSheet_(ss, rows, headers);
}

// ── 통계 차트 시트 생성 ───────────────────────────────────────
function buildChartSheet_(ss, rows, headers) {
  if (rows.length === 0) return;
  const col = name => headers.indexOf(name);

  let cs = ss.getSheetByName("통계차트");
  if (!cs) {
    cs = ss.insertSheet("통계차트");
  } else {
    cs.clearContents();
    cs.clearFormats();
    cs.getCharts().forEach(c => cs.removeChart(c));
  }

  const BLUE  = "#185FA5";
  const LBLUE = "#E6F1FB";
  const GRAY  = "#F9F9F9";
  let r = 1;

  // 섹션 데이터(A-B열) + 차트(D열~) 삽입
  // chartRows: 차트가 차지하는 예상 행 수 (겹침 방지)
  function writeSection(title, labels, values, chartType, chartH) {
    const h         = chartH || 240;
    const chartRows = Math.ceil(h / 21) + 2;
    const startRow  = r;

    cs.getRange(r, 1).setValue(title)
      .setFontWeight("bold").setFontColor(BLUE).setFontSize(11);
    r++;

    cs.getRange(r, 1, 1, 2).setValues([['항목', '건수']])
      .setBackground(LBLUE).setFontWeight("bold");
    r++;

    labels.forEach((lbl, i) => {
      cs.getRange(r, 1, 1, 2).setValues([[lbl, values[i]]])
        .setBackground(i % 2 === 0 ? GRAY : "#FFFFFF");
      r++;
    });

    // 차트: 헤더 + 데이터 범위
    const chart = cs.newChart()
      .setChartType(chartType)
      .addRange(cs.getRange(startRow + 1, 1, labels.length + 1, 2))
      .setPosition(startRow, 4, 0, 0)
      .setOption('title', title)
      .setOption('width', 420)
      .setOption('height', h)
      .setOption('legend', { position: labels.length <= 5 ? 'bottom' : 'none' })
      .build();
    cs.insertChart(chart);

    r = Math.max(r + 1, startRow + chartRows);
  }

  // 1. 내국인 / 외국인
  writeSection("내국인 / 외국인 비율",
    ["내국인", "외국인"],
    [rows.filter(x => x[col("응답자구분")] === "domestic").length,
     rows.filter(x => x[col("응답자구분")] === "foreign").length],
    Charts.ChartType.PIE);

  // 2. 소속별 분포
  const affs = ["대학교수","전문의","전임의","전공의","간호사","연구원","학생","기업관계자","기타"];
  writeSection("소속별 분포", affs,
    affs.map(a => rows.filter(x => x[col("소속직종")] === a).length),
    Charts.ChartType.COLUMN, 280);

  // 3. 참석 횟수
  const q2opts = ["처음","2-3회","4-5회","6회 이상"];
  writeSection("KCR 참석 횟수", q2opts,
    q2opts.map(v => rows.filter(x => x[col("참석횟수")] === v).length),
    Charts.ChartType.PIE);

  // 4. 인지 경로
  const q3opts = ["학회홈페이지","이메일","동료권유","기타"];
  writeSection("행사 인지 경로", q3opts,
    q3opts.map(v => rows.filter(x => x[col("인지경로")] === v).length),
    Charts.ChartType.PIE);

  // 5. 통역 경험
  writeSection("인간 통역 경험", ["예","아니오"],
    [rows.filter(x => x[col("통역경험")] === "예").length,
     rows.filter(x => x[col("통역경험")] === "아니오").length],
    Charts.ChartType.PIE);

  // 6. AI 번역 이용 수단 (복수선택)
  const q7opts = ["스크린자막","APP텍스트","APP음성","기타"];
  writeSection("AI 번역 이용 수단 (복수선택)", q7opts,
    q7opts.map(v => rows.filter(x => String(x[col("이용수단")]).includes(v)).length),
    Charts.ChartType.COLUMN);

  // 7. AI 번역 자막 인지 여부
  const q6opts = ["예","아니오","몰랐음"];
  writeSection("AI 번역 자막 인지 여부", q6opts,
    q6opts.map(v => rows.filter(x => x[col("AI번역인지")] === v).length),
    Charts.ChartType.PIE);

  // 8. 스크린 번역 만족도
  const q11opts = ["매우만족","만족","보통","불만족","매우불만족"];
  writeSection("스크린 번역 도입 만족도", q11opts,
    q11opts.map(v => rows.filter(x => x[col("스크린번역만족")] === v).length),
    Charts.ChartType.COLUMN);

  // 9. 가장 효과적인 방식
  const q12opts = ["스크린자막","전용APP","음성번역","병행사용"];
  writeSection("가장 효과적인 AI 번역 방식", q12opts,
    q12opts.map(v => rows.filter(x => x[col("효과적방식")] === v).length),
    Charts.ChartType.PIE);

  // 10. 재참석 의향
  const q24opts = ["예","아니오","미정"];
  writeSection("내년 재참석 의향", q24opts,
    q24opts.map(v => rows.filter(x => x[col("재참석의향")] === v).length),
    Charts.ChartType.PIE);

  // 11. 전반적 만족도 분포
  writeSection("전반적 만족도 분포",
    ["5점(매우만족)","4점","3점","2점","1점(매우불만족)"],
    [5,4,3,2,1].map(v => rows.filter(x => parseInt(x[col("전반적만족도")]) === v).length),
    Charts.ChartType.COLUMN);

  // 12. 추천 의향 분포
  writeSection("추천 의향 분포",
    ["5점","4점","3점","2점","1점"],
    [5,4,3,2,1].map(v => rows.filter(x => parseInt(x[col("추천의향NPS")]) === v).length),
    Charts.ChartType.COLUMN);

  // 13. 항목별 평균 점수 (가로 막대)
  const likertDefs = [
    ["AI친숙도",       "AI 기술 친숙도"],
    ["APP편리성",      "APP 설치·접속 편리성"],
    ["용어정확성",     "의학 용어 정확성"],
    ["번역자연스러움", "번역 자연스러움"],
    ["실시간성",       "실시간성(딜레이)"],
    ["가독성",         "APP 텍스트 가독성"],
    ["음성번역품질",   "음성 번역 품질"],
    ["음성vs자막도움", "음성번역 도움도"],
    ["AI지속찬성",     "AI 번역 지속 찬성"],
    ["프로그램구성",   "학술 프로그램 구성"],
    ["연자전문성",     "연자·좌장 전문성"],
    ["등록절차",       "현장 등록 절차"],
    ["APP공지",        "APP 프로그램 공지"],
    ["전시홀",         "전시 홀 유익성"],
    ["시설편의성",     "행사장 시설 편의성"],
    ["식사품질",       "식사 품질"],
    ["다과서비스",     "커피 브레이크"],
    ["전반적만족도",   "전반적 만족도"],
    ["추천의향NPS",    "추천 의향"]
  ];
  const avgLabels = likertDefs.map(p => p[1]);
  const avgValues = likertDefs.map(p => {
    const ci = col(p[0]);
    const vs = rows.map(x => parseFloat(x[ci])).filter(v => !isNaN(v) && v > 0);
    return vs.length > 0
      ? parseFloat((vs.reduce((a, b) => a + b, 0) / vs.length).toFixed(2))
      : 0;
  });
  writeSection("항목별 평균 점수 (5점 만점)", avgLabels, avgValues, Charts.ChartType.BAR, 580);

  // 열 폭 조정
  cs.setColumnWidth(1, 210);
  cs.setColumnWidth(2, 70);
  cs.setColumnWidth(3, 20);
}
