/**
 * 우리집 통합 가계부 - 서버 코드 (Google Apps Script)
 *
 * 이 파일은 구글 시트 '우리집 통합 가계부'에 붙어서 실행됩니다.
 * 시트 메뉴: 확장 프로그램 > Apps Script 에 붙여넣으세요.
 */

var SHEET_TX = '거래';
var SHEET_RULE = '규칙';
var SHEET_CAT = '카테고리';
var SHEET_SET = '설정';
var SHEET_CONSULT = '컨설팅';
var NUM_COLS = 13;

// 거래 탭 열 위치 (0부터 셈)
//  origPlace(원래 사용처): 제목을 고치면 카카오페이에 적힌 원래 이름을 여기 남깁니다.
//                         중복 확인·규칙·자동 분류는 이 원래 이름으로 합니다.
//  deleted(삭제): '삭제'면 화면과 합계에서 빠집니다. 줄은 지우지 않아서
//                 같은 엑셀을 다시 올려도 되살아나지 않고, 언제든 복원할 수 있습니다.
var COL = {
  id: 0, datetime: 1, place: 2, amount: 3, status: 4, owner: 5,
  category: 6, kind: 7, memo: 8, editor: 9, editedAt: 10, origPlace: 11, deleted: 12
};
var TX_HEADER = ['거래ID', '날짜시간', '사용처', '금액', '상태', '주인', '카테고리', '구분', '메모', '수정한 사람', '수정 시각', '원래 사용처', '삭제'];

/** 웹앱 주소로 접속하면 화면을 띄웁니다. */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('우리집 가계부')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ───────────────────────── 화면에서 부르는 함수 ─────────────────────────

/** 처음 열 때: 월별 요약 + 일별 지출 + 카테고리 목록 + 표시 이름 */
function getInitialData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var rows = readTx_(ss).filter(function (v) { return !isDeleted_(v); }).map(function (v) {
    var day = dayOf_(v[COL.datetime], tz);
    return {
      month: day.slice(0, 7),
      day: day,
      owner: v[COL.owner],
      amount: Number(v[COL.amount]) || 0,
      category: v[COL.category],
      kind: v[COL.kind]
    };
  });
  var summary = summarize(rows);
  summary.categories = readCategories_(ss);
  var set = readSettings_(ss);
  summary.names = {
    husband: set['남편 표시 이름'] || '남편',
    wife: set['아내 표시 이름'] || '아내'
  };
  return summary;
}

/**
 * 컨설팅 화면: '컨설팅' 탭에 쌓인 리포트 (최신이 위)
 * 탭 열: 작성일, 대상 기간, 제목, 내용
 */
function getConsultReports() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CONSULT);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var tz = ss.getSpreadsheetTimeZone();
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    .filter(function (r) { return r[2] || r[3]; })
    .map(function (r) {
      return {
        date: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') : String(r[0] || ''),
        period: String(r[1] || ''),
        title: String(r[2] || ''),
        body: String(r[3] || '')
      };
    })
    .sort(function (a, b) { return a.date < b.date ? 1 : (a.date > b.date ? -1 : 0); });
}

/** 거래 내역 화면: 한 달치 거래 목록 */
function getTransactions(month) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var out = [];
  readTx_(ss).forEach(function (v) {
    var d = v[COL.datetime];
    if (dayOf_(d, tz).slice(0, 7) !== month) return;
    out.push({
      id: String(v[COL.id]),
      dt: (d instanceof Date) ? Utilities.formatDate(d, tz, 'yyyy-MM-dd HH:mm') : String(d).slice(0, 16),
      place: String(v[COL.place]),
      amount: Number(v[COL.amount]) || 0,
      status: String(v[COL.status] || ''),
      owner: String(v[COL.owner]),
      category: String(v[COL.category]),
      kind: String(v[COL.kind]),
      memo: String(v[COL.memo] || ''),
      editor: String(v[COL.editor] || ''),
      editedAt: v[COL.editedAt] instanceof Date
        ? Utilities.formatDate(v[COL.editedAt], tz, 'yyyy-MM-dd HH:mm')
        : String(v[COL.editedAt] || ''),
      origPlace: String(v[COL.origPlace] || ''),
      deleted: isDeleted_(v)
    });
  });
  out.sort(function (a, b) { return a.dt < b.dt ? 1 : (a.dt > b.dt ? -1 : 0); });
  return out;
}

/**
 * 거래 한 건의 카테고리와 메모를 고칩니다.
 * applyRule이 참이면:
 *  - 규칙 탭에 '이 사용처 = 이 카테고리'를 저장하고 (다음에 올리는 엑셀에 적용)
 *  - 이미 있는 같은 사용처 거래도 모두 같은 카테고리로 바꿉니다 (결제취소 제외).
 * newPlace를 주면 제목(사용처)도 바꿉니다. 카카오페이의 원래 이름은 '원래 사용처' 열에 남겨서
 * 같은 엑셀을 다시 올려도 중복으로 들어오지 않게 합니다.
 * 돌려주는 값: { changed: 바뀐 거래 수, kind: 새 구분 }
 */
function updateTransaction(id, category, memo, applyRule, newPlace) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ensureTxCols_(ss);
    var catKinds = categoryKindMap_(ss);
    if (!catKinds[category]) throw new Error('없는 카테고리입니다: ' + category);

    var values = readAllTx_(sheet);
    var editor = currentUser_();
    var now = new Date();

    var target = findRow_(values, id);
    var row = values[target];
    var place = placeKey_(row); // 규칙·같은 사용처 찾기는 카카오페이 원래 이름으로
    var changed = 0;

    if (newPlace !== undefined && newPlace !== null) {
      var np = String(newPlace).trim();
      if (!np) throw new Error('제목을 비워 둘 수 없어요.');
      if (np !== String(row[COL.place])) {
        if (!row[COL.origPlace]) row[COL.origPlace] = String(row[COL.place]);
        row[COL.place] = np;
        if (np === String(row[COL.origPlace])) row[COL.origPlace] = ''; // 원래 이름으로 되돌린 경우
        row[COL.editor] = editor;
        row[COL.editedAt] = now;
      }
    }

    function setCat(r) {
      if (r[COL.status] === '취소') return false;
      var k = kindFor(catKinds[category], Number(r[COL.amount]) || 0);
      if (r[COL.category] === category && r[COL.kind] === k) return false;
      r[COL.category] = category;
      r[COL.kind] = k;
      r[COL.editor] = editor;
      r[COL.editedAt] = now;
      return true;
    }

    if (setCat(row)) changed++;
    if (String(row[COL.memo] || '') !== String(memo || '')) {
      row[COL.memo] = memo || '';
      row[COL.editor] = editor;
      row[COL.editedAt] = now;
    }

    if (applyRule) {
      for (var j = 0; j < values.length; j++) {
        if (j !== target && !isDeleted_(values[j]) && placeKey_(values[j]) === place && setCat(values[j])) changed++;
      }
      saveRule_(ss, place, category, editor, now);
      writeRows_(sheet, values, 0, values.length);
    } else {
      writeRows_(sheet, values, target, 1);
    }
    return { changed: changed, kind: row[COL.kind], editor: editor };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 거래를 삭제하거나(restore=false) 되살립니다(restore=true).
 * 줄을 지우지 않고 '삭제' 표시만 해서, 같은 엑셀을 다시 올려도 되살아나지 않습니다.
 */
function deleteTransaction(id, restore) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ensureTxCols_(ss);
    var values = readAllTx_(sheet);
    var target = findRow_(values, id);
    var row = values[target];
    row[COL.deleted] = restore ? '' : '삭제';
    row[COL.editor] = currentUser_();
    row[COL.editedAt] = new Date();
    writeRows_(sheet, values, target, 1);
    return { deleted: !restore };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 카카오페이 엑셀에서 읽은 거래를 거래 탭에 추가합니다.
 * owner: '남편' 또는 '아내'
 * items: [{ dt: 'yyyy-MM-dd HH:mm:ss', place, amount, status }]
 *
 * 중복 판단: 주인 + 날짜시간(초까지) + 사용처 + 금액이 같으면 같은 거래로 봅니다.
 *  같은 거래가 시트에 a건, 올린 파일에 b건 있으면 b - a건만 새로 넣습니다.
 *  (같은 시각에 같은 금액이 여러 번 찍힌 포인트 적립 같은 거래도 지키기 위해서입니다.)
 * 카테고리: 규칙 탭 → 예전에 같은 사용처에 쓴 카테고리 → 부부 이름 → 기본 키워드
 *          → 설정 탭 이체 키워드 → 지인 송금 → 기타(미분류)
 */
function addTransactions(owner, items) {
  if (owner !== '남편' && owner !== '아내') throw new Error('주인을 골라 주세요.');
  if (!items || !items.length) return { added: 0, skipped: 0, unclassified: 0, byCategory: {}, range: null };

  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = ss.getSpreadsheetTimeZone();
    var sheet = ensureTxCols_(ss);
    var existing = readTx_(ss);

    var have = {};
    var ids = {};
    existing.forEach(function (v) {
      ids[String(v[COL.id])] = true;
      if (v[COL.owner] !== owner) return;
      // 제목을 고쳤거나 삭제한 거래도 원래 이름으로 세서 다시 들어오지 않게 함
      var k = txKey(owner, secOf_(v[COL.datetime], tz), placeKey_(v), v[COL.amount]);
      have[k] = (have[k] || 0) + 1;
    });

    var ctx = classifyContext_(ss, existing);
    var catKinds = categoryKindMap_(ss);
    var prefix = owner === '남편' ? 'H' : 'W';

    var seen = {};
    var newRows = [];
    var byCategory = {};
    var unclassified = 0;
    var minDt = null, maxDt = null;

    items.forEach(function (it) {
      var dt = normalizeDt(it.dt);
      if (!dt) return;
      var place = String(it.place || '').trim();
      var amount = Number(it.amount) || 0;
      var status = String(it.status || '').trim();
      var k = txKey(owner, dt, place, amount);
      seen[k] = (seen[k] || 0) + 1;
      if (seen[k] <= (have[k] || 0)) return; // 이미 있는 거래

      var category = classifyTx(place, amount, status, ctx);
      if (!catKinds[category]) category = '기타(미분류)';
      var kind = kindFor(catKinds[category], amount);
      if (category === '기타(미분류)') unclassified++;
      byCategory[category] = (byCategory[category] || 0) + 1;

      var base = prefix + '-' + dt.replace(/-/g, '').replace(/:/g, '').replace(' ', '-') + '-' + Math.abs(amount);
      var seq = 1;
      while (ids[base + '-' + seq]) seq++;
      var id = base + '-' + seq;
      ids[id] = true;

      if (!minDt || dt < minDt) minDt = dt;
      if (!maxDt || dt > maxDt) maxDt = dt;
      // 날짜시간은 글자로 넣으면 시트가 날짜로 바꿔 줍니다 (시트 시간대 기준).
      newRows.push([id, dt, place, amount, status === '취소' ? '취소' : '', owner, category, kind, '', '', '', '', '']);
    });

    if (newRows.length) {
      var start = sheet.getLastRow() + 1;
      sheet.getRange(start, 1, newRows.length, NUM_COLS).setValues(newRows);
      sheet.getRange(2, 1, sheet.getLastRow() - 1, NUM_COLS).sort([{ column: COL.datetime + 1, ascending: true }]);
    }
    return {
      added: newRows.length,
      skipped: items.length - newRows.length,
      unclassified: unclassified,
      byCategory: byCategory,
      range: minDt ? [minDt.slice(0, 10), maxDt.slice(0, 10)] : null
    };
  } finally {
    lock.releaseLock();
  }
}

// ───────────────────────── 계산 (컴퓨터에서도 검증) ─────────────────────────

/**
 * 카테고리의 구분과 금액으로 거래의 구분을 정합니다.
 * '수입/지출' 카테고리(예: 지인 송금·입금)는 들어온 돈이면 수입, 나간 돈이면 지출입니다.
 */
function kindFor(catKind, amount) {
  if (catKind === '수입/지출') return amount > 0 ? '수입' : '지출';
  return catKind;
}

/** 날짜시간 글자를 'yyyy-MM-dd HH:mm:ss'로 맞춥니다. 못 읽으면 ''. */
function normalizeDt(s) {
  var m = String(s || '').match(/(\d{4})[-.\/](\d{1,2})[-.\/](\d{1,2})(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (!m) return '';
  function p(x) { x = String(x || '0'); return x.length < 2 ? '0' + x : x; }
  return m[1] + '-' + p(m[2]) + '-' + p(m[3]) + ' ' + p(m[4]) + ':' + p(m[5]) + ':' + p(m[6]);
}

function txKey(owner, dt, place, amount) {
  return owner + '|' + dt + '|' + String(place).trim() + '|' + Number(amount);
}

/** 기본 키워드 규칙. 위에서부터 처음 맞는 것을 씁니다. '+'로 시작하면 들어온 돈일 때만. */
var DEFAULT_RULES = [
  ['내계좌이체·저축', /^토스머니$|^238\.|토뱅 구효은|미래오픈구효은|SOL모임통장|^비상금$|^생활비$|^공과금$|토스페이머니|중개형ISA|^토스241$|하나멤버스|자유저축예탁금|토스퀵계좌이체|업비트|_모임$|→ ?내 |^내 .+ →|NH주거래/],
  ['카드대금', /^신한카드$|^롯데카드$|^씨티비씨카드$|^삼성카드$|→ ?신한카드|카드값/],
  ['+수입(급여)', /^급여$|BORGWARNER|국군재정단|특정업무비|여비|당직비/],
  ['+이자·포인트·환급', /이자|현금IC캐시백|^\d\d\.\d\d~\d\d\.\d\d$|포인트|캐시백|환불|환급|^네이버페이$|BallVic|우아한형제|적립|예탁금 수익/],
  ['대출상환', /대출|포켓론/],
  ['세금·과태료', /과태료|자동차세|지방세|지자체세입|세외수입_/],
  ['주거·통신', /아파트관리비|^가스|도시가스|주택도시보증공사|LGU|유플러스|LG전자 ?구독료|스마트로-세외수입|SK텔레콤|KT통신|olleh/],
  ['보험', /KB손|ＫＢ손|DB손|한화생명|^현대\d\d ?\d{3}$|^현대0\d|^흥국\d|교보2\d|삼성생명|삼성화재|보험/],
  ['육아·보육', /보육료|이화영아원|키즈노트|아가방|베이비배|바니키즈|우리두리앤|무스텔라|포프베베|두리 공식스토어|압소바|키스래빗|미니토이|노시부|맘맘|오가닉맘|토브옷장|쁘리마쥬|메리투투|코링코|만토바|유로덕|슈슈앤크라|슈슈봉봉|마롤로뜨|프롬랩스|베르블랑|누블|좀비베리어|윌비|midmaly|굿 바이즈|간지나라|메이드바이발레리나|삐에아떼르발레|문승련 한복|올그린스|리헤라피|BRANDEN|브랜든|보얼부티크|부띠크엠|천수인샵|자파텐|위틀스토어|키즈|몬테소리/],
  ['의료·건강', /약국|의원|이비인후과|한의원|소아청소년과|똑닥|듀오락|안국건강|콤비타|건강생활MALL|한독몰|눈 세척기|라라메디|닥터노아|유트러스|씨클|병원|크로스핏|피트니스|헬스/],
  ['장보기', /^컬리$|코스트코(?!연회비)|이마트 |롯데마트|한살림|더프레시|대한민국농수산|풍년축산|로컬푸드|통 수산|미소과메기|해브어굿팜|대한유통|네슬레|농업회사법인|올데이 올웨이즈|FRANKLiiN|에스엘|광동생활건강|LG생활건강|애경|크린포유|비누두레|폰뿌|윈홈스|국군복지단|하나로마트|홈플러스/],
  ['외식·배달', /배달의민족|요기요|쿠팡이츠|카페|커피|칼국수|스시|떡볶이|파리바게뜨|뚜레쥬르|버거킹|맥도날드|롯데리아|김밥|스타벅스|이디야|엔제리너스|팀홀튼|까사디루카|해장국|우럭|염소탕|산아래|카츠오도|제면소|비빔국수|짬뽕|모닥모닥|샤오마라|^카카오·|매머드익스프레스|지예네|희망분식|요아정|아이스 무빙|풀무원푸드|제이디베이크|이삭토스트|^강남$|나실패밀리|벤치포레스트|유월가게|백금당|나인블럭|베이커리|치킨|피자|POPEYES|우아한형제들·|빽다방|롤링핀|메가엠지씨|컴포즈|국수|분식/],
  ['편의점·자판기', /GS25|GS\)25|지에스25|씨유|CU\)|코리아세븐|세븐일레븐|이마트24|미니스톱|자판기|KIOSK/],
  ['교통·차량', /^SR$|SRT|에스알|철도|코레일|KORAIL|환불수수료|시내버스|시외버스|고속버스|경기버스|티머니버스|^버스|\d+월.*버스|지하철|티머니|경기마을|주유소|오일뱅크|칼텍스|모터스|행운열차|카카오모빌리티|카카오T|택시|이동의즐거움|블루파트너스|도로공사|도로|주차|쏘카|지바이크|하이패스/],
  ['여행·여가', /Agoda|아고다|여기어때|야놀자|진에어|JINAIR|웹투어|호텔|리조트|콘래드|놀유니버스|제주|에버랜드|서울랜드|미술관|박물관|스타필드|노래|CGV|메가박스|롯데시네마|문화비|디스트릭트|아트스토리|아쿠아 스튜디오|플라랜드|풀문랜드|포토이즘|어뮤즈먼트|테마파크/],
  ['미용·뷰티', /올리브영|헤어|미용실|케라스타즈|시세이도|로라 메르시에|에뛰드|투쿨포스쿨|한율|아비브|여신코스메틱|데자뷰|스컬피그|래쉬|^미브$|^\(주\)라엘|안경|1초재팬|티움|agazzang|노엘로힐스|샌더소니아|오더렉트|더마초이스/],
  ['쇼핑·의류·생활', /쿠팡|코오롱몰|나이키|유니클로|비너스|비비안|아울렛|백화점|갤러리아|Hmall|롯데홈쇼핑|29CM|무신사|A-RT|에이케이플라자|뉴코아|무인양품|다이소|아이디어스|롯데쇼핑|가든파이브|당근페이|문지기닷컴|크린토피아|툴베이스|무료배송오늘출발|메이크모던|쿨타임|맨크루|벨리타앤보이즈|피그먼트|스미스앤스미시|^스마틱스$|오산드빛|청운|LG전자$|11번가|G마켓|옥션|SSG/],
  ['구독·디지털', /^Google|구글|넷플릭스|디즈니|유튜브|프리미엄콘텐츠|네이버플러스멤버십|^시리즈$|컬리\(멤버스\)|코스트코연회비|자동결제|Apple|애플|쇼핑케어|토스프라임/],
  ['도서', /교보문고|알라딘|예스24|YES24/],
  ['온라인결제(가맹점미상)', /NHNKCP|이니시스|^\(주\)스마트로$|NICE|토스페이먼츠|KCP|페이앱|엔에이치엔링크|케이지파이낸셜|케이지모빌리언스|인터넷상거래|전자상거래|바로알림서비스/]
];
// 지인 송금: '홍길동(홍*동)' 모양, 카카오페이머니 송금, 또는 이름만 있는 입금
var FRIEND_RE = /\(.\*.\)|카카오페이머니 →|송금 →/;
var NAME_ONLY_RE = /^[가-힣]{2,4}$/;

/**
 * 사용처와 금액으로 카테고리를 정합니다.
 * ctx = { rules: {사용처: 카테고리}, history: {'사용처|+' 또는 '사용처|-': 카테고리},
 *         spouseRe: 부부 이름 정규식, transferRe: 설정 탭 이체 키워드 }
 */
function classifyTx(place, amount, status, ctx) {
  if (status === '취소') return '결제취소';
  if (ctx.rules && ctx.rules[place]) return ctx.rules[place];
  var hk = place + (amount > 0 ? '|+' : '|-');
  if (ctx.history && ctx.history[hk]) return ctx.history[hk];
  if (ctx.spouseRe && ctx.spouseRe.test(place)) return '내계좌이체·저축';
  for (var i = 0; i < DEFAULT_RULES.length; i++) {
    var cat = DEFAULT_RULES[i][0];
    var incomeOnly = cat.charAt(0) === '+';
    if (incomeOnly && amount < 0) continue;
    if (DEFAULT_RULES[i][1].test(place)) return incomeOnly ? cat.slice(1) : cat;
  }
  if (ctx.transferRe && ctx.transferRe.test(place)) return '내계좌이체·저축';
  if (FRIEND_RE.test(place)) return '지인 송금·입금';
  if (amount > 0 && NAME_ONLY_RE.test(place)) return '지인 송금·입금';
  return '기타(미분류)';
}

/**
 * 거래 목록을 월별로 묶어 합계를 냅니다.
 * - 지출: 구분이 '지출'인 거래. 환불(+)은 같은 카테고리에서 빠집니다. 쓴 돈을 양수로 돌려줍니다.
 * - 수입: 구분이 '수입'인 거래.
 * - 이체: 구분이 '이체'인 거래 중 나간 돈만 따로 모읍니다.
 * - 취소는 넣지 않습니다.
 * - days: 날짜별 지출 { 'yyyy-MM-dd': { husband, wife, count } }
 */
function summarize(rows) {
  var months = {};
  function blank() { return { total: 0, husband: 0, wife: 0, count: 0, cats: {} }; }
  function bucket(m) {
    if (!months[m]) {
      months[m] = { expense: blank(), income: blank(), transfer: { husband: 0, wife: 0, count: 0 }, days: {} };
    }
    return months[m];
  }
  function add(t, who, cat, v) {
    t.total += v; t[who] += v; t.count += 1;
    var c = t.cats[cat] || (t.cats[cat] = { husband: 0, wife: 0 });
    c[who] += v;
  }

  rows.forEach(function (r) {
    var who = r.owner === '남편' ? 'husband' : (r.owner === '아내' ? 'wife' : null);
    if (!who) return;
    var b = bucket(r.month);
    if (r.kind === '지출') {
      add(b.expense, who, r.category, -r.amount);
      if (r.day) {
        var d = b.days[r.day] || (b.days[r.day] = { husband: 0, wife: 0, count: 0 });
        d[who] += -r.amount;
        d.count += 1;
      }
    } else if (r.kind === '수입') add(b.income, who, r.category, r.amount);
    else if (r.kind === '이체' && r.amount < 0) {
      b.transfer[who] += -r.amount;
      b.transfer.count += 1;
    }
  });

  function finish(t) {
    var list = Object.keys(t.cats).map(function (name) {
      var c = t.cats[name];
      return { name: name, husband: c.husband, wife: c.wife, total: c.husband + c.wife };
    }).filter(function (c) { return c.total !== 0; })
      .sort(function (a, b) { return b.total - a.total; });
    return { total: t.total, husband: t.husband, wife: t.wife, count: t.count, categories: list };
  }

  var monthList = Object.keys(months).sort();
  var result = { months: monthList, data: {} };
  monthList.forEach(function (m) {
    var b = months[m];
    result.data[m] = {
      expense: finish(b.expense), income: finish(b.income), transfer: b.transfer, days: b.days
    };
  });
  return result;
}

// ───────────────────────── 시트 읽기·쓰기 도우미 ─────────────────────────

function readTx_(ss) {
  return readAllTx_(ss.getSheetByName(SHEET_TX))
    .filter(function (v) { return v[COL.datetime] !== '' && v[COL.datetime] !== null; });
}

/** 거래 탭 전체 줄 (빈 줄 포함, 줄 번호 = 위치 + 2). 새 열이 아직 없으면 빈 칸으로 채움 */
function readAllTx_(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var n = Math.min(NUM_COLS, sheet.getMaxColumns());
  return sheet.getRange(2, 1, last - 1, n).getValues().map(function (r) {
    while (r.length < NUM_COLS) r.push('');
    return r;
  });
}

/** 원래 사용처, 삭제 열이 없으면 만들고 머리글을 씁니다. */
function ensureTxCols_(ss) {
  var sheet = ss.getSheetByName(SHEET_TX);
  var max = sheet.getMaxColumns();
  if (max < NUM_COLS) sheet.insertColumnsAfter(max, NUM_COLS - max);
  var head = sheet.getRange(1, 1, 1, NUM_COLS).getValues()[0];
  if (head[COL.origPlace] !== TX_HEADER[COL.origPlace] || head[COL.deleted] !== TX_HEADER[COL.deleted]) {
    sheet.getRange(1, COL.origPlace + 1, 1, 2).setValues([[TX_HEADER[COL.origPlace], TX_HEADER[COL.deleted]]]);
  }
  return sheet;
}

/** 사용처(C)부터 삭제(M)까지만 다시 씁니다. 날짜시간 열은 건드리지 않음 */
function writeRows_(sheet, values, from, count) {
  sheet.getRange(from + 2, COL.place + 1, count, NUM_COLS - COL.place)
    .setValues(values.slice(from, from + count).map(function (r) { return r.slice(COL.place, NUM_COLS); }));
}

function findRow_(values, id) {
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][COL.id]) === String(id)) return i;
  }
  throw new Error('거래를 찾지 못했습니다. 화면을 새로고침해 주세요.');
}

/** 카카오페이에 적힌 원래 사용처 (제목을 고쳤으면 '원래 사용처' 열) */
function placeKey_(v) {
  return String(v[COL.origPlace] || v[COL.place]).trim();
}

function isDeleted_(v) { return String(v[COL.deleted] || '') === '삭제'; }

function dayOf_(d, tz) {
  return (d instanceof Date) ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : String(d).slice(0, 10);
}

function secOf_(d, tz) {
  return (d instanceof Date) ? Utilities.formatDate(d, tz, 'yyyy-MM-dd HH:mm:ss') : normalizeDt(d);
}

function readCategories_(ss) {
  var sheet = ss.getSheetByName(SHEET_CAT);
  var last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, 3).getValues()
    .filter(function (r) { return r[0]; })
    .sort(function (a, b) { return a[2] - b[2]; })
    .map(function (r) { return { name: String(r[0]), kind: String(r[1]) }; });
}

function categoryKindMap_(ss) {
  var map = {};
  readCategories_(ss).forEach(function (c) { map[c.name] = c.kind; });
  return map;
}

function readSettings_(ss) {
  var sheet = ss.getSheetByName(SHEET_SET);
  var map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function (r) {
    if (r[0]) map[String(r[0]).trim()] = String(r[1]).trim();
  });
  return map;
}

function readRules_(ss) {
  var sheet = ss.getSheetByName(SHEET_RULE);
  var map = {};
  if (sheet.getLastRow() < 2) return map;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function (r) {
    if (r[0] && r[1]) map[String(r[0])] = String(r[1]);
  });
  return map;
}

/**
 * 분류에 필요한 정보를 모읍니다.
 * existing(거래 탭 값)을 주면, 사용처별로 예전에 가장 많이 쓴 카테고리를 history로 만듭니다.
 * (결제취소, 기타(미분류)는 빼고 셉니다.)
 */
function classifyContext_(ss, existing) {
  var history = {};
  if (existing) {
    var counts = {};
    existing.forEach(function (v) {
      var cat = String(v[COL.category]);
      if (!cat || cat === '결제취소' || cat === '기타(미분류)') return;
      var k = placeKey_(v) +((Number(v[COL.amount]) || 0) > 0 ? '|+' : '|-');
      var c = counts[k] || (counts[k] = {});
      c[cat] = (c[cat] || 0) + 1;
    });
    Object.keys(counts).forEach(function (k) {
      var best = null;
      Object.keys(counts[k]).forEach(function (cat) {
        if (!best || counts[k][cat] > counts[k][best]) best = cat;
      });
      history[k] = best;
    });
  }
  var set = readSettings_(ss);
  var names = [set['남편 이름'], set['아내 이름']].filter(function (n) { return n; });
  var spouseRe = names.length ? new RegExp(names.map(escapeRe_).join('|')) : null;
  var transferRe = null;
  if (set['이체 키워드']) {
    try { transferRe = new RegExp(set['이체 키워드']); } catch (e) { transferRe = null; }
  }
  return { rules: readRules_(ss), history: history, spouseRe: spouseRe, transferRe: transferRe };
}

function escapeRe_(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

function saveRule_(ss, place, category, editor, now) {
  var sheet = ss.getSheetByName(SHEET_RULE);
  var last = sheet.getLastRow();
  if (last >= 2) {
    var places = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < places.length; i++) {
      if (String(places[i][0]) === place) {
        sheet.getRange(i + 2, 2, 1, 3).setValues([[category, editor, now]]);
        return;
      }
    }
  }
  sheet.appendRow([place, category, editor, now]);
}

function currentUser_() {
  var email = '';
  try { email = Session.getActiveUser().getEmail(); } catch (e) {}
  return email || '알 수 없음';
}

// 컴퓨터에서 계산 검증용 (Apps Script에서는 무시됩니다)
if (typeof module !== 'undefined') {
  module.exports = { summarize: summarize, kindFor: kindFor, classifyTx: classifyTx, normalizeDt: normalizeDt };
}
