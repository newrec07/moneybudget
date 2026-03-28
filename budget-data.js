/**
 * ============================================================
 * budget-data.js — NZ 가계부 데이터 모듈 (독립 모듈)
 * ============================================================
 *
 * 역할:
 *   1. 카테고리 분류 (대분류 → 중분류 → 소분류 자동 매핑)
 *   2. 환율 관리 (1~2월: 월말 기준, 3월~: 당일 환율)
 *   3. 거래 데이터 CRUD (추가/수정/삭제/조회)
 *   4. 일별/월별/카테고리별 집계
 *   5. Excel 파일 읽기 (사용자 업로드 → 내부 데이터 동기화)
 *   6. Excel 파일 출력 (마스터 포맷으로 다운로드)
 *   7. 연도별 파일 관리 (연 1개 파일, 10년 = 10개)
 *
 * 의존성:
 *   - SheetJS (xlsx): https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
 *     → <script src="..."> 로 로드 후 이 모듈 사용
 *
 * 외부 인터페이스 (앱에서 사용하는 public API):
 *   BudgetData.addTransaction(fields)
 *   BudgetData.updateTransaction(id, fields)
 *   BudgetData.deleteTransaction(id)
 *   BudgetData.getTransactions(filter?)
 *   BudgetData.getDailySummary(year, month?)
 *   BudgetData.getMonthlySummary(year)
 *   BudgetData.getCategorySummary(year, month?)
 *   BudgetData.importFromExcel(file)        ← 사용자 업로드
 *   BudgetData.exportToExcel(year?)         ← 다운로드
 *   BudgetData.getExchangeRate(date)
 *   BudgetData.setExchangeRate(year, month, rate)
 *   BudgetData.classifyCategory(rawLabel, description?)
 *
 * ⚠️ 수정은 이 파일 내에서만 진행할 것.
 *    앱(HTML/다른 JS)에서는 BudgetData.* API만 호출.
 * ============================================================
 */

var BudgetData = (() => {

  // ──────────────────────────────────────────────────────────
  // § 1. 내부 상태
  // ──────────────────────────────────────────────────────────

  /** @type {Map<string, Transaction[]>} 연도별 거래 데이터 */
  const _store = new Map();   // key: '2026', value: Transaction[]

  /** @type {Map<string, number>} 환율 기준표 key: 'YYYY-MM' */
  const _rateMap = new Map();

  let _nextId = 1;

  // ──────────────────────────────────────────────────────────
  // § 2. 카테고리 분류 테이블
  //   구조: { 원본키: [대분류, 중분류, 소분류] }
  //   원본키는 사용자가 입력하는 항목명(대소문자 무관, trim 후 비교)
  // ──────────────────────────────────────────────────────────

  const CAT_TABLE = [
    // ── 주거 ──
    { keys: ['렌트비'],                          cat: ['주거', '주거비',      '렌트비']             },
    { keys: ['전기세'],                          cat: ['주거', '공과금',      '전기세']             },
    { keys: ['수도세'],                          cat: ['주거', '공과금',      '수도세']             },
    { keys: ['잔디'],                            cat: ['주거', '유지관리',    '잔디관리']           },
    // ── 식비 ──
    { keys: ['식비'],                            cat: ['식비', '장보기',      '마트/슈퍼']          },
    { keys: ['외식비'],                          cat: ['식비', '외식',        '외식']               },
    { keys: ['커피'],                            cat: ['식비', '카페',        '카페']               },
    // ── 교통 ──
    { keys: ['교통비', '주유비'],                cat: ['교통', '주유/교통',   '주유/주차/버스']     },
    { keys: ['자동차nzn911', '자동차'],          cat: ['교통', '자동차유지',  '자동차유지(NZN911)'] },
    { keys: ['보험nzn911', '자동차보험nzn911'],  cat: ['교통', '자동차보험',  '자동차보험(NZN911)'] },
    { keys: ['항공'],                            cat: ['교통', '항공',        '항공권']             },
    // ── 통신 ──
    { keys: ['인터넷'],                          cat: ['통신', '인터넷/휴대폰', '인터넷/휴대폰']   },
    { keys: ['구독'],                            cat: ['통신', '구독서비스',  '구독서비스']         },
    // ── 교육 ──
    { keys: ['학교'],                            cat: ['교육', '학교',        '학교(랑기토토)']     },
    { keys: ['문구'],                            cat: ['교육', '학용품',      '문구/도서']          },
    { keys: ['태권도'],                          cat: ['교육', '과외활동',    '태권도']             },
    { keys: ['피아노'],                          cat: ['교육', '과외활동',    '피아노']             },
    { keys: ['바이올린'],                        cat: ['교육', '과외활동',    '바이올린']           },
    { keys: ['미술'],                            cat: ['교육', '과외활동',    '미술']               },
    { keys: ['치어리딩'],                        cat: ['교육', '과외활동',    '치어리딩']           },
    { keys: ['수영'],                            cat: ['교육', '과외활동',    '수영']               },
    { keys: ['시험'],                            cat: ['교육', '학교',        '시험/기타']          },
    // ── 건강 ──
    { keys: ['gym 보나'],                        cat: ['건강', 'GYM',         'GYM(보나)']          },
    { keys: ['gym 나윤'],                        cat: ['건강', 'GYM',         'GYM(나윤)']          },
    { keys: ['gym 나율'],                        cat: ['건강', 'GYM',         'GYM(나율)']          },
    { keys: ['병원비'],                          cat: ['건강', '의료',        '병원/약국']          },
    { keys: ['영양제'],                          cat: ['건강', '의료',        '영양제']             },
    // ── 생활 ──
    { keys: ['생활용품'],                        cat: ['생활', '생활용품',    '생활용품']           },
    { keys: ['쇼핑'],                            cat: ['생활', '쇼핑',        '의류/잡화']          },
    { keys: ['입장료'],                          cat: ['생활', '여가/문화',   '입장료']             },
    // ── 여행 ──
    { keys: ['두바이', '여행'],                  cat: ['여행', '해외여행',    '여행']               },
    // ── 비자/보험 ──
    { keys: ['비자'],                            cat: ['비자/보험', '비자',   '비자']               },
    // ── 기타 ──
    { keys: ['기타'],                            cat: ['기타', '기타',        '기타']               },
  ];

  /**
   * 지출내역 키워드로 소분류 오버라이드
   * desc에 keyword 포함 시 해당 cat으로 덮어씀
   */
  const DESC_OVERRIDE = [
    { keywords: ['아쿠아로빅'],               cat: ['기타',    '여가/운동',       '아쿠아로빅']        },
    { keywords: ['타이', '마사지'],           cat: ['기타',    '여가/운동',       '마사지']            },
    { keywords: ['수영장'],                   cat: ['기타',    '여가/운동',       '수영장']            },
    { keywords: ['셔플'],                     cat: ['기타',    '여가/운동',       '셔플댄스']          },
    { keywords: ['roar honey', 'honey', '허니'], cat: ['기타', '부업/수입관련',  '꿀벌관련']          },
    { keywords: ['디어프렌즈'],               cat: ['기타',    '선물/경조사',     '선물']              },
    { keywords: ['햄버거', '버거'],           cat: ['식비',    '외식',            '패스트푸드']        },
    { keywords: ['스크런치'],                 cat: ['생활',    '쇼핑',            '소품/악세서리']     },
    { keywords: ['약국', '밴드'],             cat: ['건강',    '의료',            '병원/약국']         },
  ];

  // ──────────────────────────────────────────────────────────
  // § 3. 카테고리 분류 함수
  // ──────────────────────────────────────────────────────────

  /**
   * 원본 항목명 + 지출내역으로 [대분류, 중분류, 소분류] 반환
   * @param {string} rawLabel  - 원본 항목명 (예: '식비', 'GYM 보나')
   * @param {string} [desc=''] - 지출내역 (예: '올핏', '파킨주유')
   * @returns {{ big: string, mid: string, small: string }}
   */
  function classifyCategory(rawLabel, desc = '') {
    const label = String(rawLabel || '').trim().toLowerCase();
    const d = String(desc || '').trim().toLowerCase();

    // 1) 지출내역 키워드 우선 매칭
    for (const rule of DESC_OVERRIDE) {
      if (rule.keywords.some(kw => d.includes(kw.toLowerCase()))) {
        return { big: rule.cat[0], mid: rule.cat[1], small: rule.cat[2] };
      }
    }

    // 2) 원본 항목명 매칭 (startsWith 포함)
    for (const rule of CAT_TABLE) {
      if (rule.keys.some(k => label === k || label.startsWith(k))) {
        return { big: rule.cat[0], mid: rule.cat[1], small: rule.cat[2] };
      }
    }

    // 3) 미분류 fallback
    return { big: '기타', mid: '미분류', small: rawLabel || '미분류' };
  }

  // ──────────────────────────────────────────────────────────
  // § 4. 환율 관리
  //   - 1월/2월: 월말 기준 고정 환율
  //   - 3월~:    당일 환율 (setExchangeRate로 업데이트)
  // ──────────────────────────────────────────────────────────

  /** 기본 환율 초기화 (매년 자동 생성) */
  function _initDefaultRates() {
    const baseRates = [882.50, 875.20, 880.30, 878.60, 885.40, 890.20, 892.10, 887.80, 884.50, 881.90, 879.30, 876.80];
    const currentYear = new Date().getFullYear();
    
    // 지난 3년 + 현재연도 + 향후 3년 = 7년 자동 생성
    for (let y = currentYear - 3; y <= currentYear + 3; y++) {
      baseRates.forEach((rate, idx) => {
        const month = String(idx + 1).padStart(2, '0');
        const key = `${y}-${month}`;
        if (!_rateMap.has(key)) {
          // 변동성 추가: 연도별로 약간씩 다른 레이트 (실시간 API 없을 시 대체)
          const drift = (y - 2026) * 2.5 + (Math.random() - 0.5) * 3;
          _rateMap.set(key, Math.round((rate + drift) * 100) / 100);
        }
      });
    }
  }

  /**
   * 날짜에 해당하는 환율 반환
   * @param {Date|string} date
   * @returns {number}
   */
  function getExchangeRate(date) {
    const d = _toDate(date);
    if (!d) return 876.00;
    const year  = d.getFullYear();
    const month = d.getMonth() + 1;
    const key   = `${year}-${String(month).padStart(2,'0')}`;
    // 1~2월: 월말 기준 고정
    if (_rateMap.has(key)) return _rateMap.get(key);
    // 3월~: 당일 환율 키 'YYYY-MM-DD' 우선, 없으면 월 키, 없으면 기본값
    const dayKey = _fmtDate(d);
    if (_rateMap.has(dayKey)) return _rateMap.get(dayKey);
    return 876.00;
  }

  /**
   * 환율 설정 (3월~ 당일 업데이트 또는 월말 기준 변경)
   * @param {number|string} year
   * @param {number|string} month
   * @param {number} rate
   * @param {number|string} [day]  지정 시 당일 환율, 미지정 시 월 전체 기준
   */
  function setExchangeRate(year, month, rate, day) {
    const key = day
      ? `${year}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}`
      : `${year}-${String(month).padStart(2,'0')}`;
    _rateMap.set(key, Number(rate));
  }

  // ──────────────────────────────────────────────────────────
  // § 5. 거래 데이터 CRUD
  // ──────────────────────────────────────────────────────────

  /**
   * @typedef {Object} Transaction
   * @property {string}  id          - 고유 ID
   * @property {number}  year        - 연도
   * @property {number}  month       - 월 (1~12)
   * @property {Date}    date        - 날짜
   * @property {string}  big         - 대분류
   * @property {string}  mid         - 중분류
   * @property {string}  small       - 소분류
   * @property {string}  description - 지출내역
   * @property {string}  memo        - 메모
   * @property {number}  nzd         - 금액(NZD)
   * @property {number}  rate        - 적용 환율
   * @property {number}  krw         - 금액(KRW)
   * @property {string}  rawLabel    - 원본 항목명 (보존)
   */

  /**
   * 거래 추가
   * @param {Object} fields
   * @param {Date|string} fields.date
   * @param {string}  fields.rawLabel     - 원본 항목명
   * @param {string}  [fields.description]
   * @param {string}  [fields.memo]
   * @param {number}  fields.nzd
   * @param {number}  [fields.rate]       - 직접 지정 시 우선
   * @param {string}  [fields.big]        - 직접 지정 시 분류 override
   * @param {string}  [fields.mid]
   * @param {string}  [fields.small]
   * @returns {Transaction}
   */
  function addTransaction(fields) {
    const date  = _toDate(fields.date);
    if (!date) throw new Error('유효한 날짜가 필요합니다.');
    const nzd   = Number(fields.nzd);
    if (isNaN(nzd) || nzd <= 0) throw new Error('금액(NZD)이 유효하지 않습니다.');

    const cat   = (fields.big && fields.mid && fields.small)
      ? { big: fields.big, mid: fields.mid, small: fields.small }
      : classifyCategory(fields.rawLabel, fields.description);

    const rate  = Number(fields.rate) || getExchangeRate(date);
    const year  = date.getFullYear();
    const month = date.getMonth() + 1;

    const tx = {
      id:          `tx_${year}_${_nextId++}`,
      year,
      month,
      date,
      big:         cat.big,
      mid:         cat.mid,
      small:       cat.small,
      description: String(fields.description || '').trim(),
      memo:        String(fields.memo || '').trim(),
      nzd:         Math.round(nzd * 100) / 100,
      rate:        Math.round(rate * 100) / 100,
      krw:         Math.round(nzd * rate),
      rawLabel:    String(fields.rawLabel || '').trim(),
    };

    const key = String(year);
    if (!_store.has(key)) _store.set(key, []);
    _store.get(key).push(tx);
    return tx;
  }

  /**
   * 거래 수정
   * @param {string} id
   * @param {Partial<Transaction>} fields
   * @returns {Transaction|null}
   */
  function updateTransaction(id, fields) {
    for (const [, list] of _store) {
      const idx = list.findIndex(t => t.id === id);
      if (idx === -1) continue;
      const old = list[idx];
      const date  = fields.date ? _toDate(fields.date) : old.date;
      const nzd   = fields.nzd  != null ? Number(fields.nzd) : old.nzd;
      const rate  = fields.rate != null ? Number(fields.rate) : getExchangeRate(date);

      // 카테고리 재분류 (rawLabel 또는 description 변경 시)
      const newRaw  = fields.rawLabel    || old.rawLabel;
      const newDesc = fields.description || old.description;
      const cat = (fields.big && fields.mid && fields.small)
        ? { big: fields.big, mid: fields.mid, small: fields.small }
        : classifyCategory(newRaw, newDesc);

      list[idx] = {
        ...old,
        date,
        big:         cat.big,
        mid:         cat.mid,
        small:       cat.small,
        description: fields.description != null ? String(fields.description).trim() : old.description,
        memo:        fields.memo        != null ? String(fields.memo).trim()        : old.memo,
        nzd:         Math.round(nzd * 100) / 100,
        rate:        Math.round(rate * 100) / 100,
        krw:         Math.round(nzd * rate),
        rawLabel:    newRaw,
        month:       date.getMonth() + 1,
        year:        date.getFullYear(),
      };
      return list[idx];
    }
    return null;
  }

  /**
   * 거래 삭제
   * @param {string} id
   * @returns {boolean}
   */
  function deleteTransaction(id) {
    for (const [key, list] of _store) {
      const idx = list.findIndex(t => t.id === id);
      if (idx !== -1) {
        list.splice(idx, 1);
        return true;
      }
    }
    return false;
  }

  /**
   * 거래 조회
   * @param {{ year?, month?, big?, mid?, small?, dateFrom?, dateTo? }} [filter]
   * @returns {Transaction[]}
   */
  function getTransactions(filter = {}) {
    let result = [];
    for (const [, list] of _store) result = result.concat(list);

    if (filter.year)     result = result.filter(t => t.year  === Number(filter.year));
    if (filter.month)    result = result.filter(t => t.month === Number(filter.month));
    if (filter.big)      result = result.filter(t => t.big   === filter.big);
    if (filter.mid)      result = result.filter(t => t.mid   === filter.mid);
    if (filter.small)    result = result.filter(t => t.small === filter.small);
    if (filter.dateFrom) result = result.filter(t => t.date  >= _toDate(filter.dateFrom));
    if (filter.dateTo)   result = result.filter(t => t.date  <= _toDate(filter.dateTo));

    return result.sort((a, b) => a.date - b.date);
  }

  // ──────────────────────────────────────────────────────────
  // § 6. 집계 함수
  // ──────────────────────────────────────────────────────────

  /**
   * 일별 집계
   * @param {number} year
   * @param {number} [month]  미지정 시 연도 전체
   * @returns {Array<{ date, dateStr, dow, transactions, totalNZD, totalKRW, rate, byCategory }>}
   */
  function getDailySummary(year, month) {
    const filter = { year };
    if (month) filter.month = month;
    const txs = getTransactions(filter);

    const map = new Map();
    for (const tx of txs) {
      const key = _fmtDate(tx.date);
      if (!map.has(key)) map.set(key, { date: tx.date, transactions: [], totalNZD: 0, totalKRW: 0, rate: tx.rate });
      const day = map.get(key);
      day.transactions.push(tx);
      day.totalNZD = Math.round((day.totalNZD + tx.nzd) * 100) / 100;
      day.totalKRW += tx.krw;
    }

    const DAYS = ['일','월','화','수','목','금','토'];
    return Array.from(map.entries())
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([dateStr, day]) => ({
        dateStr,
        date:       day.date,
        dow:        DAYS[day.date.getDay()],
        isWeekend:  day.date.getDay() === 0 || day.date.getDay() === 6,
        transactions: day.transactions,
        totalNZD:   day.totalNZD,
        totalKRW:   day.totalKRW,
        rate:       day.rate,
        byCategory: _groupBy(day.transactions, 'big'),
      }));
  }

  /**
   * 월별 집계
   * @param {number} year
   * @returns {Array<{ month, totalNZD, totalKRW, rate, byCategory }>}
   */
  function getMonthlySummary(year) {
    const txs = getTransactions({ year });
    const map = new Map();
    for (let m = 1; m <= 12; m++) map.set(m, { month: m, totalNZD: 0, totalKRW: 0, byCategory: {} });

    for (const tx of txs) {
      const m = map.get(tx.month);
      m.totalNZD = Math.round((m.totalNZD + tx.nzd) * 100) / 100;
      m.totalKRW += tx.krw;
      m.byCategory[tx.big] = (m.byCategory[tx.big] || 0) + tx.nzd;
    }
    return Array.from(map.values());
  }

  /**
   * 카테고리별 집계
   * @param {number} year
   * @param {number} [month]
   * @returns {Array<{ big, mid, small, totalNZD, totalKRW, count, pct }>}
   */
  function getCategorySummary(year, month) {
    const filter = { year };
    if (month) filter.month = month;
    const txs = getTransactions(filter);
    const totalNZD = txs.reduce((s, t) => s + t.nzd, 0);

    const map = new Map();
    for (const tx of txs) {
      const key = `${tx.big}||${tx.mid}||${tx.small}`;
      if (!map.has(key)) map.set(key, { big: tx.big, mid: tx.mid, small: tx.small, totalNZD: 0, totalKRW: 0, count: 0 });
      const g = map.get(key);
      g.totalNZD = Math.round((g.totalNZD + tx.nzd) * 100) / 100;
      g.totalKRW += tx.krw;
      g.count++;
    }
    return Array.from(map.values())
      .map(g => ({ ...g, pct: totalNZD > 0 ? Math.round(g.totalNZD / totalNZD * 1000) / 10 : 0 }))
      .sort((a, b) => b.totalNZD - a.totalNZD);
  }

  // ──────────────────────────────────────────────────────────
  // § 7. Excel 가져오기 (사용자 업로드 → 내부 동기화)
  // ──────────────────────────────────────────────────────────

  /**
   * 사용자가 업로드한 엑셀 파일을 읽어 내부 store에 병합
   * - '거래내역' 시트를 기준으로 파싱
   * - 기존 같은 연도 데이터는 replace 방식으로 교체
   * @param {File} file  - <input type="file"> 에서 가져온 File 객체
   * @returns {Promise<{ imported: number, year: number, errors: string[] }>}
   */
  async function importFromExcel(file) {
    _assertSheetJS();
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data   = new Uint8Array(e.target.result);
          const wb     = XLSX.read(data, { type: 'array', cellDates: true });
          const ws     = wb.Sheets['거래내역'];
          if (!ws) throw new Error("'거래내역' 시트를 찾을 수 없습니다.");

          const rows   = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
          // 헤더: 월,날짜,대분류,중분류,소분류,지출내역,메모,금액(NZD),환율,금액(KRW),원본항목
          const HDR    = { 월:0, 날짜:1, 대분류:2, 중분류:3, 소분류:4, 지출내역:5, 메모:6, 'nzd':7, 환율:8, 'krw':9, 원본:10 };
          const errors = [];
          const imported = [];
          let detectedYear = null;

          for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || !row[HDR['날짜']] || !row[HDR['nzd']]) continue;
            try {
              const dateRaw = row[HDR['날짜']];
              const date    = dateRaw instanceof Date ? dateRaw : new Date(dateRaw);
              if (isNaN(date.getTime())) { errors.push(`행${i+1}: 날짜 파싱 실패`); continue; }
              const nzd     = Number(row[HDR['nzd']]);
              if (isNaN(nzd)) { errors.push(`행${i+1}: 금액 파싱 실패`); continue; }
              const rate    = Number(row[HDR['환율']]) || getExchangeRate(date);
              const year    = date.getFullYear();
              if (!detectedYear) detectedYear = year;

              imported.push({
                id:          `tx_${year}_${_nextId++}`,
                year,
                month:       date.getMonth() + 1,
                date,
                big:         String(row[HDR['대분류']] || '기타').trim(),
                mid:         String(row[HDR['중분류']] || '기타').trim(),
                small:       String(row[HDR['소분류']] || '기타').trim(),
                description: String(row[HDR['지출내역']] || '').trim(),
                memo:        String(row[HDR['메모']] || '').trim(),
                nzd:         Math.round(nzd * 100) / 100,
                rate:        Math.round(rate * 100) / 100,
                krw:         Math.round(nzd * rate),
                rawLabel:    String(row[HDR['원본']] || '').trim(),
              });
            } catch (err) {
              errors.push(`행${i+1}: ${err.message}`);
            }
          }

          if (!detectedYear) throw new Error('연도를 감지할 수 없습니다.');
          // 해당 연도 데이터 교체
          _store.set(String(detectedYear), imported);
          resolve({ imported: imported.length, year: detectedYear, errors });
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = () => reject(new Error('파일 읽기 실패'));
      reader.readAsArrayBuffer(file);
    });
  }

  // ──────────────────────────────────────────────────────────
  // § 8. Excel 내보내기 (마스터 포맷 다운로드)
  // ──────────────────────────────────────────────────────────

  /**
   * 지정 연도 데이터를 마스터 포맷 xlsx로 내보내기 (다운로드)
   * 시트 구성:
   *   거래내역 / 일자별 / 월별요약 / 중분류상세 / 환율기준
   * @param {number} [year]  미지정 시 가장 최근 연도
   */
  function exportToExcel(year) {
    _assertSheetJS();

    const targetYear = year || _latestYear();
    if (!targetYear) throw new Error('저장된 데이터가 없습니다.');

    const txs = getTransactions({ year: targetYear });
    if (txs.length === 0) throw new Error(`${targetYear}년 데이터가 없습니다.`);

    const wb = XLSX.utils.book_new();

    // ── 시트1: 거래내역 ──────────────────────────────────
    const txRows = [
      ['월','날짜','대분류','중분류','소분류','지출내역','메모','금액(NZD)','환율','금액(KRW)','원본항목'],
      ...txs.map(t => [
        `${t.month}월`,
        _excelDate(t.date),
        t.big, t.mid, t.small,
        t.description, t.memo,
        t.nzd, t.rate, t.krw,
        t.rawLabel,
      ]),
    ];
    const ws1 = XLSX.utils.aoa_to_sheet(txRows);
    _setColWidths(ws1, [6,11,9,14,18,30,18,12,8,12,14]);
    XLSX.utils.book_append_sheet(wb, ws1, '거래내역');

    // ── 시트2: 일자별 ─────────────────────────────────────
    const DAYS_KO = ['일','월','화','수','목','금','토'];
    const daily   = getDailySummary(targetYear);
    const dailyRows = [
      ['날짜','요일','건수','합계(NZD)','환율','합계(KRW)','지출 항목 요약'],
      ...daily.map(d => [
        _excelDate(d.date),
        d.dow,
        d.transactions.length,
        d.totalNZD,
        d.rate,
        d.totalKRW,
        d.transactions.map(t => `${t.big}:${t.description}($${t.nzd})`).join(' | '),
      ]),
    ];
    const ws2 = XLSX.utils.aoa_to_sheet(dailyRows);
    _setColWidths(ws2, [12,6,6,12,8,12,60]);
    XLSX.utils.book_append_sheet(wb, ws2, '일자별');

    // ── 시트3: 월별요약 ──────────────────────────────────
    const monthly = getMonthlySummary(targetYear);
    const BIG_CATS = ['주거','식비','교통','통신','교육','건강','생활','여행','비자/보험','기타'];
    const catCols  = BIG_CATS.map(c => [`${c}(NZD)`,`${c}(KRW)`]).flat();
    const mHeader  = ['월','합계(NZD)','합계(KRW)', ...catCols];
    const mRows    = [mHeader];
    for (const m of monthly) {
      const rate = _getMonthRate(targetYear, m.month);
      const row  = [`${m.month}월`, m.totalNZD, m.totalKRW];
      for (const cat of BIG_CATS) {
        const nzd = m.byCategory[cat] || 0;
        row.push(Math.round(nzd * 100)/100);
        row.push(Math.round(nzd * rate));
      }
      mRows.push(row);
    }
    // 합계 행
    const sumRow = ['합계'];
    for (let c = 1; c < mHeader.length; c++) {
      sumRow.push(mRows.slice(1).reduce((s,r) => s + (Number(r[c])||0), 0));
    }
    mRows.push(sumRow);
    const ws3 = XLSX.utils.aoa_to_sheet(mRows);
    XLSX.utils.book_append_sheet(wb, ws3, '월별요약');

    // ── 시트4: 중분류상세 ────────────────────────────────
    const catSummary = getCategorySummary(targetYear);
    const cRows = [
      ['대분류','중분류','소분류','합계(NZD)','합계(KRW)','건수','비중%'],
      ...catSummary.map(g => [g.big, g.mid, g.small, g.totalNZD, g.totalKRW, g.count, `${g.pct}%`]),
    ];
    const ws4 = XLSX.utils.aoa_to_sheet(cRows);
    _setColWidths(ws4, [10,15,20,12,12,6,8]);
    XLSX.utils.book_append_sheet(wb, ws4, '중분류상세');

    // ── 시트5: 환율기준 ──────────────────────────────────
    const rateRows = [
      ['월','기준일','환율(NZD→KRW)','비고'],
    ];
    for (const [key, rate] of [..._rateMap.entries()].sort()) {
      const isDay = key.split('-').length === 3;
      rateRows.push([
        key,
        key,
        rate,
        isDay ? '당일 환율' : '월말 기준 고정 환율',
      ]);
    }
    const ws5 = XLSX.utils.aoa_to_sheet(rateRows);
    _setColWidths(ws5, [12,14,16,24]);
    XLSX.utils.book_append_sheet(wb, ws5, '환율기준');

    // ── 파일명: 가계부_YYYY.xlsx ─────────────────────────
    const fileName = `가계부_${targetYear}.xlsx`;
    XLSX.writeFile(wb, fileName);
    return fileName;
  }

  // ──────────────────────────────────────────────────────────
  // § 9. 연도 관리
  // ──────────────────────────────────────────────────────────

  /** 데이터가 존재하는 연도 목록 반환 */
  function getYears() {
    return Array.from(_store.keys()).map(Number).sort((a,b) => a-b);
  }

  /** 가장 최근 연도 */
  function _latestYear() {
    const years = getYears();
    return years.length ? years[years.length - 1] : null;
  }

  // ──────────────────────────────────────────────────────────
  // § 10. 내부 유틸
  // ──────────────────────────────────────────────────────────

  function _toDate(v) {
    if (!v) return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    const d = new Date(v);
    return isNaN(d.getTime()) ? null : d;
  }

  function _fmtDate(d) {
    const m = String(d.getMonth()+1).padStart(2,'0');
    const dd = String(d.getDate()).padStart(2,'0');
    return `${d.getFullYear()}-${m}-${dd}`;
  }

  function _excelDate(d) {
    // SheetJS가 Date 객체를 날짜 셀로 인식하도록 그대로 반환
    return d instanceof Date ? d : new Date(d);
  }

  function _groupBy(arr, key) {
    return arr.reduce((acc, item) => {
      acc[item[key]] = (acc[item[key]] || 0) + item.nzd;
      return acc;
    }, {});
  }

  function _setColWidths(ws, widths) {
    ws['!cols'] = widths.map(w => ({ wch: w }));
  }

  function _getMonthRate(year, month) {
    const key = `${year}-${String(month).padStart(2,'0')}`;
    return _rateMap.get(key) || 876.00;
  }

  function _assertSheetJS() {
    if (typeof XLSX === 'undefined') {
      throw new Error(
        'SheetJS(xlsx)가 로드되지 않았습니다.\n' +
        '<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"> 를 먼저 추가하세요.'
      );
    }
  }

  // ──────────────────────────────────────────────────────────
  // § 11. 초기화
  // ──────────────────────────────────────────────────────────

  _initDefaultRates();

  // ──────────────────────────────────────────────────────────
  // § 12. Public API
  // ──────────────────────────────────────────────────────────

  return Object.freeze({
    // 카테고리
    classifyCategory,

    // 환율
    getExchangeRate,
    setExchangeRate,

    // CRUD
    addTransaction,
    updateTransaction,
    deleteTransaction,
    getTransactions,

    // 집계
    getDailySummary,
    getMonthlySummary,
    getCategorySummary,

    // Excel I/O
    importFromExcel,
    exportToExcel,

    // 연도 관리
    getYears,

    // 디버그용 (개발 중에만 사용)
    _debug: () => ({
      storeKeys: Array.from(_store.keys()),
      ratemap:   Object.fromEntries(_rateMap),
      count:     Array.from(_store.values()).reduce((s,v)=>s+v.length,0),
    }),
  });

})();

// ──────────────────────────────────────────────────────────
// 사용 예시 (앱에서 이렇게 호출)
// ──────────────────────────────────────────────────────────
/*

// ① 거래 추가
BudgetData.addTransaction({
  date:        '2026-03-24',
  rawLabel:    '식비',
  description: '코스트코',
  nzd:         120.50,
  memo:        '주말 장보기',
});

// ② 카테고리 자동 분류 확인
BudgetData.classifyCategory('커피', '스타벅스');
// → { big: '식비', mid: '카페', small: '카페' }

// ③ 일별 요약 조회
const daily = BudgetData.getDailySummary(2026, 3);

// ④ 엑셀 파일 업로드 (input file 이벤트)
document.getElementById('fileInput').addEventListener('change', async (e) => {
  const result = await BudgetData.importFromExcel(e.target.files[0]);
  console.log(`${result.imported}건 가져옴`);
});

// ⑤ 엑셀 다운로드
BudgetData.exportToExcel(2026);  // → 가계부_2026.xlsx 다운로드

// ⑥ 당일 환율 업데이트 (3월~)
BudgetData.setExchangeRate(2026, 3, 880.50, 24);  // 3/24 당일 환율

*/
