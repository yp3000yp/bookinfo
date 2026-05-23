/**
 * 📚 교과서 채택 현황 — Google Apps Script
 * 
 * ▶ 동작 방식
 *   스프레드시트 저장 시 → data.json 생성 → GitHub API로 자동 커밋
 *   → GitHub Actions 자동 실행 → GitHub Pages 배포
 *
 * ▶ 시트 구조 (예: "2026" 시트)
 *   A1: 학교명 (헤더)   B1~: 학교 이름들
 *   A2: 1학년 (섹션 헤더, 병합행)
 *   A3: 과목명          B3~: 출판사(저자) 데이터
 *   ...
 *   A열에 "1학년" "2학년" "3학년" 이 나오면 다음 학년 시작으로 인식
 *
 * ▶ 설정 방법
 *   1. 아래 CONFIG 수정
 *   2. 스크립트 에디터 > 실행 > setupTrigger() 한 번 실행 (트리거 자동 등록)
 *   3. GitHub Personal Access Token 을 Script Properties에 저장
 *      (프로젝트 설정 > 스크립트 속성 > GITHUB_TOKEN 키로 추가)
 */

// ════════════════════════════════════════
//  ⚙️  설정 — 이 부분만 수정하세요
// ════════════════════════════════════════
const CONFIG = {
  GITHUB_OWNER: 'yp3000yp',          // GitHub 사용자명
  GITHUB_REPO:  'bookinfo',           // 저장소 이름
  GITHUB_BRANCH: 'main',             // 브랜치 (main 또는 master)
  DATA_FILE_PATH: 'data.json',        // 저장소 내 파일 경로

  // 스프레드시트 설정
  META_SHEET_NAME: 'META',           // 메타정보 시트 이름 (없으면 기본값 사용)
  YEAR_SHEET_PREFIX: '',             // 연도 시트 prefix (예: "data_" → "data_2026")
                                     // 비워두면 시트 이름이 숫자(연도)인 것만 처리

  // 기본 메타 정보
  DEFAULT_REGION: '부산광역시',
  DEFAULT_DISTRICT: '금정구',
  DEFAULT_SCHOOL_TYPE: '중학교',
};

// ════════════════════════════════════════
//  🚀  메인 함수 (트리거에 연결)
// ════════════════════════════════════════

/**
 * 스프레드시트 변경 시 자동 실행 (onEdit 트리거로 등록)
 * 단, 실제 배포는 저장 버튼 클릭 시에만 하도록 별도 메뉴도 제공
 */
function onSheetEdit(e) {
  // 디바운스: 너무 잦은 호출 방지 (30초에 1번만 실행)
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  
  try {
    syncToGitHub();
  } finally {
    lock.releaseLock();
  }
}

/**
 * 수동 실행용: 메뉴에서 "GitHub에 배포" 클릭 시
 */
function manualSync() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '📚 GitHub 배포',
    '현재 데이터를 GitHub에 업로드하고 웹사이트를 갱신합니다.\n계속할까요?',
    ui.ButtonSet.YES_NO
  );
  if (result !== ui.Button.YES) return;

  try {
    syncToGitHub();
    ui.alert('✅ 완료', 'GitHub 배포가 시작되었습니다.\n1~2분 후 웹사이트에 반영됩니다.', ui.ButtonSet.OK);
  } catch(err) {
    ui.alert('❌ 오류', err.message, ui.ButtonSet.OK);
  }
}

// ════════════════════════════════════════
//  📦  핵심 로직
// ════════════════════════════════════════

function syncToGitHub() {
  const json = buildDataJson();
  const jsonStr = JSON.stringify(json, null, 2);
  pushToGitHub(jsonStr);
  Logger.log('✅ GitHub 동기화 완료: ' + new Date().toLocaleString('ko-KR'));
}

/**
 * 스프레드시트 → data.json 구조로 변환
 */
function buildDataJson() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // 메타 정보
  let meta = {
    region: CONFIG.DEFAULT_REGION,
    district: CONFIG.DEFAULT_DISTRICT,
    schoolType: CONFIG.DEFAULT_SCHOOL_TYPE,
    lastUpdated: Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd')
  };

  // META 시트가 있으면 메타 정보 읽기
  const metaSheet = ss.getSheetByName(CONFIG.META_SHEET_NAME);
  if (metaSheet) {
    const metaData = metaSheet.getDataRange().getValues();
    metaData.forEach(row => {
      if (row[0] === 'region')    meta.region    = row[1] || meta.region;
      if (row[0] === 'district')  meta.district  = row[1] || meta.district;
      if (row[0] === 'schoolType') meta.schoolType = row[1] || meta.schoolType;
    });
  }

  // 출판사 색상 (PUB_COLORS 시트 또는 기본값)
  const pubColors = getPubColors(ss);

  // 연도별 데이터
  const years = {};
  sheets.forEach(sheet => {
    const name = sheet.getName();
    // 시트 이름이 4자리 숫자(연도)인 경우만 처리
    if (!/^\d{4}$/.test(name)) return;

    try {
      years[name] = parseYearSheet(sheet);
    } catch(e) {
      Logger.log('⚠️ 시트 파싱 오류 [' + name + ']: ' + e.message);
    }
  });

  return { meta, pubColors, years };
}

/**
 * 연도 시트 파싱
 * 
 * 시트 레이아웃:
 * 행1: [학교명] [구서여중] [금양중] ...  ← 헤더 (학교 목록)
 * 행2: [1학년]                            ← 학년 구분자 (A열만 사용)
 * 행3: [국어]   [천재(정호웅)] ...        ← 과목 + 출판사
 * 행4: [수학]   ...
 * 행N: [2학년]                            ← 다음 학년
 * ...
 */
function parseYearSheet(sheet) {
  // getValues() 대신 getDisplayValues() 사용
  // → 셀 서식(일반/텍스트)에 관계없이 화면에 보이는 값 그대로 읽음
  // → 괄호 포함 문자열(천재(안대회)) 등 수식 오인 문제 방지
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) throw new Error('데이터가 부족합니다');

  // 학교 목록 (1행, B열부터)
  const schools = data[0].slice(1).map(v => v.trim()).filter(v => v);

  const grades = {};
  let currentGrade = null;
  let currentSubjects = [];
  let currentRows = [];

  // 2행부터 파싱
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const first = row[0].trim();

    if (!first) continue; // 빈 행 무시

    // 학년 구분자 감지 (A열에 "1학년", "2학년", "3학년" 등)
    if (/학년$/.test(first) && row.slice(1).every(v => !v.trim())) {
      // 이전 학년 저장
      if (currentGrade && currentSubjects.length > 0) {
        grades[currentGrade] = { subjects: currentSubjects, rows: currentRows };
      }
      currentGrade = first;
      currentSubjects = [];
      currentRows = [];
    } else if (currentGrade) {
      // 과목 데이터 행
      const subject = first;
      const vals = row.slice(1, 1 + schools.length).map(v => {
        const s = v.trim();
        // 빈값·에러값 통일
        if (!s || s === '#ERROR!' || s === '#VALUE!' || s === '#REF!') return '—';
        return s;
      });
      currentSubjects.push(subject);
      currentRows.push(vals);
    }
  }

  // 마지막 학년 저장
  if (currentGrade && currentSubjects.length > 0) {
    grades[currentGrade] = { subjects: currentSubjects, rows: currentRows };
  }

  // label 생성
  const yearNum = sheet.getName();
  return {
    label: yearNum + '학년도',
    schools,
    grades
  };
}

/**
 * 출판사 색상 읽기 (PUB_COLORS 시트 또는 기본값)
 */
function getPubColors(ss) {
  const defaults = {
    "천재":   "#2980B9",
    "미래엔": "#27AE60",
    "동아":   "#E67E22",
    "비상":   "#8E44AD",
    "지학사": "#D4AC0D",
    "해냄":   "#E74C3C",
    "능률":   "#16A085",
    "금성":   "#B7950B",
    "신사고": "#2471A3",
    "창비":   "#6C3483",
    "교학사": "#935116",
    "교학":   "#935116",
    "YBM":    "#1A5276",
    "리베르": "#7D3C98",
    "씨마스": "#117A65",
    "길벗":   "#B03A2E",
    "박영사": "#784212"
  };

  const colorSheet = ss.getSheetByName('PUB_COLORS');
  if (!colorSheet) return defaults;

  const data = colorSheet.getDataRange().getValues();
  const colors = {};
  data.forEach(row => {
    if (row[0] && row[1]) colors[String(row[0]).trim()] = String(row[1]).trim();
  });
  return Object.keys(colors).length > 0 ? colors : defaults;
}

/**
 * GitHub API로 파일 업데이트
 */
function pushToGitHub(content) {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) throw new Error('GITHUB_TOKEN이 설정되지 않았습니다.\n스크립트 속성에서 GITHUB_TOKEN을 추가해 주세요.');

  const apiUrl = `https://api.github.com/repos/${CONFIG.GITHUB_OWNER}/${CONFIG.GITHUB_REPO}/contents/${CONFIG.DATA_FILE_PATH}`;

  // 현재 파일의 SHA 가져오기 (업데이트에 필요)
  let sha = null;
  try {
    const getRes = UrlFetchApp.fetch(apiUrl, {
      headers: {
        'Authorization': 'token ' + token,
        'Accept': 'application/vnd.github.v3+json'
      },
      muteHttpExceptions: true
    });
    if (getRes.getResponseCode() === 200) {
      sha = JSON.parse(getRes.getContentText()).sha;
    }
  } catch(e) {
    // 파일이 없는 경우 (신규 생성) — sha 없이 진행
  }

  // 파일 업데이트 (PUT)
  const body = {
    message: `📚 데이터 업데이트 (${Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm')})`,
    content: Utilities.base64Encode(content, Utilities.Charset.UTF_8),
    branch: CONFIG.GITHUB_BRANCH
  };
  if (sha) body.sha = sha;

  const putRes = UrlFetchApp.fetch(apiUrl, {
    method: 'PUT',
    headers: {
      'Authorization': 'token ' + token,
      'Accept': 'application/vnd.github.v3+json',
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  const code = putRes.getResponseCode();
  if (code !== 200 && code !== 201) {
    throw new Error(`GitHub API 오류 (HTTP ${code}): ${putRes.getContentText()}`);
  }

  // data.json 커밋 후 Actions workflow_dispatch로 배포 강제 실행
  // (push 트리거가 API 커밋을 감지 못하는 경우 대비)
  triggerGitHubActions(token);
}

/**
 * GitHub Actions workflow_dispatch 직접 호출
 * Apps Script → GitHub API 커밋만으로는 Actions push 트리거가
 * 발동하지 않을 수 있으므로, 커밋 후 명시적으로 배포를 실행
 */
function triggerGitHubActions(token) {
  const dispatchUrl = `https://api.github.com/repos/${CONFIG.GITHUB_OWNER}/${CONFIG.GITHUB_REPO}/actions/workflows/deploy.yml/dispatches`;

  const res = UrlFetchApp.fetch(dispatchUrl, {
    method: 'POST',
    headers: {
      'Authorization': 'token ' + token,
      'Accept': 'application/vnd.github.v3+json',
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({ ref: CONFIG.GITHUB_BRANCH }),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code === 204) {
    Logger.log('✅ GitHub Actions 배포 트리거 성공');
  } else if (code === 422) {
    Logger.log('⚠️ workflow_dispatch 실패 — deploy.yml에 workflow_dispatch 트리거가 있는지 확인');
    Logger.log(res.getContentText());
  } else {
    Logger.log('⚠️ Actions 트리거 응답 (HTTP ' + code + '): ' + res.getContentText());
  }
}

// ════════════════════════════════════════
//  🔧  설정 & 유틸리티
// ════════════════════════════════════════

/**
 * 트리거 등록 (최초 1회 실행)
 */
function setupTrigger() {
  // 기존 트리거 삭제
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onSheetEdit') ScriptApp.deleteTrigger(t);
  });

  // 편집 트리거 등록
  ScriptApp.newTrigger('onSheetEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert('✅ 트리거 등록 완료!\n이제 시트 수정 시 자동으로 GitHub에 반영됩니다.');
}

/**
 * 커스텀 메뉴 추가
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📚 교과서 관리')
    .addItem('GitHub에 지금 배포', 'manualSync')
    .addSeparator()
    .addItem('자동 트리거 설정', 'setupTrigger')
    .addItem('GitHub Token 설정 안내', 'showTokenGuide')
    .addToUi();
}

/**
 * GitHub Token 설정 안내
 */
function showTokenGuide() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    '🔑 GitHub Token 설정 방법',
    '1. GitHub → Settings → Developer settings → Personal access tokens → Fine-grained tokens\n' +
    '2. "Generate new token" 클릭\n' +
    '3. Repository: yp3000yp/bookinfo 선택\n' +
    '4. Permissions: Contents → Read and write\n' +
    '5. 토큰 복사\n\n' +
    '6. Apps Script 에디터 → 프로젝트 설정(⚙️)\n' +
    '7. 스크립트 속성 → 속성 추가\n' +
    '   키: GITHUB_TOKEN\n' +
    '   값: (복사한 토큰)',
    ui.ButtonSet.OK
  );
}

// ════════════════════════════════════════
//  🔬  진단 함수 (문제 해결용)
//  Apps Script 에디터에서 아래 함수를 선택 후 ▶ 실행
//  실행 후 왼쪽 메뉴 "실행 로그" 클릭해서 결과 확인
// ════════════════════════════════════════

/**
 * 진단 1 — 토큰 & GitHub 연결 확인
 * 실행 후 로그에서 HTTP 상태 코드 확인
 */
function diag_1_checkToken() {
  Logger.log('=== 진단 1: GitHub Token 및 연결 확인 ===');

  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');

  if (!token) {
    Logger.log('❌ GITHUB_TOKEN 없음! 스크립트 속성에 토큰을 추가해야 합니다.');
    return;
  }
  Logger.log('✅ GITHUB_TOKEN 존재 (길이: ' + token.length + '자)');
  Logger.log('   토큰 앞 4자리: ' + token.substring(0, 4) + '...');

  // GitHub API 연결 테스트
  const apiUrl = `https://api.github.com/repos/${CONFIG.GITHUB_OWNER}/${CONFIG.GITHUB_REPO}/contents/${CONFIG.DATA_FILE_PATH}`;
  Logger.log('\n요청 URL: ' + apiUrl);

  try {
    const res = UrlFetchApp.fetch(apiUrl, {
      headers: {
        'Authorization': 'token ' + token,
        'Accept': 'application/vnd.github.v3+json'
      },
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    Logger.log('HTTP 응답 코드: ' + code);

    if (code === 200) {
      const file = JSON.parse(res.getContentText());
      Logger.log('✅ data.json 파일 확인됨');
      Logger.log('   파일 크기: ' + file.size + ' bytes');
      Logger.log('   SHA: ' + file.sha);
      Logger.log('   마지막 수정: ' + (file.encoding || ''));
    } else if (code === 401) {
      Logger.log('❌ 인증 실패 — 토큰이 만료되었거나 잘못됨');
      Logger.log('   응답: ' + res.getContentText());
    } else if (code === 403) {
      Logger.log('❌ 권한 없음 — 토큰에 Contents 쓰기 권한 필요');
      Logger.log('   응답: ' + res.getContentText());
    } else if (code === 404) {
      Logger.log('⚠️ data.json 파일 없음 (신규 생성 예정) — 저장소/경로 확인 필요');
      Logger.log('   현재 설정: ' + CONFIG.GITHUB_OWNER + '/' + CONFIG.GITHUB_REPO);
    } else {
      Logger.log('⚠️ 예상치 못한 응답: ' + res.getContentText());
    }
  } catch(e) {
    Logger.log('❌ 네트워크 오류: ' + e.message);
  }
}

/**
 * 진단 2 — 시트 파싱 결과 확인
 * 실행 후 로그에서 JSON 구조 및 수정한 값이 포함됐는지 확인
 */
function diag_2_checkSheetParsing() {
  Logger.log('=== 진단 2: 시트 파싱 결과 확인 ===');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const sheetNames = sheets.map(s => s.getName());
  Logger.log('전체 시트 목록: ' + sheetNames.join(', '));

  // 연도 시트만 필터
  const yearSheets = sheets.filter(s => /^\d{4}$/.test(s.getName()));
  Logger.log('감지된 연도 시트: ' + yearSheets.map(s => s.getName()).join(', '));

  if (yearSheets.length === 0) {
    Logger.log('❌ 연도 시트 없음! 시트 이름이 2026 같은 4자리 숫자여야 합니다.');
    return;
  }

  yearSheets.forEach(sheet => {
    Logger.log('\n--- [' + sheet.getName() + '] 시트 파싱 ---');
    try {
      const result = parseYearSheet(sheet);
      Logger.log('학교 수: ' + result.schools.length);
      Logger.log('학교 목록: ' + result.schools.join(', '));
      Logger.log('학년 목록: ' + Object.keys(result.grades).join(', '));

      // 각 학년의 과목/행 수 및 샘플값 출력
      Object.entries(result.grades).forEach(([grade, data]) => {
        Logger.log('\n  [' + grade + '] 과목 수: ' + data.subjects.length);
        data.subjects.forEach((subj, i) => {
          // 각 과목의 첫 번째 학교 값 출력
          const sample = data.rows[i] ? data.rows[i][0] : '(없음)';
          Logger.log('    ' + subj + ' → 첫 학교: "' + sample + '"');
        });
      });

      // 수정한 값 검색 (한문 관련)
      Logger.log('\n  [한문 관련 값 검색]');
      let found = false;
      Object.entries(result.grades).forEach(([grade, data]) => {
        data.rows.forEach((row, ri) => {
          row.forEach((val, ci) => {
            if (val && val.includes('한문')) {
              Logger.log('  발견: ' + grade + ' > ' + data.subjects[ri] + ' > ' + result.schools[ci] + ' = "' + val + '"');
              found = true;
            }
          });
        });
      });
      if (!found) Logger.log('  한문 관련 값 없음');

    } catch(e) {
      Logger.log('❌ 파싱 오류: ' + e.message);
    }
  });
}

/**
 * 진단 3 — data.json 생성 후 GitHub PUT 전체 테스트
 * 실제로 GitHub에 커밋까지 실행 (반영 여부 최종 확인)
 */
function diag_3_fullTest() {
  Logger.log('=== 진단 3: 전체 동기화 테스트 ===');

  try {
    Logger.log('1. JSON 빌드 시작...');
    const json = buildDataJson();
    const jsonStr = JSON.stringify(json, null, 2);
    Logger.log('✅ JSON 빌드 완료 (' + jsonStr.length + ' bytes)');

    // 생성된 JSON에서 한문 값 확인
    Logger.log('\n[생성된 JSON에서 한문 검색]');
    const yearKeys = Object.keys(json.years);
    yearKeys.forEach(yr => {
      const yData = json.years[yr];
      Object.entries(yData.grades).forEach(([grade, gData]) => {
        gData.rows.forEach((row, ri) => {
          row.forEach(val => {
            if (val && val.includes('한문')) {
              Logger.log('  ' + yr + ' > ' + grade + ' > ' + gData.subjects[ri] + ' = "' + val + '"');
            }
          });
        });
      });
    });

    Logger.log('\n2. GitHub 업로드 시작...');
    pushToGitHub(jsonStr);
    Logger.log('✅ GitHub 업로드 완료!');
    Logger.log('→ 1~2분 후 https://' + CONFIG.GITHUB_OWNER + '.github.io/' + CONFIG.GITHUB_REPO + '/ 확인');

  } catch(e) {
    Logger.log('❌ 오류 발생: ' + e.message);
    Logger.log('스택: ' + e.stack);
  }
}

/**
 * 진단 4 — 현재 GitHub의 data.json 내용 직접 확인
 * 웹사이트에 실제 올라간 데이터가 무엇인지 확인
 */
function diag_4_readCurrentJson() {
  Logger.log('=== 진단 4: GitHub 현재 data.json 내용 확인 ===');

  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) { Logger.log('❌ GITHUB_TOKEN 없음'); return; }

  const apiUrl = `https://api.github.com/repos/${CONFIG.GITHUB_OWNER}/${CONFIG.GITHUB_REPO}/contents/${CONFIG.DATA_FILE_PATH}`;

  try {
    const res = UrlFetchApp.fetch(apiUrl, {
      headers: {
        'Authorization': 'token ' + token,
        'Accept': 'application/vnd.github.v3+json'
      },
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      Logger.log('❌ 파일 읽기 실패 (HTTP ' + res.getResponseCode() + ')');
      return;
    }

    const file = JSON.parse(res.getContentText());
    const decoded = Utilities.newBlob(
      Utilities.base64Decode(file.content.replace(/\n/g, '')),
      'application/json'
    ).getDataAsString('UTF-8');

    Logger.log('GitHub data.json 크기: ' + decoded.length + ' bytes');
    Logger.log('마지막 500자:\n' + decoded.slice(-500));

    // 한문 관련 검색
    Logger.log('\n[한문 포함 라인 검색]');
    const lines = decoded.split('\n');
    lines.forEach((line, i) => {
      if (line.includes('한문')) {
        Logger.log('  라인 ' + (i+1) + ': ' + line.trim());
      }
    });

  } catch(e) {
    Logger.log('❌ 오류: ' + e.message);
  }
}
