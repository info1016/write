<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>정보과 세특 생성기 V8 (Gemini AI 탑재)</title>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700&family=Black+Han+Sans&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  :root {
    --bg: #f8f9fa; --surface: #ffffff; --sidebar-bg: #1e293b; --border: #e2e8f0;
    --accent: #3b82f6; --accent-soft: #dbeafe; --text: #1e293b; --text-muted: #64748b;
    --success: #10b981; --warn: #f59e0b; --danger: #ef4444; --purple: #8b5cf6;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Noto Sans KR', sans-serif; background: var(--bg); color: var(--text); display: flex; flex-direction: column; height: 100vh; overflow: hidden; }
  
  header { background: #ffffff; border-bottom: 1px solid var(--border); padding: 12px 24px; display: flex; align-items: center; justify-content: space-between; z-index: 10; }
  .logo { font-family: 'Black Han Sans', sans-serif; font-size: 20px; color: var(--accent); cursor: pointer; user-select: none; }
  .logo span { color: var(--text); }

  .app-container { display: flex; flex: 1; overflow: hidden; }
  .sidebar { width: 280px; background: var(--sidebar-bg); color: white; display: flex; flex-direction: column; padding: 20px 12px; }
  .class-selector select { width: 100%; padding: 10px; border-radius: 8px; background: #334155; color: white; border: none; font-family: inherit; outline: none; margin-bottom: 20px; }
  .student-list-wrap { flex: 1; overflow-y: auto; }
  .student-item { padding: 12px 16px; margin-bottom: 4px; border-radius: 8px; cursor: pointer; display: flex; align-items: center; justify-content: space-between; font-size: 14px; transition: background 0.2s;}
  .student-item:hover { background: #334155; }
  .student-item.active { background: var(--accent); color: white; }
  .student-item .status-icon { font-size: 12px; opacity: 0.5; }
  .student-item.done .status-icon { opacity: 1; color: var(--success); }

  .main-content { flex: 1; display: flex; flex-direction: column; overflow-y: auto; padding: 24px; gap: 24px; padding-bottom: 280px; }
  .eval-section { background: var(--surface); border-radius: 12px; border: 1px solid var(--border); padding: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
  .section-title { font-size: 16px; font-weight: 700; margin-bottom: 16px; }
  .rubric-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 10px; }
  .rubric-card { border: 1.5px solid var(--border); border-radius: 10px; padding: 12px; cursor: pointer; font-size: 13.5px; line-height: 1.5; background: #fff; display: flex; align-items: flex-start; gap: 10px; transition: all 0.2s;}
  .rubric-card:hover { border-color: var(--accent); background: var(--bg); }
  .rubric-card.selected { border-color: var(--accent); background: var(--accent-soft); color: var(--accent); font-weight: 500; }
  .rubric-card .check-circle { width: 18px; height: 18px; border-radius: 50%; border: 2px solid var(--border); flex-shrink: 0; margin-top: 2px; }
  .rubric-card.selected .check-circle { background: var(--accent); border-color: var(--accent); }

  .result-area { position: fixed; bottom: 0; right: 0; width: calc(100% - 280px); background: #fff; border-top: 3px solid var(--accent); padding: 20px 24px; display: flex; flex-direction: column; gap: 12px; box-shadow: 0 -10px 25px rgba(0,0,0,0.1); z-index: 20; }
  .keyword-input-wrap { display: flex; align-items: center; gap: 12px; background: var(--bg); padding: 10px 16px; border-radius: 8px; border: 1px solid var(--border); }
  .keyword-input-wrap input { flex: 1; border: none; background: transparent; font-size: 14px; outline: none; }
  .result-header { display: flex; justify-content: space-between; align-items: center; }
  .result-text-label { font-size: 15px; font-weight: 700; color: var(--accent); }
  #generatedText { width: 100%; height: 110px; border: 1px solid var(--border); border-radius: 8px; padding: 16px; font-size: 14px; line-height: 1.8; background: var(--bg); color: var(--text); resize: none; outline: none; }
  #generatedText.loading { animation: pulse 1.5s infinite; background: #f1f5f9; color: #94a3b8; }
  @keyframes pulse { 0% { opacity: 0.6; } 50% { opacity: 1; } 100% { opacity: 0.6; } }

  .btn { padding: 8px 16px; border-radius: 8px; border: none; font-weight: 500; cursor: pointer; display: inline-flex; align-items: center; justify-content: center; gap: 6px; font-size: 14px; transition: all 0.2s;}
  .btn-primary { background: var(--accent); color: white; }
  .btn-success { background: var(--success); color: white; }
  .btn-outline { background: transparent; border: 1px solid var(--border); color: var(--text-muted); }
  .btn-warn { background: var(--warn); color: white; }
  .btn-purple { background: var(--purple); color: white; font-weight: 700; }
  .btn-purple:hover { background: #7c3aed; box-shadow: 0 4px 12px rgba(139,92,246,0.3); }

  /* Admin Panel */
  .admin-only { display: none !important; }
  body.admin-mode .admin-only { display: block !important; }
  .admin-panel { position: fixed; inset: 0; background: rgba(15, 23, 42, 0.85); z-index: 100; display: none; align-items: center; justify-content: center; padding: 20px; }
  .admin-content { background: #f8f9fa; width: 100%; max-width: 850px; max-height: 90vh; border-radius: 16px; display: flex; flex-direction: column; overflow: hidden; box-shadow: 0 20px 25px -5px rgba(0,0,0,0.2); }
  .admin-header { padding: 20px 24px; background: white; border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center; }
  .admin-body { padding: 24px; overflow-y: auto; display: flex; flex-direction: column; gap: 24px; }
  .admin-section { background: white; padding: 20px; border-radius: 12px; border: 1px solid var(--border); }
  .api-key-box { background: rgba(139,92,246,0.1); border: 1px solid rgba(139,92,246,0.3); padding: 16px; border-radius: 8px; margin-bottom: 20px;}
  .api-key-box input { width: 100%; padding: 10px; border-radius: 6px; border: 1px solid var(--border); margin-top: 8px; outline: none;}
  .api-key-box input:focus { border-color: var(--purple); }
  .file-upload-wrap { display: flex; align-items: center; gap: 10px; margin-bottom: 10px; background: var(--bg); padding: 10px; border-radius: 8px; border: 1px dashed var(--text-muted); }
  textarea#adminNameInput { width: 100%; height: 100px; padding: 12px; border-radius: 8px; border: 1px solid var(--border); outline: none; }
  .edit-category { border: 1px solid var(--border); border-radius: 8px; padding: 16px; margin-bottom: 16px; background: var(--bg); }
  .edit-category-header { display: flex; gap: 10px; margin-bottom: 12px; }
  .edit-category-header input { flex: 1; padding: 8px 12px; border-radius: 6px; border: 1px solid var(--border); font-weight: bold; }
  .edit-item { display: flex; gap: 10px; margin-bottom: 8px; align-items: center; }
  .edit-item input.item-label { flex: 1; padding: 8px; border-radius: 6px; border: 1px solid var(--border); }
</style>
</head>
<body>

<header>
  <div class="logo" id="logo" title="3번 클릭 시 관리자 모드 진입">세특<span>생성기</span> <small style="font-size: 10px; color: var(--text-muted);">V8 AI</small></div>
  <div class="header-actions">
    <button class="btn btn-outline admin-only" id="exportBtn">📊 엑셀 다운로드</button>
    <button class="btn btn-warn admin-only" id="exitAdminBtn" style="margin-left: 8px;">🚪 관리자 모드 종료</button>
    <span id="activeStudentInfo" style="font-size: 14px; font-weight: 500; margin-left: 15px; color: var(--text-muted);">학생을 선택하세요</span>
  </div>
</header>

<div class="app-container">
  <aside class="sidebar">
    <div class="class-selector">
      <select id="classSelect">
        <option value="1">1학년 1반</option><option value="2">1학년 2반</option><option value="3">1학년 3반</option>
      </select>
    </div>
    <div class="student-list-wrap" id="studentList"></div>
  </aside>

  <main class="main-content" id="mainContent"></main>

  <div class="result-area">
    <div class="keyword-input-wrap">
      <span style="font-size:16px;">📌</span>
      <label style="font-size: 13px; font-weight:bold; color:var(--text-muted); white-space:nowrap;">학생 개별 특징/키워드 :</label>
      <input type="text" id="studentKeyword" placeholder="예: '파이썬으로 방탈출 게임 제작', '논리적인 오류 수정 능력' 등 자유롭게 입력">
    </div>
    <div class="result-header">
      <div class="result-text-label">NEIS 세특 생성 문구</div>
      <div class="action-buttons">
        <button class="btn btn-purple" id="aiGenerateBtn">🤖 AI 자동 창작</button>
        <button class="btn btn-primary" id="copyBtn">📋 복사</button>
        <button class="btn btn-success" id="saveBtn">💾 완료</button>
      </div>
    </div>
    <textarea id="generatedText" spellcheck="false" placeholder="항목 체크 후 [🤖 AI 자동 창작] 버튼을 누르면 인공지능이 완벽한 문맥으로 세특을 써줍니다. (직접 수정도 가능)"></textarea>
  </div>
</div>

<div class="admin-panel" id="adminPanel">
  <div class="admin-content">
    <div class="admin-header">
      <h2 style="color: var(--sidebar-bg); font-size: 18px;">⚙️ 관리자 설정 (API 연동 및 명단)</h2>
      <div>
        <button class="btn btn-primary" id="adminSaveAllBtn">💾 모든 변경사항 저장</button>
        <button class="btn btn-outline" id="adminCloseBtn">닫기</button>
      </div>
    </div>
    <div class="admin-body">
      <div class="api-key-box">
        <h3 style="color: var(--purple); font-size: 15px; margin-bottom: 4px;">🔑 Google Gemini API Key</h3>
        <p style="font-size: 12px; color: var(--text-muted); margin-bottom: 8px;">AI 창작 기능을 사용하려면 Google AI Studio에서 발급받은 API 키가 필요합니다. (브라우저에 안전하게 암호화 저장됨)</p>
        <input type="password" id="apiKeyInput" placeholder="AIzaSy... 형식의 API 키를 여기에 붙여넣으세요">
      </div>

      <div class="admin-section">
        <h3>1. 학생 명단 입력 (<span id="targetClassLabel"></span>반)</h3>
        <div class="file-upload-wrap">
          <label style="font-size: 13px; font-weight: bold;">엑셀 업로드 (.xlsx, .csv)</label>
          <input type="file" id="excelUpload" accept=".xlsx, .xls, .csv" style="font-size: 13px;">
        </div>
        <textarea id="adminNameInput" placeholder="직접 입력 시 줄바꿈(Enter)으로 구분하세요."></textarea>
      </div>
      <div class="admin-section">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
          <h3>2. 평가 주제 및 항목명 설정</h3>
          <button class="btn btn-success" style="font-size: 12px;" onclick="addCategory()">+ 새 평가 주제 추가</button>
        </div>
        <p style="font-size: 13px; color: var(--text-muted); margin-bottom: 15px;">* AI가 항목 이름(버튼명)만 보고 문장을 창작하므로, 항목 이름만 직관적으로 적어주시면 됩니다.</p>
        <div id="rubricEditorContainer"></div>
      </div>
    </div>
  </div>
</div>

<script>
  // ── AI에 맞춘 심플해진 기본 데이터 (문장 템플릿 불필요) ──
  const defaultRubrics = [
    {
      title: "핵심 개념 이해",
      items: [
        { label: "변수와 조건문의 개념 완벽 이해" },
        { label: "변수와 조건문의 기본 원리 파악" },
        { label: "알고리즘의 논리적 구조 설계 우수" }
      ]
    },
    {
      title: "문제 해결 및 실습 과정",
      items: [
        { label: "스스로 오류를 분석하고 디버깅함" },
        { label: "적극적인 질문으로 문제 해결" },
        { label: "창의적인 기능을 스스로 추가함" }
      ]
    },
    {
      title: "참여 태도 및 정보 윤리",
      items: [
        { label: "모둠 실습에서 주도적 리더십 발휘" },
        { label: "수업 집중도가 높고 끈기 있게 탐구함" },
        { label: "정보 윤리를 준수하고 과제를 성실히 수행함" }
      ]
    }
  ];

  let studentsData = {}; let rubricData = []; let currentClass = "1"; let activeStudentIdx = null; let geminiApiKey = "";

  window.addEventListener('DOMContentLoaded', () => {
    // 저장된 데이터 로드
    const savedStudents = localStorage.getItem('teacher_eval_v8_students');
    if (savedStudents) studentsData = JSON.parse(savedStudents);
    else for(let i=1; i<=3; i++) studentsData[i] = [];

    const savedRubrics = localStorage.getItem('teacher_eval_v8_rubrics');
    if (savedRubrics) rubricData = JSON.parse(savedRubrics);
    else rubricData = JSON.parse(JSON.stringify(defaultRubrics));

    geminiApiKey = localStorage.getItem('teacher_eval_v8_apikey') || "";

    renderMainRubrics(); renderStudentList();
  });

  // ── 메인 화면 렌더링 ──
  function renderMainRubrics() {
    const mainContent = document.getElementById('mainContent'); mainContent.innerHTML = '';
    rubricData.forEach((category, cIdx) => {
      const section = document.createElement('section'); section.className = 'eval-section';
      let html = `<div class="section-title">📘 ${category.title}</div><div class="rubric-grid" data-category-idx="${cIdx}">`;
      category.items.forEach(item => {
        html += `<div class="rubric-card" data-label="${item.label}"><div class="check-circle"></div><div>${item.label}</div></div>`;
      });
      html += `</div>`; section.innerHTML = html; mainContent.appendChild(section);
    });

    document.querySelectorAll('.rubric-card').forEach(card => {
      card.onclick = function() {
        if (activeStudentIdx === null) return alert('학생을 먼저 선택하세요!');
        const parentGrid = this.closest('.rubric-grid');
        const wasSelected = this.classList.contains('selected');
        parentGrid.querySelectorAll('.rubric-card').forEach(c => c.classList.remove('selected'));
        if (!wasSelected) this.classList.add('selected');
        
        // 클릭 시 바로 문장을 만들지 않고, 어떤 버튼을 눌렀는지만 데이터로 임시 저장 (수동 연결 로직 대체)
        saveCurrentSelection();
      };
    });
  }

  function renderStudentList() {
    const listWrap = document.getElementById('studentList'); listWrap.innerHTML = '';
    const students = studentsData[currentClass] || [];
    students.forEach((s, idx) => {
      const div = document.createElement('div');
      div.className = `student-item ${idx === activeStudentIdx ? 'active' : ''} ${s.done ? 'done' : ''}`;
      div.innerHTML = `<span>${s.name}</span> <span class="status-icon">${s.done ? '✅' : '●'}</span>`;
      div.onclick = () => selectStudent(idx);
      listWrap.appendChild(div);
    });
  }

  function selectStudent(idx) {
    activeStudentIdx = idx; const student = studentsData[currentClass][idx];
    document.getElementById('activeStudentInfo').textContent = `${currentClass}반 ${student.name} 학생`;
    document.getElementById('activeStudentInfo').style.color = "var(--accent)";
    document.getElementById('studentKeyword').value = student.keyword || "";

    // 화면 초기화 후 선택 내역 복원
    document.querySelectorAll('.rubric-card').forEach(card => card.classList.remove('selected'));
    if (student.eval && student.eval.length > 0) {
      student.eval.forEach(label => {
        const card = Array.from(document.querySelectorAll('.rubric-card')).find(c => c.dataset.label === label);
        if(card) card.classList.add('selected');
      });
    }
    document.getElementById('generatedText').value = student.text || "";
    renderStudentList();
  }

  function saveCurrentSelection() {
    if (activeStudentIdx === null) return;
    const selectedCards = document.querySelectorAll('.rubric-card.selected');
    let selectedEvals = [];
    selectedCards.forEach(c => selectedEvals.push(c.dataset.label));
    studentsData[currentClass][activeStudentIdx].eval = selectedEvals;
    studentsData[currentClass][activeStudentIdx].keyword = document.getElementById('studentKeyword').value;
  }

  // ── 핵심: Gemini AI 문장 생성 로직 ──
  document.getElementById('aiGenerateBtn').onclick = async () => {
    if (activeStudentIdx === null) return alert('학생을 먼저 선택해주세요.');
    if (!geminiApiKey) {
      alert("API 키가 없습니다! 왼쪽 위 '세특생성기' 로고를 3번 클릭하여 관리자 모드에서 API 키를 먼저 입력해주세요.");
      return;
    }

    const selectedCards = document.querySelectorAll('.rubric-card.selected');
    if (selectedCards.length === 0) return alert('평가 항목을 최소 1개 이상 체크해주세요.');

    const labels = Array.from(selectedCards).map(c => c.dataset.label).join(', ');
    const keyword = document.getElementById('studentKeyword').value.trim();
    
    // AI에게 내리는 프롬프트 (명령어)
    const prompt = `
너는 대한민국 중학교 정보 과목의 전문성 있는 10년 차 교사야.
다음 [평가 항목]과 학생의 [개별 특징]을 바탕으로, 학교생활기록부(NEIS) 세부능력 및 특기사항에 바로 복사해서 붙여넣을 수 있는 완성된 글을 작성해줘.

[평가 항목]: ${labels}
[개별 특징/키워드]: ${keyword ? keyword : '특별한 키워드 없음'}

<작성 조건>
1. 길이는 2~3문장으로 간결하고 전문적으로 작성할 것.
2. 반드시 '~함.', '~임.', '~보임.' 등 명사형 종결어미로 모든 문장을 끝낼 것.
3. 접속사('또한', '그리고', '게다가' 등)의 사용을 최대한 자제하고, 문장과 문장을 자연스럽게 이어갈 것. (예: '~를 이해하며, ~하는 능력이 뛰어남.')
4. 학생의 [개별 특징]이 주어졌다면 어색하지 않게 문맥 속에 잘 녹여낼 것.
5. "세특 문구를 작성했습니다" 같은 불필요한 인사말이나 부연 설명 없이, 오직 세특에 들어갈 본문 텍스트만 출력할 것.
`;

    const textArea = document.getElementById('generatedText');
    textArea.value = "인공지능이 세특을 정성스럽게 작성 중입니다... (약 3~5초 소요)";
    textArea.classList.add('loading');

    try {
      // Gemini 1.5 Flash 모델 호출 (속도와 성능 밸런스 최적)
      const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${geminiApiKey}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }]
        })
      });

      if (!response.ok) throw new Error('API 호출 실패. 키가 유효한지 확인해주세요.');

      const data = await response.json();
      let aiText = data.candidates[0].content.parts[0].text.trim();
      
      textArea.classList.remove('loading');
      textArea.value = aiText;

      // 결과 저장
      studentsData[currentClass][activeStudentIdx].text = aiText;
      saveToLocal();

    } catch (error) {
      textArea.classList.remove('loading');
      textArea.value = "⚠️ 오류가 발생했습니다. API 키가 정확한지, 인터넷이 연결되어 있는지 확인해주세요.\n\n" + error;
    }
  };

  // ── 저장 및 기타 버튼 ──
  document.getElementById('saveBtn').onclick = () => {
    if (activeStudentIdx === null) return;
    saveCurrentSelection(); // 키워드 및 선택사항 저장
    studentsData[currentClass][activeStudentIdx].text = document.getElementById('generatedText').value;
    studentsData[currentClass][activeStudentIdx].done = true;
    saveToLocal(); renderStudentList(); alert('현재 학생 데이터가 저장되었습니다.');
  };

  document.getElementById('copyBtn').onclick = () => {
    const t = document.getElementById('generatedText').value;
    if(t) navigator.clipboard.writeText(t).then(() => alert('세특 문구가 복사되었습니다.'));
  };

  function saveToLocal() {
    localStorage.setItem('teacher_eval_v8_students', JSON.stringify(studentsData));
    localStorage.setItem('teacher_eval_v8_rubrics', JSON.stringify(rubricData));
    localStorage.setItem('teacher_eval_v8_apikey', geminiApiKey);
  }

  document.getElementById('classSelect').onchange = function() {
    currentClass = this.value; activeStudentIdx = null;
    document.getElementById('generatedText').value = ""; document.getElementById('studentKeyword').value = "";
    document.getElementById('activeStudentInfo').textContent = "학생을 선택하세요";
    document.getElementById('activeStudentInfo').style.color = "var(--text-muted)";
    renderStudentList();
  };

  document.getElementById('exportBtn').onclick = () => {
    let csvContent = "\uFEFF학년반,이름,개별키워드,세부능력 및 특기사항,평가여부\n";
    for (const cls in studentsData) {
      studentsData[cls].forEach(s => {
        csvContent += `${cls}반,${s.name},"${s.keyword||''}","${(s.text||'').replace(/"/g, '""')}",${s.done ? "완료" : "미완료"}\n`;
      });
    }
    const link = document.createElement("a");
    link.href = URL.createObjectURL(new Blob([csvContent], { type: 'text/csv;charset=utf-8;' }));
    link.download = `AI_세특_일괄데이터_${new Date().toLocaleDateString()}.csv`; link.click();
  };

  // ── 관리자 모드 ──
  let logoClicks = 0;
  document.getElementById('logo').onclick = () => {
    logoClicks++;
    if (logoClicks >= 3) {
      document.body.classList.add('admin-mode');
      document.getElementById('adminPanel').style.display = 'flex';
      document.getElementById('targetClassLabel').textContent = currentClass;
      document.getElementById('adminNameInput').value = (studentsData[currentClass] || []).map(s => s.name).join('\n');
      document.getElementById('excelUpload').value = "";
      document.getElementById('apiKeyInput').value = geminiApiKey; // 키 로드
      renderRubricEditor();
      logoClicks = 0;
    }
    setTimeout(() => logoClicks = 0, 1000);
  };
  
  function exitAdminMode() { document.body.classList.remove('admin-mode'); document.getElementById('adminPanel').style.display = 'none'; }
  document.getElementById('exitAdminBtn').onclick = exitAdminMode; document.getElementById('adminCloseBtn').onclick = exitAdminMode;

  document.getElementById('excelUpload').addEventListener('change', function(e) {
    const file = e.target.files[0]; if(!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
      const workbook = XLSX.read(new Uint8Array(e.target.result), {type: 'array'});
      const json = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {header: 1});
      const names = json.map(row => row[0]).filter(n => n && String(n).trim() !== "");
      if(names.length > 0) { document.getElementById('adminNameInput').value = names.join('\n'); alert(`데이터를 불러왔습니다.`); }
    };
    reader.readAsArrayBuffer(file);
  });

  document.getElementById('adminSaveAllBtn').onclick = () => {
    // API 키 저장
    geminiApiKey = document.getElementById('apiKeyInput').value.trim();

    const names = document.getElementById('adminNameInput').value.split('\n').map(n => n.trim()).filter(Boolean);
    const newStudentArray = [];
    names.forEach(name => {
      const existing = (studentsData[currentClass] || []).find(s => s.name === name);
      if(existing) newStudentArray.push(existing);
      else newStudentArray.push({name, eval:[], text:"", keyword:"", done:false});
    });
    studentsData[currentClass] = newStudentArray;

    const newRubrics = [];
    document.querySelectorAll('.edit-category').forEach(catDiv => {
      const title = catDiv.querySelector('.cat-title').value.trim();
      if(!title) return;
      const items = [];
      catDiv.querySelectorAll('.edit-item').forEach(itemDiv => {
        const label = itemDiv.querySelector('.item-label').value.trim();
        if(label) items.push({label}); // 텍스트는 이제 필요 없음. 라벨(버튼명)만 저장.
      });
      newRubrics.push({title, items});
    });
    if(newRubrics.length === 0) return alert('평가 주제가 1개 이상 필요합니다.');
    
    rubricData = newRubrics; saveToLocal(); renderMainRubrics(); renderStudentList();
    alert('설정이 저장되었습니다.'); exitAdminMode();
  };

  function renderRubricEditor() {
    const container = document.getElementById('rubricEditorContainer'); container.innerHTML = '';
    rubricData.forEach(category => {
      const catDiv = document.createElement('div'); catDiv.className = 'edit-category';
      let html = `<div class="edit-category-header"><input type="text" class="cat-title" value="${category.title}"><button class="btn btn-danger btn-sm" onclick="this.closest('.edit-category').remove()">주제 삭제</button></div><div class="item-list">`;
      category.items.forEach(item => html += createItemHTML(item.label));
      html += `</div><button class="btn btn-outline btn-sm" style="margin-top:8px; width:100%;" onclick="addItemToCategory(this)">+ 항목 추가</button>`;
      catDiv.innerHTML = html; container.appendChild(catDiv);
    });
  }
  
  // AI 연동 버전에서는 버튼명(label)만 관리하면 됨! (템플릿 텍스트 입력창 삭제됨)
  function createItemHTML(label = "") {
    return `<div class="edit-item"><input type="text" class="item-label" value="${label}" placeholder="평가 키워드 입력 (예: 알고리즘 설계 우수)"><button class="btn btn-outline" style="padding: 8px; border-color: var(--danger); color: var(--danger);" onclick="this.closest('.edit-item').remove()">X</button></div>`;
  }
  window.addItemToCategory = btn => btn.previousElementSibling.insertAdjacentHTML('beforeend', createItemHTML());
  window.addCategory = () => {
    const catDiv = document.createElement('div'); catDiv.className = 'edit-category';
    catDiv.innerHTML = `<div class="edit-category-header"><input type="text" class="cat-title" placeholder="새 평가 주제"><button class="btn btn-danger btn-sm" onclick="this.closest('.edit-category').remove()">삭제</button></div><div class="item-list">${createItemHTML()}</div><button class="btn btn-outline btn-sm" style="margin-top:8px; width:100%;" onclick="addItemToCategory(this)">+ 항목 추가</button>`;
    document.getElementById('rubricEditorContainer').appendChild(catDiv);
  };
</script>
</body>
</html>
