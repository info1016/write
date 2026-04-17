<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<title>정보과 세특 생성기 V9 (AI & 레이아웃 최적화)</title>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700&family=Black+Han+Sans&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  :root {
    --bg: #f0f2f5; --surface: #ffffff; --sidebar-bg: #1a237e; --border: #cfd8dc;
    --accent: #2962ff; --accent-soft: #e3f2fd; --text: #263238; --text-muted: #78909c;
    --success: #00c853; --warn: #ffd600; --danger: #ff1744; --purple: #6200ea;
  }
  
  * { box-sizing: border-box; margin: 0; padding: 0; -webkit-font-smoothing: antialiased; }
  body { font-family: 'Noto Sans KR', sans-serif; background: var(--bg); color: var(--text); height: 100vh; display: flex; flex-direction: column; overflow: hidden; }

  /* ── Header (상단바) ── */
  header { background: #fff; border-bottom: 1px solid var(--border); padding: 0 20px; height: 60px; display: flex; align-items: center; justify-content: space-between; flex-shrink: 0; z-index: 100; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
  .logo { font-family: 'Black Han Sans', sans-serif; font-size: 22px; color: var(--accent); cursor: pointer; user-select: none; }
  .logo span { color: var(--text); }
  .active-info { font-size: 14px; font-weight: 500; color: var(--text-muted); background: #f8f9fa; padding: 6px 12px; border-radius: 20px; }

  /* ── Layout (메인 컨테이너) ── */
  .app-container { display: flex; flex: 1; height: calc(100vh - 60px); overflow: hidden; }

  /* Sidebar (왼쪽 명단) */
  .sidebar { width: 260px; background: var(--sidebar-bg); color: white; display: flex; flex-direction: column; padding: 15px; flex-shrink: 0; }
  .class-selector select { width: 100%; padding: 12px; border-radius: 8px; background: rgba(255,255,255,0.1); color: white; border: 1px solid rgba(255,255,255,0.2); font-size: 15px; margin-bottom: 15px; outline: none; }
  .student-list-wrap { flex: 1; overflow-y: auto; padding-right: 5px; }
  .student-item { padding: 12px 15px; margin-bottom: 6px; border-radius: 8px; cursor: pointer; display: flex; align-items: center; justify-content: space-between; font-size: 14px; transition: all 0.2s; background: rgba(255,255,255,0.05); }
  .student-item:hover { background: rgba(255,255,255,0.15); }
  .student-item.active { background: var(--accent); color: white; box-shadow: 0 4px 8px rgba(0,0,0,0.2); }
  .status-icon { font-size: 12px; opacity: 0.6; }
  .student-item.done .status-icon { opacity: 1; color: #69f0ae; }

  /* Main Content (중앙 평가영역) */
  .main-content { flex: 1; overflow-y: auto; padding: 25px; display: flex; flex-direction: column; gap: 25px; padding-bottom: 320px; /* 미리보기 공간 */ scroll-behavior: smooth; }
  .eval-section { background: var(--surface); border-radius: 15px; border: 1px solid var(--border); padding: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.02); }
  .section-title { font-size: 17px; font-weight: 700; margin-bottom: 15px; color: var(--sidebar-bg); border-left: 4px solid var(--accent); padding-left: 10px; }
  .rubric-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(240px, 1fr)); gap: 12px; }
  .rubric-card { border: 1px solid var(--border); border-radius: 10px; padding: 15px; cursor: pointer; font-size: 14px; line-height: 1.5; background: #fff; display: flex; align-items: flex-start; gap: 12px; transition: all 0.2s; }
  .rubric-card:hover { border-color: var(--accent); transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,0.05); }
  .rubric-card.selected { border-color: var(--accent); background: var(--accent-soft); color: var(--accent); font-weight: 600; box-shadow: inset 0 0 0 1px var(--accent); }
  .check-circle { width: 20px; height: 20px; border-radius: 50%; border: 2px solid var(--border); flex-shrink: 0; margin-top: 2px; position: relative; }
  .rubric-card.selected .check-circle { background: var(--accent); border-color: var(--accent); }
  .rubric-card.selected .check-circle::after { content: '✓'; color: white; position: absolute; font-size: 12px; top: 50%; left: 50%; transform: translate(-50%, -50%); }

  /* ── Result Area (하단 고정바) ── */
  .result-area { position: fixed; bottom: 0; right: 0; width: calc(100% - 260px); background: #fff; border-top: 2px solid var(--accent); padding: 20px 30px 30px; display: flex; flex-direction: column; gap: 15px; box-shadow: 0 -10px 30px rgba(0,0,0,0.1); z-index: 90; }
  .keyword-input-wrap { display: flex; align-items: center; gap: 15px; background: #f1f3f4; padding: 12px 20px; border-radius: 30px; }
  .keyword-input-wrap input { flex: 1; border: none; background: transparent; font-size: 14px; outline: none; font-family: inherit; }
  .result-header { display: flex; justify-content: space-between; align-items: center; }
  .result-text-label { font-size: 15px; font-weight: 700; color: var(--accent); display: flex; align-items: center; gap: 8px; }
  #generatedText { width: 100%; height: 120px; border: 1px solid var(--border); border-radius: 12px; padding: 15px; font-size: 15px; line-height: 1.8; background: #fafafa; color: var(--text); resize: none; outline: none; transition: all 0.3s; }
  #generatedText:focus { background: #fff; border-color: var(--accent); box-shadow: 0 0 0 3px var(--accent-soft); }
  
  /* Loading & Error Styles */
  #generatedText.loading { animation: pulse 1.5s infinite; color: #90a4ae; }
  #generatedText.error { color: var(--danger); border-color: var(--danger); background: #fff1f0; }
  @keyframes pulse { 0% { opacity: 0.6; } 50% { opacity: 1; } 100% { opacity: 0.6; } }

  /* ── Buttons ── */
  .btn { padding: 10px 20px; border-radius: 8px; border: none; font-weight: 600; cursor: pointer; display: inline-flex; align-items: center; justify-content: center; gap: 8px; font-size: 14px; transition: all 0.2s; }
  .btn-primary { background: var(--accent); color: white; }
  .btn-success { background: var(--success); color: white; }
  .btn-outline { background: transparent; border: 1px solid var(--border); color: var(--text-muted); }
  .btn-purple { background: var(--purple); color: white; box-shadow: 0 4px 12px rgba(98,0,234,0.2); }
  .btn:hover { filter: brightness(1.1); transform: translateY(-1px); }
  .btn:active { transform: translateY(0); }

  /* ── Admin Mode & Panel ── */
  .admin-only { display: none !important; }
  body.admin-mode .admin-only { display: inline-flex !important; }
  .admin-panel { position: fixed; inset: 0; background: rgba(0,0,0,0.7); z-index: 200; display: none; align-items: center; justify-content: center; padding: 20px; }
  .admin-content { background: #f8f9fa; width: 100%; max-width: 800px; max-height: 85vh; border-radius: 20px; display: flex; flex-direction: column; overflow: hidden; box-shadow: 0 20px 40px rgba(0,0,0,0.3); }
  .admin-header { padding: 20px 25px; background: white; border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center; }
  .admin-body { padding: 25px; overflow-y: auto; display: flex; flex-direction: column; gap: 25px; }
  .admin-section { background: white; padding: 20px; border-radius: 12px; border: 1px solid var(--border); }
  .api-key-box { background: #fff9c4; border: 1px solid #fbc02d; padding: 15px; border-radius: 10px; margin-bottom: 15px; }
  .api-key-box input { width: 100%; padding: 12px; border-radius: 8px; border: 1px solid #fbc02d; margin-top: 10px; font-family: monospace; outline: none; }
  textarea#adminNameInput { width: 100%; height: 120px; padding: 12px; border-radius: 8px; border: 1px solid var(--border); outline: none; font-size: 14px; }

  /* ── Mobile Layout Fix ── */
  @media (max-width: 900px) {
    .sidebar { width: 80px; padding: 10px 5px; }
    .sidebar span, .class-selector, .status-icon { display: none; }
    .student-item { justify-content: center; padding: 15px 5px; }
    .result-area { width: calc(100% - 80px); padding: 15px; }
  }
</style>
</head>
<body class="admin-mode"> <header>
  <div class="logo" id="logo" title="3번 클릭 시 관리자 모드 진입">세특<span>생성기</span> <small style="font-size: 10px; opacity: 0.5;">V9</small></div>
  <div class="header-actions" style="display: flex; align-items: center; gap: 10px;">
    <span id="activeStudentInfo" class="active-info">학생을 선택하세요</span>
    <button class="btn btn-outline admin-only" id="exportBtn">📊 엑셀 다운로드</button>
    <button class="btn btn-warn admin-only" id="exitAdminBtn">🚪 설정 닫기</button>
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
      <span style="font-size:18px;">💡</span>
      <input type="text" id="studentKeyword" placeholder="이 학생만의 특징을 짧게 적어주세요 (예: '보고서 논리 정연', '알고리즘 최적화 아이디어 뱅크')">
    </div>
    <div class="result-header">
      <div class="result-text-label">✨ 나이스 입력용 세특 본문</div>
      <div class="action-buttons">
        <button class="btn btn-purple" id="aiGenerateBtn">🤖 AI 자동 창작</button>
        <button class="btn btn-primary" id="copyBtn">📋 복사</button>
        <button class="btn btn-success" id="saveBtn">💾 저장</button>
      </div>
    </div>
    <textarea id="generatedText" spellcheck="false" placeholder="항목 체크 후 [🤖 AI 자동 창작] 버튼을 누르세요."></textarea>
  </div>
</div>

<div class="admin-panel" id="adminPanel">
  <div class="admin-content">
    <div class="admin-header">
      <h2 style="font-size: 18px;">⚙️ 프로그램 설정</h2>
      <div style="display: flex; gap: 10px;">
        <button class="btn btn-primary" id="adminSaveAllBtn">💾 설정 저장</button>
        <button class="btn btn-outline" id="adminCloseBtn">취소</button>
      </div>
    </div>
    <div class="admin-body">
      <div class="api-key-box">
        <strong>🔑 구글 Gemini API Key 입력</strong>
        <p style="font-size: 12px; margin-top: 4px; color: #795548;">API 키가 없으면 AI 기능을 쓸 수 없습니다. [Google AI Studio]에서 발급받으세요.</p>
        <input type="password" id="apiKeyInput" placeholder="여기에 API 키를 붙여넣으세요">
      </div>
      <div class="admin-section">
        <h3>1. 학생 명단 관리 (<span id="targetClassLabel"></span>반)</h3>
        <input type="file" id="excelUpload" accept=".xlsx, .xls, .csv" style="margin: 10px 0; font-size: 13px;">
        <textarea id="adminNameInput" placeholder="이름을 엔터로 구분해서 입력"></textarea>
      </div>
      <div class="admin-section">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
          <h3>2. 평가 항목 설정</h3>
          <button class="btn btn-success" style="font-size: 12px;" onclick="addCategory()">+ 주제 추가</button>
        </div>
        <div id="rubricEditorContainer"></div>
      </div>
    </div>
  </div>
</div>

<script>
  // ── 초기 데이터 세팅 ──
  const defaultRubrics = [
    { title: "정보과학 핵심 개념", items: [{label: "변수와 연산의 원리 이해"}, {label: "제어 구조(조건, 반복) 활용 능력 우수"}] },
    { title: "실습 및 문제해결", items: [{label: "오류 발견 및 자가 디버깅 역량"}, {label: "프로그램 요구사항 분석 및 설계 우수"}] },
    { title: "태도 및 소양", items: [{label: "디지털 윤리 및 정보 보호 의식"}, {label: "협력적 문제 해결 참여도 우수"}] }
  ];

  let studentsData = {}; let rubricData = []; let currentClass = "1"; let activeStudentIdx = null; let geminiApiKey = "";

  window.addEventListener('DOMContentLoaded', () => {
    const savedStudents = localStorage.getItem('eval_v9_students');
    studentsData = savedStudents ? JSON.parse(savedStudents) : {"1":[], "2":[], "3":[]};

    const savedRubrics = localStorage.getItem('eval_v9_rubrics');
    rubricData = savedRubrics ? JSON.parse(savedRubrics) : JSON.parse(JSON.stringify(defaultRubrics));

    geminiApiKey = localStorage.getItem('eval_v9_api') || "";
    
    renderMainRubrics(); renderStudentList();
  });

  function renderMainRubrics() {
    const main = document.getElementById('mainContent'); main.innerHTML = '';
    rubricData.forEach((cat, cIdx) => {
      const sec = document.createElement('section'); sec.className = 'eval-section';
      sec.innerHTML = `<div class="section-title">${cat.title}</div><div class="rubric-grid">` + 
        cat.items.map(it => `<div class="rubric-card" data-label="${it.label}"><div class="check-circle"></div><div>${it.label}</div></div>`).join('') + `</div>`;
      main.appendChild(sec);
    });
    document.querySelectorAll('.rubric-card').forEach(card => {
      card.onclick = function() {
        if (activeStudentIdx === null) return alert('명단에서 학생을 먼저 선택하세요!');
        const grid = this.closest('.rubric-grid');
        const isSel = this.classList.contains('selected');
        grid.querySelectorAll('.rubric-card').forEach(c => c.classList.remove('selected'));
        if (!isSel) this.classList.add('selected');
        updateData();
      };
    });
  }

  function renderStudentList() {
    const list = document.getElementById('studentList'); list.innerHTML = '';
    (studentsData[currentClass] || []).forEach((s, idx) => {
      const item = document.createElement('div');
      item.className = `student-item ${idx === activeStudentIdx ? 'active' : ''} ${s.done ? 'done' : ''}`;
      item.innerHTML = `<span>${s.name}</span> <span class="status-icon">${s.done ? '✅' : '●'}</span>`;
      item.onclick = () => { activeStudentIdx = idx; selectStudent(s); };
      list.appendChild(item);
    });
  }

  function selectStudent(student) {
    document.getElementById('activeStudentInfo').textContent = `${currentClass}반 ${student.name} 학생 평가 중`;
    document.getElementById('studentKeyword').value = student.keyword || "";
    document.getElementById('generatedText').value = student.text || "";
    document.getElementById('generatedText').classList.remove('error');
    document.querySelectorAll('.rubric-card').forEach(c => {
      c.classList.toggle('selected', (student.eval || []).includes(c.dataset.label));
    });
    renderStudentList();
  }

  function updateData() {
    if (activeStudentIdx === null) return;
    const sels = Array.from(document.querySelectorAll('.rubric-card.selected')).map(c => c.dataset.label);
    studentsData[currentClass][activeStudentIdx].eval = sels;
    studentsData[currentClass][activeStudentIdx].keyword = document.getElementById('studentKeyword').value;
  }

  // ── 핵심: AI 생성 로직 (정밀 디버깅 추가) ──
  document.getElementById('aiGenerateBtn').onclick = async () => {
    if (activeStudentIdx === null) return alert('학생을 먼저 선택하세요.');
    if (!geminiApiKey) return alert('관리자 모드에서 API 키를 먼저 입력하고 [저장]을 눌러주세요.');

    const sels = studentsData[currentClass][activeStudentIdx].eval || [];
    if (sels.length === 0) return alert('항목을 최소 1개 체크해주세요.');

    const keyword = document.getElementById('studentKeyword').value.trim();
    const prompt = `너는 10년차 정보교사야. 다음 항목으로 중학교 생활기록부 세특을 써줘.
    항목: ${sels.join(', ')} / 특징: ${keyword || '없음'}.
    조건: 2-3문장, 명사형 종결(~함, ~임), 전문적 어조, 불필요한 설명 금지.`;

    const area = document.getElementById('generatedText');
    area.value = "AI가 문장을 다듬고 있습니다... (잠시만 기다려주세요)";
    area.classList.add('loading');
    area.classList.remove('error');

    try {
      const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${geminiApiKey}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] })
      });

      const result = await res.json();
      
      if (!res.ok) {
        // 구글에서 보낸 실제 에러 상세 내용을 파악
        const errDetail = result.error ? `${result.error.status}: ${result.error.message}` : '알 수 없는 API 서버 오류';
        throw new Error(errDetail);
      }

      const aiText = result.candidates[0].content.parts[0].text.trim();
      area.value = aiText;
      area.classList.remove('loading');
      studentsData[currentClass][activeStudentIdx].text = aiText;
    } catch (e) {
      area.classList.remove('loading');
      area.classList.add('error');
      area.value = `⚠️ AI 통신 오류 발생!\n\n[원인 파악용 메시지]\n${e.message}\n\n*해결 팁: API 키가 정확한지, 혹은 하루 사용량을 초과했는지 확인해 보세요.`;
      console.error('Gemini API Error:', e);
    }
  };

  // ── 공통 기능 ──
  document.getElementById('saveBtn').onclick = () => { if(activeStudentIdx!==null){ updateData(); studentsData[currentClass][activeStudentIdx].done=true; saveToLocal(); renderStudentList(); alert('저장 완료!'); } };
  document.getElementById('copyBtn').onclick = () => { const t = document.getElementById('generatedText').value; if(t) { navigator.clipboard.writeText(t); alert('복사되었습니다.'); } };
  function saveToLocal() { localStorage.setItem('eval_v9_students', JSON.stringify(studentsData)); localStorage.setItem('eval_v9_rubrics', JSON.stringify(rubricData)); localStorage.setItem('eval_v9_api', geminiApiKey); }
  document.getElementById('classSelect').onchange = function() { currentClass = this.value; activeStudentIdx = null; selectStudent({name:"학생을 선택하세요", text:"", keyword:"", eval:[]}); renderStudentList(); };
  
  document.getElementById('exportBtn').onclick = () => {
    let csv = "\uFEFF학년반,이름,특징,세부능력 및 특기사항\n";
    for (let c in studentsData) studentsData[c].forEach(s => { csv += `${c}반,${s.name},${s.keyword||''},"${(s.text||'').replace(/"/g,'""')}"\n`; });
    const link = document.createElement("a"); link.href = URL.createObjectURL(new Blob([csv], {type:'text/csv'})); link.download = `세특_결과_${new Date().toLocaleDateString()}.csv`; link.click();
  };

  // ── 관리자 모드 로직 ──
  let clicks = 0;
  document.getElementById('logo').onclick = () => { clicks++; if(clicks>=3){ openAdmin(); clicks=0; } setTimeout(()=>clicks=0, 1000); };
  function openAdmin() {
    document.getElementById('adminPanel').style.display = 'flex';
    document.getElementById('targetClassLabel').textContent = currentClass;
    document.getElementById('adminNameInput').value = (studentsData[currentClass]||[]).map(s=>s.name).join('\n');
    document.getElementById('apiKeyInput').value = geminiApiKey;
    renderRubricEditor();
  }
  document.getElementById('adminCloseBtn').onclick = () => document.getElementById('adminPanel').style.display = 'none';
  document.getElementById('exitAdminBtn').onclick = () => document.body.classList.remove('admin-mode');

  document.getElementById('adminSaveAllBtn').onclick = () => {
    geminiApiKey = document.getElementById('apiKeyInput').value.trim();
    const names = document.getElementById('adminNameInput').value.split('\n').map(n=>n.trim()).filter(Boolean);
    studentsData[currentClass] = names.map(n => (studentsData[currentClass].find(ex=>ex.name===n) || {name:n, eval:[], text:"", keyword:"", done:false}));
    
    const newRubrics = [];
    document.querySelectorAll('.edit-category').forEach(cat => {
      const title = cat.querySelector('.cat-title').value.trim();
      const items = Array.from(cat.querySelectorAll('.item-label')).map(i=>({label:i.value.trim()})).filter(i=>i.label);
      if(title) newRubrics.push({title, items});
    });
    rubricData = newRubrics;
    saveToLocal(); renderMainRubrics(); renderStudentList();
    alert('모든 설정이 저장되었습니다.');
    document.getElementById('adminPanel').style.display = 'none';
  };

  function renderRubricEditor() {
    const cont = document.getElementById('rubricEditorContainer'); cont.innerHTML = '';
    rubricData.forEach(cat => {
      const div = document.createElement('div'); div.className = 'edit-category';
      div.style.border="1px solid #ddd"; div.style.padding="10px"; div.style.marginBottom="10px";
      div.innerHTML = `<input type="text" class="cat-title" value="${cat.title}" style="font-weight:bold; width:70%; padding:5px;"> <button class="btn btn-danger btn-sm" onclick="this.parentElement.remove()">삭제</button><div class="item-list" style="margin-top:10px;">` + 
        cat.items.map(it => `<div style="display:flex; gap:5px; margin-bottom:5px;"><input type="text" class="item-label" value="${it.label}" style="flex:1; padding:5px;"> <button onclick="this.parentElement.remove()">X</button></div>`).join('') + 
        `</div><button class="btn btn-outline btn-sm" onclick="addItem(this)">+ 항목 추가</button>`;
      cont.appendChild(div);
    });
  }
  window.addItem = (btn) => { const div = document.createElement('div'); div.style="display:flex; gap:5px; margin-bottom:5px;"; div.innerHTML=`<input type="text" class="item-label" placeholder="평가 키워드" style="flex:1; padding:5px;"> <button onclick="this.parentElement.remove()">X</button>`; btn.previousElementSibling.appendChild(div); };
  window.addCategory = () => { const cont = document.getElementById('rubricEditorContainer'); const div = document.createElement('div'); div.className='edit-category'; div.style="border:1px solid #ddd; padding:10px; margin-bottom:10px;"; div.innerHTML=`<input type="text" class="cat-title" placeholder="주제명" style="font-weight:bold; width:70%; padding:5px;"> <button onclick="this.parentElement.remove()">삭제</button><div class="item-list" style="margin-top:10px;"></div><button onclick="addItem(this)">+ 항목 추가</button>`; cont.appendChild(div); };
</script>
</body>
</html>
