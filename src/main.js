// 🌟 將 CSS 全部用 import 納入模組管理，交給打包工具（Vite/Webpack）動態編譯注入
import 'bootstrap/dist/css/bootstrap.min.css';
import 'bootstrap/dist/js/bootstrap.bundle.min.js';
import './style.scss'; 

import { parseExcelWorkbook } from './parser.js';
import { exportWrongQuestionsExcel } from './exporter.js';

// 使用 DOMContentLoaded 確保 HTML 元素完全載入後才執行 JavaScript
document.addEventListener('DOMContentLoaded', () => {
  
  // 全域狀態管理
  let rawWorkbookData = { before: [], net: [], wrongFile: [] };
  let currentPool = [];       
  let selectedQuestions = []; 
  let wrongQuestions = [];    

  // DOM 元素獲取
  const fileInput = document.getElementById('excelFile');
  const rangeInputs = document.getElementsByName('sheetRange');
  const rangeSelectorGroup = document.getElementById('rangeSelectorGroup'); 
  const wrongModeAlert = document.getElementById('wrongModeAlert');         
  const quizCountInput = document.getElementById('quizCount');
  const totalAvailableText = document.getElementById('totalAvailableText');
  const startBtn = document.getElementById('startBtn');
  const setupSection = document.getElementById('setupSection');
  const quizContainer = document.getElementById('quizContainer');
  const questionsWrapper = document.getElementById('questionsWrapper');
  const submitBtn = document.getElementById('submitBtn');
  const resultContainer = document.getElementById('resultContainer');
  const scoreText = document.getElementById('scoreText');
  const resultDetails = document.getElementById('resultDetails');
  const restartBtn = document.getElementById('restartBtn');
  const exportWrongBtn = document.getElementById('exportWrongBtn');

  // 【核心除錯提示】檢查關鍵 DOM 是否成功獲取
  if (!fileInput) {
    console.error("❌ 找不到 id='excelFile' 的 HTML 元素，請檢查 index.html 是否正確設定，或 script 標籤位置是否正確。");
    return;
  }

  // 智慧上傳事件監聽
  fileInput.addEventListener('change', async function(e) {
    const file = e.target.files[0];
    console.log("📂 偵測到檔案變更：", file ? file.name : "無檔案");
    if (!file) return;

    try {
      console.log("⏳ 開始呼叫 parseExcelWorkbook 解析 Excel 檔案...");
      const parsedData = await parseExcelWorkbook(file);
      console.log("✅ Excel 解析成功，回傳原始資料：", parsedData);

      // 安全賦值，若欄位不存在則給予空陣列避免後續報錯
      rawWorkbookData.before = parsedData.before || [];
      rawWorkbookData.net = parsedData.net || [];
      rawWorkbookData.wrongFile = parsedData.wrongFile || [];

      if (parsedData.isWrongFileMode && rawWorkbookData.wrongFile.length > 0) {
        // 模式 A：識別為錯題本
        currentPool = [...rawWorkbookData.wrongFile];
        if (rangeSelectorGroup) rangeSelectorGroup.classList.add('d-none');
        if (wrongModeAlert) wrongModeAlert.classList.remove('d-none');
        
        totalAvailableText.innerText = `（🎯 錯題本模式：共 ${currentPool.length} 題）`;
        totalAvailableText.className = "badge bg-danger fs-6 p-2";
        
        quizCountInput.disabled = false;
        quizCountInput.max = currentPool.length;
        quizCountInput.value = currentPool.length; 
        startBtn.disabled = false;
        
        alert(`智慧識別成功：這是一份錯題本！\n共匯入 ${currentPool.length} 題，進入錯題盲練模式。`);
      } else if (rawWorkbookData.before.length > 0 || rawWorkbookData.net.length > 0) {
        // 模式 B：識別為大題庫
        if (rangeSelectorGroup) rangeSelectorGroup.classList.remove('d-none');
        if (wrongModeAlert) wrongModeAlert.classList.add('d-none');
        
        rangeInputs.forEach(input => input.disabled = false);
        const defaultRadio = document.getElementById('rangeBefore');
        if (defaultRadio) defaultRadio.checked = true; 
        
        updatePoolAndUI();
        alert(`智慧識別成功：這是原始大題庫！\n課前篇：${rawWorkbookData.before.length} 題\n網路資源：${rawWorkbookData.net.length} 題`);
      } else {
        alert('無法識別工作表，請確認 Sheet 名稱是否包含「課前」、「網路資源」或「錯題」。');
      }
    } catch (err) {
      console.error("❌ Excel 解析程序崩潰，詳細錯誤訊息：", err);
      alert('Excel 解析失敗，請開啟瀏覽器主控台 (F12) 確認詳細錯誤原因，或確認檔案結構。');
    }
  });

  // 範圍切換事件
  rangeInputs.forEach(input => {
    input.addEventListener('change', updatePoolAndUI);
  });

  // 更新題庫範圍與 UI 狀態
  function updatePoolAndUI() {
    if (rawWorkbookData.wrongFile.length > 0) return; 

    const checkedRadio = document.querySelector('input[name="sheetRange"]:checked');
    if (!checkedRadio) return;
    
    const selectedRange = checkedRadio.value;
    console.log(`🎯 切換題庫範圍至：${selectedRange}`);
    
    if (selectedRange === 'before') {
      currentPool = [...rawWorkbookData.before];
    } else if (selectedRange === 'net') {
      currentPool = [...rawWorkbookData.net];
    } else if (selectedRange === 'both') {
      currentPool = [...rawWorkbookData.before, ...rawWorkbookData.net];
    }
    
    if (currentPool.length > 0) {
      quizCountInput.disabled = false;
      quizCountInput.max = currentPool.length;
      quizCountInput.value = Math.min(10, currentPool.length);
      totalAvailableText.innerText = `（當前範圍共 ${currentPool.length} 題）`;
      totalAvailableText.className = "badge bg-success fs-6 p-2";
      startBtn.disabled = false;
    } else {
      quizCountInput.disabled = true;
      startBtn.disabled = true;
      totalAvailableText.innerText = `（該範圍內無題目）`;
      totalAvailableText.className = "badge bg-danger fs-6 p-2";
    }
  }

  // 開始測驗
  startBtn.addEventListener('click', function() {
    if (currentPool.length === 0) {
      alert("當前題庫沒有題目可以抽題！");
      return;
    }
    
    let count = parseInt(quizCountInput.value) || 10;
    if (count > currentPool.length) count = currentPool.length;
    if (count <= 0) count = 1;

    console.log(`🎲 開始隨取隨練，預計抽洗：${count} 題 / 總庫 ${currentPool.length} 題`);
    const shuffled = [...currentPool].sort(() => 0.5 - Math.random());
    selectedQuestions = shuffled.slice(0, count);

    renderQuiz(selectedQuestions);

    setupSection.classList.add('d-none');
    quizContainer.classList.remove('d-none');
    resultContainer.classList.add('d-none');
  });

  // 核心渲染：結合 Bootstrap 原生 btn-check 與自訂 .quiz-option-btn 大方框
  function renderQuiz(questions) {
    questionsWrapper.innerHTML = '';
    
    questions.forEach((q, qIndex) => {
      const qBlock = document.createElement('div');
      qBlock.className = 'card shadow-sm mb-4';
      
      let optionsHTML = '';
      if (q.options_en && Array.isArray(q.options_en)) {
        q.options_en.forEach((optionStr, oIndex) => {
          const match = optionStr.match(/^([A-Ea-e])/);
          const optValue = match ? match[1].toUpperCase() : '';
          const uniqueId = `q_${qIndex}_opt_${oIndex}`;
          
          optionsHTML += `
            <div class="mb-2">
              <input type="radio" 
                     name="quiz_q_${qIndex}" 
                     id="${uniqueId}" 
                     value="${optValue}" 
                     class="btn-check" 
                     required>
              <label class="btn quiz-option-btn shadow-sm text-start w-100" for="${uniqueId}">
                ${optionStr}
              </label>
            </div>
          `;
        });
      }

      qBlock.innerHTML = `
        <div class="card-body p-4">
          <h5 class="card-title fw-bold text-dark mb-4 pb-2 border-bottom">
            <span class="text-primary me-2">Q${qIndex + 1}.</span>${q.q_en || '(無英文題目)'}
          </h5>
          <div class="options-group">${optionsHTML}</div>
        </div>
      `;
      questionsWrapper.appendChild(qBlock);
    });
  }

  // 提交答案評分
  submitBtn.addEventListener('click', function() {
    const form = document.getElementById('quizForm');
    
    // 檢查 Bootstrap 表單驗證狀態
    if (!form.checkValidity()) {
      alert('請作答完所有題目後再送出評分！');
      form.classList.add('was-validated');
      
      // 智慧滾動到第一個未答題目
      const firstInvalid = form.querySelector(':invalid');
      if (firstInvalid) {
        firstInvalid.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }
      return;
    }

    let correctCount = 0;
    let htmlDetails = '';
    wrongQuestions = []; 

    selectedQuestions.forEach((q, qIndex) => {
      const selectedRadio = form.querySelector(`input[name="quiz_q_${qIndex}"]:checked`);
      const userAnswer = selectedRadio ? selectedRadio.value : '';
      const isCorrect = (userAnswer === q.answer);

      if (isCorrect) {
        correctCount++;
      } else {
        wrongQuestions.push(q); 
      }

      // 處理換行符號
      const explanationHTML = q.explanation 
        ? `<div class="alert alert-warning mt-3 border-start border-warning border-3 mb-0" role="alert">
            <strong>💡 題目解說：</strong><br>${q.explanation.replace(/\n/g, '<br>')}
           </div>` 
        : '';

      // 安全處理中文選項參考（防止 options_zh 為 undefined）
      const hasChineseOptions = q.options_zh && Array.isArray(q.options_zh) && q.options_zh.length > 0;

      htmlDetails += `
        <div class="card shadow-sm mb-4 border-start border-4 ${isCorrect ? 'border-success' : 'border-danger'}">
          <div class="card-body p-4">
            <h5 class="fw-bold text-dark mb-2">Q${qIndex + 1}. ${q.q_en || ''}</h5>
            
            <div class="alert alert-secondary py-2 px-3 my-3 fs-6 text-muted rounded-3">
              <p class="mb-1 fw-bold text-dark">【中文題目翻譯】</p>
              <p class="mb-2">${q.q_zh || '（無中文翻譯）'}</p>
              ${hasChineseOptions ? `<p class="mb-1 fw-bold text-dark">【中文選項參考】</p><p class="mb-0 small">${q.options_zh.join('<br>')}</p>` : ''}
            </div>

            <p class="fs-5 mb-1 ${isCorrect ? 'text-success fw-bold' : 'text-danger fw-bold'}">
              您的回答：選項 (${userAnswer || '未答'}) ${isCorrect ? ' ➔ 答對了！' : ' ➔ 答錯了'}
            </p>
            ${!isCorrect ? `<p class="text-success fw-bold fs-5 mb-0">正確答案是：選項 (${q.answer})</p>` : ''}
            ${explanationHTML}
          </div>
        </div>
      `;
    });

    const totalQuestions = selectedQuestions.length;
    const finalScore = Math.round((correctCount / totalQuestions) * 100);

    scoreText.innerText = finalScore;
    resultDetails.innerHTML = htmlDetails;

    if (wrongQuestions.length > 0) {
      exportWrongBtn.classList.remove('d-none');
      exportWrongBtn.innerText = `📥 匯出本次錯題本 (${wrongQuestions.length} 題)`;
    } else {
      exportWrongBtn.classList.add('d-none');
    }

    quizContainer.classList.add('d-none');
    resultContainer.classList.remove('d-none');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

  // 匯出錯題
  exportWrongBtn.addEventListener('click', function() {
    if (wrongQuestions.length === 0) return;
    exportWrongQuestionsExcel(wrongQuestions);
  });

  // 再測驗一次
  restartBtn.addEventListener('click', function() {
    document.getElementById('quizForm').classList.remove('was-validated');
    setupSection.classList.remove('d-none');
    quizContainer.classList.add('d-none');
    resultContainer.classList.add('d-none');
    
    if (rawWorkbookData.wrongFile.length > 0) {
      totalAvailableText.innerText = `（🎯 錯題本模式：共 ${currentPool.length} 題）`;
    } else {
      updatePoolAndUI();
    }
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

});