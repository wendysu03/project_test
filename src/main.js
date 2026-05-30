// 🌟 完美對接修正版：修正 await 語法衝突，100% 相容原本 parser.js 的 Promise 回傳結構

document.addEventListener('DOMContentLoaded', () => {
  
  // ==========================================
  // 1. 全域狀態管理
  // ==========================================
  let rawWorkbookData = { before: [], net: [], wrongFile: [] };
  let currentPool = [];       
  let selectedQuestions = []; 
  let wrongQuestions = [];        
  let cumulativeWrongQuestions = []; // 支援跨輪次累積

  // ==========================================
  // 2. DOM 元素獲取
  // ==========================================
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

  // 智慧狀態提示與防呆元件
  const statusMessageAlert = document.getElementById('statusMessageAlert');
  const statusMessageText = document.getElementById('statusMessageText');
  const formValidationAlert = document.getElementById('formValidationAlert');

  // UI 動態提示輔助功能
  function showPageNotification(message, type = 'success') {
    if (!statusMessageAlert || !statusMessageText) {
      console.log(`[系統提示][${type}]: ${message.replace(/<[^>]*>/g, '')}`);
      return;
    }
    statusMessageAlert.classList.remove('alert-success', 'alert-danger', 'alert-warning', 'alert-info', 'd-none');
    statusMessageAlert.classList.add(`alert-${type}`);
    statusMessageText.innerHTML = message;
    statusMessageAlert.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  if (!fileInput) {
    console.error("❌ [錯誤] 找不到 id='excelFile' 的上傳元件，請檢查 index.html。");
    return;
  }

  // ==========================================
  // 3. 智慧上傳事件監聽 (🛠️ 已在此徹底修復非同步鏈路)
  // ==========================================
  fileInput.addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    if (formValidationAlert) formValidationAlert.classList.add('d-none');

    // 安全檢查：確認全域大腦（window）是否看得到你寫的 parser 函式
    const parseFn = window.parseExcelWorkbook || (typeof parseExcelWorkbook === 'function' ? parseExcelWorkbook : null);
    
    if (!parseFn) {
      showPageNotification('❌ <strong>核心連線失敗：</strong> 瀏覽器找不到 <code>parseExcelWorkbook</code> 函式。請確認你的 <code>parser.js</code> 檔案最頂端的 <code>import</code> 和開頭的 <code>export</code> 真的都刪乾淨了。', 'danger');
      return;
    }

    console.log("⏳ [鏈路啟動] 正在透過 parser.js 解析 Excel...");

    // 🌟 核心修正：改用標準 Promise .then() 語法，完美對齊你原本 parser.js 的 return new Promise 結構
    parseFn(file)
      .then((parsedData) => {
        console.log("✅ [解析完成] 回傳原始資料內容為：", parsedData);

        // 安全初始化防止 undefined
        rawWorkbookData.before = parsedData.before || [];
        rawWorkbookData.net = parsedData.net || [];
        rawWorkbookData.wrongFile = parsedData.wrongFile || [];

        cumulativeWrongQuestions = []; // 重新上傳檔案時清空舊累積

        if (parsedData.isWrongFileMode && rawWorkbookData.wrongFile.length > 0) {
          // 模式 A：識別為錯題本
          currentPool = [...rawWorkbookData.wrongFile];
          if (rangeSelectorGroup) rangeSelectorGroup.classList.add('d-none');
          if (wrongModeAlert) wrongModeAlert.classList.remove('d-none');
          
          totalAvailableText.innerText = `（共 ${currentPool.length} 題）`;
          totalAvailableText.className = "badge bg-danger fs-6 p-2";
          
          quizCountInput.disabled = false;
          quizCountInput.max = currentPool.length;
          quizCountInput.value = currentPool.length; 
          startBtn.disabled = false;
          
          showPageNotification(`🎉 <strong>智慧識別成功：</strong> 偵測到「錯題複習本」！已成功匯入 <strong>${currentPool.length}</strong> 題並開啟盲練模式。`, 'danger');
          
        } else if (rawWorkbookData.before.length > 0 || rawWorkbookData.net.length > 0) {
          // 模式 B：識別為大題庫
          if (rangeSelectorGroup) rangeSelectorGroup.classList.remove('d-none');
          if (wrongModeAlert) wrongModeAlert.classList.add('d-none');
          
          rangeInputs.forEach(input => input.disabled = false);
          const defaultRadio = document.getElementById('rangeBefore');
          if (defaultRadio) defaultRadio.checked = true; 
          
          updatePoolAndUI();
          showPageNotification(`🎉 <strong>智慧識別成功：</strong> 成功載入原始大題庫！<br>• 課前測驗：<strong>${rawWorkbookData.before.length}</strong> 題<br>• 網路資源：<strong>${rawWorkbookData.net.length}</strong> 題`, 'success');
          
        } else {
          showPageNotification('⚠️ <strong>無法識別工作表：</strong> 請確認您的 Excel Sheet 名稱是否包含「課前」、「網路資源」或「錯題」。', 'warning');
        }
      })
      .catch((err) => {
        console.error("❌ Excel 內部解析程序潰敗:", err);
        showPageNotification(`❌ <strong>Excel 讀取失敗：</strong> 可能是 Excel 內部格式有問題。詳細原因: ${err.message || err}`, 'danger');
      });
  });

  // ==========================================
  // 4. 範圍選擇連動
  // ==========================================
  rangeInputs.forEach(input => {
    input.addEventListener('change', updatePoolAndUI);
  });

  function updatePoolAndUI() {
    if (rawWorkbookData.wrongFile.length > 0) return; 

    const checkedRadio = document.querySelector('input[name="sheetRange"]:checked');
    if (!checkedRadio) return;
    
    const selectedRange = checkedRadio.value;
    
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

  // ==========================================
  // 5. 開始測驗事件
  // ==========================================
  startBtn.addEventListener('click', function() {
    if (currentPool.length === 0) return;
    
    let count = parseInt(quizCountInput.value, 10) || 10;
    if (count > currentPool.length) count = currentPool.length;
    if (count <= 0) count = 1;

    const shuffled = [...currentPool].sort(() => 0.5 - Math.random());
    selectedQuestions = shuffled.slice(0, count);

    renderQuiz(selectedQuestions);

    setupSection.classList.add('d-none');
    quizContainer.classList.remove('d-none');
    resultContainer.classList.add('d-none');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

  // ==========================================
  // 6. 題目作答區渲染引擎
  // ==========================================
  function renderQuiz(questions) {
    questionsWrapper.innerHTML = '';
    
    questions.forEach((q, qIndex) => {
      const qBlock = document.createElement('div');
      qBlock.className = 'card shadow-sm mb-4 border-0';
      
      let optionsHTML = '';
      if (q.options_en && Array.isArray(q.options_en)) {
        q.options_en.forEach((optionStr, oIndex) => {
          const match = optionStr.match(/^([A-Ea-e])/);
          const optValue = match ? match[1].toUpperCase() : '';
          const uniqueId = `q_${qIndex}_opt_${oIndex}`;
          
          optionsHTML += `
            <div class="mb-2">
              <input type="radio" name="quiz_q_${qIndex}" id="${uniqueId}" value="${optValue}" class="btn-check" required>
              <label class="btn btn-outline-secondary text-start font-weight-bold w-100 py-3 px-4 shadow-sm" for="${uniqueId}" style="border-width: 2px; color: #020202;">
                ${optionStr}
              </label>
            </div>
          `;
        });
      }

      qBlock.innerHTML = `
        <div class="card-body p-4 bg-white rounded shadow-sm">
          <h5 class="card-title fw-bold text-dark mb-4 pb-2 border-bottom">
            <span class="text-primary me-2">Q${qIndex + 1}.</span>${q.q_en || '（題目內容缺失）'}
          </h5>
          <div class="options-group">${optionsHTML}</div>
        </div>
      `;
      questionsWrapper.appendChild(qBlock);
    });
  }

  // ==========================================
  // 7. 提交答案與結果評分引擎
  // ==========================================
  submitBtn.addEventListener('click', function() {
    const form = document.getElementById('quizForm');
    
    if (!form.checkValidity()) {
      form.classList.add('was-validated');
      if (formValidationAlert) {
        formValidationAlert.classList.remove('d-none');
        formValidationAlert.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }
      return;
    }

    if (formValidationAlert) formValidationAlert.classList.add('d-none');

    let correctCount = 0;
    let htmlDetails = '';
    wrongQuestions = []; 

    selectedQuestions.forEach((q, qIndex) => {
      const selectedRadio = form.querySelector(`input[name="quiz_q_${qIndex}"]:checked`);
      const userAnswer = selectedRadio ? selectedRadio.value : '';
      const isCorrect = (String(userAnswer).trim().toUpperCase() === String(q.answer).trim().toUpperCase());

      if (isCorrect) {
        correctCount++;
      } else {
        wrongQuestions.push(q); 
        
        // 跨輪次去重累積
        const isAlreadyAdded = cumulativeWrongQuestions.some(cq => cq.q_en === q.q_en);
        if (!isAlreadyAdded) {
          cumulativeWrongQuestions.push(q);
        }
      }

      let optionsReviewHTML = '';
      if (q.options_en && Array.isArray(q.options_en)) {
        optionsReviewHTML = '<div class="my-3 ps-2 border-start border-2 text-secondary small">';
        q.options_en.forEach(opt => {
          const match = opt.match(/^([A-Ea-e])/);
          const currentOptLetter = match ? match[1].toUpperCase() : '';
          
          let optBadge = '';
          if (currentOptLetter === q.answer) {
            optBadge = '<span class="badge bg-success me-1">正確答案</span>';
          } else if (currentOptLetter === userAnswer && !isCorrect) {
            optBadge = '<span class="badge bg-danger me-1">您的選擇</span>';
          }
          
          optionsReviewHTML += `<div class="py-1">${optBadge} ${opt}</div>`;
        });
        optionsReviewHTML += '</div>';
      }

      let explanationHTML = '';
      if (q.explanation && String(q.explanation).trim() !== "" && String(q.explanation).trim() !== "undefined") {
        explanationHTML = `
          <div class="alert alert-warning mt-3 border-start border-warning border-4 mb-0 shadow-sm" role="alert">
            <div class="fw-bold text-dark mb-1">💡 題目完整解說：</div>
            <div class="text-secondary small" style="line-height: 1.6;">
              ${String(q.explanation).replace(/\n/g, '<br>')}
            </div>
          </div>`;
      }

      const hasChineseOptions = q.options_zh && Array.isArray(q.options_zh) && q.options_zh.length > 0;

      htmlDetails += `
        <div class="card shadow-sm mb-4 border-start border-4 ${isCorrect ? 'border-success' : 'border-danger'}">
          <div class="card-body p-4 bg-white rounded">
            <h5 class="fw-bold text-dark mb-2">Q${qIndex + 1}. ${q.q_en || ''}</h5>
            ${optionsReviewHTML}
            
            <div class="alert alert-secondary py-2 px-3 my-3 fs-6 text-muted rounded-3 border-0">
              <p class="mb-1 fw-bold text-dark">【中文題目翻譯】</p>
              <p class="mb-2">${q.q_zh || '（無中文翻譯）'}</p>
              ${hasChineseOptions ? `<p class="mb-1 fw-bold text-dark">【中文選項參考】</p><p class="mb-0 small">${q.options_zh.join('<br>')}</p>` : ''}
            </div>

            <div class="d-flex align-items-center gap-2 my-2">
              <span class="fs-5 ${isCorrect ? 'text-success fw-bold' : 'text-danger fw-bold'}">
                您的回答：選項 (${userAnswer || '未答'}) ${isCorrect ? ' ➔ 答對了！' : ' ➔ 答錯了'}
              </span>
            </div>
            ${!isCorrect ? `<p class="text-success fw-bold fs-5 mb-2">正確答案是：選項 (${q.answer})</p>` : ''}
            ${explanationHTML}
          </div>
        </div>
      `;
    });

    const totalQuestions = selectedQuestions.length;
    const finalScore = totalQuestions > 0 ? Math.round((correctCount / totalQuestions) * 100) : 0;

    scoreText.innerText = finalScore;
    resultDetails.innerHTML = htmlDetails;

    // 按鈕動態更新：跨輪次累積題數
    if (cumulativeWrongQuestions.length > 0) {
      exportWrongBtn.classList.remove('d-none');
      exportWrongBtn.innerText = `📥 匯出累積錯題本 (共 ${cumulativeWrongQuestions.length} 題)`;
    } else {
      exportWrongBtn.classList.add('d-none');
    }

    quizContainer.classList.add('d-none');
    resultContainer.classList.remove('d-none');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

  // ==========================================
  // 8. 匯出累積錯題本 (內建安全自接管防護)
  // ==========================================
  exportWrongBtn.addEventListener('click', function() {
    if (cumulativeWrongQuestions.length === 0) return;
    
    const targetFunction = window.exportWrongQuestionsExcel || (typeof exportWrongQuestionsExcel === 'function' ? exportWrongQuestionsExcel : null);
    
    if (!targetFunction) {
      console.log("⚠️ 找不到外部匯出函式，啟動 main.js 本地自我包裝導出...");
      try {
        const excelRows = [["題號", "英文題目", "英文選項", "正確答案", "中文題目", "中文選項", "解說"]];
        cumulativeWrongQuestions.forEach((q, idx) => {
          excelRows.push([idx + 1, q.q_en, q.options_en.join('\n'), q.answer, q.q_zh, q.options_zh.join('\n'), q.explanation]);
        });
        const worksheet = XLSX.utils.aoa_to_sheet(excelRows);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "錯題複習");
        const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, "");
        XLSX.writeFile(workbook, `專案管理_錯題本_${dateStr}.xlsx`);
        
        cumulativeWrongQuestions = [];
        exportWrongBtn.classList.add('d-none');
        return;
      } catch(sheetErr) {
        console.error("本地導出失敗:", sheetErr);
        return;
      }
    }
    
    targetFunction(cumulativeWrongQuestions);
    cumulativeWrongQuestions = [];
    exportWrongBtn.classList.add('d-none');
  });

  // ==========================================
  // 9. 再測驗一次狀態重設
  // ==========================================
  restartBtn.addEventListener('click', function() {
    document.getElementById('quizForm').classList.remove('was-validated');
    setupSection.classList.remove('d-none');
    quizContainer.classList.add('d-none');
    resultContainer.classList.add('d-none');
    if (statusMessageAlert) statusMessageAlert.classList.add('d-none');
    
    if (rawWorkbookData.wrongFile.length > 0) {
      totalAvailableText.innerText = `（共 ${currentPool.length} 題）`;
    } else {
      updatePoolAndUI();
    }
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

});