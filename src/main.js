// 🌟 完美對接修正版：修正 await 語法衝突，100% 相容原本 parser.js 的 Promise 回傳結構
// 修正重點：
// 1. 智慧跨輪次無縫補題：輸入題數若大於未刷剩餘數，會自動撈完剩餘題目，並「立刻自動重設題池」補滿使用者要求的題數，練習不中斷！
// 2. 嚴格範圍上限防呆：輸入題數絕對無法超過當前選定範圍（課前/網路/全部）的「總題數」，超過自動修正。
// 3. 智慧地雷防護：選項含 "Both", "All of", "以上皆" 等題「不洗牌」保持原始順序。
// 4. 答題後中文題目與選項皆整合在下方灰色區塊內。
// 5. 分數旁動態顯示（答對數/總題數）。

document.addEventListener('DOMContentLoaded', () => {
  
  // ==========================================
  // 1. 全域狀態管理
  // ==========================================
  let rawWorkbookData = { before: [], net: [], wrongFile: [] };
  let currentPool = [];       // 當前選定範圍的「總題池」（如：課前、網路資源或全部）
  let unplayedPool = [];      // 目前選定範圍中「還沒被抽過」的題目池
  let selectedQuestions = []; 
  let wrongQuestions = [];        
  let cumulativeWrongQuestions = []; // 支援跨輪次累積錯題

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

  const statusMessageAlert = document.getElementById('statusMessageAlert');
  const statusMessageText = document.getElementById('statusMessageText');
  const formValidationAlert = document.getElementById('formValidationAlert');

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

  // 驗證檔案
  function validateUploadFile(file) {
    return new Promise((resolve) => {
      if (!file || !file.name) return resolve({ ok: false, reason: '找不到上傳檔案。' });
      const name = file.name;
      const ext = name.split('.').pop().toLowerCase();
      const maxSize = 10 * 1024 * 1024;
      const forbidden = ['xlsm', 'xlsb', 'xls'];
      if (forbidden.includes(ext)) return resolve({ ok: false, reason: '為了安全性，請另存為 .xlsx 或 .csv。' });
      const allowed = ['xlsx', 'csv'];
      if (!allowed.includes(ext)) return resolve({ ok: false, reason: '僅接受 .xlsx 或 .csv 檔案。' });
      if (file.size > maxSize) return resolve({ ok: false, reason: `檔案過大（上限 10MB）。` });

      if (ext === 'xlsx') {
        try {
          const blob = file.slice(0, 4);
          const reader = new FileReader();
          reader.onload = (e) => {
            const arr = new Uint8Array(e.target.result);
            const sig = String.fromCharCode(arr[0], arr[1], arr[2], arr[3]);
            if (sig === 'PK\x03\x04' || sig === 'PK\u0003\u0004') resolve({ ok: true });
            else resolve({ ok: false, reason: 'ZIP 標頭檢查失敗。' });
          };
          reader.onerror = () => resolve({ ok: false, reason: '無法讀取檔案。' });
          reader.readAsArrayBuffer(blob);
        } catch (e) { resolve({ ok: false, reason: '驗證發生錯誤。' }); }
      } else { resolve({ ok: true }); }
    });
  }

  if (!fileInput) return;

  // ==========================================
  // 3. 智慧上傳事件監聽
  // ==========================================
  fileInput.addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    if (formValidationAlert) formValidationAlert.classList.add('d-none');
    const parseFn = window.parseExcelWorkbook || (typeof parseExcelWorkbook === 'function' ? parseExcelWorkbook : null);
    
    if (!parseFn) {
      showPageNotification('❌ 核心連線失敗，找不到 parseExcelWorkbook 函式。', 'danger');
      return;
    }

    validateUploadFile(file).then(validation => {
      if (!validation.ok) {
        showPageNotification(`<strong>上傳驗證失敗：</strong> ${validation.reason}`, 'danger');
        return;
      }

      parseFn(file).then((parsedData) => {
        rawWorkbookData.before = parsedData.before || [];
        rawWorkbookData.net = parsedData.net || [];
        rawWorkbookData.wrongFile = parsedData.wrongFile || [];
        cumulativeWrongQuestions = []; 

        if (parsedData.isWrongFileMode && rawWorkbookData.wrongFile.length > 0) {
          currentPool = [...rawWorkbookData.wrongFile];
          unplayedPool = [...currentPool]; 
          if (rangeSelectorGroup) rangeSelectorGroup.classList.add('d-none');
          if (wrongModeAlert) wrongModeAlert.classList.remove('d-none');
          
          updatePoolAndUI();
          showPageNotification(`🎉 成功匯入錯題複習本共 <strong>${currentPool.length}</strong> 題。`, 'danger');
        } else if (rawWorkbookData.before.length > 0 || rawWorkbookData.net.length > 0) {
          if (rangeSelectorGroup) rangeSelectorGroup.classList.remove('d-none');
          if (wrongModeAlert) wrongModeAlert.classList.add('d-none');
          
          rangeInputs.forEach(input => input.disabled = false);
          const defaultRadio = document.getElementById('rangeBefore');
          if (defaultRadio) defaultRadio.checked = true; 
          
          updatePoolAndUI();
          showPageNotification(`🎉 成功載入原始大題庫！課前測驗：${rawWorkbookData.before.length} 題 / 網路資源：${rawWorkbookData.net.length} 題`, 'success');
        }
      });
    });
  });

  // ==========================================
  // 4. 範圍選擇連動
  // ==========================================
  rangeInputs.forEach(input => {
    input.addEventListener('change', updatePoolAndUI);
  });

  function updatePoolAndUI() {
    if (rawWorkbookData.wrongFile.length > 0) {
      currentPool = [...rawWorkbookData.wrongFile];
    } else {
      const checkedRadio = document.querySelector('input[name="sheetRange"]:checked');
      if (!checkedRadio) return;
      
      if (checkedRadio.value === 'before') currentPool = [...rawWorkbookData.before];
      else if (checkedRadio.value === 'net') currentPool = [...rawWorkbookData.net];
      else if (checkedRadio.value === 'both') currentPool = [...rawWorkbookData.before, ...rawWorkbookData.net];
    }
    
    unplayedPool = [...currentPool]; // 更換範圍或初次載入時，重設未刷題池
    updateQuizCountUI();
  }

  function updateQuizCountUI() {
    if (currentPool.length > 0) {
      quizCountInput.disabled = false;
      // 🌟 核心控制：輸入上限「絕對不能超過當前選定範圍的總題數」
      quizCountInput.max = currentPool.length; 
      quizCountInput.value = Math.min(10, currentPool.length);
      
      totalAvailableText.innerText = `（當前範圍未刷：${unplayedPool.length} / 總共：${currentPool.length} 題）`;
      totalAvailableText.className = "badge bg-success fs-6 p-2";
      startBtn.disabled = false;
    } else {
      quizCountInput.disabled = true;
      startBtn.disabled = true;
      totalAvailableText.innerText = `（該範圍內無題目）`;
      totalAvailableText.className = "badge bg-danger fs-6 p-2";
    }
  }

  // Fisher-Yates 強隨機洗牌
  function shuffleArray(array) {
    const a = Array.isArray(array) ? [...array] : [];
    for (let i = a.length - 1; i > 0; i--) {
      let rand;
      try {
        if (window && window.crypto && window.crypto.getRandomValues) {
          rand = window.crypto.getRandomValues(new Uint32Array(1))[0] / (0xFFFFFFFF + 1);
        } else { rand = Math.random(); }
      } catch (e) { rand = Math.random(); }
      const j = Math.floor(rand * (i + 1));
      [a[i], a[j]] = [a[j], a[i]];
    }
    return a;
  }

  // ==========================================
  // 5. 開始測驗事件 (全新改寫：智慧跨輪次補題與嚴格上限控制)
  // ==========================================
  startBtn.addEventListener('click', function() {
    if (currentPool.length === 0) return;
    
    let requestedCount = parseInt(quizCountInput.value, 10) || 10;
    
    // 🌟 防呆 A：要求的題數絕對不能大於當前範圍的「總題數」
    if (requestedCount > currentPool.length) {
      console.warn(`[上限防呆] 請求 ${requestedCount} 題超過當前範圍總數，已修正為最大值：${currentPool.length}`);
      requestedCount = currentPool.length;
      quizCountInput.value = currentPool.length;
    }
    if (requestedCount <= 0) {
      requestedCount = 1;
      quizCountInput.value = 1;
    }

    selectedQuestions = [];

    // 🌟 核心選題演算法：滿足使用者要求的題數
    if (requestedCount <= unplayedPool.length) {
      // 情況一：未刷題池還夠抽
      const shuffledUnplayed = shuffleArray(unplayedPool);
      selectedQuestions = shuffledUnplayed.slice(0, requestedCount);
      
      // 更新未刷題池
      unplayedPool = unplayedPool.filter(poolQ => 
        !selectedQuestions.some(selQ => selQ.q_en === poolQ.q_en)
      );
    } else {
      // 情況二：未刷題池不夠抽了！（例如只剩 5 題，使用者卻要 20 題）
      alert(`💡 溫馨提示：當前範圍即將刷完！系統已自動幫您打包最後的 ${unplayedPool.length} 題，並無縫重設題池，為您補滿所需的 ${requestedCount} 題！`);
      
      // 1. 先把剩下的未刷題全拿走
      const lastRemainingQuestions = [...unplayedPool];
      selectedQuestions = [...lastRemainingQuestions];
      
      // 2. 理論上未刷題池歸零，立刻自動開啟新一輪大循環（重撈）
      unplayedPool = [...currentPool];
      
      // 3. 計算還差多少題（例如差 15 題）
      const neededCount = requestedCount - lastRemainingQuestions.length;
      
      // 4. 從新題池中先過濾掉剛剛拿走的那幾題，避免在同一卷中看到重覆題目
      let freshPool = unplayedPool.filter(poolQ => 
        !lastRemainingQuestions.some(remQ => remQ.q_en === poolQ.q_en)
      );
      
      // 5. 洗牌並補齊不足的題數
      const shuffledFresh = shuffleArray(freshPool);
      const compensationQuestions = shuffledFresh.slice(0, neededCount);
      
      selectedQuestions = selectedQuestions.concat(compensationQuestions);
      
      // 6. 更新新一輪大循環的未刷題池（扣除剛剛用來補數量的題）
      unplayedPool = unplayedPool.filter(poolQ => 
        !compensationQuestions.some(compQ => compQ.q_en === poolQ.q_en)
      );
    }

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
        const opts = q.options_en.map((optionStr, origIdx) => {
          const m = optionStr.match(/^([A-Ea-e])\s*[\)\.：:－\-\s]*\s*(.*)$/);
          const origLetter = m ? m[1].toUpperCase() : '';
          const textEn = m ? m[2] : optionStr;
          const textZh = (q.options_zh && Array.isArray(q.options_zh)) ? (q.options_zh[origIdx] || '') : '';
          return { optionStr, origIdx, origLetter, textEn, textZh };
        });

        // 智慧判斷是否包含地雷關鍵字
        const hasStaticKeywords = opts.some(opt => {
          const text = (opt.textEn + opt.textZh).toLowerCase();
          return text.includes('both') || 
                 text.includes('all of') || 
                 text.includes('equally likely') ||
                 text.includes('all the above') ||
                 text.includes('none of') ||
                 text.includes('以上皆') || 
                 text.includes('以下皆') ||
                 text.includes('均可') ||
                 text.includes('皆有可能');
        });

        const shuffledOpts = hasStaticKeywords ? opts : shuffleArray(opts);

        const displayOptions = shuffledOpts.map((optObj, idx) => ({
          displayLabel: String.fromCharCode(65 + idx),
          origLetter: optObj.origLetter,
          text_en: optObj.textEn,
          text_zh: optObj.textZh
        }));
        q._displayOptions = displayOptions;

        displayOptions.forEach((opt, oIndex) => {
          const uniqueId = `q_${qIndex}_opt_${oIndex}`;
          optionsHTML += `
            <div class="mb-2">
              <input type="radio" name="quiz_q_${qIndex}" id="${uniqueId}" value="${opt.origLetter}" class="btn-check" required>
              <label class="btn btn-outline-secondary text-start font-weight-bold w-100 py-3 px-4 shadow-sm" for="${uniqueId}" style="border-width: 2px; color: #020202;">
                <strong class="me-2">${opt.displayLabel}.</strong> ${opt.text_en}
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
        const isAlreadyAdded = cumulativeWrongQuestions.some(cq => cq.q_en === q.q_en);
        if (!isAlreadyAdded) cumulativeWrongQuestions.push(q);
      }

      let optionsReviewHTML = '<div class="my-3 ps-2 border-start border-2 text-secondary small">';
      let chineseOptionsListHTML = ''; 

      if (q._displayOptions && Array.isArray(q._displayOptions)) {
        q._displayOptions.forEach(opt => {
          const currentOptLetter = opt.origLetter || '';
          let optBadge = '';
          if (currentOptLetter === q.answer) optBadge = '<span class="badge bg-success me-1">正確答案</span>';
          else if (currentOptLetter === userAnswer && !isCorrect) optBadge = '<span class="badge bg-danger me-1">您的選擇</span>';

          optionsReviewHTML += `<div class="py-2">${optBadge} <strong class="me-2">${opt.displayLabel}.</strong> ${opt.text_en}</div>`;
          
          if (opt.text_zh) {
            const cleanZhText = opt.text_zh.replace(/^([A-Ea-e])\s*[\)\.：:－\-\s]*/, '');
            chineseOptionsListHTML += `<div class="py-1">${opt.displayLabel}. ${cleanZhText}</div>`;
          }
        });
      }
      optionsReviewHTML += '</div>';

      let chineseOptionsBlockHTML = '';
      if (chineseOptionsListHTML !== '') {
        chineseOptionsBlockHTML = `
          <p class="mb-1 mt-3 fw-bold text-dark">【中文選項參考】</p>
          <div class="small text-muted" style="color: #6c757d !important; line-height: 1.5;">
            ${chineseOptionsListHTML}
          </div>`;
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

      htmlDetails += `
        <div class="card shadow-sm mb-4 border-start border-4 ${isCorrect ? 'border-success' : 'border-danger'}">
          <div class="card-body p-4 bg-white rounded">
            <h5 class="fw-bold text-dark mb-2">Q${qIndex + 1}. ${q.q_en || ''}</h5>
            ${optionsReviewHTML}
            
            <div class="alert alert-secondary py-3 px-3 my-3 fs-6 text-muted rounded-3 border-0">
              <p class="mb-1 fw-bold text-dark">【中文題目翻譯】</p>
              <p class="mb-0">${q.q_zh || '（無中文翻譯）'}</p>
              ${chineseOptionsBlockHTML}
            </div>

            <div class="d-flex align-items-center gap-2 my-2">
              <span class="fs-5 ${isCorrect ? 'text-success fw-bold' : 'text-danger fw-bold'}">
                您的回答：選項 (${(() => {
                  const userOpt = q._displayOptions.find(o => o.origLetter === userAnswer);
                  return userOpt ? userOpt.displayLabel : (userAnswer || '未答');
                })()}) ${isCorrect ? ' ➔ 答對了！' : ' ➔ 答錯了'}
              </span>
            </div>
            ${(() => {
              if (isCorrect) return '';
              const ansLetter = String(q.answer || '').trim().toUpperCase();
              const correctOpt = q._displayOptions.find(o => (o.origLetter || '').toUpperCase() === ansLetter);
              if (correctOpt) {
                const cleanZh = correctOpt.text_zh ? correctOpt.text_zh.replace(/^([A-Ea-e])\s*[\)\.：:－\-\s]*/, '') : '';
                return `<p class="text-success fw-bold fs-5 mb-2">正確答案是：${correctOpt.displayLabel}. ${correctOpt.text_en}${cleanZh ? ' — ' + cleanZh : ''}</p>`;
              } else {
                return `<p class="text-success fw-bold fs-5 mb-2">正確答案是：選項 (${q.answer})</p>`;
              }
            })()}
            ${explanationHTML}
          </div>
        </div>
      `;
    });

    const totalQuestions = selectedQuestions.length;
    const finalScore = totalQuestions > 0 ? Math.round((correctCount / totalQuestions) * 100) : 0;

    scoreText.innerHTML = `${finalScore} <span class="fs-4 text-white ms-2">（答對 ${correctCount} 題 / 共 ${totalQuestions} 題）</span>`;
    resultDetails.innerHTML = htmlDetails;

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
  // 8. 匯出累積錯題本
  // ==========================================
  exportWrongBtn.addEventListener('click', function() {
    if (cumulativeWrongQuestions.length === 0) return;
    const targetFunction = window.exportWrongQuestionsExcel || (typeof exportWrongQuestionsExcel === 'function' ? exportWrongQuestionsExcel : null);
    
    if (!targetFunction) {
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
      } catch(sheetErr) { return; }
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
    
    updateQuizCountUI();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

});