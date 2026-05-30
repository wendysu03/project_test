// 🌟 基礎原生模式優化：已拔除 import 與 export 語法，欄位命名與解析邏輯 100% 保持原樣

function parseExcelWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = function(evt) {
      try {
        const data = evt.target.result;
        // 核心邏輯保持原樣
        const workbook = XLSX.read(data, { type: 'binary' });
        
        const resultData = {
          before: [],      
          net: [],         
          wrongFile: [],   
          isWrongFileMode: false
        };

        workbook.SheetNames.forEach(sheetName => {
          const worksheet = workbook.Sheets[sheetName];
          if (!worksheet) return;
          
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
          let targetList = null;
          
          if (sheetName.includes('課前')) {
            targetList = resultData.before;
          } else if (sheetName.includes('網路資源')) {
            targetList = resultData.net;
          } else if (sheetName.includes('錯題') || sheetName.includes('Review')) {
            targetList = resultData.wrongFile;
            resultData.isWrongFileMode = true; 
          }
          
          if (!targetList) return; 
          
          jsonData.forEach((row, index) => {
            if (index === 0 || row.length < 4) return;
            
            let q_en = row[1];        
            let optionsRaw = row[2];  
            let answerKey = row[3];   
            let q_zh = row[4];        
            let optionsZhRaw = row[5];
            let explanation = row[6]; 
            
            if (!q_en || !optionsRaw || !answerKey) return;
            
            answerKey = answerKey.toString().trim().toUpperCase().replace(/[^A-E]/g, "");
            if (answerKey === "") return;
            
            let options_en = optionsRaw.toString().split('\n').map(o => o.trim()).filter(o => o.length > 0);
            let options_zh = optionsZhRaw ? optionsZhRaw.toString().split('\n').map(o => o.trim()).filter(o => o.length > 0) : [];
            
            targetList.push({
              q_en: q_en.toString().trim(),
              options_en: options_en,
              answer: answerKey,
              q_zh: q_zh ? q_zh.toString().trim() : "（此題無中文翻譯）",
              options_zh: options_zh,
              explanation: explanation ? explanation.toString().trim() : ""
            });
          });
        });

        resolve(resultData);
      } catch (err) {
        reject(err);
      }
    };

    reader.onerror = (err) => reject(err);
    reader.readAsBinaryString(file);
  });
}