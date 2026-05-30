// 🌟 基礎原生模式優化：已拔除 import 與 export 語法，欄位與命名邏輯 100% 保持原樣

function exportWrongQuestionsExcel(wrongQuestions) {
  if (!wrongQuestions || wrongQuestions.length === 0) return;

  console.log("📥 【全域匯出】偵測到點擊，開始將錯題匯出成 Excel...", wrongQuestions);

  // 完美保留你原本設計的欄位與格式
  const excelRows = [
    ["題號", "英文題目", "英文選項", "正確答案", "中文題目", "中文選項", "解說"]
  ];

  wrongQuestions.forEach((q, index) => {
    excelRows.push([
      index + 1,
      q.q_en,
      q.options_en.join('\n'), 
      q.answer,
      q.q_zh,
      q.options_zh.join('\n'),
      q.explanation
    ]);
  });

  try {
    // 這裡會自動無縫接軌 index.html 引入的 SheetJS (XLSX) 全域物件
    const worksheet = XLSX.utils.aoa_to_sheet(excelRows);
    const workbook = XLSX.utils.book_new();
    
    XLSX.utils.book_append_sheet(workbook, worksheet, "錯題複習");

    const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, "");
    XLSX.writeFile(workbook, `專案管理_錯題本_${dateStr}.xlsx`);
    
    console.log(`✅ [匯出成功] 已成功下載檔案：專案管理_錯題本_${dateStr}.xlsx`);
  } catch (error) {
    console.error("❌ [匯出崩潰] 產生 Excel 時發生未知錯誤:", error);
  }
}