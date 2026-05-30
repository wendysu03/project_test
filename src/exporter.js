import * as XLSX from 'xlsx';

export function exportWrongQuestionsExcel(wrongQuestions) {
  if (!wrongQuestions || wrongQuestions.length === 0) return;

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

  const worksheet = XLSX.utils.aoa_to_sheet(excelRows);
  const workbook = XLSX.utils.book_new();
  
  XLSX.utils.book_append_sheet(workbook, worksheet, "錯題複習");

  const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, "");
  XLSX.writeFile(workbook, `專案管理_錯題本_${dateStr}.xlsx`);
}