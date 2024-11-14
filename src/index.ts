import ExcelJS from "exceljs";

async function createTemplateStructure(): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Template");

  // 基本の列幅設定
  const columns = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"];
  columns.forEach((col) => {
    worksheet.getColumn(col).width = 10;
  });

  // ヘッダー部分のマージセル設定
  worksheet.mergeCells("A1:C1"); // orders.site_name
  worksheet.mergeCells("D1:H1"); // operation_categories.name
  worksheet.mergeCells("A2:C2"); // 検査員
  worksheet.mergeCells("D2:H2"); // blueprints.name:sheets.name

  // 図面エリアのマージセル
  worksheet.mergeCells("A3:H32");

  // 固定ヘッダーのセル設定
  worksheet.getCell("I1").value = "番";
  worksheet.getCell("J1").value = "符";
  worksheet.mergeCells("I1:I2");
  worksheet.mergeCells("J1:J2");

  // EOB と END の設定
  for (let i = 3; i <= 31; i++) {
    worksheet.getCell(`L${i}`).value = "EOB";
  }
  worksheet.getCell("A32").value = "END";

  // 図面エリアのプレースホルダー
  worksheet.getCell("A3").value = "#{blueprints.image}";

  // マーカー値の設定
  worksheet.getCell("A1").value = "#{orders.site_name}";
  worksheet.getCell("D1").value = "#{operation_categories.name}";
  worksheet.getCell("A2").value = "検査員 #{sheet_inspectors.names}";
  worksheet.getCell("D2").value = "#{blueprints.name:sheets.name}";

  // ファイルとして保存
  await workbook.xlsx.writeFile("template_structure.xlsx");
}

// 関数を実行
createTemplateStructure()
  .then(() => {
    console.log("Template structure has been created successfully!");
  })
  .catch((error) => {
    console.error("Error creating template structure:", error);
  });
