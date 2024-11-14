import ExcelJS from "exceljs";

const now = new Date();
const dateStr = now.toISOString().slice(0, 10).replace(/-/g, ""); // YYYYMMDD形式
const timeStr = now.toISOString().slice(11, 19).replace(/:/g, ""); // HHMMSS形式
const fileName = `${dateStr}_${timeStr}.xlsx`;

async function createTemplateStructure(): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("checklist");

  // サンプルデータ
  const checklistHeaders = [
    {
      title: "主筋",
      items: ["項目1", "項目2", "項目3"],
    },
    {
      title: "帯筋",
      items: ["項目A", "項目B"],
    },
  ];

  // 基本の列幅設定
  worksheet.getColumn("A").width = 10;
  worksheet.getColumn("I").width = 5; // 番号列
  worksheet.getColumn("J").width = 5; // 符号列

  // ヘッダー部分のマージセル設定
  worksheet.mergeCells("A1:C1"); // orders.site_name
  worksheet.mergeCells("D1:H1"); // operation_categories.name
  worksheet.mergeCells("A2:C2"); // 検査員
  worksheet.mergeCells("D2:H2"); // blueprints.name:sheets.name

  // 図面エリアのマージセル
  worksheet.mergeCells("A3:H32");

  // 固定ヘッダーのセル設定
  worksheet.getCell("I1").value = "番号";
  worksheet.getCell("J1").value = "符号";
  worksheet.mergeCells("I1:I8");
  worksheet.mergeCells("J1:J8");

  // 固定ヘッダーのセル設定（縦書き）
  const setVerticalText = (cell: ExcelJS.Cell) => {
    cell.alignment = {
      vertical: "top",
      horizontal: "center",
      textRotation: "vertical", // 255は文字を縦書きにする特別な値
    };
  };

  // チェックリストヘッダーの設定
  let currentColumn = "K";
  checklistHeaders.forEach((header) => {
    const startCol = currentColumn;
    const endCol = String.fromCharCode(
      currentColumn.charCodeAt(0) + header.items.length - 1
    );

    // タイトル行のマージ
    if (header.items.length > 1) {
      worksheet.mergeCells(`${startCol}1:${endCol}1`);
    }
    worksheet.getCell(`${startCol}1`).value = header.title;

    // 項目名の設定
    header.items.forEach((item, index) => {
      const col = String.fromCharCode(startCol.charCodeAt(0) + index);
      worksheet.mergeCells(`${col}2:${col}8`);
      const cell = worksheet.getCell(`${col}2`);
      setVerticalText(cell);
      cell.value = item;
    });

    currentColumn = String.fromCharCode(endCol.charCodeAt(0) + 1);
  });

  // EOB と END の設定
  const lastCol = String.fromCharCode(currentColumn.charCodeAt(0));
  for (let i = 3; i <= 31; i++) {
    worksheet.getCell(`${lastCol}${i}`).value = "EOB";
  }
  worksheet.getCell("A32").value = "END";

  // マーカー値の設定
  worksheet.getCell("A1").value = "#{orders.site_name}";
  worksheet.getCell("D1").value = "#{operation_categories.name}";
  worksheet.getCell("A2").value = "検査員 #{sheet_inspectors.names}";
  worksheet.getCell("D2").value = "#{blueprints.name:sheets.name}";
  worksheet.getCell("A3").value = "#{blueprints.image}";

  // ファイルとして保存
  await workbook.xlsx.writeFile(fileName);
}

// 関数を実行
createTemplateStructure()
  .then(() => {
    console.log(`${fileName}`);
  })
  .catch((error) => {
    console.error("Error creating template structure:", error);
  });
