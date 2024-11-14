import ExcelJS from "exceljs";

interface ChecklistItem {
  name: string;
}

interface ChecklistHeader {
  title: string;
  items: ChecklistItem[];
}

const now = new Date();
const dateStr = now.toISOString().slice(0, 10).replace(/-/g, "");
const timeStr = now.toISOString().slice(11, 19).replace(/:/g, "");
const fileName = `results/${dateStr}_${timeStr}.xlsx`;

async function createTemplateStructure(
  headers: ChecklistHeader[]
): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("checklist");

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

  // 固定ヘッダーのセル設定（縦書き）
  const setVerticalText = (cell: ExcelJS.Cell) => {
    cell.alignment = {
      vertical: "top",
      horizontal: "center",
      textRotation: "vertical",
    };
  };

  const setTextCenter = (cell: ExcelJS.Cell) => {
    cell.alignment = {
      vertical: "middle",
      horizontal: "center",
    };
  };

  // 固定ヘッダーのセル設定
  const idCell = worksheet.getCell("I1");
  const symbolCell = worksheet.getCell("J1");
  setTextCenter(idCell);
  setTextCenter(symbolCell);
  worksheet.getCell("I1").value = "番号";
  worksheet.getCell("J1").value = "符号";
  worksheet.mergeCells("I1:I8");
  worksheet.mergeCells("J1:J8");

  // チェックリストヘッダーの動的設定
  let currentColumn = "K".charCodeAt(0);
  headers.forEach((header) => {
    const startCol = String.fromCharCode(currentColumn);
    const endCol = String.fromCharCode(currentColumn + header.items.length - 1);

    // タイトル行のマージ（項目が1つ以上ある場合）
    if (header.items.length > 0) {
      worksheet.mergeCells(`${startCol}1:${endCol}1`);
    }
    const titleCell = worksheet.getCell(`${startCol}1`);
    titleCell.value = header.title;
    setTextCenter(titleCell);

    // 項目名の設定
    header.items.forEach((item, index) => {
      const col = String.fromCharCode(currentColumn + index);
      worksheet.mergeCells(`${col}2:${col}8`);
      const cell = worksheet.getCell(`${col}2`);
      setVerticalText(cell);
      cell.value = item.name;
      worksheet.getColumn(col).width = 8; // 各列の幅を設定
    });

    currentColumn += header.items.length;
  });

  // EOB と END の設定
  const lastCol = String.fromCharCode(currentColumn);
  for (let i = 9; i <= 32; i++) {
    worksheet.getCell(`${lastCol}${i}`).value = "EOB";
  }
  worksheet.getCell("A33").value = "END";

  // マーカー値の設定
  worksheet.getCell("A1").value = "#{orders.site_name}";
  worksheet.getCell("D1").value = "#{operation_categories.name}";
  worksheet.getCell("A2").value = "検査員 #{sheet_inspectors.names}";
  worksheet.getCell("D2").value = "#{blueprints.name:sheets.name}";
  worksheet.getCell("A3").value = "#{blueprints.image}";

  // ファイルとして保存
  await workbook.xlsx.writeFile(fileName);
}

// テスト用のサンプルデータ
const sampleHeaders: ChecklistHeader[] = [
  {
    title: "主筋",
    items: [
      { name: "項目1" },
      { name: "項目2" },
      { name: "項目3" },
      { name: "項目4" },
      { name: "項目5" },
    ],
  },
  {
    title: "帯筋",
    items: [{ name: "項目A" }, { name: "項目B" }, { name: "項目C" }],
  },
  {
    title: "継手",
    items: [{ name: "確認1" }, { name: "確認2" }],
  },
];

// 関数を実行
createTemplateStructure(sampleHeaders)
  .then(() => {
    console.log(`${fileName}`);
  })
  .catch((error) => {
    console.error("Error creating template structure:", error);
  });
