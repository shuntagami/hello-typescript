import ExcelJS from "exceljs";
import { createWriteStream } from "fs";

interface ChecklistItem {
  id: number;
  name: string;
}

interface ChecklistTemplate {
  id: number;
  name: string;
  items: ChecklistItem[];
}

interface ConstructionPart {
  id: number;
  name: string;
  checklistTemplates: ChecklistTemplate[];
}

async function createTemplateStructure(
  constructionPart: ConstructionPart,
  filePath: string
): Promise<void> {
  const stream = createWriteStream(filePath);
  const options = {
    useStyles: true,
    stream,
  };
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter(options);
  const worksheet = workbook.addWorksheet("checklist", {
    views: [{}],
    pageSetup: {
      paperSize: 8, // A3
      orientation: "landscape",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 1,
      horizontalCentered: true,
      verticalCentered: true,
    },
  });

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
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  };

  const setTextCenter = (cell: ExcelJS.Cell) => {
    cell.alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    // 枠線設定を追加
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
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
  constructionPart.checklistTemplates.forEach((template) => {
    const startCol = String.fromCharCode(currentColumn);
    const endCol = String.fromCharCode(
      currentColumn + template.items.length - 1
    );

    // タイトル行のマージ（項目が1つ以上ある場合）
    if (template.items.length > 0) {
      worksheet.mergeCells(`${startCol}1:${endCol}1`);
    }
    const titleCell = worksheet.getCell(`${startCol}1`);
    titleCell.value = template.name;
    setTextCenter(titleCell);

    // 項目名の設定
    template.items.forEach((item, index) => {
      const col = String.fromCharCode(currentColumn + index);
      worksheet.mergeCells(`${col}2:${col}8`);
      const cell = worksheet.getCell(`${col}2`);
      setVerticalText(cell);
      cell.value = item.name;
      worksheet.getColumn(col).width = 8; // 各列の幅を設定
    });

    currentColumn += template.items.length;
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

  await workbook.commit();
}

// テスト用のサンプルデータ
const sampleConstructionPart: ConstructionPart = {
  id: 1,
  name: "基礎",
  checklistTemplates: [
    {
      id: 1,
      name: "主筋",
      items: [
        { id: 1, name: "項目1" },
        { id: 2, name: "項目2" },
        { id: 3, name: "項目3" },
        { id: 4, name: "項目4" },
        { id: 5, name: "項目5" },
      ],
    },
    {
      id: 2,
      name: "帯筋",
      items: [
        { id: 6, name: "項目A" },
        { id: 7, name: "項目B" },
        { id: 8, name: "項目C" },
      ],
    },
    {
      id: 3,
      name: "継手",
      items: [
        { id: 9, name: "確認1" },
        { id: 10, name: "確認2" },
        { id: 11, name: "確認3" },
      ],
    },
  ],
};

const now = new Date();
const dateStr = now.toISOString().slice(0, 10).replace(/-/g, "");
const timeStr = now.toISOString().slice(11, 19).replace(/:/g, "");
const filePath = `results/${dateStr}_${timeStr}.xlsx`;

// 関数を実行
createTemplateStructure(sampleConstructionPart, filePath)
  .then(() => {
    console.log(`${filePath}`);
  })
  .catch((error) => {
    console.error("Error creating template structure:", error);
  });
