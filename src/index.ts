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

const getColumnLetter = (num: number): string => {
  let letter = "";
  while (num >= 0) {
    letter = String.fromCharCode((num % 26) + 65) + letter;
    num = Math.floor(num / 26) - 1;
  }
  return letter;
};

const setTextCenter = (cell: ExcelJS.Cell) => {
  cell.alignment = {
    vertical: "middle",
    horizontal: "center",
  };
  cell.border = {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" },
  };
};

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
    views: [{ state: "normal", style: "pageBreakPreview" }],
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
  worksheet.getColumn("I").width = 5;
  worksheet.getColumn("J").width = 5;

  // 使用するセル範囲を事前に把握
  const lastColumnIndex = constructionPart.checklistTemplates.reduce(
    (sum, template) => sum + template.items.length,
    0
  );

  // 全セルに対してデフォルトスタイルを適用
  for (let row = 1; row <= 32; row++) {
    for (let colIndex = 0; colIndex < lastColumnIndex + 11; colIndex++) {
      const colLetter = getColumnLetter(colIndex);
      const cell = worksheet.getCell(`${colLetter}${row}`);
      setTextCenter(cell);
    }
  }

  // ヘッダー部分のマージセル設定
  worksheet.mergeCells("A1:C1");
  worksheet.mergeCells("D1:H1");
  worksheet.mergeCells("A2:C2");
  worksheet.mergeCells("D2:H2");

  // 図面エリアのマージセル
  worksheet.mergeCells("A3:H32");

  // 固定ヘッダーのセル設定
  worksheet.mergeCells("I1:I8");
  worksheet.mergeCells("J1:J8");
  worksheet.getCell("I1").value = "番号";
  worksheet.getCell("J1").value = "符号";

  // チェックリストヘッダーの動的設定
  let currentColumnIndex = 10; // K列は10番目
  constructionPart.checklistTemplates.forEach((template) => {
    const startCol = getColumnLetter(currentColumnIndex);
    const endCol = getColumnLetter(
      currentColumnIndex + template.items.length - 1
    );

    if (template.items.length > 0) {
      worksheet.mergeCells(`${startCol}1:${endCol}1`);
    }
    const titleCell = worksheet.getCell(`${startCol}1`);
    titleCell.value = template.name;

    // 項目名の設定
    template.items.forEach((item, index) => {
      const col = getColumnLetter(currentColumnIndex + index);
      worksheet.mergeCells(`${col}2:${col}8`);
      const cell = worksheet.getCell(`${col}2`);
      setVerticalText(cell);
      cell.value = item.name;
      worksheet.getColumn(col).width = 8;
    });

    currentColumnIndex += template.items.length;
  });

  // EOB と END の設定
  const lastCol = getColumnLetter(currentColumnIndex);
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
