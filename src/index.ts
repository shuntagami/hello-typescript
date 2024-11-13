import ExcelJS from "exceljs";

// 型定義
interface CheckItem {
  value_name: string;
}

interface Header {
  part_name: string;
  items: CheckItem[];
}

interface InspectionResult {
  id: number;
  results: {
    [key: string]: number[]; // 1: OK(○), 2: NG(×), 3: Excluded(-)
  };
}

interface ColumnPosition {
  start: number;
  end: number;
}

interface ColumnPositions {
  [key: string]: ColumnPosition;
}

async function generateRebarInspectionExcel(): Promise<void> {
  // サンプルデータ
  const headers: Header[] = [
    {
      part_name: "主筋",
      items: [
        { value_name: "項目1" },
        { value_name: "項目2" },
        { value_name: "項目3" },
      ],
    },
    {
      part_name: "帯筋",
      items: [{ value_name: "項目A" }, { value_name: "項目B" }],
    },
  ];

  const records: InspectionResult[] = [
    { id: 1, results: { 主筋: [1, 2, 1], 帯筋: [1, 1] } },
    { id: 2, results: { 主筋: [1, 1, 1], 帯筋: [2, 1] } },
  ];

  // Excelワークブックの作成
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("配筋検査");

  // 列幅の設定
  worksheet.getColumn(1).width = 10; // 図面エリア用

  // 列位置の計算
  let currentColumn = 2; // 図面エリアの後から開始
  const columnPositions: ColumnPositions = {};

  headers.forEach((header) => {
    columnPositions[header.part_name] = {
      start: currentColumn,
      end: currentColumn + header.items.length - 1,
    };
    currentColumn += header.items.length;
  });

  // ヘッダー行の生成
  // 1行目: 部位名（マージセル）
  headers.forEach((header) => {
    const startCol = columnPositions[header.part_name].start;
    const endCol = columnPositions[header.part_name].end;

    // マージセルの作成
    worksheet.mergeCells(1, startCol, 1, endCol);

    // 部位名の設定
    const cell = worksheet.getCell(1, startCol);
    cell.value = header.part_name;
    cell.alignment = { horizontal: "center" };
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  });

  // 2行目: チェック項目名
  headers.forEach((header) => {
    const startCol = columnPositions[header.part_name].start;
    header.items.forEach((item, index) => {
      const cell = worksheet.getCell(2, startCol + index);
      cell.value = item.value_name;
      cell.alignment = { horizontal: "center" };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });
  });

  // データ行の生成
  records.forEach((record, index) => {
    const rowNumber = index + 3; // ヘッダー2行分オフセット

    headers.forEach((header) => {
      const startCol = columnPositions[header.part_name].start;
      const results = record.results[header.part_name];

      if (results) {
        results.forEach((result, colIndex) => {
          const cell = worksheet.getCell(rowNumber, startCol + colIndex);
          cell.value = result === 1 ? "○" : result === 2 ? "×" : "－";
          cell.alignment = { horizontal: "center" };
          cell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        });
      }
    });
  });

  // ファイルとして保存
  await workbook.xlsx.writeFile("rebar_inspection.xlsx");
}

// 関数を実行
generateRebarInspectionExcel()
  .then(() => {
    console.log("Excel file has been created successfully!");
  })
  .catch((error) => {
    console.error("Error creating Excel file:", error);
  });
