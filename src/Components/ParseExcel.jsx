import React from "react";
import { read, utils } from "xlsx";
import exceljs from "exceljs";

export const ParseExcel = () => {
  const readExcel = (file, sheetName) => {
    const parsedExcel = read(file, { type: "buffer" });
    console.log(parsedExcel);
    const workbook = {};

    if (sheetName) {
      const sheetData = utils.sheet_to_json(parsedExcel.Sheets[sheetName], {
        header: 1,
        blankrows: false,
      });
      workbook.sheetName = {
        sheetName: sheetName,
        headers: sheetData.shift(),
        data: sheetData,
      };
    } else {
      parsedExcel.SheetNames.forEach((sheet) => {
        const sheetData = utils.sheet_to_json(parsedExcel.Sheets[sheet], {
          header: 1,
          blankrows: false,
        });
        workbook.sheet = {
          sheetName: sheet,
          headers: sheetData.shift(),
          data: sheetData,
        };
      });
    }

    return workbook;
  };

  const formatWorksheet = (worksheet, formatOptions) => {
    worksheet.getRow(1).font = { bold: formatOptions.boldHeaders }
    formatOptions.colWidths.forEach((target) => {
      worksheet.getColumn(target.number).width = target.width
    })
    formatOptions.rowHeight.forEach((target) => {
      worksheet.getRow(target.number).height = target.height
    })
  }

  const writeExcel = async () => {
    const workbook = new exceljs.Workbook();
    const filename = 'testsheet'
    const sheets = [
      {
        sheetName: "firstsheet",
        headers: ["h1", "h2", "h3"],
        data: [
          [1, 1, 1],
          [2, 2, 2],
          [3, 3, 3],
        ],
      },
      {
        sheetName: "SECONDsheet",
        headers: ["h4", "h5", "h6"],
        data: [
          [11, 11, 11],
          [22, 22, 22],
          [33, 33, 33],
        ],
      },
    ];

    const format = {
      boldHeaders: true,
      colWidths: [{ number: 1, width: 20 },{ number: 2, width: 40 }],
      rowHeight: [{ number: 1, height: 10 },{ number: 2, height: 20 }]
    }

    sheets.forEach((sheet) => {
      const worksheet = workbook.addWorksheet(sheet.sheetName);
      worksheet.addRow(sheet.headers);
      worksheet.addRows(sheet.data);
      formatWorksheet(worksheet, format);
    });

    await workbook.xlsx
      .writeBuffer()
      .then((data) => {
        const blob = new Blob([data], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        const anchor = document.createElement("a");
        const url = URL.createObjectURL(blob);
        anchor.href = url;
        anchor.download = filename + ".xlsx";
        document.body.appendChild(anchor);
        anchor.click();
        document.body.removeChild(anchor);
        URL.revokeObjectURL(url);
      })
      .catch((err) => console.log(err));
  };

  const handleFile = async (event) => {
    const file = event.target.files[0];
    const data = await file.arrayBuffer();
    console.log(readExcel(data));
    await writeExcel();
  };

  return (
    <div>
      <h1>ParseExcel</h1>

      <input type="file" onChange={(e) => handleFile(e)} />
    </div>
  );
};
