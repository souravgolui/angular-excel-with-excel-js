import { Injectable } from '@angular/core';
import { Workbook } from 'exceljs';
import { header, mockData } from './constant';
import * as fs from 'file-saver';
import { style } from '@angular/animations';

@Injectable({
  providedIn: 'root'
})
export class ExcelService {
  async generateExcel(name: string) {
    let workbook = new Workbook();
    let worksheet = workbook.addWorksheet('Car Data');

    // Set Title
    let titleRow = worksheet.addRow([name]);
    titleRow.font = {
      name: 'Vardana',
      family: 4,
      size: 16,
      underline: 'double',
      bold: true
    };
    titleRow.alignment = { horizontal: 'center', vertical: 'middle' };
    titleRow.height = 50;
    worksheet.addRow([]);

    // Add Header row
    const headerRow = worksheet.addRow(header);

    worksheet.mergeCells(`A1:F1`);
    // Cell Style and accessing all cell of header row
    headerRow.eachCell((cell, number) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFF00' },
        bgColor: { argb: 'FF0000FF' }
      };

      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });
    headerRow.font = { name: 'Vardana', bold: true, size: 13 };
    headerRow.height = 50;
    headerRow.alignment = { horizontal: 'center', vertical: 'middle' };

    // All Data to excel
    mockData.forEach(d => {
      const row = worksheet.addRow(d);
      row.font = { name: 'Vardana', size: 12 };
      row.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
      const qty = row.getCell(5);
      let color = 'FF99FF99';
      if (+qty.value < 500) {
        color = 'FF9999';
      }
      qty.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: color }
      };
    });

    // add size for specific column
    worksheet.getColumn(3).width = 30;
    worksheet.getColumn(4).width = 30;
    worksheet.getColumn(5).width = 25;
    worksheet.addRow([]);

    const data = await workbook.xlsx.writeBuffer();
    this.saveExcelDoc(data, name);
  }

  saveExcelDoc(data: ArrayBuffer, name: string) {
    let blob = new Blob([data], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    fs.saveAs(blob, `${name}.xlsx`);
  }
}
