import { Injectable } from '@angular/core';
import * as FileSaver from 'file-saver';
import * as XLSX from 'xlsx';
import { DatePipe } from '@angular/common';
import { IExcelWorkSheetObj } from '../utility/app-excel-obj';

const EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';
const CSV_EXTENSION = '.csv';

@Injectable({
    providedIn: 'root'
})
export class ExportExcelService {

    constructor(private datePipe: DatePipe) { }

    public exportAsExcelFile(json: any[], excelFileName: string): void {
        const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(json);
        const workbook: XLSX.WorkBook = { Sheets: { data: worksheet }, SheetNames: ['data'] };
        const excelBuffer: any = XLSX.write(workbook, { bookType: 'xlsx', type: 'array', Props: { Author: 'Hive' } });
        this.saveAsExcelFile(excelBuffer, excelFileName);
    }

    public exportAsExcelFileWithHeadersTemplate(json: any[], excelFileName: string, headers: string[]): void {
        const worksheet: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet([headers]);
        XLSX.utils.sheet_add_json(worksheet, json, { skipHeader: true, origin: 'A2' });
        const workbook: XLSX.WorkBook = { Sheets: { data: worksheet }, SheetNames: ['data'] };

        const excelBuffer: any = XLSX.write(workbook, { bookType: 'xlsx', type: 'array', Props: { Author: 'Hive' } });
        this.saveAsExcelFile(excelBuffer, excelFileName);
    }

    public exportAsCsvWithHeadersTemplate(json: any[], fileName: string, headers: string[]): void {
        const worksheet: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet([headers]);
        XLSX.utils.sheet_add_json(worksheet, json, { skipHeader: true, origin: 'A2' });
        const workbook = XLSX.utils.book_new();

        // Convert to CSV
        XLSX.utils.book_append_sheet(workbook, worksheet, 'data'); // Add worksheet to book
        XLSX.writeFile(workbook, fileName + this.datePipe.transform(new Date(), 'yyyyMMdd') + CSV_EXTENSION);
    }

    public exportAsExcelWorksheets(excelWorkSheetObj: IExcelWorkSheetObj[], excelFileName: string): void {

        const workbook: XLSX.WorkBook = XLSX.utils.book_new();

        excelWorkSheetObj.forEach(ws => {
            const wsName: string = ws.WorkSheet_Name;
            const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(ws.WorkSheet_Obj);
            XLSX.utils.book_append_sheet(workbook, worksheet, wsName);
        });

        const excelBuffer: any = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        this.saveAsExcelFile(excelBuffer, excelFileName);
    }

    private saveAsExcelFile(buffer: any, fileName: string): void {
        const data: Blob = new Blob([buffer], { type: EXCEL_TYPE });
        FileSaver.saveAs(data, fileName + this.datePipe.transform(new Date(), 'yyyyMMdd') + EXCEL_EXTENSION);
    }
}
