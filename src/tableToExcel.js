import Parser from "./parser";
import saveAs from "file-saver";
import ExcelJS from "exceljs/dist/es5/exceljs.browser";

const TableToExcel = (function (Parser) {
    let methods = {};

    methods.initWorkBook = function () {
        let wb = new ExcelJS.Workbook();
        return wb;
    };

    methods.initSheet = function (wb, sheetName) {
        let ws = wb.addWorksheet(sheetName);
        return ws;
    };

    methods.save = function (wb, fileName) {
        wb.xlsx.writeBuffer().then(function (buffer) {
            saveAs(
                new Blob([buffer], {
                    type: "application/octet-stream"
                }),
                fileName
            );
        });
    };

    methods.tableToSheet = function (wb, tables, opts) {
        [...tables].forEach((table, index) => {
            let ws = this.initSheet(wb, opts.sheets[index].name);
            ws = Parser.parseDomToTable(ws, table, opts);
        })
        return wb;
    };

    methods.tableToBook = function (tables, opts) {
        let wb = this.initWorkBook();
        wb = this.tableToSheet(wb, tables, opts);
        return wb;
    };

    methods.convert = function (tables, opts = {}) {
        let defaultOpts = {
            name: "export.xlsx",
            autoStyle: false,
            sheets: [{
                name: "Sheet 1"
            }]
        };
        opts = {
            ...defaultOpts,
            ...opts
        };
        let wb = this.tableToBook(tables, opts);
        this.save(wb, opts.name);
    };

    return methods;
})(Parser);

export default TableToExcel;
window.TableToExcel = TableToExcel;