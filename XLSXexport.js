var XLSX = require("xlsx");
const fs = require('fs-extra');
const {
    toSystemPath
} = require('../../../lib/core/path');

//made possible by https://sheetjs.com/

exports.XLSXExport = async function (options) {

    options = this.parse(options);
    let path = this.parse(options.path);
    let data = options.data;
    let sheetname = options.name;
    let excelname = './'+path;
    if (typeof path != 'string')
        throw new Error('export.xlsx: path is required.');
    if (!Array.isArray(data) || !data.length)
        throw new Error('export.xlsx: data is required.');

    //
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "DataSheet");
    var file = XLSX.writeFile(workbook, excelname);

    return path;

}
