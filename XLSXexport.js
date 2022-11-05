var XLSX = require("xlsx");
const fs = require('fs-extra');
const {
    toSystemPath
} = require('../../../lib/core/path');

//made possible by https://sheetjs.com/

exports.XLSXExport = async function (options) {

    options = this.parse(options);
    let path1 = this.parse(options.path);
    let data = options.data;
    let sheetname = options.sheetname.substr(0, 30).replace(/[^a-zA-Z0-9 ]/g, '');
    let excelname = options.filename;
    let filetype = options.filetype;
    let datachek = false;

    //check if data , path  set or not
    if (typeof path1 != 'string')
        throw new Error('path: path is required.');
    if (!data) {
        throw new Error('data: data is required.');
    } else if (!Array.isArray(data)) {
        datachek = true;
    } else {
        datachek = false;
    }

    //check if file type set
    if (filetype) {
        //check if file type start with dot or not
        if (filetype.startsWith(".")) {
            filetype = filetype;
        } else {
            filetype = "." + filetype;
        }
    } else {
        filetype = '.xlsx';
    }
    //set th path of file
    let path = path1 + excelname + filetype;

    //check if file type start with dot or not
    if (path1.startsWith(".")) {
        path1 = path + excelname + filetype;
    } else {
        path1 = "." + path1 + excelname + filetype;
    }
    // genrate exported file
    if (datachek) {
        const worksheet = XLSX.utils.json_to_sheet([data]);
		 const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetname);
    var file = XLSX.writeFile(workbook, path1);
    } else {
        const worksheet = XLSX.utils.json_to_sheet(data);
		 const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetname);
    var file = XLSX.writeFile(workbook, path1);
    }

    return path;

}
