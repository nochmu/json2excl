#!/usr/bin/env node


/* original data */
var data = [
    ["John", "Seattle"],
    ["Mike",  "Los\nAngeles"],
    ["S",  "B"],
    ["Zach",  "NY"],

];


const XlsxPopulate = require('xlsx-populate');

// Load a new blank workbook
XlsxPopulate.fromBlankAsync()
    .then(workbook => {
        // Modify the workbook.
        var r = workbook.sheet("Sheet1").cell("A1").value(data);
        r.style({
            wrapText: true,
            horizontalAlignment: "justify",
            verticalAlignment: "justify",
            justifyLastLine: true,
        });

        // Write to file.
        return workbook.toFileAsync("./out.xlsx");
    });
