const excel = require('excel4node');

function excelize(obj, path, file, sheet, cb) {
    const titles = Object.keys(obj[0]);
    const width = titles.length;
    const height = obj.length;
    // const workbook = excelbuilder.createWorkbook(path, file);
    const workbook = new excel.Workbook();
    // const sht = workbook.createSheet(sheet, width, height + 1);
    const sht = workbook.addWorksheet(sheet);
    const style = workbook.createStyle({
        font: {
            bold: true,
            color: 'ffFF00',
            size: 14
        },
        // numberFormat: '$#,##0.00; ($#,##0.00); -',
    });
    let col, j, k;

    for (col = 1; col <= width; col++) {
        sht.cell(1, col).string(titles[col - 1]).style(style);
    }
    for (j = 1; j <= height; j++) {
        for (col = 1; col <= width; col++) {
            k = titles[col - 1];
            sht.cell(j + 1, col).string(obj[j - 1][k])
        }
    }
    workbook.write(file);
    cb()
}

module.exports = excelize;
