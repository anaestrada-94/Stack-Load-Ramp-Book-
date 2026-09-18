import ExcelJS from 'exceljs';
import {MVA_THRESHOLD, toNumber} from './utils';

const COLOR = {
    navy: 'FF1D4F91',
    navyDark: 'FF163F75',
    blueLight: 'FF21A0DB',
    blueWash: 'FFE8F1FA',
    white: 'FFFFFFFF',
    gray50: 'FFF6F8FC',
    gray200: 'FFDADDE6',
    text: 'FF31353E',
    muted: 'FF616670',
};

const LEFT_HEADERS = [
    {label: 'Campus', width: 24, align: 'left'},
    {label: 'Metro', width: 16, align: 'left'},
    {label: 'Supplier Building Code', width: 22, align: 'left'},
    {label: 'Land Status', width: 18, align: 'center'},
    {label: 'Rezoning Status', width: 18, align: 'center'},
    {label: 'Water Capacity', width: 18, align: 'center'},
    {label: 'Power Confidence Level', width: 22, align: 'left'},
    {label: 'Gross MW', width: 12, align: 'center'},
    {label: 'Live MW', width: 12, align: 'center'},
    {label: 'Pipeline MW', width: 13, align: 'center'},
];

const thinBorder = {
    top: {style: 'thin', color: {argb: COLOR.gray200}},
    left: {style: 'thin', color: {argb: COLOR.gray200}},
    bottom: {style: 'thin', color: {argb: COLOR.gray200}},
    right: {style: 'thin', color: {argb: COLOR.gray200}},
};

function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.rel = 'noopener';
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function fileStamp() {
    const now = new Date();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    return `${now.getFullYear()}-${month}-${day}`;
}

function exportBasename(view) {
    const viewLabel = view === 'director' ? 'Director' : 'Enterprise';
    return `STACK-Land-Bank-${viewLabel}-${fileStamp()}`;
}

function quarterCols(years) {
    return years.flatMap((year) => year.quarters);
}

function numberOrBlank(value) {
    if (value == null || value === '') return '';
    const n = Number(value);
    return Number.isFinite(n) ? n : '';
}

function sumField(rows, getter) {
    let total = 0;
    let hasValue = false;
    for (const row of rows) {
        const value = getter(row);
        if (value == null || value === '') continue;
        total += toNumber(value);
        hasValue = true;
    }
    return hasValue ? total : '';
}

function rowValues(row, years) {
    return [
        row.campus || '',
        row.metro || '',
        row.supplierBuildingCode || '',
        row.landStatus || '',
        row.rezoningStatus || '',
        row.waterCapacity || '',
        row.powerConfidence || '',
        numberOrBlank(row.grossMw),
        numberOrBlank(row.liveMw),
        numberOrBlank(row.pipelineMw),
        ...quarterCols(years).map((q) => numberOrBlank(row.quarters?.[q.key])),
        numberOrBlank(row.longTermMw),
    ];
}

function totalsValues(rows, years) {
    return [
        'Total',
        '',
        '',
        '',
        '',
        '',
        '',
        sumField(rows, (row) => row.grossMw),
        sumField(rows, (row) => row.liveMw),
        sumField(rows, (row) => row.pipelineMw),
        ...quarterCols(years).map((q) => sumField(rows, (row) => row.quarters?.[q.key])),
        sumField(rows, (row) => row.longTermMw),
    ];
}

function fillSolid(argb) {
    return {type: 'pattern', pattern: 'solid', fgColor: {argb}};
}

function applyRange(sheet, r1, c1, r2, c2, patch) {
    for (let row = r1; row <= r2; row++) {
        for (let col = c1; col <= c2; col++) {
            const cell = sheet.getCell(row, col);
            if (patch.fill) cell.fill = patch.fill;
            if (patch.font) cell.font = patch.font;
            if (patch.alignment) cell.alignment = patch.alignment;
            if (patch.border) cell.border = patch.border;
        }
    }
}

function isNumericCol(col) {
    return col >= 8;
}

function addStyledSheet(workbook, {name, rows, years, view, tabArgb}) {
    const sheet = workbook.addWorksheet(name, {
        properties: {tabColor: {argb: tabArgb}},
        views: [{state: 'frozen', xSplit: 2, ySplit: 5, showGridLines: false}],
        pageSetup: {
            orientation: 'landscape',
            fitToPage: true,
            fitToWidth: 1,
            fitToHeight: 0,
            paperSize: 8,
            margins: {left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2},
        },
    });

    const lastCol = LEFT_HEADERS.length + quarterCols(years).length + 1;
    const viewLabel = view === 'director' ? 'Director view' : 'Enterprise reporting';
    const subtitle =
        view === 'director'
            ? `${name}  ·  Projects view`
            : `${name}  ·  Only Load Ramp records with status Verified are included`;

    sheet.columns = [
        ...LEFT_HEADERS.map((item) => ({width: item.width})),
        ...quarterCols(years).map(() => ({width: 10})),
        {width: 16},
    ];

    sheet.mergeCells(1, 1, 1, lastCol);
    const titleCell = sheet.getCell(1, 1);
    titleCell.value = 'STACK LAND BOOK';
    applyRange(sheet, 1, 1, 1, lastCol, {
        fill: fillSolid(COLOR.navy),
        font: {name: 'Calibri', size: 18, bold: true, color: {argb: COLOR.white}},
        alignment: {vertical: 'middle', horizontal: 'left', indent: 1},
        border: thinBorder,
    });
    sheet.getRow(1).height = 30;

    sheet.mergeCells(2, 1, 2, lastCol);
    sheet.getCell(2, 1).value = viewLabel;
    applyRange(sheet, 2, 1, 2, lastCol, {
        fill: fillSolid(COLOR.blueWash),
        font: {name: 'Calibri', size: 12, bold: true, color: {argb: COLOR.navyDark}},
        alignment: {vertical: 'middle', horizontal: 'left', indent: 1},
        border: thinBorder,
    });
    sheet.getRow(2).height = 20;

    sheet.mergeCells(3, 1, 3, lastCol);
    sheet.getCell(3, 1).value = subtitle;
    applyRange(sheet, 3, 1, 3, lastCol, {
        fill: fillSolid(COLOR.blueWash),
        font: {name: 'Calibri', size: 10, color: {argb: COLOR.muted}},
        alignment: {vertical: 'middle', horizontal: 'left', indent: 1},
        border: thinBorder,
    });
    sheet.getRow(3).height = 18;

    const headTop = 4;
    const headBottom = 5;
    sheet.getRow(headTop).height = 22;
    sheet.getRow(headBottom).height = 20;

    LEFT_HEADERS.forEach((item, index) => {
        const col = index + 1;
        sheet.mergeCells(headTop, col, headBottom, col);
        sheet.getCell(headTop, col).value = item.label;
        applyRange(sheet, headTop, col, headBottom, col, {
            fill: fillSolid(COLOR.navyDark),
            font: {name: 'Calibri', size: 9, bold: true, color: {argb: COLOR.white}},
            alignment: {vertical: 'middle', horizontal: 'center', wrapText: true},
            border: thinBorder,
        });
    });

    let col = LEFT_HEADERS.length + 1;
    years.forEach((yearGroup) => {
        const start = col;
        const end = col + yearGroup.quarters.length - 1;
        if (end > start) sheet.mergeCells(headTop, start, headTop, end);
        sheet.getCell(headTop, start).value = yearGroup.label;
        applyRange(sheet, headTop, start, headTop, end, {
            fill: fillSolid(COLOR.navyDark),
            font: {name: 'Calibri', size: 10, bold: true, color: {argb: COLOR.white}},
            alignment: {vertical: 'middle', horizontal: 'center'},
            border: thinBorder,
        });
        yearGroup.quarters.forEach((quarter, offset) => {
            const quarterCol = start + offset;
            const cell = sheet.getCell(headBottom, quarterCol);
            cell.value = quarter.label;
            cell.fill = fillSolid(COLOR.navy);
            cell.font = {name: 'Calibri', size: 9, bold: true, color: {argb: COLOR.white}};
            cell.alignment = {vertical: 'middle', horizontal: 'center'};
            cell.border = thinBorder;
        });
        col = end + 1;
    });

    sheet.mergeCells(headTop, lastCol, headBottom, lastCol);
    sheet.getCell(headTop, lastCol).value = 'Long-term Delivery';
    applyRange(sheet, headTop, lastCol, headBottom, lastCol, {
        fill: fillSolid(COLOR.navyDark),
        font: {name: 'Calibri', size: 9, bold: true, color: {argb: COLOR.white}},
        alignment: {vertical: 'middle', horizontal: 'center', wrapText: true},
        border: thinBorder,
    });

    rows.forEach((row, index) => {
        const excelRow = headBottom + 1 + index;
        const values = rowValues(row, years);
        const stripe = index % 2 === 1;
        values.forEach((value, valueIndex) => {
            const cell = sheet.getCell(excelRow, valueIndex + 1);
            cell.value = value === '' ? null : value;
            cell.fill = fillSolid(stripe ? COLOR.blueWash : COLOR.white);
            cell.font = {
                name: 'Calibri',
                size: 10,
                bold: valueIndex === 7,
                color: {argb: COLOR.text},
            };
            cell.alignment = {
                vertical: 'middle',
                horizontal: LEFT_HEADERS[valueIndex]?.align || 'center',
                wrapText: valueIndex <= 6,
            };
            cell.border = thinBorder;
            if (isNumericCol(valueIndex + 1) && value !== '') {
                cell.numFmt = '#,##0.0';
            }
        });
        sheet.getRow(excelRow).height = 18;
    });

    if (rows.length) {
        const totalRow = headBottom + 1 + rows.length;
        const values = totalsValues(rows, years);
        values.forEach((value, valueIndex) => {
            const cell = sheet.getCell(totalRow, valueIndex + 1);
            cell.value = value === '' ? null : value;
            cell.fill = fillSolid(COLOR.gray50);
            cell.font = {name: 'Calibri', size: 10, bold: true, color: {argb: COLOR.text}};
            cell.alignment = {
                vertical: 'middle',
                horizontal: valueIndex === 0 ? 'right' : 'center',
            };
            cell.border = thinBorder;
            if (isNumericCol(valueIndex + 1) && value !== '') {
                cell.numFmt = '#,##0.0';
            }
        });
        if (LEFT_HEADERS.length > 1) {
            sheet.mergeCells(totalRow, 1, totalRow, 7);
            sheet.getCell(totalRow, 1).value = 'Total';
            sheet.getCell(totalRow, 1).alignment = {
                vertical: 'middle',
                horizontal: 'right',
                indent: 1,
            };
        }
        sheet.getRow(totalRow).height = 20;
    }
}

export async function downloadExcel({overRows, underRows, years, view}) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'STACK Land Book';
    workbook.created = new Date();
    workbook.views = [{x: 0, y: 0, width: 20000, height: 12000, firstSheet: 0, activeTab: 0, visibility: 'visible'}];

    addStyledSheet(workbook, {
        name: `Over ${MVA_THRESHOLD} MVA`,
        rows: overRows,
        years,
        view,
        tabArgb: COLOR.navyDark,
    });
    addStyledSheet(workbook, {
        name: `Under ${MVA_THRESHOLD} MVA`,
        rows: underRows,
        years,
        view,
        tabArgb: COLOR.navy,
    });

    const buffer = await workbook.xlsx.writeBuffer();
    downloadBlob(
        new Blob([buffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        }),
        `${exportBasename(view)}.xlsx`,
    );
}
