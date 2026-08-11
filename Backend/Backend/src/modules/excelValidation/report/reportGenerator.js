'use strict';

const ExcelJS = require('exceljs');

/**
 * Excel PASS/FAIL Report
 * -----------------------------------------------------------------
 * Combines the output of schemaValidator + dataTypeValidator into one
 * automated report: total records, successful validations, failed
 * validations, and per-failure reasons — per sheet and overall.
 *
 * A "record" here is one data row. A row counts as PASS only if it has
 * zero data-type errors AND the sheet it lives on passed structural
 * (schema) validation — a row on a sheet that's missing a required
 * column can't be meaningfully judged, so every row on that sheet is
 * reported as FAIL with the schema problem as the reason.
 * -----------------------------------------------------------------
 */

function buildReport({ schemaResult, dataTypeResult, generatedAt = new Date().toISOString() } = {}) {
    const dtBySheet = new Map((dataTypeResult?.sheetResults || []).map((r) => [r.sheet, r]));
    const schemaBySheet = new Map((schemaResult?.sheetResults || []).map((r) => [r.sheet, r]));

    const bySheet = [];
    const failures = [];
    let totalRecords = 0;
    let totalPassed = 0;

    const sheetNames = new Set([...dtBySheet.keys(), ...schemaBySheet.keys()]);

    sheetNames.forEach((sheetName) => {
        const schemaSheet = schemaBySheet.get(sheetName);
        const dtSheet = dtBySheet.get(sheetName);
        const schemaOk = !schemaSheet || schemaSheet.ok;

        if (!schemaOk) {
            const reasons = [
                ...schemaSheet.missingColumns.map((c) => `Missing required column: "${c}"`),
                ...schemaSheet.unexpectedColumns.map((c) => `Unexpected column: "${c}"`)
            ];
            const rowCount = dtSheet ? dtSheet.totalRows : 0;

            failures.push({
                sheet: sheetName,
                row: null,
                field: null,
                reason: reasons.join('; ') || 'Sheet failed structural validation'
            });

            totalRecords += rowCount;
            bySheet.push({
                sheet: sheetName,
                schemaOk: false,
                totalRecords: rowCount,
                passed: 0,
                failed: rowCount,
                schemaIssues: reasons
            });
            return;
        }

        const totalRows = dtSheet ? dtSheet.totalRows : 0;
        const validRows = dtSheet ? dtSheet.validRows : 0;
        const invalidRows = dtSheet ? dtSheet.invalidRows : 0;

        if (dtSheet) {
            dtSheet.rowResults
                .filter((r) => !r.ok)
                .forEach((r) => {
                    r.errors.forEach((e) => {
                        failures.push({
                            sheet: sheetName,
                            row: r.rowNumber,
                            field: e.header,
                            reason: e.reason,
                            value: e.value
                        });
                    });
                });
        }

        totalRecords += totalRows;
        totalPassed += validRows;

        bySheet.push({
            sheet: sheetName,
            schemaOk: true,
            totalRecords: totalRows,
            passed: validRows,
            failed: invalidRows,
            schemaIssues: []
        });
    });

    const totalFailed = totalRecords - totalPassed;

    return {
        generatedAt,
        totalRecords,
        passed: totalPassed,
        failed: totalFailed,
        passRate: totalRecords === 0 ? 1 : Number((totalPassed / totalRecords).toFixed(4)),
        schemaOk: !!schemaResult && schemaResult.ok,
        missingSheets: schemaResult?.missingSheets || [],
        unexpectedSheets: schemaResult?.unexpectedSheets || [],
        bySheet,
        failures
    };
}

/**
 * Renders a PASS/FAIL report as a downloadable .xlsx workbook — a
 * "Summary" sheet with the overall/per-sheet counts, and a "Failures"
 * sheet listing every individual reason.
 * @param {ReturnType<typeof buildReport>} report
 * @returns {Promise<Buffer>}
 */
async function reportToWorkbookBuffer(report) {
    const wb = new ExcelJS.Workbook();

    const summary = wb.addWorksheet('Summary');
    summary.addRow(['Generated At', report.generatedAt]);
    summary.addRow(['Total Records', report.totalRecords]);
    summary.addRow(['Passed', report.passed]);
    summary.addRow(['Failed', report.failed]);
    summary.addRow(['Pass Rate', `${(report.passRate * 100).toFixed(2)}%`]);
    summary.addRow(['Schema OK', report.schemaOk ? 'YES' : 'NO']);
    summary.addRow(['Missing Sheets', report.missingSheets.join(', ') || '—']);
    summary.addRow([]);
    summary.addRow(['Sheet', 'Schema OK', 'Total Records', 'Passed', 'Failed']);
    report.bySheet.forEach((s) => {
        summary.addRow([s.sheet, s.schemaOk ? 'YES' : 'NO', s.totalRecords, s.passed, s.failed]);
    });

    const failuresSheet = wb.addWorksheet('Failures');
    failuresSheet.addRow(['Sheet', 'Row', 'Field', 'Value', 'Reason']);
    report.failures.forEach((f) => {
        failuresSheet.addRow([f.sheet, f.row ?? '—', f.field ?? '—', f.value ?? '', f.reason]);
    });

    return wb.xlsx.writeBuffer();
}

module.exports = { buildReport, reportToWorkbookBuffer };
