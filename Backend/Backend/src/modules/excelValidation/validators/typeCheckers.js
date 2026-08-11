'use strict';

/**
 * Excel Data-Type Validation — primitive type/format checkers.
 * -----------------------------------------------------------------
 * exceljs hands back cell values in several native JS shapes depending
 * on how the cell was authored (plain, formula, rich text, date), so
 * every checker first normalizes through `readCellValue` before
 * applying its type/format rule. Kept dependency-free and pure —
 * easy to unit test in isolation from parsing/reporting.
 * -----------------------------------------------------------------
 */

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

/**
 * Normalize a raw exceljs cell value into a plain JS primitive.
 * Handles: null/undefined, formula results ({result, formula}),
 * rich text ({richText: [...]}), hyperlinks ({text, hyperlink}),
 * and native Date objects.
 * @param {*} raw
 * @returns {string|number|boolean|Date|null}
 */
function readCellValue(raw) {
    if (raw === null || raw === undefined) return null;

    if (raw instanceof Date) return raw;

    if (typeof raw === 'object') {
        if (Array.isArray(raw.richText)) {
            return raw.richText.map((r) => r.text).join('');
        }
        if ('result' in raw) return readCellValue(raw.result);
        if ('text' in raw) return raw.text;
        return raw; // unrecognized object shape — let the caller's type check fail it
    }

    return raw;
}

function isEmpty(value) {
    return value === null || value === undefined || (typeof value === 'string' && value.trim() === '');
}

function checkString(value) {
    const v = readCellValue(value);
    if (isEmpty(v)) return { valid: true, empty: true };
    if (typeof v === 'number' || typeof v === 'boolean') {
        // Excel happily stores "123" as a number in a text-intended
        // column — accept it as a string via coercion rather than fail
        // a value that's perfectly renderable as text.
        return { valid: true, empty: false };
    }
    if (typeof v !== 'string') {
        return { valid: false, empty: false, reason: `Expected text, got ${typeOf(v)}` };
    }
    return { valid: true, empty: false };
}

function checkNumber(value) {
    const v = readCellValue(value);
    if (isEmpty(v)) return { valid: true, empty: true };
    const n = typeof v === 'number' ? v : Number(String(v).replace(/[,$%\s]/g, ''));
    if (typeof v === 'boolean' || Number.isNaN(n)) {
        return { valid: false, empty: false, reason: `Expected a number, got "${v}"` };
    }
    return { valid: true, empty: false };
}

function checkBoolean(value) {
    const v = readCellValue(value);
    if (isEmpty(v)) return { valid: true, empty: true };
    if (typeof v === 'boolean') return { valid: true, empty: false };
    if (typeof v === 'string' && ['true', 'false', 'yes', 'no', '1', '0'].includes(v.trim().toLowerCase())) {
        return { valid: true, empty: false };
    }
    return { valid: false, empty: false, reason: `Expected true/false, got "${v}"` };
}

function checkEmail(value) {
    const v = readCellValue(value);
    if (isEmpty(v)) return { valid: true, empty: true };
    if (typeof v !== 'string' || !EMAIL_RE.test(v.trim())) {
        return { valid: false, empty: false, reason: `"${v}" is not a valid email address` };
    }
    return { valid: true, empty: false };
}

function checkDate(value) {
    const v = readCellValue(value);
    if (isEmpty(v)) return { valid: true, empty: true };
    if (v instanceof Date) {
        return Number.isNaN(v.getTime())
            ? { valid: false, empty: false, reason: 'Invalid date' }
            : { valid: true, empty: false };
    }
    const parsed = new Date(v);
    if (Number.isNaN(parsed.getTime())) {
        return { valid: false, empty: false, reason: `"${v}" is not a valid date` };
    }
    return { valid: true, empty: false };
}

function checkEnum(value, enumValues = []) {
    const v = readCellValue(value);
    if (isEmpty(v)) return { valid: true, empty: true };
    const normalized = String(v).trim().toLowerCase();
    const match = enumValues.some((allowed) => String(allowed).trim().toLowerCase() === normalized);
    if (!match) {
        return { valid: false, empty: false, reason: `"${v}" is not one of: ${enumValues.join(', ')}` };
    }
    return { valid: true, empty: false };
}

function typeOf(v) {
    if (v === null) return 'null';
    return Array.isArray(v) ? 'array' : typeof v;
}

const CHECKERS = {
    string:  (value) => checkString(value),
    number:  (value) => checkNumber(value),
    boolean: (value) => checkBoolean(value),
    email:   (value) => checkEmail(value),
    date:    (value) => checkDate(value),
    enum:    (value, column) => checkEnum(value, column && column.enumValues)
};

/**
 * Validate a single cell value against a schema column definition.
 * @param {*} value - raw exceljs cell value
 * @param {{type:string, required:boolean, enumValues?:string[]}} column
 * @returns {{ valid: boolean, reason?: string }}
 */
function checkColumnValue(value, column) {
    const checker = CHECKERS[column.type];
    if (!checker) {
        return { valid: false, reason: `Unsupported schema type "${column.type}"` };
    }

    const result = checker(value, column);

    if (result.empty) {
        if (column.required) {
            return { valid: false, reason: 'Required field is empty' };
        }
        return { valid: true };
    }

    return result.valid ? { valid: true } : { valid: false, reason: result.reason };
}

module.exports = {
    readCellValue,
    isEmpty,
    checkColumnValue,
    checkString,
    checkNumber,
    checkBoolean,
    checkEmail,
    checkDate,
    checkEnum
};
