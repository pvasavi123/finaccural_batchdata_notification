'use strict';

/**
 * Generic record-set diff engine.
 * -----------------------------------------------------------------
 * Shared by apiComparator.js (Excel vs API Validation) and
 * dbComparator.js (Excel vs Database Reconciliation) — both problems
 * are "take two arrays of plain objects, match rows by a key field,
 * report per-field mismatches plus rows only on one side."
 *
 * Kept free of any Excel/exceljs/Sequelize/axios knowledge so it's
 * trivial to unit test and to reuse for a third comparison target
 * later without touching this file.
 * -----------------------------------------------------------------
 */

function normalizeKey(v) {
    return v === undefined || v === null ? '' : String(v).trim().toLowerCase();
}

function normalizeForCompare(v) {
    if (v === undefined || v === null) return '';
    if (v instanceof Date) return v.toISOString();
    return v;
}

/**
 * @param {*} a
 * @param {*} b
 * @param {number} tolerance - absolute tolerance used when BOTH sides parse as numbers
 * @returns {boolean}
 */
function valuesEqual(a, b, tolerance = 0.01) {
    const na = normalizeForCompare(a);
    const nb = normalizeForCompare(b);

    if (na === nb) return true;

    const numA = Number(na);
    const numB = Number(nb);
    if (na !== '' && nb !== '' && !Number.isNaN(numA) && !Number.isNaN(numB)) {
        return Math.abs(numA - numB) <= tolerance;
    }

    return String(na).trim().toLowerCase() === String(nb).trim().toLowerCase();
}

/**
 * Diff two record sets keyed by `keyField`.
 * @param {object} opts
 * @param {Array<object>} opts.left - e.g. rows parsed from Excel
 * @param {Array<object>} opts.right - e.g. rows fetched from a live API or the DB
 * @param {string} opts.keyField - field name present on both sides used to match records
 * @param {string[]} [opts.fields] - fields to compare; defaults to every key on the left record that also exists on the right record
 * @param {number} [opts.tolerance] - numeric tolerance, see valuesEqual
 * @returns {{
 *   totalLeft:number, totalRight:number,
 *   matchedCount:number, mismatchedCount:number,
 *   missingInRightCount:number, missingInLeftCount:number,
 *   matched:Array, mismatched:Array, missingInRight:Array, missingInLeft:Array
 * }}
 */
function diffRecordSets({ left = [], right = [], keyField, fields, tolerance = 0.01 }) {
    if (!keyField) {
        throw new Error('diffRecordSets requires a keyField.');
    }

    const rightByKey = new Map();
    right.forEach((r) => {
        const k = normalizeKey(r[keyField]);
        if (k !== '') rightByKey.set(k, r);
    });

    const matched = [];
    const mismatched = [];
    const missingInRight = [];
    const seenRightKeys = new Set();

    left.forEach((l) => {
        const k = normalizeKey(l[keyField]);
        const r = k === '' ? undefined : rightByKey.get(k);

        if (!r) {
            missingInRight.push({ key: l[keyField], record: l });
            return;
        }

        seenRightKeys.add(k);

        const compareFields = fields && fields.length
            ? fields
            : Object.keys(l).filter((f) => f !== keyField);

        const differences = [];
        compareFields.forEach((field) => {
            if (!(field in r)) return; // field not present on the right side — not comparable, not a mismatch
            if (!valuesEqual(l[field], r[field], tolerance)) {
                differences.push({ field, leftValue: l[field] ?? null, rightValue: r[field] ?? null });
            }
        });

        if (differences.length) {
            mismatched.push({ key: l[keyField], differences });
        } else {
            matched.push({ key: l[keyField] });
        }
    });

    const missingInLeft = right
        .filter((r) => !seenRightKeys.has(normalizeKey(r[keyField])))
        .map((r) => ({ key: r[keyField], record: r }));

    return {
        totalLeft: left.length,
        totalRight: right.length,
        matchedCount: matched.length,
        mismatchedCount: mismatched.length,
        missingInRightCount: missingInRight.length,
        missingInLeftCount: missingInLeft.length,
        matched,
        mismatched,
        missingInRight,
        missingInLeft
    };
}

module.exports = { diffRecordSets, valuesEqual, normalizeKey };
