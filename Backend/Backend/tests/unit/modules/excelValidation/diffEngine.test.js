const { diffRecordSets, valuesEqual } = require('../../../../src/modules/excelValidation/compare/diffEngine');

describe('diffEngine', () => {
    describe('valuesEqual', () => {
        it('treats numbers within tolerance as equal', () => {
            expect(valuesEqual(150.5, 150.504, 0.01)).toBe(true);
            expect(valuesEqual(150.5, 151, 0.01)).toBe(false);
        });

        it('is case-insensitive for strings', () => {
            expect(valuesEqual('Active', 'active')).toBe(true);
        });

        it('treats null/undefined/empty-string as equal to each other', () => {
            expect(valuesEqual(null, undefined)).toBe(true);
            expect(valuesEqual(undefined, '')).toBe(true);
        });
    });

    describe('diffRecordSets', () => {
        const left = [
            { id: '1', name: 'Acme', balance: 100 },
            { id: '2', name: 'Globex', balance: 200 },
            { id: '3', name: 'Only In Excel', balance: 50 }
        ];
        const right = [
            { id: '1', name: 'Acme', balance: 100 },      // matches
            { id: '2', name: 'Globex Inc', balance: 200 }, // name mismatch
            { id: '4', name: 'Only In API', balance: 75 }  // missing in left
        ];

        it('matches identical records', () => {
            const { matched } = diffRecordSets({ left, right, keyField: 'id' });
            expect(matched.map((m) => m.key)).toEqual(['1']);
        });

        it('reports field-level mismatches', () => {
            const { mismatched } = diffRecordSets({ left, right, keyField: 'id' });
            expect(mismatched).toHaveLength(1);
            expect(mismatched[0].key).toBe('2');
            expect(mismatched[0].differences).toEqual([
                { field: 'name', leftValue: 'Globex', rightValue: 'Globex Inc' }
            ]);
        });

        it('reports rows present on the left but not the right', () => {
            const { missingInRight } = diffRecordSets({ left, right, keyField: 'id' });
            expect(missingInRight.map((r) => r.key)).toEqual(['3']);
        });

        it('reports rows present on the right but not the left', () => {
            const { missingInLeft } = diffRecordSets({ left, right, keyField: 'id' });
            expect(missingInLeft.map((r) => r.key)).toEqual(['4']);
        });

        it('skips fields absent on the right side rather than flagging a mismatch', () => {
            const result = diffRecordSets({
                left: [{ id: '1', name: 'Acme', extraField: 'only on excel' }],
                right: [{ id: '1', name: 'Acme' }],
                keyField: 'id'
            });
            expect(result.mismatched).toEqual([]);
            expect(result.matched).toHaveLength(1);
        });

        it('throws without a keyField', () => {
            expect(() => diffRecordSets({ left: [], right: [] })).toThrow();
        });
    });
});
