import type { IMutationInfo, IRange } from '@univerjs/core';
import { Disposable, IUniverInstanceService, Range, Rectangle } from '@univerjs/core';
import type { ISetNumfmtMutationParams } from '@univerjs/sheets';
import { factorySetNumfmtUndoMutation, INumfmtService, SetNumfmtMutation } from '@univerjs/sheets';
import type { IAutoFillHook } from '@univerjs/sheets-ui';
import { APPLY_TYPE, getRepeatRange, IAutoFillService } from '@univerjs/sheets-ui';
import { Inject, Injector } from '@wendellhu/redi';

import { SHEET_NUMFMT_PLUGIN } from '../base/const/PLUGIN_NAME';

export class NumfmtAutoFillController extends Disposable {
    constructor(
        @Inject(Injector) private _injector: Injector,
        @Inject(IUniverInstanceService) private _univerInstanceService: IUniverInstanceService,
        @Inject(INumfmtService) private _numfmtService: INumfmtService,
        @Inject(IAutoFillService) private _autoFillService: IAutoFillService
    ) {
        super();

        this._initAutoFill();
    }

    private _initAutoFill() {
        const noopReturnFunc = () => ({ redos: [], undos: [] });
        const loopFunc = (
            sourceStartCell: { row: number; col: number },
            targetStartCell: { row: number; col: number },
            relativeRange: IRange
        ) => {
            const workbookId = this._univerInstanceService.getCurrentUniverSheetInstance().getUnitId();
            const worksheetId = this._univerInstanceService
                .getCurrentUniverSheetInstance()
                .getActiveSheet()
                .getSheetId();
            const sourceRange = {
                startRow: sourceStartCell.row,
                startColumn: sourceStartCell.col,
                endColumn: sourceStartCell.col,
                endRow: sourceStartCell.row,
            };
            const targetRange = {
                startRow: targetStartCell.row,
                startColumn: targetStartCell.col,
                endColumn: targetStartCell.col,
                endRow: targetStartCell.row,
            };

            const values: ISetNumfmtMutationParams['values'] = [];

            Range.foreach(relativeRange, (row, col) => {
                const sourcePositionRange = Rectangle.getPositionRange(
                    {
                        startRow: row,
                        startColumn: col,
                        endColumn: col,
                        endRow: row,
                    },
                    sourceRange
                );
                const oldNumfmtValue = this._numfmtService.getValue(
                    workbookId,
                    worksheetId,
                    sourcePositionRange.startRow,
                    sourcePositionRange.startColumn
                );
                if (oldNumfmtValue) {
                    const targetPositionRange = Rectangle.getPositionRange(
                        {
                            startRow: row,
                            startColumn: col,
                            endColumn: col,
                            endRow: row,
                        },
                        targetRange
                    );
                    values.push({
                        pattern: oldNumfmtValue.pattern,
                        type: oldNumfmtValue.type,
                        row: targetPositionRange.startRow,
                        col: targetPositionRange.startColumn,
                    });
                }
            });
            if (values.length) {
                const redo: IMutationInfo<ISetNumfmtMutationParams> = {
                    id: SetNumfmtMutation.id,
                    params: {
                        values,
                        workbookId,
                        worksheetId,
                    },
                };
                const undo = {
                    id: SetNumfmtMutation.id,
                    params: {
                        values: factorySetNumfmtUndoMutation(this._injector, redo.params),
                        workbookId,
                        worksheetId,
                    },
                };
                return {
                    redos: [redo],
                    undos: [undo],
                };
            }
            return { redos: [], undos: [] };
        };
        const generalApplyFunc = (sourceRange: IRange, targetRange: IRange) => {
            const totalUndos: IMutationInfo[] = [];
            const totalRedos: IMutationInfo[] = [];
            const sourceStartCell = {
                row: sourceRange.startRow,
                col: sourceRange.startColumn,
            };
            const repeats = getRepeatRange(sourceRange, targetRange);
            repeats.forEach((repeat) => {
                const { undos, redos } = loopFunc(sourceStartCell, repeat.repeatStartCell, repeat.relativeRange);
                totalUndos.push(...undos);
                totalRedos.push(...redos);
            });
            return {
                undos: totalUndos,
                redos: totalRedos,
            };
        };
        const hook: IAutoFillHook = {
            hookName: SHEET_NUMFMT_PLUGIN,
            hook: {
                [APPLY_TYPE.COPY]: generalApplyFunc,
                [APPLY_TYPE.NO_FORMAT]: noopReturnFunc,
                [APPLY_TYPE.ONLY_FORMAT]: generalApplyFunc,
                [APPLY_TYPE.SERIES]: generalApplyFunc,
            },
        };
        this.disposeWithMe(this._autoFillService.addHook(hook));
    }
}
