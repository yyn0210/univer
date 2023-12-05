import type { ICellData, ICommandInfo, IRange, IUnitRange, Nullable } from '@univerjs/core';
import {
    deserializeRangeWithSheet,
    Direction,
    Disposable,
    ICommandService,
    isFormulaString,
    IUniverInstanceService,
    LifecycleStages,
    ObjectMatrix,
    OnLifecycle,
    RANGE_TYPE,
    Rectangle,
    serializeRangeToRefString,
    Tools,
} from '@univerjs/core';
import type { IFormulaData, IFormulaDataItem, ISequenceNode, IUnitSheetNameMap } from '@univerjs/engine-formula';
import { FormulaEngineService, generateStringWithSequence, sequenceNodeType } from '@univerjs/engine-formula';
import type {
    IDeleteRangeMoveLeftCommandParams,
    IDeleteRangeMoveUpCommandParams,
    IInsertColCommandParams,
    IInsertRowCommandParams,
    IInsertSheetMutationParams,
    IMoveColsCommandParams,
    IMoveRangeCommandParams,
    IMoveRowsCommandParams,
    InsertRangeMoveDownCommandParams,
    InsertRangeMoveRightCommandParams,
    IRemoveRowColCommandParams,
    IRemoveSheetMutationParams,
    ISetRangeValuesMutationParams,
    ISetWorksheetNameCommandParams,
} from '@univerjs/sheets';
import {
    DeleteRangeMoveLeftCommand,
    DeleteRangeMoveUpCommand,
    EffectRefRangId,
    handleDeleteRangeMoveLeft,
    handleDeleteRangeMoveUp,
    handleInsertCol,
    handleInsertRangeMoveDown,
    handleInsertRangeMoveRight,
    handleInsertRow,
    handleIRemoveCol,
    handleIRemoveRow,
    handleMoveRange,
    InsertColCommand,
    InsertRangeMoveDownCommand,
    InsertRangeMoveRightCommand,
    InsertRowCommand,
    InsertSheetMutation,
    MoveColsCommand,
    MoveRangeCommand,
    MoveRowsCommand,
    RemoveColCommand,
    RemoveRowCommand,
    RemoveSheetMutation,
    runRefRangeMutations,
    SelectionManagerService,
    SetRangeValuesMutation,
    SetRangeValuesUndoMutationFactory,
    SetWorksheetNameCommand,
    SheetInterceptorService,
} from '@univerjs/sheets';
import { Inject, Injector } from '@wendellhu/redi';

import { SetArrayFormulaDataMutation } from '../commands/mutations/set-array-formula-data.mutation';
import { SetFormulaDataMutation } from '../commands/mutations/set-formula-data.mutation';
import { FormulaDataModel, initSheetFormulaData } from '../models/formula-data.model';
import { offsetArrayFormula, offsetFormula } from './utils';

interface IUnitRangeWithOffset extends IUnitRange {
    refOffsetX: number;
    refOffsetY: number;
    sheetName: string;
}

enum FormulaReferenceMoveType {
    Move, // range
    InsertRow, // row
    InsertColumn, // column
    RemoveRow, // row
    RemoveColumn, // column
    DeleteMoveLeft, // range
    DeleteMoveUp, // range
    InsertMoveDown, // range
    InsertMoveRight, // range
    SetName,
}

interface IFormulaReferenceMoveParam {
    type: FormulaReferenceMoveType;
    unitId: string;
    sheetId: string;
    ranges?: IRange[];
    from?: IRange;
    to?: IRange;
    sheetName?: string;
}

enum OriginRangeEdgeType {
    UP,
    DOWN,
    LEFT,
    RIGHT,
    ALL,
}

@OnLifecycle(LifecycleStages.Ready, UpdateFormulaController)
export class UpdateFormulaController extends Disposable {
    constructor(
        @IUniverInstanceService private readonly _currentUniverService: IUniverInstanceService,
        @ICommandService private readonly _commandService: ICommandService,
        @Inject(FormulaEngineService) private readonly _formulaEngineService: FormulaEngineService,
        @Inject(FormulaDataModel) private readonly _formulaDataModel: FormulaDataModel,
        @Inject(SheetInterceptorService) private _sheetInterceptorService: SheetInterceptorService,
        @Inject(SelectionManagerService) private _selectionManagerService: SelectionManagerService,
        @Inject(Injector) readonly _injector: Injector
    ) {
        super();

        this._initialize();
    }

    private _initialize(): void {
        this._commandExecutedListener();
    }

    private _commandExecutedListener() {
        this.disposeWithMe(
            this._sheetInterceptorService.interceptCommand({
                getMutations: (command) => this._getUpdateFormula(command),
            })
        );

        this.disposeWithMe(
            this._commandService.onCommandExecuted((command: ICommandInfo) => {
                const { id, params } = command;
                if (!params) return;

                switch (id) {
                    case SetRangeValuesMutation.id:
                        this._handleSetRangeValuesMutation(params as ISetRangeValuesMutationParams);
                        break;
                    case RemoveSheetMutation.id:
                        this._handleRemoveSheetMutation(params as IRemoveSheetMutationParams);
                        break;
                    case InsertSheetMutation.id:
                        this._handleInsertSheetMutation(params as IInsertSheetMutationParams);
                        break;

                    default:
                        break;
                }
            })
        );
    }

    private _handleSetRangeValuesMutation(params: ISetRangeValuesMutationParams) {
        const { worksheetId: sheetId, workbookId: unitId, cellValue, isFormulaUpdate } = params;

        if (isFormulaUpdate === true || cellValue == null) {
            return;
        }

        this._formulaDataModel.updateFormulaData(unitId, sheetId, cellValue);

        this._commandService.executeCommand(SetFormulaDataMutation.id, {
            formulaData: this._formulaDataModel.getFormulaData(),
        });
    }

    private _handleRemoveSheetMutation(params: IRemoveSheetMutationParams) {
        const { worksheetId: sheetId, workbookId: unitId } = params;

        const formulaData = this._formulaDataModel.getFormulaData();
        delete formulaData[unitId][sheetId];

        const arrayFormulaRange = this._formulaDataModel.getArrayFormulaRange();
        delete arrayFormulaRange[unitId][sheetId];

        const arrayFormulaCellData = this._formulaDataModel.getArrayFormulaCellData();
        delete arrayFormulaCellData[unitId][sheetId];

        this._commandService.executeCommand(SetFormulaDataMutation.id, {
            formulaData,
        });
        this._commandService.executeCommand(SetArrayFormulaDataMutation.id, {
            arrayFormulaRange,
            arrayFormulaCellData,
        });
    }

    private _handleInsertSheetMutation(params: IInsertSheetMutationParams) {
        const { sheet, workbookId: unitId } = params;

        const formulaData = this._formulaDataModel.getFormulaData();
        const { id: sheetId, cellData } = sheet;
        const cellMatrix = new ObjectMatrix(cellData);
        initSheetFormulaData(formulaData, unitId, sheetId, cellMatrix);

        this._commandService.executeCommand(SetFormulaDataMutation.id, {
            formulaData,
        });
    }

    private _getUpdateFormula(command: ICommandInfo) {
        const { id } = command;
        let result: Nullable<IFormulaReferenceMoveParam> = null;

        switch (id) {
            case MoveRangeCommand.id:
                result = this._handleMoveRange(command as ICommandInfo<IMoveRangeCommandParams>);
                break;
            case MoveRowsCommand.id:
                result = this._handleMoveRows(command as ICommandInfo<IMoveRowsCommandParams>);
                break;
            case MoveColsCommand.id:
                result = this._handleMoveCols(command as ICommandInfo<IMoveColsCommandParams>);
                break;
            case InsertRowCommand.id:
                result = this._handleInsertRow(command as ICommandInfo<IInsertRowCommandParams>);
                break;
            case InsertColCommand.id:
                result = this._handleInsertCol(command as ICommandInfo<IInsertColCommandParams>);
                break;
            case InsertRangeMoveRightCommand.id:
                result = this._handleInsertRangeMoveRight(command as ICommandInfo<InsertRangeMoveRightCommandParams>);
                break;
            case InsertRangeMoveDownCommand.id:
                result = this._handleInsertRangeMoveDown(command as ICommandInfo<InsertRangeMoveDownCommandParams>);
                break;
            case RemoveRowCommand.id:
                result = this._handleRemoveRow(command as ICommandInfo<IRemoveRowColCommandParams>);
                break;
            case RemoveColCommand.id:
                result = this._handleRemoveCol(command as ICommandInfo<IRemoveRowColCommandParams>);
                break;
            case DeleteRangeMoveUpCommand.id:
                result = this._handleDeleteRangeMoveUp(command as ICommandInfo<IDeleteRangeMoveUpCommandParams>);
                break;
            case DeleteRangeMoveLeftCommand.id:
                result = this._handleDeleteRangeMoveLeft(command as ICommandInfo<IDeleteRangeMoveLeftCommandParams>);
                break;
            case SetWorksheetNameCommand.id:
                result = this._handleSetWorksheetName(command as ICommandInfo<ISetWorksheetNameCommandParams>);
                break;
        }

        if (result) {
            const { unitSheetNameMap } = this._formulaDataModel.getCalculateData();
            let oldFormulaData = this._formulaDataModel.getFormulaData();

            // change formula reference
            const formulaData = this._getFormulaReferenceMoveInfo(oldFormulaData, unitSheetNameMap, result);

            const workbook = this._currentUniverService.getCurrentUniverSheetInstance();
            const unitId = workbook.getUnitId();
            const sheetId = workbook.getActiveSheet().getSheetId();
            const selections = this._selectionManagerService.getSelections();

            // offset arrayFormula
            const arrayFormulaRange = this._formulaDataModel.getArrayFormulaRange();
            const arrayFormulaCellData = this._formulaDataModel.getArrayFormulaCellData();

            let offsetArrayFormulaRange = offsetFormula(arrayFormulaRange, command, unitId, sheetId, selections);
            offsetArrayFormulaRange = offsetArrayFormula(offsetArrayFormulaRange, unitId, sheetId);
            const offsetArrayFormulaCellData = offsetFormula(
                arrayFormulaCellData,
                command,
                unitId,
                sheetId,
                selections
            );

            // Synchronous to the worker thread
            this._commandService.executeCommand(SetArrayFormulaDataMutation.id, {
                arrayFormulaRange: offsetArrayFormulaRange,
                arrayFormulaCellData: offsetArrayFormulaCellData,
            });

            // offset formulaData
            oldFormulaData = offsetFormula(oldFormulaData, command, unitId, sheetId, selections);
            const offsetFormulaData = offsetFormula(formulaData, command, unitId, sheetId, selections);

            return this._getUpdateFormulaMutations(oldFormulaData, offsetFormulaData);
        }

        return {
            undos: [],
            redos: [],
        };
    }

    private _handleMoveRange(command: ICommandInfo<IMoveRangeCommandParams>) {
        const { params } = command;
        if (!params) return null;

        const { fromRange, toRange } = params;
        if (!fromRange || !toRange) return null;

        const { unitId, sheetId } = this._getCurrentSheetInfo();

        return {
            type: FormulaReferenceMoveType.Move,
            from: fromRange,
            to: toRange,
            unitId,
            sheetId,
        };
    }

    private _handleMoveRows(command: ICommandInfo<IMoveRowsCommandParams>) {
        const { params } = command;
        if (!params) return null;

        const { fromRow, toRow } = params;

        const workbook = this._currentUniverService.getCurrentUniverSheetInstance();
        const unitId = workbook.getUnitId();
        const worksheet = workbook.getActiveSheet();
        const sheetId = worksheet.getSheetId();

        const from = {
            startRow: fromRow,
            startColumn: 0,
            endRow: fromRow,
            endColumn: worksheet.getColumnCount() - 1,
            rangeType: RANGE_TYPE.ROW,
        };
        const to = {
            startRow: toRow,
            startColumn: 0,
            endRow: toRow,
            endColumn: worksheet.getColumnCount() - 1,
            rangeType: RANGE_TYPE.ROW,
        };

        return {
            type: FormulaReferenceMoveType.Move,
            from,
            to,
            unitId,
            sheetId,
        };
    }

    private _handleMoveCols(command: ICommandInfo<IMoveColsCommandParams>) {
        const { params } = command;
        if (!params) return null;

        const { fromCol, toCol } = params;

        const workbook = this._currentUniverService.getCurrentUniverSheetInstance();
        const unitId = workbook.getUnitId();
        const worksheet = workbook.getActiveSheet();
        const sheetId = worksheet.getSheetId();

        const from = {
            startRow: 0,
            startColumn: fromCol,
            endRow: worksheet.getRowCount() - 1,
            endColumn: fromCol,
            rangeType: RANGE_TYPE.COLUMN,
        };
        const to = {
            startRow: 0,
            startColumn: toCol,
            endRow: worksheet.getRowCount() - 1,
            endColumn: toCol,
            rangeType: RANGE_TYPE.COLUMN,
        };

        return {
            type: FormulaReferenceMoveType.Move,
            from,
            to,
            unitId,
            sheetId,
        };
    }

    private _handleInsertRow(command: ICommandInfo<IInsertRowCommandParams>) {
        const { params } = command;
        if (!params) return null;

        const { range, workbookId, worksheetId } = params;
        return {
            type: FormulaReferenceMoveType.InsertRow,
            ranges: [range],
            unitId: workbookId,
            sheetId: worksheetId,
        };
    }

    private _handleInsertCol(command: ICommandInfo<IInsertColCommandParams>) {
        const { params } = command;
        if (!params) return null;

        const { range, workbookId, worksheetId } = params;
        return {
            type: FormulaReferenceMoveType.InsertColumn,
            ranges: [range],
            unitId: workbookId,
            sheetId: worksheetId,
        };
    }

    private _handleInsertRangeMoveRight(command: ICommandInfo<InsertRangeMoveRightCommandParams>) {
        const { params } = command;
        if (!params) return null;

        const { ranges } = params;
        const { unitId, sheetId } = this._getCurrentSheetInfo();

        return {
            type: FormulaReferenceMoveType.InsertMoveRight,
            ranges,
            unitId,
            sheetId,
        };
    }

    private _handleInsertRangeMoveDown(command: ICommandInfo<InsertRangeMoveDownCommandParams>) {
        const { params } = command;
        if (!params) return null;

        const { ranges } = params;
        const { unitId, sheetId } = this._getCurrentSheetInfo();

        return {
            type: FormulaReferenceMoveType.InsertMoveDown,
            ranges,
            unitId,
            sheetId,
        };
    }

    private _handleRemoveRow(command: ICommandInfo<IRemoveRowColCommandParams>) {
        const { params } = command;
        if (!params) return null;

        const { ranges } = params;
        const { unitId, sheetId } = this._getCurrentSheetInfo();

        return {
            type: FormulaReferenceMoveType.RemoveRow,
            ranges,
            unitId,
            sheetId,
        };
    }

    private _handleRemoveCol(command: ICommandInfo<IRemoveRowColCommandParams>) {
        const { params } = command;
        if (!params) return null;

        const { ranges } = params;
        const { unitId, sheetId } = this._getCurrentSheetInfo();

        return {
            type: FormulaReferenceMoveType.RemoveColumn,
            ranges,
            unitId,
            sheetId,
        };
    }

    private _handleDeleteRangeMoveUp(command: ICommandInfo<IDeleteRangeMoveUpCommandParams>) {
        const { params } = command;
        if (!params) return null;

        const { ranges } = params;
        const { unitId, sheetId } = this._getCurrentSheetInfo();

        return {
            type: FormulaReferenceMoveType.DeleteMoveUp,
            ranges,
            unitId,
            sheetId,
        };
    }

    private _handleDeleteRangeMoveLeft(command: ICommandInfo<IDeleteRangeMoveLeftCommandParams>) {
        const { params } = command;
        if (!params) return null;

        const { ranges } = params;
        const { unitId, sheetId } = this._getCurrentSheetInfo();

        return {
            type: FormulaReferenceMoveType.DeleteMoveLeft,
            ranges,
            unitId,
            sheetId,
        };
    }

    private _handleSetWorksheetName(command: ICommandInfo<ISetWorksheetNameCommandParams>) {
        const { params } = command;
        if (!params) return null;

        const { workbookId, worksheetId, name } = params;

        const { unitId, sheetId } = this._getCurrentSheetInfo();

        return {
            type: FormulaReferenceMoveType.SetName,
            unitId: workbookId || unitId,
            sheetId: worksheetId || sheetId,
            sheetName: name,
        };
    }

    private _getUpdateFormulaMutations(oldFormulaData: IFormulaData, formulaData: IFormulaData) {
        const redos = [];
        const undos = [];
        const accessor = {
            get: this._injector.get.bind(this._injector),
        };

        const formulaDataKeys = Object.keys(formulaData);

        for (const workbookId of formulaDataKeys) {
            const sheetData = formulaData[workbookId];
            const sheetDataKeys = Object.keys(sheetData);

            for (const worksheetId of sheetDataKeys) {
                const oldFormulaMatrix = new ObjectMatrix<IFormulaDataItem>(oldFormulaData[workbookId][worksheetId]);
                const formulaMatrix = new ObjectMatrix<IFormulaDataItem>(sheetData[worksheetId]);
                const cellMatrix = new ObjectMatrix<ICellData>();

                formulaMatrix.forValue((r, c, formulaItem) => {
                    const formulaString = formulaItem?.f || '';
                    const oldFormulaString = oldFormulaMatrix.getRow(r)?.get(c)?.f || '';

                    if (isFormulaString(formulaString)) {
                        // formula with formula id
                        if (isFormulaString(oldFormulaString) && formulaString !== oldFormulaString) {
                            cellMatrix.setValue(r, c, { f: formulaString });
                        } else {
                            // formula with only id
                            cellMatrix.setValue(r, c, { f: formulaString, si: null });
                        }
                    }
                });

                const cellValue = cellMatrix.getData();
                if (Tools.isEmptyObject(cellValue)) continue;

                const setRangeValuesMutationParams: ISetRangeValuesMutationParams = {
                    worksheetId,
                    workbookId,
                    cellValue,
                };

                redos.push({
                    id: SetRangeValuesMutation.id,
                    params: setRangeValuesMutationParams,
                });

                const undoSetRangeValuesMutationParams: ISetRangeValuesMutationParams =
                    SetRangeValuesUndoMutationFactory(accessor, setRangeValuesMutationParams);

                undos.push({
                    id: SetRangeValuesMutation.id,
                    params: undoSetRangeValuesMutationParams,
                });
            }
        }

        return {
            redos,
            undos,
        };
    }

    private _getFormulaReferenceMoveInfo(
        formulaData: IFormulaData,
        unitSheetNameMap: IUnitSheetNameMap,
        formulaReferenceMoveParam: IFormulaReferenceMoveParam
    ) {
        const formulaDataKeys = Object.keys(formulaData);

        const newFormulaData: IFormulaData = {};

        for (const unitId of formulaDataKeys) {
            const sheetData = formulaData[unitId];

            const sheetDataKeys = Object.keys(sheetData);

            if (newFormulaData[unitId] == null) {
                newFormulaData[unitId] = {};
            }

            for (const sheetId of sheetDataKeys) {
                const matrixData = new ObjectMatrix(sheetData[sheetId]);

                const newFormulaDataItem = new ObjectMatrix<IFormulaDataItem>();

                matrixData.forValue((row, column, formulaDataItem) => {
                    const { f: formulaString, x, y, si } = formulaDataItem;

                    const sequenceNodes = this._formulaEngineService.buildSequenceNodes(formulaString);

                    if (sequenceNodes == null) {
                        return true;
                    }

                    let shouldModify = false;
                    const refChangeIds: number[] = [];
                    for (let i = 0, len = sequenceNodes.length; i < len; i++) {
                        const node = sequenceNodes[i];
                        if (typeof node === 'string' || node.nodeType !== sequenceNodeType.REFERENCE) {
                            continue;
                        }
                        const { token } = node;

                        const sequenceGrid = deserializeRangeWithSheet(token);

                        const { range, sheetName, unitId: sequenceUnitId } = sequenceGrid;

                        const mapUnitId =
                            sequenceUnitId == null || sequenceUnitId.length === 0 ? unitId : sequenceUnitId;

                        const sequenceSheetId = unitSheetNameMap?.[mapUnitId]?.[sheetName];

                        const sequenceUnitRangeWidthOffset = {
                            range,
                            sheetId: sequenceSheetId,
                            unitId: sequenceUnitId,
                            sheetName,
                            refOffsetX: x || 0,
                            refOffsetY: y || 0,
                        };

                        let newRefString: Nullable<string> = null;

                        if (formulaReferenceMoveParam.type === FormulaReferenceMoveType.SetName) {
                            const {
                                unitId: userUnitId,
                                sheetId: userSheetId,
                                sheetName: newSheetName,
                            } = formulaReferenceMoveParam;
                            if (newSheetName == null) {
                                continue;
                            }

                            if (sequenceSheetId == null || sequenceSheetId.length === 0) {
                                continue;
                            }

                            if (userSheetId !== sequenceSheetId) {
                                continue;
                            }

                            newRefString = serializeRangeToRefString({
                                range,
                                sheetName: newSheetName,
                                unitId: sequenceUnitId,
                            });
                        } else {
                            newRefString = this._getNewRangeByMoveParam(
                                sequenceUnitRangeWidthOffset,
                                formulaReferenceMoveParam,
                                unitId,
                                sheetId
                            );
                        }

                        if (newRefString != null) {
                            sequenceNodes[i] = {
                                ...node,
                                token: newRefString,
                            };
                            shouldModify = true;
                            refChangeIds.push(i);
                        }
                    }

                    if (!shouldModify) {
                        return true;
                    }

                    const newSequenceNodes = this._updateRefOffset(sequenceNodes, refChangeIds, x, y);

                    newFormulaDataItem.setValue(row, column, {
                        f: `=${generateStringWithSequence(newSequenceNodes)}`,
                        x,
                        y,
                        si,
                    });
                });

                newFormulaData[unitId][sheetId] = newFormulaDataItem.getData();
            }
        }

        return newFormulaData;
    }

    private _getNewRangeByMoveParam(
        unitRangeWidthOffset: IUnitRangeWithOffset,
        formulaReferenceMoveParam: IFormulaReferenceMoveParam,
        currentFormulaUnitId: string,
        currentFormulaSheetId: string
    ) {
        const { type, unitId: userUnitId, sheetId: userSheetId, ranges, from, to } = formulaReferenceMoveParam;

        const {
            range,
            sheetId: sequenceRangeSheetId,
            unitId: sequenceRangeUnitId,
            sheetName: sequenceRangeSheetName,
            refOffsetX,
            refOffsetY,
        } = unitRangeWidthOffset;

        if (
            !this._checkIsSameUnitAndSheet(
                userUnitId,
                userSheetId,
                currentFormulaUnitId,
                currentFormulaSheetId,
                sequenceRangeUnitId,
                sequenceRangeSheetId
            )
        ) {
            return;
        }

        const sequenceRange = Rectangle.moveOffset(range, refOffsetX, refOffsetY);
        let newRange: Nullable<IRange> = null;

        if (type === FormulaReferenceMoveType.Move) {
            if (from == null || to == null) {
                return;
            }

            const moveEdge = this._checkMoveEdge(sequenceRange, from);

            // const fromAndToDirection = this._checkMoveFromAndToDirection(from, to);

            // if (moveEdge == null) {
            //     return;
            // }

            const remainRange = Rectangle.getIntersects(sequenceRange, from);

            if (remainRange == null) {
                return;
            }

            const operators = handleMoveRange(
                { id: EffectRefRangId.MoveRangeCommandId, params: { toRange: to, fromRange: from } },
                remainRange
            );

            const result = runRefRangeMutations(operators, remainRange);

            if (result == null) {
                return;
            }

            newRange = this._getMoveNewRange(moveEdge, result, from, to, sequenceRange, remainRange);
        }

        if (ranges != null) {
            if (type === FormulaReferenceMoveType.InsertRow) {
                const operators = handleInsertRow(
                    {
                        id: EffectRefRangId.InsertRowCommandId,
                        params: { range: ranges[0], workbookId: '', worksheetId: '', direction: Direction.DOWN },
                    },
                    sequenceRange
                );

                const result = runRefRangeMutations(operators, sequenceRange);

                if (result == null) {
                    return;
                }

                newRange = {
                    ...sequenceRange,
                    ...result,
                };
            } else if (type === FormulaReferenceMoveType.InsertColumn) {
                const operators = handleInsertCol(
                    {
                        id: EffectRefRangId.InsertColCommandId,
                        params: { range: ranges[0], workbookId: '', worksheetId: '', direction: Direction.RIGHT },
                    },
                    sequenceRange
                );

                const result = runRefRangeMutations(operators, sequenceRange);

                if (result == null) {
                    return;
                }

                newRange = {
                    ...sequenceRange,
                    ...result,
                };
            } else if (type === FormulaReferenceMoveType.RemoveRow) {
                const operators = handleIRemoveRow(
                    {
                        id: EffectRefRangId.RemoveRowCommandId,
                        params: { ranges },
                    },
                    sequenceRange
                );

                const result = runRefRangeMutations(operators, sequenceRange);

                if (result == null) {
                    return;
                }

                newRange = {
                    ...sequenceRange,
                    ...result,
                };
            } else if (type === FormulaReferenceMoveType.RemoveColumn) {
                const operators = handleIRemoveCol(
                    {
                        id: EffectRefRangId.RemoveColCommandId,
                        params: { ranges },
                    },
                    sequenceRange
                );

                const result = runRefRangeMutations(operators, sequenceRange);

                if (result == null) {
                    return;
                }

                newRange = {
                    ...sequenceRange,
                    ...result,
                };
            } else if (type === FormulaReferenceMoveType.DeleteMoveLeft) {
                const operators = handleDeleteRangeMoveLeft(
                    {
                        id: EffectRefRangId.DeleteRangeMoveLeftCommandId,
                        params: { ranges },
                    },
                    sequenceRange
                );

                const result = runRefRangeMutations(operators, sequenceRange);

                if (result == null) {
                    return;
                }

                newRange = {
                    ...sequenceRange,
                    ...result,
                };
            } else if (type === FormulaReferenceMoveType.DeleteMoveUp) {
                const operators = handleDeleteRangeMoveUp(
                    {
                        id: EffectRefRangId.DeleteRangeMoveUpCommandId,
                        params: { ranges },
                    },
                    sequenceRange
                );

                const result = runRefRangeMutations(operators, sequenceRange);

                if (result == null) {
                    return;
                }

                newRange = {
                    ...sequenceRange,
                    ...result,
                };
            } else if (type === FormulaReferenceMoveType.InsertMoveDown) {
                const operators = handleInsertRangeMoveDown(
                    {
                        id: EffectRefRangId.InsertRangeMoveDownCommandId,
                        params: { ranges },
                    },
                    sequenceRange
                );

                const result = runRefRangeMutations(operators, sequenceRange);

                if (result == null) {
                    return;
                }

                newRange = {
                    ...sequenceRange,
                    ...result,
                };
            } else if (type === FormulaReferenceMoveType.InsertMoveRight) {
                const operators = handleInsertRangeMoveRight(
                    {
                        id: EffectRefRangId.InsertRangeMoveRightCommandId,
                        params: { ranges },
                    },
                    sequenceRange
                );

                const result = runRefRangeMutations(operators, sequenceRange);

                if (result == null) {
                    return;
                }

                newRange = {
                    ...sequenceRange,
                    ...result,
                };
            }
        }

        if (newRange == null) {
            return;
        }

        return serializeRangeToRefString({
            range: newRange,
            sheetName: sequenceRangeSheetName,
            unitId: sequenceRangeUnitId,
        });
    }

    private _checkIsSameUnitAndSheet(
        userUnitId: string,
        userSheetId: string,
        currentFormulaUnitId: string,
        currentFormulaSheetId: string,
        sequenceRangeUnitId: string,
        sequenceRangeSheetId: string
    ) {
        if (
            (sequenceRangeUnitId == null || sequenceRangeUnitId.length === 0) &&
            (sequenceRangeSheetId == null || sequenceRangeSheetId.length === 0)
        ) {
            if (userUnitId === currentFormulaUnitId && userSheetId === currentFormulaSheetId) {
                return true;
            }
        } else if (userUnitId === sequenceRangeUnitId && userSheetId === sequenceRangeSheetId) {
            return true;
        }

        return false;
    }

    /**
     * Update all ref nodes to the latest offset state.
     */
    private _updateRefOffset(
        sequenceNodes: Array<string | ISequenceNode>,
        refChangeIds: number[],
        refOffsetX: number = 0,
        refOffsetY: number = 0
    ) {
        const newSequenceNodes: Array<string | ISequenceNode> = [];
        for (let i = 0, len = sequenceNodes.length; i < len; i++) {
            const node = sequenceNodes[i];
            if (typeof node === 'string' || node.nodeType !== sequenceNodeType.REFERENCE || refChangeIds.includes(i)) {
                newSequenceNodes.push(node);
                continue;
            }

            const { token } = node;

            const sequenceGrid = deserializeRangeWithSheet(token);

            const { range, sheetName, unitId: sequenceUnitId } = sequenceGrid;

            const newRange = Rectangle.moveOffset(range, refOffsetX, refOffsetY);

            newSequenceNodes.push({
                ...node,
                token: serializeRangeToRefString({
                    range: newRange,
                    unitId: sequenceUnitId,
                    sheetName,
                }),
            });
        }

        return newSequenceNodes;
    }

    /**
     * Determine the range of the moving selection,
     * and check if it is at the edge of the reference range of the formula.
     * @param originRange
     * @param fromRange
     */
    private _checkMoveEdge(originRange: IRange, fromRange: IRange): Nullable<OriginRangeEdgeType> {
        const { startRow, startColumn, endRow, endColumn } = originRange;

        const {
            startRow: fromStartRow,
            startColumn: fromStartColumn,
            endRow: fromEndRow,
            endColumn: fromEndColumn,
        } = fromRange;

        if (
            startRow >= fromStartRow &&
            endRow <= fromEndRow &&
            startColumn >= fromStartColumn &&
            endColumn <= fromEndColumn
        ) {
            return OriginRangeEdgeType.ALL;
        }

        if (
            startColumn >= fromStartColumn &&
            endColumn <= fromEndColumn &&
            startRow >= fromStartRow &&
            startRow <= fromEndRow &&
            endRow > fromEndRow
        ) {
            return OriginRangeEdgeType.UP;
        }

        if (
            startColumn >= fromStartColumn &&
            endColumn <= fromEndColumn &&
            endRow >= fromStartRow &&
            endRow <= fromEndRow &&
            startRow < fromStartRow
        ) {
            return OriginRangeEdgeType.DOWN;
        }

        if (
            startRow >= fromStartRow &&
            endRow <= fromEndRow &&
            startColumn >= fromStartColumn &&
            startColumn <= fromEndColumn &&
            endColumn > fromEndColumn
        ) {
            return OriginRangeEdgeType.LEFT;
        }

        if (
            startRow >= fromStartRow &&
            endRow <= fromEndRow &&
            endColumn >= fromStartColumn &&
            endColumn <= fromEndColumn &&
            startColumn < fromStartColumn
        ) {
            return OriginRangeEdgeType.RIGHT;
        }
    }

    /**
     *  Calculate the new ref information for the moving selection.
     * @param moveEdge  the 'from' range lie on the edge of the original range, or does it completely cover the original range
     * @param result The original range is divided by 'from' and moved to a new position range.
     * @param from The initial range of the moving selection.
     * @param to The result range after moving the initial range.
     * @param origin The original target range.
     * @param remain "The range subtracted from the initial range by 'from'.
     * @returns
     */
    private _getMoveNewRange(
        moveEdge: Nullable<OriginRangeEdgeType>,
        result: IRange,
        from: IRange,
        to: IRange,
        origin: IRange,
        remain: IRange
    ) {
        const { startRow, endRow, startColumn, endColumn } = result;

        const {
            startRow: fromStartRow,
            startColumn: fromStartColumn,
            endRow: fromEndRow,
            endColumn: fromEndColumn,
        } = from;

        const { startRow: toStartRow, startColumn: toStartColumn, endRow: toEndRow, endColumn: toEndColumn } = to;

        const {
            startRow: remainStartRow,
            endRow: remainEndRow,
            startColumn: remainStartColumn,
            endColumn: remainEndColumn,
        } = remain;

        const {
            startRow: originStartRow,
            endRow: originEndRow,
            startColumn: originStartColumn,
            endColumn: originEndColumn,
        } = origin;

        const newRange = { ...origin };

        if (moveEdge === OriginRangeEdgeType.UP) {
            if (startColumn === originStartColumn && endColumn === originEndColumn) {
                if (startRow < originStartRow) {
                    newRange.startRow = startRow;
                } else if (endRow < originEndRow) {
                    newRange.startRow = startRow;
                } else if (endRow >= originEndRow && toStartRow <= originEndRow) {
                    newRange.startRow = fromEndRow + 1;
                    newRange.endRow = endRow;
                } else {
                    return;
                }
            } else {
                return;
            }
        } else if (moveEdge === OriginRangeEdgeType.DOWN) {
            if (startColumn === originStartColumn && endColumn === originEndColumn) {
                if (endRow > originEndRow) {
                    newRange.endRow = endRow;
                } else if (startRow > originStartRow) {
                    newRange.endRow = endRow;
                } else if (startRow <= originStartRow && toEndRow >= originStartRow) {
                    newRange.endRow = fromStartRow + 1;
                    newRange.startRow = startRow;
                } else {
                    return;
                }
            } else {
                return;
            }
        } else if (moveEdge === OriginRangeEdgeType.LEFT) {
            if (startRow === originStartRow && endRow === originEndRow) {
                if (startColumn < originStartColumn) {
                    newRange.startColumn = startColumn;
                } else if (endColumn < originEndColumn) {
                    newRange.startColumn = startColumn;
                } else if (endColumn >= originEndColumn && toStartColumn <= originEndColumn) {
                    newRange.startColumn = fromEndColumn + 1;
                    newRange.endColumn = endColumn;
                } else {
                    return;
                }
            } else {
                return;
            }
        } else if (moveEdge === OriginRangeEdgeType.RIGHT) {
            if (startRow === originStartRow && endRow === originEndRow) {
                if (endColumn > originEndColumn) {
                    newRange.endColumn = endColumn;
                } else if (startColumn > originStartColumn) {
                    newRange.endColumn = endColumn;
                } else if (startColumn <= originStartColumn && toEndColumn >= originStartColumn) {
                    newRange.endColumn = fromStartColumn + 1;
                    newRange.startColumn = startColumn;
                } else {
                    return;
                }
            } else {
                return;
            }
        } else if (moveEdge === OriginRangeEdgeType.ALL) {
            newRange.startRow = startRow;
            newRange.startColumn = startColumn;
            newRange.endRow = endRow;
            newRange.endColumn = endColumn;
        } else if (
            ((toStartColumn <= remainEndColumn + 1 && toEndColumn >= originEndColumn) ||
                (toStartColumn <= originStartColumn && toEndColumn >= remainStartColumn - 1)) &&
            toStartRow <= originStartRow &&
            toEndRow >= originEndRow
        ) {
            newRange.startRow = startRow;
            newRange.startColumn = startColumn;
            newRange.endRow = endRow;
            newRange.endColumn = endColumn;
        } else if (
            ((toStartRow <= remainEndRow + 1 && toEndRow >= originEndRow) ||
                (toStartRow <= originStartRow && toEndRow >= remainStartRow - 1)) &&
            toStartColumn <= originStartColumn &&
            toEndColumn >= originEndColumn
        ) {
            newRange.startRow = startRow;
            newRange.startColumn = startColumn;
            newRange.endRow = endRow;
            newRange.endColumn = endColumn;
        }

        return newRange;
    }

    private _getInsertNewRange() {}

    private _getCurrentSheetInfo() {
        const workbook = this._currentUniverService.getCurrentUniverSheetInstance();
        const unitId = workbook.getUnitId();
        const sheetId = workbook.getActiveSheet().getSheetId();

        return {
            unitId,
            sheetId,
        };
    }
}
