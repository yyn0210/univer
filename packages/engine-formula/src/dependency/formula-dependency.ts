import type { IRange, IUnitRange, Nullable } from '@univerjs/core';
import { Disposable, LifecycleStages, ObjectMatrix, OnLifecycle } from '@univerjs/core';
import { Inject } from '@wendellhu/redi';

import { LexerTreeBuilder } from '../analysis/lexer';
import type { LexerNode } from '../analysis/lexer-node';
import { AstTreeBuilder } from '../analysis/parser';
import type { AstRootNode, FunctionNode, PrefixNode, SuffixNode } from '../ast-node';
import type { BaseAstNode } from '../ast-node/base-ast-node';
import { ErrorNode } from '../ast-node/base-ast-node';
import { NodeType } from '../ast-node/node-type';
import { FormulaAstLRU } from '../basics/cache-lru';
import type { IDirtyUnitSheetNameMap, IFormulaData, IOtherFormulaData } from '../basics/common';
import { ErrorType } from '../basics/error-type';
import type { PreCalculateNodeType } from '../basics/node-type';
import { prefixToken, suffixToken } from '../basics/token';
import { Interpreter } from '../interpreter/interpreter';
import type { BaseReferenceObject } from '../reference-object/base-reference-object';
import { IFormulaCurrentConfigService } from '../services/current-data.service';
import {
    IReferenceExecutorManagerService,
    type ReferenceExecutorManagerService,
} from '../services/reference-executor-manager.service';
import { IFormulaRuntimeService } from '../services/runtime.service';
import { FormulaDependencyTree } from './dependency-tree';

const FORMULA_CACHE_LRU_COUNT = 100000;

export const FormulaASTCache = new FormulaAstLRU<AstRootNode>(FORMULA_CACHE_LRU_COUNT);

@OnLifecycle(LifecycleStages.Rendered, FormulaDependencyGenerator)
export class FormulaDependencyGenerator extends Disposable {
    private _updateRangeFlattenCache = new Map<string, Map<string, IRange[]>>();

    private _dirtyUnitSheetNameMap: IDirtyUnitSheetNameMap = {};

    constructor(
        @IFormulaCurrentConfigService private readonly _currentConfigService: IFormulaCurrentConfigService,
        @IFormulaRuntimeService private readonly _runtimeService: IFormulaRuntimeService,
        @IReferenceExecutorManagerService
        private readonly _referenceExecutorManagerService: ReferenceExecutorManagerService,
        @Inject(Interpreter) private readonly _interpreter: Interpreter,
        @Inject(AstTreeBuilder) private readonly _astTreeBuilder: AstTreeBuilder,
        @Inject(LexerTreeBuilder) private readonly _lexerTreeBuilder: LexerTreeBuilder
    ) {
        super();
    }

    override dispose(): void {
        this._updateRangeFlattenCache.clear();
        FormulaASTCache.clear();
        this._dirtyUnitSheetNameMap = {};
    }

    async generate() {
        this._updateRangeFlatten();

        // const formulaInterpreter = Interpreter.create(interpreterDatasetConfig);

        const formulaData = this._currentConfigService.getFormulaData();

        const otherFormulaData = this._currentConfigService.getOtherFormulaData();

        const treeList = await this._generateTreeList(formulaData, otherFormulaData);

        const updateTreeList = this._getUpdateTreeListAndMakeDependency(treeList);

        const isCycleDependency = this._checkIsCycleDependency(updateTreeList);

        if (isCycleDependency) {
            this._runtimeService.enableCycleDependency();
        }

        const finalTreeList = this._calculateRunList(updateTreeList);

        return Promise.resolve(finalTreeList);
    }

    private _isCyclicUtil(
        node: FormulaDependencyTree,
        visited: Set<FormulaDependencyTree>,
        recursionStack: Set<FormulaDependencyTree>
    ) {
        if (!visited.has(node)) {
            // Mark the current node as visited and part of recursion stack
            visited.add(node);
            recursionStack.add(node);

            // Recur for all the children of this node
            for (let i = 0; i < node.children.length; i++) {
                if (!visited.has(node.children[i]) && this._isCyclicUtil(node.children[i], visited, recursionStack)) {
                    return true;
                }
                if (recursionStack.has(node.children[i])) {
                    return true;
                }
            }
        }
        recursionStack.delete(node); // remove the node from recursion stack
        return false;
    }

    private _checkIsCycleDependency(treeList: FormulaDependencyTree[]) {
        const visited = new Set<FormulaDependencyTree>();
        const recursionStack = new Set<FormulaDependencyTree>();

        // Call the recursive helper function to detect cycle in different
        // DFS trees
        for (let i = 0, len = treeList.length; i < len; i++) {
            const tree = treeList[i];
            const isCycle = this._isCyclicUtil(tree, visited, recursionStack);
            if (isCycle === true) {
                return true;
            }
        }

        return false;
    }

    /**
     * Generate nodes for the dependency tree, where each node contains all the reference data ranges included in each formula.
     * @param formulaData
     * @returns
     */
    private async _generateTreeList(formulaData: IFormulaData, otherFormulaData: IOtherFormulaData) {
        const formulaDataKeys = Object.keys(formulaData);

        const otherFormulaDataKeys = Object.keys(otherFormulaData);

        const treeList: FormulaDependencyTree[] = [];

        /**
         * Register formulas in the sheet.
         */
        for (const unitId of formulaDataKeys) {
            const sheetData = formulaData[unitId];

            const sheetDataKeys = Object.keys(sheetData);

            for (const sheetId of sheetDataKeys) {
                const matrixData = new ObjectMatrix(sheetData[sheetId]);

                matrixData.forValue((row, column, formulaDataItem) => {
                    // const formulaString = formulaDataItem.f;
                    const { f: formulaString, x, y } = formulaDataItem;
                    const node = this._generateAstNode(formulaString, x, y);

                    const FDtree = new FormulaDependencyTree();

                    FDtree.node = node;
                    FDtree.formula = formulaString;
                    FDtree.unitId = unitId;
                    FDtree.subComponentId = sheetId;
                    FDtree.row = row;
                    FDtree.column = column;

                    treeList.push(FDtree);
                });
            }
        }

        /**
         * Register formulas in doc, slide, and other types of applications.
         */
        for (const unitId of otherFormulaDataKeys) {
            const subComponentData = otherFormulaData[unitId];

            const subComponentKeys = Object.keys(subComponentData);

            for (const subComponentId of subComponentKeys) {
                const subFormulaData = subComponentData[subComponentId];

                const subFormulaDataKeys = Object.keys(subFormulaData);

                for (const subFormulaDataId of subFormulaDataKeys) {
                    const formulaDataItem = subFormulaData[subFormulaDataId];

                    const { f: formulaString } = formulaDataItem;

                    const node = this._generateAstNode(formulaString);

                    const FDtree = new FormulaDependencyTree();

                    FDtree.node = node;
                    FDtree.formula = formulaString;
                    FDtree.unitId = unitId;
                    FDtree.subComponentId = subComponentId;

                    FDtree.formulaId = subFormulaDataId;

                    treeList.push(FDtree);
                }
            }
        }

        /**
         * Register the external application relying on 'ref' into the formula system,
         * which can determine the execution timing of the external application
         * registration Executor based on the dependency relationship.
         */
        this._referenceExecutorManagerService.getReferenceExecutorMap().forEach((params, featureId) => {
            const { unitId, subComponentId, dependencyRanges, executor } = params;
            const FDtree = new FormulaDependencyTree();

            FDtree.unitId = unitId;
            FDtree.subComponentId = subComponentId;

            FDtree.executor = executor;

            FDtree.rangeList = dependencyRanges;

            treeList.push(FDtree);
        });

        for (let i = 0, len = treeList.length; i < len; i++) {
            const tree = treeList[i];

            this._runtimeService.setCurrent(tree.row, tree.column, tree.subComponentId, tree.unitId);

            if (tree.node == null) {
                continue;
            }

            const rangeList = await this._getRangeListByNode(tree.node);

            for (let r = 0, rLen = rangeList.length; r < rLen; r++) {
                tree.pushRangeList(rangeList[r]);
            }
        }

        return treeList;
    }

    /**
     * Break down the dirty areas into ranges for subsequent matching.
     * @returns
     */
    private _updateRangeFlatten() {
        const forceCalculate = this._currentConfigService.isForceCalculate();
        const dirtyRanges = this._currentConfigService.getDirtyRanges();
        if (forceCalculate) {
            return;
        }
        this._updateRangeFlattenCache.clear();
        for (let i = 0; i < dirtyRanges.length; i++) {
            const gridRange = dirtyRanges[i];
            const range = gridRange.range;
            const sheetId = gridRange.sheetId;
            const unitId = gridRange.unitId;

            this._addFlattenCache(unitId, sheetId, range);
        }

        this._dirtyUnitSheetNameMap = this._currentConfigService.getDirtyNameMap();
    }

    private _generateAstNode(formulaString: string, refOffsetX: number = 0, refOffsetY: number = 0) {
        let astNode: Nullable<AstRootNode> = FormulaASTCache.get(`${formulaString}##${refOffsetX}${refOffsetY}`);

        if (astNode) {
            return astNode;
        }

        const lexerNode = this._lexerTreeBuilder.treeBuilder(formulaString);

        if ((lexerNode as ErrorType) in ErrorType) {
            return ErrorNode.create(lexerNode as ErrorType);
        }

        // suffix Express, 1+(3*4=4)*5+1 convert to 134*4=5*1++

        astNode = this._astTreeBuilder.parse(lexerNode as LexerNode, refOffsetX, refOffsetY);

        if (astNode == null) {
            throw new Error('astNode is null');
        }

        FormulaASTCache.set(`${formulaString}##${refOffsetX}${refOffsetY}`, astNode);

        return astNode;
    }

    private _addFlattenCache(unitId: string, sheetId: string, range: IRange) {
        let unitMatrix = this._updateRangeFlattenCache.get(unitId);
        if (unitMatrix == null) {
            unitMatrix = new Map<string, IRange[]>();
            this._updateRangeFlattenCache.set(unitId, unitMatrix);
        }

        let ranges = unitMatrix.get(sheetId);

        if (ranges == null) {
            ranges = [];
            unitMatrix.set(sheetId, ranges);
        }

        ranges.push(range);

        // let sheetMatrix = unitMatrix.get(sheetId);
        // if (!sheetMatrix) {
        //     sheetMatrix = new ObjectMatrix<IRange>();
        //     unitMatrix.set(sheetId, sheetMatrix);
        // }

        // // don't use destructuring assignment
        // const startRow = range.startRow;

        // const startColumn = range.startColumn;

        // const endRow = range.endRow;

        // const endColumn = range.endColumn;

        // // don't use chained calls
        // for (let r = startRow; r <= endRow; r++) {
        //     for (let c = startColumn; c <= endColumn; c++) {
        //         sheetMatrix.setValue(r, c, true);
        //     }
        // }
    }

    private _isPreCalculateNode(node: BaseAstNode) {
        if (node.nodeType === NodeType.UNION) {
            return true;
        }

        if (node.nodeType === NodeType.PREFIX && (node as PrefixNode).getToken() === prefixToken.AT) {
            return true;
        }

        if (node.nodeType === NodeType.SUFFIX && (node as SuffixNode).getToken() === suffixToken.POUND) {
            return true;
        }

        return false;
    }

    private _nodeTraversalRef(node: BaseAstNode, result: PreCalculateNodeType[]) {
        const children = node.getChildren();
        const childrenCount = children.length;
        for (let i = 0; i < childrenCount; i++) {
            const item = children[i];
            if (this._isPreCalculateNode(item)) {
                result.push(item as PreCalculateNodeType);
                continue;
            } else if (item.nodeType === NodeType.REFERENCE) {
                result.push(item as PreCalculateNodeType);
            }
            this._nodeTraversalRef(item, result);
        }
    }

    private _nodeTraversalReferenceFunction(node: BaseAstNode, result: FunctionNode[]) {
        const children = node.getChildren();
        const childrenCount = children.length;
        for (let i = 0; i < childrenCount; i++) {
            const item = children[i];
            if (item.nodeType === NodeType.FUNCTION && (item as FunctionNode).isAddress()) {
                result.push(item as FunctionNode);
                continue;
            }
            this._nodeTraversalReferenceFunction(item, result);
        }
    }

    private async _executeNode(node: PreCalculateNodeType | FunctionNode) {
        let value: BaseReferenceObject;
        if (this._interpreter.checkAsyncNode(node)) {
            value = (await this._interpreter.executeAsync(node)) as BaseReferenceObject;
        } else {
            value = this._interpreter.execute(node) as BaseReferenceObject;
        }
        return value;
    }

    /**
     * Calculate the range required for collection in advance,
     * including references and location functions (such as OFFSET, INDIRECT, INDEX, etc.).
     * @param node
     * @returns
     */
    private async _getRangeListByNode(node: BaseAstNode) {
        // ref function in offset indirect INDEX
        const preCalculateNodeList: PreCalculateNodeType[] = [];
        const referenceFunctionList: FunctionNode[] = [];

        this._nodeTraversalRef(node, preCalculateNodeList);

        this._nodeTraversalReferenceFunction(node, referenceFunctionList);

        const rangeList: IUnitRange[] = [];

        for (let i = 0, len = preCalculateNodeList.length; i < len; i++) {
            const node = preCalculateNodeList[i];

            const value: BaseReferenceObject = await this._executeNode(node);

            const gridRange = value.toUnitRange();

            rangeList.push(gridRange);
        }

        for (let i = 0, len = referenceFunctionList.length; i < len; i++) {
            const node = referenceFunctionList[i];
            const value: BaseReferenceObject = await this._executeNode(node);

            const gridRange = value.toUnitRange();

            rangeList.push(gridRange);
        }

        return rangeList;
    }

    /**
     * Build a formula dependency tree based on the dependency relationships.
     * @param treeList
     * @returns
     */
    private _getUpdateTreeListAndMakeDependency(treeList: FormulaDependencyTree[]) {
        const newTreeList: FormulaDependencyTree[] = [];
        const existTree = new Set<FormulaDependencyTree>();
        const forceCalculate = this._currentConfigService.isForceCalculate();
        for (let i = 0, len = treeList.length; i < len; i++) {
            const tree = treeList[i];
            for (let m = 0, mLen = treeList.length; m < mLen; m++) {
                const treeMatch = treeList[m];
                if (tree === treeMatch) {
                    continue;
                }

                if (tree.dependency(treeMatch)) {
                    tree.pushChildren(treeMatch);
                }
            }

            /**
             * forceCalculate: Mandatory calculation, adding all formulas to dependencies
             * tree.dependencyRange: Formula dependent modification range
             * includeTree: modification range contains formula
             */
            if (
                (forceCalculate ||
                    tree.dependencyRange(
                        this._updateRangeFlattenCache,
                        this._dirtyUnitSheetNameMap,
                        this._currentConfigService.getExcludedRange()
                    ) ||
                    this._includeTree(tree)) &&
                !existTree.has(tree)
            ) {
                newTreeList.push(tree);
                existTree.add(tree);
            }
        }

        return newTreeList;
    }

    /**
     * Determine whether all ranges of the current node exist within the dirty area.
     * If they are within the dirty area, return true, indicating that this node needs to be calculated.
     * @param tree
     * @returns
     */
    private _includeTree(tree: FormulaDependencyTree) {
        const unitId = tree.unitId;
        const subComponentId = tree.subComponentId;

        const excludedCell = this._currentConfigService.getExcludedRange()?.[unitId]?.[subComponentId];

        /**
         * The position of the primary cell in the array formula needs to be excluded when calculating the impact of the array formula on dependencies.
         * This is because its impact was already considered during the first calculation.
         */
        let isExclude = false;
        excludedCell?.forValue((row, column) => {
            if (tree.row === row && tree.column === column) {
                isExclude = true;
                return false;
            }
        });

        if (isExclude) {
            return false;
        }

        /**
         * When the name of a worksheet is changed, or a worksheet is inserted or deleted,
         * the formulas within it may not need to be calculated.
         */
        // if (this._dirtyUnitSheetNameMap[unitId][subComponentId] != null) {
        //     return true;
        // }

        if (!this._updateRangeFlattenCache.has(unitId)) {
            return false;
        }

        const sheetRangeMap = this._updateRangeFlattenCache.get(unitId)!;

        if (!sheetRangeMap.has(subComponentId)) {
            return false;
        }

        const ranges = sheetRangeMap.get(subComponentId)!;

        for (const range of ranges) {
            if (tree.inRangeData(range)) {
                return true;
            }
        }

        return false;
    }

    /**
     * Generate the final formula calculation order array by traversing the dependency tree established via depth-first search.
     * @param treeList
     * @returns
     */
    private _calculateRunList(treeList: FormulaDependencyTree[]) {
        let stack = treeList;
        const formulaRunList = [];
        while (stack.length > 0) {
            const tree = stack.pop();

            if (tree === undefined || tree.isSkip()) {
                continue;
            }

            if (tree.isAdded()) {
                formulaRunList.push(tree);
                continue;
            }

            const cacheStack: FormulaDependencyTree[] = [];

            for (let i = 0, len = tree.parents.length; i < len; i++) {
                const parentTree = tree.parents[i];
                cacheStack.push(parentTree);
            }

            if (cacheStack.length === 0) {
                formulaRunList.push(tree);
                tree.setSkip();
            } else {
                tree.setAdded();
                stack.push(tree);
                stack = stack.concat(cacheStack);
            }
        }

        return formulaRunList.reverse();
    }
}
