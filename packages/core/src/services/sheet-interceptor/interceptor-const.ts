import { ICellData } from '../../types/interfaces/i-cell-data';
import { ICommandInfo } from '../command/command.service';
import type { ISheetLocation } from './utils/interceptor';
import { createInterceptorKey } from './utils/interceptor';

type Rule = {
    match: (cellData: ICellData) => boolean;
    isContinue: (prev: any, cur: any) => boolean;
    applyFunctions?: Record<string, any>;
    type: string;
};
const CELL_CONTENT = createInterceptorKey<ICellData, ISheetLocation>('CELL_CONTENT');
const BEFORE_CELL_EDIT = createInterceptorKey<ICellData, ISheetLocation>('BEFORE_CELL_EDIT');
const AFTER_CELL_EDIT = createInterceptorKey<ICellData, ISheetLocation>('AFTER_CELL_EDIT');
const PERMISSION = createInterceptorKey<boolean, ICommandInfo>('PERMISSION');
const AUTO_FILL = createInterceptorKey<Rule[]>('AUTO_FILL');

export const INTERCEPTOR_POINT = {
    CELL_CONTENT,
    BEFORE_CELL_EDIT,
    AFTER_CELL_EDIT,
    PERMISSION,
    AUTO_FILL,
};
