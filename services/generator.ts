import { generatePLCPointProperty as genProp } from './generators/plcPointProperty';
import { generatePLCPointXaml as genXaml } from './generators/plcPointXaml';
import { generateSQLTable as genSql } from './generators/sqlTable';
import { generateGlobalAuto as genGlobal } from './generators/globalAuto';
import { SheetData, DeviceConfig, ByteOrder, StringByteOrder } from '../types';

export const generatePLCPointProperty = (
    sheets: SheetData[], 
    maxModules: number, 
    byteOrder: ByteOrder, 
    stringByteOrder: StringByteOrder,
    maxGap: number, 
    maxBatchSize: number
) => genProp(sheets, maxModules, byteOrder, stringByteOrder, maxGap, maxBatchSize);

export const generatePLCPointXaml = genXaml;
export const generateSQLTable = genSql;
export const generateGlobalAuto = genGlobal;