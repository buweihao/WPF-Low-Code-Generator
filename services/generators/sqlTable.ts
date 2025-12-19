import { SheetData, PointDefinition } from '../../types';
import { getCSharpType } from '../utils';

export const generateSQLTable = (sheets: SheetData[]): string => {
  const sb: string[] = [];
  sb.push(`using System;`);
  sb.push(`using SqlSugar;`);
  sb.push(`using MyDatabase.Table;`);  
  sb.push(``);
  sb.push(`namespace Core`);
  sb.push(`{`);

  const processedClasses = new Set<string>();

  sheets.forEach(sheet => {
    // Group points by Period
    const pointsByPeriod = new Map<number, PointDefinition[]>();
    sheet.points.forEach(p => {
        if (p.Period === 0) return;
        const key = p.Period;
        if (!pointsByPeriod.has(key)) pointsByPeriod.set(key, []);
        pointsByPeriod.get(key)!.push(p);
    });

    pointsByPeriod.forEach((groupPoints, period) => {
        // Sanitize period for class name (e.g. 0.1 -> 0_1)
        const safePeriod = Math.abs(period).toString().replace('.', '_');
        const className = `${sheet.name}_PeriodAbs_${safePeriod}`;
        
        if (processedClasses.has(className)) return; 
        processedClasses.add(className);

        sb.push(`    public class ${className} : TableTemplate`);
        sb.push(`    {`);
        groupPoints.forEach(p => {
             const csType = getCSharpType(p.Type);
             const safeKeyName = p.KeyName ? p.KeyName.replace(/"/g, '\\"') : "";
             
             let attrParams = `ColumnDescription = "${safeKeyName}"`;
             
             // Check if the raw Type string indicates an array (ends with []) to enable JSON storage
             if (p.Type && p.Type.trim().endsWith('[]')) {
                 attrParams += `, IsJson = true, ColumnDataType = "longtext"`;
             }

             sb.push(`        [SugarColumn(${attrParams})]`);
             sb.push(`        public ${csType} ${p.PropertyName} { get; set; }`);
        });
        sb.push(`    }`);
        sb.push(``);
    });
  });

  sb.push(`}`);
  return sb.join('\n');
};