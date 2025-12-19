import { SheetData, DeviceConfig } from '../../types';

export const generateGlobalAuto = (sheets: SheetData[], maxModules: number, devices: DeviceConfig[]): string => {
  const sb: string[] = [];
  sb.push(`using System;`);
  sb.push(`using System.Threading.Tasks;`);
  sb.push(`using System.Collections.Generic;`);
  sb.push(`using System.Text;`);
  sb.push(`using MyDatabase;`);
  sb.push(``);
  sb.push(`namespace Core`);
  sb.push(`{`);
  sb.push(`    public static partial class Global`);
  sb.push(`    {`);
  
  // 1. Static Modbus Fields
  sb.push(`        #region PLC Definitions`);
  for(let m=1; m<=maxModules; m++) {
      sheets.forEach(s => {
          sb.push(`        public static IModbusService _${m}${s.name}PLCModbus;`);
      });
  }
  sb.push(`        #endregion`);
  sb.push(``);

  // 2. InitPLC Method
  sb.push(`        public static void InitPLC(int maxModules)`);
  sb.push(`        {`);
  sb.push(`            manager.CloseAll();`);
  
  for(let m=1; m<=maxModules; m++) {
      sb.push(`            if (maxModules >= ${m})`);
      sb.push(`            {`);
      
      sheets.forEach(s => {
          // Look up device based on convention {SheetName}_M{Index}
          const targetDeviceName = `${s.name}_M${m}`;
          const device = devices.find(d => d.DeviceName === targetDeviceName);
          const ip = device ? device.IP : "127.0.0.1";
          const port = device ? device.Port : 502;
          const serviceName = `_${m}${s.name}`; // e.g., _1UnLoadModuleA

          sb.push(`                manager.AddTcp("${serviceName}", "${ip}", ${port});`);
          sb.push(`                ${serviceName}PLCModbus = Global.manager.GetService("${serviceName}");`);
      });
      
      sb.push(`            }`);
  }

  sb.push(`        }`);
  sb.push(``);

  // 3. Database Repositories
  const tableNames = new Set<string>();
  sheets.forEach(sheet => {
      const pointsByPeriod = new Map<number, any>();
      sheet.points.forEach(p => { if(p.Period !== 0) pointsByPeriod.set(p.Period, true); });
      pointsByPeriod.forEach((_, period) => {
          // Sanitize period for class name matches sqlTable.ts
          const safePeriod = Math.abs(period).toString().replace('.', '_');
          tableNames.add(`${sheet.name}_PeriodAbs_${safePeriod}`);
      });
  });

  sb.push(`        #region Database Init`);
  Array.from(tableNames).forEach(t => {
      sb.push(`        public static IRepository<${t}> repo_${t};`);
  });
  sb.push(``);
  sb.push(`        public static void SQLInit_Auto()`);
  sb.push(`        {`);
  sb.push(`            databaseHelper = new DatabaseHelper(GetValue("ConnectionStrings"), new[]`);
  sb.push(`            {`);
  Array.from(tableNames).forEach(t => {
      sb.push(`                typeof(${t}),`);
  });
  sb.push(`            });`);
  sb.push(``);
  Array.from(tableNames).forEach(t => {
      sb.push(`            repo_${t} = databaseHelper.GetRepo<${t}>();`);
  });
  sb.push(`        }`);
  sb.push(`        #endregion`);
  
  sb.push(`    }`);
  sb.push(`}`);

  return sb.join('\n');
};