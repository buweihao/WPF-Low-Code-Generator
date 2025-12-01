import { SheetData, PointDefinition } from '../types';

// Helper to clean address (remove D prefix)
const cleanAddr = (addr: string | number | undefined): string => {
  if (addr === undefined || addr === null || addr === '') return "0";
  // Force conversion to string to handle numbers from Excel, then trim and uppercase
  return String(addr).trim().toUpperCase().replace('D', '');
};

// Helper to get C# type
const getCSharpType = (type: string): string => {
  switch (type.toLowerCase()) {
    case 'bool': return 'bool';
    case 'short': return 'short';
    case 'int': return 'int';
    case 'float': return 'float';
    case 'string': return 'string';
    default: return 'string';
  }
};

// Helper to map type to Modbus method
const getReadMethod = (type: string, async: boolean = false): string => {
  const suffix = async ? 'Async' : '';
  switch (type.toLowerCase()) {
    case 'bool': return `ReadCoils${suffix}`;
    case 'short': return `ReadRegisters${suffix}`;
    case 'int': return `ReadDInt${suffix}`;
    case 'string': return `ReadString${suffix}`;
    default: return `ReadRegisters${suffix}`; // Default fallback
  }
};

const getWriteMethod = (type: string, async: boolean = false): string => {
   const suffix = async ? 'Async' : '';
   switch (type.toLowerCase()) {
     case 'bool': return `WriteCoil${suffix}`;
     case 'string': return `WriteString${suffix}`;
     default: return `WriteRegisters${suffix}`;
   }
};

// --- Generator Functions ---

export const generatePLCPointProperty = (sheets: SheetData[], maxModules: number): string => {
  const sb: string[] = [];

  sb.push(`using System;`);
  sb.push(`using System.Collections.Generic;`);
  sb.push(`using System.ComponentModel;`);
  sb.push(`using System.Threading;`);
  sb.push(`using System.Threading.Tasks;`);
  sb.push(`using System.Diagnostics;`);
  sb.push(`using System.Text;`);
  sb.push(``);
  sb.push(`namespace Core`);
  sb.push(`{`);
  sb.push(`    public class PLCPointProperty : INotifyPropertyChanged`);
  sb.push(`    {`);
  sb.push(`        public static PLCPointProperty Instance { get; } = new PLCPointProperty();`);
  sb.push(`        public event PropertyChangedEventHandler PropertyChanged;`);
  sb.push(`        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));`);
  sb.push(``);
  sb.push(`        private int _frequency = 1000;`);
  sb.push(`        public int Frequency`);
  sb.push(`        {`);
  sb.push(`            get => _frequency;`);
  sb.push(`            set { _frequency = value; OnPropertyChanged(nameof(Frequency)); StartAllTasks(); }`);
  sb.push(`        }`);
  sb.push(``);
  sb.push(`#region 属性`);
  // Generate Properties
  for (let m = 1; m <= maxModules; m++) {
    sheets.forEach(sheet => {
      sheet.points.forEach(p => {
        const propName = `${p.PropertyName}_M${m}`;
        const csType = getCSharpType(p.Type);
        sb.push(`        private ${csType} _${propName};`);
        sb.push(`        public ${csType} ${propName}`);
        sb.push(`        {`);
        sb.push(`            get => _${propName};`);
        sb.push(`            set { _${propName} = value; OnPropertyChanged(nameof(${propName})); }`);
        sb.push(`        }`);
      });
    });
  }
  sb.push(`#endregion`);
  sb.push(``);
  sb.push(`        private Dictionary<string, CancellationTokenSource> _tasks = new Dictionary<string, CancellationTokenSource>();`);
  sb.push(``);
  sb.push(`        public void StartAllTasks()`);
  sb.push(`        {`);
  sb.push(`            // Cancel existing tasks`);
  sb.push(`            foreach (var kvp in _tasks) kvp.Value.Cancel();`);
  sb.push(`            _tasks.Clear();`);
  sb.push(``);
  sb.push(`            if (Frequency <= 0) return;`);
  sb.push(``);

  // 1. Generate Monitor Tasks (Independent task per property for UI updates)
  for (let m = 1; m <= maxModules; m++) {
      sheets.forEach(sheet => {
          sheet.points.forEach(p => {
              const propName = `${p.PropertyName}_M${m}`;
              const taskName = `Monitor_${propName}`;
              sb.push(`            StartTask("${taskName}", (token) => MonitorTask_${propName}(token));`);
          });
      });
  }
  sb.push(``);

  // 2. Generate Logging/Handshake Tasks (Based on Period)
  for (let m = 1; m <= maxModules; m++) {
    sheets.forEach(sheet => {
      // Group by Period
      const pointsByPeriod = new Map<number, PointDefinition[]>();
      sheet.points.forEach(p => {
          if (p.Period === 0) return;
          const key = p.Period;
          if (!pointsByPeriod.has(key)) pointsByPeriod.set(key, []);
          pointsByPeriod.get(key)!.push(p);
      });

      pointsByPeriod.forEach((groupPoints, period) => {
          if (period < 0) {
              const groupByTrigger = new Map<string, PointDefinition[]>();
              groupPoints.forEach(p => {
                  const t = cleanAddr(p.TriggerAddress);
                  if(!groupByTrigger.has(t)) groupByTrigger.set(t, []);
                  groupByTrigger.get(t)!.push(p);
              });

              groupByTrigger.forEach((_, triggerAddr) => {
                  const taskName = `${sheet.name}_M${m}_Handshake_T${triggerAddr}`;
                  sb.push(`            StartTask("${taskName}", (token) => RunHandshakeTask_${m}_${sheet.name}_T${triggerAddr}(token));`);
              });
          } else {
               const taskName = `${sheet.name}_M${m}_Period_${period}`;
               sb.push(`            StartTask("${taskName}", (token) => RunPeriodicTask_${m}_${sheet.name}_P${period}(token));`);
          }
      });
    });
  }

  sb.push(`        }`);
  sb.push(``);
  sb.push(`        private void StartTask(string name, Func<CancellationToken, Task> action)`);
  sb.push(`        {`);
  sb.push(`            var cts = new CancellationTokenSource();`);
  sb.push(`            _tasks[name] = cts;`);
  sb.push(`            Task.Run(() => action(cts.Token), cts.Token);`);
  sb.push(`        }`);
  sb.push(``);

  // --- Generate Task Bodies ---
  
  // A. Monitor Tasks Definitions
  for (let m = 1; m <= maxModules; m++) {
      sheets.forEach(sheet => {
          const plcName = `Global._${m}${sheet.name}PLCModbus`;
          sheet.points.forEach(p => {
             const propName = `${p.PropertyName}_M${m}`;
             const addr = cleanAddr(p.ValueAddress);
             const method = getReadMethod(p.Type, false);
             const type = p.Type.toLowerCase();

             sb.push(`        private async Task MonitorTask_${propName}(CancellationToken token)`);
             sb.push(`        {`);
             sb.push(`            while (!token.IsCancellationRequested)`);
             sb.push(`            {`);
             sb.push(`                try`);
             sb.push(`                {`);
             sb.push(`                    await Task.Delay(Frequency, token);`);
             sb.push(`                    if (${plcName} == null || !${plcName}.IsConnected) continue;`);
             sb.push(``);
             sb.push(`                    string val = ${plcName}.${method}("${addr}", ${p.Length});`);
             if (type === 'string') {
                  sb.push(`                    this.${propName} = val;`);
             } else if (type === 'bool') {
                  sb.push(`                    this.${propName} = bool.TryParse(val, out var b) ? b : false;`);
             } else if (type === 'float') {
                  sb.push(`                    this.${propName} = float.TryParse(val, out var f) ? f : 0f;`);
             } else if (type === 'int') {
                  sb.push(`                    this.${propName} = int.TryParse(val, out var i) ? i : 0;`);
             } else {
                  sb.push(`                    this.${propName} = short.TryParse(val, out var s) ? s : (short)0;`);
             }
             sb.push(`                }`);
             sb.push(`                catch {}`);
             sb.push(`            }`);
             sb.push(`        }`);
          });
      });
  }
  sb.push(``);

  // B. Logging/Handshake Tasks Definitions
  for (let m = 1; m <= maxModules; m++) {
      sheets.forEach(sheet => {
          const plcName = `Global._${m}${sheet.name}PLCModbus`;
          
          const pointsByPeriod = new Map<number, PointDefinition[]>();
          sheet.points.forEach(p => {
              if (p.Period === 0) return;
              const key = p.Period;
              if (!pointsByPeriod.has(key)) pointsByPeriod.set(key, []);
              pointsByPeriod.get(key)!.push(p);
          });

          pointsByPeriod.forEach((groupPoints, period) => {
              const tableName = `${sheet.name}_PeriodAbs_${Math.abs(period)}`;

              if (period > 0) {
                  // --- Periodic Task ---
                  sb.push(`        private async Task RunPeriodicTask_${m}_${sheet.name}_P${period}(CancellationToken token)`);
                  sb.push(`        {`);
                  sb.push(`            while (!token.IsCancellationRequested)`);
                  sb.push(`            {`);
                  sb.push(`                try`);
                  sb.push(`                {`);
                  sb.push(`                    await Task.Delay(${period}, token);`); 
                  sb.push(``);
                  sb.push(`                    if (${plcName} == null || !${plcName}.IsConnected) continue;`);
                  sb.push(``);
                  sb.push(`                    var entity = new ${tableName}();`);
                  sb.push(`                    entity.ModuleNum = ${m};`);
                  sb.push(``);
                  groupPoints.forEach(p => {
                      const addr = cleanAddr(p.ValueAddress);
                      const method = getReadMethod(p.Type, false);
                      sb.push(`                    string val_${p.PropertyName} = ${plcName}.${method}("${addr}", ${p.Length});`);
                      
                      const type = p.Type.toLowerCase();
                      if (type === 'string') {
                           sb.push(`                    entity.${p.PropertyName} = val_${p.PropertyName};`);
                      } else if (type === 'bool') {
                           sb.push(`                    entity.${p.PropertyName} = bool.TryParse(val_${p.PropertyName}, out var b_${p.PropertyName}) ? b_${p.PropertyName} : false;`);
                      } else if (type === 'float') {
                           sb.push(`                    entity.${p.PropertyName} = float.TryParse(val_${p.PropertyName}, out var f_${p.PropertyName}) ? f_${p.PropertyName} : 0f;`);
                      } else if (type === 'int') {
                           sb.push(`                    entity.${p.PropertyName} = int.TryParse(val_${p.PropertyName}, out var i_${p.PropertyName}) ? i_${p.PropertyName} : 0;`);
                      } else {
                           sb.push(`                    entity.${p.PropertyName} = short.TryParse(val_${p.PropertyName}, out var s_${p.PropertyName}) ? s_${p.PropertyName} : (short)0;`);
                      }
                  });
                  sb.push(``);
                  sb.push(`                    await Global.repo_${tableName}.InsertAsync(entity);`);
                  sb.push(`                }`);
                  sb.push(`                catch (Exception ex) { Console.WriteLine(ex.Message); }`);
                  sb.push(`            }`);
                  sb.push(`        }`);
              } else {
                  // --- Handshake Task ---
                  const groupByTrigger = new Map<string, PointDefinition[]>();
                  groupPoints.forEach(p => {
                      const t = cleanAddr(p.TriggerAddress);
                      if(!groupByTrigger.has(t)) groupByTrigger.set(t, []);
                      groupByTrigger.get(t)!.push(p);
                  });

                  groupByTrigger.forEach((tPoints, triggerAddr) => {
                      const returnAddr = cleanAddr(tPoints[0].ReturnAddress);
                      
                      sb.push(`        private async Task RunHandshakeTask_${m}_${sheet.name}_T${triggerAddr}(CancellationToken token)`);
                      sb.push(`        {`);
                      sb.push(`            while (!token.IsCancellationRequested)`);
                      sb.push(`            {`);
                      sb.push(`                try`);
                      sb.push(`                {`);
                      sb.push(`                     await Task.Delay(${Math.abs(period)}, token);`);
                      sb.push(`                     if (${plcName} == null || !${plcName}.IsConnected) continue;`);
                      sb.push(``);
                      sb.push(`                     string triggerValStr = ${plcName}.ReadRegisters("${triggerAddr}", 1);`);
                      sb.push(`                     if (!int.TryParse(triggerValStr, out int triggerVal) || triggerVal != 11) continue;`);
                      sb.push(``);
                      sb.push(`                    var entity = new ${tableName}();`);
                      sb.push(`                    entity.ModuleNum = ${m};`);
                      tPoints.forEach(p => {
                          const addr = cleanAddr(p.ValueAddress);
                          const method = getReadMethod(p.Type, false);
                          sb.push(`                    string val_${p.PropertyName} = ${plcName}.${method}("${addr}", ${p.Length});`);
                          const type = p.Type.toLowerCase();
                          if (type === 'string') {
                               sb.push(`                    entity.${p.PropertyName} = val_${p.PropertyName};`);
                          } else if (type === 'bool') {
                               sb.push(`                    entity.${p.PropertyName} = bool.TryParse(val_${p.PropertyName}, out var b_${p.PropertyName}) ? b_${p.PropertyName} : false;`);
                          } else if (type === 'float') {
                               sb.push(`                    entity.${p.PropertyName} = float.TryParse(val_${p.PropertyName}, out var f_${p.PropertyName}) ? f_${p.PropertyName} : 0f;`);
                          } else if (type === 'int') {
                               sb.push(`                    entity.${p.PropertyName} = int.TryParse(val_${p.PropertyName}, out var i_${p.PropertyName}) ? i_${p.PropertyName} : 0;`);
                          } else {
                               sb.push(`                    entity.${p.PropertyName} = short.TryParse(val_${p.PropertyName}, out var s_${p.PropertyName}) ? s_${p.PropertyName} : (short)0;`);
                          }
                      });
                      sb.push(``);
                      sb.push(`                     await Global.repo_${tableName}.InsertAsync(entity);`);
                      sb.push(``);
                      sb.push(`                     ${plcName}.WriteRegisters("${returnAddr}", 11);`);
                      sb.push(``);
                      sb.push(`                     Stopwatch sw = Stopwatch.StartNew();`);
                      sb.push(`                     while (sw.ElapsedMilliseconds < 5000)`);
                      sb.push(`                     {`);
                      sb.push(`                         string rVal = ${plcName}.ReadRegisters("${returnAddr}", 1);`);
                      sb.push(`                         if (int.TryParse(rVal, out int rInt) && rInt == 0) break;`);
                      sb.push(`                         await Task.Delay(100, token);`);
                      sb.push(`                     }`);
                      sb.push(``);
                      sb.push(`                     sw.Restart();`);
                      sb.push(`                     while (sw.ElapsedMilliseconds < 5000)`);
                      sb.push(`                     {`);
                      sb.push(`                         string tVal = ${plcName}.ReadRegisters("${triggerAddr}", 1);`);
                      sb.push(`                         if (int.TryParse(tVal, out int tInt) && tInt == 0) break;`);
                      sb.push(`                         await Task.Delay(100, token);`);
                      sb.push(`                     }`);
                      sb.push(`                }`);
                      sb.push(`                catch (Exception ex) { Console.WriteLine($"Error in Handshake: {ex.Message}"); }`);
                      sb.push(`            }`);
                      sb.push(`        }`);
                  });
              }
          });
      });
  }

  sb.push(`    }`);
  sb.push(`}`);
  return sb.join('\n');
};

export const generatePLCPointXaml = (sheets: SheetData[], maxModules: number): string => {
  const sb: string[] = [];
  sb.push(`<UserControl x:Class="Core.PLCPoint"`);
  sb.push(`             xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"`);
  sb.push(`             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"`);
  sb.push(`             xmlns:local="clr-namespace:Core">`);
  sb.push(`    <UserControl.DataContext>`);
  sb.push(`        <x:Static Member="local:PLCPointProperty.Instance"/>`);
  sb.push(`    </UserControl.DataContext>`);
  sb.push(`    <Grid>`);
  sb.push(`        <Grid.RowDefinitions>`);
  sb.push(`            <RowDefinition Height="Auto"/>`);
  sb.push(`            <RowDefinition Height="*"/>`);
  sb.push(`        </Grid.RowDefinitions>`);
  sb.push(`        <StackPanel Orientation="Horizontal" Margin="5">`);
  sb.push(`            <TextBlock Text="Frequency (ms): " VerticalAlignment="Center" Foreground="White"/>`);
  sb.push(`            <ComboBox SelectedItem="{Binding Frequency}" Width="100">`);
  sb.push(`                <sys:Int32 xmlns:sys="clr-namespace:System;assembly=mscorlib">0</sys:Int32>`);
  sb.push(`                <sys:Int32 xmlns:sys="clr-namespace:System;assembly=mscorlib">500</sys:Int32>`);
  sb.push(`                <sys:Int32 xmlns:sys="clr-namespace:System;assembly=mscorlib">1000</sys:Int32>`);
  sb.push(`                <sys:Int32 xmlns:sys="clr-namespace:System;assembly=mscorlib">3000</sys:Int32>`);
  sb.push(`                <sys:Int32 xmlns:sys="clr-namespace:System;assembly=mscorlib">5000</sys:Int32>`);
  sb.push(`            </ComboBox>`);
  sb.push(`        </StackPanel>`);
  sb.push(`        <ScrollViewer Grid.Row="1">`);
  sb.push(`            <StackPanel>`);

  for (let m = 1; m <= maxModules; m++) {
    sb.push(`                <GroupBox Header="Module ${m}" Margin="5">`);
    sb.push(`                    <StackPanel>`);
    
    sheets.forEach(sheet => {
      sb.push(`                        <Expander Header="${sheet.name} (Device: ${sheet.deviceInfo?.DeviceName})" Margin="2">`);
      sb.push(`                            <StackPanel>`);
      
      sheet.points.forEach(p => {
        sb.push(`                                <Grid Margin="2">`);
        sb.push(`                                    <Grid.ColumnDefinitions>`);
        sb.push(`                                        <ColumnDefinition Width="200"/>`);
        sb.push(`                                        <ColumnDefinition Width="*"/>`);
        sb.push(`                                    </Grid.ColumnDefinitions>`);
        sb.push(`                                    <TextBlock Text="${p.KeyName} (${p.PropertyName}):" VerticalAlignment="Center"/>`);
        sb.push(`                                    <TextBox Grid.Column="1" Text="{Binding ${p.PropertyName}_M${m}}" IsReadOnly="True"/>`);
        sb.push(`                                </Grid>`);
      });

      sb.push(`                            </StackPanel>`);
      sb.push(`                        </Expander>`);
    });

    sb.push(`                    </StackPanel>`);
    sb.push(`                </GroupBox>`);
  }

  sb.push(`            </StackPanel>`);
  sb.push(`        </ScrollViewer>`);
  sb.push(`    </Grid>`);
  sb.push(`</UserControl>`);
  return sb.join('\n');
};

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
        const className = `${sheet.name}_PeriodAbs_${Math.abs(period)}`;
        if (processedClasses.has(className)) return; 
        processedClasses.add(className);

        sb.push(`    public class ${className} : TableTemplate`);
        sb.push(`    {`);
        groupPoints.forEach(p => {
             const csType = getCSharpType(p.Type);
             sb.push(`        [SugarColumn(ColumnDescription = "${p.KeyName}")]`);
             sb.push(`        public ${csType} ${p.PropertyName} { get; set; }`);
        });
        sb.push(`    }`);
        sb.push(``);
    });
  });

  sb.push(`}`);
  return sb.join('\n');
};

export const generateGlobalAuto = (sheets: SheetData[], maxModules: number): string => {
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
  sb.push(`            for (int i = 1; i <= maxModules; i++)`);
  sb.push(`            {`);
  
  sheets.forEach(s => {
      const ip = s.deviceInfo?.IP || "127.0.0.1";
      const port = s.deviceInfo?.Port || 502;
      sb.push(`                manager.AddTcp($"_{i}${s.name}", "${ip}", ${port});`);
  });
  
  sb.push(``);

  for(let m=1; m<=maxModules; m++) {
      sheets.forEach(s => {
          sb.push(`                if(i == ${m}) _${m}${s.name}PLCModbus = Global.manager.GetService($"_{i}${s.name}");`);
      });
  }

  sb.push(`            }`);
  sb.push(`        }`);
  sb.push(``);

  // 3. Database Repositories
  const tableNames = new Set<string>();
  sheets.forEach(sheet => {
      const pointsByPeriod = new Map<number, any>();
      sheet.points.forEach(p => { if(p.Period !== 0) pointsByPeriod.set(p.Period, true); });
      pointsByPeriod.forEach((_, period) => {
          tableNames.add(`${sheet.name}_PeriodAbs_${Math.abs(period)}`);
      });
  });

  sb.push(`        #region Database Init`);
  // Note: databaseHelper definition removed as requested. Assumed to be in other partial class.
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