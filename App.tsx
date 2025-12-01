import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { AppState, DeviceConfig, SheetData, PointDefinition } from './types';
import { 
  generatePLCPointProperty, 
  generatePLCPointXaml, 
  generateSQLTable, 
  generateGlobalAuto 
} from './services/generator';

const App: React.FC = () => {
  const [state, setState] = useState<AppState>({
    maxModules: 2,
    sheets: [],
    devices: [],
    isProcessing: false,
    error: null,
    logs: []
  });

  const addLog = (msg: string) => {
    setState(prev => ({ ...prev, logs: [...prev.logs, msg] }));
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setState(prev => ({ ...prev, isProcessing: true, error: null, logs: ['正在加载文件...'] }));

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data);

      // 1. Parse IP_Port
      if (!workbook.Sheets['IP_Port']) {
        throw new Error('缺失 "IP_Port" 工作表');
      }
      
      const rawIpPort = XLSX.utils.sheet_to_json<any>(workbook.Sheets['IP_Port']);
      // Sanitize IP_Port data (trim whitespace)
    const ipPortData: DeviceConfig[] = rawIpPort.map(row => ({
    // 1. Excel 表头是 "Device"，不是 "DeviceName"
    DeviceName: String(row.Device || '').trim(), 
    
    IP: String(row.IP || '').trim(),
    
    // 2. Excel 表头是 "PORT" (全大写)，建议做个兼容处理
    Port: Number(row.PORT || row.Port) || 502 
})).filter(d => d.DeviceName); // 现在 DeviceName 有值了，就不会被过滤了

      addLog(`发现 ${ipPortData.length} 个设备在 IP_Port 表中。`);

      // 2. Parse Modules
      const sheets: SheetData[] = [];
      const allPropNames = new Set<string>();

      for (const originalSheetName of workbook.SheetNames) {
        // Trim sheet name for comparison
        const sheetName = originalSheetName.trim();
        
        if (sheetName === 'IP_Port') continue;

        // Check if device exists using trimmed name
        const matchedDevice = ipPortData.find(d => d.DeviceName === sheetName);
        
        if (!matchedDevice) {
             addLog(`警告: 工作表 ${sheetName} 未在 IP_Port 设备列表中找到。映射可能会失败。`);
        }

        const rawPoints = XLSX.utils.sheet_to_json<any>(workbook.Sheets[originalSheetName]);
        const points: PointDefinition[] = rawPoints.map((row: any) => {
            // Validate unique property name
            if (allPropNames.has(row.PropertyName)) {
                throw new Error(`在表 ${sheetName} 中发现重复的属性名: ${row.PropertyName}`);
            }
            allPropNames.add(row.PropertyName);

            return {
                PropertyName: row.PropertyName,
                KeyName: row.KeyName,
                // Explicitly cast addresses to String to handle numeric values from Excel
                ValueAddress: String(row.ValueAddress || ''),
                Type: row.Type,
                Length: Number(row.Length),
                TriggerAddress: row.TriggerAddress ? String(row.TriggerAddress) : undefined,
                ReturnAddress: row.ReturnAddress ? String(row.ReturnAddress) : undefined,
                Period: Number(row.Period)
            };
        });

        // Validate Trigger Consistency
        const triggerMap = new Map<string, number>();
        points.forEach(p => {
            if (p.TriggerAddress) {
                if (triggerMap.has(p.TriggerAddress)) {
                    if (triggerMap.get(p.TriggerAddress) !== p.Period) {
                        throw new Error(`一致性错误: 触发器 ${p.TriggerAddress} 在 ${sheetName} 中包含多个不同的周期值`);
                    }
                } else {
                    triggerMap.set(p.TriggerAddress, p.Period);
                }
            }
        });

        sheets.push({
            name: sheetName, // Use the trimmed name
            points,
            deviceInfo: matchedDevice
        });
        addLog(`已解析工作表 ${sheetName}: ${points.length} 个点位。`);
      }

      setState(prev => ({
        ...prev,
        devices: ipPortData,
        sheets,
        isProcessing: false
      }));

    } catch (err: any) {
      setState(prev => ({ ...prev, isProcessing: false, error: err.message }));
    }
  };

  const handleGenerate = async () => {
    if (state.sheets.length === 0) {
        setState(prev => ({...prev, error: "未加载任何数据。"}));
        return;
    }

    try {
        const zip = new JSZip();

        // 1. PLCPointProperty.cs
        const csProperty = generatePLCPointProperty(state.sheets, state.maxModules);
        zip.file("PLCPointProperty.cs", csProperty);
        addLog("已生成 PLCPointProperty.cs");

        // 2. PLCPoint.xaml
        const xaml = generatePLCPointXaml(state.sheets, state.maxModules);
        zip.file("PLCPoint.xaml", xaml);
        addLog("已生成 PLCPoint.xaml");

        // 3. SQLTable_Auto.cs (Renamed from SQLTable.cs)
        const sqlTable = generateSQLTable(state.sheets);
        zip.file("SQLTable_Auto.cs", sqlTable);
        addLog("已生成 SQLTable_Auto.cs");

        // 4. Global_Auto.cs
        const globalAuto = generateGlobalAuto(state.sheets, state.maxModules);
        zip.file("Global_Auto.cs", globalAuto);
        addLog("已生成 Global_Auto.cs");

        const content = await zip.generateAsync({ type: "blob" });
        
        // Use native download
        const url = URL.createObjectURL(content);
        const link = document.createElement('a');
        link.href = url;
        link.download = "WPF_Generated_Code.zip";
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        
        addLog("下载已开始。");

    } catch (err: any) {
        setState(prev => ({...prev, error: "生成失败: " + err.message}));
    }
  };

  return (
    <div className="min-h-screen bg-gray-100 p-8 font-sans">
      <div className="max-w-4xl mx-auto bg-white shadow-xl rounded-lg p-6">
        <h1 className="text-3xl font-bold text-gray-800 mb-2">WPF 低代码生成器</h1>
        <p className="text-gray-500 mb-8">基于 Excel 定义生成 C# .NET 8 WPF 基础代码（PLC 通讯与 SQL 存储）。</p>

        {/* Step 1: Configuration */}
        <div className="mb-8 p-4 bg-blue-50 border border-blue-200 rounded">
          <h2 className="text-xl font-semibold text-blue-800 mb-4">1. 配置</h2>
          <div className="flex items-center gap-4">
            <label className="font-medium">最大模组数 (Max Modules):</label>
            <input 
              type="number" 
              min="1" 
              value={state.maxModules} 
              onChange={(e) => setState(prev => ({...prev, maxModules: parseInt(e.target.value) || 1}))}
              className="border p-2 rounded w-24"
            />
          </div>
        </div>

        {/* Step 2: Upload */}
        <div className="mb-8 p-4 bg-blue-50 border border-blue-200 rounded">
          <h2 className="text-xl font-semibold text-blue-800 mb-4">2. 导入定义 (Excel)</h2>
          <input 
            type="file" 
            accept=".xlsx, .xls"
            onChange={handleFileUpload}
            className="block w-full text-sm text-gray-500
              file:mr-4 file:py-2 file:px-4
              file:rounded-full file:border-0
              file:text-sm file:font-semibold
              file:bg-blue-600 file:text-white
              hover:file:bg-blue-700
            "
          />
          {state.error && (
             <div className="mt-4 p-3 bg-red-100 text-red-700 rounded border border-red-300">
                <strong>错误:</strong> {state.error}
             </div>
          )}
        </div>

        {/* Step 3: Status & Action */}
        <div className="mb-8">
           <div className="flex justify-between items-center mb-4">
              <h2 className="text-xl font-semibold text-gray-800">3. 状态</h2>
              <button 
                onClick={handleGenerate}
                disabled={state.sheets.length === 0 || state.isProcessing}
                className={`px-6 py-2 rounded font-bold text-white transition-colors ${
                    state.sheets.length === 0 ? 'bg-gray-400 cursor-not-allowed' : 'bg-green-600 hover:bg-green-700'
                }`}
              >
                生成并下载代码
              </button>
           </div>
           
           <div className="bg-gray-900 text-green-400 p-4 rounded h-64 overflow-y-auto font-mono text-sm">
              {state.logs.length === 0 ? (
                  <span className="text-gray-500">// 等待输入...</span>
              ) : (
                  state.logs.map((log, i) => <div key={i}>{`> ${log}`}</div>)
              )}
           </div>
        </div>
      </div>
    </div>
  );
};

export default App;