import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { AppState, DeviceConfig, SheetData, PointDefinition, ByteOrder, StringByteOrder } from './types';
import { 
  generatePLCPointProperty, 
  generatePLCPointXaml, 
  generateSQLTable, 
  generateGlobalAuto 
} from './services/generator';

const App: React.FC = () => {
  const [state, setState] = useState<AppState>({
    maxModules: 2,
    byteOrder: 'ABCD', // Default standard numeric
    stringByteOrder: 'ABCD', // Default standard string
    maxGap: 20,
    maxBatchSize: 100,
    sheets: [],
    devices: [],
    isProcessing: false,
    error: null,
    logs: [],
    showReadme: false
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

      if (!workbook.Sheets['IP_Port']) {
        throw new Error('缺失 "IP_Port" 工作表');
      }
      
      const rawIpPort = XLSX.utils.sheet_to_json<any>(workbook.Sheets['IP_Port']);
      const ipPortData: DeviceConfig[] = rawIpPort.map(row => ({
        DeviceName: String(row.Device || '').trim(), 
        IP: String(row.IP || '').trim(),
        Port: Number(row.PORT || row.Port) || 502 
      })).filter(d => d.DeviceName); 

      const ips = new Set<string>();
      for (const d of ipPortData) {
          if (ips.has(d.IP)) {
              throw new Error(`IP冲突错误: IP地址 ${d.IP} 在 IP_Port 表中被重复定义。每个设备的IP必须唯一。`);
          }
          ips.add(d.IP);
      }

      addLog(`发现 ${ipPortData.length} 个设备在 IP_Port 表中。`);

      const sheets: SheetData[] = [];
      const allPropNames = new Set<string>();

      for (const originalSheetName of workbook.SheetNames) {
        const sheetName = originalSheetName.trim();
        if (sheetName === 'IP_Port') continue;

        const rawPoints = XLSX.utils.sheet_to_json<any>(workbook.Sheets[originalSheetName]);
        const points: PointDefinition[] = rawPoints.map((row: any) => {
            if (allPropNames.has(row.PropertyName)) {
                throw new Error(`在表 ${sheetName} 中发现重复的属性名: ${row.PropertyName}`);
            }
            allPropNames.add(row.PropertyName);

            return {
                PropertyName: row.PropertyName,
                KeyName: row.KeyName,
                ValueAddress: String(row.ValueAddress || ''),
                Type: row.Type,
                Length: Number(row.Length),
                TriggerAddress: row.TriggerAddress ? String(row.TriggerAddress) : undefined,
                ReturnAddress: row.ReturnAddress ? String(row.ReturnAddress) : undefined,
                Period: Number(row.Period)
            };
        });

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
            name: sheetName,
            points
        });
        addLog(`已解析工作表 ${sheetName}: ${points.length} 个点位。`);
      }

      for (let m = 1; m <= state.maxModules; m++) {
          for (const sheet of sheets) {
              const expectedDeviceName = `${sheet.name}_M${m}`;
              const exists = ipPortData.find(d => d.DeviceName === expectedDeviceName);
              if (!exists) {
                  throw new Error(`配置缺失: 未在 IP_Port 表中找到设备 "${expectedDeviceName}" (对应工作表 ${sheet.name}, 模组 ${m})。`);
              }
          }
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

        // Pass Configuration to Property Generator
        const csProperty = generatePLCPointProperty(
            state.sheets, 
            state.maxModules, 
            state.byteOrder, 
            state.stringByteOrder,
            state.maxGap,
            state.maxBatchSize
        );
        zip.file("PLCPointProperty.cs", csProperty);
        addLog(`已生成 PLCPointProperty.cs (Gap: ${state.maxGap}, Batch: ${state.maxBatchSize})`);

        const xaml = generatePLCPointXaml(state.sheets, state.maxModules);
        zip.file("PLCPoint.xaml", xaml);
        addLog("已生成 PLCPoint.xaml");

        const sqlTable = generateSQLTable(state.sheets);
        zip.file("SQLTable_Auto.cs", sqlTable);
        addLog("已生成 SQLTable_Auto.cs");

        const globalAuto = generateGlobalAuto(state.sheets, state.maxModules, state.devices);
        zip.file("Global_Auto.cs", globalAuto);
        addLog("已生成 Global_Auto.cs");

        const content = await zip.generateAsync({ type: "blob" });
        
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

  const toggleReadme = () => {
      setState(prev => ({ ...prev, showReadme: !prev.showReadme }));
  };

  return (
    <div className="min-h-screen bg-gray-100 p-8 font-sans relative">
      <div className="max-w-5xl mx-auto bg-white shadow-xl rounded-lg p-6">
        <div className="flex justify-between items-center mb-6">
            <div className="flex flex-col">
                <h1 className="text-3xl font-bold text-gray-800">WPF 低代码生成器</h1>
                <p className="text-gray-500">基于 Excel 定义生成 C# .NET 8 WPF 基础代码。</p>
            </div>
            <button 
                onClick={toggleReadme}
                className="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded shadow flex items-center gap-2"
            >
                <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className="w-5 h-5">
                    <path strokeLinecap="round" strokeLinejoin="round" d="M9.879 7.519c1.171-1.025 3.071-1.025 4.242 0 1.172 1.025 1.172 2.687 0 3.712-.203.179-.43.326-.67.442-.745.361-1.45.999-1.45 1.827v.75M21 12a9 9 0 1 1-18 0 9 9 0 0 1 18 0Zm-9 5.25h.008v.008H12v-.008Z" />
                </svg>
                使用说明
            </button>
        </div>

        {/* Step 1: Configuration */}
        <div className="mb-8 p-5 bg-blue-50 border border-blue-200 rounded-lg">
          <h2 className="text-xl font-semibold text-blue-800 mb-4 flex items-center gap-2">
            <span className="bg-blue-200 text-blue-800 w-8 h-8 rounded-full flex items-center justify-center text-sm">1</span> 
            全局配置
          </h2>
          
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
             {/* Core Settings */}
             <div className="flex flex-col">
                <label className="font-bold text-gray-700 mb-2">最大模组数 (Modules)</label>
                <input 
                  type="number" 
                  min="1" 
                  value={state.maxModules} 
                  onChange={(e) => setState(prev => ({...prev, maxModules: parseInt(e.target.value) || 1}))}
                  className="border border-gray-300 p-2 rounded focus:ring-2 focus:ring-blue-400 outline-none"
                />
             </div>
             
             {/* Optimization Settings */}
             <div className="flex flex-col">
                <label className="font-bold text-gray-700 mb-2 flex justify-between">
                    合并间隔 (Gap)
                    <span className="text-xs font-normal text-gray-500 self-center">跳过未定义地址数</span>
                </label>
                <input 
                  type="number" 
                  min="0" 
                  value={state.maxGap} 
                  onChange={(e) => setState(prev => ({...prev, maxGap: parseInt(e.target.value) || 0}))}
                  className="border border-gray-300 p-2 rounded focus:ring-2 focus:ring-blue-400 outline-none"
                />
             </div>

             <div className="flex flex-col">
                <label className="font-bold text-gray-700 mb-2 flex justify-between">
                    最大包长 (Batch Size)
                    <span className="text-xs font-normal text-gray-500 self-center">单次请求最大字数</span>
                </label>
                <input 
                  type="number" 
                  min="1" 
                  value={state.maxBatchSize} 
                  onChange={(e) => setState(prev => ({...prev, maxBatchSize: parseInt(e.target.value) || 100}))}
                  className="border border-gray-300 p-2 rounded focus:ring-2 focus:ring-blue-400 outline-none"
                />
             </div>

             {/* Endianness */}
             <div className="flex flex-col">
                <label className="font-bold text-gray-700 mb-2">数值字节序 (Numeric)</label>
                <select 
                  value={state.byteOrder} 
                  onChange={(e) => setState(prev => ({...prev, byteOrder: e.target.value as ByteOrder}))}
                  className="border border-gray-300 p-2 rounded bg-white focus:ring-2 focus:ring-blue-400 outline-none"
                >
                  <option value="ABCD">ABCD (Big Endian - Std)</option>
                  <option value="CDAB">CDAB (Little Endian Word)</option>
                  <option value="DCBA">DCBA (Little Endian)</option>
                  <option value="BADC">BADC (Big Endian Byte)</option>
                </select>
             </div>

             <div className="flex flex-col">
                <label className="font-bold text-gray-700 mb-2">字符串字节序 (String)</label>
                <select 
                  value={state.stringByteOrder} 
                  onChange={(e) => setState(prev => ({...prev, stringByteOrder: e.target.value as StringByteOrder}))}
                  className="border border-gray-300 p-2 rounded bg-white focus:ring-2 focus:ring-blue-400 outline-none"
                >
                  <option value="ABCD">ABCD (标准 - HighByte First)</option>
                  <option value="BADC">BADC (交换 - LowByte First)</option>
                </select>
             </div>
          </div>
        </div>

        {/* Step 2: Upload */}
        <div className="mb-8 p-5 bg-indigo-50 border border-indigo-200 rounded-lg">
          <h2 className="text-xl font-semibold text-indigo-800 mb-4 flex items-center gap-2">
            <span className="bg-indigo-200 text-indigo-800 w-8 h-8 rounded-full flex items-center justify-center text-sm">2</span> 
            导入定义 (Excel)
          </h2>
          <div className="flex items-center gap-4">
              <input 
                type="file" 
                accept=".xlsx, .xls"
                onChange={handleFileUpload}
                className="block w-full text-sm text-gray-500
                  file:mr-4 file:py-2 file:px-4
                  file:rounded-full file:border-0
                  file:text-sm file:font-semibold
                  file:bg-indigo-600 file:text-white
                  hover:file:bg-indigo-700
                  cursor-pointer
                "
              />
          </div>
          {state.error && (
             <div className="mt-4 p-3 bg-red-100 text-red-700 rounded border border-red-300 flex items-start">
                <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className="w-6 h-6 mr-2 shrink-0">
                  <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v3.75m9-.75a9 9 0 1 1-18 0 9 9 0 0 1 18 0Zm-9 3.75h.008v.008H12v-.008Z" />
                </svg>
                <span><strong>错误:</strong> {state.error}</span>
             </div>
          )}
        </div>

        {/* Step 3: Status & Action */}
        <div className="mb-8">
           <div className="flex justify-between items-center mb-4">
              <h2 className="text-xl font-semibold text-gray-800 flex items-center gap-2">
                <span className="bg-gray-200 text-gray-800 w-8 h-8 rounded-full flex items-center justify-center text-sm">3</span> 
                生成与日志
              </h2>
              <button 
                onClick={handleGenerate}
                disabled={state.sheets.length === 0 || state.isProcessing}
                className={`px-6 py-2 rounded font-bold text-white transition-colors shadow-lg ${
                    state.sheets.length === 0 ? 'bg-gray-400 cursor-not-allowed' : 'bg-green-600 hover:bg-green-700'
                }`}
              >
                生成并下载代码 (ZIP)
              </button>
           </div>
           
           <div className="bg-gray-900 text-green-400 p-4 rounded-lg h-64 overflow-y-auto font-mono text-sm border border-gray-700 shadow-inner">
              {state.logs.length === 0 ? (
                  <span className="text-gray-500 opacity-50">// 等待文件导入...</span>
              ) : (
                  state.logs.map((log, i) => <div key={i} className="mb-1 border-b border-gray-800 pb-1 last:border-0">{`> ${log}`}</div>)
              )}
           </div>
        </div>
      </div>

      {/* Readme Modal */}
      {state.showReadme && (
          <div className="fixed inset-0 bg-black bg-opacity-50 z-50 flex items-center justify-center p-4">
              <div className="bg-white rounded-lg shadow-2xl max-w-4xl w-full max-h-[90vh] flex flex-col">
                  <div className="flex justify-between items-center p-6 border-b">
                      <h3 className="text-2xl font-bold text-gray-800">使用说明与逻辑介绍</h3>
                      <button onClick={toggleReadme} className="text-gray-500 hover:text-gray-700">
                        <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className="w-6 h-6">
                          <path strokeLinecap="round" strokeLinejoin="round" d="M6 18 18 6M6 6l12 12" />
                        </svg>
                      </button>
                  </div>
                  <div className="p-6 overflow-y-auto">
                      <div className="prose max-w-none text-gray-700 space-y-4">
                          <section>
                              <h4 className="text-lg font-bold text-blue-800 border-b border-blue-100 pb-1 mb-2">1. Excel 模板格式</h4>
                              <p>Excel 文件必须包含一个名为 <strong>IP_Port</strong> 的 Sheet，以及若干个定义点位的 Sheet。</p>
                              <ul className="list-disc pl-5 mt-2 space-y-1">
                                  <li><strong>IP_Port 表:</strong> 包含 <code>Device</code> (设备名), <code>IP</code>, <code>Port</code>。</li>
                                  <li><strong>点位定义表:</strong> 任意名称，列包含: 
                                      <ul className="list-circle pl-5 mt-1 text-sm text-gray-600">
                                          <li><code>PropertyName</code>: 属性名 (英文，无空格)</li>
                                          <li><code>KeyName</code>: 中文描述 (用于 UI 显示或数据库注释)</li>
                                          <li><code>ValueAddress</code>: Modbus 地址 (例如 D100, 100)</li>
                                          <li><code>Type</code>: 数据类型 (bool, short, int, float, string, short[] 等)</li>
                                          <li><code>Length</code>: 数组长度或字符串长度 (对于基础类型 int/float 自动计算)</li>
                                          <li><code>Period</code>: 采样/存储周期 (毫秒)。如果是 0 则仅在 UI 刷新，不存库。</li>
                                          <li><code>TriggerAddress / ReturnAddress</code>: 握手信号地址 (用于握手日志)</li>
                                      </ul>
                                  </li>
                              </ul>
                          </section>

                          <section>
                              <h4 className="text-lg font-bold text-blue-800 border-b border-blue-100 pb-1 mb-2">2. 生成器核心逻辑</h4>
                              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                  <div className="bg-gray-50 p-3 rounded">
                                      <h5 className="font-bold text-gray-900">数据监视 (Monitor)</h5>
                                      <p className="text-sm mt-1">
                                          所有定义的点位都会自动生成后台读取任务。系统会根据 <strong>合并间隔 (Gap)</strong> 
                                          自动将连续的地址合并为一个 Modbus 请求，以减少网络开销。
                                      </p>
                                  </div>
                                  <div className="bg-gray-50 p-3 rounded">
                                      <h5 className="font-bold text-gray-900">定时存储 (Periodic Log)</h5>
                                      <p className="text-sm mt-1">
                                          当 <code>Period >= 1</code> 时，系统按固定周期 (ms) 自动快照数据并写入数据库。
                                      </p>
                                  </div>
                                  <div className="bg-gray-50 p-3 rounded">
                                      <h5 className="font-bold text-gray-900">变化存储 (Change Log)</h5>
                                      <p className="text-sm mt-1">
                                          当 <code>0 &lt; Period &lt; 1</code> (如 0.1) 时，系统进入高频检测模式。只有当数据发生变化时才写入数据库。
                                          Period 字段此时代表检测频率 (如 0.1 代表 100ms 检测一次)。
                                      </p>
                                  </div>
                                  <div className="bg-gray-50 p-3 rounded">
                                      <h5 className="font-bold text-gray-900">握手存储 (Handshake Log)</h5>
                                      <p className="text-sm mt-1">
                                          当 <code>Period &lt; 0</code> 且填写了 Trigger/Return 时，启用握手模式。
                                          流程: 检测 Trigger=11 -> 读取数据并存库 -> 写入 Return=11 -> 等待 PLC 复位。
                                      </p>
                                  </div>
                              </div>
                          </section>

                          <section>
                              <h4 className="text-lg font-bold text-blue-800 border-b border-blue-100 pb-1 mb-2">3. 优化参数说明</h4>
                              <ul className="list-disc pl-5 space-y-2 text-sm">
                                  <li><strong>合并间隔 (Gap):</strong> 当两个点位的地址差值小于此值时，生成器会将其视为“连续”段，用一条指令读取。
                                      <br/><span className="text-gray-500">调大此值可减少请求次数，但会读取更多无用数据增加带宽。默认建议 20。</span>
                                  </li>
                                  <li><strong>最大包长 (Batch Size):</strong> 单个 Modbus 读取指令允许请求的最大寄存器数量 (Word)。
                                      <br/><span className="text-gray-500">标准 Modbus TCP 通常支持到 120-125 左右。默认建议 100。</span>
                                  </li>
                              </ul>
                          </section>
                      </div>
                  </div>
                  <div className="p-4 border-t bg-gray-50 rounded-b-lg flex justify-end">
                      <button 
                        onClick={toggleReadme}
                        className="bg-blue-600 hover:bg-blue-700 text-white px-6 py-2 rounded font-bold"
                      >
                        关闭
                      </button>
                  </div>
              </div>
          </div>
      )}
    </div>
  );
};

export default App;