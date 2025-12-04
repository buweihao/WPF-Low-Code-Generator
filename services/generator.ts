import { SheetData, PointDefinition, DeviceConfig } from '../types';

// Helper to clean address (remove D prefix)
const cleanAddr = (addr: string | number | undefined): string => {
  if (addr === undefined || addr === null || addr === '') return "0";
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

// Helper to get parse method name
const getParseMethod = (type: string): string => {
  switch (type.toLowerCase()) {
    case 'bool': return 'ParseBool';
    case 'short': return 'ParseShort';
    case 'int': return 'ParseInt';
    case 'float': return 'ParseFloat';
    default: return ''; // string doesn't need parsing
  }
};

// Helper to map type to Modbus method
const getReadMethod = (type: string): string => {
  switch (type.toLowerCase()) {
    case 'bool': return `ReadCoils`;
    case 'short': return `ReadRegisters`;
    case 'int': return `ReadDInt`;
    case 'string': return `ReadString`;
    default: return `ReadRegisters`;
  }
};

// --- Optimization Logic ---

interface TagInfo {
    Name: string; // Full property name (e.g., Prop_M1)
    Address: number;
    Length: number;
    Type: string;
}

interface RequestBlock {
    StartAddress: number;
    Length: number;
    IncludedTags: TagInfo[];
}

const optimizeRequests = (tags: TagInfo[]): RequestBlock[] => {
    if (!tags || tags.length === 0) return [];

    // 1. Sort by address
    tags.sort((a, b) => a.Address - b.Address);

    const blocks: RequestBlock[] = [];
    const MAX_GAP = 20;
    const MAX_BATCH_SIZE = 100;

    let currentBlock: RequestBlock = {
        StartAddress: tags[0].Address,
        Length: tags[0].Length,
        IncludedTags: [tags[0]]
    };

    for (let i = 1; i < tags.length; i++) {
        const tag = tags[i];
        
        const currentEnd = currentBlock.StartAddress + currentBlock.Length;
        const gap = tag.Address - currentEnd;
        const newLength = (tag.Address + tag.Length) - currentBlock.StartAddress;

        // Merge if gap is small and total size is within limit
        // Note: gap < 0 means overlap, which is fine to merge
        if (gap <= MAX_GAP && newLength <= MAX_BATCH_SIZE) {
            currentBlock.Length = newLength;
            currentBlock.IncludedTags.push(tag);
        } else {
            blocks.push(currentBlock);
            currentBlock = {
                StartAddress: tag.Address,
                Length: tag.Length,
                IncludedTags: [tag]
            };
        }
    }
    blocks.push(currentBlock);

    return blocks;
};

// --- Generator Functions ---

export const generatePLCPointProperty = (sheets: SheetData[], maxModules: number): string => {
  const sb: string[] = [];

  sb.push(`using System;`);
  sb.push(`using System.Collections.Generic;`);
  sb.push(`using System.Collections.Concurrent;`);
  sb.push(`using System.IO;`);
  sb.push(`using System.ComponentModel;`);
  sb.push(`using System.Threading;`);
  sb.push(`using System.Threading.Tasks;`);
  sb.push(`using System.Diagnostics;`);
  sb.push(`using System.Text;`);
  sb.push(`using System.Linq;`);
  sb.push(`using System.Reflection;`);
  sb.push(`using PropertyChanged;`); // For Fody
  sb.push(``);
  sb.push(`namespace Core`);
  sb.push(`{`);
  sb.push(`    [AddINotifyPropertyChangedInterface]`);
  sb.push(`    public class PLCPointProperty : INotifyPropertyChanged`);
  sb.push(`    {`);
  sb.push(`        public static PLCPointProperty Instance { get; } = new PLCPointProperty();`);
  sb.push(``);
  sb.push(`        // 构造函数`);
  sb.push(`        private PLCPointProperty()`);
  sb.push(`        {`);
  sb.push(`            InitializePropertyMap();`);
  sb.push(`            UpdateAllPropertyCaches();`);
  sb.push(`            _ = Task.Run(ProcessLogQueue);`);
  sb.push(`        }`);
  sb.push(``);
  sb.push(`        // 频率控制 (Frequency)`);
  sb.push(`        private int _frequency = 1000;`);
  sb.push(`        public int Frequency`);
  sb.push(`        {`);
  sb.push(`            get => _frequency;`);
  sb.push(`            set { _frequency = value; StartAllTasks(); }`);
  sb.push(`        }`);
  sb.push(``);
  
  // --- Recording Logic ---
  sb.push(`        // ================== 录制逻辑 (Recording) ==================`);
  sb.push(`        private string _recordState = "Stopped";`);
  sb.push(`        public string RecordState`);
  sb.push(`        {`);
  sb.push(`            get => _recordState;`);
  sb.push(`            set`);
  sb.push(`            {`);
  sb.push(`                _recordState = value;`);
  sb.push(`                IsRecording = (value == "Recording");`);
  sb.push(`            }`);
  sb.push(`        }`);
  sb.push(`        public bool IsRecording { get; private set; }`);
  sb.push(``);
  sb.push(`        private ConcurrentQueue<string> _logQueue = new ConcurrentQueue<string>();`);
  sb.push(`        private class LogMetadata { public string KeyName; public string Type; public PropertyInfo Info; }`);
  sb.push(`        private Dictionary<string, LogMetadata> _logMetadata = new Dictionary<string, LogMetadata>();`);
  sb.push(``);
  sb.push(`        private async Task ProcessLogQueue()`);
  sb.push(`        {`);
  sb.push(`            while (true)`);
  sb.push(`            {`);
  sb.push(`                try`);
  sb.push(`                {`);
  sb.push(`                    if (_logQueue.IsEmpty) { await Task.Delay(200); continue; }`);
  sb.push(`                    var lines = new List<string>();`);
  sb.push(`                    while (_logQueue.TryDequeue(out var line)) lines.Add(line);`);
  sb.push(`                    if (lines.Count > 0)`);
  sb.push(`                    {`);
  sb.push(`                        string path = $"Log_{DateTime.Now:yyyyMMdd}.csv";`);
  sb.push(`                        // 使用带BOM的UTF8编码，解决Excel打开中文乱码问题`);
  sb.push(`                        using (var sw = new StreamWriter(path, true, new System.Text.UTF8Encoding(true)))`);
  sb.push(`                        {`);
  sb.push(`                            foreach (var line in lines) sw.WriteLine(line);`);
  sb.push(`                        }`);
  sb.push(`                    }`);
  sb.push(`                }`);
  sb.push(`                catch { await Task.Delay(1000); }`);
  sb.push(`            }`);
  sb.push(`        }`);
  sb.push(``);

  // --- Dynamic Module Switching Logic ---
  sb.push(`        // ================== 1. 映射配置区域 ==================`);
  sb.push(`        private Dictionary<string, PropertyInfo> _reflectionCache = new Dictionary<string, PropertyInfo>();`);
  sb.push(`        private Dictionary<string, string> _nameMapping = new Dictionary<string, string>();`);
  sb.push(``);
  sb.push(`        private void InitializePropertyMap()`);
  sb.push(`        {`);
  sb.push(`            // UI Mapping`);
  sheets.forEach(sheet => {
      sheet.points.forEach(p => {
          sb.push(`            _nameMapping.Add(nameof(Current${p.PropertyName}), "${p.PropertyName}_M");`);
      });
  });
  sb.push(``);
  sb.push(`            // Log Metadata Initialization`);
  for(let m=1; m<=maxModules; m++) {
      sheets.forEach(sheet => {
          sheet.points.forEach(p => {
               const propName = `${p.PropertyName}_M${m}`;
               // KeyName可能包含引号，简单转义
               const safeKeyName = p.KeyName ? p.KeyName.replace(/"/g, '\\"') : "";
               sb.push(`            _logMetadata["${propName}"] = new LogMetadata { KeyName = "${safeKeyName}", Type = "${p.Type}", Info = this.GetType().GetProperty("${propName}") };`);
          });
      });
  }
  sb.push(`        }`);
  sb.push(``);
  sb.push(`        // ================== 2. 模组切换逻辑 ==================`);
  sb.push(`        public int MaxModuleCount { get; set; } = ${maxModules};`);
  sb.push(`        private int _currentModuleIndex = 1;`);
  sb.push(`        public int CurrentModuleIndex`);
  sb.push(`        {`);
  sb.push(`            get => _currentModuleIndex;`);
  sb.push(`            set`);
  sb.push(`            {`);
  sb.push(`                if (value < 1) value = 1;`);
  sb.push(`                if (value > MaxModuleCount) value = MaxModuleCount;`);
  sb.push(``);
  sb.push(`                if (_currentModuleIndex != value)`);
  sb.push(`                {`);
  sb.push(`                    _currentModuleIndex = value;`);
  sb.push(`                    UpdateAllPropertyCaches();`);
  sb.push(`                    NotifyAllCurrentProperties();`);
  sb.push(`                }`);
  sb.push(`            }`);
  sb.push(`        }`);
  sb.push(``);
  sb.push(`        private void UpdateAllPropertyCaches()`);
  sb.push(`        {`);
  sb.push(`            foreach (var kvp in _nameMapping)`);
  sb.push(`            {`);
  sb.push(`                string uiName = kvp.Key;`);
  sb.push(`                string prefix = kvp.Value;`);
  sb.push(`                string realName = $"{prefix}{CurrentModuleIndex}";`);
  sb.push(`                var propInfo = this.GetType().GetProperty(realName);`);
  sb.push(`                if (_reflectionCache.ContainsKey(uiName)) _reflectionCache[uiName] = propInfo;`);
  sb.push(`                else _reflectionCache.Add(uiName, propInfo);`);
  sb.push(`            }`);
  sb.push(`        }`);
  sb.push(``);
  sb.push(`        private void NotifyAllCurrentProperties()`);
  sb.push(`        {`);
  sb.push(`            foreach (var key in _nameMapping.Keys)`);
  sb.push(`            {`);
  sb.push(`                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(key));`);
  sb.push(`            }`);
  sb.push(`        }`);
  sb.push(``);
  sb.push(`        // ================== 3. 通用取值/赋值辅助方法 ==================`);
  sb.push(`        private T GetDynamicValue<T>(string uiPropertyName)`);
  sb.push(`        {`);
  sb.push(`            if (_reflectionCache.TryGetValue(uiPropertyName, out PropertyInfo info) && info != null)`);
  sb.push(`            {`);
  sb.push(`                var val = info.GetValue(this);`);
  sb.push(`                try { return (T)Convert.ChangeType(val, typeof(T)); } catch { return default(T); }`);
  sb.push(`            }`);
  sb.push(`            return default(T);`);
  sb.push(`        }`);
  sb.push(``);
  sb.push(`        private void SetDynamicValue(string uiPropertyName, object value)`);
  sb.push(`        {`);
  sb.push(`            if (_reflectionCache.TryGetValue(uiPropertyName, out PropertyInfo info) && info != null)`);
  sb.push(`            {`);
  sb.push(`                info.SetValue(this, value);`);
  sb.push(`            }`);
  sb.push(`        }`);
  sb.push(``);
  
  sb.push(`        // ================== 4. 界面绑定属性 (CurrentX) ==================`);
  sheets.forEach(sheet => {
      sheet.points.forEach(p => {
          const csType = getCSharpType(p.Type);
          sb.push(`        public ${csType} Current${p.PropertyName}`);
          sb.push(`        {`);
          sb.push(`            get => GetDynamicValue<${csType}>(nameof(Current${p.PropertyName}));`);
          sb.push(`            set => SetDynamicValue(nameof(Current${p.PropertyName}), value);`);
          sb.push(`        }`);
      });
  });
  sb.push(``);

  sb.push(`        public event PropertyChangedEventHandler PropertyChanged;`);
  sb.push(`        protected void OnPropertyChanged(string name)`);
  sb.push(`        {`);
  sb.push(`            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));`);
  sb.push(``);
  
  // Logging Hook
  sb.push(`            // Check Logging`);
  sb.push(`            if (IsRecording && _logMetadata.TryGetValue(name, out var meta))`);
  sb.push(`            {`);
  sb.push(`                try {`);
  sb.push(`                    var val = meta.Info.GetValue(this);`);
  sb.push(`                    // 处理KeyName包含逗号的情况`);
  sb.push(`                    string safeKey = (meta.KeyName != null && meta.KeyName.Contains(",")) ? $"\\"{meta.KeyName}\\"" : meta.KeyName;`);
  sb.push(`                    _logQueue.Enqueue($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff},{safeKey},{meta.Type},{val}");`);
  sb.push(`                } catch {}`);
  sb.push(`            }`);
  sb.push(``);
  
  // CurrentModule Update Hook
  sb.push(`            if (!name.EndsWith($"_M{CurrentModuleIndex}")) return;`);
  sb.push(`            var targetEntry = _reflectionCache.FirstOrDefault(x => x.Value != null && x.Value.Name == name);`);
  sb.push(`            if (targetEntry.Key != null) PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(targetEntry.Key));`);
  sb.push(`        }`);
  sb.push(``);

  sb.push(`        #region 辅助解析方法 (For Batch & Single)`);
  sb.push(`        private bool ParseBool(string val) => bool.TryParse(val, out var b) ? b : false;`);
  sb.push(`        private short ParseShort(string val) => short.TryParse(val, out var s) ? s : (short)0;`);
  sb.push(`        private int ParseInt(string val) => int.TryParse(val, out var i) ? i : 0;`);
  sb.push(`        private float ParseFloat(string val) => float.TryParse(val, out var f) ? f : 0f;`);
  sb.push(``);
  sb.push(`        // 批量读取辅助`);
  sb.push(`        private ushort[] ParseRegisters(string raw)`);
  sb.push(`        {`);
  sb.push(`            if (string.IsNullOrWhiteSpace(raw)) return new ushort[0];`);
  sb.push(`            // 假设 IModbusService 返回类似 "123, 456" 或 "123 456" 的格式`);
  sb.push(`            var parts = raw.Split(new[] { ',', ' ', '[', ']' }, StringSplitOptions.RemoveEmptyEntries);`);
  sb.push(`            var list = new List<ushort>();`);
  sb.push(`            foreach (var p in parts) if (ushort.TryParse(p, out var v)) list.Add(v);`);
  sb.push(`            return list.ToArray();`);
  sb.push(``);
  sb.push(`        }`);
  sb.push(``);
  sb.push(`        private bool[] ParseCoils(string raw)`);
  sb.push(`        {`);
  sb.push(`            if (string.IsNullOrWhiteSpace(raw)) return new bool[0];`);
  sb.push(`            var parts = raw.Split(new[] { ',', ' ', '[', ']' }, StringSplitOptions.RemoveEmptyEntries);`);
  sb.push(`            var list = new List<bool>();`);
  sb.push(`            foreach (var p in parts)`);
  sb.push(`            {`);
  sb.push(`                 if (bool.TryParse(p, out var b)) list.Add(b);`);
  sb.push(`                 else if (int.TryParse(p, out var i)) list.Add(i != 0);`);
  sb.push(`            }`);
  sb.push(`            return list.ToArray();`);
  sb.push(`        }`);
  sb.push(``);
  sb.push(`        private short GetShort(ushort[] data, int offset) => (offset < data.Length) ? (short)data[offset] : (short)0;`);
  sb.push(`        private int GetInt(ushort[] data, int offset) => (offset + 1 < data.Length) ? (int)((data[offset] << 16) | data[offset + 1]) : 0;`);
  sb.push(`        private float GetFloat(ushort[] data, int offset)`);
  sb.push(`        {`);
  sb.push(`            if (offset + 1 >= data.Length) return 0f;`);
  sb.push(`            // 简易转换，假定大端字序 (High Word First)`);
  sb.push(`            byte[] bytes = new byte[4];`);
  sb.push(`            byte[] high = BitConverter.GetBytes(data[offset]);`);
  sb.push(`            byte[] low = BitConverter.GetBytes(data[offset + 1]);`);
  sb.push(`            // 根据实际PLC字序可能需要调整复制顺序`);
  sb.push(`            Buffer.BlockCopy(high, 0, bytes, 2, 2);`);
  sb.push(`            Buffer.BlockCopy(low, 0, bytes, 0, 2);`);
  sb.push(`            return BitConverter.ToSingle(bytes, 0);`);
  sb.push(`        }`);
  sb.push(`        private string GetString(ushort[] data, int offset, int length)`);
  sb.push(`        {`);
  sb.push(`            if (offset + length > data.Length) return string.Empty;`);
  sb.push(`            var bytes = new List<byte>();`);
  sb.push(`            for (int i = 0; i < length; i++)`);
  sb.push(`            {`);
  sb.push(`                ushort val = data[offset + i];`);
  sb.push(`                bytes.Add((byte)(val & 0xFF));`);
  sb.push(`                bytes.Add((byte)(val >> 8));`);
  sb.push(`            }`);
  sb.push(`            return Encoding.ASCII.GetString(bytes.ToArray()).Trim('\\0');`);
  sb.push(`        }`);
  sb.push(`        #endregion`);
  sb.push(``);

  sb.push(`        #region 底层属性 (Mx) - Fody Auto Properties`);
  for (let m = 1; m <= maxModules; m++) {
    sheets.forEach(sheet => {
      sheet.points.forEach(p => {
        const propName = `${p.PropertyName}_M${m}`;
        const csType = getCSharpType(p.Type);
        sb.push(`        public ${csType} ${propName} { get; set; }`);
      });
    });
  }
  sb.push(`        #endregion`);
  sb.push(``);
  
  // --- Task Helpers ---
  sb.push(`        #region Task Helpers`);
  sb.push(`        private Dictionary<string, CancellationTokenSource> _tasks = new Dictionary<string, CancellationTokenSource>();`);
  sb.push(``);
  sb.push(`        private void StartTask(string name, Func<CancellationToken, Task> action)`);
  sb.push(`        {`);
  sb.push(`            var cts = new CancellationTokenSource();`);
  sb.push(`            _tasks[name] = cts;`);
  sb.push(`            Task.Run(() => action(cts.Token), cts.Token);`);
  sb.push(`        }`);
  sb.push(``);
  
  // Helper: Monitor Value (Single)
  sb.push(`        private async Task MonitorValue<T>(CancellationToken token, IModbusService plc, Func<IModbusService, T> readFunc, Action<T> setProperty)`);
  sb.push(`        {`);
  sb.push(`            while (!token.IsCancellationRequested)`);
  sb.push(`            {`);
  sb.push(`                try`);
  sb.push(`                {`);
  sb.push(`                    await Task.Delay(Frequency, token);`);
  sb.push(`                    if (plc == null || !plc.IsConnected) continue;`);
  sb.push(`                    setProperty(readFunc(plc));`);
  sb.push(`                }`);
  sb.push(`                catch {}`);
  sb.push(`            }`);
  sb.push(`        }`);
  sb.push(``);

  // Helper: Batch Monitor Registers
  sb.push(`        private async Task BatchMonitorRegisters(CancellationToken token, IModbusService plc, int start, int length, Action<ushort[]> mapData)`);
  sb.push(`        {`);
  sb.push(`            while (!token.IsCancellationRequested)`);
  sb.push(`            {`);
  sb.push(`                try`);
  sb.push(`                {`);
  sb.push(`                    await Task.Delay(Frequency, token);`);
  sb.push(`                    if (plc == null || !plc.IsConnected) continue;`);
  sb.push(`                    // 假设 ReadRegistersAsync 返回 string, 需解析`);
  sb.push(`                    string raw = await plc.ReadRegistersAsync(start.ToString(), (ushort)length);`);
  sb.push(`                    ushort[] data = ParseRegisters(raw);`);
  sb.push(`                    // 简单校验长度，防止越界`);
  sb.push(`                    if (data.Length > 0) mapData(data);`);
  sb.push(`                }`);
  sb.push(`                catch {}`);
  sb.push(`            }`);
  sb.push(`        }`);
  sb.push(``);

  // Helper: Batch Monitor Coils
  sb.push(`        private async Task BatchMonitorCoils(CancellationToken token, IModbusService plc, int start, int length, Action<bool[]> mapData)`);
  sb.push(`        {`);
  sb.push(`            while (!token.IsCancellationRequested)`);
  sb.push(`            {`);
  sb.push(`                try`);
  sb.push(`                {`);
  sb.push(`                    await Task.Delay(Frequency, token);`);
  sb.push(`                    if (plc == null || !plc.IsConnected) continue;`);
  sb.push(`                    string raw = await plc.ReadCoilsAsync(start.ToString(), (ushort)length);`);
  sb.push(`                    bool[] data = ParseCoils(raw);`);
  sb.push(`                    if (data.Length > 0) mapData(data);`);
  sb.push(`                }`);
  sb.push(`                catch {}`);
  sb.push(`            }`);
  sb.push(`        }`);
  sb.push(``);

  // Helper: Periodic Log
  sb.push(`        private async Task RunPeriodicLog<T>(CancellationToken token, int period, IModbusService plc, Func<T> createEntity, Action<T, IModbusService> fillEntity, Func<T, Task> saveEntity)`);
  sb.push(`        {`);
  sb.push(`            while (!token.IsCancellationRequested)`);
  sb.push(`            {`);
  sb.push(`                try`);
  sb.push(`                {`);
  sb.push(`                    await Task.Delay(period, token);`);
  sb.push(`                    if (plc == null || !plc.IsConnected) continue;`);
  sb.push(`                    var entity = createEntity();`);
  sb.push(`                    fillEntity(entity, plc);`);
  sb.push(`                    await saveEntity(entity);`);
  sb.push(`                }`);
  sb.push(`                catch (Exception ex) { Console.WriteLine(ex.Message); }`);
  sb.push(`            }`);
  sb.push(`        }`);
  sb.push(``);

  // Helper: Handshake Log
  sb.push(`        private async Task RunHandshakeLog<T>(CancellationToken token, int period, IModbusService plc, string triggerAddr, string returnAddr, Func<T> createEntity, Action<T, IModbusService> fillEntity, Func<T, Task> saveEntity)`);
  sb.push(`        {`);
  sb.push(`            while (!token.IsCancellationRequested)`);
  sb.push(`            {`);
  sb.push(`                try`);
  sb.push(`                {`);
  sb.push(`                    await Task.Delay(period, token);`);
  sb.push(`                    if (plc == null || !plc.IsConnected) continue;`);
  sb.push(``);
  sb.push(`                    string tStr = plc.ReadRegisters(triggerAddr, 1);`);
  sb.push(`                    if (!int.TryParse(tStr, out int tVal) || tVal != 11) continue;`);
  sb.push(``);
  sb.push(`                    var entity = createEntity();`);
  sb.push(`                    fillEntity(entity, plc);`);
  sb.push(`                    await saveEntity(entity);`);
  sb.push(``);
  sb.push(`                    plc.WriteRegisters(returnAddr, 11);`);
  sb.push(``);
  sb.push(`                    Stopwatch sw = Stopwatch.StartNew();`);
  sb.push(`                    while (sw.ElapsedMilliseconds < 5000)`);
  sb.push(`                    {`);
  sb.push(`                        string rStr = plc.ReadRegisters(returnAddr, 1);`);
  sb.push(`                        if (int.TryParse(rStr, out int rInt) && rInt == 0) break;`);
  sb.push(`                        await Task.Delay(100, token);`);
  sb.push(`                    }`);
  sb.push(``);
  sb.push(`                    sw.Restart();`);
  sb.push(`                    while (sw.ElapsedMilliseconds < 5000)`);
  sb.push(`                    {`);
  sb.push(`                        string tEndStr = plc.ReadRegisters(triggerAddr, 1);`);
  sb.push(`                        if (int.TryParse(tEndStr, out int tEndInt) && tEndInt == 0) break;`);
  sb.push(`                        await Task.Delay(100, token);`);
  sb.push(`                    }`);
  sb.push(`                }`);
  sb.push(`                catch (Exception ex) { Console.WriteLine(ex.Message); }`);
  sb.push(`            }`);
  sb.push(`        }`);
  sb.push(`        #endregion`);

  sb.push(``);
  sb.push(`        public void StartAllTasks()`);
  sb.push(`        {`);
  sb.push(`            foreach (var kvp in _tasks) kvp.Value.Cancel();`);
  sb.push(`            _tasks.Clear();`);
  sb.push(``);
  sb.push(`            if (Frequency <= 0) return;`);
  sb.push(``);
  
  // 1. Generate Monitor Tasks (Optimized)
  for (let m = 1; m <= maxModules; m++) {
      sheets.forEach(sheet => {
          const plcName = `Global._${m}${sheet.name}PLCModbus`;
          const allPoints = sheet.points;

          // -- Optimize Coils (Bool) --
          const coilTags = allPoints.filter(p => p.Type.toLowerCase() === 'bool').map(p => ({
              Name: `${p.PropertyName}_M${m}`,
              Address: parseInt(cleanAddr(p.ValueAddress)),
              Length: p.Length || 1,
              Type: 'bool'
          }));
          const coilBlocks = optimizeRequests(coilTags);

          coilBlocks.forEach((block, idx) => {
               const taskName = `${sheet.name}_M${m}_Coils_${idx}`;
               sb.push(`            StartTask("${taskName}", t => BatchMonitorCoils(t, ${plcName}, ${block.StartAddress}, ${block.Length}, data => {`);
               // Map data back to properties
               block.IncludedTags.forEach(tag => {
                   const offset = tag.Address - block.StartAddress;
                   sb.push(`                if (${offset} < data.Length) this.${tag.Name} = data[${offset}];`);
               });
               sb.push(`            }));`);
          });

          // -- Optimize Registers (Short, Int, Float, String) --
          // String is now treated as a sequence of registers
          const regTags = allPoints.filter(p => !['bool'].includes(p.Type.toLowerCase())).map(p => {
              let len = 1;
              const type = p.Type.toLowerCase();
              if (type === 'int' || type === 'float') len = 2;
              if (type === 'string') len = p.Length || 10;
              
              return {
                  Name: `${p.PropertyName}_M${m}`,
                  Address: parseInt(cleanAddr(p.ValueAddress)),
                  Length: len,
                  Type: type
              };
          });
          const regBlocks = optimizeRequests(regTags);

          regBlocks.forEach((block, idx) => {
              const taskName = `${sheet.name}_M${m}_Regs_${idx}`;
              sb.push(`            StartTask("${taskName}", t => BatchMonitorRegisters(t, ${plcName}, ${block.StartAddress}, ${block.Length}, data => {`);
              block.IncludedTags.forEach(tag => {
                  const offset = tag.Address - block.StartAddress;
                  if (tag.Type === 'short') {
                       sb.push(`                this.${tag.Name} = GetShort(data, ${offset});`);
                  } else if (tag.Type === 'int') {
                       sb.push(`                this.${tag.Name} = GetInt(data, ${offset});`);
                  } else if (tag.Type === 'float') {
                       sb.push(`                this.${tag.Name} = GetFloat(data, ${offset});`);
                  } else if (tag.Type === 'string') {
                       sb.push(`                this.${tag.Name} = GetString(data, ${offset}, ${tag.Length});`);
                  }
              });
              sb.push(`            }));`);
          });
      });
  }
  sb.push(``);

  // 2. Generate Logging/Handshake Task Calls (Same as before)
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
          const repoName = `Global.repo_${tableName}`;

          if (period > 0) {
              const taskName = `${sheet.name}_M${m}_Period_${period}`;
              sb.push(`            StartTask("${taskName}", t => RunPeriodicLog(t, ${period}, ${plcName},`);
              sb.push(`                () => new ${tableName} { ModuleNum = ${m} },`);
              sb.push(`                (e, plc) => {`);
              groupPoints.forEach(p => {
                  const addr = cleanAddr(p.ValueAddress);
                  const readMethod = getReadMethod(p.Type);
                  const parseMethod = getParseMethod(p.Type);
                  if(p.Type.toLowerCase() === 'string') {
                      sb.push(`                    e.${p.PropertyName} = plc.${readMethod}("${addr}", ${p.Length});`);
                  } else {
                      sb.push(`                    e.${p.PropertyName} = ${parseMethod}(plc.${readMethod}("${addr}", ${p.Length}));`);
                  }
              });
              sb.push(`                },`);
              sb.push(`                e => ${repoName}.InsertAsync(e)`);
              sb.push(`            ));`);

          } else {
              // Handshake Log
              const groupByTrigger = new Map<string, PointDefinition[]>();
              groupPoints.forEach(p => {
                  const t = cleanAddr(p.TriggerAddress);
                  if(!groupByTrigger.has(t)) groupByTrigger.set(t, []);
                  groupByTrigger.get(t)!.push(p);
              });

              groupByTrigger.forEach((tPoints, triggerAddr) => {
                  const returnAddr = cleanAddr(tPoints[0].ReturnAddress);
                  const taskName = `${sheet.name}_M${m}_Handshake_T${triggerAddr}`;
                  const absPeriod = Math.abs(period);

                  sb.push(`            StartTask("${taskName}", t => RunHandshakeLog(t, ${absPeriod}, ${plcName}, "${triggerAddr}", "${returnAddr}",`);
                  sb.push(`                () => new ${tableName} { ModuleNum = ${m} },`);
                  sb.push(`                (e, plc) => {`);
                  tPoints.forEach(p => {
                      const addr = cleanAddr(p.ValueAddress);
                      const readMethod = getReadMethod(p.Type);
                      const parseMethod = getParseMethod(p.Type);
                      if(p.Type.toLowerCase() === 'string') {
                          sb.push(`                    e.${p.PropertyName} = plc.${readMethod}("${addr}", ${p.Length});`);
                      } else {
                          sb.push(`                    e.${p.PropertyName} = ${parseMethod}(plc.${readMethod}("${addr}", ${p.Length}));`);
                      }
                  });
                  sb.push(`                },`);
                  sb.push(`                e => ${repoName}.InsertAsync(e)`);
                  sb.push(`            ));`);
              });
          }
      });
    });
  }

  sb.push(`        }`);
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
  
  // Header: Frequency + Module Selector + Recorder
  sb.push(`        <StackPanel Orientation="Horizontal" Margin="5">`);
  sb.push(`            <TextBlock Text="Frequency (ms): " VerticalAlignment="Center" Margin="0,0,5,0"/>`);
  sb.push(`            <ComboBox SelectedItem="{Binding Frequency}" Width="80" Margin="0,0,20,0">`);
  sb.push(`                <sys:Int32 xmlns:sys="clr-namespace:System;assembly=mscorlib">0</sys:Int32>`);
  sb.push(`                <sys:Int32 xmlns:sys="clr-namespace:System;assembly=mscorlib">500</sys:Int32>`);
  sb.push(`                <sys:Int32 xmlns:sys="clr-namespace:System;assembly=mscorlib">1000</sys:Int32>`);
  sb.push(`                <sys:Int32 xmlns:sys="clr-namespace:System;assembly=mscorlib">3000</sys:Int32>`);
  sb.push(`            </ComboBox>`);
  
  sb.push(`            <TextBlock Text="Current Module: " VerticalAlignment="Center" Margin="0,0,5,0"/>`);
  sb.push(`            <ComboBox SelectedItem="{Binding CurrentModuleIndex}" Width="80" Margin="0,0,20,0">`);
  for(let i=1; i<=maxModules; i++) {
     sb.push(`                <sys:Int32 xmlns:sys="clr-namespace:System;assembly=mscorlib">${i}</sys:Int32>`);
  }
  sb.push(`            </ComboBox>`);

  sb.push(`            <TextBlock Text="Record: " VerticalAlignment="Center" Margin="0,0,5,0"/>`);
  sb.push(`            <ComboBox SelectedItem="{Binding RecordState}" Width="100">`);
  sb.push(`                <sys:String xmlns:sys="clr-namespace:System;assembly=mscorlib">Stopped</sys:String>`);
  sb.push(`                <sys:String xmlns:sys="clr-namespace:System;assembly=mscorlib">Recording</sys:String>`);
  sb.push(`            </ComboBox>`);

  sb.push(`        </StackPanel>`);

  sb.push(`        <ScrollViewer Grid.Row="1">`);
  sb.push(`            <StackPanel>`);
  
  // Single View: Displays CurrentModule Properties
  sb.push(`                <GroupBox Header="Current Module View" Margin="5">`);
  sb.push(`                    <StackPanel>`);
  
  sheets.forEach(sheet => {
    // Note: Header removed DeviceInfo since it changes per module
    sb.push(`                        <Expander Header="${sheet.name}" Margin="2" IsExpanded="True">`);
    sb.push(`                            <StackPanel>`);
    
    sheet.points.forEach(p => {
      sb.push(`                                <Grid Margin="2">`);
      sb.push(`                                    <Grid.ColumnDefinitions>`);
      sb.push(`                                        <ColumnDefinition Width="200"/>`);
      sb.push(`                                        <ColumnDefinition Width="*"/>`);
      sb.push(`                                    </Grid.ColumnDefinitions>`);
      sb.push(`                                    <TextBlock Text="${p.KeyName} (${p.PropertyName}):" VerticalAlignment="Center"/>`);
      // Bind to CurrentX
      sb.push(`                                    <TextBox Grid.Column="1" Text="{Binding Current${p.PropertyName}}" IsReadOnly="True"/>`);
      sb.push(`                                </Grid>`);
    });

    sb.push(`                            </StackPanel>`);
    sb.push(`                        </Expander>`);
  });

  sb.push(`                    </StackPanel>`);
  sb.push(`                </GroupBox>`);

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