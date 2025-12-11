import { SheetData, PointDefinition } from '../../types';
import { cleanAddr, getCSharpType, getParseMethod, getReadMethod } from '../utils';
import { optimizeRequests } from '../optimizer';

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
  sb.push(`using System.Windows.Data;`);
  sb.push(`using System.Globalization;`);
  sb.push(`using PropertyChanged;`); // For Fody
  sb.push(`using System.Windows;`);
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
  sb.push(`        private class LogMetadata { public string KeyName; public string Address; public string Type; public PropertyInfo Info; }`);
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
               const safeKeyName = p.KeyName ? p.KeyName.replace(/"/g, '\\"') : "";
               const address = p.ValueAddress || "";
               sb.push(`            _logMetadata["${propName}"] = new LogMetadata { KeyName = "${safeKeyName}", Address = "${address}", Type = "${p.Type}", Info = this.GetType().GetProperty("${propName}") };`);
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
  sb.push(`                if (val is T t) return t;`); // Direct cast check to support arrays
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
  // Handle Arrays in Log
  sb.push(`                    string valStr = val?.ToString();`);
  sb.push(`                    if (val is System.Collections.IEnumerable enumerable && !(val is string))`);
  sb.push(`                    {`);
  sb.push(`                         valStr = string.Join(";", System.Linq.Enumerable.Cast<object>(enumerable));`);
  sb.push(`                    }`);
  sb.push(`                    string safeKey = (meta.KeyName != null && meta.KeyName.Contains(",")) ? $"\\"{meta.KeyName}\\"" : meta.KeyName;`);
  sb.push(`                    _logQueue.Enqueue($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff},{safeKey},{meta.Address},{meta.Type},{valStr}");`);
  sb.push(`                } catch {}`);
  sb.push(`            }`);
  sb.push(``);
  
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
  // Array Parsing from String (CSV style return from ReadRegisters)
  sb.push(`        private short[] ParseShortArray(string raw)`);
  sb.push(`        {`);
  sb.push(`            var us = ParseRegisters(raw);`);
  sb.push(`            var res = new short[us.Length];`);
  sb.push(`            for(int i=0; i<us.Length; i++) res[i] = (short)us[i];`);
  sb.push(`            return res;`);
  sb.push(`        }`);
  sb.push(`        private int[] ParseIntArray(string raw)`);
  sb.push(`        {`);
  sb.push(`            var us = ParseRegisters(raw);`);
  sb.push(`            var res = new int[us.Length / 2];`);
  sb.push(`            for(int i=0; i<res.Length; i++) res[i] = GetInt(us, i*2);`);
  sb.push(`            return res;`);
  sb.push(`        }`);
  sb.push(`        private float[] ParseFloatArray(string raw)`);
  sb.push(`        {`);
  sb.push(`            var us = ParseRegisters(raw);`);
  sb.push(`            var res = new float[us.Length / 2];`);
  sb.push(`            for(int i=0; i<res.Length; i++) res[i] = GetFloat(us, i*2);`);
  sb.push(`            return res;`);
  sb.push(`        }`);
  sb.push(``);
  
  sb.push(`        // 批量读取辅助`);
  sb.push(`        private ushort[] ParseRegisters(string raw)`);
  sb.push(`        {`);
  sb.push(`            if (string.IsNullOrWhiteSpace(raw)) return new ushort[0];`);
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
  sb.push(`            byte[] bytes = new byte[4];`);
  sb.push(`            byte[] high = BitConverter.GetBytes(data[offset]);`);
  sb.push(`            byte[] low = BitConverter.GetBytes(data[offset + 1]);`);
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
  sb.push(``);
  // Array Extraction Methods
  sb.push(`        private short[] GetShortArray(ushort[] data, int offset, int len)`);
  sb.push(`        {`);
  sb.push(`            var res = new short[len];`);
  sb.push(`            for(int i=0; i<len; i++) res[i] = GetShort(data, offset + i);`);
  sb.push(`            return res;`);
  sb.push(`        }`);
  sb.push(`        private int[] GetIntArray(ushort[] data, int offset, int len)`);
  sb.push(`        {`);
  sb.push(`            var res = new int[len];`);
  sb.push(`            for(int i=0; i<len; i++) res[i] = GetInt(data, offset + i * 2);`);
  sb.push(`            return res;`);
  sb.push(`        }`);
  sb.push(`        private float[] GetFloatArray(ushort[] data, int offset, int len)`);
  sb.push(`        {`);
  sb.push(`            var res = new float[len];`);
  sb.push(`            for(int i=0; i<len; i++) res[i] = GetFloat(data, offset + i * 2);`);
  sb.push(`            return res;`);
  sb.push(`        }`);
  sb.push(`        #endregion`);
  sb.push(``);

  sb.push(`        #region 底层属性 (Mx) - Fody Auto Properties`);
  for (let m = 1; m <= maxModules; m++) {
    sheets.forEach(sheet => {
      sheet.points.forEach(p => {
        const propName = `${p.PropertyName}_M${m}`;
        const csType = getCSharpType(p.Type);
        
        // --- Fix: Initialize arrays ---
        let init = "";
        if (p.Type.trim().toLowerCase().endsWith("[]")) {
             const arrLen = p.Length || 5; 
             const baseType = getCSharpType(p.Type).replace("[]","");
             init = ` = new ${baseType}[${arrLen}];`;
        }
        
        sb.push(`        public ${csType} ${propName} { get; set; }${init}`);
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
  sb.push(`                    string raw = await plc.ReadRegistersAsync(start.ToString(), (ushort)length);`);
  sb.push(`                    ushort[] data = ParseRegisters(raw);`);
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
              Type: 'bool',
              ArrayLen: 1
          }));
          const coilBlocks = optimizeRequests(coilTags);

          coilBlocks.forEach((block, idx) => {
               const taskName = `${sheet.name}_M${m}_Coils_${idx}`;
               sb.push(`            StartTask("${taskName}", t => BatchMonitorCoils(t, ${plcName}, ${block.StartAddress}, ${block.Length}, data => {`);
               block.IncludedTags.forEach(tag => {
                   const offset = tag.Address - block.StartAddress;
                   sb.push(`                if (${offset} < data.Length) this.${tag.Name} = data[${offset}];`);
               });
               sb.push(`            }));`);
          });

          // -- Optimize Registers (Short, Int, Float, String, Arrays) --
          const regTags = allPoints.filter(p => !['bool'].includes(p.Type.toLowerCase())).map(p => {
              const lowerType = p.Type.toLowerCase();
              let regCount = 1;
              let arrayLen = p.Length || 1; 

              if (lowerType === 'int' || lowerType === 'float') regCount = 2;
              if (lowerType === 'string') regCount = p.Length || 10;
              
              // Handle Arrays
              if (lowerType.endsWith('[]')) {
                  arrayLen = p.Length || 5; // Default array length if not specified
                  const baseType = lowerType.replace('[]', '');
                  let elementRegs = 1;
                  if (baseType === 'int' || baseType === 'float') elementRegs = 2;
                  regCount = arrayLen * elementRegs;
              }

              return {
                  Name: `${p.PropertyName}_M${m}`,
                  Address: parseInt(cleanAddr(p.ValueAddress)),
                  Length: regCount, // Total registers to read
                  Type: lowerType,
                  ArrayLen: arrayLen
              };
          });

          const regBlocks = optimizeRequests(regTags);

          regBlocks.forEach((block, idx) => {
              const taskName = `${sheet.name}_M${m}_Regs_${idx}`;
              sb.push(`            StartTask("${taskName}", t => BatchMonitorRegisters(t, ${plcName}, ${block.StartAddress}, ${block.Length}, data => {`);
              block.IncludedTags.forEach(tag => {
                  const offset = tag.Address - block.StartAddress;
                  // Handle different types
                  if (tag.Type === 'short') {
                       sb.push(`                this.${tag.Name} = GetShort(data, ${offset});`);
                  } else if (tag.Type === 'int') {
                       sb.push(`                this.${tag.Name} = GetInt(data, ${offset});`);
                  } else if (tag.Type === 'float') {
                       sb.push(`                this.${tag.Name} = GetFloat(data, ${offset});`);
                  } else if (tag.Type === 'string') {
                       sb.push(`                this.${tag.Name} = GetString(data, ${offset}, ${tag.Length});`);
                  } else if (tag.Type.endsWith('[]')) {
                       // Array extraction
                       if (tag.Type.includes('short')) {
                           sb.push(`                this.${tag.Name} = GetShortArray(data, ${offset}, ${tag.ArrayLen});`);
                       } else if (tag.Type.includes('int')) {
                           sb.push(`                this.${tag.Name} = GetIntArray(data, ${offset}, ${tag.ArrayLen});`);
                       } else if (tag.Type.includes('float')) {
                           sb.push(`                this.${tag.Name} = GetFloatArray(data, ${offset}, ${tag.ArrayLen});`);
                       }
                  }
              });
              sb.push(`            }));`);
          });
      });
  }
  sb.push(``);

  // 2. Generate Logging/Handshake Task Calls
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
                  let readLen = p.Length;
                  // Calculate correct read length for arrays (Modbus registers)
                  if (p.Type.toLowerCase().endsWith('[]')) {
                      const base = p.Type.toLowerCase().replace('[]', '');
                      const count = p.Length || 1;
                      readLen = (base === 'int' || base === 'float') ? count * 2 : count;
                  }

                  if(p.Type.toLowerCase() === 'string') {
                      sb.push(`                    e.${p.PropertyName} = plc.${readMethod}("${addr}", ${p.Length});`);
                  } else if (p.Type.toLowerCase().endsWith('[]')) {
                      // Array handling: ParseXArray(plc.ReadRegisters(...))
                       sb.push(`                    e.${p.PropertyName} = ${parseMethod}(plc.${readMethod}("${addr}", ${readLen}));`);
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
                      let readLen = p.Length;
                      if (p.Type.toLowerCase().endsWith('[]')) {
                          const base = p.Type.toLowerCase().replace('[]', '');
                          const count = p.Length || 1;
                          readLen = (base === 'int' || base === 'float') ? count * 2 : count;
                      }

                      if(p.Type.toLowerCase() === 'string') {
                          sb.push(`                    e.${p.PropertyName} = plc.${readMethod}("${addr}", ${p.Length});`);
                      } else if (p.Type.toLowerCase().endsWith('[]')) {
                          sb.push(`                    e.${p.PropertyName} = ${parseMethod}(plc.${readMethod}("${addr}", ${readLen}));`);
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
  
  // --- Array Converter Class ---
  sb.push(``);
  sb.push(`    public class ArrayToStringConverter : IValueConverter`);
  sb.push(`    {`);
  sb.push(`        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)`);
  sb.push(`        {`);
  sb.push(`            if (value is System.Collections.IEnumerable enumerable && !(value is string))`);
  sb.push(`            {`);
  sb.push(`                var list = new List<string>();`);
  sb.push(`                foreach (var item in enumerable) list.Add(item?.ToString() ?? "");`);
  sb.push(`                return string.Join(", ", list);`);
  sb.push(`            }`);
  sb.push(`            return value?.ToString();`);
  sb.push(`        }`);
  sb.push(``);
  sb.push(`        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)`);
  sb.push(`        {`);
  sb.push(`            return DependencyProperty.UnsetValue;`);
  sb.push(`        }`);
  sb.push(`    }`);
  
  sb.push(`}`);
  return sb.join('\n');
};
