import { SheetData, PointDefinition, ByteOrder, StringByteOrder } from '../../types';
import { cleanAddr, getCSharpType, getParseMethod, getReadMethod } from '../utils';
import { optimizeRequests } from '../optimizer';

export const generatePLCPointProperty = (
    sheets: SheetData[], 
    maxModules: number, 
    byteOrder: ByteOrder, 
    stringByteOrder: StringByteOrder,
    maxGap: number,
    maxBatchSize: number
): string => {
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
  sb.push(`using CommunityToolkit.Mvvm.ComponentModel;`);
  sb.push(`using System.Windows;`);
  sb.push(``);
  sb.push(`namespace Core`);
  sb.push(`{`);
  sb.push(`    // Numeric Byte Order: ${byteOrder}`);
  sb.push(`    // String Byte Order:  ${stringByteOrder}`);
  sb.push(`    // Optimization: Gap=${maxGap}, Batch=${maxBatchSize}`);
  sb.push(`    public partial class PLCPointProperty : ObservableObject`);
  sb.push(`    {`);
  sb.push(`        public static PLCPointProperty Instance { get; } = new PLCPointProperty();`);
  sb.push(``);
  sb.push(`        private PLCPointProperty()`);
  sb.push(`        {`);
  sb.push(`            InitializePropertyMap();`);
  sb.push(`            UpdateAllPropertyCaches();`);
  sb.push(`            _ = Task.Run(ProcessLogQueue);`);
  sb.push(`        }`);
  sb.push(``);
  sb.push(`        private int _frequency = 1000;`);
  sb.push(`        public int Frequency`);
  sb.push(`        {`);
  sb.push(`            get => _frequency;`);
  sb.push(`            set`);
  sb.push(`            {`);
  sb.push(`                if (SetProperty(ref _frequency, value)) StartAllTasks();`);
  sb.push(`            }`);
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
  sb.push(`                if (SetProperty(ref _recordState, value)) IsRecording = (value == "Recording");`);
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
  sheets.forEach(sheet => {
      sheet.points.forEach(p => {
          sb.push(`            _nameMapping.Add(nameof(Current${p.PropertyName}), "${p.PropertyName}_M");`);
      });
  });
  sb.push(``);
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
  sb.push(`                if (SetProperty(ref _currentModuleIndex, value))`);
  sb.push(`                {`);
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
  sb.push(`            foreach (var key in _nameMapping.Keys) OnPropertyChanged(key);`);
  sb.push(`        }`);
  sb.push(``);
  sb.push(`        // ================== 3. 通用取值/赋值辅助方法 ==================`);
  sb.push(`        private T GetDynamicValue<T>(string uiPropertyName)`);
  sb.push(`        {`);
  sb.push(`            if (_reflectionCache.TryGetValue(uiPropertyName, out PropertyInfo info) && info != null)`);
  sb.push(`            {`);
  sb.push(`                var val = info.GetValue(this);`);
  sb.push(`                if (val is T t) return t;`);
  sb.push(`                try { return (T)Convert.ChangeType(val, typeof(T)); } catch { return default(T); }`);
  sb.push(`            }`);
  sb.push(`            return default(T);`);
  sb.push(`        }

        private void SetDynamicValue(string uiPropertyName, object value)
        {
            if (_reflectionCache.TryGetValue(uiPropertyName, out PropertyInfo info) && info != null)
            {
                info.SetValue(this, value);
            }
        }`);
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
  
  sb.push(`        protected override void OnPropertyChanged(PropertyChangedEventArgs e)`);
  sb.push(`        {`);
  sb.push(`            base.OnPropertyChanged(e);`);
  sb.push(`            string name = e.PropertyName;`);
  sb.push(`            if (IsRecording && _logMetadata.TryGetValue(name, out var meta))`);
  sb.push(`            {`);
  sb.push(`                try {`);
  sb.push(`                    var val = meta.Info.GetValue(this);`);
  sb.push(`                    string valStr = val?.ToString();`);
  sb.push(`                    if (val is System.Collections.IEnumerable enumerable && !(val is string))`);
  sb.push(`                        valStr = string.Join(";", System.Linq.Enumerable.Cast<object>(enumerable));`);
  sb.push(`                    string safeKey = (meta.KeyName != null && meta.KeyName.Contains(",")) ? $"\\"{meta.KeyName}\\"" : meta.KeyName;`);
  sb.push(`                    _logQueue.Enqueue($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff},{safeKey},{meta.Address},{meta.Type},{valStr}");`);
  sb.push(`                } catch {}`);
  sb.push(`            }`);
  sb.push(``);
  sb.push(`            if (!name.EndsWith($"_M{CurrentModuleIndex}")) return;`);
  sb.push(`            var targetEntry = _reflectionCache.FirstOrDefault(x => x.Value != null && x.Value.Name == name);`);
  sb.push(`            if (targetEntry.Key != null) OnPropertyChanged(targetEntry.Key);`);
  sb.push(`        }`);
  sb.push(``);

  // ================== PARSING LOGIC GENERATION (Endiannness) ==================
  sb.push(`        #region 辅助解析方法 (Numeric: ${byteOrder}, String: ${stringByteOrder})`);
  sb.push(`        private bool ParseBool(string val) => bool.TryParse(val, out var b) ? b : false;`);
  sb.push(``);
  
  // -- GET SHORT --
  sb.push(`        private short GetShort(ushort[] data, int offset)`);
  sb.push(`        {`);
  sb.push(`            if (offset >= data.Length) return 0;`);
  if (byteOrder === 'BADC' || byteOrder === 'DCBA') {
      // Byte Swap needed within the register
      sb.push(`            ushort v = data[offset];`);
      sb.push(`            return (short)((v >> 8) | (v << 8));`);
  } else {
      // ABCD or CDAB (No byte swap within register)
      sb.push(`            return (short)data[offset];`);
  }
  sb.push(`        }`);
  sb.push(``);

  // -- GET INT (32-bit) --
  sb.push(`        private int GetInt(ushort[] data, int offset)`);
  sb.push(`        {`);
  sb.push(`            if (offset + 1 >= data.Length) return 0;`);
  // Map words based on endianness
  if (byteOrder === 'ABCD') {
      sb.push(`            // ABCD: Big Endian`);
      sb.push(`            return (int)((data[offset] << 16) | data[offset + 1]);`);
  } else if (byteOrder === 'CDAB') {
      sb.push(`            // CDAB: Little Endian Word Swap`);
      sb.push(`            return (int)((data[offset + 1] << 16) | data[offset]);`);
  } else if (byteOrder === 'BADC') {
      sb.push(`            // BADC: Big Endian Byte Swap`);
      sb.push(`            ushort w1 = (ushort)((data[offset] >> 8) | (data[offset] << 8));`);
      sb.push(`            ushort w2 = (ushort)((data[offset + 1] >> 8) | (data[offset + 1] << 8));`);
      sb.push(`            return (int)((w1 << 16) | w2);`);
  } else { // DCBA
      sb.push(`            // DCBA: Little Endian`);
      sb.push(`            ushort w1 = (ushort)((data[offset] >> 8) | (data[offset] << 8));`);
      sb.push(`            ushort w2 = (ushort)((data[offset + 1] >> 8) | (data[offset + 1] << 8));`);
      sb.push(`            return (int)((w2 << 16) | w1);`);
  }
  sb.push(`        }`);
  sb.push(``);

  // -- GET FLOAT --
  sb.push(`        private float GetFloat(ushort[] data, int offset)`);
  sb.push(`        {`);
  sb.push(`            if (offset + 1 >= data.Length) return 0f;`);
  
  if (byteOrder === 'ABCD') {
      sb.push(`            byte[] bytes = new byte[4];`);
      sb.push(`            bytes[0] = (byte)(data[offset + 1] & 0xFF);`);
      sb.push(`            bytes[1] = (byte)(data[offset + 1] >> 8);`);
      sb.push(`            bytes[2] = (byte)(data[offset] & 0xFF);`);
      sb.push(`            bytes[3] = (byte)(data[offset] >> 8);`);
  } else if (byteOrder === 'CDAB') {
      sb.push(`            byte[] bytes = new byte[4];`);
      sb.push(`            bytes[0] = (byte)(data[offset] & 0xFF);`);
      sb.push(`            bytes[1] = (byte)(data[offset] >> 8);`);
      sb.push(`            bytes[2] = (byte)(data[offset + 1] & 0xFF);`);
      sb.push(`            bytes[3] = (byte)(data[offset + 1] >> 8);`);
  } else if (byteOrder === 'BADC') {
      sb.push(`            byte[] bytes = new byte[4];`);
      sb.push(`            bytes[0] = (byte)(data[offset + 1] >> 8);`);
      sb.push(`            bytes[1] = (byte)(data[offset + 1] & 0xFF);`);
      sb.push(`            bytes[2] = (byte)(data[offset] >> 8);`);
      sb.push(`            bytes[3] = (byte)(data[offset + 1] & 0xFF);`);
  } else { // DCBA
      sb.push(`            byte[] bytes = new byte[4];`);
      sb.push(`            bytes[0] = (byte)(data[offset] >> 8);`);
      sb.push(`            bytes[1] = (byte)(data[offset] & 0xFF);`);
      sb.push(`            bytes[2] = (byte)(data[offset + 1] >> 8);`);
      sb.push(`            bytes[3] = (byte)(data[offset + 1] & 0xFF);`);
  }
  
  sb.push(`            return BitConverter.ToSingle(bytes, 0);`);
  sb.push(`        }`);
  
  // -- STRING --
  sb.push(`        private string GetString(ushort[] data, int offset, int length)`);
  sb.push(`        {`);
  sb.push(`            if (offset + length > data.Length) return string.Empty;`);
  sb.push(`            var bytes = new List<byte>();`);
  sb.push(`            for (int i = 0; i < length; i++)`);
  sb.push(`            {`);
  sb.push(`                ushort val = data[offset + i];`);
  
  if (stringByteOrder === 'BADC') {
      sb.push(`                bytes.Add((byte)(val & 0xFF));`);
      sb.push(`                bytes.Add((byte)(val >> 8));`);
  } else {
      sb.push(`                bytes.Add((byte)(val >> 8));`);
      sb.push(`                bytes.Add((byte)(val & 0xFF));`);
  }
  
  sb.push(`            }`);
  sb.push(`            return Encoding.ASCII.GetString(bytes.ToArray()).Trim('\\0');`);
  sb.push(`        }`);
  sb.push(``);

  // -- ARRAY GETTERS --
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
  sb.push(``);

  // -- PARSE WRAPPERS (Used by RunPeriodicLog) --
  sb.push(`        private ushort[] ParseRegisters(string raw)`);
  sb.push(`        {`);
  sb.push(`            if (string.IsNullOrWhiteSpace(raw)) return new ushort[0];`);
  sb.push(`            var parts = raw.Split(new[] { ',', ' ', '[', ']' }, StringSplitOptions.RemoveEmptyEntries);`);
  sb.push(`            var list = new List<ushort>();`);
  sb.push(`            foreach (var p in parts) if (ushort.TryParse(p, out var v)) list.Add(v);`);
  sb.push(`            return list.ToArray();`);
  sb.push(`        }`);
  sb.push(``);
  // Wrappers to convert Raw Registers -> Type using internal endianness logic
  sb.push(`        private short ParseShort(string rawRegs) => GetShort(ParseRegisters(rawRegs), 0);`);
  sb.push(`        private int ParseInt(string rawRegs) => GetInt(ParseRegisters(rawRegs), 0);`);
  sb.push(`        private float ParseFloat(string rawRegs) => GetFloat(ParseRegisters(rawRegs), 0);`);
  sb.push(``);
  sb.push(`        private short[] ParseShortArray(string rawRegs, int len) => GetShortArray(ParseRegisters(rawRegs), 0, len);`);
  sb.push(`        private int[] ParseIntArray(string rawRegs, int len) => GetIntArray(ParseRegisters(rawRegs), 0, len);`);
  sb.push(`        private float[] ParseFloatArray(string rawRegs, int len) => GetFloatArray(ParseRegisters(rawRegs), 0, len);`);
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
  sb.push(`        #endregion`);
  sb.push(``);

  sb.push(`        #region 底层属性 (Mx)`);
  for (let m = 1; m <= maxModules; m++) {
    sheets.forEach(sheet => {
      sheet.points.forEach(p => {
        const propName = `${p.PropertyName}_M${m}`;
        const fieldName = `_${propName}`;
        const csType = getCSharpType(p.Type);
        
        let init = "";
        if (p.Type.trim().toLowerCase().endsWith("[]")) {
             const arrLen = p.Length || 5; 
             const baseType = getCSharpType(p.Type).replace("[]","");
             init = ` = new ${baseType}[${arrLen}]`;
        }
        
        sb.push(`        private ${csType} ${fieldName}${init};`);
        sb.push(`        public ${csType} ${propName}`);
        sb.push(`        {`);
        sb.push(`            get => ${fieldName};`);
        sb.push(`            set => SetProperty(ref ${fieldName}, value);`);
        sb.push(`        }`);
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
  
  // Batch Monitor Registers
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

  // Periodic Log: Uses Action<T, IModbusService>
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

  // New: RunChangeLog for Period < 1.0 (Store only on change)
  sb.push(`        private async Task RunChangeLog<T>(CancellationToken token, int period, IModbusService plc, Func<T> createEntity, Action<T, IModbusService> fillEntity, Func<T, Task> saveEntity, Func<T, T, bool> isDifferent)`);
  sb.push(`        {`);
  sb.push(`            T lastEntity = default;`);
  sb.push(`            while (!token.IsCancellationRequested)`);
  sb.push(`            {`);
  sb.push(`                try`);
  sb.push(`                {`);
  sb.push(`                    await Task.Delay(period, token);`);
  sb.push(`                    if (plc == null || !plc.IsConnected) continue;`);
  sb.push(`                    var entity = createEntity();`);
  sb.push(`                    fillEntity(entity, plc);`);
  sb.push(``);
  sb.push(`                    if (lastEntity == null || isDifferent(entity, lastEntity))`);
  sb.push(`                    {`);
  sb.push(`                        await saveEntity(entity);`);
  sb.push(`                        lastEntity = entity;`);
  sb.push(`                    }`);
  sb.push(`                }`);
  sb.push(`                catch (Exception ex) { Console.WriteLine(ex.Message); }`);
  sb.push(`            }`);
  sb.push(`        }`);
  sb.push(``);

  // Handshake Log
  sb.push(`        private async Task RunHandshakeLog<T>(CancellationToken token, int period, IModbusService plc, string triggerAddr, string returnAddr, Func<T> createEntity, Action<T, IModbusService> fillEntity, Func<T, Task> saveEntity)`);
  sb.push(`        {`);
  sb.push(`            while (!token.IsCancellationRequested)`);
  sb.push(`            {`);
  sb.push(`                try`);
  sb.push(`                {`);
  sb.push(`                    await Task.Delay(period, token);`);
  sb.push(`                    if (plc == null || !plc.IsConnected) continue;`);
  sb.push(`                    // Use ReadRegisters for trigger check to avoid byte order issues`);
  sb.push(`                    string tStr = plc.ReadRegisters(triggerAddr, 1);`);
  sb.push(`                    if (ParseInt(tStr) != 11) continue;`); // trigger value 11
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
  sb.push(`                        if (ParseInt(rStr) == 0) break;`);
  sb.push(`                        await Task.Delay(100, token);`);
  sb.push(`                    }`);
  sb.push(``);
  sb.push(`                    sw.Restart();`);
  sb.push(`                    while (sw.ElapsedMilliseconds < 5000)`);
  sb.push(`                    {`);
  sb.push(`                        string tEndStr = plc.ReadRegisters(triggerAddr, 1);`);
  sb.push(`                        if (ParseInt(tEndStr) == 0) break;`);
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
  
  // 1. Generate Monitor Tasks (Existing Logic)
  for (let m = 1; m <= maxModules; m++) {
      sheets.forEach(sheet => {
          const plcName = `Global._${m}${sheet.name}PLCModbus`;
          const allPoints = sheet.points;
          const coilTags = allPoints.filter(p => p.Type.toLowerCase() === 'bool').map(p => ({
              Name: `${p.PropertyName}_M${m}`,
              Address: parseInt(cleanAddr(p.ValueAddress)),
              Length: p.Length || 1,
              Type: 'bool',
              ArrayLen: 1
          }));
          const coilBlocks = optimizeRequests(coilTags, maxGap, maxBatchSize);

          coilBlocks.forEach((block, idx) => {
               const taskName = `${sheet.name}_M${m}_Coils_${idx}`;
               sb.push(`            StartTask("${taskName}", t => BatchMonitorCoils(t, ${plcName}, ${block.StartAddress}, ${block.Length}, data => {`);
               block.IncludedTags.forEach(tag => {
                   const offset = tag.Address - block.StartAddress;
                   sb.push(`                if (${offset} < data.Length) this.${tag.Name} = data[${offset}];`);
               });
               sb.push(`            }));`);
          });

          const regTags = allPoints.filter(p => !['bool'].includes(p.Type.toLowerCase())).map(p => {
              const lowerType = p.Type.toLowerCase();
              let regCount = 1;
              let arrayLen = p.Length || 1; 

              if (lowerType === 'int' || lowerType === 'float') regCount = 2;
              if (lowerType === 'string') regCount = p.Length || 10;
              if (lowerType.endsWith('[]')) {
                  arrayLen = p.Length || 5; 
                  const baseType = lowerType.replace('[]', '');
                  let elementRegs = 1;
                  if (baseType === 'int' || baseType === 'float') elementRegs = 2;
                  regCount = arrayLen * elementRegs;
              }
              return {
                  Name: `${p.PropertyName}_M${m}`,
                  Address: parseInt(cleanAddr(p.ValueAddress)),
                  Length: regCount,
                  Type: lowerType,
                  ArrayLen: arrayLen
              };
          });

          const regBlocks = optimizeRequests(regTags, maxGap, maxBatchSize);

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
                  } else if (tag.Type.endsWith('[]')) {
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
          const safePeriod = Math.abs(period).toString().replace('.', '_');
          const tableName = `${sheet.name}_PeriodAbs_${safePeriod}`;
          const repoName = `Global.repo_${tableName}`;

          if (period > 0) {
              // CASE: High Frequency Change Log (0 < P < 1)
              if (period < 1) {
                  const frequency = Math.floor(period * 1000);
                  const taskName = `${sheet.name}_M${m}_ChangeLog_${safePeriod}`;
                  
                  // Generate Comparison Logic
                  const compareLogic = groupPoints.map(p => {
                     if (p.Type.endsWith('[]')) {
                         const baseType = getCSharpType(p.Type).replace('[]', '');
                         return `!Enumerable.SequenceEqual(a.${p.PropertyName} ?? new ${baseType}[0], b.${p.PropertyName} ?? new ${baseType}[0])`;
                     }
                     return `a.${p.PropertyName} != b.${p.PropertyName}`;
                  }).join(' || ');

                  sb.push(`            StartTask("${taskName}", t => RunChangeLog(t, ${frequency}, ${plcName},`);
                  sb.push(`                () => new ${tableName} { ModuleNum = ${m} },`);
                  sb.push(`                (e, plc) => {`);
                  generateReadLogic(sb, groupPoints); // Helper function extracted below or inline logic
                  sb.push(`                },`);
                  sb.push(`                e => ${repoName}.InsertAsync(e),`);
                  sb.push(`                (a, b) => ${compareLogic}`); // isDifferent lambda
                  sb.push(`            ));`);
              } 
              // CASE: Standard Periodic Log (P >= 1)
              else {
                  const taskName = `${sheet.name}_M${m}_Period_${safePeriod}`;
                  sb.push(`            StartTask("${taskName}", t => RunPeriodicLog(t, ${period}, ${plcName},`);
                  sb.push(`                () => new ${tableName} { ModuleNum = ${m} },`);
                  sb.push(`                (e, plc) => {`);
                  generateReadLogic(sb, groupPoints);
                  sb.push(`                },`);
                  sb.push(`                e => ${repoName}.InsertAsync(e)`);
                  sb.push(`            ));`);
              }

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
                  generateReadLogic(sb, tPoints);
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

// Helper function to deduplicate reading logic in the template
function generateReadLogic(sb: string[], points: PointDefinition[]) {
    points.forEach(p => {
        const addr = cleanAddr(p.ValueAddress);
        let readLen = 1;
        const t = p.Type.toLowerCase();
        if (t === 'int' || t === 'float') readLen = 2;
        else if (t === 'string') readLen = p.Length || 10;
        else if (t.endsWith('[]')) {
            const count = p.Length || 1;
            if (t.includes('int') || t.includes('float')) readLen = count * 2;
            else readLen = count;
        }

        if (t === 'bool') {
            sb.push(`                    e.${p.PropertyName} = ParseBool(plc.ReadCoils("${addr}", 1));`);
        } else if (t.endsWith('[]')) {
             let parseCall = "ParseShortArray";
             if (t.includes('int')) parseCall = "ParseIntArray";
             if (t.includes('float')) parseCall = "ParseFloatArray";
             sb.push(`                    e.${p.PropertyName} = ${parseCall}(plc.ReadRegisters("${addr}", ${readLen}), ${p.Length});`);
        } else {
            let parseCall = "ParseShort";
            if (t === 'int') parseCall = "ParseInt";
            if (t === 'float') parseCall = "ParseFloat";
            if (t === 'string') {
                 sb.push(`                    e.${p.PropertyName} = GetString(ParseRegisters(plc.ReadRegisters("${addr}", ${readLen})), 0, ${readLen});`);
            } else {
                 sb.push(`                    e.${p.PropertyName} = ${parseCall}(plc.ReadRegisters("${addr}", ${readLen}));`);
            }
        }
    });
}