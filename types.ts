export interface DeviceConfig {
  DeviceName: string;
  IP: string;
  Port: number;
}

export interface PointDefinition {
  PropertyName: string;
  KeyName: string;
  ValueAddress: string;
  Type: 'bool' | 'short' | 'float' | 'string' | 'int';
  Length: number;
  TriggerAddress?: string; // Optional, can be empty
  ReturnAddress?: string; // Optional, can be empty
  Period: number;
}

export interface SheetData {
  name: string;
  points: PointDefinition[];
  deviceInfo?: DeviceConfig; // Linked from IP_Port sheet
}

export interface AppState {
  maxModules: number;
  sheets: SheetData[];
  devices: DeviceConfig[];
  isProcessing: boolean;
  error: string | null;
  logs: string[];
}