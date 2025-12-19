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
}

export type ByteOrder = 'ABCD' | 'CDAB' | 'BADC' | 'DCBA';
export type StringByteOrder = 'ABCD' | 'BADC'; // Only allow byte swaps for strings, not word swaps

export interface AppState {
  maxModules: number;
  byteOrder: ByteOrder;
  stringByteOrder: StringByteOrder;
  maxGap: number;        // Max address gap to merge
  maxBatchSize: number;  // Max registers per request
  sheets: SheetData[];
  devices: DeviceConfig[];
  isProcessing: boolean;
  error: string | null;
  logs: string[];
  showReadme: boolean;   // UI state for Readme modal
}