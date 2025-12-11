// Helper to clean address (remove D prefix)
export const cleanAddr = (addr: string | number | undefined): string => {
  if (addr === undefined || addr === null || addr === '') return "0";
  return String(addr).trim().toUpperCase().replace('D', '');
};

export const isArrayType = (type: string): boolean => type.trim().toUpperCase().endsWith('[]');

// Helper to get C# type
export const getCSharpType = (type: string): string => {
  const t = type.trim().toLowerCase();
  if (t.endsWith('[]')) {
    const base = t.substring(0, t.length - 2);
    switch (base) {
      case 'bool': return 'bool[]';
      case 'short': return 'short[]';
      case 'int': return 'int[]';
      case 'float': return 'float[]';
      case 'string': return 'string[]'; // Uncommon but possible
      default: return 'string[]';
    }
  }

  switch (t) {
    case 'bool': return 'bool';
    case 'short': return 'short';
    case 'int': return 'int';
    case 'float': return 'float';
    case 'string': return 'string';
    default: return 'string';
  }
};

// Helper to get parse method name
export const getParseMethod = (type: string): string => {
  const t = type.trim().toLowerCase();
  if (t.endsWith('[]')) {
      const base = t.substring(0, t.length - 2);
      switch (base) {
          case 'bool': return 'ParseBoolArray';
          case 'short': return 'ParseShortArray';
          case 'int': return 'ParseIntArray';
          case 'float': return 'ParseFloatArray';
          default: return '';
      }
  }

  switch (t) {
    case 'bool': return 'ParseBool';
    case 'short': return 'ParseShort';
    case 'int': return 'ParseInt';
    case 'float': return 'ParseFloat';
    default: return ''; // string doesn't need parsing
  }
};

// Helper to map type to Modbus method
export const getReadMethod = (type: string): string => {
  const t = type.trim().toLowerCase();
  // Arrays read registers directly
  if (t.endsWith('[]') && !t.startsWith('bool')) return 'ReadRegisters';
  if (t === 'bool[]') return 'ReadCoils';

  switch (t) {
    case 'bool': return `ReadCoils`;
    case 'short': return `ReadRegisters`;
    case 'int': return `ReadDInt`;
    case 'string': return `ReadString`;
    default: return `ReadRegisters`;
  }
};
