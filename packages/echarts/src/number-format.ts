export function createNumberFormatter(format: string | undefined): ((value: unknown) => string) | undefined {
  if (!format) {
    return undefined;
  }

  const normalized = format.toLowerCase();
  if (normalized === 'general') {
    return undefined;
  }
  if (normalized.includes('%')) {
    const decimals = decimalPlaces(format);
    return (value) => {
      const number = toFiniteNumber(value);
      return number === undefined ? '' : `${(number * 100).toFixed(decimals)}%`;
    };
  }

  if (normalized.includes('#,##') || normalized.includes('0,0')) {
    const decimals = decimalPlaces(format);
    return (value) => {
      const number = toFiniteNumber(value);
      return number === undefined ? '' : number.toLocaleString('en-US', {
        maximumFractionDigits: decimals,
        minimumFractionDigits: decimals
      });
    };
  }

  const decimals = decimalPlaces(format);
  if (decimals > 0) {
    return (value) => {
      const number = toFiniteNumber(value);
      return number === undefined ? '' : number.toFixed(decimals);
    };
  }

  return (value) => {
    const number = toFiniteNumber(value);
    return number === undefined ? '' : String(Math.round(number));
  };
}

function decimalPlaces(format: string): number {
  const dotIndex = format.indexOf('.');
  if (dotIndex < 0) {
    return 0;
  }
  const decimalPattern = format.slice(dotIndex + 1).match(/[0#]+/);
  return decimalPattern?.[0].length ?? 0;
}

function toFiniteNumber(value: unknown): number | undefined {
  const number = typeof value === 'number' ? value : Number(value);
  return Number.isFinite(number) ? number : undefined;
}
