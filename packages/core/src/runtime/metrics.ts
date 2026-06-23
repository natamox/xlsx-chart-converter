export interface MetricsSink {
  increment(name: string, tags?: Record<string, string>): void;
  timing(name: string, milliseconds: number, tags?: Record<string, string>): void;
}
