export interface WorkerPool {
  close(): Promise<void>;
}
