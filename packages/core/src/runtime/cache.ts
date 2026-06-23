export class WorkbookCache<K, V> {
  private readonly values = new Map<K, V>();

  get(key: K): V | undefined {
    return this.values.get(key);
  }

  set(key: K, value: V): void {
    this.values.set(key, value);
  }
}
