export interface ExternalDataProvider {
  resolve(formula: string): Promise<readonly unknown[]>;
}
