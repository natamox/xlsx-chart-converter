import { describe, expect, it } from 'vitest';
import { UnsupportedOperationError, createChartEngine, openWorkbook } from '../src/index.js';

describe('workbook facade', () => {
  it('opens a workbook source and returns an empty chart list before discovery is implemented', async () => {
    const workbook = await openWorkbook(Buffer.from([]));

    await expect(workbook.listCharts()).resolves.toEqual([]);
    await workbook.close();
  });

  it('creates an engine that opens workbook handles', async () => {
    const engine = createChartEngine();
    const workbook = await engine.open(new Uint8Array());

    await expect(workbook.listCharts()).resolves.toEqual([]);
    await workbook.close();
  });

  it('rejects chart model requests with a structured unsupported operation error', async () => {
    const workbook = await openWorkbook(Buffer.from([]));

    await expect(workbook.getChartModel('chart-1')).rejects.toMatchObject({
      code: 'ERR_NOT_IMPLEMENTED',
      name: 'UnsupportedOperationError'
    });
    await workbook.close();
  });

  it('rejects operations after close', async () => {
    const workbook = await openWorkbook(Buffer.from([]));

    await workbook.close();

    await expect(workbook.listCharts()).rejects.toBeInstanceOf(UnsupportedOperationError);
    await expect(workbook.listCharts()).rejects.toMatchObject({
      code: 'ERR_WORKBOOK_CLOSED'
    });
  });
});
