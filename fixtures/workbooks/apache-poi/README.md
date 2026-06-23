# Apache POI chart corpus

These `.xlsx` workbooks are copied from Apache POI's public test data:

https://github.com/apache/poi/tree/trunk/test-data/spreadsheet

They are used as external OOXML chart discovery, parsing, rendering, and visual
preview regression fixtures.
The files cover worksheets, chartsheets, drawing-only workbooks, multiple charts,
scatter charts, chart titles, and workbooks whose relationship targets differ in
filename casing.

Run the local preview exporter from the repository root:

```bash
make export-chart-corpus
```

The generated preview is written to `output/chart-corpus-preview/index.html`.
The `output/` directory is ignored by git.
