#!/usr/bin/env node

import { printExportNotImplemented } from './commands/export.js';
import { printInspectNotImplemented } from './commands/inspect.js';

const [, , command] = process.argv;

const help = `excel-chart

Usage:
  excel-chart list <file> [--json] [--include-unsupported]
  excel-chart inspect <file> <chart-id> [--ir] [--raw-xml] [--diagnostics]
  excel-chart export <file> --chart <id|name|all> --format <svg|png> --out <dir>

Status:
  The project scaffold is ready; chart discovery and rendering are planned for M1-M5.
`;

if (command === undefined || command === '--help' || command === '-h') {
  process.stdout.write(help);
  process.exitCode = 0;
} else if (command === 'export') {
  printExportNotImplemented();
  process.exitCode = 2;
} else if (command === 'inspect') {
  printInspectNotImplemented();
  process.exitCode = 2;
} else {
  process.stderr.write(`Command "${command}" is not implemented yet.\n\n${help}`);
  process.exitCode = 2;
}
