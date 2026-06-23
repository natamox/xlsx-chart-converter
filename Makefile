.PHONY: install lint test typecheck build check clean export-chart-corpus

install:
	pnpm install

lint:
	pnpm run lint

test:
	pnpm run test

typecheck:
	pnpm run typecheck

build:
	pnpm run build

check:
	pnpm run check

clean:
	pnpm run clean

export-chart-corpus: build
	node scripts/export-chart-corpus.mjs
