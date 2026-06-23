.PHONY: install lint test typecheck build check clean

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
