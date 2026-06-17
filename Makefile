install:
	uv sync --all-extras && uv run pre-commit install
dev:
	uv run recipe-automation --help
test:
	uv run pytest --cov=src --cov-report=term-missing
lint:
	uv run ruff check . && uv run ruff format --check .
typecheck:
	uv run mypy src/
check:
	make lint && make typecheck && make test
clean:
	rm -rf .venv dist build *.egg-info .pytest_cache .mypy_cache
