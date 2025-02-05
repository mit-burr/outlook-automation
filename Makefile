.PHONY: install test lint format run docker-build docker-run

install:
    poetry install

test:
    poetry run pytest

lint:
    poetry run flake8 .
    poetry run mypy .
    poetry run black . --check
    poetry run isort . --check-only

format:
    poetry run black .
    poetry run isort .

run:
    poetry run uvicorn services.service_template.main:app --reload

docker-build:
    docker-compose build

docker-run:
    docker-compose up
