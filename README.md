# Python Microservices Template

A template for creating Python microservices projects.

## Setup

1. Install Poetry:
```bash
curl -sSL https://install.python-poetry.org | python3 -
```

2. Install dependencies:
```bash
make install
```

3. Set up environment:
```bash
cp .env.example .env
```

4. Run the service:
```bash
make run
```

## Development

- Run tests: `make test`
- Format code: `make format`
- Run linting: `make lint`

## Docker

- Build: `make docker-build`
- Run: `make docker-run`

## Creating a New Service

1. Copy the service template:
```bash
cp -r services/service_template services/your_new_service
```

2. Update the service files:
- Update service name and description in `main.py`
- Define your models in `models.py`
- Implement your business logic in `services.py`
- Define your routes in `routes.py`
- Add tests in `tests/`

## Project Structure

```
+-- services/               # Microservices
�   +-- service_template/   # Template service
+-- shared/                # Shared utilities
+-- tests/                 # Project-level tests
+-- ...                    # Configuration files
```
