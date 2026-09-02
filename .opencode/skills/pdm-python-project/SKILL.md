---
name: pdm-python-project
description: Gestion dependencias PDM, scripts, build, publish, entornos virtuales
license: MIT
compatibility: opencode
metadata:
  audience: developers
  domain: python-packaging
---

## Qué cubre

- **PDM**: `pdm init`, `pdm add`, `pdm install`, `pdm run`, `pdm build`, `pdm publish`
- **Scripts**: `[tool.pdm.scripts]` en `pyproject.toml`
- **Dependencias**: `dependencies`, `optional-dependencies.dev`
- **Python**: `requires-python`, version constraints
- **Build**: `build-system`, `setuptools`, `wheel`

## Patrones clave

```toml
# pyproject.toml
[project]
name = "rpa-excel-ui-automation"
version = "0.1.0"
dependencies = ["uiautomation>=2.0.29", "openpyxl>=3.1.2"]
requires-python = ">=3.12"

[project.optional-dependencies]
dev = ["pytest", "ruff", "mypy", "pre-commit", "commitizen"]

[tool.pdm.scripts]
test = "pytest -v --cov=src --cov-fail-under=90"
lint = "ruff check src tests && mypy src"
format = "ruff format src tests"
pre-commit = "pre-commit run --all-files"
sync-fork = "gh repo sync AngBan2x/rpa-excel-ui-automation --branch main"

[tool.ruff]
line-length = 100
target-version = "py312"

[tool.mypy]
python_version = "3.12"
strict = true
```