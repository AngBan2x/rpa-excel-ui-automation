---
name: test-driven-development
description: Estructura pytest, fixtures, mocking, coverage, TDD cycle
license: MIT
compatibility: opencode
metadata:
  audience: developers
  domain: testing
---

## Ciclo TDD

1. **RED**: Escribir test que falla
2. **GREEN**: Implementar minimo para pasar
3. **REFACTOR**: Limpiar codigo, mantener tests verdes

## Estructura pytest

```
tests/
├── __init__.py
├── conftest.py          # Fixtures compartidas
├── test_excel_manager.py
└── test_file_explorer.py
```

## Fixtures (conftest.py)

```python
import pytest
from unittest.mock import MagicMock
from pathlib import Path

@pytest.fixture
def mock_uia():
    with patch('rpa_excel_ui_automation.excel_manager.auto') as m:
        yield m

@pytest.fixture
def sample_path():
    return Path(".data/input/origen.xlsx")
```

## Mocking UI Automation

```python
# Mock WindowControl chain
mock_app = MagicMock()
mock_app.Exists.return_value = True
mock_app.WaitForExist.return_value = True

mock_dlg = MagicMock()
mock_dlg.WaitForExist.return_value = True

mock_uia.WindowControl.side_effect = [mock_app, mock_dlg]
```

## Coverage Objetivos

```bash
# Minimo 90%
pytest --cov=src --cov-fail-under=90 --cov-report=term-missing

# Excluir tests
# [tool.coverage.run] omit = ["tests/*"]
```

## Parametrización

```python
@pytest.mark.parametrize("path,expected", [
    (Path("file.xlsx"), True),
    (Path("missing.xlsx"), False),
])
def test_open_file_dialog(path, expected, file_explorer, mock_uia):
    ...
```