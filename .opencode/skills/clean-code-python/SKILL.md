---
name: clean-code-python
description: Principios SOLID, typing, logging, anti-patrones, nomenclatura, documentacion
license: MIT
compatibility: opencode
metadata:
  audience: developers
  domain: software-engineering
---

## Principios SOLID aplicados

- **S** - Single Responsibility: `ExcelManager` (app) vs `FileExplorer` (dialogos)
- **O** - Open/Closed: Extender via herencia, no modificar clases base
- **L** - Liskov: Subclases sustituibles
- **I** - Interface Segregation: Interfaces pequeñas y especificas
- **D** - Dependency Inversion: Depender de abstracciones (protocols), no implementaciones

## Reglas obligatorias

### Typing
```python
from pathlib import Path
from typing import Optional

def method(self, path: Path) -> bool:  # Siempre type hints
    ...
```

### Logging
```python
import logging
logger = logging.getLogger(__name__)

def method(self) -> bool:
    logger.info("Iniciando %s", "method")
    try:
        ...
        logger.debug("Detalle: %s", var)
        return True
    except Exception as e:
        logger.exception("Error en %s: %s", "method", e)
        return False
```

### Anti-patrones PROHIBIDOS
- `time.sleep()` -> `.WaitForExist()`, `.Exists(timeout)`
- `send_keys("{TAB}")` -> Selectores directos
- `os.path` / string concat rutas -> `pathlib.Path`
- `except:` sin especificar -> `except Exception as e:`
- Funciones >50 lineas -> Refactorizar
- Clases >200 lineas -> Separar responsabilidades

### Nomenclatura
- Clases: `PascalCase` (`ExcelManager`)
- Metodos/funciones: `snake_case` (`open_file`)
- Constantes: `UPPER_SNAKE_CASE` (`MAX_RETRIES`)
- Privados: `_leading_underscore`