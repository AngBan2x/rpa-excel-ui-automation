---
name: git-github-automation
description: Workflows GitHub Actions, conventional commits, PRs, sync fork, releases
license: MIT
compatibility: opencode
metadata:
  audience: developers
  domain: devops
---

## GitHub Actions CI/CD

```yaml
# .github/workflows/ci.yml
name: CI/CD
on: [push, pull_request]
jobs:
  test:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v4
      - uses: pdm-project/setup-pdm@v4
        with: {python-version: "3.12"}
      - run: pdm install -d
      - run: pdm run test
      - run: pdm run lint
```

## Conventional Commits (commitizen)

```bash
# Commit interactivo
cz commit

# Formato: <type>(<scope>): <subject>
# Types: feat, fix, refactor, test, chore, docs, perf
# Ejemplo: feat(excel): implement open_file with robust UIA selectors
```

## PR Automatizado

```bash
gh pr create --fill --base main
gh pr merge --squash --delete-branch
```

## Sync Fork

```bash
gh repo sync AngBan2x/rpa-excel-ui-automation --branch main
```

## Versionado Semántico

```bash
pdm version patch  # 0.1.0 -> 0.1.1
pdm version minor  # 0.1.0 -> 0.2.0
pdm version major  # 0.1.0 -> 1.0.0
git push --tags
```

## Pre-commit Hooks

```yaml
# .pre-commit-config.yaml
repos:
  - repo: https://github.com/astral-sh/ruff-pre-commit
    rev: v0.5.0
    hooks: [{id: ruff}, {id: ruff-format}]
  - repo: https://github.com/pre-commit/mirrors-mypy
    rev: v1.10.0
    hooks: [{id: mypy, additional_dependencies: [uiautomation]}]
```