from pydantic import BaseModel, Field
from typing import Literal, Optional, Dict, Any, List
from datetime import datetime
import uuid


class TaskSpec(BaseModel):
    """Especificacion de tarea para subagentes."""
    objective: str = Field(..., description="Objetivo alto nivel de la tarea")
    task_type: Literal[
        "architect",
        "impl-excel-manager",
        "impl-file-explorer",
        "test-writer",
        "reviewer",
        "git-manager",
        "fix"
    ] = Field(..., description="Tipo de subagente a invocar")
    context: Dict[str, Any] = Field(default_factory=dict, description="Contexto: spec, files, prior_results")
    acceptance_criteria: List[str] = Field(default_factory=list, description="Criterios de aceptacion")
    iteration: int = Field(default=1, description="Numero de iteracion actual")


class TaskResult(BaseModel):
    """Resultado de ejecucion de subagente."""
    agent: str = Field(..., description="Nombre del subagente que ejecuto")
    success: bool = Field(..., description="Si la tarea tuvo exito")
    files_changed: List[str] = Field(default_factory=list, description="Archivos modificados/creados")
    output_summary: str = Field(default="", description="Resumen de lo hecho")
    errors: List[str] = Field(default_factory=list, description="Errores encontrados")
    warnings: List[str] = Field(default_factory=list, description="Warnings/violations")
    feedback: Optional[str] = Field(default=None, description="Feedback para siguiente iteracion")


class TestRunResult(BaseModel):
    """Resultado de ejecucion de tests."""
    passed: int = 0
    failed: int = 0
    skipped: int = 0
    coverage_percent: float = 0.0
    failed_tests: List[str] = Field(default_factory=list)
    error_output: str = ""


class OrchestratorState(BaseModel):
    """Estado persistente del orquestador entre invocaciones."""
    session_id: str = Field(default_factory=lambda: str(uuid.uuid4())[:8])
    objective: str = Field(..., description="Objetivo original del usuario")
    phase: Literal[
        "planning",
        "implementing",
        "testing",
        "reviewing",
        "committing",
        "done",
        "failed"
    ] = Field(default="planning")
    current_iteration: int = Field(default=1)
    max_iterations: int = Field(default=3)
    completed_tasks: Dict[str, TaskResult] = Field(default_factory=dict)
    pending_tasks: List[TaskSpec] = Field(default_factory=list)
    decisions: List[str] = Field(default_factory=list, description="Decisiones tecnicas tomadas")
    test_results: Optional[TestRunResult] = None
    created_at: datetime = Field(default_factory=datetime.now)
    updated_at: datetime = Field(default_factory=datetime.now)

    def update_timestamp(self) -> None:
        self.updated_at = datetime.now()

    def add_decision(self, decision: str) -> None:
        self.decisions.append(f"[{datetime.now().isoformat()}] {decision}")

    def add_task_result(self, result: TaskResult) -> None:
        self.completed_tasks[result.agent] = result
        self.update_timestamp()


class Spec(BaseModel):
    """Especificacion tecnica compartida (output de architect)."""
    project: str = "rpa-excel-ui-automation"
    classes: List[Dict[str, Any]] = Field(default_factory=list)
    paths: Dict[str, str] = Field(default_factory=dict)
    rules: List[str] = Field(default_factory=list)
    selectors: Dict[str, Dict[str, str]] = Field(default_factory=dict)