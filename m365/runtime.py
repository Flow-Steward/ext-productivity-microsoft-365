"""Shared operation framework for the mail, files and excel packages.

A service package declares its operations as :class:`Operation` rows. The
framework validates the envelope, resolves the connection, applies Test Run
suppression to writes, builds the Graph transport lazily and turns
:class:`M365Error` into the closed failure envelope (with external-effect facts
for writes).
"""

from __future__ import annotations

from collections.abc import Callable, Mapping
from dataclasses import dataclass, field
from typing import Any

from flowsteward_extension_sdk import runtime_context_from_payload

from .auth import transport_for
from .connection import ConnectionContext, mapping, resolve_connection, text
from .errors import M365Error, error_from_exception, error_response
from .transport import GraphTransport

TransportFactory = Callable[[ConnectionContext], GraphTransport]


@dataclass
class OperationContext:
    payload: dict[str, Any]
    operation_id: str
    input: dict[str, Any]
    connection: ConnectionContext
    test_mode: bool
    transport_factory: TransportFactory
    _transport: GraphTransport | None = field(default=None, repr=False)

    @property
    def transport(self) -> GraphTransport:
        if self._transport is None:
            self._transport = self.transport_factory(self.connection)
        return self._transport

    @property
    def idempotency_key(self) -> str:
        """Stable per step execution: the same on an automatic step retry."""
        try:
            runtime = runtime_context_from_payload(self.payload)
            step_execution_id = runtime.step_execution_id or ""
        except ValueError:
            step_execution_id = ""
        explicit = text(
            mapping(self.payload.get("runtime_context")).get("idempotency_key"), limit=256
        )
        key = explicit or step_execution_id
        if not key:
            raise M365Error("invalid_payload", "The step has no execution id for idempotency")
        return f"{self.operation_id}:{key}"


Handler = Callable[[OperationContext], dict[str, Any]]


@dataclass(frozen=True)
class Operation:
    operation_id: str
    service: str
    handler: Handler
    #: Declared in external_effects; suppressed in Test Run.
    effectful: bool = False
    #: Input fields that are allowed besides ``connection_ref``.
    fields: frozenset[str] = frozenset()


def operation_envelope(payload: Mapping[str, Any]) -> tuple[str, dict[str, Any]]:
    action = mapping(payload.get("action"))
    operation_id = text(action.get("action_id") or action.get("operation_id"), limit=128)
    raw_input = action.get("input")
    if raw_input is None:
        raw_input = {}
    if not operation_id or not isinstance(raw_input, Mapping):
        raise M365Error("invalid_payload", "The operation request is invalid")
    return operation_id, dict(raw_input)


def test_mode(payload: Mapping[str, Any]) -> bool:
    runtime = mapping(payload.get("runtime_context"))
    value = runtime.get("test_mode", False)
    if not isinstance(value, bool):
        raise M365Error("invalid_payload", "runtime_context.test_mode must be boolean")
    return value


def suppressed_response() -> dict[str, Any]:
    return {
        "ok": True,
        "result": {
            "test_mode_status": "suppressed",
            "external_effect_status": "suppressed",
            "definitely_no_external_effect": True,
        },
        "external_effect_status": "suppressed",
        "definitely_no_external_effect": True,
    }


def run_operation(
    payload: dict[str, Any],
    registry: Mapping[str, Operation],
    *,
    transport_factory: TransportFactory = transport_for,
) -> dict[str, Any]:
    operation: Operation | None = None
    try:
        operation_id, raw_input = operation_envelope(payload)
        operation = registry.get(operation_id)
        if operation is None:
            raise M365Error("invalid_payload", "Unknown Microsoft 365 operation")
        unknown = set(raw_input) - set(operation.fields) - {"connection_ref"}
        if unknown:
            raise M365Error(
                "invalid_payload", f"Unknown input fields: {', '.join(sorted(unknown))}"
            )
        is_test = test_mode(payload)
        if operation.effectful and is_test:
            return suppressed_response()
        connection = resolve_connection(
            payload,
            service=operation.service,
            connection_ref=text(raw_input.get("connection_ref"), limit=256),
        )
        context = OperationContext(
            payload=payload,
            operation_id=operation_id,
            input=raw_input,
            connection=connection,
            test_mode=is_test,
            transport_factory=transport_factory,
        )
        result = operation.handler(context)
        response: dict[str, Any] = {"ok": True, "result": result, "errors": []}
        if operation.effectful:
            status = str(result.get("external_effect_status") or "succeeded")
            response["external_effect_status"] = status
            response["definitely_no_external_effect"] = bool(
                result.get("definitely_no_external_effect", False)
            )
        return response
    except M365Error as exc:
        return error_from_exception(exc, effectful=bool(operation and operation.effectful))
    except Exception:
        return error_response("internal_error", "The Microsoft 365 operation failed")


__all__ = [
    "Operation",
    "OperationContext",
    "operation_envelope",
    "run_operation",
    "suppressed_response",
    "test_mode",
]
