"""
measurement.py
Version: measurement-fixed-2026-07-30-v4

CATIA V5 MCP measurement and parameter tools.

Important corrections:
- SPAWorkbench is obtained from the active Document, not CATIA.Application.
- Measurable.Volume is preserved as m^3 and converted to cm^3/mm^3.
- Measurable.Area is preserved as m^2 and converted to cm^2/mm^2.
- Parameter values expose ValueAsString, parameter type, unit and write status.
- Numeric dimension parameters are set with ValuateFromString when a unit exists.
- Parameter update failures attempt to restore the previous value.
- Parameter assignments are read back and verified; silent CATIA rejection is an error.
- CATSafeArrayVariant outputs use BYREF VARIANT marshalling and write detection.
- Unwritten or physically invalid zero inertia arrays are never reported as valid data.
- CATIA SystemService.Evaluate is used as the official CATSafeArrayVariant fallback.
- The fallback runs fixed in-memory VBScript only; no external macro file is created.
- Distance searches reject ambiguous duplicate names by default.
"""

from __future__ import annotations

import math
import re
from typing import Any, Optional

from catia_mcp.connection import CATIAError


IMPLEMENTATION_VERSION = "measurement-fixed-2026-07-30-v4"


# ---------------------------------------------------------------------------
# Standard results
# ---------------------------------------------------------------------------

def _success(data: Any, warnings: Optional[list[str]] = None) -> dict[str, Any]:
    warning_list = list(warnings or [])
    return {
        "implementation_version": IMPLEMENTATION_VERSION,
        "ok": True,
        "status": "success_with_warnings" if warning_list else "success",
        "data": data,
        "warnings": warning_list,
    }


def _error(
    message: str,
    *,
    data: Any = None,
    warnings: Optional[list[str]] = None,
    status: str = "error",
) -> dict[str, Any]:
    result: dict[str, Any] = {
        "implementation_version": IMPLEMENTATION_VERSION,
        "ok": False,
        "status": status,
        "error": str(message),
        "warnings": list(warnings or []),
    }
    if data is not None:
        result["data"] = data
    return result


# ---------------------------------------------------------------------------
# Connection and COM helpers
# ---------------------------------------------------------------------------

def _get_connection(ctx: Any) -> Any:
    return getattr(ctx, "conn", ctx)


def _ensure_connected(conn: Any) -> None:
    method = getattr(conn, "ensure_connected", None)
    if callable(method):
        method()


def _get_application(conn: Any) -> Any:
    app = getattr(conn, "app", None)
    if app is None:
        app = getattr(conn, "_app", None)

    if app is not None:
        return app

    method = getattr(conn, "ensure_connected", None)
    if callable(method):
        app = method()
        if app is not None:
            return app

    raise CATIAError("Cannot access the CATIA Application object.")


def _active_document(conn: Any) -> Any:
    app = getattr(conn, "app", None)
    if app is None:
        app = getattr(conn, "_app", None)

    if app is not None:
        try:
            return app.ActiveDocument
        except Exception:
            pass

    getter = getattr(conn, "get_active_document", None)
    if callable(getter):
        try:
            return getter()
        except Exception:
            pass

    raise CATIAError("Cannot access the active CATIA document.")


def _get_active_part(conn: Any) -> Any:
    getter = getattr(conn, "get_active_part", None)
    if callable(getter):
        return getter()

    document = _active_document(conn)
    try:
        return document.Part
    except Exception as exc:
        raise CATIAError(
            "The active CATIA document is not a CATPart document."
        ) from exc


def _get_active_body(conn: Any, part: Any) -> Any:
    getter = getattr(conn, "get_active_part_body", None)
    if callable(getter):
        try:
            return getter()
        except Exception:
            pass

    try:
        return part.MainBody
    except Exception:
        pass

    try:
        return part.PartBody
    except Exception as exc:
        raise CATIAError(
            "Cannot access the active CATPart main solid body."
        ) from exc


def _normalise_name(value: Any, parameter_name: str) -> str:
    text = str(value).strip()
    if not text:
        raise CATIAError(f"{parameter_name} cannot be empty.")
    return text


def _finite_number(value: Any, parameter_name: str) -> float:
    if isinstance(value, bool):
        raise CATIAError(f"{parameter_name} must be a finite number.")

    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise CATIAError(
            f"{parameter_name} must be a finite number."
        ) from exc

    if not math.isfinite(number):
        raise CATIAError(f"{parameter_name} must be finite.")

    return number


def _positive_integer(value: Any, parameter_name: str) -> int:
    if isinstance(value, bool):
        raise CATIAError(f"{parameter_name} must be a positive integer.")

    try:
        integer = int(value)
        numeric = float(value)
    except (TypeError, ValueError) as exc:
        raise CATIAError(
            f"{parameter_name} must be a positive integer."
        ) from exc

    if numeric != float(integer) or integer <= 0:
        raise CATIAError(f"{parameter_name} must be a positive integer.")

    return integer


def _describe_com_object(value: Any) -> dict[str, Any]:
    return {
        "python_type": type(value).__name__,
        "python_module": type(value).__module__,
        "has_oleobj": bool(hasattr(value, "_oleobj_")),
    }


def _dispatch_com(value: Any) -> Any:
    import win32com.client  # type: ignore

    return win32com.client.Dispatch(value)


def _dispatch_if_possible(value: Any) -> tuple[Any, dict[str, Any]]:
    details = {
        "raw": _describe_com_object(value),
        "dispatch_used": False,
    }

    try:
        dispatched = _dispatch_com(value)
        details["dispatch_used"] = True
        details["resolved"] = _describe_com_object(dispatched)
        return dispatched, details
    except Exception as exc:
        details["dispatch_error"] = str(exc)
        details["resolved"] = details["raw"]
        return value, details


def _object_name(value: Any, fallback: str = "") -> str:
    try:
        name = str(value.Name).strip()
        return name or fallback
    except Exception:
        return fallback


def _clear_selection(document: Any) -> None:
    try:
        document.Selection.Clear()
    except Exception:
        pass


def _refresh_display(conn: Any) -> list[str]:
    warnings: list[str] = []
    method = getattr(conn, "refresh_display", None)
    if callable(method):
        try:
            method()
        except Exception as exc:
            warnings.append(f"Display refresh failed: {exc}")
    return warnings


def _json_scalar(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value

    try:
        return float(value)
    except Exception:
        pass

    try:
        return str(value)
    except Exception:
        return None


def _numeric_property(
    value: Any,
    property_name: str,
) -> tuple[Optional[float], Optional[str]]:
    try:
        result = getattr(value, property_name)
        if callable(result):
            result = result()
        number = float(result)
        if not math.isfinite(number):
            return None, f"{property_name} returned a non-finite value."
        return number, None
    except Exception as exc:
        return None, str(exc)


def _numeric_sequence(
    candidate: Any,
    length: int,
) -> tuple[Optional[list[float]], Optional[str]]:
    if candidate is None:
        return None, "No array value was returned."

    try:
        if hasattr(candidate, "value"):
            candidate = candidate.value
    except Exception:
        pass

    try:
        sequence = list(candidate)
    except Exception as exc:
        return None, f"The returned value is not iterable: {exc}"

    if len(sequence) < length:
        return None, (
            f"The returned array contains {len(sequence)} values; "
            f"{length} values were required."
        )

    try:
        values = [float(sequence[index]) for index in range(length)]
    except (TypeError, ValueError) as exc:
        return None, f"The returned array contains non-numeric values: {exc}"

    if not all(math.isfinite(item) for item in values):
        return None, "The returned array contains non-finite values."

    return values, None


def _array_is_unchanged(
    values: list[float],
    sentinel: list[float],
) -> bool:
    return len(values) == len(sentinel) and all(
        actual == expected
        for actual, expected in zip(values, sentinel)
    )


def _array_is_effectively_zero(
    values: list[float],
    tolerance: float = 1.0e-30,
) -> bool:
    return all(abs(item) <= tolerance for item in values)


def _call_out_array(
    value: Any,
    method_name: str,
    length: int,
    *,
    allow_all_zero: bool = True,
) -> tuple[
    Optional[list[float]],
    Optional[str],
    dict[str, Any],
]:
    """Call a CATIA method with a CATSafeArrayVariant output argument.

    CATIA exposes methods such as GetInertiaMatrix as VB-style Subs with an
    output SAFEARRAY. Depending on the pywin32 dispatch mode, the values may be
    returned by a no-argument call or may require a BYREF VARIANT array.

    A normal Python list initialized with zeros is unsafe: CATIA may accept the
    call without writing into the list, making an untouched buffer look like a
    valid all-zero result. This helper uses distinctive sentinels, records each
    strategy, and rejects unchanged buffers.
    """

    diagnostics: dict[str, Any] = {
        "method": method_name,
        "required_length": length,
        "allow_all_zero": bool(allow_all_zero),
        "attempts": [],
        "selected_strategy": None,
        "written_output_verified": False,
    }

    raw_method = getattr(value, method_name, None)
    if not callable(raw_method):
        return (
            None,
            f"{method_name} is not available.",
            diagnostics,
        )

    candidates: list[tuple[str, Any]] = [("raw_dispatch", value)]

    try:
        import win32com.client  # type: ignore

        try:
            generated = win32com.client.gencache.EnsureDispatch(value)
            if generated is not value:
                candidates.append(("generated_dispatch", generated))
        except Exception as exc:
            diagnostics["attempts"].append(
                {
                    "strategy": "generated_dispatch",
                    "success": False,
                    "error": str(exc),
                }
            )
    except Exception as exc:
        win32com = None  # type: ignore
        diagnostics["attempts"].append(
            {
                "strategy": "import_win32com",
                "success": False,
                "error": str(exc),
            }
        )

    errors: list[str] = []

    def accept_values(
        values: Optional[list[float]],
        error: Optional[str],
        strategy: str,
        *,
        written_verified: bool,
    ) -> Optional[list[float]]:
        attempt: dict[str, Any] = {
            "strategy": strategy,
            "success": values is not None,
            "written_output_verified": written_verified,
        }

        if error:
            attempt["error"] = error

        if values is not None:
            attempt["values"] = values

            if not allow_all_zero and _array_is_effectively_zero(values):
                attempt["success"] = False
                attempt["error"] = (
                    "The returned array is entirely zero, which is not "
                    "physically valid for this non-empty inertia result."
                )
                diagnostics["attempts"].append(attempt)
                errors.append(f"{strategy}: {attempt['error']}")
                return None

            diagnostics["attempts"].append(attempt)
            diagnostics["selected_strategy"] = strategy
            diagnostics["written_output_verified"] = written_verified
            return values

        diagnostics["attempts"].append(attempt)
        errors.append(f"{strategy}: {error or 'No values returned.'}")
        return None

    # Strategy 1: generated or dynamic wrappers may expose the output as a
    # normal tuple returned by a zero-argument call.
    for candidate_name, candidate in candidates:
        method = getattr(candidate, method_name, None)
        if not callable(method):
            continue

        strategy = f"{candidate_name}.no_argument_return"
        try:
            result = method()
            values, error = _numeric_sequence(result, length)
            accepted = accept_values(
                values,
                error,
                strategy,
                written_verified=(values is not None),
            )
            if accepted is not None:
                return accepted, None, diagnostics
        except Exception as exc:
            error = str(exc)
            diagnostics["attempts"].append(
                {
                    "strategy": strategy,
                    "success": False,
                    "error": error,
                }
            )
            errors.append(f"{strategy}: {error}")

    # Strategies 2 and 3: explicitly marshal a BYREF SAFEARRAY. CATIA calls
    # the parameter CATSafeArrayVariant; different pywin32 dispatch modes may
    # accept either SAFEARRAY(VARIANT) or SAFEARRAY(DOUBLE).
    try:
        import pythoncom  # type: ignore
        from win32com.client import VARIANT  # type: ignore

        variant_types = [
            (
                "byref_safearray_variant",
                pythoncom.VT_BYREF
                | pythoncom.VT_ARRAY
                | pythoncom.VT_VARIANT,
            ),
            (
                "byref_safearray_double",
                pythoncom.VT_BYREF
                | pythoncom.VT_ARRAY
                | pythoncom.VT_R8,
            ),
        ]

        for candidate_name, candidate in candidates:
            method = getattr(candidate, method_name, None)
            if not callable(method):
                continue

            for variant_name, variant_type in variant_types:
                strategy = f"{candidate_name}.{variant_name}"
                sentinel = [
                    9_876_543.125 + (index * 101.25)
                    for index in range(length)
                ]

                try:
                    output_variant = VARIANT(
                        variant_type,
                        list(sentinel),
                    )
                    result = method(output_variant)

                    candidate_value = (
                        result
                        if result is not None
                        else output_variant.value
                    )
                    values, error = _numeric_sequence(
                        candidate_value,
                        length,
                    )

                    if (
                        values is not None
                        and _array_is_unchanged(values, sentinel)
                    ):
                        values = None
                        error = (
                            "CATIA returned without changing the BYREF "
                            "SAFEARRAY sentinel."
                        )

                    accepted = accept_values(
                        values,
                        error,
                        strategy,
                        written_verified=(values is not None),
                    )
                    if accepted is not None:
                        return accepted, None, diagnostics
                except Exception as exc:
                    error = str(exc)
                    diagnostics["attempts"].append(
                        {
                            "strategy": strategy,
                            "success": False,
                            "error": error,
                        }
                    )
                    errors.append(f"{strategy}: {error}")
    except Exception as exc:
        error = f"BYREF VARIANT marshalling unavailable: {exc}"
        diagnostics["attempts"].append(
            {
                "strategy": "byref_variant_setup",
                "success": False,
                "error": error,
            }
        )
        errors.append(error)

    # Last compatibility strategy: a mutable Python list. It is initialized
    # with a sentinel, never zeros, so an unmodified list cannot be mistaken
    # for valid output.
    for candidate_name, candidate in candidates:
        method = getattr(candidate, method_name, None)
        if not callable(method):
            continue

        strategy = f"{candidate_name}.sentinel_python_list"
        sentinel = [
            8_765_432.875 + (index * 97.5)
            for index in range(length)
        ]
        buffer = list(sentinel)

        try:
            result = method(buffer)
            candidate_value = result if result is not None else buffer
            values, error = _numeric_sequence(candidate_value, length)

            if values is not None and _array_is_unchanged(values, sentinel):
                values = None
                error = (
                    "CATIA returned without modifying the Python-list "
                    "sentinel; the output was not written."
                )

            accepted = accept_values(
                values,
                error,
                strategy,
                written_verified=(values is not None),
            )
            if accepted is not None:
                return accepted, None, diagnostics
        except Exception as exc:
            error = str(exc)
            diagnostics["attempts"].append(
                {
                    "strategy": strategy,
                    "success": False,
                    "error": error,
                }
            )
            errors.append(f"{strategy}: {error}")

    return (
        None,
        "; ".join(errors) or f"{method_name} returned no usable values.",
        diagnostics,
    )


_CATVB_SCRIPT_LANGUAGE = 1

_INERTIA_ARRAY_SCRIPT_SPECS: dict[str, dict[str, Any]] = {
    "GetCOGPosition": {
        "function_name": "MCP_GetCOGPosition",
        "upper_bound": 2,
    },
    "GetInertiaMatrix": {
        "function_name": "MCP_GetInertiaMatrix",
        "upper_bound": 8,
    },
    "GetPrincipalAxes": {
        "function_name": "MCP_GetPrincipalAxes",
        "upper_bound": 8,
    },
    "GetPrincipalMoments": {
        "function_name": "MCP_GetPrincipalMoments",
        "upper_bound": 2,
    },
}


def _get_system_service(
    application: Any,
) -> tuple[Any, dict[str, Any]]:
    try:
        raw_service = application.SystemService
    except Exception as exc:
        raise CATIAError(
            f"Cannot access CATIA.Application.SystemService: {exc}"
        ) from exc

    service, dispatch_details = _dispatch_if_possible(raw_service)
    return service, {
        "acquisition": "CATIA.Application.SystemService",
        "interface": dispatch_details,
    }


def _evaluate_cat_safe_array(
    application: Any,
    com_object: Any,
    method_name: str,
    length: int,
    *,
    allow_all_zero: bool,
) -> tuple[
    Optional[list[float]],
    Optional[str],
    dict[str, Any],
]:
    """Retrieve a CATSafeArrayVariant through CATIA SystemService.Evaluate.

    CATIA's Evaluate method runs a fixed in-memory VBScript function inside
    CATIA. The VBScript allocates the native array, invokes the requested
    CATIA Automation method, and returns the array as the function result.

    Only the four fixed Inertia method names in _INERTIA_ARRAY_SCRIPT_SPECS
    are accepted. No user-provided script, method name or file path is used.
    """

    diagnostics: dict[str, Any] = {
        "strategy": "SystemService.Evaluate.CATVBScriptLanguage",
        "method": method_name,
        "required_length": length,
        "allow_all_zero": bool(allow_all_zero),
        "script_language_value": _CATVB_SCRIPT_LANGUAGE,
        "fixed_script": True,
        "external_macro_file": False,
        "success": False,
        "written_output_verified": False,
        "values": None,
        "error": None,
        "system_service": None,
    }

    spec = _INERTIA_ARRAY_SCRIPT_SPECS.get(method_name)
    if spec is None:
        error = (
            f"SystemService.Evaluate fallback is not allowed for "
            f"method '{method_name}'."
        )
        diagnostics["error"] = error
        return None, error, diagnostics

    expected_length = int(spec["upper_bound"]) + 1
    if expected_length != length:
        error = (
            f"Configured VBScript array length for {method_name} is "
            f"{expected_length}, but {length} was requested."
        )
        diagnostics["error"] = error
        return None, error, diagnostics

    function_name = str(spec["function_name"])
    upper_bound = int(spec["upper_bound"])

    script = (
        f"Public Function {function_name}(inertiaObject)\n"
        f"    Dim outputValues({upper_bound})\n"
        f"    inertiaObject.{method_name} outputValues\n"
        f"    {function_name} = outputValues\n"
        f"End Function"
    )

    try:
        system_service, service_details = _get_system_service(application)
        diagnostics["system_service"] = service_details

        evaluate = getattr(system_service, "Evaluate", None)
        if not callable(evaluate):
            dispatched, dispatch_details = _dispatch_if_possible(
                system_service
            )
            diagnostics["system_service"]["evaluate_dispatch"] = (
                dispatch_details
            )
            evaluate = getattr(dispatched, "Evaluate", None)

        if not callable(evaluate):
            raise CATIAError(
                "CATIA SystemService does not expose Evaluate."
            )

        # A tuple/list of COM parameters is the documented CATSafeArrayVariant
        # input for SystemService.Evaluate.
        result = evaluate(
            script,
            _CATVB_SCRIPT_LANGUAGE,
            function_name,
            [com_object],
        )

        values, parse_error = _numeric_sequence(result, length)
        if values is None:
            raise CATIAError(
                parse_error
                or "SystemService.Evaluate returned no numeric array."
            )

        if not allow_all_zero and _array_is_effectively_zero(values):
            raise CATIAError(
                "SystemService.Evaluate returned an entirely zero array, "
                "which is not physically valid for this non-empty inertia "
                "result."
            )

        diagnostics["success"] = True
        diagnostics["written_output_verified"] = True
        diagnostics["values"] = values
        return values, None, diagnostics
    except Exception as exc:
        error = str(exc)
        diagnostics["error"] = error
        return None, error, diagnostics


def _call_inertia_array(
    application: Any,
    inertia: Any,
    method_name: str,
    length: int,
    *,
    allow_all_zero: bool,
) -> tuple[
    Optional[list[float]],
    Optional[str],
    dict[str, Any],
]:
    """Use direct COM strategies, then CATIA's official Evaluate fallback."""

    values, direct_error, direct_diagnostics = _call_out_array(
        inertia,
        method_name,
        length,
        allow_all_zero=allow_all_zero,
    )

    combined: dict[str, Any] = {
        "method": method_name,
        "required_length": length,
        "allow_all_zero": bool(allow_all_zero),
        "selected_strategy": None,
        "written_output_verified": False,
        "direct_com": direct_diagnostics,
        "system_service_evaluate": None,
        "fallback_used": False,
    }

    if values is not None:
        combined["selected_strategy"] = (
            direct_diagnostics.get("selected_strategy")
        )
        combined["written_output_verified"] = bool(
            direct_diagnostics.get("written_output_verified")
        )
        return values, None, combined

    evaluate_values, evaluate_error, evaluate_diagnostics = (
        _evaluate_cat_safe_array(
            application,
            inertia,
            method_name,
            length,
            allow_all_zero=allow_all_zero,
        )
    )
    combined["system_service_evaluate"] = evaluate_diagnostics
    combined["fallback_used"] = True

    if evaluate_values is not None:
        combined["selected_strategy"] = (
            "SystemService.Evaluate.CATVBScriptLanguage"
        )
        combined["written_output_verified"] = True
        return evaluate_values, None, combined

    errors = [
        error
        for error in (direct_error, evaluate_error)
        if error
    ]
    return (
        None,
        "; ".join(errors) or f"{method_name} returned no usable values.",
        combined,
    )


def _validate_inertia_matrix(
    matrix: Optional[list[float]],
) -> dict[str, Any]:
    validation: dict[str, Any] = {
        "available": matrix is not None,
        "nonzero": False,
        "positive_diagonal": False,
        "symmetric": False,
        "valid": False,
    }
    if matrix is None or len(matrix) < 9:
        return validation

    validation["nonzero"] = not _array_is_effectively_zero(matrix)
    validation["positive_diagonal"] = all(
        matrix[index] > 0.0
        for index in (0, 4, 8)
    )
    validation["symmetric"] = all(
        _numbers_close(matrix[a], matrix[b], relative_tolerance=1.0e-7)
        for a, b in ((1, 3), (2, 6), (5, 7))
    )
    validation["valid"] = bool(
        validation["nonzero"]
        and validation["positive_diagonal"]
        and validation["symmetric"]
    )
    return validation


def _validate_principal_axes(
    axes: Optional[list[float]],
) -> dict[str, Any]:
    validation: dict[str, Any] = {
        "available": axes is not None,
        "nonzero": False,
        "norms": None,
        "dot_products": None,
        "approximately_orthonormal": False,
        "valid": False,
    }
    if axes is None or len(axes) < 9:
        return validation

    # CATIA returns [A1x,A2x,A3x,A1y,A2y,A3y,A1z,A2z,A3z].
    vectors = [
        (axes[0], axes[3], axes[6]),
        (axes[1], axes[4], axes[7]),
        (axes[2], axes[5], axes[8]),
    ]
    norms = [
        math.sqrt(sum(component * component for component in vector))
        for vector in vectors
    ]
    dots = [
        sum(vectors[a][i] * vectors[b][i] for i in range(3))
        for a, b in ((0, 1), (0, 2), (1, 2))
    ]

    validation["nonzero"] = not _array_is_effectively_zero(axes)
    validation["norms"] = norms
    validation["dot_products"] = dots
    validation["approximately_orthonormal"] = bool(
        all(abs(norm - 1.0) <= 1.0e-6 for norm in norms)
        and all(abs(dot) <= 1.0e-6 for dot in dots)
    )
    validation["valid"] = bool(
        validation["nonzero"]
        and validation["approximately_orthonormal"]
    )
    return validation


def _validate_principal_moments(
    moments: Optional[list[float]],
) -> dict[str, Any]:
    validation: dict[str, Any] = {
        "available": moments is not None,
        "nonzero": False,
        "all_positive": False,
        "valid": False,
    }
    if moments is None or len(moments) < 3:
        return validation

    validation["nonzero"] = not _array_is_effectively_zero(moments)
    validation["all_positive"] = all(value > 0.0 for value in moments)
    validation["valid"] = bool(
        validation["nonzero"]
        and validation["all_positive"]
    )
    return validation


def _get_spa_workbench(document: Any) -> tuple[Any, dict[str, Any]]:
    try:
        raw_spa = document.GetWorkbench("SPAWorkbench")
    except Exception as exc:
        raise CATIAError(
            "Cannot obtain SPAWorkbench from the active Document. "
            f"Document.GetWorkbench('SPAWorkbench') failed: {exc}"
        ) from exc

    spa, dispatch_details = _dispatch_if_possible(raw_spa)
    return spa, {
        "acquisition": "ActiveDocument.GetWorkbench('SPAWorkbench')",
        "interface": dispatch_details,
    }


def _get_measurable(
    spa: Any,
    reference: Any,
) -> tuple[Any, dict[str, Any]]:
    candidates = [spa]
    dispatched, dispatch_details = _dispatch_if_possible(spa)
    if dispatched is not spa:
        candidates.append(dispatched)

    errors: list[str] = []
    for candidate in candidates:
        for method_name in ("GetMeasurable", "Measurable"):
            method = getattr(candidate, method_name, None)
            if not callable(method):
                continue
            try:
                raw_measurable = method(reference)
                measurable, measurable_dispatch = _dispatch_if_possible(
                    raw_measurable
                )
                return measurable, {
                    "method": method_name,
                    "spa_dispatch": dispatch_details,
                    "measurable_dispatch": measurable_dispatch,
                }
            except Exception as exc:
                errors.append(f"{method_name}: {exc}")

    detail = "; ".join(errors) if errors else "No measurable method exposed."
    raise CATIAError(f"Cannot create a Measurable object: {detail}")


# ---------------------------------------------------------------------------
# Parameter helpers
# ---------------------------------------------------------------------------

def _value_as_string(parameter: Any) -> tuple[Optional[str], Optional[str]]:
    try:
        method = getattr(parameter, "ValueAsString")
        value = method() if callable(method) else method
        return str(value), None
    except Exception as exc:
        return None, str(exc)


def _parameter_unit(parameter: Any) -> tuple[Optional[str], Optional[str]]:
    try:
        raw_unit = parameter.Unit
    except Exception as exc:
        return None, str(exc)

    unit, _ = _dispatch_if_possible(raw_unit)

    for attribute_name in ("Symbol", "Name"):
        try:
            value = getattr(unit, attribute_name)
            if callable(value):
                value = value()
            text = str(value).strip()
            if text:
                return text, None
        except Exception:
            continue

    return None, "The parameter Unit object did not expose Symbol or Name."


def _unit_from_value_string(value_as_string: Optional[str]) -> Optional[str]:
    if not value_as_string:
        return None

    text = value_as_string.strip()
    # Keeps common CATIA symbols such as mm, deg, rad, kg, N, Pa and %.
    match = re.search(
        r"[-+]?(?:\d+(?:[.,]\d*)?|[.,]\d+)(?:[Ee][-+]?\d+)?\s*(.*)$",
        text,
    )
    if not match:
        return None

    suffix = match.group(1).strip()
    return suffix or None


def _parameter_type(
    parameter: Any,
    raw_value: Any,
    unit_symbol: Optional[str],
) -> str:
    if unit_symbol:
        return "Dimension"

    try:
        if bool(parameter.IsTrueParameter) is False:
            return "GeometricalParameter"
    except Exception:
        pass

    if isinstance(raw_value, bool):
        return "Boolean"
    if isinstance(raw_value, int):
        return "Integer"
    if isinstance(raw_value, float):
        return "Real"
    if isinstance(raw_value, str):
        return "String"

    return type(raw_value).__name__


def _optional_property(parameter: Any, name: str) -> Any:
    try:
        value = getattr(parameter, name)
        if callable(value):
            value = value()
        return _json_scalar(value)
    except Exception:
        return None


def _parameter_snapshot(
    parameter: Any,
    *,
    index: Optional[int] = None,
) -> tuple[dict[str, Any], list[str]]:
    warnings: list[str] = []

    name = _object_name(parameter)
    raw_value: Any = None
    try:
        raw_value = parameter.Value
    except Exception as exc:
        warnings.append(f"Parameter '{name}' Value could not be read: {exc}")

    value_as_string, value_string_error = _value_as_string(parameter)
    if value_string_error:
        warnings.append(
            f"Parameter '{name}' ValueAsString could not be read: "
            f"{value_string_error}"
        )

    unit_symbol, unit_error = _parameter_unit(parameter)
    if unit_symbol is None:
        unit_symbol = _unit_from_value_string(value_as_string)

    # Lack of Unit is normal for Real/String/Boolean parameters.
    if unit_error and unit_symbol:
        unit_error = None

    snapshot: dict[str, Any] = {
        "name": name,
        "value": _json_scalar(raw_value),
        "value_as_string": value_as_string,
        "unit": unit_symbol,
        "parameter_type": _parameter_type(
            parameter,
            raw_value,
            unit_symbol,
        ),
        "read_only": _optional_property(parameter, "ReadOnly"),
        "hidden": _optional_property(parameter, "Hidden"),
        "is_true_parameter": _optional_property(
            parameter,
            "IsTrueParameter",
        ),
        "user_access_mode": _optional_property(
            parameter,
            "UserAccessMode",
        ),
    }

    if index is not None:
        snapshot["index"] = index

    try:
        snapshot["comment"] = str(parameter.Comment)
    except Exception:
        snapshot["comment"] = None

    return snapshot, warnings



_UNIT_FACTORS: dict[str, tuple[str, float]] = {
    # Length, base unit: metre.
    "m": ("length", 1.0),
    "meter": ("length", 1.0),
    "metre": ("length", 1.0),
    "mm": ("length", 1.0e-3),
    "millimeter": ("length", 1.0e-3),
    "millimetre": ("length", 1.0e-3),
    "cm": ("length", 1.0e-2),
    "centimeter": ("length", 1.0e-2),
    "centimetre": ("length", 1.0e-2),
    "um": ("length", 1.0e-6),
    "µm": ("length", 1.0e-6),
    "μm": ("length", 1.0e-6),
    "in": ("length", 0.0254),
    "inch": ("length", 0.0254),
    "ft": ("length", 0.3048),
    # Angle, base unit: radian.
    "rad": ("angle", 1.0),
    "radian": ("angle", 1.0),
    "deg": ("angle", math.pi / 180.0),
    "degree": ("angle", math.pi / 180.0),
    "°": ("angle", math.pi / 180.0),
    # Mass, base unit: kilogram.
    "kg": ("mass", 1.0),
    "g": ("mass", 1.0e-3),
    "mg": ("mass", 1.0e-6),
    # Time, base unit: second.
    "s": ("time", 1.0),
    "sec": ("time", 1.0),
    "min": ("time", 60.0),
    "h": ("time", 3600.0),
    # Dimensionless percentage.
    "%": ("dimensionless", 0.01),
}


def _normalise_unit_symbol(unit: Any) -> str:
    return str(unit or "").strip().lower().replace(" ", "")


def _parse_value_with_unit(
    value_as_string: Optional[str],
) -> tuple[Optional[float], Optional[str], Optional[str]]:
    if not value_as_string:
        return None, None, "ValueAsString is empty."

    text = str(value_as_string).strip()
    match = re.match(
        r"^([-+]?(?:\d+(?:[.,]\d*)?|[.,]\d+)"
        r"(?:[Ee][-+]?\d+)?)\s*(.*?)\s*$",
        text,
    )
    if not match:
        return None, None, (
            f"Could not parse numeric ValueAsString: {value_as_string!r}."
        )

    number_text = match.group(1).replace(",", ".")
    try:
        number = float(number_text)
    except ValueError as exc:
        return None, None, str(exc)

    if not math.isfinite(number):
        return None, None, "ValueAsString contains a non-finite value."

    return number, match.group(2).strip() or None, None


def _convert_between_units(
    value: float,
    source_unit: str,
    target_unit: str,
) -> tuple[Optional[float], Optional[str]]:
    source = _normalise_unit_symbol(source_unit)
    target = _normalise_unit_symbol(target_unit)

    if source == target:
        return value, None

    source_spec = _UNIT_FACTORS.get(source)
    target_spec = _UNIT_FACTORS.get(target)

    if source_spec is None:
        return None, f"Unsupported requested unit for verification: {source_unit!r}."
    if target_spec is None:
        return None, f"Unsupported CATIA result unit for verification: {target_unit!r}."
    if source_spec[0] != target_spec[0]:
        return None, (
            f"Unit dimension mismatch: requested {source_unit!r}, "
            f"CATIA returned {target_unit!r}."
        )

    base_value = value * source_spec[1]
    return base_value / target_spec[1], None


def _numbers_close(
    actual: float,
    expected: float,
    *,
    relative_tolerance: float = 1.0e-9,
    absolute_tolerance: float = 1.0e-9,
) -> bool:
    return math.isclose(
        actual,
        expected,
        rel_tol=relative_tolerance,
        abs_tol=absolute_tolerance,
    )


def _verify_parameter_assignment(
    requested_value: float,
    requested_unit: str,
    effective_unit: Optional[str],
    before: dict[str, Any],
    after: dict[str, Any],
) -> dict[str, Any]:
    actual_number, actual_unit, parse_error = _parse_value_with_unit(
        after.get("value_as_string")
    )

    if actual_number is None:
        raw_actual = after.get("value")
        try:
            actual_number = float(raw_actual)
        except (TypeError, ValueError):
            actual_number = None
        actual_unit = after.get("unit") or effective_unit

    source_unit = requested_unit or effective_unit or ""
    target_unit = actual_unit or after.get("unit") or ""

    verification: dict[str, Any] = {
        "verified": False,
        "requested_value": requested_value,
        "requested_unit": requested_unit or None,
        "effective_unit": effective_unit,
        "actual_value": actual_number,
        "actual_unit": target_unit or None,
        "expected_value_in_actual_unit": None,
        "parse_warning": parse_error,
        "reason": None,
    }

    if actual_number is None:
        verification["reason"] = (
            "CATIA parameter value could not be read after assignment."
        )
        return verification

    if source_unit or target_unit:
        if not source_unit or not target_unit:
            verification["reason"] = (
                "Requested and actual parameter units could not both be "
                "resolved for assignment verification."
            )
            return verification

        expected_value, conversion_error = _convert_between_units(
            requested_value,
            source_unit,
            target_unit,
        )
        if expected_value is None:
            verification["reason"] = conversion_error
            return verification
    else:
        expected_value = requested_value

    verification["expected_value_in_actual_unit"] = expected_value
    verification["verified"] = _numbers_close(
        actual_number,
        expected_value,
    )

    if not verification["verified"]:
        verification["reason"] = (
            "CATIA did not apply the requested parameter value. "
            f"Expected approximately {expected_value:g}"
            f"{target_unit}, but read back {actual_number:g}{target_unit}."
        )
    else:
        verification["reason"] = "Read-back value matches the request."

    # This is useful for diagnosing silent CATIA rejection.
    verification["value_changed_from_before"] = (
        before.get("value_as_string") != after.get("value_as_string")
        or before.get("value") != after.get("value")
    )
    return verification


def _assign_numeric_parameter(
    parameter: Any,
    value: float,
    unit: str,
    existing_unit: Optional[str],
) -> dict[str, Any]:
    requested_unit = unit.strip()
    effective_unit = requested_unit or (existing_unit or "")

    if effective_unit:
        expression = f"{value:.17g}{effective_unit}"
        method = getattr(parameter, "ValuateFromString", None)
        if not callable(method):
            dispatched, _ = _dispatch_if_possible(parameter)
            method = getattr(dispatched, "ValuateFromString", None)

        if not callable(method):
            raise CATIAError(
                "The parameter has a unit, but ValuateFromString is not "
                "available on its COM interface."
            )

        method(expression)
        return {
            "assignment_method": "ValuateFromString",
            "assignment_expression": expression,
            "requested_unit": requested_unit or None,
            "effective_unit": effective_unit,
        }

    parameter.Value = value
    return {
        "assignment_method": "Value",
        "assignment_expression": None,
        "requested_unit": requested_unit or None,
        "effective_unit": None,
    }


def _restore_parameter(
    parameter: Any,
    old_value: Any,
    old_value_as_string: Optional[str],
    part: Any,
) -> tuple[bool, list[str]]:
    warnings: list[str] = []

    if old_value_as_string:
        try:
            method = getattr(parameter, "ValuateFromString", None)
            if not callable(method):
                dispatched, _ = _dispatch_if_possible(parameter)
                method = getattr(dispatched, "ValuateFromString", None)
            if callable(method):
                method(old_value_as_string)
                part.Update()
                return True, warnings
        except Exception as exc:
            warnings.append(
                f"Rollback with ValuateFromString failed: {exc}"
            )

    try:
        parameter.Value = old_value
        part.Update()
        return True, warnings
    except Exception as exc:
        warnings.append(f"Rollback with Value failed: {exc}")
        return False, warnings


# ---------------------------------------------------------------------------
# Named object search
# ---------------------------------------------------------------------------

def _search_named_object(
    document: Any,
    object_name: str,
    occurrence: int,
    require_unique: bool,
) -> tuple[Any, dict[str, Any]]:
    name = _normalise_name(object_name, "element name")
    index = _positive_integer(occurrence, "occurrence")
    selection = document.Selection

    selection.Clear()
    try:
        selection.Search(f"Name={name},all")
        count = int(selection.Count)

        if count == 0:
            raise CATIAError(f"Element not found: {name}")

        if bool(require_unique) and count != 1:
            raise CATIAError(
                f"Element name '{name}' is ambiguous: {count} objects "
                "matched. Rename the objects uniquely or call with "
                "require_unique=false and an explicit occurrence."
            )

        if index > count:
            raise CATIAError(
                f"Occurrence {index} was requested for '{name}', but only "
                f"{count} objects matched."
            )

        selected = selection.Item(index)
        obj = selected.Value
        resolved_name = _object_name(obj, name)

        return obj, {
            "requested_name": name,
            "resolved_name": resolved_name,
            "match_count": count,
            "occurrence": index,
            "require_unique": bool(require_unique),
            "object": _describe_com_object(obj),
        }
    finally:
        selection.Clear()


# ---------------------------------------------------------------------------
# Technological Inertia helper
# ---------------------------------------------------------------------------

def _get_technological_inertia(
    document: Any,
) -> tuple[Any, dict[str, Any]]:
    try:
        product = document.Product
    except Exception as exc:
        raise CATIAError(
            f"The active CATPart document does not expose Product: {exc}"
        ) from exc

    try:
        raw_inertia = product.GetTechnologicalObject("Inertia")
    except Exception as exc:
        raise CATIAError(
            "Product.GetTechnologicalObject('Inertia') failed: "
            f"{exc}"
        ) from exc

    inertia, dispatch_details = _dispatch_if_possible(raw_inertia)
    return inertia, {
        "acquisition": "ActiveDocument.Product."
        "GetTechnologicalObject('Inertia')",
        "interface": dispatch_details,
    }


# ---------------------------------------------------------------------------
# MCP tools
# ---------------------------------------------------------------------------

def register_tools(mcp: Any, ctx: Any) -> list[str]:
    conn = _get_connection(ctx)
    names: list[str] = []

    @mcp.tool()
    def catia_list_parameters(
        name_contains: str = "",
        limit: int = 200,
    ) -> dict[str, Any]:
        """List active CATPart parameters with values, units and metadata."""

        warnings: list[str] = []

        try:
            max_items = _positive_integer(limit, "limit")
            if max_items > 5000:
                raise CATIAError("limit cannot exceed 5000.")

            _ensure_connected(conn)
            part = _get_active_part(conn)
            parameters = part.Parameters
            needle = str(name_contains).strip().lower()

            items: list[dict[str, Any]] = []
            total_count = int(parameters.Count)
            matched_before_limit = 0

            for index in range(1, total_count + 1):
                parameter = parameters.Item(index)
                name = _object_name(parameter)

                if needle and needle not in name.lower():
                    continue

                matched_before_limit += 1
                if len(items) >= max_items:
                    continue

                item, item_warnings = _parameter_snapshot(
                    parameter,
                    index=index,
                )
                items.append(item)
                warnings.extend(item_warnings)

            return _success(
                {
                    "parameters": items,
                    "total_parameter_count": total_count,
                    "matched_parameter_count": matched_before_limit,
                    "returned_parameter_count": len(items),
                    "truncated": matched_before_limit > len(items),
                    "name_contains": str(name_contains),
                    "limit": max_items,
                    "unit_semantics": {
                        "value": (
                            "Raw Parameter.Value as exposed by CATIA."
                        ),
                        "value_as_string": (
                            "CATIA ValueAsString representation; dimension "
                            "parameters normally include the display unit."
                        ),
                        "unit": (
                            "Dimension.Unit.Symbol when available, otherwise "
                            "parsed from ValueAsString."
                        ),
                    },
                },
                warnings,
            )
        except Exception as exc:
            return _error(
                str(exc),
                data={
                    "parameters": [],
                    "returned_parameter_count": 0,
                },
                warnings=warnings,
            )

    names.append("catia_list_parameters")

    @mcp.tool()
    def catia_set_parameter(
        name: str,
        value: float,
        unit: str = "",
        update: bool = True,
    ) -> dict[str, Any]:
        """Set a numeric CATPart parameter, optionally with an explicit unit."""

        warnings: list[str] = []
        parameter = None
        assignment_started = False
        old_value: Any = None
        old_value_as_string: Optional[str] = None

        try:
            parameter_name = _normalise_name(name, "name")
            requested_value = _finite_number(value, "value")
            requested_unit = str(unit).strip()

            _ensure_connected(conn)
            part = _get_active_part(conn)
            parameters = part.Parameters

            try:
                parameter = parameters.Item(parameter_name)
            except Exception as exc:
                raise CATIAError(
                    f"Parameter not found: {parameter_name}. Detail: {exc}"
                ) from exc

            before, before_warnings = _parameter_snapshot(parameter)
            warnings.extend(before_warnings)

            if before.get("read_only") is True:
                raise CATIAError(
                    f"Parameter '{parameter_name}' is read-only."
                )

            old_value = before.get("value")
            old_value_as_string = before.get("value_as_string")

            assignment = _assign_numeric_parameter(
                parameter,
                requested_value,
                requested_unit,
                before.get("unit"),
            )
            assignment_started = True

            update_succeeded: Optional[bool] = None
            if bool(update):
                try:
                    part.Update()
                    update_succeeded = True
                except Exception as exc:
                    rollback_succeeded, rollback_warnings = _restore_parameter(
                        parameter,
                        old_value,
                        old_value_as_string,
                        part,
                    )
                    warnings.extend(rollback_warnings)

                    return _error(
                        "The parameter value was assigned, but CATIA Part.Update "
                        f"failed: {exc}",
                        data={
                            "name": parameter_name,
                            "requested_value": requested_value,
                            "requested_unit": requested_unit or None,
                            "assignment": assignment,
                            "assignment_started": True,
                            "update_requested": True,
                            "update_succeeded": False,
                            "rollback_attempted": True,
                            "rollback_succeeded": rollback_succeeded,
                            "parameter_persisted": not rollback_succeeded,
                            "before": before,
                        },
                        warnings=warnings,
                        status=(
                            "error"
                            if rollback_succeeded
                            else "partial_success"
                        ),
                    )

                warnings.extend(_refresh_display(conn))

            after, after_warnings = _parameter_snapshot(parameter)
            warnings.extend(after_warnings)

            verification = _verify_parameter_assignment(
                requested_value,
                requested_unit,
                assignment.get("effective_unit"),
                before,
                after,
            )

            if not verification["verified"]:
                rollback_succeeded, rollback_warnings = _restore_parameter(
                    parameter,
                    old_value,
                    old_value_as_string,
                    part,
                )
                warnings.extend(rollback_warnings)

                final_snapshot, final_warnings = _parameter_snapshot(parameter)
                warnings.extend(final_warnings)

                return _error(
                    verification["reason"]
                    or "CATIA did not apply the requested parameter value.",
                    data={
                        "name": parameter_name,
                        "requested_value": requested_value,
                        "requested_unit": requested_unit or None,
                        "assignment": assignment,
                        "assignment_started": True,
                        "assignment_verified": False,
                        "assignment_verification": verification,
                        "before": before,
                        "after_assignment": after,
                        "after_rollback": final_snapshot,
                        "update_requested": bool(update),
                        "update_succeeded": update_succeeded,
                        "rollback_attempted": True,
                        "rollback_succeeded": rollback_succeeded,
                        "parameter_persisted": not rollback_succeeded,
                    },
                    warnings=warnings,
                    status=(
                        "error"
                        if rollback_succeeded
                        else "partial_success"
                    ),
                )

            return _success(
                {
                    "name": parameter_name,
                    "requested_value": requested_value,
                    "requested_unit": requested_unit or None,
                    "assignment": assignment,
                    "assignment_started": True,
                    "assignment_verified": True,
                    "assignment_verification": verification,
                    "before": before,
                    "after": after,
                    "update_requested": bool(update),
                    "update_succeeded": update_succeeded,
                    "parameter_persisted": True,
                    "rollback_attempted": False,
                    "rollback_succeeded": None,
                },
                warnings,
            )
        except Exception as exc:
            rollback_attempted = False
            rollback_succeeded: Optional[bool] = None

            if parameter is not None and assignment_started:
                rollback_attempted = True
                try:
                    part = _get_active_part(conn)
                    rollback_succeeded, rollback_warnings = _restore_parameter(
                        parameter,
                        old_value,
                        old_value_as_string,
                        part,
                    )
                    warnings.extend(rollback_warnings)
                except Exception as rollback_exc:
                    rollback_succeeded = False
                    warnings.append(
                        f"Parameter rollback could not be executed: "
                        f"{rollback_exc}"
                    )

            return _error(
                str(exc),
                data={
                    "name": str(name),
                    "requested_value": _json_scalar(value),
                    "requested_unit": str(unit).strip() or None,
                    "assignment_started": assignment_started,
                    "assignment_verified": False,
                    "update_requested": bool(update),
                    "update_succeeded": False,
                    "rollback_attempted": rollback_attempted,
                    "rollback_succeeded": rollback_succeeded,
                    "parameter_persisted": bool(
                        assignment_started
                        and rollback_succeeded is not True
                    ),
                },
                warnings=warnings,
                status=(
                    "partial_success"
                    if assignment_started
                    and rollback_succeeded is False
                    else "error"
                ),
            )

    names.append("catia_set_parameter")

    @mcp.tool()
    def catia_get_inertia(
        density_kg_m3: float = 0.0,
    ) -> dict[str, Any]:
        """Measure active PartBody geometry and retrieve mass/inertia data.

        density_kg_m3=0 keeps CATIA material/default density behavior.
        A positive value requests a temporary density override for the
        technological Inertia object and is also used for a geometric mass
        calculation from measured volume.
        """

        warnings: list[str] = []

        try:
            requested_density = _finite_number(
                density_kg_m3,
                "density_kg_m3",
            )
            if requested_density < 0.0:
                raise CATIAError(
                    "density_kg_m3 must be 0 or a positive value."
                )

            _ensure_connected(conn)
            application = _get_application(conn)
            document = _active_document(conn)
            part = _get_active_part(conn)
            body = _get_active_body(conn, part)

            data: dict[str, Any] = {
                "target_body": _object_name(body, "MainBody"),
                "requested_density_kg_m3": requested_density,
                "units": {
                    "volume_m3": "m^3",
                    "volume_cm3": "cm^3",
                    "volume_mm3": "mm^3",
                    "area_m2": "m^2",
                    "area_cm2": "cm^2",
                    "area_mm2": "mm^2",
                    "center_of_gravity_mm": "mm",
                    "technological_cog_m": "m",
                    "technological_cog_mm": "mm",
                    "mass_kg": "kg",
                    "mass_g": "g",
                    "density_kg_m3": "kg/m^3",
                    "inertia_matrix_kg_m2": "kg*m^2",
                    "principal_moments_kg_m2": "kg*m^2",
                },
                "measurement_methods": {
                    "cat_safe_array_policy": {
                        "primary": "direct_pywin32_COM",
                        "fallback": (
                            "CATIA.Application.SystemService.Evaluate "
                            "with fixed in-memory CATVBScript"
                        ),
                        "script_language_value": _CATVB_SCRIPT_LANGUAGE,
                        "external_macro_file": False,
                        "user_script_input": False,
                    },
                },
                "geometry": {
                    "available": False,
                    "volume_m3": None,
                    "volume_cm3": None,
                    "volume_mm3": None,
                    "area_m2": None,
                    "area_cm2": None,
                    "area_mm2": None,
                    "center_of_gravity_mm": None,
                },
                "mass_properties": {
                    "available": False,
                    "density_kg_m3": None,
                    "mass_kg": None,
                    "mass_g": None,
                    "center_of_gravity_m": None,
                    "center_of_gravity_mm": None,
                    "inertia_matrix_kg_m2": None,
                    "principal_axes": None,
                    "principal_moments_kg_m2": None,
                    "array_call_diagnostics": {},
                    "cat_safe_array_fallback_used": False,
                    "cat_safe_array_fallback_methods": [],
                    "inertia_validation": {
                        "matrix": None,
                        "principal_axes": None,
                        "principal_moments": None,
                        "all_valid": False,
                    },
                    "density_override_applied": False,
                    "density_restored": None,
                },
            }

            # Geometrical measurement through Document-level SPAWorkbench.
            try:
                spa, spa_details = _get_spa_workbench(document)
                reference = part.CreateReferenceFromObject(body)
                measurable, measurable_details = _get_measurable(
                    spa,
                    reference,
                )
                data["measurement_methods"]["geometry"] = {
                    "spa": spa_details,
                    "measurable": measurable_details,
                }

                geometry = data["geometry"]
                geometry["available"] = True

                volume_m3, volume_error = _numeric_property(
                    measurable,
                    "Volume",
                )
                if volume_m3 is not None:
                    geometry["volume_m3"] = volume_m3
                    geometry["volume_cm3"] = volume_m3 * 1_000_000.0
                    geometry["volume_mm3"] = volume_m3 * 1_000_000_000.0
                else:
                    warnings.append(
                        f"Measurable.Volume was unavailable: {volume_error}"
                    )

                area_m2, area_error = _numeric_property(
                    measurable,
                    "Area",
                )
                if area_m2 is not None:
                    geometry["area_m2"] = area_m2
                    geometry["area_cm2"] = area_m2 * 10_000.0
                    geometry["area_mm2"] = area_m2 * 1_000_000.0
                else:
                    warnings.append(
                        f"Measurable.Area was unavailable: {area_error}"
                    )

                cog_mm, cog_error, cog_diagnostics = _call_out_array(
                    measurable,
                    "GetCOG",
                    3,
                    allow_all_zero=True,
                )
                data["measurement_methods"]["geometry_cog"] = (
                    cog_diagnostics
                )
                if cog_mm is not None:
                    geometry["center_of_gravity_mm"] = {
                        "x": cog_mm[0],
                        "y": cog_mm[1],
                        "z": cog_mm[2],
                    }
                else:
                    warnings.append(
                        f"Measurable.GetCOG was unavailable: {cog_error}"
                    )

                if requested_density > 0.0 and volume_m3 is not None:
                    geometric_mass_kg = requested_density * volume_m3
                    geometry["density_kg_m3"] = requested_density
                    geometry["calculated_mass_kg"] = geometric_mass_kg
                    geometry["calculated_mass_g"] = (
                        geometric_mass_kg * 1000.0
                    )
            except Exception as exc:
                warnings.append(
                    f"SPA geometrical measurement was unavailable: {exc}"
                )

            # Product technological Inertia provides real mass and inertia data.
            inertia = None
            original_density: Optional[float] = None
            density_overridden = False

            try:
                inertia, inertia_details = _get_technological_inertia(
                    document
                )
                data["measurement_methods"]["mass_properties"] = (
                    inertia_details
                )
                mass_properties = data["mass_properties"]
                mass_properties["available"] = True

                original_density, _ = _numeric_property(
                    inertia,
                    "Density",
                )

                if requested_density > 0.0:
                    inertia.Density = requested_density
                    density_overridden = True
                    mass_properties["density_override_applied"] = True

                actual_density, density_error = _numeric_property(
                    inertia,
                    "Density",
                )
                mass_properties["density_kg_m3"] = actual_density
                if density_error:
                    warnings.append(
                        f"Inertia.Density could not be read: {density_error}"
                    )

                mass_kg, mass_error = _numeric_property(inertia, "Mass")
                if mass_kg is not None:
                    mass_properties["mass_kg"] = mass_kg
                    mass_properties["mass_g"] = mass_kg * 1000.0
                else:
                    warnings.append(
                        f"Inertia.Mass could not be read: {mass_error}"
                    )

                cog_m, cog_error, cog_diagnostics = _call_inertia_array(
                    application,
                    inertia,
                    "GetCOGPosition",
                    3,
                    allow_all_zero=True,
                )
                mass_properties["array_call_diagnostics"][
                    "GetCOGPosition"
                ] = cog_diagnostics
                if cog_m is not None:
                    mass_properties["center_of_gravity_m"] = {
                        "x": cog_m[0],
                        "y": cog_m[1],
                        "z": cog_m[2],
                    }
                    mass_properties["center_of_gravity_mm"] = {
                        "x": cog_m[0] * 1000.0,
                        "y": cog_m[1] * 1000.0,
                        "z": cog_m[2] * 1000.0,
                    }
                else:
                    warnings.append(
                        f"Inertia.GetCOGPosition was unavailable: "
                        f"{cog_error}"
                    )

                matrix, matrix_error, matrix_diagnostics = _call_inertia_array(
                    application,
                    inertia,
                    "GetInertiaMatrix",
                    9,
                    allow_all_zero=False,
                )
                mass_properties["array_call_diagnostics"][
                    "GetInertiaMatrix"
                ] = matrix_diagnostics
                matrix_validation = _validate_inertia_matrix(matrix)
                mass_properties["inertia_validation"]["matrix"] = (
                    matrix_validation
                )
                if matrix is not None and matrix_validation["valid"]:
                    mass_properties["inertia_matrix_kg_m2"] = {
                        "ixx": matrix[0],
                        "ixy": matrix[1],
                        "ixz": matrix[2],
                        "iyx": matrix[3],
                        "iyy": matrix[4],
                        "iyz": matrix[5],
                        "izx": matrix[6],
                        "izy": matrix[7],
                        "izz": matrix[8],
                    }
                else:
                    warnings.append(
                        "Inertia.GetInertiaMatrix did not produce a valid "
                        f"physical matrix: {matrix_error or matrix_validation}"
                    )

                axes, axes_error, axes_diagnostics = _call_inertia_array(
                    application,
                    inertia,
                    "GetPrincipalAxes",
                    9,
                    allow_all_zero=False,
                )
                mass_properties["array_call_diagnostics"][
                    "GetPrincipalAxes"
                ] = axes_diagnostics
                axes_validation = _validate_principal_axes(axes)
                mass_properties["inertia_validation"][
                    "principal_axes"
                ] = axes_validation
                if axes is not None and axes_validation["valid"]:
                    mass_properties["principal_axes"] = axes
                else:
                    warnings.append(
                        "Inertia.GetPrincipalAxes did not produce valid "
                        f"orthonormal axes: {axes_error or axes_validation}"
                    )

                moments, moments_error, moments_diagnostics = _call_inertia_array(
                    application,
                    inertia,
                    "GetPrincipalMoments",
                    3,
                    allow_all_zero=False,
                )
                mass_properties["array_call_diagnostics"][
                    "GetPrincipalMoments"
                ] = moments_diagnostics
                moments_validation = _validate_principal_moments(
                    moments
                )
                mass_properties["inertia_validation"][
                    "principal_moments"
                ] = moments_validation
                if moments is not None and moments_validation["valid"]:
                    mass_properties["principal_moments_kg_m2"] = {
                        "m1": moments[0],
                        "m2": moments[1],
                        "m3": moments[2],
                    }
                else:
                    warnings.append(
                        "Inertia.GetPrincipalMoments did not produce valid "
                        f"positive moments: {moments_error or moments_validation}"
                    )

                fallback_methods = [
                    method_name
                    for method_name, diagnostics in (
                        mass_properties["array_call_diagnostics"].items()
                    )
                    if diagnostics.get("fallback_used")
                    and diagnostics.get("selected_strategy")
                    == "SystemService.Evaluate.CATVBScriptLanguage"
                ]
                mass_properties["cat_safe_array_fallback_methods"] = (
                    fallback_methods
                )
                mass_properties["cat_safe_array_fallback_used"] = bool(
                    fallback_methods
                )

                mass_properties["inertia_validation"]["all_valid"] = bool(
                    matrix_validation["valid"]
                    and axes_validation["valid"]
                    and moments_validation["valid"]
                )
            except Exception as exc:
                warnings.append(
                    f"Technological Inertia was unavailable: {exc}"
                )
            finally:
                if (
                    inertia is not None
                    and density_overridden
                    and original_density is not None
                ):
                    try:
                        inertia.Density = original_density
                        data["mass_properties"]["density_restored"] = True
                    except Exception as exc:
                        data["mass_properties"]["density_restored"] = False
                        warnings.append(
                            f"The original Inertia density could not be "
                            f"restored: {exc}"
                        )

            geometry = data["geometry"]
            mass_properties = data["mass_properties"]
            meaningful_result = any(
                value is not None
                for value in (
                    geometry.get("volume_m3"),
                    geometry.get("area_m2"),
                    mass_properties.get("mass_kg"),
                    mass_properties.get("inertia_matrix_kg_m2"),
                )
            )

            if not meaningful_result:
                return _error(
                    "Neither SPA geometrical measurements nor technological "
                    "Inertia data could be obtained.",
                    data=data,
                    warnings=warnings,
                )

            return _success(data, warnings)
        except Exception as exc:
            return _error(str(exc), warnings=warnings)

    names.append("catia_get_inertia")

    @mcp.tool()
    def catia_measure_distance_by_name(
        element1_name: str,
        element2_name: str,
        occurrence1: int = 1,
        occurrence2: int = 1,
        require_unique: bool = True,
    ) -> dict[str, Any]:
        """Measure minimum distance between two named CATPart elements.

        By default each name must match exactly one object. Set
        require_unique=false and provide occurrence indexes only when duplicate
        names are intentional.
        """

        warnings: list[str] = []
        document = None

        try:
            name1 = _normalise_name(
                element1_name,
                "element1_name",
            )
            name2 = _normalise_name(
                element2_name,
                "element2_name",
            )
            index1 = _positive_integer(occurrence1, "occurrence1")
            index2 = _positive_integer(occurrence2, "occurrence2")

            _ensure_connected(conn)
            document = _active_document(conn)
            part = _get_active_part(conn)
            _clear_selection(document)

            obj1, details1 = _search_named_object(
                document,
                name1,
                index1,
                bool(require_unique),
            )
            obj2, details2 = _search_named_object(
                document,
                name2,
                index2,
                bool(require_unique),
            )

            try:
                reference1 = part.CreateReferenceFromObject(obj1)
                reference2 = part.CreateReferenceFromObject(obj2)
            except Exception as exc:
                raise CATIAError(
                    f"Cannot create CATIA References for the selected "
                    f"objects: {exc}"
                ) from exc

            spa, spa_details = _get_spa_workbench(document)
            measurable, measurable_details = _get_measurable(
                spa,
                reference1,
            )

            try:
                distance = float(
                    measurable.GetMinimumDistance(reference2)
                )
            except Exception as exc:
                raise CATIAError(
                    "GetMinimumDistance failed. CATIA does not support "
                    "minimum-distance measurement between some container "
                    "objects such as Body/HybridBody; select measurable "
                    f"points, curves, edges, faces or surfaces instead. "
                    f"Detail: {exc}"
                ) from exc

            if not math.isfinite(distance):
                raise CATIAError(
                    "GetMinimumDistance returned a non-finite distance."
                )

            return _success(
                {
                    "element1": details1,
                    "element2": details2,
                    "distance_mm": distance,
                    "units": {
                        "distance_mm": "mm",
                    },
                    "measurement_method": {
                        "spa": spa_details,
                        "measurable": measurable_details,
                        "distance_method": "GetMinimumDistance",
                    },
                    "selection_cleared": True,
                    "model_modified": False,
                },
                warnings,
            )
        except Exception as exc:
            return _error(
                str(exc),
                data={
                    "element1_name": str(element1_name),
                    "element2_name": str(element2_name),
                    "selection_cleared": True,
                    "model_modified": False,
                },
                warnings=warnings,
            )
        finally:
            if document is not None:
                _clear_selection(document)

    names.append("catia_measure_distance_by_name")

    return names


# ---------------------------------------------------------------------------
# Registry Entry Point (Required by registry.py)
# ---------------------------------------------------------------------------
