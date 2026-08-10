"""
assembly.py
Version: assembly-fixed-2026-07-31-v3

CATIA V5 MCP assembly tools.

Key corrections:
- CATIA Position and Move SAFEARRAY calls are executed inside CATIA with
  fixed in-memory CATVBScript through SystemService.Evaluate.
- Position.GetComponents is no longer called with a normal Python list as
  a simulated Visual Basic ByRef array.
- Product.Move.Apply receives a real CATIA-side 12-element transformation
  array.
- Every position matrix is validated as a rigid transform; an all-zero
  rotation block is rejected rather than reported as a readable position.
- Every move is read back and numerically verified.
- Position.SetComponents is retained only as a CATIA-side verified fallback.
- A blocked or ineffective move returns an error instead of false success.
- Component rotation supports incremental X/Y/Z rotation around the
  component origin in either local-component or parent-product axes.
- Rotation uses CATIA-side Position.SetComponents with a real SAFEARRAY,
  followed by rigid-transform readback and numerical verification.
- Partial or unexpected rotation attempts are rolled back to the exact
  pre-operation 12-element position matrix.
- Component lookup is explicit and duplicate-safe.
- Component insertion, translation and recursive listing retain their
  previously verified behavior.
"""

from __future__ import annotations

import math
import os
from pathlib import Path
from typing import Any, Optional


IMPLEMENTATION_VERSION = "assembly-fixed-2026-07-31-v3"

_ALLOWED_COMPONENT_EXTENSIONS = {".catpart", ".catproduct"}
_CATVB_SCRIPT_LANGUAGE = 1
_DESIGN_MODE = 2
_IDENTITY_ROTATION_COLUMN_MAJOR = [
    1.0, 0.0, 0.0,
    0.0, 1.0, 0.0,
    0.0, 0.0, 1.0,
]


# ---------------------------------------------------------------------------
# Standard result helpers
# ---------------------------------------------------------------------------

def _success(
    data: Any,
    warnings: Optional[list[str]] = None,
) -> dict[str, Any]:
    warning_list = list(warnings or [])
    return {
        "implementation_version": IMPLEMENTATION_VERSION,
        "ok": True,
        "status": (
            "success_with_warnings"
            if warning_list
            else "success"
        ),
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
# Connection helpers
# ---------------------------------------------------------------------------

def _get_connection(ctx: Any) -> Any:
    return getattr(ctx, "conn", ctx)


def _ensure_connected(conn: Any) -> Any:
    method = getattr(conn, "ensure_connected", None)
    if callable(method):
        result = method()
        if result is not None:
            return result

    app = getattr(conn, "app", None)
    if app is None:
        app = getattr(conn, "_app", None)
    if app is not None:
        return app

    raise RuntimeError("Cannot access the CATIA Application object.")


def _get_application(conn: Any) -> Any:
    app = getattr(conn, "app", None)
    if app is None:
        app = getattr(conn, "_app", None)
    if app is not None:
        return app
    return _ensure_connected(conn)


def _get_active_document(conn: Any) -> Any:
    getter = getattr(conn, "get_active_document", None)
    if callable(getter):
        try:
            return getter()
        except Exception:
            pass

    application = _get_application(conn)
    try:
        return application.ActiveDocument
    except Exception as exc:
        raise RuntimeError(
            "Cannot access the active CATIA document."
        ) from exc


def _document_name(document: Any) -> str:
    try:
        return str(document.Name).strip()
    except Exception:
        return ""


def _document_saved(document: Any) -> Optional[bool]:
    try:
        return bool(document.Saved)
    except Exception:
        return None


def _require_active_product_document(
    conn: Any,
) -> tuple[Any, Any]:
    document = _get_active_document(conn)
    name = _document_name(document)

    if name and not name.lower().endswith(".catproduct"):
        raise RuntimeError(
            "The active document must be a CATProduct document; "
            f"current document is '{name}'."
        )

    try:
        product = document.Product
        _ = product.Products
    except Exception as exc:
        raise RuntimeError(
            "The active document does not expose a CATProduct "
            "root Product and Products collection."
        ) from exc

    return document, product


def _prepare_component_for_positioning(
    root_product: Any,
    component: Any,
) -> tuple[dict[str, Any], list[str]]:
    """Load product representations before Position/Move operations.

    CatWorkModeType is declared as DEFAULT_MODE, VISUALIZATION_MODE,
    DESIGN_MODE, so the Automation value for DESIGN_MODE is 2.
    """
    warnings: list[str] = []
    details: dict[str, Any] = {
        "design_mode_value": _DESIGN_MODE,
        "root_apply_work_mode_attempted": True,
        "root_apply_work_mode_succeeded": False,
        "component_apply_work_mode_attempted": True,
        "component_apply_work_mode_succeeded": False,
        "activate_default_shape_attempted": True,
        "activate_default_shape_succeeded": False,
    }

    try:
        root_product.ApplyWorkMode(_DESIGN_MODE)
        details["root_apply_work_mode_succeeded"] = True
    except Exception as exc:
        details["root_apply_work_mode_error"] = str(exc)
        warnings.append(
            f"Root Product.ApplyWorkMode(DESIGN_MODE) failed: {exc}"
        )

    try:
        component.ApplyWorkMode(_DESIGN_MODE)
        details["component_apply_work_mode_succeeded"] = True
    except Exception as exc:
        details["component_apply_work_mode_error"] = str(exc)
        warnings.append(
            f"Component.ApplyWorkMode(DESIGN_MODE) failed: {exc}"
        )

    try:
        component.ActivateDefaultShape()
        details["activate_default_shape_succeeded"] = True
    except Exception as exc:
        details["activate_default_shape_error"] = str(exc)
        warnings.append(
            f"Component.ActivateDefaultShape failed: {exc}"
        )

    return details, warnings


def _refresh_display(conn: Any) -> list[str]:
    warnings: list[str] = []

    method = getattr(conn, "refresh_display", None)
    if callable(method):
        try:
            method()
            return warnings
        except Exception as exc:
            warnings.append(f"Connection display refresh failed: {exc}")

    try:
        application = _get_application(conn)
        application.ActiveWindow.ActiveViewer.Update()
    except Exception as exc:
        warnings.append(f"Active viewer refresh failed: {exc}")

    return warnings


def _update_product(product: Any) -> list[str]:
    warnings: list[str] = []
    try:
        product.Update()
    except Exception as exc:
        warnings.append(f"Product.Update failed: {exc}")
    return warnings


# ---------------------------------------------------------------------------
# CATIA SAFEARRAY bridge
# ---------------------------------------------------------------------------

def _evaluate(
    application: Any,
    script: str,
    function_name: str,
    parameters: list[Any],
) -> Any:
    try:
        service = application.SystemService
    except Exception as exc:
        raise RuntimeError(
            f"Cannot access CATIA SystemService: {exc}"
        ) from exc

    try:
        return service.Evaluate(
            script,
            _CATVB_SCRIPT_LANGUAGE,
            function_name,
            parameters,
        )
    except Exception as exc:
        raise RuntimeError(
            f"SystemService.Evaluate failed for {function_name}: {exc}"
        ) from exc


def _numeric_sequence(
    value: Any,
    expected_length: int,
) -> list[float]:
    try:
        sequence = list(value)
    except Exception as exc:
        raise RuntimeError(
            f"CATIA did not return an array: {exc}"
        ) from exc

    if len(sequence) != expected_length:
        raise RuntimeError(
            f"CATIA returned {len(sequence)} values; "
            f"{expected_length} were required."
        )

    result: list[float] = []
    for index, item in enumerate(sequence):
        try:
            number = float(item)
        except (TypeError, ValueError) as exc:
            raise RuntimeError(
                f"CATIA array element {index} is not numeric."
            ) from exc
        if not math.isfinite(number):
            raise RuntimeError(
                f"CATIA array element {index} is not finite."
            )
        result.append(number)
    return result


def _position_quality(
    matrix: list[float],
    tolerance: float = 1e-5,
) -> dict[str, Any]:
    columns = [
        (matrix[0], matrix[1], matrix[2]),
        (matrix[3], matrix[4], matrix[5]),
        (matrix[6], matrix[7], matrix[8]),
    ]

    norms = [
        math.sqrt(sum(value * value for value in column))
        for column in columns
    ]
    dot_products = [
        sum(columns[0][i] * columns[1][i] for i in range(3)),
        sum(columns[0][i] * columns[2][i] for i in range(3)),
        sum(columns[1][i] * columns[2][i] for i in range(3)),
    ]

    determinant = (
        matrix[0] * (matrix[4] * matrix[8] - matrix[7] * matrix[5])
        - matrix[3] * (matrix[1] * matrix[8] - matrix[7] * matrix[2])
        + matrix[6] * (matrix[1] * matrix[5] - matrix[4] * matrix[2])
    )

    all_zero_rotation = all(
        abs(matrix[index]) <= tolerance
        for index in range(9)
    )
    orthonormal = bool(
        not all_zero_rotation
        and all(abs(norm - 1.0) <= tolerance for norm in norms)
        and all(abs(value) <= tolerance for value in dot_products)
        and abs(abs(determinant) - 1.0) <= tolerance
    )

    return {
        "valid_rigid_transform": orthonormal,
        "all_zero_rotation": all_zero_rotation,
        "column_norms": norms,
        "column_dot_products": dot_products,
        "rotation_determinant": determinant,
        "validation_tolerance": tolerance,
    }


def _validate_position_matrix(
    matrix: list[float],
) -> dict[str, Any]:
    if len(matrix) != 12:
        raise RuntimeError(
            f"Position matrix must contain 12 values, got {len(matrix)}."
        )

    quality = _position_quality(matrix)
    if not quality["valid_rigid_transform"]:
        raise RuntimeError(
            "CATIA returned an invalid rigid-transform position matrix. "
            f"Quality: {quality}"
        )
    return quality


def _read_position_via_evaluate(
    component: Any,
    application: Any,
) -> tuple[list[float], str, dict[str, Any]]:
    script = (
        "Public Function MCP_ReadProductPosition(productObject)\n"
        "    Dim values(11)\n"
        "    productObject.Position.GetComponents values\n"
        "    MCP_ReadProductPosition = Array("
        "CDbl(values(0)), CDbl(values(1)), CDbl(values(2)), "
        "CDbl(values(3)), CDbl(values(4)), CDbl(values(5)), "
        "CDbl(values(6)), CDbl(values(7)), CDbl(values(8)), "
        "CDbl(values(9)), CDbl(values(10)), CDbl(values(11)))\n"
        "End Function"
    )
    matrix = _numeric_sequence(
        _evaluate(
            application,
            script,
            "MCP_ReadProductPosition",
            [component],
        ),
        12,
    )
    quality = _validate_position_matrix(matrix)
    return (
        matrix,
        "SystemService.Evaluate.Position.GetComponents",
        quality,
    )


def _read_position_via_direct_return(
    component: Any,
) -> tuple[list[float], str, dict[str, Any]]:
    value = component.Position.GetComponents()
    matrix = _numeric_sequence(value, 12)
    quality = _validate_position_matrix(matrix)
    return matrix, "Position.GetComponents_return_value", quality


def _read_position_via_typed_variant(
    component: Any,
) -> tuple[list[float], str, dict[str, Any]]:
    try:
        import pythoncom  # type: ignore
        from win32com.client import VARIANT  # type: ignore
    except Exception as exc:
        raise RuntimeError(
            f"Typed VARIANT support is unavailable: {exc}"
        ) from exc

    buffer = VARIANT(
        pythoncom.VT_BYREF
        | pythoncom.VT_ARRAY
        | pythoncom.VT_R8,
        [0.0] * 12,
    )
    component.Position.GetComponents(buffer)
    matrix = _numeric_sequence(buffer.value, 12)
    quality = _validate_position_matrix(matrix)
    return matrix, "Position.GetComponents_typed_BYREF_VARIANT", quality


def _read_position_matrix(
    component: Any,
    application: Any,
) -> tuple[list[float], str, dict[str, Any], list[dict[str, Any]]]:
    attempts: list[dict[str, Any]] = []
    strategies = (
        (
            "SystemService.Evaluate.Position.GetComponents",
            lambda: _read_position_via_evaluate(
                component,
                application,
            ),
        ),
        (
            "Position.GetComponents_return_value",
            lambda: _read_position_via_direct_return(component),
        ),
        (
            "Position.GetComponents_typed_BYREF_VARIANT",
            lambda: _read_position_via_typed_variant(component),
        ),
    )

    for strategy_name, strategy in strategies:
        try:
            matrix, method, quality = strategy()
            attempts.append(
                {
                    "method": strategy_name,
                    "succeeded": True,
                    "error": None,
                }
            )
            return matrix, method, quality, attempts
        except Exception as exc:
            attempts.append(
                {
                    "method": strategy_name,
                    "succeeded": False,
                    "error": str(exc),
                }
            )

    raise RuntimeError(
        "All CATIA position SAFEARRAY read strategies failed: "
        f"{attempts}"
    )


def _apply_relative_move_via_evaluate(
    component: Any,
    delta_matrix: list[float],
    application: Any,
) -> None:
    script = (
        "Public Function MCP_ApplyProductMove("
        "productObject, tx, ty, tz)\n"
        "    Dim movement(11)\n"
        "    movement(0) = 1.0\n"
        "    movement(1) = 0.0\n"
        "    movement(2) = 0.0\n"
        "    movement(3) = 0.0\n"
        "    movement(4) = 1.0\n"
        "    movement(5) = 0.0\n"
        "    movement(6) = 0.0\n"
        "    movement(7) = 0.0\n"
        "    movement(8) = 1.0\n"
        "    movement(9) = CDbl(tx)\n"
        "    movement(10) = CDbl(ty)\n"
        "    movement(11) = CDbl(tz)\n"
        "    productObject.Move.Apply movement\n"
        "    MCP_ApplyProductMove = True\n"
        "End Function"
    )
    _evaluate(
        application,
        script,
        "MCP_ApplyProductMove",
        [
            component,
            delta_matrix[9],
            delta_matrix[10],
            delta_matrix[11],
        ],
    )


def _set_position_via_evaluate(
    component: Any,
    target_matrix: list[float],
    application: Any,
) -> None:
    script = (
        "Public Function MCP_SetProductPosition("
        "productObject, a0, a1, a2, a3, a4, a5, "
        "a6, a7, a8, a9, a10, a11)\n"
        "    Dim values(11)\n"
        "    values(0) = CDbl(a0)\n"
        "    values(1) = CDbl(a1)\n"
        "    values(2) = CDbl(a2)\n"
        "    values(3) = CDbl(a3)\n"
        "    values(4) = CDbl(a4)\n"
        "    values(5) = CDbl(a5)\n"
        "    values(6) = CDbl(a6)\n"
        "    values(7) = CDbl(a7)\n"
        "    values(8) = CDbl(a8)\n"
        "    values(9) = CDbl(a9)\n"
        "    values(10) = CDbl(a10)\n"
        "    values(11) = CDbl(a11)\n"
        "    productObject.Position.SetComponents values\n"
        "    MCP_SetProductPosition = True\n"
        "End Function"
    )
    _evaluate(
        application,
        script,
        "MCP_SetProductPosition",
        [component, *target_matrix],
    )


# ---------------------------------------------------------------------------
# Validation helpers
# ---------------------------------------------------------------------------

def _nonempty_text(value: Any, parameter_name: str) -> str:
    text = str(value).strip()
    if not text:
        raise ValueError(f"{parameter_name} cannot be empty.")
    return text


def _finite_float(value: Any, parameter_name: str) -> float:
    if isinstance(value, bool):
        raise ValueError(f"{parameter_name} must be a finite number.")

    try:
        result = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(
            f"{parameter_name} must be a finite number."
        ) from exc

    if not math.isfinite(result):
        raise ValueError(f"{parameter_name} must be finite.")

    return result


def _positive_integer(value: Any, parameter_name: str) -> int:
    if isinstance(value, bool):
        raise ValueError(
            f"{parameter_name} must be a positive integer."
        )

    try:
        integer = int(value)
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(
            f"{parameter_name} must be a positive integer."
        ) from exc

    if number != float(integer) or integer < 1:
        raise ValueError(
            f"{parameter_name} must be a positive integer."
        )
    return integer


def _nonnegative_integer(value: Any, parameter_name: str) -> int:
    if isinstance(value, bool):
        raise ValueError(
            f"{parameter_name} must be a non-negative integer."
        )

    try:
        integer = int(value)
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(
            f"{parameter_name} must be a non-negative integer."
        ) from exc

    if number != float(integer) or integer < 0:
        raise ValueError(
            f"{parameter_name} must be a non-negative integer."
        )
    return integer


def _normalise_component_file(file_path: Any) -> str:
    raw_path = _nonempty_text(file_path, "file_path")
    expanded = os.path.expandvars(os.path.expanduser(raw_path))
    path = Path(expanded)

    try:
        path = path.resolve(strict=False)
    except Exception:
        path = Path(os.path.abspath(expanded))

    if not path.exists():
        raise FileNotFoundError(
            f"Component file does not exist: {path}"
        )
    if not path.is_file():
        raise ValueError(
            f"Component path is not a file: {path}"
        )

    extension = path.suffix.lower()
    if extension not in _ALLOWED_COMPONENT_EXTENSIONS:
        allowed = ", ".join(sorted(_ALLOWED_COMPONENT_EXTENSIONS))
        raise ValueError(
            "Component file must be CATPart or CATProduct; "
            f"received '{path.suffix}'. Allowed: {allowed}."
        )

    return str(path)


# ---------------------------------------------------------------------------
# Component and transformation helpers
# ---------------------------------------------------------------------------

def _safe_attribute(
    value: Any,
    attribute_name: str,
    default: Any = "",
) -> Any:
    try:
        return getattr(value, attribute_name)
    except Exception:
        return default


def _component_name(component: Any) -> str:
    return str(_safe_attribute(component, "Name", "") or "")


def _part_number(component: Any) -> str:
    return str(_safe_attribute(component, "PartNumber", "") or "")


def _component_source_path(component: Any) -> Optional[str]:
    try:
        reference = component.ReferenceProduct
    except Exception:
        reference = None

    candidates = [reference, component]
    for candidate in candidates:
        if candidate is None:
            continue

        try:
            parent = candidate.Parent
            full_name = str(parent.FullName).strip()
            if full_name:
                return full_name
        except Exception:
            pass

    return None


def _rotation_rows(matrix: list[float]) -> list[list[float]]:
    return [
        [matrix[0], matrix[3], matrix[6]],
        [matrix[1], matrix[4], matrix[7]],
        [matrix[2], matrix[5], matrix[8]],
    ]


def _matrix_payload(
    matrix: list[float],
    *,
    read_method: Optional[str] = None,
    quality: Optional[dict[str, Any]] = None,
    read_attempts: Optional[list[dict[str, Any]]] = None,
) -> dict[str, Any]:
    return {
        "components_column_major": list(matrix),
        "rotation_matrix_rows": _rotation_rows(matrix),
        "translation_mm": {
            "x": matrix[9],
            "y": matrix[10],
            "z": matrix[11],
        },
        "read_method": read_method,
        "quality": quality,
        "read_attempts": read_attempts,
        "matrix_semantics": {
            "components_0_to_8": (
                "rotation matrix stored by successive columns"
            ),
            "components_9_to_11": (
                "translation vector in millimetres"
            ),
            "position_kind": (
                "relative component position under its parent product"
            ),
        },
    }


def _component_payload(
    component: Any,
    index: int,
    application: Any,
    *,
    parent_path: str = "",
    depth: int = 1,
) -> dict[str, Any]:
    name = _component_name(component)
    position_error: Optional[str] = None
    matrix_data: Optional[dict[str, Any]] = None

    try:
        (
            matrix,
            read_method,
            quality,
            read_attempts,
        ) = _read_position_matrix(component, application)
        matrix_data = _matrix_payload(
            matrix,
            read_method=read_method,
            quality=quality,
            read_attempts=read_attempts,
        )
    except Exception as exc:
        position_error = str(exc)

    try:
        child_count = int(component.Products.Count)
    except Exception:
        child_count = 0

    try:
        move_object = component.Move
        move_available = move_object is not None
    except Exception:
        move_available = False

    component_path = (
        f"{parent_path}/{name}"
        if parent_path
        else name
    )

    return {
        "index": index,
        "name": name,
        "part_number": _part_number(component),
        "component_path": component_path,
        "depth": depth,
        "direct_child_count": child_count,
        "source_file_path": _component_source_path(component),
        "move_available": move_available,
        "position_readable": matrix_data is not None,
        "position_valid_rigid_transform": bool(
            matrix_data
            and matrix_data["quality"][
                "valid_rigid_transform"
            ]
        ),
        "position": matrix_data,
        "position_error": position_error,
    }


def _find_direct_components(
    product: Any,
    component_name: str,
) -> list[tuple[int, Any]]:
    requested = _nonempty_text(
        component_name,
        "component_name",
    )
    products = product.Products
    exact: list[tuple[int, Any]] = []

    for index in range(1, int(products.Count) + 1):
        component = products.Item(index)
        if _component_name(component) == requested:
            exact.append((index, component))

    return exact


def _resolve_direct_component(
    product: Any,
    component_name: str,
    occurrence: int,
    require_unique: bool,
) -> tuple[int, Any, dict[str, Any]]:
    requested = _nonempty_text(
        component_name,
        "component_name",
    )
    requested_occurrence = _positive_integer(
        occurrence,
        "occurrence",
    )
    matches = _find_direct_components(product, requested)
    match_count = len(matches)

    if match_count == 0:
        raise RuntimeError(
            f"Direct component not found: '{requested}'."
        )

    if require_unique and match_count != 1:
        raise RuntimeError(
            f"Component name '{requested}' is ambiguous: "
            f"{match_count} direct components matched. Rename the "
            "instances uniquely or set require_unique=false and "
            "provide an explicit occurrence."
        )

    if requested_occurrence > match_count:
        raise RuntimeError(
            f"Occurrence {requested_occurrence} was requested for "
            f"'{requested}', but only {match_count} components matched."
        )

    index, component = matches[requested_occurrence - 1]
    return index, component, {
        "requested_name": requested,
        "resolved_name": _component_name(component),
        "match_count": match_count,
        "occurrence": requested_occurrence,
        "require_unique": bool(require_unique),
        "direct_child_index": index,
    }


def _identity_translation_matrix(
    tx: float,
    ty: float,
    tz: float,
) -> list[float]:
    return [
        *_IDENTITY_ROTATION_COLUMN_MAJOR,
        float(tx),
        float(ty),
        float(tz),
    ]


def _vector_difference(
    after: list[float],
    before: list[float],
) -> dict[str, float]:
    return {
        "x": after[9] - before[9],
        "y": after[10] - before[10],
        "z": after[11] - before[11],
    }


def _max_rotation_difference(
    after: list[float],
    before: list[float],
) -> float:
    return max(
        abs(after[index] - before[index])
        for index in range(9)
    )


def _translation_matches(
    actual: dict[str, float],
    requested: dict[str, float],
    tolerance_mm: float,
) -> bool:
    return all(
        abs(actual[axis] - requested[axis]) <= tolerance_mm
        for axis in ("x", "y", "z")
    )



def _normalise_rotation_axis(value: Any) -> str:
    axis = _nonempty_text(value, "axis").lower()
    aliases = {
        "x": "x",
        "x_axis": "x",
        "x-axis": "x",
        "y": "y",
        "y_axis": "y",
        "y-axis": "y",
        "z": "z",
        "z_axis": "z",
        "z-axis": "z",
    }
    if axis not in aliases:
        raise ValueError(
            "axis must be one of: x, y, z."
        )
    return aliases[axis]


def _normalise_coordinate_system(value: Any) -> str:
    coordinate_system = _nonempty_text(
        value,
        "coordinate_system",
    ).lower()
    aliases = {
        "local": "local",
        "component": "local",
        "component_local": "local",
        "parent": "parent",
        "product": "parent",
        "parent_product": "parent",
    }
    if coordinate_system not in aliases:
        raise ValueError(
            "coordinate_system must be 'local' or 'parent'."
        )
    return aliases[coordinate_system]


def _normalise_signed_angle_deg(angle_deg: float) -> float:
    value = (float(angle_deg) + 180.0) % 360.0 - 180.0
    if abs(value) < 1e-12:
        return 0.0
    return value


def _matrix3_multiply(
    left: list[list[float]],
    right: list[list[float]],
) -> list[list[float]]:
    return [
        [
            sum(left[row][k] * right[k][column] for k in range(3))
            for column in range(3)
        ]
        for row in range(3)
    ]


def _matrix3_transpose(
    matrix: list[list[float]],
) -> list[list[float]]:
    return [
        [matrix[column][row] for column in range(3)]
        for row in range(3)
    ]


def _rotation_rows_to_column_major(
    rows: list[list[float]],
) -> list[float]:
    return [
        rows[0][0], rows[1][0], rows[2][0],
        rows[0][1], rows[1][1], rows[2][1],
        rows[0][2], rows[1][2], rows[2][2],
    ]


def _axis_angle_rotation_rows(
    axis: str,
    angle_deg: float,
) -> list[list[float]]:
    radians = math.radians(float(angle_deg))
    cosine = math.cos(radians)
    sine = math.sin(radians)

    if axis == "x":
        return [
            [1.0, 0.0, 0.0],
            [0.0, cosine, -sine],
            [0.0, sine, cosine],
        ]
    if axis == "y":
        return [
            [cosine, 0.0, sine],
            [0.0, 1.0, 0.0],
            [-sine, 0.0, cosine],
        ]
    return [
        [cosine, -sine, 0.0],
        [sine, cosine, 0.0],
        [0.0, 0.0, 1.0],
    ]


def _compose_rotation_target(
    before_matrix: list[float],
    axis: str,
    angle_deg: float,
    coordinate_system: str,
) -> tuple[list[float], list[list[float]]]:
    before_rows = _rotation_rows(before_matrix)
    delta_rows = _axis_angle_rotation_rows(
        axis,
        angle_deg,
    )

    if coordinate_system == "local":
        target_rows = _matrix3_multiply(
            before_rows,
            delta_rows,
        )
    else:
        target_rows = _matrix3_multiply(
            delta_rows,
            before_rows,
        )

    target_matrix = [
        *_rotation_rows_to_column_major(target_rows),
        before_matrix[9],
        before_matrix[10],
        before_matrix[11],
    ]
    _validate_position_matrix(target_matrix)
    return target_matrix, delta_rows


def _relative_rotation_rows(
    before_matrix: list[float],
    after_matrix: list[float],
    coordinate_system: str,
) -> list[list[float]]:
    before_rows = _rotation_rows(before_matrix)
    after_rows = _rotation_rows(after_matrix)
    before_transpose = _matrix3_transpose(before_rows)

    if coordinate_system == "local":
        return _matrix3_multiply(
            before_transpose,
            after_rows,
        )
    return _matrix3_multiply(
        after_rows,
        before_transpose,
    )


def _signed_axis_angle_deg(
    rotation_rows: list[list[float]],
    axis: str,
) -> float:
    if axis == "x":
        angle = math.degrees(
            math.atan2(
                rotation_rows[2][1],
                rotation_rows[1][1],
            )
        )
    elif axis == "y":
        angle = math.degrees(
            math.atan2(
                rotation_rows[0][2],
                rotation_rows[0][0],
            )
        )
    else:
        angle = math.degrees(
            math.atan2(
                rotation_rows[1][0],
                rotation_rows[0][0],
            )
        )
    return _normalise_signed_angle_deg(angle)


def _wrapped_angle_error_deg(
    actual_deg: float,
    expected_deg: float,
) -> float:
    return _normalise_signed_angle_deg(
        actual_deg - expected_deg
    )


def _max_matrix_component_difference(
    first: list[float],
    second: list[float],
    start: int = 0,
    end: int = 12,
) -> float:
    return max(
        abs(first[index] - second[index])
        for index in range(start, end)
    )


def _rotation_axis_vector_in_parent(
    before_matrix: list[float],
    axis: str,
    coordinate_system: str,
) -> dict[str, float]:
    if coordinate_system == "parent":
        vectors = {
            "x": (1.0, 0.0, 0.0),
            "y": (0.0, 1.0, 0.0),
            "z": (0.0, 0.0, 1.0),
        }
        vector = vectors[axis]
    else:
        column_start = {"x": 0, "y": 3, "z": 6}[axis]
        vector = (
            before_matrix[column_start],
            before_matrix[column_start + 1],
            before_matrix[column_start + 2],
        )

    return {
        "x": vector[0],
        "y": vector[1],
        "z": vector[2],
    }


def _apply_absolute_position(
    component: Any,
    target_matrix: list[float],
    application: Any,
) -> tuple[str, list[str], list[dict[str, Any]]]:
    warnings: list[str] = []
    attempts: list[dict[str, Any]] = []

    try:
        _set_position_via_evaluate(
            component,
            target_matrix,
            application,
        )
        attempts.append(
            {
                "method": (
                    "SystemService.Evaluate.Position.SetComponents"
                ),
                "succeeded": True,
                "error": None,
            }
        )
        return (
            "SystemService.Evaluate.Position.SetComponents",
            warnings,
            attempts,
        )
    except Exception as exc:
        attempts.append(
            {
                "method": (
                    "SystemService.Evaluate.Position.SetComponents"
                ),
                "succeeded": False,
                "error": str(exc),
            }
        )
        warnings.append(
            "CATIA-side Position.SetComponents failed; "
            f"trying tuple SAFEARRAY input: {exc}"
        )

    try:
        component.Position.SetComponents(tuple(target_matrix))
        attempts.append(
            {
                "method": (
                    "Position.SetComponents_tuple_SAFEARRAY"
                ),
                "succeeded": True,
                "error": None,
            }
        )
        return (
            "Position.SetComponents_tuple_SAFEARRAY",
            warnings,
            attempts,
        )
    except Exception as exc:
        attempts.append(
            {
                "method": (
                    "Position.SetComponents_tuple_SAFEARRAY"
                ),
                "succeeded": False,
                "error": str(exc),
            }
        )
        raise RuntimeError(
            "All CATIA absolute-position SAFEARRAY strategies "
            f"failed: {attempts}"
        ) from exc


def _attempt_position_matrix_rollback(
    component: Any,
    original_matrix: list[float],
    product: Any,
    application: Any,
    rotation_matrix_tolerance: float,
    translation_tolerance_mm: float,
) -> tuple[bool, dict[str, Any], list[str]]:
    warnings: list[str] = []

    try:
        (
            method,
            method_warnings,
            apply_attempts,
        ) = _apply_absolute_position(
            component,
            original_matrix,
            application,
        )
        warnings.extend(method_warnings)
        warnings.extend(_update_product(product))
    except Exception as exc:
        return False, {
            "attempted": True,
            "succeeded": False,
            "method": None,
            "error": str(exc),
        }, warnings

    try:
        (
            after_matrix,
            read_method,
            quality,
            read_attempts,
        ) = _read_position_matrix(component, application)
    except Exception as exc:
        return False, {
            "attempted": True,
            "succeeded": False,
            "method": method,
            "apply_attempts": apply_attempts,
            "error": (
                "Rollback was applied but could not be verified: "
                f"{exc}"
            ),
        }, warnings

    rotation_error = _max_matrix_component_difference(
        after_matrix,
        original_matrix,
        0,
        9,
    )
    translation_error = _max_matrix_component_difference(
        after_matrix,
        original_matrix,
        9,
        12,
    )
    succeeded = bool(
        rotation_error <= rotation_matrix_tolerance
        and translation_error <= translation_tolerance_mm
    )

    return succeeded, {
        "attempted": True,
        "succeeded": succeeded,
        "method": method,
        "apply_attempts": apply_attempts,
        "max_rotation_component_error": rotation_error,
        "max_translation_error_mm": translation_error,
        "position_after_rollback": _matrix_payload(
            after_matrix,
            read_method=read_method,
            quality=quality,
            read_attempts=read_attempts,
        ),
    }, warnings


def _apply_move(
    component: Any,
    delta_matrix: list[float],
    before_matrix: list[float],
    application: Any,
) -> tuple[str, list[str], list[dict[str, Any]]]:
    warnings: list[str] = []
    attempts: list[dict[str, Any]] = []

    try:
        _apply_relative_move_via_evaluate(
            component,
            delta_matrix,
            application,
        )
        attempts.append(
            {
                "method": (
                    "SystemService.Evaluate.Product.Move.Apply"
                ),
                "succeeded": True,
                "error": None,
            }
        )
        return (
            "SystemService.Evaluate.Product.Move.Apply",
            warnings,
            attempts,
        )
    except Exception as exc:
        attempts.append(
            {
                "method": (
                    "SystemService.Evaluate.Product.Move.Apply"
                ),
                "succeeded": False,
                "error": str(exc),
            }
        )
        warnings.append(
            "CATIA-side Move.Apply failed; trying a tuple SAFEARRAY "
            f"call: {exc}"
        )

    try:
        component.Move.Apply(tuple(delta_matrix))
        attempts.append(
            {
                "method": "Product.Move.Apply_tuple_SAFEARRAY",
                "succeeded": True,
                "error": None,
            }
        )
        return (
            "Product.Move.Apply_tuple_SAFEARRAY",
            warnings,
            attempts,
        )
    except Exception as exc:
        attempts.append(
            {
                "method": "Product.Move.Apply_tuple_SAFEARRAY",
                "succeeded": False,
                "error": str(exc),
            }
        )
        warnings.append(
            "Tuple Move.Apply failed; trying CATIA-side absolute "
            f"Position.SetComponents fallback: {exc}"
        )

    target = list(before_matrix)
    target[9] += delta_matrix[9]
    target[10] += delta_matrix[10]
    target[11] += delta_matrix[11]

    try:
        _set_position_via_evaluate(
            component,
            target,
            application,
        )
        attempts.append(
            {
                "method": (
                    "SystemService.Evaluate.Position.SetComponents"
                ),
                "succeeded": True,
                "error": None,
            }
        )
        return (
            "SystemService.Evaluate.Position.SetComponents",
            warnings,
            attempts,
        )
    except Exception as exc:
        attempts.append(
            {
                "method": (
                    "SystemService.Evaluate.Position.SetComponents"
                ),
                "succeeded": False,
                "error": str(exc),
            }
        )
        raise RuntimeError(
            "All CATIA movement SAFEARRAY strategies failed: "
            f"{attempts}"
        ) from exc


def _attempt_translation_rollback(
    component: Any,
    actual_delta: dict[str, float],
    product: Any,
    tolerance_mm: float,
    application: Any,
) -> tuple[bool, dict[str, Any], list[str]]:
    warnings: list[str] = []
    (
        rollback_before,
        rollback_read_method,
        rollback_quality,
        rollback_read_attempts,
    ) = _read_position_matrix(component, application)

    inverse = _identity_translation_matrix(
        -actual_delta["x"],
        -actual_delta["y"],
        -actual_delta["z"],
    )

    try:
        (
            rollback_method,
            apply_warnings,
            rollback_apply_attempts,
        ) = _apply_move(
            component,
            inverse,
            rollback_before,
            application,
        )
        warnings.extend(apply_warnings)
    except Exception as exc:
        return False, {
            "attempted": True,
            "succeeded": False,
            "method": None,
            "error": str(exc),
        }, warnings

    warnings.extend(_update_product(product))
    (
        rollback_after,
        rollback_after_method,
        rollback_after_quality,
        rollback_after_attempts,
    ) = _read_position_matrix(component, application)
    rollback_delta = _vector_difference(
        rollback_after,
        rollback_before,
    )

    expected = {
        "x": -actual_delta["x"],
        "y": -actual_delta["y"],
        "z": -actual_delta["z"],
    }
    succeeded = _translation_matches(
        rollback_delta,
        expected,
        tolerance_mm,
    )

    return succeeded, {
        "attempted": True,
        "succeeded": succeeded,
        "method": rollback_method,
        "apply_attempts": rollback_apply_attempts,
        "expected_delta_mm": expected,
        "actual_delta_mm": rollback_delta,
        "position_before_rollback": _matrix_payload(
            rollback_before,
            read_method=rollback_read_method,
            quality=rollback_quality,
            read_attempts=rollback_read_attempts,
        ),
        "position_after_rollback": _matrix_payload(
            rollback_after,
            read_method=rollback_after_method,
            quality=rollback_after_quality,
            read_attempts=rollback_after_attempts,
        ),
    }, warnings


def _collect_components(
    product: Any,
    application: Any,
    recursive: bool,
    max_depth: int,
    *,
    parent_path: str = "",
    depth: int = 1,
) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    products = product.Products

    for index in range(1, int(products.Count) + 1):
        component = products.Item(index)
        payload = _component_payload(
            component,
            index,
            application,
            parent_path=parent_path,
            depth=depth,
        )
        result.append(payload)

        if recursive and depth < max_depth:
            try:
                child_count = int(component.Products.Count)
            except Exception:
                child_count = 0

            if child_count > 0:
                result.extend(
                    _collect_components(
                        component,
                        application,
                        True,
                        max_depth,
                        parent_path=payload[
                            "component_path"
                        ],
                        depth=depth + 1,
                    )
                )

    return result


# ---------------------------------------------------------------------------
# MCP registration
# ---------------------------------------------------------------------------

def register_tools(mcp: Any, ctx: Any) -> list[str]:
    conn = _get_connection(ctx)
    names: list[str] = []

    @mcp.tool()
    def catia_add_component(
        file_path: str,
    ) -> dict[str, Any]:
        """Add one CATPart or CATProduct file to the active CATProduct.

        The tool validates the file, records the Products count before and
        after insertion, identifies the newly added direct component and
        returns its relative transformation.
        """

        warnings: list[str] = []
        document = None
        product = None
        products = None
        before_count: Optional[int] = None
        after_count: Optional[int] = None
        normalised_path: Optional[str] = None
        insertion_started = False

        try:
            normalised_path = _normalise_component_file(
                file_path
            )
            application = _ensure_connected(conn)
            document, product = (
                _require_active_product_document(conn)
            )
            products = product.Products
            before_count = int(products.Count)
            document_saved_before = _document_saved(document)

            products.AddComponentsFromFiles(
                [normalised_path],
                "All",
            )
            insertion_started = True

            warnings.extend(_update_product(product))
            warnings.extend(_refresh_display(conn))

            after_count = int(products.Count)
            added_count = after_count - before_count

            if added_count != 1:
                return _error(
                    "CATIA did not add exactly one direct component. "
                    f"Products count changed from {before_count} "
                    f"to {after_count}.",
                    data={
                        "requested_file_path": file_path,
                        "normalised_file_path": normalised_path,
                        "component_count_before": before_count,
                        "component_count_after": after_count,
                        "added_count": added_count,
                        "insertion_started": insertion_started,
                        "insertion_verified": False,
                        "model_modified": added_count > 0,
                        "document_save_required": added_count > 0,
                    },
                    warnings=warnings,
                    status=(
                        "partial_success"
                        if added_count > 0
                        else "error"
                    ),
                )

            added_index = after_count
            added_component = products.Item(added_index)
            added_payload = _component_payload(
                added_component,
                added_index,
                application,
            )

            return _success(
                {
                    "requested_file_path": file_path,
                    "normalised_file_path": normalised_path,
                    "file_extension": (
                        Path(normalised_path).suffix
                    ),
                    "active_document": _document_name(
                        document
                    ),
                    "root_product_name": _component_name(
                        product
                    ),
                    "component_count_before": before_count,
                    "component_count_after": after_count,
                    "added_count": 1,
                    "added_component": added_payload,
                    "insertion_started": True,
                    "insertion_verified": True,
                    "source_file_modified": False,
                    "feature_created": False,
                    "model_modified": True,
                    "document_save_required": True,
                    "document_saved_before": (
                        document_saved_before
                    ),
                    "document_saved_after": (
                        _document_saved(document)
                    ),
                    "api_method": (
                        "Products.AddComponentsFromFiles"
                    ),
                },
                warnings,
            )
        except Exception as exc:
            if products is not None and before_count is not None:
                try:
                    after_count = int(products.Count)
                except Exception:
                    pass

            modified = bool(
                insertion_started
                and after_count is not None
                and before_count is not None
                and after_count > before_count
            )

            return _error(
                str(exc),
                data={
                    "requested_file_path": file_path,
                    "normalised_file_path": normalised_path,
                    "component_count_before": before_count,
                    "component_count_after": after_count,
                    "insertion_started": insertion_started,
                    "insertion_verified": False,
                    "source_file_modified": False,
                    "feature_created": False,
                    "model_modified": modified,
                    "document_save_required": modified,
                },
                warnings=warnings,
                status=(
                    "partial_success"
                    if modified
                    else "error"
                ),
            )

    names.append("catia_add_component")

    @mcp.tool()
    def catia_list_components(
        recursive: bool = False,
        max_depth: int = 20,
    ) -> dict[str, Any]:
        """List components in the active CATProduct.

        By default only direct children are returned. Set recursive=true to
        include nested subassemblies up to max_depth.
        """

        try:
            depth_limit = _positive_integer(
                max_depth,
                "max_depth",
            )
            application = _ensure_connected(conn)
            document, product = (
                _require_active_product_document(conn)
            )

            direct_count = int(product.Products.Count)
            components = _collect_components(
                product,
                application,
                bool(recursive),
                depth_limit,
            )
            unreadable_count = sum(
                1
                for item in components
                if not item["position_readable"]
            )

            warnings: list[str] = []
            if unreadable_count:
                warnings.append(
                    f"{unreadable_count} component position(s) "
                    "could not be read. See position_error."
                )

            return _success(
                {
                    "active_document": _document_name(
                        document
                    ),
                    "root_product_name": _component_name(
                        product
                    ),
                    "recursive": bool(recursive),
                    "max_depth": depth_limit,
                    "direct_component_count": direct_count,
                    "returned_component_count": len(
                        components
                    ),
                    "position_readable_count": (
                        len(components) - unreadable_count
                    ),
                    "position_unreadable_count": (
                        unreadable_count
                    ),
                    "components": components,
                    "model_modified": False,
                    "document_save_required": False,
                },
                warnings,
            )
        except Exception as exc:
            return _error(
                str(exc),
                data={
                    "recursive": recursive,
                    "max_depth": max_depth,
                    "model_modified": False,
                    "document_save_required": False,
                },
            )

    names.append("catia_list_components")

    @mcp.tool()
    def catia_move_component(
        component_name: str,
        tx: float = 0.0,
        ty: float = 0.0,
        tz: float = 0.0,
        occurrence: int = 1,
        require_unique: bool = True,
        tolerance_mm: float = 0.001,
    ) -> dict[str, Any]:
        """Incrementally translate one direct component.

        Translation is applied with a CATIA-side SAFEARRAY through
        SystemService.Evaluate.Product.Move.Apply and verified with a
        CATIA-side Position.GetComponents readback. If CATIA accepts the
        call but the component does not move, the tool returns an error.

        A fixed or constrained component may reject or undo the requested
        movement. In that case remove or deactivate the relevant assembly
        constraints before retrying.
        """

        warnings: list[str] = []
        document = None
        product = None
        component = None
        resolution: Optional[dict[str, Any]] = None
        before_matrix: Optional[list[float]] = None
        after_matrix: Optional[list[float]] = None
        movement_method: Optional[str] = None
        movement_apply_attempts: list[dict[str, Any]] = []
        positioning_preparation: Optional[dict[str, Any]] = None
        before_read_method: Optional[str] = None
        before_quality: Optional[dict[str, Any]] = None
        before_read_attempts: list[dict[str, Any]] = []
        after_read_method: Optional[str] = None
        after_quality: Optional[dict[str, Any]] = None
        after_read_attempts: list[dict[str, Any]] = []
        movement_started = False
        rollback: dict[str, Any] = {
            "attempted": False,
            "succeeded": None,
        }

        try:
            requested = {
                "x": _finite_float(tx, "tx"),
                "y": _finite_float(ty, "ty"),
                "z": _finite_float(tz, "tz"),
            }
            tolerance = _finite_float(
                tolerance_mm,
                "tolerance_mm",
            )
            if tolerance <= 0.0:
                raise ValueError(
                    "tolerance_mm must be greater than zero."
                )

            requested_occurrence = _positive_integer(
                occurrence,
                "occurrence",
            )

            application = _ensure_connected(conn)
            document, product = (
                _require_active_product_document(conn)
            )
            (
                component_index,
                component,
                resolution,
            ) = _resolve_direct_component(
                product,
                component_name,
                requested_occurrence,
                bool(require_unique),
            )

            (
                positioning_preparation,
                preparation_warnings,
            ) = _prepare_component_for_positioning(
                product,
                component,
            )
            warnings.extend(preparation_warnings)

            (
                before_matrix,
                before_read_method,
                before_quality,
                before_read_attempts,
            ) = _read_position_matrix(component, application)
            document_saved_before = _document_saved(document)

            no_change_requested = all(
                abs(value) <= tolerance
                for value in requested.values()
            )

            if no_change_requested:
                return _success(
                    {
                        "component": resolution,
                        "positioning_preparation": (
                            positioning_preparation
                        ),
                        "requested_translation_delta_mm": (
                            requested
                        ),
                        "actual_translation_delta_mm": {
                            "x": 0.0,
                            "y": 0.0,
                            "z": 0.0,
                        },
                        "tolerance_mm": tolerance,
                        "no_change_requested": True,
                        "position_before": _matrix_payload(
                            before_matrix,
                            read_method=before_read_method,
                            quality=before_quality,
                            read_attempts=before_read_attempts,
                        ),
                        "position_after": _matrix_payload(
                            before_matrix,
                            read_method=before_read_method,
                            quality=before_quality,
                            read_attempts=before_read_attempts,
                        ),
                        "movement_method": None,
                        "movement_started": False,
                        "movement_verified": True,
                        "rotation_unchanged": True,
                        "rollback": rollback,
                        "feature_created": False,
                        "model_modified": False,
                        "document_save_required": False,
                        "document_saved_before": (
                            document_saved_before
                        ),
                        "document_saved_after": (
                            _document_saved(document)
                        ),
                    },
                    warnings,
                )

            delta_matrix = _identity_translation_matrix(
                requested["x"],
                requested["y"],
                requested["z"],
            )

            (
                movement_method,
                method_warnings,
                movement_apply_attempts,
            ) = _apply_move(
                component,
                delta_matrix,
                before_matrix,
                application,
            )
            warnings.extend(method_warnings)
            movement_started = True

            warnings.extend(_update_product(product))
            warnings.extend(_refresh_display(conn))

            (
                after_matrix,
                after_read_method,
                after_quality,
                after_read_attempts,
            ) = _read_position_matrix(component, application)
            actual_delta = _vector_difference(
                after_matrix,
                before_matrix,
            )
            rotation_difference = _max_rotation_difference(
                after_matrix,
                before_matrix,
            )
            translation_verified = _translation_matches(
                actual_delta,
                requested,
                tolerance,
            )
            rotation_unchanged = (
                rotation_difference <= 1e-9
            )
            movement_verified = bool(
                translation_verified
                and rotation_unchanged
            )

            if movement_verified:
                return _success(
                    {
                        "component": resolution,
                        "positioning_preparation": (
                            positioning_preparation
                        ),
                        "component_index": component_index,
                        "requested_translation_delta_mm": (
                            requested
                        ),
                        "actual_translation_delta_mm": (
                            actual_delta
                        ),
                        "translation_error_mm": {
                            axis: (
                                actual_delta[axis]
                                - requested[axis]
                            )
                            for axis in ("x", "y", "z")
                        },
                        "tolerance_mm": tolerance,
                        "no_change_requested": False,
                        "position_before": _matrix_payload(
                            before_matrix,
                            read_method=before_read_method,
                            quality=before_quality,
                            read_attempts=before_read_attempts,
                        ),
                        "position_after": _matrix_payload(
                            after_matrix,
                            read_method=after_read_method,
                            quality=after_quality,
                            read_attempts=after_read_attempts,
                        ),
                        "movement_method": movement_method,
                        "movement_apply_attempts": (
                            movement_apply_attempts
                        ),
                        "movement_started": True,
                        "movement_verified": True,
                        "rotation_unchanged": True,
                        "max_rotation_component_change": (
                            rotation_difference
                        ),
                        "rollback": rollback,
                        "constraint_note": (
                            "The move was accepted and verified. "
                            "Assembly constraints may reposition a "
                            "component on a later update."
                        ),
                        "feature_created": False,
                        "model_modified": True,
                        "document_save_required": True,
                        "document_saved_before": (
                            document_saved_before
                        ),
                        "document_saved_after": (
                            _document_saved(document)
                        ),
                    },
                    warnings,
                )

            actual_movement_magnitude = max(
                abs(actual_delta["x"]),
                abs(actual_delta["y"]),
                abs(actual_delta["z"]),
            )

            if actual_movement_magnitude > tolerance:
                (
                    rollback_succeeded,
                    rollback,
                    rollback_warnings,
                ) = _attempt_translation_rollback(
                    component,
                    actual_delta,
                    product,
                    tolerance,
                    application,
                )
                warnings.extend(rollback_warnings)
                warnings.extend(_refresh_display(conn))
            else:
                rollback_succeeded = True
                rollback = {
                    "attempted": False,
                    "succeeded": True,
                    "reason": (
                        "No measurable component translation "
                        "occurred."
                    ),
                }

            constraint_message = (
                "CATIA did not retain the requested component "
                "translation after update and readback. The component "
                "may be fixed, constrained, not independently movable, "
                "or affected by assembly update rules."
            )

            return _error(
                constraint_message,
                data={
                    "component": resolution,
                    "positioning_preparation": (
                        positioning_preparation
                    ),
                    "component_index": component_index,
                    "requested_translation_delta_mm": requested,
                    "actual_translation_delta_mm": actual_delta,
                    "translation_error_mm": {
                        axis: (
                            actual_delta[axis]
                            - requested[axis]
                        )
                        for axis in ("x", "y", "z")
                    },
                    "tolerance_mm": tolerance,
                    "position_before": _matrix_payload(
                        before_matrix,
                        read_method=before_read_method,
                        quality=before_quality,
                        read_attempts=before_read_attempts,
                    ),
                    "position_after_attempt": _matrix_payload(
                        after_matrix,
                        read_method=after_read_method,
                        quality=after_quality,
                        read_attempts=after_read_attempts,
                    ),
                    "movement_method": movement_method,
                    "movement_apply_attempts": (
                        movement_apply_attempts
                    ),
                    "movement_started": True,
                    "movement_verified": False,
                    "rotation_unchanged": rotation_unchanged,
                    "max_rotation_component_change": (
                        rotation_difference
                    ),
                    "rollback": rollback,
                    "feature_created": False,
                    "model_modified": (
                        rollback_succeeded is not True
                    ),
                    "document_save_required": (
                        rollback_succeeded is not True
                    ),
                    "document_saved_before": (
                        document_saved_before
                    ),
                    "document_saved_after": (
                        _document_saved(document)
                    ),
                },
                warnings=warnings,
                status=(
                    "error"
                    if rollback_succeeded
                    else "partial_success"
                ),
            )
        except Exception as exc:
            modified = bool(
                movement_started
                and before_matrix is not None
                and after_matrix is not None
                and rollback.get("succeeded") is not True
            )

            return _error(
                str(exc),
                data={
                    "component": resolution,
                    "positioning_preparation": (
                        positioning_preparation
                    ),
                    "requested_translation_delta_mm": {
                        "x": tx,
                        "y": ty,
                        "z": tz,
                    },
                    "occurrence": occurrence,
                    "require_unique": require_unique,
                    "tolerance_mm": tolerance_mm,
                    "position_before": (
                        _matrix_payload(
                            before_matrix,
                            read_method=before_read_method,
                            quality=before_quality,
                            read_attempts=before_read_attempts,
                        )
                        if before_matrix is not None
                        else None
                    ),
                    "position_after_attempt": (
                        _matrix_payload(
                            after_matrix,
                            read_method=after_read_method,
                            quality=after_quality,
                            read_attempts=after_read_attempts,
                        )
                        if after_matrix is not None
                        else None
                    ),
                    "movement_method": movement_method,
                    "movement_apply_attempts": (
                        movement_apply_attempts
                    ),
                    "movement_started": movement_started,
                    "movement_verified": False,
                    "rollback": rollback,
                    "feature_created": False,
                    "model_modified": modified,
                    "document_save_required": modified,
                },
                warnings=warnings,
                status=(
                    "partial_success"
                    if modified
                    else "error"
                ),
            )

    names.append("catia_move_component")

    @mcp.tool()
    def catia_rotate_component(
        component_name: str,
        axis: str = "z",
        angle_deg: float = 0.0,
        coordinate_system: str = "local",
        occurrence: int = 1,
        require_unique: bool = True,
        angle_tolerance_deg: float = 0.001,
        translation_tolerance_mm: float = 0.001,
    ) -> dict[str, Any]:
        """Incrementally rotate one direct component about its own origin.

        axis may be x, y, or z.

        coordinate_system:
        - local: rotate around the component's current local axis.
        - parent: rotate around the active parent CATProduct axis.

        Positive angles follow the right-hand rule. The component origin
        remains fixed, so a successful rotation does not change translation.
        Multi-axis orientations should be built with consecutive calls,
        making the rotation order explicit.
        """

        warnings: list[str] = []
        document = None
        product = None
        component = None
        resolution: Optional[dict[str, Any]] = None
        positioning_preparation: Optional[dict[str, Any]] = None
        before_matrix: Optional[list[float]] = None
        after_matrix: Optional[list[float]] = None
        target_matrix: Optional[list[float]] = None
        before_read_method: Optional[str] = None
        before_quality: Optional[dict[str, Any]] = None
        before_read_attempts: list[dict[str, Any]] = []
        after_read_method: Optional[str] = None
        after_quality: Optional[dict[str, Any]] = None
        after_read_attempts: list[dict[str, Any]] = []
        rotation_method: Optional[str] = None
        rotation_apply_attempts: list[dict[str, Any]] = []
        rotation_started = False
        rollback: dict[str, Any] = {
            "attempted": False,
            "succeeded": None,
        }

        try:
            requested_axis = _normalise_rotation_axis(axis)
            requested_coordinate_system = (
                _normalise_coordinate_system(
                    coordinate_system
                )
            )
            requested_angle = _finite_float(
                angle_deg,
                "angle_deg",
            )
            effective_angle = _normalise_signed_angle_deg(
                requested_angle
            )
            angle_tolerance = _finite_float(
                angle_tolerance_deg,
                "angle_tolerance_deg",
            )
            translation_tolerance = _finite_float(
                translation_tolerance_mm,
                "translation_tolerance_mm",
            )
            if angle_tolerance <= 0.0:
                raise ValueError(
                    "angle_tolerance_deg must be greater than zero."
                )
            if translation_tolerance <= 0.0:
                raise ValueError(
                    "translation_tolerance_mm must be greater "
                    "than zero."
                )

            requested_occurrence = _positive_integer(
                occurrence,
                "occurrence",
            )
            rotation_matrix_tolerance = max(
                1e-10,
                2.0 * math.sin(
                    math.radians(angle_tolerance) / 2.0
                ),
            )

            application = _ensure_connected(conn)
            document, product = (
                _require_active_product_document(conn)
            )
            (
                component_index,
                component,
                resolution,
            ) = _resolve_direct_component(
                product,
                component_name,
                requested_occurrence,
                bool(require_unique),
            )

            (
                positioning_preparation,
                preparation_warnings,
            ) = _prepare_component_for_positioning(
                product,
                component,
            )
            warnings.extend(preparation_warnings)

            (
                before_matrix,
                before_read_method,
                before_quality,
                before_read_attempts,
            ) = _read_position_matrix(
                component,
                application,
            )
            document_saved_before = _document_saved(document)
            axis_vector_parent = (
                _rotation_axis_vector_in_parent(
                    before_matrix,
                    requested_axis,
                    requested_coordinate_system,
                )
            )

            no_change_requested = bool(
                abs(effective_angle) <= angle_tolerance
            )

            if no_change_requested:
                return _success(
                    {
                        "component": resolution,
                        "component_index": component_index,
                        "positioning_preparation": (
                            positioning_preparation
                        ),
                        "requested_rotation": {
                            "axis": requested_axis,
                            "angle_deg": requested_angle,
                            "effective_angle_deg": (
                                effective_angle
                            ),
                            "coordinate_system": (
                                requested_coordinate_system
                            ),
                            "pivot": "component_origin",
                            "positive_direction": (
                                "right_hand_rule"
                            ),
                            "axis_vector_in_parent_before": (
                                axis_vector_parent
                            ),
                        },
                        "actual_rotation": {
                            "signed_angle_deg": 0.0,
                            "angle_error_deg": 0.0,
                        },
                        "angle_tolerance_deg": angle_tolerance,
                        "rotation_matrix_tolerance": (
                            rotation_matrix_tolerance
                        ),
                        "translation_tolerance_mm": (
                            translation_tolerance
                        ),
                        "no_change_requested": True,
                        "position_before": _matrix_payload(
                            before_matrix,
                            read_method=before_read_method,
                            quality=before_quality,
                            read_attempts=before_read_attempts,
                        ),
                        "target_position": _matrix_payload(
                            before_matrix,
                            quality=before_quality,
                        ),
                        "position_after": _matrix_payload(
                            before_matrix,
                            read_method=before_read_method,
                            quality=before_quality,
                            read_attempts=before_read_attempts,
                        ),
                        "rotation_method": None,
                        "rotation_apply_attempts": [],
                        "rotation_started": False,
                        "rotation_verified": True,
                        "translation_unchanged": True,
                        "rollback": rollback,
                        "feature_created": False,
                        "model_modified": False,
                        "document_save_required": False,
                        "document_saved_before": (
                            document_saved_before
                        ),
                        "document_saved_after": (
                            _document_saved(document)
                        ),
                    },
                    warnings,
                )

            target_matrix, requested_delta_rows = (
                _compose_rotation_target(
                    before_matrix,
                    requested_axis,
                    effective_angle,
                    requested_coordinate_system,
                )
            )
            target_quality = _validate_position_matrix(
                target_matrix
            )

            (
                rotation_method,
                method_warnings,
                rotation_apply_attempts,
            ) = _apply_absolute_position(
                component,
                target_matrix,
                application,
            )
            warnings.extend(method_warnings)
            rotation_started = True

            warnings.extend(_update_product(product))
            warnings.extend(_refresh_display(conn))

            (
                after_matrix,
                after_read_method,
                after_quality,
                after_read_attempts,
            ) = _read_position_matrix(
                component,
                application,
            )

            relative_rows = _relative_rotation_rows(
                before_matrix,
                after_matrix,
                requested_coordinate_system,
            )
            actual_angle = _signed_axis_angle_deg(
                relative_rows,
                requested_axis,
            )
            angle_error = _wrapped_angle_error_deg(
                actual_angle,
                effective_angle,
            )

            max_target_rotation_error = (
                _max_matrix_component_difference(
                    after_matrix,
                    target_matrix,
                    0,
                    9,
                )
            )
            max_translation_change = (
                _max_matrix_component_difference(
                    after_matrix,
                    before_matrix,
                    9,
                    12,
                )
            )
            max_translation_target_error = (
                _max_matrix_component_difference(
                    after_matrix,
                    target_matrix,
                    9,
                    12,
                )
            )

            translation_unchanged = bool(
                max_translation_change
                <= translation_tolerance
            )
            rotation_verified = bool(
                max_target_rotation_error
                <= rotation_matrix_tolerance
                and abs(angle_error) <= angle_tolerance
                and translation_unchanged
            )

            result_data = {
                "component": resolution,
                "component_index": component_index,
                "positioning_preparation": (
                    positioning_preparation
                ),
                "requested_rotation": {
                    "axis": requested_axis,
                    "angle_deg": requested_angle,
                    "effective_angle_deg": effective_angle,
                    "coordinate_system": (
                        requested_coordinate_system
                    ),
                    "pivot": "component_origin",
                    "positive_direction": "right_hand_rule",
                    "axis_vector_in_parent_before": (
                        axis_vector_parent
                    ),
                    "delta_rotation_matrix_rows": (
                        requested_delta_rows
                    ),
                },
                "actual_rotation": {
                    "signed_angle_deg": actual_angle,
                    "angle_error_deg": angle_error,
                    "relative_rotation_matrix_rows": (
                        relative_rows
                    ),
                },
                "angle_tolerance_deg": angle_tolerance,
                "rotation_matrix_tolerance": (
                    rotation_matrix_tolerance
                ),
                "translation_tolerance_mm": (
                    translation_tolerance
                ),
                "max_target_rotation_component_error": (
                    max_target_rotation_error
                ),
                "max_translation_change_mm": (
                    max_translation_change
                ),
                "max_translation_target_error_mm": (
                    max_translation_target_error
                ),
                "no_change_requested": False,
                "position_before": _matrix_payload(
                    before_matrix,
                    read_method=before_read_method,
                    quality=before_quality,
                    read_attempts=before_read_attempts,
                ),
                "target_position": _matrix_payload(
                    target_matrix,
                    quality=target_quality,
                ),
                "position_after": _matrix_payload(
                    after_matrix,
                    read_method=after_read_method,
                    quality=after_quality,
                    read_attempts=after_read_attempts,
                ),
                "rotation_method": rotation_method,
                "rotation_apply_attempts": (
                    rotation_apply_attempts
                ),
                "rotation_started": True,
                "rotation_verified": rotation_verified,
                "translation_unchanged": (
                    translation_unchanged
                ),
                "rollback": rollback,
                "feature_created": False,
                "model_modified": rotation_verified,
                "document_save_required": rotation_verified,
                "document_saved_before": (
                    document_saved_before
                ),
                "document_saved_after": (
                    _document_saved(document)
                ),
            }

            if rotation_verified:
                result_data["constraint_note"] = (
                    "The rotation was applied and verified. "
                    "Assembly constraints may reposition the "
                    "component on a later update."
                )
                return _success(result_data, warnings)

            actual_rotation_change = (
                _max_matrix_component_difference(
                    after_matrix,
                    before_matrix,
                    0,
                    9,
                )
            )
            actual_translation_change = (
                _max_matrix_component_difference(
                    after_matrix,
                    before_matrix,
                    9,
                    12,
                )
            )
            measurable_change = bool(
                actual_rotation_change
                > rotation_matrix_tolerance
                or actual_translation_change
                > translation_tolerance
            )

            if measurable_change:
                (
                    rollback_succeeded,
                    rollback,
                    rollback_warnings,
                ) = _attempt_position_matrix_rollback(
                    component,
                    before_matrix,
                    product,
                    application,
                    rotation_matrix_tolerance,
                    translation_tolerance,
                )
                warnings.extend(rollback_warnings)
                warnings.extend(_refresh_display(conn))
            else:
                rollback_succeeded = True
                rollback = {
                    "attempted": False,
                    "succeeded": True,
                    "reason": (
                        "No measurable component rotation or "
                        "translation occurred."
                    ),
                }

            result_data["rollback"] = rollback
            result_data["model_modified"] = (
                rollback_succeeded is not True
            )
            result_data["document_save_required"] = (
                rollback_succeeded is not True
            )

            return _error(
                "CATIA did not retain the requested component "
                "rotation after update and readback. The component "
                "may be fixed, constrained, not independently "
                "rotatable, or affected by assembly update rules.",
                data=result_data,
                warnings=warnings,
                status=(
                    "error"
                    if rollback_succeeded
                    else "partial_success"
                ),
            )
        except Exception as exc:
            modified = bool(
                rotation_started
                and before_matrix is not None
                and after_matrix is not None
                and rollback.get("succeeded") is not True
            )

            return _error(
                str(exc),
                data={
                    "component": resolution,
                    "positioning_preparation": (
                        positioning_preparation
                    ),
                    "requested_rotation": {
                        "axis": axis,
                        "angle_deg": angle_deg,
                        "coordinate_system": coordinate_system,
                        "pivot": "component_origin",
                    },
                    "occurrence": occurrence,
                    "require_unique": require_unique,
                    "angle_tolerance_deg": (
                        angle_tolerance_deg
                    ),
                    "translation_tolerance_mm": (
                        translation_tolerance_mm
                    ),
                    "position_before": (
                        _matrix_payload(
                            before_matrix,
                            read_method=before_read_method,
                            quality=before_quality,
                            read_attempts=before_read_attempts,
                        )
                        if before_matrix is not None
                        else None
                    ),
                    "target_position": (
                        _matrix_payload(target_matrix)
                        if target_matrix is not None
                        else None
                    ),
                    "position_after_attempt": (
                        _matrix_payload(
                            after_matrix,
                            read_method=after_read_method,
                            quality=after_quality,
                            read_attempts=after_read_attempts,
                        )
                        if after_matrix is not None
                        else None
                    ),
                    "rotation_method": rotation_method,
                    "rotation_apply_attempts": (
                        rotation_apply_attempts
                    ),
                    "rotation_started": rotation_started,
                    "rotation_verified": False,
                    "rollback": rollback,
                    "feature_created": False,
                    "model_modified": modified,
                    "document_save_required": modified,
                },
                warnings=warnings,
                status=(
                    "partial_success"
                    if modified
                    else "error"
                ),
            )

    names.append("catia_rotate_component")

    return names


# ---------------------------------------------------------------------------
# Registry Entry Point (Required by registry.py)
# ---------------------------------------------------------------------------
