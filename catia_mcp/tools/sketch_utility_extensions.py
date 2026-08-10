"""
sketch_utility_extensions.py
============================

CATIA V5 MCP sketch utility tools.

The CATIA V5 Factory2D Automation interface exposes primitive-creation methods,
but it does not expose CreateSymmetry, CreateTrim, CreateTranslate, or
CreateRotate.  These tools therefore read the source geometry, calculate the
required 2D transformation, and reconstruct supported geometry with Factory2D.

Supported source geometry:
- Point2D
- Line2D / finite line segment
- Circle2D closed circle
- Circle2D arc

Trim and extend currently support line-to-line operations only.
"""

from __future__ import annotations

import json
import logging
import math
import os
import threading
from dataclasses import asdict, dataclass, field
from typing import Any, Callable, Dict, List, Sequence, Tuple

logger = logging.getLogger(__name__)

# CatConstraintType values verified against the CATIA V5 Automation enum.
_CAT_CST_TYPE_ON = 2
_CAT_CST_TYPE_SYMMETRY = 15

# CatGeometricType values from CATIA V5 Automation.
_CAT_GEO_TYPE_UNKNOWN = 0
_CAT_GEO_TYPE_AXIS_2D = 1
_CAT_GEO_TYPE_POINT_2D = 2
_CAT_GEO_TYPE_LINE_2D = 3
_CAT_GEO_TYPE_CONTROL_POINT_2D = 4
_CAT_GEO_TYPE_CIRCLE_2D = 5

_EPSILON = 1.0e-9
_ANGLE_EPSILON = 1.0e-8

# GeometricElements.Item() is commonly wrapped by pywin32 as the base
# MecModInterfaces.GeometricElement interface.  Concrete sketch members may be
# unavailable on that Python wrapper even though CATIA's internal automation
# runtime can resolve the object's actual Line2D/Circle2D/Point2D interface.
# SystemService.Evaluate is therefore the primary geometry bridge.  The
# SketcherInterfaces type-library/QueryInterface path remains an optional
# fallback only and must never be a hard prerequisite for utility operations.
_SKETCHER_TYPELIB_SPEC: Any = None
_SKETCHER_TYPELIB_LOOKUP_DONE = False


@dataclass
class SketchOperationResult:
    """Structured result returned by every sketch utility operation."""

    operation: str = ""
    status: str = "success"
    elements_created: int = 0
    elements_modified: int = 0
    originals_deleted: int = 0
    constraints_added: int = 0
    constraints_removed: int = 0
    constraint_status: str = "unknown"
    element_names: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    details: Dict[str, Any] = field(default_factory=dict)
    sketch_closed: bool = False
    update_succeeded: bool = False
    document_unit: str = "mm"


@dataclass(frozen=True)
class GeometrySnapshot:
    """COM-independent representation of supported 2D sketch geometry."""

    kind: str
    name: str
    data: Dict[str, Any]
    interface_details: Dict[str, Any] = field(default_factory=dict)


def _result_dict(result: SketchOperationResult) -> Dict[str, Any]:
    return asdict(result)


def _normalise_text(value: Any) -> str:
    return str(value).strip().lower()


def _describe_exception(exc: BaseException) -> str:
    """Return pywin32/COM exception details without losing HRESULT context."""

    parts = [
        f"type={type(exc).__name__}",
        f"message={exc}",
    ]

    hresult = getattr(exc, "hresult", None)
    if hresult is not None:
        try:
            parts.append(f"hresult=0x{int(hresult) & 0xFFFFFFFF:08X}")
        except Exception:
            parts.append(f"hresult={hresult!r}")

    excepinfo = getattr(exc, "excepinfo", None)
    if excepinfo:
        parts.append(f"excepinfo={excepinfo!r}")

    argerror = getattr(exc, "argerror", None)
    if argerror is not None:
        parts.append(f"argerror={argerror!r}")

    return ", ".join(parts)


def _describe_com_object(obj: Any) -> Dict[str, Any]:
    """Describe the wrapper received by the MCP worker without probing geometry."""

    description: Dict[str, Any] = {
        "python_type": type(obj).__name__,
        "python_module": type(obj).__module__,
    }

    try:
        description["repr"] = repr(obj)
    except Exception as exc:
        description["repr_error"] = _describe_exception(exc)

    try:
        description["name"] = str(obj.Name)
    except Exception as exc:
        description["name_error"] = _describe_exception(exc)

    try:
        description["has_oleobj"] = getattr(obj, "_oleobj_", None) is not None
    except Exception as exc:
        description["oleobj_error"] = _describe_exception(exc)

    return description



def _parse_typelib_version(value: Any) -> int:
    """Parse the registry's inconsistent type-library version representation."""

    if isinstance(value, int):
        return value
    raw = str(value).strip()
    for base in (16, 10):
        try:
            return int(raw, base)
        except (TypeError, ValueError):
            continue
    return 0


def _find_sketcher_typelib() -> Any:
    """Locate the CATIA V5 SketcherInterfaces type library registration."""

    global _SKETCHER_TYPELIB_SPEC, _SKETCHER_TYPELIB_LOOKUP_DONE

    if _SKETCHER_TYPELIB_LOOKUP_DONE:
        if _SKETCHER_TYPELIB_SPEC is None:
            raise RuntimeError(
                "CATIA V5 SketcherInterfaces Object Library was not found in "
                "the registered COM type libraries."
            )
        return _SKETCHER_TYPELIB_SPEC

    _SKETCHER_TYPELIB_LOOKUP_DONE = True

    try:
        from win32com.client import selecttlb  # type: ignore
    except Exception as exc:
        raise RuntimeError(
            "win32com.client.selecttlb is unavailable: "
            f"{_describe_exception(exc)}"
        ) from exc

    candidates: List[Any] = []
    exact_description = "CATIA V5 SketcherInterfaces Object Library"

    try:
        candidates.extend(selecttlb.FindTlbsWithDescription(exact_description))
    except Exception:
        pass

    if not candidates:
        try:
            for spec in selecttlb.EnumTlbs():
                searchable = " ".join(
                    str(getattr(spec, attribute, "") or "")
                    for attribute in ("desc", "ver_desc", "dll")
                ).casefold()
                if "catia" in searchable and "sketcherinterfaces" in searchable:
                    candidates.append(spec)
        except Exception as exc:
            raise RuntimeError(
                "Cannot enumerate registered COM type libraries: "
                f"{_describe_exception(exc)}"
            ) from exc

    if not candidates:
        available_catia: List[str] = []
        try:
            for spec in selecttlb.EnumTlbs():
                description = str(getattr(spec, "desc", "") or "")
                if "catia" in description.casefold():
                    available_catia.append(description)
        except Exception:
            pass

        suffix = ""
        if available_catia:
            suffix = " Registered CATIA libraries include: " + ", ".join(
                sorted(set(available_catia))[:20]
            )
        raise RuntimeError(
            "CATIA V5 SketcherInterfaces Object Library was not found." + suffix
        )

    candidates.sort(
        key=lambda spec: (
            _parse_typelib_version(getattr(spec, "major", 0)),
            _parse_typelib_version(getattr(spec, "minor", 0)),
        )
    )
    _SKETCHER_TYPELIB_SPEC = candidates[-1]
    return _SKETCHER_TYPELIB_SPEC


def _typelib_description(spec: Any) -> Dict[str, Any]:
    return {
        "description": str(getattr(spec, "desc", "") or ""),
        "version_description": str(getattr(spec, "ver_desc", "") or ""),
        "clsid": str(getattr(spec, "clsid", "") or ""),
        "lcid": getattr(spec, "lcid", None),
        "major": getattr(spec, "major", None),
        "minor": getattr(spec, "minor", None),
        "dll": str(getattr(spec, "dll", "") or ""),
    }


def _interface_matches(obj: Any, interface_name: str) -> bool:
    target = interface_name.casefold()
    try:
        return any(base.__name__.casefold() == target for base in type(obj).__mro__)
    except Exception:
        return type(obj).__name__.casefold() == target


def _get_sketcher_interface_metadata(interface_name: str) -> Tuple[Any, Any, Any]:
    """Return (typelib_spec, interface_iid, generated_wrapper_class)."""

    typelib_spec = _find_sketcher_typelib()
    try:
        from win32com.client import gencache  # type: ignore
    except Exception as exc:
        raise RuntimeError(
            "win32com.client.gencache is unavailable: "
            f"{_describe_exception(exc)}"
        ) from exc

    module = gencache.MakeModuleForTypelib(
        typelib_spec.clsid,
        typelib_spec.lcid,
        int(str(typelib_spec.major), 16),
        int(str(typelib_spec.minor), 16),
    )
    target_iid = module.NamesToIIDMap.get(interface_name)
    if target_iid is None:
        raise RuntimeError(
            f"Interface '{interface_name}' is absent from the generated "
            "CATIA SketcherInterfaces module."
        )

    interface_module = gencache.GetModuleForCLSID(target_iid)
    if interface_module is None:
        raise RuntimeError(
            f"No generated wrapper module exists for {interface_name}."
        )

    target_class = getattr(interface_module, interface_name)
    target_class = getattr(target_class, "default_interface", target_class)
    return typelib_spec, target_iid, target_class


def _cast_geometry_interface(
    element: Any,
    interface_name: str,
) -> Tuple[Any, Dict[str, Any]]:
    """Obtain and validate a concrete CATIA Sketcher interface.

    ``win32com.client.CastTo`` can produce a generated Python wrapper even when
    the underlying COM object does not support the requested interface because
    pywin32 deliberately tolerates ``E_NOINTERFACE`` in its generated wrapper
    constructor.  That behavior made a Line2D appear to cast successfully to
    Circle2D.  This function performs an explicit QueryInterface first and only
    constructs the target wrapper after COM confirms support for the IID.
    """

    diagnostics: Dict[str, Any] = {
        "source": _describe_com_object(element),
        "target_interface": interface_name,
        "attempts": [],
    }

    if _interface_matches(element, interface_name):
        diagnostics.update(
            {
                "strategy": "already_specific_interface",
                "result_python_type": type(element).__name__,
                "result_python_module": type(element).__module__,
                "query_interface_validated": True,
            }
        )
        return element, diagnostics

    try:
        import pythoncom  # type: ignore

        typelib_spec, target_iid, target_class = _get_sketcher_interface_metadata(
            interface_name
        )
        diagnostics["sketcher_typelib"] = _typelib_description(typelib_spec)
        diagnostics["target_iid"] = str(target_iid)

        raw_dispatch = getattr(element, "_oleobj_", element)
        queried_dispatch = raw_dispatch.QueryInterface(
            target_iid,
            pythoncom.IID_IDispatch,
        )
        diagnostics["attempts"].append(
            {
                "strategy": "explicit_query_interface",
                "status": "success",
            }
        )

        casted = target_class(queried_dispatch)
        diagnostics.update(
            {
                "strategy": "explicit_query_interface",
                "result_python_type": type(casted).__name__,
                "result_python_module": type(casted).__module__,
                "query_interface_validated": True,
            }
        )
        return casted, diagnostics
    except Exception as exc:
        diagnostics["attempts"].append(
            {
                "strategy": "explicit_query_interface",
                "status": "error",
                "error": _describe_exception(exc),
            }
        )
        raise RuntimeError(
            f"Cannot query sketch element '{_get_element_name(element)}' as "
            f"{interface_name}. Diagnostics: "
            f"{json.dumps(diagnostics, ensure_ascii=False, default=str)}"
        ) from exc


def _get_geometric_type(element: Any) -> Tuple[int | None, Dict[str, Any]]:
    """Read the language-independent CatGeometricType value."""

    diagnostics: Dict[str, Any] = {}
    try:
        value = int(element.GeometricType)
        diagnostics.update(
            {
                "strategy": "GeometricType_property",
                "value": value,
            }
        )
        return value, diagnostics
    except Exception as exc:
        diagnostics.update(
            {
                "strategy": "GeometricType_property",
                "error": _describe_exception(exc),
            }
        )
        return None, diagnostics

def _execution_context(ctx: Any) -> Dict[str, Any]:
    """Capture MCP worker/thread facts without changing COM initialization."""

    context: Dict[str, Any] = {
        "process_id": os.getpid(),
        "thread_id": threading.get_ident(),
        "thread_name": threading.current_thread().name,
    }

    try:
        app = ctx.conn.app
        context["app_python_type"] = type(app).__name__
        context["app_python_module"] = type(app).__module__
        context["app_has_oleobj"] = getattr(app, "_oleobj_", None) is not None
    except Exception as exc:
        context["app_description_error"] = _describe_exception(exc)

    # Diagnostic only: do not call CoInitialize here, because that would alter
    # the behavior we are trying to observe.
    try:
        import pythoncom  # type: ignore

        get_apartment_type = getattr(pythoncom, "CoGetApartmentType", None)
        if get_apartment_type is None:
            context["com_apartment"] = "CoGetApartmentType unavailable"
        else:
            try:
                context["com_apartment"] = repr(get_apartment_type())
            except Exception as exc:
                context["com_apartment_error"] = _describe_exception(exc)
    except Exception as exc:
        context["pythoncom_error"] = _describe_exception(exc)

    return context


def _coerce_vector(value: Any, expected_size: int) -> Tuple[float, ...] | None:
    """Convert COM/VARIANT/nested sequence output into a finite float tuple."""

    if value is None:
        return None

    # win32com.client.VARIANT exposes the wrapped SAFEARRAY through ``value``.
    if hasattr(value, "value"):
        try:
            value = value.value
        except Exception:
            pass

    if isinstance(value, (list, tuple)):
        # First prefer a flat sequence containing the requested values.
        if len(value) >= expected_size:
            try:
                candidate = tuple(float(value[index]) for index in range(expected_size))
            except (TypeError, ValueError):
                candidate = None
            if candidate is not None and all(math.isfinite(item) for item in candidate):
                return candidate

        # Generated COM wrappers sometimes nest the out-array in another tuple,
        # or return additional values beside it. Search nested members safely.
        for item in value:
            nested = _coerce_vector(item, expected_size)
            if nested is not None:
                return nested

    return None


def _system_service_read_array(
    catia_app: Any,
    obj: Any,
    method_name: str,
    expected_size: int,
) -> Tuple[float, ...]:
    """Read CATSafeArrayVariant output through CATIA's VBScript evaluator.

    CATIA exposes several sketch methods as ``Sub Method(CATSafeArrayVariant)``.
    pywin32 cannot reliably marshal these output arrays in every generated or
    dynamic-dispatch configuration.  CATIA's own SystemService.Evaluate executes
    the native VBScript call inside CATIA and returns the populated array.
    """

    if catia_app is None:
        raise RuntimeError("CATIA Application is required for SystemService fallback.")

    bounds = expected_size - 1
    function_name = f"MCP_{method_name}_{expected_size}"
    script = (
        f"Public Function {function_name}(target)\n"
        f"  Dim values({bounds})\n"
        f"  target.{method_name} values\n"
        f"  {function_name} = values\n"
        "End Function"
    )

    try:
        result = catia_app.SystemService.Evaluate(
            script,
            0,  # CATVBScriptLanguage
            function_name,
            [obj],
        )
    except Exception as exc:
        raise RuntimeError(
            f"SystemService.Evaluate fallback for {method_name} failed: "
            f"{_describe_exception(exc)}"
        ) from exc

    values = _coerce_vector(result, expected_size)
    if values is None:
        raise RuntimeError(
            f"SystemService.Evaluate returned unusable {method_name} data: "
            f"{result!r}"
        )
    return values



def _system_service_evaluate(
    catia_app: Any,
    script: str,
    function_name: str,
    arguments: Sequence[Any],
) -> Any:
    """Execute a small CATIA-side VBScript function with consistent errors."""

    if catia_app is None:
        raise RuntimeError("CATIA Application is required for SystemService.Evaluate.")

    try:
        return catia_app.SystemService.Evaluate(
            script,
            0,  # CATVBScriptLanguage
            function_name,
            list(arguments),
        )
    except Exception as exc:
        raise RuntimeError(
            f"SystemService.Evaluate function {function_name} failed: "
            f"{_describe_exception(exc)}"
        ) from exc


def _system_service_read_point(
    catia_app: Any,
    element: Any,
) -> Tuple[Tuple[float, float], Dict[str, Any]]:
    """Read Point2D coordinates inside CATIA without a Python Point2D wrapper."""

    function_name = "MCP_ReadSketchPoint"
    script = (
        "Public Function MCP_ReadSketchPoint(target)\n"
        "  Dim values(1)\n"
        "  target.GetCoordinates values\n"
        "  MCP_ReadSketchPoint = values\n"
        "End Function"
    )
    raw = _system_service_evaluate(catia_app, script, function_name, [element])
    values = _coerce_vector(raw, 2)
    if values is None:
        raise RuntimeError(f"CATIA returned unusable Point2D data: {raw!r}")
    return (values[0], values[1]), {
        "strategy": "system_service_evaluate",
        "function": function_name,
        "query_interface_required": False,
        "values": [values[0], values[1]],
    }


def _system_service_read_line(
    catia_app: Any,
    element: Any,
) -> Tuple[Tuple[float, float, float, float], Dict[str, Any]]:
    """Read Line2D endpoints inside CATIA from a generic GeometricElement."""

    function_name = "MCP_ReadSketchLine"
    script = (
        "Public Function MCP_ReadSketchLine(target)\n"
        "  Dim values(3)\n"
        "  target.GetEndPoints values\n"
        "  MCP_ReadSketchLine = values\n"
        "End Function"
    )
    raw = _system_service_evaluate(catia_app, script, function_name, [element])
    values = _coerce_vector(raw, 4)
    if values is None:
        raise RuntimeError(f"CATIA returned unusable Line2D data: {raw!r}")
    return (values[0], values[1], values[2], values[3]), {
        "strategy": "system_service_evaluate",
        "function": function_name,
        "query_interface_required": False,
        "values": list(values),
    }


def _system_service_read_circle(
    catia_app: Any,
    element: Any,
) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    """Read Circle2D/arc data inside CATIA from a generic GeometricElement.

    The result array contains flags so optional Curve2D calls can fail without
    preventing a closed circle from being read:
      0..2  center_x, center_y, radius
      3     periodic flag
      4     parameter extents available
      5..6  parameter extents
      7     endpoints available
      8..11 start_x, start_y, end_x, end_y
    """

    function_name = "MCP_ReadSketchCircle"
    script = (
        "Public Function MCP_ReadSketchCircle(target)\n"
        "  Dim result(11)\n"
        "  Dim center(1)\n"
        "  Dim params(1)\n"
        "  Dim endpoints(3)\n"
        "  Dim periodicValue\n"
        "  target.GetCenter center\n"
        "  result(0) = center(0)\n"
        "  result(1) = center(1)\n"
        "  result(2) = target.Radius\n"
        "  result(3) = 0\n"
        "  result(4) = 0\n"
        "  result(7) = 0\n"
        "  On Error Resume Next\n"
        "  Err.Clear\n"
        "  periodicValue = target.IsPeriodic()\n"
        "  If Err.Number = 0 Then\n"
        "    If periodicValue Then result(3) = 1\n"
        "  End If\n"
        "  Err.Clear\n"
        "  target.GetParamExtents params\n"
        "  If Err.Number = 0 Then\n"
        "    result(4) = 1\n"
        "    result(5) = params(0)\n"
        "    result(6) = params(1)\n"
        "  End If\n"
        "  Err.Clear\n"
        "  target.GetEndPoints endpoints\n"
        "  If Err.Number = 0 Then\n"
        "    result(7) = 1\n"
        "    result(8) = endpoints(0)\n"
        "    result(9) = endpoints(1)\n"
        "    result(10) = endpoints(2)\n"
        "    result(11) = endpoints(3)\n"
        "  End If\n"
        "  On Error GoTo 0\n"
        "  MCP_ReadSketchCircle = result\n"
        "End Function"
    )
    raw = _system_service_evaluate(catia_app, script, function_name, [element])
    values = _coerce_vector(raw, 12)
    if values is None:
        raise RuntimeError(f"CATIA returned unusable Circle2D data: {raw!r}")

    data: Dict[str, Any] = {
        "center": (values[0], values[1]),
        "radius": values[2],
        "periodic": bool(round(values[3])),
        "params": None,
        "endpoints": None,
    }
    if bool(round(values[4])):
        data["params"] = (values[5], values[6])
    if bool(round(values[7])):
        data["endpoints"] = (values[8], values[9], values[10], values[11])

    return data, {
        "strategy": "system_service_evaluate",
        "function": function_name,
        "query_interface_required": False,
        "center": [values[0], values[1]],
        "radius": values[2],
        "periodic": data["periodic"],
        "parameter_extents_available": data["params"] is not None,
        "endpoints_available": data["endpoints"] is not None,
    }


def _system_service_move_line_endpoint(
    catia_app: Any,
    sketch: Any,
    source: Any,
    boundary: Any,
    endpoint_name: str,
    x_value: float,
    y_value: float,
) -> Tuple[Tuple[float, float], bool, Dict[str, Any]]:
    """Move a Line2D endpoint and optionally add its on-boundary constraint.

    All concrete Line2D/Point2D member access happens inside CATIA, so neither
    the SketcherInterfaces type library nor a Python Point2D wrapper is needed.
    """

    function_name = "MCP_MoveSketchLineEndpoint"
    script = (
        "Public Function MCP_MoveSketchLineEndpoint(sketch, source, boundary, endpointName, xValue, yValue)\n"
        "  Dim result(2)\n"
        "  Dim point\n"
        "  Dim coords(1)\n"
        "  Dim constraintObject\n"
        "  If LCase(CStr(endpointName)) = \"start\" Then\n"
        "    Set point = source.StartPoint\n"
        "  Else\n"
        "    Set point = source.EndPoint\n"
        "  End If\n"
        "  point.SetData CDbl(xValue), CDbl(yValue)\n"
        "  point.GetCoordinates coords\n"
        "  result(0) = coords(0)\n"
        "  result(1) = coords(1)\n"
        "  result(2) = 0\n"
        "  On Error Resume Next\n"
        "  Err.Clear\n"
        "  Set constraintObject = sketch.Constraints.AddBiEltCst(2, point, boundary)\n"
        "  If Err.Number = 0 Then result(2) = 1\n"
        "  On Error GoTo 0\n"
        "  MCP_MoveSketchLineEndpoint = result\n"
        "End Function"
    )
    raw = _system_service_evaluate(
        catia_app,
        script,
        function_name,
        [sketch, source, boundary, endpoint_name, float(x_value), float(y_value)],
    )
    values = _coerce_vector(raw, 3)
    if values is None:
        raise RuntimeError(f"CATIA returned unusable endpoint update data: {raw!r}")
    constraint_added = bool(round(values[2]))
    return (values[0], values[1]), constraint_added, {
        "strategy": "system_service_evaluate",
        "function": function_name,
        "query_interface_required": False,
        "endpoint": endpoint_name,
        "confirmed_coordinates": [values[0], values[1]],
        "constraint_added": constraint_added,
    }

def _read_out_array(
    obj: Any,
    method_name: str,
    expected_size: int,
    catia_app: Any = None,
) -> Tuple[float, ...]:
    """Read a CATSafeArrayVariant output across pywin32 dispatch modes."""

    errors: List[str] = []
    method: Any = None

    def _accept(value: Any, source: str) -> Tuple[float, ...] | None:
        values = _coerce_vector(value, expected_size)
        if values is not None:
            return values
        if value is not None:
            errors.append(f"{source}: unusable return {value!r}")
        return None

    try:
        method = getattr(obj, method_name)
    except Exception as exc:
        errors.append(
            f"wrapper has no usable {method_name}: {_describe_exception(exc)}"
        )

    if method is not None:
        # makepy wrappers may return the out-array directly.
        try:
            values = _accept(method(), "direct call")
            if values is not None:
                return values
        except Exception as exc:
            errors.append(f"direct call: {_describe_exception(exc)}")

        # Some wrappers require an explicit mutable argument.
        for seed_kind, seed in (
            ("python-list", [float("nan")] * expected_size),
            ("python-tuple", tuple(float("nan") for _ in range(expected_size))),
        ):
            try:
                returned = method(seed)
                values = _accept(returned, f"{seed_kind} return")
                if values is not None:
                    return values
                values = _coerce_vector(seed, expected_size)
                if values is not None:
                    return values
                errors.append(f"{seed_kind} buffer was not populated")
            except Exception as exc:
                errors.append(f"{seed_kind} call: {_describe_exception(exc)}")

        try:
            import pythoncom  # type: ignore
            from win32com.client import VARIANT  # type: ignore

            variant_values = (
                pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_R8,
                pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
                pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
            )
            for variant_type in variant_values:
                try:
                    buffer = VARIANT(
                        variant_type,
                        [float("nan")] * expected_size,
                    )
                    returned = method(buffer)
                    values = _accept(returned, f"VARIANT return {variant_type}")
                    if values is not None:
                        return values
                    values = _coerce_vector(buffer, expected_size)
                    if values is not None:
                        return values
                    errors.append(
                        f"VARIANT buffer {variant_type} was not populated"
                    )
                except Exception as exc:
                    errors.append(
                        f"VARIANT call {variant_type}: {_describe_exception(exc)}"
                    )
        except Exception as exc:
            errors.append(f"VARIANT unavailable: {_describe_exception(exc)}")

    # CATIA-native fallback. This is the most reliable path for
    # CATSafeArrayVariant output parameters and is the same approach used by
    # mature CATIA Python wrappers.
    try:
        return _system_service_read_array(
            catia_app,
            obj,
            method_name,
            expected_size,
        )
    except Exception as exc:
        errors.append(f"SystemService fallback: {_describe_exception(exc)}")

    raise RuntimeError(
        f"{method_name} did not return {expected_size} numeric value(s). "
        + "; ".join(errors)
    )


def _get_point_coordinates(
    point: Any,
    catia_app: Any = None,
) -> Tuple[float, float]:
    point_interface = point
    if not _interface_matches(point, "Point2D"):
        point_interface, _ = _cast_geometry_interface(point, "Point2D")
    values = _read_out_array(
        point_interface,
        "GetCoordinates",
        2,
        catia_app=catia_app,
    )
    return values[0], values[1]


def _set_point_coordinates(point: Any, x_value: float, y_value: float) -> None:
    """Point2D coordinates are modified with SetData."""

    point_interface = point
    if not _interface_matches(point, "Point2D"):
        point_interface, _ = _cast_geometry_interface(point, "Point2D")
    point_interface.SetData(float(x_value), float(y_value))


def _get_curve_endpoints(
    curve: Any,
    catia_app: Any = None,
) -> Tuple[float, float, float, float]:
    """Prefer Point2D endpoint properties, then CATSafeArrayVariant output."""

    property_error: BaseException | None = None
    try:
        start_point = curve.StartPoint
        end_point = curve.EndPoint
        start_x, start_y = _get_point_coordinates(start_point, catia_app)
        end_x, end_y = _get_point_coordinates(end_point, catia_app)
        return start_x, start_y, end_x, end_y
    except Exception as exc:
        property_error = exc

    try:
        values = _read_out_array(
            curve,
            "GetEndPoints",
            4,
            catia_app=catia_app,
        )
        return values[0], values[1], values[2], values[3]
    except Exception as array_error:
        raise RuntimeError(
            "Cannot read curve endpoints. "
            f"StartPoint/EndPoint error: [{_describe_exception(property_error)}]; "
            f"GetEndPoints error: [{_describe_exception(array_error)}]"
        ) from array_error


def _get_circle_center(
    circle: Any,
    catia_app: Any = None,
) -> Tuple[float, float]:
    """Prefer the CenterPoint property, then CATSafeArrayVariant GetCenter."""

    property_error: BaseException | None = None
    try:
        return _get_point_coordinates(circle.CenterPoint, catia_app)
    except Exception as exc:
        property_error = exc

    try:
        values = _read_out_array(
            circle,
            "GetCenter",
            2,
            catia_app=catia_app,
        )
        return values[0], values[1]
    except Exception as array_error:
        raise RuntimeError(
            "Cannot read circle center. "
            f"CenterPoint error: [{_describe_exception(property_error)}]; "
            f"GetCenter error: [{_describe_exception(array_error)}]"
        ) from array_error


def _get_curve_param_extents(
    curve: Any,
    catia_app: Any = None,
) -> Tuple[float, float] | None:
    try:
        values = _read_out_array(
            curve,
            "GetParamExtents",
            2,
            catia_app=catia_app,
        )
        return values[0], values[1]
    except Exception:
        return None

def _get_element_name(element: Any, fallback: str = "") -> str:
    try:
        value = str(element.Name).strip()
        if value:
            return value
    except Exception:
        pass
    return fallback


def _is_periodic_curve(curve: Any) -> bool:
    try:
        return bool(curve.IsPeriodic())
    except Exception:
        return False


def _snapshot_geometry(
    element: Any,
    fallback_name: str = "",
    catia_app: Any = None,
) -> GeometrySnapshot:
    """Read supported geometry by GeometricType with CATIA-side evaluation.

    GeometricType provides language-independent routing.  SystemService.Evaluate
    is the primary bridge because GeometricElements.Item() may expose only the
    base MecModInterfaces.GeometricElement wrapper in Python.  A concrete
    SketcherInterfaces QueryInterface is attempted only as a secondary fallback.
    """

    name = _get_element_name(element, fallback_name)
    geometric_type, type_details = _get_geometric_type(element)
    diagnostics: Dict[str, Any] = {
        "element_name": name or fallback_name,
        "object": _describe_com_object(element),
        "geometric_type": type_details,
        "attempts": [],
    }

    if geometric_type == _CAT_GEO_TYPE_LINE_2D:
        candidate_kind = "line"
        target_interface = "Line2D"
    elif geometric_type == _CAT_GEO_TYPE_CIRCLE_2D:
        candidate_kind = "circle"
        target_interface = "Circle2D"
    elif geometric_type in {
        _CAT_GEO_TYPE_POINT_2D,
        _CAT_GEO_TYPE_CONTROL_POINT_2D,
    }:
        candidate_kind = "point"
        target_interface = "Point2D"
    else:
        raise RuntimeError(
            f"Unsupported or unavailable CatGeometricType {geometric_type} for "
            f"'{name or fallback_name}'. Diagnostics: "
            f"{json.dumps(diagnostics, ensure_ascii=False, default=str)}"
        )

    # Primary path: let CATIA resolve the object's real sketch interface.
    try:
        if candidate_kind == "line":
            endpoints, read_details = _system_service_read_line(catia_app, element)
            x1, y1, x2, y2 = endpoints
            if math.hypot(x2 - x1, y2 - y1) <= _EPSILON:
                raise RuntimeError(f"Line '{name}' has zero length.")
            return GeometrySnapshot(
                kind="line",
                name=name,
                data={"start": (x1, y1), "end": (x2, y2)},
                interface_details={
                    "geometric_type": type_details,
                    "target_interface": target_interface,
                    "primary_access": read_details,
                    "query_interface_required": False,
                },
            )

        if candidate_kind == "point":
            point, read_details = _system_service_read_point(catia_app, element)
            return GeometrySnapshot(
                kind="point",
                name=name,
                data={"point": point},
                interface_details={
                    "geometric_type": type_details,
                    "target_interface": target_interface,
                    "primary_access": read_details,
                    "query_interface_required": False,
                },
            )

        circle_data, read_details = _system_service_read_circle(catia_app, element)
        center_x, center_y = circle_data["center"]
        radius = float(circle_data["radius"])
        params = circle_data["params"]
        endpoints = circle_data["endpoints"]
        periodic = bool(circle_data["periodic"])

        if radius <= _EPSILON:
            raise RuntimeError(f"Circle/arc '{name}' has a non-positive radius.")

        if periodic:
            return GeometrySnapshot(
                kind="circle",
                name=name,
                data={"center": (center_x, center_y), "radius": radius},
                interface_details={
                    "geometric_type": type_details,
                    "target_interface": target_interface,
                    "primary_access": read_details,
                    "query_interface_required": False,
                },
            )

        sweep: float | None = None
        if params is not None:
            sweep = float(params[1] - params[0])
            while sweep <= 0.0:
                sweep += 2.0 * math.pi

        if sweep is not None and sweep >= 2.0 * math.pi - _ANGLE_EPSILON:
            return GeometrySnapshot(
                kind="circle",
                name=name,
                data={"center": (center_x, center_y), "radius": radius},
                interface_details={
                    "geometric_type": type_details,
                    "target_interface": target_interface,
                    "primary_access": read_details,
                    "query_interface_required": False,
                },
            )

        if endpoints is None:
            # Closed circles in some CATIA dispatch configurations do not expose
            # IsPeriodic/GetParamExtents reliably.  If no endpoints exist, the
            # only valid Circle2D interpretation is a closed circle.
            return GeometrySnapshot(
                kind="circle",
                name=name,
                data={"center": (center_x, center_y), "radius": radius},
                interface_details={
                    "geometric_type": type_details,
                    "target_interface": target_interface,
                    "primary_access": read_details,
                    "closed_circle_inferred_from_missing_endpoints": True,
                    "query_interface_required": False,
                },
            )

        start_point = (endpoints[0], endpoints[1])
        end_point = (endpoints[2], endpoints[3])
        if sweep is None:
            start_angle = math.atan2(
                start_point[1] - center_y,
                start_point[0] - center_x,
            )
            end_angle = math.atan2(
                end_point[1] - center_y,
                end_point[0] - center_x,
            )
            sweep = (end_angle - start_angle) % (2.0 * math.pi)

        if sweep <= _ANGLE_EPSILON:
            raise RuntimeError(f"Arc '{name}' has an invalid zero sweep.")

        return GeometrySnapshot(
            kind="arc",
            name=name,
            data={
                "center": (center_x, center_y),
                "radius": radius,
                "start_point": start_point,
                "end_point": end_point,
                "sweep": sweep,
            },
            interface_details={
                "geometric_type": type_details,
                "target_interface": target_interface,
                "primary_access": read_details,
                "query_interface_required": False,
            },
        )
    except Exception as evaluate_error:
        diagnostics["attempts"].append(
            {
                "strategy": "system_service_evaluate",
                "status": "error",
                "error": _describe_exception(evaluate_error),
            }
        )

    # Optional fallback: use a concrete Python interface only when the local
    # SketcherInterfaces registration is actually available.
    try:
        specific, cast_details = _cast_geometry_interface(element, target_interface)
        if candidate_kind == "line":
            x1, y1, x2, y2 = _get_curve_endpoints(specific, catia_app)
            if math.hypot(x2 - x1, y2 - y1) <= _EPSILON:
                raise RuntimeError(f"Line '{name}' has zero length.")
            return GeometrySnapshot(
                kind="line",
                name=name,
                data={"start": (x1, y1), "end": (x2, y2)},
                interface_details={
                    "geometric_type": type_details,
                    "target_interface": target_interface,
                    "fallback_access": cast_details,
                    "query_interface_required": True,
                },
            )
        if candidate_kind == "point":
            point_x, point_y = _get_point_coordinates(specific, catia_app)
            return GeometrySnapshot(
                kind="point",
                name=name,
                data={"point": (point_x, point_y)},
                interface_details={
                    "geometric_type": type_details,
                    "target_interface": target_interface,
                    "fallback_access": cast_details,
                    "query_interface_required": True,
                },
            )

        radius = float(specific.Radius)
        center_x, center_y = _get_circle_center(specific, catia_app)
        params = _get_curve_param_extents(specific, catia_app)
        periodic = _is_periodic_curve(specific)
        if periodic:
            return GeometrySnapshot(
                kind="circle",
                name=name,
                data={"center": (center_x, center_y), "radius": radius},
                interface_details={
                    "geometric_type": type_details,
                    "target_interface": target_interface,
                    "fallback_access": cast_details,
                    "query_interface_required": True,
                },
            )
        endpoints = _get_curve_endpoints(specific, catia_app)
        start_point = (endpoints[0], endpoints[1])
        end_point = (endpoints[2], endpoints[3])
        if params is not None:
            sweep = float(params[1] - params[0])
            while sweep <= 0.0:
                sweep += 2.0 * math.pi
        else:
            start_angle = math.atan2(start_point[1] - center_y, start_point[0] - center_x)
            end_angle = math.atan2(end_point[1] - center_y, end_point[0] - center_x)
            sweep = (end_angle - start_angle) % (2.0 * math.pi)
        if sweep >= 2.0 * math.pi - _ANGLE_EPSILON:
            return GeometrySnapshot(
                kind="circle",
                name=name,
                data={"center": (center_x, center_y), "radius": radius},
                interface_details={
                    "geometric_type": type_details,
                    "target_interface": target_interface,
                    "fallback_access": cast_details,
                    "query_interface_required": True,
                },
            )
        return GeometrySnapshot(
            kind="arc",
            name=name,
            data={
                "center": (center_x, center_y),
                "radius": radius,
                "start_point": start_point,
                "end_point": end_point,
                "sweep": sweep,
            },
            interface_details={
                "geometric_type": type_details,
                "target_interface": target_interface,
                "fallback_access": cast_details,
                "query_interface_required": True,
            },
        )
    except Exception as fallback_error:
        diagnostics["attempts"].append(
            {
                "strategy": "optional_query_interface_fallback",
                "status": "error",
                "error": _describe_exception(fallback_error),
            }
        )

    raise RuntimeError(
        f"Cannot snapshot sketch geometry '{name or fallback_name}'. Diagnostics: "
        f"{json.dumps(diagnostics, ensure_ascii=False, default=str)}"
    )

def _normalise_arc_start(angle: float) -> float:
    return float(angle % (2.0 * math.pi))


def _create_from_snapshot(
    factory2d: Any,
    snapshot: GeometrySnapshot,
    transform: Callable[[float, float], Tuple[float, float]],
    orientation_reversed: bool = False,
) -> Any:
    """Create one transformed copy from a geometry snapshot."""

    if snapshot.kind == "point":
        point_x, point_y = transform(*snapshot.data["point"])
        return factory2d.CreatePoint(point_x, point_y)

    if snapshot.kind == "line":
        start_x, start_y = transform(*snapshot.data["start"])
        end_x, end_y = transform(*snapshot.data["end"])
        if math.hypot(end_x - start_x, end_y - start_y) <= _EPSILON:
            raise RuntimeError(f"Transformed line '{snapshot.name}' has zero length.")
        return factory2d.CreateLine(start_x, start_y, end_x, end_y)

    if snapshot.kind == "circle":
        center_x, center_y = transform(*snapshot.data["center"])
        return factory2d.CreateClosedCircle(
            center_x,
            center_y,
            float(snapshot.data["radius"]),
        )

    if snapshot.kind == "arc":
        center_x, center_y = transform(*snapshot.data["center"])
        sweep = float(snapshot.data["sweep"])

        # Reflection reverses curve orientation.  Swapping transformed endpoints
        # preserves the same arc locus while satisfying CATIA's CCW parameters.
        source_start = snapshot.data["start_point"]
        source_end = snapshot.data["end_point"]
        if orientation_reversed:
            start_x, start_y = transform(*source_end)
        else:
            start_x, start_y = transform(*source_start)

        start_angle = _normalise_arc_start(
            math.atan2(start_y - center_y, start_x - center_x)
        )
        end_angle = start_angle + sweep
        return factory2d.CreateCircle(
            center_x,
            center_y,
            float(snapshot.data["radius"]),
            start_angle,
            end_angle,
        )

    raise RuntimeError(f"Unsupported snapshot kind: {snapshot.kind}")


def _get_sketch_in_edit(
    catia_app: Any,
    sketch_name: str,
) -> Tuple[Any, Any, Any, Any]:
    """Find a named sketch in the active CATPart and open its edition."""

    if not str(sketch_name).strip():
        raise RuntimeError("sketch_name cannot be empty.")

    try:
        doc = catia_app.ActiveDocument
    except Exception as exc:
        raise RuntimeError("No active CATIA document.") from exc

    try:
        document_name = str(doc.Name)
    except Exception:
        document_name = ""

    if not document_name.lower().endswith(".catpart"):
        raise RuntimeError("Active document must be a CATPart for sketch operations.")

    try:
        part = doc.Part
    except Exception as exc:
        raise RuntimeError("Cannot access Part from the active CATPart.") from exc

    try:
        doc.Selection.Clear()
    except Exception:
        pass

    sketch = None

    try:
        bodies = part.Bodies
        for body_index in range(1, int(bodies.Count) + 1):
            body = bodies.Item(body_index)
            try:
                sketch = body.Sketches.Item(sketch_name)
                break
            except Exception:
                continue
    except Exception:
        pass

    if sketch is None:
        try:
            hybrid_bodies = part.HybridBodies
            for body_index in range(1, int(hybrid_bodies.Count) + 1):
                hybrid_body = hybrid_bodies.Item(body_index)
                try:
                    sketch = hybrid_body.HybridSketches.Item(sketch_name)
                    break
                except Exception:
                    continue
        except Exception:
            pass

    if sketch is None:
        raise RuntimeError(
            f"Sketch '{sketch_name}' was not found in the active CATPart."
        )

    try:
        factory2d = sketch.OpenEdition()
    except Exception as exc:
        raise RuntimeError(
            f"Cannot open sketch '{sketch_name}' for editing: {exc}"
        ) from exc

    return sketch, factory2d, part, doc


def _close_sketch(
    sketch: Any,
    part: Any,
) -> Tuple[bool, bool, List[str]]:
    """Close sketch edition, then update the sketch after it is closed."""

    warnings: List[str] = []
    closed = False
    updated = False

    try:
        sketch.CloseEdition()
        closed = True
    except Exception as exc:
        warnings.append(f"Sketch CloseEdition failed: {exc}")
        return closed, updated, warnings

    try:
        part.UpdateObject(sketch)
        updated = True
    except Exception as update_object_error:
        try:
            part.Update()
            updated = True
        except Exception as update_error:
            warnings.append(
                "Part update after sketch close failed: "
                f"UpdateObject={update_object_error}; Update={update_error}"
            )

    return closed, updated, warnings


def _get_element_by_name(sketch: Any, name: str) -> Any:
    element_name = str(name).strip()
    if not element_name:
        raise RuntimeError("Sketch element name cannot be empty.")

    try:
        return sketch.GeometricElements.Item(element_name)
    except Exception as direct_error:
        # CATIA names may be localised; preserve an exact case-insensitive fallback.
        try:
            elements = sketch.GeometricElements
            for index in range(1, int(elements.Count) + 1):
                candidate = elements.Item(index)
                if _get_element_name(candidate).casefold() == element_name.casefold():
                    return candidate
        except Exception:
            pass
        raise RuntimeError(
            f"Element '{element_name}' was not found in sketch."
        ) from direct_error


def _delete_elements(doc: Any, elements: Sequence[Any]) -> int:
    if not elements:
        return 0

    selection = doc.Selection
    selection.Clear()
    added = 0
    try:
        for element in elements:
            selection.Add(element)
            added += 1
        selection.Delete()
        return added
    finally:
        try:
            selection.Clear()
        except Exception:
            pass


def _constraint_count(sketch: Any) -> int:
    try:
        return int(sketch.Constraints.Count)
    except Exception:
        return 0


def _evaluate_sketch(sketch: Any) -> str:
    try:
        sketch.Evaluate()
        return "evaluated"
    except Exception:
        return "unknown"


def _line_intersection(
    source: Tuple[float, float, float, float],
    boundary: Tuple[float, float, float, float],
) -> Tuple[float, float, float, float]:
    """Return intersection x/y plus parameters on source and boundary supports."""

    x1, y1, x2, y2 = source
    x3, y3, x4, y4 = boundary

    source_dx = x2 - x1
    source_dy = y2 - y1
    boundary_dx = x4 - x3
    boundary_dy = y4 - y3

    denominator = source_dx * boundary_dy - source_dy * boundary_dx
    if abs(denominator) <= _EPSILON:
        raise RuntimeError("Source and boundary lines are parallel or coincident.")

    delta_x = x3 - x1
    delta_y = y3 - y1
    source_parameter = (
        delta_x * boundary_dy - delta_y * boundary_dx
    ) / denominator
    boundary_parameter = (
        delta_x * source_dy - delta_y * source_dx
    ) / denominator

    intersection_x = x1 + source_parameter * source_dx
    intersection_y = y1 + source_parameter * source_dy
    return intersection_x, intersection_y, source_parameter, boundary_parameter


def _reflect_transform(
    axis_start: Tuple[float, float],
    axis_end: Tuple[float, float],
) -> Callable[[float, float], Tuple[float, float]]:
    axis_x1, axis_y1 = axis_start
    axis_x2, axis_y2 = axis_end
    axis_dx = axis_x2 - axis_x1
    axis_dy = axis_y2 - axis_y1
    length_squared = axis_dx * axis_dx + axis_dy * axis_dy

    if length_squared <= _EPSILON:
        raise RuntimeError("Mirror axis has zero length.")

    def reflect(point_x: float, point_y: float) -> Tuple[float, float]:
        projection = (
            (point_x - axis_x1) * axis_dx
            + (point_y - axis_y1) * axis_dy
        ) / length_squared
        projected_x = axis_x1 + projection * axis_dx
        projected_y = axis_y1 + projection * axis_dy
        return 2.0 * projected_x - point_x, 2.0 * projected_y - point_y

    return reflect


def _translation_transform(
    delta_x: float,
    delta_y: float,
) -> Callable[[float, float], Tuple[float, float]]:
    def translate(point_x: float, point_y: float) -> Tuple[float, float]:
        return point_x + delta_x, point_y + delta_y

    return translate


def _rotation_transform(
    center_x: float,
    center_y: float,
    angle_rad: float,
) -> Callable[[float, float], Tuple[float, float]]:
    cosine = math.cos(angle_rad)
    sine = math.sin(angle_rad)

    def rotate(point_x: float, point_y: float) -> Tuple[float, float]:
        delta_x = point_x - center_x
        delta_y = point_y - center_y
        return (
            center_x + delta_x * cosine - delta_y * sine,
            center_y + delta_x * sine + delta_y * cosine,
        )

    return rotate


def mirror_sketch_geometry(
    catia_app: Any,
    sketch_name: str,
    element_names: List[str],
    axis_name: str,
    copy: bool = True,
) -> Dict[str, Any]:
    """Mirror supported sketch geometry about a line in the same sketch."""

    result = SketchOperationResult(operation="mirror_sketch_geometry")
    sketch = factory2d = part = doc = None
    created_elements: List[Any] = []

    try:
        if not element_names:
            raise RuntimeError("element_names must contain at least one element.")
        if any(not str(name).strip() for name in element_names):
            raise RuntimeError("element_names cannot contain empty names.")
        if str(axis_name).strip().casefold() in {
            str(name).strip().casefold() for name in element_names
        }:
            raise RuntimeError("The mirror axis cannot also be a source element.")

        sketch, factory2d, part, doc = _get_sketch_in_edit(catia_app, sketch_name)

        axis_element = _get_element_by_name(sketch, axis_name)
        axis_snapshot = _snapshot_geometry(axis_element, axis_name, catia_app)
        if axis_snapshot.kind != "line":
            raise RuntimeError("axis_name must identify a Line2D element.")

        reflect = _reflect_transform(
            axis_snapshot.data["start"],
            axis_snapshot.data["end"],
        )

        source_elements: List[Any] = []
        source_snapshots: List[GeometrySnapshot] = []
        for element_name in element_names:
            source = _get_element_by_name(sketch, element_name)
            source_elements.append(source)
            source_snapshots.append(_snapshot_geometry(source, element_name, catia_app))

        # All source geometry is validated before the first CATIA modification.
        try:
            for snapshot in source_snapshots:
                mirrored = _create_from_snapshot(
                    factory2d,
                    snapshot,
                    reflect,
                    orientation_reversed=True,
                )
                created_elements.append(mirrored)
                result.element_names.append(
                    _get_element_name(mirrored, f"Mirror_{snapshot.name}")
                )
        except Exception:
            try:
                _delete_elements(doc, created_elements)
            except Exception as rollback_error:
                result.warnings.append(f"Mirror rollback failed: {rollback_error}")
            raise

        result.elements_created = len(created_elements)

        if copy:
            try:
                constraints = sketch.Constraints
                for source, mirrored in zip(source_elements, created_elements):
                    try:
                        constraints.AddTriEltCst(
                            _CAT_CST_TYPE_SYMMETRY,
                            source,
                            mirrored,
                            axis_element,
                        )
                        result.constraints_added += 1
                    except Exception as exc:
                        result.warnings.append(
                            "Mirrored geometry was created, but its symmetry "
                            f"constraint could not be added: {exc}"
                        )
            except Exception as exc:
                result.warnings.append(
                    f"Cannot access sketch constraints for mirror relation: {exc}"
                )
        else:
            constraints_before = _constraint_count(sketch)
            try:
                result.originals_deleted = _delete_elements(doc, source_elements)
            except Exception as exc:
                result.warnings.append(
                    f"Mirrored copies were created, but source deletion failed: {exc}"
                )
            constraints_after = _constraint_count(sketch)
            result.constraints_removed = max(0, constraints_before - constraints_after)

        result.details["geometry_interfaces"] = {
            "axis": axis_snapshot.interface_details,
            "sources": [
                {
                    "name": snapshot.name,
                    "kind": snapshot.kind,
                    "interface": snapshot.interface_details,
                }
                for snapshot in source_snapshots
            ],
        }
        result.constraint_status = _evaluate_sketch(sketch)
        result.status = (
            "success" if not result.warnings else "success_with_warnings"
        )

    except Exception as exc:
        result.status = "error"
        result.warnings.append(str(exc))
    finally:
        if sketch is not None and part is not None:
            closed, updated, close_warnings = _close_sketch(sketch, part)
            result.sketch_closed = closed
            result.update_succeeded = updated
            result.warnings.extend(close_warnings)
            if result.status == "success" and result.warnings:
                result.status = "success_with_warnings"
        if doc is not None:
            try:
                doc.Selection.Clear()
            except Exception:
                pass

    return _result_dict(result)


def trim_extend_sketch_element(
    catia_app: Any,
    sketch_name: str,
    element_name: str,
    boundary_name: str,
    mode: str = "trim",
    extend_side: str = "nearest",
) -> Dict[str, Any]:
    """
    Trim or extend a finite line segment to another line's infinite support.

    The selected source endpoint is moved with Point2D.SetData.  No nonexistent
    Factory2D CreateTrim/CreateExtend method is called.
    """

    result = SketchOperationResult(operation="trim_extend_sketch_element")
    sketch = factory2d = part = doc = None

    try:
        normalised_mode = _normalise_text(mode)
        normalised_side = _normalise_text(extend_side)
        if normalised_mode not in {"trim", "extend"}:
            raise RuntimeError("mode must be 'trim' or 'extend'.")
        if normalised_side not in {"start", "end", "nearest"}:
            raise RuntimeError("extend_side must be 'start', 'end', or 'nearest'.")
        if _normalise_text(element_name) == _normalise_text(boundary_name):
            raise RuntimeError("element_name and boundary_name must be different.")

        sketch, factory2d, part, doc = _get_sketch_in_edit(catia_app, sketch_name)
        source_element = _get_element_by_name(sketch, element_name)
        boundary_element = _get_element_by_name(sketch, boundary_name)

        source_snapshot = _snapshot_geometry(source_element, element_name, catia_app)
        boundary_snapshot = _snapshot_geometry(boundary_element, boundary_name, catia_app)
        if source_snapshot.kind != "line":
            raise RuntimeError("Trim/extend currently supports a Line2D source only.")
        if boundary_snapshot.kind != "line":
            raise RuntimeError("Trim/extend currently supports a Line2D boundary only.")

        source_coords = (
            source_snapshot.data["start"][0],
            source_snapshot.data["start"][1],
            source_snapshot.data["end"][0],
            source_snapshot.data["end"][1],
        )
        boundary_coords = (
            boundary_snapshot.data["start"][0],
            boundary_snapshot.data["start"][1],
            boundary_snapshot.data["end"][0],
            boundary_snapshot.data["end"][1],
        )
        intersection_x, intersection_y, source_parameter, _ = _line_intersection(
            source_coords,
            boundary_coords,
        )

        if normalised_mode == "trim":
            if source_parameter < -_EPSILON or source_parameter > 1.0 + _EPSILON:
                raise RuntimeError(
                    "The boundary does not intersect the finite source segment. "
                    "Use mode='extend' instead."
                )
            if (
                abs(source_parameter) <= _EPSILON
                or abs(source_parameter - 1.0) <= _EPSILON
            ):
                raise RuntimeError("Source endpoint already lies on the boundary.")

            if normalised_side == "start":
                endpoint_name = "start"
            elif normalised_side == "end":
                endpoint_name = "end"
            else:
                endpoint_name = (
                    "start" if source_parameter <= 0.5 else "end"
                )
        else:
            if -_EPSILON <= source_parameter <= 1.0 + _EPSILON:
                raise RuntimeError(
                    "The boundary intersects the existing source segment. "
                    "Use mode='trim' instead."
                )

            required_side = "start" if source_parameter < 0.0 else "end"
            if normalised_side == "nearest":
                endpoint_name = required_side
            else:
                endpoint_name = normalised_side
                if endpoint_name != required_side:
                    raise RuntimeError(
                        f"Intersection lies beyond the {required_side} endpoint; "
                        f"extend_side='{endpoint_name}' would not extend the line."
                    )

        original_point = (
            source_snapshot.data["start"]
            if endpoint_name == "start"
            else source_snapshot.data["end"]
        )
        original_x, original_y = original_point

        try:
            confirmed_point, constraint_added, move_details = (
                _system_service_move_line_endpoint(
                    catia_app,
                    sketch,
                    source_element,
                    boundary_element,
                    endpoint_name,
                    intersection_x,
                    intersection_y,
                )
            )
        except Exception as exc:
            raise RuntimeError(
                f"Could not move the source {endpoint_name} endpoint inside CATIA: "
                f"{_describe_exception(exc)}"
            ) from exc

        result.elements_modified = 1
        result.element_names.append(_get_element_name(source_element, element_name))
        if constraint_added:
            result.constraints_added += 1
        else:
            result.warnings.append(
                "Geometry was modified, but CATIA could not add the "
                "endpoint-on-boundary constraint."
            )

        result.constraint_status = _evaluate_sketch(sketch)

        confirmed_x, confirmed_y = confirmed_point
        if math.hypot(
            confirmed_x - intersection_x,
            confirmed_y - intersection_y,
        ) > 1.0e-6:
            result.warnings.append(
                "CATIA solver changed the moved endpoint after evaluation."
            )

        result.details.update(
            {
                "mode": normalised_mode,
                "moved_endpoint": endpoint_name,
                "original_coordinates": [original_x, original_y],
                "intersection_coordinates": [intersection_x, intersection_y],
                "source_parameter": source_parameter,
                "geometry_interfaces": {
                    "source": source_snapshot.interface_details,
                    "boundary": boundary_snapshot.interface_details,
                },
                "endpoint_update": move_details,
            }
        )
        result.status = (
            "success" if not result.warnings else "success_with_warnings"
        )

    except Exception as exc:
        result.status = "error"
        result.warnings.append(str(exc))
    finally:
        if sketch is not None and part is not None:
            closed, updated, close_warnings = _close_sketch(sketch, part)
            result.sketch_closed = closed
            result.update_succeeded = updated
            result.warnings.extend(close_warnings)
            if result.status == "success" and result.warnings:
                result.status = "success_with_warnings"
        if doc is not None:
            try:
                doc.Selection.Clear()
            except Exception:
                pass

    return _result_dict(result)


def pattern_sketch_elements(
    catia_app: Any,
    sketch_name: str,
    element_names: List[str],
    pattern_type: str = "rectangular",
    count_x: int = 3,
    count_y: int = 1,
    spacing_x_mm: float = 20.0,
    spacing_y_mm: float = 20.0,
    circular_count: int = 6,
    circular_radius_mm: float = 50.0,
    center_x_mm: float = 0.0,
    center_y_mm: float = 0.0,
    angle_start_deg: float = 0.0,
    angle_span_deg: float = 360.0,
) -> Dict[str, Any]:
    """
    Reconstruct rectangular or circular copies of supported sketch geometry.

    Counts include the original instance.  Therefore, a 3x2 rectangular pattern
    creates five copies per source element, and a circular_count of four creates
    three copies per source element.
    """

    del circular_radius_mm  # Kept in the MCP schema for backward compatibility.

    result = SketchOperationResult(operation="pattern_sketch_elements")
    sketch = factory2d = part = doc = None
    created_elements: List[Any] = []

    try:
        normalised_pattern = _normalise_text(pattern_type)
        if normalised_pattern not in {"rectangular", "circular"}:
            raise RuntimeError("pattern_type must be 'rectangular' or 'circular'.")
        if not element_names:
            raise RuntimeError("element_names must contain at least one element.")
        if any(not str(name).strip() for name in element_names):
            raise RuntimeError("element_names cannot contain empty names.")

        if normalised_pattern == "rectangular":
            if not isinstance(count_x, int) or isinstance(count_x, bool):
                raise RuntimeError("count_x must be an integer.")
            if not isinstance(count_y, int) or isinstance(count_y, bool):
                raise RuntimeError("count_y must be an integer.")
            if count_x < 1 or count_y < 1:
                raise RuntimeError("count_x and count_y must be at least 1.")
            if count_x > 100 or count_y > 100:
                raise RuntimeError("count_x and count_y cannot exceed 100.")
            if count_x > 1 and abs(float(spacing_x_mm)) <= _EPSILON:
                raise RuntimeError("spacing_x_mm cannot be zero when count_x > 1.")
            if count_y > 1 and abs(float(spacing_y_mm)) <= _EPSILON:
                raise RuntimeError("spacing_y_mm cannot be zero when count_y > 1.")
        else:
            if not isinstance(circular_count, int) or isinstance(circular_count, bool):
                raise RuntimeError("circular_count must be an integer.")
            if circular_count < 2 or circular_count > 360:
                raise RuntimeError("circular_count must be between 2 and 360.")
            if not math.isfinite(float(angle_start_deg)):
                raise RuntimeError("angle_start_deg must be finite.")
            if (
                not math.isfinite(float(angle_span_deg))
                or float(angle_span_deg) <= 0.0
                or float(angle_span_deg) > 360.0
            ):
                raise RuntimeError("angle_span_deg must be in the range (0, 360].")

        numeric_values = (
            spacing_x_mm,
            spacing_y_mm,
            center_x_mm,
            center_y_mm,
        )
        if not all(math.isfinite(float(value)) for value in numeric_values):
            raise RuntimeError("Pattern coordinates and spacing values must be finite.")

        sketch, factory2d, part, doc = _get_sketch_in_edit(catia_app, sketch_name)

        source_snapshots: List[GeometrySnapshot] = []
        for element_name in element_names:
            source = _get_element_by_name(sketch, element_name)
            source_snapshots.append(_snapshot_geometry(source, element_name, catia_app))

        transforms: List[Callable[[float, float], Tuple[float, float]]] = []

        if normalised_pattern == "rectangular":
            for y_index in range(count_y):
                for x_index in range(count_x):
                    if x_index == 0 and y_index == 0:
                        continue
                    transforms.append(
                        _translation_transform(
                            x_index * float(spacing_x_mm),
                            y_index * float(spacing_y_mm),
                        )
                    )
        else:
            full_circle = math.isclose(
                float(angle_span_deg),
                360.0,
                rel_tol=0.0,
                abs_tol=1.0e-9,
            )
            denominator = circular_count if full_circle else circular_count - 1
            angle_step = float(angle_span_deg) / denominator

            for instance_index in range(1, circular_count):
                rotation_deg = float(angle_start_deg) + instance_index * angle_step
                transforms.append(
                    _rotation_transform(
                        float(center_x_mm),
                        float(center_y_mm),
                        math.radians(rotation_deg),
                    )
                )

        expected_created = len(transforms) * len(source_snapshots)

        try:
            for transform in transforms:
                for snapshot in source_snapshots:
                    created = _create_from_snapshot(
                        factory2d,
                        snapshot,
                        transform,
                        orientation_reversed=False,
                    )
                    created_elements.append(created)
                    result.element_names.append(
                        _get_element_name(
                            created,
                            f"Pattern_{len(created_elements)}_{snapshot.name}",
                        )
                    )
        except Exception:
            try:
                _delete_elements(doc, created_elements)
                created_elements.clear()
                result.element_names.clear()
            except Exception as rollback_error:
                result.warnings.append(f"Pattern rollback failed: {rollback_error}")
            raise

        result.elements_created = len(created_elements)
        if result.elements_created != expected_created:
            raise RuntimeError(
                "Pattern creation count mismatch: "
                f"expected {expected_created}, created {result.elements_created}."
            )

        # Automatically inventing distance/fix constraints changes design intent
        # and previously used invalid enum values.  Geometry copies are created
        # deterministically; source constraints are not cloned.
        result.constraint_status = _evaluate_sketch(sketch)
        result.details.update(
            {
                "pattern_type": normalised_pattern,
                "source_element_count": len(source_snapshots),
                "expected_elements_created": expected_created,
                "counts_include_original": True,
                "source_constraints_cloned": False,
                "source_interfaces": [
                    {
                        "name": snapshot.name,
                        "kind": snapshot.kind,
                        "interface": snapshot.interface_details,
                    }
                    for snapshot in source_snapshots
                ],
            }
        )
        if normalised_pattern == "circular":
            result.details["circular_radius_parameter"] = (
                "compatibility_only; actual radius is defined by source geometry "
                "position relative to the pattern center"
            )
        result.status = "success"

    except Exception as exc:
        result.status = "error"
        result.warnings.append(str(exc))
    finally:
        if sketch is not None and part is not None:
            closed, updated, close_warnings = _close_sketch(sketch, part)
            result.sketch_closed = closed
            result.update_succeeded = updated
            result.warnings.extend(close_warnings)
            if result.status == "success" and result.warnings:
                result.status = "success_with_warnings"
        if doc is not None:
            try:
                doc.Selection.Clear()
            except Exception:
                pass

    return _result_dict(result)


SKETCH_UTILITY_MCP_TOOLS = [
    {
        "name": "catia_mirror_sketch_geometry",
        "description": (
            "Mirror Point2D, Line2D, circles, or circular arcs about a sketch "
            "line. copy=true keeps the source and attempts to add a CATIA "
            "symmetry constraint; copy=false creates the mirror and deletes "
            "the source after successful creation."
        ),
        "parameters": {
            "type": "object",
            "properties": {
                "sketch_name": {
                    "type": "string",
                    "description": "Name of the sketch in the feature tree.",
                },
                "element_names": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Names of supported geometry to mirror.",
                },
                "axis_name": {
                    "type": "string",
                    "description": "Name of the Line2D mirror axis.",
                },
                "copy": {
                    "type": "boolean",
                    "description": "Keep source geometry when true.",
                    "default": True,
                },
            },
            "required": ["sketch_name", "element_names", "axis_name"],
        },
    },
    {
        "name": "catia_trim_extend_sketch_element",
        "description": (
            "Trim or extend a finite Line2D to another Line2D support by moving "
            "one source endpoint with Point2D.SetData. In trim mode, nearest "
            "moves the endpoint closest to the intersection. In extend mode, "
            "nearest selects the endpoint beyond which the intersection lies."
        ),
        "parameters": {
            "type": "object",
            "properties": {
                "sketch_name": {
                    "type": "string",
                    "description": "Name of the sketch.",
                },
                "element_name": {
                    "type": "string",
                    "description": "Name of the source Line2D.",
                },
                "boundary_name": {
                    "type": "string",
                    "description": "Name of the boundary Line2D.",
                },
                "mode": {
                    "type": "string",
                    "enum": ["trim", "extend"],
                    "default": "trim",
                },
                "extend_side": {
                    "type": "string",
                    "enum": ["start", "end", "nearest"],
                    "description": "Endpoint to move; also applies to trim mode.",
                    "default": "nearest",
                },
            },
            "required": ["sketch_name", "element_name", "boundary_name"],
        },
    },
    {
        "name": "catia_pattern_sketch_elements",
        "description": (
            "Create rectangular or circular copies of Point2D, Line2D, circles, "
            "or circular arcs. Counts include the original instance. Copies are "
            "reconstructed from geometry coordinates; source constraints are "
            "not cloned automatically."
        ),
        "parameters": {
            "type": "object",
            "properties": {
                "sketch_name": {
                    "type": "string",
                    "description": "Name of the sketch.",
                },
                "element_names": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Names of supported source geometry.",
                },
                "pattern_type": {
                    "type": "string",
                    "enum": ["rectangular", "circular"],
                    "default": "rectangular",
                },
                "count_x": {
                    "type": "integer",
                    "description": "Total X instances including the original.",
                    "default": 3,
                    "minimum": 1,
                    "maximum": 100,
                },
                "count_y": {
                    "type": "integer",
                    "description": "Total Y instances including the original.",
                    "default": 1,
                    "minimum": 1,
                    "maximum": 100,
                },
                "spacing_x_mm": {
                    "type": "number",
                    "default": 20.0,
                },
                "spacing_y_mm": {
                    "type": "number",
                    "default": 20.0,
                },
                "circular_count": {
                    "type": "integer",
                    "description": "Total circular instances including original.",
                    "default": 6,
                    "minimum": 2,
                    "maximum": 360,
                },
                "circular_radius_mm": {
                    "type": "number",
                    "description": (
                        "Compatibility parameter. Actual radius is determined "
                        "by source geometry position relative to the center."
                    ),
                    "default": 50.0,
                },
                "center_x_mm": {
                    "type": "number",
                    "default": 0.0,
                },
                "center_y_mm": {
                    "type": "number",
                    "default": 0.0,
                },
                "angle_start_deg": {
                    "type": "number",
                    "description": "Additional rotation offset in degrees.",
                    "default": 0.0,
                },
                "angle_span_deg": {
                    "type": "number",
                    "default": 360.0,
                    "exclusiveMinimum": 0.0,
                    "maximum": 360.0,
                },
            },
            "required": ["sketch_name", "element_names"],
        },
    },
]


def register_tools(mcp: Any, ctx: Any) -> list[str]:
    """Register sketch utility tools using FastMCP function decorators.

    FastMCP derives each tool's input schema from the wrapper function
    signature.  This avoids the obsolete ``add_tool(parameters=..., handler=...)``
    registration form that is incompatible with the server's installed SDK.
    """

    @mcp.tool()
    def catia_mirror_sketch_geometry(
        sketch_name: str,
        element_names: List[str],
        axis_name: str,
        copy: bool = True,
    ) -> Dict[str, Any]:
        """Mirror supported sketch geometry about a line in the same sketch."""
        execution_context: Dict[str, Any] = {"before_connect": _execution_context(ctx)}
        try:
            ctx.conn.ensure_connected()
            execution_context["after_connect"] = _execution_context(ctx)
            response = mirror_sketch_geometry(
                ctx.conn.app,
                sketch_name=sketch_name,
                element_names=element_names,
                axis_name=axis_name,
                copy=copy,
            )
            response.setdefault("details", {})["execution_context"] = execution_context
            return response
        except Exception as exc:
            return {
                "operation": "mirror_sketch_geometry",
                "status": "error",
                "elements_created": 0,
                "elements_modified": 0,
                "originals_deleted": 0,
                "constraints_added": 0,
                "constraints_removed": 0,
                "constraint_status": "unknown",
                "element_names": [],
                "warnings": [_describe_exception(exc)],
                "details": {
                    "execution_context": execution_context,
                    "wrapper_exception": _describe_exception(exc),
                },
                "sketch_closed": False,
                "update_succeeded": False,
                "document_unit": "mm",
            }

    @mcp.tool()
    def catia_trim_extend_sketch_element(
        sketch_name: str,
        element_name: str,
        boundary_name: str,
        mode: str = "trim",
        extend_side: str = "nearest",
    ) -> Dict[str, Any]:
        """Trim or extend a Line2D to another Line2D boundary."""
        execution_context: Dict[str, Any] = {"before_connect": _execution_context(ctx)}
        try:
            ctx.conn.ensure_connected()
            execution_context["after_connect"] = _execution_context(ctx)
            response = trim_extend_sketch_element(
                ctx.conn.app,
                sketch_name=sketch_name,
                element_name=element_name,
                boundary_name=boundary_name,
                mode=mode,
                extend_side=extend_side,
            )
            response.setdefault("details", {})["execution_context"] = execution_context
            return response
        except Exception as exc:
            return {
                "operation": "trim_extend_sketch_element",
                "status": "error",
                "elements_created": 0,
                "elements_modified": 0,
                "originals_deleted": 0,
                "constraints_added": 0,
                "constraints_removed": 0,
                "constraint_status": "unknown",
                "element_names": [],
                "warnings": [_describe_exception(exc)],
                "details": {
                    "execution_context": execution_context,
                    "wrapper_exception": _describe_exception(exc),
                },
                "sketch_closed": False,
                "update_succeeded": False,
                "document_unit": "mm",
            }

    @mcp.tool()
    def catia_pattern_sketch_elements(
        sketch_name: str,
        element_names: List[str],
        pattern_type: str = "rectangular",
        count_x: int = 3,
        count_y: int = 1,
        spacing_x_mm: float = 20.0,
        spacing_y_mm: float = 20.0,
        circular_count: int = 6,
        circular_radius_mm: float = 50.0,
        center_x_mm: float = 0.0,
        center_y_mm: float = 0.0,
        angle_start_deg: float = 0.0,
        angle_span_deg: float = 360.0,
    ) -> Dict[str, Any]:
        """Create rectangular or circular copies of supported sketch geometry."""
        execution_context: Dict[str, Any] = {"before_connect": _execution_context(ctx)}
        try:
            ctx.conn.ensure_connected()
            execution_context["after_connect"] = _execution_context(ctx)
            response = pattern_sketch_elements(
                ctx.conn.app,
                sketch_name=sketch_name,
                element_names=element_names,
                pattern_type=pattern_type,
                count_x=count_x,
                count_y=count_y,
                spacing_x_mm=spacing_x_mm,
                spacing_y_mm=spacing_y_mm,
                circular_count=circular_count,
                circular_radius_mm=circular_radius_mm,
                center_x_mm=center_x_mm,
                center_y_mm=center_y_mm,
                angle_start_deg=angle_start_deg,
                angle_span_deg=angle_span_deg,
            )
            response.setdefault("details", {})["execution_context"] = execution_context
            return response
        except Exception as exc:
            return {
                "operation": "pattern_sketch_elements",
                "status": "error",
                "elements_created": 0,
                "elements_modified": 0,
                "originals_deleted": 0,
                "constraints_added": 0,
                "constraints_removed": 0,
                "constraint_status": "unknown",
                "element_names": [],
                "warnings": [_describe_exception(exc)],
                "details": {
                    "execution_context": execution_context,
                    "wrapper_exception": _describe_exception(exc),
                },
                "sketch_closed": False,
                "update_succeeded": False,
                "document_unit": "mm",
            }

    return [
        "catia_mirror_sketch_geometry",
        "catia_trim_extend_sketch_element",
        "catia_pattern_sketch_elements",
    ]


# ---------------------------------------------------------------------------
# Registry Entry Point (Required by registry.py)
# ---------------------------------------------------------------------------
