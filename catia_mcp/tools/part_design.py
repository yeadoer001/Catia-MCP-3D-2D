"""
part_design.py
Version: part-design-fixed-2026-08-04-v2

CATIA V5 MCP Part Design tools.

v2 corrections:
- Pad/Pocket success can be verified from actual solid-volume change, not only
  from feature-tree creation and Part.Update.
- direction="auto" retries the opposite prism orientation when an offset-plane
  feature updates but does not add/remove material.
- Pocket and circular-hole tools support dimension, up_to_next, through_all
  (up_to_last), and through_next limit modes with CATIA readback.
- Circular boss/hole tools return an explicit sketch-plane coordinate contract.
- Failed geometry-effect verification removes the feature, updates the Part,
  and reports partial_success if rollback cannot restore the model.
"""

from __future__ import annotations

import math
from typing import Any

from catia_mcp.connection import CATIAError


IMPLEMENTATION_VERSION = "part-design-fixed-2026-08-04-v2"

# CATIA V5 PartInterfaces enumerations are zero based.
_CAT_REGULAR_ORIENTATION = 0
_CAT_INVERSE_ORIENTATION = 1
_CAT_LIMIT_MODE = {
    "dimension": 0,       # catOffsetLimit
    "up_to_next": 1,      # catUpToNextLimit
    "through_all": 2,     # catUpToLastLimit
    "up_to_last": 2,      # alias
    "through_next": 5,    # catUpThruNextLimit
}

_PLANE_CONTRACTS = {
    "xy": {
        "origin_for_offset": lambda d: [0.0, 0.0, d],
        "local_u_global": [1.0, 0.0, 0.0],
        "local_v_global": [0.0, 1.0, 0.0],
        "nominal_normal_global": [0.0, 0.0, 1.0],
    },
    "yz": {
        "origin_for_offset": lambda d: [d, 0.0, 0.0],
        "local_u_global": [0.0, 1.0, 0.0],
        "local_v_global": [0.0, 0.0, 1.0],
        "nominal_normal_global": [1.0, 0.0, 0.0],
    },
    "zx": {
        "origin_for_offset": lambda d: [0.0, d, 0.0],
        "local_u_global": [0.0, 0.0, 1.0],
        "local_v_global": [1.0, 0.0, 0.0],
        "nominal_normal_global": [0.0, 1.0, 0.0],
    },
}


def _success(data: Any, warnings: list[str] | None = None) -> dict[str, Any]:
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
    data: Any | None = None,
    warnings: list[str] | None = None,
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



def _finite_positive(value: Any, name: str) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise CATIAError(f"{name} must be a number.") from exc
    if not math.isfinite(number):
        raise CATIAError(f"{name} must be finite.")
    if number <= 0.0:
        raise CATIAError(f"{name} must be greater than 0.")
    return number


def _finite_number(value: Any, name: str) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise CATIAError(f"{name} must be a number.") from exc
    if not math.isfinite(number):
        raise CATIAError(f"{name} must be finite.")
    return number


def _normalise_direction(direction: str, allowed: set[str]) -> str:
    value = str(direction).strip().lower()
    if value not in allowed:
        expected = ", ".join(sorted(allowed))
        raise CATIAError(f"direction must be one of: {expected}.")
    return value


def _get_sketch(conn: Any, sketch_name: str = "") -> Any:
    body = conn.get_active_part_body()
    sketches = body.Sketches

    if str(sketch_name).strip():
        try:
            return sketches.Item(str(sketch_name).strip())
        except Exception as exc:
            raise CATIAError(f"Sketch not found in active PartBody: {sketch_name}") from exc

    if int(sketches.Count) < 1:
        raise CATIAError("No sketch found in active PartBody.")

    return sketches.Item(sketches.Count)


def _active_document(conn: Any) -> Any:
    try:
        return conn.app.ActiveDocument
    except Exception as exc:
        raise CATIAError("Cannot access the active CATIA document.") from exc


def _describe_com_object(value: Any) -> dict[str, Any]:
    return {
        "python_type": type(value).__name__,
        "python_module": type(value).__module__,
        "has_oleobj": bool(hasattr(value, "_oleobj_")),
    }


def _has_callable(value: Any, method_name: str) -> bool:
    try:
        return callable(getattr(value, method_name, None))
    except Exception:
        return False


def _get_shape_factory(part: Any, required_method: str) -> tuple[Any, dict[str, Any]]:
    """Resolve CATIA ShapeFactory even when pywin32 returns generic Factory."""

    try:
        raw_factory = part.ShapeFactory
    except Exception as exc:
        raise CATIAError(f"Cannot access Part.ShapeFactory: {exc}") from exc

    details: dict[str, Any] = {
        "required_method": required_method,
        "raw": _describe_com_object(raw_factory),
        "dispatch_used": False,
    }

    if _has_callable(raw_factory, required_method):
        details["resolved"] = _describe_com_object(raw_factory)
        details["required_method_available"] = True
        return raw_factory, details

    try:
        import win32com.client  # type: ignore

        shape_factory = win32com.client.Dispatch(raw_factory)
    except Exception as exc:
        details["dispatch_error"] = str(exc)
        raise CATIAError(
            "Part.ShapeFactory was returned as a generic COM Factory and "
            f"could not be dynamically dispatched: {exc}"
        ) from exc

    details["dispatch_used"] = True
    details["resolved"] = _describe_com_object(shape_factory)
    details["required_method_available"] = _has_callable(
        shape_factory, required_method
    )

    if not details["required_method_available"]:
        raise CATIAError(
            "ShapeFactory dispatch completed, but the resolved COM object "
            f"does not expose {required_method}."
        )

    return shape_factory, details


def _delete_object(conn: Any, obj: Any) -> bool:
    if obj is None:
        return True

    try:
        document = _active_document(conn)
        selection = document.Selection
        selection.Clear()
        selection.Add(obj)
        selection.Delete()
        selection.Clear()
        return True
    except Exception:
        try:
            _active_document(conn).Selection.Clear()
        except Exception:
            pass
        return False


def _set_feature_name(feature: Any, feature_name: str) -> list[str]:
    warnings: list[str] = []
    name = str(feature_name).strip()
    if not name:
        return warnings
    try:
        feature.Name = name
    except Exception as exc:
        warnings.append(f"Feature was created but could not be renamed to '{name}': {exc}")
    return warnings


def _update_feature(part: Any, feature: Any) -> tuple[bool, str, list[str]]:
    warnings: list[str] = []
    try:
        part.UpdateObject(feature)
        return True, "UpdateObject", warnings
    except Exception as update_object_error:
        warnings.append(f"UpdateObject failed: {update_object_error}")

    try:
        part.Update()
        warnings.append("Part.Update fallback succeeded.")
        return True, "Part.Update", warnings
    except Exception as update_error:
        warnings.append(f"Part.Update fallback failed: {update_error}")
        return False, "failed", warnings


def _normalise_plane(plane: Any) -> str:
    value = str(plane).strip().lower().replace("xz", "zx")
    if value not in _PLANE_CONTRACTS:
        raise CATIAError("plane must be one of: xy, yz, zx.")
    return value


def _plane_coordinate_contract(
    plane: Any,
    offset: Any,
    center_x: Any | None = None,
    center_y: Any | None = None,
) -> dict[str, Any]:
    plane_key = _normalise_plane(plane)
    offset_value = _finite_number(offset, "offset")
    base = _PLANE_CONTRACTS[plane_key]
    result: dict[str, Any] = {
        "plane": plane_key,
        "offset_mm": offset_value,
        "sketch_origin_global_mm": base["origin_for_offset"](offset_value),
        "local_u_global": list(base["local_u_global"]),
        "local_v_global": list(base["local_v_global"]),
        "nominal_normal_global": list(base["nominal_normal_global"]),
        "normal_sign_verified_from_catia": False,
        "direction_policy": (
            "Use direction='auto' when the material side of an offset plane is "
            "not known. The tool verifies the solid-volume effect and retries "
            "the inverse CATIA prism orientation when necessary."
        ),
    }
    if center_x is not None and center_y is not None:
        u = _finite_number(center_x, "center_x")
        v = _finite_number(center_y, "center_y")
        origin = result["sketch_origin_global_mm"]
        u_axis = result["local_u_global"]
        v_axis = result["local_v_global"]
        result["local_center_mm"] = [u, v]
        result["nominal_global_center_mm"] = [
            origin[i] + u * u_axis[i] + v * v_axis[i]
            for i in range(3)
        ]
    return result


def _normalise_limit_type(value: Any) -> str:
    aliases = {
        "blind": "dimension",
        "offset": "dimension",
        "up-to-next": "up_to_next",
        "up_to_last": "through_all",
        "up-to-last": "through_all",
        "through": "through_all",
        "thru_all": "through_all",
        "up-thru-next": "through_next",
        "thru_next": "through_next",
    }
    key = str(value or "dimension").strip().lower()
    key = aliases.get(key, key)
    if key not in {"dimension", "up_to_next", "through_all", "through_next"}:
        raise CATIAError(
            "limit_type must be one of: dimension, up_to_next, through_all, "
            "through_next."
        )
    return key


def _apply_prism_limit(feature: Any, limit_type: Any, dimension_value: float) -> dict[str, Any]:
    mode = _normalise_limit_type(limit_type)
    code = _CAT_LIMIT_MODE[mode]
    errors: list[str] = []
    try:
        limit = feature.FirstLimit
    except Exception as exc:
        raise CATIAError(f"Cannot access feature.FirstLimit: {exc}") from exc

    try:
        limit.LimitMode = code
    except Exception as exc:
        errors.append(f"LimitMode write: {exc}")
        raise CATIAError(
            f"Cannot set first limit mode '{mode}' (code {code}): {exc}"
        ) from exc

    if mode == "dimension":
        try:
            limit.Dimension.Value = float(dimension_value)
        except Exception as exc:
            errors.append(f"Dimension.Value write: {exc}")
            raise CATIAError(f"Cannot set dimensional first limit: {exc}") from exc

    try:
        actual_code = int(limit.LimitMode)
    except Exception as exc:
        actual_code = None
        errors.append(f"LimitMode readback: {exc}")

    actual_dimension = None
    if mode == "dimension":
        try:
            actual_dimension = float(limit.Dimension.Value)
        except Exception as exc:
            errors.append(f"Dimension.Value readback: {exc}")

    verified = actual_code == code and (
        mode != "dimension"
        or actual_dimension is None
        or abs(actual_dimension - float(dimension_value)) <= 1.0e-6
    )
    if actual_code is not None and not verified:
        raise CATIAError(
            f"First-limit readback mismatch: requested {mode}/{code}, "
            f"actual code={actual_code}, dimension={actual_dimension}."
        )

    return {
        "requested": mode,
        "requested_code": code,
        "actual_code": actual_code,
        "dimension_mm": float(dimension_value) if mode == "dimension" else None,
        "actual_dimension_mm": actual_dimension,
        "readback_verified": verified if actual_code is not None else None,
        "errors": errors,
    }


def _body_volume_snapshot(conn: Any, part: Any, body: Any) -> dict[str, Any]:
    details: dict[str, Any] = {
        "available": False,
        "volume_m3": None,
        "volume_mm3": None,
        "method": None,
        "error": None,
    }
    try:
        document = _active_document(conn)
        spa = document.GetWorkbench("SPAWorkbench")
        reference = part.CreateReferenceFromObject(body)
        measurable = spa.GetMeasurable(reference)
        volume_m3 = float(measurable.Volume)
        if not math.isfinite(volume_m3) or volume_m3 < 0.0:
            raise CATIAError(f"Invalid Measurable.Volume value: {volume_m3!r}")
        details.update(
            {
                "available": True,
                "volume_m3": volume_m3,
                "volume_mm3": volume_m3 * 1_000_000_000.0,
                "method": "ActiveDocument.GetWorkbench('SPAWorkbench').GetMeasurable",
            }
        )
    except Exception as exc:
        details["error"] = str(exc)
    return details


def _volume_effect(
    before: dict[str, Any],
    after: dict[str, Any],
    *,
    expected: str,
    tolerance_mm3: float,
) -> dict[str, Any]:
    tolerance = _finite_positive(tolerance_mm3, "volume_tolerance_mm3")
    before_value = before.get("volume_mm3") if before.get("available") else None
    after_value = after.get("volume_mm3") if after.get("available") else None
    result: dict[str, Any] = {
        "expected": expected,
        "before": before,
        "after": after,
        "delta_mm3": None,
        "threshold_mm3": tolerance,
        "verified": None,
        "reason": None,
    }

    if after_value is None:
        result["reason"] = "Post-operation solid volume is unavailable."
        return result

    if before_value is None:
        if expected == "increase" and float(after_value) > tolerance:
            result["delta_mm3"] = float(after_value)
            result["verified"] = True
            result["reason"] = "No measurable pre-existing solid; post volume is positive."
        else:
            result["reason"] = "Pre-operation solid volume is unavailable."
        return result

    delta = float(after_value) - float(before_value)
    threshold = max(tolerance, abs(float(before_value)) * 1.0e-9)
    result["delta_mm3"] = delta
    result["threshold_mm3"] = threshold
    if expected == "increase":
        result["verified"] = delta > threshold
    elif expected == "decrease":
        result["verified"] = delta < -threshold
    else:
        raise CATIAError(f"Unsupported volume-effect expectation: {expected}")
    result["reason"] = (
        "Solid volume changed in the expected direction."
        if result["verified"]
        else "Feature updated, but the solid volume did not change in the expected direction."
    )
    return result


def _set_prism_orientation(feature: Any, orientation_code: int) -> dict[str, Any]:
    code = int(orientation_code)
    if code not in {_CAT_REGULAR_ORIENTATION, _CAT_INVERSE_ORIENTATION}:
        raise CATIAError(f"Unsupported prism orientation code: {code}")
    try:
        feature.DirectionOrientation = code
    except Exception as exc:
        raise CATIAError(f"Cannot set prism DirectionOrientation={code}: {exc}") from exc
    try:
        actual = int(feature.DirectionOrientation)
    except Exception:
        actual = None
    if actual is not None and actual != code:
        raise CATIAError(
            f"Prism direction readback mismatch: requested {code}, actual {actual}."
        )
    return {
        "requested_code": code,
        "actual_code": actual,
        "readback_verified": actual == code if actual is not None else None,
        "name": "normal" if code == 0 else "reverse",
    }


def _rollback_feature(
    conn: Any,
    feature: Any,
    *,
    volume_before: dict[str, Any] | None = None,
    volume_tolerance_mm3: float = 1.0e-3,
) -> dict[str, Any]:
    details: dict[str, Any] = {
        "delete_succeeded": True,
        "update_succeeded": True,
        "volume_restored": None,
        "volume_after_rollback": None,
        "verified": True,
        "errors": [],
    }
    if feature is None:
        return details

    details["delete_succeeded"] = _delete_object(conn, feature)
    if not details["delete_succeeded"]:
        details["errors"].append("Selection.Delete failed for the feature.")

    try:
        part = conn.get_active_part()
        part.Update()
    except Exception as exc:
        details["update_succeeded"] = False
        details["errors"].append(f"Part.Update after rollback failed: {exc}")

    if volume_before and volume_before.get("available"):
        try:
            part = conn.get_active_part()
            body = conn.get_active_part_body()
            after = _body_volume_snapshot(conn, part, body)
            details["volume_after_rollback"] = after
            if after.get("available"):
                before_value = float(volume_before["volume_mm3"])
                after_value = float(after["volume_mm3"])
                threshold = max(
                    float(volume_tolerance_mm3),
                    abs(before_value) * 1.0e-9,
                )
                details["volume_restored"] = abs(after_value - before_value) <= threshold
            else:
                details["volume_restored"] = None
        except Exception as exc:
            details["errors"].append(f"Rollback volume verification failed: {exc}")

    details["verified"] = bool(
        details["delete_succeeded"]
        and details["update_succeeded"]
        and details["volume_restored"] is not False
    )
    return details


def _feature_failure(
    conn: Any,
    feature: Any,
    message: str,
    *,
    feature_type: str,
    warnings: list[str] | None = None,
    extra_data: dict[str, Any] | None = None,
    volume_before: dict[str, Any] | None = None,
    volume_tolerance_mm3: float = 1.0e-3,
) -> dict[str, Any]:
    warning_list = list(warnings or [])
    feature_created = feature is not None
    rollback = _rollback_feature(
        conn,
        feature,
        volume_before=volume_before,
        volume_tolerance_mm3=volume_tolerance_mm3,
    ) if feature_created else {
        "delete_succeeded": True,
        "update_succeeded": True,
        "volume_restored": None,
        "verified": True,
        "errors": [],
    }
    feature_persisted = feature_created and not bool(rollback.get("verified"))

    data: dict[str, Any] = {
        "type": feature_type,
        "feature_created": feature_created,
        "feature_persisted": feature_persisted,
        "update_succeeded": False,
        "rollback_succeeded": bool(rollback.get("verified")),
        "rollback": rollback,
    }
    if extra_data:
        data.update(extra_data)

    if feature_persisted:
        warning_list.append(
            "The operation failed and rollback was not fully verified; the CATPart may have been modified."
        )
        return _error(
            message,
            data=data,
            warnings=warning_list,
            status="partial_success",
        )

    return _error(message, data=data, warnings=warning_list)



def _apply_pad_direction(pad: Any, direction: str, symmetric: bool) -> dict[str, Any]:
    use_symmetric = bool(symmetric) or direction == "both"
    if use_symmetric:
        try:
            pad.IsSymmetric = True
        except Exception as exc:
            raise CATIAError(f"Cannot enable symmetric Pad: {exc}") from exc
        return {
            "direction_requested": direction,
            "direction_resolved": "both",
            "symmetric": True,
            "orientation": None,
        }

    orientation = _set_prism_orientation(
        pad,
        _CAT_INVERSE_ORIENTATION if direction == "reverse" else _CAT_REGULAR_ORIENTATION,
    )
    return {
        "direction_requested": direction,
        "direction_resolved": orientation["name"],
        "symmetric": False,
        "orientation": orientation,
    }



def _apply_pocket_direction(pocket: Any, direction: str) -> dict[str, Any]:
    orientation = _set_prism_orientation(
        pocket,
        _CAT_INVERSE_ORIENTATION if direction == "reverse" else _CAT_REGULAR_ORIENTATION,
    )
    return {
        "direction_requested": direction,
        "direction_resolved": orientation["name"],
        "orientation": orientation,
    }



def _set_angle_parameter(feature: Any, attribute_name: str, value: float) -> dict[str, Any]:
    errors: list[str] = []

    try:
        parameter = getattr(feature, attribute_name)
        parameter.Value = float(value)
        try:
            actual = float(parameter.Value)
        except Exception:
            actual = float(value)
        return {
            "attribute": attribute_name,
            "strategy": "parameter.Value",
            "requested": float(value),
            "actual": actual,
        }
    except Exception as exc:
        errors.append(f"parameter.Value: {exc}")

    try:
        setattr(feature, attribute_name, float(value))
        try:
            actual_value = getattr(feature, attribute_name)
            if hasattr(actual_value, "Value"):
                actual = float(actual_value.Value)
            else:
                actual = float(actual_value)
        except Exception:
            actual = float(value)
        return {
            "attribute": attribute_name,
            "strategy": "direct_attribute",
            "requested": float(value),
            "actual": actual,
            "fallback_errors": errors,
        }
    except Exception as exc:
        errors.append(f"direct_attribute: {exc}")

    raise CATIAError(
        f"Cannot set {attribute_name} to {value}. Attempts: {'; '.join(errors)}"
    )


def _create_pad_impl(
    conn: Any,
    *,
    height: float,
    sketch_name: str,
    direction: str,
    symmetric: bool,
    feature_name: str,
    require_geometry_change: bool = False,
    volume_tolerance_mm3: float = 1.0e-3,
) -> dict[str, Any]:
    height_value = _finite_positive(height, "height")
    requested_direction = _normalise_direction(
        direction, {"normal", "reverse", "both", "auto"}
    )
    # A symmetric Pad has no single material-side orientation.  Treat an
    # auto+symmetric request as the explicit bidirectional mode rather than
    # attempting meaningless regular/inverse retries.
    direction_key = (
        "both" if bool(symmetric) and requested_direction == "auto"
        else requested_direction
    )
    tolerance = _finite_positive(volume_tolerance_mm3, "volume_tolerance_mm3")

    conn.ensure_connected()
    part = conn.get_active_part()
    body = conn.get_active_part_body()
    sketch = _get_sketch(conn, sketch_name)
    shape_factory, factory_details = _get_shape_factory(part, "AddNewPad")
    part.InWorkObject = body

    feature = None
    warnings: list[str] = []
    volume_before = _body_volume_snapshot(conn, part, body)
    orientation_attempts: list[dict[str, Any]] = []
    try:
        feature = shape_factory.AddNewPad(sketch, height_value)
        initial_direction = "normal" if direction_key == "auto" else direction_key
        direction_details = _apply_pad_direction(feature, initial_direction, symmetric)
        direction_details["direction_requested"] = requested_direction
        if requested_direction == "auto" and direction_key == "both":
            direction_details["direction_resolved"] = "both"
        warnings.extend(_set_feature_name(feature, feature_name))

        updated, update_strategy, update_warnings = _update_feature(part, feature)
        warnings.extend(update_warnings)
        volume_after = _body_volume_snapshot(conn, part, body) if updated else {
            "available": False,
            "volume_m3": None,
            "volume_mm3": None,
            "method": None,
            "error": "Feature update failed before post-volume measurement.",
        }
        effect = _volume_effect(
            volume_before,
            volume_after,
            expected="increase",
            tolerance_mm3=tolerance,
        ) if updated else {
            "expected": "increase",
            "before": volume_before,
            "after": volume_after,
            "delta_mm3": None,
            "threshold_mm3": tolerance,
            "verified": False,
            "reason": "Feature update failed.",
        }
        orientation_attempts.append({
            "orientation": direction_details,
            "update_succeeded": updated,
            "update_strategy": update_strategy,
            "volume_effect": effect,
        })

        if direction_key == "auto" and (not updated or effect.get("verified") is not True):
            warnings.append(
                "The regular prism orientation did not produce a verified material increase; retrying inverse orientation."
            )
            inverse = _set_prism_orientation(feature, _CAT_INVERSE_ORIENTATION)
            updated, update_strategy, retry_warnings = _update_feature(part, feature)
            warnings.extend(retry_warnings)
            volume_after = _body_volume_snapshot(conn, part, body) if updated else {
                "available": False,
                "volume_m3": None,
                "volume_mm3": None,
                "method": None,
                "error": "Inverse-orientation update failed.",
            }
            effect = _volume_effect(
                volume_before,
                volume_after,
                expected="increase",
                tolerance_mm3=tolerance,
            ) if updated else {
                "expected": "increase",
                "before": volume_before,
                "after": volume_after,
                "delta_mm3": None,
                "threshold_mm3": tolerance,
                "verified": False,
                "reason": "Inverse-orientation update failed.",
            }
            direction_details = {
                "direction_requested": "auto",
                "direction_resolved": "reverse",
                "symmetric": False,
                "orientation": inverse,
            }
            orientation_attempts.append({
                "orientation": direction_details,
                "update_succeeded": updated,
                "update_strategy": update_strategy,
                "volume_effect": effect,
            })
        elif direction_key == "auto":
            direction_details = {
                **direction_details,
                "direction_requested": "auto",
                "direction_resolved": "normal",
            }

        if not updated:
            return _feature_failure(
                conn,
                feature,
                "Pad was created but CATIA could not update the feature in either permitted orientation.",
                feature_type="Pad",
                warnings=warnings,
                volume_before=volume_before,
                volume_tolerance_mm3=tolerance,
                extra_data={
                    "factory": factory_details,
                    "height": height_value,
                    "sketch": getattr(sketch, "Name", ""),
                    "direction": direction_details,
                    "orientation_attempts": orientation_attempts,
                    "update_strategy": update_strategy,
                    "geometry_verification": effect,
                },
            )

        if require_geometry_change and effect.get("verified") is not True:
            return _feature_failure(
                conn,
                feature,
                "Pad updated, but a material increase could not be verified. The feature was rolled back to avoid a false success.",
                feature_type="Pad",
                warnings=warnings,
                volume_before=volume_before,
                volume_tolerance_mm3=tolerance,
                extra_data={
                    "factory": factory_details,
                    "height": height_value,
                    "sketch": getattr(sketch, "Name", ""),
                    "direction": direction_details,
                    "orientation_attempts": orientation_attempts,
                    "update_strategy": update_strategy,
                    "geometry_verification": effect,
                    "require_geometry_change": True,
                },
            )

        if effect.get("verified") is None:
            warnings.append(
                "Pad update succeeded, but solid-volume verification was unavailable."
            )
        elif effect.get("verified") is False:
            warnings.append(
                "Pad update succeeded, but the measured solid volume did not increase."
            )

        conn.refresh_display()
        return _success(
            {
                "feature": getattr(feature, "Name", ""),
                "type": "Pad",
                "height": height_value,
                "sketch": getattr(sketch, "Name", ""),
                "feature_created": True,
                "feature_persisted": True,
                "update_succeeded": True,
                "update_strategy": update_strategy,
                "rollback_succeeded": None,
                "factory": factory_details,
                "direction": direction_details,
                "orientation_attempts": orientation_attempts,
                "geometry_verification": effect,
                "require_geometry_change": bool(require_geometry_change),
            },
            warnings,
        )
    except Exception as exc:
        return _feature_failure(
            conn,
            feature,
            str(exc),
            feature_type="Pad",
            warnings=warnings,
            volume_before=volume_before,
            volume_tolerance_mm3=tolerance,
            extra_data={
                "factory": factory_details,
                "height": height_value,
                "sketch": getattr(sketch, "Name", ""),
                "orientation_attempts": orientation_attempts,
            },
        )



def _create_pocket_impl(
    conn: Any,
    *,
    depth: float,
    sketch_name: str,
    direction: str,
    feature_name: str,
    limit_type: str = "dimension",
    require_material_removed: bool = False,
    volume_tolerance_mm3: float = 1.0e-3,
) -> dict[str, Any]:
    depth_value = _finite_positive(depth, "depth")
    direction_key = _normalise_direction(direction, {"normal", "reverse", "auto"})
    limit_key = _normalise_limit_type(limit_type)
    tolerance = _finite_positive(volume_tolerance_mm3, "volume_tolerance_mm3")

    conn.ensure_connected()
    part = conn.get_active_part()
    body = conn.get_active_part_body()
    sketch = _get_sketch(conn, sketch_name)
    shape_factory, factory_details = _get_shape_factory(part, "AddNewPocket")
    part.InWorkObject = body

    feature = None
    warnings: list[str] = []
    volume_before = _body_volume_snapshot(conn, part, body)
    orientation_attempts: list[dict[str, Any]] = []
    limit_details: dict[str, Any] | None = None
    try:
        feature = shape_factory.AddNewPocket(sketch, depth_value)
        limit_details = _apply_prism_limit(feature, limit_key, depth_value)
        initial_direction = "normal" if direction_key == "auto" else direction_key
        direction_details = _apply_pocket_direction(feature, initial_direction)
        warnings.extend(_set_feature_name(feature, feature_name))

        updated, update_strategy, update_warnings = _update_feature(part, feature)
        warnings.extend(update_warnings)
        volume_after = _body_volume_snapshot(conn, part, body) if updated else {
            "available": False,
            "volume_m3": None,
            "volume_mm3": None,
            "method": None,
            "error": "Feature update failed before post-volume measurement.",
        }
        effect = _volume_effect(
            volume_before,
            volume_after,
            expected="decrease",
            tolerance_mm3=tolerance,
        ) if updated else {
            "expected": "decrease",
            "before": volume_before,
            "after": volume_after,
            "delta_mm3": None,
            "threshold_mm3": tolerance,
            "verified": False,
            "reason": "Feature update failed.",
        }
        orientation_attempts.append({
            "orientation": direction_details,
            "update_succeeded": updated,
            "update_strategy": update_strategy,
            "volume_effect": effect,
        })

        if direction_key == "auto" and (not updated or effect.get("verified") is not True):
            warnings.append(
                "The regular prism orientation did not produce a verified material decrease; retrying inverse orientation."
            )
            inverse = _set_prism_orientation(feature, _CAT_INVERSE_ORIENTATION)
            updated, update_strategy, retry_warnings = _update_feature(part, feature)
            warnings.extend(retry_warnings)
            volume_after = _body_volume_snapshot(conn, part, body) if updated else {
                "available": False,
                "volume_m3": None,
                "volume_mm3": None,
                "method": None,
                "error": "Inverse-orientation update failed.",
            }
            effect = _volume_effect(
                volume_before,
                volume_after,
                expected="decrease",
                tolerance_mm3=tolerance,
            ) if updated else {
                "expected": "decrease",
                "before": volume_before,
                "after": volume_after,
                "delta_mm3": None,
                "threshold_mm3": tolerance,
                "verified": False,
                "reason": "Inverse-orientation update failed.",
            }
            direction_details = {
                "direction_requested": "auto",
                "direction_resolved": "reverse",
                "orientation": inverse,
            }
            orientation_attempts.append({
                "orientation": direction_details,
                "update_succeeded": updated,
                "update_strategy": update_strategy,
                "volume_effect": effect,
            })
        elif direction_key == "auto":
            direction_details = {
                **direction_details,
                "direction_requested": "auto",
                "direction_resolved": "normal",
            }

        if not updated:
            return _feature_failure(
                conn,
                feature,
                "Pocket was created but CATIA could not update the feature in either permitted orientation.",
                feature_type="Pocket",
                warnings=warnings,
                volume_before=volume_before,
                volume_tolerance_mm3=tolerance,
                extra_data={
                    "factory": factory_details,
                    "depth": depth_value,
                    "limit": limit_details,
                    "sketch": getattr(sketch, "Name", ""),
                    "direction": direction_details,
                    "orientation_attempts": orientation_attempts,
                    "update_strategy": update_strategy,
                    "geometry_verification": effect,
                },
            )

        if require_material_removed and effect.get("verified") is not True:
            return _feature_failure(
                conn,
                feature,
                "Pocket updated, but material removal could not be verified. The feature was rolled back to avoid a false success.",
                feature_type="Pocket",
                warnings=warnings,
                volume_before=volume_before,
                volume_tolerance_mm3=tolerance,
                extra_data={
                    "factory": factory_details,
                    "depth": depth_value,
                    "limit": limit_details,
                    "sketch": getattr(sketch, "Name", ""),
                    "direction": direction_details,
                    "orientation_attempts": orientation_attempts,
                    "update_strategy": update_strategy,
                    "geometry_verification": effect,
                    "require_material_removed": True,
                },
            )

        if effect.get("verified") is None:
            warnings.append(
                "Pocket update succeeded, but solid-volume verification was unavailable."
            )
        elif effect.get("verified") is False:
            warnings.append(
                "Pocket update succeeded, but the measured solid volume did not decrease."
            )

        conn.refresh_display()
        return _success(
            {
                "feature": getattr(feature, "Name", ""),
                "type": "Pocket",
                "depth": depth_value,
                "limit": limit_details,
                "sketch": getattr(sketch, "Name", ""),
                "feature_created": True,
                "feature_persisted": True,
                "update_succeeded": True,
                "update_strategy": update_strategy,
                "rollback_succeeded": None,
                "factory": factory_details,
                "direction": direction_details,
                "orientation_attempts": orientation_attempts,
                "geometry_verification": effect,
                "require_material_removed": bool(require_material_removed),
            },
            warnings,
        )
    except Exception as exc:
        return _feature_failure(
            conn,
            feature,
            str(exc),
            feature_type="Pocket",
            warnings=warnings,
            volume_before=volume_before,
            volume_tolerance_mm3=tolerance,
            extra_data={
                "factory": factory_details,
                "depth": depth_value,
                "limit": limit_details,
                "sketch": getattr(sketch, "Name", ""),
                "orientation_attempts": orientation_attempts,
            },
        )



def _create_circle_sketch(
    conn: Any,
    *,
    radius: float,
    center_x: float,
    center_y: float,
    plane: str,
    offset: float,
    sketch_name: str,
) -> Any:
    radius_value = _finite_positive(radius, "radius")
    center_x_value = _finite_number(center_x, "center_x")
    center_y_value = _finite_number(center_y, "center_y")
    offset_value = _finite_number(offset, "offset")

    part = conn.get_active_part()
    body = conn.get_active_part_body()
    plane_ref = conn.create_offset_plane_reference(str(plane).strip().lower(), offset_value)

    sketch = None
    opened = False
    try:
        part.InWorkObject = body
        sketch = body.Sketches.Add(plane_ref)
        sketch.Name = str(sketch_name).strip()
        factory = sketch.OpenEdition()
        opened = True
        geometry_before = int(sketch.GeometricElements.Count)
        factory.CreateClosedCircle(center_x_value, center_y_value, radius_value)
        geometry_after = int(sketch.GeometricElements.Count)
        if geometry_after != geometry_before + 1:
            raise CATIAError(
                "Circular sketch geometry verification failed: "
                f"before={geometry_before}, after={geometry_after}."
            )
        sketch.CloseEdition()
        opened = False

        updated, _, warnings = _update_feature(part, sketch)
        if not updated:
            raise CATIAError("; ".join(warnings))
        return sketch
    except Exception:
        if opened and sketch is not None:
            try:
                sketch.CloseEdition()
            except Exception:
                pass
        if sketch is not None:
            _delete_object(conn, sketch)
        raise


def _cleanup_failed_circle_sketch(conn: Any, sketch: Any, result: dict[str, Any]) -> None:
    data = result.get("data") if isinstance(result, dict) else None
    feature_persisted = bool(data.get("feature_persisted")) if isinstance(data, dict) else False
    if result.get("ok") or feature_persisted:
        return

    removed = _delete_object(conn, sketch)
    if isinstance(data, dict):
        data["orphan_sketch_cleanup_succeeded"] = removed
    if not removed:
        result.setdefault("warnings", []).append(
            "Feature creation failed and the temporary circle sketch could not be removed."
        )
        result["status"] = "partial_success"



def _normalise_axis_mode(axis_mode: str, axis_element_name: str = "") -> str:
    aliases = {
        "v": "vertical",
        "vertical_axis": "vertical",
        "vertical-reference": "vertical",
        "h": "horizontal",
        "horizontal_axis": "horizontal",
        "horizontal-reference": "horizontal",
        "named": "explicit",
        "line": "explicit",
    }
    mode = str(axis_mode or "vertical").strip().lower()
    mode = aliases.get(mode, mode)

    # Supplying a named element is an explicit request even if the caller kept
    # the default axis_mode for backward compatibility.
    if str(axis_element_name).strip() and mode != "explicit":
        mode = "explicit"

    if mode not in {"vertical", "horizontal", "explicit"}:
        raise CATIAError(
            "axis_mode must be one of: vertical, horizontal, explicit."
        )
    return mode


def _get_sketch_element(sketch: Any, element_name: str) -> Any:
    name = str(element_name).strip()
    if not name:
        raise CATIAError("axis_element_name cannot be empty in explicit mode.")

    elements = sketch.GeometricElements
    try:
        return elements.Item(name)
    except Exception as direct_error:
        try:
            for index in range(1, int(elements.Count) + 1):
                candidate = elements.Item(index)
                candidate_name = str(getattr(candidate, "Name", "")).strip()
                if candidate_name.casefold() == name.casefold():
                    return candidate
        except Exception:
            pass
        raise CATIAError(
            f"Axis element not found in sketch '{getattr(sketch, 'Name', '')}': {name}"
        ) from direct_error


def _geometric_type_value(element: Any) -> Any:
    try:
        value = element.GeometricType
        try:
            return int(value)
        except (TypeError, ValueError):
            return str(value)
    except Exception:
        return None


def _validate_axis_line(element: Any, *, source: str) -> dict[str, Any]:
    geometric_type = _geometric_type_value(element)
    object_details = _describe_com_object(element)
    type_text = (
        f"{object_details.get('python_type', '')} "
        f"{object_details.get('python_module', '')}"
    ).lower()

    # GeometricType=3 is Line2D in CATIA Sketcher.  Generated wrappers may
    # expose Line2D directly without GeometricType, so retain a type-name
    # fallback for those environments.
    is_line = geometric_type == 3 or "line2d" in type_text
    if not is_line:
        raise CATIAError(
            f"The selected revolution axis from {source} is not a Line2D "
            f"(GeometricType={geometric_type!r}, "
            f"python_type={object_details.get('python_type')})."
        )

    return {
        "geometric_type": geometric_type,
        "object": object_details,
    }


def _reference_display_name(reference: Any) -> str:
    for attribute_name in ("DisplayName", "Name"):
        try:
            value = str(getattr(reference, attribute_name)).strip()
            if value:
                return value
        except Exception:
            continue
    return ""


def _resolve_revolution_axis(
    conn: Any,
    part: Any,
    profile_sketch: Any,
    *,
    axis_mode: str,
    axis_sketch_name: str,
    axis_element_name: str,
) -> tuple[Any, dict[str, Any], list[str]]:
    mode = _normalise_axis_mode(axis_mode, axis_element_name)
    warnings: list[str] = []

    requested_axis_sketch = str(axis_sketch_name).strip()
    requested_axis_element = str(axis_element_name).strip()

    profile_sketch_name = str(getattr(profile_sketch, "Name", "")).strip()
    same_sketch = True

    if mode == "explicit":
        if not requested_axis_sketch:
            raise CATIAError(
                "axis_sketch_name is required when axis_mode='explicit'."
            )
        if not requested_axis_element:
            raise CATIAError(
                "axis_element_name is required when axis_mode='explicit'."
            )

        # Do not compare pywin32 COM proxies with ``is``.  Re-fetching the same
        # CATIA Sketch can produce a different Python wrapper even though both
        # wrappers refer to the same native object.  Compare normalized CATIA
        # sketch names and, when they match, reuse the original profile proxy.
        same_sketch = (
            bool(profile_sketch_name)
            and requested_axis_sketch.casefold() == profile_sketch_name.casefold()
        )

        if same_sketch:
            axis_sketch = profile_sketch
            strategy = "explicit_same_sketch_line"
        else:
            axis_sketch = _get_sketch(conn, requested_axis_sketch)
            strategy = "explicit_cross_sketch_line"

        axis_element = _get_sketch_element(axis_sketch, requested_axis_element)
        source = (
            f"sketch '{getattr(axis_sketch, 'Name', requested_axis_sketch)}' "
            f"element '{requested_axis_element}'"
        )
    else:
        axis_sketch = profile_sketch
        try:
            absolute_axis = profile_sketch.AbsoluteAxis
            property_name = (
                "VerticalReference" if mode == "vertical" else "HorizontalReference"
            )
            axis_element = getattr(absolute_axis, property_name)
        except Exception as exc:
            raise CATIAError(
                f"Cannot access the profile sketch absolute {mode} axis: {exc}"
            ) from exc
        strategy = f"profile_absolute_axis_{mode}"
        source = f"profile sketch absolute {mode} axis"

    line_details = _validate_axis_line(axis_element, source=source)

    # For axes belonging to the profile sketch, also mark the line as the
    # sketch center line before creating the Shaft/Groove.  CATIA uses a
    # sketch CenterLine as the default revolution axis.  Failure here is not
    # fatal because RevoluteAxis is assigned explicitly after feature creation.
    center_line_assigned = False
    if same_sketch:
        try:
            profile_sketch.CenterLine = axis_element
            center_line_assigned = True
        except Exception as exc:
            warnings.append(
                f"Profile sketch CenterLine assignment failed; "
                f"RevoluteAxis assignment will still be attempted: {exc}"
            )

    try:
        axis_reference = part.CreateReferenceFromObject(axis_element)
    except Exception as exc:
        raise CATIAError(
            f"Cannot create a CATIA Reference from the selected revolution axis: {exc}"
        ) from exc

    details = {
        "requested_mode": str(axis_mode or "vertical"),
        "resolved_mode": mode,
        "strategy": strategy,
        "profile_sketch": profile_sketch_name,
        "axis_sketch": str(getattr(axis_sketch, "Name", "")),
        "same_sketch": same_sketch,
        "axis_element": str(
            getattr(axis_element, "Name", requested_axis_element or property_name)
        ),
        "center_line_assigned": center_line_assigned,
        "reference_display_name": _reference_display_name(axis_reference),
        **line_details,
    }
    return axis_reference, details, warnings


def _assign_revolute_axis(feature: Any, axis_reference: Any) -> dict[str, Any]:
    errors: list[str] = []

    try:
        feature.RevoluteAxis = axis_reference
        try:
            readback = feature.RevoluteAxis
            verified = readback is not None
            readback_name = _reference_display_name(readback)
        except Exception as exc:
            verified = False
            readback_name = ""
            errors.append(f"direct readback: {exc}")
        return {
            "strategy": "direct_attribute",
            "dispatch_used": False,
            "verified": verified,
            "readback_display_name": readback_name,
            "warnings": errors,
        }
    except Exception as exc:
        errors.append(f"direct_attribute: {exc}")

    try:
        import win32com.client  # type: ignore

        dispatched_feature = win32com.client.Dispatch(feature)
        dispatched_feature.RevoluteAxis = axis_reference
        try:
            readback = dispatched_feature.RevoluteAxis
            verified = readback is not None
            readback_name = _reference_display_name(readback)
        except Exception as exc:
            verified = False
            readback_name = ""
            errors.append(f"dispatch readback: {exc}")
        return {
            "strategy": "dynamic_dispatch_attribute",
            "dispatch_used": True,
            "verified": verified,
            "readback_display_name": readback_name,
            "warnings": errors,
        }
    except Exception as exc:
        errors.append(f"dynamic_dispatch_attribute: {exc}")

    raise CATIAError(
        "Cannot assign Revolution.RevoluteAxis. Attempts: " + "; ".join(errors)
    )


def _create_revolution_impl(
    conn: Any,
    *,
    feature_type: str,
    angle: float,
    sketch_name: str,
    feature_name: str,
    axis_mode: str = "vertical",
    axis_sketch_name: str = "",
    axis_element_name: str = "",
) -> dict[str, Any]:
    angle_value = _finite_positive(angle, "angle")
    if angle_value > 360.0:
        raise CATIAError("angle must be less than or equal to 360 degrees.")

    method_name = "AddNewShaft" if feature_type == "Shaft" else "AddNewGroove"

    conn.ensure_connected()
    part = conn.get_active_part()
    body = conn.get_active_part_body()
    sketch = _get_sketch(conn, sketch_name)
    shape_factory, factory_details = _get_shape_factory(part, method_name)
    part.InWorkObject = body

    feature = None
    warnings: list[str] = []
    axis_details: dict[str, Any] | None = None
    axis_assignment: dict[str, Any] | None = None

    try:
        axis_reference, axis_details, axis_warnings = _resolve_revolution_axis(
            conn,
            part,
            sketch,
            axis_mode=axis_mode,
            axis_sketch_name=axis_sketch_name,
            axis_element_name=axis_element_name,
        )
        warnings.extend(axis_warnings)

        feature = getattr(shape_factory, method_name)(sketch)
        axis_assignment = _assign_revolute_axis(feature, axis_reference)
        warnings.extend(axis_assignment.pop("warnings", []))

        first_angle = _set_angle_parameter(feature, "FirstAngle", angle_value)

        second_angle: dict[str, Any] | None = None
        try:
            second_angle = _set_angle_parameter(feature, "SecondAngle", 0.0)
        except Exception as exc:
            warnings.append(f"SecondAngle could not be set to 0: {exc}")

        warnings.extend(_set_feature_name(feature, feature_name))

        updated, update_strategy, update_warnings = _update_feature(part, feature)
        warnings.extend(update_warnings)
        angle_data = {
            "first_angle": first_angle,
            "second_angle": second_angle,
        }
        axis_data = {
            "axis": axis_details,
            "axis_assignment": axis_assignment,
        }

        if not updated:
            return _feature_failure(
                conn,
                feature,
                f"{feature_type} was created but CATIA could not update the feature.",
                feature_type=feature_type,
                warnings=warnings,
                extra_data={
                    "factory": factory_details,
                    "angle": angle_value,
                    "sketch": getattr(sketch, "Name", ""),
                    "update_strategy": update_strategy,
                    **angle_data,
                    **axis_data,
                },
            )

        conn.refresh_display()
        return _success(
            {
                "feature": getattr(feature, "Name", ""),
                "type": feature_type,
                "angle": angle_value,
                "sketch": getattr(sketch, "Name", ""),
                "feature_created": True,
                "feature_persisted": True,
                "update_succeeded": True,
                "update_strategy": update_strategy,
                "rollback_succeeded": None,
                "factory": factory_details,
                **angle_data,
                **axis_data,
            },
            warnings,
        )
    except Exception as exc:
        return _feature_failure(
            conn,
            feature,
            str(exc),
            feature_type=feature_type,
            warnings=warnings,
            extra_data={
                "factory": factory_details,
                "angle": angle_value,
                "sketch": getattr(sketch, "Name", ""),
                "axis": axis_details,
                "axis_assignment": axis_assignment,
            },
        )


def register_tools(mcp: Any, ctx: Any) -> list[str]:
    conn = ctx.conn
    names: list[str] = []

    @mcp.tool()
    def catia_pad(
        height: float,
        sketch_name: str = "",
        direction: str = "normal",
        symmetric: bool = False,
        feature_name: str = "",
        require_geometry_change: bool = False,
        volume_tolerance_mm3: float = 0.001,
    ) -> dict[str, Any]:
        """Create a Pad and optionally require a measured material increase."""
        try:
            return _create_pad_impl(
                conn,
                height=height,
                sketch_name=sketch_name,
                direction=direction,
                symmetric=symmetric,
                feature_name=feature_name,
                require_geometry_change=require_geometry_change,
                volume_tolerance_mm3=volume_tolerance_mm3,
            )
        except Exception as exc:
            return _error(str(exc))

    names.append("catia_pad")

    @mcp.tool()
    def catia_pocket(
        depth: float,
        sketch_name: str = "",
        direction: str = "normal",
        feature_name: str = "",
        limit_type: str = "dimension",
        require_material_removed: bool = False,
        volume_tolerance_mm3: float = 0.001,
    ) -> dict[str, Any]:
        """Create a Pocket with dimensional or through-solid limit modes."""
        try:
            return _create_pocket_impl(
                conn,
                depth=depth,
                sketch_name=sketch_name,
                direction=direction,
                feature_name=feature_name,
                limit_type=limit_type,
                require_material_removed=require_material_removed,
                volume_tolerance_mm3=volume_tolerance_mm3,
            )
        except Exception as exc:
            return _error(str(exc))

    names.append("catia_pocket")

    @mcp.tool()
    def catia_circular_pad(
        radius: float,
        height: float,
        center_x: float = 0.0,
        center_y: float = 0.0,
        plane: str = "xy",
        offset: float = 0.0,
        sketch_name: str = "",
        feature_name: str = "",
        direction: str = "auto",
        require_geometry_change: bool = True,
        volume_tolerance_mm3: float = 0.001,
    ) -> dict[str, Any]:
        """Create a circular boss, auto-resolving offset-plane extrusion direction."""
        sketch = None
        try:
            conn.ensure_connected()
            radius_value = _finite_positive(radius, "radius")
            height_value = _finite_positive(height, "height")
            plane_key = _normalise_plane(plane)
            _normalise_direction(direction, {"normal", "reverse", "auto"})
            final_sketch_name = str(sketch_name).strip() or "MCP_CircularPadSketch"
            contract = _plane_coordinate_contract(
                plane_key,
                offset,
                center_x,
                center_y,
            )
            sketch = _create_circle_sketch(
                conn,
                radius=radius_value,
                center_x=center_x,
                center_y=center_y,
                plane=plane_key,
                offset=offset,
                sketch_name=final_sketch_name,
            )
            result = _create_pad_impl(
                conn,
                height=height_value,
                sketch_name=getattr(sketch, "Name", final_sketch_name),
                direction=direction,
                symmetric=False,
                feature_name=str(feature_name).strip() or "MCP_CircularPad",
                require_geometry_change=require_geometry_change,
                volume_tolerance_mm3=volume_tolerance_mm3,
            )
            if isinstance(result.get("data"), dict):
                result["data"].update(
                    {
                        "radius": radius_value,
                        "nominal_profile_area_mm2": math.pi * radius_value * radius_value,
                        "nominal_untrimmed_volume_mm3": math.pi * radius_value * radius_value * height_value,
                        "coordinate_contract": contract,
                    }
                )
            _cleanup_failed_circle_sketch(conn, sketch, result)
            return result
        except Exception as exc:
            if sketch is not None:
                _delete_object(conn, sketch)
            return _error(str(exc))

    names.append("catia_circular_pad")

    @mcp.tool()
    def catia_cut_circular_hole(
        diameter: float,
        depth: float,
        center_x: float = 0.0,
        center_y: float = 0.0,
        plane: str = "xy",
        offset: float = 0.0,
        direction: str = "auto",
        sketch_name: str = "",
        feature_name: str = "",
        limit_type: str = "through_all",
        require_material_removed: bool = True,
        volume_tolerance_mm3: float = 0.001,
    ) -> dict[str, Any]:
        """Create a circular Pocket, defaulting to verified through-all removal."""
        sketch = None
        try:
            conn.ensure_connected()
            diameter_value = _finite_positive(diameter, "diameter")
            depth_value = _finite_positive(depth, "depth")
            plane_key = _normalise_plane(plane)
            _normalise_direction(direction, {"normal", "reverse", "auto"})
            limit_key = _normalise_limit_type(limit_type)
            final_sketch_name = str(sketch_name).strip() or "MCP_CircularHoleSketch"
            contract = _plane_coordinate_contract(
                plane_key,
                offset,
                center_x,
                center_y,
            )
            sketch = _create_circle_sketch(
                conn,
                radius=diameter_value / 2.0,
                center_x=center_x,
                center_y=center_y,
                plane=plane_key,
                offset=offset,
                sketch_name=final_sketch_name,
            )
            result = _create_pocket_impl(
                conn,
                depth=depth_value,
                sketch_name=getattr(sketch, "Name", final_sketch_name),
                direction=direction,
                feature_name=str(feature_name).strip() or "MCP_CircularHole",
                limit_type=limit_key,
                require_material_removed=require_material_removed,
                volume_tolerance_mm3=volume_tolerance_mm3,
            )
            if isinstance(result.get("data"), dict):
                result["data"].update(
                    {
                        "diameter": diameter_value,
                        "radius": diameter_value / 2.0,
                        "coordinate_contract": contract,
                        "through_all_requested": limit_key == "through_all",
                    }
                )
            _cleanup_failed_circle_sketch(conn, sketch, result)
            return result
        except Exception as exc:
            if sketch is not None:
                _delete_object(conn, sketch)
            return _error(str(exc))

    names.append("catia_cut_circular_hole")

    @mcp.tool()
    def catia_shaft(
        angle: float = 360.0,
        sketch_name: str = "",
        feature_name: str = "",
        axis_mode: str = "vertical",
        axis_sketch_name: str = "",
        axis_element_name: str = "",
    ) -> dict[str, Any]:
        """Create a Shaft with vertical, horizontal, or explicit revolution axis."""
        try:
            return _create_revolution_impl(
                conn,
                feature_type="Shaft",
                angle=angle,
                sketch_name=sketch_name,
                feature_name=feature_name,
                axis_mode=axis_mode,
                axis_sketch_name=axis_sketch_name,
                axis_element_name=axis_element_name,
            )
        except Exception as exc:
            return _error(str(exc))

    names.append("catia_shaft")

    @mcp.tool()
    def catia_groove(
        angle: float = 360.0,
        sketch_name: str = "",
        feature_name: str = "",
        axis_mode: str = "vertical",
        axis_sketch_name: str = "",
        axis_element_name: str = "",
    ) -> dict[str, Any]:
        """Create a Groove with vertical, horizontal, or explicit revolution axis."""
        try:
            return _create_revolution_impl(
                conn,
                feature_type="Groove",
                angle=angle,
                sketch_name=sketch_name,
                feature_name=feature_name,
                axis_mode=axis_mode,
                axis_sketch_name=axis_sketch_name,
                axis_element_name=axis_element_name,
            )
        except Exception as exc:
            return _error(str(exc))

    names.append("catia_groove")

    @mcp.tool()
    def catia_list_features() -> dict[str, Any]:
        """List PartBody shape features and return a current solid-volume snapshot."""
        try:
            conn.ensure_connected()
            part = conn.get_active_part()
            body = conn.get_active_part_body()
            shapes = body.Shapes
            result: list[dict[str, Any]] = []

            for index in range(1, int(shapes.Count) + 1):
                item = shapes.Item(index)
                result.append(
                    {
                        "index": index,
                        "name": getattr(item, "Name", ""),
                        "type": getattr(item, "Type", ""),
                    }
                )

            return _success(
                {
                    "features": result,
                    "shape_count": int(shapes.Count),
                    "sketch_count": int(body.Sketches.Count),
                    "solid_volume": _body_volume_snapshot(conn, part, body),
                }
            )
        except Exception as exc:
            return _error(str(exc))

    names.append("catia_list_features")

    @mcp.tool()
    def catia_update_active_part() -> dict[str, Any]:
        """Update/rebuild the active CATPart and return actual volume evidence."""
        try:
            conn.ensure_connected()
            part = conn.get_active_part()
            body = conn.get_active_part_body()
            before = _body_volume_snapshot(conn, part, body)
            part.Update()
            after = _body_volume_snapshot(conn, part, body)
            conn.refresh_display()
            return _success(
                {
                    "message": "Active part updated.",
                    "update_succeeded": True,
                    "volume_before": before,
                    "volume_after": after,
                }
            )
        except Exception as exc:
            return _error(str(exc))

    names.append("catia_update_active_part")

    return names



# ---------------------------------------------------------------------------
# Registry Entry Point (Required by registry.py)
# ---------------------------------------------------------------------------
