"""Column Internals tools for RadFrac (Aspen Plus V15 / 41.0).

Column Internals live in the block's own Input tree, keyed by configuration
and section:  Data\\Blocks\\{blk}\\Input\\CA_xxx\\{config}\\{section}.
A configuration (INT-1, ...) is created with
Subobjects\\Column Internals.Elements.Add(); a section is a labelled row,
created with InsertRow + SetLabel on CA_STAGE1\\{config} (Elements.Add is
"not yet implemented" there).  Results go to Output\\CA_xxx\\{config}\\{section}.

All COM operations run inside aspen.call() lambdas on the dedicated
COM apartment thread.
"""

from __future__ import annotations

import math

from ..com_bridge import aspen


def _walk(root, *parts):
    node = root
    for p in parts:
        try:
            node = node.Elements(p)
        except Exception:
            return None
        if node is None:
            return None
    return node


def _kids(node) -> list[str]:
    if node is None:
        return []
    try:
        return [node.Elements.Item(i).Name for i in range(node.Elements.Count)]
    except Exception:
        return []


def _val(node):
    if node is None:
        return None
    try:
        return node.Value
    except Exception:
        return None


def _unit(node) -> str:
    try:
        return node.UnitString or ""
    except Exception:
        return ""


# Length units the Aspen unit sets use for internals geometry.
_LEN_TO_M = {"meter": 1.0, "m": 1.0, "mm": 1e-3, "cm": 1e-2, "ft": 0.3048, "in": 0.0254}


def _set_length(node, value_m: float):
    """Write a length given in metres, converting to the node's display unit."""
    u = _unit(node).strip().lower()
    factor = _LEN_TO_M.get(u)
    node.Value = value_m / factor if factor else value_m


def _get_length_m(node):
    v = _val(node)
    if v is None:
        return None
    factor = _LEN_TO_M.get(_unit(node).strip().lower())
    return v * factor if factor else v


def _options(node) -> list[str] | None:
    try:
        opt = node.AttributeValue(5)
    except Exception:
        return None
    if opt is None:
        return None
    try:
        vals = [opt.Elements.Item(i).Name for i in range(opt.Elements.Count)]
        return vals or None
    except Exception:
        return None


def _active_config(inp, config: str | None) -> str | None:
    if config:
        return config
    return _val(_walk(inp, "CA_CONFIG"))


def _sections(inp, config: str) -> list[str]:
    return _kids(_walk(inp, "CA_STAGE1", config))


def _nstage(tree, block_name):
    return _val(_walk(tree, "Data", "Blocks", block_name, "Input", "NSTAGE"))


# Inputs worth showing for a section (key -> label).
_SECTION_INPUTS = {
    "CA_STAGE1": "start_stage", "CA_STAGE2": "end_stage",
    "CA_INTERNAL": "type", "CA_SIZING": "mode",
    "CA_TRAYTYPE": "tray_type", "CA_NPASS": "passes",
    "CA_TRAY_SPC": "tray_spacing", "CA_DIAM": "diameter",
    "CA_DECK_TCK": "deck_thickness", "CA_DECK_TCK2": "deck_thickness_value",
    "CA_WEIR_HT": "weir_height", "CA_HOLE_DIAM": "hole_diameter",
    "CA_HOLE_AREA": "hole_area_fraction", "CA_DC_CLEAR": "downcomer_clearance",
    "CA_EFF": "overall_efficiency", "CA_SYSFAC": "foaming_factor",
    "CA_P_UPDATE": "update_pressure",
    "CA_PACKTYPE": "packing_type", "CA_VENDOR": "vendor",
    "CA_PACK_MAT": "packing_material", "CA_PACK_SIZE": "packing_size",
    "CA_PACK_HT": "packed_height", "CA_HETP": "hetp",
}

# Section results (key -> label); trays and packing use different keys in V15.
_SECTION_RESULTS = {
    # packing
    "CA_DIAM2": "diameter", "CA_FLD_FAC2": "max_flooding_pct",
    "CA_DPSEC2": "section_pressure_drop", "CA_HEIGHT2": "section_height",
    "CA_HETP2": "hetp", "CA_HOLDUP2": "liquid_holdup",
    "CA_MAXFLOOD4": "stage_of_max_flooding",
    # trays, interactive sizing
    "CA_DIAM6": "tray_diameter", "CA_DPSEC4": "tray_section_pressure_drop",
    "CA_TS_TOTHT": "tray_section_height", "CA_TS_HLOSS": "tray_section_head_loss",
    "CA_MAXSTAGE": "limiting_stage", "CA_TS_MAXDC": "stage_of_max_downcomer_backup",
    # trays, rating
    "CA_FLD_FAC5": "max_flooding_pct", "CA_DCBSPACE1": "max_downcomer_backup_pct_of_spacing",
    "CA_DPSEC3": "rated_section_pressure_drop", "CA_TR_TOTHT": "rated_section_height",
    "CA_MAXFLOOD": "stage_of_max_flooding", "CA_MAXBACKUP": "stage_of_max_backup",
    "CA_DC_WIDTH1": "side_downcomer_width", "CA_WEIRLEN": "weir_length",
    "CA_FLOWP": "flow_path_length",
}


def _fmt(v, node=None):
    if v is None:
        return None
    u = _unit(node) if node is not None else ""
    if isinstance(v, float):
        v = round(v, 6)
    return f"{v} {u}".strip() if u else v


def tool_list_column_internals(block_name: str, config: str | None = None) -> dict:
    """Describe the Column Internals of a RadFrac: configs, sections, inputs, results."""
    try:
        def impl():
            tree = aspen._app.RootModel("")
            inp = _walk(tree, "Data", "Blocks", block_name, "Input")
            if inp is None:
                return {"error": f"block '{block_name}' not found"}
            out = _walk(tree, "Data", "Blocks", block_name, "Output")
            configs = _kids(_walk(tree, "Data", "Blocks", block_name, "Subobjects", "Column Internals"))
            cfg = _active_config(inp, config)
            res = {"block": block_name, "nstage": _nstage(tree, block_name),
                   "configs": configs, "active_config": _val(_walk(inp, "CA_CONFIG")),
                   "sections": {}}
            if not cfg or cfg not in configs:
                return res
            for sec in _sections(inp, cfg):
                d = {}
                for key, label in _SECTION_INPUTS.items():
                    n = _walk(inp, key, cfg, sec)
                    v = _val(n)
                    if v is not None:
                        d[label] = _fmt(v, n)
                r = {}
                for key, label in _SECTION_RESULTS.items():
                    n = _walk(out, key, cfg, sec)
                    v = _val(n)
                    if v not in (None, 0):
                        r[label] = _fmt(v, n)
                d["results"] = r
                res["sections"][sec] = d
            res["pressure_fix"] = _val(_walk(inp, "CA_P_FIX", cfg))
            return res
        return aspen.call(impl)
    except Exception as exc:
        return {"error": str(exc)}


def tool_add_column_internals(block_name: str, config: str = "INT-1") -> str:
    """Create a Column Internals configuration (e.g. INT-1) on a RadFrac and make it active."""
    try:
        def impl():
            tree = aspen._app.RootModel("")
            ci = _walk(tree, "Data", "Blocks", block_name, "Subobjects", "Column Internals")
            if ci is None:
                return f"Error: '{block_name}' has no Column Internals folder (is it a RadFrac?)"
            if config in _kids(ci):
                return f"{block_name}: internals '{config}' already exists"
            ci.Elements.Add(config)
            act = _walk(tree, "Data", "Blocks", block_name, "Input", "CA_CONFIG")
            if act is not None:
                act.Value = config
            return f"Created internals '{config}' on {block_name} (active)"
        return aspen.call(impl)
    except Exception as exc:
        return "Error: " + str(exc)


def tool_add_internals_section(block_name: str, section: str, stage_start: int,
                               stage_end: int, internal_type: str = "TRAY",
                               config: str | None = None) -> str:
    """Add a section (labelled row) to a Column Internals configuration.

    stage_start / stage_end are RadFrac stage numbers (with a condenser the
    first tray is stage 2; with a kettle reboiler the last tray is NSTAGE-1).
    internal_type: TRAY or PACKING.
    """
    internal_type = internal_type.upper()
    if internal_type not in ("TRAY", "PACKING"):
        return "Error: internal_type must be TRAY or PACKING"
    try:
        def impl():
            tree = aspen._app.RootModel("")
            inp = _walk(tree, "Data", "Blocks", block_name, "Input")
            if inp is None:
                return f"Error: block '{block_name}' not found"
            cfg = _active_config(inp, config)
            if not cfg:
                return "Error: no internals configuration; call add_column_internals first"
            n = _nstage(tree, block_name)
            if n and not (1 <= stage_start <= stage_end <= n):
                return f"Error: stages must satisfy 1 <= {stage_start} <= {stage_end} <= {n}"
            st1 = _walk(inp, "CA_STAGE1", cfg)
            if st1 is None:
                return f"Error: configuration '{cfg}' not found on {block_name}"
            if section not in _kids(st1):
                els = st1.Elements
                idx = els.Count
                els.InsertRow(0, idx)
                els.SetLabel(0, idx, False, section)
            _walk(inp, "CA_STAGE1", cfg, section).Value = stage_start
            _walk(inp, "CA_STAGE2", cfg, section).Value = stage_end
            _walk(inp, "CA_INTERNAL", cfg, section).Value = internal_type
            return (f"{block_name}/{cfg}: section {section} = stages {stage_start}-{stage_end}"
                    f" ({internal_type})")
        return aspen.call(impl)
    except Exception as exc:
        return "Error: " + str(exc)


def tool_remove_internals_section(block_name: str, section: str,
                                  config: str | None = None) -> str:
    """Remove a section from a Column Internals configuration."""
    try:
        def impl():
            tree = aspen._app.RootModel("")
            inp = _walk(tree, "Data", "Blocks", block_name, "Input")
            cfg = _active_config(inp, config)
            st1 = _walk(inp, "CA_STAGE1", cfg)
            names = _kids(st1)
            if section not in names:
                return f"Error: section '{section}' not found in {cfg} (have {names})"
            st1.Elements.RemoveRow(0, names.index(section))
            return f"Removed section {section} from {block_name}/{cfg}"
        return aspen.call(impl)
    except Exception as exc:
        return "Error: " + str(exc)


def _set_section(inp, cfg, sec, key, value):
    n = _walk(inp, key, cfg, sec)
    if n is None:
        raise KeyError(f"{key}\\{cfg}\\{sec} not found")
    n.Value = value
    return n


def tool_configure_tray_section(block_name: str, section: str,
                                tray_type: str = "SIEVE",
                                n_passes: int = 1,
                                tray_spacing_m: float | None = None,
                                mode: str = "INTER-SIZING",
                                diameter_m: float | None = None,
                                deck_thickness: str | None = None,
                                deck_thickness_mm: float | None = None,
                                weir_height_mm: float | None = None,
                                hole_diameter_mm: float | None = None,
                                hole_area_fraction: float | None = None,
                                downcomer_clearance_mm: float | None = None,
                                overall_efficiency: float | None = None,
                                foaming_factor: float | None = None,
                                update_pressure: bool | None = None,
                                config: str | None = None) -> str:
    """Set the geometry of a tray section.

    mode: INTER-SIZING (Aspen sizes the diameter for the flooding target)
          or RATING (diameter_m is required; Aspen reports flooding/backup).
    deck_thickness: '3 GAUGE', '10 GAUGE', '12 GAUGE', '14 GAUGE' or 'OTHER'
          (OTHER uses deck_thickness_mm).
    overall_efficiency: theoretical stages / actual trays (CA_EFF); Aspen uses
          it to turn the section's stages into real trays and height.
    Lengths are given in SI (m / mm) and converted to the block's unit set.
    """
    mode = mode.upper()
    if mode not in ("INTER-SIZING", "RATING"):
        return "Error: mode must be INTER-SIZING or RATING"
    if mode == "RATING" and diameter_m is None:
        return "Error: RATING mode needs diameter_m"
    if not 1 <= n_passes <= 4:
        return "Error: n_passes must be 1..4"
    try:
        def impl():
            tree = aspen._app.RootModel("")
            inp = _walk(tree, "Data", "Blocks", block_name, "Input")
            cfg = _active_config(inp, config)
            if section not in _sections(inp, cfg):
                return f"Error: section '{section}' not in {block_name}/{cfg}; add it first"
            done = []
            _set_section(inp, cfg, section, "CA_INTERNAL", "TRAY")
            n = _walk(inp, "CA_TRAYTYPE", cfg, section)
            opts = _options(n)
            if opts and tray_type.upper() not in opts:
                return f"Error: tray_type must be one of {opts}"
            n.Value = tray_type.upper()
            done.append(f"tray={tray_type.upper()}")
            _set_section(inp, cfg, section, "CA_NPASS", n_passes)
            done.append(f"passes={n_passes}")
            if tray_spacing_m is not None:
                _set_length(_walk(inp, "CA_TRAY_SPC", cfg, section), tray_spacing_m)
                done.append(f"spacing={tray_spacing_m} m")
            _set_section(inp, cfg, section, "CA_SIZING", mode)
            done.append(f"mode={mode}")
            if diameter_m is not None:
                _set_length(_walk(inp, "CA_DIAM", cfg, section), diameter_m)
                done.append(f"D={diameter_m} m")
            if deck_thickness is not None:
                _set_section(inp, cfg, section, "CA_DECK_TCK", deck_thickness.upper())
                done.append(f"deck={deck_thickness.upper()}")
            if deck_thickness_mm is not None:
                _set_length(_walk(inp, "CA_DECK_TCK2", cfg, section), deck_thickness_mm / 1000)
                done.append(f"deck={deck_thickness_mm} mm")
            for key, v, label in (("CA_WEIR_HT", weir_height_mm, "weir"),
                                  ("CA_HOLE_DIAM", hole_diameter_mm, "hole"),
                                  ("CA_DC_CLEAR", downcomer_clearance_mm, "dc_clear")):
                if v is not None:
                    _set_length(_walk(inp, key, cfg, section), v / 1000)
                    done.append(f"{label}={v} mm")
            if hole_area_fraction is not None:
                _set_section(inp, cfg, section, "CA_HOLE_AREA", hole_area_fraction)
                done.append(f"hole_area={hole_area_fraction}")
            if overall_efficiency is not None:
                _set_section(inp, cfg, section, "CA_EFF", overall_efficiency)
                done.append(f"eff={overall_efficiency}")
            if foaming_factor is not None:
                _set_section(inp, cfg, section, "CA_SYSFAC", foaming_factor)
                done.append(f"foaming={foaming_factor}")
            if update_pressure is not None:
                _set_section(inp, cfg, section, "CA_P_UPDATE", "YES" if update_pressure else "NO")
                done.append(f"update_dP={'YES' if update_pressure else 'NO'}")
                fix = _walk(inp, "CA_P_FIX", cfg)
                if update_pressure and fix is not None and _val(fix) in (None, "NOUPDATE"):
                    fix.Value = "TOP"  # top pressure held, hydraulics build the profile
                    done.append("P_FIX=TOP")
            return f"{block_name}/{cfg}/{section}: " + ", ".join(done)
        return aspen.call(impl)
    except Exception as exc:
        return "Error: " + str(exc)


def tool_configure_packed_section(block_name: str, section: str,
                                  packing_type: str, vendor: str | None = None,
                                  material: str | None = None, size: str | None = None,
                                  packed_height_m: float | None = None,
                                  hetp_m: float | None = None,
                                  mode: str = "INTER-SIZING",
                                  diameter_m: float | None = None,
                                  update_pressure: bool | None = None,
                                  config: str | None = None) -> str:
    """Set a packed section (e.g. MELLAPAK / SULZER / STANDARD / 250Y).

    Give packed_height_m or hetp_m (Aspen derives the other from the stage count).
    """
    mode = mode.upper()
    if mode == "RATING" and diameter_m is None:
        return "Error: RATING mode needs diameter_m"
    try:
        def impl():
            tree = aspen._app.RootModel("")
            inp = _walk(tree, "Data", "Blocks", block_name, "Input")
            cfg = _active_config(inp, config)
            if section not in _sections(inp, cfg):
                return f"Error: section '{section}' not in {block_name}/{cfg}; add it first"
            _set_section(inp, cfg, section, "CA_INTERNAL", "PACKING")
            _set_section(inp, cfg, section, "CA_PACKTYPE", packing_type.upper())
            if vendor:
                _set_section(inp, cfg, section, "CA_VENDOR", vendor.upper())
            if material:
                _set_section(inp, cfg, section, "CA_PACK_MAT", material.upper())
            if size:
                _set_section(inp, cfg, section, "CA_PACK_SIZE", size.upper())
            if packed_height_m is not None:
                _set_length(_walk(inp, "CA_PACK_HT", cfg, section), packed_height_m)
                n = _walk(inp, "OPT_CA_HETP", cfg, section)
                if n is not None:
                    n.Value = "HEIGHT"
            if hetp_m is not None:
                _set_length(_walk(inp, "CA_HETP", cfg, section), hetp_m)
                n = _walk(inp, "OPT_CA_HETP", cfg, section)
                if n is not None:
                    n.Value = "HETP"
            _set_section(inp, cfg, section, "CA_SIZING", mode)
            if diameter_m is not None:
                _set_length(_walk(inp, "CA_DIAM", cfg, section), diameter_m)
            if update_pressure is not None:
                _set_section(inp, cfg, section, "CA_P_UPDATE", "YES" if update_pressure else "NO")
                fix = _walk(inp, "CA_P_FIX", cfg)
                if update_pressure and fix is not None and _val(fix) in (None, "NOUPDATE"):
                    fix.Value = "TOP"
            return f"{block_name}/{cfg}/{section}: packing {packing_type.upper()} {size or ''} ({mode})"
        return aspen.call(impl)
    except Exception as exc:
        return "Error: " + str(exc)


def tool_set_internals_param(block_name: str, section: str | None, param: str,
                             value: str | float | int,
                             config: str | None = None) -> str:
    """Set any Column Internals input (CA_xxx) for a section, or for the whole
    configuration when section is None (e.g. CA_P_FIX, CA_LIQ_RTIME).
    Option-type inputs are validated against Aspen's option list."""
    try:
        def impl():
            tree = aspen._app.RootModel("")
            inp = _walk(tree, "Data", "Blocks", block_name, "Input")
            cfg = _active_config(inp, config)
            n = _walk(inp, param, cfg, section) if section else _walk(inp, param, cfg)
            if n is None:
                return f"Error: {param}\\{cfg}" + (f"\\{section}" if section else "") + " not found"
            opts = _options(n)
            if opts and isinstance(value, str) and value.upper() not in [o.upper() for o in opts]:
                return f"Error: {param} must be one of {opts}"
            n.Value = value.upper() if (opts and isinstance(value, str)) else value
            return f"Set {block_name} {param}[{cfg}{'/' + section if section else ''}] = {_fmt(_val(n), n)}"
        return aspen.call(impl)
    except Exception as exc:
        return "Error: " + str(exc)


def tool_describe_internals_params(block_name: str, name_filter: str = "",
                                   config: str | None = None) -> list[dict]:
    """List Column Internals inputs (CA_xxx) with Aspen's description, unit,
    option list and current value of the first section. Use name_filter
    (e.g. 'DC', 'WEIR', 'TRAY') to narrow the list."""
    try:
        def impl():
            tree = aspen._app.RootModel("")
            inp = _walk(tree, "Data", "Blocks", block_name, "Input")
            cfg = _active_config(inp, config)
            secs = _sections(inp, cfg) if cfg else []
            rows = []
            for i in range(inp.Elements.Count):
                p = inp.Elements.Item(i)
                name = p.Name
                if not name.startswith("CA_") or (name_filter and name_filter.upper() not in name):
                    continue
                leaf = _walk(p, cfg, secs[0]) if (cfg and secs) else None
                if leaf is None and cfg:
                    leaf = _walk(p, cfg)
                node = leaf if leaf is not None else p
                try:
                    prompt = str(node.AttributeValue(19))
                except Exception:
                    prompt = ""
                rows.append({"param": name, "description": prompt, "unit": _unit(node),
                             "options": _options(node), "value": _val(node)})
            return rows
        return aspen.call(impl)
    except Exception as exc:
        return [{"error": str(exc)}]


def tool_get_column_hydraulics(block_name: str, config: str | None = None,
                               per_stage: bool = False) -> dict:
    """Read Column Internals results: per section diameter, max flooding,
    downcomer backup, pressure drop, height; optionally per-stage flooding/dP."""
    try:
        def impl():
            tree = aspen._app.RootModel("")
            inp = _walk(tree, "Data", "Blocks", block_name, "Input")
            out = _walk(tree, "Data", "Blocks", block_name, "Output")
            cfg = _active_config(inp, config)
            if not cfg:
                return {"error": f"{block_name} has no Column Internals"}
            res = {"block": block_name, "config": cfg, "sections": {}}
            for sec in _sections(inp, cfg):
                d = {}
                for i in range(out.Elements.Count):
                    k = out.Elements.Item(i)
                    name = k.Name
                    if not name.startswith("CA_"):
                        continue
                    leaf = _walk(k, cfg, sec)
                    if leaf is None:
                        continue
                    try:
                        cnt = leaf.Elements.Count
                    except Exception:
                        cnt = 0
                    if cnt == 0:
                        v = _val(leaf)
                        if v not in (None, 0, ""):
                            d[name] = _fmt(v, leaf)
                    elif per_stage:
                        d[name] = {leaf.Elements.Item(j).Name: _fmt(_val(leaf.Elements.Item(j)), leaf.Elements.Item(j))
                                   for j in range(cnt)}
                res["sections"][sec] = d
            return res
        return aspen.call(impl)
    except Exception as exc:
        return {"error": str(exc)}


def _oconnell_eo(mu_cp: float, alpha: float) -> float:
    """O'Connell (1946) overall tray efficiency, as fitted by Lockett:
    Eo = 0.492 (mu_L alpha)^-0.245, mu_L in cP at average column conditions."""
    x = max(mu_cp * alpha, 1e-6)
    return 0.492 * x ** -0.245


def tool_oconnell_efficiency(block_name: str, light_key: str, heavy_key: str,
                             stage_from: int | None = None,
                             stage_to: int | None = None) -> dict:
    """Estimate the overall tray efficiency (O'Connell) from the converged
    column: alpha = K_LK/K_HK and liquid viscosity per stage, averaged over the
    trays (condenser and reboiler excluded by default)."""
    try:
        def impl():
            tree = aspen._app.RootModel("")
            out = _walk(tree, "Data", "Blocks", block_name, "Output")
            n = int(_nstage(tree, block_name))
            a, b = stage_from or 2, stage_to or n - 1
            rows = []
            for st in range(a, b + 1):
                k_lk = _val(_walk(out, "B_K", str(st), light_key))
                k_hk = _val(_walk(out, "B_K", str(st), heavy_key))
                if not (k_lk and k_hk):
                    # three-phase RadFrac leaves B_K empty (K1/K2 per liquid): use y/x
                    k_lk, k_hk = None, None
                    xl = _val(_walk(out, "X", str(st), light_key))
                    yl = _val(_walk(out, "Y", str(st), light_key))
                    xh = _val(_walk(out, "X", str(st), heavy_key))
                    yh = _val(_walk(out, "Y", str(st), heavy_key))
                    if xl and yl and xh and yh and xl > 1e-12 and xh > 1e-12:
                        k_lk, k_hk = yl / xl, yh / xh
                mu = _val(_walk(out, "HYD_MUL", str(st)))
                if not (k_lk and k_hk and mu):
                    continue
                alpha = k_lk / k_hk
                if alpha < 1:
                    alpha = 1 / alpha
                rows.append((st, alpha, mu, _oconnell_eo(mu, alpha)))
            if not rows:
                return {"error": "no converged stage data (run the simulation first)"}
            mu_avg = sum(r[2] for r in rows) / len(rows)
            alpha_avg = math.exp(sum(math.log(r[1]) for r in rows) / len(rows))
            eo = _oconnell_eo(mu_avg, alpha_avg)
            return {"block": block_name, "keys": [light_key, heavy_key],
                    "stages": [a, b], "alpha_geo_mean": round(alpha_avg, 3),
                    "mu_L_avg_cP": round(mu_avg, 4), "Eo": round(eo, 3),
                    "Eo_min": round(min(r[3] for r in rows), 3),
                    "Eo_max": round(max(r[3] for r in rows), 3),
                    "theoretical_trays": b - a + 1,
                    "actual_trays": math.ceil((b - a + 1) / eo)}
        return aspen.call(impl)
    except Exception as exc:
        return {"error": str(exc)}


def tool_column_height_estimate(block_name: str, top_space_m: float = 1.0,
                                bottom_space_m: float | None = None,
                                sump_residence_min: float = 5.0,
                                feed_extra_spacing: float = 0.5,
                                config: str | None = None) -> dict:
    """Tangent-to-tangent height from the internals.

    trays per section = ceil(stages / CA_EFF); tray height = (trays - 1) x spacing
    (packing: packed height); + top space + one extra 'feed_extra_spacing' x
    spacing per feed tray + bottom space. If bottom_space_m is None the sump is
    sized for `sump_residence_min` of bottoms liquid plus 0.6 m for the reboiler
    return nozzle and 0.3 m disengagement.
    """
    try:
        def impl():
            tree = aspen._app.RootModel("")
            inp = _walk(tree, "Data", "Blocks", block_name, "Input")
            out = _walk(tree, "Data", "Blocks", block_name, "Output")
            cfg = _active_config(inp, config)
            if not cfg:
                return {"error": f"{block_name} has no Column Internals"}
            n = int(_nstage(tree, block_name))
            feeds = [_val(_walk(inp, "FEED_STAGE", f)) for f in _kids(_walk(inp, "FEED_STAGE"))]
            secs, h_int, dmax, spacing_ref = [], 0.0, 0.0, None
            for sec in _sections(inp, cfg):
                s1 = _val(_walk(inp, "CA_STAGE1", cfg, sec))
                s2 = _val(_walk(inp, "CA_STAGE2", cfg, sec))
                kind = _val(_walk(inp, "CA_INTERNAL", cfg, sec))
                sizing = _val(_walk(inp, "CA_SIZING", cfg, sec)) != "RATING"
                d = ((sizing and (_get_length_m(_walk(out, "CA_DIAM6", cfg, sec))
                                  or _get_length_m(_walk(out, "CA_DIAM2", cfg, sec))))
                     or _get_length_m(_walk(inp, "CA_DIAM", cfg, sec)) or 0.0)
                dmax = max(dmax, d)
                if kind == "PACKING":
                    h = _get_length_m(_walk(out, "CA_HEIGHT2", cfg, sec)) or _get_length_m(_walk(inp, "CA_PACK_HT", cfg, sec)) or 0.0
                    secs.append({"section": sec, "type": "PACKING", "stages": [s1, s2],
                                 "diameter_m": round(d, 3), "height_m": round(h, 3)})
                    h_int += h
                else:
                    ts = _get_length_m(_walk(inp, "CA_TRAY_SPC", cfg, sec)) or 0.6096
                    spacing_ref = spacing_ref or ts
                    eff = _val(_walk(inp, "CA_EFF", cfg, sec)) or 1.0
                    trays = math.ceil((s2 - s1 + 1) / eff - 1e-9)
                    h = trays * ts
                    nfeed = sum(1 for f in feeds if f and s1 <= f <= s2)
                    h += nfeed * feed_extra_spacing * ts
                    secs.append({"section": sec, "type": "TRAY", "stages": [s1, s2],
                                 "efficiency": eff, "actual_trays": trays,
                                 "tray_spacing_m": round(ts, 4), "diameter_m": round(d, 3),
                                 "feed_trays": nfeed, "height_m": round(h, 3)})
                    h_int += h
            if spacing_ref:
                h_int -= spacing_ref  # n trays span (n - 1) spacings
            bottom = bottom_space_m
            if bottom is None:
                # liquid leaving the last tray feeds the sump (kettle reboiler);
                # three-phase RadFrac has no LIQ_FLOW_MS, the hydraulic HYD_LMF has it
                mflow = (_val(_walk(out, "LIQ_FLOW_MS", str(n - 1)))
                         or _val(_walk(out, "HYD_LMF", str(n - 1))))  # kg/hr
                rho = _val(_walk(out, "HYD_RHOL", str(n - 1)))  # gm/cc
                q = 0.0
                if mflow and rho:
                    q = (mflow / (rho * 1000.0)) / 60.0  # m3/min
                area = math.pi * dmax ** 2 / 4 if dmax else 0
                level = (q * sump_residence_min / area) if area else 0.0
                bottom = level + 0.6 + 0.3
            total = top_space_m + h_int + bottom
            return {"block": block_name, "config": cfg, "nstage": n, "sections": secs,
                    "internals_height_m": round(h_int, 3), "top_space_m": top_space_m,
                    "bottom_space_m": round(bottom, 3),
                    "tangent_to_tangent_m": round(total, 2),
                    "max_diameter_m": round(dmax, 3),
                    "L_over_D": round(total / dmax, 1) if dmax else None}
        return aspen.call(impl)
    except Exception as exc:
        return {"error": str(exc)}


def tool_resize_column_stages(block_name: str, nstage: int,
                              feed_stages: dict[str, int] | None = None,
                              update_internals: bool = True) -> str:
    """Change NSTAGE of a RadFrac and keep dependent inputs consistent:
    feed stages (given explicitly), product stages that pointed at the old
    bottom stage, and internals sections that ended at the old last tray."""
    try:
        def impl():
            tree = aspen._app.RootModel("")
            inp = _walk(tree, "Data", "Blocks", block_name, "Input")
            old = int(_nstage(tree, block_name))
            msgs = []
            _walk(inp, "NSTAGE").Value = nstage
            msgs.append(f"NSTAGE {old}->{nstage}")
            for s, st in (feed_stages or {}).items():
                node = _walk(inp, "FEED_STAGE", s)
                if node is None:
                    msgs.append(f"feed '{s}' not found")
                    continue
                node.Value = st
                msgs.append(f"feed {s}->{st}")
            for p in _kids(_walk(inp, "PROD_STAGE")):
                node = _walk(inp, "PROD_STAGE", p)
                if _val(node) == old:
                    node.Value = nstage
                    msgs.append(f"product {p}->{nstage}")
            if update_internals:
                cfg = _active_config(inp, None)
                if cfg:
                    for sec in _sections(inp, cfg):
                        e = _walk(inp, "CA_STAGE2", cfg, sec)
                        if _val(e) is not None and _val(e) >= nstage:
                            e.Value = nstage - 1
                            msgs.append(f"{cfg}/{sec} end->{nstage - 1}")
            return f"{block_name}: " + ", ".join(msgs)
        return aspen.call(impl)
    except Exception as exc:
        return "Error: " + str(exc)
