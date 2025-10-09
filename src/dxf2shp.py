#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
DXF -> SHP converter
- Splits output into three shapefiles per input: *_points.shp, *_lines.shp, *_polygons.shp
- Attributes written to .dbf: layer, etype, handle, R, G, B
- CPG file is written with UTF-8 if missing for correct attribute encoding
"""
import os
import shapefile  # pyshp
import ezdxf
from ezdxf import colors as ezcolors

INPUT_DIR = os.path.join(".", "Files2Convert")
OUTPUT_DIR = os.path.join(".", "ConvertedDXF2SHP")
DBF_CPG = "UTF-8"


def ensure_dir(path: str) -> None:
    if not os.path.isdir(path):
        os.makedirs(path)


def write_cpg_if_missing(shp_base_path: str) -> None:
    cpg = shp_base_path + ".cpg"
    if not os.path.exists(cpg):
        with open(cpg, "w", encoding="ascii") as f:
            f.write(DBF_CPG)


def get_point_xy(e):
    loc = e.dxf.location
    return float(loc.x), float(loc.y)


def get_line_xy(e):
    s = e.dxf.start
    e2 = e.dxf.end
    return (float(s.x), float(s.y)), (float(e2.x), float(e2.y))


def lwpolyline_xy(e):
    pts = []
    gp = getattr(e, "get_points", None)
    if callable(gp):
        try:
            for x, y in gp("xy"):
                pts.append((float(x), float(y)))
            return pts
        except TypeError:
            for p in gp():
                pts.append((float(p[0]), float(p[1])))
            return pts
    if isinstance(gp, (list, tuple)):
        for p in gp:
            pts.append((float(p[0]), float(p[1])))
        return pts
    try:
        for p in e:
            pts.append((float(p[0]), float(p[1])))
        return pts
    except Exception:
        pass
    raise ValueError("Unsupported LWPOLYLINE point format for this ezdxf version.")


def polyline_xy(e):
    pts = []
    verts_attr = getattr(e, "vertices", None)
    if callable(verts_attr):
        it = verts_attr()
    elif isinstance(verts_attr, (list, tuple)):
        it = verts_attr
    else:
        it = []
    for v in it:
        d = getattr(v, "dxf", v)
        loc = getattr(d, "location", None)
        if loc is not None:
            pts.append((float(getattr(loc, "x", 0.0)), float(getattr(loc, "y", 0.0))))
        else:
            x = float(getattr(d, "x", 0.0))
            y = float(getattr(d, "y", 0.0))
            pts.append((x, y))
    return pts


def is_closed(e) -> bool:
    v = getattr(e, "closed", None)
    if isinstance(v, bool):
        return v
    v = getattr(e, "is_closed", None)
    if isinstance(v, bool):
        return v
    return False


def close_if_needed(pts):
    return pts + [pts[0]] if pts and pts[0] != pts[-1] else pts


class WriterSet:
    def __init__(self, base_out: str):
        self.base = base_out
        self.points = None
        self.lines = None
        self.polygons = None
        self.counts = {"points": 0, "lines": 0, "polygons": 0}

    def _init(self, kind: str):
        path = f"{self.base}_{kind}"
        if kind == "points":
            w = shapefile.Writer(path, shapeType=shapefile.POINT)
        elif kind == "lines":
            w = shapefile.Writer(path, shapeType=shapefile.POLYLINE)
        elif kind == "polygons":
            w = shapefile.Writer(path, shapeType=shapefile.POLYGON)
        else:
            raise ValueError(kind)
        w.autoBalance = 1
        # Attributes
        w.field("layer", "C", size=64)
        w.field("etype", "C", size=16)
        # Removed handle field
        # GIS-common color storage
        w.field("RGB_text", "C", size=16)
        setattr(self, kind, w)
        return w

    def get(self, kind: str):
        w = getattr(self, kind)
        return w if w is not None else self._init(kind)

    def close_all(self):
        for kind in ("points", "lines", "polygons"):
            w = getattr(self, kind)
            if w is not None:
                try:
                    w.close()
                finally:
                    write_cpg_if_missing(f"{self.base}_{kind}")


def convert_one_dxf(dxf_path: str, out_dir: str) -> None:
    name = os.path.splitext(os.path.basename(dxf_path))[0]
    base_out = os.path.join(out_dir, name)
    print(f"Converting '{dxf_path}' -> '{base_out}_*.shp' ...")

    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()
    entities = msp.query("POINT LINE LWPOLYLINE POLYLINE")

    writers = WriterSet(base_out)

    def entity_rgb(ent) -> tuple:
        """Get RGB for an entity: prefer true_color, then entity ACI, then layer ACI."""
        d = ent.dxf
        # 1) true_color if present
        if d.hasattr("true_color"):
            r, g, b = ezcolors.int2rgb(d.true_color)
            return int(r), int(g), int(b)
        # 2) explicit entity ACI (1..255), 256=ByLayer, 0=ByBlock
        aci = int(getattr(d, "color", 256) or 256)
        if aci not in (0, 256):
            r, g, b = ezcolors.aci2rgb(aci)
            return int(r), int(g), int(b)
        # 3) layer color
        lay = getattr(d, "layer", None) or "0"
        try:
            layer_obj = doc.layers.get(lay)
            r, g, b = ezcolors.aci2rgb(int(layer_obj.color or 7))
            return int(r), int(g), int(b)
        except Exception:
            # Fallback black
            return 0, 0, 0

    for e in entities:
        et = e.dxftype()
        layer = getattr(e.dxf, "layer", "")
        R, G, B = entity_rgb(e)
        rgb_text = f"rgb({R},{G},{B})"

        if et == "POINT":
            x, y = get_point_xy(e)
            w = writers.get("points")
            w.point(x, y)
            w.record(layer, et, rgb_text)
            writers.counts["points"] += 1

        elif et == "LINE":
            (x1, y1), (x2, y2) = get_line_xy(e)
            w = writers.get("lines")
            w.line([[(x1, y1), (x2, y2)]])
            w.record(layer, et, rgb_text)
            writers.counts["lines"] += 1

        elif et == "LWPOLYLINE":
            pts = lwpolyline_xy(e)
            if is_closed(e) and len(pts) >= 3:
                w = writers.get("polygons")
                w.poly([close_if_needed(pts)])
                w.record(layer, et, rgb_text)
                writers.counts["polygons"] += 1
            else:
                w = writers.get("lines")
                w.line([pts])
                w.record(layer, et, rgb_text)
                writers.counts["lines"] += 1

        elif et == "POLYLINE":
            pts = polyline_xy(e)
            if is_closed(e) and len(pts) >= 3:
                w = writers.get("polygons")
                w.poly([close_if_needed(pts)])
                w.record(layer, et, rgb_text)
                writers.counts["polygons"] += 1
            else:
                w = writers.get("lines")
                w.line([pts])
                w.record(layer, et, rgb_text)
                writers.counts["lines"] += 1

    writers.close_all()
    print(
        f"  Done: {writers.counts['points']} points, "
        f"{writers.counts['lines']} lines, "
        f"{writers.counts['polygons']} polygons\n"
    )


def convert_all():
    if not os.path.isdir(INPUT_DIR):
        print(f"Input directory '{INPUT_DIR}' does not exist.")
        return
    ensure_dir(OUTPUT_DIR)
    if not os.path.isdir(OUTPUT_DIR):
        print(f"Failed to create output directory '{OUTPUT_DIR}'.")
        return
    print(f"Output directory: {OUTPUT_DIR}")
    for fn in os.listdir(INPUT_DIR):
        if fn.lower().endswith(".dxf"):
            dxf_path = os.path.join(INPUT_DIR, fn)
            try:
                convert_one_dxf(dxf_path, OUTPUT_DIR)
            except Exception as e:
                import traceback
                print(f"Failed to convert '{fn}'. Error: {e}")
                traceback.print_exc()
                print()
    print("All conversions completed.")


if __name__ == "__main__":
    convert_all()
