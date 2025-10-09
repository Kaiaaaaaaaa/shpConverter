# -*- coding: utf-8 -*-
"""
SHP -> DXF converter
- Converts POINT, POLYLINE, and POLYGON to DXF entities
- Reads attributes to set entity layer and color:
  Preferred: integer fields R,G,B (0..255)
  Fallbacks: 'color' as '#RRGGBB' or 'r,g,b' or 'R,G,B' string; 'aci' as AutoCAD Color Index
- Writes DXF true_color (exact RGB). If a 'layer' attribute exists it is applied to the entity.
"""
import os
import shapefile
import ezdxf
from ezdxf import colors as ezcolors


def _split_parts(shape):
    """
    Yield parts for polyline/polygon shapes.
    Each part is a list of (x, y) tuples.
    """
    pts = shape.points
    parts = getattr(shape, "parts", [0])
    # parts is a list of starting indices; add end sentinel
    idxs = list(parts) + [len(pts)]
    for i in range(len(parts)):
        a, b = idxs[i], idxs[i + 1]
        yield pts[a:b]


def _parse_rgb_from_record(rec, field_names_lower):
    """
    Extract (R,G,B) from a record using common GIS conventions.
    Priority:
        1) 'RGB_text' column as 'rgb(R,G,B)'
        2) Integer fields R,G,B
        3) 'color' hex '#RRGGBB' or 'RRGGBB'
        4) 'color' as 'r,g,b' (any case, with spaces)
        5) 'aci' (AutoCAD Color Index)
    Returns (R,G,B) ints in 0..255; defaults to (0,0,0) if not found/invalid.
    """
    name_to_idx = {n: i for i, n in enumerate(field_names_lower)}
    # 1) RGB_text column
    rgb_text_idx = name_to_idx.get("rgb_text")
    if rgb_text_idx is not None:
        val = rec[rgb_text_idx]
        if isinstance(val, str) and val.lower().startswith("rgb(") and val.endswith(")"):
            try:
                parts = val[4:-1].split(",")
                if len(parts) == 3:
                    R, G, B = int(parts[0]), int(parts[1]), int(parts[2])
                    if 0 <= R <= 255 and 0 <= G <= 255 and 0 <= B <= 255:
                        return R, G, B
            except Exception:
                pass

    # 2) R,G,B integer columns
    def get_val(*candidates):
        for c in candidates:
            idx = name_to_idx.get(c.lower())
            if idx is not None:
                return rec[idx]
        return None

    r = get_val("r")
    g = get_val("g")
    b = get_val("b")
    try:
        if r is not None and g is not None and b is not None:
            R = int(r)
            G = int(g)
            B = int(b)
            if 0 <= R <= 255 and 0 <= G <= 255 and 0 <= B <= 255:
                return R, G, B
    except Exception:
        pass

    # 3/4) color string field
    color_val = get_val("color", "colour", "clr")
    if isinstance(color_val, str):
        s = color_val.strip()
        # Hex forms
        if s.startswith("#"):
            s = s[1:]
        if len(s) == 6 and all(ch in "0123456789aAbBcCdDeEfF" for ch in s):
            try:
                R = int(s[0:2], 16)
                G = int(s[2:4], 16)
                B = int(s[4:6], 16)
                return R, G, B
            except Exception:
                pass
        # CSV 'r,g,b'
        parts = [p.strip() for p in s.replace("(", "").replace(")", "").split(",")]
        if len(parts) == 3:
            try:
                R, G, B = int(parts[0]), int(parts[1]), int(parts[2])
                if 0 <= R <= 255 and 0 <= G <= 255 and 0 <= B <= 255:
                    return R, G, B
            except Exception:
                pass

    # 5) ACI -> RGB
    aci_val = get_val("aci", "autocadcolorindex")
    try:
        if aci_val is not None:
            aci = int(aci_val)
            if aci not in (0, 256):
                R, G, B = ezcolors.aci2rgb(aci)
                return int(R), int(G), int(B)
    except Exception:
        pass

    # Default
    return 0, 0, 0


def convert_shp_to_dxf(input_dir, output_dir):
    """
    Converts all .shp files in input_dir to .dxf format and saves them in output_dir.
    - Supports POINT, POLYLINE, POLYGON
    - Applies color and layer from attributes when available
    """

    # Check input directory
    if not os.path.isdir(input_dir):
        print(f"Input directory '{input_dir}' does not exist.")
        return

    # Ensure output directory
    if not os.path.isdir(output_dir):
        os.makedirs(output_dir)
        print(f"Created output directory '{output_dir}'.")

    for filename in os.listdir(input_dir):
        if not filename.lower().endswith(".shp"):
            continue

        shp_path = os.path.join(input_dir, filename)
        dxf_filename = os.path.splitext(filename)[0] + ".dxf"
        dxf_path = os.path.join(output_dir, dxf_filename)

        print(f"Converting '{shp_path}' to '{dxf_path}'...")

        try:
            sf = shapefile.Reader(shp_path)
            # Fields: list of tuples; first item is DeletionFlag placeholder
            fields = [f[0] for f in sf.fields[1:]]
            fields_lower = [f.lower() for f in fields]

            # New DXF
            doc = ezdxf.new(dxfversion="R2010")
            msp = doc.modelspace()

            # Iterate shape+record pairs to keep attributes aligned
            for sr in sf.iterShapeRecords():
                shape = sr.shape
                rec = sr.record  # behaves like a list-like sequence (values ordered as fields)
                # Layer from attribute if present
                layer_name = None
                try:
                    if "layer" in fields_lower:
                        layer_name = rec[fields_lower.index("layer")]
                        if isinstance(layer_name, bytes):
                            layer_name = layer_name.decode("utf-8", errors="ignore")
                        if not isinstance(layer_name, str):
                            layer_name = str(layer_name)
                        layer_name = layer_name[:255]
                        if layer_name not in doc.layers:
                            doc.layers.add(layer_name)
                except Exception:
                    layer_name = None

                # Color as true_color from attributes
                R, G, B = _parse_rgb_from_record(rec, fields_lower)
                true_col = ezcolors.rgb2int((int(R), int(G), int(B)))

                dxfattribs = {"true_color": true_col}
                if layer_name:
                    dxfattribs["layer"] = layer_name

                geom_type = shape.shapeType

                if geom_type == shapefile.POINT:
                    if shape.points:
                        msp.add_point(shape.points[0], dxfattribs=dxfattribs)

                elif geom_type in (shapefile.POLYLINE, shapefile.POLYGON):
                    is_closed = (geom_type == shapefile.POLYGON)
                    # Handle multipart shapes (rings/parts)
                    for part in _split_parts(shape):
                        if not part:
                            continue
                        # Ensure closure for polygons
                        pts = list(part)
                        if is_closed and pts[0] != pts[-1]:
                            pts = pts + [pts[0]]
                        msp.add_lwpolyline(pts, close=is_closed, dxfattribs=dxfattribs)

                else:
                    print(
                        f"Unsupported geometry type ({geom_type}) in '{filename}'. Skipping shape."
                    )

            # Save DXF
            doc.saveas(dxf_path)
            print(f"Saved '{dxf_path}'.\n")

        except Exception as e:
            print(f"Failed to convert '{filename}'. Error: {e}\n")

    print("All conversions completed.")


if __name__ == "__main__":
    # Define input and output directories
    input_directory = os.path.join(".", "Files2Convert")
    output_directory = os.path.join(".", "ConvertedSHP2DXF")
    if not os.path.isdir(output_directory):
        os.makedirs(output_directory)
        print(f"Created output directory: {output_directory}")
    convert_shp_to_dxf(input_directory, output_directory)
