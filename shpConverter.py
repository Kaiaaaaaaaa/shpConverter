import os
import shapefile
import ezdxf

def convert_shp_to_dxf(input_dir, output_dir):
    """
    Converts all .shp files in the input_dir to .dxf format and saves them in output_dir.
    Only Points and Polylines are converted.
    
    :param input_dir: Path to the directory containing .shp files
    :param output_dir: Path to the directory where .dxf files will be saved
    """
    # Check if input directory exists
    if not os.path.isdir(input_dir):
        print(f"Input directory '{input_dir}' does not exist.")
        return

    # Create output directory if it doesn't exist
    if not os.path.isdir(output_dir):
        os.makedirs(output_dir)
        print(f"Created output directory '{output_dir}'.")

    # Iterate through all files in the input directory
    for filename in os.listdir(input_dir):
        if filename.lower().endswith('.shp'):
            shp_path = os.path.join(input_dir, filename)
            dxf_filename = os.path.splitext(filename)[0] + '.dxf'
            dxf_path = os.path.join(output_dir, dxf_filename)

            print(f"Converting '{shp_path}' to '{dxf_path}'...")

            try:
                # Read Shapefile
                sf = shapefile.Reader(shp_path)
                shapes = sf.shapes()
                shape_types = sf.shapeType

                # Create DXF document
                doc = ezdxf.new(dxfversion='R2010')
                msp = doc.modelspace()

                for shape in shapes:
                    geom_type = shape.shapeType
                    points = shape.points

                    if geom_type == shapefile.POINT:
                        # Add point to DXF
                        msp.add_point(points[0])
                    elif geom_type in [shapefile.POLYLINE, shapefile.POLYGON]:
                        # Determine if shape is closed
                        is_closed = (geom_type == shapefile.POLYGON)
                        # Add polyline to DXF
                        msp.add_lwpolyline(points, close=is_closed)
                    else:
                        print(f"Unsupported geometry type ({geom_type}) in '{filename}'. Skipping shape.")
                
                # Save DXF
                doc.saveas(dxf_path)
                print(f"Saved '{dxf_path}'.\n")

            except Exception as e:
                print(f"Failed to convert '{filename}'. Error: {e}\n")

    print("All conversions completed.")

if __name__ == "__main__":
    # Define input and output directories
    input_directory = os.path.join('.', 'Files')       # .\ConvertFiles
    output_directory = os.path.join('.', 'Files')      # .\ConvertedDXF

    convert_shp_to_dxf(input_directory, output_directory)
