import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from typing import List, Tuple, Dict
from PIL import Image

def generate_report(data_list: List[Tuple[Dict[str, Tuple[str, str, float]], Dict[str, float], float, str, str]]) -> None:
    """
    Generate a report in Excel with the given data.
    data_list: List of tuples (assignments, load_zones, execution_time, instance_name, method_name)
    """
    report_data = []
    image_paths = []

    for assignments, load_zones, execution_time, instance_name, method_name in data_list:
        Wmax = max(load_zones.values())
        Wmin = min(load_zones.values())
        Wmax_Wmin = Wmax - Wmin
        total_W = sum(load_zones.values())
        average_W = total_W / len(load_zones)

        report_data.append({
            'Instance': instance_name,
            'Method': method_name,
            'Wmax': round(Wmax, 2),
            'Wmin': round(Wmin, 2),
            'Wmax-Wmin': round(Wmax_Wmin, 2),
            'Total W': round(total_W, 2),
            'Average W': round(average_W, 2),
            'Execution Time': execution_time
        })

        # Generate bar chart for load_zones
        plt.figure(figsize=(10, 6))
        colors = plt.cm.cividis(np.linspace(0, 1, len(load_zones)))
        plt.bar(load_zones.keys(), load_zones.values(), color=colors)
        plt.xlabel('Zones')
        plt.ylabel('Time [s]')
        plt.title(f'{instance_name} + {method_name}')
        plt.tight_layout()

        # Save the plot as an image
        output_dir = 'constructive-method/reports/bar_images'
        os.makedirs(output_dir, exist_ok=True)
        image_path = f'{output_dir}/{instance_name}_{method_name}.png'
        plt.savefig(image_path)
        plt.close()
        image_paths.append(image_path)

    df = pd.DataFrame(report_data)
    output_dir = 'constructive-method/reports'
    os.makedirs(output_dir, exist_ok=True)
    output_file = f'{output_dir}/report.xlsx'

    # Save the DataFrame to an Excel file
    df.to_excel(output_file, index=False)

    # Adjust column widths
    workbook = load_workbook(output_file)
    worksheet = workbook.active

    for column in worksheet.columns:
        max_length = 0
        column = list(column)
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 3) * 1.2
        worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

    workbook.save(output_file)

    # Create a collage with 2 columns per row
    images = [Image.open(image_path) for image_path in image_paths]
    num_images = len(images)
    columns = 2
    rows = int(np.ceil(num_images / columns))
    max_width = max(img.width for img in images)
    max_height = max(img.height for img in images)

    collage_width = columns * max_width
    collage_height = rows * max_height

    collage = Image.new('RGB', (collage_width, collage_height), (255, 255, 255))

    x_offset = 0
    y_offset = 0
    for i, img in enumerate(images):
        collage.paste(img, (x_offset, y_offset))
        x_offset += max_width
        if (i + 1) % columns == 0:
            x_offset = 0
            y_offset += max_height

    collage.save(f'{output_dir}/comparison-workload-balancing-bar-chart.png')