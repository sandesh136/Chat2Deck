from pptx import Presentation
from pptx.util import Pt
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR_TYPE
from pptx.util import Inches
from pptx.dml.color import RGBColor
from io import BytesIO
import os
import uuid
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def create_pptx_from_code(python_code: str) -> BytesIO:
    decks_dir = "decks"
    try:
        os.makedirs(decks_dir, exist_ok=True)

        filename = f"generated_presentation_{uuid.uuid4()}.pptx"
        filepath = os.path.abspath(os.path.join(decks_dir, filename))

        code_with_new_path = python_code.replace("'generated_presentation.pptx'", f"'{filepath}'").replace('"generated_presentation.pptx"', f'"{filepath}"')

        logging.info(f"Executing AI-generated code to create {filepath}")

        # Define a safe but comprehensive global scope for exec
        exec_globals = {
            "Presentation": Presentation, "Pt": Pt, "Inches": Inches,
            "CategoryChartData": CategoryChartData, "XL_CHART_TYPE": XL_CHART_TYPE,
            "XL_LEGEND_POSITION": XL_LEGEND_POSITION, "MSO_SHAPE": MSO_SHAPE,
            "MSO_CONNECTOR_TYPE": MSO_CONNECTOR_TYPE, "RGBColor": RGBColor
        }

        exec(code_with_new_path, exec_globals)

        if not os.path.exists(filepath):
            raise FileNotFoundError(f"AI-generated code did not create the expected presentation file at {filepath}.")

        with open(filepath, "rb") as f:
            buffer = BytesIO(f.read())

        logging.info("Successfully created and read presentation file.")
        return buffer

    except Exception as e:
        # Re-raise the exception to be caught by the retry loop in the interaction service
        raise e
