from pptx import Presentation
from pptx.util import Pt
from io import BytesIO
import os
import uuid
import logging
import traceback

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def create_pptx_from_code(python_code: str) -> BytesIO:
    decks_dir = "decks"
    try:
        os.makedirs(decks_dir, exist_ok=True)
        
        # Create a unique filename and join with decks_dir
        filename = f"generated_presentation_{uuid.uuid4()}.pptx"
        relative_filepath = os.path.join(decks_dir, filename)
        
        # Convert to absolute path to ensure correct file location
        filepath = os.path.abspath(relative_filepath)
        
        # Replace original filename string with absolute filepath
        code_with_new_path = (
            python_code.replace("'generated_presentation.pptx'", f"'{filepath}'")
                       .replace('"generated_presentation.pptx"', f'"{filepath}"')
        )

        logging.info(f"Executing AI-generated code to create {filepath}")
        logging.debug(f"Code to exec:\n{code_with_new_path}")

        exec_globals = {"__builtins__": __builtins__, "Presentation": Presentation, "Pt": Pt}

        try:
            compiled_code = compile(code_with_new_path, "<string>", "exec")  # syntax check
            exec(compiled_code, exec_globals)
        except Exception as exec_err:
            logging.error("Execution inside AI-generated code failed:")
            logging.error(traceback.format_exc())
            print("AI-generated code execution error:", exec_err)
            print("Failing code snippet:\n", code_with_new_path)
            raise

        if not os.path.exists(filepath):
            raise FileNotFoundError(f"AI-generated code did not create the expected presentation file at {filepath}.")

        with open(filepath, "rb") as f:
            buffer = BytesIO(f.read())
        
        logging.info("Successfully created and read presentation file.")
        return buffer

    except Exception as e:
        logging.error(f"Failed to create presentation from code. Error: {e}")
        logging.error(f"Failing code:\n---\n{code_with_new_path}\n---")
        raise RuntimeError(f"An error occurred while generating the presentation: {e}")
