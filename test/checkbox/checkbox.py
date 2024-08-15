import xlwings as xw
import logging

logging.basicConfig(
    level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s"
)


def explore_shape(workbook_path, sheet_name, checkbox_name):
    try:
        with xw.App(visible=False) as app:
            logging.info(f"Opening workbook: {workbook_path}")
            wb = app.books.open(workbook_path)

            logging.info(f"Accessing sheet: {sheet_name}")
            ws = wb.sheets[sheet_name]

            logging.info(f"Searching for checkbox: {checkbox_name}")
            for shape in ws.shapes:
                if shape.name == checkbox_name:
                    logging.info(f"Checkbox '{checkbox_name}' found.")

                    api = shape.api

                    print(f"Shape type: {shape.type}")
                    print(f"Shape api type: {type(api)}")

                    print("\nTrying to access properties:")
                    properties_to_try = [
                        "value",
                        "checked",
                        "state",
                        "enabled",
                        "locked",
                        "visible",
                        "form_control_type",
                        "control_format",
                        "option_button",
                        "properties",
                        "object",
                        "form_control",
                    ]

                    for prop in properties_to_try:
                        try:
                            value = getattr(api, prop)()
                            print(f"{prop}: {value}")
                        except Exception as e:
                            print(f"{prop}: Error - {str(e)}")

                    print("\nTrying to get properties:")
                    try:
                        properties = api.properties.get()
                        for prop in properties:
                            print(f"Property: {prop}")
                    except Exception as e:
                        print(f"Error getting properties: {str(e)}")

                    return

            logging.warning(f"Checkbox '{checkbox_name}' not found")

    except Exception as e:
        logging.error(f"An error occurred: {str(e)}")


# Usage
workbook_path = "/Users/remillieux/Documents/Proneo/logiciel/data/Plan et Rapport d'audit certification V32.xlsm"
sheet_name = "OPAC"
checkbox_name = "Check Box 59"

explore_shape(workbook_path, sheet_name, checkbox_name)
