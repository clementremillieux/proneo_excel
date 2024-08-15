import subprocess
import sys


def check_checkbox(checkbox_name):
    script_path = "/Users/remillieux/Documents/Proneo/logiciel/test/appel_script/test_checkbox.scpt"
    try:
        result = subprocess.run(['osascript', script_path, checkbox_name],
                                capture_output=True,
                                text=True,
                                check=True)
        return result.stdout.strip()
    except subprocess.CalledProcessError as e:
        return f"Erreur lors de l'exécution du script : {e.stderr.strip()}"


# Exemple d'utilisation
if __name__ == "__main__":
    checkboxes_to_check = [
        "Check Box 58", "Check Box 59", "Check Box 60", "Check Box 61",
        "Check Box 80"
    ]

    for checkbox in checkboxes_to_check:
        state = check_checkbox(checkbox)
        print(state)
