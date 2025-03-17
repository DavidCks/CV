# auto generated using `datamodel-codegen --input cv.json --output prg/__generated__/cv_model.py`
from .__generated__.cv_model import Model
import json
from pathlib import Path


def loadCV(json_path: str) -> Model:
    """
    Reads a JSON file and returns a Model instance.

    :param json_path: Path to the JSON file.
    :return: Instance of Model.
    """
    json_file = Path(json_path)

    if not json_file.exists():
        raise FileNotFoundError(f"File not found: {json_path}")

    with json_file.open("r", encoding="utf-8") as f:
        data = json.load(f)

    try:
        return Model(**data)
    except Exception as e:
        raise ValueError(f"Error loading JSON into Model: {e}")


# Example usage:
# cv_model = load_cv_model("cv.json")
# print(cv_model.Profile.Name)
