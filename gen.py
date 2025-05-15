# Make sure to run `datamodel-codegen --input cv.json --output prg/__generated__/cv_model.py`
# when cv.json changes and to update the selectors if needed.

from _cv import runCV
from _resume import runResume

import argparse

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Generate content in a specific language."
    )
    parser.add_argument(
        "lang", choices=["ja", "en", "de"], help="Language code (ja, en, de)"
    )

    args = parser.parse_args()

    print(f"Running with language: {args.lang}")
    runResume(args.lang)
    runCV(args.lang)
