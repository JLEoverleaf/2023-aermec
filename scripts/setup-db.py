#!/usr/bin/env python3

# setup-dp.py
from pathlib import Path

project_dir = Path(__file__).parent.parent
import sys

sys.path.append(str(project_dir))
from dotenv import dotenv_values
from classes import DB, Translations

config = dotenv_values(project_dir / ".env")


def get_translations(filename: Path | str) -> Translations:
    return Translations.from_file(filename)


def create_db(filename: Path | str, translations: Translations) -> DB:
    db = DB(filename)
    if db.create(translations):
        sys.exit(1)
    return db


def main():
    translations = get_translations(Path(config["DATA_LOCATION"]) / "translation.xlsx")
    db = create_db(Path(config["DB_NAME"]), translations)
    sys.exit(0)


if __name__ == "__main__":
    main()
