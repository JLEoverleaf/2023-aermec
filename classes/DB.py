# DB.py

from pathlib import Path
import sqlite3

from classes import Translations


class DB:
    def __init__(self, filename: str | Path):
        self.filename: Path = filename if isinstance(filename, Path) else Path(filename)

    def create(self, translations: Translations) -> int:
        if self.filename.is_file():
            self.filename.unlink()
        con = sqlite3.connect(self.filename)
        ret_value = -1
        try:
            cur = con.cursor()
            sql = "CREATE TABLE headers(italian TEXT, english TEXT, cleaned TEXT);"
            cur.execute(sql)
            data = []
            for left, right, cleaned in zip(translations.italian, translations.english, translations.cleaned_english):
                data.append({"left": left, "right": right, "cleaned": cleaned})
            cur.executemany("INSERT INTO headers VALUES (:left, :right, :cleaned)", data)
            con.commit()
            ret_value = 0

        except Exception as ex:
            print(f"Error: {str(ex)}")
        finally:
            con.close()
        return ret_value
