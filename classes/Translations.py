# Translations.py
# Header translations were provided in a file from aermec.
# This class encapsulates those translations.

from .Utility import XLSXDataSource
from pathlib import Path
from io import StringIO
import re


class Translations:
    def __init__(self, source: XLSXDataSource):
        self._source = source
        self.italian = []
        self.english = []
        self.cleaned_english = []
        self._populate()

    def _clean(self, inp) -> str:
        oup = re.sub(r"[\s\-\:\,\.\%\(\)\=\/]", "_", inp)
        return oup

    def _populate(self):
        i = 0
        while True:
            left = self._source.cell
            right = self._source.cell_right
            if left == None or right == None or not left or not right:
                break
            self.italian.append(left)
            self.english.append(right)
            self.cleaned_english.append(self._clean(right))
            self._source.move_down_one()

    @classmethod
    def from_file(cls, filename: Path | str):
        return cls(XLSXDataSource.from_file(filename))

    def __str__(self):
        sio = StringIO()
        for left, right, cleaned in zip(self.italian, self.english, self.cleaned_english):
            sio.write(f"{left} : {right} ({cleaned})\n")
        sio.write(f"\nFound {len(self.italian)} rows\n")
        return sio.getvalue()
