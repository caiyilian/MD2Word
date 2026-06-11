from __future__ import annotations
import os
import sys


class Win32PageCounter:
    def __init__(self):
        self._word = None

    def count(self, docx_path: str) -> int:
        abs_path = os.path.abspath(docx_path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"Document not found: {abs_path}")

        try:
            from win32com.client import Dispatch
        except ImportError:
            raise RuntimeError(
                "pywin32 is required for page counting. "
                "Install it with: pip install pywin32"
            )

        try:
            word = Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            doc = word.Documents.Open(abs_path)
            pages = doc.ComputeStatistics(2)  # wdStatisticPages = 2
            doc.Close(False)
            return pages
        except Exception as e:
            raise RuntimeError(f"Failed to count pages via Word: {e}") from e
        finally:
            try:
                word.Quit()
            except Exception:
                pass

    @staticmethod
    def available() -> bool:
        try:
            from win32com.client import Dispatch
            word = Dispatch("Word.Application")
            word.Quit()
            return True
        except Exception:
            return False
