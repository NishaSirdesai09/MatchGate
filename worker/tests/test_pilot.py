import sys
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from pilot import (  # noqa: E402
    HEADERS,
    changed,
    read_sample,
    reunion_in_range,
    state_in_region,
    state_is_ny,
)


class PilotTests(unittest.TestCase):
    def test_reads_ny_sample_without_network(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "sample.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            headers = list(HEADERS.values())
            sheet.append(headers)
            sheet.append(
                [
                    "Ada",
                    "Lovelace",
                    "Ada Lovelace",
                    "New York",
                    "NY",
                    "United States",
                    "Director",
                    "Example Co",
                    "MBA",
                    2000,
                ]
            )
            sheet.append(
                [
                    "Grace",
                    "Hopper",
                    "Grace Hopper",
                    "New York",
                    "New York",
                    "United States",
                    "VP",
                    "Example Co",
                    "MS",
                    2001,
                ]
            )
            workbook.save(path)

            read_headers, people, mismatches = read_sample(path, None, 100)

            self.assertEqual(headers, read_headers)
            self.assertEqual(2, len(people))
            self.assertEqual([], mismatches)

    def test_state_normalization(self):
        self.assertTrue(state_is_ny("NY"))
        self.assertTrue(state_is_ny(" New York "))
        self.assertFalse(state_is_ny("NJ"))

    def test_change_comparison(self):
        self.assertEqual("NO", changed("Example Co", " example   co "))
        self.assertEqual("YES", changed("Old Co", "New Co"))
        self.assertEqual("", changed("Old Co", ""))

    def test_region_and_reunion_filters(self):
        self.assertTrue(state_in_region("NJ", "tri-state"))
        self.assertFalse(state_in_region("CA", "tri-state"))
        self.assertTrue(reunion_in_range(1990, 1980, 1999))
        self.assertFalse(reunion_in_range(2005, 1980, 1999))


if __name__ == "__main__":
    unittest.main()
