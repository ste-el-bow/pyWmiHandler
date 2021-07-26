from unittest import TestCase
import pyWmiHandler


class TestHelpers(TestCase):
    def test_helper_to_human(self):
        five_tb = pyWmiHandler.helpers.convert_to_human_readable(5 * 1024 * 1024 * 1024 * 1024, False)
        self.assertEqual(five_tb, '5 TB')
