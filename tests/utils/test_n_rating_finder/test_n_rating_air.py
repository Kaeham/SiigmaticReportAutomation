import unittest
from src.utils.NRatingFinder import n_rating_air_results

class TestAirNRating(unittest.TestCase):
    def __init__(self, methodName = "runTest"):
        super().__init__(methodName)

    def test_valid_rating_data(self):
        res = n_rating_air_results(2, 2)
        self.assertNotEqual(res, ("", ""))

    def test_invalid_rating_data(self):
        res = n_rating_air_results("akb", "nmb")
        self.assertEqual(res, ("", ""))

    def test_none_rating_data(self):
        res = n_rating_air_results(None, None)
        self.assertEqual(res, ("", ""))