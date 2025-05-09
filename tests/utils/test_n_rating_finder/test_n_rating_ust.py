import unittest
from src.utils.NRatingFinder import n_rating_ust

class TestUltimateStrengthNRating(unittest.TestCase):
    def __init__(self, methodName = "runTest"):
        super().__init__(methodName)

    def test_valid_rating_data(self):
        res = n_rating_ust(1000)
        self.assertEqual(type(res), tuple)

    def test_invalid_rating_data(self):
        res = n_rating_ust("akb")
        self.assertEqual(res, ("", ""))

    def test_none_rating_data(self):
        res = n_rating_ust(None)
        self.assertEqual(res, ("", ""))