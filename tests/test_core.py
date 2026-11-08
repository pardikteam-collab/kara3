import unittest
from core.automation import build_employee_plan
from core.mapping import DEFAULT_MAPPING

class ShiftPlanTests(unittest.TestCase):
    def employee(self, values):
        return {"days": {i + 1: v for i, v in enumerate(values)}}

    def test_mapping_matches_user_rules(self):
        emp = self.employee(["ص", "ع", "م", "ن", ""])
        plan, errors = build_employee_plan(emp, DEFAULT_MAPPING)
        self.assertEqual(errors, [])
        self.assertEqual(plan, {1:"M921", 2:"M1400", 3:"SHIFT", 4:"OFF", 5:"OFF"})

    def test_unknown_codes_block_employee(self):
        for code in ("ض", "غ", "نامشخص"):
            with self.subTest(code=code):
                plan, errors = build_employee_plan(self.employee([code]), DEFAULT_MAPPING)
                self.assertEqual(plan, {})
                self.assertEqual(len(errors), 1)

    def test_leading_day_keys_are_ints(self):
        plan, errors = build_employee_plan({"days": {"01":"ص", "02":"ن"}}, DEFAULT_MAPPING)
        self.assertEqual(errors, [])
        self.assertEqual(plan, {1:"M921", 2:"OFF"})

if __name__ == "__main__":
    unittest.main()
