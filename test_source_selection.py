import os
import tempfile
import unittest
from copy import deepcopy

from wps_source_selection import plan_selected_source_groups, validate_source_selection


class SourceSelectionTests(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp()
        self.a = os.path.join(self.root, "a.xlsx")
        self.b = os.path.join(self.root, "b.xlsx")
        self.c = os.path.join(self.root, "c.xlsx")

    def rule(self, **kwargs):
        value = {"scope": "source", "source_paths": [self.a]}
        value.update(kwargs)
        return value

    def test_all_five_operations_group_and_preserve_first_group_order(self):
        settings = {}
        for key in ("CustomFormula", "TextToColumns", "RegionFilter", "Deduplicate", "FormatConvert"):
            settings[key] = {"enabled": True, "rules": [self.rule(marker=key)]}
        groups = plan_selected_source_groups(settings)
        self.assertEqual(len(groups), 1)
        self.assertEqual(groups[0]["source_paths"], [os.path.abspath(self.a)])
        self.assertEqual(groups[0]["output_filename"], "a_已处理.xlsx")
        self.assertEqual(list(groups[0]["rules"]), [
            "CustomFormula", "TextToColumns", "RegionFilter", "Deduplicate", "FormatConvert"
        ])
        self.assertEqual([groups[0]["rules"][key][0]["marker"] for key in groups[0]["rules"]], list(groups[0]["rules"]))

    def test_single_and_multi_source_selection_and_distinct_order(self):
        settings = {
            "CustomFormula": {"enabled": True, "rules": [self.rule(source_paths=[self.a, self.b])]},
            "TextToColumns": {"enabled": True, "rules": [self.rule(source_paths=[self.b, self.a])]},
        }
        groups = plan_selected_source_groups(settings)
        self.assertEqual(len(groups), 2)
        self.assertEqual(groups[0]["source_paths"], [os.path.abspath(self.a), os.path.abspath(self.b)])
        self.assertEqual(groups[1]["source_paths"], [os.path.abspath(self.b), os.path.abspath(self.a)])

    def test_output_name_normalization_and_explicit_name_grouping(self):
        settings = {
            "CustomFormula": {"enabled": True, "rules": [self.rule(source_output_filename="daily")]},
            "RegionFilter": {"enabled": True, "rules": [self.rule(source_output_filename="daily.xlsx")]},
        }
        groups = plan_selected_source_groups(settings)
        self.assertEqual(len(groups), 1)
        self.assertEqual(groups[0]["output_filename"], "daily.xlsx")

    def test_disabled_or_target_rules_are_not_selected(self):
        settings = {
            "CustomFormula": {"enabled": False, "rules": [self.rule()]},
            "TextToColumns": {"enabled": True, "rules": [self.rule(scope="target", source_paths=[])]},
        }
        self.assertEqual(plan_selected_source_groups(settings), [])

    def test_legacy_missing_and_empty_fields_are_compatible(self):
        settings = {
            "CustomFormula": {"enabled": True, "rules": [{"scope": "source"}]},
            "TextToColumns": {"enabled": True, "rules": [self.rule(source_paths=[])]},
        }
        validate_source_selection({"scope": "target"})
        validate_source_selection({"scope": "source", "source_paths": []})
        self.assertEqual(plan_selected_source_groups(settings), [])

    def test_deep_copy_of_grouped_rules(self):
        original = self.rule(nested={"values": [1]})
        settings = {"Deduplicate": {"enabled": True, "rules": [original]}}
        groups = plan_selected_source_groups(settings)
        original["nested"]["values"].append(2)
        self.assertEqual(groups[0]["rules"]["Deduplicate"][0]["nested"], {"values": [1]})
        self.assertIsNot(groups[0]["rules"]["Deduplicate"][0], original)

    def test_invalid_source_paths_and_duplicates(self):
        bad_values = [None, "x.xlsx", (self.a,), [self.a, "  "], [self.a, self.a], [self.a, self.a.replace("a.xlsx", "." + os.sep + "a.xlsx")]]
        for paths in bad_values:
            with self.subTest(paths=paths):
                with self.assertRaises(ValueError):
                    validate_source_selection(self.rule(source_paths=paths))
        with self.assertRaises(ValueError):
            validate_source_selection(self.rule(source_paths=[self.a, "bad.csv"]))
        with self.assertRaises(ValueError):
            validate_source_selection({"scope": "target", "source_paths": [self.a]})
        with self.assertRaises(ValueError):
            validate_source_selection(self.rule(source_paths=[self.a + "\n"]))

    def test_invalid_output_names(self):
        names = ["bad/name", "bad\\name", "bad:name", "bad?.xlsx", "bad.xlsx.txt", "CON.foo.xlsx", "NUL.xlsx", ".xlsx", "bad. ", "bad\tname"]
        for name in names:
            with self.subTest(name=name):
                with self.assertRaises(ValueError):
                    validate_source_selection(self.rule(source_output_filename=name))
        with self.assertRaises(ValueError):
            validate_source_selection(self.rule(source_output_filename=123))

    def test_no_file_existence_check(self):
        validate_source_selection(self.rule(source_paths=[os.path.join(self.root, "does-not-exist.xlsx")]))


if __name__ == "__main__":
    unittest.main()
