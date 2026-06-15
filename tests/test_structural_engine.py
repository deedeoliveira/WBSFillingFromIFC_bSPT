import pytest
import sys, os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from app.core.structural_engine import (
    migrate_rule_v1_to_v2,
    load_and_migrate_rules,
    IFCInvestigator,
    NO_ELEMENTS_FOUND,
)


class TestMigrateRuleV1ToV2:

    def test_v1_rule_is_migrated(self):
        v1 = {
            "filter": {
                "ifc_class": "IfcWall",
                "predefined": "SOLIDWALL",
                "props": [],
            },
            "material": "Concrete",
            "quantity": {"pset": "Qto_WallBaseQuantities", "prop": "NetVolume"},
            "agrupamento": {"pset": "bSPT", "prop": "WbsGrouping"},
        }
        v2 = migrate_rule_v1_to_v2(v1)
        assert "mappings" in v2
        assert len(v2["mappings"]) == 1
        assert v2["mappings"][0]["filter"]["ifc_class"] == "IfcWall"
        assert v2["mappings"][0]["quantity_detail"]["pset"] == "Qto_WallBaseQuantities"
        assert v2["mappings"][0]["quantity_detail"]["prop"] == "NetVolume"
        assert v2["material"] == "Concrete"
        assert v2["quantity"]["type"] == "prop"

    def test_v2_rule_is_unchanged(self):
        v2 = {
            "mappings": [{"filter": {"ifc_class": "IfcWall", "predefined": "SOLIDWALL"}}],
            "material": "",
            "quantity": {"type": "count"},
            "agrupamento": {},
        }
        result = migrate_rule_v1_to_v2(v2)
        assert result is v2

    def test_v1_without_quantity_migrates_as_prop(self):
        v1 = {
            "filter": {"ifc_class": "IfcSlab", "predefined": "FLOOR", "props": []},
            "material": "",
            "quantity": {},
            "agrupamento": {},
        }
        v2 = migrate_rule_v1_to_v2(v1)
        assert v2["quantity"]["type"] == "prop"


class TestLoadAndMigrateRules:

    def test_version1_file(self):
        data = {
            "version": 1,
            "rules": {
                "08.01": {
                    "filter": {"ifc_class": "IfcWall", "predefined": "SOLIDWALL", "props": []},
                    "material": "",
                    "quantity": {"pset": "Qto", "prop": "NetVolume"},
                    "agrupamento": {},
                }
            }
        }
        rules = load_and_migrate_rules(data)
        assert "08.01" in rules
        assert "mappings" in rules["08.01"]

    def test_version2_file(self):
        data = {
            "version": 2,
            "rules": {
                "08.01": {
                    "mappings": [{"filter": {"ifc_class": "IfcWall", "predefined": "SOLIDWALL"}}],
                    "quantity": {"type": "count"},
                    "material": "",
                    "agrupamento": {},
                }
            }
        }
        rules = load_and_migrate_rules(data)
        assert rules["08.01"]["quantity"]["type"] == "count"

    def test_empty_rules(self):
        rules = load_and_migrate_rules({"version": 2, "rules": {}})
        assert rules == {}

    def test_partial_flag_preserved(self):
        data = {
            "version": 2,
            "partial": True,
            "rules": {}
        }
        rules = load_and_migrate_rules(data)
        assert isinstance(rules, dict)


class TestCountElements:

    def _make_mock_elements(self, n):
        class MockEl:
            def __init__(self, i):
                self.GlobalId = f"GUID_{i:04d}"
        return [MockEl(i) for i in range(n)]

    def test_count_empty(self):
        inv = IFCInvestigator()
        count, details = inv.count_elements([])
        assert count == 0
        assert details == []

    def test_count_correct(self):
        inv = IFCInvestigator()
        elems = self._make_mock_elements(5)
        count, details = inv.count_elements(elems)
        assert count == 5
        assert len(details) == 5

    def test_each_detail_has_valor_1(self):
        inv = IFCInvestigator()
        elems = self._make_mock_elements(3)
        _, details = inv.count_elements(elems)
        for d in details:
            assert d["valor"] == 1.0

    def test_guid_captured(self):
        inv = IFCInvestigator()
        elems = self._make_mock_elements(2)
        _, details = inv.count_elements(elems)
        assert details[0]["guid"] == "GUID_0000"
        assert details[1]["guid"] == "GUID_0001"


class TestSumQuantity:

    def _make_mock_element(self, guid, pset_name, prop_name, value):
        import unittest.mock as mock

        el = mock.MagicMock()
        el.GlobalId = guid
        el.is_a.return_value = "IfcWall"

        import ifcopenshell.util.element as util_el
        el._mock_psets = {pset_name: {prop_name: value}}
        return el

    def test_sum_with_mock(self, monkeypatch):
        import unittest.mock as mock
        import ifcopenshell.util.element as util_el

        class FakeEl:
            GlobalId = "G001"

        el = FakeEl()

        monkeypatch.setattr(
            util_el, "get_psets",
            lambda e: {"Qto_WallBaseQuantities": {"NetVolume": 3.5}}
        )

        inv = IFCInvestigator()
        total, details = inv.sum_quantity([el], "Qto_WallBaseQuantities", "NetVolume")
        assert abs(total - 3.5) < 1e-9
        assert len(details) == 1
        assert details[0]["valor"] == 3.5

    def test_sum_multiple_elements(self, monkeypatch):
        import ifcopenshell.util.element as util_el

        class FakeEl:
            def __init__(self, g):
                self.GlobalId = g

        elems = [FakeEl("G1"), FakeEl("G2"), FakeEl("G3")]
        values = [1.0, 2.5, 0.5]

        call_count = [0]
        def fake_psets(e):
            v = values[call_count[0] % len(values)]
            call_count[0] += 1
            return {"MyPset": {"MyProp": v}}

        monkeypatch.setattr(util_el, "get_psets", fake_psets)

        inv = IFCInvestigator()
        total, details = inv.sum_quantity(elems, "MyPset", "MyProp")
        assert abs(total - 4.0) < 1e-9


class TestExtractQuantities:

    def test_count_type_returns_int(self, monkeypatch):
        inv = IFCInvestigator()

        class FakeEl:
            GlobalId = "G1"
            def is_a(self): return "IfcDoor"
            PredefinedType = "DOOR"

        monkeypatch.setattr(inv, "filter_elements_for_mapping", lambda m: [FakeEl()] * 3)
        monkeypatch.setattr(inv, "_apply_material_filter", lambda elems, mat: elems)

        rule = {
            "mappings": [{"filter": {"ifc_class": "IfcDoor", "predefined": "DOOR"}}],
            "material": "",
            "quantity": {"type": "count"},
            "agrupamento": {},
        }
        total, details, found_any = inv.extract_quantities(rule)
        assert found_any is True
        assert total == 3
        assert isinstance(total, int)

    def test_no_elements_returns_found_any_false(self, monkeypatch):
        inv = IFCInvestigator()
        monkeypatch.setattr(inv, "filter_elements_for_mapping", lambda m: [])
        monkeypatch.setattr(inv, "_apply_material_filter", lambda elems, mat: elems)

        rule = {
            "mappings": [{"filter": {"ifc_class": "IfcWall", "predefined": "SOLIDWALL"}}],
            "material": "",
            "quantity": {"type": "count"},
            "agrupamento": {},
        }
        total, details, found_any = inv.extract_quantities(rule)
        assert found_any is False
        assert total == 0

    def test_multi_mapping_aggregates(self, monkeypatch):
        inv = IFCInvestigator()

        call_count = [0]
        def fake_filter(m):
            call_count[0] += 1
            class E:
                GlobalId = f"G{call_count[0]}"
                def is_a(self): return "IfcDoor"
                PredefinedType = "DOOR"
            return [E(), E()]

        monkeypatch.setattr(inv, "filter_elements_for_mapping", fake_filter)
        monkeypatch.setattr(inv, "_apply_material_filter", lambda elems, mat: elems)

        rule = {
            "mappings": [
                {"filter": {"ifc_class": "IfcDoor",   "predefined": "DOOR"}},
                {"filter": {"ifc_class": "IfcWindow",  "predefined": "WINDOW"}},
            ],
            "material": "",
            "quantity": {"type": "count"},
            "agrupamento": {},
        }
        total, details, found_any = inv.extract_quantities(rule)
        assert found_any is True
        assert total == 4

    def test_v1_rule_migrated_automatically(self, monkeypatch):
        inv = IFCInvestigator()

        class FakeEl:
            GlobalId = "G1"
            def is_a(self): return "IfcWall"
            PredefinedType = "SOLIDWALL"

        import ifcopenshell.util.element as util_el
        monkeypatch.setattr(util_el, "get_psets",
                            lambda e: {"Qto": {"NetVolume": 2.0}})
        monkeypatch.setattr(inv, "filter_elements_for_mapping", lambda m: [FakeEl()])
        monkeypatch.setattr(inv, "_apply_material_filter", lambda elems, mat: elems)

        v1_rule = {
            "filter": {"ifc_class": "IfcWall", "predefined": "SOLIDWALL", "props": []},
            "material": "",
            "quantity": {"pset": "Qto", "prop": "NetVolume"},
            "agrupamento": {},
        }
        total, details, found_any = inv.extract_quantities(v1_rule)
        assert found_any is True
        assert abs(total - 2.0) < 1e-9
