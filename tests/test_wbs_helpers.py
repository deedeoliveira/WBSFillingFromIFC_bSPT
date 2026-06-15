import pytest
import pandas as pd
import sys, os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from app.gui.wbs_helpers import (
    _build_filter_specs,
    extract_partial_mapping,
    find_wbs_columns,
    unpack_core_columns,
    split_levels,
)


class TestBuildFilterSpecs:

    def test_case1_dot_notation_single(self):
        specs = _build_filter_specs("IfcFooting.PILE_CAP", "", "")
        assert len(specs) == 1
        assert specs[0]["ifc_class"] == "IfcFooting"
        assert specs[0]["predefined"] == "PILE_CAP"

    def test_case1_dot_notation_multiple_tokens(self):
        specs = _build_filter_specs("IfcDoor.DOOR / IfcWindow.WINDOW", "", "")
        assert len(specs) == 2
        assert specs[0]["ifc_class"] == "IfcDoor"
        assert specs[0]["predefined"] == "DOOR"
        assert specs[1]["ifc_class"] == "IfcWindow"
        assert specs[1]["predefined"] == "WINDOW"

    def test_case2_multiple_classes_paired(self):
        specs = _build_filter_specs("IfcDoor / IfcWindow", "DOOR / WINDOW", "")
        assert len(specs) == 2
        assert specs[0]["ifc_class"] == "IfcDoor"
        assert specs[0]["predefined"] == "DOOR"
        assert specs[1]["ifc_class"] == "IfcWindow"
        assert specs[1]["predefined"] == "WINDOW"

    def test_case3_single_class_multiple_predefined(self):
        specs = _build_filter_specs("IfcCovering", "CEILING / FLOORING / CLADDING", "")
        assert len(specs) == 3
        for s in specs:
            assert s["ifc_class"] == "IfcCovering"
        predefs = [s["predefined"] for s in specs]
        assert predefs == ["CEILING", "FLOORING", "CLADDING"]

    def test_case4_simple_one_to_one(self):
        specs = _build_filter_specs("IfcWall", "SOLIDWALL", "")
        assert len(specs) == 1
        assert specs[0]["ifc_class"] == "IfcWall"
        assert specs[0]["predefined"] == "SOLIDWALL"

    def test_empty_class_returns_empty(self):
        specs = _build_filter_specs("", "SOLIDWALL", "")
        assert specs == []

    def test_object_type_set_when_userdefined(self):
        specs = _build_filter_specs("IfcWall", "USERDEFINED", "MyWallType")
        assert specs[0]["object_type"] == "MyWallType"

    def test_object_type_empty_when_not_userdefined(self):
        specs = _build_filter_specs("IfcWall", "SOLIDWALL", "ShouldBeIgnored")
        assert specs[0]["object_type"] == ""

    def test_whitespace_stripped(self):
        specs = _build_filter_specs("  IfcSlab  ", "  FLOOR  ", "")
        assert specs[0]["ifc_class"] == "IfcSlab"
        assert specs[0]["predefined"] == "FLOOR"

    def test_props_always_empty_list(self):
        specs = _build_filter_specs("IfcWall", "SOLIDWALL", "")
        assert specs[0]["props"] == []


class TestExtractPartialMapping:

    def _make_df(self, rows):
        return pd.DataFrame(rows)

    def _cols(self):
        return {
            "col_wbs":       "WBS",
            "col_nivel":     "NIVEL",
            "col_desc":      "DESC",
            "col_ifc_class": "IFC Class",
            "col_predef":    "PredefinedType",
            "col_objtype":   "ObjectType",
            "col_ifc_prop":  "IFC Property",
            "col_unidades":  "UNIDADES",
        }

    def test_basic_row_with_prop(self):
        df = self._make_df([{
            "NIVEL": 3, "WBS": "08.01.01", "DESC": "Test",
            "IFC Class": "IfcWall", "PredefinedType": "SOLIDWALL",
            "ObjectType": "", "IFC Property": "NetVolume", "UNIDADES": "m3",
        }])
        rules = extract_partial_mapping(df, self._cols())
        assert "08.01.01" in rules
        r = rules["08.01.01"]
        assert r["quantity"]["type"] == "prop"
        assert r["mappings"][0]["quantity_detail"]["prop"] == "NetVolume"
        assert r["mappings"][0]["filter"]["ifc_class"] == "IfcWall"

    def test_row_without_ifc_prop_defaults_to_count(self):
        df = self._make_df([{
            "NIVEL": 3, "WBS": "08.01.02", "DESC": "Test",
            "IFC Class": "IfcDoor", "PredefinedType": "DOOR",
            "ObjectType": "", "IFC Property": "", "UNIDADES": "un",
        }])
        rules = extract_partial_mapping(df, self._cols())
        assert rules["08.01.02"]["quantity"]["type"] == "count"

    def test_row_without_ifc_class_is_skipped(self):
        df = self._make_df([{
            "NIVEL": 3, "WBS": "08.01.03", "DESC": "Test",
            "IFC Class": "", "PredefinedType": "",
            "ObjectType": "", "IFC Property": "", "UNIDADES": "",
        }])
        rules = extract_partial_mapping(df, self._cols())
        assert "08.01.03" not in rules

    def test_level10_row_is_skipped(self):
        df = self._make_df([{
            "NIVEL": 10, "WBS": "", "DESC": "descrição do utilizador",
            "IFC Class": "IfcWall", "PredefinedType": "SOLIDWALL",
            "ObjectType": "", "IFC Property": "NetVolume", "UNIDADES": "",
        }])
        rules = extract_partial_mapping(df, self._cols())
        assert rules == {}

    def test_ifc_prop_with_dot_notation_pset_and_prop(self):
        df = self._make_df([{
            "NIVEL": 3, "WBS": "08.01.04", "DESC": "Test",
            "IFC Class": "IfcWall", "PredefinedType": "SOLIDWALL",
            "ObjectType": "", "IFC Property": "Qto_WallBaseQuantities.NetVolume",
            "UNIDADES": "m3",
        }])
        rules = extract_partial_mapping(df, self._cols())
        qd = rules["08.01.04"]["mappings"][0]["quantity_detail"]
        assert qd["pset"] == "Qto_WallBaseQuantities"
        assert qd["prop"] == "NetVolume"

    def test_multiple_predefined_types_generate_multiple_mappings(self):
        df = self._make_df([{
            "NIVEL": 3, "WBS": "13.02.01", "DESC": "Test",
            "IFC Class": "IfcCovering",
            "PredefinedType": "CEILING / FLOORING",
            "ObjectType": "", "IFC Property": "", "UNIDADES": "m2",
        }])
        rules = extract_partial_mapping(df, self._cols())
        assert len(rules["13.02.01"]["mappings"]) == 2

    def test_multiple_wbs_rows(self):
        df = self._make_df([
            {"NIVEL": 3, "WBS": "08.01.01", "DESC": "A",
             "IFC Class": "IfcWall", "PredefinedType": "SOLIDWALL",
             "ObjectType": "", "IFC Property": "NetVolume", "UNIDADES": "m3"},
            {"NIVEL": 3, "WBS": "08.01.02", "DESC": "B",
             "IFC Class": "IfcSlab", "PredefinedType": "FLOOR",
             "ObjectType": "", "IFC Property": "", "UNIDADES": "m2"},
            {"NIVEL": 3, "WBS": "08.01.03", "DESC": "C",
             "IFC Class": "", "PredefinedType": "",
             "ObjectType": "", "IFC Property": "", "UNIDADES": ""},
        ])
        rules = extract_partial_mapping(df, self._cols())
        assert "08.01.01" in rules
        assert "08.01.02" in rules
        assert "08.01.03" not in rules


class TestFindWbsColumns:

    def test_detects_core_columns(self):
        df = pd.DataFrame(columns=["NÍVEL", "WBS", "DESCRIÇÃO", "UNIDADES"])
        cols = find_wbs_columns(df)
        col_wbs, col_desc, col_nivel = unpack_core_columns(cols)
        assert col_wbs   == "WBS"
        assert col_nivel == "NÍVEL"
        assert col_desc  == "DESCRIÇÃO"

    def test_detects_ifc_columns(self):
        df = pd.DataFrame(columns=[
            "NÍVEL", "WBS", "DESCRIÇÃO", "UNIDADES",
            "IFC Class", "PredefinedType", "ObjectType", "IFC Property",
        ])
        cols = find_wbs_columns(df)
        assert cols.get("col_ifc_class") == "IFC Class"
        assert cols.get("col_predef")    == "PredefinedType"
        assert cols.get("col_objtype")   == "ObjectType"
        assert cols.get("col_ifc_prop")  == "IFC Property"

    def test_missing_ifc_columns_return_none(self):
        df = pd.DataFrame(columns=["NÍVEL", "WBS", "DESCRIÇÃO"])
        cols = find_wbs_columns(df)
        assert cols.get("col_ifc_class") is None
        assert cols.get("col_predef")    is None


class TestSplitLevels:

    def test_integer_values(self):
        df = pd.DataFrame({"NIVEL": [1, 2, 3, 10]})
        result = split_levels(df, "NIVEL")
        assert list(result) == [1, 2, 3, 10]

    def test_float_values(self):
        df = pd.DataFrame({"NIVEL": [1.0, 2.0, 10.0]})
        result = split_levels(df, "NIVEL")
        assert list(result) == [1, 2, 10]

    def test_nan_returns_none(self):
        import math
        df = pd.DataFrame({"NIVEL": [1, float("nan"), 3]})
        result = split_levels(df, "NIVEL")
        val = result.iat[1]
        assert val is None or (val != val)

    def test_string_numbers(self):
        df = pd.DataFrame({"NIVEL": ["1", "2", "10"]})
        result = split_levels(df, "NIVEL")
        assert list(result) == [1, 2, 10]
