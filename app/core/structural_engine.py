import ifcopenshell
import ifcopenshell.util.element

NO_ELEMENTS_FOUND = "NO_ELEMENTS_FOUND"


def migrate_rule_v1_to_v2(rule: dict) -> dict:
    if "mappings" in rule:
        return rule

    old_filter = rule.get("filter", {})
    old_qty = rule.get("quantity", {})

    mapping_entry = {"filter": old_filter}

    qty_type = "prop"
    if old_qty.get("pset") or old_qty.get("prop"):
        mapping_entry["quantity_detail"] = {
            "pset": old_qty.get("pset", ""),
            "prop": old_qty.get("prop", ""),
        }

    return {
        "mappings": [mapping_entry],
        "material": rule.get("material", ""),
        "quantity": {"type": qty_type},
        "agrupamento": rule.get("agrupamento", {}),
    }


def migrate_rules_v1_to_v2(rules: dict) -> dict:
    return {code: migrate_rule_v1_to_v2(r) for code, r in rules.items()}


def load_and_migrate_rules(data: dict) -> dict:
    version = str(data.get("version", "1"))
    rules = data.get("rules", data)
    if not isinstance(rules, dict):
        rules = {}
    if version == "1":
        rules = migrate_rules_v1_to_v2(rules)
    return rules


class IFCInvestigator:
    def __init__(self):
        self.ifc_file = None
        self.index_by_class = {}
        self.predefs_by_class = {}

    def open_ifc(self, path: str):
        self.ifc_file = ifcopenshell.open(path)
        self.index_by_class.clear()
        self.predefs_by_class.clear()

        for e in self.ifc_file.by_type("IfcProduct"):
            etype = e.is_a()
            self.index_by_class.setdefault(etype, []).append(e)
            predef = getattr(e, "PredefinedType", None)
            if predef:
                self.predefs_by_class.setdefault(etype, set()).add(str(predef))

    def list_classes(self):
        return sorted(self.index_by_class.keys())

    def list_predefined_types(self, ifc_class: str):
        return sorted(self.predefs_by_class.get(ifc_class, []))

    def _filter_single(self, filter_spec: dict) -> list:
        if not self.ifc_file:
            raise RuntimeError("No IFC file loaded.")

        etype = filter_spec.get("ifc_class")
        if not etype:
            return []

        elems = self.index_by_class.get(etype, [])

        predef = filter_spec.get("predefined_type") or filter_spec.get("predefined")
        if predef:
            elems = [e for e in elems if str(getattr(e, "PredefinedType", "")) == predef]

        objtype = filter_spec.get("object_type")
        if predef == "USERDEFINED" and objtype:
            elems = [e for e in elems
                     if str(getattr(e, "ObjectType", "")).upper() == objtype.upper()]

        extras = filter_spec.get("extra_filters") or filter_spec.get("props", [])
        if extras:
            def _bool_from_ifc(v) -> bool | None:
                if isinstance(v, bool):
                    return v
                if v is None:
                    return None
                s = str(v).strip().upper()
                if s in ("TRUE", ".T.", "T", "1", "YES"):
                    return True
                if s in ("FALSE", ".F.", "F", "0", "NO"):
                    return False
                return None

            def _match_props(e):
                psets = ifcopenshell.util.element.get_psets(e)
                try:
                    type_psets = ifcopenshell.util.element.get_psets(e, psets_only=False)
                    merged = {**type_psets, **psets}
                except Exception:
                    merged = psets

                for ff in extras:
                    pset = ff.get("pset", "")
                    prop = ff.get("prop", "")
                    val  = ff.get("value")
                    if not pset or not prop or val is None:
                        continue
                    cur_val = merged.get(pset, {}).get(prop)
                    if isinstance(val, bool):
                        cur_bool = _bool_from_ifc(cur_val)
                        if cur_bool is None or cur_bool != val:
                            return False
                    elif isinstance(val, (int, float)):
                        try:
                            if float(cur_val) != float(val):
                                return False
                        except (ValueError, TypeError):
                            return False
                    else:
                        val_s = str(val).strip().lower()
                        cur_s = str(cur_val).strip().lower()
                        bool_map = {
                            ".t.": "true", "t": "true", "1": "true", "yes": "true",
                            ".f.": "false", "f": "false", "0": "false", "no": "false",
                        }
                        val_s = bool_map.get(val_s, val_s)
                        cur_s = bool_map.get(cur_s, cur_s)
                        if val_s != cur_s:
                            return False
                return True

            elems = [e for e in elems if _match_props(e)]

        return elems

    def _apply_material_filter(self, elems: list, mat_filter) -> list:
        if not mat_filter or str(mat_filter).strip() in ("", "None"):
            return elems

        if isinstance(mat_filter, dict):
            wanted_cat  = str(mat_filter.get("category", "") or "").strip().lower()
            wanted_name = str(mat_filter.get("name", "")     or "").strip().lower()
        else:
            wanted_cat  = str(mat_filter or "").strip().lower()
            wanted_name = ""

        def _materials_of(el):
            mats = []

            def _add(m):
                if not m:
                    return
                try:
                    kind = m.is_a()
                except Exception:
                    kind = ""
                if kind == "IfcMaterial":
                    mats.append({
                        "name":     (getattr(m, "Name",     "") or "").strip(),
                        "category": (getattr(m, "Category", "") or "").strip(),
                    })
                elif kind == "IfcMaterialLayer":
                    _add(getattr(m, "Material", None))
                elif kind == "IfcMaterialLayerSet":
                    for lyr in (getattr(m, "MaterialLayers", None) or []):
                        _add(getattr(lyr, "Material", None))
                elif kind == "IfcMaterialConstituent":
                    _add(getattr(m, "Material", None))
                elif kind == "IfcMaterialConstituentSet":
                    for c in (getattr(m, "MaterialConstituents", None) or []):
                        _add(getattr(c, "Material", None))
                else:
                    maybe = getattr(m, "Material", None)
                    if maybe:
                        _add(maybe)

            try:
                for rel in (getattr(el, "HasAssociations", None) or []):
                    if rel and getattr(rel, "is_a", lambda: "")() == "IfcRelAssociatesMaterial":
                        _add(getattr(rel, "RelatingMaterial", None))
            except Exception:
                pass
            return mats

        def _match(el):
            for d in _materials_of(el):
                nm  = (d.get("name",     "") or "").strip().lower()
                cat = (d.get("category", "") or "").strip().lower()
                if wanted_cat  and cat == wanted_cat:
                    return True
                if wanted_name and nm  == wanted_name:
                    return True
            return False

        return [e for e in elems if _match(e)]

    def filter_elements_for_mapping(self, mapping_entry: dict) -> list:
        elems = self._filter_single(mapping_entry.get("filter", {}))
        if not elems:
            return []
        return elems

    def filter_elements(self, rule: dict) -> list:
        rule = migrate_rule_v1_to_v2(rule)
        all_elems = []
        for m in rule.get("mappings", []):
            all_elems.extend(self.filter_elements_for_mapping(m))
        mat = rule.get("material")
        if mat:
            all_elems = self._apply_material_filter(all_elems, mat)
        return all_elems

    def sum_quantity(self, elements, q_pset: str, q_prop: str):
        total = 0.0
        details = []
        for e in elements:
            try:
                psets = ifcopenshell.util.element.get_psets(e)
                if not psets:
                    continue
                val = psets.get(q_pset, {}).get(q_prop)
                if val is None:
                    continue
                try:
                    num = float(str(val))
                    total += num
                    details.append({"element": e, "guid": e.GlobalId, "valor": num})
                except Exception:
                    pass
            except Exception:
                pass
        return total, details

    def count_elements(self, elements: list):
        count = len(elements)
        details = [
            {"element": e, "guid": getattr(e, "GlobalId", "?"), "valor": 1.0}
            for e in elements
        ]
        return count, details

    def extract_quantities(self, rule: dict):
        rule = migrate_rule_v1_to_v2(rule)
        qty_type = rule.get("quantity", {}).get("type", "prop")
        mat = rule.get("material")

        all_details = []

        for mapping_entry in rule.get("mappings", []):
            elems = self.filter_elements_for_mapping(mapping_entry)
            if mat:
                elems = self._apply_material_filter(elems, mat)
            if not elems:
                continue

            if qty_type == "count":
                _, dets = self.count_elements(elems)
                all_details.extend(dets)
            else:
                qty_detail = mapping_entry.get("quantity_detail", {})
                q_pset = qty_detail.get("pset", "")
                q_prop = qty_detail.get("prop", "")
                if not q_pset or not q_prop:
                    continue
                _, dets = self.sum_quantity(elems, q_pset, q_prop)
                all_details.extend(dets)

        found_any = bool(all_details)
        total = sum(d.get("valor", 0.0) for d in all_details)

        if qty_type == "count":
            total = int(total)

        return total, all_details, found_any

    def get_prop_values(self, elements, pset: str, prop: str):
        out = []
        seen = set()
        for e in elements or []:
            try:
                psets = ifcopenshell.util.element.get_psets(e)
                if not psets:
                    continue
                val = psets.get(pset, {}).get(prop)
                vals = val if isinstance(val, (list, tuple, set)) else [val]
                for v in vals:
                    s = "" if v is None else str(v).strip()
                    if s and s not in seen:
                        seen.add(s)
                        out.append(s)
            except Exception:
                continue
        return out

    def get_element_material(self, element):
        try:
            rels = getattr(element, "HasAssociations", [])
            for rel in rels:
                if rel.is_a("IfcRelAssociatesMaterial"):
                    mat = rel.RelatingMaterial
                    if mat.is_a("IfcMaterialLayerSet") and hasattr(mat, "MaterialLayers"):
                        layers = [l.Material.Name for l in mat.MaterialLayers if l.Material]
                        return ", ".join(layers)
                    elif hasattr(mat, "Name"):
                        return str(mat.Name)
            return "n/a"
        except Exception:
            return "n/a"

    def extract_all_materials(self):
        if not self.ifc_file:
            return {}
        materials = {}
        try:
            products = self.ifc_file.by_type("IfcProduct")
        except Exception:
            products = []

        for element in products:
            try:
                rels = getattr(element, "HasAssociations", [])
                for rel in rels:
                    if rel.is_a("IfcRelAssociatesMaterial"):
                        mat = rel.RelatingMaterial
                        if mat.is_a("IfcMaterialLayerSet") and hasattr(mat, "MaterialLayers"):
                            for layer in mat.MaterialLayers:
                                if layer.Material:
                                    m_name = str(layer.Material.Name) if hasattr(layer.Material, "Name") else "n/a"
                                    m_cat  = str(layer.Material.Category) if hasattr(layer.Material, "Category") else "n/a"
                                    if m_cat and m_cat != "n/a" and m_cat not in materials:
                                        materials[m_cat] = m_name
                        elif mat.is_a("IfcMaterial"):
                            m_name = str(mat.Name) if hasattr(mat, "Name") else "n/a"
                            m_cat  = str(mat.Category) if hasattr(mat, "Category") else "n/a"
                            if m_cat and m_cat != "n/a" and m_cat not in materials:
                                materials[m_cat] = m_name
                        elif mat.is_a("IfcMaterialList") and hasattr(mat, "Materials"):
                            for m in mat.Materials:
                                m_name = str(m.Name) if hasattr(m, "Name") else "n/a"
                                m_cat  = str(m.Category) if hasattr(m, "Category") else "n/a"
                                if m_cat and m_cat != "n/a" and m_cat not in materials:
                                    materials[m_cat] = m_name
            except Exception:
                pass
        return materials

    def get_project_info(self):
        if not self.ifc_file:
            return {"project": "n/a", "site": "n/a", "building": "n/a"}
        result = {"project": "n/a", "site": "n/a", "building": "n/a"}
        for key, ifc_type in [("project", "IfcProject"), ("site", "IfcSite"), ("building", "IfcBuilding")]:
            try:
                items = self.ifc_file.by_type(ifc_type)
                if items:
                    name = getattr(items[0], "Name", "n/a")
                    result[key] = str(name) if name else "n/a"
            except Exception:
                pass
        return result

    def _build_classification_index(self):
        self._class_by_obj = {}
        if not self.ifc_file:
            return
        try:
            rels = self.ifc_file.by_type("IfcRelAssociatesClassification")
        except Exception:
            rels = []
        for rel in rels or []:
            ref = getattr(rel, "RelatingClassification", None)
            code = None
            if ref is not None:
                if ref.is_a("IfcClassificationReference"):
                    code = (getattr(ref, "Identification", None) or
                            getattr(ref, "ItemReference", None) or
                            getattr(ref, "Name", None))
                elif ref.is_a("IfcClassification"):
                    code = (getattr(ref, "Identification", None) or
                            getattr(ref, "Name", None))
            if not code:
                code = "n/a"
            for obj in getattr(rel, "RelatedObjects", []) or []:
                try:
                    self._class_by_obj[obj.id()] = str(code)
                except Exception:
                    pass

    def get_classification_code(self, element):
        if element is None:
            return "n/a"
        if not hasattr(self, "_class_by_obj"):
            self._build_classification_index()
        return self._class_by_obj.get(element.id(), "n/a")

    def get_building_storey(self, element):
        if not element or not self.ifc_file:
            return "n/a"
        try:
            import ifcopenshell.util.element as util_element
        except Exception:
            util_element = None
        container = None
        if util_element:
            try:
                container = util_element.get_container(element)
            except Exception:
                container = None
        visited = set()
        cur = container
        while cur and id(cur) not in visited:
            visited.add(id(cur))
            try:
                if cur.is_a("IfcBuildingStorey"):
                    return getattr(cur, "Name", None) or getattr(cur, "LongName", None) or "n/a"
            except Exception:
                pass
            try:
                dec = getattr(cur, "Decomposes", None)
                if dec and len(dec) > 0:
                    cur = getattr(dec[0], "RelatingObject", None)
                    continue
            except Exception:
                pass
            break
        return "n/a"
