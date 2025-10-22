# structural_engine.py
import ifcopenshell
import ifcopenshell.util.element

class IFCInvestigator:
    def __init__(self):
        self.ifc_file = None
        self.index_by_class = {}
        self.predefs_by_class = {}

    def open_ifc(self, path: str):
        """Abre o ficheiro IFC e indexa entidades por classe."""
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
        """Lista todas as IfcClasses encontradas no IFC."""
        return sorted(self.index_by_class.keys())

    def list_predefined_types(self, ifc_class: str):
        """Lista os PredefinedTypes únicos para uma dada classe."""
        return sorted(self.predefs_by_class.get(ifc_class, []))

    def filter_elements(self, rule: dict):
        if not self.ifc_file:
            raise RuntimeError("Nenhum IFC carregado.")

        f = rule.get("filter", {})
        etype = f.get("ifc_class")
        if not etype:
            print("[DEBUG] Nenhuma IFC class definida na regra.")
            return []

        elems = self.index_by_class.get(etype, [])
        print(f"[DEBUG] Classe {etype} => {len(elems)} elementos antes de filtros")

        predef = f.get("predefined_type") or f.get("predefined")
        if predef:
            elems = [e for e in elems if str(getattr(e, "PredefinedType", "")) == predef]
            print(f"[DEBUG] Filtro PredefinedType={predef} => {len(elems)} elementos")

        objtype = f.get("object_type")
        if predef == "USERDEFINED" and objtype:
            elems = [e for e in elems if str(getattr(e, "ObjectType", "")).upper() == objtype.upper()]
            print(f"[DEBUG] Filtro ObjectType={objtype} => {len(elems)} elementos")

        mat_filter = rule.get("material")
        if mat_filter and str(mat_filter).strip() not in ("", "None"):
            before_mat = len(elems)
            print("\n[DBG] ===== FILTRO DE MATERIAL =====")
            print(f"[DBG] Regra.material recebido do JSON: {mat_filter!r}")
            print(f"[DBG] Elementos antes do filtro de material: {before_mat}")

            wanted_cat  = ""
            wanted_name = ""
            if isinstance(mat_filter, dict):
                wanted_cat  = (str(mat_filter.get("category", "") or "").strip().lower())
                wanted_name = (str(mat_filter.get("name", "") or "").strip().lower())
            else:
                wanted_cat = str(mat_filter or "").strip().lower()

            print(f"[DBG] Categoria esperada: {wanted_cat!r} | Nome esperado: {wanted_name!r}")

            def _materials_of(el):
                """Coleta materiais do elemento (Name/Category) via RelAssociatesMaterial e estruturas comuns."""
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
                            "name":     (getattr(m, "Name", "") or "").strip(),
                            "category": (getattr(m, "Category", "") or "").strip(),
                        })
                        return
                    if kind == "IfcMaterialLayer":
                        _add(getattr(m, "Material", None)); return
                    if kind == "IfcMaterialLayerSet":
                        for lyr in (getattr(m, "MaterialLayers", None) or []):
                            _add(getattr(lyr, "Material", None)); 
                        return
                    if kind == "IfcMaterialConstituent":
                        _add(getattr(m, "Material", None)); return
                    if kind == "IfcMaterialConstituentSet":
                        for c in (getattr(m, "MaterialConstituents", None) or []):
                            _add(getattr(c, "Material", None)); 
                        return
                    maybe = getattr(m, "Material", None)
                    if maybe:
                        _add(maybe)

                try:
                    for rel in (getattr(el, "HasAssociations", None) or []):
                        if rel and getattr(rel, "is_a", lambda: "")() == "IfcRelAssociatesMaterial":
                            _add(getattr(rel, "RelatingMaterial", None))
                except Exception as ex:
                    print(f"[DBG] Aviso: falha ao varrer materiais de {getattr(el, 'GlobalId', '?')}: {ex}")
                return mats

            sample_n = min(5, len(elems))
            print(f"[DBG] A inspecionar materiais dos primeiros {sample_n} elementos (antes do filtro)…")
            for i, el in enumerate(elems[:sample_n], 1):
                gid = getattr(el, "GlobalId", "?")
                lst = _materials_of(el)
                pretty = [f"{d.get('name','')} | cat={d.get('category','')}" for d in lst] or ["— sem materiais —"]
                print(f"  [DBG] {i:02d}) {el.is_a()} {gid} -> {pretty}")

            def _match_material(el):
                mats = _materials_of(el)
                if not mats:
                    return False
                for d in mats:
                    nm = (d.get("name") or "").strip().lower()
                    cat = (d.get("category") or "").strip().lower()
                    if wanted_cat and cat == wanted_cat:
                        return True
                    if wanted_name and nm == wanted_name:
                        return True
                return False

            elems = [e for e in elems if _match_material(e)]
            after_mat = len(elems)
            print(f"[DBG] Elementos após filtro de material: {after_mat} (removidos: {before_mat - after_mat})")
            print("[DBG] ===== FIM FILTRO DE MATERIAL =====\n")
        else:

            print("[DBG] Regra sem filtro de material: ignorado (material opcional).")

        extras = f.get("extra_filters") or f.get("props", [])
        if extras:
            def match(e):
                psets = ifcopenshell.util.element.get_psets(e)
                for ff in extras:
                    pset = ff.get("pset", "")
                    prop = ff.get("prop", "")
                    val = ff.get("value")
                    
                    if not pset or not prop or val is None:
                        continue
                    
                    cur_val = psets.get(pset, {}).get(prop)
                    
                    if isinstance(val, bool):
                        if cur_val != val:
                            return False

                    elif isinstance(val, (int, float)):
                        try:
                            if float(cur_val) != float(val):
                                return False
                        except (ValueError, TypeError):
                            return False

                    else:
                        val_str = str(val).strip().lower()
                        cur_str = str(cur_val).strip().lower()
                        if cur_str != val_str:
                            return False
                return True
            elems = [e for e in elems if match(e)]
            print(f"[DEBUG] Filtros adicionais aplicados => {len(elems)} elementos")

        print(f"[DEBUG] Resultado final para {etype}: {len(elems)} elementos")
        return elems
        
    def sum_quantity(self, elements, q_pset, q_prop):
        print(f"[DEBUG] Somando quantidade: pset={q_pset}, prop={q_prop}, elementos={len(elements)}")

        total = 0.0
        details = []

        for e in elements:
            try:
                psets = ifcopenshell.util.element.get_psets(e)
                if not psets:
                    print(f"[DEBUG] Elemento {getattr(e, 'GlobalId', '?')} sem Psets.")
                    continue

                val = psets.get(q_pset, {}).get(q_prop)
                if val is None:
                    print(f"[DEBUG] Elemento {getattr(e, 'GlobalId', '?')} não tem {q_pset}:{q_prop}")
                    continue

                print(f"[DEBUG] Elemento {e.is_a()} {getattr(e, 'GlobalId', '?')} → valor bruto = {val}")

                try:
                    num = float(str(val))
                    total += num
                    details.append({"element": e, "guid": e.GlobalId, "valor": num})
                    print(f"[VALOR] {e.is_a()} {e.GlobalId}: {num}")
                except Exception as conv_err:
                    print(f"[ERRO] Não consegui converter valor '{val}' para número ({conv_err})")

            except Exception as ex:
                print(f"[ERRO] Falha ao processar {e.is_a()} {getattr(e, 'GlobalId', '?')}: {ex}")

        print(f"[DEBUG] Soma final para {q_pset}:{q_prop} = {total}")
        return total, details
        
    def get_prop_values(self, elements, pset: str, prop: str):
        """Retorna lista de strings únicas com os valores encontrados no pset/prop dos elementos."""
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
                                        m_cat = str(layer.Material.Category) if hasattr(layer.Material, "Category") else "n/a"
                                        if m_cat and m_cat != "n/a" and m_cat not in materials:
                                            materials[m_cat] = m_name

                            elif mat.is_a("IfcMaterial"):
                                m_name = str(mat.Name) if hasattr(mat, "Name") else "n/a"
                                m_cat = str(mat.Category) if hasattr(mat, "Category") else "n/a"
                                if m_cat and m_cat != "n/a" and m_cat not in materials:
                                    materials[m_cat] = m_name

                            elif mat.is_a("IfcMaterialList") and hasattr(mat, "Materials"):
                                for m in mat.Materials:
                                    m_name = str(m.Name) if hasattr(m, "Name") else "n/a"
                                    m_cat = str(m.Category) if hasattr(m, "Category") else "n/a"
                                    if m_cat and m_cat != "n/a" and m_cat not in materials:
                                        materials[m_cat] = m_name

                except Exception as ex:
                    print(f"[DEBUG] Error extracting material from {element.is_a()}: {ex}")
                    continue

            return materials
            
    def get_project_info(self):
            if not self.ifc_file:
                return {"project": "n/a", "site": "n/a", "building": "n/a"}

            result = {"project": "n/a", "site": "n/a", "building": "n/a"}

            try:
                projects = self.ifc_file.by_type("IfcProject")
                if projects:
                    proj_name = getattr(projects[0], "Name", "n/a")
                    result["project"] = str(proj_name) if proj_name else "n/a"
            except Exception as ex:
                print(f"[DEBUG] Error extracting IfcProject: {ex}")

            try:
                sites = self.ifc_file.by_type("IfcSite")
                if sites:
                    site_name = getattr(sites[0], "Name", "n/a")
                    result["site"] = str(site_name) if site_name else "n/a"
            except Exception as ex:
                print(f"[DEBUG] Error extracting IfcSite: {ex}")

            try:
                buildings = self.ifc_file.by_type("IfcBuilding")
                if buildings:
                    bldg_name = getattr(buildings[0], "Name", "n/a")
                    result["building"] = str(bldg_name) if bldg_name else "n/a"
            except Exception as ex:
                print(f"[DEBUG] Error extracting IfcBuilding: {ex}")

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

                    code = getattr(ref, "Identification", None) or \
                           getattr(ref, "ItemReference", None) or \
                           getattr(ref, "Name", None)
                elif ref.is_a("IfcClassification"):

                    code = getattr(ref, "Identification", None) or \
                           getattr(ref, "Name", None)

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

                    rel = dec[0]
                    cur = getattr(rel, "RelatingObject", None)
                    continue
            except Exception:
                pass


            break


        return "n/a"




