# -*- coding: utf-8 -*-
"""Extrai metricas BIM brutas dos .rvt recebidos via pyRevit CLI.

Compativel com IronPython 2.7 (engine do pyRevit no Revit 2020).

Funciona de duas formas:
  1) pyRevit CLI: o modelo e' passado na variavel global __models__
     e' aberto explicitamente com Application.OpenDocumentFile
  2) GUI (clique no pushbutton): processa o revit.doc ativo
"""
import os
import re
import sys
import json
import codecs
import traceback

# garante imports locais (bim_paths.py na mesma pasta)
sys.path.append(os.path.dirname(__file__))

from bim_paths import PASTA_MODELOS, PASTA_SAIDA

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, ImportInstance,
    FamilyInstance, ViewSheet, View, Level, IndependentTag,
    ViewSchedule, ViewType, ProjectLocation, ElementId,
    RevitLinkInstance, CADLinkType, FamilySymbol, WorksetTable
)
from pyrevit import revit, forms, script

log = script.get_logger()


# ============================================================
#  REGEX DE CONFORMIDADE (IronPython 2.7 usa ur"...")
# ============================================================
RX_ARQUIVO = re.compile(ur"^(ME|PA|AG)[-_].+_ARQ_EP_R\d+", re.IGNORECASE)
RX_NIVEL   = re.compile(ur"^(NIVEL|PAV|N)\s*[-_]?\s*\d+", re.IGNORECASE)
RX_VISTA   = re.compile(ur"^(PLANTA|CORTE|FACHADA|ELEV|DETALHE|3D|VISTA|IMPLANTA)",
                        re.IGNORECASE)
RX_FOLHA   = re.compile(ur"^[A-Z]{2,4}-\d{2,3}", re.IGNORECASE)
RX_TABELA  = re.compile(ur"^(QT|TB|LISTA|QUANT)[-_]", re.IGNORECASE)


# ============================================================
#  COLETA DE METRICAS
# ============================================================
def coletar_metricas(doc):
    """Coleta resiliente: cada bloco em try/except.
    Uma falha pontual nao derruba o JSON inteiro."""
    m = {}

    # ---- metadados do arquivo ----
    try:
        m["nome_arquivo"] = os.path.basename(doc.PathName)
    except Exception:
        m["nome_arquivo"] = u"desconhecido"

    try:
        m["nomenclatura_arquivo"] = int(bool(RX_ARQUIVO.match(m["nome_arquivo"])))
    except Exception:
        m["nomenclatura_arquivo"] = 0

    # ---- elementos ----
    try:
        m["elementos_voados"] = _conta_elementos_voados(doc)
    except Exception:
        m["elementos_voados"] = -1

    try:
        m["dwg_refs"] = _conta_dwgs(doc)
    except Exception:
        m["dwg_refs"] = -1

    try:
        m["modelagem_local"] = _conta_inplace(doc)
    except Exception:
        m["modelagem_local"] = -1

    # ---- niveis e vistas ----
    try:
        niveis = list(FilteredElementCollector(doc).OfClass(Level))
    except Exception:
        niveis = []

    try:
        vistas = []
        for v in FilteredElementCollector(doc).OfClass(View):
            try:
                if v.ViewType not in (ViewType.DrawingSheet, ViewType.Internal,
                                      ViewType.ProjectBrowser):
                    vistas.append(v)
            except Exception:
                pass
    except Exception:
        vistas = []

    try:
        folhas = list(FilteredElementCollector(doc).OfClass(ViewSheet))
    except Exception:
        folhas = []

    # ---- contagens ----
    try:
        m["total_niveis"] = len(niveis)
    except Exception:
        m["total_niveis"] = -1

    try:
        m["nivel_fora_padrao"] = sum(1 for n in niveis
                                     if not RX_NIVEL.match(n.Name or u""))
    except Exception:
        m["nivel_fora_padrao"] = -1

    try:
        m["total_vistas"] = len(vistas)
    except Exception:
        m["total_vistas"] = -1

    try:
        m["vista_fora_padrao"] = sum(1 for v in vistas
                                     if not RX_VISTA.match(v.Name or u""))
    except Exception:
        m["vista_fora_padrao"] = -1

    try:
        m["nivel_vista_fora_padrao"] = (
            (m.get("nivel_fora_padrao") or 0) + (m.get("vista_fora_padrao") or 0))
    except Exception:
        m["nivel_vista_fora_padrao"] = -1

    try:
        m["total_folhas"] = len(folhas)
    except Exception:
        m["total_folhas"] = -1

    try:
        m["folhas_fora_padrao"] = sum(1 for f in folhas
                                      if not RX_FOLHA.match(f.SheetNumber or u""))
    except Exception:
        m["folhas_fora_padrao"] = -1

    # ---- carimbos, genericos, tabelas ----
    try:
        m["carimbo_incompleto"] = _conta_carimbos_incompletos(folhas)
    except Exception:
        m["carimbo_incompleto"] = -1

    try:
        m["familias_generico"] = _conta_genericos(doc)
    except Exception:
        m["familias_generico"] = -1

    try:
        m["tabelas_fora_padrao"] = _conta_tabelas_fora_padrao(doc)
    except Exception:
        m["tabelas_fora_padrao"] = -1

    # ---- ponto zero ----
    try:
        m["ponto_zero_ok"] = int(_ponto_zero_ok(doc))
    except Exception:
        m["ponto_zero_ok"] = -1

    # ---- niveis fixados (Level.IsPinned NAO existe no Revit 2020) ----
    # Heuristica: tenta IsPinned, senao tenta parametro, senao assume 0 fixados
    try:
        nao_fixados = 0
        for n in niveis:
            fixado = getattr(n, "IsPinned", None)
            if fixado is None:
                # tenta via parametro interno
                try:
                    p = n.get_Parameter(
                        __import__("Autodesk.Revit.DB", fromlist=["BuiltInParameter"])
                        .BuiltInParameter.ELEM_PIN_PARAM)
                    fixado = bool(p and p.HasValue and p.AsInteger() == 1)
                except Exception:
                    fixado = True   # nao conseguiu ler, assume fixado (nao penaliza)
            if not fixado:
                nao_fixados += 1
        m["niveis_nao_fixados"] = nao_fixados
    except Exception:
        m["niveis_nao_fixados"] = -1

    # ---- tags ----
    try:
        m["total_tags"] = _conta_tags(doc)
    except Exception:
        m["total_tags"] = -1

    # ---- avisos ----
    try:
        m["avisos_graves"] = len(doc.GetWarnings())
    except Exception:
        m["avisos_graves"] = -1

    # ---- NOVAS METRICAS (out/2026) ----
    try:
        m["tamanho_arquivo_mb"] = _tamanho_arquivo_mb(doc)
    except Exception:
        m["tamanho_arquivo_mb"] = -1

    try:
        m["qtd_familias_unicas"] = _conta_familias_unicas(doc)
    except Exception:
        m["qtd_familias_unicas"] = -1

    try:
        m["qtd_instancias_familia"] = _conta_instancias_familia(doc)
    except Exception:
        m["qtd_instancias_familia"] = -1

    try:
        m["qtd_vinculos_rvt"] = _conta_vinculos_rvt(doc)
    except Exception:
        m["qtd_vinculos_rvt"] = -1

    try:
        m["qtd_vinculos_cad"] = _conta_vinculos_cad(doc)
    except Exception:
        m["qtd_vinculos_cad"] = -1

    try:
        m["qtd_worksets"] = _conta_worksets(doc)
    except Exception:
        m["qtd_worksets"] = -1

    try:
        m["qtd_vistas_sem_template"] = _conta_vistas_sem_template(doc)
    except Exception:
        m["qtd_vistas_sem_template"] = -1

    return m


# ============================================================
#  HELPERS DE COLETA
# ============================================================
def _conta_dwgs(doc):
    return len(list(FilteredElementCollector(doc).OfClass(ImportInstance)))


def _conta_inplace(doc):
    inst = FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType()
    return sum(1 for i in inst if getattr(i.Symbol, "IsInPlace", False))


def _conta_genericos(doc):
    cat_id = ElementId(int(BuiltInCategory.OST_GenericModel))
    inst = FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType()
    return sum(1 for i in inst if i.Symbol.Family.FamilyCategoryId == cat_id)


def _conta_tabelas_fora_padrao(doc):
    total = 0
    for s in FilteredElementCollector(doc).OfClass(ViewSchedule):
        try:
            if s.IsTitleblockRevisionSchedule:
                continue
            if not RX_TABELA.match(s.Name or u""):
                total += 1
        except Exception:
            pass
    return total


def _conta_tags(doc):
    return len(list(FilteredElementCollector(doc)
                    .OfClass(IndependentTag)
                    .WhereElementIsNotElementType()))


def _conta_carimbos_incompletos(folhas):
    PARAMS_CARIMBO = [u"Projeto", u"Cliente", u"Revisao", u"Data",
                      u"Escala", u"Prancha"]
    incompletas = 0
    for f in folhas:
        for p in PARAMS_CARIMBO:
            try:
                par = f.LookupParameter(p)
                if par is None or not par.HasValue or not (par.AsString() or u"").strip():
                    incompletas += 1
                    break
            except Exception:
                pass
    return incompletas


def _ponto_zero_ok(doc):
    for loc in FilteredElementCollector(doc).OfClass(ProjectLocation):
        try:
            p = loc.Point
            if abs(p.X) < 0.01 and abs(p.Y) < 0.01 and abs(p.Z) < 0.01:
                return True
        except Exception:
            pass
    return False


def _conta_elementos_voados(doc):
    """Conta FamilyInstances REALMENTE fora do envelope do modelo.

    Heuristica em 3 camadas:
      1) Z: elemento cuja bbox INTEIRA esta acima do ultimo nivel
         (+ 1.5x altura de um pe-direito) ou abaixo do primeiro nivel (-2m)
      2) X: elemento cuja bbox INTEIRA esta fora do bbox X do modelo (margem 20%)
      3) Y: idem para Y

    Nao conta elementos dentro da faixa de niveis (mesmo que estejam
    fora do prumo, o Revit nao modela assim com frequencia).
    """
    niveis = list(FilteredElementCollector(doc).OfClass(Level))
    if not niveis:
        return 0

    zs_niveis = sorted([n.Elevation for n in niveis])
    z_min_nivel = zs_niveis[0]
    z_max_nivel = zs_niveis[-1]

    # altura media de um pe-direito
    if len(zs_niveis) >= 2:
        altura_pav = (z_max_nivel - z_min_nivel) / float(len(zs_niveis) - 1)
    else:
        altura_pav = 3.0

    z_limite_baixo = z_min_nivel - 2.0               # fundacao, baldrame
    z_limite_alto = z_max_nivel + altura_pav * 1.5   # caixa d'agua, cobertura

    # coleta todos os bboxes uma vez (evita refazer collector)
    todos = []
    for fi in FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType():
        try:
            bb = fi.get_BoundingBox(None)
            if bb is not None:
                todos.append(bb)
        except Exception:
            pass

    if not todos:
        return 0

    # envelope horizontal do modelo
    xs_min = min(b.Min.X for b in todos)
    xs_max = max(b.Max.X for b in todos)
    ys_min = min(b.Min.Y for b in todos)
    ys_max = max(b.Max.Y for b in todos)

    # margem horizontal: 20% do tamanho
    dx = (xs_max - xs_min) * 0.20
    dy = (ys_max - ys_min) * 0.20

    x_lim_baixo = xs_min - dx
    x_lim_alto = xs_max + dx
    y_lim_baixo = ys_min - dy
    y_lim_alto = ys_max + dy

    voando = 0
    for bb in todos:
        # ---- Z: bbox INTEIRA fora do envelope vertical ----
        if bb.Min.Z > z_limite_alto:
            voando += 1
            continue
        if bb.Max.Z < z_limite_baixo:
            voando += 1
            continue

        # ---- X: bbox INTEIRA fora ----
        if bb.Min.X > x_lim_alto or bb.Max.X < x_lim_baixo:
            voando += 1
            continue

        # ---- Y: bbox INTEIRA fora ----
        if bb.Min.Y > y_lim_alto or bb.Max.Y < y_lim_baixo:
            voando += 1
            continue

    return voando
# ============================================================
#  NOVOS HELPERS (adicionados em out/2026)
# ============================================================

def _tamanho_arquivo_mb(doc):
    """Tamanho do .rvt em disco, em MB. -1 se falhar."""
    try:
        p = doc.PathName
        if not p or not os.path.isfile(p):
            return -1
        return round(os.path.getsize(p) / (1024.0 * 1024.0), 2)
    except Exception:
        return -1


def _conta_familias_unicas(doc):
    """Familias distintas (por nome) carregadas no projeto."""
    try:
        nomes = set()
        for s in FilteredElementCollector(doc).OfClass(FamilySymbol):
            try:
                nome = s.Family.Name
                if nome:
                    nomes.add(nome)
            except Exception:
                pass
        return len(nomes)
    except Exception:
        return -1


def _conta_instancias_familia(doc):
    """Total de FamilyInstance no modelo (elementos de familia colocados)."""
    try:
        return len(list(
            FilteredElementCollector(doc)
            .OfClass(FamilyInstance)
            .WhereElementIsNotElementType()
        ))
    except Exception:
        return -1


def _conta_vinculos_rvt(doc):
    """Vinculos RVT (RevitLinkInstance). Diferente de ImportInstance."""
    try:
        return len(list(
            FilteredElementCollector(doc).OfClass(RevitLinkInstance)
        ))
    except Exception:
        return -1


def _conta_vinculos_cad(doc):
    """CADs vinculados (link), nao importados. Diferente de ImportInstance."""
    try:
        return len(list(
            FilteredElementCollector(doc).OfClass(CADLinkType)
        ))
    except Exception:
        return -1


def _conta_worksets(doc):
    """Total de worksets. 0 se nao for workshared. -1 se falhar.
    Nota: usa .Count do ICollection .NET, nao list() (falha no IronPython 2.7)."""
    try:
        if not doc.IsWorkshared:
            return 0
        tabela = doc.GetWorksetTable()
        if tabela is None:
            return -1
        ids = tabela.GetWorksetIds()
        if ids is None:
            return -1
        # .NET ICollection<T> expoe .Count como propriedade
        try:
            return int(ids.Count)
        except Exception:
            # fallback: itera manualmente
            n = 0
            for _ in ids:
                n += 1
            return n
    except Exception:
        return -1


def _conta_vistas_sem_template(doc):
    """Vistas de planta/corte/fachada/3D sem view template aplicado."""
    try:
        tipos_alvo = (ViewType.FloorPlan, ViewType.CeilingPlan,
                      ViewType.Section, ViewType.Elevation,
                      ViewType.ThreeD)
        total = 0
        for v in FilteredElementCollector(doc).OfClass(View):
            try:
                if v.ViewType not in tipos_alvo:
                    continue
                # ViewTemplateId == -1 significa "sem template"
                if v.ViewTemplateId.IntegerValue == -1:
                    total += 1
            except Exception:
                pass
        return total
    except Exception:
        return -1
# ============================================================
#  LISTA DE MODELOS A PROCESSAR
# ============================================================
def _get_models_to_open():
    """pyRevit CLI injeta __models__ com a lista de .rvt a processar.
    Na GUI, cai no fallback: processa o doc ativo."""
    try:
        models = __models__
        if models:
            return list(models)
    except NameError:
        pass

    # fallback GUI: doc ativo
    try:
        if revit.doc is not None:
            return [revit.doc.PathName]
    except Exception:
        pass

    return []


# ============================================================
#  MAIN
# ============================================================
def main():
    """Abre cada modelo de __models__ e extrai metricas em JSON."""
    debug_path = os.path.join(PASTA_SAIDA, u"_debug_script.log")

    if not os.path.isdir(PASTA_SAIDA):
        try:
            os.makedirs(PASTA_SAIDA)
        except Exception as e:
            forms.alert(u"Nao foi possivel criar a pasta:\n{}\n\n{}".format(PASTA_SAIDA, e))
            return

    with codecs.open(debug_path, "a", "utf-8") as dbg:
        dbg.write(u"\n==== INICIO ====\n")
        try:
            dbg.write(u"Revit versao: {}\n".format(
                __revit__.Application.VersionNumber))
        except Exception:
            pass

        modelos = _get_models_to_open()
        dbg.write(u"Modelos recebidos: {}\n".format(len(modelos)))
        for m in modelos:
            dbg.write(u"  - {}\n".format(m))

        if not modelos:
            dbg.write(u"[ERRO] __models__ vazio e nenhum doc ativo.\n")
            forms.alert(u"Nenhum modelo recebido em __models__.")
            return

        try:
            app = __revit__.Application
        except Exception as e:
            dbg.write(u"[ERRO] __revit__.Application indisponivel: {}\n".format(e))
            forms.alert(u"__revit__.Application indisponivel.")
            return

        for modelo_path in modelos:
            dbg.write(u"\n--- Processando: {} ---\n".format(modelo_path))

            doc = None
            try:
                doc = app.OpenDocumentFile(modelo_path)
                dbg.write(u"Documento aberto OK.\n")
            except Exception as e:
                dbg.write(u"[ERRO] OpenDocumentFile falhou:\n{}\n".format(
                    traceback.format_exc()))
                continue

            if doc is None:
                dbg.write(u"[ERRO] OpenDocumentFile retornou None.\n")
                continue

            try:
                nome = os.path.basename(doc.PathName)
                metricas = coletar_metricas(doc)
                dbg.write(u"Metricas coletadas: {} campos\n".format(len(metricas)))

                destino = os.path.join(PASTA_SAIDA, u"metricas_{}.json".format(
                    os.path.splitext(nome)[0]))
                with codecs.open(destino, "w", "utf-8") as fp:
                    fp.write(json.dumps({nome: metricas}, indent=2, ensure_ascii=False))
                dbg.write(u"JSON gravado: {}\n".format(destino))

            except Exception as e:
                dbg.write(u"[ERRO] processamento falhou:\n{}\n".format(
                    traceback.format_exc()))
            finally:
                try:
                    doc.Close(False)   # False = nao salva
                    dbg.write(u"Documento fechado.\n")
                except Exception as e:
                    dbg.write(u"[aviso] Close falhou: {}\n".format(e))

        dbg.write(u"\n==== FIM ====\n")


if __name__ == "__main__":
    main()