import os
import sys
import pandas as pd
import zipfile
import re
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm

# =========================================================
#  LAUDO (VISTORIA) - GERADOR DOCX via DocxTemplate
#  Adaptado para rodar localmente e no Render (FastAPI),
#  usando LAUDO_BASE_DIR como "workspace" temporário.
# =========================================================

# Diretório base: onde está este script (local) ou pasta temporária (Render)
BASE_DIR = os.getenv("LAUDO_BASE_DIR", os.path.dirname(os.path.abspath(__file__)))

EXCEL_PATH = os.path.join(BASE_DIR, "Vistoria.xlsx")
TEMPLATE_PATH = os.path.join(BASE_DIR, "Modelo_Vistoria.docx")
OUTPUT_DIR = os.path.join(BASE_DIR, "saida")
os.makedirs(OUTPUT_DIR, exist_ok=True)


def refresh_paths():
    """Recalcula caminhos globais a partir do LAUDO_BASE_DIR (se definido)."""
    global BASE_DIR, EXCEL_PATH, TEMPLATE_PATH, OUTPUT_DIR
    BASE_DIR = os.getenv("LAUDO_BASE_DIR", os.path.dirname(os.path.abspath(__file__)))
    EXCEL_PATH = os.path.join(BASE_DIR, "Vistoria.xlsx")
    TEMPLATE_PATH = os.path.join(BASE_DIR, "Modelo_Vistoria.docx")
    OUTPUT_DIR = os.path.join(BASE_DIR, "saida")
    os.makedirs(OUTPUT_DIR, exist_ok=True)


# ----------------- Helpers básicos ----------------- #

def safe_str(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)) or pd.isna(val):
        return ""
    s = str(val)
    return "" if s.strip().lower() == "nan" else s.strip()


def get_ci(row: pd.Series, target: str) -> str:
    """Busca case-insensitive de uma coluna no Series do pandas."""
    tnorm = target.strip().lower()
    for col in row.index:
        if str(col).strip().lower() == tnorm:
            return safe_str(row[col])
    return ""


def decimal_to_dms(value, is_lat=True) -> str:
    """Converte graus decimais em string DMS."""
    if value is None or pd.isna(value):
        return ""
    value = float(value)
    hemi = ("N" if value >= 0 else "S") if is_lat else ("E" if value >= 0 else "W")
    abs_val = abs(value)
    degrees = int(abs_val)
    minutes_float = (abs_val - degrees) * 60
    minutes = int(minutes_float)
    seconds = (minutes_float - minutes) * 60
    return f"{degrees:02d}°{minutes:02d}'{seconds:04.1f}\"{hemi}"


def _fix_template_tags_inplace(path_docx: str):
    """
    Corrige erros comuns de tags Jinja dentro do DOCX.
    (No seu Modelo_Vistoria.docx existe '{% dfor %}' em vez de '{% endfor %}'.)
    """
    if not os.path.isfile(path_docx):
        return

    tmp = path_docx + ".tmp"
    with zipfile.ZipFile(path_docx, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                xml = data.decode("utf-8", errors="ignore")
                # correções
                xml = xml.replace("{% dfor %}", "{% endfor %}")
                xml = xml.replace("{% dfor%}", "{% endfor %}")
                data = xml.encode("utf-8")
            zout.writestr(item, data)

    os.replace(tmp, path_docx)


def carregar_planilhas():
    refresh_paths()
    xls = pd.ExcelFile(EXCEL_PATH)

    def read(name):
        return pd.read_excel(xls, name) if name in xls.sheet_names else pd.DataFrame()

    return {
        "vistoria": read("Vistoria"),
        "endereco": read("Endereco"),
        "imovel": read("Imovel"),
        "info_est": read("Info_Est"),
        "ambientes": read("Ambientes"),
        "fotos_imovel": read("Fotos_imovel"),
        "foto_ambiente": read("Foto_ambiente"),
    }


def encontrar_imagem(path_str: str):
    """
    Resolve o caminho da imagem a partir da coluna Foto.
    Aceita 'Pasta/arquivo.jpg' ou apenas 'arquivo.jpg'.
    """
    if not isinstance(path_str, str) or not path_str.strip():
        return None

    rel_path = path_str.strip().replace("\\", "/")
    full_path = os.path.join(BASE_DIR, rel_path)
    if os.path.exists(full_path):
        return full_path

    filename = os.path.basename(rel_path)

    # Pastas típicas deste app
    for pasta in ["Fotos_imovel_Images", "Foto_ambiente_Images", "Fotos_ambiente_Images"]:
        teste = os.path.join(BASE_DIR, pasta, filename)
        if os.path.exists(teste):
            return teste

    print(f"[AVISO] Imagem não encontrada: {path_str}")
    return None


def inline_image(doc: DocxTemplate, path: str, width_cm: float):
    """Cria um InlineImage com largura fixa em cm e altura proporcional."""
    if not path:
        return ""
    return InlineImage(doc, path, width=Cm(width_cm))


def _normalizar_bool(val):
    """Transforma True/False em 'Sim'/'Não' quando fizer sentido no Word."""
    if isinstance(val, (bool,)):
        return "Sim" if val else "Não"
    return safe_str(val)


# ----------------- Montagem de fotos ----------------- #

def montar_geral_rows(doc: DocxTemplate, fotos_imovel: pd.DataFrame, id_vistoria: str, start_fig: int = 1):
    """
    Bloco "Relatório fotográfico geral do imóvel" (geral_rows),
    em 2 colunas. Retorna (rows2, next_fig).
    """
    if fotos_imovel is None or fotos_imovel.empty:
        return [], start_fig

    df = fotos_imovel.copy()

    if "ID_Vistoria" in df.columns:
        df = df[df["ID_Vistoria"].astype(str) == str(id_vistoria)]

    if "Incluir_no_Laudo" in df.columns:
        df = df[df["Incluir_no_Laudo"] == True]

    if df.empty:
        return [], start_fig

    ordem_col = "Ordem" if "Ordem" in df.columns else None
    if ordem_col:
        df = df.sort_values(ordem_col)

    registros = []
    fig = start_fig
    for _, r in df.iterrows():
        img_path = encontrar_imagem(safe_str(r.get("Foto", "")))
        img = inline_image(doc, img_path, width_cm=8)
        legenda = safe_str(r.get("Legenda", ""))
        caption = f"Figura {fig} - {legenda}" if legenda else f"Figura {fig}"
        registros.append({"img": img, "caption": caption})
        fig += 1

    rows2 = []
    for i in range(0, len(registros), 2):
        r1 = registros[i]
        r2 = registros[i + 1] if i + 1 < len(registros) else None
        rows2.append({
            "col1_img": r1["img"],
            "col1_caption": r1["caption"],
            "col2_img": r2["img"] if r2 else "",
            "col2_caption": r2["caption"] if r2 else "",
        })
    return rows2, fig


def _montar_legenda_foto_amb(r: pd.Series) -> str:
    """Legenda preferencial: coluna Legenda; fallback: compõe com Registro/Ocorrencia/Local."""
    leg = safe_str(r.get("Legenda", ""))
    if leg:
        return leg

    registro = safe_str(r.get("Registro", ""))
    do_da = safe_str(r.get("do/da", ""))
    ocorr = safe_str(r.get("Ocorrencia", ""))
    no_na = safe_str(r.get("no/na", ""))
    local = safe_str(r.get("Local", ""))

    partes = [registro]
    if do_da:
        partes.append(do_da)
    if ocorr:
        partes.append(ocorr)
    if no_na:
        partes.append(no_na)
    if local:
        partes.append(local)

    return " ".join([p for p in partes if p]).strip()


def montar_ambientes(doc: DocxTemplate, ambientes: pd.DataFrame, foto_ambiente: pd.DataFrame, id_vistoria: str, start_fig: int = 1):
    """
    Monta a lista 'ambientes' no formato exigido pelo modelo:
    cada Ambiente tem 'nome' e 'rows' (2 colunas) com figura + legenda.

    Retorna (ambientes_ctx, next_fig).
    """
    if ambientes is None or ambientes.empty:
        return [], start_fig

    amb = ambientes.copy()
    amb = amb[amb["ID_Vistoria"].astype(str) == str(id_vistoria)].copy()
    if amb.empty:
        return [], start_fig

    # ordena ambientes pela ordem do AppSheet
    if "Ordem" in amb.columns:
        amb = amb.sort_values("Ordem")

    fa = (foto_ambiente.copy() if foto_ambiente is not None else pd.DataFrame())

    # alguns exports do AppSheet deixam ID_Vistoria vazio em Foto_ambiente
    # -> inferimos pelo ID_ambiente
    if not fa.empty:
        if "ID_Vistoria" in fa.columns:
            mask = fa["ID_Vistoria"].isna() | fa["ID_Vistoria"].astype(str).str.strip().eq("")
            if mask.any():
                mapa = dict(zip(amb["ID_ambiente"].astype(str), amb["ID_Vistoria"].astype(str)))
                fa.loc[mask, "ID_Vistoria"] = fa.loc[mask, "ID_ambiente"].astype(str).map(mapa)

        # filtra por vistoria e por incluir
        if "ID_Vistoria" in fa.columns:
            fa = fa[fa["ID_Vistoria"].astype(str) == str(id_vistoria)]
        if "Incluir_no_Laudo" in fa.columns:
            fa = fa[fa["Incluir_no_Laudo"] == True]

    fig = start_fig
    ambientes_ctx = []

    for _, a in amb.iterrows():
        id_amb = safe_str(a.get("ID_ambiente", ""))
        nome = safe_str(a.get("Ambiente", "")) or id_amb

        fotos = fa[fa["ID_ambiente"].astype(str) == str(id_amb)] if (not fa.empty and "ID_ambiente" in fa.columns) else pd.DataFrame()

        if not fotos.empty and "Ordem" in fotos.columns:
            fotos = fotos.sort_values("Ordem")

        registros = []
        for _, r in fotos.iterrows():
            img_path = encontrar_imagem(safe_str(r.get("Foto", "")))
            img = inline_image(doc, img_path, width_cm=8)
            legenda = _montar_legenda_foto_amb(r)
            registros.append({"img": img, "fig": fig, "caption": legenda})
            fig += 1

        # agrupa em 2 colunas
        rows2 = []
        for i in range(0, len(registros), 2):
            r1 = registros[i]
            r2 = registros[i + 1] if i + 1 < len(registros) else None
            rows2.append({
                "col1_img": r1["img"],
                "col2_img": r2["img"] if r2 else "",
                "col1_fig": r1["fig"],
                "col2_fig": r2["fig"] if r2 else "",
                "col1_caption": r1["caption"],
                "col2_caption": r2["caption"] if r2 else "",
            })

        ambientes_ctx.append({"nome": nome, "rows": rows2})

    return ambientes_ctx, fig


# ----------------- Função principal ----------------- #

def gerar_laudo(id_vistoria: str):
    refresh_paths()
    sheets = carregar_planilhas()

    vistoria = sheets["vistoria"]
    endereco = sheets["endereco"]
    imovel = sheets["imovel"]
    info = sheets["info_est"]
    ambientes = sheets["ambientes"]
    fotos_imovel = sheets["fotos_imovel"]
    foto_ambiente = sheets["foto_ambiente"]

    row_v = vistoria[vistoria["ID_Vistoria"].astype(str) == str(id_vistoria)]
    if row_v.empty:
        raise ValueError(f"ID_Vistoria '{id_vistoria}' não encontrado na aba Vistoria.")
    row_v = row_v.iloc[0]

    row_end = endereco[endereco["ID_Vistoria"].astype(str) == str(id_vistoria)]
    row_end = row_end.iloc[0] if not row_end.empty else pd.Series(dtype=object)

    row_im = imovel[imovel["ID_Vistoria"].astype(str) == str(id_vistoria)]
    row_im = row_im.iloc[0] if not row_im.empty else pd.Series(dtype=object)

    row_info = info[info["ID_Vistoria"].astype(str) == str(id_vistoria)]
    row_info = row_info.iloc[0] if not row_info.empty else pd.Series(dtype=object)

    # Coordenada decimal -> DMS
    coord_raw = get_ci(row_end, "Coordenada")
    coord_dms = ""
    if coord_raw:
        try:
            parts = str(coord_raw).replace(";", ",").split(",")
            if len(parts) >= 2:
                lat = float(parts[0].strip())
                lon = float(parts[1].strip())
                coord_dms = f"{decimal_to_dms(lat, True)} {decimal_to_dms(lon, False)}"
            else:
                coord_dms = str(coord_raw)
        except Exception:
            coord_dms = str(coord_raw)

    # Corrige tags do template antes de renderizar
    _fix_template_tags_inplace(TEMPLATE_PATH)

    doc = DocxTemplate(TEMPLATE_PATH)

    # Foto da capa
    capa_path = encontrar_imagem(get_ci(row_v, "Foto_da_capa"))
    foto_capa = inline_image(doc, capa_path, width_cm=16) if capa_path else ""

    # Monta blocos fotográficos (numeração contínua)
    geral_rows, next_fig = montar_geral_rows(doc, fotos_imovel, id_vistoria, start_fig=1)
    ambientes_ctx, _ = montar_ambientes(doc, ambientes, foto_ambiente, id_vistoria, start_fig=next_fig)

    # Data da vistoria
    data_v = get_ci(row_im, "Data_da_vistoria")
    if data_v:
        try:
            data_v = pd.to_datetime(data_v).strftime("%d/%m/%Y")
        except Exception:
            pass

    # Contexto (campos do modelo)
    context = {
        # Cabeçalho
        "LAUDO": get_ci(row_v, "Laudo") or "Vistoria",
        "Laudo": get_ci(row_v, "Laudo") or "Vistoria",

        "Contratante": get_ci(row_v, "Contratante"),
        "Referencia": get_ci(row_v, "Referencia"),
        "Representante": get_ci(row_v, "Representante"),
        "Cargo": get_ci(row_v, "Cargo"),
        "Artigo_": get_ci(row_v, "Artigo_") or "A",
        "Artigo": get_ci(row_v, "Artigo") or "a",
        "ART": get_ci(row_v, "ART"),

        "Foto_da_capa": foto_capa,

        # Localização
        "Endereco_imovel": get_ci(row_end, "Endereco_imovel"),
        "Coordenada_DMS": coord_dms,

        # Vistoria
        "Data_da_vistoria": data_v,
        "Acompanhante": get_ci(row_im, "Acompanhante"),

        # Caracterização do imóvel
        "Tipo_imovel": get_ci(row_im, "Tipo_imovel"),
        "pavimentos": get_ci(row_im, "pavimentos"),
        "a_construida": get_ci(row_im, "a_construida"),
        "a_terreno": get_ci(row_im, "a_terreno"),
        "denominacao": get_ci(row_im, "denominacao"),
        "Idade": get_ci(row_im, "Idade"),
        "tipo_idade": get_ci(row_im, "tipo_idade"),
        "funcao": get_ci(row_im, "funcao"),
        "uso": get_ci(row_im, "uso"),
        "padrao": get_ci(row_im, "padrao"),
        "Fechamento": get_ci(row_im, "Fechamento"),
        "Esquadria": get_ci(row_im, "Esquadria"),
        "piso": get_ci(row_im, "piso"),
        "parede": get_ci(row_im, "parede"),
        "Forro": get_ci(row_im, "Forro"),
        "Cobertura": get_ci(row_im, "Cobertura"),
        "Singularidades": get_ci(row_im, "Singularidades"),
        "plano": _normalizar_bool(row_im.get("plano", "")),
        "intervencao": get_ci(row_im, "intervencao"),

        # Estrutura / entorno / utilização (Info_Est)
        "projetista": get_ci(row_info, "projetista"),
        "Construtora": get_ci(row_info, "Construtora"),
        "Estrutura": get_ci(row_info, "Estrutura"),
        "Cobrimento": get_ci(row_info, "Cobrimento"),
        "rev_estrutura": get_ci(row_info, "rev_estrutura"),
        "lajes": get_ci(row_info, "lajes"),
        "tipo_lajes": get_ci(row_info, "tipo_lajes"),
        "secao_pilar": get_ci(row_info, "secao_pilar"),
        "secao_viga": get_ci(row_info, "secao_viga"),
        "junta": _normalizar_bool(row_info.get("junta", "")),
        "junta_estado": get_ci(row_info, "junta_estado"),

        "Fundacao": get_ci(row_info, "Fundacao"),
        "cota_fund": get_ci(row_info, "cota_fund"),
        "solo": get_ci(row_info, "solo"),
        "lencol_freatico": get_ci(row_info, "lencol_freatico"),
        "arvore": _normalizar_bool(row_info.get("arvore", "")),

        "limitrofes": _normalizar_bool(row_info.get("limitrofes", "")),
        "lim_tipo": get_ci(row_info, "lim_tipo"),
        "drenagem": get_ci(row_info, "drenagem"),
        "taludes": _normalizar_bool(row_info.get("taludes", "")),
        "prot_taludes": get_ci(row_info, "prot_taludes"),
        "topografia": get_ci(row_info, "topografia"),
        "microclima": get_ci(row_info, "microclima") or get_ci(row_info, "microclima "),
        "classe_agre": get_ci(row_info, "classe_agre"),

        "incendio": _normalizar_bool(row_info.get("incendio", "")),
        "esp_incendio": get_ci(row_info, "esp_incendio"),
        "agressivo": _normalizar_bool(row_info.get("agressivo", "")),
        "esp_agressivo": get_ci(row_info, "esp_agressivo"),
        "carregamento": _normalizar_bool(row_info.get("carregamento", "")),

        # Blocos fotográficos
        "geral_rows": geral_rows,
        "ambientes": ambientes_ctx,
    }

    doc.render(context)

    contratante = get_ci(row_v, "Contratante") or str(id_vistoria)
    # Sanitiza para nome de arquivo (Windows/Linux)
    contratante_safe = re.sub(r'[\\/:*?"<>|]+', '_', contratante).strip()
    contratante_safe = re.sub(r'\s+', '_', contratante_safe)
    contratante_safe = contratante_safe.strip('._-') or str(id_vistoria)

    out_name = f"Laudo_{contratante_safe}.docx"
    out_path = os.path.join(OUTPUT_DIR, out_name)
    doc.save(out_path)

    print(f"[OK] Laudo gerado em: {out_path}")
    return out_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python gerar_laudo.py <ID_VISTORIA>")
        sys.exit(1)
    gerar_laudo(sys.argv[1])
