"""Módulo compartilhado de geração de relatórios de share (LATAM e GOL).

Produz matrizes de cálculo por canal (CONS, CORP, GOV), visualização na web,
planilhas Excel formatadas (idênticas ao modelo GOL) e relatórios executivos em PDF.
"""
from __future__ import annotations

from pathlib import Path
import json
import re
from typing import Any
import pandas as pd
import numpy as np

import gol_legacy as legacy
from processor import read_csv_robust

CIASS = ("JJ", "G3", "AD")
REPORTS = ("CONS", "CORP", "GOV")
FOOTER_NOTE_DEFAULT = "CONS exclui ART, MAX, -CORP e -GOV. CORP usa -CORP. GOV usa -GOV."
MONTHS_PT = ("", "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
             "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO")


def read_any_table(path: Path) -> pd.DataFrame:
    """Lê CSV (com detecção de separador e encoding) ou XLSX."""
    suffix = path.suffix.lower()
    if suffix == ".csv":
        df, _ = read_csv_robust(path)
        return df
    if suffix in (".xlsx", ".xls"):
        return pd.read_excel(path, dtype=str).fillna("")
    raise ValueError(f"Formato de arquivo não suportado: {suffix}. Envie CSV ou XLSX.")


def normalize_colname(name: Any) -> str:
    """Normaliza nome de coluna para maiúsculas sem caracteres especiais."""
    s = str(name).strip().upper()
    return re.sub(r"[^A-Z0-9]+", "_", s).strip("_")


def normalize_base_for_share(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza nomes e formatos essenciais de qualquer base (LATAM ou GOL)."""
    out = df.copy()
    col_map = {normalize_colname(c): c for c in out.columns}

    def resolve(candidates: list[str]) -> str | None:
        for cand in candidates:
            c_norm = normalize_colname(cand)
            if c_norm in col_map:
                return col_map[c_norm]
        return None

    # Mapeamento de colunas canônicas
    cia_col = resolve(["CIA_BILHETE", "CIA", "COMPANHIA", "CIA_AEREA", "EMPRESA_AEREA"])
    valor_col = resolve(["VALOR_BRL", "VALOR", "TARIFA_BRL_S", "TARIFA_BRL", "TARIFA", "VALOR_TOTAL"])
    base_col = resolve(["BASE_VENDA", "BASE", "BASE VENDA", "BASE_DE_VENDA"])
    cupom_col = resolve(["CUPOM", "QTD_CUPOM", "QTD_CUPONS", "NUM_CUPOM"])
    tipo_col = resolve(["TIPO", "TIPO_VENDA", "TIPO_REGISTRO"])
    emissao_col = resolve(["DATA_EMISSAO", "DATA_EMI", "EMISSAO", "DT_EMISSAO", "DATA EMISSAO"])
    embarque_col = resolve(["DATA_EMBARQUE", "DATA_EMB", "EMBARQUE", "DT_EMBARQUE", "DATA EMBARQUE"])
    bilhete_col = resolve(["BILHETE", "NUMERO_BILHETE", "TICKET", "NUM_BILHETE"])
    gov_col = resolve(["GOV", "GOVERNO"])
    cons2_col = resolve(["CONSOLIDADA_2", "CONSOLIDADA 2", "CONSOLIDADA"])

    if not cia_col:
        raise ValueError("Coluna de companhia aérea (CIA / CIA_BILHETE) não foi encontrada.")
    if not valor_col:
        raise ValueError("Coluna de valor (VALOR_BRL / VALOR / TARIFA) não foi encontrada.")

    # Padroniza nomes
    rename_dict = {
        cia_col: "CIA_BILHETE",
        valor_col: "VALOR_BRL",
    }
    if base_col:
        rename_dict[base_col] = "BASE_VENDA"
    if cupom_col:
        rename_dict[cupom_col] = "CUPOM"
    if tipo_col:
        rename_dict[tipo_col] = "TIPO"
    if emissao_col:
        rename_dict[emissao_col] = "DATA_EMISSAO"
    if embarque_col:
        rename_dict[embarque_col] = "DATA_EMBARQUE"
    if bilhete_col:
        rename_dict[bilhete_col] = "BILHETE"

    out = out.rename(columns=rename_dict)

    # Garantir existência de colunas mínimas
    if "BASE_VENDA" not in out.columns:
        out["BASE_VENDA"] = "CGB"
    if "CUPOM" not in out.columns:
        out["CUPOM"] = "1"
    if "TIPO" not in out.columns:
        out["TIPO"] = "N"
    if "DATA_EMISSAO" not in out.columns:
        out["DATA_EMISSAO"] = ""
    if "DATA_EMBARQUE" not in out.columns:
        out["DATA_EMBARQUE"] = ""
    if "BILHETE" not in out.columns:
        out["BILHETE"] = [f"TKT_{i}" for i in range(len(out))]

    # Limpeza e tipagem
    out["CIA_BILHETE"] = out["CIA_BILHETE"].fillna("").astype(str).str.strip().str.upper()
    out["BASE_VENDA"] = out["BASE_VENDA"].fillna("").astype(str).str.strip().str.upper()
    out["CUPOM"] = out["CUPOM"].fillna("").astype(str).str.strip().str.replace(r"\.0+$", "", regex=True)
    out["TIPO"] = out["TIPO"].fillna("").astype(str).str.strip().str.upper()

    # Se a base não possuir sufixos -CORP e -GOV, utiliza coluna GOV e CONSOLIDADA como apoio
    has_corp_gov_suffix = out["BASE_VENDA"].str.endswith(("-CORP", "-GOV")).any()
    if not has_corp_gov_suffix:
        if gov_col and gov_col in out.columns:
            gov_mask = out[gov_col].fillna("").astype(str).str.strip().str.upper().isin(["SIM", "S", "1", "TRUE"])
            out.loc[gov_mask, "BASE_VENDA"] = out.loc[gov_mask, "BASE_VENDA"] + "-GOV"
        if cons2_col and cons2_col in out.columns:
            corp_mask = out[cons2_col].fillna("").astype(str).str.strip().str.upper().str.contains("CORP", regex=False)
            # Apenas se não for GOV
            corp_mask &= ~out["BASE_VENDA"].str.endswith("-GOV")
            out.loc[corp_mask, "BASE_VENDA"] = out.loc[corp_mask, "BASE_VENDA"] + "-CORP"

    return out


def extract_filter_options(df: pd.DataFrame) -> dict:
    """Extrai opções dinâmicas para os filtros da tela."""
    df_clean = normalize_base_for_share(df)
    dates = pd.to_datetime(df_clean["DATA_EMISSAO"], errors="coerce", dayfirst=True).dropna()
    bases = df_clean["BASE_VENDA"]
    upper = bases.str.upper()

    cupons = sorted(set(df_clean["CUPOM"]) - {""})
    tipos = sorted(set(df_clean["TIPO"]) - {""})

    return {
        "cupons": cupons if cupons else ["1"],
        "tipos": tipos if tipos else ["N"],
        "bases": {
            "CONS": sorted(set(bases.loc[~upper.isin(["ART", "MAX"]) & ~upper.str.endswith(("-CORP", "-GOV"))]) - {""}),
            "CORP": sorted(set(bases.loc[upper.str.endswith("-CORP")]) - {""}),
            "GOV": sorted(set(bases.loc[upper.str.endswith("-GOV")]) - {""}),
        },
        "date_start": dates.min().strftime("%Y-%m-%d") if not dates.empty else "",
        "date_end": dates.max().strftime("%Y-%m-%d") if not dates.empty else "",
    }


def default_filters_for_share(df: pd.DataFrame) -> dict:
    """Retorna filtros padrão do negócio: CUPOM 1, TIPO N, sem WHITE LABEL e período completo."""
    options = extract_filter_options(df)
    return {
        "cupons": ["1"] if "1" in options["cupons"] else options["cupons"],
        "tipos": ["N"] if "N" in options["tipos"] else options["tipos"],
        "bases": {k: [v for v in values if v.upper() != "WHITE LABEL"]
                  for k, values in options["bases"].items()},
        "date_start": options["date_start"],
        "date_end": options["date_end"],
    }


def filter_share_dataframe(df: pd.DataFrame, filters: dict | None) -> pd.DataFrame:
    """Aplica os filtros de cupom, tipo e período de emissão."""
    out = df.copy()
    if filters is None:
        filters = {"cupons": ["1"], "tipos": ["N"]}

    mask = pd.Series(True, index=out.index)
    if filters.get("cupons"):
        mask &= out["CUPOM"].isin([str(x) for x in filters["cupons"]])
    if filters.get("tipos"):
        mask &= out["TIPO"].isin([str(x).upper() for x in filters["tipos"]])

    if filters.get("date_start") or filters.get("date_end"):
        dates = pd.to_datetime(out["DATA_EMISSAO"], errors="coerce", dayfirst=True)
        if filters.get("date_start"):
            mask &= dates.ge(pd.Timestamp(filters["date_start"]))
        if filters.get("date_end"):
            mask &= dates.le(pd.Timestamp(filters["date_end"]) + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1))

    return out.loc[mask].copy()


def report_kind_rows(df: pd.DataFrame, kind: str,
                      advp: tuple[int | None, int | None] | None = None,
                      filters: dict | None = None) -> pd.DataFrame:
    """Filtra as linhas do DataFrame para um tipo específico de relatório (CONS, CORP, GOV)."""
    base = df["BASE_VENDA"].fillna("").astype(str).str.strip().str.upper()
    if kind == "CONS":
        mask = ~base.isin(["ART", "MAX"]) & ~base.str.endswith(("-CORP", "-GOV"))
    elif kind == "CORP":
        mask = base.str.endswith("-CORP")
    elif kind == "GOV":
        mask = base.str.endswith("-GOV")
    else:
        raise ValueError(f"Tipo de relatório inválido: {kind}")

    selected = filter_share_dataframe(df, filters)
    mask &= df.index.isin(selected.index)

    if filters is not None and "bases" in filters:
        allowed = filters["bases"].get(kind, [])
        if allowed:
            mask &= df["BASE_VENDA"].isin(allowed)

    if advp is not None and "DATA_EMBARQUE" in df and "DATA_EMISSAO" in df:
        dates_emb = pd.to_datetime(df["DATA_EMBARQUE"], errors="coerce", dayfirst=True)
        dates_emi = pd.to_datetime(df["DATA_EMISSAO"], errors="coerce", dayfirst=True)
        days = (dates_emb - dates_emi).dt.days
        lo, hi = advp
        if lo is not None:
            mask &= days.ge(lo)
        if hi is not None:
            mask &= days.le(hi)

    return df.loc[mask].copy()


def calculate_summary(df: pd.DataFrame) -> dict:
    """Calcula total em BRL e share percentual por companhia aérea."""
    val = pd.to_numeric(df["VALOR_BRL"].map(legacy.converter_numero_brasileiro), errors="coerce").fillna(0.0)
    cia = df["CIA_BILHETE"].fillna("").astype(str).str.strip().str.upper()
    vals = {c: float(val.loc[cia.eq(c)].sum()) for c in CIASS}
    total = sum(vals.values())
    shares = {c: (vals[c] / total if total > 0 else 0.0) for c in CIASS}
    return {
        "rows": len(df),
        "total": total,
        "values": vals,
        "shares": shares,
        "gap_pp": ((vals["JJ"] - vals["G3"]) / total * 100 if total > 0 else 0.0),
    }


def period_span_label(df: pd.DataFrame, filters: dict | None = None) -> str:
    """Retorna o rótulo do intervalo de datas formatado (ex.: '01 A 08 OUTUBRO 2026')."""
    selected = filter_share_dataframe(df, filters)
    dates = pd.to_datetime(selected["DATA_EMISSAO"], errors="coerce", dayfirst=True).dropna()
    if dates.empty:
        return "SEM DADOS NO PERÍODO"
    first, last = dates.min(), dates.max()
    if first.year == last.year and first.month == last.month:
        if first.day == last.day:
            return f"{first.day:02d} DE {MONTHS_PT[last.month]} DE {last.year}"
        return f"{first.day:02d} A {last.day:02d} DE {MONTHS_PT[last.month]} DE {last.year}"
    if first.year == last.year:
        return f"{first.day:02d} DE {MONTHS_PT[first.month]} A {last.day:02d} DE {MONTHS_PT[last.month]} DE {last.year}"
    return f"{first.day:02d}/{first.month:02d}/{first.year} A {last.day:02d}/{last.month:02d}/{last.year}"


def calculate_share_matrix_and_layout(df: pd.DataFrame,
                                      groups: dict | None = None,
                                      advp_kinds: set[str] | None = None,
                                      filters: dict | None = None) -> list[dict]:
    """Calcula a matriz de relatórios e organiza no layout das tabelas (conforme imagem).

    Retorna uma lista de dicionários, um para cada relatório (CONS, CORP, GOV),
    onde cada um contém:
      - title: 'RELATÓRIO CONS' (ou CORP, GOV)
      - bands: lista de colunas agrupadas (ex: [{'kind': 'CONS', 'label': 'CONS ADVP GERAL'}])
      - rows: lista de linhas ordenadas por maior valor decrescente (JJ, AD, G3) + 'Total geral'
    """
    scoped = set(advp_kinds) if advp_kinds else set()
    tables = []

    for kind in REPORTS:
        bands_spec = []
        if kind in scoped and groups:
            for name, bounds in groups.items():
                bands_spec.append((name, bounds, f"ADVP {name}"))
            bands_spec.append(("GERAL", None, "ADVP TOTAL"))
        else:
            bands_spec.append(("GERAL", None, f"{kind} ADVP GERAL"))

        band_columns = []
        summaries = {}
        for name, bounds, label in bands_spec:
            rows_band = report_kind_rows(df, kind, advp=bounds, filters=filters)
            summ = calculate_summary(rows_band)
            summaries[name] = summ
            band_columns.append({"name": name, "label": label, "kind": kind})

        # Ordenação das companhias pelo valor total no relatório (GERAL)
        geral_summary = summaries["GERAL"]
        sorted_cias = sorted(CIASS, key=lambda c: geral_summary["values"].get(c, 0.0), reverse=True)

        rows = []
        for cia in (*sorted_cias, "Total geral"):
            cells = []
            for col in band_columns:
                s = summaries[col["name"]]
                if cia == "Total geral":
                    val = s["total"]
                    share = 1.0 if s["total"] > 0 else 0.0
                else:
                    val = s["values"].get(cia, 0.0)
                    share = s["shares"].get(cia, 0.0)
                cells.append({"value": val, "share": share})
            rows.append({"cia": cia, "cells": cells})

        tables.append({
            "title": f"RELATÓRIO {kind}",
            "kind": kind,
            "bands": band_columns,
            "rows": rows,
            "summary_geral": geral_summary,
        })

    return tables


def write_pivot_sheet_to_writer(writer: pd.ExcelWriter,
                               sheet_name: str,
                               tables: list[dict],
                               period_title: str,
                               program_title: str = "RELATÓRIO DE SHARE — BASE AJUSTADA",
                               footer_note: str = FOOTER_NOTE_DEFAULT):
    """Escreve uma aba dinâmica de share em um pd.ExcelWriter aberto."""
    book = writer.book
    ws = book.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = ws
    ws.hide_gridlines(2)

    title_fmt = book.add_format({
        "bold": True, "font_size": 14, "font_name": "Arial", "font_color": "#172033"
    })
    subtitle_fmt = book.add_format({
        "italic": True, "font_size": 9, "font_name": "Arial", "font_color": "#5A6A85"
    })
    table_title_fmt = book.add_format({
        "bold": True, "font_size": 11, "font_name": "Arial", "font_color": "#172033"
    })
    cia_head_fmt = book.add_format({
        "bold": True, "bg_color": "#DCE6F1", "border": 1,
        "align": "center", "valign": "vcenter", "font_name": "Arial", "font_size": 9
    })
    subhead_fmt = book.add_format({
        "bold": True, "bg_color": "#F1F2EE", "border": 1,
        "align": "center", "valign": "vcenter", "font_name": "Arial", "font_size": 9
    })
    band_colors = {
        "CONS": "#ECEAE3",
        "CORP": "#DCE6F1",
        "GOV": "#E9EFE4",
    }
    band_formats = {
        k: book.add_format({
            "bold": True, "bg_color": color, "border": 1,
            "align": "center", "valign": "vcenter", "font_name": "Arial", "font_size": 9
        })
        for k, color in band_colors.items()
    }
    cia_cell_fmt = book.add_format({
        "border": 1, "font_name": "Arial", "font_size": 9, "align": "left"
    })
    cia_total_fmt = book.add_format({
        "bold": True, "bg_color": "#E5ECF2", "border": 1,
        "font_name": "Arial", "font_size": 9, "align": "left"
    })
    money_fmt = book.add_format({
        "num_format": 'R$ #,##0.00', "border": 1, "font_name": "Arial", "font_size": 9, "align": "right"
    })
    money_total_fmt = book.add_format({
        "num_format": 'R$ #,##0.00', "bold": True, "bg_color": "#E5ECF2", "border": 1,
        "font_name": "Arial", "font_size": 9, "align": "right"
    })
    pct_fmt = book.add_format({
        "num_format": "0.00%", "border": 1, "font_name": "Arial", "font_size": 9, "align": "right"
    })
    pct_total_fmt = book.add_format({
        "num_format": "0.00%", "bold": True, "bg_color": "#E5ECF2", "border": 1,
        "font_name": "Arial", "font_size": 9, "align": "right"
    })
    footer_fmt = book.add_format({
        "italic": True, "font_size": 8.5, "font_name": "Arial", "font_color": "#6B7280"
    })

    ws.set_column(0, 0, 22)
    ws.set_column(1, 14, 18)

    ws.write(0, 0, program_title, title_fmt)
    ws.write(1, 0, f"Período: {period_title}", subtitle_fmt)

    current_row = 3
    for table in tables:
        ws.write(current_row, 0, table["title"], table_title_fmt)
        head_row = current_row + 1

        ws.merge_range(head_row, 0, head_row + 1, 0, "CIA_BILHETE", cia_head_fmt)

        for j, band in enumerate(table["bands"]):
            col_idx = 1 + j * 2
            kind = band.get("kind", "CONS")
            band_fmt = band_formats.get(kind, band_formats["CONS"])
            ws.merge_range(head_row, col_idx, head_row, col_idx + 1, band["label"], band_fmt)
            ws.write(head_row + 1, col_idx, "TARIFA", subhead_fmt)
            ws.write(head_row + 1, col_idx + 1, "SHARE", subhead_fmt)

        for i, row in enumerate(table["rows"]):
            y = head_row + 2 + i
            is_total = (row["cia"] == "Total geral")
            ws.write(y, 0, row["cia"], cia_total_fmt if is_total else cia_cell_fmt)
            for j, cell in enumerate(row["cells"]):
                col_idx = 1 + j * 2
                ws.write_number(y, col_idx, cell["value"], money_total_fmt if is_total else money_fmt)
                ws.write_number(y, col_idx + 1, cell["share"], pct_total_fmt if is_total else pct_fmt)

        current_row = head_row + 2 + len(table["rows"]) + 2

    ws.write(current_row, 0, footer_note, footer_fmt)
    return ws


def write_adjusted_share_excel(path: Path,
                               tables: list[dict],
                               period_title: str,
                               program_title: str = "RELATÓRIO DE SHARE — BASE AJUSTADA",
                               df_data: pd.DataFrame | None = None,
                               footer_note: str = FOOTER_NOTE_DEFAULT) -> Path:
    """Gera planilha Excel (.xlsx) com layout idêntico ao modelo da GOL."""
    import xlsxwriter

    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        write_pivot_sheet_to_writer(writer, "RELATORIO_SHARE", tables, period_title, program_title, footer_note)

        # Aba opcional de dados
        if df_data is not None and not df_data.empty:
            sample_cols = [c for c in ["DATA_EMISSAO", "CIA_BILHETE", "BILHETE", "CUPOM", "TIPO", "VALOR_BRL", "BASE_VENDA"]
                           if c in df_data.columns]
            export_df = df_data[sample_cols].copy() if sample_cols else df_data.iloc[:, :15].copy()
            export_df.to_excel(writer, sheet_name="DADOS_ANALISADOS", index=False)

    return path


def write_adjusted_share_pdf(path: Path,
                             program_name: str,
                             tables: list[dict],
                             period_title: str,
                             filters_summary: dict | None = None,
                             stats_summary: dict | None = None,
                             footer_note: str = FOOTER_NOTE_DEFAULT) -> Path:
    """Gera PDF de relatório executivo em formato paisagem (A4) idêntico ao modelo executivo."""
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, KeepTogether

    doc = SimpleDocTemplate(
        str(path),
        pagesize=(297 * mm, 210 * mm),
        leftMargin=14 * mm,
        rightMargin=14 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "ReportTitle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=15,
        leading=18,
        textColor=colors.HexColor("#102A43"),
        alignment=TA_LEFT,
        spaceAfter=3,
    )
    subtitle_style = ParagraphStyle(
        "ReportSubtitle",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#486581"),
        spaceAfter=10,
    )
    table_title_style = ParagraphStyle(
        "TableHeading",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=14,
        textColor=colors.HexColor("#102A43"),
        spaceBefore=8,
        spaceAfter=4,
    )
    footer_style = ParagraphStyle(
        "FooterNote",
        parent=styles["Normal"],
        fontName="Helvetica-Oblique",
        fontSize=8,
        leading=11,
        textColor=colors.HexColor("#627D98"),
        spaceBefore=8,
    )

    def money_str(val: float) -> str:
        return "R$ " + f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    story = [
        Paragraph(f"{program_name} — Relatório de Share (Base Ajustada)", title_style),
        Paragraph(f"<b>Período apurado:</b> {period_title}", subtitle_style),
    ]

    # Caixa de resumo se houver estatísticas
    if stats_summary:
        meta_items = []
        if "total_rows" in stats_summary:
            meta_items.append(["Total de linhas:", f"{int(stats_summary['total_rows']):,}"])
        if "total_tickets" in stats_summary:
            meta_items.append(["Bilhetes únicos:", f"{int(stats_summary['total_tickets']):,}"])
        if "total_value" in stats_summary:
            meta_items.append(["Faturamento BRL:", money_str(stats_summary["total_value"])])
        if filters_summary:
            cupons_str = ", ".join(map(str, filters_summary.get("cupons", ["1"])))
            tipos_str = ", ".join(map(str, filters_summary.get("tipos", ["N"])))
            meta_items.append(["Filtros aplicados:", f"CUPOM {cupons_str} | TIPO {tipos_str}"])

        if meta_items:
            m_table = Table(meta_items, colWidths=[38 * mm, 95 * mm], hAlign="LEFT")
            m_table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F0F4F8")),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D9E2EC")),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTNAME", (1, 0), (1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]))
            story.append(m_table)
            story.append(Spacer(1, 4 * mm))

    # Construção das tabelas
    for table in tables:
        heading_row_0 = ["CIA_BILHETE"]
        heading_row_1 = [""]
        for band in table["bands"]:
            heading_row_0.extend([band["label"], ""])
            heading_row_1.extend(["TARIFA", "SHARE"])

        table_data = [heading_row_0, heading_row_1]

        for r in table["rows"]:
            row_vals = [r["cia"]]
            for cell in r["cells"]:
                row_vals.append(money_str(cell["value"]))
                row_vals.append(f"{cell['share'] * 100:.2f}%".replace(".", ","))
            table_data.append(row_vals)

        available_w = 269 * mm
        num_cols = len(heading_row_0)
        first_w = 40 * mm
        other_w = (available_w - first_w) / (num_cols - 1)
        col_widths = [first_w] + [other_w] * (num_cols - 1)

        t = Table(table_data, colWidths=col_widths, repeatRows=2)
        style_cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DCE6F1")),
            ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#F1F2EE")),
            ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E5ECF2")),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#9AABBC")),
            ("FONTNAME", (0, 0), (-1, 1), "Helvetica-Bold"),
            ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7.5),
            ("ALIGN", (0, 0), (0, -1), "LEFT"),
            ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
            ("ALIGN", (0, 0), (-1, 1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 3.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
            ("SPAN", (0, 0), (0, 1)),
        ]
        # Mesclar as colunas dos blocos superiores
        for b_idx in range(len(table["bands"])):
            c_start = 1 + b_idx * 2
            style_cmds.append(("SPAN", (c_start, 0), (c_start + 1, 0)))

        t.setStyle(TableStyle(style_cmds))
        story.append(KeepTogether([Paragraph(table["title"], table_title_style), t]))
        story.append(Spacer(1, 4 * mm))

    story.append(Paragraph(footer_note, footer_style))
    doc.build(story)
    return path
