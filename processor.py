
from __future__ import annotations

from pathlib import Path
import csv
import io
import json
import re
import shutil
import zipfile
import numpy as np
import pandas as pd
from xml.sax.saxutils import escape as xml_escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    Image as RLImage, PageBreak, KeepTogether
)

WEST_IATAS = {"57516196", "57516197"}
DEFAULT_PRIORITY_IATAS = ["GRU", "POA", "CGB", "FOR", "GIG"]
TARGET_CIAS = ["JJ", "G3", "AD"]


# ============================================================
# LEITURA / UTILITÁRIOS
# ============================================================

def read_csv_robust(path: Path):
    raw = path.read_bytes()
    text = None
    encoding = None
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            text = raw.decode(enc)
            encoding = enc
            break
        except UnicodeDecodeError:
            pass
    if text is None:
        raise ValueError("Não foi possível identificar a codificação do CSV.")

    try:
        dialect = csv.Sniffer().sniff(text[:100_000], delimiters=";,|\t,")
        sep = dialect.delimiter
    except Exception:
        sep = ";"

    df = pd.read_csv(
        io.StringIO(text),
        sep=sep,
        dtype=str,
        keep_default_na=False,
        low_memory=False,
    )
    return df, {"encoding": encoding, "separator": sep}


def norm(s):
    return re.sub(r"[^A-Z0-9]+", "_", str(s).strip().upper()).strip("_")


def find_col(df, candidates, required=True):
    mapping = {norm(c): c for c in df.columns}
    for c in candidates:
        if norm(c) in mapping:
            return mapping[norm(c)]
    if required:
        raise ValueError(f"Coluna não encontrada. Procurei: {candidates}")
    return None


def identify_consolidadas(df):
    cons_cols = [c for c in df.columns if norm(c).startswith("CONSOLIDADA")]
    if len(cons_cols) >= 2:
        return cons_cols[0], cons_cols[1]
    if len(df.columns) >= 21:
        return df.columns[19], df.columns[20]
    raise ValueError("Não foi possível identificar CONSOLIDADA 1 e CONSOLIDADA 2.")


def parse_money(series):
    s = series.astype(str).str.strip().str.replace("R$", "", regex=False).str.replace(" ", "", regex=False)
    has_comma = s.str.contains(",", regex=False)
    out = pd.Series(np.nan, index=s.index, dtype=float)
    out.loc[has_comma] = pd.to_numeric(
        s[has_comma].str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
        errors="coerce",
    )
    out.loc[~has_comma] = pd.to_numeric(s[~has_comma], errors="coerce")
    return out.fillna(0.0)


def parse_date(series):
    return pd.to_datetime(series, errors="coerce", dayfirst=True)


def clean_digits(v):
    return re.sub(r"\D", "", str(v or ""))


def normalize_ticket(v):
    """Retorna somente a parte numérica do bilhete.

    Exemplos:
    N.9570000046590 -> 9570000046590
    95745-1234567   -> 957451234567
    """
    return clean_digits(v)


def blocked_ticket(v):
    t = normalize_ticket(v)
    if not t:
        return False
    if t.startswith(("95744", "95745", "95719")):
        return True
    # 957 seguido de pelo menos 3 zeros: 957000..., 9570000..., etc.
    return re.match(r"^9570{3,}", t) is not None


def iata_ends_2412(v):
    return clean_digits(v).endswith("2412")


def route_codes(v):
    return re.findall(r"[A-Z]{3}", str(v).upper())


def is_direct_route(v):
    p = route_codes(v)
    return len(p) == 2 or (len(p) == 3 and p[0] == p[-1])


def contains_priority_iata(route, priorities):
    return bool(set(route_codes(route)).intersection(set(priorities)))


# ============================================================
# ETAPA 1 — FORMATAÇÃO
# ============================================================

def stage1_format(df):
    """Executa a etapa 1 e retorna bases + auditoria detalhada."""
    df = df.copy()
    original_rows = len(df)
    original_cols = len(df.columns)

    # --------------------------------------------------------
    # 1) GOV: pela regra do processo, GOV é a última coluna.
    #    Remove por nome se existir; se a última coluna normalizar
    #    como GOV, também remove. Registra o que foi feito.
    # --------------------------------------------------------
    gov_removed = False
    gov_col = find_col(df, ["GOV"], required=False)
    if gov_col is not None:
        df = df.drop(columns=[gov_col])
        gov_removed = True
    elif len(df.columns) and norm(df.columns[-1]) == "GOV":
        df = df.iloc[:, :-1].copy()
        gov_removed = True

    cons1, cons2 = identify_consolidadas(df)

    bilhete = find_col(df, ["BILHETE", "NUMERO_BILHETE", "TICKET"])
    cupom = find_col(df, ["CUPOM", "QTD_CUPOM", "QTD_CUPONS"])
    data_reserva = find_col(df, ["DATA_RESERVA", "DATA_RES", "RESERVA"])
    data_emissao = find_col(df, ["DATA_EMISSAO", "DATA_EMI", "EMISSAO"])
    grupo_empresa = find_col(df, ["GRUPO_EMPRESA", "GRUPO EMPRESA"], required=False)
    empresa = find_col(df, ["EMPRESA"], required=False)
    n2 = find_col(df, ["N2"], required=False)
    grupo_cons = find_col(df, ["GRUPO_CONSOLIDADA", "GRUPO CONSOLIDADA"], required=False)
    base_col = find_col(df, ["BASE_VENDA", "BASE", "BASE VENDA"], required=False)
    cia = find_col(df, ["CIA", "CIA_BILHETE", "COMPANHIA"])
    iata = find_col(df, ["IATA", "IATA_AGENCIA", "IATA AGENCIA"])
    tipo = find_col(df, ["TIPO"])
    valor = find_col(df, ["VALOR_BRL", "TARIFA_BRL_S", "TARIFA_BRL"])
    trecho = find_col(df, ["TRECHO", "ROTA", "ROUTE"], required=False)

    excluded_parts = []
    exclusion_counts = {}

    def exclude(mask, reason):
        nonlocal df
        mask = pd.Series(mask, index=df.index).fillna(False).astype(bool)
        qty = int(mask.sum())
        exclusion_counts[reason] = exclusion_counts.get(reason, 0) + qty
        if qty:
            part = df.loc[mask].copy()
            part["MOTIVO_EXCLUSAO"] = reason
            excluded_parts.append(part)
            df = df.loc[~mask].copy()

    # --------------------------------------------------------
    # 2) Bilhetes bloqueados.
    #    IMPORTANTE: o arquivo traz valores como N.957000..., então
    #    normalizamos para dígitos antes de testar o prefixo.
    # --------------------------------------------------------
    ticket_norm = df[bilhete].map(normalize_ticket)
    mask_95744 = ticket_norm.str.startswith("95744", na=False)
    mask_95745 = ticket_norm.str.startswith("95745", na=False)
    mask_95719 = ticket_norm.str.startswith("95719", na=False)
    mask_957zeros = ticket_norm.str.match(r"^9570{3,}", na=False)

    # Separado por motivo para auditoria clara
    exclude(mask_95744, "BILHETE INICIA 95744")
    ticket_norm = df[bilhete].map(normalize_ticket)
    exclude(ticket_norm.str.startswith("95745", na=False), "BILHETE INICIA 95745")
    ticket_norm = df[bilhete].map(normalize_ticket)
    exclude(ticket_norm.str.startswith("95719", na=False), "BILHETE INICIA 95719")
    ticket_norm = df[bilhete].map(normalize_ticket)
    exclude(ticket_norm.str.match(r"^9570{3,}", na=False), "BILHETE 957 + 3 OU MAIS ZEROS")

    # --------------------------------------------------------
    # 3) CONSOLIDADA 2 iniciando com CONF SR.
    #    Normaliza espaços para pegar ' CONF   SR ...'.
    # --------------------------------------------------------
    cons2_norm = (
        df[cons2].astype(str)
        .str.strip().str.upper()
        .str.replace(r"\s+", " ", regex=True)
    )
    exclude(cons2_norm.str.match(r"^CONF\s+SR", na=False), "CONSOLIDADA 2 INICIA COM CONF SR")

    # --------------------------------------------------------
    # 4) CUPOM = 0 (aceita 0, 0.0, 0,00 etc.).
    # --------------------------------------------------------
    cup_text = df[cupom].astype(str).str.strip().str.replace(",", ".", regex=False)
    cup_num = pd.to_numeric(cup_text, errors="coerce")
    exclude(cup_num.eq(0), "CUPOM = 0")

    # --------------------------------------------------------
    # 5) DATA_RESERVA vazia => DATA_EMISSAO.
    #    Também trata '', NaN textual, None/NaT.
    # --------------------------------------------------------
    reserva_txt = df[data_reserva].astype(str).str.strip().str.upper()
    empty_res = reserva_txt.isin(["", "NAN", "NONE", "NAT", "NULL"])
    data_reserva_filled = int(empty_res.sum())
    df.loc[empty_res, data_reserva] = df.loc[empty_res, data_emissao]

    # --------------------------------------------------------
    # 6) GRUPO_EMPRESA = EMPRESA; se EMPRESA não existir, usa N2.
    # --------------------------------------------------------
    grupo_empresa_source = None
    if grupo_empresa is not None:
        if empresa is not None:
            df[grupo_empresa] = df[empresa]
            grupo_empresa_source = empresa
        elif n2 is not None:
            df[grupo_empresa] = df[n2]
            grupo_empresa_source = n2

    # --------------------------------------------------------
    # 7) GRUPO_CONSOLIDADA = CONSOLIDADA 1.
    # --------------------------------------------------------
    grupo_consolidada_updated = 0
    if grupo_cons is not None:
        grupo_consolidada_updated = len(df)
        df[grupo_cons] = df[cons1]

    # --------------------------------------------------------
    # 8) ART / MAX.
    # --------------------------------------------------------
    artmax_log = []
    if base_col is not None:
        # recalcula a cada iteração porque df é alterado em lugar
        for base_name in ("ART", "MAX"):
            base_upper = df[base_col].astype(str).str.strip().str.upper()
            cia_upper = df[cia].astype(str).str.strip().str.upper()
            mask_base = base_upper.eq(base_name)
            mask_jj = mask_base & cia_upper.eq("JJ")
            jj_count = int(mask_jj.sum())
            total_base = int(mask_base.sum())

            if jj_count == 0:
                artmax_log.append({
                    "base": base_name,
                    "total_base": total_base,
                    "jj": 0,
                    "jj_2412": 0,
                    "alterados": 0,
                    "regra": "SEM JJ - SEM ALTERAÇÃO",
                })
                continue

            flags_2412 = df.loc[mask_jj, iata].map(iata_ends_2412)
            jj_2412_count = int(flags_2412.sum())
            all_2412 = jj_2412_count == jj_count

            if all_2412:
                changed = int(mask_base.sum())
                df.loc[mask_base, cons1] = "958"
                df.loc[mask_base, cons2] = "WHITE LABEL"
                regra = "BASE INTEIRA => 958 / WHITE LABEL"
            else:
                mask_jj_2412 = mask_jj & df[iata].map(iata_ends_2412)
                changed = int(mask_jj_2412.sum())
                df.loc[mask_jj_2412, cons1] = "43"
                df.loc[mask_jj_2412, cons2] = "ARARAUNA TURISMO"
                regra = "SOMENTE JJ FINAL 2412 => 43 / ARARAUNA TURISMO"

            artmax_log.append({
                "base": base_name,
                "total_base": total_base,
                "jj": jj_count,
                "jj_2412": jj_2412_count,
                "alterados": changed,
                "regra": regra,
            })

    formatted_full = df.copy()

    # --------------------------------------------------------
    # 9) WEST: gera arquivos e retira os IATAs da base principal.
    # --------------------------------------------------------
    iata_digits = df[iata].map(clean_digits)
    west_mask = iata_digits.isin(WEST_IATAS)
    west = df.loc[west_mask].copy()
    main_after_west = df.loc[~west_mask].copy()

    excluded_format = (
        pd.concat(excluded_parts, ignore_index=True)
        if excluded_parts
        else pd.DataFrame(columns=list(formatted_full.columns) + ["MOTIVO_EXCLUSAO"])
    )

    columns = {
        "cons1": cons1,
        "cons2": cons2,
        "bilhete": bilhete,
        "cupom": cupom,
        "data_reserva": data_reserva,
        "data_emissao": data_emissao,
        "grupo_empresa": grupo_empresa,
        "empresa": empresa if empresa is not None else n2,
        "grupo_consolidada": grupo_cons,
        "base": base_col,
        "cia": cia,
        "iata": iata,
        "tipo": tipo,
        "valor": valor,
        "trecho": trecho,
    }

    stats = {
        "original_rows": original_rows,
        "original_cols": original_cols,
        "gov_removed": gov_removed,
        "formatted_rows": len(formatted_full),
        "excluded_format_rows": len(excluded_format),
        "west_rows": len(west),
        "main_rows": len(main_after_west),
        "columns_after_gov": len(formatted_full.columns),
        "data_reserva_filled": data_reserva_filled,
        "grupo_empresa_source": grupo_empresa_source,
        "grupo_consolidada_updated": grupo_consolidada_updated,
        "exclusion_counts": exclusion_counts,
        "artmax_log": artmax_log,
        "detected_columns": columns,
    }

    return {
        "formatted_full": formatted_full,
        "main_after_west": main_after_west,
        "excluded_format": excluded_format,
        "west": west,
        "columns": columns,
        "stats": stats,
    }


# ============================================================
# FILTROS VISÍVEIS DO USUÁRIO
# ============================================================

def get_filter_options(df, columns):
    c = columns
    cupons = sorted({
        str(x).strip()
        for x in df[c["cupom"]].astype(str)
        if str(x).strip() != ""
    })
    tipos = sorted({
        str(x).strip().upper()
        for x in df[c["tipo"]].astype(str)
        if str(x).strip() != ""
    })
    consolidadas = sorted({
        str(x).strip()
        for x in df[c["cons2"]].astype(str)
        if str(x).strip() != ""
    })

    dates = parse_date(df[c["data_emissao"]]).dropna()
    date_start = dates.min().strftime("%Y-%m-%d") if not dates.empty else ""
    date_end = dates.max().strftime("%Y-%m-%d") if not dates.empty else ""

    return {
        "cupons": cupons,
        "tipos": tipos,
        "consolidadas": consolidadas,
        "date_start": date_start,
        "date_end": date_end,
    }


def apply_user_filters(df, columns, filters):
    c = columns
    out = df.copy()

    cup = out[c["cupom"]].astype(str).str.strip()
    tipo = out[c["tipo"]].astype(str).str.strip().str.upper()
    cons2 = out[c["cons2"]].astype(str).str.strip()
    data = parse_date(out[c["data_emissao"]])

    if filters.get("cupons"):
        out = out.loc[cup.isin([str(x) for x in filters["cupons"]])].copy()
        tipo = out[c["tipo"]].astype(str).str.strip().str.upper()
        cons2 = out[c["cons2"]].astype(str).str.strip()
        data = parse_date(out[c["data_emissao"]])

    if filters.get("tipos"):
        out = out.loc[tipo.isin([str(x).upper() for x in filters["tipos"]])].copy()
        cons2 = out[c["cons2"]].astype(str).str.strip()
        data = parse_date(out[c["data_emissao"]])

    if filters.get("consolidadas"):
        out = out.loc[cons2.isin(filters["consolidadas"])].copy()
        data = parse_date(out[c["data_emissao"]])

    if filters.get("date_start"):
        start = pd.to_datetime(filters["date_start"], errors="coerce")
        if not pd.isna(start):
            out = out.loc[data >= start].copy()
            data = parse_date(out[c["data_emissao"]])

    if filters.get("date_end"):
        end = pd.to_datetime(filters["date_end"], errors="coerce")
        if not pd.isna(end):
            out = out.loc[data <= end + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1)].copy()

    return out


def analysis_view(df, columns):
    c = columns
    out = df.copy()
    out["_CIA"] = out[c["cia"]].astype(str).str.strip().str.upper()
    out["_VALOR"] = parse_money(out[c["valor"]])
    out["_DATA"] = parse_date(out[c["data_emissao"]]).dt.normalize()
    out["_TRECHO"] = (
        out[c["trecho"]].astype(str).str.strip().str.upper()
        if c["trecho"] else ""
    )
    out["_BILHETE"] = out[c["bilhete"]].astype(str).str.strip()
    return out.loc[out["_CIA"].isin(TARGET_CIAS)].copy()


def shares(df):
    total = float(df["_VALOR"].sum())
    vals = {
        cia: float(df.loc[df["_CIA"].eq(cia), "_VALOR"].sum())
        for cia in TARGET_CIAS
    }
    sh = {cia: (vals[cia] / total if total else 0.0) for cia in TARGET_CIAS}
    return total, vals, sh


def build_share_preview(filtered_df, columns):
    e = analysis_view(filtered_df, columns)
    total, vals, sh = shares(e)
    dates = e["_DATA"].dropna()

    display_columns = [
        columns["data_emissao"],
        columns["bilhete"],
        columns["cupom"],
        columns["cia"],
        columns["iata"],
        columns["trecho"] if columns["trecho"] else columns["cia"],
        columns["tipo"],
        columns["cons1"],
        columns["cons2"],
        columns["valor"],
    ]
    # remove duplicados preservando ordem
    display_columns = list(dict.fromkeys([c for c in display_columns if c]))

    return {
        "period_start": dates.min().strftime("%d/%m/%Y") if not dates.empty else "-",
        "period_end": dates.max().strftime("%d/%m/%Y") if not dates.empty else "-",
        "jj": sh["JJ"],
        "g3": sh["G3"],
        "ad": sh["AD"],
        "total": total,
        "rows": len(e),
        "unique_tickets": int(e["_BILHETE"].nunique()),
        "display_columns": display_columns,
    }


# ============================================================
# TABELAS DE AUDITORIA
# ============================================================

def daily_table(df):
    rows = []
    for dt, g in df.groupby("_DATA", dropna=False):
        total, vals, sh = shares(g)
        rows.append({
            "DATA": dt,
            "TOTAL": total,
            "JJ_BRL": vals["JJ"], "JJ_SHARE": sh["JJ"],
            "G3_BRL": vals["G3"], "G3_SHARE": sh["G3"],
            "AD_BRL": vals["AD"], "AD_SHARE": sh["AD"],
        })
    return pd.DataFrame(rows)


def route_table(df):
    rows = []
    for route, g in df.groupby("_TRECHO", dropna=False):
        total, vals, sh = shares(g)
        rows.append({
            "TRECHO": route,
            "DIRETO": "SIM" if is_direct_route(route) else "NÃO",
            "TOTAL": total,
            "JJ_BRL": vals["JJ"], "JJ_SHARE": sh["JJ"],
            "G3_BRL": vals["G3"], "G3_SHARE": sh["G3"],
            "AD_BRL": vals["AD"], "AD_SHARE": sh["AD"],
            "GAP_JJ_G3_PP": (sh["JJ"] - sh["G3"]) * 100,
        })
    return pd.DataFrame(rows)


def priority_weights(params):
    order = params.get("priority_order") or [1, 2, 3, 4]
    pos = {rule: i for i, rule in enumerate(order)}
    return {
        1: 1000 / (1 + pos.get(1, 0)),
        2: 300 / (1 + pos.get(2, 1)),
        3: 220 / (1 + pos.get(3, 2)),
        4: 160 / (1 + pos.get(4, 3)),
    }


def score_state(before, after, params):
    w = priority_weights(params)
    _, _, sh = shares(after)

    # 1) Meta LATAM
    score = abs(sh["JJ"] - params["target_share"]) * w[1]

    # 2) Faixa G3 / AD
    company_pen = 0.0
    for cia, lo, hi in [
        ("G3", params.get("g3_min"), params.get("g3_max")),
        ("AD", params.get("ad_min"), params.get("ad_max")),
    ]:
        if lo is not None:
            company_pen += max(0.0, lo - sh[cia])
        if hi is not None:
            company_pen += max(0.0, sh[cia] - hi)
    score += company_pen * w[2]

    # 3) Preservação diária
    if params.get("daily_control") and params.get("daily_preserve_from") is not None:
        threshold = params["daily_preserve_from"]
        db = daily_table(before).set_index("DATA")
        da = daily_table(after).set_index("DATA")
        p = 0.0
        for dt in db.index:
            sb = float(db.loc[dt, "JJ_SHARE"])
            sa = float(da.loc[dt, "JJ_SHARE"]) if dt in da.index else sb
            if params.get("avoid_high_days") and sb >= threshold:
                p += max(0.0, sa - sb) * 20.0
            else:
                soft_cap = max(threshold, params["target_share"])
                p += max(0.0, sa - soft_cap) * 4.0
        score += p * w[3]

    # 4) Gap por trecho prioritário
    if params.get("route_gap") is not None:
        rt = route_table(after)
        p = 0.0
        priorities = params.get("priority_iatas") or DEFAULT_PRIORITY_IATAS
        for _, r in rt.iterrows():
            if r["DIRETO"] != "SIM":
                continue
            if not contains_priority_iata(r["TRECHO"], priorities):
                continue
            shortfall = params["route_gap"] - (r["JJ_SHARE"] - r["G3_SHARE"])
            if shortfall > 0:
                p += shortfall
        score += p * w[4]

    return score


# ============================================================
# AJUSTE DE SHARE — UNIDADE DE EXCLUSÃO = BILHETE COMPLETO
# ============================================================

def run_share_adjustment(main_df, columns, params):
    filters = params["filters"]

    # Share antes: exatamente os filtros escolhidos na tela
    filtered_before_raw = apply_user_filters(main_df, columns, filters)
    before = analysis_view(filtered_before_raw, columns)

    if before.empty:
        raise ValueError("Os filtros selecionados não retornaram registros para análise.")

    total_b, vals_b, sh_b = shares(before)

    # Meta igual ao share atual (considerando arredondamento normal de tela):
    # não remove nenhum bilhete. Meta abaixo do atual também não é permitida,
    # porque este processo só aumenta o share JJ retirando AD/G3.
    target = params["target_share"]
    tolerance = 0.00005
    if target < sh_b["JJ"] - tolerance:
        raise ValueError(
            f"A meta informada ({target:.2%}) é menor que o share atual da LATAM ({sh_b['JJ']:.2%}). "
            "Este processo só aumenta o share retirando bilhetes AD/G3."
        )

    no_adjustment_needed = abs(target - sh_b["JJ"]) <= tolerance

    # Bilhetes candidatos: somente bilhetes que no universo filtrado
    # tenham AD ou G3. O bilhete completo será removido da base física.
    ticket_info = []
    for ticket, g in before.groupby("_BILHETE"):
        cias = set(g["_CIA"])
        if not cias.intersection({"AD", "G3"}):
            continue

        # Evita escolher bilhete que também tenha JJ no universo analisado.
        if "JJ" in cias:
            continue

        dom_cia = g.groupby("_CIA")["_VALOR"].sum().idxmax()
        dom_date = g.groupby("_DATA")["_VALOR"].sum().idxmax()
        dom_route = g.groupby("_TRECHO")["_VALOR"].sum().idxmax()
        value = float(g["_VALOR"].sum())

        ticket_info.append({
            "ticket": ticket,
            "cia": dom_cia,
            "date": dom_date,
            "route": dom_route,
            "value": value,
        })

    if not ticket_info and not no_adjustment_needed:
        raise ValueError("Não há bilhetes AD/G3 elegíveis para ajuste com os filtros escolhidos.")

    tickets_df = pd.DataFrame(ticket_info)

    daily_before = daily_table(before).set_index("DATA")
    route_before = route_table(before).set_index("TRECHO")
    priorities = params.get("priority_iatas") or DEFAULT_PRIORITY_IATAS

    def rank_ticket(row):
        original_day_share = (
            float(daily_before.loc[row["date"], "JJ_SHARE"])
            if row["date"] in daily_before.index else 0.0
        )

        high_day = (
            params.get("daily_control")
            and params.get("avoid_high_days")
            and params.get("daily_preserve_from") is not None
            and original_day_share >= params["daily_preserve_from"]
        )

        # Preserva G3 por padrão: AD primeiro.
        cia_penalty = 1 if row["cia"] == "G3" else 0

        need_gap = False
        if (
            params.get("route_gap") is not None
            and row["route"] in route_before.index
            and is_direct_route(row["route"])
            and contains_priority_iata(row["route"], priorities)
        ):
            rr = route_before.loc[row["route"]]
            need_gap = (float(rr["JJ_SHARE"]) - float(rr["G3_SHARE"])) < params["route_gap"]

        # Se precisa corrigir gap JJ x G3, remover G3 daquela rota pode ganhar prioridade.
        if need_gap and row["cia"] == "G3":
            cia_penalty = -1

        return (
            1 if high_day else 0,
            cia_penalty,
            0 if need_gap else 1,
            row["value"],
        )

    if not tickets_df.empty:
        tickets_df["_rank"] = tickets_df.apply(rank_ticket, axis=1)
        tickets_df = tickets_df.sort_values("_rank")

    removed_tickets = set()
    current_full = main_df.copy()
    current_filtered = before.copy()
    current_score = score_state(before, current_filtered, params)

    # Para desempenho, testa em ordem heurística.
    # Se a meta já é o share existente, tickets_df fica sem uso e nada é removido.
    for _, cand in tickets_df.iterrows():
        ticket = cand["ticket"]
        if ticket in removed_tickets:
            continue

        # Regra crítica:
        # remove TODAS as linhas/cupons do bilhete da base física.
        trial_full = current_full.loc[
            current_full[columns["bilhete"]].astype(str).str.strip() != str(ticket).strip()
        ].copy()

        # Reaplica os mesmos filtros do usuário após remover o bilhete inteiro.
        trial_filtered_raw = apply_user_filters(trial_full, columns, filters)
        trial_filtered = analysis_view(trial_filtered_raw, columns)

        new_score = score_state(before, trial_filtered, params)
        _, _, sh_cur = shares(current_filtered)
        _, _, sh_new = shares(trial_filtered)

        closer = (
            abs(sh_new["JJ"] - params["target_share"])
            < abs(sh_cur["JJ"] - params["target_share"])
        )
        accept = new_score + 1e-10 < current_score

        # Enquanto ainda está abaixo da meta, aceita avanço controlado.
        if sh_cur["JJ"] < params["target_share"]:
            if sh_new["JJ"] <= params["target_share"] + 0.001:
                accept = accept or closer

        if accept:
            removed_tickets.add(ticket)
            current_full = trial_full
            current_filtered = trial_filtered
            current_score = new_score

        _, _, sh_now = shares(current_filtered)

        # Próximo o bastante da meta e dentro das faixas, encerra.
        company_ok = True
        for cia, lo, hi in [
            ("G3", params.get("g3_min"), params.get("g3_max")),
            ("AD", params.get("ad_min"), params.get("ad_max")),
        ]:
            if lo is not None and sh_now[cia] < lo:
                company_ok = False
            if hi is not None and sh_now[cia] > hi:
                company_ok = False

        if abs(sh_now["JJ"] - params["target_share"]) <= 0.00015 and company_ok:
            break

    # Linhas físicas excluídas: TODOS os cupons dos bilhetes escolhidos.
    bil_series = main_df[columns["bilhete"]].astype(str).str.strip()
    removed_all_rows = main_df.loc[bil_series.isin(removed_tickets)].copy()
    main_final = main_df.loc[~bil_series.isin(removed_tickets)].copy()

    # Marca se cada linha participou do filtro de share.
    participating_index = set(before.index)
    removed_all_rows["PARTICIPOU_CALCULO_SHARE"] = np.where(
        removed_all_rows.index.isin(participating_index),
        "SIM",
        "NÃO",
    )
    removed_all_rows["MOTIVO_EXCLUSAO"] = "AJUSTE DE SHARE - BILHETE COMPLETO"

    # Tabelas comparativas
    after_filtered_raw = apply_user_filters(main_final, columns, filters)
    after = analysis_view(after_filtered_raw, columns)

    daily_b = daily_table(before).rename(columns={
        "TOTAL": "TOTAL_ORIGINAL",
        "JJ_BRL": "JJ_BRL_ORIGINAL", "JJ_SHARE": "JJ_SHARE_ORIGINAL",
        "G3_BRL": "G3_BRL_ORIGINAL", "G3_SHARE": "G3_SHARE_ORIGINAL",
        "AD_BRL": "AD_BRL_ORIGINAL", "AD_SHARE": "AD_SHARE_ORIGINAL",
    })
    daily_a = daily_table(after).rename(columns={
        "TOTAL": "TOTAL_ATUAL",
        "JJ_BRL": "JJ_BRL_ATUAL", "JJ_SHARE": "JJ_SHARE_ATUAL",
        "G3_BRL": "G3_BRL_ATUAL", "G3_SHARE": "G3_SHARE_ATUAL",
        "AD_BRL": "AD_BRL_ATUAL", "AD_SHARE": "AD_SHARE_ATUAL",
    })
    daily_compare = daily_b.merge(daily_a, on="DATA", how="outer")

    route_b = route_table(before).rename(columns={
        "DIRETO": "DIRETO_ORIGINAL",
        "TOTAL": "TOTAL_ORIGINAL",
        "JJ_BRL": "JJ_BRL_ORIGINAL", "JJ_SHARE": "JJ_SHARE_ORIGINAL",
        "G3_BRL": "G3_BRL_ORIGINAL", "G3_SHARE": "G3_SHARE_ORIGINAL",
        "AD_BRL": "AD_BRL_ORIGINAL", "AD_SHARE": "AD_SHARE_ORIGINAL",
        "GAP_JJ_G3_PP": "GAP_JJ_G3_PP_ORIGINAL",
    })
    route_a = route_table(after).rename(columns={
        "DIRETO": "DIRETO_ATUAL",
        "TOTAL": "TOTAL_ATUAL",
        "JJ_BRL": "JJ_BRL_ATUAL", "JJ_SHARE": "JJ_SHARE_ATUAL",
        "G3_BRL": "G3_BRL_ATUAL", "G3_SHARE": "G3_SHARE_ATUAL",
        "AD_BRL": "AD_BRL_ATUAL", "AD_SHARE": "AD_SHARE_ATUAL",
        "GAP_JJ_G3_PP": "GAP_JJ_G3_PP_ATUAL",
    })
    route_compare = route_b.merge(route_a, on="TRECHO", how="outer")

    total_a, vals_a, sh_a = shares(after)
    company_compare = pd.DataFrame([
        {
            "CIA": cia,
            "VALOR_ORIGINAL": vals_b[cia],
            "SHARE_ORIGINAL": sh_b[cia],
            "VALOR_ATUAL": vals_a[cia],
            "SHARE_ATUAL": sh_a[cia],
        }
        for cia in TARGET_CIAS
    ])

    # Anexa shares antes/depois por dia e trecho às linhas removidas.
    if not removed_all_rows.empty:
        tmp = removed_all_rows.copy()
        tmp["_DATA"] = parse_date(tmp[columns["data_emissao"]]).dt.normalize()
        tmp["_TRECHO"] = (
            tmp[columns["trecho"]].astype(str).str.strip().str.upper()
            if columns["trecho"] else ""
        )

        dmap_b = daily_compare.set_index("DATA")["JJ_SHARE_ORIGINAL"].to_dict()
        dmap_a = daily_compare.set_index("DATA")["JJ_SHARE_ATUAL"].to_dict()
        rmap_b = route_compare.set_index("TRECHO")["JJ_SHARE_ORIGINAL"].to_dict()
        rmap_a = route_compare.set_index("TRECHO")["JJ_SHARE_ATUAL"].to_dict()

        removed_all_rows["SHARE_LATAM_DIA_ANTES"] = tmp["_DATA"].map(dmap_b)
        removed_all_rows["SHARE_LATAM_DIA_DEPOIS"] = tmp["_DATA"].map(dmap_a)
        removed_all_rows["SHARE_LATAM_TRECHO_ANTES"] = tmp["_TRECHO"].map(rmap_b)
        removed_all_rows["SHARE_LATAM_TRECHO_DEPOIS"] = tmp["_TRECHO"].map(rmap_a)

    summary = {
        "original_share_jj": sh_b["JJ"],
        "final_share_jj": sh_a["JJ"],
        "original_share_g3": sh_b["G3"],
        "final_share_g3": sh_a["G3"],
        "original_share_ad": sh_b["AD"],
        "final_share_ad": sh_a["AD"],
        "target_share": params["target_share"],
        "removed_tickets": len(removed_tickets),
        "removed_rows": len(removed_all_rows),
        "removed_value_physical": float(parse_money(removed_all_rows[columns["valor"]]).sum()) if not removed_all_rows.empty else 0.0,
        "final_rows": len(main_final),
    }

    return {
        "main_final": main_final,
        "removed_share": removed_all_rows,
        "daily_compare": daily_compare,
        "route_compare": route_compare,
        "company_compare": company_compare,
        "summary": summary,
    }



# ============================================================
# PDF - RELATÓRIO INICIAL / AUDITORIA
# ============================================================

def _brl(v):
    return ("R$ {:,.2f}".format(float(v or 0))
            .replace(",", "X").replace(".", ",").replace("X", "."))


def _pct(v):
    return f"{float(v or 0) * 100:.2f}%".replace(".", ",")


def _pdf_text(v):
    return xml_escape(str(v if v is not None else ""))


def create_initial_report_pdf(path, preview, stage1_stats, filters, logo_path=None):
    """
    Gera um PDF equivalente ao quadro 'Share atual do relatório' da tela,
    acrescido da auditoria da formatação e dos filtros realmente aplicados.
    Os valores são derivados do arquivo carregado pelo usuário.
    """
    doc = SimpleDocTemplate(
        str(path),
        pagesize=landscape(A4),
        rightMargin=12 * mm,
        leftMargin=12 * mm,
        topMargin=10 * mm,
        bottomMargin=10 * mm,
        title="Relatório inicial de share e formatação",
        author="Tratamento BDA",
    )

    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "TitleCustom", parent=styles["Title"], fontName="Helvetica-Bold",
        fontSize=18, leading=22, textColor=colors.HexColor("#172033"),
        spaceAfter=5,
    )
    subtitle = ParagraphStyle(
        "SubCustom", parent=styles["BodyText"], fontSize=9.5, leading=13,
        textColor=colors.HexColor("#657086"), spaceAfter=8,
    )
    h2 = ParagraphStyle(
        "H2Custom", parent=styles["Heading2"], fontName="Helvetica-Bold",
        fontSize=12, leading=15, textColor=colors.HexColor("#075F91"),
        spaceBefore=8, spaceAfter=6, keepWithNext=1,
    )
    body = ParagraphStyle(
        "BodyCustom", parent=styles["BodyText"], fontSize=8.5, leading=12,
        textColor=colors.HexColor("#334155"),
    )
    small = ParagraphStyle(
        "SmallCustom", parent=body, fontSize=7.7, leading=10.5,
        textColor=colors.HexColor("#657086"),
    )
    metric_label = ParagraphStyle(
        "MetricLabel", parent=small, fontSize=7.5, leading=9,
        textColor=colors.HexColor("#64748B"),
    )
    metric_value = ParagraphStyle(
        "MetricValue", parent=body, fontName="Helvetica-Bold",
        fontSize=15, leading=17, textColor=colors.HexColor("#172033"),
    )

    story = []

    # Cabeçalho
    head_data = []
    if logo_path and Path(logo_path).exists():
        img = RLImage(str(logo_path), width=22 * mm, height=22 * mm)
        head_data = [[
            img,
            [
                Paragraph("Relatório atual do processamento", title),
                Paragraph(
                    "Share do relatório após a formatação e conforme os filtros selecionados.",
                    subtitle,
                ),
            ]
        ]]
        ht = Table(head_data, colWidths=[27 * mm, 240 * mm])
        ht.setStyle(TableStyle([
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("LEFTPADDING", (0,0), (-1,-1), 0),
            ("RIGHTPADDING", (0,0), (-1,-1), 0),
            ("TOPPADDING", (0,0), (-1,-1), 0),
            ("BOTTOMPADDING", (0,0), (-1,-1), 0),
        ]))
        story.append(ht)
    else:
        story.append(Paragraph("Relatório atual do processamento", title))
        story.append(Paragraph(
            "Share do relatório após a formatação e conforme os filtros selecionados.",
            subtitle,
        ))

    story.append(Spacer(1, 3 * mm))

    # Cards de share
    metric_rows = [[
        [Paragraph("LATAM (JJ)", metric_label), Paragraph(_pct(preview["jj"]), metric_value)],
        [Paragraph("G3", metric_label), Paragraph(_pct(preview["g3"]), metric_value)],
        [Paragraph("AD", metric_label), Paragraph(_pct(preview["ad"]), metric_value)],
        [Paragraph("Total BRL", metric_label), Paragraph(_brl(preview["total"]), metric_value)],
    ]]
    mt = Table(metric_rows, colWidths=[66 * mm] * 4)
    mt.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#F8FBFE")),
        ("BOX", (0,0), (-1,-1), 0.7, colors.HexColor("#D8E3EF")),
        ("INNERGRID", (0,0), (-1,-1), 0.7, colors.HexColor("#D8E3EF")),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 7),
        ("RIGHTPADDING", (0,0), (-1,-1), 7),
        ("TOPPADDING", (0,0), (-1,-1), 8),
        ("BOTTOMPADDING", (0,0), (-1,-1), 8),
    ]))
    story.append(mt)
    story.append(Spacer(1, 4 * mm))

    info = [
        ["Período", f'{preview["period_start"]} a {preview["period_end"]}'],
        ["Linhas consideradas", f'{preview["rows"]:,}'.replace(",", ".")],
        ["Bilhetes únicos", f'{preview["unique_tickets"]:,}'.replace(",", ".")],
    ]
    it = Table(info, colWidths=[48 * mm, 216 * mm])
    it.setStyle(TableStyle([
        ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
        ("FONTNAME", (1,0), (1,-1), "Helvetica"),
        ("FONTSIZE", (0,0), (-1,-1), 8.5),
        ("TEXTCOLOR", (0,0), (0,-1), colors.HexColor("#64748B")),
        ("TEXTCOLOR", (1,0), (1,-1), colors.HexColor("#334155")),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING", (0,0), (-1,-1), 4),
        ("LINEBELOW", (0,-1), (-1,-1), 0.5, colors.HexColor("#E2E8F0")),
    ]))
    story.append(it)

    # Filtros
    story.append(Paragraph("Filtros usados no share", h2))
    cons = filters.get("consolidadas") or []
    cons_display = ", ".join(cons[:12])
    if len(cons) > 12:
        cons_display += f" ... (+{len(cons)-12})"
    filter_rows = [
        ["CUPOM", ", ".join(filters.get("cupons") or []) or "Todos"],
        ["TIPO", ", ".join(filters.get("tipos") or []) or "Todos"],
        ["CONSOLIDADA 2", cons_display or "Todas"],
        ["DATA_EMISSAO", f'{filters.get("date_start") or "início"} até {filters.get("date_end") or "fim"}'],
    ]
    ft = Table(
        [[Paragraph(f"<b>{_pdf_text(a)}</b>", body), Paragraph(_pdf_text(b), body)] for a,b in filter_rows],
        colWidths=[48 * mm, 216 * mm],
    )
    ft.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (0,-1), colors.HexColor("#EFF6FB")),
        ("BOX", (0,0), (-1,-1), 0.5, colors.HexColor("#D8E3EF")),
        ("INNERGRID", (0,0), (-1,-1), 0.4, colors.HexColor("#E5EDF5")),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
        ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ("TOPPADDING", (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
    ]))
    story.append(ft)

    # Resumo da formatação
    story.append(Paragraph("Auditoria da formatação", h2))
    s = stage1_stats or {}
    audit_rows = [
        ["Original", f'{s.get("original_rows", 0):,} linhas / {s.get("original_cols", 0)} colunas'.replace(",", ".")],
        ["GOV removida", "SIM" if s.get("gov_removed") else "NÃO"],
        ["Após formatação", f'{s.get("formatted_rows", 0):,}'.replace(",", ".")],
        ["Excluídas na formatação", f'{s.get("excluded_format_rows", 0):,}'.replace(",", ".")],
        ["DATA_RESERVA preenchidas", f'{s.get("data_reserva_filled", 0):,}'.replace(",", ".")],
        ["WEST", f'{s.get("west_rows", 0):,}'.replace(",", ".")],
        ["Base principal", f'{s.get("main_rows", 0):,}'.replace(",", ".")],
    ]
    at = Table(
        [[Paragraph(f"<b>{_pdf_text(a)}</b>", body), Paragraph(_pdf_text(b), body)] for a,b in audit_rows],
        colWidths=[70 * mm, 194 * mm],
    )
    at.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (0,-1), colors.HexColor("#F8FAFC")),
        ("BOX", (0,0), (-1,-1), 0.5, colors.HexColor("#D8E3EF")),
        ("INNERGRID", (0,0), (-1,-1), 0.4, colors.HexColor("#E5EDF5")),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
        ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ("TOPPADDING", (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ]))
    story.append(at)

    # Exclusões detalhadas
    exc = s.get("exclusion_counts", {}) or {}
    if exc:
        story.append(Paragraph("Exclusões realizadas na formatação", h2))
        exc_rows = [["Regra", "Quantidade"]]
        for motivo, qtd in exc.items():
            exc_rows.append([_pdf_text(motivo), f"{int(qtd):,}".replace(",", ".")])
        et = Table(exc_rows, colWidths=[220 * mm, 44 * mm], repeatRows=1)
        et.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#0B84C6")),
            ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTNAME", (0,1), (-1,-1), "Helvetica"),
            ("FONTSIZE", (0,0), (-1,-1), 8),
            ("ALIGN", (1,1), (1,-1), "RIGHT"),
            ("BOX", (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
            ("INNERGRID", (0,0), (-1,-1), 0.35, colors.HexColor("#E2E8F0")),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F8FAFC")]),
            ("LEFTPADDING", (0,0), (-1,-1), 5),
            ("RIGHTPADDING", (0,0), (-1,-1), 5),
            ("TOPPADDING", (0,0), (-1,-1), 4),
            ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ]))
        story.append(et)

    artmax = s.get("artmax_log", []) or []
    if artmax:
        story.append(Paragraph("Tratamento ART / MAX", h2))
        rows = [["Base", "Regra aplicada", "JJ", "JJ final 2412", "Alterados"]]
        for item in artmax:
            rows.append([
                Paragraph(_pdf_text(item.get("base", "")), small),
                Paragraph(_pdf_text(item.get("regra", "")), small),
                str(item.get("jj", 0)),
                str(item.get("jj_2412", 0)),
                str(item.get("alterados", 0)),
            ])
        tt = Table(rows, colWidths=[22*mm, 154*mm, 22*mm, 34*mm, 32*mm], repeatRows=1)
        tt.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#075F91")),
            ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE", (0,0), (-1,-1), 7.2),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("BOX", (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
            ("INNERGRID", (0,0), (-1,-1), 0.35, colors.HexColor("#E2E8F0")),
            ("LEFTPADDING", (0,0), (-1,-1), 4),
            ("RIGHTPADDING", (0,0), (-1,-1), 4),
            ("TOPPADDING", (0,0), (-1,-1), 4),
            ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ]))
        story.append(tt)

    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph(
        "Observação: este PDF registra o cenário inicial do relatório, antes das exclusões para ajuste de share. "
        "O share é calculado por valor BRL e respeita exatamente os filtros escolhidos na tela.",
        small,
    ))

    doc.build(story)


def create_initial_report_txt(path, preview, stage1_stats, filters):
    """Registra em texto os indicadores do painel antes do ajuste de share."""
    s = stage1_stats or {}
    count = lambda value: f"{int(value or 0):,}".replace(",", ".")
    lines = [
        "SHARE ATUAL DO RELATÓRIO",
        "Calculado por valor BRL sobre a tabela filtrada, antes do ajuste de share.",
        "",
        f"LATAM (JJ): {_pct(preview['jj'])}",
        f"G3: {_pct(preview['g3'])}",
        f"AD: {_pct(preview['ad'])}",
        f"Total BRL: {_brl(preview['total'])}",
        "",
        f"Período: {preview['period_start']} a {preview['period_end']}",
        f"Linhas consideradas: {count(preview['rows'])}",
        f"Bilhetes únicos: {count(preview['unique_tickets'])}",
        "",
        "AUDITORIA DA FORMATAÇÃO",
        f"Original: {count(s.get('original_rows'))} linhas / {count(s.get('original_cols'))} colunas",
        f"GOV removida: {'SIM' if s.get('gov_removed') else 'NÃO'}",
        f"Após formatação: {count(s.get('formatted_rows'))}",
        f"Excluídas na formatação: {count(s.get('excluded_format_rows'))}",
        f"DATA_RESERVA preenchidas: {count(s.get('data_reserva_filled'))}",
        f"WEST: {count(s.get('west_rows'))}",
        f"Base principal: {count(s.get('main_rows'))}",
        "",
        "EXCLUSÕES",
    ]
    exclusion_counts = s.get("exclusion_counts") or {}
    lines.extend(
        f"{reason}: {count(quantity)}"
        for reason, quantity in exclusion_counts.items()
    )
    if not exclusion_counts:
        lines.append("Nenhuma exclusão registrada.")

    lines.extend([
        "",
        "PREENCHIMENTOS",
        f"DATA_RESERVA: {count(s.get('data_reserva_filled'))}",
        f"GRUPO_EMPRESA origem: {s.get('grupo_empresa_source') or 'não encontrada'}",
        f"GRUPO_CONSOLIDADA atualizadas: {count(s.get('grupo_consolidada_updated'))}",
        "",
        "ART / MAX",
    ])
    artmax_log = s.get("artmax_log") or []
    lines.extend(
        f"{item.get('base', '')}: {item.get('regra', '')} - "
        f"JJ: {count(item.get('jj'))}; "
        f"JJ final 2412: {count(item.get('jj_2412'))}; "
        f"alterados: {count(item.get('alterados'))}"
        for item in artmax_log
    )
    if not artmax_log:
        lines.append("Nenhum tratamento ART / MAX registrado.")

    lines.extend([
        "",
        "FILTROS USADOS NO SHARE",
        f"CUPOM: {', '.join(filters.get('cupons') or []) or 'Todos'}",
        f"TIPO: {', '.join(filters.get('tipos') or []) or 'Todos'}",
        f"CONSOLIDADA 2: {count(len(filters.get('consolidadas') or []))} selecionadas",
        f"WHITE LABEL: {'incluída' if any(str(x).strip().upper() == 'WHITE LABEL' for x in filters.get('consolidadas') or []) else 'fora do share'}",
        f"DATA_EMISSAO: {filters.get('date_start') or 'início'} até {filters.get('date_end') or 'fim'}",
        "",
    ])
    Path(path).write_text("\n".join(lines), encoding="utf-8-sig")


# ============================================================
# SAÍDAS / ZIP
# ============================================================

def write_csv(df, path):
    df.to_csv(path, sep=";", index=False, encoding="utf-8-sig")


def period_label(df, date_col, include_year=False):
    dates = parse_date(df[date_col]).dropna()
    if dates.empty:
        return "SEM_PERIODO"
    a = dates.min()
    b = dates.max()
    if include_year:
        return f"{a.strftime('%d-%m-%Y')} A {b.strftime('%d-%m-%Y')}"
    return f"{a.strftime('%d')} A {b.strftime('%d')}"


def build_zip_outputs(
    job_dir,
    original_path,
    formatted_full,
    excluded_format,
    west,
    main_before_share,
    adjustment,
    columns,
    params,
    stage1_stats=None,
):
    out_dir = job_dir / "saida"
    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir()

    p1 = out_dir / "01_ORIGINAL"
    p2 = out_dir / "02_FORMATACAO"
    p3 = out_dir / "03_WEST"
    p4 = out_dir / "04_AJUSTE_SHARE"
    p5 = out_dir / "05_FINAL"
    for p in (p1, p2, p3, p4, p5):
        p.mkdir()

    shutil.copy2(original_path, p1 / original_path.name)

    write_csv(formatted_full, p2 / "BASE_FORMATADA.csv")
    write_csv(excluded_format, p2 / "LINHAS_EXCLUIDAS_FORMATACAO.csv")

    west_period = period_label(west, columns["data_emissao"], include_year=True)
    write_csv(
        west,
        p3 / f"WEST ({west_period}) COM CONSOLIDADA.csv",
    )
    write_csv(
        west.drop(columns=[columns["cons2"]], errors="ignore"),
        p3 / f"WEST ({west_period}).csv",
    )

    removed = adjustment["removed_share"].copy()
    write_csv(removed, p4 / "LINHAS_EXCLUIDAS_AJUSTE_SHARE.csv")

    report_path = p4 / "RELATORIO_SHARE.xlsx"
    with pd.ExcelWriter(report_path, engine="xlsxwriter") as writer:
        s = adjustment["summary"]
        pd.DataFrame([
            ["Share LATAM original", s["original_share_jj"]],
            ["Share LATAM final", s["final_share_jj"]],
            ["Share G3 original", s["original_share_g3"]],
            ["Share G3 final", s["final_share_g3"]],
            ["Share AD original", s["original_share_ad"]],
            ["Share AD final", s["final_share_ad"]],
            ["Meta LATAM", s["target_share"]],
            ["Bilhetes retirados", s["removed_tickets"]],
            ["Linhas/cupons retirados", s["removed_rows"]],
            ["Valor físico retirado", s["removed_value_physical"]],
        ], columns=["INDICADOR", "VALOR"]).to_excel(writer, sheet_name="RESUMO", index=False)

        adjustment["company_compare"].to_excel(writer, sheet_name="SHARE_CIA", index=False)
        adjustment["daily_compare"].to_excel(writer, sheet_name="SHARE_DIA_A_DIA", index=False)
        adjustment["route_compare"].to_excel(writer, sheet_name="SHARE_POR_TRECHO", index=False)
        removed.to_excel(writer, sheet_name="BILHETES_RETIRAR", index=False)

        wb = writer.book
        pct = wb.add_format({"num_format": "0.00%"})
        money = wb.add_format({"num_format": 'R$ #,##0.00'})
        header = wb.add_format({
            "bold": True,
            "bg_color": "#0B78B7",
            "font_color": "white",
            "border": 1,
        })

        for sheet_name, df_sheet in [
            ("SHARE_CIA", adjustment["company_compare"]),
            ("SHARE_DIA_A_DIA", adjustment["daily_compare"]),
            ("SHARE_POR_TRECHO", adjustment["route_compare"]),
            ("BILHETES_RETIRAR", removed),
        ]:
            ws = writer.sheets[sheet_name]
            ws.freeze_panes(1, 0)
            if len(df_sheet.columns):
                ws.autofilter(0, 0, max(len(df_sheet), 1), len(df_sheet.columns) - 1)

            for ci, col in enumerate(df_sheet.columns):
                ws.write(0, ci, col, header)
                width = min(max(len(str(col)) + 2, 14), 34)
                fmt = None
                if "SHARE" in str(col):
                    fmt = pct
                elif any(x in str(col) for x in ("BRL", "VALOR", "TOTAL")):
                    fmt = money
                ws.set_column(ci, ci, width, fmt)

    # PDF do cenário inicial (equivalente ao painel 'Share atual do relatório')
    initial_filtered = apply_user_filters(main_before_share, columns, params.get("filters", {}))
    initial_preview = build_share_preview(initial_filtered, columns)
    pdf_path = p4 / "RELATORIO_INICIAL_SHARE_E_FORMATACAO.pdf"
    create_initial_report_pdf(
        pdf_path,
        preview=initial_preview,
        stage1_stats=stage1_stats or {},
        filters=params.get("filters", {}),
        logo_path=None,
    )
    shutil.copy2(pdf_path, out_dir / "SHARE_ATUAL_E_AUDITORIA.pdf")
    create_initial_report_txt(
        out_dir / "SHARE_ATUAL_E_AUDITORIA.txt",
        preview=initial_preview,
        stage1_stats=stage1_stats or {},
        filters=params.get("filters", {}),
    )

    final_50 = adjustment["main_final"].copy()
    period_short = period_label(final_50, columns["data_emissao"], include_year=False)

    write_csv(final_50, p5 / "BASE FINAL 50 COLUNAS.csv")
    write_csv(
        final_50.drop(columns=[columns["cons2"]], errors="ignore"),
        p5 / f"{period_short}.csv",
    )

    s = adjustment["summary"]
    txt = f"""RESUMO DO PROCESSAMENTO

Período final: {period_short}

FILTROS USADOS NO SHARE:
{json.dumps(params.get('filters', {}), ensure_ascii=False, indent=2)}

Share LATAM original: {s['original_share_jj']:.4%}
Share LATAM final:    {s['final_share_jj']:.4%}
Meta LATAM:           {s['target_share']:.4%}

Share G3 original:    {s['original_share_g3']:.4%}
Share G3 final:       {s['final_share_g3']:.4%}

Share AD original:    {s['original_share_ad']:.4%}
Share AD final:       {s['final_share_ad']:.4%}

Bilhetes retirados: {s['removed_tickets']}
Linhas/cupons retirados: {s['removed_rows']}
Valor físico retirado: R$ {s['removed_value_physical']:,.2f}

IMPORTANTE:
A seleção do bilhete é feita com base no universo filtrado.
Quando um bilhete é escolhido, TODAS as linhas/cupons desse bilhete
são retiradas da base final, mesmo que não tenham participado do filtro.
"""
    (out_dir / "RESUMO_PROCESSAMENTO.txt").write_text(txt, encoding="utf-8")

    zip_name = f"PROCESSAMENTO_BDA_{period_short.replace(' ', '_')}.zip"
    zip_path = job_dir / zip_name
    if zip_path.exists():
        zip_path.unlink()

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for file in out_dir.rglob("*"):
            if file.is_file():
                z.write(file, file.relative_to(out_dir))

    return zip_path
