"""Tratamento e relatórios da BDA GOL.

As regras de limpeza originais estão em gol_legacy.py, cópia do arquivo V4
fornecido para esta integração. O fluxo web não executa a interface Tkinter.
"""
from __future__ import annotations

from pathlib import Path
import json
import re
import zipfile

import pandas as pd
import numpy as np

import gol_legacy as legacy
from processor import read_csv_robust

CIASS = ("JJ", "G3", "AD")
REPORTS = ("CONS", "CORP", "GOV")


def read_table(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        return read_csv_robust(path)[0]
    if path.suffix.lower() == ".xlsx":
        return pd.read_excel(path, dtype=object)
    raise ValueError("Envie CSV ou XLSX.")


def tidy(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    if out.columns.duplicated().any():
        out.columns = legacy.tornar_nomes_unicos(list(out.columns))
    for c in out.select_dtypes(include=["object", "string"]).columns:
        out[c] = out[c].map(lambda v: v.strip() if isinstance(v, str) else v)
    return out


def _dates(df: pd.DataFrame) -> pd.DataFrame:
    for c in ("DATA_EMISSAO", "DATA_EMBARQUE"):
        df[c] = parse_dates(df[c])
    return df


def parse_dates(values: pd.Series) -> pd.Series:
    """Lê datas textuais e números seriais gravados em planilhas Excel."""
    numeric = pd.to_numeric(values, errors="coerce")
    serial = numeric.between(20000, 80000)
    result = pd.to_datetime(values.where(~serial), errors="coerce", dayfirst=True)
    result.loc[serial] = pd.to_datetime(numeric.loc[serial], unit="D",
                                         origin="1899-12-30", errors="coerce")
    return result


def recover_emission_dates(raw: pd.DataFrame, source: pd.DataFrame) -> pd.DataFrame:
    """Recupera DATA_EMISSAO somente de uma origem que corresponde às vendas."""
    target = tidy(raw)
    target.columns = legacy.tornar_nomes_unicos([legacy.normalizar_nome_coluna(c) for c in target])
    origin = tidy(source)
    origin.columns = legacy.tornar_nomes_unicos([legacy.normalizar_nome_coluna(c) for c in origin])
    needed = {"BILHETE", "CUPOM", "CIA_BILHETE", "IATA_AGENCIA", "VALOR_BRL", "DATA_EMBARQUE"}
    if not needed <= set(target) or not needed | {"DATA_EMISSAO"} <= set(origin):
        raise ValueError("A origem das datas deve conter DATA_EMISSAO, BILHETE, CUPOM, CIA_BILHETE, IATA_AGENCIA, VALOR_BRL e DATA_EMBARQUE.")
    if len(target) != len(origin):
        raise ValueError("A origem das datas tem quantidade diferente de vendas.")
    def keys(frame):
        return pd.MultiIndex.from_arrays([
            frame["BILHETE"].map(legacy.somente_digitos),
            frame["CUPOM"].map(legacy.somente_digitos),
        ], names=["BILHETE", "CUPOM"])
    target_key, origin_key = keys(target), keys(origin)
    if not target_key.is_unique or not origin_key.is_unique or not target_key.isin(origin_key).all():
        raise ValueError("Não foi possível relacionar todas as vendas por BILHETE e CUPOM únicos.")
    aligned = origin.set_axis(origin_key).reindex(target_key).reset_index(drop=True)
    def normalized(frame, col):
        return frame[col].fillna("").astype(str).str.strip().str.upper()
    for col in ("CIA_BILHETE", "IATA_AGENCIA"):
        a = target[col].map(legacy.somente_digitos) if col == "IATA_AGENCIA" else normalized(target, col)
        b = aligned[col].map(legacy.somente_digitos) if col == "IATA_AGENCIA" else normalized(aligned, col)
        if not a.reset_index(drop=True).equals(b.reset_index(drop=True)):
            raise ValueError(f"A origem das datas não corresponde à coluna {col} da base complementar.")
    a_value = target["VALOR_BRL"].map(legacy.converter_numero_brasileiro).round(2).reset_index(drop=True)
    b_value = aligned["VALOR_BRL"].map(legacy.converter_numero_brasileiro).round(2).reset_index(drop=True)
    if not a_value.equals(b_value):
        raise ValueError("A origem das datas não corresponde aos valores da base complementar.")
    if not parse_dates(target["DATA_EMBARQUE"]).reset_index(drop=True).eq(
            parse_dates(aligned["DATA_EMBARQUE"]).reset_index(drop=True)).all():
        raise ValueError("A origem das datas não corresponde às datas de embarque da base complementar.")
    emission = parse_dates(aligned["DATA_EMISSAO"])
    if emission.isna().any():
        raise ValueError("A origem contém DATA_EMISSAO vazia ou inválida.")
    target["DATA_EMISSAO"] = emission.to_numpy()
    return target


def format_gol(raw: pd.DataFrame, corp_iatas: set[int]):
    raw = tidy(raw)
    cons_sources = [c for c in raw if legacy.normalizar_nome_coluna(c).startswith("CONSOLIDADA")]
    df = legacy.preparar_colunas(raw.copy())
    if len(cons_sources) >= 2:
        df["CONSOLIDADA_1"] = raw[cons_sources[0]].map(legacy.normalizar_texto)
        df["CONSOLIDADA_2"] = raw[cons_sources[1]].map(legacy.normalizar_texto)
    legacy.validar_colunas(df)
    df.insert(0, "_LINHA_ORIGINAL", range(2, len(df) + 2))
    reasons = legacy.montar_motivos_exclusao(df)
    excluded = df.loc[reasons.map(bool)].copy()
    excluded.insert(0, "MOTIVO_EXCLUSAO", reasons.loc[excluded.index].map("; ".join))
    df = df.loc[~reasons.map(bool)].copy()
    df = _dates(df)
    df["VALOR_BRL"] = df["VALOR_BRL"].map(legacy.converter_numero_brasileiro)
    df["IATA_AGENCIA"] = pd.array(
        [int(d) if d else pd.NA for d in df["IATA_AGENCIA"].map(legacy.somente_digitos)],
        dtype="Int64",
    )
    if "CONSOLIDADA_1" not in df.columns:
        df["CONSOLIDADA_1"] = ""
    if "CONSOLIDADA_2" not in df.columns:
        df["CONSOLIDADA_2"] = df.get("CONSOLIDADA", "")
    base = df["BASE_VENDA"].fillna("").astype(str).str.replace(r"\s+", "", regex=True)
    artmax = base.str.replace(r"-(CORP|GOV)$", "", case=False, regex=True).str.upper().isin(["ART", "MAX"])
    base = base.str.replace(r"-(CORP|GOV)$", "", case=False, regex=True)
    gov = (df["GOV"].map(legacy.remover_acentos_maiusculo).eq("SIM")
           if "GOV" in df else pd.Series(False, index=df.index))
    corp = df["IATA_AGENCIA"].isin(corp_iatas).fillna(False) & ~artmax & ~gov
    df["BASE_VENDA"] = base + gov.map({True: "-GOV", False: ""})
    df.loc[corp, "BASE_VENDA"] = base.loc[corp] + "-CORP"
    df.loc[artmax, "BASE_VENDA"] = base.loc[artmax]
    iata_final_255 = df["IATA_AGENCIA"].astype("string").str.endswith("255").fillna(False)
    gol_in_artmax = (df["CIA_BILHETE"].fillna("").astype(str).str.strip().str.upper().eq("G3")
                     & artmax)
    missing = gol_in_artmax & ~iata_final_255
    df.loc[missing, "BASE_VENDA"] = "CGB"
    df.loc[missing, "CONSOLIDADA_1"] = "43"
    df.loc[missing, "CONSOLIDADA_2"] = "ARARAUNA TURISMO"
    df.loc[missing, "CONSOLIDADA"] = "ARARAUNA TURISMO"
    df["CNPJ"] = legacy.CNPJ_PADRAO
    df["NUM_VOO"] = df["NUM_VOO"].map(legacy.gerar_numero_voo)
    source_class = "CABINE" if "CABINE" in df else "CLASSE_SERVICO"
    df["CLASSE_SERVICO"] = df[source_class].map(legacy.gerar_classe_servico)
    df["BILHETE"] = df["BILHETE"].map(legacy.tratar_bilhete)
    df["TRECHO"] = df["TRECHO"].map(legacy.tratar_trecho)
    df, coupon_audit = legacy.corrigir_cupons(df)
    for c in legacy.COLUNAS_FINAIS:
        if c not in df:
            df[c] = ""
    cols = list(dict.fromkeys(legacy.COLUNAS_FINAIS + ["CONSOLIDADA_1", "CONSOLIDADA_2"]))
    result = tidy(df[cols])
    return result, tidy(excluded), tidy(coupon_audit), {
        "original_rows": len(raw), "formatted_rows": len(result),
        "excluded_rows": len(excluded), "cgb_rows": int(missing.sum()),
        "artmax_rows": int(artmax.sum()),
        "gol_artmax_rows": int(gol_in_artmax.sum()),
        "gol_artmax_iata_255_rows": int((gol_in_artmax & iata_final_255).sum()),
        "corp_iatas": sorted(corp_iatas),
    }


def merge_period(current: pd.DataFrame, extra: pd.DataFrame) -> pd.DataFrame:
    extra = tidy(extra)
    extra.columns = legacy.tornar_nomes_unicos([legacy.normalizar_nome_coluna(c) for c in extra])
    required = {"DATA_EMISSAO", "DATA_EMBARQUE", "CUPOM", "TIPO", "BASE_VENDA",
                "CIA_BILHETE", "VALOR_BRL", "BILHETE", "IATA_AGENCIA"}
    absent = required - set(extra)
    if absent:
        raise ValueError("Base complementar sem colunas: " + ", ".join(sorted(absent)))
    extra = extra.reindex(columns=[c for c in current if c != "_SOURCE_PERIOD"], fill_value="")
    if "CONSOLIDADA_2" in extra:
        blank = extra["CONSOLIDADA_2"].fillna("").astype(str).str.strip().eq("")
        extra.loc[blank, "CONSOLIDADA_2"] = extra.loc[blank, "CONSOLIDADA"]
    extra = _dates(extra)
    extra["VALOR_BRL"] = extra["VALOR_BRL"].map(legacy.converter_numero_brasileiro)
    extra["IATA_AGENCIA"] = pd.array(
        [int(d) if d else pd.NA for d in extra["IATA_AGENCIA"].map(legacy.somente_digitos)],
        dtype="Int64",
    )
    extra["_SOURCE_PERIOD"] = "COMPLEMENTAR"
    return tidy(pd.concat([current, extra], ignore_index=True))


def prepare_complement(raw: pd.DataFrame, corp_iatas: set[int]):
    """Detecta se a base complementar já tem o layout GOL final."""
    candidate = tidy(raw)
    candidate.columns = legacy.tornar_nomes_unicos([legacy.normalizar_nome_coluna(c) for c in candidate])
    names = set(candidate)
    formatted = set(legacy.COLUNAS_FINAIS) <= names and "CABINE" not in names and len(names) <= 25
    if formatted and not candidate.empty:
        date_ok = all(parse_dates(candidate[c]).notna().all()
                      for c in ("DATA_EMISSAO", "DATA_EMBARQUE"))
        iata_ok = candidate["IATA_AGENCIA"].map(legacy.somente_digitos).str.fullmatch(r"\d+").fillna(False).all()
        ticket_ok = candidate["BILHETE"].map(legacy.somente_digitos).str.fullmatch(r"\d{13}").fillna(False).all()
        value_ok = candidate["VALOR_BRL"].map(legacy.converter_numero_brasileiro).notna().all()
        formatted = bool(date_ok and iata_ok and ticket_ok and value_ok)
    if formatted:
        return candidate, "JÁ FORMATADA", None
    result, excluded, coupons, stats = format_gol(raw, corp_iatas)
    stats["coupon_corrections"] = len(coupons)
    return result, "FORMATADA AUTOMATICAMENTE", stats


def week_dates(current: pd.DataFrame) -> list[str]:
    dates = pd.to_datetime(current["DATA_EMISSAO"], errors="coerce").dropna()
    return sorted(dates.dt.strftime("%Y-%m-%d").unique().tolist())


def _coupon_text(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip().str.replace(r"\.0+$", "", regex=True)


def filter_options(df: pd.DataFrame) -> dict:
    dates = pd.to_datetime(df["DATA_EMISSAO"], errors="coerce", dayfirst=True).dropna()
    bases = df["BASE_VENDA"].fillna("").astype(str).str.strip()
    upper = bases.str.upper()
    return {
        "cupons": sorted(set(_coupon_text(df["CUPOM"])) - {""}),
        "tipos": sorted(set(df["TIPO"].fillna("").astype(str).str.strip().str.upper()) - {""}),
        "bases": {
            "CONS": sorted(set(bases.loc[~upper.isin(["ART", "MAX"]) & ~upper.str.endswith(("-CORP", "-GOV"))]) - {""}),
            "CORP": sorted(set(bases.loc[upper.str.endswith("-CORP")]) - {""}),
            "GOV": sorted(set(bases.loc[upper.str.endswith("-GOV")]) - {""}),
        },
        "date_start": dates.min().strftime("%Y-%m-%d") if not dates.empty else "",
        "date_end": dates.max().strftime("%Y-%m-%d") if not dates.empty else "",
    }


def default_filters(df: pd.DataFrame) -> dict:
    options = filter_options(df)
    return {"cupons": ["1"] if "1" in options["cupons"] else options["cupons"],
            "tipos": ["N"] if "N" in options["tipos"] else options["tipos"],
            "bases": {k: [v for v in values if v.upper() != "WHITE LABEL"]
                      for k, values in options["bases"].items()},
            "date_start": options["date_start"], "date_end": options["date_end"]}


def filtered_rows(df: pd.DataFrame, filters: dict | None) -> pd.DataFrame:
    if filters is None:
        filters = {"cupons": ["1"], "tipos": ["N"]}
    mask = pd.Series(True, index=df.index)
    if filters.get("cupons"):
        mask &= _coupon_text(df["CUPOM"]).isin(filters["cupons"])
    if filters.get("tipos"):
        mask &= df["TIPO"].fillna("").astype(str).str.strip().str.upper().isin(filters["tipos"])
    dates = pd.to_datetime(df["DATA_EMISSAO"], errors="coerce", dayfirst=True)
    if filters.get("date_start"):
        mask &= dates.ge(pd.Timestamp(filters["date_start"]))
    if filters.get("date_end"):
        mask &= dates.le(pd.Timestamp(filters["date_end"]) + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1))
    return df.loc[mask].copy()


MONTHS_PT = ("", "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
             "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO")


def period_titles(df: pd.DataFrame, week: set[str], filters: dict | None,
                  include_complete: bool) -> dict[str, str]:
    selected = filtered_rows(df, filters)
    def label(rows, prefix):
        dates = pd.to_datetime(rows["DATA_EMISSAO"], errors="coerce", dayfirst=True).dropna()
        if dates.empty:
            return prefix + " — SEM DADOS NOS FILTROS"
        first, last = dates.min(), dates.max()
        if first.year == last.year and first.month == last.month:
            span = f"{first.day:02d} A {last.day:02d} {MONTHS_PT[last.month]}"
        else:
            left = f"{first.day:02d} {MONTHS_PT[first.month]}"
            right = f"{last.day:02d} {MONTHS_PT[last.month]}"
            span = f"{left} A {right}" if first.year == last.year else f"{left} {first.year} A {right} {last.year}"
        return f"{prefix} {span}"
    original = selected.loc[selected.get("_SOURCE_PERIOD", pd.Series("ORIGINAL", index=selected.index)).eq("ORIGINAL")]
    titles = {"SEMANA": label(original, "RELATÓRIO DA SEMANA DE")}
    if include_complete:
        titles["COMPLETO"] = label(selected, "RELATÓRIO COMPLETO")
    return titles


def report_rows(df: pd.DataFrame, kind: str, dates: set[str] | None = None,
                advp: tuple[int | None, int | None] | None = None,
                filters: dict | None = None) -> pd.DataFrame:
    base = df["BASE_VENDA"].fillna("").astype(str).str.strip().str.upper()
    if kind == "CONS":
        mask = ~base.isin(["ART", "MAX"]) & ~base.str.endswith(("-CORP", "-GOV"))
    elif kind == "CORP":
        mask = base.str.endswith("-CORP")
    elif kind == "GOV":
        mask = base.str.endswith("-GOV")
    else:
        raise ValueError("Tipo de relatório inválido.")
    selected = filtered_rows(df, filters)
    mask &= df.index.isin(selected.index)
    if filters is not None and "bases" in filters:
        allowed = filters["bases"].get(kind, [])
        mask &= df["BASE_VENDA"].fillna("").astype(str).str.strip().isin(allowed)
    if dates is not None:
        mask &= pd.to_datetime(df["DATA_EMISSAO"], errors="coerce").dt.strftime("%Y-%m-%d").isin(dates)
        if "_SOURCE_PERIOD" in df:
            mask &= df["_SOURCE_PERIOD"].eq("ORIGINAL")
    if advp is not None:
        days = (pd.to_datetime(df["DATA_EMBARQUE"], errors="coerce") -
                pd.to_datetime(df["DATA_EMISSAO"], errors="coerce")).dt.days
        lo, hi = advp
        if lo is not None:
            mask &= days.ge(lo)
        if hi is not None:
            mask &= days.le(hi)
    return df.loc[mask].copy()


def summary(df: pd.DataFrame) -> dict:
    value = df["VALOR_BRL"].map(legacy.converter_numero_brasileiro).fillna(0)
    cia = df["CIA_BILHETE"].fillna("").astype(str).str.strip().str.upper()
    vals = {c: float(value.loc[cia.eq(c)].sum()) for c in CIASS}
    total = sum(vals.values())
    return {"rows": len(df), "total": total, "values": vals,
            "shares": {c: vals[c] / total if total else 0 for c in CIASS},
            "gap_pp": ((vals["JJ"] - vals["G3"]) / total * 100 if total else 0)}


def report_matrix(df: pd.DataFrame, dates: set[str], groups: dict,
                  advp_kinds: set[str] | None = None,
                  filters: dict | None = None, include_complete: bool = True) -> list[dict]:
    out = []
    scoped = set(REPORTS) if advp_kinds is None else set(advp_kinds)
    for kind in REPORTS:
        periods = [("SEMANA", dates)] + ([("COMPLETO", None)] if include_complete else [])
        for period, period_dates in periods:
            bands = [("GERAL", None), *(groups.items() if kind in scoped else [])]
            for name, bounds in bands:
                rows = report_rows(df, kind, period_dates, bounds, filters)
                out.append({"kind": kind, "period": period, "advp": name, **summary(rows)})
    return out


def report_layouts(matrix: list[dict], groups: dict,
                   advp_kinds: set[str] | None = None) -> dict[str, list[dict]]:
    """Organiza os resumos como as tabelas de duas colunas por faixa.

    As companhias aéreas são sempre ordenadas do maior para o menor valor (tarifa/share),
    com o 'Total geral' na última linha.
    """
    scoped = set(REPORTS) if advp_kinds is None else set(advp_kinds)
    lookup = {(r["period"], r["kind"], r["advp"]): r for r in matrix}
    def band_label(kind, name):
        if name == "GERAL":
            return "ADVP TOTAL" if kind in scoped and groups else f"{kind} ADVP GERAL"
        lo, hi = groups[name]
        return f"ADVP {lo} A {hi}" if hi is not None else f"ADVP {lo}+"
    def make_table(title, bands, period):
        columns = [{"kind": k, "name": n, "label": band_label(k, n)} for k, n in bands]

        def cia_table_total(cia):
            geral = [(k, n) for k, n in bands if n == "GERAL"]
            target_bands = geral if geral else bands
            return sum(lookup[(period, k, n)]["values"].get(cia, 0) for k, n in target_bands)

        sorted_cias = sorted(CIASS, key=cia_table_total, reverse=True)

        rows = []
        for cia in (*sorted_cias, "Total geral"):
            cells = []
            for k, n in bands:
                record = lookup[(period, k, n)]
                value = record["total"] if cia == "Total geral" else record["values"][cia]
                share = (1.0 if record["total"] else 0.0) if cia == "Total geral" else record["shares"][cia]
                cells.append({"value": value, "share": share})
            rows.append({"cia": cia, "cells": cells})
        return {"title": title, "bands": columns, "rows": rows}
    layouts = {}
    for period in dict.fromkeys(r["period"] for r in matrix):
        tables = []
        for kind in REPORTS:
            if kind in scoped and groups:
                bands = [(kind, name) for name in groups] + [(kind, "GERAL")]
            else:
                bands = [(kind, "GERAL")]
            tables.append(make_table(f"RELATÓRIO {kind}", bands, period))
        layouts[period] = tables
    return layouts


def adjust_shares(df: pd.DataFrame, week: set[str], targets: dict,
                  skip_days: set[str], groups: dict, ad_drop_pp: float = 2.0,
                  advp_kinds: set[str] | None = None, filters: dict | None = None,
                  include_complete: bool = True):
    """Busca exclusões de bilhetes JJ/AD da semana, com metas por relatório.

    Uma venda é o bilhete completo. O algoritmo só aceita uma retirada que
    melhora a distância às metas da GOL (G3) e mantém a LATAM em média ao menos 3 pp à frente da GOL,
    tanto nos relatórios gerais quanto em cada faixa de ADVP (curto e longo) e no dia a dia.
    """
    current = df.copy()
    removed_parts = []
    original = report_matrix(df, week, groups, advp_kinds, filters, include_complete)
    active_week_days = sorted(week - skip_days)
    for kind in REPORTS:
        spec = targets.get(kind, {})
        # 1. Universos com metas estipuladas pelo usuário
        target_specs = []
        for key, period in (("general", "COMPLETO"), ("week", "SEMANA")):
            if period == "COMPLETO" and not include_complete:
                continue
            if spec.get(key) is not None:
                target_specs.append((period, None, None, spec[key], f"{kind} {period} GERAL"))
        for name, target in spec.get("advp", {}).items():
            if name in groups:
                target_specs.append(("COMPLETO" if include_complete else "SEMANA", groups[name], None, target, f"{kind} ADVP {name}"))
        if not target_specs:
            continue

        pool = report_rows(current, kind, week, filters=filters)
        day = pd.to_datetime(pool["DATA_EMISSAO"], errors="coerce").dt.strftime("%Y-%m-%d")
        cia = pool["CIA_BILHETE"].fillna("").astype(str).str.strip().str.upper()
        pool = pool.loc[day.isin(active_week_days) & cia.isin(["JJ", "AD"])]
        if pool.empty:
            continue
        full_ticket = current["BILHETE"].astype(str)
        ticket_meta = current.assign(
            _DAY=pd.to_datetime(current["DATA_EMISSAO"], errors="coerce").dt.strftime("%Y-%m-%d"),
            _CIA=current["CIA_BILHETE"].fillna("").astype(str).str.strip().str.upper(),
        ).groupby(full_ticket).agg(
            dates=("_DAY", lambda s: set(s)),
            cias=("_CIA", lambda s: set(s)),
            sources=("_SOURCE_PERIOD", lambda s: set(s)) if "_SOURCE_PERIOD" in current else ("_CIA", lambda s: {"ORIGINAL"}),
        )
        eligible = [t for t in pool["BILHETE"].astype(str).unique()
                    if t in ticket_meta.index and
                    ticket_meta.at[t, "dates"] <= set(active_week_days) and
                    ticket_meta.at[t, "sources"] == {"ORIGINAL"} and
                    ticket_meta.at[t, "cias"] <= {"JJ", "AD"}]
        if not eligible:
            continue
        tickets = pd.Index(eligible)

        # 2. Universos de proteção (Guard): LATAM >= GOL + 3 pp no total, em cada ADVP (Curto, Longo) e no dia a dia
        guard_specs = []
        guard_specs.append(("SEMANA", None, None, None, f"{kind} SEMANA TOTAL"))
        if include_complete:
            guard_specs.append(("COMPLETO", None, None, None, f"{kind} COMPLETO TOTAL"))
        for name, bounds in groups.items():
            guard_specs.append(("SEMANA", bounds, None, None, f"{kind} SEMANA ADVP {name}"))
            if include_complete:
                guard_specs.append(("COMPLETO", bounds, None, None, f"{kind} COMPLETO ADVP {name}"))
        for d in active_week_days:
            guard_specs.append(("DAY", None, {d}, None, f"{kind} DIA {d} TOTAL"))
            if kind in (advp_kinds or set()) or (advp_kinds is None and groups):
                for name, bounds in groups.items():
                    guard_specs.append(("DAY", bounds, {d}, None, f"{kind} DIA {d} ADVP {name}"))

        all_specs = target_specs + guard_specs
        num_targets = len(target_specs)
        num_total = len(all_specs)

        totals = []
        effects = []
        for period_type, bounds, day_filter, _, _ in all_specs:
            if period_type == "DAY":
                rows = report_rows(current, kind, dates=day_filter, advp=bounds, filters=filters)
            elif period_type == "SEMANA":
                rows = report_rows(current, kind, dates=week, advp=bounds, filters=filters)
            else:  # COMPLETO
                rows = report_rows(current, kind, dates=None, advp=bounds, filters=filters)

            values = rows["VALOR_BRL"].map(legacy.converter_numero_brasileiro).fillna(0)
            companies = rows["CIA_BILHETE"].fillna("").astype(str).str.strip().str.upper()
            rows = rows.assign(_VALUE=values, _CIA=companies, _TICKET=rows["BILHETE"].astype(str))
            sums = rows.groupby("_CIA")["_VALUE"].sum()
            total = np.array([float(sums.get(c, 0)) for c in CIASS])
            pivot = rows.pivot_table(index="_TICKET", columns="_CIA", values="_VALUE", aggfunc="sum", fill_value=0)
            effect = pivot.reindex(index=tickets, columns=CIASS, fill_value=0).to_numpy(dtype=float)
            totals.append(total)
            effects.append(effect)

        base = np.stack(totals)  # shape (num_total, 3)
        effect = np.stack(effects, axis=1)  # shape (num_tickets, num_total, 3)
        active = np.ones(len(tickets), dtype=bool)

        base_den = base.sum(axis=1)
        base_jj = np.divide(base[:, 0], base_den, out=np.zeros_like(base_den), where=base_den > 0)
        base_g3 = np.divide(base[:, 1], base_den, out=np.zeros_like(base_den), where=base_den > 0)
        base_ad = np.divide(base[:, 2], base_den, out=np.zeros_like(base_den), where=base_den > 0)
        base_gap = base_jj - base_g3

        def scores(candidate_totals):
            den = candidate_totals.sum(axis=-1)
            jj = np.divide(candidate_totals[..., 0], den, out=np.zeros_like(den), where=den > 0)
            g3 = np.divide(candidate_totals[..., 1], den, out=np.zeros_like(den), where=den > 0)
            ad = np.divide(candidate_totals[..., 2], den, out=np.zeros_like(den), where=den > 0)
            gap = jj - g3

            valid = np.ones(candidate_totals.shape[:-2], dtype=bool) if candidate_totals.ndim == 3 else np.array([True])

            # Hard constraints para todos os universos
            for u in range(num_total):
                if base_den[u] <= 0:
                    continue
                valid &= (den[..., u] > 0)
                # Proteção LATAM: no mínimo 3 pp à frente da GOL (ou preservação da folga inicial se menor)
                if base_gap[u] >= 0.03 - 1e-7:
                    valid &= (gap[..., u] >= 0.03 - 1e-7)
                elif base_gap[u] >= 0:
                    valid &= (gap[..., u] >= base_gap[u] - 1e-7) & (jj[..., u] >= g3[..., u])
                else:
                    valid &= (gap[..., u] >= base_gap[u] - 1e-7)

            # Limite de queda do share da Azul no total
            for u in range(num_total):
                if all_specs[u][0] in ("SEMANA", "COMPLETO") and all_specs[u][1] is None and base_den[u] > 0:
                    valid &= (ad[..., u] >= base_ad[u] - ad_drop_pp / 100 - 1e-7)

            objective = np.zeros(candidate_totals.shape[0])
            for j in range(num_targets):
                target = target_specs[j][3]
                bounds = target_specs[j][1]
                err = np.abs(g3[..., j] - target)
                if bounds is not None:
                    overshoot = np.maximum(0, g3[..., j] - target)
                    objective += err + 2.5 * overshoot
                else:
                    objective += err
                objective += 10.0 * np.maximum(0, 0.03 - gap[..., j])

            for u in range(num_targets, num_total):
                if base_den[u] > 0:
                    objective += 1.0 * np.maximum(0, 0.03 - gap[..., u])

            objective = np.where(valid, objective, np.inf)
            return objective

        current_score = float(scores(base[None, :, :])[0])
        chosen = []
        while active.any():
            candidates = base[None, :, :] - effect
            candidate_scores = scores(candidates)
            candidate_scores[~active] = np.inf
            best = int(np.argmin(candidate_scores))
            if not np.isfinite(candidate_scores[best]) or candidate_scores[best] >= current_score - 1e-10:
                break
            current_score = float(candidate_scores[best])
            base -= effect[best]
            chosen.append(tickets[best])
            active[best] = False

        if chosen:
            mask = full_ticket.isin(chosen)
            part = current.loc[mask].copy()
            part["RELATORIO_AJUSTE"] = kind
            part["DIA_EXCLUSAO"] = pd.to_datetime(part["DATA_EMISSAO"]).dt.strftime("%Y-%m-%d")
            removed_parts.append(part)
            current = current.loc[~mask].copy()

    removed = pd.concat(removed_parts, ignore_index=True) if removed_parts else pd.DataFrame(columns=list(df) + ["RELATORIO_AJUSTE", "DIA_EXCLUSAO"])
    return current, removed, original, report_matrix(current, week, groups, advp_kinds, filters, include_complete)


def daily_comparison_tables(before_df: pd.DataFrame, after_df: pd.DataFrame,
                            week: set[str], groups: dict,
                            advp_kinds: set[str] | None = None,
                            filters: dict | None = None) -> list[dict]:
    """Gera as tabelas comparativas dia a dia de share antes x depois e diferenças.

    Cobre por tipo de relatório:
    - CONS: ADVP CURTO por dia, ADVP LONGO por dia, ADVP TOTAL por dia
    - GOV: ADVP TOTAL por dia
    - CORP: ADVP TOTAL por dia
    (e faixas de ADVP extras se configuradas).
    """
    scoped = set(REPORTS) if advp_kinds is None else set(advp_kinds)
    tables = []
    days = sorted(week)

    def band_label(name):
        if name == "GERAL":
            return "ADVP TOTAL"
        lo, hi = groups[name]
        return f"ADVP {lo} A {hi}" if hi is not None else f"ADVP {lo}+"

    def calc_metrics(df_sub):
        val = df_sub["VALOR_BRL"].map(legacy.converter_numero_brasileiro).fillna(0)
        cia = df_sub["CIA_BILHETE"].fillna("").astype(str).str.strip().str.upper()
        cia_vals = {c: float(val.loc[cia.eq(c)].sum()) for c in CIASS}
        tot = sum(cia_vals.values())
        shares = {c: (cia_vals[c] / tot) if tot > 0 else 0.0 for c in CIASS}
        return cia_vals, tot, shares

    for kind in REPORTS:
        bands = []
        if kind in scoped and groups:
            for name in groups:
                bands.append((name, groups[name]))
        bands.append(("GERAL", None))

        for name, bounds in bands:
            title = f"{kind} — {band_label(name)} POR DIA"
            rows = []
            for d in days:
                d_before = report_rows(before_df, kind, dates={d}, advp=bounds, filters=filters)
                d_after = report_rows(after_df, kind, dates={d}, advp=bounds, filters=filters)
                v_b, t_b, s_b = calc_metrics(d_before)
                v_a, t_a, s_a = calc_metrics(d_after)

                dt = pd.to_datetime(d, errors="coerce")
                date_label = dt.strftime("%d/%m/%Y") if pd.notna(dt) else d

                row = {
                    "date": d,
                    "date_label": date_label,
                    "is_total": False,
                    "jj": {
                        "orig_share": s_b["JJ"],
                        "curr_share": s_a["JJ"],
                        "diff_pp": (s_a["JJ"] - s_b["JJ"]) * 100,
                        "orig_val": v_b["JJ"],
                        "curr_val": v_a["JJ"],
                        "diff_val": v_a["JJ"] - v_b["JJ"],
                    },
                    "g3": {
                        "orig_share": s_b["G3"],
                        "curr_share": s_a["G3"],
                        "diff_pp": (s_a["G3"] - s_b["G3"]) * 100,
                        "orig_val": v_b["G3"],
                        "curr_val": v_a["G3"],
                        "diff_val": v_a["G3"] - v_b["G3"],
                    },
                    "ad": {
                        "orig_share": s_b["AD"],
                        "curr_share": s_a["AD"],
                        "diff_pp": (s_a["AD"] - s_b["AD"]) * 100,
                        "orig_val": v_b["AD"],
                        "curr_val": v_a["AD"],
                        "diff_val": v_a["AD"] - v_b["AD"],
                    },
                    "total": {
                        "orig_val": t_b,
                        "curr_val": t_a,
                        "diff_val": t_a - t_b,
                    }
                }
                rows.append(row)

            # Linha TOTAL DA SEMANA
            w_before = report_rows(before_df, kind, dates=week, advp=bounds, filters=filters)
            w_after = report_rows(after_df, kind, dates=week, advp=bounds, filters=filters)
            wv_b, wt_b, ws_b = calc_metrics(w_before)
            wv_a, wt_a, ws_a = calc_metrics(w_after)

            total_row = {
                "date": "TOTAL",
                "date_label": "TOTAL DA SEMANA",
                "is_total": True,
                "jj": {
                    "orig_share": ws_b["JJ"],
                    "curr_share": ws_a["JJ"],
                    "diff_pp": (ws_a["JJ"] - ws_b["JJ"]) * 100,
                    "orig_val": wv_b["JJ"],
                    "curr_val": wv_a["JJ"],
                    "diff_val": wv_a["JJ"] - wv_b["JJ"],
                },
                "g3": {
                    "orig_share": ws_b["G3"],
                    "curr_share": ws_a["G3"],
                    "diff_pp": (ws_a["G3"] - ws_b["G3"]) * 100,
                    "orig_val": wv_b["G3"],
                    "curr_val": wv_a["G3"],
                    "diff_val": wv_a["G3"] - wv_b["G3"],
                },
                "ad": {
                    "orig_share": ws_b["AD"],
                    "curr_share": ws_a["AD"],
                    "diff_pp": (ws_a["AD"] - ws_b["AD"]) * 100,
                    "orig_val": wv_b["AD"],
                    "curr_val": wv_a["AD"],
                    "diff_val": wv_a["AD"] - wv_b["AD"],
                },
                "total": {
                    "orig_val": wt_b,
                    "curr_val": wt_a,
                    "diff_val": wt_a - wt_b,
                }
            }
            rows.append(total_row)

            tables.append({
                "kind": kind,
                "advp": name,
                "title": title,
                "rows": rows,
            })

    return tables


def daily_comparison_dataframe(tables: list[dict]) -> pd.DataFrame:
    rows = []
    for t in tables:
        for r in t["rows"]:
            rows.append({
                "RELATORIO": t["kind"],
                "TABELA": t["title"],
                "DATA": r["date_label"],
                "JJ_SHARE_ORIGINAL": r["jj"]["orig_share"],
                "JJ_SHARE_ATUAL": r["jj"]["curr_share"],
                "JJ_DIF_PP": r["jj"]["diff_pp"],
                "G3_SHARE_ORIGINAL": r["g3"]["orig_share"],
                "G3_SHARE_ATUAL": r["g3"]["curr_share"],
                "G3_DIF_PP": r["g3"]["diff_pp"],
                "AD_SHARE_ORIGINAL": r["ad"]["orig_share"],
                "AD_SHARE_ATUAL": r["ad"]["curr_share"],
                "AD_DIF_PP": r["ad"]["diff_pp"],
                "TOTAL_ORIGINAL_BRL": r["total"]["orig_val"],
                "TOTAL_ATUAL_BRL": r["total"]["curr_val"],
                "TOTAL_DIF_BRL": r["total"]["diff_val"],
            })
    return pd.DataFrame(rows)


def write_zip(path: Path, original_path: Path, formatted: pd.DataFrame,
              full_before_adjustment: pd.DataFrame,
              excluded: pd.DataFrame, coupon_audit: pd.DataFrame,
              final: pd.DataFrame, removed: pd.DataFrame,
              before: list[dict], after: list[dict], week: set[str], stats: dict,
              groups: dict | None = None, advp_kinds: set[str] | None = None,
              filters: dict | None = None, include_complete: bool = True,
              complement_path: Path | None = None,
              date_source_path: Path | None = None):
    out = path.parent / "gol_saida"
    out.mkdir(exist_ok=True)
    for old in out.iterdir():
        if old.is_file():
            old.unlink()
    def csv(frame, filename):
        tidy(frame).to_csv(out / filename, index=False, sep=";", encoding="utf-8-sig", date_format="%d/%m/%Y")
    csv(formatted, "GOL_FORMATADO.csv")
    if include_complete and full_before_adjustment is not None:
        csv(full_before_adjustment.drop(columns="_SOURCE_PERIOD", errors="ignore"), "GOL_FORMATADO_COMPLETO.csv")

    def export_gol_20(frame, filename):
        export = tidy(frame.reindex(columns=legacy.COLUNAS_FINAIS, fill_value=""))
        for column in ("DATA_EMISSAO", "DATA_EMBARQUE"):
            export[column] = pd.to_datetime(export[column], errors="coerce", dayfirst=True).dt.strftime("%d/%m/%Y").fillna("")
        export["VALOR_BRL"] = export["VALOR_BRL"].map(
            lambda v: "" if pd.isna(v) else f"{float(v):.2f}".replace(".", ","))
        export["IATA_AGENCIA"] = pd.array(export["IATA_AGENCIA"], dtype="Int64")
        export.to_csv(out / filename, index=False, sep=";", encoding="utf-8-sig")

    # Arquivo principal (período atual / semana de ajuste após o ajuste de share)
    is_original = final.get("_SOURCE_PERIOD", pd.Series("ORIGINAL", index=final.index)).eq("ORIGINAL")
    final_principal = final.loc[is_original].copy()
    export_gol_20(final_principal, "GOL FINAL.csv")
    export_gol_20(final_principal, "GOL PRINCIPAL.csv")

    # Base completa (base dos demais períodos + base do período atual após o ajuste)
    export_gol_20(final, "GOL BASE COMPLETA.csv")
    export_gol_20(final, "GOL FINAL COMPLETO.csv")
    csv(excluded, "GOL_EXCLUIDOS_FORMATACAO.csv")
    csv(coupon_audit, "GOL_CUPONS_CORRIGIDOS.csv")
    csv(removed, "GOL_VENDAS_EXCLUIDAS_SHARE.csv")
    for day, part in removed.groupby("DIA_EXCLUSAO"):
        csv(part, f"GOL_EXCLUIDAS_{day}.csv")
    daily_tables = daily_comparison_tables(full_before_adjustment, final, week, groups or {}, advp_kinds, filters)
    csv(daily_comparison_dataframe(daily_tables), "GOL_COMPARATIVO_DIA_A_DIA.csv")
    def flat(records):
        return pd.DataFrame([{
            "RELATORIO": r["kind"], "PERIODO": r["period"], "ADVP": r["advp"],
            "LINHAS": r["rows"], "TOTAL_BRL": r["total"],
            **{f"{c}_BRL": r["values"][c] for c in CIASS},
            **{f"{c}_SHARE": r["shares"][c] for c in CIASS},
            "GAP_JJ_G3_PP": r["gap_pp"],
        } for r in records])
    with pd.ExcelWriter(out / "GOL_FORMATADO_COM_DINAMICA.xlsx", engine="xlsxwriter", datetime_format="dd/mm/yyyy") as writer:
        book = writer.book
        heading = book.add_format({"bold": True, "font_size": 13, "font_name": "Arial"})
        label_head = book.add_format({"bold": True, "bg_color": "#DCE6F1", "border": 1,
                                      "align": "center", "valign": "vcenter", "font_name": "Arial"})
        subhead = book.add_format({"bold": True, "bg_color": "#F0F2F4", "border": 1,
                                   "align": "center", "font_name": "Arial"})
        band_formats = {
            "CONS": book.add_format({"bold": True, "bg_color": "#ECEAE3", "border": 1,
                                     "align": "center", "font_name": "Arial"}),
            "CORP": book.add_format({"bold": True, "bg_color": "#DCE6F1", "border": 1,
                                     "align": "center", "font_name": "Arial"}),
            "GOV": book.add_format({"bold": True, "bg_color": "#E9EFE4", "border": 1,
                                    "align": "center", "font_name": "Arial"}),
        }
        money = book.add_format({"num_format": '"R$ "#,##0.00', "border": 1, "font_name": "Arial"})
        pct = book.add_format({"num_format": "0.00%", "border": 1, "font_name": "Arial"})
        money_total = book.add_format({"num_format": '"R$ "#,##0.00', "bold": True,
                                        "bg_color": "#E5ECF2", "border": 1, "font_name": "Arial"})
        pct_total = book.add_format({"num_format": "0.00%", "bold": True,
                                      "bg_color": "#E5ECF2", "border": 1, "font_name": "Arial"})
        cia_fmt = book.add_format({"border": 1, "font_name": "Arial"})
        cia_total = book.add_format({"bold": True, "bg_color": "#E5ECF2", "border": 1,
                                     "font_name": "Arial"})
        def pivot_sheet(name, title, tables):
            ws = book.add_worksheet(name)
            writer.sheets[name] = ws
            ws.hide_gridlines(2)
            ws.set_column(0, 0, 21)
            ws.set_column(1, 14, 19)
            ws.write(0, 0, title, heading)
            top = 2
            for table in tables:
                ws.write(top, 0, table["title"], heading)
                head = top + 1
                ws.merge_range(head, 0, head + 1, 0, "CIA_BILHETE", label_head)
                for j, band in enumerate(table["bands"]):
                    col = 1 + j * 2
                    ws.merge_range(head, col, head, col + 1,
                                   band["label"], band_formats[band["kind"]])
                    ws.write(head + 1, col, "TARIFA", subhead)
                    ws.write(head + 1, col + 1, "SHARE", subhead)
                for i, row in enumerate(table["rows"]):
                    y = head + 2 + i
                    is_total = row["cia"] == "Total geral"
                    ws.write(y, 0, row["cia"], cia_total if is_total else cia_fmt)
                    for j, cell in enumerate(row["cells"]):
                        col = 1 + j * 2
                        ws.write_number(y, col, cell["value"], money_total if is_total else money)
                        ws.write_number(y, col + 1, cell["share"], pct_total if is_total else pct)
                top = head + 7 + 2
            return ws

        def daily_sheet(tables):
            ws = book.add_worksheet("DIA_A_DIA")
            writer.sheets["DIA_A_DIA"] = ws
            ws.hide_gridlines(2)
            ws.set_column(0, 0, 20)
            for c in range(1, 13):
                ws.set_column(c, c, 15)
            ws.write(0, 0, "RELATÓRIO DIA A DIA — COMPARATIVO ANTES × DEPOIS", heading)
            top = 2
            diff_pct = book.add_format({"num_format": "+0.00%;-0.00%;0.00%", "border": 1, "font_name": "Arial"})
            diff_pct_total = book.add_format({"num_format": "+0.00%;-0.00%;0.00%", "bold": True,
                                              "bg_color": "#E5ECF2", "border": 1, "font_name": "Arial"})
            diff_money = book.add_format({"num_format": '"+R$ "#,##0.00;"-R$ "#,##0.00;"R$ "0.00', "border": 1, "font_name": "Arial"})
            diff_money_total = book.add_format({"num_format": '"+R$ "#,##0.00;"-R$ "#,##0.00;"R$ "0.00', "bold": True,
                                                "bg_color": "#E5ECF2", "border": 1, "font_name": "Arial"})
            for table in tables:
                ws.write(top, 0, table["title"], heading)
                head = top + 1
                ws.merge_range(head, 0, head + 1, 0, "DATA", label_head)
                ws.merge_range(head, 1, head, 3, "LATAM (JJ)", band_formats["CORP"])
                ws.merge_range(head, 4, head, 6, "GOL (G3)", band_formats["CONS"])
                ws.merge_range(head, 7, head, 9, "AZUL (AD)", band_formats["GOV"])
                ws.merge_range(head, 10, head, 12, "TOTAL GERAL", subhead)

                sub_labels = [
                    "SHARE ORIG.", "SHARE ATUAL", "DIF. (pp)",
                    "SHARE ORIG.", "SHARE ATUAL", "DIF. (pp)",
                    "SHARE ORIG.", "SHARE ATUAL", "DIF. (pp)",
                    "TARIFA ORIG.", "TARIFA ATUAL", "DIF. (R$)",
                ]
                for j, lbl in enumerate(sub_labels):
                    ws.write(head + 1, 1 + j, lbl, subhead)

                for i, row in enumerate(table["rows"]):
                    y = head + 2 + i
                    is_tot = row["is_total"]
                    r_fmt = cia_total if is_tot else cia_fmt
                    p_fmt = pct_total if is_tot else pct
                    dp_fmt = diff_pct_total if is_tot else diff_pct
                    m_fmt = money_total if is_tot else money
                    dm_fmt = diff_money_total if is_tot else diff_money

                    ws.write(y, 0, row["date_label"], r_fmt)
                    ws.write_number(y, 1, row["jj"]["orig_share"], p_fmt)
                    ws.write_number(y, 2, row["jj"]["curr_share"], p_fmt)
                    ws.write_number(y, 3, row["jj"]["diff_pp"] / 100, dp_fmt)

                    ws.write_number(y, 4, row["g3"]["orig_share"], p_fmt)
                    ws.write_number(y, 5, row["g3"]["curr_share"], p_fmt)
                    ws.write_number(y, 6, row["g3"]["diff_pp"] / 100, dp_fmt)

                    ws.write_number(y, 7, row["ad"]["orig_share"], p_fmt)
                    ws.write_number(y, 8, row["ad"]["curr_share"], p_fmt)
                    ws.write_number(y, 9, row["ad"]["diff_pp"] / 100, dp_fmt)

                    ws.write_number(y, 10, row["total"]["orig_val"], m_fmt)
                    ws.write_number(y, 11, row["total"]["curr_val"], m_fmt)
                    ws.write_number(y, 12, row["total"]["diff_val"], dm_fmt)

                top = head + 2 + len(table["rows"]) + 2
            return ws

        periods = ("SEMANA", "COMPLETO") if include_complete else ("SEMANA",)
        for period in periods:
            pivot_sheet(f"DINAMICA_{period}", f"Relatórios {period.lower()}",
                        report_layouts(before, groups or {}, advp_kinds)[period])
            pivot_sheet(f"AJUSTADO_{period}", f"Relatórios {period.lower()} após ajuste",
                        report_layouts(after, groups or {}, advp_kinds)[period])
        daily_sheet(daily_tables)
        tidy(formatted).to_excel(writer, sheet_name="BASE_FORMATADA", index=False)
        if include_complete:
            tidy(full_before_adjustment.drop(columns="_SOURCE_PERIOD", errors="ignore")).to_excel(
                writer, sheet_name="BASE_COMPLETA", index=False)
        flat(before).to_excel(writer, sheet_name="RELATORIOS_COMPLETO_SEMANA" if include_complete else "RELATORIOS_SEMANA", index=False)
        flat(after).to_excel(writer, sheet_name="RELATORIOS_APOS_AJUSTE", index=False)
        for kind in REPORTS:
            report_rows(final, kind, week, filters=filters).drop(columns="_SOURCE_PERIOD", errors="ignore").to_excel(
                writer, sheet_name=f"{kind}_SEMANA", index=False)
            if include_complete:
                report_rows(final, kind, filters=filters).drop(columns="_SOURCE_PERIOD", errors="ignore").to_excel(
                    writer, sheet_name=f"{kind}_COMPLETO", index=False)
    (out / "GOL_RESUMO.json").write_text(json.dumps({"semana_ajuste": sorted(week), "estatisticas": stats,
        "antes": before, "depois": after}, ensure_ascii=False, indent=2), encoding="utf-8")
    write_report_pdf(out / "GOL_RELATORIOS.pdf", full_before_adjustment, before, after,
                     week, stats, len(final), len(removed), groups or {}, advp_kinds,
                     filters, include_complete, daily_tables=daily_tables)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(original_path, f"01_ORIGINAL/{original_path.name}")
        if complement_path is not None:
            z.write(complement_path, f"01_ORIGINAL/{complement_path.name}")
        if date_source_path is not None:
            z.write(date_source_path, f"01_ORIGINAL/{date_source_path.name}")
        for p in out.iterdir():
            z.write(p, p.name)
    return path


def write_report_pdf(path: Path, full: pd.DataFrame, before: list[dict],
                     after: list[dict], week: set[str], stats: dict,
                     final_rows: int, removed_rows: int, groups: dict,
                     advp_kinds: set[str] | None, filters: dict | None,
                     include_complete: bool,
                     daily_tables: list[dict] | None = None) -> None:
    """PDF de conferência com indicadores e shares antes/depois."""
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, KeepTogether, PageBreak

    styles = getSampleStyleSheet()
    styles["Title"].fontName = "Helvetica-Bold"
    styles["Title"].fontSize = 16
    styles["Heading2"].fontSize = 11
    styles["Heading2"].spaceBefore = 12
    styles["Normal"].fontSize = 8
    styles["Normal"].leading = 11
    doc = SimpleDocTemplate(str(path), pagesize=(297*mm, 210*mm),
                            leftMargin=13*mm, rightMargin=13*mm,
                            topMargin=14*mm, bottomMargin=13*mm)
    story = [Paragraph("GOL - Relatório de tratamento e share", styles["Title"]), Spacer(1, 6*mm)]
    labels = [
        ("original_rows", "Original"), ("formatted_rows", "Formatadas"),
        ("excluded_rows", "Excluídas"), ("gol_artmax_rows", "GOL em ART/MAX"),
        ("gol_artmax_iata_255_rows", "GOL em ART/MAX com IATA final 255"),
        ("cgb_rows", "Redirecionadas para CGB"), ("artmax_rows", "ART/MAX no original"),
    ]
    metrics = [[label, f"{int(stats.get(key, 0)):,}".replace(",", ".")]
               for key, label in labels]
    metrics += [["Excluídas no ajuste", f"{removed_rows:,}".replace(",", ".")],
                ["Linhas na base final", f"{final_rows:,}".replace(",", ".")]]
    metric_table = Table(metrics, colWidths=[78*mm, 28*mm], hAlign="LEFT")
    metric_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F5F8FB")),
        ("GRID", (0, 0), (-1, -1), .35, colors.HexColor("#CBD5DF")),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTNAME", (1, 0), (1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story += [metric_table, Spacer(1, 5*mm)]
    titles = period_titles(full, week, filters, include_complete)
    def money(value):
        return "R$ " + f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    for stage, records in (("ORIGINAL - ANTES DO AJUSTE", before),
                           ("AJUSTADO - DEPOIS DO AJUSTE", after)):
        if stage.startswith("AJUSTADO"):
            story.append(PageBreak())
        story.append(Paragraph(stage, styles["Heading2"]))
        layouts = report_layouts(records, groups, advp_kinds)
        for period, title in titles.items():
            story.append(Paragraph(title, styles["Heading2"]))
            for section in layouts[period]:
                heading = ["CIA_BILHETE"] + [f"{band['label']} - {field}"
                    for band in section["bands"] for field in ("TARIFA", "SHARE")]
                data = [heading]
                for row in section["rows"]:
                    data.append([row["cia"]] + [item for cell in row["cells"]
                        for item in (money(cell["value"]), f"{cell['share']*100:.2f}%".replace(".", ","))])
                available = 271*mm
                first_width = 28*mm
                other_width = (available-first_width)/(len(heading)-1)
                table = Table(data, colWidths=[first_width]+[other_width]*(len(heading)-1), repeatRows=1)
                table.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DCE6F1")),
                    ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E5ECF2")),
                    ("GRID", (0, 0), (-1, -1), .3, colors.HexColor("#9AABBC")),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), 6.7),
                    ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]))
                story.append(KeepTogether([Paragraph(section["title"], styles["Heading2"]), table]))
        story.append(Spacer(1, 5*mm))

    if daily_tables:
        story.append(PageBreak())
        story.append(Paragraph("RELATÓRIO DIA A DIA — COMPARATIVO ANTES × DEPOIS E DIFERENÇA", styles["Heading2"]))
        for t in daily_tables:
            data = [
                ["DATA", "JJ ORIG", "JJ ATUAL", "JJ DIF", "G3 ORIG", "G3 ATUAL", "G3 DIF",
                 "AD ORIG", "AD ATUAL", "AD DIF", "TOT ORIG", "TOT ATUAL", "TOT DIF"]
            ]
            for r in t["rows"]:
                data.append([
                    r["date_label"],
                    f"{r['jj']['orig_share']*100:.2f}%", f"{r['jj']['curr_share']*100:.2f}%", f"{r['jj']['diff_pp']:+.2f}pp",
                    f"{r['g3']['orig_share']*100:.2f}%", f"{r['g3']['curr_share']*100:.2f}%", f"{r['g3']['diff_pp']:+.2f}pp",
                    f"{r['ad']['orig_share']*100:.2f}%", f"{r['ad']['curr_share']*100:.2f}%", f"{r['ad']['diff_pp']:+.2f}pp",
                    money(r['total']['orig_val']), money(r['total']['curr_val']), money(r['total']['diff_val'])
                ])
            col_widths = [25*mm] + [18*mm]*9 + [28*mm]*3
            d_table = Table(data, colWidths=col_widths, repeatRows=1)
            d_table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DCE6F1")),
                ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E5ECF2")),
                ("GRID", (0, 0), (-1, -1), .3, colors.HexColor("#9AABBC")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 6.2),
                ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]))
            story.append(KeepTogether([Paragraph(t["title"], styles["Heading2"]), d_table]))
            story.append(Spacer(1, 4*mm))

    doc.build(story)
