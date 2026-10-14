
from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from pathlib import Path
from werkzeug.utils import secure_filename
import uuid
import json
import shutil
import pandas as pd
import re

from gol_processor import (
    read_table, format_gol, merge_period, prepare_complement, recover_emission_dates, week_dates,
    filter_options as gol_filter_options, default_filters as gol_default_filters,
    period_titles as gol_period_titles, report_matrix, report_layouts,
    adjust_shares, write_zip, daily_comparison_tables,
)
from gol_legacy import IATAS_CORP

from processor import (
    read_csv_robust,
    stage1_format,
    get_filter_options,
    apply_user_filters,
    build_share_preview,
    run_share_adjustment,
    build_zip_outputs,
)

BASE_DIR = Path(__file__).resolve().parent
JOBS_DIR = BASE_DIR / "jobs"
JOBS_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.secret_key = "troque-esta-chave-em-producao"
app.config["MAX_CONTENT_LENGTH"] = 300 * 1024 * 1024


def load_job(job_id):
    if not re.fullmatch(r"[0-9a-f]{32}", job_id):
        return None, None
    job_dir = JOBS_DIR / job_id
    if not job_dir.exists():
        return None, None
    meta_path = job_dir / "meta.json"
    if not meta_path.exists():
        return None, None
    meta = json.loads(meta_path.read_text(encoding="utf-8"))
    return job_dir, meta


@app.get("/")
def index():
    return render_template("home.html")


def gol_page(job_id=None, step=1, stats=None, matrix=None, week=None,
             groups=None, result=None, corp_text=None, advp_kinds=None,
             filter_options=None, current_filters=None, titles=None,
             has_complement=False, complement_status=None, complement_stats=None,
             daily_tables=None):
    groups = groups or {}
    advp_kinds = set(advp_kinds or [])
    return render_template("gol.html", job_id=job_id, step=step,
        stats=stats, matrix=matrix, week=week or [], groups=groups,
        layouts=report_layouts(matrix, groups, advp_kinds) if matrix else {},
        advp_kinds=advp_kinds, result=result,
        filter_options=filter_options, current_filters=current_filters,
        titles=titles or {}, has_complement=has_complement,
        complement_status=complement_status, complement_stats=complement_stats,
        corp_text=corp_text or ", ".join(sorted(IATAS_CORP)),
        daily_tables=daily_tables or [])


def gol_results_page(job_id, job_dir, meta, step=3, matrix=None, result=None, daily_tables=None):
    full = pd.read_pickle(job_dir / "full.pkl")
    groups = {k: tuple(v) for k, v in meta.get("groups", {}).items()}
    scope = set(meta.get("advp_kinds", []))
    filters = meta["filters"]
    has_complement = bool(meta.get("complementar_name"))
    if matrix is None:
        matrix = report_matrix(full, set(meta["week"]), groups, scope,
                               filters, has_complement)
    return gol_page(job_id, step, meta["stats"], matrix, meta["week"], groups,
                    result=result, corp_text=meta["corp_text"], advp_kinds=scope,
                    filter_options=gol_filter_options(full), current_filters=filters,
                    titles=gol_period_titles(full, set(meta["week"]), filters, has_complement),
                    has_complement=has_complement,
                    complement_status=meta.get("complement_status"),
                    complement_stats=meta.get("complement_stats"),
                    daily_tables=daily_tables)


@app.template_filter("brl")
def brl(value):
    return "R$ " + f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


@app.get("/gol")
def gol():
    return gol_page()


@app.post("/gol/formatar")
def gol_formatar():
    original = request.files.get("arquivo")
    if not original or not original.filename:
        flash("Envie a base GOL.")
        return redirect(url_for("gol"))
    job_id = uuid.uuid4().hex
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir()
    try:
        original_path = job_dir / ("ORIGINAL_" + secure_filename(original.filename))
        original.save(original_path)
        corp_text = request.form.get("corp_iatas", "")
        corp = {int(x) for x in re.findall(r"\d+", corp_text)}
        formatted, excluded, coupons, stats = format_gol(read_table(original_path), corp)
        formatted.to_pickle(job_dir / "formatted.pkl")
        excluded.to_pickle(job_dir / "excluded.pkl")
        coupons.to_pickle(job_dir / "coupons.pkl")
        meta = {"original_name": original_path.name,
                "stats": stats, "week": week_dates(formatted), "corp_text": corp_text}
        (job_dir / "meta.json").write_text(json.dumps(meta, ensure_ascii=False), encoding="utf-8")
        return gol_page(job_id, 2, stats, week=meta["week"], corp_text=corp_text)
    except Exception as exc:
        shutil.rmtree(job_dir, ignore_errors=True)
        flash(f"Erro ao formatar GOL: {exc}")
        return redirect(url_for("gol"))


@app.post("/gol/relatorios/<job_id>")
def gol_relatorios(job_id):
    job_dir, meta = load_job(job_id)
    if not job_dir or not (job_dir / "formatted.pkl").exists():
        flash("Sessão GOL não encontrada.")
        return redirect(url_for("gol"))
    try:
        current = pd.read_pickle(job_dir / "formatted.pkl")
        current["_SOURCE_PERIOD"] = "ORIGINAL"
        extra_file = request.files.get("complementar")
        if extra_file and extra_file.filename:
            extra_path = job_dir / ("COMPLEMENTAR_" + secure_filename(extra_file.filename))
            extra_file.save(extra_path)
            extra_raw = read_table(extra_path)
            date_source = request.files.get("origem_datas")
            if date_source and date_source.filename:
                source_path = job_dir / ("ORIGEM_DATAS_" + secure_filename(date_source.filename))
                date_source.save(source_path)
                extra_raw = recover_emission_dates(extra_raw, read_table(source_path))
                meta["origem_datas_name"] = source_path.name
            elif "DATA_EMISSAO" not in {str(c).strip().upper() for c in extra_raw.columns}:
                raise ValueError("A base complementar não contém DATA_EMISSAO. Envie também o CSV original com essa coluna no campo 'Origem das datas'.")
            else:
                meta.pop("origem_datas_name", None)
            extra, status, extra_stats = prepare_complement(
                extra_raw, {int(x) for x in re.findall(r"\d+", meta["corp_text"])})
            full = merge_period(current, extra)
            meta["complementar_name"] = extra_path.name
            meta["complement_status"] = status
            meta["complement_stats"] = extra_stats
        else:
            full = current.copy()
            meta.pop("complementar_name", None)
            meta.pop("complement_status", None)
            meta.pop("complement_stats", None)
            meta.pop("origem_datas_name", None)
        groups = {}
        advp_kinds = set()
        if request.form.get("use_advp") == "S":
            scope = request.form.get("advp_scope", "CONS").upper()
            if scope not in {"CONS", "CORP", "GOV", "TODOS"}:
                raise ValueError("Escolha um tipo de relatório para o ADVP.")
            advp_kinds = {"CONS", "CORP", "GOV"} if scope == "TODOS" else {scope}
            short_start = int(request.form.get("short_start", "0"))
            short_end = int(request.form.get("short_end", "27"))
            if short_start < 0 or short_end < short_start:
                raise ValueError("Faixa ADVP curto inválida.")
            groups["CURTO"] = (short_start, short_end)
            long_start = short_end + 1
            if request.form.get("use_intermediate") == "S":
                middle_start = int(request.form["middle_start"])
                middle_end = int(request.form["middle_end"])
                if middle_start != long_start or middle_end < middle_start:
                    raise ValueError("O ADVP intermediário deve começar logo após o curto.")
                groups["INTERMEDIARIO"] = (middle_start, middle_end)
                long_start = middle_end + 1
            groups["LONGO"] = (long_start, None)
        full.to_pickle(job_dir / "full.pkl")
        meta["groups"] = groups
        meta["advp_kinds"] = sorted(advp_kinds)
        meta["filters"] = gol_default_filters(full)
        (job_dir / "meta.json").write_text(json.dumps(meta, ensure_ascii=False), encoding="utf-8")
        return gol_results_page(job_id, job_dir, meta)
    except Exception as exc:
        flash(f"Erro nos relatórios GOL: {exc}")
        return gol_page(job_id, 2, meta["stats"], week=meta["week"], corp_text=meta["corp_text"])


@app.post("/gol/filtros/<job_id>")
def gol_filtros(job_id):
    job_dir, meta = load_job(job_id)
    if not job_dir or not (job_dir / "full.pkl").exists():
        flash("Prepare a base GOL primeiro.")
        return redirect(url_for("gol"))
    try:
        options = gol_filter_options(pd.read_pickle(job_dir / "full.pkl"))
        selected = {}
        for name, key in (("cupons", "cupons"), ("tipos", "tipos")):
            values = request.form.getlist(name)
            if any(v not in options[key] for v in values):
                raise ValueError(f"Valor inválido no filtro {name}.")
            selected[name] = values or options[key]
        selected["bases"] = {}
        for kind in ("CONS", "CORP", "GOV"):
            values = request.form.getlist(f"bases_{kind}")
            if any(v not in options["bases"][kind] for v in values):
                raise ValueError(f"BASE inválida no relatório {kind}.")
            selected["bases"][kind] = values or options["bases"][kind]
        selected["date_start"] = request.form.get("date_start", "").strip()
        selected["date_end"] = request.form.get("date_end", "").strip()
        for key in ("date_start", "date_end"):
            if selected[key]:
                pd.Timestamp(selected[key])
        if selected["date_start"] and selected["date_end"] and selected["date_start"] > selected["date_end"]:
            raise ValueError("A data inicial não pode ser depois da final.")
        meta["filters"] = selected
        (job_dir / "meta.json").write_text(json.dumps(meta, ensure_ascii=False), encoding="utf-8")
    except Exception as exc:
        flash(f"Erro nos filtros GOL: {exc}")
    return gol_results_page(job_id, job_dir, meta)


@app.post("/gol/exportar/<job_id>")
def gol_exportar(job_id):
    job_dir, meta = load_job(job_id)
    if not job_dir or not (job_dir / "full.pkl").exists():
        flash("Prepare os relatórios GOL primeiro.")
        return redirect(url_for("gol"))
    try:
        groups = {k: tuple(v) for k, v in meta.get("groups", {}).items()}
        advp_kinds = set(meta.get("advp_kinds", []))
        has_complement = bool(meta.get("complementar_name"))
        filters = meta["filters"]
        week = set(meta["week"])
        skip = set(request.form.getlist("skip_days"))
        if not skip <= week:
            raise ValueError("Dia ignorado fora da semana de ajuste.")
        def percent(field):
            raw = request.form.get(field, "").strip()
            if not raw:
                return None
            value = float(raw.replace(",", ".")) / 100
            if not 0 <= value <= 1:
                raise ValueError(f"Share inválido: {field}.")
            return value
        targets = {}
        for kind in ("CONS", "CORP", "GOV"):
            targets[kind] = {"general": percent(f"{kind}_general") if has_complement else None,
                             "week": percent(f"{kind}_week"), "advp": {}}
            for name in groups if kind in advp_kinds else []:
                value = percent(f"{kind}_{name}")
                if value is not None:
                    targets[kind]["advp"][name] = value
        full = pd.read_pickle(job_dir / "full.pkl")
        final, removed, before, after = adjust_shares(
            full, week, targets, skip, groups, advp_kinds=advp_kinds,
            filters=filters, include_complete=has_complement)
        warnings = []
        for row in after:
            if row["total"] and row["gap_pp"] < 3 - 1e-7:
                warnings.append(f"{row['kind']} {row['period']} {row['advp']}: LATAM está menos de 3 pp à frente da GOL.")
            spec = targets[row["kind"]]
            target = (spec["advp"].get(row["advp"]) if row["advp"] != "GERAL"
                      else spec["general"] if row["period"] == "COMPLETO"
                      else spec["week"])
            if target is not None and row["total"] and abs(row["shares"]["G3"] - target) > 0.0005:
                warnings.append(f"{row['kind']} {row['period']} {row['advp']}: meta GOL {target:.2%} não atingida; resultado {row['shares']['G3']:.2%}.")
        zip_path = job_dir / "GOL_RESULTADO.zip"
        write_zip(zip_path, job_dir / meta["original_name"],
            pd.read_pickle(job_dir / "formatted.pkl"),
            full,
            pd.read_pickle(job_dir / "excluded.pkl"),
            pd.read_pickle(job_dir / "coupons.pkl"), final, removed,
            before, after, week, {**meta["stats"], "complementar": meta.get("complement_stats")},
            groups, advp_kinds, filters, has_complement,
            job_dir / meta["complementar_name"] if has_complement else None,
            job_dir / meta["origem_datas_name"] if meta.get("origem_datas_name") else None)
        daily_tables = daily_comparison_tables(full, final, week, groups, advp_kinds, filters)
        return gol_results_page(job_id, job_dir, meta, step=4, matrix=after,
                                result={"removed": len(removed), "final": len(final),
                                        "zip_name": zip_path.name, "warnings": warnings},
                                daily_tables=daily_tables)
    except Exception as exc:
        flash(f"Erro no ajuste GOL: {exc}")
        return gol_results_page(job_id, job_dir, meta)


@app.get("/gol/download/<job_id>")
def gol_download(job_id):
    job_dir, _ = load_job(job_id)
    if not job_dir or not (job_dir / "GOL_RESULTADO.zip").exists():
        flash("ZIP GOL não encontrado.")
        return redirect(url_for("gol"))
    return send_file(job_dir / "GOL_RESULTADO.zip", as_attachment=True)


@app.get("/latam")
def latam():
    return render_template(
        "index.html",
        preview=None,
        result=None,
        table_rows=[],
        filter_options=None,
        current_filters=None,
        job_id=None,
        stage1_stats=None,
    )


@app.post("/analisar")
def analisar():
    f = request.files.get("arquivo")
    if not f or not f.filename:
        flash("Selecione um arquivo CSV.")
        return redirect(url_for("latam"))

    if not f.filename.lower().endswith(".csv"):
        flash("Envie um arquivo CSV.")
        return redirect(url_for("latam"))

    job_id = uuid.uuid4().hex
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    original_name = secure_filename(f.filename) or "arquivo_original.csv"
    original_path = job_dir / original_name
    f.save(original_path)

    try:
        df, csv_meta = read_csv_robust(original_path)
        s1 = stage1_format(df)

        s1["formatted_full"].to_pickle(job_dir / "formatted_full.pkl")
        s1["main_after_west"].to_pickle(job_dir / "main_after_west.pkl")
        s1["excluded_format"].to_pickle(job_dir / "excluded_format.pkl")
        s1["west"].to_pickle(job_dir / "west.pkl")

        meta = {
            "job_id": job_id,
            "original_name": original_name,
            "csv_meta": csv_meta,
            "columns": s1["columns"],
            "stage1_stats": s1["stats"],
        }
        (job_dir / "meta.json").write_text(
            json.dumps(meta, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        options = get_filter_options(s1["main_after_west"], s1["columns"])

        default_filters = {
            "cupons": ["1"] if "1" in options["cupons"] else options["cupons"],
            "tipos": ["N"] if "N" in options["tipos"] else options["tipos"],
            "consolidadas": [
                x for x in options["consolidadas"]
                if x.strip().upper() != "WHITE LABEL"
            ],
            "date_start": options["date_start"],
            "date_end": options["date_end"],
        }

        (job_dir / "filters.json").write_text(
            json.dumps(default_filters, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        filtered = apply_user_filters(
            s1["main_after_west"],
            s1["columns"],
            default_filters,
        )
        preview = build_share_preview(filtered, s1["columns"])

        table_rows = filtered.head(300).fillna("").to_dict("records")

        return render_template(
            "index.html",
            preview=preview,
            result=None,
            table_rows=table_rows,
            table_columns=preview["display_columns"],
            filter_options=options,
            current_filters=default_filters,
            job_id=job_id,
            stage1_stats=s1["stats"],
        )

    except Exception as exc:
        shutil.rmtree(job_dir, ignore_errors=True)
        flash(f"Erro ao processar o arquivo: {exc}")
        return redirect(url_for("latam"))


@app.post("/filtros/<job_id>")
def filtros(job_id):
    job_dir, meta = load_job(job_id)
    if not job_dir:
        flash("Sessão não encontrada. Envie o arquivo novamente.")
        return redirect(url_for("latam"))

    main_df = pd.read_pickle(job_dir / "main_after_west.pkl")
    columns = meta["columns"]
    options = get_filter_options(main_df, columns)

    selected = {
        "cupons": request.form.getlist("cupons"),
        "tipos": request.form.getlist("tipos"),
        "consolidadas": request.form.getlist("consolidadas"),
        "date_start": request.form.get("date_start", "").strip(),
        "date_end": request.form.get("date_end", "").strip(),
    }

    # Se o usuário não selecionou nada em um multi-select,
    # interpreta como "todos".
    if not selected["cupons"]:
        selected["cupons"] = options["cupons"]
    if not selected["tipos"]:
        selected["tipos"] = options["tipos"]
    if not selected["consolidadas"]:
        selected["consolidadas"] = options["consolidadas"]

    (job_dir / "filters.json").write_text(
        json.dumps(selected, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    filtered = apply_user_filters(main_df, columns, selected)
    preview = build_share_preview(filtered, columns)
    table_rows = filtered.head(300).fillna("").to_dict("records")

    return render_template(
        "index.html",
        preview=preview,
        result=None,
        table_rows=table_rows,
        table_columns=preview["display_columns"],
        filter_options=options,
        current_filters=selected,
        job_id=job_id,
        stage1_stats=meta.get("stage1_stats", {}),
    )


@app.post("/processar/<job_id>")
def processar(job_id):
    job_dir, meta = load_job(job_id)
    if not job_dir:
        flash("Sessão de processamento não encontrada.")
        return redirect(url_for("latam"))

    try:
        main_df = pd.read_pickle(job_dir / "main_after_west.pkl")
        formatted_full = pd.read_pickle(job_dir / "formatted_full.pkl")
        excluded_format = pd.read_pickle(job_dir / "excluded_format.pkl")
        west = pd.read_pickle(job_dir / "west.pkl")
        columns = meta["columns"]

        filters = json.loads((job_dir / "filters.json").read_text(encoding="utf-8"))

        target_share = float(request.form["target_share"].replace(",", ".")) / 100

        use_company_bounds = request.form.get("use_company_bounds") == "S"

        def pct_optional(name):
            v = request.form.get(name, "").strip()
            return None if not v else float(v.replace(",", ".")) / 100

        g3_min = pct_optional("g3_min") if use_company_bounds else None
        g3_max = pct_optional("g3_max") if use_company_bounds else None
        ad_min = pct_optional("ad_min") if use_company_bounds else None
        ad_max = pct_optional("ad_max") if use_company_bounds else None

        daily_control = request.form.get("daily_control") == "S"
        daily_preserve_from = pct_optional("daily_preserve_from") if daily_control else None
        avoid_high_days = request.form.get("avoid_high_days") == "S"

        use_route_gap = request.form.get("use_route_gap") == "S"
        route_gap_pp = (
            float(request.form.get("route_gap_pp", "5").replace(",", ".")) / 100
            if use_route_gap else None
        )

        default_iatas = ["GRU", "POA", "CGB", "FOR", "GIG"]
        extra_iatas = [
            x.strip().upper()
            for x in request.form.get("extra_iatas", "").split(",")
            if x.strip()
        ]
        priority_iatas = list(dict.fromkeys(default_iatas + extra_iatas))

        priority_order = [
            int(x.strip())
            for x in request.form.get("priority_order", "1,2,3,4").split(",")
            if x.strip().isdigit()
        ]
        if sorted(set(priority_order)) != [1, 2, 3, 4]:
            priority_order = [1, 2, 3, 4]

        params = {
            "target_share": target_share,
            "filters": filters,
            "g3_min": g3_min,
            "g3_max": g3_max,
            "ad_min": ad_min,
            "ad_max": ad_max,
            "daily_control": daily_control,
            "daily_preserve_from": daily_preserve_from,
            "avoid_high_days": avoid_high_days,
            "route_gap": route_gap_pp,
            "priority_iatas": priority_iatas,
            "priority_order": priority_order,
        }

        result = run_share_adjustment(main_df, columns, params)

        zip_path = build_zip_outputs(
            job_dir=job_dir,
            original_path=job_dir / meta["original_name"],
            formatted_full=formatted_full,
            excluded_format=excluded_format,
            west=west,
            main_before_share=main_df,
            adjustment=result,
            columns=columns,
            params=params,
            stage1_stats=meta.get("stage1_stats", {}),
        )

        options = get_filter_options(main_df, columns)
        filtered = apply_user_filters(main_df, columns, filters)
        preview = build_share_preview(filtered, columns)
        table_rows = filtered.head(300).fillna("").to_dict("records")

        result_summary = result["summary"].copy()
        result_summary["zip_name"] = zip_path.name

        return render_template(
            "index.html",
            preview=preview,
            result=result_summary,
            table_rows=table_rows,
            table_columns=preview["display_columns"],
            filter_options=options,
            current_filters=filters,
            job_id=job_id,
            stage1_stats=meta.get("stage1_stats", {}),
        )

    except Exception as exc:
        flash(f"Erro no ajuste de share: {exc}")
        return redirect(url_for("latam"))


@app.get("/download/<job_id>")
def download(job_id):
    job_dir, _ = load_job(job_id)
    if not job_dir:
        flash("Sessão não encontrada.")
        return redirect(url_for("latam"))

    candidates = sorted(job_dir.glob("*.zip"))
    if not candidates:
        flash("Arquivo ZIP não encontrado.")
        return redirect(url_for("latam"))

    return send_file(candidates[-1], as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
