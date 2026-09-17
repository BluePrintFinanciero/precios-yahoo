import io
from datetime import date, timedelta

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st
import yfinance as yf

# ---------------------------------------------------------------------------
# Configuración de la página
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Blueprint Financiero | Herramientas",
    page_icon="📈",
    layout="centered",
)

st.image("logo.jpg", width=260)

TRADING_DAYS = 252

BENCHMARKS = {
    "S&P 500 (^GSPC)": "^GSPC",
    "Nasdaq 100 (^NDX)": "^NDX",
    "Merval (^MERV)": "^MERV",
    "Bonos USD corto plazo (BIL)": "BIL",
    "Otro (escribir ticker)": None,
}

INTERVALS = {"Diaria": "1d", "Semanal": "1wk", "Mensual": "1mo"}


# ---------------------------------------------------------------------------
# Utilidades compartidas
# ---------------------------------------------------------------------------
def parse_tickers(raw: str) -> list[str]:
    seen = []
    for t in raw.replace(";", ",").replace("\n", ",").split(","):
        t = t.strip().upper()
        if t and t not in seen:
            seen.append(t)
    return seen


@st.cache_data(show_spinner=False, ttl=60 * 60)
def download_prices(tickers: tuple, start: date, end: date, interval: str, auto_adjust: bool) -> pd.DataFrame:
    df = yf.download(
        list(tickers),
        start=start.isoformat(),
        end=(end + timedelta(days=1)).isoformat(),  # yfinance excluye "end"
        interval=interval,
        auto_adjust=auto_adjust,
        progress=False,
        group_by="column",
    )
    if df.empty:
        return pd.DataFrame()
    close = df["Close"]
    if isinstance(close, pd.Series):
        close = close.to_frame(name=tickers[0])
    close.index = pd.to_datetime(close.index).tz_localize(None)
    return close


def autosize(ws, max_rows=200):
    for col in ws.columns:
        width = max(len(str(c.value)) if c.value is not None else 0 for c in col[:max_rows])
        ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 12), 40)


# ---------------------------------------------------------------------------
# PESTAÑA 1 · Descarga de precios
# ---------------------------------------------------------------------------
def build_output(close: pd.DataFrame, tickers: list[str]) -> tuple[pd.DataFrame, list[str]]:
    close = close.copy()
    missing = [t for t in tickers if t not in close.columns or close[t].isna().all()]
    valid = [t for t in tickers if t not in missing]
    close = close[valid].reset_index()
    date_col = close.columns[0]
    close[date_col] = pd.to_datetime(close[date_col]).dt.date
    close = close.rename(columns={date_col: "Fecha"})
    for t in valid:
        close[f"{t} Var. %"] = close[t].pct_change() * 100
    ordered = ["Fecha"] + valid + [f"{t} Var. %" for t in valid]
    return close[ordered], missing


def prices_to_excel(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Precios")
        ws = writer.sheets["Precios"]
        ws.freeze_panes = "B2"
        autosize(ws)
    return buffer.getvalue()


def tab_descarga():
    st.header("Descarga de precios históricos")
    st.caption("Bajá precios de cierre ajustados desde Yahoo Finance, con la variación en %, listos en un Excel.")

    with st.form("params_descarga"):
        tickers_raw = st.text_input(
            "Tickers (separados por coma)",
            value="AMZN, GOOGL, NVDA, ^NDX",
            help="Cualquier símbolo de Yahoo Finance. Ej: AAPL, YPF.BA (BYMA), ^GSPC (S&P 500).",
        )
        c1, c2 = st.columns(2)
        start = c1.date_input("Desde", value=date(2017, 1, 2), min_value=date(1970, 1, 1), key="d_start")
        end = c2.date_input("Hasta", value=date.today(), key="d_end")
        c3, c4 = st.columns(2)
        interval_label = c3.selectbox("Frecuencia", options=list(INTERVALS))
        auto_adjust = c4.checkbox("Ajustar por splits y dividendos", value=True)
        submitted = st.form_submit_button("Descargar precios", type="primary", use_container_width=True)

    if not submitted:
        return

    tickers = parse_tickers(tickers_raw)
    if not tickers:
        st.error("Ingresá al menos un ticker.")
        return
    if start >= end:
        st.error("La fecha 'Desde' tiene que ser anterior a 'Hasta'.")
        return

    with st.spinner("Descargando desde Yahoo Finance..."):
        try:
            close = download_prices(tuple(tickers), start, end, INTERVALS[interval_label], auto_adjust)
        except Exception as e:  # noqa: BLE001
            st.error(f"No se pudo descargar la información. Detalle: {e}")
            return

    if close.empty:
        st.error("Yahoo Finance no devolvió datos. Revisá los tickers y el rango de fechas.")
        return

    result, missing = build_output(close, tickers)
    if missing:
        st.warning(f"Sin datos para: {', '.join(missing)}. Revisá que el símbolo exista en Yahoo Finance.")
    if result.shape[1] <= 1:
        return

    st.success(
        f"{len(result)} filas · {len(tickers) - len(missing)} tickers · "
        f"{result['Fecha'].iloc[0]} → {result['Fecha'].iloc[-1]}"
    )
    st.download_button(
        "⬇️ Descargar Excel",
        data=prices_to_excel(result),
        file_name="Prices_Yahoo_Finance.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
    st.subheader("Vista previa")
    st.dataframe(result.tail(15), use_container_width=True, hide_index=True)


# ---------------------------------------------------------------------------
# PESTAÑA 2 · Radiografía de la cartera
# ---------------------------------------------------------------------------
def max_drawdown(series: pd.Series) -> tuple[float, int, str]:
    """Devuelve (caída máxima en %, meses hasta recuperar o -1 si no recuperó, fecha del piso)."""
    running_max = series.cummax()
    dd = series / running_max - 1
    trough = dd.idxmin()
    mdd = dd.min() * 100
    peak = series.loc[:trough].idxmax()
    after = series.loc[trough:]
    recovered = after[after >= series.loc[peak]]
    if recovered.empty:
        months = -1
    else:
        months = max(1, round((recovered.index[0] - peak).days / 30.4))
    return mdd, months, trough.strftime("%b %Y")


def corr_phrase(c: float) -> str:
    if c >= 0.9:
        return "se mueve prácticamente igual que el índice: en la práctica, casi la misma exposición que comprar el índice."
    if c >= 0.75:
        return "sigue de cerca al índice. La diversificación respecto del mercado es limitada."
    if c >= 0.5:
        return "está bastante ligada al índice, aunque tiene comportamiento propio."
    if c >= 0.2:
        return "tiene una relación moderada con el índice."
    return "se mueve de forma bastante independiente del índice."


def vol_phrase(vol_port: float, vol_bench: float) -> str:
    ratio = vol_port / vol_bench if vol_bench else np.nan
    if np.isnan(ratio):
        return ""
    if ratio > 1.3:
        return f"Tu cartera es {ratio:.1f} veces más volátil que el índice: asumís bastante más riesgo."
    if ratio > 1.05:
        return "Tu cartera es algo más volátil que el índice."
    if ratio >= 0.95:
        return "Tu cartera tiene un riesgo similar al del índice."
    return f"Tu cartera es menos volátil que el índice ({ratio:.2f} veces)."


def analyze(close: pd.DataFrame, weights: dict[str, float], bench: str) -> dict:
    assets = list(weights)
    data = close[assets + [bench]].dropna()
    rets = data.pct_change().dropna()
    w = pd.Series(weights, dtype=float)
    w = w / w.sum()

    port_ret = (rets[assets] * w).sum(axis=1)
    bench_ret = rets[bench]

    growth = pd.DataFrame(
        {"Tu cartera": (1 + port_ret).cumprod() * 100, "Índice": (1 + bench_ret).cumprod() * 100}
    )
    growth.loc[rets.index[0] - pd.Timedelta(days=1)] = [100.0, 100.0]
    growth = growth.sort_index()

    years = (rets.index[-1] - rets.index[0]).days / 365.25
    ann = lambda s: ((1 + s).prod() ** (1 / years) - 1) * 100 if years > 0 else np.nan  # noqa: E731

    vol = rets.std() * np.sqrt(TRADING_DAYS) * 100
    vol_port = port_ret.std() * np.sqrt(TRADING_DAYS) * 100

    corr_matrix = rets[assets].corr()
    corr_pb = port_ret.corr(bench_ret)
    beta = port_ret.cov(bench_ret) / bench_ret.var()

    dd_rows = []
    for name, s in [("Tu cartera", growth["Tu cartera"]), ("Índice", growth["Índice"])] + [
        (a, (1 + rets[a]).cumprod()) for a in assets
    ]:
        mdd, months, when = max_drawdown(s)
        dd_rows.append({"Activo": name, "Peor caída %": mdd, "Meses para recuperar": months, "Piso": when})
    dd = pd.DataFrame(dd_rows)

    pairs = []
    for i, a in enumerate(assets):
        for b in assets[i + 1 :]:
            pairs.append((a, b, corr_matrix.loc[a, b]))
    high_pairs = [p for p in pairs if p[2] >= 0.8]

    summary = pd.DataFrame(
        {
            "Peso %": (w * 100).round(1),
            "Rendimiento anual %": [ann(rets[a]) for a in assets],
            "Volatilidad anual %": [vol[a] for a in assets],
            "Correlación c/ índice": [rets[a].corr(bench_ret) for a in assets],
        },
        index=assets,
    )

    return dict(
        start=rets.index[0].date(),
        end=rets.index[-1].date(),
        years=years,
        growth=growth,
        ann_port=ann(port_ret),
        ann_bench=ann(bench_ret),
        vol=vol,
        vol_port=vol_port,
        vol_bench=vol[bench],
        corr_matrix=corr_matrix,
        corr_pb=corr_pb,
        beta=beta,
        dd=dd,
        pairs=pairs,
        high_pairs=high_pairs,
        summary=summary,
        weights=w,
    )


def analysis_to_excel(res: dict, bench: str) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        resumen = pd.DataFrame(
            {
                "Métrica": [
                    "Período analizado",
                    "Rendimiento anual cartera %",
                    f"Rendimiento anual índice ({bench}) %",
                    "Volatilidad anual cartera %",
                    f"Volatilidad anual índice ({bench}) %",
                    "Correlación cartera vs índice",
                    "Beta cartera vs índice",
                ],
                "Valor": [
                    f"{res['start']} a {res['end']}",
                    round(res["ann_port"], 2),
                    round(res["ann_bench"], 2),
                    round(res["vol_port"], 2),
                    round(res["vol_bench"], 2),
                    round(res["corr_pb"], 3),
                    round(res["beta"], 3),
                ],
            }
        )
        resumen.to_excel(writer, index=False, sheet_name="Resumen")
        res["summary"].round(3).rename_axis("Ticker").to_excel(writer, sheet_name="Por activo")
        res["corr_matrix"].round(3).to_excel(writer, sheet_name="Correlaciones")
        res["dd"].round(2).to_excel(writer, index=False, sheet_name="Caídas")
        g = res["growth"].copy()
        g.index = g.index.date
        g.round(2).rename_axis("Fecha").to_excel(writer, sheet_name="Base 100")
        for ws in writer.sheets.values():
            autosize(ws)
    return buffer.getvalue()


def tab_cartera():
    st.header("Radiografía de tu cartera")
    st.caption(
        "Cargá tus activos y descubrí qué tan diversificada está tu cartera, cuánto se mueve con el "
        "mercado y qué tan fuerte podría caer. Todo con datos históricos reales."
    )

    tickers_raw = st.text_input(
        "Tickers de tu cartera (separados por coma)",
        value="AAPL, MSFT, NVDA, KO, GLD",
        help="Ej: GGAL.BA para acciones argentinas en pesos, AAPL.BA para CEDEARs, BTC-USD para cripto.",
        key="c_tickers",
    )
    tickers = parse_tickers(tickers_raw)

    st.markdown("**Pesos (% de la cartera).** Si no los tocás, se reparte en partes iguales.")
    default_w = round(100 / len(tickers), 1) if tickers else 0
    weights_df = st.data_editor(
        pd.DataFrame({"Ticker": tickers, "Peso %": [default_w] * len(tickers)}),
        hide_index=True,
        use_container_width=True,
        disabled=["Ticker"],
        column_config={"Peso %": st.column_config.NumberColumn(min_value=0, max_value=100, step=0.5, format="%.1f")},
        key=f"weights_{','.join(tickers)}",
    )

    total_w = float(pd.to_numeric(weights_df["Peso %"], errors="coerce").fillna(0).sum())
    if abs(total_w - 100) < 0.05:
        st.success(f"Total cargado: {total_w:.1f}% ✓")
    elif total_w < 100:
        st.warning(f"Total cargado: {total_w:.1f}% · te faltan {100 - total_w:.1f}% para llegar al 100%")
    else:
        st.warning(f"Total cargado: {total_w:.1f}% · te pasaste {total_w - 100:.1f}% del 100%")

    c1, c2 = st.columns(2)
    bench_label = c1.selectbox("Comparar contra", options=list(BENCHMARKS))
    bench = BENCHMARKS[bench_label]
    if bench is None:
        bench = c1.text_input("Ticker del índice / ETF", value="SPY").strip().upper()
    years_back = c2.selectbox("Período", options=[3, 5, 10, 15], index=1, format_func=lambda y: f"Últimos {y} años")

    run = st.button("Analizar mi cartera", type="primary", use_container_width=True)
    if not run:
        return

    if len(tickers) < 2:
        st.error("Cargá al menos dos activos para analizar la cartera.")
        return
    if not bench:
        st.error("Elegí un índice para comparar.")
        return

    weights = {t: float(p) for t, p in zip(weights_df["Ticker"], weights_df["Peso %"]) if float(p) > 0}
    if len(weights) < 2:
        st.error("Al menos dos activos necesitan un peso mayor a cero.")
        return
    if abs(sum(weights.values()) - 100) > 0.5:
        st.info(f"Los pesos suman {sum(weights.values()):.1f}%. Se normalizan automáticamente a 100%.")

    end = date.today()
    start = end - timedelta(days=int(years_back * 365.25))

    with st.spinner("Descargando datos y calculando..."):
        try:
            close = download_prices(tuple(sorted(set(list(weights) + [bench]))), start, end, "1d", True)
        except Exception as e:  # noqa: BLE001
            st.error(f"No se pudo descargar la información. Detalle: {e}")
            return

    if close.empty:
        st.error("Yahoo Finance no devolvió datos. Revisá los tickers.")
        return

    missing = [t for t in list(weights) + [bench] if t not in close.columns or close[t].isna().all()]
    if bench in missing:
        st.error(f"No hay datos para el índice {bench}. Probá con otro.")
        return
    if missing:
        st.warning(f"Sin datos para: {', '.join(missing)}. Se analizan los demás.")
        weights = {t: w for t, w in weights.items() if t not in missing}
        if len(weights) < 2:
            st.error("Quedaron menos de dos activos con datos.")
            return

    res = analyze(close, weights, bench)
    assets = list(res["weights"].index)

    if res["years"] < years_back * 0.8:
        st.info(
            f"Algún activo tiene historia más corta: el análisis cubre desde {res['start']} "
            f"({res['years']:.1f} años), el tramo en que todos tienen datos."
        )

    st.caption(f"Período analizado: {res['start']} → {res['end']} · Datos diarios ajustados por dividendos y splits.")

    # 1 · Correlación con el índice -------------------------------------------------
    st.subheader("1 · Cuánto se mueve tu cartera con el mercado")
    m1, m2, m3 = st.columns(3)
    m1.metric("Correlación con el índice", f"{res['corr_pb']:.0%}")
    m2.metric("Beta", f"{res['beta']:.2f}", help="Si el índice sube 1%, tu cartera tiende a moverse este número en %.")
    m3.metric("Activos", f"{len(assets)}")
    st.write(f"Tu cartera **{corr_phrase(res['corr_pb'])}**")

    # 2 · Mapa de correlaciones -----------------------------------------------------
    st.subheader("2 · ¿Tus activos se mueven distinto entre sí?")
    cm = res["corr_matrix"].rename_axis(index="a", columns=None).reset_index().melt(id_vars="a", var_name="b", value_name="corr")
    heat = (
        alt.Chart(cm)
        .mark_rect()
        .encode(
            x=alt.X("a:N", title=None, sort=assets),
            y=alt.Y("b:N", title=None, sort=assets),
            color=alt.Color(
                "corr:Q",
                scale=alt.Scale(domain=[-1, 0, 1], range=["#3b8f5e", "#16203a", "#c8503f"]),
                legend=alt.Legend(title="Correlación"),
            ),
            tooltip=[alt.Tooltip("a:N", title="Activo"), alt.Tooltip("b:N", title="Activo"), alt.Tooltip("corr:Q", format=".2f")],
        )
    )
    text = heat.mark_text(fontSize=12).encode(
        text=alt.Text("corr:Q", format=".2f"),
        color=alt.condition("abs(datum.corr) > 0.6", alt.value("white"), alt.value("#c1c2c4")),
    )
    st.altair_chart((heat + text).properties(height=60 + 42 * len(assets)), use_container_width=True)

    n_pairs = len(res["pairs"])
    n_high = len(res["high_pairs"])
    if n_high == 0:
        st.write("Ningún par de activos tiene correlación mayor a 0,80. **Tus activos se mueven de forma distinta entre sí: buena diversificación interna.**")
    else:
        names = ", ".join(f"{a}–{b}" for a, b, _ in res["high_pairs"][:4])
        extra = f" y {n_high - 4} más" if n_high > 4 else ""
        st.write(
            f"**{n_high} de {n_pairs} pares se mueven casi igual** (correlación ≥ 0,80): {names}{extra}. "
            "En la práctica, esos activos funcionan como una sola apuesta."
        )
    st.caption("Verde: se mueven distinto (diversifican). Rojo: se mueven igual (no diversifican).")

    # 3 · Volatilidad ---------------------------------------------------------------
    st.subheader("3 · Cuánto riesgo tiene cada activo")
    vol_df = pd.concat(
        [res["vol"][assets], pd.Series({"Tu cartera": res["vol_port"], bench_label.split(" (")[0]: res["vol_bench"]})]
    ).rename("Volatilidad anual %").rename_axis("Activo").reset_index()
    vol_df["Tipo"] = np.where(vol_df["Activo"].isin(assets), "Activo", "Referencia")
    bars = (
        alt.Chart(vol_df)
        .mark_bar()
        .encode(
            x=alt.X("Volatilidad anual %:Q", title="Volatilidad anual (%)"),
            y=alt.Y("Activo:N", sort="-x", title=None),
            color=alt.Color("Tipo:N", scale=alt.Scale(domain=["Activo", "Referencia"], range=["#6b7fa8", "#c1c2c4"]), legend=None),
            tooltip=["Activo", alt.Tooltip("Volatilidad anual %:Q", format=".1f")],
        )
        .properties(height=40 + 30 * len(vol_df))
    )
    st.altair_chart(bars, use_container_width=True)
    most = res["vol"][assets].idxmax()
    least = res["vol"][assets].idxmin()
    ratio_ml = res["vol"][most] / res["vol"][least]
    st.write(
        f"{vol_phrase(res['vol_port'], res['vol_bench'])} "
        f"**{most}** es tu activo más volátil ({res['vol'][most]:.0f}% anual) y **{least}** el más estable "
        f"({res['vol'][least]:.0f}%): darles el mismo peso no significa asumir el mismo riesgo en cada uno "
        f"(uno se mueve {ratio_ml:.1f} veces más que el otro)."
    )

    # 4 · Drawdowns -----------------------------------------------------------------
    st.subheader("4 · Cuánto podría caer")
    dd = res["dd"].copy()
    dd_port = dd.iloc[0]
    rec = "todavía no recuperó ese nivel" if dd_port["Meses para recuperar"] == -1 else f"tardó {int(dd_port['Meses para recuperar'])} meses en recuperar"
    st.write(
        f"En el período analizado, tu cartera **hubiese caído hasta un {abs(dd_port['Peor caída %']):.0f}%** "
        f"desde su máximo (piso en {dd_port['Piso']}) y {rec}. ¿Lo aguantarías sin vender?"
    )
    dd_show = dd.copy()
    dd_show["Peor caída %"] = dd_show["Peor caída %"].map(lambda v: f"{v:.1f}%")
    dd_show["Meses para recuperar"] = dd_show["Meses para recuperar"].map(lambda m: "Sin recuperar" if m == -1 else f"{int(m)}")
    st.dataframe(dd_show, hide_index=True, use_container_width=True)

    # 5 · Base 100 ------------------------------------------------------------------
    st.subheader("5 · Tu cartera contra el índice")
    g = res["growth"].rename_axis("Fecha").reset_index().melt(id_vars="Fecha", var_name="Serie", value_name="Valor")
    line = (
        alt.Chart(g)
        .mark_line(strokeWidth=2)
        .encode(
            x=alt.X("Fecha:T", title=None, axis=alt.Axis(format="%m/%Y", labelAngle=0)),
            y=alt.Y("Valor:Q", title="Base 100", scale=alt.Scale(zero=False)),
            color=alt.Color("Serie:N", scale=alt.Scale(domain=["Tu cartera", "Índice"], range=["#6b7fa8", "#c1c2c4"]), legend=alt.Legend(title=None, orient="top")),
            tooltip=[alt.Tooltip("Fecha:T"), "Serie", alt.Tooltip("Valor:Q", format=".0f")],
        )
        .properties(height=320)
    )
    st.altair_chart(line, use_container_width=True)
    final_p = res["growth"]["Tu cartera"].iloc[-1]
    final_b = res["growth"]["Índice"].iloc[-1]
    won = final_p >= final_b
    st.write(
        f"\\$100 invertidos en tu cartera al inicio hoy serían **\\${final_p:,.0f}**; en el índice, **\\${final_b:,.0f}**. "
        f"Rendimiento anual: {res['ann_port']:.1f}% vs. {res['ann_bench']:.1f}%, con volatilidad de "
        f"{res['vol_port']:.0f}% vs. {res['vol_bench']:.0f}%. "
        + ("Le ganaste al índice, pero mirá si fue asumiendo más riesgo." if won and res["vol_port"] > res["vol_bench"] * 1.05
           else "Le ganaste al índice con un riesgo similar o menor." if won
           else "El índice rindió más: vale preguntarse qué aporta la selección de activos.")
    )

    # Detalle y descarga ------------------------------------------------------------
    st.subheader("Detalle por activo")
    st.dataframe(
        res["summary"].style.format({"Peso %": "{:.1f}", "Rendimiento anual %": "{:.1f}", "Volatilidad anual %": "{:.1f}", "Correlación c/ índice": "{:.2f}"}),
        use_container_width=True,
    )
    st.download_button(
        "⬇️ Descargar análisis en Excel",
        data=analysis_to_excel(res, bench),
        file_name="Radiografia_Cartera.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
    st.caption(
        "Los cálculos suponen pesos constantes (rebalanceo diario) y usan precios de cierre ajustados. "
        "Todo mide el comportamiento pasado: el riesgo histórico no garantiza el futuro."
    )


# ---------------------------------------------------------------------------
# Layout
# ---------------------------------------------------------------------------
tab1, tab2 = st.tabs(["🔍 Radiografía de tu cartera", "⬇️ Descarga de precios"])
with tab1:
    tab_cartera()
with tab2:
    tab_descarga()

st.divider()
st.caption(
    "Blueprint Financiero · Datos provistos por Yahoo Finance a través de la librería yfinance. "
    "Uso informativo y educativo; no constituye recomendación de inversión."
)
