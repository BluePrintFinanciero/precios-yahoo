import io
from datetime import date, timedelta

import pandas as pd
import streamlit as st
import yfinance as yf

# ---------------------------------------------------------------------------
# Configuración de la página (cambiá nombre/emoji a gusto)
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Blueprint Financiero | Descarga de precios",
    page_icon="📈",
    layout="centered",
)

st.image("logo.jpg", width=260)
st.title("Descarga de precios históricos")
st.caption(
    "Bajá precios de cierre ajustados desde Yahoo Finance, con la variación "
    "diaria en %, listos en un Excel."
)

# ---------------------------------------------------------------------------
# Formulario
# ---------------------------------------------------------------------------
with st.form("params"):
    tickers_raw = st.text_input(
        "Tickers (separados por coma)",
        value="AMZN, GOOGL, NVDA, ^NDX",
        help="Cualquier símbolo de Yahoo Finance. Ej: AAPL, YPF.BA (BYMA), ^GSPC (S&P 500).",
    )

    col1, col2 = st.columns(2)
    with col1:
        start = st.date_input("Desde", value=date(2017, 1, 2), min_value=date(1970, 1, 1))
    with col2:
        end = st.date_input("Hasta", value=date.today())

    col3, col4 = st.columns(2)
    with col3:
        interval_label = st.selectbox(
            "Frecuencia",
            options=["Diaria", "Semanal", "Mensual"],
        )
    with col4:
        auto_adjust = st.checkbox(
            "Ajustar por splits y dividendos",
            value=True,
            help="Desmarcá para obtener el precio de mercado sin ajustar.",
        )

    submitted = st.form_submit_button("Descargar precios", type="primary", use_container_width=True)

INTERVALS = {"Diaria": "1d", "Semanal": "1wk", "Mensual": "1mo"}


# ---------------------------------------------------------------------------
# Lógica de descarga y armado del Excel
# ---------------------------------------------------------------------------
def parse_tickers(raw: str) -> list[str]:
    seen = []
    for t in raw.replace(";", ",").split(","):
        t = t.strip().upper()
        if t and t not in seen:
            seen.append(t)
    return seen


@st.cache_data(show_spinner=False, ttl=60 * 60)
def download_prices(tickers: tuple, start: date, end: date, interval: str, auto_adjust: bool) -> pd.DataFrame:
    df = yf.download(
        list(tickers),
        start=start.isoformat(),
        # yfinance excluye el día "end", sumamos 1 para incluirlo
        end=(end + timedelta(days=1)).isoformat(),
        interval=interval,
        auto_adjust=auto_adjust,
        progress=False,
        group_by="column",
    )
    if df.empty:
        return pd.DataFrame()

    close = df["Close"]
    if isinstance(close, pd.Series):  # un solo ticker en versiones viejas
        close = close.to_frame(name=tickers[0])
    return close


def build_output(close: pd.DataFrame, tickers: list[str]) -> tuple[pd.DataFrame, list[str]]:
    """Devuelve (tabla final, tickers sin datos)."""
    close = close.copy()
    missing = [t for t in tickers if t not in close.columns or close[t].isna().all()]
    valid = [t for t in tickers if t not in missing]
    close = close[valid]

    close = close.reset_index()
    date_col = close.columns[0]
    close[date_col] = pd.to_datetime(close[date_col]).dt.date
    close = close.rename(columns={date_col: "Fecha"})

    for t in valid:
        close[f"{t} Var. %"] = close[t].pct_change() * 100

    # Orden: Fecha, precios, variaciones
    ordered = ["Fecha"] + valid + [f"{t} Var. %" for t in valid]
    return close[ordered], missing


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Precios")
        ws = writer.sheets["Precios"]
        ws.freeze_panes = "B2"
        for col in ws.columns:
            width = max(len(str(c.value)) if c.value is not None else 0 for c in col[:200])
            ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 12), 40)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# Ejecución
# ---------------------------------------------------------------------------
if submitted:
    tickers = parse_tickers(tickers_raw)

    if not tickers:
        st.error("Ingresá al menos un ticker.")
        st.stop()
    if start >= end:
        st.error("La fecha 'Desde' tiene que ser anterior a 'Hasta'.")
        st.stop()

    with st.spinner("Descargando desde Yahoo Finance..."):
        try:
            close = download_prices(tuple(tickers), start, end, INTERVALS[interval_label], auto_adjust)
        except Exception as e:  # noqa: BLE001
            st.error(f"No se pudo descargar la información. Detalle: {e}")
            st.stop()

    if close.empty:
        st.error("Yahoo Finance no devolvió datos. Revisá los tickers y el rango de fechas.")
        st.stop()

    result, missing = build_output(close, tickers)

    if missing:
        st.warning(f"Sin datos para: {', '.join(missing)}. Revisá que el símbolo exista en Yahoo Finance.")
    if result.shape[1] <= 1:
        st.stop()

    st.success(f"{len(result)} filas · {len(tickers) - len(missing)} tickers · {result['Fecha'].iloc[0]} → {result['Fecha'].iloc[-1]}")

    st.download_button(
        label="⬇️ Descargar Excel",
        data=to_excel_bytes(result),
        file_name="Prices_Yahoo_Finance.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    st.subheader("Vista previa")
    st.dataframe(result.tail(15), use_container_width=True, hide_index=True)

st.divider()
st.caption(
    "Blueprint Financiero · Datos provistos por Yahoo Finance a través de la librería yfinance. "
    "Uso informativo y educativo; no constituye recomendación de inversión."
)
