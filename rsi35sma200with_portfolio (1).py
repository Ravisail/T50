#!/usr/bin/env python3
"""
NIFTY 200 Stock Scanner & Backtesting Platform v3.4.3
====================================================

Production-ready Streamlit app with realistic portfolio simulation,
state-based signal generation, batch data downloads, and comprehensive
performance metrics.

STRATEGY (v3.3+) — RSI Tracking + EMA Alignment (state-based)
------------------------------------------------------------
A two-phase, state-machine strategy. The RSI event only *arms* tracking;
the actual entry waits for an early reclaim above the 200 EMA while the
faster EMAs are still below it.

  Indicators (every candle): RSI(14), EMA(200), EMA(50), EMA(21)

  1. START TRACKING (arming, NOT a buy):
       When on the SAME candle:  RSI(14) < 35  AND  Close < EMA(200)
       -> set Tracking = TRUE, save the tracking-start date. Do NOT buy.
  2. TRACKING MODE:
       Ignore all further RSI signals. Remain tracking until a valid buy.
  3. BUY — only while Tracking == TRUE, all FIVE true on one candle:
       (1) Tracking == True
       (2) Close   > EMA(200)
       (3) EMA(50) > EMA(21)
       (4) EMA(50) < EMA(200)
       (5) EMA(21) < EMA(200)
       i.e. structure  Close > EMA200 > EMA50 > EMA21  (early reversal).
       -> generate BUY, execute per the backtest execution model
          (T+1 fill), then reset Tracking = FALSE.

  Rules enforced:
    - The RSI-trigger candle can NEVER buy (arming only).
    - Repeated RSI dips while tracking are ignored (no re-arming, no
      duplicate tracking states).
    - Exactly ONE buy per tracking cycle; a fresh RSI+below-EMA200 event
      is required to arm the next cycle.
    - No repainting: only completed candles, indicators use min_periods.
    - Scanner, chart signals, and backtest share this one code path.

  Portfolio layer (unchanged from prior versions):
    - SELL: price-based profit target (default 3.14%).
    - AVERAGE: a genuinely new buy signal whose T+1 fill is >=trigger%
      below the previous execution price, up to N entries (Averaging mode).
    - Optional fixed-% Stop Loss mode (mutually exclusive with Averaging).

Author: AI Assistant
Version: 3.4.3
Python: 3.12+
"""

# =============================================================================
# IMPORTS
# =============================================================================

import os
import sys
import json
import time
import random
import hashlib
import logging
import warnings
import sqlite3
from pathlib import Path
from datetime import datetime, timedelta, date as dt_date
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple, Union, Any, Callable, Set
from concurrent.futures import ThreadPoolExecutor, as_completed
from functools import wraps
import io
import html as _html

import numpy as np
import pandas as pd
import yfinance as yf
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

warnings.filterwarnings('ignore')
logging.getLogger('yfinance').setLevel(logging.ERROR)

# =============================================================================
# CONSTANTS
# =============================================================================

APP_NAME = "NIFTY 200 Strategy Scanner"
APP_VERSION = "3.4.3"
APP_ICON = "📈"

CACHE_DIR = Path.home() / ".nifty200_scanner_cache"
try:
    CACHE_DIR.mkdir(exist_ok=True)
except (PermissionError, OSError):
    # Read-only or unwritable home (e.g. some container / cloud sandboxes):
    # fall back to a temp dir instead of crashing at import time.
    import tempfile
    CACHE_DIR = Path(tempfile.gettempdir()) / "nifty200_scanner_cache"
    try:
        CACHE_DIR.mkdir(exist_ok=True)
    except (PermissionError, OSError):
        pass  # DataManager still works from its in-memory cache if this fails

# NIFTY 200 Stock Symbols (deduplicated, sorted)
# NIFTY 200 constituents: NIFTY 50 + NIFTY Next 50 (well-established, stable)
# plus a curated set of additional large/mid-cap NSE names, cross-checked to be a
# subset of the verified NIFTY 500 list below. NSE rebalances these indices
# semi-annually (Jan 31 / Jul 31 cut-off), so this may drift slightly from the
# live official constituents over time. For an exact match, use the 'Upload
# Excel/CSV' custom-universe option in the sidebar instead.
NIFTY200_SYMBOLS = sorted(list(set([
    "ABB.NS", "ABBOTINDIA.NS", "ACC.NS", "ADANIENSOL.NS", "ADANIENT.NS", "ADANIGREEN.NS", "ADANIPORTS.NS", "ADANIPOWER.NS",
    "ALKEM.NS", "AMBUJACEM.NS", "ANGELONE.NS", "APLAPOLLO.NS", "APOLLOHOSP.NS", "APOLLOTYRE.NS", "ASHOKLEY.NS", "ASIANPAINT.NS",
    "ASTRAL.NS", "ATGL.NS", "AUROPHARMA.NS", "AXISBANK.NS", "BAJAJ-AUTO.NS", "BAJAJFINSV.NS", "BAJAJHLDNG.NS", "BAJFINANCE.NS",
    "BALKRISIND.NS", "BANKBARODA.NS", "BATAINDIA.NS", "BEL.NS", "BHARATFORG.NS", "BHARTIARTL.NS", "BHEL.NS", "BIOCON.NS",
    "BOSCHLTD.NS", "BRITANNIA.NS", "BSE.NS", "CAMPUS.NS", "CAMS.NS", "CANBK.NS", "CANFINHOME.NS", "CDSL.NS",
    "CENTURYPLY.NS", "CGPOWER.NS", "CHENNPETRO.NS", "CHOLAFIN.NS", "CIPLA.NS", "COALINDIA.NS", "COFORGE.NS", "COLPAL.NS",
    "CONCOR.NS", "CROMPTON.NS", "CUMMINSIND.NS", "DABUR.NS", "DELHIVERY.NS", "DIVISLAB.NS", "DIXON.NS", "DLF.NS",
    "DRREDDY.NS", "EICHERMOT.NS", "ESCORTS.NS", "ETERNAL.NS", "EXIDEIND.NS", "FINCABLES.NS", "FINPIPE.NS", "FORTIS.NS",
    "GAIL.NS", "GESHIP.NS", "GLAND.NS", "GLENMARK.NS", "GODREJCP.NS", "GODREJPROP.NS", "GRANULES.NS", "GRASIM.NS",
    "GRINDWELL.NS", "GSPL.NS", "GUJGASLTD.NS", "HAL.NS", "HAVELLS.NS", "HCLTECH.NS", "HDFCBANK.NS", "HDFCLIFE.NS",
    "HEROMOTOCO.NS", "HINDALCO.NS", "HINDPETRO.NS", "HINDUNILVR.NS", "HINDZINC.NS", "HONASA.NS", "HUDCO.NS", "ICICIBANK.NS",
    "ICICIGI.NS", "ICICIPRULI.NS", "IEX.NS", "IGL.NS", "IIFL.NS", "INDIGO.NS", "INDUSINDBK.NS", "INDUSTOWER.NS",
    "INFY.NS", "IOC.NS", "IPCALAB.NS", "IRCTC.NS", "IRFC.NS", "ITC.NS", "JINDALSAW.NS", "JINDALSTEL.NS",
    "JIOFIN.NS", "JSWENERGY.NS", "JSWSTEEL.NS", "JUBLFOOD.NS", "KAJARIACER.NS", "KALYANKJIL.NS", "KEI.NS", "KOTAKBANK.NS",
    "KPITTECH.NS", "LALPATHLAB.NS", "LAURUSLABS.NS", "LICHSGFIN.NS", "LICI.NS", "LODHA.NS", "LT.NS", "LTIM.NS",
    "LTTS.NS", "LUPIN.NS", "M&M.NS", "MANAPPURAM.NS", "MARICO.NS", "MARUTI.NS", "MAXHEALTH.NS", "MCX.NS",
    "METROBRAND.NS", "MFSL.NS", "MGL.NS", "MOTHERSON.NS", "MOTILALOFS.NS", "MPHASIS.NS", "MRF.NS", "MRPL.NS",
    "MUTHOOTFIN.NS", "NATCOPHARM.NS", "NAUKRI.NS", "NESTLEIND.NS", "NHPC.NS", "NLCINDIA.NS", "NTPC.NS", "NYKAA.NS",
    "OBEROIRLTY.NS", "OIL.NS", "ONGC.NS", "PAGEIND.NS", "PATANJALI.NS", "PAYTM.NS", "PERSISTENT.NS", "PETRONET.NS",
    "PFC.NS", "PHOENIXLTD.NS", "PIDILITIND.NS", "PNB.NS", "PNBHOUSING.NS", "POLICYBZR.NS", "POLYCAB.NS", "POWERGRID.NS",
    "PRESTIGE.NS", "RAJESHEXPO.NS", "RATNAMANI.NS", "RECLTD.NS", "RELIANCE.NS", "RVNL.NS", "SBICARD.NS", "SBILIFE.NS",
    "SBIN.NS", "SCHAEFFLER.NS", "SHREECEM.NS", "SHRIRAMFIN.NS", "SIEMENS.NS", "SJVN.NS", "SKFINDIA.NS", "SUNPHARMA.NS",
    "SUPREMEIND.NS", "SYNGENE.NS", "TATACOMM.NS", "TATACONSUM.NS", "TATAMOTORS.NS", "TATAPOWER.NS", "TATASTEEL.NS", "TCS.NS",
    "TECHM.NS", "THERMAX.NS", "TIMKEN.NS", "TITAN.NS", "TORNTPHARM.NS", "TORNTPOWER.NS", "TRENT.NS", "TVSMOTOR.NS",
    "ULTRACEMCO.NS", "UNITDSPR.NS", "VBL.NS", "VEDL.NS", "VIPIND.NS", "VOLTAS.NS", "WELCORP.NS", "WIPRO.NS",
])))

# NIFTY 500 constituents, sourced from NSE-published data (Wikipedia mirror,
# snapshot dated 26 Nov 2024). Same semi-annual-rebalance caveat as above applies.
NIFTY500_SYMBOLS = sorted(list(set([
    "360ONE.NS", "3MINDIA.NS", "AARTIIND.NS", "AAVAS.NS", "ABB.NS", "ABBOTINDIA.NS", "ABCAPITAL.NS", "ABFRL.NS",
    "ABREL.NS", "ACC.NS", "ACE.NS", "ACI.NS", "ADANIENSOL.NS", "ADANIENT.NS", "ADANIGREEN.NS", "ADANIPORTS.NS",
    "ADANIPOWER.NS", "AEGISLOG.NS", "AETHER.NS", "AFFLE.NS", "AIAENG.NS", "AJANTPHARM.NS", "ALIVUS.NS", "ALKEM.NS",
    "ALKYLAMINE.NS", "ALLCARGO.NS", "ALOKINDS.NS", "AMBER.NS", "AMBUJACEM.NS", "ANANDRATHI.NS", "ANGELONE.NS", "ANURAS.NS",
    "APARINDS.NS", "APLAPOLLO.NS", "APLLTD.NS", "APOLLOHOSP.NS", "APOLLOTYRE.NS", "APTUS.NS", "ARE&M.NS", "ASAHIINDIA.NS",
    "ASHOKLEY.NS", "ASIANPAINT.NS", "ASTERDM.NS", "ASTRAL.NS", "ASTRAZEN.NS", "ATGL.NS", "ATUL.NS", "AUBANK.NS",
    "AUROPHARMA.NS", "AVANTIFEED.NS", "AWL.NS", "AXISBANK.NS", "BAJAJ-AUTO.NS", "BAJAJFINSV.NS", "BAJAJHLDNG.NS", "BAJFINANCE.NS",
    "BALAMINES.NS", "BALKRISIND.NS", "BALRAMCHIN.NS", "BANDHANBNK.NS", "BANKBARODA.NS", "BANKINDIA.NS", "BATAINDIA.NS", "BAYERCROP.NS",
    "BBTC.NS", "BDL.NS", "BEL.NS", "BEML.NS", "BERGEPAINT.NS", "BHARATFORG.NS", "BHARTIARTL.NS", "BHEL.NS",
    "BIKAJI.NS", "BIOCON.NS", "BIRLACORPN.NS", "BLS.NS", "BLUEDART.NS", "BLUESTARCO.NS", "BORORENEW.NS", "BOSCHLTD.NS",
    "BPCL.NS", "BRIGADE.NS", "BRITANNIA.NS", "BSE.NS", "BSOFT.NS", "CAMPUS.NS", "CAMS.NS", "CANBK.NS",
    "CANFINHOME.NS", "CAPLIPOINT.NS", "CARBORUNIV.NS", "CASTROLIND.NS", "CCL.NS", "CDSL.NS", "CEATLTD.NS", "CELLO.NS",
    "CENTRALBK.NS", "CENTURYPLY.NS", "CERA.NS", "CESC.NS", "CGCL.NS", "CGPOWER.NS", "CHALET.NS", "CHAMBLFERT.NS",
    "CHEMPLASTS.NS", "CHENNPETRO.NS", "CHOLAFIN.NS", "CHOLAHLDNG.NS", "CIEINDIA.NS", "CIPLA.NS", "CLEAN.NS", "COALINDIA.NS",
    "COCHINSHIP.NS", "COFORGE.NS", "COLPAL.NS", "CONCOR.NS", "CONCORDBIO.NS", "COROMANDEL.NS", "CRAFTSMAN.NS", "CREDITACC.NS",
    "CRISIL.NS", "CROMPTON.NS", "CSBBANK.NS", "CUB.NS", "CUMMINSIND.NS", "CYIENT.NS", "DABUR.NS", "DALBHARAT.NS",
    "DATAPATTNS.NS", "DCMSHRIRAM.NS", "DEEPAKFERT.NS", "DEEPAKNTR.NS", "DELHIVERY.NS", "DEVYANI.NS", "DIVISLAB.NS", "DIXON.NS",
    "DLF.NS", "DMART.NS", "DOMS.NS", "DRREDDY.NS", "EASEMYTRIP.NS", "ECLERX.NS", "EICHERMOT.NS", "EIDPARRY.NS",
    "EIHOTEL.NS", "ELECON.NS", "ELGIEQUIP.NS", "EMAMILTD.NS", "ENDURANCE.NS", "ENGINERSIN.NS", "EPL.NS", "EQUITASBNK.NS",
    "ERIS.NS", "ESCORTS.NS", "ETERNAL.NS", "EXIDEIND.NS", "FACT.NS", "FDC.NS", "FEDERALBNK.NS", "FINCABLES.NS",
    "FINEORG.NS", "FINPIPE.NS", "FIVESTAR.NS", "FLUOROCHEM.NS", "FORTIS.NS", "FSL.NS", "GAEL.NS", "GAIL.NS",
    "GESHIP.NS", "GICRE.NS", "GILLETTE.NS", "GLAND.NS", "GLAXO.NS", "GLENMARK.NS", "GMDCLTD.NS", "GMMPFAUDLR.NS",
    "GMRINFRASTRUCT.NS", "GNFC.NS", "GODFRYPHLP.NS", "GODREJCP.NS", "GODREJIND.NS", "GODREJPROP.NS", "GPIL.NS", "GPPL.NS",
    "GRANULES.NS", "GRAPHITE.NS", "GRASIM.NS", "GRINDWELL.NS", "GRSE.NS", "GSFC.NS", "GSPL.NS", "GUJGASLTD.NS",
    "HAL.NS", "HAPPSTMNDS.NS", "HAPPYFORGE.NS", "HAVELLS.NS", "HBLENGINE.NS", "HCLTECH.NS", "HDFCAMC.NS", "HDFCBANK.NS",
    "HDFCLIFE.NS", "HEG.NS", "HEROMOTOCO.NS", "HFCL.NS", "HINDALCO.NS", "HINDCOPPER.NS", "HINDPETRO.NS", "HINDUNILVR.NS",
    "HINDZINC.NS", "HOMEFIRST.NS", "HONASA.NS", "HONAUT.NS", "HSCL.NS", "HUDCO.NS", "ICICIBANK.NS", "ICICIGI.NS",
    "ICICIPRULI.NS", "IDBI.NS", "IDEA.NS", "IDFCFIRSTB.NS", "IEX.NS", "IFCI.NS", "IGL.NS", "IIFL.NS",
    "INDHOTEL.NS", "INDIACEM.NS", "INDIAMART.NS", "INDIANB.NS", "INDIGO.NS", "INDIGOPNTS.NS", "INDUSINDBK.NS", "INDUSTOWER.NS",
    "INFY.NS", "INOXWIND.NS", "INTELLECT.NS", "IOB.NS", "IOC.NS", "IPCALAB.NS", "IRB.NS", "IRCON.NS",
    "IRCTC.NS", "IRFC.NS", "ISEC.NS", "ITC.NS", "ITI.NS", "J&KBANK.NS", "JAIBALAJI.NS", "JBCHEPHARM.NS",
    "JBMA.NS", "JINDALSAW.NS", "JINDALSTEL.NS", "JIOFIN.NS", "JKCEMENT.NS", "JKLAKSHMI.NS", "JKPAPER.NS", "JMFINANCIL.NS",
    "JSL.NS", "JSWENERGY.NS", "JSWINFRA.NS", "JSWSTEEL.NS", "JUBLFOOD.NS", "JUBLINGREA.NS", "JUBLPHARMA.NS", "JUSTDIAL.NS",
    "JWL.NS", "JYOTHYLAB.NS", "KAJARIACER.NS", "KALYANKJIL.NS", "KANSAINER.NS", "KARURVYSYA.NS", "KAYNES.NS", "KEC.NS",
    "KEI.NS", "KFINTECH.NS", "KIMS.NS", "KNRCON.NS", "KOTAKBANK.NS", "KPIL.NS", "KPITTECH.NS", "KPRMILL.NS",
    "KRBL.NS", "KSB.NS", "LALPATHLAB.NS", "LATENTVIEW.NS", "LAURUSLABS.NS", "LEMONTREE.NS", "LICHSGFIN.NS", "LICI.NS",
    "LINDEINDIA.NS", "LLOYDSME.NS", "LODHA.NS", "LT.NS", "LTF.NS", "LTIM.NS", "LTTS.NS", "LUPIN.NS",
    "LXCHEM.NS", "M&M.NS", "M&MFIN.NS", "MAHABANK.NS", "MAHLIFE.NS", "MAHSEAMLES.NS", "MANAPPURAM.NS", "MANKIND.NS",
    "MANYAVAR.NS", "MAPMYINDIA.NS", "MARICO.NS", "MARUTI.NS", "MASTEK.NS", "MAXHEALTH.NS", "MAZDOCK.NS", "MCX.NS",
    "MEDANTA.NS", "MEDPLUS.NS", "METROBRAND.NS", "METROPOLIS.NS", "MFSL.NS", "MGL.NS", "MHRIL.NS", "MINDACORP.NS",
    "MMTC.NS", "MOTHERSON.NS", "MOTILALOFS.NS", "MPHASIS.NS", "MRF.NS", "MRPL.NS", "MSUMI.NS", "MTARTECH.NS",
    "MUTHOOTFIN.NS", "NAM-INDIA.NS", "NATCOPHARM.NS", "NATIONALUM.NS", "NAUKRI.NS", "NAVINFLUOR.NS", "NBCC.NS", "NCC.NS",
    "NESTLEIND.NS", "NETWORK18.NS", "NH.NS", "NHPC.NS", "NIACL.NS", "NLCINDIA.NS", "NMDC.NS", "NSLNISP.NS",
    "NTPC.NS", "NUVAMA.NS", "NUVOCO.NS", "NYKAA.NS", "OBEROIRLTY.NS", "OFSS.NS", "OIL.NS", "OLECTRA.NS",
    "ONGC.NS", "PAGEIND.NS", "PATANJALI.NS", "PAYTM.NS", "PCBL.NS", "PEL.NS", "PERSISTENT.NS", "PETRONET.NS",
    "PFC.NS", "PGHH.NS", "PHOENIXLTD.NS", "PIDILITIND.NS", "PIIND.NS", "PNB.NS", "PNBHOUSING.NS", "PNCINFRA.NS",
    "POLICYBZR.NS", "POLYCAB.NS", "POLYMED.NS", "POONAWALLA.NS", "POWERGRID.NS", "POWERINDIA.NS", "PPLPHARMA.NS", "PRAJIND.NS",
    "PRESTIGE.NS", "PRINCEPIPE.NS", "PRSMJOHNSN.NS", "PVRINOX.NS", "QUESS.NS", "RADICO.NS", "RAILTEL.NS", "RAINBOW.NS",
    "RAJESHEXPO.NS", "RAMCOCEM.NS", "RATNAMANI.NS", "RAYMOND.NS", "RBA.NS", "RBLBANK.NS", "RCF.NS", "RECLTD.NS",
    "REDINGTON.NS", "RELIANCE.NS", "RENUKA.NS", "RHIM.NS", "RITES.NS", "RKFORGE.NS", "ROUTE.NS", "RRKABEL.NS",
    "RTNINDIA.NS", "RVNL.NS", "SAFARI.NS", "SAIL.NS", "SAMMAANCAP.NS", "SANOFI.NS", "SAPPHIRE.NS", "SAREGAMA.NS",
    "SBFC.NS", "SBICARD.NS", "SBILIFE.NS", "SBIN.NS", "SCHAEFFLER.NS", "SCHNEIDER.NS", "SHREECEM.NS", "SHRIRAMFIN.NS",
    "SHYAMMETL.NS", "SIEMENS.NS", "SIGNATURE.NS", "SJVN.NS", "SKFINDIA.NS", "SOBHA.NS", "SOLARINDS.NS", "SONACOMS.NS",
    "SONATSOFTW.NS", "SPARC.NS", "SRF.NS", "STARHEALTH.NS", "STLTECH.NS", "SUMICHEM.NS", "SUNDARMFIN.NS", "SUNDRMFAST.NS",
    "SUNPHARMA.NS", "SUNTECK.NS", "SUNTV.NS", "SUPREMEIND.NS", "SUVENPHAR.NS", "SUZLON.NS", "SWANENERGY.NS", "SWSOLAR.NS",
    "SYNGENE.NS", "SYRMA.NS", "TANLA.NS", "TATACHEM.NS", "TATACOMM.NS", "TATACONSUM.NS", "TATAELXSI.NS", "TATAINVEST.NS",
    "TATAMOTORS.NS", "TATAPOWER.NS", "TATASTEEL.NS", "TATATECH.NS", "TBOTEK.NS", "TCS.NS", "TECHM.NS", "TEJASNET.NS",
    "THERMAX.NS", "TIINDIA.NS", "TIMKEN.NS", "TITAGARH.NS", "TITAN.NS", "TMB.NS", "TORNTPHARM.NS", "TORNTPOWER.NS",
    "TRENT.NS", "TRIDENT.NS", "TRITURBINE.NS", "TRIVENI.NS", "TTML.NS", "TVSMOTOR.NS", "TVSSCS.NS", "UBL.NS",
    "UCOBANK.NS", "UJJIVANSFB.NS", "ULTRACEMCO.NS", "UNIONBANK.NS", "UNITDSPR.NS", "UNOMINDA.NS", "UPL.NS", "USHAMART.NS",
    "UTIAMC.NS", "VAIBHAVGBL.NS", "VARROC.NS", "VBL.NS", "VEDL.NS", "VGUARD.NS", "VIJAYA.NS", "VIPIND.NS",
    "VOLTAS.NS", "VTL.NS", "WELCORP.NS", "WELSPUNLIV.NS", "WESTLIFE.NS", "WHIRLPOOL.NS", "WIPRO.NS", "YESBANK.NS",
    "ZEEL.NS", "ZENSARTECH.NS", "ZFCVINDIA.NS", "ZYDUSLIFE.NS",
])))

# =============================================================================
# LOGGING
# =============================================================================

def setup_logging() -> logging.Logger:
    logger = logging.getLogger("nifty200_scanner")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        handler = logging.StreamHandler()
        handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
        logger.addHandler(handler)
    return logger

logger = setup_logging()

# =============================================================================
# UTILITIES
# =============================================================================

def hash_params(**kwargs) -> str:
    param_str = json.dumps(kwargs, sort_keys=True, default=str)
    return hashlib.md5(param_str.encode()).hexdigest()

def retry_on_error(max_retries: int = 3, delay: float = 1.0) -> Callable:
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except Exception:
                    if attempt == max_retries - 1:
                        raise
                    time.sleep(delay * (attempt + 1))
            return None
        return wrapper
    return decorator


def _is_rate_limit_error(exc: BaseException) -> bool:
    """
    Detect Yahoo/yfinance rate-limit signatures across the different
    exception shapes it raises (requests.HTTPError, YFRateLimitError,
    generic Exception with '429' in the message, etc).
    """
    msg = str(exc).lower()
    type_name = type(exc).__name__.lower()
    markers = ("429", "too many requests", "rate limit", "ratelimit")
    return any(m in msg for m in markers) or "ratelimit" in type_name

# =============================================================================
# DATA MANAGER (with cache expiry and batch download)
# =============================================================================

class DataManager:
    """Manages stock data download, caching, and retrieval with batch support."""

    def __init__(self, cache_dir: Path = CACHE_DIR, cache_ttl_hours: int = 24):
        self.cache_dir = cache_dir
        try:
            self.cache_dir.mkdir(exist_ok=True)
        except (PermissionError, OSError):
            logger.warning("Cache dir %s not writable; using in-memory cache only.", self.cache_dir)
        self._data_cache: Dict[str, pd.DataFrame] = {}
        self.cache_ttl = timedelta(hours=cache_ttl_hours)
        self.last_rate_limit_hits: int = 0

    def _get_cache_path(self, symbol: str, period: str, interval: str) -> Path:
        cache_key = hash_params(symbol=symbol, period=period, interval=interval)
        return self.cache_dir / f"{cache_key}.parquet"

    def _get_metadata_path(self, symbol: str) -> Path:
        return self.cache_dir / f"{symbol}_metadata.json"

    @staticmethod
    def _expected_bars(period: str) -> int:
        """Rough number of trading days a period should contain."""
        return StrategyConfig.PERIOD_TRADING_DAYS.get(period, 252)

    @classmethod
    def _looks_truncated(cls, df: Optional[pd.DataFrame], period: str) -> bool:
        """
        True if a frame is suspiciously short for the requested period.
        Yahoo, when rate-limited, often returns a PARTIAL frame instead of an
        error. Caching such a frame poisons every run for the next 24h, so we
        must detect it. Threshold is 50% of expected bars - lenient enough
        that genuinely newly-listed stocks (short history) usually pass after
        a confirming retry.
        """
        if df is None or df.empty:
            return True
        return len(df) < cls._expected_bars(period) * 0.5

    @staticmethod
    def _last_bar_date(df: pd.DataFrame):
        if 'date' in df.columns:
            return pd.to_datetime(df['date'].iloc[-1])
        return pd.to_datetime(df.index[-1])

    @classmethod
    def _is_stale(cls, df: pd.DataFrame, max_age_days: int = 7) -> bool:
        """True if the most recent bar is well before today. 7 days tolerates
        weekends + holidays. A genuinely newly-listed stock is short but NOT
        stale (its data runs right up to the latest session); a throttled
        partial of an old stock is typically short AND stale."""
        try:
            last = cls._last_bar_date(df).normalize()
        except Exception:
            return True
        return (pd.Timestamp.now().normalize() - last).days > max_age_days

    @classmethod
    def _cacheable(cls, df: Optional[pd.DataFrame], period: str) -> bool:
        """Whether a frame is worth writing to / serving from the 24h cache.

        Caches when the frame has enough bars OR when it is a SHORT history that
        is still CURRENT — the signature of a genuinely newly-listed stock. This
        fixes 'cache starvation' where such stocks were never cached and forced
        a fresh network request on every run. A short AND stale frame (the usual
        shape of a throttled/partial Yahoo response) is still refused, and the
        live-fetch path additionally retries truncated responses with an
        explicit start date before this decision is reached."""
        if df is None or df.empty:
            return False
        if not cls._looks_truncated(df, period):
            return True
        return not cls._is_stale(df)

    @staticmethod
    def _period_to_start_date(period: str) -> Optional[str]:
        """Convert a yfinance period string to an explicit start date, adding
        ~25% margin so warmup bars are always covered."""
        days_map = {"6mo": 230, "1y": 460, "2y": 920, "3y": 1370, "5y": 2280, "10y": 4560}
        days = days_map.get(period)
        if days is None:
            return None  # "max" - let yfinance handle it
        return (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d")

    def _normalize_frame(self, df: pd.DataFrame) -> Optional[pd.DataFrame]:
        """Shared post-processing: flatten columns, lower-case, tz-strip,
        validate OHLC. Returns None if the frame is unusable."""
        if df is None or df.empty:
            return None

        df = df.reset_index()

        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [
                next((str(x) for x in col if x not in (None, '')), '')
                for col in df.columns
            ]

        df.columns = [str(c).replace(' ', '_').lower() for c in df.columns]

        if 'date' not in df.columns and 'datetime' in df.columns:
            df = df.rename(columns={'datetime': 'date'})
        if 'date' not in df.columns and 'index' in df.columns:
            df = df.rename(columns={'index': 'date'})

        if 'date' in df.columns:
            df['date'] = pd.to_datetime(df['date'])
            if getattr(df['date'].dt, 'tz', None) is not None:
                df['date'] = df['date'].dt.tz_localize(None)

        required_cols = {'open', 'high', 'low', 'close'}
        if not required_cols.issubset(df.columns):
            return None

        # Drop rows where OHLC is entirely NaN (bulk downloads pad missing
        # symbols with NaN rows)
        df = df.dropna(subset=['open', 'high', 'low', 'close'], how='all')
        if df.empty:
            return None
        return df

    def _is_cache_valid(self, cache_path: Path) -> bool:
        if not cache_path.exists():
            return False
        mtime = datetime.fromtimestamp(cache_path.stat().st_mtime)
        return datetime.now() - mtime < self.cache_ttl

    @retry_on_error(max_retries=3, delay=1.0)
    def download_stock_data(
        self,
        symbol: str,
        period: str = "2y",
        interval: str = "1d",
        force_refresh: bool = False
    ) -> Optional[pd.DataFrame]:
        cache_path = self._get_cache_path(symbol, period, interval)
        cache_key = f"{symbol}_{period}_{interval}"

        # BUGFIX (poisoned cache): a rate-limited Yahoo response is often a
        # SHORT PARTIAL frame, not an error. The old code cached it for 24h,
        # so every later run "downloaded" 20-60 bars per symbol and the
        # backtest reported "fewer than N bars of history" forever. Now any
        # cached frame that looks truncated is treated as INVALID and
        # re-downloaded.
        if cache_key in self._data_cache and not force_refresh:
            cached = self._data_cache[cache_key]
            if self._cacheable(cached, period):
                return cached.copy()
            del self._data_cache[cache_key]

        if cache_path.exists() and not force_refresh and self._is_cache_valid(cache_path):
            try:
                df = pd.read_parquet(cache_path)
                if self._cacheable(df, period):
                    self._data_cache[cache_key] = df
                    return df.copy()
                # Truncated cache entry: delete it so it can't poison again
                cache_path.unlink(missing_ok=True)
            except Exception:
                pass

        try:
            ticker = yf.Ticker(symbol)
            df = self._normalize_frame(ticker.history(period=period, interval=interval))

            # RETRY with an explicit start date if the period request came back
            # truncated. Yahoo sometimes honours start/end when it shortchanges
            # range/period requests.
            if self._looks_truncated(df, period):
                start = self._period_to_start_date(period)
                if start:
                    time.sleep(0.5)
                    df2 = self._normalize_frame(
                        ticker.history(start=start, interval=interval)
                    )
                    # Keep whichever attempt returned MORE history
                    if df2 is not None and (df is None or len(df2) > len(df)):
                        df = df2

            if df is None or df.empty:
                return None

            # Only CACHE the frame if it doesn't look truncated. A short frame
            # can still be returned for THIS run (a genuinely newly-listed
            # stock has short history), but it must never be written to the
            # 24h cache where a throttled partial response would poison every
            # subsequent run.
            if self._cacheable(df, period):
                try:
                    df.to_parquet(cache_path)
                except Exception as e:
                    logger.warning("Could not write cache for %s: %s", symbol, e)
                self._data_cache[cache_key] = df
            else:
                logger.warning(
                    f"{symbol}: only {len(df)} bars for period={period} "
                    f"(expected ~{self._expected_bars(period)}); not caching."
                )

            try:
                info = ticker.info
                metadata = {
                    'symbol': symbol,
                    'name': info.get('longName', symbol),
                    'sector': info.get('sector', 'Unknown'),
                    'market_cap': info.get('marketCap', 0),
                    'currency': info.get('currency', 'INR'),
                    'last_updated': datetime.now().isoformat()
                }
            except Exception:
                metadata = {
                    'symbol': symbol, 'name': symbol, 'sector': 'Unknown',
                    'market_cap': 0, 'currency': 'INR',
                    'last_updated': datetime.now().isoformat()
                }

            with open(self._get_metadata_path(symbol), 'w') as f:
                json.dump(metadata, f)

            return df.copy()

        except Exception as e:
            if _is_rate_limit_error(e):
                # Let this propagate - callers (batch_download's fallback
                # pass) need to see it to trigger their own backoff/retry.
                # Swallowing it here would make rate limiting invisible.
                raise
            logger.error(f"Failed to download {symbol}: {e}")
            return None

    def batch_download(
        self,
        symbols: List[str],
        period: str = "2y",
        interval: str = "1d",
        max_workers: int = 8,
        progress_callback: Optional[Callable] = None,
        status_callback: Optional[Callable] = None,
    ) -> Dict[str, pd.DataFrame]:
        """
        Download multiple symbols with ADAPTIVE rate-limit handling.

        1. Serve everything possible from valid cache (no network at all).
        2. Bulk-download the rest with yf.download() in small chunks.
           Chunk size shrinks and the inter-chunk cooldown grows
           exponentially the moment a 429 / rate-limit response is
           detected, instead of plowing ahead on a fixed 0.5s sleep.
        3. Fall back to a slow, staggered per-ticker pass (low
           concurrency + jittered delay between requests, with its own
           backoff-on-429) for anything the bulk pass still missed.

        status_callback(str) is called with human-readable progress text
        (e.g. "Rate limited - cooling down 8s...") so the UI can show
        the person WHY a run is slow, instead of it looking stuck.
        """
        results: Dict[str, pd.DataFrame] = {}
        remaining: List[str] = []
        total = max(len(symbols), 1)
        self.last_rate_limit_hits = 0

        def status(msg: str):
            if status_callback:
                status_callback(msg)

        # --- Pass 1: valid cache (zero network calls) ---
        for sym in symbols:
            cache_path = self._get_cache_path(sym, period, interval)
            cache_key = f"{sym}_{period}_{interval}"
            df = None
            if cache_key in self._data_cache:
                df = self._data_cache[cache_key]
            elif cache_path.exists() and self._is_cache_valid(cache_path):
                try:
                    df = pd.read_parquet(cache_path)
                except Exception:
                    df = None
            if df is not None and self._cacheable(df, period):
                self._data_cache[cache_key] = df
                results[sym] = df.copy()
            else:
                remaining.append(sym)

        if progress_callback:
            progress_callback(len(results) / total)
        status(f"{len(results)}/{total} served from cache. "
               f"{len(remaining)} left to download.")

        # --- Pass 2: chunked bulk download with adaptive backoff ---
        BASE_CHUNK = 25          # smaller than before - fewer symbols per HTTP call
        MIN_CHUNK = 5
        MAX_CHUNK_RETRIES = 8    # give up on a chunk after this many rate-limit retries
        base_cooldown = 1.5      # seconds between chunks when things are healthy
        cooldown = base_cooldown
        chunk_size = BASE_CHUNK
        idx = 0
        chunk_retries = 0        # consecutive rate-limit retries at the current idx

        while idx < len(remaining):
            chunk = remaining[idx: idx + chunk_size]
            rate_limited = False
            try:
                bulk = yf.download(
                    tickers=" ".join(chunk),
                    period=period,
                    interval=interval,
                    group_by="ticker",
                    auto_adjust=True,
                    threads=False,
                    progress=False,
                )
            except Exception as e:
                bulk = None
                if _is_rate_limit_error(e):
                    rate_limited = True
                else:
                    logger.error(f"Bulk download failed for chunk starting {chunk[0]}: {e}")

            got_any = False
            if bulk is not None and not bulk.empty:
                for sym in chunk:
                    try:
                        if isinstance(bulk.columns, pd.MultiIndex):
                            if sym not in bulk.columns.get_level_values(0):
                                continue
                            sub = bulk[sym]
                        else:
                            sub = bulk  # single-symbol chunk: flat columns
                        df = self._normalize_frame(sub)
                        if df is None:
                            continue
                        if self._cacheable(df, period):
                            cache_path = self._get_cache_path(sym, period, interval)
                            try:
                                df.to_parquet(cache_path)
                            except Exception as e:
                                logger.warning("Could not write cache for %s: %s", sym, e)
                            self._data_cache[f"{sym}_{period}_{interval}"] = df
                            results[sym] = df.copy()
                            got_any = True
                        # truncated bulk result -> leave for per-ticker retry
                    except Exception as e:
                        logger.error(f"Parsing bulk data for {sym}: {e}")

            # A chunk that returns nothing at all for every symbol, right
            # after a healthy previous chunk, is also a rate-limit signature
            # even if yfinance didn't raise (it can return an empty frame
            # silently instead of throwing).
            if not rate_limited and not got_any and bulk is not None and bulk.empty:
                rate_limited = True

            if rate_limited:
                self.last_rate_limit_hits += 1
                chunk_retries += 1
                if chunk_retries > MAX_CHUNK_RETRIES:
                    # Persistent rate limiting on this chunk: stop retrying so
                    # the run can't hang forever. Leave these symbols
                    # undownloaded (they simply won't appear in results) and
                    # move on to the next chunk.
                    logger.warning(
                        "Giving up on chunk starting %s after %d rate-limit retries.",
                        chunk[0], chunk_retries,
                    )
                    status(f"Skipping {len(chunk)} symbol(s) after repeated rate limits.")
                    idx += len(chunk)
                    chunk_retries = 0
                    chunk_size = BASE_CHUNK
                    cooldown = base_cooldown
                    continue
                cooldown = min(cooldown * 2, 60.0)          # exponential backoff, capped at 60s
                chunk_size = max(chunk_size // 2, MIN_CHUNK)  # shrink chunk under pressure
                status(f"Rate limited by Yahoo - cooling down "
                       f"{cooldown:.0f}s and shrinking batch size to "
                       f"{chunk_size}...")
                time.sleep(cooldown + random.uniform(0, 1.0))
                # Don't advance idx - retry this same chunk (now smaller next loop)
                continue
            else:
                # Healthy chunk: ease cooldown back down and step forward
                cooldown = max(cooldown * 0.7, base_cooldown)
                chunk_retries = 0
                idx += len(chunk)

            if progress_callback:
                progress_callback(min((len(results)) / total, 0.95))
            time.sleep(cooldown + random.uniform(0, 0.5))  # gentle, slightly jittered gap

        # --- Pass 3: slow, staggered per-ticker fallback for anything still missing ---
        still_missing = [s for s in remaining if s not in results]
        completed = len(results)

        if still_missing:
            status(f"Bulk pass done. {len(still_missing)} symbols still "
                   f"missing - retrying individually (slower, to stay under "
                   f"Yahoo's rate limit)...")

        def process_with_backoff(sym: str) -> Tuple[str, Optional[pd.DataFrame]]:
            local_delay = 1.0
            for attempt in range(3):
                try:
                    return sym, self.download_stock_data(sym, period=period, interval=interval)
                except Exception as e:
                    if _is_rate_limit_error(e) and attempt < 2:
                        time.sleep(local_delay + random.uniform(0, 1.0))
                        local_delay *= 2
                        continue
                    logger.error(f"Batch fallback error for {sym}: {e}")
                    return sym, None
            return sym, None

        if still_missing:
            # Very low concurrency + small stagger between submissions so we
            # never fire a burst of near-simultaneous requests at Yahoo.
            with ThreadPoolExecutor(max_workers=min(max_workers, 2)) as executor:
                futures = {}
                for sym in still_missing:
                    futures[executor.submit(process_with_backoff, sym)] = sym
                    time.sleep(0.3)  # stagger submissions
                for future in as_completed(futures):
                    sym, df = future.result()
                    completed += 1
                    if df is not None:
                        results[sym] = df
                    if progress_callback:
                        progress_callback(min(completed / total, 1.0))

        if progress_callback:
            progress_callback(1.0)

        if self.last_rate_limit_hits:
            status(f"Done. Hit Yahoo's rate limit {self.last_rate_limit_hits} "
                   f"time(s) during this run - backed off automatically. "
                   f"{len(results)}/{total} symbols downloaded.")
        else:
            status(f"Done. {len(results)}/{total} symbols downloaded, no rate limiting.")

        return results

    def get_stock_metadata(self, symbol: str) -> dict:
        metadata_path = self._get_metadata_path(symbol)
        if metadata_path.exists():
            try:
                with open(metadata_path, 'r') as f:
                    return json.load(f)
            except Exception:
                pass
        return {'symbol': symbol, 'name': symbol, 'sector': 'Unknown', 'market_cap': 0}

    def clear_cache(self) -> None:
        for f in self.cache_dir.glob("*.parquet"):
            f.unlink()
        for f in self.cache_dir.glob("*.json"):
            f.unlink()
        self._data_cache.clear()

# =============================================================================
# TECHNICAL INDICATORS
# =============================================================================

class TechnicalIndicators:
    """Vectorized technical indicator calculations."""

    @staticmethod
    def rsi(prices: pd.Series, period: int = 14) -> pd.Series:
        delta = prices.diff()
        gain = delta.where(delta > 0, 0.0)
        loss = (-delta).where(delta < 0, 0.0)
        avg_gain = gain.ewm(alpha=1/period, min_periods=period, adjust=False).mean()
        avg_loss = loss.ewm(alpha=1/period, min_periods=period, adjust=False).mean()
        rs = avg_gain / avg_loss
        return 100 - (100 / (1 + rs))

    @staticmethod
    def ema(prices: pd.Series, period: int = 200) -> pd.Series:
        """Exponential moving average.

        Uses adjust=False (standard recursive EMA) and min_periods=period so
        the series stays NaN until `period` completed candles exist. This is
        what prevents early-bar repainting / spurious signals during warmup.
        """
        return prices.ewm(span=period, adjust=False, min_periods=period).mean()

    @staticmethod
    def sma(prices: pd.Series, period: int = 50) -> pd.Series:
        return prices.rolling(window=period, min_periods=period).mean()

# =============================================================================
# STRATEGY CONFIGURATION
# =============================================================================

@dataclass
class StrategyConfig:
    """Configuration for the trading strategy."""
    rsi_period: int = 14
    rsi_threshold: float = 35.0
    # RSI Tracking + 200 EMA Trend Confirmation strategy parameters.
    # ema_long  = the trend filter the price must be BELOW to arm tracking
    #             (with RSI < threshold) and ABOVE to trigger the buy.
    # ema_mid / ema_short = the fast-cross confirmation (ema_mid > ema_short).
    ema_long: int = 200
    ema_mid: int = 50
    ema_short: int = 21
    profit_target_pct: float = 3.14
    average_trigger_pct: float = 2.0
    max_average_entries: int = 5
    position_size: float = 50000.0
    initial_capital: float = 1000000.0
    max_positions: int = 20
    brokerage_pct: float = 0.05
    slippage_pct: float = 0.05
    transaction_charges_pct: float = 0.01
    data_period: str = "2y"  # NEW: how much history to download for scanning/backtesting

    # --- Optional Stop Loss module (user-controlled, OFF by default) ---
    # When enabled, this REPLACES averaging entirely for that run (the two
    # risk-management approaches are mutually exclusive - see BacktestEngine).
    # Entry logic, profit target, and everything else is unaffected.
    stop_loss_mode: bool = False
    stop_loss_pct: float = 3.0

    VALID_PERIODS = ("6mo", "1y", "2y", "3y", "5y", "10y", "max")
    # Approximate trading days per period, used only to warn if a period is
    # too short for the configured indicator warmup (SMA / highest-close lookback).
    PERIOD_TRADING_DAYS = {
        "6mo": 125, "1y": 252, "2y": 504, "3y": 756,
        "5y": 1260, "10y": 2520, "max": 10_000,
    }

    def __post_init__(self):
        assert self.rsi_period > 0
        assert 0 < self.rsi_threshold < 100
        assert self.ema_long > 0
        assert self.ema_mid > 0
        assert self.ema_short > 0
        # The strategy's buy alignment (Close > EMA_long > EMA_mid > EMA_short)
        # only makes sense when the PERIODS are ordered short < mid < long.
        # This is a soft check (warn, don't crash) so an unusual config just
        # produces few/no signals rather than failing app startup.
        if not (self.ema_short < self.ema_mid < self.ema_long):
            logger.warning(
                "StrategyConfig: EMA periods are not short<mid<long "
                "(short=%s, mid=%s, long=%s); signal conditions may never align.",
                self.ema_short, self.ema_mid, self.ema_long,
            )
        assert self.profit_target_pct > 0
        assert self.average_trigger_pct > 0
        assert self.max_average_entries >= 0
        assert self.position_size > 0
        assert self.initial_capital > 0
        assert self.max_positions > 0
        assert self.data_period in self.VALID_PERIODS
        assert 0.5 <= self.stop_loss_pct <= 20.0

    @property
    def total_transaction_cost_pct(self) -> float:
        return self.brokerage_pct + self.slippage_pct + self.transaction_charges_pct

    def calculate_entry_cost(self, amount: float) -> float:
        return amount * self.total_transaction_cost_pct / 100

    def calculate_exit_cost(self, amount: float) -> float:
        return amount * self.total_transaction_cost_pct / 100

    @property
    def min_required_bars(self) -> int:
        """Minimum bars needed before the first valid signal can be generated.

        Driven by the 200 EMA (the longest-warmup indicator): its value is
        NaN until `ema_long` completed candles exist, and the trend filter
        can't be evaluated until then.
        """
        return max(self.ema_long, self.ema_mid, self.ema_short) + 10

    @property
    def period_likely_insufficient(self) -> bool:
        """True if the selected data_period probably won't cover warmup + a
        reasonable number of tradeable bars afterward."""
        available = self.PERIOD_TRADING_DAYS.get(self.data_period, 10_000)
        return available < self.min_required_bars + 60  # leave room for actual trades


# =============================================================================
# SIGNAL GENERATOR (Vectorized, shared logic)
# =============================================================================

class SignalGenerator:
    """
    Shared signal generation logic used by both Scanner and BacktestEngine.
    All operations are fully vectorized for performance.
    """

    @staticmethod
    def generate_signals(df: pd.DataFrame, config: StrategyConfig) -> pd.DataFrame:
        """
        RSI Tracking + EMA Alignment — state machine.

        This strategy is inherently sequential (the state carried from one
        candle to the next decides the meaning of the next candle), so the
        core is a single forward pass. It is NOT a rolling window and must
        not be re-expressed as one: once armed, tracking persists across an
        arbitrary number of candles until the reclaim buy fires.

            NORMAL ──[RSI<thr AND Close<EMA200]──▶ TRACKING ──▶ BUY ─▶ reset ─▶ NORMAL

        The BUY candle must satisfy ALL FIVE conditions simultaneously:
            1. Tracking == True
            2. Close   > EMA200
            3. EMA50   > EMA21
            4. EMA50   < EMA200
            5. EMA21   < EMA200
        i.e. price has *just* closed back above the 200 EMA while both the
        50 and 21 EMAs are still below it (early reversal), with the 50
        still above the 21. Structure:  Close > EMA200 > EMA50 > EMA21.

        Guarantees (see module docstring):
          * The arming candle never buys (we `continue` past the buy check).
          * Further RSI dips while TRACKING are ignored (the arm condition is
            only ever tested while NOT tracking).
          * Exactly one buy per cycle; a fresh arm is required afterwards.
          * No repainting — indicators are NaN during warmup and such candles
            can neither arm nor buy.
        """
        df = df.copy()

        # ---- Indicators (every candle) ----
        df['rsi'] = TechnicalIndicators.rsi(df['close'], config.rsi_period)
        df['ema_long'] = TechnicalIndicators.ema(df['close'], config.ema_long)     # EMA(200)
        df['ema_mid'] = TechnicalIndicators.ema(df['close'], config.ema_mid)       # EMA(50)
        df['ema_short'] = TechnicalIndicators.ema(df['close'], config.ema_short)   # EMA(21)

        n = len(df)
        close_v = df['close'].values
        rsi_v = df['rsi'].values
        ema_long_v = df['ema_long'].values
        ema_mid_v = df['ema_mid'].values
        ema_short_v = df['ema_short'].values

        # Arming condition, per candle: RSI < threshold AND Close < EMA(200).
        # (NaN comparisons during warmup are False, so warmup never arms.)
        arm_v = (rsi_v < config.rsi_threshold) & (close_v < ema_long_v)

        tracking_started = np.zeros(n, dtype=bool)   # marker: the arming candle
        tracking_active = np.zeros(n, dtype=bool)     # True on every candle the
                                                      # machine is in TRACKING,
                                                      # including arm & buy candle
        buy_signal = np.zeros(n, dtype=bool)          # the reclaim buy candle

        tracking = False
        for i in range(n):
            # Warmup guard: if the trend/RSI can't be evaluated yet, this
            # candle can neither arm nor buy. Carry state unchanged.
            if np.isnan(ema_long_v[i]) or np.isnan(ema_mid_v[i]) \
                    or np.isnan(ema_short_v[i]) or np.isnan(rsi_v[i]):
                tracking_active[i] = tracking
                continue

            if not tracking:
                # NORMAL state — the only place the RSI arm is ever tested.
                if arm_v[i]:
                    tracking = True
                    tracking_started[i] = True
                    tracking_active[i] = True
                    # Arming candle NEVER buys — stop processing this candle.
                    continue
                # Still NORMAL, nothing to do (tracking_active stays False).
                continue

            # TRACKING state — ignore any further RSI signals; only look for
            # the reclaim buy. All five conditions must hold on THIS candle:
            #   Close > EMA200, EMA50 > EMA21, EMA50 < EMA200, EMA21 < EMA200
            # (Tracking == True is guaranteed by being in this branch.)
            tracking_active[i] = True
            if (close_v[i] > ema_long_v[i]
                    and ema_mid_v[i] > ema_short_v[i]
                    and ema_mid_v[i] < ema_long_v[i]
                    and ema_short_v[i] < ema_long_v[i]):
                buy_signal[i] = True
                tracking = False  # reset — one buy per cycle

        df['tracking_started'] = tracking_started
        df['tracking_active'] = tracking_active
        # Backward-compatible alias used by the scanner/UI.
        df['tracking_enabled'] = tracking_active
        df['buy_signal'] = buy_signal

        # Human-readable condition flags (handy for the scanner + debugging).
        df['close_above_ema_long'] = df['close'] > df['ema_long']
        df['ema_mid_gt_short'] = df['ema_mid'] > df['ema_short']
        df['ema_mid_below_long'] = df['ema_mid'] < df['ema_long']
        df['ema_short_below_long'] = df['ema_short'] < df['ema_long']
        # Full early-reversal EMA alignment: Close > EMA200 > EMA50 > EMA21.
        df['ema_alignment_ok'] = (
            df['close_above_ema_long'] & df['ema_mid_gt_short']
            & df['ema_mid_below_long'] & df['ema_short_below_long']
        )

        return df

    @staticmethod
    def get_latest_signal(df: pd.DataFrame, config: StrategyConfig) -> Optional[Dict]:
        """Get the latest signal state for scanner."""
        if df.empty or len(df) < config.min_required_bars:
            return None

        df_signals = SignalGenerator.generate_signals(df, config)
        latest = df_signals.iloc[-1]

        close = float(latest['close'])
        rsi_val = float(latest['rsi']) if not pd.isna(latest['rsi']) else 50
        ema_long = float(latest['ema_long']) if not pd.isna(latest['ema_long']) else close
        ema_mid = float(latest['ema_mid']) if not pd.isna(latest['ema_mid']) else close
        ema_short = float(latest['ema_short']) if not pd.isna(latest['ema_short']) else close

        return {
            'symbol': '',
            'close': close,
            'rsi': rsi_val,
            'ema_long': ema_long,     # EMA(200)
            'ema_mid': ema_mid,       # EMA(50)
            'ema_short': ema_short,   # EMA(21)
            'close_above_ema_long': bool(latest['close_above_ema_long']),
            'ema_mid_gt_short': bool(latest['ema_mid_gt_short']),
            'ema_alignment_ok': bool(latest['ema_alignment_ok']),
            'tracking_enabled': bool(latest['tracking_active']),
            'tracking_started': bool(latest['tracking_started']),
            'buy_signal': bool(latest['buy_signal']),
            'volume': float(latest.get('volume', 0)),
            'date': latest['date'] if 'date' in latest else df.index[-1]
        }

# =============================================================================
# TRADE & PORTFOLIO
# =============================================================================

@dataclass
class Position:
    """Represents an open position in a single stock."""
    symbol: str
    entry_dates: List[datetime] = field(default_factory=list)      # Execution dates (T+1 open fill)
    entry_prices: List[float] = field(default_factory=list)        # Execution prices (T+1 open)
    entry_quantities: List[float] = field(default_factory=list)
    entry_costs: List[float] = field(default_factory=list)
    signal_dates: List[datetime] = field(default_factory=list)     # Day T signal candle date
    signal_closes: List[float] = field(default_factory=list)       # Day T signal candle close
    average_count: int = 0
    total_invested: float = 0.0
    total_quantity: float = 0.0

    @property
    def average_entry_price(self) -> float:
        if self.total_quantity == 0:
            return 0.0
        return self.total_invested / self.total_quantity

    def add_entry(self, date: datetime, price: float, quantity: float, cost: float,
                  signal_date: Optional[datetime] = None, signal_close: Optional[float] = None):
        # Guard against corrupt inputs: a zero/negative/NaN quantity or price
        # would desync average_count from the real share count and poison P&L.
        if not (np.isfinite(quantity) and quantity > 0) or not (np.isfinite(price) and price > 0):
            logger.warning(
                "Position.add_entry ignored invalid entry (price=%s, qty=%s) for %s",
                price, quantity, self.symbol,
            )
            return
        self.entry_dates.append(date)
        self.entry_prices.append(price)
        self.entry_quantities.append(quantity)
        self.entry_costs.append(cost)
        self.signal_dates.append(signal_date if signal_date is not None else date)
        self.signal_closes.append(signal_close if signal_close is not None else price)
        self.total_invested += (price * quantity) + cost
        self.total_quantity += quantity
        self.average_count += 1

    def calculate_unrealized_pnl(self, current_price: float, config: StrategyConfig) -> Tuple[float, float]:
        if self.total_quantity == 0:
            return 0.0, 0.0

        current_value = self.total_quantity * current_price
        exit_cost = config.calculate_exit_cost(current_value)
        net_value = current_value - exit_cost
        gross_pnl = net_value - self.total_invested
        profit_pct = (gross_pnl / self.total_invested) * 100 if self.total_invested > 0 else 0.0

        return profit_pct, gross_pnl

    def close_position(self, date: datetime, price: float, config: StrategyConfig) -> Tuple[float, float, float]:
        current_value = self.total_quantity * price
        exit_cost = config.calculate_exit_cost(current_value)
        net_value = current_value - exit_cost
        gross_pnl = net_value - self.total_invested
        profit_pct = (gross_pnl / self.total_invested) * 100 if self.total_invested > 0 else 0.0

        return profit_pct, gross_pnl, net_value

    # ---------------------------------------------------------------
    # MODIFIED EXIT LOGIC (averaged positions) — price-based target.
    #
    #   Entries == 1 : target = Buy Price * (1 + target_pct/100)
    #   Entries  > 1 : target = Weighted Average Buy Price * (1 + target_pct/100)
    #                  Weighted Average Buy Price = Total Cost of All
    #                  Entries (qty * price, no charges) / Total Quantity.
    #
    # Recalculated fresh from pos.entry_prices / pos.entry_quantities on
    # every call, so a new averaging entry automatically shifts the target.
    # This does not touch total_invested/total_quantity (which still
    # include transaction costs and drive cash/P&L accounting elsewhere).
    # ---------------------------------------------------------------
    def weighted_average_buy_price_raw(self) -> float:
        """Weighted average buy price from raw entry cost (qty*price), no charges."""
        if not self.entry_quantities or self.total_quantity == 0:
            return 0.0
        if self.average_count <= 1:
            return self.entry_prices[0]
        total_cost = sum(p * q for p, q in zip(self.entry_prices, self.entry_quantities))
        total_qty = sum(self.entry_quantities)
        return total_cost / total_qty if total_qty > 0 else 0.0

    def exit_trigger_price(self, config: StrategyConfig) -> float:
        """Price at/above which the entire (averaged) position should exit."""
        base_price = self.weighted_average_buy_price_raw()
        return base_price * (1 + config.profit_target_pct / 100)

    def stop_loss_trigger_price(self, config: StrategyConfig) -> float:
        """Fixed-percentage Stop Loss price, based on the ORIGINAL entry
        price. Only meaningful in Stop Loss mode, where averaging is
        disabled so there is always exactly one entry (entry_prices[0])."""
        if not self.entry_prices:
            return 0.0
        return self.entry_prices[0] * (1 - config.stop_loss_pct / 100)


@dataclass
class CompletedTrade:
    """A completed trade record."""
    symbol: str
    entry_dates: List[datetime]        # Execution dates (T+1 open fill)
    entry_prices: List[float]          # Execution prices (T+1 open)
    signal_dates: List[datetime]       # Signal candle dates (Day T close)
    signal_closes: List[float]         # Signal candle close prices
    average_count: int
    exit_date: Optional[datetime]
    exit_price: Optional[float]
    holding_days: int
    profit_pct: float
    profit_inr: float
    total_invested: float
    total_returned: float
    status: str
    exit_reason: str


@dataclass
class PortfolioState:
    """Tracks portfolio state day by day with realistic cash management."""
    cash: float
    positions: Dict[str, Position] = field(default_factory=dict)
    completed_trades: List[CompletedTrade] = field(default_factory=list)
    daily_equity: List[Tuple[datetime, float]] = field(default_factory=list)
    # NEW (additive): daily mark-to-market value of open positions only
    # (equity minus cash), used solely for the new "Capital Utilization"
    # report metric. Does not affect daily_equity/total_equity/cash.
    daily_invested: List[Tuple[datetime, float]] = field(default_factory=list)
    # FIX: tracks the most recent observed market price per symbol so that
    # missing-data days fall back to the last real price, not the entry price.
    last_known_prices: Dict[str, float] = field(default_factory=dict)
    # NEW (additive): tracks the highest observed mark-to-market value per
    # open symbol since it was opened, used solely for the "Current
    # Drawdown %" field on the Open Positions report. Reset (popped) when
    # a position is closed. Does not affect any other calculation.
    position_peak_value: Dict[str, float] = field(default_factory=dict)

    @property
    def total_equity(self) -> float:
        # FIX: mark-to-market using last known prices, not cost basis
        position_value = 0.0
        for symbol, pos in self.positions.items():
            price = self.last_known_prices.get(symbol, pos.entry_prices[-1] if pos.entry_prices else 0)
            position_value += pos.total_quantity * price
        return self.cash + position_value

    @property
    def open_position_count(self) -> int:
        return len(self.positions)

    def record_equity(self, date: datetime, prices: Dict[str, float], config: StrategyConfig):
        # FIX: update the last-known-price cache with today's observed prices
        self.last_known_prices.update(prices)

        position_value = 0.0
        for symbol, pos in self.positions.items():
            # FIX: fall back to the last known market price (not the stale entry
            # price) when today's price is missing for a held symbol
            current_price = prices.get(
                symbol,
                self.last_known_prices.get(symbol, pos.entry_prices[-1] if pos.entry_prices else 0)
            )
            current_value = pos.total_quantity * current_price
            exit_cost = config.calculate_exit_cost(current_value)
            position_value += current_value - exit_cost

            # NEW (additive): track this position's peak mark-to-market value
            # for the "Current Drawdown %" field in the Open Positions report.
            self.position_peak_value[symbol] = max(
                self.position_peak_value.get(symbol, current_value), current_value
            )

        equity = self.cash + position_value
        self.daily_equity.append((date, equity))
        self.daily_invested.append((date, position_value))

    def can_open_position(self, config: StrategyConfig) -> bool:
        position_size = config.position_size
        entry_cost = config.calculate_entry_cost(position_size)
        return self.cash >= position_size + entry_cost and self.open_position_count < config.max_positions

    def open_position(self, symbol: str, date: datetime, price: float, config: StrategyConfig,
                       signal_date: Optional[datetime] = None, signal_close: Optional[float] = None) -> bool:
        if not (np.isfinite(price) and price > 0):
            return False
        if not self.can_open_position(config):
            return False

        position_size = config.position_size
        quantity = position_size / price
        entry_cost = config.calculate_entry_cost(position_size)
        total_needed = position_size + entry_cost

        if self.cash < total_needed:
            return False

        self.cash -= total_needed

        if symbol not in self.positions:
            self.positions[symbol] = Position(symbol=symbol)

        self.positions[symbol].add_entry(date, price, quantity, entry_cost,
                                          signal_date=signal_date, signal_close=signal_close)
        return True

    def add_to_position(self, symbol: str, date: datetime, price: float, config: StrategyConfig,
                         signal_date: Optional[datetime] = None, signal_close: Optional[float] = None) -> bool:
        # FIX v3.1: Only check cash availability, ignore max_positions for averaging
        if not (np.isfinite(price) and price > 0):
            return False
        position_size = config.position_size
        quantity = position_size / price
        entry_cost = config.calculate_entry_cost(position_size)
        total_needed = position_size + entry_cost

        if self.cash < total_needed:
            return False

        self.cash -= total_needed
        self.positions[symbol].add_entry(date, price, quantity, entry_cost,
                                          signal_date=signal_date, signal_close=signal_close)
        return True

    def close_position(self, symbol: str, date: datetime, price: float, config: StrategyConfig, reason: str):
        if symbol not in self.positions:
            return
        if not (np.isfinite(price) and price > 0):
            logger.warning("close_position skipped for %s: invalid exit price %s", symbol, price)
            return

        pos = self.positions[symbol]
        profit_pct, profit_inr, net_returned = pos.close_position(date, price, config)
        holding_days = (date - pos.entry_dates[0]).days if pos.entry_dates else 0

        self.cash += net_returned

        trade = CompletedTrade(
            symbol=symbol,
            entry_dates=pos.entry_dates.copy(),
            entry_prices=pos.entry_prices.copy(),
            signal_dates=pos.signal_dates.copy(),
            signal_closes=pos.signal_closes.copy(),
            average_count=pos.average_count,
            exit_date=date,
            exit_price=price,
            holding_days=holding_days,
            profit_pct=profit_pct,
            profit_inr=profit_inr,
            total_invested=pos.total_invested,
            total_returned=net_returned,
            status="CLOSED",
            exit_reason=reason
        )

        self.completed_trades.append(trade)
        del self.positions[symbol]
        self.position_peak_value.pop(symbol, None)

    def close_all_positions(self, date: datetime, prices: Dict[str, float], config: StrategyConfig, reason: str = "End of Data"):
        for symbol in list(self.positions.keys()):
            price = prices.get(symbol, 0)
            if price > 0:
                self.close_position(symbol, date, price, config, reason)

# =============================================================================
# BACKTEST ENGINE (Optimized with date-indexed lookup)
# =============================================================================

def _is_valid_price(x: Any) -> bool:
    """True only for a finite, strictly-positive price. Guards the execution
    paths against bad market data (NaN / 0 / negative) that would otherwise
    corrupt fills and P&L."""
    try:
        xf = float(x)
    except (TypeError, ValueError):
        return False
    return np.isfinite(xf) and xf > 0


def _resolve_gap_or_touch_exit(open_price: float, extreme_price: float,
                                trigger_price: float, direction: str) -> Optional[float]:
    """
    BUGFIX: realistic same-day fill price for a price-level order (profit
    target sell, or stop loss), instead of exiting at the day's CLOSE
    whenever the close happened to be past the trigger. A resting order at
    `trigger_price` cannot fill better than what the market actually
    offered:
      direction='up'   (profit target): triggers when price rises to/through
                        trigger_price. Detected via the day's HIGH.
      direction='down' (stop loss):     triggers when price falls to/through
                        trigger_price. Detected via the day's LOW.
    If the day's OPEN itself already gapped past the trigger, the order
    fills AT THE OPEN (can't get the trigger price - the market skipped
    over it). Otherwise, if the day's extreme (high/low) reached the
    trigger, the order fills exactly AT the trigger price (no price
    improvement assumed). Returns None if the trigger wasn't reached today
    or if the day's prices are invalid (NaN / non-positive data errors).
    """
    if trigger_price <= 0:
        return None
    if not _is_valid_price(open_price) or not _is_valid_price(extreme_price):
        return None
    if direction == 'up':
        if open_price >= trigger_price:
            return open_price
        if extreme_price >= trigger_price:
            return trigger_price
        return None
    else:
        if open_price <= trigger_price:
            return open_price
        if extreme_price <= trigger_price:
            return trigger_price
        return None


class BacktestEngine:
    """Optimized portfolio-based backtesting engine with realistic T+1 execution."""

    # A pending T+1 order fills on the stock's NEXT available session. If the
    # stock is halted / has no bar that day, the order is carried forward across
    # up to this many missing sessions before being discarded (so a delisted or
    # long-halted name doesn't fill at a meaningless much-later price, but a
    # routine 1-2 day halt no longer silently drops the buy).
    STALE_ORDER_MAX_MISSES = 5

    def __init__(self, config: StrategyConfig):
        self.config = config
        self.last_diagnostics: Dict[str, Any] = {}
        self.last_daily_invested: List[Tuple[datetime, float]] = []
        # NEW (additive): snapshot of positions still open when the backtest
        # data ends. Positions are NEVER force-closed - see run_backtest().
        self.last_open_positions: List[Dict[str, Any]] = []

    def run_backtest(self, data_dict: Dict[str, pd.DataFrame]) -> Tuple[List[CompletedTrade], List[Tuple[datetime, float]]]:
        """
        Run portfolio-level backtest across all symbols.
        Uses date-indexed DataFrames for O(1) lookup.

        EXECUTION MODEL (T+1, unconditional open fill - per strategy spec):
        - Day T: Generate signals using completed candle (close). No trade is
          executed on the signal candle itself.
        - Day T: Store Signal Date + Signal Close, create a pending order.
        - Day T+1: Execute the pending order unconditionally at that day's
          OPEN price (no limit-price / no-fill logic - always fills as long
          as the symbol has data on T+1 and portfolio constraints allow it).

        For open positions (Averaging mode - config.stop_loss_mode == False):
        - Profit target is a resting sell-limit at entry/WAP * (1+target%).
          BUGFIX: previously triggered on CLOSE >= target and filled AT the
          close (which could overshoot the stated target by several % on a
          volatile day). Now resolved realistically via
          _resolve_gap_or_touch_exit(): fills at OPEN if the day gapped
          past the target, otherwise fills AT the target price itself if
          the day's HIGH touched it.
        - Averaging requires a genuinely NEW buy_signal on Day T (Condition 1)
          AND the Day T+1 OPEN execution price must be >=average_trigger_pct%
          below the previous EXECUTION price (Condition 2, checked at fill
          time - not the signal-day close). Otherwise the signal is ignored.

        For open positions (Stop Loss mode - config.stop_loss_mode == True):
        - Averaging is fully disabled.
        - Exit on whichever hits first: Stop Loss (fills at OPEN on a
          gap-down through the stop, else AT the stop price if the day's
          LOW touched it) or Profit Target (same open-gap / high-touch
          resolution as Averaging mode).
        """
        # Prepare all data with signals
        prepared_data: Dict[str, pd.DataFrame] = {}
        all_dates: Set[datetime] = set()

        # BUGFIX: the warmup requirement must be max(sma, lookback) (whichever
        # indicator needs the most history), NOT their sum. Using the sum
        # (176 bars) silently discarded valid symbols and disagreed with the
        # Scanner / UI, which use max(...)+10 (136 bars). Use the single shared
        # definition everywhere so the two engines never diverge again.
        min_required = self.config.min_required_bars

        # Diagnostics so a "zero trades" result is never silent - the UI can
        # tell the user WHY (no data downloaded / too little history / no
        # signals fired / no cash), instead of just showing 0.
        diag = {
            'symbols_received': len(data_dict),
            'symbols_empty': 0,
            'symbols_too_short': 0,
            'symbols_missing_cols': 0,
            'symbols_prepared': 0,
            'total_buy_signals': 0,
            'buy_orders_executed': 0,
            'average_orders_executed': 0,
            'bar_counts': [],
        }

        for symbol, df in data_dict.items():
            if df is None or df.empty:
                diag['symbols_empty'] += 1
                continue
            diag['bar_counts'].append(len(df))
            if len(df) < min_required:
                diag['symbols_too_short'] += 1
                continue

            df_signals = SignalGenerator.generate_signals(df, self.config)

            # Ensure date column exists and is datetime
            if 'date' not in df_signals.columns:
                diag['symbols_missing_cols'] += 1
                continue
            df_signals['date'] = pd.to_datetime(df_signals['date'])

            # BUGFIX: strip timezone. Fresh yfinance data is tz-aware
            # (Asia/Kolkata); if even one symbol is tz-naive (older cache, a
            # different source, an index), sorted(all_dates) raises
            # "Cannot compare tz-naive and tz-aware timestamps" and the ENTIRE
            # backtest aborts to zero. Normalising every symbol to tz-naive
            # makes the merged date axis always comparable.
            if getattr(df_signals['date'].dt, 'tz', None) is not None:
                df_signals['date'] = df_signals['date'].dt.tz_localize(None)

            # Drop duplicate dates to keep the index unique (one row per day)
            df_signals = df_signals.drop_duplicates(subset=['date'], keep='last')

            diag['total_buy_signals'] += int(df_signals['buy_signal'].sum())

            # Set date as index for O(1) lookup
            df_signals = df_signals.set_index('date')
            prepared_data[symbol] = df_signals
            all_dates.update(df_signals.index.tolist())
            diag['symbols_prepared'] += 1

        self.last_diagnostics = diag

        if not all_dates:
            return [], []

        sorted_dates = sorted(all_dates)
        portfolio = PortfolioState(cash=self.config.initial_capital)

        # Pending orders: {symbol: {'signal_date': date, 'signal_close': float}}
        # Per spec: no limit-price logic. A signal generated on Day T's close
        # ALWAYS fills on Day T+1 at that day's OPEN price (if the symbol has
        # data on T+1 and portfolio constraints - cash / max positions - allow it).
        pending_orders: Dict[str, Dict] = {}
        pending_average_orders: Dict[str, Dict] = {}  # For averaging down

        for current_date in sorted_dates:
            daily_prices: Dict[str, float] = {}

            # =====================================================================
            # STEP 1: Execute pending orders from previous day (T+1 execution)
            # FIX: unconditional buy at next day's OPEN price - no limit-order
            # simulation. This matches the spec exactly: "Buy at the OPEN PRICE"
            # on the next trading day, full stop.
            # =====================================================================

            # Execute pending NEW position orders. "T+1" means the stock's next
            # AVAILABLE session — if the stock is halted / has no bar today we
            # carry the order forward (bounded by STALE_ORDER_MAX_MISSES) rather
            # than silently discarding a valid buy signal.
            for symbol, order in list(pending_orders.items()):
                have_bar = (symbol in prepared_data
                            and current_date in prepared_data[symbol].index
                            and current_date > order['signal_date'])
                if have_bar:
                    row = prepared_data[symbol].loc[current_date]
                    execute_price = float(row['open'])

                    if not _is_valid_price(execute_price):
                        # Bad price data (NaN / non-positive): treat like a
                        # no-data day and carry the order forward rather than
                        # filling at a corrupt price.
                        order['misses'] = order.get('misses', 0) + 1
                        if order['misses'] > self.STALE_ORDER_MAX_MISSES:
                            del pending_orders[symbol]
                            self.last_diagnostics['orders_expired_no_data'] = \
                                self.last_diagnostics.get('orders_expired_no_data', 0) + 1
                        continue

                    filled = False
                    if portfolio.can_open_position(self.config):
                        if portfolio.open_position(
                            symbol, current_date, execute_price, self.config,
                            signal_date=order['signal_date'], signal_close=order['signal_close']
                        ):
                            filled = True
                            self.last_diagnostics['buy_orders_executed'] = \
                                self.last_diagnostics.get('buy_orders_executed', 0) + 1

                    if filled:
                        del pending_orders[symbol]
                    else:
                        # Capacity full or insufficient cash on the fill day —
                        # the buy signal is still valid, so carry it forward
                        # (bounded) instead of silently dropping it. A slot /
                        # cash freeing up within the window will fill it.
                        order['misses'] = order.get('misses', 0) + 1
                        if order['misses'] > self.STALE_ORDER_MAX_MISSES:
                            del pending_orders[symbol]
                            self.last_diagnostics['orders_expired_capacity'] = \
                                self.last_diagnostics.get('orders_expired_capacity', 0) + 1
                else:
                    # No bar for this symbol today (halt / missing data). Keep the
                    # order pending for the next session, up to the carry cap.
                    order['misses'] = order.get('misses', 0) + 1
                    if order['misses'] > self.STALE_ORDER_MAX_MISSES:
                        del pending_orders[symbol]
                        self.last_diagnostics['orders_expired_no_data'] = \
                            self.last_diagnostics.get('orders_expired_no_data', 0) + 1

            # Execute pending AVERAGING orders
            # FIX (averaging spec): a pending averaging order only ever gets
            # queued (see STEP 2 below) when Condition 1 - a genuinely NEW
            # buy_signal - has fired. This block enforces Condition 2: the
            # ACTUAL next-day OPEN execution price must be at least
            # average_trigger_pct% below the PREVIOUS EXECUTION price (not
            # the signal-day close). If the open price doesn't clear that
            # gate, the signal is ignored - no averaging happens today, and
            # the strategy simply waits for the next qualifying buy_signal.
            # As with new orders, a no-data (halt) day carries the order
            # forward instead of dropping it.
            for symbol, order in list(pending_average_orders.items()):
                # Position already exited, or averaging disabled -> order is moot.
                if self.config.stop_loss_mode or symbol not in portfolio.positions:
                    del pending_average_orders[symbol]
                    continue

                have_bar = (symbol in prepared_data
                            and current_date in prepared_data[symbol].index
                            and current_date > order['signal_date'])
                if not have_bar:
                    order['misses'] = order.get('misses', 0) + 1
                    if order['misses'] > self.STALE_ORDER_MAX_MISSES:
                        del pending_average_orders[symbol]
                        self.last_diagnostics['orders_expired_no_data'] = \
                            self.last_diagnostics.get('orders_expired_no_data', 0) + 1
                    continue

                row = prepared_data[symbol].loc[current_date]
                execute_price = float(row['open'])
                if not _is_valid_price(execute_price):
                    order['misses'] = order.get('misses', 0) + 1
                    if order['misses'] > self.STALE_ORDER_MAX_MISSES:
                        del pending_average_orders[symbol]
                        self.last_diagnostics['orders_expired_no_data'] = \
                            self.last_diagnostics.get('orders_expired_no_data', 0) + 1
                    continue
                pos = portfolio.positions[symbol]

                last_entry_price = pos.entry_prices[-1]
                price_gate_met = execute_price <= last_entry_price * (1 - self.config.average_trigger_pct / 100)

                # Only check cash, ignore max_positions for averaging
                position_size = self.config.position_size
                entry_cost = self.config.calculate_entry_cost(position_size)
                total_needed = position_size + entry_cost
                cash_ok = portfolio.cash >= total_needed
                count_ok = pos.average_count < self.config.max_average_entries

                if price_gate_met and cash_ok and count_ok:
                    if portfolio.add_to_position(
                        symbol, current_date, execute_price, self.config,
                        signal_date=order['signal_date'], signal_close=order['signal_close']
                    ):
                        self.last_diagnostics['average_orders_executed'] = \
                            self.last_diagnostics.get('average_orders_executed', 0) + 1
                # else: Condition 2 (or cash / max-entries) not met -
                # per spec, ignore this signal entirely.

                del pending_average_orders[symbol]

            # =====================================================================
            # STEP 2: Process today's data - check exits and generate new signals
            # =====================================================================

            for symbol, df in prepared_data.items():
                if current_date not in df.index:
                    continue

                row = df.loc[current_date]
                close = float(row['close'])
                # Skip a day with corrupt close data (NaN / non-positive): it
                # must not seed daily_prices (equity/P&L) or drive exits/signals.
                if not _is_valid_price(close):
                    continue
                buy_signal = bool(row['buy_signal'])
                daily_prices[symbol] = close

                has_position = symbol in portfolio.positions

                if has_position:
                    pos = portfolio.positions[symbol]

                    if self.config.stop_loss_mode:
                        # ------------------------------------------------
                        # STOP LOSS MODE (optional, user-controlled):
                        # averaging is fully disabled - every position has
                        # exactly one entry. Exit on whichever hits first:
                        #   - Stop Loss: day's OPEN gapped through it (fill
                        #     at open), else day's LOW touched it (fill AT
                        #     the stop price).
                        #   - Profit Target: day's OPEN gapped through it
                        #     (fill at open), else day's HIGH touched it
                        #     (fill AT the target price).
                        # BUGFIX: previously this checked CLOSE >= target
                        # and exited AT the close, which could silently
                        # fill far above the stated 3.14% target on a
                        # volatile day. A resting sell-limit/stop order
                        # can't fill better than the market actually
                        # offered - see _resolve_gap_or_touch_exit().
                        # Entry logic and the profit target/stop % values
                        # themselves are completely unchanged.
                        # ------------------------------------------------
                        stop_loss_price = pos.stop_loss_trigger_price(self.config)
                        exit_target = pos.exit_trigger_price(self.config)
                        open_price = float(row['open'])
                        high = float(row['high'])
                        low = float(row['low'])

                        sl_fill = _resolve_gap_or_touch_exit(open_price, low, stop_loss_price, 'down')
                        target_fill = _resolve_gap_or_touch_exit(open_price, high, exit_target, 'up')

                        if sl_fill is not None:
                            profit_pct, _ = pos.calculate_unrealized_pnl(sl_fill, self.config)
                            gap_note = " (Gap)" if open_price <= stop_loss_price else ""
                            reason = f"Stop Loss ₹{sl_fill:.2f}{gap_note} ({profit_pct:.2f}%)"
                            portfolio.close_position(symbol, current_date, sl_fill, self.config, reason)
                        elif target_fill is not None:
                            profit_pct, _ = pos.calculate_unrealized_pnl(target_fill, self.config)
                            gap_note = " (Gap)" if open_price >= exit_target else ""
                            reason = f"Profit Target ₹{target_fill:.2f}{gap_note} ({profit_pct:.2f}%)"
                            portfolio.close_position(symbol, current_date, target_fill, self.config, reason)
                        # else: hold. No averaging is ever attempted in Stop Loss mode.

                    else:
                        # ------------------------------------------------
                        # AVERAGING MODE (original strategy - unchanged
                        # profit-target VALUE (3.14%) and averaging rule
                        # eligibility, FIXED fill-price realism).
                        # ------------------------------------------------
                        # MODIFIED EXIT LOGIC (per spec): pure price-based target,
                        # not net-of-cost profit_pct.
                        #   entries == 1 -> Buy Price * 1.0314
                        #   entries  > 1 -> Weighted Average Buy Price * 1.0314
                        # Recalculated live from current entries on every check.
                        exit_target = pos.exit_trigger_price(self.config)
                        open_price = float(row['open'])
                        high = float(row['high'])

                        # BUGFIX: was `close >= exit_target` -> exit AT close,
                        # which let the reported exit price (and profit%)
                        # run far past the stated target whenever the day's
                        # close simply happened to be higher. Now resolved
                        # realistically off OPEN (gap fill) / HIGH (intraday
                        # touch, fills exactly at the target) - see
                        # _resolve_gap_or_touch_exit().
                        target_fill = _resolve_gap_or_touch_exit(open_price, high, exit_target, 'up')

                        if target_fill is not None:
                            profit_pct, _ = pos.calculate_unrealized_pnl(target_fill, self.config)
                            gap_note = " (Gap)" if open_price >= exit_target else ""
                            if pos.average_count > 1:
                                reason = (f"Averaged Exit - WAP Target ₹{target_fill:.2f}{gap_note} "
                                          f"({profit_pct:.2f}%)")
                            else:
                                reason = f"Profit Target ₹{target_fill:.2f}{gap_note} ({profit_pct:.2f}%)"
                            portfolio.close_position(symbol, current_date, target_fill, self.config, reason)
                        else:
                            # Averaging requires a genuinely NEW, complete buy
                            # signal (Condition 1) — i.e. today's buy_signal must
                            # be True again. Under the RSI Tracking + 200 EMA
                            # state machine this can only happen after a FRESH
                            # cycle: SignalGenerator resets tracking to False the
                            # instant any buy fires (including the original
                            # entry), so a later buy_signal=True bar necessarily
                            # means the stock re-armed (RSI < threshold while
                            # Close < EMA200) and then reclaimed (Close > EMA200
                            # and EMA50 > EMA21) all over again — exactly one buy
                            # per tracking cycle, never a re-used state.
                            # Condition 2 (execution price >= average_trigger_pct%
                            # below the PREVIOUS EXECUTION price) is enforced at
                            # T+1 fill time in STEP 1 above, using the actual
                            # open price rather than the signal-day close.
                            if buy_signal and pos.average_count < self.config.max_average_entries:
                                if symbol not in pending_average_orders:
                                    pending_average_orders[symbol] = {
                                        'signal_date': current_date,
                                        'signal_close': close
                                    }
                else:
                    # No position - check for new buy signal
                    if buy_signal and portfolio.can_open_position(self.config):
                        # Signal fires on Day T's close; execution deferred to
                        # Day T+1's open (see STEP 1 above). Never buy on the
                        # signal candle itself.
                        if symbol not in pending_orders:
                            pending_orders[symbol] = {
                                'signal_date': current_date,
                                'signal_close': close
                            }

            # Record daily equity
            portfolio.record_equity(current_date, daily_prices, self.config)

        # Positions still open when the data ends are NEVER force-closed:
        # no synthetic sell, no "End of Data" exit, not counted in
        # win/loss stats or realized P&L. Instead, build a read-only
        # snapshot of each open position for the separate "Open Positions"
        # report. The equity curve above already marks these positions to
        # market via portfolio.record_equity() on every bar, so total
        # portfolio value stays accurate without turning them into trades.
        open_positions_snapshot: List[Dict[str, Any]] = []
        for symbol, pos in portfolio.positions.items():
            if symbol not in prepared_data or len(prepared_data[symbol]) == 0:
                continue
            df = prepared_data[symbol]
            last_row = df.iloc[-1]
            current_price = float(last_row['close'])
            last_date = df.index[-1]
            if isinstance(last_date, pd.Timestamp):
                last_date = last_date.to_pydatetime()

            profit_pct, gross_pnl = pos.calculate_unrealized_pnl(current_price, self.config)
            current_value = pos.total_quantity * current_price
            peak_value = portfolio.position_peak_value.get(symbol, current_value)
            drawdown_pct = ((peak_value - current_value) / peak_value * 100) if peak_value > 0 else 0.0
            entry_date0 = pos.entry_dates[0] if pos.entry_dates else last_date
            days_held = (last_date - entry_date0).days if pos.entry_dates else 0

            open_positions_snapshot.append({
                'symbol': symbol,
                'entry_date': entry_date0,
                'entry_price': pos.entry_prices[0] if pos.entry_prices else 0.0,
                'average_entry_price': pos.average_entry_price,
                'num_entries': pos.average_count,
                'current_price': current_price,
                'quantity': pos.total_quantity,
                'invested_amount': pos.total_invested,
                'current_market_value': current_value,
                'unrealized_pnl_inr': gross_pnl,
                'unrealized_pnl_pct': profit_pct,
                'current_drawdown_pct': drawdown_pct,
                'profit_target_price': pos.exit_trigger_price(self.config),
                'stop_loss_price': pos.stop_loss_trigger_price(self.config) if self.config.stop_loss_mode else None,
                'days_held': days_held,
                'status': 'Open',
            })

        self.last_open_positions = open_positions_snapshot

        # Additive (does not change the return signature/behavior): expose
        # the daily invested-capital series for the new "Capital
        # Utilization" report metric.
        self.last_daily_invested = portfolio.daily_invested

        return portfolio.completed_trades, portfolio.daily_equity
# =============================================================================
# METRICS CALCULATOR
# =============================================================================

def calculate_xirr(cashflows: List[Tuple[Any, float]]) -> float:
    """
    Portfolio-level XIRR (annualized, irregular-cashflow IRR) as a percentage.

    cashflows: list of (date, amount) - negative amounts for money invested
    (buys), positive for money returned (sells). Pure-Python bisection root
    solve on NPV(rate) - no scipy dependency, since the rest of the app has
    none either.
    """
    if not cashflows or len(cashflows) < 2:
        return 0.0

    dates = [d for d, _ in cashflows]
    t0 = min(dates)

    # A sign change in the cash flows is necessary for a real IRR to exist.
    # All-positive or all-negative series have no root at all -> report 0.
    amounts = [a for _, a in cashflows]
    if all(a >= 0 for a in amounts) or all(a <= 0 for a in amounts):
        return 0.0

    def npv(rate: float) -> float:
        total = 0.0
        for d, amt in cashflows:
            days = (d - t0).days
            total += amt / ((1 + rate) ** (days / 365.25))
        return total

    lo = -0.999
    f_lo = npv(lo)
    if f_lo == 0:
        return lo * 100

    # Bracket the root on the high side. A fixed 1000% ceiling silently failed
    # for very high annualized returns (common on ultra-short holds), where
    # npv(hi) keeps f_lo's sign and the old bracket check wrongly reported 0.
    # Expand hi geometrically until the sign flips or a large cap; if the root
    # is beyond the cap (astronomically high return), report the capped bound
    # rather than 0 so the value is never silently lost.
    hi = 1.0
    f_hi = npv(hi)
    expansions = 0
    while f_lo * f_hi > 0 and hi < 1e12 and expansions < 80:
        hi *= 2.0
        f_hi = npv(hi)
        expansions += 1
    if f_hi == 0:
        return hi * 100
    if f_lo * f_hi > 0:
        # Sign change exists (checked above) but lies beyond hi -> IRR exceeds
        # the cap. Return the capped bound instead of 0.
        return hi * 100

    mid = 0.0
    for _ in range(200):
        mid = (lo + hi) / 2
        f_mid = npv(mid)
        if abs(f_mid) < 1e-6:
            break
        if f_lo * f_mid < 0:
            hi = mid
        else:
            lo = mid
            f_lo = f_mid
    return mid * 100


class MetricsCalculator:
    """Calculate comprehensive portfolio-level metrics."""

    @staticmethod
    def calculate_metrics(
        completed_trades: List[CompletedTrade],
        daily_equity: List[Tuple[datetime, float]],
        initial_capital: float,
        daily_invested: Optional[List[Tuple[datetime, float]]] = None,
    ) -> Dict[str, Any]:
        metrics = {
            'total_trades': len(completed_trades),
            'winning_trades': 0,
            'losing_trades': 0,
            'win_rate': 0.0,
            'avg_win': 0.0,
            'avg_loss': 0.0,
            'avg_loss_pct': 0.0,
            'profit_factor': 0.0,
            'net_profit': 0.0,
            'total_return_pct': 0.0,
            'max_drawdown_pct': 0.0,
            'sharpe_ratio': 0.0,
            'sortino_ratio': 0.0,
            'calmar_ratio': 0.0,
            'cagr': 0.0,
            'xirr_pct': 0.0,
            'expectancy': 0.0,
            'recovery_factor': 0.0,
            'avg_holding_days': 0.0,
            'largest_winner': 0.0,
            'largest_loser': 0.0,
            'stop_loss_hits': 0,
            'target_hits': 0,
            'stop_loss_hit_rate': 0.0,
            'capital_utilization_pct': 0.0,
            'monthly_returns': {},
            'yearly_returns': {},
            'equity_curve': daily_equity,
        }

        if not completed_trades:
            return metrics

        # Trade-level metrics
        wins = [t.profit_inr for t in completed_trades if t.profit_inr > 0]
        losses = [t.profit_inr for t in completed_trades if t.profit_inr <= 0]

        metrics['winning_trades'] = len(wins)
        metrics['losing_trades'] = len(losses)
        metrics['win_rate'] = (len(wins) / len(completed_trades) * 100) if completed_trades else 0
        metrics['avg_win'] = np.mean(wins) if wins else 0
        # FIX v3.1: Store avg_loss as absolute magnitude for cleaner calculations
        metrics['avg_loss'] = abs(np.mean(losses)) if losses else 0
        metrics['largest_winner'] = max(wins) if wins else 0
        metrics['largest_loser'] = min(losses) if losses else 0

        total_wins = sum(wins) if wins else 0
        total_losses = abs(sum(losses)) if losses else 0
        # Cap at a large sentinel instead of float('inf') when there are no
        # losing trades: 'inf' renders poorly in the UI and breaks Excel export.
        if total_losses > 0:
            metrics['profit_factor'] = total_wins / total_losses
        else:
            metrics['profit_factor'] = 999.99 if total_wins > 0 else 0.0

        metrics['net_profit'] = sum(t.profit_inr for t in completed_trades)
        metrics['total_return_pct'] = (metrics['net_profit'] / initial_capital) * 100
        metrics['avg_holding_days'] = np.mean([t.holding_days for t in completed_trades])

        # NEW: Stop Loss vs Target reporting (Optional Stop Loss module).
        # Classified from the exit_reason text set by BacktestEngine - a
        # trade closed on Stop Loss always starts with "Stop Loss", a
        # target-style exit always starts with "Profit Target" or
        # "Averaged Exit" (WAP target hit). "End of Data" trades count as
        # neither. Harmless / all-zero when Stop Loss mode wasn't used.
        loss_pcts = [t.profit_pct for t in completed_trades if t.profit_inr <= 0]
        metrics['avg_loss_pct'] = abs(np.mean(loss_pcts)) if loss_pcts else 0.0

        stop_loss_hits = sum(1 for t in completed_trades if (t.exit_reason or '').startswith('Stop Loss'))
        target_hits = sum(
            1 for t in completed_trades
            if (t.exit_reason or '').startswith('Profit Target') or (t.exit_reason or '').startswith('Averaged Exit')
        )
        metrics['stop_loss_hits'] = stop_loss_hits
        metrics['target_hits'] = target_hits
        metrics['stop_loss_hit_rate'] = (stop_loss_hits / len(completed_trades) * 100) if completed_trades else 0.0

        # NEW: Portfolio-level XIRR from actual entry/exit cash flows. Every
        # entry (initial or averaging) invests the same ~position_size, so
        # total_invested / average_count recovers each entry's exact cash
        # outflow without needing extra state on CompletedTrade.
        cashflows: List[Tuple[Any, float]] = []
        for t in completed_trades:
            per_entry = t.total_invested / t.average_count if t.average_count else t.total_invested
            for d in t.entry_dates:
                cashflows.append((d, -per_entry))
            if t.exit_date:
                cashflows.append((t.exit_date, t.total_returned))
        metrics['xirr_pct'] = calculate_xirr(cashflows)

        # NEW: Capital Utilization - average capital actually deployed in
        # open positions over the run, as a % of initial capital.
        if daily_invested:
            invested_values = [v for _, v in daily_invested]
            if invested_values:
                metrics['capital_utilization_pct'] = (np.mean(invested_values) / initial_capital) * 100

        # Daily equity-based metrics
        if daily_equity and len(daily_equity) > 1:
            equity_values = [e for _, e in daily_equity]
            dates = [d for d, _ in daily_equity]

            # Max drawdown using daily equity
            peak = equity_values[0]
            max_dd = 0
            for eq in equity_values:
                if eq > peak:
                    peak = eq
                dd = (peak - eq) / peak * 100 if peak > 0 else 0
                if dd > max_dd:
                    max_dd = dd
            metrics['max_drawdown_pct'] = max_dd

            # Daily returns
            daily_returns = []
            for i in range(1, len(equity_values)):
                if equity_values[i-1] > 0:
                    ret = (equity_values[i] - equity_values[i-1]) / equity_values[i-1]
                    daily_returns.append(ret)

            if daily_returns:
                returns_arr = np.array(daily_returns)
                avg_return = np.mean(returns_arr)
                # Sample standard deviation (ddof=1) is the convention for
                # financial return series; population std slightly overstates
                # the Sharpe/Sortino ratios. Needs >=2 points to be defined.
                std_return = np.std(returns_arr, ddof=1) if len(returns_arr) > 1 else 0.0

                if std_return > 0:
                    metrics['sharpe_ratio'] = (avg_return / std_return) * np.sqrt(252)

                downside_returns = returns_arr[returns_arr < 0]
                if len(downside_returns) > 1:
                    downside_std = np.std(downside_returns, ddof=1)
                    if downside_std > 0:
                        metrics['sortino_ratio'] = (avg_return / downside_std) * np.sqrt(252)

                # FIX v3.1: Calmar Ratio = CAGR / Max Drawdown
                # Calculate CAGR properly first
                start_date = dates[0]
                end_date = dates[-1]
                years = max((end_date - start_date).days / 365.25, 0.01)
                final_equity = equity_values[-1]
                metrics['cagr'] = ((final_equity / initial_capital) ** (1/years) - 1) * 100

                # Calmar = CAGR / Max Drawdown (both as percentages)
                if metrics['max_drawdown_pct'] > 0:
                    metrics['calmar_ratio'] = metrics['cagr'] / metrics['max_drawdown_pct']


        if completed_trades:
            win_prob = metrics['winning_trades'] / len(completed_trades)
            loss_prob = metrics['losing_trades'] / len(completed_trades)
            metrics['expectancy'] = (win_prob * metrics['avg_win']) - (loss_prob * metrics['avg_loss'])

        if metrics['max_drawdown_pct'] > 0:
            metrics['recovery_factor'] = metrics['total_return_pct'] / metrics['max_drawdown_pct']

        monthly_rets = {}
        yearly_rets = {}
        for trade in completed_trades:
            if trade.exit_date:
                month_key = trade.exit_date.strftime('%Y-%m')
                year_key = trade.exit_date.strftime('%Y')
                monthly_rets[month_key] = monthly_rets.get(month_key, 0) + trade.profit_inr
                yearly_rets[year_key] = yearly_rets.get(year_key, 0) + trade.profit_inr

        metrics['monthly_returns'] = monthly_rets
        metrics['yearly_returns'] = yearly_rets

        return metrics

    @staticmethod
    def metrics_to_dataframe(metrics: Dict[str, Any]) -> pd.DataFrame:
        # FIX v3.1: All values cast to str for PyArrow/Streamlit compatibility
        data = {
            'Metric': [
                'Total Trades', 'Winning Trades', 'Losing Trades', 'Win Rate (%)',
                'Average Win (₹)', 'Average Loss (₹)', 'Average Loss (%)', 'Profit Factor',
                'Net Profit (₹)', 'Total Return (%)', 'Max Drawdown (%)',
                'Sharpe Ratio', 'Sortino Ratio', 'Calmar Ratio', 'CAGR (%)', 'XIRR (%)',
                'Expectancy (₹)', 'Recovery Factor', 'Avg Holding Days',
                'Capital Utilization (%)',
                'Target Hits', 'Stop Loss Hits', 'Stop Loss Hit Rate (%)',
                'Largest Winner (₹)', 'Largest Loser (₹)'
            ],
            'Value': [
                str(metrics['total_trades']),
                str(metrics['winning_trades']),
                str(metrics['losing_trades']),
                f"{metrics['win_rate']:.2f}",
                f"{metrics['avg_win']:,.2f}",
                f"{metrics['avg_loss']:,.2f}",
                f"{metrics.get('avg_loss_pct', 0):.2f}",
                f"{metrics['profit_factor']:.2f}",
                f"{metrics['net_profit']:,.2f}",
                f"{metrics['total_return_pct']:.2f}",
                f"{metrics['max_drawdown_pct']:.2f}",
                f"{metrics['sharpe_ratio']:.2f}",
                f"{metrics['sortino_ratio']:.2f}",
                f"{metrics['calmar_ratio']:.2f}",
                f"{metrics['cagr']:.2f}",
                f"{metrics.get('xirr_pct', 0):.2f}",
                f"{metrics['expectancy']:,.2f}",
                f"{metrics['recovery_factor']:.2f}",
                f"{metrics['avg_holding_days']:.1f}",
                f"{metrics.get('capital_utilization_pct', 0):.2f}",
                str(metrics.get('target_hits', 0)),
                str(metrics.get('stop_loss_hits', 0)),
                f"{metrics.get('stop_loss_hit_rate', 0):.2f}",
                f"{metrics['largest_winner']:,.2f}",
                f"{metrics['largest_loser']:,.2f}"
            ]
        }
        return pd.DataFrame(data)

    # -----------------------------------------------------------------
    # NEW: Unrealized Performance (open positions only). Kept entirely
    # separate from calculate_metrics()/metrics_to_dataframe() above,
    # which remain purely closed-trade ("Realized Performance") stats -
    # Win Rate, Avg Win/Loss, Profit Factor, etc. are never touched by
    # open positions. This aggregates BacktestEngine.last_open_positions.
    # -----------------------------------------------------------------
    @staticmethod
    def summarize_open_positions(open_positions: List[Dict[str, Any]]) -> Dict[str, float]:
        summary = {
            'open_position_count': len(open_positions),
            'total_unrealized_pnl_inr': 0.0,
            'total_unrealized_pnl_pct': 0.0,
            'current_market_value': 0.0,
            'invested_capital': 0.0,
            'unrealized_return_pct': 0.0,
        }
        if not open_positions:
            return summary

        total_invested = sum(p.get('invested_amount', 0.0) for p in open_positions)
        total_market_value = sum(p.get('current_market_value', 0.0) for p in open_positions)
        total_pnl = sum(p.get('unrealized_pnl_inr', 0.0) for p in open_positions)

        summary['total_unrealized_pnl_inr'] = total_pnl
        summary['current_market_value'] = total_market_value
        summary['invested_capital'] = total_invested
        summary['unrealized_return_pct'] = (total_pnl / total_invested * 100) if total_invested > 0 else 0.0
        # Aggregate % is invested-capital-weighted, not a simple mean of
        # each position's own %, so it's consistent with the ₹ P&L above.
        summary['total_unrealized_pnl_pct'] = summary['unrealized_return_pct']

        return summary

    @staticmethod
    def open_positions_to_dataframe(open_positions: List[Dict[str, Any]]) -> pd.DataFrame:
        """Display-ready DataFrame for the 'Open Positions' report table."""
        if not open_positions:
            return pd.DataFrame(columns=[
                'Symbol', 'Entry Date', 'Entry Price', 'Avg Entry Price', 'Entries',
                'Current Price', 'Quantity', 'Invested (₹)', 'Market Value (₹)',
                'Unrealized P&L (₹)', 'Unrealized P&L (%)', 'Drawdown (%)',
                'Target Price', 'Stop Loss Price', 'Days Held', 'Status'
            ])

        rows = []
        for p in open_positions:
            entry_date = p.get('entry_date')
            entry_date_str = entry_date.strftime('%Y-%m-%d') if hasattr(entry_date, 'strftime') else str(entry_date)
            sl_price = p.get('stop_loss_price')
            rows.append({
                'Symbol': p.get('symbol', ''),
                'Entry Date': entry_date_str,
                'Entry Price': f"{p.get('entry_price', 0):.2f}",
                'Avg Entry Price': f"{p.get('average_entry_price', 0):.2f}",
                'Entries': p.get('num_entries', 0),
                'Current Price': f"{p.get('current_price', 0):.2f}",
                'Quantity': f"{p.get('quantity', 0):.2f}",
                'Invested (₹)': f"{p.get('invested_amount', 0):,.2f}",
                'Market Value (₹)': f"{p.get('current_market_value', 0):,.2f}",
                'Unrealized P&L (₹)': f"{p.get('unrealized_pnl_inr', 0):,.2f}",
                'Unrealized P&L (%)': f"{p.get('unrealized_pnl_pct', 0):.2f}",
                'Drawdown (%)': f"{p.get('current_drawdown_pct', 0):.2f}",
                'Target Price': f"{p.get('profit_target_price', 0):.2f}",
                'Stop Loss Price': f"{sl_price:.2f}" if sl_price is not None else "—",
                'Days Held': p.get('days_held', 0),
                'Status': p.get('status', 'Open'),
            })
        return pd.DataFrame(rows)

# =============================================================================
# SCANNER (Multithreaded, full universe)
# =============================================================================

class Scanner:
    """Real-time scanner using shared SignalGenerator for consistency."""

    def __init__(self, config: StrategyConfig):
        self.config = config

    def scan_stock(self, symbol: str, df: pd.DataFrame) -> Optional[Dict]:
        result = SignalGenerator.get_latest_signal(df, self.config)
        if result is None:
            return None
        result['symbol'] = symbol
        return result

    def scan_universe(
        self,
        symbols: List[str],
        data_manager: DataManager,
        max_workers: int = 8,
        progress_callback=None,
        status_callback=None,
    ) -> List[Dict]:
        """
        Scan complete universe.

        BUGFIX (rate limiting): this used to fire one concurrent
        download_stock_data() call per symbol via ThreadPoolExecutor - the
        exact pattern that triggers Yahoo's rate limit (see batch_download's
        docstring for the same bug on the backtest side). Now it prefetches
        everything through DataManager.batch_download(), which already
        handles chunking + adaptive backoff, THEN runs signal generation
        locally (no network) against the results - no network calls happen
        during the scan itself.
        """
        results = []

        data_dict = data_manager.batch_download(
            symbols,
            period=self.config.data_period,
            max_workers=max_workers,
            progress_callback=(lambda p: progress_callback(p * 0.9)) if progress_callback else None,
            status_callback=status_callback,
        )

        total = max(len(symbols), 1)
        for i, symbol in enumerate(symbols):
            df = data_dict.get(symbol)
            if df is not None:
                try:
                    result = self.scan_stock(symbol, df)
                    if result and result['buy_signal']:
                        meta = data_manager.get_stock_metadata(symbol)
                        result['company'] = meta.get('name', symbol)
                        result['sector'] = meta.get('sector', 'Unknown')
                        results.append(result)
                except Exception as e:
                    logger.error(f"Scanner error for {symbol}: {e}")
            if progress_callback:
                progress_callback(0.9 + 0.1 * ((i + 1) / total))

        return results

# =============================================================================
# VISUALIZATION
# =============================================================================

class ChartBuilder:
    """Build interactive Plotly charts with unique keys."""

    @staticmethod
    def create_trade_chart(df: pd.DataFrame, trades: List[CompletedTrade], symbol: str, config: StrategyConfig) -> go.Figure:
        df = SignalGenerator.generate_signals(df, config)

        fig = make_subplots(
            rows=3, cols=1,
            shared_xaxes=True,
            vertical_spacing=0.05,
            row_heights=[0.6, 0.2, 0.2],
            subplot_titles=(f'{symbol} - Price & Signals', 'RSI', 'Volume')
        )

        x_vals = df['date'] if 'date' in df.columns else df.index

        fig.add_trace(
            go.Candlestick(
                x=x_vals, open=df['open'], high=df['high'],
                low=df['low'], close=df['close'], name='Price'
            ), row=1, col=1
        )

        fig.add_trace(
            go.Scatter(x=x_vals, y=df['ema_long'], name=f'EMA {config.ema_long}',
                      line=dict(color='#8B5CF6', width=1.8)), row=1, col=1
        )
        fig.add_trace(
            go.Scatter(x=x_vals, y=df['ema_mid'], name=f'EMA {config.ema_mid}',
                      line=dict(color='#F59E0B', width=1.3)), row=1, col=1
        )
        fig.add_trace(
            go.Scatter(x=x_vals, y=df['ema_short'], name=f'EMA {config.ema_short}',
                      line=dict(color='#22D3EE', width=1.1)), row=1, col=1
        )

        # "Tracking Started" markers — the RSI<thr & Close<EMA200 arming
        # candles (arming only; these never buy).
        if 'tracking_started' in df.columns:
            ts_mask = df['tracking_started'].fillna(False).astype(bool)
            if ts_mask.any():
                ts_x = (df['date'] if 'date' in df.columns else pd.Series(df.index))[ts_mask.values]
                ts_y = df['close'][ts_mask.values]
                fig.add_trace(
                    go.Scatter(x=ts_x, y=ts_y, mode='markers',
                              marker=dict(color='#FACC15', size=11, symbol='x-thin',
                                          line=dict(width=2, color='#FACC15')),
                              name='Tracking Started'), row=1, col=1
                )

        for trade in trades:
            for i, (date, price) in enumerate(zip(trade.entry_dates, trade.entry_prices)):
                color = '#10B981' if i == 0 else '#3B82F6'
                symbol_marker = 'triangle-up' if i == 0 else 'diamond'
                fig.add_trace(
                    go.Scatter(x=[date], y=[price], mode='markers',
                              marker=dict(color=color, size=12, symbol=symbol_marker),
                              name=f'Buy #{i+1} (Execution)', showlegend=False), row=1, col=1
                )
            # FIX: also mark the signal candle (Day T close) separately from the
            # execution point (Day T+1 open), so the T+1 lag is visible on the chart
            for i, (sig_date, sig_close) in enumerate(zip(trade.signal_dates, trade.signal_closes)):
                fig.add_trace(
                    go.Scatter(x=[sig_date], y=[sig_close], mode='markers',
                              marker=dict(color='#F59E0B', size=8, symbol='circle-open', line=dict(width=2)),
                              name=f'Signal #{i+1}', showlegend=False), row=1, col=1
                )

            if trade.exit_date and trade.exit_price:
                fig.add_trace(
                    go.Scatter(x=[trade.exit_date], y=[trade.exit_price], mode='markers',
                              marker=dict(color='#EF4444', size=12, symbol='triangle-down'),
                              name='Sell', showlegend=False), row=1, col=1
                )

        fig.add_trace(
            go.Scatter(x=x_vals, y=df['rsi'], name='RSI',
                      line=dict(color='#3B82F6', width=1)), row=2, col=1
        )

        fig.add_hline(y=config.rsi_threshold, line_dash="dash", line_color="#EF4444", row=2, col=1)

        if 'volume' in df.columns:
            fig.add_trace(
                go.Bar(x=x_vals, y=df['volume'], name='Volume', marker_color='#475569'),
                row=3, col=1
            )

        # Cosmetic-only theming (colors/background/grid/fonts) — no trace
        # data or signal logic is touched above.
        apply_plotly_theme(fig, height=800)
        fig.update_layout(
            title=f'{symbol} - Strategy Analysis',
            yaxis_title='Price (₹)', xaxis_title='Date',
            showlegend=True,
            legend=dict(yanchor="top", y=0.99, xanchor="left", x=0.01)
        )
        fig.update_xaxes(rangeslider_visible=False)

        return fig

# =============================================================================
# EXPORT
# =============================================================================

class Exporter:
    """Handle data export functionality."""

    @staticmethod
    def trades_to_csv(trades: List[CompletedTrade]) -> str:
        if not trades:
            return ""
        data = []
        for t in trades:
            data.append({
                'Symbol': t.symbol,
                'Signal_Date': t.signal_dates[0].strftime('%Y-%m-%d') if t.signal_dates else '',
                'Signal_Close': t.signal_closes[0] if t.signal_closes else 0,
                'Execution_Date': t.entry_dates[0].strftime('%Y-%m-%d') if t.entry_dates else '',
                'Execution_Open_Price': t.entry_prices[0] if t.entry_prices else 0,
                'Average_Count': t.average_count,
                'Exit_Date': t.exit_date.strftime('%Y-%m-%d') if t.exit_date else '',
                'Exit_Price': t.exit_price if t.exit_price else 0,
                'Holding_Days': t.holding_days,
                'Profit_Pct': f"{t.profit_pct:.2f}",
                'Profit_INR': f"{t.profit_inr:.2f}",
                'Total_Invested': f"{t.total_invested:.2f}",
                'Total_Returned': f"{t.total_returned:.2f}",
                'Status': t.status,
                'Exit_Reason': t.exit_reason
            })
        return pd.DataFrame(data).to_csv(index=False)

    @staticmethod
    def scanner_to_csv(results: List[Dict]) -> str:
        if not results:
            return ""
        return pd.DataFrame(results).to_csv(index=False)

    @staticmethod
    def create_excel_report(
        metrics: Dict[str, Any],
        trades: List[CompletedTrade],
        scanner_results: Optional[List[Dict]] = None
    ) -> bytes:
        if not OPENPYXL_AVAILABLE:
            return b""

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            MetricsCalculator.metrics_to_dataframe(metrics).to_excel(writer, sheet_name='Metrics', index=False)

            if trades:
                trades_data = []
                for t in trades:
                    trades_data.append({
                        'Symbol': t.symbol,
                        'Signal_Date': t.signal_dates[0].strftime('%Y-%m-%d') if t.signal_dates else '',
                        'Signal_Close': t.signal_closes[0] if t.signal_closes else 0,
                        'Execution_Date': t.entry_dates[0].strftime('%Y-%m-%d') if t.entry_dates else '',
                        'Execution_Open_Price': t.entry_prices[0] if t.entry_prices else 0,
                        'Average_Count': t.average_count,
                        'Exit_Date': t.exit_date.strftime('%Y-%m-%d') if t.exit_date else '',
                        'Exit_Price': t.exit_price if t.exit_price else 0,
                        'Holding_Days': t.holding_days,
                        'Profit_%': t.profit_pct,
                        'Profit_INR': t.profit_inr,
                        'Total_Invested': t.total_invested,
                        'Total_Returned': t.total_returned,
                        'Status': t.status,
                        'Exit_Reason': t.exit_reason
                    })
                pd.DataFrame(trades_data).to_excel(writer, sheet_name='Trades', index=False)

            if scanner_results:
                pd.DataFrame(scanner_results).to_excel(writer, sheet_name='Scanner', index=False)

            if metrics.get('monthly_returns'):
                pd.DataFrame([
                    {'Month': k, 'Profit_INR': v}
                    for k, v in sorted(metrics['monthly_returns'].items())
                ]).to_excel(writer, sheet_name='Monthly Returns', index=False)

            if metrics.get('yearly_returns'):
                pd.DataFrame([
                    {'Year': k, 'Profit_INR': v}
                    for k, v in sorted(metrics['yearly_returns'].items())
                ]).to_excel(writer, sheet_name='Yearly Returns', index=False)

        return output.getvalue()

# =============================================================================
# PORTFOLIO MODULE (Investment Tracking) — separate feature, own storage
# =============================================================================
# This module tracks REAL positions the user has actually decided to take
# (added from a Scanner buy signal, or manually). It is fully independent of
# SignalGenerator / Scanner / BacktestEngine above: nothing in this section
# reads from, or writes into, those classes or their state. It has its own
# SQLite-backed persistence so portfolio data survives an app restart.

PORTFOLIO_DB_PATH = CACHE_DIR / "portfolio.db"
DEFAULT_AVERAGE_GAP_PCT = 2.0
DEFAULT_AVERAGE_LEVELS = 5
DEFAULT_CAPITAL_BASE = 1_000_000.0


def _pf_rerun() -> None:
    fn = getattr(st, "rerun", None) or getattr(st, "experimental_rerun", None)
    if fn:
        fn()
    else:
        logger.warning(
            "No st.rerun / st.experimental_rerun available (Streamlit too old); "
            "the page won't auto-refresh — the user may need to interact to update."
        )


# ------------------------------- Storage ------------------------------- #

def get_portfolio_conn() -> sqlite3.Connection:
    # timeout: wait (rather than immediately erroring) if another connection
    # holds the write lock. WAL mode lets readers proceed during a write and
    # substantially reduces 'database is locked' errors in multi-session
    # (Streamlit / container) deployments. PRAGMAs are per-connection but
    # WAL is a persistent database-level setting once set.
    conn = sqlite3.connect(str(PORTFOLIO_DB_PATH), check_same_thread=False, timeout=30.0)
    conn.row_factory = sqlite3.Row
    try:
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA busy_timeout=30000")
        conn.execute("PRAGMA synchronous=NORMAL")
    except sqlite3.Error:
        pass
    return conn


def init_portfolio_db() -> None:
    conn = get_portfolio_conn()
    try:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS positions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                symbol TEXT NOT NULL,
                company TEXT,
                buy_date TEXT NOT NULL,
                strategy TEXT,
                notes TEXT,
                stop_loss REAL,
                target_price REAL,
                average_gap_pct REAL DEFAULT 2.0,
                avg1_price REAL, avg2_price REAL, avg3_price REAL,
                avg4_price REAL, avg5_price REAL,
                entries_json TEXT NOT NULL DEFAULT '[]',
                exits_json TEXT NOT NULL DEFAULT '[]',
                status TEXT NOT NULL DEFAULT 'OPEN',
                close_date TEXT,
                close_price REAL,
                created_at TEXT,
                updated_at TEXT
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS portfolio_meta (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)
        conn.commit()
    finally:
        conn.close()


def pf_get_meta(key: str, default: Any = None) -> Any:
    conn = get_portfolio_conn()
    try:
        row = conn.execute("SELECT value FROM portfolio_meta WHERE key = ?", (key,)).fetchone()
        if row is None:
            return default
        try:
            return json.loads(row['value'])
        except Exception:
            return row['value']
    finally:
        conn.close()


def pf_set_meta(key: str, value: Any) -> None:
    conn = get_portfolio_conn()
    try:
        conn.execute(
            "INSERT INTO portfolio_meta (key, value) VALUES (?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            (key, json.dumps(value))
        )
        conn.commit()
    finally:
        conn.close()


def pf_fetch_positions(status: Optional[str] = None) -> List[Dict[str, Any]]:
    conn = get_portfolio_conn()
    try:
        if status:
            rows = conn.execute(
                "SELECT * FROM positions WHERE status = ? ORDER BY id DESC", (status,)
            ).fetchall()
        else:
            rows = conn.execute("SELECT * FROM positions ORDER BY id DESC").fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def pf_fetch_position(position_id: int) -> Optional[Dict[str, Any]]:
    conn = get_portfolio_conn()
    try:
        row = conn.execute("SELECT * FROM positions WHERE id = ?", (position_id,)).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def pf_symbol_in_portfolio(symbol: str) -> bool:
    conn = get_portfolio_conn()
    try:
        row = conn.execute(
            "SELECT id FROM positions WHERE symbol = ? AND status = 'OPEN' LIMIT 1", (symbol,)
        ).fetchone()
        return row is not None
    finally:
        conn.close()


# Columns pf_update_position is permitted to write. `id` (WHERE key) and
# `created_at` (set once at insert) are intentionally excluded. Any key outside
# this set is rejected before it can reach the SQL string.
PF_UPDATABLE_COLUMNS = frozenset({
    'symbol', 'company', 'buy_date', 'strategy', 'notes', 'stop_loss',
    'target_price', 'average_gap_pct', 'avg1_price', 'avg2_price', 'avg3_price',
    'avg4_price', 'avg5_price', 'entries_json', 'exits_json', 'status',
    'close_date', 'close_price', 'updated_at',
})


def pf_update_position(position_id: int, **fields) -> None:
    if not fields:
        return
    fields['updated_at'] = datetime.now().isoformat()

    # SECURITY: column names are interpolated into the SQL string (values are
    # always parameterized), so every key MUST be validated against a fixed
    # allow-list. This stays safe even if a caller ever forwards a **kwargs
    # dict built from user input. `id` is intentionally excluded (it's the
    # WHERE key, never an update target).
    invalid = set(fields) - PF_UPDATABLE_COLUMNS
    if invalid:
        raise ValueError(f"pf_update_position: refusing unknown/unsafe column(s): {sorted(invalid)}")

    cols = ", ".join(f"{k} = ?" for k in fields)
    values = list(fields.values()) + [position_id]
    conn = get_portfolio_conn()
    try:
        conn.execute(f"UPDATE positions SET {cols} WHERE id = ?", values)
        conn.commit()
    finally:
        conn.close()


def pf_delete_position(position_id: int) -> None:
    conn = get_portfolio_conn()
    try:
        conn.execute("DELETE FROM positions WHERE id = ?", (position_id,))
        conn.commit()
    finally:
        conn.close()


def pf_insert_position(symbol: str, company: str, buy_date, entry_price: float,
                        quantity: float, amount: float, stop_loss: float,
                        target_price: float, average_gap_pct: float,
                        avg_prices: List[float], strategy: str, notes: str) -> int:
    now = datetime.now().isoformat()
    entries = [{
        'date': buy_date.isoformat() if hasattr(buy_date, 'isoformat') else str(buy_date),
        'price': float(entry_price),
        'qty': float(quantity),
        'amount': float(amount),
        'type': 'INITIAL',
    }]
    conn = get_portfolio_conn()
    try:
        cur = conn.execute("""
            INSERT INTO positions (
                symbol, company, buy_date, strategy, notes, stop_loss, target_price,
                average_gap_pct, avg1_price, avg2_price, avg3_price, avg4_price, avg5_price,
                entries_json, exits_json, status, created_at, updated_at
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            symbol, company, str(buy_date), strategy, notes, stop_loss, target_price,
            average_gap_pct,
            avg_prices[0] if len(avg_prices) > 0 else None,
            avg_prices[1] if len(avg_prices) > 1 else None,
            avg_prices[2] if len(avg_prices) > 2 else None,
            avg_prices[3] if len(avg_prices) > 3 else None,
            avg_prices[4] if len(avg_prices) > 4 else None,
            json.dumps(entries), '[]', 'OPEN', now, now
        ))
        conn.commit()
        return cur.lastrowid
    finally:
        conn.close()


def pf_add_average_entry(position_id: int, date_val, price: float, qty: float,
                          amount: float, entry_type: str = 'AVERAGE') -> None:
    pos = pf_fetch_position(position_id)
    if not pos:
        return
    # Never mutate a closed/exited position's entries.
    if pos.get('status') != 'OPEN':
        logger.warning(
            "pf_add_average_entry: refusing to add to non-OPEN position %s (status=%s)",
            position_id, pos.get('status'),
        )
        return
    entries = json.loads(pos['entries_json'])
    entries.append({
        'date': date_val.isoformat() if hasattr(date_val, 'isoformat') else str(date_val),
        'price': float(price), 'qty': float(qty),
        'amount': float(amount), 'type': entry_type
    })

    # Keep the denormalized avg1..avg5 columns in sync with the averaging
    # entries (everything after the initial buy), derived from entries_json so
    # they can never drift from the source of truth.
    avg_prices = [e['price'] for e in entries[1:]]
    updates: Dict[str, Any] = {'entries_json': json.dumps(entries)}
    for i in range(5):
        updates[f'avg{i + 1}_price'] = float(avg_prices[i]) if i < len(avg_prices) else None
    pf_update_position(position_id, **updates)


def pf_exit_position(position_id: int, date_val, price: float, qty: float, reason: str) -> None:
    pos = pf_fetch_position(position_id)
    if not pos:
        return
    metrics = pf_position_metrics(pos)
    qty = min(qty, metrics['total_quantity'])
    if qty <= 0:
        return
    avg_cost = metrics['average_cost']
    pnl = (price - avg_cost) * qty
    exits = json.loads(pos['exits_json'])
    exits.append({
        'date': date_val.isoformat() if hasattr(date_val, 'isoformat') else str(date_val),
        'price': float(price), 'qty': float(qty),
        'amount': float(price * qty), 'reason': reason, 'pnl': float(pnl)
    })
    remaining_qty = metrics['total_quantity'] - qty
    updates: Dict[str, Any] = {'exits_json': json.dumps(exits)}
    if remaining_qty <= 1e-6:
        updates['status'] = 'CLOSED'
        updates['close_date'] = date_val.isoformat() if hasattr(date_val, 'isoformat') else str(date_val)
        updates['close_price'] = float(price)
    pf_update_position(position_id, **updates)


# ---------------------------- Calculations ------------------------------ #

def compute_average_levels(entry_price: float, gap_pct: float = DEFAULT_AVERAGE_GAP_PCT,
                            levels: int = DEFAULT_AVERAGE_LEVELS) -> List[float]:
    """Cascading averaging levels — each level is gap_pct below the previous
    one, starting from the entry price. Mirrors the strategy's own averaging
    rule ("New signal <= Previous Buy Price * (1 - gap%)")."""
    prices = []
    price = entry_price
    for _ in range(levels):
        price = price * (1 - gap_pct / 100)
        prices.append(round(price, 2))
    return prices


def compute_averaging_plan_summary(entry_price: float, investment_amount: float,
                                    avg_prices: List[float], profit_target_pct: float) -> Dict[str, float]:
    """Projected outcome IF every averaging level is eventually filled with
    the same investment amount as the initial entry (matches how the
    scanner/backtest strategy sizes averaging entries)."""
    all_prices = [entry_price] + list(avg_prices)
    total_investment = investment_amount * len(all_prices)
    total_quantity = sum((investment_amount / p) for p in all_prices if p and p > 0)
    average_cost = total_investment / total_quantity if total_quantity > 0 else 0.0
    new_target = average_cost * (1 + profit_target_pct / 100)
    return {
        'total_investment': total_investment,
        'total_quantity': total_quantity,
        'average_cost': average_cost,
        'breakeven_price': average_cost,
        'new_target_price': new_target,
    }


def pf_position_metrics(pos: Dict[str, Any]) -> Dict[str, Any]:
    """Recomputes live quantity/cost/P&L for a position from its entries and
    exits, using the average-cost method (each exit realizes P&L against the
    average cost at the time, and reduces the remaining cost basis by
    avg_cost * qty_exited — standard weighted-average accounting)."""
    entries = json.loads(pos['entries_json']) if pos.get('entries_json') else []
    exits = json.loads(pos['exits_json']) if pos.get('exits_json') else []

    total_bought_qty = sum(e['qty'] for e in entries)
    total_bought_amt = sum(e['amount'] for e in entries)
    total_sold_qty = sum(x['qty'] for x in exits)
    total_sold_amt = sum(x['amount'] for x in exits)
    realized_pnl = sum(x.get('pnl', 0.0) for x in exits)

    average_cost = (total_bought_amt / total_bought_qty) if total_bought_qty > 0 else 0.0
    remaining_qty = max(total_bought_qty - total_sold_qty, 0.0)
    remaining_invested = average_cost * remaining_qty

    entry_price = entries[0]['price'] if entries else 0.0
    buy_date = entries[0]['date'] if entries else pos.get('buy_date')

    holding_days = 0
    if buy_date:
        # Tolerant parse (handles most stored formats); fall back to strict ISO
        # on the leading yyyy-mm-dd. Only if BOTH fail do we log and leave 0 —
        # the failure is surfaced instead of being swallowed silently.
        bd = pd.to_datetime(buy_date, errors='coerce')
        if pd.isna(bd):
            try:
                bd = pd.Timestamp(datetime.fromisoformat(str(buy_date)[:10]))
            except (ValueError, TypeError):
                bd = pd.NaT
        if pd.isna(bd):
            logger.warning(
                "pf_position_metrics: could not parse buy_date %r (symbol=%s, id=%s); "
                "holding_days defaulted to 0.",
                buy_date, pos.get('symbol'), pos.get('id'),
            )
        else:
            holding_days = (pd.Timestamp.now().normalize() - bd.normalize()).days

    return {
        'total_bought_qty': total_bought_qty,
        'total_bought_amt': total_bought_amt,
        'total_sold_qty': total_sold_qty,
        'total_sold_amt': total_sold_amt,
        'realized_pnl': realized_pnl,
        'average_cost': average_cost,
        'total_quantity': remaining_qty,
        'invested': remaining_invested,
        'entry_price': entry_price,
        'buy_date': buy_date,
        'average_count': len(entries),
        'holding_days': holding_days,
    }


def pf_get_current_prices(symbols: List[str], data_manager: 'DataManager',
                           force_refresh: bool = False) -> Dict[str, float]:
    """Cheap, cached current-price lookup for portfolio symbols. Reuses the
    existing DataManager (same cache as Scanner/Backtest) so viewing the
    Portfolio tab never fires extra network calls unless the cache is stale
    or a refresh was explicitly requested — keeps reruns fast."""
    if not symbols:
        return {}
    cache = st.session_state.setdefault('portfolio_price_cache', {})
    prev_cache = st.session_state.setdefault('portfolio_prev_close', {})

    missing = list(symbols) if force_refresh else [s for s in symbols if s not in cache]
    if missing:
        data_dict = data_manager.batch_download(missing, period="5d", interval="1d", max_workers=4)
        for sym, df in data_dict.items():
            if df is not None and not df.empty:
                cache[sym] = float(df.iloc[-1]['close'])
                prev_cache[sym] = float(df.iloc[-2]['close']) if len(df) > 1 else cache[sym]
        st.session_state['portfolio_price_cache_time'] = datetime.now().isoformat()

    return {s: cache.get(s, 0.0) for s in symbols}


def pf_portfolio_summary(positions: List[Dict[str, Any]], prices: Dict[str, float],
                          capital_base: float) -> Dict[str, float]:
    invested = 0.0
    current_value = 0.0
    unrealized_pnl = 0.0
    realized_pnl = 0.0
    todays_pnl = 0.0
    total_holdings = 0
    total_bought_all = 0.0
    total_sold_all = 0.0

    prev_close_map = st.session_state.get('portfolio_prev_close', {})

    for pos in positions:
        m = pf_position_metrics(pos)
        total_bought_all += m['total_bought_amt']
        total_sold_all += m['total_sold_amt']
        realized_pnl += m['realized_pnl']

        if pos['status'] == 'OPEN' and m['total_quantity'] > 0:
            total_holdings += 1
            price = prices.get(pos['symbol'], m['average_cost'])
            value = m['total_quantity'] * price
            invested += m['invested']
            current_value += value
            unrealized_pnl += (value - m['invested'])
            prev_close = prev_close_map.get(pos['symbol'])
            if prev_close:
                todays_pnl += (price - prev_close) * m['total_quantity']

    available_cash = capital_base - total_bought_all + total_sold_all
    portfolio_value = available_cash + current_value
    return_pct = ((portfolio_value - capital_base) / capital_base * 100) if capital_base > 0 else 0.0

    return {
        'portfolio_value': portfolio_value,
        'invested_amount': invested,
        'available_cash': available_cash,
        'todays_pnl': todays_pnl,
        'unrealized_pnl': unrealized_pnl,
        'realized_pnl': realized_pnl,
        'return_pct': return_pct,
        'total_holdings': total_holdings,
    }


# ------------------------------ UI: Add form ----------------------------- #

def _render_add_position_form(ctx: Dict[str, Any]) -> None:
    config = st.session_state.config
    is_manual = not ctx.get('symbol')

    with st.container(border=True):
        st.markdown(f'<div class="dlg-card-title">{icon("briefcase", 13)} Position Details</div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            if is_manual:
                symbol = st.text_input("Symbol (e.g. RELIANCE.NS)", value="", key="pf_form_symbol").strip().upper()
                if symbol and not symbol.endswith(('.NS', '.BO')):
                    symbol = f"{symbol}.NS"
            else:
                symbol = ctx['symbol']
                st.text_input("Symbol", value=symbol, disabled=True, key="pf_form_symbol_ro")
        with c2:
            if is_manual:
                company = st.text_input("Company", value="", key="pf_form_company")
            else:
                company = ctx.get('company', symbol)
                st.text_input("Company", value=company, disabled=True, key="pf_form_company_ro")

        c3, c4 = st.columns(2)
        with c3:
            buy_date = st.date_input("📅 Buy Date", value=dt_date.today(), key="pf_form_buy_date")
        with c4:
            entry_price = st.number_input(
                "💰 Entry Price (₹)", min_value=0.01,
                value=float(ctx.get('close') or 100.0), step=0.05, key="pf_form_entry_price"
            )

    with st.container(border=True):
        st.markdown(f'<div class="dlg-card-title">{icon("sliders", 13)} Position Sizing</div>', unsafe_allow_html=True)
        mode = st.radio("Size by", options=["Investment Amount", "Quantity"], horizontal=True, key="pf_form_mode")
        c5, c6 = st.columns(2)
        default_amount = float(config.position_size)
        if mode == "Investment Amount":
            with c5:
                amount = st.number_input("Investment Amount (₹)", min_value=0.0, value=default_amount,
                                          step=1000.0, key="pf_form_amount")
            quantity = (amount / entry_price) if entry_price > 0 else 0.0
            with c6:
                st.number_input("Quantity (auto)", value=round(quantity, 4), disabled=True, key="pf_form_qty_disp")
        else:
            with c5:
                default_qty = round(default_amount / entry_price, 2) if entry_price else 0.0
                quantity = st.number_input("Quantity", min_value=0.0, value=default_qty,
                                            step=1.0, key="pf_form_qty")
            amount = quantity * entry_price
            with c6:
                st.number_input("Investment Amount (auto, ₹)", value=round(amount, 2), disabled=True,
                                 key="pf_form_amount_disp")

    with st.container(border=True):
        st.markdown(f'<div class="dlg-card-title">{icon("shield", 13)} Risk Management</div>', unsafe_allow_html=True)
        c7, c8 = st.columns(2)
        with c7:
            stop_loss = st.number_input(
                "🛑 Stop Loss (₹)", min_value=0.0,
                value=float(ctx.get('stop_loss') or round(entry_price * 0.9, 2)),
                step=0.05, key="pf_form_sl"
            )
        with c8:
            target_price = st.number_input(
                "🎯 Target Price (₹)", min_value=0.0,
                value=float(ctx.get('target_price') or round(entry_price * (1 + config.profit_target_pct / 100), 2)),
                step=0.05, key="pf_form_target"
            )
        c9, c10 = st.columns(2)
        with c9:
            strategy = st.text_input("Strategy", value="RSI(14)<35 below EMA200 → Early EMA-Alignment Reclaim", key="pf_form_strategy")
        with c10:
            gap_pct = st.number_input(
                "Averaging Gap (%)", min_value=0.1, max_value=50.0,
                value=float(st.session_state.get('portfolio_default_gap_pct', DEFAULT_AVERAGE_GAP_PCT)),
                step=0.5, key="pf_form_gap"
            )
        notes = st.text_area("Notes", value="", key="pf_form_notes", height=80)

    avg_levels = compute_average_levels(entry_price, gap_pct, DEFAULT_AVERAGE_LEVELS)
    summary = compute_averaging_plan_summary(entry_price, amount, avg_levels, config.profit_target_pct)

    with st.container(border=True):
        st.markdown(f'<div class="dlg-card-title">{icon("layers", 13)} Automatic Averaging Plan</div>', unsafe_allow_html=True)
        avg_cols = st.columns(5)
        for i, (col, price) in enumerate(zip(avg_cols, avg_levels), 1):
            with col:
                st.markdown(render_mini_stat(f"Average {i}", f"₹{price:,.2f}"), unsafe_allow_html=True)

        st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)
        s1, s2, s3, s4 = st.columns(4)
        with s1:
            st.markdown(render_summary_highlight("Avg. Cost (fully avg.)", f"₹{summary['average_cost']:.2f}"), unsafe_allow_html=True)
        with s2:
            st.markdown(render_summary_highlight("Total Investment", f"₹{summary['total_investment']:,.0f}"), unsafe_allow_html=True)
        with s3:
            st.markdown(render_summary_highlight("Total Quantity", f"{summary['total_quantity']:.2f}"), unsafe_allow_html=True)
        with s4:
            st.markdown(render_summary_highlight("New Target (post-avg)", f"₹{summary['new_target_price']:.2f}"), unsafe_allow_html=True)
        st.caption(f"Break-even Price (projected): ₹{summary['breakeven_price']:.2f}")

    if is_manual and not symbol:
        st.caption(f"{icon('alert-triangle', 12, '#F59E0B')} Enter a symbol to enable Add Position.", unsafe_allow_html=True)

    b1, b2 = st.columns(2)
    with b1:
        submit = st.button("✅ Add Position", type="primary", use_container_width=True, key="pf_form_submit")
    with b2:
        cancel = st.button("Cancel", use_container_width=True, key="pf_form_cancel")

    if cancel:
        st.session_state.pop('pf_inline_add_ctx', None)
        _pf_rerun()

    if submit:
        if not symbol:
            st.error("Please enter a valid symbol.")
        else:
            pf_insert_position(
                symbol=symbol, company=company or symbol, buy_date=buy_date,
                entry_price=entry_price, quantity=quantity, amount=amount,
                stop_loss=stop_loss, target_price=target_price,
                average_gap_pct=gap_pct, avg_prices=avg_levels,
                strategy=strategy, notes=notes,
            )
            st.session_state['portfolio_default_gap_pct'] = gap_pct
            st.session_state.pop('pf_inline_add_ctx', None)
            st.success(f"Added {symbol} to Portfolio.")
            _pf_rerun()


if hasattr(st, "dialog"):
    @st.dialog("➕ Add Position to Portfolio")
    def open_add_position_dialog(ctx: Dict[str, Any]) -> None:
        _render_add_position_form(ctx)
else:
    def open_add_position_dialog(ctx: Dict[str, Any]) -> None:
        # Older Streamlit without st.dialog support: fall back to an inline
        # form rendered at the top of the Scanner tab.
        st.session_state['pf_inline_add_ctx'] = ctx


# --------------------------- UI: Edit dialog ------------------------------ #

def _render_edit_position_form(position_id: int) -> None:
    pos = pf_fetch_position(position_id)
    if not pos:
        st.error("Position not found.")
        return
    st.markdown(
        f"<div class='dlg-card-title'>{icon('briefcase', 13)} {pos['symbol'].replace('.NS', '')} "
        f"— {pos.get('company') or ''}</div>", unsafe_allow_html=True
    )
    with st.container(border=True):
        c1, c2 = st.columns(2)
        with c1:
            new_sl = st.number_input("🛑 Stop Loss (₹)", value=float(pos['stop_loss'] or 0.0),
                                      key=f"edit_sl_{position_id}")
            new_strategy = st.text_input("Strategy", value=pos.get('strategy') or '',
                                          key=f"edit_strat_{position_id}")
        with c2:
            new_target = st.number_input("🎯 Target Price (₹)", value=float(pos['target_price'] or 0.0),
                                          key=f"edit_target_{position_id}")
        new_notes = st.text_area("Notes", value=pos.get('notes') or '', key=f"edit_notes_{position_id}", height=80)

    b1, b2 = st.columns(2)
    with b1:
        if st.button("💾 Save Changes", type="primary", use_container_width=True, key=f"edit_save_{position_id}"):
            pf_update_position(position_id, stop_loss=new_sl, target_price=new_target,
                                strategy=new_strategy, notes=new_notes)
            st.success("Position updated.")
            _pf_rerun()
    with b2:
        if st.button("Cancel", use_container_width=True, key=f"edit_cancel_{position_id}"):
            _pf_rerun()


if hasattr(st, "dialog"):
    @st.dialog("✏️ Edit Position")
    def open_edit_position_dialog(position_id: int) -> None:
        _render_edit_position_form(position_id)
else:
    def open_edit_position_dialog(position_id: int) -> None:
        st.session_state['pf_inline_edit_id'] = position_id


# --------------------------- UI: Close dialog ------------------------------ #

def _render_close_position_form(position_id: int, current_price: float) -> None:
    pos = pf_fetch_position(position_id)
    if not pos:
        st.error("Position not found.")
        return
    m = pf_position_metrics(pos)
    st.markdown(
        f"<div class='dlg-card-title'>{icon('lock', 13)} {pos['symbol'].replace('.NS', '')} "
        f"— {m['total_quantity']:.2f} qty @ avg ₹{m['average_cost']:.2f}</div>", unsafe_allow_html=True
    )
    with st.container(border=True):
        c1, c2 = st.columns(2)
        with c1:
            close_date = st.date_input("📅 Close Date", value=dt_date.today(), key=f"close_date_{position_id}")
        with c2:
            close_price = st.number_input("💰 Exit Price (₹)", value=float(current_price),
                                           key=f"close_price_{position_id}")
        reason = st.text_input("Reason", value="Target Achieved", key=f"close_reason_{position_id}")

    projected_pnl = (close_price - m['average_cost']) * m['total_quantity']
    st.markdown(render_summary_highlight(
        "Projected Realized P&L",
        f"<span class='{_pnl_class(projected_pnl)}'>₹{projected_pnl:,.2f}</span>"
    ), unsafe_allow_html=True)

    b1, b2 = st.columns(2)
    with b1:
        if st.button("🔒 Confirm Close", type="primary", use_container_width=True, key=f"close_confirm_{position_id}"):
            pf_exit_position(position_id, close_date, close_price, m['total_quantity'], reason)
            st.success("Position closed.")
            _pf_rerun()
    with b2:
        if st.button("Cancel", use_container_width=True, key=f"close_cancel_{position_id}"):
            _pf_rerun()


if hasattr(st, "dialog"):
    @st.dialog("🔒 Close Position")
    def open_close_position_dialog(position_id: int, current_price: float) -> None:
        _render_close_position_form(position_id, current_price)
else:
    def open_close_position_dialog(position_id: int, current_price: float) -> None:
        st.session_state['pf_inline_close'] = (position_id, current_price)


# ------------------------- UI: More-actions dialog ------------------------ #

def _render_more_actions_form(position_id: int, current_price: float) -> None:
    pos = pf_fetch_position(position_id)
    if not pos:
        st.error("Position not found.")
        return
    m = pf_position_metrics(pos)

    action = st.radio(
        "Action", ["Record Average Buy", "Partial Exit", "Delete"],
        horizontal=True, key=f"pf_action_{position_id}"
    )

    if action == "Record Average Buy":
        avg_prices = [pos.get(f'avg{i}_price') for i in range(1, 6)]
        next_idx = min(m['average_count'], 5)
        suggested_price = avg_prices[next_idx - 1] if next_idx >= 1 and avg_prices[next_idx - 1] else current_price
        with st.container(border=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                avg_date = st.date_input("Buy Date", value=dt_date.today(), key=f"pf_avg_date_{position_id}")
            with c2:
                avg_price = st.number_input("Price (₹)", value=float(suggested_price or current_price),
                                             key=f"pf_avg_price_{position_id}")
            with c3:
                avg_amount = st.number_input("Investment Amount (₹)",
                                              value=float(st.session_state.config.position_size),
                                              key=f"pf_avg_amount_{position_id}")
            avg_qty = (avg_amount / avg_price) if avg_price > 0 else 0.0
            st.caption(f"Quantity (auto): {avg_qty:.2f}")
        if m['average_count'] >= 6:
            st.warning("Maximum of 5 averaging levels already recorded for this position.")
        elif st.button("➕ Record Average Entry", type="primary", use_container_width=True, key=f"pf_avg_submit_{position_id}"):
            pf_add_average_entry(position_id, avg_date, avg_price, avg_qty, avg_amount,
                                  entry_type=f"AVERAGE{next_idx}")
            st.success("Average entry recorded.")
            _pf_rerun()

    elif action == "Partial Exit":
        with st.container(border=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                exit_date = st.date_input("Exit Date", value=dt_date.today(), key=f"pf_pexit_date_{position_id}")
            with c2:
                exit_price = st.number_input("Exit Price (₹)", value=float(current_price),
                                              key=f"pf_pexit_price_{position_id}")
            with c3:
                max_qty = max(float(m['total_quantity']), 0.01)
                exit_qty = st.number_input("Quantity to Exit", min_value=0.0, max_value=max_qty,
                                            value=round(max_qty / 2, 2), key=f"pf_pexit_qty_{position_id}")
            reason = st.text_input("Reason", value="Partial Profit Booking", key=f"pf_pexit_reason_{position_id}")
        if st.button("✂️ Execute Partial Exit", type="primary", use_container_width=True, key=f"pf_pexit_submit_{position_id}"):
            pf_exit_position(position_id, exit_date, exit_price, exit_qty, reason)
            st.success("Partial exit recorded.")
            _pf_rerun()

    elif action == "Delete":
        st.warning(f"This permanently deletes {pos['symbol']} and all its history.")
        confirm = st.checkbox("I understand, delete this position", key=f"pf_del_confirm_{position_id}")
        if st.button("🗑️ Delete Position", use_container_width=True, key=f"pf_del_submit_{position_id}", disabled=not confirm):
            pf_delete_position(position_id)
            st.success("Position deleted.")
            _pf_rerun()


if hasattr(st, "dialog"):
    @st.dialog("⚙️ More Actions")
    def open_more_actions_dialog(position_id: int, current_price: float) -> None:
        _render_more_actions_form(position_id, current_price)
else:
    def open_more_actions_dialog(position_id: int, current_price: float) -> None:
        st.session_state['pf_inline_more'] = (position_id, current_price)


# ------------------------------ UI: Dashboard ----------------------------- #

def render_investment_portfolio() -> None:
    render_page_header("My Portfolio", "Live tracking of positions you've actually taken — independent of the Backtest engine.", "briefcase")
    st.caption("Prices are daily closes (not intraday), so P&L updates once per trading day or on refresh.")

    init_portfolio_db()
    data_manager = st.session_state.data_manager

    top1, top2, top3 = st.columns([2, 1, 1])
    with top1:
        stored_capital = float(pf_get_meta('capital_base', DEFAULT_CAPITAL_BASE))
        capital_base = st.number_input(
            "Total Capital Allocated to This Portfolio (₹)",
            min_value=10000.0, value=stored_capital, step=10000.0, key="pf_capital_base_input"
        )
        if capital_base != stored_capital:
            pf_set_meta('capital_base', capital_base)
    with top2:
        refresh = st.button("🔄 Refresh Prices", use_container_width=True)
    with top3:
        manual_add = st.button("➕ Manual Position", use_container_width=True)

    if manual_add:
        open_add_position_dialog({'symbol': '', 'company': '', 'close': 0.0,
                                   'stop_loss': 0.0, 'target_price': 0.0,
                                   'avg_levels': [0, 0, 0, 0, 0]})

    # Inline fallbacks for Streamlit builds without st.dialog support.
    if not hasattr(st, "dialog"):
        if st.session_state.get('pf_inline_add_ctx'):
            with st.container(border=True):
                st.markdown("#### ➕ Add Position")
                _render_add_position_form(st.session_state['pf_inline_add_ctx'])
        if st.session_state.get('pf_inline_edit_id'):
            with st.container(border=True):
                st.markdown("#### ✏️ Edit Position")
                _render_edit_position_form(st.session_state['pf_inline_edit_id'])
        if st.session_state.get('pf_inline_close'):
            with st.container(border=True):
                st.markdown("#### 🔒 Close Position")
                _render_close_position_form(*st.session_state['pf_inline_close'])
        if st.session_state.get('pf_inline_more'):
            with st.container(border=True):
                st.markdown("#### ⚙️ More Actions")
                _render_more_actions_form(*st.session_state['pf_inline_more'])

    positions = pf_fetch_positions()
    open_positions = [p for p in positions if p['status'] == 'OPEN']
    symbols = sorted({p['symbol'] for p in open_positions})
    prices = pf_get_current_prices(symbols, data_manager, force_refresh=refresh) if symbols else {}

    summary = pf_portfolio_summary(positions, prices, capital_base)

    render_section_title("Portfolio Dashboard", "bar-chart-2")
    r1 = st.columns(4)
    with r1[0]:
        render_kpi_card("Portfolio Value", f"₹{summary['portfolio_value']:,.0f}", "Cash + holdings at market",
                         icon_name="wallet", accent="purple")
    with r1[1]:
        render_kpi_card("Invested Capital", f"₹{summary['invested_amount']:,.0f}", "Cost basis of open positions",
                         icon_name="layers", accent="blue")
    with r1[2]:
        render_kpi_card("Available Cash", f"₹{summary['available_cash']:,.0f}", "Free to deploy",
                         icon_name="dollar-sign", accent="neutral")
    with r1[3]:
        render_kpi_card("Holdings", f"{summary['total_holdings']}", "Open positions",
                         icon_name="briefcase", accent="purple")

    r2 = st.columns(4)
    with r2[0]:
        render_kpi_card("Today's P&L", f"₹{summary['todays_pnl']:,.0f}", "vs. previous close",
                         icon_name="trending-up" if summary['todays_pnl'] >= 0 else "trending-down",
                         accent="green" if summary['todays_pnl'] >= 0 else "red")
    with r2[1]:
        render_kpi_card("Unrealized P&L", f"₹{summary['unrealized_pnl']:,.0f}", "Open positions, mark-to-market",
                         icon_name="activity", accent="green" if summary['unrealized_pnl'] >= 0 else "red")
    with r2[2]:
        render_kpi_card("Realized P&L", f"₹{summary['realized_pnl']:,.0f}", "Booked from closes/partial exits",
                         icon_name="check-circle", accent="green" if summary['realized_pnl'] >= 0 else "red")
    with r2[3]:
        render_kpi_card("Return %", f"{summary['return_pct']:.2f}%", "On allocated capital",
                         icon_name="percent", accent="green" if summary['return_pct'] >= 0 else "red")

    render_section_title("Holdings", "briefcase")

    if not open_positions:
        st.info("No open positions yet. Add one from the Scanner tab, or use 'Manual Position' above.")
    else:
        holding_rows = []
        for pos in open_positions:
            m = pf_position_metrics(pos)
            price = prices.get(pos['symbol'], m['average_cost'])
            current_value = m['total_quantity'] * price
            unreal_pnl = current_value - m['invested']
            ret_pct = (unreal_pnl / m['invested'] * 100) if m['invested'] > 0 else 0.0
            prev_close = st.session_state.get('portfolio_prev_close', {}).get(pos['symbol'])
            todays_pnl = (price - prev_close) * m['total_quantity'] if prev_close else None
            holding_days = 0
            try:
                bd = datetime.fromisoformat(str(m['buy_date'])[:10])
                holding_days = (datetime.now() - bd).days
            except Exception:
                pass
            holding_rows.append({
                'id': pos['id'], 'symbol': pos['symbol'].replace('.NS', ''),
                'buy_date': str(m['buy_date'])[:10], 'holding_days': holding_days,
                'entry_price': m['entry_price'], 'average_cost': m['average_cost'],
                'current_price': price, 'quantity': m['total_quantity'],
                'invested': m['invested'], 'current_value': current_value,
                'todays_pnl': todays_pnl, 'unrealized_pnl': unreal_pnl, 'return_pct': ret_pct,
                'stop_loss': pos['stop_loss'], 'target_price': pos['target_price'],
                'average_count': m['average_count'], 'status': pos['status'],
            })

        page_rows, filtered_rows = render_table_toolbar(
            holding_rows,
            search_keys=['symbol'],
            sort_options={
                'Return %': 'return_pct', 'Unrealized P&L': 'unrealized_pnl',
                'Holding Days': 'holding_days', 'Investment': 'invested', 'Symbol': 'symbol',
            },
            state_prefix="holdings",
            default_page_size=25,
        )

        def _status_badge(v):
            return ("🟢 Open", "badge-green") if v == 'OPEN' else ("⚪ Closed", "badge-gray")

        holdings_columns = [
            {'key': 'symbol', 'label': 'Symbol', 'fmt': lambda v: f"<b>{v}</b>"},
            {'key': 'buy_date', 'label': 'Buy Date'},
            {'key': 'holding_days', 'label': 'Days', 'fmt': lambda v: f"{v}d"},
            {'key': 'entry_price', 'label': 'Entry', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'average_cost', 'label': 'Avg Cost', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'current_price', 'label': 'Current', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'quantity', 'label': 'Qty', 'fmt': lambda v: f"{v:,.2f}"},
            {'key': 'invested', 'label': 'Investment', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'current_value', 'label': 'Value', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'todays_pnl', 'label': "Today's P&L", 'fmt': lambda v: f"₹{v:,.2f}", 'kind': 'pnl'},
            {'key': 'unrealized_pnl', 'label': 'Unrealized P&L', 'fmt': lambda v: f"₹{v:,.2f}", 'kind': 'pnl'},
            {'key': 'return_pct', 'label': 'Return %', 'fmt': lambda v: f"{v:+.2f}%", 'kind': 'pnl'},
            {'key': 'stop_loss', 'label': 'Stop Loss', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'target_price', 'label': 'Target', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'average_count', 'label': 'Avg Count'},
            {'key': 'status', 'label': 'Status', 'kind': 'badge', 'badge_fn': _status_badge},
        ]
        render_premium_table(page_rows, holdings_columns, height=420)

        render_section_title("Actions", "settings")
        act_header = st.columns([1.6, 1, 1, 1])
        for label, c in zip(["Position", "Edit", "Close", "More"], act_header):
            c.markdown(f"<span style='color:var(--text-secondary);font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.04em;'>{label}</span>", unsafe_allow_html=True)

        for row in page_rows:
            price = row['current_price']
            c1, c2, c3, c4 = st.columns([1.6, 1, 1, 1])
            c1.markdown(
                f"**{row['symbol']}** <span class='badge badge-gray' style='margin-left:6px;'>#{row['id']}</span>",
                unsafe_allow_html=True
            )
            with c2:
                if st.button("✏️ Edit", key=f"hedit_{row['id']}", use_container_width=True):
                    open_edit_position_dialog(row['id'])
            with c3:
                if st.button("🔒 Close", key=f"hclose_{row['id']}", use_container_width=True):
                    open_close_position_dialog(row['id'], price)
            with c4:
                if st.button("⋯ More", key=f"hmore_{row['id']}", use_container_width=True):
                    open_more_actions_dialog(row['id'], price)

    closed_positions = [p for p in positions if p['status'] == 'CLOSED']
    if closed_positions:
        render_section_title("Closed Positions", "check-circle")
        closed_rows = []
        for pos in closed_positions:
            m = pf_position_metrics(pos)
            closed_rows.append({
                'symbol': pos['symbol'].replace('.NS', ''), 'buy_date': str(m['buy_date'])[:10],
                'close_date': str(pos.get('close_date') or '')[:10],
                'entry_price': m['entry_price'], 'average_cost': m['average_cost'],
                'close_price': pos.get('close_price') or 0.0,
                'realized_pnl': m['realized_pnl'], 'average_count': m['average_count'],
            })
        closed_columns = [
            {'key': 'symbol', 'label': 'Symbol', 'fmt': lambda v: f"<b>{v}</b>"},
            {'key': 'buy_date', 'label': 'Buy Date'},
            {'key': 'close_date', 'label': 'Close Date'},
            {'key': 'entry_price', 'label': 'Entry', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'average_cost', 'label': 'Avg Cost', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'close_price', 'label': 'Close Price', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'realized_pnl', 'label': 'Realized P&L', 'fmt': lambda v: f"₹{v:,.2f}", 'kind': 'pnl'},
            {'key': 'average_count', 'label': 'Avg Count'},
        ]
        render_premium_table(closed_rows, closed_columns, height=280)


# =============================================================================
# STREAMLIT UI
# =============================================================================

# =============================================================================
# UI KIT — Premium Design System
# =============================================================================
# Pure presentation layer: CSS theme, icon set, and generic render helpers
# (cards, tables, toolbars). Nothing here touches Scanner / SignalGenerator /
# BacktestEngine / Portfolio data logic — nothing here holds any strategy or
# accounting state. Every function takes already-computed values and renders
# them.

ACCENT_COLORS = {
    'purple': '#8B5CF6',
    'blue': '#3B82F6',
    'green': '#10B981',
    'red': '#EF4444',
    'amber': '#F59E0B',
    'neutral': '#64748B',
}

_ICON_PATHS = {
    'activity': '<polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline>',
    'wallet': '<path d="M20 12V8H6a2 2 0 0 1-2-2c0-1.1.9-2 2-2h12v4"></path><path d="M4 6v12a2 2 0 0 0 2 2h14v-4"></path><path d="M18 12a2 2 0 0 0 0 4h4v-4Z"></path>',
    'trending-up': '<polyline points="23 6 13.5 15.5 8.5 10.5 1 18"></polyline><polyline points="17 6 23 6 23 12"></polyline>',
    'trending-down': '<polyline points="23 18 13.5 8.5 8.5 13.5 1 6"></polyline><polyline points="17 18 23 18 23 12"></polyline>',
    'dollar-sign': '<line x1="12" y1="1" x2="12" y2="23"></line><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"></path>',
    'layers': '<polygon points="12 2 2 7 12 12 22 7 12 2"></polygon><polyline points="2 17 12 22 22 17"></polyline><polyline points="2 12 12 17 22 12"></polyline>',
    'target': '<circle cx="12" cy="12" r="10"></circle><circle cx="12" cy="12" r="6"></circle><circle cx="12" cy="12" r="2"></circle>',
    'shield': '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10Z"></path>',
    'clock': '<circle cx="12" cy="12" r="10"></circle><polyline points="12 6 12 12 16 14"></polyline>',
    'search': '<circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line>',
    'briefcase': '<rect x="2" y="7" width="20" height="14" rx="2" ry="2"></rect><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"></path>',
    'plus-circle': '<circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="16"></line><line x1="8" y1="12" x2="16" y2="12"></line>',
    'check-circle': '<path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline>',
    'edit': '<path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path><path d="M18.5 2.5a2.12 2.12 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5Z"></path>',
    'lock': '<rect x="3" y="11" width="18" height="11" rx="2" ry="2"></rect><path d="M7 11V7a5 5 0 0 1 10 0v4"></path>',
    'trash': '<polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>',
    'scissors': '<circle cx="6" cy="6" r="3"></circle><circle cx="6" cy="18" r="3"></circle><line x1="20" y1="4" x2="8.12" y2="15.88"></line><line x1="14.47" y1="14.48" x2="20" y2="20"></line><line x1="8.12" y1="8.12" x2="12" y2="12"></line>',
    'more-horizontal': '<circle cx="12" cy="12" r="1"></circle><circle cx="19" cy="12" r="1"></circle><circle cx="5" cy="12" r="1"></circle>',
    'refresh-cw': '<polyline points="23 4 23 10 17 10"></polyline><polyline points="1 20 1 14 7 14"></polyline><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"></path>',
    'arrow-up-right': '<line x1="7" y1="17" x2="17" y2="7"></line><polyline points="7 7 17 7 17 17"></polyline>',
    'arrow-down-right': '<line x1="7" y1="7" x2="17" y2="17"></line><polyline points="17 7 17 17 7 17"></polyline>',
    'bar-chart-2': '<line x1="18" y1="20" x2="18" y2="10"></line><line x1="12" y1="20" x2="12" y2="4"></line><line x1="6" y1="20" x2="6" y2="14"></line>',
    'sliders': '<line x1="4" y1="21" x2="4" y2="14"></line><line x1="4" y1="10" x2="4" y2="3"></line><line x1="12" y1="21" x2="12" y2="12"></line><line x1="12" y1="8" x2="12" y2="3"></line><line x1="20" y1="21" x2="20" y2="16"></line><line x1="20" y1="12" x2="20" y2="3"></line><line x1="1" y1="14" x2="7" y2="14"></line><line x1="9" y1="8" x2="15" y2="8"></line><line x1="17" y1="16" x2="23" y2="16"></line>',
    'zap': '<polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"></polygon>',
    'list': '<line x1="8" y1="6" x2="21" y2="6"></line><line x1="8" y1="12" x2="21" y2="12"></line><line x1="8" y1="18" x2="21" y2="18"></line><line x1="3" y1="6" x2="3.01" y2="6"></line><line x1="3" y1="12" x2="3.01" y2="12"></line><line x1="3" y1="18" x2="3.01" y2="18"></line>',
    'calendar': '<rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect><line x1="16" y1="2" x2="16" y2="6"></line><line x1="8" y1="2" x2="8" y2="6"></line><line x1="3" y1="10" x2="21" y2="10"></line>',
    'settings': '<circle cx="12" cy="12" r="3"></circle><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1Z"></path>',
    'alert-triangle': '<path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0Z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line>',
    'percent': '<line x1="19" y1="5" x2="5" y2="19"></line><circle cx="6.5" cy="6.5" r="2.5"></circle><circle cx="17.5" cy="17.5" r="2.5"></circle>',
}


def icon(name: str, size: int = 16, color: str = "currentColor", stroke_width: float = 2.0) -> str:
    """Inline SVG icon (lucide-style), stroke color driven by `color` so it
    inherits the surrounding accent automatically."""
    path = _ICON_PATHS.get(name, _ICON_PATHS['activity'])
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" '
        f'viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="{stroke_width}" '
        f'stroke-linecap="round" stroke-linejoin="round" style="vertical-align:middle;flex-shrink:0;">'
        f'{path}</svg>'
    )


PREMIUM_CSS = """
<style>
:root {
    --bg-primary: #0B1220;
    --bg-card: #111827;
    --bg-secondary: #1F2937;
    --border-color: rgba(148, 163, 184, 0.12);
    --border-strong: rgba(148, 163, 184, 0.24);
    --accent-purple: #8B5CF6;
    --accent-blue: #3B82F6;
    --accent-gradient: linear-gradient(135deg, #8B5CF6 0%, #3B82F6 100%);
    --profit: #10B981;
    --profit-bg: rgba(16, 185, 129, 0.14);
    --loss: #EF4444;
    --loss-bg: rgba(239, 68, 68, 0.14);
    --neutral: #94A3B8;
    --neutral-bg: rgba(148, 163, 184, 0.14);
    --warn: #F59E0B;
    --warn-bg: rgba(245, 158, 11, 0.14);
    --text-primary: #F1F5F9;
    --text-secondary: #94A3B8;
    --text-muted: #64748B;
    --radius-sm: 6px;
    --radius-md: 10px;
    --radius-lg: 14px;
    --shadow-soft: 0 4px 18px rgba(0,0,0,0.35);
}

html, body, .stApp { background: var(--bg-primary) !important; color: var(--text-primary); }
* { font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif !important; }

.block-container { padding-top: 1.1rem !important; padding-bottom: 2rem !important; max-width: 100% !important; }
div[data-testid="stToolbar"] { right: 8px; }

::-webkit-scrollbar { width: 8px; height: 8px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: var(--bg-secondary); border-radius: 4px; }
::-webkit-scrollbar-thumb:hover { background: var(--border-strong); }

section[data-testid="stSidebar"] { background: var(--bg-card) !important; border-right: 1px solid var(--border-color); }
section[data-testid="stSidebar"] .block-container { padding-top: 1.2rem !important; }

h1, h2, h3, h4 { color: var(--text-primary) !important; letter-spacing: -0.01em; }
h1 { font-weight: 700 !important; }
h2, h3 { font-weight: 600 !important; }
p, span, div, label { color: var(--text-primary); }
.stCaption, [data-testid="stCaptionContainer"] { color: var(--text-secondary) !important; }

hr { border-color: var(--border-color) !important; margin: 0.9rem 0 !important; }

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    gap: 4px; background: var(--bg-card); padding: 6px; border-radius: var(--radius-lg);
    border: 1px solid var(--border-color); flex-wrap: wrap;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px; padding: 8px 16px; color: var(--text-secondary) !important;
    font-weight: 500; transition: all .15s ease; background: transparent;
}
.stTabs [data-baseweb="tab"] p { color: inherit !important; font-weight: 500; }
.stTabs [aria-selected="true"] { background: var(--accent-gradient) !important; }
.stTabs [aria-selected="true"] p { color: #fff !important; }
.stTabs [data-baseweb="tab-highlight"] { background: transparent !important; }
.stTabs [data-baseweb="tab-border"] { display: none !important; }

/* Buttons */
.stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {
    border-radius: 8px !important; border: 1px solid var(--border-strong) !important;
    background: var(--bg-secondary) !important; color: var(--text-primary) !important;
    font-weight: 500 !important; padding: 0.42rem 0.9rem !important;
    transition: transform .12s ease, box-shadow .12s ease, border-color .12s ease !important;
    box-shadow: none !important;
}
.stButton > button:hover:not(:disabled), .stDownloadButton > button:hover, .stFormSubmitButton > button:hover {
    transform: translateY(-1px); border-color: var(--accent-blue) !important;
    box-shadow: 0 4px 14px rgba(59,130,246,0.25) !important;
}
.stButton > button[kind="primary"] {
    background: var(--accent-gradient) !important; border: none !important; color: #fff !important;
    box-shadow: 0 4px 14px rgba(139,92,246,0.35) !important;
}
.stButton > button[kind="primary"]:hover { box-shadow: 0 6px 22px rgba(139,92,246,0.5) !important; }
.stButton > button:disabled { opacity: 0.5 !important; transform: none !important; box-shadow: none !important; }

/* Inputs */
.stTextInput input, .stNumberInput input, .stDateInput input, .stTextArea textarea {
    background: var(--bg-secondary) !important; border: 1px solid var(--border-strong) !important;
    border-radius: 8px !important; color: var(--text-primary) !important;
}
.stTextInput input:focus, .stNumberInput input:focus, .stTextArea textarea:focus {
    border-color: var(--accent-blue) !important; box-shadow: 0 0 0 2px rgba(59,130,246,0.18) !important;
}
.stSelectbox [data-baseweb="select"] > div, .stMultiSelect [data-baseweb="select"] > div {
    background: var(--bg-secondary) !important; border-radius: 8px !important; border: 1px solid var(--border-strong) !important;
}
.stRadio [role="radiogroup"] label, .stCheckbox label { color: var(--text-primary) !important; }

/* Expander / bordered containers */
.stExpander, div[data-testid="stExpander"] {
    background: var(--bg-card); border: 1px solid var(--border-color) !important; border-radius: var(--radius-lg) !important;
}
div[data-testid="stVerticalBlockBorderWrapper"] > div {
    border-color: var(--border-color) !important; border-radius: var(--radius-lg) !important;
}

/* Native dataframe */
div[data-testid="stDataFrame"] { border-radius: var(--radius-lg); overflow: hidden; border: 1px solid var(--border-color); }

/* Metric */
div[data-testid="stMetric"] {
    background: var(--bg-card); border: 1px solid var(--border-color); border-radius: var(--radius-md);
    padding: 10px 14px;
}

/* ---------------- KPI Cards ---------------- */
.kpi-card {
    background: var(--bg-card); border: 1px solid var(--border-color); border-left: 3px solid var(--neutral);
    border-radius: var(--radius-lg); padding: 14px 16px; box-shadow: var(--shadow-soft);
    transition: transform .15s ease, box-shadow .15s ease; margin-bottom: 10px; height: 100%;
}
.kpi-card:hover { transform: translateY(-2px); box-shadow: 0 10px 26px rgba(0,0,0,0.45); }
.kpi-card-top { display: flex; align-items: center; gap: 8px; margin-bottom: 10px; }
.kpi-card-icon { width: 30px; height: 30px; border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
.kpi-card-title { color: var(--text-secondary); font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: .04em; }
.kpi-card-value { font-size: 23px; font-weight: 700; color: var(--text-primary); line-height: 1.2; }
.kpi-card-subtitle { font-size: 12px; color: var(--text-muted); margin-top: 3px; }
.kpi-card-trend { display: flex; align-items: center; gap: 4px; margin-top: 8px; font-size: 12px; font-weight: 600; }

/* ---------------- Badges ---------------- */
.badge {
    display: inline-flex; align-items: center; gap: 4px; padding: 3px 10px; border-radius: 999px;
    font-size: 11px; font-weight: 600; white-space: nowrap; line-height: 1.6;
}
.badge-green { background: var(--profit-bg); color: var(--profit); }
.badge-red { background: var(--loss-bg); color: var(--loss); }
.badge-blue { background: rgba(59,130,246,0.14); color: var(--accent-blue); }
.badge-purple { background: rgba(139,92,246,0.14); color: var(--accent-purple); }
.badge-yellow { background: var(--warn-bg); color: var(--warn); }
.badge-gray { background: var(--neutral-bg); color: var(--neutral); }

.pnl-pos { color: var(--profit); font-weight: 600; }
.pnl-neg { color: var(--loss); font-weight: 600; }
.pnl-flat { color: var(--text-secondary); font-weight: 500; }

/* ---------------- Premium Table ---------------- */
.pt-wrap { border: 1px solid var(--border-color); border-radius: var(--radius-lg); overflow: hidden; box-shadow: var(--shadow-soft); background: var(--bg-card); }
.pt-scroll { overflow: auto; }
table.pt { width: 100%; border-collapse: separate; border-spacing: 0; font-size: 12.5px; }
table.pt thead th {
    position: sticky; top: 0; background: var(--bg-secondary); color: var(--text-secondary);
    text-transform: uppercase; font-size: 10.5px; letter-spacing: .04em; font-weight: 700;
    padding: 10px 12px; border-bottom: 1px solid var(--border-strong); text-align: left; z-index: 2; white-space: nowrap;
}
table.pt tbody td { padding: 8px 12px; border-bottom: 1px solid var(--border-color); color: var(--text-primary); white-space: nowrap; }
table.pt tbody tr { transition: background .1s ease; }
table.pt tbody tr:nth-child(even) { background: rgba(255,255,255,0.015); }
table.pt tbody tr:hover { background: rgba(139,92,246,0.07); }
table.pt tbody tr:last-child td { border-bottom: none; }
table.pt thead th:first-child, table.pt tbody td:first-child {
    position: sticky; left: 0; background: var(--bg-card); z-index: 1; font-weight: 700; box-shadow: 2px 0 4px rgba(0,0,0,0.25);
}
table.pt thead th:first-child { z-index: 3; background: var(--bg-secondary); }
table.pt tbody tr:nth-child(even) td:first-child { background: #131c2e; }

.pt-toolbar-caption { color: var(--text-muted); font-size: 11.5px; padding: 4px 2px 8px 2px; }

/* ---------------- Dialog / form cards ---------------- */
.dlg-card-title { font-size: 12px; font-weight: 700; color: var(--text-secondary); text-transform: uppercase; letter-spacing: .05em; margin: 4px 0 8px 0; display: flex; align-items: center; gap: 6px; }
.stat-chip { background: var(--bg-card); border: 1px solid var(--border-color); border-radius: 8px; padding: 10px 8px; text-align: center; }
.stat-chip-label { font-size: 10.5px; color: var(--text-muted); text-transform: uppercase; letter-spacing: .03em; }
.stat-chip-value { font-size: 15px; font-weight: 700; color: var(--text-primary); margin-top: 3px; }
.summary-highlight {
    background: linear-gradient(135deg, rgba(139,92,246,0.14), rgba(59,130,246,0.08));
    border: 1px solid rgba(139,92,246,0.32); border-radius: var(--radius-md); padding: 12px 14px; text-align: center;
}
.summary-highlight-label { font-size: 10.5px; color: var(--text-secondary); text-transform: uppercase; letter-spacing: .04em; }
.summary-highlight-value { font-size: 19px; font-weight: 700; margin-top: 4px; color: var(--text-primary); }

/* ---------------- Page headers ---------------- */
.page-title { font-size: 24px; font-weight: 700; color: var(--text-primary); display: flex; align-items: center; gap: 10px; margin-bottom: 2px; }
.page-subtitle { color: var(--text-secondary); font-size: 12.5px; margin-bottom: 14px; }
.section-title { font-size: 14px; font-weight: 700; color: var(--text-primary); display: flex; align-items: center; gap: 8px; margin: 16px 0 8px 0; }

/* ---------------- Skeleton loader ---------------- */
@keyframes pt-shimmer { 0% { background-position: -400px 0; } 100% { background-position: 400px 0; } }
.pt-skeleton-row { height: 14px; border-radius: 4px; margin: 6px 0; background: linear-gradient(90deg, var(--bg-secondary) 25%, rgba(255,255,255,0.07) 37%, var(--bg-secondary) 63%); background-size: 400px 100%; animation: pt-shimmer 1.3s ease infinite; }

/* Plotly chart container */
.chart-card { background: var(--bg-card); border: 1px solid var(--border-color); border-radius: var(--radius-lg); padding: 10px; box-shadow: var(--shadow-soft); }

/* ---------------- Section titles: subtle divider for rhythm ---------------- */
.section-title {
    padding-bottom: 8px; border-bottom: 1px solid var(--border-color);
}

/* ---------------- Native alerts (st.info/warning/error/success) ----------------
   Unthemed Streamlit alerts render as light boxes and clash hard against the
   dark terminal background. Recolor them as left-accent cards consistent
   with the rest of the design system, matching each semantic color. */
div[data-testid="stAlertContentInfo"],
div[data-testid="stAlertContentWarning"],
div[data-testid="stAlertContentError"],
div[data-testid="stAlertContentSuccess"] {
    color: var(--text-primary) !important;
}
div[data-testid="stNotification"], div.stAlert {
    background: var(--bg-card) !important; border-radius: var(--radius-md) !important;
    border: 1px solid var(--border-color) !important; box-shadow: var(--shadow-soft) !important;
}
div[data-testid="stNotification"]:has(div[data-testid="stAlertContentInfo"]),
div.stAlert:has(div[data-testid="stAlertContentInfo"]) { border-left: 3px solid var(--accent-blue) !important; }
div[data-testid="stNotification"]:has(div[data-testid="stAlertContentWarning"]),
div.stAlert:has(div[data-testid="stAlertContentWarning"]) { border-left: 3px solid var(--warn) !important; }
div[data-testid="stNotification"]:has(div[data-testid="stAlertContentError"]),
div.stAlert:has(div[data-testid="stAlertContentError"]) { border-left: 3px solid var(--loss) !important; }
div[data-testid="stNotification"]:has(div[data-testid="stAlertContentSuccess"]),
div.stAlert:has(div[data-testid="stAlertContentSuccess"]) { border-left: 3px solid var(--profit) !important; }
div[data-testid="stNotification"] svg, div.stAlert svg { opacity: 0.85; }

/* ---------------- Progress bar ---------------- */
div[data-testid="stProgress"] > div > div > div { background: var(--accent-gradient) !important; }
div[data-testid="stProgress"] > div > div { background: var(--bg-secondary) !important; border-radius: 999px !important; }

/* ---------------- Toggle / checkbox accent (default Streamlit red clashes with theme) ---------------- */
[data-testid="stCheckbox"] [role="checkbox"][aria-checked="true"],
[data-baseweb="checkbox"] [aria-checked="true"] > div:first-child {
    background: var(--accent-purple) !important; border-color: var(--accent-purple) !important;
}
[data-testid="stToggle"] [aria-checked="true"],
[data-baseweb="switch"] [aria-checked="true"] {
    background: var(--accent-gradient) !important;
}
[data-testid="stToggle"] label div[data-baseweb="switch"] { background: var(--bg-secondary); }

/* ---------------- Popover menus (select/multiselect dropdown lists) ----------------
   BaseWeb renders these in a portal outside .stApp, so they inherit the
   Streamlit default (light) theme unless targeted directly - normally shows
   up as a jarring white dropdown against the dark app. */
ul[role="listbox"], div[data-baseweb="popover"] div[data-baseweb="menu"] {
    background: var(--bg-secondary) !important; border: 1px solid var(--border-strong) !important;
    border-radius: 8px !important; box-shadow: 0 10px 30px rgba(0,0,0,0.5) !important;
}
ul[role="listbox"] li, div[data-baseweb="popover"] div[data-baseweb="menu"] li {
    background: transparent !important; color: var(--text-primary) !important;
}
ul[role="listbox"] li:hover, div[data-baseweb="popover"] div[data-baseweb="menu"] li:hover {
    background: rgba(139,92,246,0.14) !important;
}
ul[role="listbox"] li[aria-selected="true"] { background: rgba(139,92,246,0.22) !important; }

/* ---------------- File uploader dropzone ---------------- */
[data-testid="stFileUploaderDropzone"] {
    background: var(--bg-secondary) !important; border: 1px dashed var(--border-strong) !important;
    border-radius: var(--radius-md) !important;
}
[data-testid="stFileUploaderDropzone"] * { color: var(--text-secondary) !important; }

/* ---------------- Spinner ---------------- */
div[data-testid="stSpinner"] > div { color: var(--text-secondary) !important; }

/* ---------------- Radio pills ---------------- */
.stRadio [role="radiogroup"] { gap: 4px; }

/* ---------------- Font smoothing ---------------- */
html, body { -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale; }
</style>
"""


def render_mini_stat(label: str, value: str) -> str:
    return f"""<div class="stat-chip"><div class="stat-chip-label">{label}</div><div class="stat-chip-value">{value}</div></div>"""


def render_summary_highlight(label: str, value: str) -> str:
    return f"""<div class="summary-highlight"><div class="summary-highlight-label">{label}</div><div class="summary-highlight-value">{value}</div></div>"""


def render_kpi_card(title: str, value: str, subtitle: Optional[str] = None,
                     icon_name: str = "activity", accent: str = "neutral",
                     trend: Optional[str] = None, trend_positive: Optional[bool] = None) -> None:
    """Premium KPI card: icon chip, large value, subtitle, colored left
    border, optional trend indicator. `accent` picks the border/icon color;
    `trend_positive` overrides auto sign-detection on the trend string."""
    border_color = ACCENT_COLORS.get(accent, ACCENT_COLORS['neutral'])
    trend_html = ""
    if trend:
        is_pos = trend_positive if trend_positive is not None else (not trend.strip().startswith('-'))
        trend_color = ACCENT_COLORS['green'] if is_pos else ACCENT_COLORS['red']
        arrow = icon('arrow-up-right', 12, trend_color) if is_pos else icon('arrow-down-right', 12, trend_color)
        trend_html = f'<div class="kpi-card-trend" style="color:{trend_color};">{arrow}{trend}</div>'
    icon_svg = icon(icon_name, 16, border_color)
    st.markdown(f"""
        <div class="kpi-card" style="border-left-color:{border_color};">
            <div class="kpi-card-top">
                <div class="kpi-card-icon" style="background:{border_color}22;">{icon_svg}</div>
                <div class="kpi-card-title">{title}</div>
            </div>
            <div class="kpi-card-value">{value}</div>
            {f'<div class="kpi-card-subtitle">{subtitle}</div>' if subtitle else ''}
            {trend_html}
        </div>
    """, unsafe_allow_html=True)


def _pnl_class(numeric: float) -> str:
    if numeric > 1e-9:
        return "pnl-pos"
    if numeric < -1e-9:
        return "pnl-neg"
    return "pnl-flat"


def render_premium_table(rows: List[Dict[str, Any]], columns: List[Dict[str, Any]],
                          height: int = 420, empty_message: str = "No data to display.") -> None:
    """Generic premium HTML table: sticky header, zebra rows, hover
    highlight, frozen first column, badges and colored P&L cells.

    Each column spec: {'key', 'label', 'fmt' (value->str), 'kind' in
    {'text','pnl','badge'}, 'badge_fn' (value -> (label, css_class)),
    'raw_key' (optional numeric field used for pnl sign / sorting instead of key)}
    """
    if not rows:
        st.info(empty_message)
        return

    header_html = "".join(f"<th>{c['label']}</th>" for c in columns)
    body_rows = []
    for row in rows:
        cells = []
        for c in columns:
            key = c['key']
            raw = row.get(key)
            kind = c.get('kind', 'text')
            fmt = c.get('fmt', str)
            if kind == 'badge' and c.get('badge_fn'):
                label, css_class = c['badge_fn'](raw)
                cells.append(f'<td><span class="badge {css_class}">{_html.escape(str(label))}</span></td>')
            elif kind == 'pnl':
                if raw is None:
                    cells.append('<td>—</td>')
                else:
                    try:
                        numeric = float(row.get(c.get('raw_key', key), raw))
                    except (TypeError, ValueError):
                        numeric = 0.0
                    cells.append(f'<td><span class="{_pnl_class(numeric)}">{_html.escape(str(fmt(raw)))}</span></td>')
            else:
                display_val = "—" if raw is None else _html.escape(str(fmt(raw)))
                cells.append(f"<td>{display_val}</td>")
        body_rows.append(f"<tr>{''.join(cells)}</tr>")

    html = f"""
    <div class="pt-wrap"><div class="pt-scroll" style="max-height:{height}px;">
    <table class="pt"><thead><tr>{header_html}</tr></thead><tbody>{''.join(body_rows)}</tbody></table>
    </div></div>
    """
    st.markdown(html, unsafe_allow_html=True)


def render_table_toolbar(source_rows: List[Dict[str, Any]], search_keys: List[str],
                          sort_options: Dict[str, str], state_prefix: str,
                          filter_options: Optional[Dict[str, List[str]]] = None,
                          default_page_size: int = 25) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """Search box + sort + optional filters + pagination, applied to a list
    of plain dicts (kept independent of pandas so it works the same for
    Scanner rows and Portfolio holdings). Returns (current_page_rows, all_filtered_rows)."""
    n_filters = len(filter_options) if filter_options else 0
    top_cols = st.columns([2.4, 1.3, 1.0, 0.9] + [1.1] * n_filters)

    with top_cols[0]:
        query = st.text_input("Search", value="", key=f"{state_prefix}_search",
                               placeholder=f"{('🔎 ')}Search symbol, company…", label_visibility="collapsed")
    with top_cols[1]:
        sort_label = st.selectbox("Sort by", list(sort_options.keys()), key=f"{state_prefix}_sort",
                                   label_visibility="collapsed")
    with top_cols[2]:
        sort_dir = st.selectbox("Direction", ["↓ Desc", "↑ Asc"], key=f"{state_prefix}_dir",
                                 label_visibility="collapsed")
    with top_cols[3]:
        page_size = st.selectbox("Rows", [10, 25, 50, 100], index=[10, 25, 50, 100].index(default_page_size)
                                  if default_page_size in [10, 25, 50, 100] else 1,
                                  key=f"{state_prefix}_pagesize", label_visibility="collapsed")

    filtered = list(source_rows)
    if query:
        q = query.strip().lower()
        filtered = [r for r in filtered if any(q in str(r.get(k, '')).lower() for k in search_keys)]

    if filter_options:
        for i, (fkey, options) in enumerate(filter_options.items()):
            with top_cols[4 + i]:
                chosen = st.selectbox(fkey, ["All"] + options, key=f"{state_prefix}_filter_{fkey}",
                                       label_visibility="collapsed")
            if chosen != "All":
                filtered = [r for r in filtered if str(r.get(fkey)) == chosen]

    sort_key = sort_options[sort_label]
    reverse = (sort_dir == "↓ Desc")
    try:
        filtered = sorted(filtered, key=lambda r: (r.get(sort_key) is None, r.get(sort_key)), reverse=reverse)
    except TypeError:
        pass

    total = len(filtered)
    total_pages = max(1, (total + page_size - 1) // page_size)
    page_key = f"{state_prefix}_page"
    if page_key not in st.session_state:
        st.session_state[page_key] = 1
    st.session_state[page_key] = max(1, min(st.session_state[page_key], total_pages))

    pcol1, pcol2, pcol3 = st.columns([1, 3, 1])
    with pcol1:
        if st.button("‹ Prev", key=f"{state_prefix}_prev", use_container_width=True,
                      disabled=st.session_state[page_key] <= 1):
            st.session_state[page_key] -= 1
    with pcol2:
        st.markdown(
            f'<div class="pt-toolbar-caption" style="text-align:center;padding-top:9px;">'
            f'Page {st.session_state[page_key]} of {total_pages} · {total} result(s)</div>',
            unsafe_allow_html=True
        )
    with pcol3:
        if st.button("Next ›", key=f"{state_prefix}_next", use_container_width=True,
                      disabled=st.session_state[page_key] >= total_pages):
            st.session_state[page_key] += 1

    start = (st.session_state[page_key] - 1) * page_size
    end = start + page_size
    return filtered[start:end], filtered


# ---------------------------- Plotly theming ------------------------------ #

PLOTLY_COLORWAY = ['#8B5CF6', '#3B82F6', '#10B981', '#EF4444', '#F59E0B', '#EC4899', '#14B8A6']


def apply_plotly_theme(fig: go.Figure, height: Optional[int] = None) -> go.Figure:
    """Cosmetic-only theming shared by every chart in the app (equity curve,
    drawdown, histogram, trade chart). Never touches trace data/values —
    only colors, fonts, backgrounds, grid and hover styling."""
    fig.update_layout(
        template='plotly_dark',
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(family='Inter, -apple-system, Segoe UI, sans-serif', color='#94A3B8', size=12),
        colorway=PLOTLY_COLORWAY,
        margin=dict(l=44, r=24, t=36, b=40),
        legend=dict(bgcolor='rgba(0,0,0,0)', font=dict(color='#CBD5E1', size=11)),
        hoverlabel=dict(bgcolor='#1F2937', font=dict(color='#F1F5F9', size=12), bordercolor='#334155'),
    )
    fig.update_xaxes(gridcolor='rgba(148,163,184,0.08)', zerolinecolor='rgba(148,163,184,0.15)', showline=False)
    fig.update_yaxes(gridcolor='rgba(148,163,184,0.08)', zerolinecolor='rgba(148,163,184,0.15)', showline=False)
    if height:
        fig.update_layout(height=height)
    return fig


def render_chart_card(fig: go.Figure, key: Optional[str] = None, height: Optional[int] = None) -> None:
    """Renders a Plotly figure inside a bordered, rounded card — uses a real
    st.container(border=True) so the chart is an actual DOM child of the
    styled wrapper (a plain st.markdown div can't contain a widget)."""
    if height:
        fig.update_layout(height=height)
    with st.container(border=True):
        st.plotly_chart(fig, use_container_width=True, key=key)


def render_page_header(title: str, subtitle: str = "", icon_name: str = "activity") -> None:
    st.markdown(f"""
        <div class="page-title">{icon(icon_name, 22, '#8B5CF6')} {title}</div>
        {f'<div class="page-subtitle">{subtitle}</div>' if subtitle else ''}
    """, unsafe_allow_html=True)


def render_section_title(title: str, icon_name: str = "list",
                          badge_text: Optional[str] = None, badge_class: str = "badge-blue") -> None:
    """Section header. Optional `badge_text` renders a small pill after the
    title (e.g. flagging a section as unrealized/mark-to-market) without
    changing any existing call site's behavior."""
    badge_html = (
        f'<span class="badge {badge_class}" style="margin-left:8px;">{badge_text}</span>'
        if badge_text else ''
    )
    st.markdown(
        f'<div class="section-title">{icon(icon_name, 15, "#94A3B8")} {title}{badge_html}</div>',
        unsafe_allow_html=True,
    )


def parse_symbol_file(uploaded_file) -> List[str]:
    """
    Parse an uploaded Excel (.xlsx/.xls) or CSV file into a clean list of
    NSE ticker symbols.

    Looks for a column named Symbol/Symbols/Ticker/Tickers/Scrip/Stock
    (case-insensitive); falls back to the first column if none match.
    Adds the '.NS' suffix automatically if the symbol doesn't already end
    in '.NS' or '.BO'. Strips whitespace, upper-cases, and de-duplicates
    while preserving order.
    """
    filename = getattr(uploaded_file, "name", "").lower()

    if filename.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    if df is None or df.empty:
        return []

    candidate_cols = {
        "symbol", "symbols", "ticker", "tickers",
        "scrip", "scrip code", "stock", "stock symbol", "nse symbol",
    }
    col_name = None
    for col in df.columns:
        if str(col).strip().lower() in candidate_cols:
            col_name = col
            break
    if col_name is None:
        col_name = df.columns[0]

    raw_symbols = df[col_name].dropna().astype(str).tolist()

    cleaned: List[str] = []
    seen: Set[str] = set()
    for s in raw_symbols:
        s = s.strip().upper().replace(" ", "")
        if not s or s in ("NAN", "NONE"):
            continue
        if s.endswith(".NSE"):
            s = s[:-4]
        if not s.endswith(".NS") and not s.endswith(".BO"):
            s = f"{s}.NS"
        if s not in seen:
            seen.add(s)
            cleaned.append(s)

    return cleaned


def init_session_state():
    defaults = {
        'data_manager': DataManager(),
        'backtest_results': {},
        'scanner_results': [],
        'all_trades': [],
        'metrics': None,
        'config': StrategyConfig(),
        'selected_symbol': None,
        'backtest_complete': False,
        'scan_complete': False,
        'daily_equity': [],
        'backtest_diagnostics': None,
        'open_positions_snapshot': [],
        'unrealized_summary': None,
        'active_symbols': NIFTY200_SYMBOLS,
        'universe_source_label': f"NIFTY 200 ({len(NIFTY200_SYMBOLS)} symbols)",
        # Portfolio module (investment tracking) defaults
        'portfolio_default_gap_pct': DEFAULT_AVERAGE_GAP_PCT,
        'portfolio_price_cache': {},
        'portfolio_prev_close': {},
        'portfolio_db_ready': False,
        # Shared OHLC cache so Buy Signal Analysis can reuse data already
        # downloaded by the Backtest (no re-download when the period matches).
        'signal_universe_data': {'data': {}, 'period': None},
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value
    if not st.session_state.get('portfolio_db_ready'):
        init_portfolio_db()
        st.session_state['portfolio_db_ready'] = True


def render_sidebar():
    with st.sidebar:
        st.title(f"{APP_ICON} {APP_NAME}")
        st.markdown(f"**Version:** {APP_VERSION}")
        st.markdown("---")

        st.header("Strategy Parameters")

        st.caption("RSI Tracking + EMA Alignment. Arm on RSI < threshold "
                   "**and** Close < EMA(long); buy on the early reclaim "
                   "**Close > EMA(long) > EMA(mid) > EMA(short)** "
                   "(price back above the long EMA while both faster EMAs "
                   "are still below it, mid above short).")
        rsi_period = st.number_input("RSI Period", min_value=5, max_value=50, value=14, step=1)
        rsi_threshold = st.number_input("RSI Threshold", min_value=10.0, max_value=50.0, value=35.0, step=1.0)
        ema_long = st.number_input("EMA Long (trend filter)", min_value=50, max_value=400, value=200, step=1,
                                   help="Price must be BELOW this to arm tracking, and ABOVE it to trigger the buy.")
        ema_mid = st.number_input("EMA Mid", min_value=5, max_value=200, value=50, step=1,
                                  help="Buy needs EMA(Mid) > EMA(Short).")
        ema_short = st.number_input("EMA Short", min_value=2, max_value=100, value=21, step=1)
        profit_target = st.number_input("Profit Target (%)", min_value=0.5, max_value=20.0, value=3.14, step=0.01)

        if not (ema_short < ema_mid < ema_long):
            st.warning("Tip: this strategy assumes EMA Short < EMA Mid < EMA Long "
                       "(e.g. 21 < 50 < 200). Your current values don't follow that "
                       "ordering — signals will still compute, but may behave oddly.")

        st.markdown("---")
        st.header("Risk Management Mode")
        st.caption("Compare the original Averaging strategy against an optional fixed-% Stop Loss. Only one can be active per backtest.")

        stop_loss_mode = st.toggle("Enable Stop Loss", value=False, key="sidebar_stop_loss_mode")

        if stop_loss_mode:
            st.info("Stop Loss mode disables Averaging because both risk management methods cannot be used together.")
            stop_loss_pct = st.number_input(
                "Stop Loss (%)", min_value=0.5, max_value=20.0, value=3.0, step=0.25,
                key="sidebar_stop_loss_pct"
            )
        else:
            stop_loss_pct = 3.0

        average_trigger = st.number_input(
            "Average Trigger (%)", min_value=0.5, max_value=10.0, value=2.0, step=0.5,
            disabled=stop_loss_mode,
            help="Next-day execution price must be at least this much below the previous execution price to average."
        )
        max_averages = st.number_input(
            "Max Average Entries", min_value=0, max_value=20, value=5, step=1,
            disabled=stop_loss_mode
        )

        st.markdown("---")

        st.header("Capital & Position")

        initial_capital = st.number_input("Initial Capital (₹)", min_value=100000, max_value=100000000, value=1000000, step=50000)
        position_size = st.number_input("Position Size (₹)", min_value=10000, max_value=500000, value=50000, step=5000)
        max_positions = st.number_input("Max Open Positions", min_value=1, max_value=100, value=20, step=1)

        st.markdown("---")

        st.header("Transaction Costs")

        brokerage = st.number_input("Brokerage (%)", min_value=0.0, max_value=1.0, value=0.05, step=0.01)
        slippage = st.number_input("Slippage (%)", min_value=0.0, max_value=1.0, value=0.05, step=0.01)
        transaction_charges = st.number_input("Transaction Charges (%)", min_value=0.0, max_value=1.0, value=0.01, step=0.01)

        st.markdown("---")

        st.header("Data")

        period_options = list(StrategyConfig.VALID_PERIODS)
        data_period = st.selectbox(
            "Backtest / Scan Period",
            options=period_options,
            index=period_options.index("2y"),
            help="How much historical daily data to download per stock. "
                 "Must be long enough to cover the EMA Long warmup "
                 "(e.g. 200 bars for EMA 200) plus room for actual trades."
        )

        min_required_bars = max(ema_long, ema_mid, ema_short) + 10
        approx_bars = StrategyConfig.PERIOD_TRADING_DAYS.get(data_period, 10_000)
        if approx_bars < min_required_bars + 60:
            st.warning(
                f"⚠️ '{data_period}' (~{approx_bars} trading days) may be too short for "
                f"EMA Long={ema_long} "
                f"(needs ≥{min_required_bars} bars for warmup, plus room to trade). "
                f"Consider a longer period (2y+ is recommended for EMA 200)."
            )

        if st.button("Clear Cache"):
            st.session_state.data_manager.clear_cache()
            st.success("Cache cleared!")

        st.markdown("---")
        st.header("Stock Universe")

        universe_mode = st.radio(
            "Symbol Source",
            options=["NIFTY 200", "NIFTY 500", "Custom Symbol"],
            index=0,
            help="Choose which universe to scan/backtest. NIFTY 200 and "
                 "NIFTY 500 are curated built-in lists; 'Custom Symbol' lets "
                 "you upload your own Excel/CSV of tickers."
        )

        if universe_mode == "NIFTY 200":
            st.session_state.active_symbols = NIFTY200_SYMBOLS
            st.session_state.universe_source_label = (
                f"NIFTY 200 ({len(NIFTY200_SYMBOLS)} symbols)"
            )
            st.caption(
                "Curated NIFTY 50 + NIFTY Next 50 + additional large/mid-caps, "
                "cross-checked against the NIFTY 500 list. NSE rebalances "
                "semi-annually, so this can drift slightly from the live "
                "official constituents."
            )

        elif universe_mode == "NIFTY 500":
            st.session_state.active_symbols = NIFTY500_SYMBOLS
            st.session_state.universe_source_label = (
                f"NIFTY 500 ({len(NIFTY500_SYMBOLS)} symbols)"
            )
            st.caption(
                "Sourced from NSE-published constituent data (snapshot dated "
                "26 Nov 2024). Same semi-annual-rebalance caveat applies."
            )

        else:  # Custom Symbol
            uploaded_symbol_file = st.file_uploader(
                "Upload symbol file",
                type=["xlsx", "xls", "csv"],
                help="File needs one column of stock symbols, e.g. header "
                     "'Symbol' (also accepts 'Ticker', 'Scrip', 'Stock'). "
                     "'.NS' is appended automatically if missing."
            )
            if uploaded_symbol_file is not None:
                try:
                    custom_symbols = parse_symbol_file(uploaded_symbol_file)
                except Exception as e:
                    custom_symbols = []
                    # Keep the user message friendly, but preserve the full
                    # traceback in the logs so malformed-file issues are
                    # debuggable instead of being discarded.
                    logger.exception(
                        "parse_symbol_file failed for upload %r",
                        getattr(uploaded_symbol_file, 'name', '<unknown>'),
                    )
                    st.error(f"Couldn't read that file: {e}")

                if custom_symbols:
                    st.session_state.active_symbols = custom_symbols
                    st.session_state.universe_source_label = (
                        f"Custom upload ({len(custom_symbols)} symbols)"
                    )
                    st.success(f"Loaded {len(custom_symbols)} symbols from "
                               f"'{uploaded_symbol_file.name}'")
                    with st.expander("Preview symbols"):
                        st.write(custom_symbols)
                else:
                    st.error(
                        "No valid symbols found in that file. Add a column "
                        "named 'Symbol' (or 'Ticker'/'Scrip'/'Stock') with "
                        "one ticker per row."
                    )
                    st.session_state.active_symbols = NIFTY200_SYMBOLS
                    st.session_state.universe_source_label = (
                        f"NIFTY 200 ({len(NIFTY200_SYMBOLS)} symbols) — fallback"
                    )
            else:
                st.info("Upload a file to use a custom stock universe.")
                st.session_state.active_symbols = NIFTY200_SYMBOLS
                st.session_state.universe_source_label = (
                    f"NIFTY 200 ({len(NIFTY200_SYMBOLS)} symbols) — fallback"
                )

        st.caption(f"Active universe: **{st.session_state.universe_source_label}**")

        config = StrategyConfig(
            rsi_period=rsi_period,
            rsi_threshold=rsi_threshold,
            ema_long=ema_long,
            ema_mid=ema_mid,
            ema_short=ema_short,
            profit_target_pct=profit_target,
            average_trigger_pct=average_trigger,
            max_average_entries=max_averages,
            position_size=position_size,
            initial_capital=initial_capital,
            max_positions=max_positions,
            brokerage_pct=brokerage,
            slippage_pct=slippage,
            transaction_charges_pct=transaction_charges,
            data_period=data_period,
            stop_loss_mode=stop_loss_mode,
            stop_loss_pct=stop_loss_pct,
        )

        st.session_state.config = config
        return config


def render_dashboard(metrics: Optional[Dict[str, Any]]):
    render_page_header("Dashboard", "Backtest performance overview", "bar-chart-2")

    if metrics is None:
        st.info("Run a backtest to see metrics")
        return

    # --- Closed Performance (realized, closed trades only) ---
    render_section_title("Closed Performance", "check-circle", badge_text="Realized", badge_class="badge-green")

    # KPI Cards — row 1
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        render_kpi_card("Closed Trades", f"{metrics['total_trades']}", "Executed round-trips",
                         icon_name="layers", accent="blue")
    with col2:
        render_kpi_card("Win Rate", f"{metrics['win_rate']:.1f}%", f"{metrics['winning_trades']} winners",
                         icon_name="target", accent="purple")
    with col3:
        net_positive = metrics['net_profit'] >= 0
        render_kpi_card("Realized Profit", f"₹{metrics['net_profit']:,.0f}", "Realized across closed trades",
                         icon_name="dollar-sign", accent="green" if net_positive else "red",
                         trend=f"{metrics['total_return_pct']:+.2f}%", trend_positive=net_positive)
    with col4:
        render_kpi_card("Total Return", f"{metrics['total_return_pct']:.2f}%", "On initial capital",
                         icon_name="trending-up" if metrics['total_return_pct'] >= 0 else "trending-down",
                         accent="green" if metrics['total_return_pct'] >= 0 else "red")

    # KPI Cards — row 2
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        render_kpi_card("Max Drawdown", f"{metrics['max_drawdown_pct']:.2f}%", "Peak-to-trough",
                         icon_name="trending-down", accent="red")
    with col2:
        render_kpi_card("Sharpe Ratio", f"{metrics['sharpe_ratio']:.2f}", "Risk-adjusted return",
                         icon_name="activity", accent="blue")
    with col3:
        render_kpi_card("Profit Factor", f"{metrics['profit_factor']:.2f}", "Gross win / gross loss",
                         icon_name="percent", accent="purple")
    with col4:
        render_kpi_card("Avg Holding", f"{metrics['avg_holding_days']:.1f}d", "Per trade",
                         icon_name="clock", accent="neutral")

    # --- Open Positions (unrealized, mark-to-market) ---
    unrealized = st.session_state.get('unrealized_summary')
    if unrealized:
        render_section_title("Open Positions", "briefcase", badge_text="Unrealized · Mark-to-Market", badge_class="badge-blue")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            render_kpi_card("Open Positions", f"{unrealized['open_position_count']}", "Still held (not force-closed)",
                             icon_name="layers", accent="neutral")
        with col2:
            pnl_positive = unrealized['total_unrealized_pnl_inr'] >= 0
            render_kpi_card("Unrealized P&L", f"₹{unrealized['total_unrealized_pnl_inr']:,.0f}",
                             "Mark-to-market, not yet realized",
                             icon_name="dollar-sign", accent="green" if pnl_positive else "red",
                             trend=f"{unrealized['unrealized_return_pct']:+.2f}%", trend_positive=pnl_positive)
        with col3:
            render_kpi_card("Unrealized Return", f"{unrealized['unrealized_return_pct']:.2f}%",
                             "On invested capital (open positions)",
                             icon_name="trending-up" if unrealized['unrealized_return_pct'] >= 0 else "trending-down",
                             accent="green" if unrealized['unrealized_return_pct'] >= 0 else "red")
        with col4:
            render_kpi_card("Current Portfolio Value", f"₹{unrealized['current_market_value']:,.0f}",
                             "Market value of open positions",
                             icon_name="briefcase", accent="blue")

    if metrics.get('equity_curve'):
        render_section_title("Equity Curve", "trending-up")
        dates = [x[0] for x in metrics['equity_curve']]
        equity = [x[1] for x in metrics['equity_curve']]

        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=dates, y=equity, mode='lines', name='Equity',
            line=dict(color='#8B5CF6', width=2),
            fill='tonexty', fillcolor='rgba(139,92,246,0.12)'
        ))
        apply_plotly_theme(fig, height=380)
        fig.update_layout(xaxis_title='Date', yaxis_title='Portfolio Value (₹)', showlegend=False)
        render_chart_card(fig, key="dashboard_equity_chart")

    if metrics.get('monthly_returns'):
        render_section_title("Monthly Returns", "calendar")
        monthly_rows = [{'Month': m, 'profit_raw': p, 'Profit': f"₹{p:,.0f}"}
                         for m, p in sorted(metrics['monthly_returns'].items())]
        render_premium_table(
            monthly_rows,
            columns=[
                {'key': 'Month', 'label': 'Month'},
                {'key': 'Profit', 'label': 'Profit', 'kind': 'pnl', 'raw_key': 'profit_raw'},
            ],
            height=260,
        )


def _scanner_portfolio_badge(in_portfolio: bool) -> Tuple[str, str]:
    if in_portfolio:
        return ("🟢 In Portfolio", "badge-green")
    return ("🔵 Buy Signal", "badge-blue")


def _sl_mode_badge(enabled: bool) -> Tuple[str, str]:
    return ("🟡 Yes", "badge-yellow") if enabled else ("⚪ No", "badge-gray")


def render_scanner():
    render_page_header("Scanner", "Live buy-signal scan across the selected universe", "zap")
    config = st.session_state.config

    col1, col2 = st.columns([1, 3])
    with col1:
        scan_button = st.button("▶ Run Scanner", type="primary", use_container_width=True)

    active_symbols = st.session_state.get('active_symbols', NIFTY200_SYMBOLS)
    st.caption(f"Universe: **{st.session_state.get('universe_source_label', f'{len(active_symbols)} symbols')}**")

    if scan_button:
        with st.spinner(f"Scanning {len(active_symbols)} symbols..."):
            scanner = Scanner(config)
            progress_bar = st.progress(0)
            status_box = st.empty()
            data_manager = st.session_state.data_manager

            def update_progress(pct):
                progress_bar.progress(min(pct, 1.0))

            def update_status(msg):
                status_box.caption(f"📡 {msg}")

            results = scanner.scan_universe(
                active_symbols,
                data_manager,
                max_workers=8,
                progress_callback=update_progress,
                status_callback=update_status,
            )

            st.session_state.scanner_results = results
            st.session_state.scan_complete = True
            st.session_state['scanner_page'] = 1
            progress_bar.empty()
            status_box.empty()

    # Inline "Add Position" form for Streamlit versions without st.dialog
    # support (see open_add_position_dialog in the Portfolio module).
    if not hasattr(st, "dialog") and st.session_state.get('pf_inline_add_ctx'):
        with st.container(border=True):
            st.markdown("#### ➕ Add Position to Portfolio")
            _render_add_position_form(st.session_state['pf_inline_add_ctx'])

    if st.session_state.scan_complete and st.session_state.scanner_results:
        results = st.session_state.scanner_results

        k1, k2, k3 = st.columns(3)
        with k1:
            render_kpi_card("Buy Signals Found", f"{len(results)}", icon_name="zap", accent="purple")
        with k2:
            render_kpi_card("Universe Scanned", f"{len(active_symbols)}", icon_name="layers", accent="blue")
        with k3:
            render_kpi_card("Avg RSI", f"{(sum(r['rsi'] for r in results) / len(results)):.1f}" if results else "—",
                             icon_name="activity", accent="neutral")

        # Cheap single query so the table + action row below don't hit
        # SQLite once per row.
        open_positions_by_symbol = {p['symbol']: p for p in pf_fetch_positions(status='OPEN')}
        avg_gap = st.session_state.get('portfolio_default_gap_pct', DEFAULT_AVERAGE_GAP_PCT)

        sl_enabled = bool(config.stop_loss_mode)

        scan_rows: List[Dict[str, Any]] = []
        for r in results:
            entry_price = r['close']
            avg_levels = compute_average_levels(entry_price, avg_gap, DEFAULT_AVERAGE_LEVELS)
            target_price = round(entry_price * (1 + config.profit_target_pct / 100), 2)
            symbol = r['symbol']
            pos = open_positions_by_symbol.get(symbol)
            in_portfolio = pos is not None
            avg_count = None
            return_pct = None
            if pos is not None:
                m = pf_position_metrics(pos)
                avg_count = m['average_count']
                if m['average_cost'] > 0:
                    return_pct = (entry_price - m['average_cost']) / m['average_cost'] * 100

            if in_portfolio and pos.get('stop_loss'):
                # Show the position's actual recorded Stop Loss, not a
                # fresh estimate off today's price.
                stop_loss = float(pos['stop_loss'])
            elif sl_enabled:
                # Fixed-% Stop Loss module (Strategy Settings).
                stop_loss = round(entry_price * (1 - config.stop_loss_pct / 100), 2)
            else:
                # Averaging mode default suggestion: just below the last
                # averaging level.
                stop_loss = round(avg_levels[-1] * 0.98, 2)

            scan_rows.append({
                'symbol': symbol, 'symbol_short': symbol.replace('.NS', ''),
                'company': r.get('company', symbol), 'sector': r.get('sector', 'Unknown'),
                'close': entry_price, 'rsi': r['rsi'],
                'ema_short': r.get('ema_short', entry_price),
                'ema_mid': r.get('ema_mid', entry_price),
                'ema_long': r.get('ema_long', entry_price),
                'volume': r['volume'],
                'entry_price': entry_price, 'stop_loss': stop_loss, 'target_price': target_price,
                'avg1': avg_levels[0], 'avg2': avg_levels[1], 'avg3': avg_levels[2],
                'avg4': avg_levels[3], 'avg5': avg_levels[4],
                'avg_count': avg_count, 'return_pct': return_pct,
                'in_portfolio': in_portfolio,
                'sl_enabled': sl_enabled, 'sl_pct': config.stop_loss_pct if sl_enabled else None,
                'company_ctx': r.get('company', symbol), 'signal_date': r.get('date'),
            })

        render_section_title("Buy Signals", "list")
        sectors = sorted({r['sector'] for r in scan_rows})
        page_rows, filtered_rows = render_table_toolbar(
            scan_rows,
            search_keys=['symbol_short', 'company', 'sector'],
            sort_options={
                'RSI': 'rsi', 'EMA 200': 'ema_long', 'Close': 'close',
                'Volume': 'volume', 'Return %': 'return_pct', 'Symbol': 'symbol_short',
            },
            state_prefix="scanner",
            filter_options={'sector': sectors} if sectors else None,
            default_page_size=25,
        )

        table_columns = [
            {'key': 'symbol_short', 'label': 'Symbol', 'fmt': lambda v: f"<b>{v}</b>"},
            {'key': 'company', 'label': 'Company'},
            {'key': 'sector', 'label': 'Sector'},
            {'key': 'close', 'label': 'Close', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'rsi', 'label': 'RSI', 'fmt': lambda v: f"{v:.1f}"},
            {'key': 'ema_short', 'label': 'EMA 21', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'ema_mid', 'label': 'EMA 50', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'ema_long', 'label': 'EMA 200', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'entry_price', 'label': 'Entry', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'stop_loss', 'label': 'Stop Loss', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'target_price', 'label': 'Target', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'avg1', 'label': 'Avg 1', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'avg2', 'label': 'Avg 2', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'avg3', 'label': 'Avg 3', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'avg4', 'label': 'Avg 4', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'avg5', 'label': 'Avg 5', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'volume', 'label': 'Volume', 'fmt': lambda v: f"{v:,.0f}"},
            {'key': 'avg_count', 'label': 'Avg Count', 'fmt': lambda v: str(int(v))},
            {'key': 'return_pct', 'label': 'Return %', 'fmt': lambda v: f"{v:+.2f}%", 'kind': 'pnl'},
            {'key': 'sl_enabled', 'label': 'SL Mode', 'kind': 'badge', 'badge_fn': _sl_mode_badge},
            {'key': 'sl_pct', 'label': 'SL %', 'fmt': lambda v: f"{v:.2f}%"},
            {'key': 'in_portfolio', 'label': 'Portfolio', 'kind': 'badge', 'badge_fn': _scanner_portfolio_badge},
        ]
        render_premium_table(page_rows, table_columns, height=440)

        if st.button("⬇ Export Scanner CSV"):
            csv = Exporter.scanner_to_csv(st.session_state.scanner_results)
            st.download_button("Download CSV", csv, "scanner_results.csv", "text/csv")

        render_section_title("Add to Portfolio", "plus-circle")
        st.caption("Every row is a live BUY signal. Add any of them to your tracked Portfolio.")

        header_cols = st.columns([1.4, 1, 1, 1, 1.3])
        for label, c in zip(["Symbol", "Entry", "Stop Loss", "Target", "Action"], header_cols):
            c.markdown(f"<span style='color:var(--text-secondary);font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.04em;'>{label}</span>", unsafe_allow_html=True)

        for meta in page_rows:
            symbol = meta['symbol']
            c1, c2, c3, c4, c5 = st.columns([1.4, 1, 1, 1, 1.3])
            c1.markdown(f"**{meta['symbol_short']}**")
            c2.write(f"₹{meta['entry_price']:.2f}")
            c3.write(f"₹{meta['stop_loss']:.2f}")
            c4.write(f"₹{meta['target_price']:.2f}")
            with c5:
                if meta['in_portfolio']:
                    st.button("✓ In Portfolio", key=f"added_{symbol}", disabled=True, use_container_width=True)
                else:
                    if st.button("➕ Add to Portfolio", key=f"add_{symbol}", use_container_width=True):
                        open_add_position_dialog({
                            'symbol': symbol,
                            'company': meta['company_ctx'],
                            'close': meta['entry_price'],
                            'stop_loss': meta['stop_loss'],
                            'target_price': meta['target_price'],
                            'avg_levels': [meta['avg1'], meta['avg2'], meta['avg3'], meta['avg4'], meta['avg5']],
                            'signal_date': meta['signal_date'],
                        })
    else:
        st.info("Click 'Run Scanner' to find stocks matching buy criteria")


def render_backtest():
    render_page_header("Backtest", "Portfolio-level strategy simulation with realistic T+1 execution", "clock")
    config = st.session_state.config

    mode_label = "🛑 Stop Loss Mode" if config.stop_loss_mode else "📊 Averaging Mode"
    mode_badge_class = "badge-yellow" if config.stop_loss_mode else "badge-purple"
    st.markdown(f"<span class='badge {mode_badge_class}'>{mode_label}</span>", unsafe_allow_html=True)

    st.caption(f"Data period: **{config.data_period}** · "
               f"Warmup required: **{config.min_required_bars} bars** "
               f"(EMA {config.ema_long} / {config.ema_mid} / {config.ema_short})")
    if config.period_likely_insufficient:
        st.warning(
            f"⚠️ The selected data period ('{config.data_period}') may not leave enough "
            f"bars after warmup for meaningful backtest results. Increase the period in "
            f"the sidebar (2y+ recommended for EMA {config.ema_long})."
        )

    active_symbols = st.session_state.get('active_symbols', NIFTY200_SYMBOLS)
    st.caption(f"Universe: **{st.session_state.get('universe_source_label', f'{len(active_symbols)} symbols')}**")

    col1, col2 = st.columns([1, 3])
    with col1:
        if st.button("Run Backtest", type="primary", use_container_width=True):
            with st.spinner(f"Running backtest on {len(active_symbols)} symbols..."):
                data_manager = st.session_state.data_manager

                # Download all data using batch download with adaptive rate-limit backoff
                progress_bar = st.progress(0)
                status_box = st.empty()

                def update_dl_progress(pct):
                    progress_bar.progress(min(pct, 1.0))

                def update_dl_status(msg):
                    status_box.caption(f"📡 {msg}")

                data_dict = data_manager.batch_download(
                    active_symbols,
                    period=config.data_period,
                    interval="1d",
                    max_workers=8,
                    progress_callback=update_dl_progress,
                    status_callback=update_dl_status,
                )

                progress_bar.empty()
                status_box.empty()

                # Share the freshly downloaded OHLC with the Buy Signal
                # Analysis page so it doesn't have to re-download.
                st.session_state['signal_universe_data'] = {
                    'data': data_dict, 'period': config.data_period
                }

                # Run optimized portfolio-level backtest
                engine = BacktestEngine(config)
                all_trades, daily_equity = engine.run_backtest(data_dict)
                st.session_state.backtest_diagnostics = dict(engine.last_diagnostics)
                st.session_state.backtest_diagnostics['symbols_downloaded'] = len(data_dict)
                st.session_state.backtest_diagnostics['symbols_requested'] = len(active_symbols)

                # Calculate metrics (daily_invested is additive - powers the
                # new Capital Utilization metric, doesn't change anything else)
                # NOTE: purely Realized Performance - computed only from
                # completed_trades (closed positions). Positions still open
                # when the data ends are NEVER included here.
                metrics = MetricsCalculator.calculate_metrics(
                    all_trades, daily_equity, config.initial_capital,
                    daily_invested=engine.last_daily_invested,
                )

                # NEW: Unrealized Performance - open positions still held
                # when the backtest data ended. Kept fully separate from
                # `metrics` above (never mixed into Realized stats).
                st.session_state.open_positions_snapshot = list(engine.last_open_positions)
                st.session_state.unrealized_summary = MetricsCalculator.summarize_open_positions(
                    engine.last_open_positions
                )

                # Store per-symbol data for charts
                backtest_data = {}
                for symbol, df in data_dict.items():
                    symbol_trades = [t for t in all_trades if t.symbol == symbol]
                    if symbol_trades:
                        backtest_data[symbol] = {'data': df, 'trades': symbol_trades}

                st.session_state.all_trades = all_trades
                st.session_state.backtest_results = backtest_data
                st.session_state.metrics = metrics
                st.session_state.daily_equity = daily_equity
                st.session_state.backtest_complete = True

                # Keep a copy per Risk Management Mode so Averaging vs Stop
                # Loss can be compared side by side (see Mode Comparison
                # section below) without needing to re-run both at once.
                mode_key = "Stop Loss Mode" if config.stop_loss_mode else "Averaging Mode"
                st.session_state.setdefault('metrics_by_mode', {})[mode_key] = metrics

    # Diagnostics panel: explains WHY a run produced few/zero trades instead
    # of failing silently.
    diag = st.session_state.get('backtest_diagnostics')
    if diag:
        downloaded = diag.get('symbols_downloaded', 0)
        requested = diag.get('symbols_requested', len(NIFTY200_SYMBOLS))
        prepared = diag.get('symbols_prepared', 0)
        buy_signals = diag.get('total_buy_signals', 0)
        executed = diag.get('buy_orders_executed', 0)

        with st.expander("🔎 Backtest diagnostics", expanded=(buy_signals == 0)):
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Symbols downloaded", f"{downloaded}/{requested}")
            c2.metric("Enough history", prepared)
            c3.metric("Buy-signal bars", buy_signals)
            c4.metric("Orders executed", executed)

            if downloaded == 0:
                st.error(
                    "No data was downloaded. This is almost always a network / "
                    "yfinance rate-limit issue (300 symbols at once can get "
                    "throttled). Try again, reduce the universe, or clear the "
                    "cache and retry."
                )
            elif prepared == 0:
                bars = diag.get('bar_counts') or [0]
                st.error(
                    f"Data downloaded but symbols averaged only "
                    f"~{int(np.median(bars))} bars (min {min(bars)}, max {max(bars)}); "
                    f"the strategy needs {config.min_required_bars} (warmup for "
                    f"EMA {config.ema_long}). "
                    f"If your period is 1y+ this means Yahoo returned throttled "
                    f"partial data. Click **Clear Cache** in the sidebar, wait a "
                    f"minute, then run again - the downloader will re-fetch in "
                    f"bulk and will no longer cache truncated responses."
                )
            elif buy_signals == 0:
                st.warning(
                    "Data is fine, but the strategy conditions never aligned. "
                    f"Arming needs RSI(14) < {config.rsi_threshold:.0f} while "
                    f"Close is below EMA {config.ema_long}; the buy then needs "
                    f"the early-reclaim alignment "
                    f"**Close > EMA{config.ema_long} > EMA{config.ema_mid} > "
                    f"EMA{config.ema_short}** on a single candle. That window is "
                    "deliberately narrow (price reclaims the long EMA before the "
                    "faster EMAs cross above it), so zero signals over a short "
                    "period is common. Try a longer data period or a larger "
                    "universe."
                )
            elif executed == 0:
                st.warning(
                    "Signals fired but no orders filled - check that "
                    "Position Size fits within Initial Capital and Max Positions."
                )

    if st.session_state.backtest_complete and st.session_state.metrics:
        metrics = st.session_state.metrics

        render_section_title("Realized Performance", "bar-chart-2", badge_text="Realized", badge_class="badge-green")
        st.caption("Computed purely from closed trades. Open positions are never included here.")
        metrics_rows = MetricsCalculator.metrics_to_dataframe(metrics).to_dict('records')
        render_premium_table(
            metrics_rows,
            columns=[{'key': 'Metric', 'label': 'Metric'}, {'key': 'Value', 'label': 'Value'}],
            height=480,
        )

        # --- Open Positions (unrealized, never force-closed) ---
        open_positions = st.session_state.get('open_positions_snapshot', [])
        unrealized = st.session_state.get('unrealized_summary')
        render_section_title("Open Positions", "briefcase", badge_text="Unrealized · Mark-to-Market", badge_class="badge-blue")
        if open_positions:
            st.caption(
                f"{len(open_positions)} position(s) still open when the data ended. "
                "They are NOT closed, NOT counted as trades, and NOT included in the "
                "Realized Performance table above - only marked to market."
            )
            if unrealized:
                oc1, oc2, oc3, oc4 = st.columns(4)
                with oc1:
                    render_kpi_card("Open Positions", f"{unrealized['open_position_count']}",
                                     icon_name="layers", accent="neutral")
                with oc2:
                    pnl_positive = unrealized['total_unrealized_pnl_inr'] >= 0
                    render_kpi_card("Unrealized P&L", f"₹{unrealized['total_unrealized_pnl_inr']:,.0f}",
                                     icon_name="dollar-sign", accent="green" if pnl_positive else "red",
                                     trend=f"{unrealized['unrealized_return_pct']:+.2f}%", trend_positive=pnl_positive)
                with oc3:
                    render_kpi_card("Current Market Value", f"₹{unrealized['current_market_value']:,.0f}",
                                     icon_name="briefcase", accent="blue")
                with oc4:
                    render_kpi_card("Invested Capital", f"₹{unrealized['invested_capital']:,.0f}",
                                     icon_name="dollar-sign", accent="purple")

            open_rows = MetricsCalculator.open_positions_to_dataframe(open_positions).to_dict('records')
            open_columns = [
                {'key': 'Symbol', 'label': 'Symbol'},
                {'key': 'Entry Date', 'label': 'Entry Date'},
                {'key': 'Entry Price', 'label': 'Entry Price'},
                {'key': 'Avg Entry Price', 'label': 'Avg Entry Price'},
                {'key': 'Entries', 'label': 'Entries'},
                {'key': 'Current Price', 'label': 'Current Price'},
                {'key': 'Quantity', 'label': 'Quantity'},
                {'key': 'Invested (₹)', 'label': 'Invested (₹)'},
                {'key': 'Market Value (₹)', 'label': 'Market Value (₹)'},
                {'key': 'Unrealized P&L (₹)', 'label': 'Unrealized P&L (₹)'},
                {'key': 'Unrealized P&L (%)', 'label': 'Unrealized P&L (%)'},
                {'key': 'Drawdown (%)', 'label': 'Drawdown (%)'},
                {'key': 'Target Price', 'label': 'Target Price'},
                {'key': 'Stop Loss Price', 'label': 'Stop Loss Price'},
                {'key': 'Days Held', 'label': 'Days Held'},
                {'key': 'Status', 'label': 'Status'},
            ]
            render_premium_table(open_rows, open_columns, height=420)
        else:
            st.info("No open positions remained when the backtest data ended - every position was closed via a real exit.")

        metrics_by_mode = st.session_state.get('metrics_by_mode', {})
        if len(metrics_by_mode) >= 2:
            render_section_title("Mode Comparison — Averaging vs Stop Loss", "sliders")
            st.caption("Run a backtest in each mode (toggle Stop Loss in the sidebar) to keep both results here for direct comparison.")

            compare_fields = [
                ('CAGR (%)', 'cagr', '{:.2f}'),
                ('XIRR (%)', 'xirr_pct', '{:.2f}'),
                ('Total Return (%)', 'total_return_pct', '{:.2f}'),
                ('Win Rate (%)', 'win_rate', '{:.2f}'),
                ('Max Drawdown (%)', 'max_drawdown_pct', '{:.2f}'),
                ('Profit Factor', 'profit_factor', '{:.2f}'),
                ('Avg Holding Days', 'avg_holding_days', '{:.1f}'),
                ('Number of Trades', 'total_trades', '{:.0f}'),
                ('Capital Utilization (%)', 'capital_utilization_pct', '{:.2f}'),
                ('Sharpe Ratio', 'sharpe_ratio', '{:.2f}'),
                ('Calmar Ratio', 'calmar_ratio', '{:.2f}'),
            ]
            mode_names = list(metrics_by_mode.keys())
            compare_rows = []
            for label, key, fmt in compare_fields:
                row = {'Metric': label}
                for mn in mode_names:
                    row[mn] = fmt.format(metrics_by_mode[mn].get(key, 0) or 0)
                compare_rows.append(row)

            compare_columns = [{'key': 'Metric', 'label': 'Metric'}] + [
                {'key': mn, 'label': mn} for mn in mode_names
            ]
            render_premium_table(compare_rows, compare_columns, height=420)

        col1, col2 = st.columns(2)
        with col1:
            if st.button("⬇ Export Trades CSV"):
                csv = Exporter.trades_to_csv(st.session_state.all_trades)
                st.download_button("Download Trades CSV", csv, "trades.csv", "text/csv")

        with col2:
            if st.button("⬇ Export Excel Report"):
                if OPENPYXL_AVAILABLE:
                    excel = Exporter.create_excel_report(
                        metrics,
                        st.session_state.all_trades,
                        st.session_state.scanner_results
                    )
                    st.download_button(
                        "Download Excel Report", excel, "backtest_report.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("openpyxl not available")
    else:
        st.info("Click 'Run Backtest' to start backtesting")


def render_trade_log():
    render_page_header("Trade Log", "Every simulated trade from the last backtest run", "list")

    if not st.session_state.all_trades:
        st.info("Run a backtest to see trades")
        return

    trades = st.session_state.all_trades

    col1, col2, col3 = st.columns(3)
    with col1:
        symbols = list(set(t.symbol for t in trades))
        selected_symbol = st.selectbox("Filter by Symbol", ["All"] + sorted(symbols))
    with col2:
        status_filter = st.selectbox("Status", ["All", "CLOSED"])
    with col3:
        min_profit = st.number_input("Min Profit %", value=-100.0, step=1.0)

    filtered = trades
    if selected_symbol != "All":
        filtered = [t for t in filtered if t.symbol == selected_symbol]
    if status_filter != "All":
        filtered = [t for t in filtered if t.status == status_filter]
    filtered = [t for t in filtered if t.profit_pct >= min_profit]

    trade_rows = []
    for t in filtered:
        trade_rows.append({
            'symbol': t.symbol,
            'entry_date': t.entry_dates[0].strftime('%Y-%m-%d') if t.entry_dates else '',
            'entry_price': t.entry_prices[0] if t.entry_prices else 0,
            'averages': t.average_count,
            'exit_date': t.exit_date.strftime('%Y-%m-%d') if t.exit_date else '',
            'exit_price': t.exit_price if t.exit_price else 0,
            'holding_days': t.holding_days,
            'profit_pct': t.profit_pct,
            'profit_inr': t.profit_inr,
            'invested': t.total_invested,
            'returned': t.total_returned,
            'reason': t.exit_reason,
        })

    render_premium_table(
        trade_rows,
        columns=[
            {'key': 'symbol', 'label': 'Symbol', 'fmt': lambda v: f"<b>{v}</b>"},
            {'key': 'entry_date', 'label': 'Entry Date'},
            {'key': 'entry_price', 'label': 'Entry Price', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'averages', 'label': 'Averages'},
            {'key': 'exit_date', 'label': 'Exit Date'},
            {'key': 'exit_price', 'label': 'Exit Price', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'holding_days', 'label': 'Holding Days', 'fmt': lambda v: f"{v}d"},
            {'key': 'profit_pct', 'label': 'Profit %', 'fmt': lambda v: f"{v:+.2f}%", 'kind': 'pnl'},
            {'key': 'profit_inr', 'label': 'Profit ₹', 'fmt': lambda v: f"₹{v:,.2f}", 'kind': 'pnl'},
            {'key': 'invested', 'label': 'Invested', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'returned', 'label': 'Returned', 'fmt': lambda v: f"₹{v:,.2f}"},
            {'key': 'reason', 'label': 'Reason'},
        ],
        height=460,
    )
    st.caption(f"Showing {len(filtered)} of {len(trades)} trades")


def render_charts():
    render_page_header("Charts", "Per-symbol price/RSI/volume chart with entry, average and exit markers", "activity")

    if not st.session_state.backtest_results:
        st.info("Run a backtest to see charts")
        return

    symbols = list(st.session_state.backtest_results.keys())
    if not symbols:
        st.warning("No backtest data available")
        return

    selected = st.selectbox("Select Symbol", symbols)

    if selected:
        data = st.session_state.backtest_results[selected]
        df = data['data']
        trades = data['trades']

        fig = ChartBuilder.create_trade_chart(df, trades, selected, st.session_state.config)
        render_chart_card(fig, key=f"trade_chart_{selected}")

        render_section_title("Trade Details", "list")
        for i, trade in enumerate(trades, 1):
            with st.expander(f"Trade #{i} - {trade.symbol}"):
                # Per-entry Signal vs Execution breakdown (each buy, including averages)
                entries_rows = [
                    {
                        'signal_date': sd.strftime('%Y-%m-%d'), 'signal_close': sc,
                        'execution_date': ed.strftime('%Y-%m-%d'), 'buy_price': ep,
                    }
                    for sd, sc, ed, ep in zip(trade.signal_dates, trade.signal_closes,
                                               trade.entry_dates, trade.entry_prices)
                ]
                st.write("**Entries (Signal → Next-Day-Open Execution):**")
                render_premium_table(
                    entries_rows,
                    columns=[
                        {'key': 'signal_date', 'label': 'Signal Date'},
                        {'key': 'signal_close', 'label': 'Signal Close', 'fmt': lambda v: f"₹{v:,.2f}"},
                        {'key': 'execution_date', 'label': 'Execution Date'},
                        {'key': 'buy_price', 'label': 'Buy Price (Next Open)', 'fmt': lambda v: f"₹{v:,.2f}"},
                    ],
                    height=220,
                )

                st.write(f"**Average Count:** {trade.average_count}")
                st.write(f"**Exit Date:** {trade.exit_date.strftime('%Y-%m-%d') if trade.exit_date else 'N/A'}")
                st.write(f"**Exit Price:** ₹{trade.exit_price:.2f}" if trade.exit_price else "N/A")
                st.write(f"**Profit:** {trade.profit_pct:.2f}% (₹{trade.profit_inr:,.2f})")
                st.write(f"**Invested:** ₹{trade.total_invested:,.2f}")
                st.write(f"**Returned:** ₹{trade.total_returned:,.2f}")
                st.write(f"**Reason:** {trade.exit_reason}")


def render_portfolio():
    render_page_header("Portfolio", "Backtest-level equity curve and drawdown analysis", "trending-up")

    if not st.session_state.metrics or not st.session_state.metrics.get('equity_curve'):
        st.info("Run a backtest to see portfolio")
        return

    metrics = st.session_state.metrics

    render_section_title("Equity Curve", "trending-up")
    equity_df = pd.DataFrame([
        {'Date': d, 'Equity': e} for d, e in metrics['equity_curve']
    ])

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=equity_df['Date'], y=equity_df['Equity'],
        mode='lines', fill='tonexty',
        line=dict(color='#10B981', width=2),
        fillcolor='rgba(16,185,129,0.12)'
    ))
    apply_plotly_theme(fig, height=460)
    fig.update_layout(xaxis_title='Date', yaxis_title='Portfolio Value (₹)')
    render_chart_card(fig, key="portfolio_equity_chart")

    render_section_title("Drawdown Analysis", "trending-down")

    equity_values = [e for _, e in metrics['equity_curve']]
    peak = equity_values[0]
    drawdowns = []

    for eq in equity_values:
        if eq > peak:
            peak = eq
        dd = (peak - eq) / peak * 100 if peak > 0 else 0
        drawdowns.append(dd)

    dd_df = pd.DataFrame({
        'Date': [d for d, _ in metrics['equity_curve']],
        'Drawdown %': drawdowns
    })

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=dd_df['Date'], y=dd_df['Drawdown %'],
        mode='lines', fill='tozeroy',
        line=dict(color='#EF4444', width=1),
        fillcolor='rgba(239,68,68,0.16)'
    ))
    apply_plotly_theme(fig, height=380)
    fig.update_layout(xaxis_title='Date', yaxis_title='Drawdown (%)')
    render_chart_card(fig, key="portfolio_drawdown_chart")


def render_statistics():
    render_page_header("Statistics", "Monthly/yearly returns and trade distribution", "sliders")

    if not st.session_state.metrics:
        st.info("Run a backtest to see statistics")
        return

    metrics = st.session_state.metrics

    col1, col2 = st.columns(2)
    with col1:
        render_section_title("Monthly Returns", "calendar")
        if metrics.get('monthly_returns'):
            monthly_rows = [{'Month': k, 'profit_raw': v, 'Profit': f"₹{v:,.0f}"}
                             for k, v in sorted(metrics['monthly_returns'].items())]
            render_premium_table(
                monthly_rows,
                columns=[
                    {'key': 'Month', 'label': 'Month'},
                    {'key': 'Profit', 'label': 'Profit', 'kind': 'pnl', 'raw_key': 'profit_raw'},
                ],
                height=280,
            )

    with col2:
        render_section_title("Yearly Returns", "calendar")
        if metrics.get('yearly_returns'):
            yearly_rows = [{'Year': k, 'profit_raw': v, 'Profit': f"₹{v:,.0f}"}
                            for k, v in sorted(metrics['yearly_returns'].items())]
            render_premium_table(
                yearly_rows,
                columns=[
                    {'key': 'Year', 'label': 'Year'},
                    {'key': 'Profit', 'label': 'Profit', 'kind': 'pnl', 'raw_key': 'profit_raw'},
                ],
                height=280,
            )

    render_section_title("Trade Distribution", "bar-chart-2")
    if st.session_state.all_trades:
        profits = [t.profit_pct for t in st.session_state.all_trades]

        fig = go.Figure()
        fig.add_trace(go.Histogram(
            x=profits, nbinsx=20,
            marker_color='#8B5CF6', name='Profit Distribution',
            opacity=0.85
        ))
        fig.add_vline(x=0, line_dash="dash", line_color="#94A3B8", line_width=1)
        apply_plotly_theme(fig, height=380)
        fig.update_layout(xaxis_title='Profit %', yaxis_title='Frequency', showlegend=False)
        render_chart_card(fig, key="statistics_histogram_chart")


# =============================================================================
# MAIN
# =============================================================================

# =============================================================================
# BUY SIGNAL ANALYSIS  (date-range signal counting — reuses SignalGenerator)
# =============================================================================

def _ordinal(k: int) -> str:
    """1 -> '1st', 2 -> '2nd', 11 -> '11th', 23 -> '23rd'."""
    if 10 <= (k % 100) <= 20:
        suffix = 'th'
    else:
        suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(k % 10, 'th')
    return f"{k}{suffix}"


def _compute_buy_signal_rows(data_dict: Dict[str, pd.DataFrame], config: StrategyConfig,
                             start_ts: pd.Timestamp, end_ts: pd.Timestamp
                             ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], Dict[str, Any]]:
    """Run the SAME SignalGenerator.generate_signals() used by the Scanner and
    Backtest over every symbol and collect, for the range [start_ts, end_ts]
    (inclusive): one row per buy_signal candle AND one row per tracking_started
    candle.

    No new signal logic: buy_signal / tracking_started / rsi / ema_* all come
    straight from generate_signals(), so counts match the Scanner and Backtest
    exactly. Each buy is paired with its cycle's arm (the most recent
    tracking_started at/before the buy) and the symbol's immediately preceding
    buy (full history). Each tracking-started row is paired with the buy it
    eventually produced (if any).

    Returns (buy_rows, tracking_rows, stats).
    """
    start_ts = pd.Timestamp(start_ts).normalize()
    end_ts = pd.Timestamp(end_ts).normalize()

    rows: List[Dict[str, Any]] = []
    tracking_rows: List[Dict[str, Any]] = []
    stocks_scanned = 0
    earliest = None
    latest = None

    for symbol, df in data_dict.items():
        if df is None or len(df) < config.min_required_bars:
            continue
        stocks_scanned += 1

        sig = SignalGenerator.generate_signals(df, config)
        if 'date' in sig.columns:
            dt_series = pd.to_datetime(sig['date'])
        else:
            dt_series = pd.to_datetime(pd.Series(sig.index))
        # Strip timezone so comparisons against the (naive) start/end range can
        # never raise or silently drop rows on a tz mismatch.
        if getattr(dt_series.dt, 'tz', None) is not None:
            dt_series = dt_series.dt.tz_localize(None)
        dates = dt_series.dt.normalize().to_numpy()

        n = len(sig)
        if n == 0:
            continue
        is_buy = sig['buy_signal'].to_numpy()
        is_arm = (sig['tracking_started'].to_numpy() if 'tracking_started' in sig.columns
                  else np.zeros(n, dtype=bool))
        close = sig['close'].to_numpy()
        rsi = sig['rsi'].to_numpy()
        ema_s = sig['ema_short'].to_numpy()
        ema_m = sig['ema_mid'].to_numpy()
        ema_l = sig['ema_long'].to_numpy()

        d0 = pd.Timestamp(dates[0]); d1 = pd.Timestamp(dates[-1])
        earliest = d0 if earliest is None else min(earliest, d0)
        latest = d1 if latest is None else max(latest, d1)

        # Single pass: collect arms, pair each buy with its cycle arm, and
        # map each arm to the buy it eventually produced (if any).
        last_arm = None
        buys: List[Tuple[int, Optional[int]]] = []
        arms: List[int] = []
        arm_to_buy: Dict[int, int] = {}
        for i in range(n):
            if is_arm[i]:
                last_arm = i
                arms.append(i)
            if is_buy[i]:
                buys.append((i, last_arm))
                if last_arm is not None:
                    arm_to_buy[last_arm] = i

        sym_short = symbol.replace('.NS', '').replace('.BO', '')

        # ---- Buy-signal rows (date in range) ----
        for seq, (i, arm_idx) in enumerate(buys):
            bd = pd.Timestamp(dates[i])
            if not (start_ts <= bd <= end_ts):
                continue
            arm_d = pd.Timestamp(dates[arm_idx]) if arm_idx is not None else None
            days_tracking = int((bd - arm_d).days) if arm_d is not None else None
            prev_buy_d = pd.Timestamp(dates[buys[seq - 1][0]]) if seq > 0 else None
            rows.append({
                'symbol': sym_short,
                'symbol_full': symbol,
                'buy_date': bd,
                'buy_date_str': bd.strftime('%d-%b-%Y'),
                'close': float(close[i]),
                'rsi': (float(rsi[i]) if not np.isnan(rsi[i]) else None),
                'ema21': (float(ema_s[i]) if not np.isnan(ema_s[i]) else None),
                'ema50': (float(ema_m[i]) if not np.isnan(ema_m[i]) else None),
                'ema200': (float(ema_l[i]) if not np.isnan(ema_l[i]) else None),
                'tracking_start_str': (arm_d.strftime('%d-%b-%Y') if arm_d is not None else '—'),
                'days_in_tracking': days_tracking,
                'prev_buy_str': (prev_buy_d.strftime('%d-%b-%Y') if prev_buy_d is not None else '—'),
            })

        # ---- Tracking-started rows (arm date in range) ----
        for a in arms:
            ad = pd.Timestamp(dates[a])
            if not (start_ts <= ad <= end_ts):
                continue
            buy_j = arm_to_buy.get(a)
            buy_d = pd.Timestamp(dates[buy_j]) if buy_j is not None else None
            tracking_rows.append({
                'symbol': sym_short,
                'symbol_full': symbol,
                'track_date': ad,
                'track_date_str': ad.strftime('%d-%b-%Y'),
                'close': float(close[a]),
                'rsi': (float(rsi[a]) if not np.isnan(rsi[a]) else None),
                'ema21': (float(ema_s[a]) if not np.isnan(ema_s[a]) else None),
                'ema50': (float(ema_m[a]) if not np.isnan(ema_m[a]) else None),
                'ema200': (float(ema_l[a]) if not np.isnan(ema_l[a]) else None),
                'buy_date': buy_d,
                'buy_date_str': (buy_d.strftime('%d-%b-%Y') if buy_d is not None else '— (still tracking / no buy)'),
                'resulted_in_buy': buy_d is not None,
            })

    # Signal number WITHIN the selected period, per symbol, in date order.
    rows.sort(key=lambda r: (r['symbol'], r['buy_date']))
    per_symbol_counter: Dict[str, int] = {}
    for r in rows:
        c = per_symbol_counter.get(r['symbol'], 0) + 1
        per_symbol_counter[r['symbol']] = c
        r['signal_number'] = c

    tracking_rows.sort(key=lambda r: (r['symbol'], r['track_date']))
    track_counts: Dict[str, int] = {}
    for r in tracking_rows:
        track_counts[r['symbol']] = track_counts.get(r['symbol'], 0) + 1

    stats = {
        'total_signals': len(rows),
        'total_tracking_started': len(tracking_rows),
        'stocks_scanned': stocks_scanned,
        'stocks_with_signals': len(per_symbol_counter),
        'stocks_tracking': len(track_counts),
        'avg_per_stock': (len(rows) / len(per_symbol_counter)) if per_symbol_counter else 0.0,
        'sym_counts': dict(per_symbol_counter),
        'track_counts': dict(track_counts),
        'earliest': earliest,
        'latest': latest,
    }
    return rows, tracking_rows, stats


def render_buy_signal_analysis():
    render_page_header(
        "Buy Signal Analysis",
        "Count strategy BUY signals over any custom date range — identical logic to Scanner & Backtest",
        "target",
    )
    config = st.session_state.config
    store = st.session_state.get('signal_universe_data', {'data': {}, 'period': None})
    data_dict = store.get('data') or {}
    active_symbols = st.session_state.get('active_symbols', NIFTY200_SYMBOLS)

    st.caption(f"Universe: **{st.session_state.get('universe_source_label', f'{len(active_symbols)} symbols')}** · "
               f"Data period: **{config.data_period}** (set in sidebar) · "
               f"Strategy: RSI(14)&lt;{config.rsi_threshold:.0f} arm → "
               f"Close&gt;EMA{config.ema_long}&gt;EMA{config.ema_mid}&gt;EMA{config.ema_short} buy")

    # ---------------- Date range + data controls ----------------
    today = datetime.now().date()
    default_start = today - timedelta(days=180)
    c1, c2, c3 = st.columns([1, 1, 1.15])
    with c1:
        start_date = st.date_input("Start Date", value=default_start, key="bsa_start")
    with c2:
        end_date = st.date_input("End Date", value=today, key="bsa_end")
    with c3:
        if data_dict and store.get('period') == config.data_period:
            st.success(f"Reusing {len(data_dict)} downloaded symbols", icon="✅")
        elif data_dict:
            st.warning(f"Cached period '{store.get('period')}' ≠ current '{config.data_period}' — reload to match.",
                       icon="⚠️")
        else:
            st.info("No data loaded yet — click below.", icon="ℹ️")
        load = st.button("⬇️ Load / Refresh Data", use_container_width=True, key="bsa_load",
                         help="Reuses on-disk cache (24h TTL); only re-downloads symbols that aren't cached.")

    if load:
        with st.spinner(f"Loading {len(active_symbols)} symbols (cache reused when fresh)..."):
            dm = st.session_state.data_manager
            pb = st.progress(0); sb = st.empty()
            data_dict = dm.batch_download(
                active_symbols, period=config.data_period, interval="1d", max_workers=8,
                progress_callback=lambda p: pb.progress(min(p, 1.0)),
                status_callback=lambda m: sb.caption(f"📡 {m}"),
            )
            pb.empty(); sb.empty()
            st.session_state['signal_universe_data'] = {'data': data_dict, 'period': config.data_period}
            store = st.session_state['signal_universe_data']

    if start_date > end_date:
        st.error("Start Date must be on or before End Date.")
        return
    if not data_dict:
        st.info("Load data to run the analysis. It automatically reuses anything already "
                "downloaded by the Backtest tab for the same data period.")
        return

    # ---------------- Compute (reuses SignalGenerator) ----------------
    with st.spinner("Counting tracking events and buy signals across the universe..."):
        rows, tracking_rows, stats = _compute_buy_signal_rows(
            data_dict, config, pd.Timestamp(start_date), pd.Timestamp(end_date)
        )

    if stats['earliest'] is not None and pd.Timestamp(start_date) < stats['earliest']:
        st.caption(f"⚠️ Selected start {start_date:%d-%b-%Y} precedes the earliest available data "
                   f"({stats['earliest']:%d-%b-%Y}). Increase the sidebar Data Period to cover older ranges.")

    # ---------------- Summary metrics ----------------
    render_section_title("Summary", "bar-chart-2",
                         badge_text=f"{start_date:%d-%b-%Y} → {end_date:%d-%b-%Y}", badge_class="badge-purple")
    m1, m2, m3, m4, m5 = st.columns(5)
    with m1:
        render_kpi_card("Total Tracking Started", f"{stats['total_tracking_started']}", icon_name="target", accent="amber")
    with m2:
        render_kpi_card("Total Buy Signals", f"{stats['total_signals']}", icon_name="zap", accent="purple")
    with m3:
        render_kpi_card("Stocks Scanned", f"{stats['stocks_scanned']}", icon_name="layers", accent="blue")
    with m4:
        render_kpi_card("Stocks with Signals", f"{stats['stocks_with_signals']}", icon_name="check-circle", accent="green")
    with m5:
        render_kpi_card("Avg Signals / Stock", f"{stats['avg_per_stock']:.2f}", icon_name="activity", accent="neutral")

    st.caption(f"Of {stats['total_tracking_started']} tracking cycles started, "
               f"{stats['total_signals']} produced a buy signal in this range "
               f"(across {stats['stocks_tracking']} stocks tracking, "
               f"{stats['stocks_with_signals']} with a buy).")

    if not rows and not tracking_rows:
        st.info("No tracking events or buy signals were generated in this date range for the loaded universe.")
        return
    if not rows:
        st.info("No buy signals were generated in this date range for the loaded universe.")
        return

    # ---------------- Filters ----------------
    render_section_title("Detailed Signals", "list")
    f1, f2, f3 = st.columns([1.2, 1, 1])
    with f1:
        min_sigs = st.number_input("Min buy signals per stock", min_value=1, max_value=50,
                                   value=1, step=1, key="bsa_min")
    with f2:
        st.checkbox("Only stocks with ≥1 signal", value=True, key="bsa_only", disabled=True,
                    help="Every detailed row is a buy signal, so this is always applied.")
    with f3:
        sort_by = st.selectbox("Sort by", ["Signal Date", "Symbol"], key="bsa_sort")

    view = [r for r in rows if stats['sym_counts'].get(r['symbol'], 0) >= min_sigs]
    if sort_by == "Signal Date":
        view.sort(key=lambda r: (r['buy_date'], r['symbol']))
    else:
        view.sort(key=lambda r: (r['symbol'], r['buy_date']))

    table_rows = [{
        'symbol': r['symbol'], 'buy_date_str': r['buy_date_str'], 'close': r['close'],
        'rsi': r['rsi'], 'ema21': r['ema21'], 'ema50': r['ema50'], 'ema200': r['ema200'],
        'tracking_start_str': r['tracking_start_str'], 'days_in_tracking': r['days_in_tracking'],
        'signal_ord': _ordinal(r['signal_number']), 'prev_buy_str': r['prev_buy_str'],
    } for r in view]

    render_premium_table(table_rows, columns=[
        {'key': 'symbol', 'label': 'Symbol'},
        {'key': 'buy_date_str', 'label': 'Buy Date'},
        {'key': 'close', 'label': 'Close', 'fmt': lambda v: f"₹{v:,.2f}"},
        {'key': 'rsi', 'label': 'RSI', 'fmt': lambda v: (f"{v:.1f}" if v is not None else "—")},
        {'key': 'ema21', 'label': 'EMA21', 'fmt': lambda v: (f"₹{v:,.2f}" if v is not None else "—")},
        {'key': 'ema50', 'label': 'EMA50', 'fmt': lambda v: (f"₹{v:,.2f}" if v is not None else "—")},
        {'key': 'ema200', 'label': 'EMA200', 'fmt': lambda v: (f"₹{v:,.2f}" if v is not None else "—")},
        {'key': 'tracking_start_str', 'label': 'Tracking Start'},
        {'key': 'days_in_tracking', 'label': 'Days Tracking', 'fmt': lambda v: (f"{v}" if v is not None else "—")},
        {'key': 'signal_ord', 'label': 'Signal #'},
        {'key': 'prev_buy_str', 'label': 'Prev Buy'},
    ], height=460)

    st.caption(f"Showing {len(view)} of {stats['total_signals']} signals"
               + (f" · filtered to stocks with ≥{min_sigs} signals" if min_sigs > 1 else ""))

    # ---------------- Charts (from the filtered view) ----------------
    render_section_title("Signal Distribution", "trending-up")
    df_v = pd.DataFrame(view)
    df_v['buy_date'] = pd.to_datetime(df_v['buy_date'])

    cc1, cc2 = st.columns(2)
    with cc1:
        mth = df_v.groupby(df_v['buy_date'].dt.to_period('M')).size()
        mfig = go.Figure(go.Bar(x=[str(p) for p in mth.index], y=list(mth.values), marker_color='#8B5CF6'))
        mfig.update_layout(title="Buy Signals per Month")
        render_chart_card(apply_plotly_theme(mfig, height=300), key="bsa_month")
    with cc2:
        wk = df_v.groupby(df_v['buy_date'].dt.to_period('W')).size()
        wfig = go.Figure(go.Bar(x=[p.start_time.strftime('%d-%b') for p in wk.index],
                                y=list(wk.values), marker_color='#3B82F6'))
        wfig.update_layout(title="Buy Signals per Week (week starting)")
        render_chart_card(apply_plotly_theme(wfig, height=300), key="bsa_week")

    cc3, cc4 = st.columns(2)
    with cc3:
        bs = df_v.groupby('symbol').size().sort_values(ascending=False).head(20)
        bfig = go.Figure(go.Bar(x=list(bs.values), y=list(bs.index), orientation='h', marker_color='#10B981'))
        bfig.update_layout(title="Buy Signals by Stock (Top 20)", yaxis=dict(autorange="reversed"))
        render_chart_card(apply_plotly_theme(bfig, height=430), key="bsa_stock")
    with cc4:
        daily = df_v.groupby(df_v['buy_date'].dt.normalize()).size()
        dfig = go.Figure(go.Scatter(x=list(daily.index), y=list(daily.values), mode='lines+markers',
                                    line=dict(color='#F59E0B', width=2), marker=dict(size=5)))
        dfig.update_layout(title="Daily Buy Signal Frequency")
        render_chart_card(apply_plotly_theme(dfig, height=430), key="bsa_daily")

    # ---------------- Tracking Started (detail) ----------------
    with st.expander(f"🎯 Tracking Started events in range ({stats['total_tracking_started']})", expanded=False):
        if tracking_rows:
            tview = sorted(tracking_rows, key=lambda r: (r['track_date'], r['symbol']))
            render_premium_table([{
                'symbol': r['symbol'], 'track_date_str': r['track_date_str'],
                'rsi': r['rsi'], 'close': r['close'], 'ema21': r['ema21'],
                'ema50': r['ema50'], 'ema200': r['ema200'], 'buy_date_str': r['buy_date_str'],
            } for r in tview], columns=[
                {'key': 'symbol', 'label': 'Symbol'},
                {'key': 'track_date_str', 'label': 'Tracking Start (Signal Date)'},
                {'key': 'rsi', 'label': 'RSI', 'fmt': lambda v: (f"{v:.1f}" if v is not None else "—")},
                {'key': 'close', 'label': 'Close', 'fmt': lambda v: f"₹{v:,.2f}"},
                {'key': 'ema21', 'label': 'EMA21', 'fmt': lambda v: (f"₹{v:,.2f}" if v is not None else "—")},
                {'key': 'ema50', 'label': 'EMA50', 'fmt': lambda v: (f"₹{v:,.2f}" if v is not None else "—")},
                {'key': 'ema200', 'label': 'EMA200', 'fmt': lambda v: (f"₹{v:,.2f}" if v is not None else "—")},
                {'key': 'buy_date_str', 'label': 'Resulting Buy'},
            ], height=360)
        else:
            st.info("No tracking-started events in this date range.")

    # ---------------- Export ----------------
    render_section_title("Export", "refresh-cw")

    # Buy-signals detail (as shown in the table above)
    export_df = pd.DataFrame([{
        'Symbol': r['symbol'], 'Buy Signal Date': r['buy_date_str'], 'Close Price': r['close'],
        'RSI(14)': r['rsi'], 'EMA21': r['ema21'], 'EMA50': r['ema50'], 'EMA200': r['ema200'],
        'Tracking Start Date': r['tracking_start_str'], 'Days in Tracking': r['days_in_tracking'],
        'Signal Number': r['signal_number'], 'Previous Buy Date': r['prev_buy_str'],
    } for r in view])

    # Tracking-started detail (exact column set requested)
    tracking_df = pd.DataFrame([{
        'Symbol': r['symbol'], 'Signal Date': r['track_date_str'], 'RSI': r['rsi'],
        'Close': r['close'], 'EMA21': r['ema21'], 'EMA50': r['ema50'], 'EMA200': r['ema200'],
        'Tracking Start Date': r['track_date_str'], 'Buy Signal Date': r['buy_date_str'],
    } for r in sorted(tracking_rows, key=lambda r: (r['track_date'], r['symbol']))])

    # Unified "All Signals": one row per tracking-start AND per buy, sharing the
    # exact columns requested (Symbol, Signal Date, RSI, Close, EMA21/50/200,
    # Tracking Start Date, Buy Signal Date) plus a Type flag.
    all_events = []
    for r in tracking_rows:
        all_events.append({
            'Type': 'TRACKING_START', 'Symbol': r['symbol'], 'Signal Date': r['track_date_str'],
            'RSI': r['rsi'], 'Close': r['close'], 'EMA21': r['ema21'], 'EMA50': r['ema50'],
            'EMA200': r['ema200'], 'Tracking Start Date': r['track_date_str'],
            'Buy Signal Date': (r['buy_date_str'] if r['resulted_in_buy'] else ''),
            '_sort': r['track_date'],
        })
    for r in view:
        all_events.append({
            'Type': 'BUY', 'Symbol': r['symbol'], 'Signal Date': r['buy_date_str'],
            'RSI': r['rsi'], 'Close': r['close'], 'EMA21': r['ema21'], 'EMA50': r['ema50'],
            'EMA200': r['ema200'], 'Tracking Start Date': r['tracking_start_str'],
            'Buy Signal Date': r['buy_date_str'], '_sort': r['buy_date'],
        })
    all_events.sort(key=lambda x: (x['_sort'], x['Type'], x['Symbol']))
    all_signals_df = pd.DataFrame(
        [{k: v for k, v in e.items() if k != '_sort'} for e in all_events]
    )

    e1, e2, e3 = st.columns(3)
    with e1:
        st.download_button(
            "⬇️ Buy Signals CSV", export_df.to_csv(index=False),
            f"buy_signals_{start_date:%Y%m%d}_{end_date:%Y%m%d}.csv", "text/csv",
            use_container_width=True, key="bsa_csv",
        )
    with e2:
        st.download_button(
            "⬇️ All Signals CSV", all_signals_df.to_csv(index=False),
            f"all_signals_{start_date:%Y%m%d}_{end_date:%Y%m%d}.csv", "text/csv",
            use_container_width=True, key="bsa_csv_all",
        )
    with e3:
        xbuf = io.BytesIO()
        with pd.ExcelWriter(xbuf, engine='openpyxl') as writer:
            all_signals_df.to_excel(writer, sheet_name='All Signals', index=False)
            export_df.to_excel(writer, sheet_name='Buy Signals', index=False)
            (tracking_df if not tracking_df.empty
             else pd.DataFrame(columns=['Symbol', 'Signal Date', 'RSI', 'Close', 'EMA21',
                                        'EMA50', 'EMA200', 'Tracking Start Date', 'Buy Signal Date'])
             ).to_excel(writer, sheet_name='Tracking Started', index=False)
            pd.DataFrame([
                {'Metric': 'Total Tracking Started', 'Value': stats['total_tracking_started']},
                {'Metric': 'Total Buy Signals', 'Value': stats['total_signals']},
                {'Metric': 'Stocks Scanned', 'Value': stats['stocks_scanned']},
                {'Metric': 'Stocks Tracking', 'Value': stats['stocks_tracking']},
                {'Metric': 'Stocks with Signals', 'Value': stats['stocks_with_signals']},
                {'Metric': 'Avg Signals per Stock', 'Value': round(stats['avg_per_stock'], 2)},
                {'Metric': 'Date Range', 'Value': f"{start_date:%d-%b-%Y} to {end_date:%d-%b-%Y}"},
            ]).to_excel(writer, sheet_name='Summary', index=False)
            (pd.Series(stats['sym_counts']).sort_values(ascending=False)
             .rename_axis('Symbol').reset_index(name='Buy Signals')
             ).to_excel(writer, sheet_name='Buy Signals by Symbol', index=False)
            (pd.Series(stats['track_counts']).sort_values(ascending=False)
             .rename_axis('Symbol').reset_index(name='Tracking Started')
             ).to_excel(writer, sheet_name='Tracking by Symbol', index=False)
        st.download_button(
            "⬇️ Download Excel (all sheets)", xbuf.getvalue(),
            f"signal_analysis_{start_date:%Y%m%d}_{end_date:%Y%m%d}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, key="bsa_xlsx",
        )


def main():
    st.set_page_config(
        page_title=f"{APP_NAME} v{APP_VERSION}",
        page_icon=APP_ICON,
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # Premium dark theme (design system defined in the UI KIT section above).
    st.markdown(PREMIUM_CSS, unsafe_allow_html=True)

    init_session_state()

    st.markdown(f"""
        <div style="display:flex;align-items:center;justify-content:space-between;
                    padding:2px 2px 12px 2px;margin-bottom:6px;border-bottom:1px solid var(--border-color);">
            <div style="display:flex;align-items:center;gap:10px;">
                <div style="width:34px;height:34px;border-radius:9px;background:var(--accent-gradient);
                            display:flex;align-items:center;justify-content:center;font-size:17px;">📈</div>
                <div>
                    <div style="font-size:17px;font-weight:700;color:var(--text-primary);line-height:1.1;">{APP_NAME}</div>
                    <div style="font-size:11px;color:var(--text-muted);">v{APP_VERSION} · Scanner · Signals · Backtester · Portfolio</div>
                </div>
            </div>
            <span class="badge badge-purple">{icon('activity', 11)} Live</span>
        </div>
    """, unsafe_allow_html=True)

    config = render_sidebar()

    tabs = st.tabs([
        "📊 Dashboard", "🔍 Scanner", "💰 My Portfolio", "📈 Backtest",
        "🎯 Buy Signals", "📋 Trade Log", "📉 Charts", "💼 Portfolio", "📑 Statistics"
    ])

    with tabs[0]:
        render_dashboard(st.session_state.metrics)
    with tabs[1]:
        render_scanner()
    with tabs[2]:
        render_investment_portfolio()
    with tabs[3]:
        render_backtest()
    with tabs[4]:
        render_buy_signal_analysis()
    with tabs[5]:
        render_trade_log()
    with tabs[6]:
        render_charts()
    with tabs[7]:
        render_portfolio()
    with tabs[8]:
        render_statistics()


if __name__ == "__main__":
    main()
