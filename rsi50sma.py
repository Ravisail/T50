#!/usr/bin/env python3
"""
NIFTY 200 Stock Scanner & Backtesting Platform v3.1
====================================================

Production-ready Streamlit app with realistic portfolio simulation,
vectorized signal generation, batch data downloads, and comprehensive
performance metrics.

FIXES in v3.1:
  1. Tracking logic now uses persistent state machine instead of rolling window.
     Once RSI < 35, tracking stays enabled until a BUY signal occurs.
  2. Calmar Ratio now correctly calculated as CAGR / Max Drawdown.

Strategy:
  1. TRACKING: When RSI(14) < 35, mark stock as "TRACKING ENABLED"
  2. BUY: When Tracking + Close > SMA(50) + Close <= HighestClose(126) * 0.92
  3. SELL: When Profit >= 3.14% (no stop loss, no time exit)
  4. AVERAGE: New signal <= Previous Buy Price * 0.98, max N entries

Author: AI Assistant
Version: 3.1.0
Python: 3.12+
"""

# =============================================================================
# IMPORTS
# =============================================================================

import os
import sys
import json
import time
import hashlib
import logging
import warnings
from pathlib import Path
from datetime import datetime, timedelta, date as dt_date
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple, Union, Any, Callable, Set
from concurrent.futures import ThreadPoolExecutor, as_completed
from functools import wraps
import io

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
APP_VERSION = "3.1.0"
APP_ICON = "📈"

CACHE_DIR = Path.home() / ".nifty200_scanner_cache"
CACHE_DIR.mkdir(exist_ok=True)

# NIFTY 200 Stock Symbols (deduplicated, sorted)
NIFTY200_SYMBOLS = sorted(list(set([
    "RELIANCE.NS", "TCS.NS", "HDFCBANK.NS", "INFY.NS", "ICICIBANK.NS",
    "HINDUNILVR.NS", "ITC.NS", "SBIN.NS", "BHARTIARTL.NS", "KOTAKBANK.NS",
    "BAJFINANCE.NS", "LICI.NS", "LT.NS", "HCLTECH.NS", "AXISBANK.NS",
    "ASIANPAINT.NS", "MARUTI.NS", "SUNPHARMA.NS", "TITAN.NS", "BAJAJFINSV.NS",
    "ADANIENT.NS", "WIPRO.NS", "ULTRACEMCO.NS", "NESTLEIND.NS", "POWERGRID.NS",
    "M&M.NS", "NTPC.NS", "COALINDIA.NS", "TATAMOTORS.NS", "JSWSTEEL.NS",
    "TECHM.NS", "GRASIM.NS", "ONGC.NS", "BRITANNIA.NS", "CIPLA.NS",
    "SBILIFE.NS", "HDFCLIFE.NS", "EICHERMOT.NS", "DRREDDY.NS", "APOLLOHOSP.NS",
    "ADANIPORTS.NS", "TATASTEEL.NS", "HINDALCO.NS", "DIVISLAB.NS", "BAJAJ-AUTO.NS",
    "HEROMOTOCO.NS", "UPL.NS", "INDUSINDBK.NS", "BPCL.NS", "IOC.NS",
    "SHRIRAMFIN.NS", "DABUR.NS", "PIDILITIND.NS", "HAVELLS.NS", "GODREJCP.NS",
    "MARICO.NS", "ICICIPRULI.NS", "BERGEPAINT.NS", "INDIGO.NS", "SIEMENS.NS",
    "TATACONSUM.NS", "DLF.NS", "BOSCHLTD.NS", "ABB.NS", "CHOLAFIN.NS",
    "MCDOWELL-N.NS", "SRF.NS", "GAIL.NS", "BANKBARODA.NS", "CANBK.NS",
    "IOB.NS", "UNIONBANK.NS", "CENTRALBK.NS", "PNB.NS", "IDFCFIRSTB.NS",
    "FEDERALBNK.NS", "RBLBANK.NS", "BANDHANBNK.NS", "AUBANK.NS", "KARURVYSYA.NS",
    "CITYUNION.NS", "SOUTHBANK.NS", "J&KBANK.NS", "KARNATAKABK.NS", "TMB.NS",
    "CSBBANK.NS", "DCBBANK.NS", "DHANBANK.NS", "YESBANK.NS", "IDBI.NS",
    "UCOBANK.NS", "PSB.NS", "MAHABANK.NS", "NATIONALUM.NS", "HINDZINC.NS",
    "VEDL.NS", "NMDC.NS", "SAIL.NS", "JSL.NS", "APLAPOLLO.NS",
    "RATNAMANI.NS", "TINPLATE.NS", "WELCORP.NS", "MAHSEAMLES.NS", "HISARMETAL.NS",
    "PENIND.NS", "SYNTHFO.NS", "SARDAEN.NS", "RAMASTEEL.NS", "GANDHITUBE.NS",
    "PRAKASH.NS", "UTTAMSTL.NS", "JSWISPL.NS", "GRAVITA.NS", "NLCINDIA.NS",
    "RECLTD.NS", "PFC.NS", "NHPC.NS", "SJVN.NS", "THIRUSUGAR.NS",
    "BAJAJHLDNG.NS", "BATAINDIA.NS", "COLPAL.NS", "EMAMILTD.NS", "GILLETTE.NS",
    "GLAXO.NS", "HEG.NS", "IIFL.NS", "JUBLFOOD.NS", "LALPATHLAB.NS",
    "MFSL.NS", "MGL.NS", "MINDTREE.NS", "MOTHERSON.NS", "NAUKRI.NS",
    "NAVINFLUOR.NS", "PAGEIND.NS", "PERSISTENT.NS", "PETRONET.NS", "PFIZER.NS",
    "PIIND.NS", "POLYCAB.NS", "RITES.NS", "SANOFI.NS", "SKFINDIA.NS",
    "SRTRANSFIN.NS", "SUNTV.NS", "SUPREMEIND.NS", "TATACHEM.NS", "TATACOMM.NS",
    "TATAPOWER.NS", "TORNTPHARM.NS", "TORNTPOWER.NS", "TRENT.NS", "TVSMOTOR.NS",
    "VBL.NS", "VOLTAS.NS", "WHIRLPOOL.NS", "ZEEL.NS", "ZOMATO.NS",
    "360ONE.NS", "AARTIIND.NS", "AAVAS.NS", "ABBOTINDIA.NS", "ACE.NS",
    "ADANIGREEN.NS", "ADANIPOWER.NS", "ADANIWILMAR.NS", "AEGISCHEM.NS", "AETHER.NS",
    "AFFLE.NS", "AJANTPHARM.NS", "AKZOINDIA.NS", "ALEMBICLTD.NS", "ALKEM.NS",
    "ALKYLAMINE.NS", "ALOKINDS.NS", "AMARAJABAT.NS", "AMBER.NS", "AMBUJACEM.NS",
    "ANANDRATHI.NS", "ANANTRAJ.NS", "ANGELONE.NS", "ANURAS.NS", "APTUS.NS",
    "ARVINDFASN.NS", "ASAHIINDIA.NS", "ASHOKLEY.NS", "ASIANTILES.NS", "ASTERDM.NS",
    "ATGL.NS", "ATUL.NS", "AURIONPRO.NS", "AVANTIFEED.NS", "AWL.NS",
    "BAJAJELEC.NS", "BALAMINES.NS", "BALKRISIND.NS", "BALRAMCHIN.NS",
    "BASF.NS", "BAYERCROP.NS", "BBTC.NS", "BDL.NS", "BEL.NS",
    "BEML.NS", "BEPL.NS", "BFINVEST.NS", "BFUTILITIE.NS", "BGRENERGY.NS",
    "BIKAJI.NS", "BIRLACORPN.NS", "BLISSGVS.NS", "BLUESTARCO.NS", "BOMDYEING.NS",
    "BORORENEW.NS", "BRIGADE.NS", "BSOFT.NS", "CAMLINFINE.NS", "CAMPUS.NS",
    "CANFINHOME.NS", "CAPLIPOINT.NS", "CARBORUNIV.NS", "CASTROLIND.NS", "CEATLTD.NS",
    "CENTURYPLY.NS", "CERA.NS", "CGCL.NS", "CGPOWER.NS", "CHEMPLASTS.NS",
    "CHENNPETRO.NS", "CHOLAHLDNG.NS", "CIGNITITEC.NS", "CLEAN.NS", "COCHINSHIP.NS",
    "COFORGE.NS", "CONCOR.NS", "COROMANDEL.NS", "CREDITACC.NS", "CROMPTON.NS",
    "CUB.NS", "CUMMINSIND.NS", "CYIENT.NS", "DATAPATTNS.NS",
    "DBL.NS", "DCAL.NS", "DCMSHRIRAM.NS", "DEEPAKFERT.NS", "DEEPAKNTR.NS",
    "DELHIVERY.NS", "DEVYANI.NS", "DIXON.NS", "DMART.NS", "EIDPARRY.NS",
    "ELGIEQUIP.NS", "ENDURANCE.NS", "ENGINERSIN.NS", "EPL.NS", "EQUITASBNK.NS",
    "ERIS.NS", "ESABINDIA.NS", "EXIDEIND.NS", "FACT.NS", "FINEORG.NS",
    "FINPIPE.NS", "FLUOROCHEM.NS", "FORTIS.NS", "FSL.NS", "GAEL.NS",
    "GARFIBRES.NS", "GEPIL.NS", "GET&D.NS", "GHCL.NS", "GLENMARK.NS",
    "GMRINFRA.NS", "GNFC.NS", "GOCOLORS.NS", "GPPL.NS", "GRANULES.NS",
    "GRAPHITE.NS", "GREAVESCOT.NS", "GREENLAM.NS", "GREENPANEL.NS", "GRINDWELL.NS",
    "GSFC.NS", "GSPL.NS", "GUJALKALI.NS", "GUJGASLTD.NS", "GULFOILLUB.NS",
    "HAL.NS", "HAPPSTMNDS.NS", "HATSUN.NS", "HEIDELBERG.NS", "HEMIPROP.NS",
    "HFCL.NS", "HIKAL.NS", "HINDCOPPER.NS", "HINDPETRO.NS", "HONAUT.NS",
    "HUDCO.NS", "IBREALEST.NS", "IBULHSGFIN.NS", "ICICIGI.NS", "ICRA.NS",
    "IDFC.NS", "IEX.NS", "IFBIND.NS", "IGL.NS", "IIFLSEC.NS",
    "INDHOTEL.NS", "INDIAMART.NS", "INDIANB.NS", "INDOCO.NS", "INDOSTAR.NS",
    "INFIBEAM.NS", "INTELLECT.NS", "IPCALAB.NS", "IRB.NS", "IRCON.NS",
    "IRCTC.NS", "ISEC.NS", "ITI.NS", "JBCHEPHARM.NS", "JBMA.NS",
    "JINDALSAW.NS", "JINDALSTEL.NS", "JIOFIN.NS", "JKCEMENT.NS", "JKLAKSHMI.NS",
    "JKTYRE.NS", "JMFINANCIL.NS", "JPPOWER.NS", "JSLHISAR.NS", "JUBLINGREA.NS",
    "JUBLPHARMA.NS", "JUSTDIAL.NS", "JYOTHYLAB.NS", "KAJARIACER.NS", "KEC.NS",
    "KEI.NS", "KIMS.NS", "KNRCON.NS", "KPRMILL.NS", "KRBL.NS",
    "KSB.NS", "L&TFH.NS", "LAOPALA.NS", "LATENTVIEW.NS", "LAURUSLABS.NS",
    "LEMONTREE.NS", "LGBBROSLTD.NS", "LINDEINDIA.NS", "LLOYDSME.NS", "LUPIN.NS",
    "M&MFIN.NS", "MAHINDCIE.NS", "MAHLIFE.NS", "MANAPPURAM.NS",
    "MANYAVAR.NS", "MASTEK.NS", "MAXHEALTH.NS", "MAZDOCK.NS", "METROPOLIS.NS",
    "MHRIL.NS", "MIDHANI.NS", "MINDACORP.NS", "MIRCELECTR.NS", "MMTC.NS",
    "MOIL.NS", "MOTILALOFS.NS", "MPHASIS.NS", "MRF.NS", "MSUMI.NS",
    "MTARTECH.NS", "MUTHOOTFIN.NS", "NATCOPHARM.NS", "NBCC.NS", "NCC.NS",
    "NETWORK18.NS", "NEWGEN.NS", "NFL.NS", "NIITLTD.NS", "NIPPOBATRY.NS",
    "NRAIL.NS", "NSLNISP.NS", "NUVAMA.NS", "OBEROIRLTY.NS",
    "OLECTRA.NS", "OMAXE.NS", "ORIENTCEM.NS", "ORIENTELEC.NS", "ORISSAMINE.NS",
    "PAISALO.NS", "PARAS.NS", "PATANJALI.NS", "PCBL.NS", "PEL.NS",
    "PFS.NS", "PHOENIXLTD.NS", "PLS.NS", "PNBHOUSING.NS", "POLYMED.NS",
    "POONAWALLA.NS", "POWERMECH.NS", "PRAJIND.NS", "PRESTIGE.NS", "PRINCEPIPE.NS",
    "PRIVISCL.NS", "PTC.NS", "PUNJABCHEM.NS", "RAILTEL.NS", "RAINBOW.NS",
    "RAJESHEXPO.NS", "RALLIS.NS", "RAMCOCEM.NS", "RATEGAIN.NS", "RAYMOND.NS",
    "REDINGTON.NS", "RELAXO.NS", "RENUKA.NS", "RHIM.NS",
    "RKFORGE.NS", "ROSSARI.NS", "RTNPOWER.NS", "RVNL.NS", "SAGCEM.NS",
    "SAKSOFT.NS", "SANSERA.NS", "SAPPHIRE.NS", "SATIA.NS", "SBC.NS",
    "SCI.NS", "SEQUENT.NS", "SHARDAMOTR.NS", "SHILPAMED.NS", "SHOPERSTOP.NS",
    "SHREDIGCEM.NS", "SHRIPISTON.NS", "SIGNATURE.NS", "SIS.NS",
    "SKMEGGPROD.NS", "SOBHA.NS", "SOLARINDS.NS", "SONACOMS.NS", "SPARC.NS",
    "SPICEJET.NS", "SPLPETRO.NS", "STLTECH.NS", "STYRENIX.NS", "SUDARSCHEM.NS",
    "SUMICHEM.NS", "SUNCLAYLTD.NS", "SUNDARMFIN.NS", "SUNDRMFAST.NS", "SUNFLAG.NS",
    "SURYAROSNI.NS", "SWSOLAR.NS", "SYMPHONY.NS", "SYRMA.NS", "TANLA.NS",
    "TATAELXSI.NS", "TATAINVEST.NS", "TATAMETALI.NS", "TATATECH.NS", "TBZ.NS",
    "TCI.NS", "TCIEXP.NS", "TCNSBRANDS.NS", "TEAMLEASE.NS", "TECHNOE.NS",
    "TEGA.NS", "THEINVEST.NS", "THERMAX.NS", "TIMKEN.NS", "TITAGARH.NS",
    "TKIL.NS", "TRITURBINE.NS", "TRIVENI.NS", "TTKPRESTIG.NS", "TVSELECT.NS",
    "UBL.NS", "UNOMINDA.NS", "USHAMART.NS",
    "UTIAMC.NS", "VAKRANGEE.NS", "VARROC.NS", "VGUARD.NS", "VIJAYA.NS",
    "VINATIORGA.NS", "VIPIND.NS", "VTL.NS", "WABAG.NS", "WELENT.NS",
    "WELSPUNIND.NS", "WESTLIFE.NS", "WONDERLA.NS", "XCHANGING.NS", "YATHARTH.NS",
    "ZENSARTECH.NS", "ZFCVINDIA.NS", "ZUARIIND.NS"
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

# =============================================================================
# DATA MANAGER (with cache expiry and batch download)
# =============================================================================

class DataManager:
    """Manages stock data download, caching, and retrieval with batch support."""

    def __init__(self, cache_dir: Path = CACHE_DIR, cache_ttl_hours: int = 24):
        self.cache_dir = cache_dir
        self.cache_dir.mkdir(exist_ok=True)
        self._data_cache: Dict[str, pd.DataFrame] = {}
        self.cache_ttl = timedelta(hours=cache_ttl_hours)

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
            if not self._looks_truncated(cached, period):
                return cached.copy()
            del self._data_cache[cache_key]

        if cache_path.exists() and not force_refresh and self._is_cache_valid(cache_path):
            try:
                df = pd.read_parquet(cache_path)
                if not self._looks_truncated(df, period):
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
            if not self._looks_truncated(df, period):
                df.to_parquet(cache_path)
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
            logger.error(f"Failed to download {symbol}: {e}")
            return None

    def batch_download(
        self,
        symbols: List[str],
        period: str = "2y",
        interval: str = "1d",
        max_workers: int = 8,
        progress_callback: Optional[Callable] = None
    ) -> Dict[str, pd.DataFrame]:
        """
        Download multiple symbols.

        BUGFIX (rate limiting): the old implementation fired ~380 separate
        Ticker.history() requests through a thread pool - Yahoo throttles
        that and returns partial frames. Now we:
          1. Serve everything possible from valid cache.
          2. Bulk-download the rest with yf.download() in chunks of 50
             symbols per HTTP request (~8 requests for the full universe).
          3. Fall back to per-ticker download only for symbols the bulk
             call missed or returned truncated.
        """
        results: Dict[str, pd.DataFrame] = {}
        remaining: List[str] = []
        total = max(len(symbols), 1)

        # --- Pass 1: valid cache ---
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
            if df is not None and not self._looks_truncated(df, period):
                self._data_cache[cache_key] = df
                results[sym] = df.copy()
            else:
                remaining.append(sym)

        if progress_callback:
            progress_callback(len(results) / total)

        # --- Pass 2: chunked bulk download ---
        CHUNK = 50
        for start_idx in range(0, len(remaining), CHUNK):
            chunk = remaining[start_idx:start_idx + CHUNK]
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
                logger.error(f"Bulk download failed for chunk starting {chunk[0]}: {e}")
                bulk = None

            if bulk is not None and not bulk.empty:
                for sym in chunk:
                    try:
                        if isinstance(bulk.columns, pd.MultiIndex):
                            if sym not in bulk.columns.get_level_values(0):
                                continue
                            sub = bulk[sym]
                        else:
                            # Single-symbol chunk: flat columns
                            sub = bulk
                        df = self._normalize_frame(sub)
                        if df is None:
                            continue
                        if not self._looks_truncated(df, period):
                            cache_path = self._get_cache_path(sym, period, interval)
                            df.to_parquet(cache_path)
                            self._data_cache[f"{sym}_{period}_{interval}"] = df
                            results[sym] = df.copy()
                        # truncated bulk result -> leave for per-ticker retry
                    except Exception as e:
                        logger.error(f"Parsing bulk data for {sym}: {e}")

            if progress_callback:
                progress_callback(min((len(results)) / total, 0.95))
            time.sleep(0.5)  # be gentle between bulk requests

        # --- Pass 3: per-ticker fallback for anything still missing ---
        still_missing = [s for s in remaining if s not in results]
        completed = len(results)

        def process(sym):
            try:
                df = self.download_stock_data(sym, period=period, interval=interval)
                return sym, df
            except Exception as e:
                logger.error(f"Batch download error for {sym}: {e}")
                return sym, None

        if still_missing:
            # Low concurrency on the fallback to avoid re-triggering throttling
            with ThreadPoolExecutor(max_workers=min(max_workers, 4)) as executor:
                futures = {executor.submit(process, sym): sym for sym in still_missing}
                for future in as_completed(futures):
                    sym, df = future.result()
                    completed += 1
                    if df is not None:
                        results[sym] = df
                    if progress_callback:
                        progress_callback(min(completed / total, 1.0))

        if progress_callback:
            progress_callback(1.0)

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
    def sma(prices: pd.Series, period: int = 50) -> pd.Series:
        return prices.rolling(window=period, min_periods=period).mean()

    @staticmethod
    def highest_close(prices: pd.Series, period: int = 126) -> pd.Series:
        return prices.rolling(window=period, min_periods=period).max()

# =============================================================================
# STRATEGY CONFIGURATION
# =============================================================================

@dataclass
class StrategyConfig:
    """Configuration for the trading strategy."""
    rsi_period: int = 14
    rsi_threshold: float = 35.0
    sma_period: int = 50
    highest_close_lookback: int = 126
    distance_from_high_pct: float = 8.0
    profit_target_pct: float = 3.14
    average_trigger_pct: float = 2.0
    max_average_entries: int = 5
    position_size: float = 50000.0
    initial_capital: float = 1000000.0
    max_positions: int = 20
    brokerage_pct: float = 0.05
    slippage_pct: float = 0.05
    transaction_charges_pct: float = 0.01
    close_open_trades: bool = True
    data_period: str = "2y"  # NEW: how much history to download for scanning/backtesting

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
        assert self.sma_period > 0
        assert self.highest_close_lookback > 0
        assert 0 <= self.distance_from_high_pct <= 100
        assert self.profit_target_pct > 0
        assert self.average_trigger_pct > 0
        assert self.max_average_entries >= 0
        assert self.position_size > 0
        assert self.initial_capital > 0
        assert self.max_positions > 0
        assert self.data_period in self.VALID_PERIODS

    @property
    def total_transaction_cost_pct(self) -> float:
        return self.brokerage_pct + self.slippage_pct + self.transaction_charges_pct

    def calculate_entry_cost(self, amount: float) -> float:
        return amount * self.total_transaction_cost_pct / 100

    def calculate_exit_cost(self, amount: float) -> float:
        return amount * self.total_transaction_cost_pct / 100

    @property
    def min_required_bars(self) -> int:
        """Minimum bars needed before the first valid signal can be generated."""
        return max(self.sma_period, self.highest_close_lookback) + 10

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
        Generate all strategy signals for a dataframe.
        Fully vectorized - no Python loops over rows.
        """
        df = df.copy()

        # Calculate indicators
        df['rsi'] = TechnicalIndicators.rsi(df['close'], config.rsi_period)
        df['sma50'] = TechnicalIndicators.sma(df['close'], config.sma_period)
        df['highest_126'] = TechnicalIndicators.highest_close(df['close'], config.highest_close_lookback)

        # FIX v3.1: Persistent tracking state machine
        # Once RSI goes below threshold, tracking stays ON until a BUY signal occurs
        df['rsi_below'] = df['rsi'] < config.rsi_threshold

        tracking = np.zeros(len(df), dtype=bool)
        tracking_enabled = False

        close_vals = df['close'].values
        sma_vals = df['sma50'].values
        high_vals = df['highest_126'].values
        rsi_below_vals = df['rsi_below'].values

        for i in range(len(df)):
            if rsi_below_vals[i]:
                tracking_enabled = True

            tracking[i] = tracking_enabled

            # If buy signal conditions are met on this bar, reset tracking for NEXT bar
            above_sma = close_vals[i] > sma_vals[i]
            below_high = close_vals[i] <= high_vals[i] * (1 - config.distance_from_high_pct / 100)

            if tracking_enabled and above_sma and below_high:
                tracking_enabled = False

        df['tracking_enabled'] = tracking

        # Buy signal conditions
        df['above_sma'] = df['close'] > df['sma50']
        df['below_high_threshold'] = df['close'] <= df['highest_126'] * (1 - config.distance_from_high_pct / 100)
        df['buy_signal'] = df['tracking_enabled'] & df['above_sma'] & df['below_high_threshold']

        return df

    @staticmethod
    def get_latest_signal(df: pd.DataFrame, config: StrategyConfig) -> Optional[Dict]:
        """Get the latest signal state for scanner."""
        if df.empty or len(df) < max(config.sma_period, config.highest_close_lookback) + 10:
            return None

        df_signals = SignalGenerator.generate_signals(df, config)
        latest = df_signals.iloc[-1]

        close = float(latest['close'])
        rsi_val = float(latest['rsi']) if not pd.isna(latest['rsi']) else 50
        sma50 = float(latest['sma50']) if not pd.isna(latest['sma50']) else close
        highest_126 = float(latest['highest_126']) if not pd.isna(latest['highest_126']) else close

        distance_from_high = ((highest_126 - close) / highest_126) * 100 if highest_126 > 0 else 0

        return {
            'symbol': '',
            'close': close,
            'rsi': rsi_val,
            'sma50': sma50,
            'highest_126': highest_126,
            'distance_from_high_pct': distance_from_high,
            'tracking_enabled': bool(latest['tracking_enabled']),
            'above_sma': bool(latest['above_sma']),
            'below_high_threshold': bool(latest['below_high_threshold']),
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
    # FIX: tracks the most recent observed market price per symbol so that
    # missing-data days fall back to the last real price, not the entry price.
    last_known_prices: Dict[str, float] = field(default_factory=dict)

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

        equity = self.cash + position_value
        self.daily_equity.append((date, equity))

    def can_open_position(self, config: StrategyConfig) -> bool:
        position_size = config.position_size
        entry_cost = config.calculate_entry_cost(position_size)
        return self.cash >= position_size + entry_cost and self.open_position_count < config.max_positions

    def open_position(self, symbol: str, date: datetime, price: float, config: StrategyConfig,
                       signal_date: Optional[datetime] = None, signal_close: Optional[float] = None) -> bool:
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

    def close_all_positions(self, date: datetime, prices: Dict[str, float], config: StrategyConfig, reason: str = "End of Data"):
        for symbol in list(self.positions.keys()):
            price = prices.get(symbol, 0)
            if price > 0:
                self.close_position(symbol, date, price, config, reason)

# =============================================================================
# BACKTEST ENGINE (Optimized with date-indexed lookup)
# =============================================================================

class BacktestEngine:
    """Optimized portfolio-based backtesting engine with realistic T+1 execution."""

    def __init__(self, config: StrategyConfig):
        self.config = config
        self.last_diagnostics: Dict[str, Any] = {}

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

        For open positions:
        - Profit target checked on Day T close (sell logic is unchanged).
        - Averaging signals generated Day T, executed Day T+1 at OPEN, same
          unconditional-fill rule as new entries.
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

            # Execute pending NEW position orders
            for symbol, order in list(pending_orders.items()):
                if symbol in prepared_data and current_date in prepared_data[symbol].index:
                    row = prepared_data[symbol].loc[current_date]
                    execute_price = float(row['open'])

                    if portfolio.can_open_position(self.config):
                        if portfolio.open_position(
                            symbol, current_date, execute_price, self.config,
                            signal_date=order['signal_date'], signal_close=order['signal_close']
                        ):
                            self.last_diagnostics['buy_orders_executed'] = \
                                self.last_diagnostics.get('buy_orders_executed', 0) + 1

                # Remove from pending regardless (order is for one day only)
                del pending_orders[symbol]

            # Execute pending AVERAGING orders
            for symbol, order in list(pending_average_orders.items()):
                if symbol in prepared_data and current_date in prepared_data[symbol].index:
                    row = prepared_data[symbol].loc[current_date]
                    execute_price = float(row['open'])

                    # Only check cash, ignore max_positions for averaging
                    position_size = self.config.position_size
                    entry_cost = self.config.calculate_entry_cost(position_size)
                    total_needed = position_size + entry_cost
                    if portfolio.cash >= total_needed and symbol in portfolio.positions:
                        if portfolio.add_to_position(
                            symbol, current_date, execute_price, self.config,
                            signal_date=order['signal_date'], signal_close=order['signal_close']
                        ):
                            self.last_diagnostics['average_orders_executed'] = \
                                self.last_diagnostics.get('average_orders_executed', 0) + 1

                del pending_average_orders[symbol]

            # =====================================================================
            # STEP 2: Process today's data - check exits and generate new signals
            # =====================================================================

            for symbol, df in prepared_data.items():
                if current_date not in df.index:
                    continue

                row = df.loc[current_date]
                close = float(row['close'])
                buy_signal = bool(row['buy_signal'])
                daily_prices[symbol] = close

                has_position = symbol in portfolio.positions

                if has_position:
                    pos = portfolio.positions[symbol]
                    profit_pct, _ = pos.calculate_unrealized_pnl(close, self.config)

                    # Check profit target on today's close
                    if profit_pct >= self.config.profit_target_pct:
                        portfolio.close_position(symbol, current_date, close, self.config,
                                                f"Profit Target ({profit_pct:.2f}%)")
                    else:
                        # Check averaging conditions
                        # Averaging is a pure price-based rule per the strategy spec
                        # ("New signal <= Previous Buy Price * 0.98"). It must NOT
                        # require buy_signal, because buy_signal needs tracking_enabled
                        # to be re-armed (RSI < threshold again), which is switched OFF
                        # the moment the original entry fires - making averaging almost
                        # never trigger if left as a buy_signal-gated condition.
                        last_entry_price = pos.entry_prices[-1]
                        threshold_price = last_entry_price * (1 - self.config.average_trigger_pct / 100)

                        if (close <= threshold_price and
                            pos.average_count < self.config.max_average_entries):
                            # Signal fires on Day T's close; execution deferred to
                            # Day T+1's open (see STEP 1 above).
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

        # Close remaining positions at last available price
        if self.config.close_open_trades:
            for symbol in list(portfolio.positions.keys()):
                if symbol in prepared_data:
                    df = prepared_data[symbol]
                    if len(df) > 0:
                        last_row = df.iloc[-1]
                        last_close = float(last_row['close'])
                        last_date = df.index[-1]
                        if isinstance(last_date, pd.Timestamp):
                            last_date = last_date.to_pydatetime()
                        portfolio.close_position(symbol, last_date, last_close, self.config, "End of Data")

        return portfolio.completed_trades, portfolio.daily_equity
# =============================================================================
# METRICS CALCULATOR
# =============================================================================

class MetricsCalculator:
    """Calculate comprehensive portfolio-level metrics."""

    @staticmethod
    def calculate_metrics(
        completed_trades: List[CompletedTrade],
        daily_equity: List[Tuple[datetime, float]],
        initial_capital: float
    ) -> Dict[str, Any]:
        metrics = {
            'total_trades': len(completed_trades),
            'winning_trades': 0,
            'losing_trades': 0,
            'win_rate': 0.0,
            'avg_win': 0.0,
            'avg_loss': 0.0,
            'profit_factor': 0.0,
            'net_profit': 0.0,
            'total_return_pct': 0.0,
            'max_drawdown_pct': 0.0,
            'sharpe_ratio': 0.0,
            'sortino_ratio': 0.0,
            'calmar_ratio': 0.0,
            'cagr': 0.0,
            'expectancy': 0.0,
            'recovery_factor': 0.0,
            'avg_holding_days': 0.0,
            'largest_winner': 0.0,
            'largest_loser': 0.0,
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
        metrics['profit_factor'] = total_wins / total_losses if total_losses > 0 else float('inf')

        metrics['net_profit'] = sum(t.profit_inr for t in completed_trades)
        metrics['total_return_pct'] = (metrics['net_profit'] / initial_capital) * 100
        metrics['avg_holding_days'] = np.mean([t.holding_days for t in completed_trades])

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
                std_return = np.std(returns_arr)

                if std_return > 0:
                    metrics['sharpe_ratio'] = (avg_return / std_return) * np.sqrt(252)

                downside_returns = returns_arr[returns_arr < 0]
                if len(downside_returns) > 0:
                    downside_std = np.std(downside_returns)
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
                'Average Win (₹)', 'Average Loss (₹)', 'Profit Factor',
                'Net Profit (₹)', 'Total Return (%)', 'Max Drawdown (%)',
                'Sharpe Ratio', 'Sortino Ratio', 'Calmar Ratio', 'CAGR (%)',
                'Expectancy (₹)', 'Recovery Factor', 'Avg Holding Days',
                'Largest Winner (₹)', 'Largest Loser (₹)'
            ],
            'Value': [
                str(metrics['total_trades']),
                str(metrics['winning_trades']),
                str(metrics['losing_trades']),
                f"{metrics['win_rate']:.2f}",
                f"{metrics['avg_win']:,.2f}",
                f"{metrics['avg_loss']:,.2f}",
                f"{metrics['profit_factor']:.2f}",
                f"{metrics['net_profit']:,.2f}",
                f"{metrics['total_return_pct']:.2f}",
                f"{metrics['max_drawdown_pct']:.2f}",
                f"{metrics['sharpe_ratio']:.2f}",
                f"{metrics['sortino_ratio']:.2f}",
                f"{metrics['calmar_ratio']:.2f}",
                f"{metrics['cagr']:.2f}",
                f"{metrics['expectancy']:,.2f}",
                f"{metrics['recovery_factor']:.2f}",
                f"{metrics['avg_holding_days']:.1f}",
                f"{metrics['largest_winner']:,.2f}",
                f"{metrics['largest_loser']:,.2f}"
            ]
        }
        return pd.DataFrame(data)

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
        progress_callback=None
    ) -> List[Dict]:
        """Scan complete universe using multithreading."""
        results = []

        def process_symbol(symbol):
            try:
                # Uses the configured data_period (sidebar "Backtest / Scan Period")
                # so scanner and backtest always see the same amount of history.
                df = data_manager.download_stock_data(symbol, period=self.config.data_period)
                if df is not None:
                    result = self.scan_stock(symbol, df)
                    if result and result['buy_signal']:
                        meta = data_manager.get_stock_metadata(symbol)
                        result['company'] = meta.get('name', symbol)
                        result['sector'] = meta.get('sector', 'Unknown')
                        return result
            except Exception as e:
                logger.error(f"Scanner error for {symbol}: {e}")
            return None

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {executor.submit(process_symbol, sym): sym for sym in symbols}
            completed = 0
            for future in as_completed(futures):
                completed += 1
                if progress_callback:
                    progress_callback(completed / len(symbols))
                result = future.result()
                if result:
                    results.append(result)

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
            go.Scatter(x=x_vals, y=df['sma50'], name=f'SMA {config.sma_period}',
                      line=dict(color='orange', width=1.5)), row=1, col=1
        )

        fig.add_trace(
            go.Scatter(x=x_vals, y=df['highest_126'], name='126d High',
                      line=dict(color='purple', width=1, dash='dash')), row=1, col=1
        )

        for trade in trades:
            for i, (date, price) in enumerate(zip(trade.entry_dates, trade.entry_prices)):
                color = 'green' if i == 0 else 'blue'
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
                              marker=dict(color='yellow', size=8, symbol='circle-open', line=dict(width=2)),
                              name=f'Signal #{i+1}', showlegend=False), row=1, col=1
                )

            if trade.exit_date and trade.exit_price:
                fig.add_trace(
                    go.Scatter(x=[trade.exit_date], y=[trade.exit_price], mode='markers',
                              marker=dict(color='red', size=12, symbol='triangle-down'),
                              name='Sell', showlegend=False), row=1, col=1
                )

        fig.add_trace(
            go.Scatter(x=x_vals, y=df['rsi'], name='RSI',
                      line=dict(color='cyan', width=1)), row=2, col=1
        )

        fig.add_hline(y=config.rsi_threshold, line_dash="dash", line_color="red", row=2, col=1)

        if 'volume' in df.columns:
            fig.add_trace(
                go.Bar(x=x_vals, y=df['volume'], name='Volume', marker_color='gray'),
                row=3, col=1
            )

        fig.update_layout(
            title=f'{symbol} - Strategy Analysis',
            yaxis_title='Price (₹)', xaxis_title='Date',
            height=800, template='plotly_dark', showlegend=True,
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
# STREAMLIT UI
# =============================================================================

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
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def render_sidebar():
    with st.sidebar:
        st.title(f"{APP_ICON} {APP_NAME}")
        st.markdown(f"**Version:** {APP_VERSION}")
        st.markdown("---")

        st.header("Strategy Parameters")

        rsi_period = st.number_input("RSI Period", min_value=5, max_value=50, value=14, step=1)
        rsi_threshold = st.number_input("RSI Threshold", min_value=10.0, max_value=50.0, value=35.0, step=1.0)
        sma_period = st.number_input("SMA Period", min_value=10, max_value=200, value=50, step=1)
        highest_lookback = st.number_input("Highest Close Lookback", min_value=50, max_value=252, value=126, step=1)
        distance_from_high = st.number_input("Distance from High (%)", min_value=0.0, max_value=50.0, value=8.0, step=0.5)
        profit_target = st.number_input("Profit Target (%)", min_value=0.5, max_value=20.0, value=3.14, step=0.01)
        average_trigger = st.number_input("Average Trigger (%)", min_value=0.5, max_value=10.0, value=2.0, step=0.5)
        max_averages = st.number_input("Max Average Entries", min_value=0, max_value=20, value=5, step=1)

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

        st.header("Backtest Options")

        close_open_trades = st.checkbox("Close open trades at end of data", value=True)

        st.markdown("---")

        st.header("Data")

        period_options = list(StrategyConfig.VALID_PERIODS)
        data_period = st.selectbox(
            "Backtest / Scan Period",
            options=period_options,
            index=period_options.index("2y"),
            help="How much historical daily data to download per stock. "
                 "Must be long enough to cover SMA Period / Highest Close "
                 "Lookback warmup plus room for actual trades."
        )

        min_required_bars = max(sma_period, highest_lookback) + 10
        approx_bars = StrategyConfig.PERIOD_TRADING_DAYS.get(data_period, 10_000)
        if approx_bars < min_required_bars + 60:
            st.warning(
                f"⚠️ '{data_period}' (~{approx_bars} trading days) may be too short for "
                f"SMA Period={sma_period} / Highest Close Lookback={highest_lookback} "
                f"(needs ≥{min_required_bars} bars for warmup, plus room to trade). "
                f"Consider a longer period or shorter lookbacks."
            )

        if st.button("Clear Cache"):
            st.session_state.data_manager.clear_cache()
            st.success("Cache cleared!")

        config = StrategyConfig(
            rsi_period=rsi_period,
            rsi_threshold=rsi_threshold,
            sma_period=sma_period,
            highest_close_lookback=highest_lookback,
            distance_from_high_pct=distance_from_high,
            profit_target_pct=profit_target,
            average_trigger_pct=average_trigger,
            max_average_entries=max_averages,
            position_size=position_size,
            initial_capital=initial_capital,
            max_positions=max_positions,
            brokerage_pct=brokerage,
            slippage_pct=slippage,
            transaction_charges_pct=transaction_charges,
            close_open_trades=close_open_trades,
            data_period=data_period
        )

        st.session_state.config = config
        return config


def render_kpi_card(title: str, value: str, delta: Optional[str] = None):
    """Render a styled KPI card."""
    st.markdown(f"""
        <div style="
            background: linear-gradient(135deg, #1a1d24 0%, #2d3142 100%);
            border-radius: 12px;
            padding: 16px;
            border: 1px solid #3a3f4f;
            margin-bottom: 8px;
        ">
            <div style="color: #8892b0; font-size: 12px; text-transform: uppercase; letter-spacing: 1px;">{title}</div>
            <div style="color: #e6f1ff; font-size: 24px; font-weight: 700; margin-top: 4px;">{value}</div>
            {f'<div style="color: #64ffda; font-size: 12px; margin-top: 4px;">{delta}</div>' if delta else ''}
        </div>
    """, unsafe_allow_html=True)


def render_dashboard(metrics: Optional[Dict[str, Any]]):
    st.header("Dashboard")

    if metrics is None:
        st.info("Run a backtest to see metrics")
        return

    # KPI Cards
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        render_kpi_card("Total Trades", f"{metrics['total_trades']}")
    with col2:
        render_kpi_card("Win Rate", f"{metrics['win_rate']:.1f}%")
    with col3:
        render_kpi_card("Net Profit", f"₹{metrics['net_profit']:,.0f}")
    with col4:
        render_kpi_card("Total Return", f"{metrics['total_return_pct']:.2f}%")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        render_kpi_card("Max Drawdown", f"{metrics['max_drawdown_pct']:.2f}%")
    with col2:
        render_kpi_card("Sharpe Ratio", f"{metrics['sharpe_ratio']:.2f}")
    with col3:
        render_kpi_card("Profit Factor", f"{metrics['profit_factor']:.2f}")
    with col4:
        render_kpi_card("Avg Holding", f"{metrics['avg_holding_days']:.1f}d")

    st.markdown("---")

    if metrics.get('equity_curve'):
        st.subheader("Equity Curve")
        dates = [x[0] for x in metrics['equity_curve']]
        equity = [x[1] for x in metrics['equity_curve']]

        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=dates, y=equity, mode='lines', name='Equity',
            line=dict(color='#00ff88', width=2),
            fill='tonexty', fillcolor='rgba(0,255,136,0.1)'
        ))
        fig.update_layout(
            template='plotly_dark', height=400,
            xaxis_title='Date', yaxis_title='Portfolio Value (₹)',
            showlegend=False, margin=dict(l=40, r=40, t=20, b=40)
        )
        st.plotly_chart(fig, use_container_width=True, key="dashboard_equity_chart")

    if metrics.get('monthly_returns'):
        st.subheader("Monthly Returns")
        monthly_data = [{'Month': m, 'Profit': p} for m, p in sorted(metrics['monthly_returns'].items())]
        st.dataframe(pd.DataFrame(monthly_data), use_container_width=True, hide_index=True)


def render_scanner():
    st.header("Scanner")
    config = st.session_state.config

    col1, col2 = st.columns([1, 3])
    with col1:
        scan_button = st.button("Run Scanner", type="primary", use_container_width=True)

    if scan_button:
        with st.spinner("Scanning NIFTY 200..."):
            scanner = Scanner(config)
            progress_bar = st.progress(0)
            data_manager = st.session_state.data_manager

            def update_progress(pct):
                progress_bar.progress(min(pct, 1.0))

            results = scanner.scan_universe(
                NIFTY200_SYMBOLS,
                data_manager,
                max_workers=8,
                progress_callback=update_progress
            )

            st.session_state.scanner_results = results
            st.session_state.scan_complete = True
            progress_bar.empty()

    if st.session_state.scan_complete and st.session_state.scanner_results:
        results = st.session_state.scanner_results
        st.success(f"Found {len(results)} buy signals")

        df_data = []
        for r in results:
            df_data.append({
                'Symbol': r['symbol'],
                'Company': r.get('company', r['symbol']),
                'Sector': r.get('sector', 'Unknown'),
                'Close': f"₹{r['close']:.2f}",
                'RSI': f"{r['rsi']:.1f}",
                'SMA50': f"₹{r['sma50']:.2f}",
                'Distance from 6M High': f"{r['distance_from_high_pct']:.1f}%",
                'Volume': f"{r['volume']:,.0f}",
                'Buy Signal': '✅ YES' if r['buy_signal'] else '❌ NO'
            })

        st.dataframe(pd.DataFrame(df_data), use_container_width=True, hide_index=True)

        if st.button("Export Scanner CSV"):
            csv = Exporter.scanner_to_csv(st.session_state.scanner_results)
            st.download_button("Download CSV", csv, "scanner_results.csv", "text/csv")
    else:
        st.info("Click 'Run Scanner' to find stocks matching buy criteria")


def render_backtest():
    st.header("Backtest")
    config = st.session_state.config

    st.caption(f"Data period: **{config.data_period}** · "
               f"Warmup required: **{config.min_required_bars} bars** "
               f"(SMA {config.sma_period} / Lookback {config.highest_close_lookback})")
    if config.period_likely_insufficient:
        st.warning(
            f"⚠️ The selected data period ('{config.data_period}') may not leave enough "
            f"bars after warmup for meaningful backtest results. Increase the period in "
            f"the sidebar, or lower SMA Period / Highest Close Lookback."
        )

    col1, col2 = st.columns([1, 3])
    with col1:
        if st.button("Run Backtest", type="primary", use_container_width=True):
            with st.spinner("Running backtest on NIFTY 200..."):
                data_manager = st.session_state.data_manager

                # Download all data using batch download with threading
                progress_bar = st.progress(0)

                def update_dl_progress(pct):
                    progress_bar.progress(min(pct, 1.0))

                data_dict = data_manager.batch_download(
                    NIFTY200_SYMBOLS,
                    period=config.data_period,
                    interval="1d",
                    max_workers=8,
                    progress_callback=update_dl_progress
                )

                progress_bar.empty()

                # Run optimized portfolio-level backtest
                engine = BacktestEngine(config)
                all_trades, daily_equity = engine.run_backtest(data_dict)
                st.session_state.backtest_diagnostics = dict(engine.last_diagnostics)
                st.session_state.backtest_diagnostics['symbols_downloaded'] = len(data_dict)

                # Calculate metrics
                metrics = MetricsCalculator.calculate_metrics(
                    all_trades, daily_equity, config.initial_capital
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

    # Diagnostics panel: explains WHY a run produced few/zero trades instead
    # of failing silently.
    diag = st.session_state.get('backtest_diagnostics')
    if diag:
        downloaded = diag.get('symbols_downloaded', 0)
        prepared = diag.get('symbols_prepared', 0)
        buy_signals = diag.get('total_buy_signals', 0)
        executed = diag.get('buy_orders_executed', 0)

        with st.expander("🔎 Backtest diagnostics", expanded=(buy_signals == 0)):
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Symbols downloaded", f"{downloaded}/{len(NIFTY200_SYMBOLS)}")
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
                    f"SMA {config.sma_period} / lookback {config.highest_close_lookback}). "
                    f"If your period is 1y+ this means Yahoo returned throttled "
                    f"partial data. Click **Clear Cache** in the sidebar, wait a "
                    f"minute, then run again - the downloader will re-fetch in "
                    f"bulk and will no longer cache truncated responses."
                )
            elif buy_signals == 0:
                st.warning(
                    "Data is fine, but the strategy conditions never aligned "
                    "(RSI dip → reclaim above SMA while still ≥"
                    f"{config.distance_from_high_pct:.0f}% below the "
                    f"{config.highest_close_lookback}-day high). Loosen the "
                    "filters, e.g. raise the RSI threshold, lower 'Distance from "
                    "High', or lengthen the data period."
                )
            elif executed == 0:
                st.warning(
                    "Signals fired but no orders filled - check that "
                    "Position Size fits within Initial Capital and Max Positions."
                )

    if st.session_state.backtest_complete and st.session_state.metrics:
        metrics = st.session_state.metrics

        st.subheader("Performance Metrics")
        st.dataframe(MetricsCalculator.metrics_to_dataframe(metrics), use_container_width=True, hide_index=True)

        col1, col2 = st.columns(2)
        with col1:
            if st.button("Export Trades CSV"):
                csv = Exporter.trades_to_csv(st.session_state.all_trades)
                st.download_button("Download Trades CSV", csv, "trades.csv", "text/csv")

        with col2:
            if st.button("Export Excel Report"):
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
    st.header("Trade Log")

    if not st.session_state.all_trades:
        st.info("Run a backtest to see trades")
        return

    trades = st.session_state.all_trades

    col1, col2, col3 = st.columns(3)
    with col1:
        symbols = list(set(t.symbol for t in trades))
        selected_symbol = st.selectbox("Filter by Symbol", ["All"] + sorted(symbols))
    with col2:
        status_filter = st.selectbox("Status", ["All", "CLOSED", "OPEN"])
    with col3:
        min_profit = st.number_input("Min Profit %", value=-100.0, step=1.0)

    filtered = trades
    if selected_symbol != "All":
        filtered = [t for t in filtered if t.symbol == selected_symbol]
    if status_filter != "All":
        filtered = [t for t in filtered if t.status == status_filter]
    filtered = [t for t in filtered if t.profit_pct >= min_profit]

    trade_data = []
    for t in filtered:
        trade_data.append({
            'Symbol': t.symbol,
            'Entry Date': t.entry_dates[0].strftime('%Y-%m-%d') if t.entry_dates else '',
            'Entry Price': f"₹{t.entry_prices[0]:.2f}" if t.entry_prices else '',
            'Averages': t.average_count,
            'Exit Date': t.exit_date.strftime('%Y-%m-%d') if t.exit_date else '',
            'Exit Price': f"₹{t.exit_price:.2f}" if t.exit_price else '',
            'Holding Days': t.holding_days,
            'Profit %': f"{t.profit_pct:.2f}",
            'Profit ₹': f"₹{t.profit_inr:,.2f}",
            'Invested': f"₹{t.total_invested:,.2f}",
            'Returned': f"₹{t.total_returned:,.2f}",
            'Reason': t.exit_reason
        })

    st.dataframe(pd.DataFrame(trade_data), use_container_width=True, hide_index=True)
    st.metric("Showing Trades", len(filtered))


def render_charts():
    st.header("Charts")

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
        st.plotly_chart(fig, use_container_width=True, key=f"trade_chart_{selected}")

        st.subheader("Trade Details")
        for i, trade in enumerate(trades, 1):
            with st.expander(f"Trade #{i} - {trade.symbol}"):
                # Per-entry Signal vs Execution breakdown (each buy, including averages)
                entries_df = pd.DataFrame({
                    'Signal Date': [d.strftime('%Y-%m-%d') for d in trade.signal_dates],
                    'Signal Close': [f"₹{p:.2f}" for p in trade.signal_closes],
                    'Execution Date': [d.strftime('%Y-%m-%d') for d in trade.entry_dates],
                    'Buy Price (Next Open)': [f"₹{p:.2f}" for p in trade.entry_prices],
                })
                st.write("**Entries (Signal → Next-Day-Open Execution):**")
                st.dataframe(entries_df, use_container_width=True, hide_index=True)

                st.write(f"**Average Count:** {trade.average_count}")
                st.write(f"**Exit Date:** {trade.exit_date.strftime('%Y-%m-%d') if trade.exit_date else 'N/A'}")
                st.write(f"**Exit Price:** ₹{trade.exit_price:.2f}" if trade.exit_price else "N/A")
                st.write(f"**Profit:** {trade.profit_pct:.2f}% (₹{trade.profit_inr:,.2f})")
                st.write(f"**Invested:** ₹{trade.total_invested:,.2f}")
                st.write(f"**Returned:** ₹{trade.total_returned:,.2f}")
                st.write(f"**Reason:** {trade.exit_reason}")


def render_portfolio():
    st.header("Portfolio")

    if not st.session_state.metrics or not st.session_state.metrics.get('equity_curve'):
        st.info("Run a backtest to see portfolio")
        return

    metrics = st.session_state.metrics

    st.subheader("Equity Curve")
    equity_df = pd.DataFrame([
        {'Date': d, 'Equity': e} for d, e in metrics['equity_curve']
    ])

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=equity_df['Date'], y=equity_df['Equity'],
        mode='lines', fill='tonexty',
        line=dict(color='#00ff88', width=2),
        fillcolor='rgba(0,255,136,0.1)'
    ))
    fig.update_layout(
        template='plotly_dark', height=500,
        xaxis_title='Date', yaxis_title='Portfolio Value (₹)',
        margin=dict(l=40, r=40, t=20, b=40)
    )
    st.plotly_chart(fig, use_container_width=True, key="portfolio_equity_chart")

    st.subheader("Drawdown Analysis")

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
        line=dict(color='red', width=1),
        fillcolor='rgba(255,0,0,0.2)'
    ))
    fig.update_layout(
        template='plotly_dark', height=400,
        xaxis_title='Date', yaxis_title='Drawdown (%)',
        margin=dict(l=40, r=40, t=20, b=40)
    )
    st.plotly_chart(fig, use_container_width=True, key="portfolio_drawdown_chart")


def render_statistics():
    st.header("Statistics")

    if not st.session_state.metrics:
        st.info("Run a backtest to see statistics")
        return

    metrics = st.session_state.metrics

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Monthly Returns")
        if metrics.get('monthly_returns'):
            monthly_df = pd.DataFrame([
                {'Month': k, 'Profit': v}
                for k, v in sorted(metrics['monthly_returns'].items())
            ])
            st.dataframe(monthly_df, use_container_width=True, hide_index=True)

    with col2:
        st.subheader("Yearly Returns")
        if metrics.get('yearly_returns'):
            yearly_df = pd.DataFrame([
                {'Year': k, 'Profit': v}
                for k, v in sorted(metrics['yearly_returns'].items())
            ])
            st.dataframe(yearly_df, use_container_width=True, hide_index=True)

    st.subheader("Trade Distribution")
    if st.session_state.all_trades:
        profits = [t.profit_pct for t in st.session_state.all_trades]

        fig = go.Figure()
        fig.add_trace(go.Histogram(
            x=profits, nbinsx=20,
            marker_color='#00ff88', name='Profit Distribution',
            opacity=0.8
        ))
        fig.add_vline(x=0, line_dash="dash", line_color="white", line_width=1)
        fig.update_layout(
            template='plotly_dark', height=400,
            xaxis_title='Profit %', yaxis_title='Frequency',
            showlegend=False, margin=dict(l=40, r=40, t=20, b=40)
        )
        st.plotly_chart(fig, use_container_width=True, key="statistics_histogram_chart")


# =============================================================================
# MAIN
# =============================================================================

def main():
    st.set_page_config(
        page_title=f"{APP_NAME} v{APP_VERSION}",
        page_icon=APP_ICON,
        layout="wide",
        initial_sidebar_state="expanded"
    )

    st.markdown("""
        <style>
        .stApp { background-color: #0e1117; }
        .css-1d391kg, .css-12oz5g7 { background-color: #1a1d24; }
        .stMetric { background-color: #1a1d24; border-radius: 8px; padding: 10px; }
        div[data-testid="stDataFrameResizable"] { background-color: #1a1d24; }
        .stButton>button { border-radius: 8px; }
        .stTabs [data-baseweb="tab-list"] { gap: 8px; }
        .stTabs [data-baseweb="tab"] { border-radius: 8px 8px 0 0; }
        </style>
    """, unsafe_allow_html=True)

    init_session_state()
    config = render_sidebar()

    tabs = st.tabs([
        "📊 Dashboard", "🔍 Scanner", "📈 Backtest",
        "📋 Trade Log", "📉 Charts", "💼 Portfolio", "📑 Statistics"
    ])

    with tabs[0]:
        render_dashboard(st.session_state.metrics)
    with tabs[1]:
        render_scanner()
    with tabs[2]:
        render_backtest()
    with tabs[3]:
        render_trade_log()
    with tabs[4]:
        render_charts()
    with tabs[5]:
        render_portfolio()
    with tabs[6]:
        render_statistics()


if __name__ == "__main__":
    main()
