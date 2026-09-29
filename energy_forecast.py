# ============================================================
# ⚡ Enerji Tüketim Tahmin Sistemi — Adaptif Rolling Forecast v4
# ============================================================
# v4 Değişiklikleri (v3'e göre):
#  1. Frekans otomatik tespiti — veri 15dk, 1 saat, 1 gün ya da
#     başka herhangi bir periyotta olabilir; kod veriden kendisi
#     anlar ve tüm lag/rolling/periyot sabitlerini buna göre
#     dinamik olarak ölçekler.
#  2. Excel/PNG çıktıları kaldırıldı — sonuçlar artık MySQL
#     veritabanına yazılıyor (.env dosyasından bağlantı bilgisi
#     okunur).
#  3. Strateji seçim mantığı v3 ile birebir aynı bırakıldı
#     (ONOFF fallback davranışı değiştirilmedi).
# ============================================================

import sys, os, warnings, logging  # EN İYİ SON VERSİYON
import pandas as pd
import numpy as np
from prophet import Prophet
import xgboost as xgb
from sklearn.metrics import mean_absolute_percentage_error
from dotenv import load_dotenv
from sqlalchemy import create_engine, text

warnings.filterwarnings('ignore')
logging.getLogger('prophet').setLevel(logging.ERROR)
logging.getLogger('cmdstanpy').setLevel(logging.ERROR)

# ══════════════════════════════════════════════════════════════
# AYARLAR — .env dosyasından okunur
# ══════════════════════════════════════════════════════════════
_script_dir = os.path.dirname(os.path.abspath(__file__))
_env_path   = os.path.join(_script_dir, '.env')

# Betiğin klasöründe .env yoksa çalışma dizininde de ara
if not os.path.exists(_env_path):
    _env_path = os.path.join(os.getcwd(), '.env')

if not os.path.exists(_env_path):
    print(f'❌ .env dosyası bulunamadı. Aranan yollar:\n'
          f'   {os.path.join(_script_dir, ".env")}\n'
          f'   {os.path.join(os.getcwd(), ".env")}')
    sys.exit(1)

print(f'ℹ️  .env dosyası: {_env_path}')
load_dotenv(_env_path, override=True)

MYSQL_HOST     = os.getenv('MYSQL_HOST', 'localhost')
MYSQL_PORT     = os.getenv('MYSQL_PORT', '3306')
MYSQL_USER     = os.getenv('MYSQL_USER', 'root')
MYSQL_PASSWORD = os.getenv('MYSQL_PASSWORD', '')
MYSQL_DATABASE = os.getenv('MYSQL_DATABASE', 'enerji_tahmin')

DATA_FILE_PATH = os.getenv('DATA_FILE_PATH', 'subat25- mayıs26.xlsx')
BACKTEST_AYLAR = int(os.getenv('BACKTEST_AYLAR', '2'))

# TAHMIN_AYI: .env'de YYYY-MM formatında girilmeli (örn. 2026-04)
_env_ay = os.getenv('TAHMIN_AYI', '').strip()
if not _env_ay:
    print('❌ .env dosyasında TAHMIN_AYI tanımlı değil. Örnek: TAHMIN_AYI=2026-04')
    sys.exit(1)
TAHMIN_AYI = _env_ay
print(f'ℹ️  TAHMIN_AYI: {TAHMIN_AYI}')

BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = DATA_FILE_PATH if os.path.isabs(DATA_FILE_PATH) else os.path.join(BASE_DIR, DATA_FILE_PATH)

if not os.path.exists(DATA_FILE):
    print(f'❌ Dosya bulunamadı: {DATA_FILE}')
    sys.exit(1)

# ══════════════════════════════════════════════════════════════
# MySQL BAĞLANTISI
# ══════════════════════════════════════════════════════════════
def get_engine():
    url = f"mysql+pymysql://{MYSQL_USER}:{MYSQL_PASSWORD}@{MYSQL_HOST}:{MYSQL_PORT}/{MYSQL_DATABASE}?charset=utf8mb4"
    return create_engine(url, pool_pre_ping=True)

def ensure_database_exists():
    """Veritabanı yoksa oluşturur (DATABASE belirtmeden bağlanıp CREATE DATABASE çalıştırır)."""
    url = f"mysql+pymysql://{MYSQL_USER}:{MYSQL_PASSWORD}@{MYSQL_HOST}:{MYSQL_PORT}/?charset=utf8mb4"
    eng = create_engine(url, pool_pre_ping=True)
    with eng.connect() as conn:
        conn.execute(text(
            f"CREATE DATABASE IF NOT EXISTS `{MYSQL_DATABASE}` "
            f"CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
        ))
        conn.commit()
    eng.dispose()

def ensure_tables_exist(engine):
    """Gerekli tabloları (yoksa) oluşturur."""
    ddl_statements = [
        """
        CREATE TABLE IF NOT EXISTS forecast_runs (
            run_id          BIGINT AUTO_INCREMENT PRIMARY KEY,
            tahmin_ayi      VARCHAR(7)   NOT NULL,
            freq_label      VARCHAR(20)  NOT NULL,
            freq_minutes    INT          NOT NULL,
            train_start     DATE         NOT NULL,
            train_end       DATE         NOT NULL,
            backtest_aylar  INT          NOT NULL,
            created_at      DATETIME     DEFAULT CURRENT_TIMESTAMP
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
        """,
        """
        CREATE TABLE IF NOT EXISTS forecast_detail (
            id              BIGINT AUTO_INCREMENT PRIMARY KEY,
            run_id          BIGINT       NOT NULL,
            machine         VARCHAR(128) NOT NULL,
            datetime_ts     DATETIME     NOT NULL,
            tahmin          DOUBLE       NOT NULL,
            alt_sinir       DOUBLE       NOT NULL,
            ust_sinir       DOUBLE       NOT NULL,
            gercek          DOUBLE       NULL,
            sapma           DOUBLE       NULL,
            INDEX idx_run_machine (run_id, machine),
            INDEX idx_machine_dt (machine, datetime_ts),
            CONSTRAINT fk_detail_run FOREIGN KEY (run_id) REFERENCES forecast_runs(run_id) ON DELETE CASCADE
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
        """,
        """
        CREATE TABLE IF NOT EXISTS forecast_daily (
            id              BIGINT AUTO_INCREMENT PRIMARY KEY,
            run_id          BIGINT       NOT NULL,
            machine         VARCHAR(128) NOT NULL,
            tarih           DATE         NOT NULL,
            tahmin_toplam   DOUBLE       NOT NULL,
            alt_sinir       DOUBLE       NOT NULL,
            ust_sinir       DOUBLE       NOT NULL,
            gercek_toplam   DOUBLE       NULL,
            hata_pct        DOUBLE       NULL,
            INDEX idx_run_machine_daily (run_id, machine),
            UNIQUE KEY uq_run_machine_tarih (run_id, machine, tarih),
            CONSTRAINT fk_daily_run FOREIGN KEY (run_id) REFERENCES forecast_runs(run_id) ON DELETE CASCADE
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
        """,
        """
        CREATE TABLE IF NOT EXISTS forecast_summary (
            id              BIGINT AUTO_INCREMENT PRIMARY KEY,
            run_id          BIGINT       NOT NULL,
            machine         VARCHAR(128) NOT NULL,
            strateji        VARCHAR(20)  NOT NULL,
            durum           VARCHAR(5)   NULL,
            gercek_kwh      DOUBLE       NULL,
            tahmin_kwh      DOUBLE       NOT NULL,
            fark_kwh        DOUBLE       NULL,
            hata_pct        DOUBLE       NULL,
            mape_pct        DOUBLE       NULL,
            zero_pct        DOUBLE       NOT NULL,
            cv_pct          DOUBLE       NOT NULL,
            bt_xgb          DOUBLE       NULL,
            bt_hybrid       DOUBLE       NULL,
            bt_rolling      DOUBLE       NULL,
            bt_onoff        DOUBLE       NULL,
            alt_sinir_toplam DOUBLE      NULL,
            ust_sinir_toplam DOUBLE      NULL,
            eğitim_baslangic DATE        NULL,
            eğitim_bitis     DATE        NULL,
            UNIQUE KEY uq_run_machine (run_id, machine),
            CONSTRAINT fk_summary_run FOREIGN KEY (run_id) REFERENCES forecast_runs(run_id) ON DELETE CASCADE
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
        """,
    ]
    with engine.connect() as conn:
        for ddl in ddl_statements:
            conn.execute(text(ddl))
        conn.commit()


# ══════════════════════════════════════════════════════════════
# FREKANS OTOMATİK TESPİTİ
# ══════════════════════════════════════════════════════════════
print(f'{"="*60}')
print(f'  Tahmin ayı : {TAHMIN_AYI}')
print(f'  Veri       : {os.path.basename(DATA_FILE)}')
print(f'  Backtest   : son {BACKTEST_AYLAR} ay')
print(f'{"="*60}')

# ══════════════════════════════════════════════════════════════
# FREKANS OTOMATİK TESPİTİ
# ══════════════════════════════════════════════════════════════
def detect_frequency(df_all):
    diffs = df_all['DateTime'].sort_values().diff().dropna()
    diffs = diffs[diffs > pd.Timedelta(0)]
    if len(diffs) == 0:
        raise ValueError('Frekans tespiti için yeterli veri yok.')

    delta = diffs.mode().iloc[0] if len(diffs.mode()) > 0 else diffs.median()
    freq_minutes = int(round(delta.total_seconds() / 60))

    if freq_minutes <= 0:
        raise ValueError('Tespit edilen frekans 0 veya negatif; veri zaman damgalarını kontrol edin.')

    if freq_minutes % (24 * 60) == 0 and freq_minutes >= 24 * 60:
        days = freq_minutes // (24 * 60)
        pandas_freq = f'{days}D'
        freq_label  = 'Günlük' if days == 1 else f'{days} Günlük'
    elif freq_minutes % 60 == 0:
        hours = freq_minutes // 60
        pandas_freq = f'{hours}h'
        freq_label  = 'Saatlik' if hours == 1 else f'{hours} Saatlik'
    else:
        pandas_freq = f'{freq_minutes}min'
        freq_label  = f'{freq_minutes} Dakikalık'

    periods_per_day = max(1, round((24 * 60) / freq_minutes))
    return pandas_freq, freq_minutes, periods_per_day, freq_label

# ── Veri okuma ───────────────────────────────────────────────
def parse_dt(series):
    # Önce boşlukları normalize et: "2025-02-01 00: 15" → "2025-02-01 00:15"
    cleaned = (
        series.astype(str)
        .str.replace(r':\s+', ':', regex=True)   # "HH: MM" → "HH:MM"
        .str.strip()
    )
    parsed = pd.to_datetime(cleaned, errors='coerce')
    # Hâlâ çok fazla NaT varsa farklı bir format dene
    if parsed.isna().mean() > 0.5:
        parsed = pd.to_datetime(cleaned, infer_datetime_format=True, errors='coerce')
    return parsed

print('\n📂 Veri okunuyor...')
ext = os.path.splitext(DATA_FILE)[1].lower()
if ext in ('.xlsx', '.xls'):
    df_all = pd.read_excel(DATA_FILE, header=0)
elif ext == '.csv':
    df_all = pd.read_csv(DATA_FILE)
else:
    print(f'❌ Desteklenmeyen dosya türü: {ext}')
    sys.exit(1)

df_all = df_all.rename(columns={df_all.columns[0]: 'DateTime'})
df_all['DateTime'] = parse_dt(df_all['DateTime'])
df_all = df_all.dropna(subset=['DateTime']).sort_values('DateTime').reset_index(drop=True)
machine_cols = [c for c in df_all.columns if c != 'DateTime']

if len(df_all) < 2:
    print('❌ Eğitim için yeterli veri yok.')
    sys.exit(1)

# ── Frekansı veriden otomatik tespit et ─────────────────────
FREQ, FREQ_MINUTES, PERIODS_DAY, FREQ_LABEL = detect_frequency(df_all)
print(f'\n🔍 Tespit edilen veri frekansı: {FREQ_LABEL} ({FREQ_MINUTES} dakika, periyot/gün={PERIODS_DAY})')

# Referans adımlar 15dk frekansına göre tanımlanmıştı; yeni frekansa oranla.
_BASE_FREQ_MIN  = 15
_BASE_LAGS      = [1, 2, 4, 96, 192, 672]
_BASE_ROLLS     = [96, 672]

def _scale(base_steps):
    ratio = _BASE_FREQ_MIN / FREQ_MINUTES
    return sorted(set(max(1, int(round(s * ratio))) for s in base_steps))

# LAG_STEPS / ROLL_WINS / LAG_1Y_STEPS / LAG_2W_STEPS
# veri boyutuna göre train split'ten sonra hesaplanır (aşağıda)

tp          = pd.Period(TAHMIN_AYI, freq='M')
train_end   = (tp - 1).end_time
# Veride kaç aylık geçmiş varsa ona göre eğitim başlangıcını belirle
# (12 ay isteriz ama verinin başından daha geriye gidemeyiz)
ideal_train_start = (tp - 13).start_time
actual_data_start = df_all['DateTime'].min()
train_start = max(ideal_train_start, actual_data_start)

df_train = df_all[
    (df_all['DateTime'] >= train_start) &
    (df_all['DateTime'] <= train_end)
].reset_index(drop=True)

df_test = df_all[
    (df_all['DateTime'].dt.year  == tp.year) &
    (df_all['DateTime'].dt.month == tp.month)
].reset_index(drop=True)

# Kaç aylık eğitim verisi var?
n_train_months = max(1, round((train_end - train_start).days / 30))

# Backtest: en fazla (eğitim ay sayısı - 1) ay geriye gidebiliriz
# ve en az 1 aylık eğitim verisi kalsın
BACKTEST_AYLAR_EFF = min(BACKTEST_AYLAR, max(0, n_train_months - 1))
if BACKTEST_AYLAR_EFF < BACKTEST_AYLAR:
    print(f'   ⚠️  Veri kısa ({n_train_months} ay eğitim) → backtest {BACKTEST_AYLAR_EFF} aya indirildi')

# Lag adımlarını veri boyutuyla sınırla
max_lag = max(1, len(df_train) - 1)
LAG_STEPS    = [l for l in _scale(_BASE_LAGS)  if l < max_lag]
ROLL_WINS    = [w for w in _scale(_BASE_ROLLS) if w < max_lag]
LAG_1Y_STEPS = min(max(1, round(365 * PERIODS_DAY)), max_lag)
LAG_2W_STEPS = min(max(1, round(14  * PERIODS_DAY)), max_lag)

# Minimum 1 lag adımı garantisi
if not LAG_STEPS:
    LAG_STEPS = [1]
if not ROLL_WINS:
    ROLL_WINS = [max(1, len(df_train) // 4)]

print(f'\n   Tüm veri  : {df_all["DateTime"].min().date()} → {df_all["DateTime"].max().date()}')
print(f'   Eğitim    : {df_train["DateTime"].min().date()} → {df_train["DateTime"].max().date()} ({len(df_train)} periyot, {n_train_months} ay)')
print(f'   Test      : {df_test["DateTime"].min().date()} → {df_test["DateTime"].max().date()} ({len(df_test)} periyot)')

if len(df_test) == 0:
    print(f'\n❌ {TAHMIN_AYI} için gerçek veri yok.')
    sys.exit(1)

if len(df_train) == 0:
    print(f'\n❌ {TAHMIN_AYI} öncesinde eğitim verisi yok.')
    sys.exit(1)

# ── Makine profili ────────────────────────────────────────────
def machine_profile(df, machine):
    vals    = pd.to_numeric(df[machine], errors='coerce').fillna(0).clip(lower=0)
    monthly = vals.groupby(df['DateTime'].dt.to_period('M')).sum()
    # std() için en az 2 ay gerekir; tek ay varsa CV hesaplanamaz → 999
    if len(monthly) >= 2 and monthly.mean() > 0:
        cv = monthly.std() / monthly.mean() * 100
    else:
        cv = 999
    zero    = (vals == 0).mean() * 100
    on_mean = vals[vals > 0].mean() if (vals > 0).any() else 0.0
    return {'cv': cv, 'zero': zero, 'on_mean': on_mean}

def get_strategy_rule(profile):
    """v3 ile birebir aynı kural tabanlı fallback (değiştirilmedi)."""
    if profile['zero'] > 60:
        return 'ONOFF'
    elif profile['zero'] > 40:
        return 'ROLLING'
    elif profile['cv'] < 10:
        return 'XGB'
    elif profile['cv'] < 25:
        return 'HYBRID'
    else:
        return 'ROLLING'

# ── Seri çıkarma ─────────────────────────────────────────────
def get_series(wide_df, machine):
    sub = wide_df[['DateTime', machine]].copy()
    sub[machine] = pd.to_numeric(sub[machine], errors='coerce').fillna(0).clip(lower=0)
    sub = sub.set_index('DateTime').resample(FREQ)[machine].mean().fillna(0).reset_index()
    sub.columns = ['DateTime', 'Consumption']
    return sub

# ══════════════════════════════════════════════════════════════
# ADAPTİF HİPERPARAMETRE SEÇİMİ
# ══════════════════════════════════════════════════════════════
def get_xgb_params(profile):
    cv   = profile['cv']
    zero = profile['zero']
    if zero > 40:
        return dict(n_estimators=300, max_depth=4, learning_rate=0.05,
                    subsample=0.7, colsample_bytree=0.7, min_child_weight=5,
                    reg_alpha=0.3, reg_lambda=2.0, objective='reg:squarederror',
                    tree_method='hist', random_state=42, n_jobs=-1)
    elif cv < 10:
        return dict(n_estimators=300, max_depth=4, learning_rate=0.05,
                    subsample=0.8, colsample_bytree=0.8, min_child_weight=3,
                    reg_alpha=0.05, reg_lambda=0.5, objective='reg:squarederror',
                    tree_method='hist', random_state=42, n_jobs=-1)
    elif cv < 25:
        return dict(n_estimators=500, max_depth=6, learning_rate=0.03,
                    subsample=0.8, colsample_bytree=0.8, min_child_weight=3,
                    reg_alpha=0.1, reg_lambda=1.0, objective='reg:squarederror',
                    tree_method='hist', random_state=42, n_jobs=-1)
    else:
        return dict(n_estimators=600, max_depth=7, learning_rate=0.02,
                    subsample=0.75, colsample_bytree=0.75, min_child_weight=4,
                    reg_alpha=0.2, reg_lambda=1.5, objective='reg:squarederror',
                    tree_method='hist', random_state=42, n_jobs=-1)

# ── XGBoost özellik üretimi (frekansa göre dinamik lag/rolling) ─
def build_xgb_features(series):
    f   = pd.DataFrame(index=series.index)
    idx = pd.DatetimeIndex(series['DateTime'])

    f['hour']       = idx.hour
    f['minute']     = idx.minute
    f['slot']       = (idx.hour * 60 + idx.minute) // FREQ_MINUTES
    f['dow']        = idx.dayofweek
    f['month']      = idx.month
    f['is_weekend'] = (idx.dayofweek >= 5).astype(int)
    f['hour_sin']   = np.sin(2 * np.pi * idx.hour / 24)
    f['hour_cos']   = np.cos(2 * np.pi * idx.hour / 24)
    f['slot_sin']   = np.sin(2 * np.pi * f['slot'] / max(PERIODS_DAY, 1))
    f['slot_cos']   = np.cos(2 * np.pi * f['slot'] / max(PERIODS_DAY, 1))
    f['dow_sin']    = np.sin(2 * np.pi * idx.dayofweek / 7)
    f['dow_cos']    = np.cos(2 * np.pi * idx.dayofweek / 7)
    f['month_sin']  = np.sin(2 * np.pi * idx.month / 12)
    f['month_cos']  = np.cos(2 * np.pi * idx.month / 12)

    s = series['Consumption']
    for lag in LAG_STEPS:
        f[f'lag_{lag}'] = s.shift(lag).values
    s1 = s.shift(1)
    for w in ROLL_WINS:
        f[f'roll_mean_{w}'] = s1.rolling(w, min_periods=1).mean().values
        f[f'roll_std_{w}']  = s1.rolling(w, min_periods=1).std().fillna(0).values

    f['same_slot_1y']   = s.shift(LAG_1Y_STEPS).values
    f['same_slot_2w']   = s.shift(LAG_2W_STEPS).values
    f['same_slot_mean'] = (f['same_slot_1y'].fillna(0) + f['same_slot_2w'].fillna(0)) / 2

    return f

def train_xgb(df_m, profile=None):
    params    = get_xgb_params(profile) if profile else dict(
        n_estimators=500, max_depth=6, learning_rate=0.03,
        subsample=0.8, colsample_bytree=0.8, min_child_weight=3,
        reg_alpha=0.1, reg_lambda=1.0, objective='reg:squarederror',
        tree_method='hist', random_state=42, n_jobs=-1
    )
    feat      = build_xgb_features(df_m).dropna()
    feat_cols = feat.columns.tolist()
    y         = df_m.loc[feat.index, 'Consumption']

    val_cut   = feat.index[-(max(1, len(feat) // 7))]
    X_tr,  y_tr  = feat[feat.index <= val_cut], y[feat.index <= val_cut]
    X_val, y_val = feat[feat.index >  val_cut], y[feat.index >  val_cut]

    model = xgb.XGBRegressor(**params)
    if len(X_val) > 0:
        model.fit(X_tr, y_tr, eval_set=[(X_val, y_val)], verbose=False)
    else:
        model.fit(X_tr, y_tr, verbose=False)
    return model, feat_cols

# ══════════════════════════════════════════════════════════════
# RECURSİVE XGB TAHMİNİ (blok boyutu artık frekansa göre 1 gün)
# ══════════════════════════════════════════════════════════════
def predict_xgb_recursive(model, feat_cols, df_m, future_index):
    BLOCK = max(1, PERIODS_DAY)  # 1 günlük blok

    combined = pd.concat([
        df_m,
        pd.DataFrame({'DateTime': future_index, 'Consumption': np.nan})
    ], ignore_index=True).reset_index(drop=True)

    preds_dict = {}
    n_future = len(future_index)
    n_blocks  = int(np.ceil(n_future / BLOCK))

    for b in range(n_blocks):
        start_i = b * BLOCK
        end_i   = min(start_i + BLOCK, n_future)
        block_idx = list(range(start_i, end_i))

        combined_len = len(df_m)
        row_indices  = [combined_len + i for i in block_idx]

        feat_all = build_xgb_features(combined)

        for ci, ri in zip(block_idx, row_indices):
            feat_row = feat_all.iloc[[ri]].copy()
            for col in feat_cols:
                if col not in feat_row.columns:
                    feat_row[col] = 0.0
            feat_row = feat_row[feat_cols].fillna(0)

            pred_val = float(model.predict(feat_row.values).clip(0)[0])
            preds_dict[future_index[ci]] = pred_val
            combined.at[ri, 'Consumption'] = pred_val

    preds = pd.Series(
        [preds_dict[ts] for ts in future_index],
        index=future_index, name='yhat'
    )
    return preds

def predict_xgb(model, feat_cols, df_m, future_index):
    future_stub = pd.DataFrame({'DateTime': future_index, 'Consumption': 0.0})
    combined    = pd.concat([df_m, future_stub], ignore_index=True).reset_index(drop=True)
    feat_all    = build_xgb_features(combined)
    feat_fut    = feat_all.iloc[-len(future_index):].copy()
    for col in feat_cols:
        if col not in feat_fut.columns:
            feat_fut[col] = 0.0
    feat_fut = feat_fut[feat_cols].fillna(0)
    preds    = model.predict(feat_fut.values).clip(0)
    return pd.Series(preds, index=future_index, name='yhat')

# ── Prophet ──────────────────────────────────────────────────
def train_prophet(df_m, zero_pct=0):
    pdf = df_m.rename(columns={'DateTime': 'ds', 'Consumption': 'y'}).copy()
    cps  = 0.05 if zero_pct > 30 else 0.10
    m = Prophet(
        seasonality_mode='additive',
        daily_seasonality=(PERIODS_DAY > 1),
        weekly_seasonality=True,
        yearly_seasonality=True,
        changepoint_prior_scale=cps,
        uncertainty_samples=200
    )
    m.fit(pdf)
    return m

def predict_prophet(model, df_m, future_index):
    end     = future_index[-1]
    last    = df_m['DateTime'].max()
    periods = max(
        int((end - last).total_seconds() / 60 / FREQ_MINUTES) + PERIODS_DAY * 2,
        31 * PERIODS_DAY
    )
    future  = model.make_future_dataframe(periods=periods, freq=FREQ)
    fc      = model.predict(future)
    mask    = (fc['ds'] >= future_index[0]) & (fc['ds'] <= future_index[-1])
    fc_f    = fc[mask].set_index('ds')
    yhat    = fc_f['yhat'].reindex(future_index).fillna(0).clip(lower=0)
    yhat_lo = fc_f['yhat_lower'].reindex(future_index).fillna(0).clip(lower=0)
    yhat_hi = fc_f['yhat_upper'].reindex(future_index).fillna(0).clip(lower=0)
    return yhat, yhat_lo, yhat_hi

# ── On/Off sınıflandırıcı + regresör ─────────────────────────
def train_onoff(df_m, profile=None):
    feat      = build_xgb_features(df_m).dropna()
    feat_cols = feat.columns.tolist()
    s         = df_m.loc[feat.index, 'Consumption']
    y_cls     = (s > 0).astype(int).values
    y_reg     = s[s > 0].values

    n_off = (y_cls == 0).sum()
    n_on  = (y_cls == 1).sum()
    spw   = n_off / n_on if n_on > 0 else 1.0

    clf = xgb.XGBClassifier(
        n_estimators=200, max_depth=5, learning_rate=0.05,
        scale_pos_weight=spw,
        random_state=42, n_jobs=-1, verbosity=0
    )
    clf.fit(feat.values, y_cls)

    feat_on = feat[s > 0]
    if len(feat_on) == 0:
        # Hiç aktif periyot yoksa regresör eğitilemez; sıfır tahmin döndürecek basit bir model kur.
        reg = xgb.XGBRegressor(n_estimators=10, max_depth=2, random_state=42, n_jobs=-1, verbosity=0)
        reg.fit(feat.values, np.zeros(len(feat)))
    else:
        reg = xgb.XGBRegressor(
            n_estimators=300, max_depth=5, learning_rate=0.05,
            random_state=42, n_jobs=-1, verbosity=0
        )
        reg.fit(feat_on.values, y_reg)
    return clf, reg, feat_cols

def predict_onoff(clf, reg, feat_cols, df_m, future_index):
    future_stub = pd.DataFrame({'DateTime': future_index, 'Consumption': 0.0})
    combined    = pd.concat([df_m, future_stub], ignore_index=True).reset_index(drop=True)
    feat_all    = build_xgb_features(combined)
    feat_fut    = feat_all.iloc[-len(future_index):].copy()
    for col in feat_cols:
        if col not in feat_fut.columns:
            feat_fut[col] = 0.0
    feat_mat = feat_fut[feat_cols].fillna(0).values

    prob_on = clf.predict_proba(feat_mat)[:, 1]
    is_on   = prob_on >= 0.35
    preds   = np.where(is_on, reg.predict(feat_mat).clip(0), 0.0)
    return pd.Series(preds, index=future_index, name='yhat')

# ── Fallback: geçen yıl ──────────────────────────────────────
def fallback_last_year(df_m, future_index):
    ref_start = future_index[0]  - pd.DateOffset(years=1)
    ref_end   = future_index[-1] - pd.DateOffset(years=1)
    ref_mask  = (df_m['DateTime'] >= ref_start) & (df_m['DateTime'] <= ref_end)
    ref_s     = df_m[ref_mask].set_index('DateTime')['Consumption']
    yhat = (
        ref_s.resample(FREQ).mean()
        .reindex(future_index - pd.DateOffset(years=1))
        .ffill()
        .fillna(0)
        .clip(0)
    )
    yhat.index = future_index
    return yhat

# ══════════════════════════════════════════════════════════════
# DİNAMİK BLEND AĞIRLIKLARI
# ══════════════════════════════════════════════════════════════
def dynamic_rolling_blend(prophet_pred, xgb_pred, df_m, future_index,
                           w_prophet=0.35, w_xgb=0.40, w_lastyear=0.25):
    ref_start  = future_index[0]  - pd.DateOffset(years=1)
    ref_end    = future_index[-1] - pd.DateOffset(years=1)
    ref_mask   = (df_m['DateTime'] >= ref_start) & (df_m['DateTime'] <= ref_end)
    ref_series = df_m[ref_mask].set_index('DateTime')['Consumption']

    recent_2m_start = future_index[0] - pd.DateOffset(months=2)
    recent_1m_start = future_index[0] - pd.DateOffset(months=1)
    last_2m = df_m[df_m['DateTime'] >= recent_2m_start]['Consumption'].mean()
    last_1m = df_m[df_m['DateTime'] >= recent_1m_start]['Consumption'].mean()
    trend_factor = np.clip((last_1m / last_2m) if last_2m > 0 else 1.0, 0.7, 1.3)

    ref_resampled = (
        ref_series.resample(FREQ).mean()
        .reindex(future_index - pd.DateOffset(years=1))
        .ffill()
        .fillna(0)
    )
    ref_scaled = (ref_resampled.values * trend_factor).clip(0)

    blended = w_lastyear * ref_scaled + w_prophet * prophet_pred.values + w_xgb * xgb_pred.values
    return pd.Series(blended.clip(0), index=future_index, name='yhat')

# ══════════════════════════════════════════════════════════════
# BACKTEST İLE OTOMATİK MODEL SEÇİMİ
# ══════════════════════════════════════════════════════════════
STRATEGIES_TO_TEST = ['XGB', 'HYBRID', 'ROLLING', 'ONOFF']

def run_single_strategy(strategy, df_m_bt, future_idx_bt, profile):
    try:
        if strategy == 'XGB':
            model, feat_cols = train_xgb(df_m_bt, profile)
            yhat = predict_xgb_recursive(model, feat_cols, df_m_bt, future_idx_bt)

        elif strategy == 'ONOFF':
            if profile['zero'] <= 20:
                return np.nan
            clf, reg, feat_cols = train_onoff(df_m_bt, profile)
            yhat = predict_onoff(clf, reg, feat_cols, df_m_bt, future_idx_bt)

        elif strategy == 'HYBRID':
            p_model = train_prophet(df_m_bt, zero_pct=profile['zero'])
            yhat, _, _ = predict_prophet(p_model, df_m_bt, future_idx_bt)

        elif strategy == 'ROLLING':
            p_model = train_prophet(df_m_bt, zero_pct=profile['zero'])
            yhat_p, yhat_lo, yhat_hi = predict_prophet(p_model, df_m_bt, future_idx_bt)
            xgb_m, feat_cols = train_xgb(df_m_bt, profile)
            yhat_x = predict_xgb_recursive(xgb_m, feat_cols, df_m_bt, future_idx_bt)
            yhat   = dynamic_rolling_blend(yhat_p, yhat_x, df_m_bt, future_idx_bt)
        else:
            return np.nan

        return yhat

    except Exception:
        return None

def backtest_model_selection(df_m_full, machine, profile, bt_aylar=2):
    scores = {s: [] for s in STRATEGIES_TO_TEST}
    rolling_preds  = []

    for offset in range(1, bt_aylar + 1):
        bt_tp      = tp - offset
        bt_end     = (bt_tp - 1).end_time
        bt_start   = (bt_tp - 13).start_time

        bt_train = df_m_full[
            (df_m_full['DateTime'] >= bt_start) &
            (df_m_full['DateTime'] <= bt_end)
        ].reset_index(drop=True)

        bt_real_df = df_m_full[
            (df_m_full['DateTime'].dt.year  == bt_tp.year) &
            (df_m_full['DateTime'].dt.month == bt_tp.month)
        ].reset_index(drop=True)

        # Minimum eğitim: en az 1 aylık veri (veri kısaysa 1 haftaya kadar düşür)
        min_train = max(PERIODS_DAY * 7, min(PERIODS_DAY * 30, len(bt_train) // 2))
        if len(bt_train) < min_train or len(bt_real_df) == 0:
            continue

        bt_idx = pd.date_range(
            bt_real_df['DateTime'].min(),
            bt_real_df['DateTime'].max(),
            freq=FREQ
        )
        real_s = bt_real_df.set_index('DateTime')['Consumption']

        for strategy in STRATEGIES_TO_TEST:
            yhat = run_single_strategy(strategy, bt_train, bt_idx, profile)
            if yhat is None or not isinstance(yhat, pd.Series):
                scores[strategy].append(np.nan)
                continue

            merged = pd.DataFrame({'yhat': yhat.values, 'real': real_s.reindex(bt_idx).fillna(0).values})
            mask_p = merged['real'] > 0.001
            if mask_p.sum() > 10:
                mape = mean_absolute_percentage_error(
                    merged.loc[mask_p, 'real'],
                    merged.loc[mask_p, 'yhat'].clip(lower=0)
                ) * 100
                scores[strategy].append(mape)
            else:
                scores[strategy].append(np.nan)

        try:
            p_model_bt = train_prophet(bt_train, zero_pct=profile['zero'])
            yhat_p_bt, _, _ = predict_prophet(p_model_bt, bt_train, bt_idx)
            xgb_bt, fc_bt = train_xgb(bt_train, profile)
            yhat_x_bt = predict_xgb_recursive(xgb_bt, fc_bt, bt_train, bt_idx)
            real_vals  = real_s.reindex(bt_idx).fillna(0).values
            rolling_preds.append((yhat_p_bt.values, yhat_x_bt.values, real_vals, bt_train, bt_idx))
        except Exception:
            pass

    avg_scores = {}
    for s, sc_list in scores.items():
        valid = [x for x in sc_list if not np.isnan(x)]
        avg_scores[s] = np.mean(valid) if valid else 999.0

    best_strategy = min(avg_scores, key=avg_scores.get)
    best_score    = avg_scores[best_strategy]

    rule_strategy = get_strategy_rule(profile)

    # Hiç backtest yapılamadıysa veya tüm skorlar 999 ise direkt kural tabanlı
    if all(v >= 999 for v in avg_scores.values()):
        best_strategy = rule_strategy
        print(f'       ⚠️  Backtest yapılamadı (yetersiz veri) → kural tabanlı: {rule_strategy}')
    elif best_score > 50:
        best_strategy = rule_strategy
        print(f'       ⚠️  Backtest skoru yüksek (%{best_score:.0f}) → kural tabanlı: {rule_strategy}')

    opt_weights = None
    if best_strategy == 'ROLLING' and len(rolling_preds) >= 1:
        opt_weights = optimize_blend_weights(rolling_preds)

    return best_strategy, avg_scores, opt_weights

def optimize_blend_weights(rolling_preds_list):
    best_mape   = np.inf
    best_weights = (0.35, 0.40, 0.25)

    for w_p in np.arange(0.1, 0.65, 0.1):
        for w_x in np.arange(0.1, 0.65, 0.1):
            w_ly = round(1.0 - w_p - w_x, 2)
            if w_ly < 0.1 or w_ly > 0.7:
                continue

            total_mape = 0
            count      = 0
            for (yhat_p, yhat_x, real_vals, bt_train, bt_idx) in rolling_preds_list:
                yhat_ly = fallback_last_year(bt_train, bt_idx).values
                blended = w_ly * yhat_ly + w_p * yhat_p + w_x * yhat_x
                mask_p  = real_vals > 0.001
                if mask_p.sum() > 10:
                    m = mean_absolute_percentage_error(real_vals[mask_p], blended[mask_p].clip(0)) * 100
                    total_mape += m
                    count      += 1

            if count > 0 and (total_mape / count) < best_mape:
                best_mape    = total_mape / count
                best_weights = (round(w_p, 1), round(w_x, 1), round(w_ly, 1))

    return best_weights

# ═══════════════════════════════════════════════════════════
# ANA DÖNGÜ
# ═══════════════════════════════════════════════════════════
print(f'\n📐 Strateji: Rolling pencere {train_start.date()} → {train_end.date()}')
print(f'   Test    : {TAHMIN_AYI}')
print(f'   Backtest: son {BACKTEST_AYLAR} ay (otomatik model seçimi)\n')

future_index = pd.date_range(
    df_test['DateTime'].min(),
    df_test['DateTime'].max(),
    freq=FREQ
)

results = {}
total   = len(machine_cols)

for i, machine in enumerate(machine_cols, 1):
    print(f'[{i}/{total}] {machine}')

    prof     = machine_profile(df_train, machine)
    df_m     = get_series(df_train, machine)
    df_r     = get_series(df_test,  machine)
    df_m_all = get_series(df_all, machine)

    print(f'       zero={prof["zero"]:.0f}% | CV={prof["cv"]:.0f}% | XGB params: adaptif')

    print(f'       🔍 Backtest çalışıyor ({BACKTEST_AYLAR} ay)...')
    strategy, bt_scores, opt_weights = backtest_model_selection(
        df_m_all, machine, prof, bt_aylar=BACKTEST_AYLAR_EFF
    )
    scores_str = ' | '.join([f'{s}=%{v:.0f}' for s, v in bt_scores.items() if v < 999])
    print(f'       Backtest → {scores_str}')
    print(f'       ✅ Seçilen strateji: {strategy}', end='')
    if opt_weights and strategy == 'ROLLING':
        print(f' | Blend: P={opt_weights[0]} X={opt_weights[1]} LY={opt_weights[2]}', end='')
    print()

    yhat    = None
    yhat_lo = None
    yhat_hi = None

    try:
        if strategy == 'XGB':
            model, feat_cols = train_xgb(df_m, prof)
            yhat    = predict_xgb_recursive(model, feat_cols, df_m, future_index)
            yhat_lo = (yhat * 0.90).clip(0)
            yhat_hi = (yhat * 1.10).clip(0)

        elif strategy == 'ONOFF':
            clf, reg, feat_cols = train_onoff(df_m, prof)
            yhat    = predict_onoff(clf, reg, feat_cols, df_m, future_index)
            yhat_lo = (yhat * 0.85).clip(0)
            yhat_hi = (yhat * 1.15).clip(0)

        elif strategy == 'HYBRID':
            p_model = train_prophet(df_m, zero_pct=prof['zero'])
            yhat, yhat_lo, yhat_hi = predict_prophet(p_model, df_m, future_index)

        elif strategy == 'ROLLING':
            p_model = train_prophet(df_m, zero_pct=prof['zero'])
            yhat_p, yhat_lo, yhat_hi = predict_prophet(p_model, df_m, future_index)
            xgb_m, feat_cols = train_xgb(df_m, prof)
            yhat_x  = predict_xgb_recursive(xgb_m, feat_cols, df_m, future_index)

            if opt_weights:
                w_p, w_x, w_ly = opt_weights
            else:
                w_p, w_x, w_ly = 0.35, 0.40, 0.25

            yhat    = dynamic_rolling_blend(yhat_p, yhat_x, df_m, future_index,
                                            w_prophet=w_p, w_xgb=w_x, w_lastyear=w_ly)
            band_w  = (yhat_hi - yhat_lo) * 0.5
            yhat_lo = (yhat - band_w).clip(lower=0)
            yhat_hi = (yhat + band_w).clip(lower=0)

    except Exception as e:
        print(f'       ⚠️  Hata: {e} → fallback: geçen yıl aynı ay')
        yhat    = fallback_last_year(df_m, future_index)
        yhat_lo = (yhat * 0.85).clip(0)
        yhat_hi = (yhat * 1.15).clip(0)

    # Son güvenlik ağı: her ihtimale karşı yhat hâlâ None ise sıfır seri oluştur.
    if yhat is None:
        yhat    = pd.Series(0.0, index=future_index, name='yhat')
        yhat_lo = pd.Series(0.0, index=future_index, name='yhat')
        yhat_hi = pd.Series(0.0, index=future_index, name='yhat')

    total_pred = yhat.sum()
    total_real = df_r['Consumption'].sum()
    hata_pct   = abs(total_pred - total_real) / total_real * 100 if total_real > 0 else np.nan

    yhat_df = pd.DataFrame({'DateTime': yhat.index, 'yhat': yhat.values})
    merged  = pd.merge(yhat_df, df_r[['DateTime', 'Consumption']], on='DateTime', how='inner')
    if len(merged) > 0 and merged['Consumption'].sum() > 0:
        mask_pos = merged['Consumption'] > 0.001
        if mask_pos.sum() > 10:
            mape = mean_absolute_percentage_error(
                merged.loc[mask_pos, 'Consumption'],
                merged.loc[mask_pos, 'yhat'].clip(lower=0)
            ) * 100
        else:
            mape = np.nan
    else:
        mape = np.nan

    durum = '✅' if not np.isnan(hata_pct) and hata_pct <= 15 else \
            ('⚠️' if not np.isnan(hata_pct) and hata_pct <= 25 else '❌')

    print(f'       Tahmin={total_pred:.0f} | Gerçek={total_real:.0f} | Hata=%{hata_pct:.1f} | MAPE=%{mape:.1f} {durum}')

    fc_df = pd.DataFrame({
        'DateTime': future_index, 'yhat': yhat.values,
        'yhat_lo' : yhat_lo.values, 'yhat_hi': yhat_hi.values
    })
    daily_fc = (fc_df.assign(Tarih=fc_df['DateTime'].dt.date)
                .groupby('Tarih').agg(Tahmin=('yhat', 'sum'),
                                      Alt=('yhat_lo', 'sum'),
                                      Ust=('yhat_hi', 'sum'))
                .round(4).reset_index())
    daily_real = (df_r.assign(Tarih=df_r['DateTime'].dt.date)
                  .groupby('Tarih')['Consumption'].sum().reset_index()
                  .rename(columns={'Consumption': 'Gercek'}))

    results[machine] = {
        'yhat': yhat, 'yhat_lo': yhat_lo, 'yhat_hi': yhat_hi,
        'df_r': df_r, 'df_m': df_m,
        'daily_fc': daily_fc, 'daily_real': daily_real,
        'total_pred': total_pred, 'total_real': total_real,
        'hata_pct': hata_pct, 'mape': mape,
        'strategy': strategy, 'profile': prof, 'durum': durum,
        'bt_scores': bt_scores,
        'opt_weights': opt_weights
    }

print(f'\n✅ {len(results)} makine tamamlandı.')

# ══════════════════════════════════════════════════════════════
# GRAFİKLER
# ══════════════════════════════════════════════════════════════
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

DARK = {
    'figure.facecolor': '#0f1117', 'axes.facecolor': '#1a1d27',
    'axes.edgecolor': '#2d3045',   'axes.labelcolor': '#c8ccd8',
    'xtick.color': '#8b8fa8',      'ytick.color': '#8b8fa8',
    'text.color': '#c8ccd8',       'grid.color': '#2d3045',
    'grid.alpha': 0.6,             'font.family': 'monospace'
}
plt.rcParams.update(DARK)
C_FORE = '#4fc97f'
C_REAL = '#f7834f'
C_BAND = '#4fc97f'

OUTPUT_DIR = os.path.join(_script_dir, f'grafikler_{TAHMIN_AYI}')
os.makedirs(OUTPUT_DIR, exist_ok=True)

print(f'\n📊 Grafikler çiziliyor → {OUTPUT_DIR}')

for machine, r in results.items():
    daily_cmp = pd.merge(r['daily_fc'], r['daily_real'], on='Tarih', how='outer').sort_values('Tarih')

    fig, axes = plt.subplots(3, 1, figsize=(16, 14), facecolor='#0f1117',
                             gridspec_kw={'height_ratios': [1.4, 1, 0.8]})

    bt_str = ' | '.join([f'{s}:{v:.0f}%' for s, v in r['bt_scores'].items() if v < 999])
    fig.suptitle(
        f'{machine} — {TAHMIN_AYI}  [{r["strategy"]}]\n'
        f'Tahmin: {r["total_pred"]:.0f} kWh  |  Gerçek: {r["total_real"]:.0f} kWh  |  '
        f'Hata: %{r["hata_pct"]:.1f}  |  MAPE: %{r["mape"]:.1f}\n'
        f'Backtest → {bt_str}',
        fontsize=10, color='white', fontweight='bold', y=0.99
    )

    days = daily_cmp['Tarih'].values
    x, w = np.arange(len(days)), 0.38

    # Panel 1 — Günlük çubuk grafik
    ax = axes[0]
    ax.bar(x - w/2, daily_cmp['Tahmin'].fillna(0), width=w, color=C_FORE, alpha=0.85, label='Tahmin')
    ax.bar(x + w/2, daily_cmp['Gercek'].fillna(0), width=w, color=C_REAL, alpha=0.85, label='Gerçek')
    ax.fill_between(x, daily_cmp['Alt'].fillna(0), daily_cmp['Ust'].fillna(0),
                    alpha=0.12, color=C_BAND, label='Güven aralığı')
    ax.set_xticks(x)
    ax.set_xticklabels([str(d)[-5:] for d in days], rotation=45, ha='right', fontsize=7)
    ax.set_ylabel('Günlük kWh')
    ax.set_title(f'Günlük Tahmin vs Gerçek  (zero=%{r["profile"]["zero"]:.0f}, CV=%{r["profile"]["cv"]:.0f})',
                 fontsize=9, pad=8)
    ax.legend(fontsize=9)
    ax.grid(axis='y')

    # Panel 2 — Periyot bazlı zaman serisi
    ax = axes[1]
    ax.plot(r['yhat'].index, r['yhat'].values, color=C_FORE, lw=1.0, label='Tahmin', zorder=3)
    ax.fill_between(r['yhat'].index, r['yhat_lo'].values, r['yhat_hi'].values,
                    alpha=0.18, color=C_BAND, label='Güven aralığı')
    ax.plot(r['df_r']['DateTime'], r['df_r']['Consumption'],
            color=C_REAL, lw=0.9, alpha=0.85, label='Gerçek', zorder=4)
    ax.set_ylabel(f'{FREQ_LABEL} (kWh)')
    ax.set_title(f'{FREQ_LABEL} Tahmin vs Gerçek', fontsize=9, pad=8)
    ax.legend(fontsize=9)
    ax.grid(True)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d'))
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=max(1, len(days) // 10)))
    ax.tick_params(axis='x', rotation=30)

    # Panel 3 — Günlük hata çubuğu
    ax = axes[2]
    hata_arr = daily_cmp['Tahmin'].fillna(0) - daily_cmp['Gercek'].fillna(0)
    ax.bar(x, hata_arr, color=[C_FORE if v >= 0 else C_REAL for v in hata_arr], alpha=0.85, width=0.7)
    ax.axhline(0, color='white', lw=0.8, linestyle='--')
    ax.set_xticks(x)
    ax.set_xticklabels([str(d)[-5:] for d in days], rotation=45, ha='right', fontsize=7)
    ax.set_ylabel('Fark (kWh)')
    ax.set_title('Günlük Hata (yeşil=fazla tahmin, kırmızı=eksik)', fontsize=9, pad=8)
    ax.grid(axis='y')

    plt.tight_layout()
    safe_name = machine.replace('/', '_').replace('\\', '_').replace(':', '_')
    fname = os.path.join(OUTPUT_DIR, f'{safe_name}_{TAHMIN_AYI}_karsilastirma.png')
    fig.savefig(fname, dpi=120, bbox_inches='tight', facecolor='#0f1117')
    plt.close()
    print(f'   ✅ {machine}')

print(f'   Grafikler kaydedildi → {OUTPUT_DIR}')


print('\n🗄️  MySQL bağlantısı kuruluyor...')
try:
    ensure_database_exists()
    engine = get_engine()
    ensure_tables_exist(engine)
except Exception as e:
    print(f'❌ MySQL bağlantı/tablo hatası: {e}')
    print('   .env dosyasındaki MYSQL_HOST / MYSQL_USER / MYSQL_PASSWORD / MYSQL_DATABASE ayarlarını kontrol edin.')
    sys.exit(1)

print('   Bağlantı başarılı, tablolar hazır.')

with engine.connect() as conn:
    res = conn.execute(text("""
        INSERT INTO forecast_runs (tahmin_ayi, freq_label, freq_minutes, train_start, train_end, backtest_aylar)
        VALUES (:tahmin_ayi, :freq_label, :freq_minutes, :train_start, :train_end, :backtest_aylar)
    """), {
        'tahmin_ayi': TAHMIN_AYI,
        'freq_label': FREQ_LABEL,
        'freq_minutes': FREQ_MINUTES,
        'train_start': train_start.date(),
        'train_end': train_end.date(),
        'backtest_aylar': BACKTEST_AYLAR,
    })
    conn.commit()
    run_id = res.lastrowid

print(f'   Yeni run_id: {run_id}')

print('\n📥 Sonuçlar MySQL\'e yazılıyor...')
for machine, r in results.items():
    print(f'   → {machine}')

    # ── forecast_detail ──────────────────────────────────────
    real_lookup = r['df_r'].set_index('DateTime')['Consumption']
    detail_rows = []
    for ts, yh, ylo, yhi in zip(r['yhat'].index, r['yhat'].values, r['yhat_lo'].values, r['yhat_hi'].values):
        rv = real_lookup.get(ts, np.nan)
        if pd.isna(rv):
            gercek, sapma = None, None
        else:
            if rv < ylo:
                sapma = float(rv - ylo)
            elif rv > yhi:
                sapma = float(rv - yhi)
            else:
                sapma = None  # bulut içinde — boş
            gercek = float(rv)
        detail_rows.append({
            'run_id': run_id, 'machine': machine, 'datetime_ts': ts,
            'tahmin': float(yh), 'alt_sinir': float(ylo), 'ust_sinir': float(yhi),
            'gercek': gercek, 'sapma': sapma
        })
    detail_df = pd.DataFrame(detail_rows)
    detail_df.to_sql('forecast_detail', engine, if_exists='append', index=False, method='multi', chunksize=1000)

    # ── forecast_daily ───────────────────────────────────────
    daily_cmp = pd.merge(r['daily_fc'], r['daily_real'], on='Tarih', how='outer').sort_values('Tarih')
    daily_cmp['hata_pct'] = np.where(
        daily_cmp['Gercek'].fillna(0) > 0,
        (daily_cmp['Tahmin'].fillna(0) - daily_cmp['Gercek'].fillna(0)).abs() / daily_cmp['Gercek'] * 100,
        np.nan
    )
    daily_rows = [{
        'run_id': run_id, 'machine': machine, 'tarih': row['Tarih'],
        'tahmin_toplam': float(row['Tahmin']) if pd.notna(row['Tahmin']) else 0.0,
        'alt_sinir': float(row['Alt']) if pd.notna(row['Alt']) else 0.0,
        'ust_sinir': float(row['Ust']) if pd.notna(row['Ust']) else 0.0,
        'gercek_toplam': float(row['Gercek']) if pd.notna(row['Gercek']) else None,
        'hata_pct': float(row['hata_pct']) if pd.notna(row['hata_pct']) else None,
    } for _, row in daily_cmp.iterrows()]
    pd.DataFrame(daily_rows).to_sql('forecast_daily', engine, if_exists='append', index=False, method='multi', chunksize=1000)

    # ── forecast_summary ─────────────────────────────────────
    bt = r['bt_scores']
    summary_row = {
        'run_id': run_id, 'machine': machine, 'strateji': r['strategy'], 'durum': r['durum'],
        'gercek_kwh': float(r['total_real']), 'tahmin_kwh': float(r['total_pred']),
        'fark_kwh': float(r['total_pred'] - r['total_real']),
        'hata_pct': float(r['hata_pct']) if not np.isnan(r['hata_pct']) else None,
        'mape_pct': float(r['mape']) if not np.isnan(r['mape']) else None,
        'zero_pct': float(r['profile']['zero']), 'cv_pct': float(r['profile']['cv']),
        'bt_xgb': float(bt.get('XGB')) if bt.get('XGB', 999) < 999 else None,
        'bt_hybrid': float(bt.get('HYBRID')) if bt.get('HYBRID', 999) < 999 else None,
        'bt_rolling': float(bt.get('ROLLING')) if bt.get('ROLLING', 999) < 999 else None,
        'bt_onoff': float(bt.get('ONOFF')) if bt.get('ONOFF', 999) < 999 else None,
        'alt_sinir_toplam': float(r['daily_fc']['Alt'].sum()),
        'ust_sinir_toplam': float(r['daily_fc']['Ust'].sum()),
        'eğitim_baslangic': train_start.date(),
        'eğitim_bitis': train_end.date(),
    }
    pd.DataFrame([summary_row]).to_sql('forecast_summary', engine, if_exists='append', index=False, method='multi')

print(f'\n✅ Tüm sonuçlar MySQL\'e yazıldı (run_id={run_id}, veritabanı="{MYSQL_DATABASE}").')

# ══════════════════════════════════════════════════════════════
# KONSOL ÖZET
# ══════════════════════════════════════════════════════════════
print(f'\n{"="*60}')
print(f'  SONUÇLAR — {TAHMIN_AYI}  (Rolling: {train_start.date()} → {train_end.date()})')
print(f'{"="*60}')

rows_for_summary = []
for machine, r in results.items():
    rows_for_summary.append({
        'Makine': machine, 'Strateji': r['strategy'],
        'Hata (%)': r['hata_pct'], 'MAPE (%)': r['mape'], 'CV (%)': r['profile']['cv']
    })
df_out = pd.DataFrame(rows_for_summary).sort_values('Hata (%)')

h_vals = df_out['Hata (%)']
iyi    = df_out[h_vals <= 15]
orta   = df_out[(h_vals > 15) & (h_vals <= 25)]
kotu   = df_out[h_vals > 25]
print(f'  ✅ İyi   (≤%15)  : {len(iyi):2d} makine')
print(f'  ⚠️  Orta  (%15-25): {len(orta):2d} makine')
print(f'  ❌ Kötü  (>%25)  : {len(kotu):2d} makine')
print(f'\n  Ort. Hata : %{h_vals.mean():.1f}  |  Ort. MAPE: %{df_out["MAPE (%)"].mean():.1f}')

print(f'\n  Strateji dağılımı (backtest seçimi):')
for s in ['XGB', 'HYBRID', 'ROLLING', 'ONOFF']:
    cnt = df_out[df_out['Strateji'] == s]
    if len(cnt):
        print(f'    {s:8s} → {len(cnt)} makine | ort. hata: %{cnt["Hata (%)"].mean():.1f}')

print(f'\n  En iyi 5:')
print(df_out.head(5)[['Makine', 'Strateji', 'Hata (%)', 'MAPE (%)']].to_string(index=False))
if len(kotu):
    print(f'\n  ❌ Kötü makineler:')
    print(df_out[h_vals > 25][['Makine', 'Strateji', 'Hata (%)', 'CV (%)']].to_string(index=False))

print(f'\n✅ MySQL veritabanı : {MYSQL_DATABASE}  (host={MYSQL_HOST})')
print(f'   run_id           : {run_id}')
print(f'   Tablolar         : forecast_runs, forecast_detail, forecast_daily, forecast_summary')
print(f'{"="*60}')
