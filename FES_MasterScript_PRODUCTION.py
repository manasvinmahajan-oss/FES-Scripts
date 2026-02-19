#!/usr/bin/env python
# coding: utf-8

# FES Master Script - PRODUCTION VERSION
# - Dynamic file paths based on trading date
# - Production SQL table names
# - Organized file structure

import pandas as pd
from datetime import datetime, timedelta
from zeep import Client, Settings
import os
import calendar
import datetime as dt
import numpy as np
from pathlib import Path
from sqlalchemy import create_engine
import urllib
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# Import PPT Generator (optional - only if python-pptx is installed)
try:
    from FES_PPT_Generator import generate_forecast_presentation
    PPT_AVAILABLE = True
except ImportError:
    PPT_AVAILABLE = False
    print("[WARNING] python-pptx not installed - PowerPoint generation disabled")
    print("[INFO] Install with: pip install python-pptx")

# ==========================================
# 1. DATABASE UPLOAD FUNCTION
# ==========================================
def upload_to_fabric(df, file_name):
    """
    Uploads data to Fabric SQL using a strict column mapping to match specific DB column names.
    PRODUCTION: Uses Generation_D_Minus_1 or Generation_D_Minus_X tables
    
    Returns:
        bool: True if upload succeeded, False if failed
    """
    # 1. CONFIGURATION
    server = 'g3hsqkj33hsejptu6vliyt5gny-6novrz7kmrcuriuozi2uqi5sy4.datawarehouse.fabric.microsoft.com'
    database = 'trading_data'

    # 2. TABLE SELECTION
    if "D-1" in file_name:
        table_name = "Generation_D_Minus_1"
    else:
        table_name = "Generation_D_Minus_X"

    print(f"Detected File: {file_name}")
    print(f"Target Table: {table_name}")

    # 3. PREPARE DATAFRAME
    sql_df = df.copy()

    # --- STRICT COLUMN MAPPING ---
    column_map = {
        'DateTime': 'DateTime',
        'Meteo ROI (MW)': 'Meteo ROI _MW_',
        'Meteo NI (MW)': 'Meteo NI _MW_',
        'Meteo TB (MW)': 'Meteo TB _MW_',
        'Meteo CK (MW)': 'Meteo CK _MW_',
        'Meteo LD (MW)': 'Meteo LD _MW_',
        'Meteo CD (MW)': 'Meteo CD _MW_',
        'Naïve Nonwind (MW)': 'Naïve Nonwind _MW_',
        'Self-forecast (MW)': 'Self-forecast _MW_',
        'Meteo DT (MW)': 'Meteo DT _MW_',
        'Meteo MUR (MW)': 'Meteo MUR _MW_',
        'Meteo S1 (MW)': 'Meteo S1 _MW_',
        'Meteo S2 (MW)': 'Meteo S2 _MW_'
    }

    # Add mappings for S3 to S25 (Future-proofing)
    for i in range(3, 26):
        source_col = f'Meteo S{i} (MW)'
        target_col = f'Meteo S{i} _MW_'
        column_map[source_col] = target_col

    # Apply the renaming
    sql_df.rename(columns=column_map, inplace=True)

    # Add Upload Timestamp
    sql_df['Upload_Timestamp'] = datetime.now()

    # Filter DataFrame to strictly match the columns we mapped (plus Timestamp)
    valid_db_columns = list(column_map.values()) + ['Upload_Timestamp']
    final_cols = [c for c in valid_db_columns if c in sql_df.columns]
    sql_df = sql_df[final_cols]

    # Ensure DateTime is correct format
    if 'DateTime' in sql_df.columns:
        sql_df['DateTime'] = pd.to_datetime(sql_df['DateTime'], format='%d/%m/%Y %H:%M')

    # 4. UPLOAD
    conn_str = (
        f"Driver={{ODBC Driver 18 for SQL Server}};"
        f"Server={server};"
        f"Database={database};"
        f"Authentication=ActiveDirectoryInteractive;"
        f"Encrypt=yes;TrustServerCertificate=no;Connection Timeout=60;"
    )
    params = urllib.parse.quote_plus(conn_str)
    engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")

    try:
        sql_df.to_sql(table_name, engine, if_exists='append', index=False)
        print(f"[OK] Uploaded {len(sql_df)} rows to {table_name} at {datetime.now().strftime('%H:%M:%S')}")
        return True
    except Exception as e:
        print(f"[FAIL] Upload Failed: {e}")
        return False

# ==========================================
# 2. HELPER FUNCTIONS
# ==========================================

lolu = ['0275', '402050', '402280']

def large_unit_availability(input_date):
    date_obj_1 = datetime.strptime(input_date, "%d/%m/%Y") - timedelta(days=1)
    date_obj_2 = datetime.strptime(input_date, "%d/%m/%Y")

    y1, m1, d1 = date_obj_1.year, date_obj_1.month, date_obj_1.day
    y2, m2, d2 = date_obj_2.year, date_obj_2.month, date_obj_2.day

    roi_unit_ids = ['Vayu_0275','Vayu_402050','Vayu_GU_402280']
    remaining = []

    settings = Settings(strict=False)
    wsdl = 'https://webservice.meteologica.com/api/wsdl/MeteologicaDataExchangeService.wsdl'
    client = Client(wsdl=wsdl, settings=settings)

    login_req = client.get_type('ns0:LoginReq')(
            username="Flogas",
            password="uc:DF824")

    login_response = client.service.login(request=login_req)
    session_token = login_response.header.sessionToken

    availability_req_type = client.get_type('ns0:getAvailabilityMultiReq')
    req = availability_req_type(
        header={'sessionToken': session_token},
        fromDate=datetime(y1, m1, d1, 22, 0, 0).isoformat() + 'Z',
        toDate=datetime(y2, m2, d2, 23, 0, 0).isoformat() + 'Z',
        facilitiesId=[],
        unit='MW')

    availability_response = client.service.getAvailabilityMulti(request=req)

    for item in availability_response.facilityAvailabilityData.item:
        if item.facilityId in roi_unit_ids :
            remaining.append(item) 

    logout_req_type = client.get_type('ns0:LogoutReq')(
        header={'sessionToken': session_token})
    client.service.logout(request=logout_req_type)

    input_date_obj = datetime.strptime(input_date, "%d/%m/%Y")
    previous_day = input_date_obj - timedelta(days=1)
    timestamps = []
    for minute in [0, 30]:
        timestamp = previous_day.replace(hour=23, minute=minute)
        timestamps.append(timestamp)
    for hour in range(0, 23):
        for minute in [0, 30]:
            timestamp = input_date_obj.replace(hour=hour, minute=minute)
            timestamps.append(timestamp)
    df = pd.DataFrame(timestamps, columns=['datetime'])

    id_availability_map = {'Vayu_0275':'CD Availability', 'Vayu_402050':'TB Availability' , 'Vayu_GU_402280':'CK Availability' }

    df['CD Availability'] = 1
    df['TB Availability'] = 1
    df['CK Availability'] = 1

    df['CD Max'] = 4.5
    df['TB Max'] = 13.8
    df['CK Max'] = 11.5

    for facility in remaining:
        if facility.availabilityData == None:
            pass
        else:
            for event in facility.availabilityData['item']:
                json_format = "%Y-%m-%dT%H:%M:%S%z"
                start_date = datetime.strptime(event['fromDate'], json_format)
                end_date = datetime.strptime(event['toDate'], json_format)

                start_date = start_date.replace(tzinfo=None)
                end_date = end_date.replace(tzinfo=None)
                date_series = pd.date_range(start=start_date, end=end_date, freq='30min')

                temp_df = pd.DataFrame({'datetime': date_series,'availability': float(event['powerPercentage'])/100})   

                col_name = id_availability_map[facility.facilityId]

                temp_df = temp_df.rename(columns={'availability': col_name})
                temp_df[col_name] = temp_df[col_name].replace(0, 1)
                temp_df['datetime'] = temp_df['datetime'].dt.tz_localize(None)

                df.set_index('datetime', inplace=True)
                temp_df.set_index('datetime', inplace=True)

                df.update(temp_df)
                df.reset_index(inplace=True)

    df['CD Current Output'] = df['CD Availability'] * df['CD Max']
    df['TB Current Output'] = df['TB Availability'] * df['TB Max']
    df['CK Current Output'] = df['CK Availability'] * df['CK Max']
    return df

def to_excel_serial_date(d):
    excel_epoch = dt.datetime(1899, 12, 30)
    if isinstance(d, str):
        d = pd.to_datetime(d, dayfirst=True)
    if hasattr(d, 'tzinfo') and d.tzinfo is not None:
        d = d.tz_localize(None) if hasattr(d, 'tz_localize') else d.replace(tzinfo=None)
    delta = d - excel_epoch
    return delta.total_seconds() / (24 * 60 * 60)

def find_latest_self_forecast(input_date_str, max_lookback_days=14):
    base_root = Path(r"V:\Renewables\Self-Forecasting\2) Forecasts Received by Trading Day")
    dt0 = datetime.strptime(input_date_str, "%d/%m/%Y")

    for delta in range(0, max_lookback_days + 1):
        dt = dt0 - timedelta(days=delta)
        year = dt.year
        month = dt.month
        month_name = calendar.month_name[month]
        day_folder = dt.strftime('%d.%m.%Y')

        month_folder = base_root / f"{year}" / f"{month}) {month_name}" / day_folder
        if not month_folder.exists():
            continue 

        expected_name = f"1) Aggregated Naturgy_Self_Forecast_Template_v1_{day_folder}.xlsx"
        candidate = month_folder / expected_name
        if candidate.exists():
            return candidate 

        matches = sorted(month_folder.glob("1) Aggregated Naturgy_Self_Forecast_Template_v1_*.xlsx"))
        if matches:
            return matches[-1]

    raise FileNotFoundError(f"No self-forecast file found within {max_lookback_days} days back from {input_date_str}.")

def process_forecast_data(input_date, base_df):
    file_path = find_latest_self_forecast(input_date)
    self_forecast = pd.read_excel(
        file_path,
        skiprows=16,
        nrows=24,
        usecols="C",
        header=None,
        names=['Self-forecast (MW)']
    )

    final_df = base_df.copy()
    final_df['time'] = pd.to_datetime(final_df['time'], format='%d/%m/%Y %H:%M')

    min_time = final_df['time'].min()
    max_time = final_df['time'].max()

    half_hourly_times = pd.date_range(start=min_time, end=max_time, freq='30min')
    half_hourly_df = pd.DataFrame({'time': half_hourly_times})
    final_df = pd.merge_asof(
        half_hourly_df.sort_values('time'),
        final_df.sort_values('time'),
        on='time',
        direction='backward'
    )

    final_df = final_df.ffill()
    hourly_forecast = self_forecast['Self-forecast (MW)'].values
    half_hourly_forecast = np.repeat(hourly_forecast, 2)

    if len(half_hourly_forecast) > len(final_df):
        half_hourly_forecast = half_hourly_forecast[:len(final_df)]
    elif len(half_hourly_forecast) < len(final_df):
        padding = np.full(len(final_df) - len(half_hourly_forecast), half_hourly_forecast[-1])
        half_hourly_forecast = np.append(half_hourly_forecast, padding)

    final_df['Self-forecast (MW)'] = half_hourly_forecast
    final_df['Naïve Nonwind (MW)'] = 0.7
    final_df['time'] = final_df['time'].dt.strftime('%d/%m/%Y %H:%M')

    column_order = [
        'time',
        'Meteo ROI (MW)', 
        'Meteo NI (MW)', 
        'Meteo TB (MW)', 
        'Meteo CK (MW)',
        'Meteo LD (MW)',
        'Meteo CD (MW)',
        'Naïve Nonwind (MW)',
        'Self-forecast (MW)',
        'Meteo DT (MW)', 
        'Meteo MUR (MW)',
        'Meteo S1 (MW)',
        'Meteo S2 (MW)'
    ]
    for col in column_order:
        if col not in final_df.columns:
            final_df[col] = 0

    return final_df[column_order]

# ==========================================
# 3. MAIN FUNCTION - WITH UPLOAD CONTROL
# ==========================================
def grab_forecast_data(input_date, upload_to_sql=False):
    """
    Get generation forecast data and save to production location.
    Automatically determines lag (D-1, D-2, D-3, etc.) based on trading date.

    Args:
        input_date: Date string in format 'dd/mm/YYYY' (trading day)
        upload_to_sql: Boolean - True to upload to Fabric SQL, False to skip upload

    Returns:
        tuple: (DataFrame with forecast data, lag_string like 'D-1' or 'D-2')
    """
    # Calculate lag dynamically
    trading_date = datetime.strptime(input_date, "%d/%m/%Y")
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    days_ahead = (trading_date - today).days
    lag = f"D-{days_ahead}"  # e.g., D-1, D-2, D-3
    
    print(f"[INFO] Trading Date: {input_date}")
    print(f"[INFO] Days Ahead: {days_ahead}")
    print(f"[INFO] Lag: {lag}")
    
    date_obj_1 = datetime.strptime(input_date, "%d/%m/%Y") - timedelta(days=1)
    date_obj_2 = datetime.strptime(input_date, "%d/%m/%Y")

    y1, m1, d1 = date_obj_1.year, date_obj_1.month, date_obj_1.day
    y2, m2, d2 = date_obj_2.year, date_obj_2.month, date_obj_2.day

    settings = Settings(strict=False)
    wsdl = "https://webservice.meteologica.com/api/wsdl/MeteologicaDataExchangeService.wsdl"
    client = Client(wsdl=wsdl)

    login_req_type = client.get_type('ns0:LoginReq')
    login_req = login_req_type(
        username="Flogas",
        password="uc:DF824"
    )
    login_response = client.service.login(request=login_req)
    session_token = login_response.header.sessionToken

    forecast_req_type = client.get_type('ns0:GetForecastMultiReq')
    req = forecast_req_type(
        header={'sessionToken': session_token},
        variableId='prod',
        predictorId='aggregated',
        fromDate=datetime(y1, m1, d1, 23, 0, 0).isoformat(),
        toDate=datetime(y2, m2, d2, 22, 30, 0).isoformat(),
        granularity='30',
        percentiles='50',
        facilitiesId=[])

    forecast_response = client.service.getForecastMulti(request=req)

    all_timestamps = set()
    facility_data = {}

    for facility in forecast_response.facilitiesForecastData.item:
        facility_id = facility.facilityId
        facility_data[facility_id] = {}

        forecast_pairs = facility.forecastData.split(':')
        start_idx = 1 if forecast_pairs[0] == '' else 0

        for i in range(start_idx, len(forecast_pairs)):
            parts = forecast_pairs[i].split('~')
            if len(parts) >= 2:
                timestamp = int(parts[0])
                value = float(parts[1])
                dt_obj = datetime.utcfromtimestamp(timestamp)
                all_timestamps.add(dt_obj)
                facility_data[facility_id][dt_obj] = value / 1000  # Convert to MW

    df = pd.DataFrame(
        index=pd.to_datetime(sorted(all_timestamps)),
        columns=[
            'Meteo ROI (MW)', 'Meteo NI (MW)', 'Meteo TB (MW)',
            'Meteo CK (MW)', 'Meteo LD (MW)' , 'Meteo CD (MW)', 'Meteo DT (MW)', 'Meteo MUR (MW)','Meteo S1 (MW)','Meteo S2 (MW)'
        ]
    )

    column_mapping = {
        'Vayu_Cluster1': 'Meteo ROI (MW)',
        'Vayu_Cluster2': 'Meteo NI (MW)',
        'Vayu_402050': 'Meteo TB (MW)',
        'Vayu_GU_402280': 'Meteo CK (MW)',
        'Flogas-solar_0587' : 'Meteo LD (MW)',
        'Vayu_0275': 'Meteo CD (MW)',
        'Flogas-solar_0378__': 'Meteo DT (MW)',
        'Vayu_GEN_504260': 'Meteo MUR (MW)',
        'Flogas-solar_0670' : 'Meteo S1 (MW)',
        'Flogas-solar_0684' : 'Meteo S2 (MW)'
    }

    for facility_id, col_name in column_mapping.items():
        if facility_id in facility_data:
            df[col_name] = [facility_data[facility_id].get(ts, 0) for ts in df.index]

    df = df.reset_index().rename(columns={'index': 'time'})
    df['time'] = df['time'] + pd.to_timedelta('1 hour')
    df['time'] = df['time'].dt.strftime('%d/%m/%Y %H:%M')

    df = process_forecast_data(input_date, df)
    df.rename(columns={"time": "DateTime"}, inplace=True)

    input_date_obj = datetime.strptime(input_date, "%d/%m/%Y")
    previous_day = input_date_obj - timedelta(days=1)
    timestamps = []
    for minute in [0, 30]:
        timestamp = previous_day.replace(hour=23, minute=minute)
        timestamps.append(timestamp)
    for hour in range(0, 23):
        for minute in [0, 30]:
            timestamp = input_date_obj.replace(hour=hour, minute=minute)
            timestamps.append(timestamp)

    df['DateTime'] = timestamps
    df['DateTime'] = df['DateTime'].dt.strftime('%d/%m/%Y %H:%M')

    temp = df.copy(deep=True)

    # Commented out large unit availability - uncomment if needed
    #  lsa = large_unit_availability(input_date)
    # df['Meteo ROI (MW)'] = df['Meteo ROI (MW)'] + 1.158*(df['Meteo CD (MW)']/lsa['CD Current Output']) + 1.158*(df['Meteo TB (MW)']/lsa['TB Current Output']) + 1.158*(df['Meteo CK (MW)']/lsa['CK Current Output'])

    temp['ROI Linear'] = df['Meteo ROI (MW)']
    temp['NI Linear'] = df['Meteo NI (MW)']

    df.loc[:, df.columns != 'DateTime'] = df.loc[:, df.columns != 'DateTime'].round(1)

    # === SAVE TO EXCEL - DYNAMIC PATH BASED ON TRADING DATE ===
    date_obj = datetime.strptime(input_date, "%d/%m/%Y")
    year = str(date_obj.year)
    month = date_obj.strftime("%B")  # Full month name
    
    output_path = Path(rf"I:\Daily Generation Forecasts\Daily Generation to Submit\{year}\{month}")
    output_path.mkdir(parents=True, exist_ok=True)
    
    file_name = f"Generation Forecast {date_obj.strftime('%d.%m.%Y')} {lag}.xlsx"
    full_path = output_path / file_name
    df.to_excel(full_path, index=False)
    print(f"[OK] File saved: {file_name}")

    # === UPLOAD TO FABRIC (ONLY IF ENABLED) ===
    upload_success = False
    if upload_to_sql:
        print("[UPLOAD] Uploading to Fabric SQL...")
        upload_success = upload_to_fabric(df, file_name)  # upload_to_fabric will auto-route to D_Minus_1 or D_Minus_X
    else:
        print("[SKIP] SQL upload disabled")

    logout_req_type = client.get_type('ns0:LogoutReq')(
        header={'sessionToken': session_token})
    client.service.logout(request=logout_req_type)

    return df, lag, upload_success  # Return both forecast, lag string, and upload status


# ==========================================
# 4. MURLEY GU COMPILER CLASS
# ==========================================
class MurleyGUCompiler:
    def __init__(self):
        self.cwd = Path.cwd()
        self.fabric_server = (
            "g3hsqkj33hsejptu6vliyt5gny-6novrz7kmrcuriuozi2uqi5sy4.datawarehouse.fabric.microsoft.com"
        )

    def find_gen_file(self, bid_date, lag="D-1"):
        """Finds Generation Forecast BID_DATE {lag}.xlsx"""
        date_obj = datetime.strptime(bid_date, "%d/%m/%Y")
        day_str = date_obj.strftime("%d.%m.%Y")
        year = str(date_obj.year)
        month = date_obj.strftime("%B")

        gen_base = Path(r"I:\Daily Generation Forecasts\Daily Generation to Submit")
        gen_file = gen_base / year / month / f"Generation Forecast {day_str} {lag}.xlsx"

        print(f"Looking for: {gen_file}")
        if not gen_file.exists():
            raise FileNotFoundError(f"Not found: {gen_file}")

        return gen_file

    def load_gen_data(self, bid_date, lag="D-1"):
        """Load generation forecast"""
        gen_file = self.find_gen_file(bid_date, lag)
        df = pd.read_excel(gen_file)

        print(f"Loaded {len(df)} rows from generation forecast")
        df["DateTime"] = pd.to_datetime(df["DateTime"], format="%d/%m/%Y %H:%M")
        return df.sort_values("DateTime")

    def create_aggregation(self, bid_date, lag="D-1"):
        """Create GU Traders Table"""
        gen_df = self.load_gen_data(bid_date, lag)

        # Trading day starts at 23:00 on D-1
        delivery_date = datetime.strptime(bid_date, "%d/%m/%Y")
        start_time = delivery_date - timedelta(days=1) + timedelta(hours=23)
        times = pd.date_range(start_time, periods=48, freq="30min")

        agg_df = pd.DataFrame({"DateTime": times})

        # Merge generation data
        gen_df["time_str"] = gen_df["DateTime"].dt.strftime("%d/%m/%Y %H:%M")
        agg_df["time_str"] = agg_df["DateTime"].dt.strftime("%d/%m/%Y %H:%M")

        gen_merge = gen_df[["time_str", "Meteo MUR (MW)"]].copy()

        agg_df = agg_df.merge(gen_merge, on="time_str", how="left").fillna(0.0)
        agg_df.drop("time_str", axis=1, inplace=True)

        # GU traders table format - round to 1 decimal
        agg_df["GU_504260"] = (-agg_df["Meteo MUR (MW)"]).round(1)
        agg_df["Price"] = ""
        agg_df["Time"] = agg_df["DateTime"].dt.strftime("%H:%M")

        return agg_df[["DateTime", "GU_504260", "Price", "Time"]]

    def generate_ets_bids(self, agg_df):
        """ETS bids format - GU always sells at columns 3&4 only
        
        GU Logic (based on actual Excel output):
        - Columns 1&2 must be EMPTY (not 0) to avoid ETS rejection
        - Columns 3&4 contain the negative quantity
        """
        ets_data = []

        for i, row in agg_df.iterrows():
            qty = row["GU_504260"]
            # GU always selling - values only in columns 3&4, columns 1&2 are EMPTY
            ets_data.append([
                row["Time"],
                i + 1,
                0,     # -1500 price: EMPTY (not 0!)
                0,     # -41.7 price: EMPTY (not 0!)
                qty,    # -41.7 price: negative quantity
                qty     # 9000 price: negative quantity (same as column 3)
            ])

        ets_df = pd.DataFrame(ets_data, columns=["", "Period", "-1500", "-41.7", "-41.7", "9000"])
        return ets_df

    def generate_dam_bids(self, agg_df):
        """DAM bids CSV format"""
        dam_rows = []
        for idx, row in agg_df.iterrows():
            qty = row["GU_504260"]
            buy_sell = "SELL" if qty < 0 else "BUY"

            dam_rows.append({
                "Period": idx + 1,
                "DateTime": row["DateTime"],
                "BuySell": buy_sell,
                "Curve-Price 1": -1500,
                "Curve-Qty 1": 0.0,
                "Curve-Price 2": -41.7,
                "Curve-Qty 2": 0.0,
                "Curve-Price 3": -41.7,
                "Curve-Qty 3": abs(qty),
                "Curve-Price 4": 9000.0,
                "Curve-Qty 4": abs(qty),
            })
        return pd.DataFrame(dam_rows)

    def create_bid_chart(self, agg_df, bid_date):
        """Create bid submission chart - saved to local output folder"""
        output_dir = self.cwd / "output"
        output_dir.mkdir(exist_ok=True)

        day_str = datetime.strptime(bid_date, "%d/%m/%Y").strftime("%d.%m.%Y")

        fig, ax = plt.subplots(figsize=(14, 8))

        times = agg_df["DateTime"]
        bids = agg_df["GU_504260"]

        ax.plot(times, bids, linewidth=2.5, color='#1f77b4', label='GU_504260 Bids', marker='o', markersize=4)
        ax.fill_between(times, bids, 0, alpha=0.3, color='#1f77b4')

        ax.axhline(y=0, color='red', linestyle='--', linewidth=1, alpha=0.5, label='Zero Line')

        ax.set_xlabel('Time', fontsize=14, fontweight='bold')
        ax.set_ylabel('Bid Quantity (MW)', fontsize=14, fontweight='bold')
        ax.set_title(f'Murley GU504260 Bid Submission\nDelivery Date: {day_str}', 
                     fontsize=16, fontweight='bold', pad=20)

        ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
        plt.xticks(rotation=45, ha='right')

        ax.grid(True, alpha=0.3, linestyle='--')
        ax.set_axisbelow(True)

        ax.legend(loc='upper right', fontsize=11, framealpha=0.9)

        total = bids.sum()
        avg = bids.mean()
        min_bid = bids.min()
        max_bid = bids.max()

        stats_text = f'Total: {total:.1f} MW\nAverage: {avg:.1f} MW\nMin: {min_bid:.1f} MW\nMax: {max_bid:.1f} MW'
        ax.text(0.02, 0.98, stats_text, transform=ax.transAxes, 
                fontsize=10, verticalalignment='top',
                bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.8))

        plt.tight_layout()

        chart_file = output_dir / f"GU_504260_Bid_Chart_{day_str}.png"
        plt.savefig(chart_file, dpi=300, bbox_inches='tight')
        print(f"Chart saved: {chart_file}")

        plt.close()
        return chart_file

    def save_files(self, bid_date, agg_df, dam_df, ets_df, lag="D-1"):
        """Save output files to production locations"""
        date_obj = datetime.strptime(bid_date, "%d/%m/%Y")
        day_str = date_obj.strftime("%d.%m.%Y")
        year = str(date_obj.year)
        month = date_obj.strftime("%B")

        # Add totals row for traders table
        totals = pd.Series({
            "DateTime": "",
            "GU_504260": agg_df["GU_504260"].sum(),
            "Price": "",
            "Time": ""
        })
        agg_with_totals = pd.concat([agg_df, pd.DataFrame([totals])], ignore_index=True)

        # === 1. ETS BID FILE ===
        ets_path = Path(rf"I:\ETS Bids\DAM Bids\{year}\{month}\DAM Bids {day_str}")
        ets_path.mkdir(parents=True, exist_ok=True)
        ets_file = ets_path / f"DAM GU_504260--ALL {lag}.csv"
        ets_df.to_csv(ets_file, index=False)

        # === 2. TRADERS TABLE ===
        traders_path = Path(rf"I:\Day-Ahead Process\Traders' Tables\{year}\{month}")
        traders_path.mkdir(parents=True, exist_ok=True)
        traders_file = traders_path / f"DAM Traders' Table {day_str} {lag} GU_504260.xlsx"
        
        with pd.ExcelWriter(traders_file, engine='openpyxl') as writer:
            agg_with_totals.to_excel(writer, index=False, sheet_name='Sheet1')
            worksheet = writer.sheets['Sheet1']
            for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
                if isinstance(row[1].value, (int, float)) and row[1].value != "":
                    row[1].number_format = '0.0'

        # === 3. DAM AUCTION RECONCILIATION ===
        auction_path = Path(rf"I:\Auction reconciliation tables\DAM\{year}\{month}\{day_str}")
        auction_path.mkdir(parents=True, exist_ok=True)
        dam_file = auction_path / f"DAM GU_504260--ALL {lag} GU_504260.csv"
        dam_df.to_csv(dam_file, index=False)

        print(f"\nFiles saved:")
        print(f"  ETS Bid: {ets_file}")
        print(f"  Traders Table: {traders_file}")
        print(f"  DAM Auction: {dam_file}")

        self.create_bid_chart(agg_df, bid_date)

        return dam_file, ets_file, traders_file

    def upload_to_fabric(self, agg_df, bid_date, lag="D-1", use_production=True):
        """Upload GU bids to Fabric warehouse
        
        Args:
            lag: Lag string (D-1, D-2, etc.) to determine which table to use
            use_production: If True, uses production tables (Bids_Murley_D_Minus_1 or D_Minus_X)
                           If False, uses test_Bids_Murley (test table)
        
        Returns:
            bool: True if upload succeeded, False if failed
        """

        # Prepare upload dataframe - simple structure for GU
        upload_df = pd.DataFrame()
        upload_df["DateTime"] = agg_df["DateTime"]
        upload_df["GU_504260"] = agg_df["GU_504260"]
        upload_df["Upload_Timestamp"] = datetime.now()

        # Choose table: production (D-1 or D-X) or test
        if use_production:
            if lag == "D-1":
                table_name = "Bids_Murley_D_Minus_1"
            else:
                table_name = "Bids_Murley_D_Minus_X"
        else:
            table_name = "test_Bids_Murley"

        conn_str = (
            f"Driver={{ODBC Driver 18 for SQL Server}};"
            f"Server={self.fabric_server};"
            f"Database=trading_data;"
            f"Authentication=ActiveDirectoryInteractive;"
            f"Encrypt=yes;TrustServerCertificate=no;Connection Timeout=60;"
        )
        params = urllib.parse.quote_plus(conn_str)
        engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")

        try:
            upload_df.to_sql(table_name, engine, if_exists="append", index=False, method='multi')
            print(f"\n[OK] Fabric upload: {len(upload_df)} rows to {table_name}")
            print(f"[OK] Uploaded columns: DateTime, GU_504260, Upload_Timestamp")
            return True
        except Exception as e:
            print(f"\n[FAIL] Fabric upload failed: {str(e)}")
            return False

    def print_analysis(self, agg_df):
        """Print traders table preview and metrics"""
        print("\n" + "="*70)
        print("TRADERS TABLE PREVIEW")
        print("="*70)
        print(agg_df.head(10).to_string(index=False))
        print("...")
        print(agg_df.tail(5).to_string(index=False))

        print("\n" + "="*70)
        print("SUMMARY METRICS")
        print("="*70)
        print(f"Total periods: {len(agg_df)}")
        print(f"Trading period: {agg_df['DateTime'].min()} to {agg_df['DateTime'].max()}")
        print(f"\nGU_504260 Statistics:")
        print(f"  Total: {agg_df['GU_504260'].sum():.2f} MW")
        print(f"  Average: {agg_df['GU_504260'].mean():.2f} MW")
        print(f"  Minimum: {agg_df['GU_504260'].min():.2f} MW")
        print(f"  Maximum: {agg_df['GU_504260'].max():.2f} MW")
        print(f"  Std Dev: {agg_df['GU_504260'].std():.2f} MW")
        print("="*70 + "\n")

    def run(self, bid_date="21/01/2026", lag="D-1", upload_sql=False, use_production=True):
        """Full workflow execution
        
        Args:
            bid_date: Trading date in DD/MM/YYYY format
            lag: Lag string (D-1, D-2, etc.)
            upload_sql: Whether to upload to SQL database
            use_production: If True, uses production table (Bids_Murley_D_Minus_1 or D_Minus_X)
                           If False, uses test table (test_Bids_Murley)
        """
        print("\n" + "="*70)
        print("MURLEY GU504260 BID COMPILATION WORKFLOW")
        print("="*70)
        print(f"Trading Day: {bid_date}")
        print(f"Lag: {lag}")
        print(f"Execution Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        if upload_sql:
            if use_production:
                table_name = "Bids_Murley_D_Minus_1" if lag == "D-1" else "Bids_Murley_D_Minus_X"
            else:
                table_name = "test_Bids_Murley"
            print(f"SQL Upload: ENABLED -> {table_name}")
        else:
            print(f"SQL Upload: DISABLED")
        print("="*70 + "\n")

        agg_df = self.create_aggregation(bid_date, lag)
        print(f"Aggregation table created: {len(agg_df)} periods")

        dam_df = self.generate_dam_bids(agg_df)
        ets_df = self.generate_ets_bids(agg_df)
        print(f"Bid files generated: DAM and ETS formats")

        self.save_files(bid_date, agg_df, dam_df, ets_df, lag)

        upload_success = False
        if upload_sql:
            upload_success = self.upload_to_fabric(agg_df, bid_date, lag, use_production=use_production)
        else:
            print("[SKIP] SQL upload disabled")

        self.print_analysis(agg_df)

        return agg_df, dam_df, ets_df, upload_success


# ==========================================
# 5. SUPPLY UNIT COMPILER CLASS
# ==========================================
class SupplyUnitCompiler:
    def __init__(self):
        self.cwd = Path.cwd()
        self.fabric_server = (
            "g3hsqkj33hsejptu6vliyt5gny-6novrz7kmrcuriuozi2uqi5sy4.datawarehouse.fabric.microsoft.com"
        )

    def find_demand_file(self, bid_date):
        """Find QH demand file (CSV: 2026-01-18.csv for trading day 18th)"""
        date_obj = datetime.strptime(bid_date, "%d/%m/%Y")
        file_name = date_obj.strftime("%Y-%m-%d") + ".csv"

        demand_path = Path(r"I:\Daily Forecasts\Daily Demand Forecast - QH\D-1")
        demand_file = demand_path / file_name

        print(f"Looking for QH Demand: {demand_file}")
        if not demand_file.exists():
            raise FileNotFoundError(f"QH Demand file not found: {demand_file}")

        return demand_file

    def find_gen_file(self, bid_date, lag="D-1"):
        """Find generation forecast file"""
        date_obj = datetime.strptime(bid_date, "%d/%m/%Y")
        day_str = date_obj.strftime("%d.%m.%Y")
        year = str(date_obj.year)
        month = date_obj.strftime("%B")

        gen_base = Path(r"I:\Daily Generation Forecasts\Daily Generation to Submit")
        gen_file = gen_base / year / month / f"Generation Forecast {day_str} {lag}.xlsx"

        print(f"Looking for Generation: {gen_file}")
        if not gen_file.exists():
            raise FileNotFoundError(f"Generation file not found: {gen_file}")

        return gen_file

    def load_forecasts(self, bid_date, lag="D-1"):
        """Load QH demand (CSV) and generation (Excel) forecasts"""
        demand_file = self.find_demand_file(bid_date)
        gen_file = self.find_gen_file(bid_date, lag)

        # Load QH demand from CSV
        demand_df = pd.read_csv(demand_file)
        demand_df["DateTime"] = pd.to_datetime(demand_df["DateTime"], format="%d/%m/%Y %H:%M")

        # Convert kWh to MW: (kWh * 2) / 1000 - keep 1 decimal
        demand_df["QH_MW"] = ((demand_df["Demand"] * 2) / 1000).round(1)

        # Load generation from Excel
        gen_df = pd.read_excel(gen_file)
        gen_df["DateTime"] = pd.to_datetime(gen_df["DateTime"], format="%d/%m/%Y %H:%M")

        print(f"Loaded: QH Demand={len(demand_df)} rows, Generation={len(gen_df)} rows")

        return demand_df.sort_values("DateTime"), gen_df.sort_values("DateTime")

    def create_aggregation(self, bid_date, lag="D-1"):
        """Create SU Traders Table"""
        demand_df, gen_df = self.load_forecasts(bid_date, lag)

        # Trading day starts at 23:00 on D-1
        delivery_date = datetime.strptime(bid_date, "%d/%m/%Y")
        start_time = delivery_date - timedelta(days=1) + timedelta(hours=23)
        times = pd.date_range(start_time, periods=48, freq="30min")

        agg_df = pd.DataFrame({"DateTime": times})

        # Merge data
        demand_df["time_str"] = demand_df["DateTime"].dt.strftime("%d/%m/%Y %H:%M")
        gen_df["time_str"] = gen_df["DateTime"].dt.strftime("%d/%m/%Y %H:%M")
        agg_df["time_str"] = agg_df["DateTime"].dt.strftime("%d/%m/%Y %H:%M")

        # Merge QH demand
        agg_df = agg_df.merge(demand_df[["time_str", "QH_MW"]], on="time_str", how="left")

        # Merge generation columns
        gen_merge = gen_df[[
            "time_str", "Meteo ROI (MW)", "Meteo NI (MW)", "Meteo TB (MW)",
            "Meteo CK (MW)", "Meteo LD (MW)", "Meteo CD (MW)", 
            "Naïve Nonwind (MW)", "Self-forecast (MW)", "Meteo DT (MW)",
            "Meteo S1 (MW)", "Meteo S2 (MW)"
        ]].copy()

        agg_df = agg_df.merge(gen_merge, on="time_str", how="left").fillna(0.0)
        agg_df.drop("time_str", axis=1, inplace=True)

        # Build Traders Table columns with proper rounding
        agg_df["Adj. QH (MW)"] = agg_df["QH_MW"].round(1)
        agg_df["Adj. NQH (MW)"] = 0
        agg_df["Unmetered (MW)"] = 0.5

        # Generation (negative = supply) - round to 1 decimal
        agg_df["Adj. ROI Wind (MW)"] = (-agg_df["Meteo ROI (MW)"]).round(1)
        agg_df["Adj. NI Wind (MW)"] = (-agg_df["Meteo NI (MW)"]).round(1)
        agg_df["Adj. Tullabrack (MW)"] = (-agg_df["Meteo TB (MW)"]).round(1)
        agg_df["Adj. Cloghaneleskirt (MW)"] = (-agg_df["Meteo CK (MW)"]).round(1)
        agg_df["Adj. Lisdowney (MW)"] = (-agg_df["Meteo LD (MW)"]).round(1)
        agg_df["Adj. Curraghderrig (MW)"] = (-agg_df["Meteo CD (MW)"]).round(1)
        agg_df["Adj. Nonwind (MW)"] = (-agg_df["Naïve Nonwind (MW)"]).round(1)
        agg_df["Self-forecast (MW)"] = (-agg_df["Self-forecast (MW)"]).round(1)
        agg_df["Adj. Davidstown (MW)"] = (-agg_df["Meteo DT (MW)"]).round(1)

        # S1-S25 columns
        for i in range(1, 26):
            if i <= 2:
                agg_df[f"S{i}"] = (-agg_df[f"Meteo S{i} (MW)"]).round(1)
            else:
                agg_df[f"S{i}"] = 0

        # Trading Qty = Net Demand
        demand_cols = ["Adj. QH (MW)", "Adj. NQH (MW)", "Unmetered (MW)"]
        gen_cols = ["Adj. ROI Wind (MW)", "Adj. NI Wind (MW)", "Adj. Tullabrack (MW)",
                    "Adj. Cloghaneleskirt (MW)", "Adj. Lisdowney (MW)", "Adj. Curraghderrig (MW)",
                    "Adj. Nonwind (MW)", "Self-forecast (MW)", "Adj. Davidstown (MW)"] + [f"S{i}" for i in range(1, 26)]

        agg_df["Trading Qty (MW)"] = agg_df[demand_cols + gen_cols].sum(axis=1).round(1)
        agg_df["SU_400130"] = agg_df["Trading Qty (MW)"]
        agg_df["Price"] = ""

        # Reorder columns
        final_cols = (["DateTime"] + demand_cols + 
                      ["Adj. ROI Wind (MW)", "Adj. NI Wind (MW)", "Adj. Tullabrack (MW)",
                       "Adj. Cloghaneleskirt (MW)", "Adj. Lisdowney (MW)", "Adj. Curraghderrig (MW)",
                       "Adj. Nonwind (MW)", "Self-forecast (MW)", "Adj. Davidstown (MW)"] + 
                      [f"S{i}" for i in range(1, 26)] + 
                      ["Trading Qty (MW)", "SU_400130", "Price"])

        return agg_df[final_cols]

    def generate_ets_bids(self, agg_df):
        """ETS bids - SU format matching Excel VBA output exactly
        
        VBA Logic:
        - Column 1 (-500 price): ALWAYS contains the bid quantity (negative or positive)
        - When SELLING (negative): Column 4 (4000 price) also gets the same negative value
        - When BUYING (positive): Column 2 (500 price) also gets the same positive value
        """
        ets_data = []

        for i, row in agg_df.iterrows():
            qty = row["SU_400130"]
            
            if qty < 0:  # SELLING (excess generation)
                # Column 1 has the negative value, Column 4 has same negative value
                # Columns 2&3 must be EMPTY (not 0) to avoid ETS rejection
                ets_data.append([
                    row["DateTime"].strftime("%H:%M:%S"),
                    i + 1,
                    qty,           # -500 price: negative quantity
                    "",            # 500 price: EMPTY (not 0!)
                    "",            # 500 price: EMPTY (not 0!)
                    qty            # 4000 price: negative quantity (same as column 1)
                ])
            else:  # BUYING (demand exceeds generation)
                # Column 1 and Column 2 both have positive value
                # Columns 3&4 are 0 (matching VBA output)
                ets_data.append([
                    row["DateTime"].strftime("%H:%M:%S"),
                    i + 1,
                    qty,           # -500 price: positive quantity
                    qty,           # 500 price: positive quantity (same as column 1)
                    0,             # 500 price: 0
                    0              # 4000 price: 0
                ])

        ets_df = pd.DataFrame(ets_data, columns=["", "Period", "-500", "500", "500", "4000"])
        return ets_df

    def generate_dam_bids(self, agg_df):
        """DAM bids CSV with correct price placement logic (matching VBA)"""
        dam_rows = []
        for idx, row in agg_df.iterrows():
            qty = row["SU_400130"]
            abs_qty = abs(qty)  # Always work with absolute values
            
            # VBA Logic - quantities are ALWAYS positive
            # Sign only determines placement
            
            if qty < 0:  # SELLING
                dam_rows.append({
                    "Period": idx + 1,
                    "DateTime": row["DateTime"],
                    "BuySell": "SELL",
                    "Curve-Price 1": -500,
                    "Curve-Qty 1": 0.0,
                    "Curve-Price 2": 500,
                    "Curve-Qty 2": 0.0,
                    "Curve-Price 3": 500,
                    "Curve-Qty 3": abs_qty,  # POSITIVE quantity
                    "Curve-Price 4": 4000,
                    "Curve-Qty 4": abs_qty,  # POSITIVE quantity
                })
            else:  # BUYING
                dam_rows.append({
                    "Period": idx + 1,
                    "DateTime": row["DateTime"],
                    "BuySell": "BUY",
                    "Curve-Price 1": -500,
                    "Curve-Qty 1": abs_qty,  # POSITIVE quantity
                    "Curve-Price 2": 500,
                    "Curve-Qty 2": abs_qty,  # POSITIVE quantity
                    "Curve-Price 3": 500,
                    "Curve-Qty 3": 0.0,
                    "Curve-Price 4": 4000,
                    "Curve-Qty 4": 0.0,
                })
        return pd.DataFrame(dam_rows)

    def create_bid_chart(self, agg_df, bid_date):
        """Create SU bid chart - saved to local output folder"""
        output_dir = self.cwd / "output"
        output_dir.mkdir(exist_ok=True)

        day_str = datetime.strptime(bid_date, "%d/%m/%Y").strftime("%d.%m.%Y")

        fig, ax = plt.subplots(figsize=(14, 8))

        times = agg_df["DateTime"]
        bids = agg_df["SU_400130"]

        ax.plot(times, bids, linewidth=2.5, color='#2ecc71', label='SU_400130 Bids', marker='o', markersize=4)
        ax.fill_between(times, bids, 0, alpha=0.3, color='#2ecc71')

        ax.axhline(y=0, color='red', linestyle='--', linewidth=1, alpha=0.5, label='Zero Line')

        ax.set_xlabel('Time', fontsize=14, fontweight='bold')
        ax.set_ylabel('Bid Quantity (MW)', fontsize=14, fontweight='bold')
        ax.set_title(f'Supply Unit SU400130 Bid Submission\nDelivery Date: {day_str}', 
                     fontsize=16, fontweight='bold', pad=20)

        ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
        plt.xticks(rotation=45, ha='right')

        ax.grid(True, alpha=0.3, linestyle='--')
        ax.set_axisbelow(True)

        ax.legend(loc='upper right', fontsize=11, framealpha=0.9)

        total = bids.sum()
        avg = bids.mean()
        min_bid = bids.min()
        max_bid = bids.max()

        stats_text = f'Total: {total:.1f} MW\nAverage: {avg:.1f} MW\nMin: {min_bid:.1f} MW\nMax: {max_bid:.1f} MW'
        ax.text(0.02, 0.98, stats_text, transform=ax.transAxes, 
                fontsize=10, verticalalignment='top',
                bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.8))

        plt.tight_layout()

        chart_file = output_dir / f"SU_400130_Bid_Chart_{day_str}.png"
        plt.savefig(chart_file, dpi=300, bbox_inches='tight')
        print(f"Chart saved: {chart_file}")

        plt.close()
        return chart_file

    def save_files(self, bid_date, agg_df, dam_df, ets_df, lag="D-1"):
        """Save SU files to production locations"""
        date_obj = datetime.strptime(bid_date, "%d/%m/%Y")
        day_str = date_obj.strftime("%d.%m.%Y")
        year = str(date_obj.year)
        month = date_obj.strftime("%B")

        # Add totals + Net Demand row
        totals = agg_df.select_dtypes(include=['number']).sum()
        totals["DateTime"] = ""

        net_demand_row = pd.Series({col: "" for col in agg_df.columns})
        net_demand_row["DateTime"] = "Net Demand"
        net_demand_row["Trading Qty (MW)"] = totals["Trading Qty (MW)"]

        agg_with_totals = pd.concat([agg_df, pd.DataFrame([totals]), pd.DataFrame([net_demand_row])], ignore_index=True)

        # === 1. ETS BID FILE ===
        ets_path = Path(rf"I:\ETS Bids\DAM Bids\{year}\{month}\DAM Bids {day_str}")
        ets_path.mkdir(parents=True, exist_ok=True)
        ets_file = ets_path / f"DAM SU_400130--ALL {lag}.csv"
        ets_df.to_csv(ets_file, index=False)

        # === 2. TRADERS TABLE ===
        traders_path = Path(rf"I:\Day-Ahead Process\Traders' Tables\{year}\{month}")
        traders_path.mkdir(parents=True, exist_ok=True)
        traders_file = traders_path / f"DAM Traders' Table {day_str} {lag} SU_400130.xlsx"
        
        with pd.ExcelWriter(traders_file, engine='openpyxl') as writer:
            agg_with_totals.to_excel(writer, index=False, sheet_name='Sheet1')

            # Format numeric columns to show 1 decimal
            worksheet = writer.sheets['Sheet1']
            for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
                for cell in row:
                    if isinstance(cell.value, (int, float)) and cell.value != "":
                        cell.number_format = '0.0'

        # === 3. DAM AUCTION RECONCILIATION ===
        auction_path = Path(rf"I:\Auction reconciliation tables\DAM\{year}\{month}\{day_str}")
        auction_path.mkdir(parents=True, exist_ok=True)
        dam_file = auction_path / f"DAM SU_400130--ALL {lag} SU_400130.csv"
        dam_df.to_csv(dam_file, index=False)

        print(f"\nFiles saved:")
        print(f"  ETS Bid: {ets_file}")
        print(f"  Traders Table: {traders_file}")
        print(f"  DAM Auction: {dam_file}")

        self.create_bid_chart(agg_df, bid_date)

        return dam_file, ets_file, traders_file

    def upload_to_fabric(self, agg_df, lag="D-1", use_production=True):
        """Upload SU bids to Fabric warehouse
        
        Args:
            use_production: If True, uses Bids_SU_D_Minus_1 (production table)
                           If False, uses test_Bids_SU (test table)
        
        Returns:
            bool: True if upload succeeded, False if failed
        """

        # Prepare upload dataframe with database column names
        upload_df = pd.DataFrame()
        upload_df["DateTime"] = agg_df["DateTime"]

        # Map Python column names to database column names (with SPACE before underscore)
        column_mapping = {
            "Adj. QH (MW)": "Adj. QH _MW_",
            "Adj. NQH (MW)": "Adj. NQH _MW_",
            "Unmetered (MW)": "Unmetered _MW_",
            "Adj. ROI Wind (MW)": "Adj. ROI Wind _MW_",
            "Adj. NI Wind (MW)": "Adj. NI Wind _MW_",
            "Adj. Tullabrack (MW)": "Adj. Tullabrack _MW_",
            "Adj. Cloghaneleskirt (MW)": "Adj. Cloghaneleskirt _MW_",
            "Adj. Lisdowney (MW)": "Adj. Lisdowney _MW_",
            "Adj. Curraghderrig (MW)": "Adj. Curraghderrig _MW_",
            "Adj. Nonwind (MW)": "Adj. Nonwind _MW_",
            "Self-forecast (MW)": "Self-forecast _MW_",
            "Adj. Davidstown (MW)": "Adj. Davidstown _MW_",
            "Trading Qty (MW)": "Trading Qty _MW_",
            "SU_400130": "SU_400130"
        }

        # Add mapped columns
        for src_col, db_col in column_mapping.items():
            if src_col in agg_df.columns:
                upload_df[db_col] = agg_df[src_col]
            else:
                upload_df[db_col] = None

        # Add S1-S25 columns
        for i in range(1, 26):
            col_name = f"S{i}"
            if col_name in agg_df.columns:
                upload_df[col_name] = agg_df[col_name]
            else:
                upload_df[col_name] = None

        upload_df["Upload_Timestamp"] = datetime.now()

        # Choose table: production (D-1 or D-X) or test
        if use_production:
            if lag == "D-1":
                table_name = "Bids_SU_D_Minus_1"
            else:
                table_name = "Bids_SU_D_Minus_X"
        else:
            table_name = "test_Bids_SU"

        # Connect to Fabric
        conn_str = (
            f"Driver={{ODBC Driver 18 for SQL Server}};"
            f"Server={self.fabric_server};"
            f"Database=trading_data;"
            f"Authentication=ActiveDirectoryInteractive;"
            f"Encrypt=yes;TrustServerCertificate=no;Connection Timeout=60;"
        )
        params = urllib.parse.quote_plus(conn_str)
        engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")

        try:
            upload_df.to_sql(table_name, engine, if_exists="append", index=False, method='multi')
            print(f"\n{'='*70}")
            print("[OK] FABRIC UPLOAD SUCCESSFUL")
            print(f"{'='*70}")
            print(f"Table: {table_name}")
            print(f"Rows uploaded: {len(upload_df)}")
            print(f"Columns: {len(upload_df.columns)} (DateTime, Demand, Generation, S1-S25, SU_400130, Upload_Timestamp)")
            print(f"Time range: {upload_df['DateTime'].min()} to {upload_df['DateTime'].max()}")
            print(f"{'='*70}")
            return True
        except Exception as e:
            print(f"\n{'='*70}")
            print("[FAIL] FABRIC UPLOAD FAILED")
            print(f"{'='*70}")
            print(f"Error: {str(e)}")
            print(f"{'='*70}")
            return False

    def print_analysis(self, agg_df):
        """Print traders table preview and metrics"""
        print("\n" + "="*70)
        print("TRADERS TABLE PREVIEW")
        print("="*70)

        # Select key columns for preview
        preview_cols = ["DateTime", "Adj. QH (MW)", "Adj. ROI Wind (MW)", 
                        "Adj. NI Wind (MW)", "Adj. Nonwind (MW)", "Trading Qty (MW)", "SU_400130"]

        print(agg_df[preview_cols].head(10).to_string(index=False))
        print("...")
        print(agg_df[preview_cols].tail(5).to_string(index=False))

        print("\n" + "="*70)
        print("SUMMARY METRICS")
        print("="*70)
        print(f"Total periods: {len(agg_df)}")
        print(f"Trading period: {agg_df['DateTime'].min()} to {agg_df['DateTime'].max()}")

        print(f"\nDemand Statistics:")
        print(f"  QH Total: {agg_df['Adj. QH (MW)'].sum():.2f} MW")
        print(f"  QH Average: {agg_df['Adj. QH (MW)'].mean():.2f} MW")

        print(f"\nGeneration Statistics:")
        print(f"  ROI Wind Total: {agg_df['Adj. ROI Wind (MW)'].sum():.2f} MW")
        print(f"  NI Wind Total: {agg_df['Adj. NI Wind (MW)'].sum():.2f} MW")
        print(f"  Nonwind Total: {agg_df['Adj. Nonwind (MW)'].sum():.2f} MW")

        print(f"\nNet Position (SU_400130):")
        print(f"  Total: {agg_df['SU_400130'].sum():.2f} MW")
        print(f"  Average: {agg_df['SU_400130'].mean():.2f} MW")
        print(f"  Minimum: {agg_df['SU_400130'].min():.2f} MW")
        print(f"  Maximum: {agg_df['SU_400130'].max():.2f} MW")
        print(f"  Std Dev: {agg_df['SU_400130'].std():.2f} MW")
        print("="*70 + "\n")

    def run(self, bid_date="21/01/2026", lag="D-1", upload_sql=False, use_production=True):
        """Full SU workflow
        
        Args:
            bid_date: Trading date in DD/MM/YYYY format
            lag: Lag string (D-1, D-2, etc.)
            upload_sql: Whether to upload to SQL database
            use_production: If True, uses production table (Bids_SU_D_Minus_1 or D_Minus_X)
                           If False, uses test table (test_Bids_SU)
        """
        print("\n" + "="*70)
        print("SUPPLY UNIT SU400130 BID COMPILATION WORKFLOW")
        print("="*70)
        print(f"Trading Day: {bid_date}")
        print(f"Lag: {lag}")
        print(f"Execution Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        if upload_sql:
            if use_production:
                table_name = "Bids_SU_D_Minus_1" if lag == "D-1" else "Bids_SU_D_Minus_X"
            else:
                table_name = "test_Bids_SU"
            print(f"SQL Upload: ENABLED -> {table_name}")
        else:
            print(f"SQL Upload: DISABLED")
        print("="*70 + "\n")

        agg_df = self.create_aggregation(bid_date, lag)
        print(f"Aggregation table created: {len(agg_df)} periods")

        dam_df = self.generate_dam_bids(agg_df)
        ets_df = self.generate_ets_bids(agg_df)
        print(f"Bid files generated: DAM and ETS formats")

        self.save_files(bid_date, agg_df, dam_df, ets_df, lag)

        upload_success = False
        if upload_sql:
            upload_success = self.upload_to_fabric(agg_df, lag, use_production=use_production)
        else:
            print("[SKIP] SQL upload disabled")

        self.print_analysis(agg_df)

        return agg_df, dam_df, ets_df, upload_success


# ==========================================
# 6. POWERPOINT GENERATION FUNCTION
# ==========================================
def create_forecast_presentation(trading_date_str, gu_chart_path=None, su_chart_path=None, send_email=False, force_friday_mode=False):
    """
    Create PowerPoint presentation with forecast and bid charts
    Optionally send email to trading team
    For Friday presentations, automatically includes weekend forecasts
    
    Args:
        trading_date_str: Trading date in DD/MM/YYYY format
        gu_chart_path: Path to GU bid chart PNG (optional)
        su_chart_path: Path to SU bid chart PNG (optional)
        send_email: If True, send email to trading team with D-1 forecast attached
        force_friday_mode: If True, force Friday mode (load weekend forecasts)
    
    Returns:
        Path to presentation file or None if PPT generation unavailable
    """
    if not PPT_AVAILABLE:
        print("[SKIP] PowerPoint generation not available (python-pptx not installed)")
        return None
    
    try:
        ppt_path = generate_forecast_presentation(trading_date_str, gu_chart_path, su_chart_path, send_email, force_friday_mode)
        return ppt_path
    except Exception as e:
        print(f"[ERROR] PowerPoint generation failed: {e}")
        return None


# ==========================================
# 7. NO AUTO-EXECUTION - GUI CONTROLLED
# ==========================================
# The following lines are COMMENTED OUT so the script doesn't auto-run
# The GUI will call the functions as needed

# if __name__ == "__main__":
#     # Example usage - uncomment to test manually
#     input_date = "22/01/2026"
#     
#     # Get forecast (with upload) - PRODUCTION MODE
#     # forecast_df = grab_forecast_data(input_date, upload_to_sql=True)
#     
#     # Compile GU bids (with upload) - PRODUCTION MODE
#     # gu_compiler = MurleyGUCompiler()
#     # agg, dam, ets = gu_compiler.run(input_date, upload_sql=True, use_production=True)
#     
#     # Compile SU bids (with upload) - PRODUCTION MODE
#     # su_compiler = SupplyUnitCompiler()
#     # agg, dam, ets = su_compiler.run(input_date, upload_sql=True, use_production=True)
