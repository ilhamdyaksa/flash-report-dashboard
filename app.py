import io
import os
import json
import re
import math
import zipfile
import requests
import time
from datetime import datetime, timedelta
from io import BytesIO
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.io as pio
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import streamlit as st
from playwright.sync_api import sync_playwright

# ==========================================
# 1. FUNGSI API & AUTOMASI PLAYWRIGHT
# ==========================================
@st.cache_data(ttl=300)
def fetch_live_api_status(user_4g, pass_4g, token_uso, target_sites):
    site_status_dict = {}
    log_msgs = []
    
    try:
        res_login = requests.post("https://api.baktioss.id/v2/auth/login", json={"username": user_4g, "password": pass_4g}, timeout=30)
        if res_login.status_code in [200, 201]:
            token_4g = res_login.json().get("data", {}).get("accessToken")
            if token_4g:
                headers = {"Authorization": f"Bearer {token_4g}"}
                for sid in target_sites:
                    try:
                        res_site = requests.get(f"https://api.baktioss.id/v1/sites/detail/information?sites={sid}", headers=headers, timeout=10)
                        if res_site.status_code == 200:
                            data = res_site.json()
                            if "status" in data and data["status"]:
                                st_val = str(data["status"]).strip().capitalize()
                                site_status_dict[sid] = st_val
                    except Exception:
                        pass
        else: log_msgs.append(f"❌ Login API 4G Gagal: {res_login.status_code}")
    except Exception as e: log_msgs.append(f"❌ Gagal tarik API 4G: {e}")

    try:
        res_uso = requests.get("http://nmsapi.btsbakti.com/api/v1/itjen_kominfo/site?limit=10000", headers={"Authorization": f"Bearer {token_uso}"}, timeout=30)
        if res_uso.status_code == 200:
            for site in res_uso.json().get("result", []):
                sid = str(site.get("site_id", "")).strip().upper()
                if sid in target_sites: 
                    st_val = str(site.get("site_status", "")).strip().capitalize()
                    if st_val.lower() == "other_issue": st_val = "Unmonitor"
                    if sid: site_status_dict[sid] = st_val
    except Exception as e: log_msgs.append(f"❌ Gagal tarik API USO: {e}")
        
    return site_status_dict, log_msgs

def fetch_raw_hourly_via_playwright(username, password, target_date=None, paket_mss=None):
    if not target_date: target_date = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")

    list_paket = []
    if paket_mss:
        if isinstance(paket_mss, (int, str)):
            raw_items = [p.strip() for p in str(paket_mss).split(',') if p.strip()]
            for item in raw_items: list_paket.extend(re.findall(r'\b(?:[1-9]|1[0-5])\b', item))
        list_paket = sorted(list(set(list_paket)), key=lambda x: int(x))

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True, slow_mo=300)
            context = browser.new_context(accept_downloads=True, viewport={'width': 1366, 'height': 768})
            page = context.new_page()

            page.goto("https://baktioss.id/login", timeout=90000)
            page.wait_for_selector('input[type="password"]', timeout=30000)
            page.fill('input[type="text"], input[name="username"], input[placeholder*="sername"], input[placeholder*="mail"]', username)
            page.fill('input[type="password"]', password)
            page.click('button[type="submit"], button:has-text("Log in"), button:has-text("Masuk")')
            page.wait_for_selector('text="Dashboard Operation"', timeout=60000)

            page.goto("https://baktioss.id/site/exports?tab=bts", timeout=90000)
            time.sleep(4)

            page.get_by_text("Site Level", exact=True).locator("visible=true").first.click()
            time.sleep(1)
            page.get_by_text("Hourly", exact=True).locator("visible=true").first.click()
            time.sleep(1)

            if list_paket:
                box_paket_mss = page.locator('.ant-select').filter(has=page.locator('.ant-select-selection-placeholder:has-text("Paket MSS"), .ant-select-selection-item:has-text("Paket MSS")')).first
                if not box_paket_mss.is_visible(): box_paket_mss = page.locator('div.ant-col:has-text("Paket MSS") .ant-select').first
                if not box_paket_mss.is_visible(): box_paket_mss = page.locator('div.ant-select:has([placeholder*="Paket MSS"])').first
                
                if box_paket_mss.is_visible():
                    box_paket_mss.scroll_into_view_if_needed()
                    box_paket_mss.click()
                    time.sleep(1.5)
                    for num in list_paket:
                        page.keyboard.type(str(num), delay=150)
                        time.sleep(1)
                        page.keyboard.press("Enter")
                        time.sleep(1)
                    page.keyboard.press("Escape")
                    time.sleep(1)

            btn_kpi = page.locator('label, span').filter(has_text=re.compile(r"^KPI BTS$", re.IGNORECASE)).locator("visible=true").first
            if not btn_kpi.is_visible(): btn_kpi = page.get_by_text("KPI BTS", exact=True).locator("visible=true").first
            if btn_kpi.is_visible(): btn_kpi.click()
            time.sleep(1)

            date_inputs = page.locator('.ant-picker-input input').locator("visible=true")
            if date_inputs.count() >= 2:
                date_inputs.nth(0).click(force=True)
                page.keyboard.press("Control+A")
                page.keyboard.press("Backspace")
                date_inputs.nth(0).press_sequentially(f"{target_date} 00:00:00", delay=50)
                page.keyboard.press("Enter")
                time.sleep(1)
                
                date_inputs.nth(1).click(force=True)
                page.keyboard.press("Control+A")
                page.keyboard.press("Backspace")
                date_inputs.nth(1).press_sequentially(f"{target_date} 23:00:00", delay=50)
                page.keyboard.press("Enter")
                time.sleep(1)
                
                btn_ok = page.locator('.ant-picker-ok button, button:has-text("OK")').locator("visible=true").last
                if btn_ok.is_visible(): btn_ok.click(force=True)
            time.sleep(1.5)

            btn_ekspor = page.get_by_text("Ekspor", exact=True).locator("visible=true").last
            if not btn_ekspor.is_visible(): btn_ekspor = page.locator('button:has-text("Ekspor"), button:has-text("Export")').locator("visible=true").last
            btn_ekspor.click(force=True)
            time.sleep(2)

            input_file_name = page.locator('.ant-modal-content input, div[role="dialog"] input').locator("visible=true").first
            nama_file_eks = f"Report_KPI_BTS_{target_date}"
            if list_paket: nama_file_eks += "_pkt" + "-".join(list_paket[:4])
            
            waktu_eksekusi = datetime.now()
            if input_file_name.is_visible():
                input_file_name.click()
                input_file_name.fill(nama_file_eks)
                time.sleep(1)
                btn_confirm = page.locator('.ant-modal-content button:has-text("Export"), div[role="dialog"] button:has-text("Export")').locator("visible=true").first
                if btn_confirm.is_visible(): btn_confirm.click(force=True)
                else: page.keyboard.press("Enter")
            else:
                page.keyboard.press("Enter")

            page.wait_for_selector('text="Daftar Permintaan Saya"', timeout=30000)
            target_times = [waktu_eksekusi.strftime("%d/%m/%Y %H:%M"), (waktu_eksekusi - timedelta(minutes=1)).strftime("%d/%m/%Y %H:%M"), (waktu_eksekusi + timedelta(minutes=1)).strftime("%d/%m/%Y %H:%M")]
            
            found_ready, target_row = False, None
            for i in range(25):
                for t_time in target_times:
                    row = page.locator('tr.ant-table-row', has_text=t_time).first
                    if row.is_visible(): target_row = row; break
                
                if target_row and target_row.is_visible():
                    if target_row.locator('button, a, span').filter(has_text="Download").first.is_visible():
                        found_ready = True
                        break
                
                if (i + 1) % 5 == 0:
                    page.reload(wait_until="domcontentloaded")
                    time.sleep(5)
                else:
                    tab_form = page.locator('div[role="tab"]').filter(has_text="Form Permintaan KPI BTS").first
                    tab_daftar = page.locator('div[role="tab"]').filter(has_text="Daftar Permintaan Saya").first
                    if tab_form.is_visible() and tab_daftar.is_visible():
                        tab_form.click(force=True); time.sleep(1); tab_daftar.click(force=True)
                time.sleep(5)

            if not found_ready:
                browser.close()
                return None, "Waktu antrean habis (lebih dari 10 menit)."

            with page.expect_download(timeout=180000) as download_info:
                final_btn = target_row.locator('button, a, span').filter(has_text="Download").first
                final_btn.scroll_into_view_if_needed()
                try: final_btn.click(force=True, timeout=5000)
                except: final_btn.dispatch_event("click")

            download = download_info.value
            save_path = os.path.join(os.getcwd(), download.suggested_filename)
            download.save_as(save_path)
            
            if save_path.endswith('.zip'):
                with zipfile.ZipFile(save_path, 'r') as zip_ref:
                    zip_ref.extractall(os.getcwd())
                    extracted_files = zip_ref.namelist()
                    if extracted_files: save_path = os.path.join(os.getcwd(), extracted_files[0])

            try: df = pd.read_csv(save_path)
            except: df = pd.read_excel(save_path)
                
            browser.close()
            return df, f"Berhasil mengunduh otomatis: {download.suggested_filename}"

    except Exception as e:
        return None, f"Terjadi kesalahan internal bot: {e}"

def extract_dataframe_from_file(uploaded_file):
    try:
        if uploaded_file.name.endswith('.csv'): return pd.read_csv(uploaded_file)
        elif uploaded_file.name.endswith(('.xls', '.xlsx')): return pd.read_excel(uploaded_file)
        elif uploaded_file.name.endswith('.zip'):
            with zipfile.ZipFile(uploaded_file, 'r') as z:
                for file_name in z.namelist():
                    if file_name.endswith('.csv'): return pd.read_csv(z.open(file_name))
                    elif file_name.endswith(('.xls', '.xlsx')): return pd.read_excel(z.open(file_name))
    except Exception as e: st.error(f"Gagal membaca file: {e}")
    return pd.DataFrame()


# ==========================================
# 2. FUNGSI UTILITAS & PERHITUNGAN
# ==========================================
def hitung_jarak_haversine_vec(lat1, lon1, lat2_series, lon2_series):
    R = 6371.0
    lat1_rad, lon1_rad = np.radians(lat1), np.radians(lon1)
    lat2_clean, lon2_clean = lat2_series.fillna(0), lon2_series.fillna(0)
    lat2_rad, lon2_rad = np.radians(lat2_clean.to_numpy()), np.radians(lon2_clean.to_numpy())
    dlat, dlon = lat2_rad - lat1_rad, lon2_rad - lon1_rad
    a = np.sin(dlat / 2.0) ** 2 + np.cos(lat1_rad) * np.cos(lat2_rad) * np.sin(dlon / 2.0) ** 2
    c = 2 * np.arctan2(np.sqrt(a), np.sqrt(1 - a))
    return R * c

def buat_polygon_lingkaran(center_lat, center_lon, radius_km):
    R = 6371.0
    points = []
    for angle in range(0, 360, 5):
        theta = math.radians(angle)
        lat_rad, lon_rad = math.radians(center_lat), math.radians(center_lon)
        dist = radius_km / R
        new_lat = math.asin(math.sin(lat_rad) * math.cos(dist) + math.cos(lat_rad) * math.sin(dist) * math.cos(theta))
        new_lon = lon_rad + math.atan2(math.sin(theta) * math.sin(dist) * math.cos(lat_rad), math.cos(dist) - math.sin(lat_rad) * math.sin(new_lat))
        points.append([math.degrees(new_lon), math.degrees(new_lat)])
    points.append(points[0]) 
    return {"type": "Feature", "geometry": {"type": "Polygon", "coordinates": [points]}}

def hitung_auto_zoom(lat_series, lon_series, active_radius=None):
    if lat_series.empty or lon_series.empty: return 8.0
    lat_min, lat_max = lat_series.min(), lat_series.max()
    lon_min, lon_max = lon_series.min(), lon_series.max()
    max_diff = max(abs(lat_max - lat_min), abs(lon_max - lon_min))
    if active_radius is not None:
        min_deg = ((active_radius + 20) * 2) / 111.0
        max_diff = max(max_diff, min_deg)
    if max_diff == 0: max_diff = 0.5
    zoom = 8.2 - math.log2(max_diff)
    return max(4.0, min(zoom, 12.5))

def temukan_kolom(df, keywords):
    for col in df.columns:
        col_clean = re.sub(r'\s+', ' ', str(col).strip().lower())
        if col_clean in keywords: return col
    for col in df.columns:
        if any(kw in str(col).strip().lower() for kw in keywords): return col
    return None

def clean_to_numeric(series):
    return pd.to_numeric(series.astype(str).str.replace('%', '', regex=False).str.replace(',', '.', regex=False).str.strip(), errors='coerce').fillna(0)

def hitung_ringkasan_wilayah(df):
    prov_col = temukan_kolom(df, ['provinsi', 'province', 'prov'])
    kab_col = temukan_kolom(df, ['kabupaten', 'regency', 'kab', 'kab/kota'])
    kec_col = temukan_kolom(df, ['kecamatan', 'district', 'kec'])
    desa_col = temukan_kolom(df, ['desa', 'kelurahan', 'village'])
    return {
        'n_prov': df[prov_col].dropna().nunique() if prov_col else 0,
        'n_kab': df[kab_col].dropna().nunique() if kab_col else 0,
        'n_kec': df[kec_col].dropna().nunique() if kec_col else 0,
        'n_desa': df[desa_col].dropna().nunique() if desa_col else 0
    }

def hitung_breakdown_site_transmisi(df):
    cat_col = temukan_kolom(df, ['program', 'kategori', 'category', 'jenis_bts', 'bts_type', 'tipe_bts', 'sumber_dana', 'tipe site', 'tipe_site'])
    trans_col = temukan_kolom(df, ['transmisi', 'transmission', 'transport', 'backhaul', 'tipe_transmisi', 'sistem_transmisi'])

    if cat_col and cat_col in df.columns:
        cat_str = df[cat_col].astype(str).str.upper()
        uso_sites, g4_sites = df[cat_str.str.contains('USO', na=False)], df[cat_str.str.contains('4G', na=False)]
    else:
        uso_mask, g4_mask = pd.Series(False, index=df.index), pd.Series(False, index=df.index)
        for c in df.select_dtypes(include=['object']):
            s = df[c].astype(str).str.upper()
            uso_mask |= s.str.contains('USO', na=False)
            g4_mask |= s.str.contains('4G', na=False)
        uso_sites, g4_sites = df[uso_mask], df[g4_mask]

    uso_vsat = uso_mw = g4_vsat = g4_mw = 0
    if trans_col and trans_col in df.columns:
        uso_tr, g4_tr = uso_sites[trans_col].astype(str).str.upper(), g4_sites[trans_col].astype(str).str.upper()
        uso_vsat, uso_mw = len(uso_sites[uso_tr.str.contains('VSAT', na=False)]), len(uso_sites[uso_tr.str.contains('MW|MICROWAVE', regex=True, na=False)])
        g4_vsat, g4_mw = len(g4_sites[g4_tr.str.contains('VSAT', na=False)]), len(g4_sites[g4_tr.str.contains('MW|MICROWAVE', regex=True, na=False)])

    return {'n_uso': len(uso_sites), 'n_4g': len(g4_sites), 'uso_vsat': uso_vsat, 'uso_mw': uso_mw, 'g4_vsat': g4_vsat, 'g4_mw': g4_mw}

def hitung_breakdown_power_tower(df):
    power_col = df.columns[17] if len(df.columns) > 17 else temukan_kolom(df, ['tipe power', 'power', 'sumber_daya', 'catergory_power', 'tipe_power', 'catu_daya'])
    tower_col = df.columns[18] if len(df.columns) > 18 else temukan_kolom(df, ['tipe tower', 'tower', 'tipe_tower', 'tower_type', 'jenis_tower', 'structure'])
    p_summary = df[power_col].astype(str).str.strip().str.title().value_counts().to_dict() if power_col and power_col in df.columns else {}
    t_summary = df[tower_col].astype(str).str.strip().str.upper().value_counts().to_dict() if tower_col and tower_col in df.columns else {}
    return p_summary, t_summary

# ==========================================
# 3. FUNGSI PPTX
# ==========================================
def replace_placeholder_shape_with_image(slide, img_bytes, target_keys, force_width=None, force_height=None):
    target_keys_upper = [k.upper() for k in target_keys]
    for shape in list(slide.shapes):
        if any(k in shape.name.strip().upper() for k in target_keys_upper):
            left, top = shape.left, shape.top
            width, height = force_width if force_width else shape.width, force_height if force_height else shape.height
            slide.shapes.add_picture(img_bytes, left, top, width=width, height=height)
            sp = shape._element
            sp.getparent().remove(sp)
            return True
    return False

def insert_kabupaten_table_to_pptx(slide, kab_summary, left=None, top=None, width=None):
    rows, cols = len(kab_summary) + 2, 9
    pos_left = left if (left is not None and left < Inches(5.0)) else Inches(0.8)
    pos_top = top if top is not None else Inches(1.5)
    pos_width = width if (width is not None and width >= Inches(6.5)) else Inches(7.5)
    pos_height = Inches(0.28 * rows)

    table = slide.shapes.add_table(rows, cols, pos_left, pos_top, pos_width, pos_height).table
    for i, w_ratio in enumerate([0.06, 0.30, 0.09, 0.09, 0.09, 0.09, 0.09, 0.09, 0.09]):
        table.columns[i].width = int(pos_width * w_ratio)

    def format_cell(cell, text, bold=False, size_pt=8, color_rgb=None, bg_rgb=None):
        cell.text, cell.margin_left, cell.margin_right, cell.margin_top, cell.margin_bottom = text, Inches(0.04), Inches(0.04), Inches(0.02), Inches(0.02)
        if bg_rgb: cell.fill.solid(); cell.fill.fore_color.rgb = bg_rgb
        p = cell.text_frame.paragraphs[0]
        p.alignment, p.font.bold, p.font.size = PP_ALIGN.CENTER, bold, Pt(size_pt)
        if color_rgb: p.font.color.rgb = color_rgb

    for c_idx, text in enumerate(["No", "Kabupaten", "Total Site", "BTS 4G", "BTS USO", "Status BTS 4G", "", "Status BTS USO", ""]):
        format_cell(table.cell(0, c_idx), text, bold=True, size_pt=9, color_rgb=RGBColor(255, 255, 255), bg_rgb=RGBColor(0, 153, 153))
    table.cell(0, 5).merge(table.cell(0, 6)); table.cell(0, 7).merge(table.cell(0, 8))

    for c_idx, text in enumerate(["", "", "", "", "", "Site Up", "Site Down", "Site Up", "Site Down"]):
        format_cell(table.cell(1, c_idx), text, bold=True, size_pt=8, color_rgb=RGBColor(255, 255, 255), bg_rgb=RGBColor(0, 153, 153))
    for c_idx in range(5): table.cell(0, c_idx).merge(table.cell(1, c_idx))

    for r_idx, row in kab_summary.iterrows():
        vals = [str(r_idx + 1), str(row['Kabupaten']), str(int(row['Total Site'])), str(int(row['BTS 4G'])), str(int(row['BTS USO'])), str(int(row['4G_Up'])), str(int(row['4G_Down'])), str(int(row['USO_Up'])), str(int(row['USO_Down']))]
        for c_idx, val in enumerate(vals):
            c_color = RGBColor(211, 47, 47) if (c_idx in [6, 8] and int(val) > 0) else None
            format_cell(table.cell(r_idx + 2, c_idx), val, bold=(c_idx in [1, 6, 8] and (c_idx == 1 or int(val) > 0)), size_pt=8, color_rgb=c_color, bg_rgb=RGBColor(248, 250, 251) if (r_idx % 2 == 1) else RGBColor(255, 255, 255))

def generate_standard_pptx_report(jenis_kejadian, detail_kejadian, waktu_kejadian, loc_label, mode_label, df_result, fig_map=None, fig_avail=None, fig_traffic=None, fig_user=None):
    template_path = "template_baku.pptx.pptx" if (not os.path.exists("template_baku.pptx") and os.path.exists("template_baku.pptx.pptx")) else "template_baku.pptx"
    prs = Presentation(template_path) if os.path.exists(template_path) else Presentation()
    if not os.path.exists(template_path):
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
        for _ in range(4): prs.slides.add_slide(prs.slide_layouts[6])

    if 'status' not in df_result.columns: df_result['status'] = 'Unmonitor'
    status_counts = df_result['status'].astype(str).str.lower().value_counts()
    bts_down, bts_up, unmonitor = status_counts.get('down', 0), status_counts.get('up', 0), len(df_result) - status_counts.get('down', 0) - status_counts.get('up', 0)

    total_traffic_mb = df_result['traffic_mb'].sum() if 'traffic_mb' in df_result.columns else 0.0
    total_period_gb = (df_result['total_traffic_mb_period'].sum() if 'total_traffic_mb_period' in df_result.columns else total_traffic_mb) / 1024.0

    up_stats = hitung_breakdown_site_transmisi(df_result[df_result['status'].str.lower() == 'up'])
    down_stats = hitung_breakdown_site_transmisi(df_result[df_result['status'].str.lower() == 'down'])
    unmon_stats = hitung_breakdown_site_transmisi(df_result[~df_result['status'].str.lower().isin(['up', 'down'])])
    site_info = hitung_breakdown_site_transmisi(df_result)
    wilayah_info = hitung_ringkasan_wilayah(df_result)
    
    cat_col = temukan_kolom(df_result, ['program', 'kategori', 'category', 'jenis_bts', 'bts_type', 'tipe_bts', 'tipe site'])
    p_4g, t_4g = hitung_breakdown_power_tower(df_result[df_result[cat_col].astype(str).str.upper().str.contains('4G', na=False)] if cat_col else pd.DataFrame())
    p_uso, t_uso = hitung_breakdown_power_tower(df_result[df_result[cat_col].astype(str).str.upper().str.contains('USO', na=False)] if cat_col else pd.DataFrame())

    timeline_df = st.session_state.get("timeline_df", pd.DataFrame())
    avg_avail_val = 0.0
    str_timestamp_raw, str_tgl_raw, str_jam_raw = "-", "-", "-"

    if not timeline_df.empty:
        trend_raw = timeline_df[timeline_df['site_id_clean'].isin(df_result['site_id_clean'].unique())].copy()
        if not trend_raw.empty:
            trend_raw['hours_count'] = trend_raw.groupby(['site_id_clean', trend_raw['parsed_time'].dt.date])['parsed_time'].transform(lambda x: x.dt.hour.nunique())
            v_df = trend_raw[trend_raw['hours_count'] >= 24]
            avg_avail_val = v_df['avail_num'].mean() if not v_df.empty else trend_raw['avail_num'].mean()
            unique_times = sorted(trend_raw['parsed_time'].dropna().unique())
            latest_raw_time = unique_times[-2] if len(unique_times) >= 2 else (unique_times[-1] if len(unique_times) == 1 else None)
            if latest_raw_time:
                str_timestamp_raw, str_tgl_raw, str_jam_raw = latest_raw_time.strftime("%d/%m/%Y %H:%M WIB"), latest_raw_time.strftime("%d/%m/%Y"), latest_raw_time.strftime("%H:%M WIB")
    else:
        avg_avail_val = (bts_up / (bts_up + bts_down) * 100.0) if (bts_up + bts_down) > 0 else 0.0

    kab_col_summary = temukan_kolom(df_result, ['kabupaten', 'regency', 'kab', 'kab/kota'])
    top_kab_name, top_kab_down_total, top_kab_4g_down, top_kab_uso_down = "-", 0, 0, 0
    df_down = df_result[df_result['status'].str.lower() == 'down']
    
    if kab_col_summary and bts_down > 0 and not df_down.empty:
        top_kab_series = df_down[kab_col_summary].value_counts()
        if not top_kab_series.empty:
            top_kab_name, top_kab_down_total = str(top_kab_series.index[0]), int(top_kab_series.iloc[0])
            df_top_kab = df_down[df_down[kab_col_summary] == top_kab_name]
            if cat_col and cat_col in df_top_kab.columns:
                cat_str = df_top_kab[cat_col].astype(str).str.upper()
                top_kab_4g_down, top_kab_uso_down = len(df_top_kab[cat_str.str.contains('4G', na=False)]), len(df_top_kab[cat_str.str.contains('USO', na=False)])

    max_user_sum = int(df_result.get('max_user', pd.Series([0])).sum())
    
    replacement_dict = {
        "{{jenis_kejadian}}": str(jenis_kejadian), "{{detail_kejadian}}": str(detail_kejadian), "{{waktu_kejadian}}": str(waktu_kejadian),
        "{{lokasi}}": str(loc_label), "{{mode}}": str(mode_label), "{{timestamp_data}}": str_timestamp_raw, "{{timestamp_raw}}": str_timestamp_raw,
        "{{total_site}}": str(len(df_result)), "{{site_up}}": str(bts_up), "{{site_down}}": str(bts_down), "{{unmonitor}}": str(unmonitor),
        "{{total_traffic_gb}}": f"{total_period_gb:,.2f}", "{{daily_traffic_mb}}": f"{total_traffic_mb:,.2f}", "{{daily_traffic_gb}}": f"{(total_traffic_mb/1024.0):,.2f}",
        "{{active_users}}": f"{max_user_sum:,}", "{{avg_avail}}": f"{avg_avail_val:.2f}%",
        "{{n_prov}}": str(wilayah_info['n_prov']), "{{n_kab}}": str(wilayah_info['n_kab']), "{{n_kec}}": str(wilayah_info['n_kec']), "{{n_desa}}": str(wilayah_info['n_desa']),
        "{{n_4g}}": str(site_info['n_4g']), "{{g4_vsat}}": str(site_info['g4_vsat']), "{{g4_mw}}": str(site_info['g4_mw']),
        "{{n_uso}}": str(site_info['n_uso']), "{{uso_vsat}}": str(site_info['uso_vsat']), "{{uso_mw}}": str(site_info['uso_mw']),
        "{{4g_up}}": str(up_stats['n_4g']), "{{uso_up}}": str(up_stats['n_uso']), "{{4g_up_vsat}}": str(up_stats['g4_vsat']), "{{4g_up_mw}}": str(up_stats['g4_mw']),
        "{{uso_up_vsat}}": str(up_stats['uso_vsat']), "{{uso_up_mw}}": str(up_stats['uso_mw']), "{{4g_down}}": str(down_stats['n_4g']), "{{uso_down}}": str(down_stats['n_uso']),
        "{{4g_down_vsat}}": str(down_stats['g4_vsat']), "{{4g_down_mw}}": str(down_stats['g4_mw']), "{{uso_down_vsat}}": str(down_stats['uso_vsat']), "{{uso_down_mw}}": str(down_stats['uso_mw']),
        "{{top_kab_down}}": str(top_kab_name), "{{top_kab_down_count}}": str(top_kab_down_total), "{{top_kab_down_4g}}": str(top_kab_4g_down), "{{top_kab_down_uso}}": str(top_kab_uso_down),
        "{{tipe_power_4g}}": ", ".join([f"{k}:{v}" for k, v in p_4g.items()]) if p_4g else "-", "{{tipe_power_uso}}": ", ".join([f"{k}:{v}" for k, v in p_uso.items()]) if p_uso else "-",
        "{{tipe_tower_4g}}": ", ".join([f"{k}:{v}" for k, v in t_4g.items()]) if t_4g else "-", "{{tipe_tower_uso}}": ", ".join([f"{k}:{v}" for k, v in t_uso.items()]) if t_uso else "-",
        "{{4g_unmon_vsat}}": str(unmon_stats['g4_vsat']), "{{4g_unmon_mw}}": str(unmon_stats['g4_mw']), "{{uso_unmon_vsat}}": str(unmon_stats['uso_vsat']), "{{uso_unmon_mw}}": str(unmon_stats['uso_mw'])
    }

    def proses_shape(shape):
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                p_text = paragraph.text.replace('\u200b', '').replace('\u200d', '')
                if any(key in p_text for key in replacement_dict.keys()):
                    for key, val in replacement_dict.items(): p_text = p_text.replace(key, val)
                    if paragraph.runs:
                        for i in range(len(paragraph.runs)): paragraph.runs[i].text = "" 
                        paragraph.runs[0].text = p_text 
                    else: paragraph.text = p_text
        elif shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells: proses_shape(cell)
        elif shape.shape_type == 6: 
            for sub_shape in shape.shapes: proses_shape(sub_shape)

    for slide in prs.slides:
        for shape in slide.shapes: proses_shape(shape)

    try:
        if len(prs.slides) >= 2 and fig_map is not None:
            slide_map = prs.slides[1]
            map_bytes = BytesIO(pio.to_image(fig_map, format="png", width=800, height=676, scale=2))
            if not replace_placeholder_shape_with_image(slide_map, map_bytes, ["PH_MAP", "PETA", "MAP"], force_width=Inches(5.88), force_height=Inches(4.97)):
                slide_map.shapes.add_picture(map_bytes, Inches(0.5), Inches(1.5), width=Inches(5.88), height=Inches(4.97))

        if len(prs.slides) >= 3:
            slide_chart = prs.slides[2]
            if fig_traffic is not None: replace_placeholder_shape_with_image(slide_chart, BytesIO(pio.to_image(fig_traffic, format="png", width=800, height=450, scale=2)), ["PH_TRAFFIC"])
            if fig_user is not None: replace_placeholder_shape_with_image(slide_chart, BytesIO(pio.to_image(fig_user, format="png", width=800, height=450, scale=2)), ["PH_USER"])
            if fig_avail is not None: replace_placeholder_shape_with_image(slide_chart, BytesIO(pio.to_image(fig_avail, format="png", width=800, height=450, scale=2)), ["PH_AVAIL"])
    except Exception as e: print(f"Render grafik PPTX dilewati: {e}")

    if kab_col_summary and kab_col_summary in df_result.columns:
        df_calc_ppt = df_result.copy()
        df_calc_ppt['is_4g'] = df_calc_ppt[cat_col].astype(str).str.upper().str.contains('4G', na=False) if cat_col else False
        df_calc_ppt['is_uso'] = df_calc_ppt[cat_col].astype(str).str.upper().str.contains('USO', na=False) if cat_col else False
        st_lower = df_calc_ppt['status'].astype(str).str.lower()
        df_calc_ppt['is_up'], df_calc_ppt['is_down'] = st_lower == 'up', st_lower == 'down'

        kab_summary_ppt = df_calc_ppt.groupby(kab_col_summary).apply(lambda g: pd.Series({
            'Total Site': len(g), 'BTS 4G': len(g[g['is_4g']]), 'BTS USO': len(g[g['is_uso']]),
            '4G_Up': len(g[g['is_4g'] & g['is_up']]), '4G_Down': len(g[g['is_4g'] & g['is_down']]),
            'USO_Up': len(g[g['is_uso'] & g['is_up']]), 'USO_Down': len(g[g['is_uso'] & g['is_down']]),
        })).reset_index().rename(columns={kab_col_summary: 'Kabupaten'}).sort_values(by=['4G_Down', 'USO_Down', 'Total Site'], ascending=[False, False, False]).reset_index(drop=True)

        target_slide = prs.slides[3] if len(prs.slides) >= 4 else prs.slides.add_slide(prs.slide_layouts[6])
        target_shape = next((shape for slide in prs.slides for shape in slide.shapes if "PH_TABEL" in shape.name.strip().upper() or "PH_TABLE" in shape.name.strip().upper()), None)

        if target_shape:
            pos_left, pos_top, pos_width = target_shape.left, target_shape.top, target_shape.width
            target_shape._element.getparent().remove(target_shape._element)
            insert_kabupaten_table_to_pptx(target_slide, kab_summary_ppt, left=pos_left, top=pos_top, width=pos_width)
        else: insert_kabupaten_table_to_pptx(target_slide, kab_summary_ppt)

    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer


# ==========================================
# 4. INISIALISASI SESSION STATE
# ==========================================
for key in ["result_df", "timeline_df", "raw_hourly_df"]:
    if key not in st.session_state: st.session_state[key] = pd.DataFrame()
for key in ["center_coords", "loc_label", "mode_label", "generated_pptx_bytes", "active_radius", "filter_mode"]:
    if key not in st.session_state: st.session_state[key] = None if "coords" in key or "bytes" in key or "radius" in key else ""
for key in ["jenis_kejadian", "detail_kejadian", "waktu_kejadian"]:
    if key not in st.session_state: st.session_state[key] = ""
if "api_status_dict" not in st.session_state: st.session_state["api_status_dict"] = {}


# ==========================================
# 5. STREAMLIT UI & LOGIC UTAMA
# ==========================================
st.set_page_config(page_title="Flash Report Bencana", layout="wide")
st.title("Dashboard Laporan Dampak Kejadian")
st.caption("Kementerian Komunikasi dan Digital — BAKTI")

# --- SIDEBAR 1: DATA MASTER ---
st.sidebar.header("1. Upload Data Master")
file_master = st.sidebar.file_uploader("1. Upload Master Site (Koordinat):", type=["xlsx", "xls", "csv"])

df_master = pd.DataFrame()
if file_master is not None:
    df_master = extract_dataframe_from_file(file_master)
    site_col_master = temukan_kolom(df_master, ['site id', 'site_id', 'id site', 'id_site', 'site', 'siteid'])
    
    if site_col_master:
        df_master['site_id_clean'] = df_master[site_col_master].astype(str).str.strip().str.upper()
        
        # Cegah hilangnya status pada saat Streamlit Rerun
        stored_api_dict = st.session_state.get("api_status_dict", {})
        if stored_api_dict:
            df_master['status'] = df_master['site_id_clean'].map(stored_api_dict).fillna('Unmonitor')
        else:
            df_master['status'] = 'Unmonitor'
            
        for c in ['traffic_mb', 'total_traffic_mb_period', 'max_user']:
            if c not in df_master.columns: df_master[c] = 0.0

# --- SIDEBAR 2: PENGATURAN AREA & 1-CLICK AUTOMATION ---
st.sidebar.markdown("---")
st.sidebar.header("2. Area & Automasi (1-Click)")
main_category = st.sidebar.radio("Metode Filter Area:", ["Pilih Berdasarkan Area / Identitas Site", "Input Koordinat Manual", "Data Live Gempa BMKG"])

if not df_master.empty:
    lat_col = temukan_kolom(df_master, ['latitude', 'lat', 'y'])
    lon_col = temukan_kolom(df_master, ['longitude', 'long', 'lon', 'x'])

    if lat_col and lon_col:
        df_master[lat_col] = pd.to_numeric(df_master[lat_col], errors='coerce')
        df_master[lon_col] = pd.to_numeric(df_master[lon_col], errors='coerce')

        temp_center_coords, temp_result_df, temp_loc_label, temp_mode_label, temp_active_radius, temp_filter_mode = None, pd.DataFrame(), "", "", None, ""
        temp_jenis, temp_detail, temp_waktu = "", "", ""

        if main_category == "Data Live Gempa BMKG":
            if st.sidebar.button("⬇️ Tarik Riwayat BMKG"):
                try:
                    resp = requests.get("https://data.bmkg.go.id/DataMKG/TEWS/gempaterkini.json", timeout=10)
                    if resp.status_code == 200: st.session_state["bmkg_list"] = resp.json()['Infogempa']['gempa']
                except: st.sidebar.error("Gagal terhubung ke BMKG.")

            if "bmkg_list" in st.session_state:
                gempa_opt = {f"{g['Tanggal']} {g['Jam']} | {g['Magnitude']} M - {g['Wilayah']}": g for g in st.session_state["bmkg_list"]}
                data_bmkg = gempa_opt[st.sidebar.selectbox("Pilih Gempa:", list(gempa_opt.keys()))]
                bmkg_tgl_jam, bmkg_mag, bmkg_wilayah, bmkg_kedalaman = f"{data_bmkg['Tanggal']} {data_bmkg['Jam']}", data_bmkg['Magnitude'], data_bmkg['Wilayah'], data_bmkg['Kedalaman']
                bmkg_lat, bmkg_lon = float(data_bmkg['Coordinates'].split(",")[0]), float(data_bmkg['Coordinates'].split(",")[1])
                st.sidebar.info(f"📍 **{bmkg_wilayah}**\n\n🧲 {bmkg_mag} Mag")
                
                radius_km = st.sidebar.slider("Radius Terdampak (KM):", 1, 100, 10)
                df_master['jarak_km'] = hitung_jarak_haversine_vec(bmkg_lat, bmkg_lon, df_master[lat_col], df_master[lon_col])
                
                temp_center_coords = (bmkg_lat, bmkg_lon)
                temp_result_df = df_master[df_master['jarak_km'] <= radius_km].copy()
                temp_loc_label, temp_mode_label = bmkg_wilayah, f"Radius {radius_km} KM"
                temp_active_radius, temp_filter_mode = radius_km, "Radius Jarak (KM) dari Pusat Area"
                temp_jenis, temp_detail, temp_waktu = "Gempa Bumi", f"{bmkg_mag} Mag, Kedalaman {bmkg_kedalaman}", bmkg_tgl_jam

        else:
            in_jenis = st.sidebar.text_input("Kejadian:", st.session_state.get("jenis_kejadian", "Banjir Bandang"))
            in_detail = st.sidebar.text_input("Detail:", st.session_state.get("detail_kejadian", "Cuaca Ekstrem"))
            in_waktu = st.sidebar.text_input("Waktu:", st.session_state.get("waktu_kejadian", "Oktober 2026"))

            if main_category == "Pilih Berdasarkan Area / Identitas Site":
                kategori_area = st.sidebar.selectbox("Kategori Pencarian:", ["Kabupaten", "Kecamatan", "Desa", "Provinsi", "Site ID"])
                target_col = {"Provinsi": temukan_kolom(df_master, ['provinsi', 'province', 'prov']), "Kabupaten": temukan_kolom(df_master, ['kabupaten', 'regency', 'kab']), "Kecamatan": temukan_kolom(df_master, ['kecamatan', 'district', 'kec']), "Desa": temukan_kolom(df_master, ['desa', 'kelurahan', 'village']), "Site ID": 'site_id_clean'}.get(kategori_area)

                if target_col and target_col in df_master.columns:
                    list_options = sorted(df_master[target_col].dropna().astype(str).unique().tolist())
                    selected_val = st.sidebar.selectbox(f"Pilih {kategori_area}:", options=["-- Pilih --"] + list_options)
                    filter_mode = st.sidebar.radio("Cakupan Area:", ["Batas Administrasi Murni (Eksak)", "Radius Jarak (KM) dari Pusat Area"])
                    radius_km = st.sidebar.slider("Radius (KM):", 1, 100, 10) if filter_mode == "Radius Jarak (KM) dari Pusat Area" else 10

                    if selected_val != "-- Pilih --":
                        matched_df = df_master[df_master[target_col].astype(str).str.lower() == selected_val.lower()]
                        if not matched_df.empty:
                            center_lat, center_lon = matched_df[lat_col].mean(), matched_df[lon_col].mean()
                            df_master['jarak_km'] = hitung_jarak_haversine_vec(center_lat, center_lon, df_master[lat_col], df_master[lon_col])
                            
                            temp_center_coords = (center_lat, center_lon)
                            temp_result_df = matched_df.copy() if filter_mode == "Batas Administrasi Murni (Eksak)" else df_master[df_master['jarak_km'] <= radius_km].copy()
                            temp_loc_label = f"{kategori_area} {selected_val}"
                            temp_mode_label = f"Batas Administrasi Murni ({kategori_area}: {selected_val})" if filter_mode == "Batas Administrasi Murni (Eksak)" else f"Radius {radius_km} KM dari Pusat {selected_val}"
                            temp_active_radius, temp_filter_mode = (None if filter_mode == "Batas Administrasi Murni (Eksak)" else radius_km), filter_mode
                            temp_jenis, temp_detail, temp_waktu = in_jenis, in_detail, in_waktu

            elif main_category == "Input Koordinat Manual":
                c_l, c_r = st.sidebar.columns(2)
                with c_l: input_lat = st.number_input("Latitude:", format="%.6f", value=-4.512300)
                with c_r: input_lon = st.number_input("Longitude:", format="%.6f", value=140.412300)
                radius_km = st.sidebar.slider("Radius Terdampak (KM):", 1, 100, 10)
                
                df_master['jarak_km'] = hitung_jarak_haversine_vec(input_lat, input_lon, df_master[lat_col], df_master[lon_col])
                temp_center_coords = (input_lat, input_lon)
                temp_result_df = df_master[df_master['jarak_km'] <= radius_km].copy()
                temp_loc_label, temp_mode_label = f"Koordinat ({input_lat}, {input_lon})", f"Radius {radius_km} KM"
                temp_active_radius, temp_filter_mode = radius_km, "Radius Jarak (KM) dari Pusat Area"
                temp_jenis, temp_detail, temp_waktu = in_jenis, in_detail, in_waktu

        with st.sidebar.expander("🔑 Konfigurasi Akun BAKTI OSS", expanded=False):
            pl_user = st.text_input("Username:", "ilham.dyaksa")
            pl_pass = st.text_input("Password:", type="password", value="Bakti@2024")
            token_uso = st.text_input("Token USO:", type="password", value="eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJmcmVzaCI6ZmFsc2UsImlhdCI6MTc4NTQ4ODYwMSwianRpIjoiY2VhMmM1YTAtZWJmYi00ZDk1LThiMGMtMzUxYzQwODliMzliIiwidHlwZSI6ImFjY2VzcyIsInN1YiI6eyJ1c2VyX2lkIjo0MjUsInVzZXJuYW1lIjoiaXRqZW5fYWRtaW4iLCJmdWxsbmFtZSI6Ikl0amVuIEFkbWluIiwiZW1haWwiOiJpdGplbl9lbWFpbEBkZWZhdWx0LmNvbSIsInJvbGVzIjpbIkFsbCBSZXBvcnQiLCJIaXN0b3J5IiwiSW52ZW50b3J5IFNpdGUiLCJJbnZlbnRvcnkgU291cmNlIiwiSXRqZW4gS29taW5mbyIsIlJlcG9ydCBMb3cgU0xBIiwiUmVwb3J0IFBlcmZvcm1hbmNlIERlZ3JhZGVkIiwiUmVwb3J0IFJlcGV0aXRpdmUiLCJSZXBvcnQgUmVwZXRpdGl2ZSBTdW1tYXJ5IiwiUmVwb3J0IFNpdGUgTFZEIERvd24iLCJSZXBvcnQgU2l0ZSBMaXN0IiwiUmVwb3J0IFNpdGUgUG93ZXIgRG93biIsIlJlcG9ydCBTaXRlIFBvd2VyIFByb2JsZW0iLCJSZXBvcnQgU2l0ZSBTb21vIiwiUmVwb3J0IFNpdGUgVHJhZmZpYyBCdXN5IEhvdXIiLCJSZXBvcnQgU2l0ZSBUcmFmZmljIE1heCBBdmVyYWdlIiwiUmVwb3J0IFNpdGUgVnNhdCBEb3duIiwiUmVwb3J0IFN0YXRlIENvZGUiLCJSZXBvcnQgV2FybmluZyBMZXR0ZXIiLCJSZXBvcnQgWmVybyBUcmFmZmljIiwiU2l0ZSBFeHBvcnQiLCJTaXRlIFNlbnNvciBFeHBvcnQiXSwibGFuZGluZ19wYWdlX3BhdGgiOiIvZGFzaGJvYXJkIiwiZGFzaGJvYXJkX2lkcyI6WzEsMzEsMzMsNTddLCJwcm92aWRlcl9pZHMiOltdLCJjZWxsdWxhcl9vcGVyYXRvcl9pZHMiOltdLCJncm91cF9pZCI6MTYyLCJncm91cF9uYW1lIjoiSXRqZW4ifSwibmJmIjoxNzg1NDg4NjAxLCJjc3JmIjoiMWZjN2U1MzEtM2ZhOC00MjVjLTk3YjQtOGNlY2FiOTY5ZmIxIiwiZXhwIjoyMTAwODQ4NjAxfQ.kb67aTVP23JVhZ25I7zjcJ9ptt34O9FQnEGfbYqFI4Y")
            pl_date = st.date_input("Target Tgl Data:", datetime.now() - timedelta(days=1))

        if not temp_result_df.empty:
            st.sidebar.markdown(f"🎯 **Terdeteksi {len(temp_result_df)} Site Terdampak**")
            
            if st.sidebar.button("🚀 Cari Area & Tarik Semua Data (Scraping + API)", type="primary"):
                st.session_state.update({
                    "center_coords": temp_center_coords, "result_df": temp_result_df.copy(), "loc_label": temp_loc_label, 
                    "mode_label": temp_mode_label, "active_radius": temp_active_radius, "filter_mode": temp_filter_mode,
                    "jenis_kejadian": temp_jenis, "detail_kejadian": temp_detail, "waktu_kejadian": temp_waktu
                })
                
                with st.status("📡 Memulai Eksekusi 1-Click Automation...", expanded=True) as status:
                    
                    st.write("🌍 BAKTI 3T Tertinggal, Terdepan, Terluar, PASTI BISA !!!!")
                    
                    auto_paket_str = ""
                    paket_col = temukan_kolom(st.session_state["result_df"], ['paket penyedia mss', 'paket mss', 'paket'])
                    if paket_col:
                        raw_pakets = st.session_state["result_df"][paket_col].dropna().astype(str).tolist()
                        extracted_nums = []
                        for val in raw_pakets: extracted_nums.extend(re.findall(r'\b(?:[1-9]|1[0-5])\b', val))
                        if extracted_nums: auto_paket_str = ", ".join(sorted(list(set(extracted_nums)), key=lambda x: int(x)))
                    
                    st.write(f"⏳ [1/3] Menjalankan bot Chromium untuk mengunduh paket MSS [{auto_paket_str}]...")
                    df_scraped, msg = fetch_raw_hourly_via_playwright(pl_user, pl_pass, target_date=pl_date.strftime("%Y-%m-%d"), paket_mss=auto_paket_str)
                    
                    if df_scraped is not None:
                        st.session_state["raw_hourly_df"] = df_scraped
                        st.write("✅ Scraping data trafik Hourly selesai!")
                    else:
                        st.write(f"⚠️ Scraping gagal/dilewati: {msg}")

                    st.write("⏳ [2/3] Mengakses API server BAKTI OSS & USO untuk status Live...")
                    target_sites = tuple(st.session_state["result_df"]['site_id_clean'].unique().tolist())
                    
                    api_status_dict, err_logs = fetch_live_api_status(pl_user, pl_pass, token_uso, target_sites)
                    st.session_state["api_status_dict"] = api_status_dict
                    df_master['status'] = df_master['site_id_clean'].map(api_status_dict).fillna('Unmonitor')
                    st.session_state["result_df"]['status'] = st.session_state["result_df"]['site_id_clean'].map(api_status_dict).fillna('Unmonitor')
                    
                    if err_logs:
                        for err in err_logs: st.sidebar.warning(err)
                    st.write("✅ Sinkronisasi API Live selesai (100% Anti-Timeout)!")

                    st.write("⏳ [3/3] Memproses kalkulasi dampak dan merapikan DataFrame...")
                    if not st.session_state["raw_hourly_df"].empty:
                        df_status = st.session_state["raw_hourly_df"]
                        site_col_status = temukan_kolom(df_status, ['site id', 'site_id', 'id site', 'id_site', 'site'])
                        avail_col = temukan_kolom(df_status, ['availability (%)', 'availability(%)', 'availability', 'avail'])
                        traffic_col = temukan_kolom(df_status, ['traffic/payload (mb)', 'payload (mb)', 'traffic (mb)', 'payload_mb'])
                        user_col = temukan_kolom(df_status, ['max active user', 'max_active_user', 'active user', 'user', 'peak user'])

                        if site_col_status and avail_col:
                            df_status['site_id_clean'] = df_status[site_col_status].astype(str).str.strip().str.upper()
                            df_status['avail_num'] = clean_to_numeric(df_status[avail_col])
                            df_status['traffic_num'] = clean_to_numeric(df_status[traffic_col]) if traffic_col else 0.0
                            df_status['user_num'] = clean_to_numeric(df_status[user_col]) if user_col else 0.0

                            combined_dt = df_status.iloc[:, 0].astype(str).str.strip() + " " + df_status.iloc[:, 1].astype(str).str.strip()
                            df_status['parsed_time'] = pd.to_datetime(combined_dt, errors='coerce', dayfirst=False)
                            if df_status['parsed_time'].isna().all(): df_status['parsed_time'] = pd.to_datetime(combined_dt, errors='coerce')
                            
                            df_status = df_status.dropna(subset=['parsed_time'])
                            if not df_status.empty:
                                df_status = df_status[df_status['parsed_time'] >= (df_status['parsed_time'].max() - pd.Timedelta(days=4))].copy()

                            df_status['date_str'] = df_status['parsed_time'].dt.strftime('%Y-%m-%d')
                            st.session_state["timeline_df"] = df_status[['site_id_clean', 'parsed_time', 'avail_num', 'traffic_num', 'user_num']].copy()

                            df_metrics = df_status.groupby(['site_id_clean', 'date_str']).agg(total_traffic_day=('traffic_num', 'sum'), peak_user_day=('user_num', 'max')).reset_index().groupby('site_id_clean').agg(traffic_mb=('total_traffic_day', 'mean'), total_traffic_mb_period=('total_traffic_day', 'sum'), max_user=('peak_user_day', 'mean')).reset_index()

                            for col_dup in ['traffic_mb', 'total_traffic_mb_period', 'max_user']: 
                                if col_dup in st.session_state["result_df"].columns: 
                                    st.session_state["result_df"] = st.session_state["result_df"].drop(columns=[col_dup])

                            st.session_state["result_df"] = pd.merge(st.session_state["result_df"], df_metrics, on='site_id_clean', how='left')
                            for col in ['traffic_mb', 'total_traffic_mb_period', 'max_user']: 
                                st.session_state["result_df"][col] = st.session_state["result_df"][col].fillna(0.0)
                    else:
                        st.session_state["timeline_df"] = pd.DataFrame()
                        st.session_state["result_df"]['traffic_mb'] = st.session_state["result_df"]['total_traffic_mb_period'] = st.session_state["result_df"]['max_user'] = 0.0

                    status.update(label="🎉 Seluruh Proses Selesai! Menampilkan Dashboard...", state="complete", expanded=False)
                    st.rerun()


# ==========================================
# 6. TAMPILAN DASHBOARD UTAMA
# ==========================================
result_df = st.session_state["result_df"]

# Pastikan status API selalu dimapping ulang saat terjadi rerun (seperti klik tombol PPT)
stored_api_dict = st.session_state.get("api_status_dict", {})
if stored_api_dict and not result_df.empty:
    result_df['status'] = result_df['site_id_clean'].map(stored_api_dict).fillna('Unmonitor')
elif not result_df.empty and 'status' not in result_df.columns:
    result_df['status'] = 'Unmonitor'

if not result_df.empty:
    j_kejadian, d_kejadian, w_kejadian = st.session_state["jenis_kejadian"], st.session_state["detail_kejadian"], st.session_state["waktu_kejadian"]
    st.success(f"**Ringkasan Laporan:** {j_kejadian} - {d_kejadian} | {w_kejadian}")
    st.markdown(f"📍 **Lokasi Fokus:** {st.session_state['loc_label']} | **Metode:** {st.session_state['mode_label']}")
    
    site_info = hitung_breakdown_site_transmisi(result_df)
    cat_col = temukan_kolom(result_df, ['program', 'kategori', 'category', 'jenis_bts', 'bts_type', 'tipe_bts', 'tipe site'])
    df_uso = result_df[result_df[cat_col].astype(str).str.upper().str.contains('USO', na=False)] if cat_col else pd.DataFrame()
    df_4g = result_df[result_df[cat_col].astype(str).str.upper().str.contains('4G', na=False)] if cat_col else pd.DataFrame()
    power_4g, tower_4g = hitung_breakdown_power_tower(df_4g)
    power_uso, tower_uso = hitung_breakdown_power_tower(df_uso)

    df_up = result_df[result_df['status'].str.lower() == 'up']
    df_down = result_df[result_df['status'].str.lower() == 'down']
    df_unmon = result_df[~result_df['status'].str.lower().isin(['up', 'down'])]

    bts_down, bts_up, unmonitor = len(df_down), len(df_up), len(df_unmon)
    
    up_stats = hitung_breakdown_site_transmisi(df_up)
    down_stats = hitung_breakdown_site_transmisi(df_down)
    unmon_stats = hitung_breakdown_site_transmisi(df_unmon)

    st.subheader("⚡ Status Jaringan Terdampak (Realtime API)")
    c1, c2, c3, c4 = st.columns(4)
    
    c1.metric("Total Site", len(result_df))
    c1.caption(f"4G: **{site_info['n_4g']}** | USO: **{site_info['n_uso']}**")
    
    c2.metric("Site Up", bts_up)
    c2.caption(f"4G: **{up_stats['n_4g']}** | USO: **{up_stats['n_uso']}**")
    
    c3.metric("Site Down", bts_down, delta="- Kritis" if bts_down > 0 else "Normal", delta_color="inverse")
    c3.caption(f"4G: **{down_stats['n_4g']}** | USO: **{down_stats['n_uso']}**")
    
    c4.metric("Unmonitor", unmonitor)
    c4.caption(f"4G: **{unmon_stats['n_4g']}** | USO: **{unmon_stats['n_uso']}**")
    
    st.markdown("---")
    
    st.subheader("📡 Kategori Transmisi & Infrastruktur")
    col_4g, col_uso = st.columns(2)
    with col_4g:
        st.markdown("#### **BTS 4G**")
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Site", site_info['n_4g']); m2.metric("VSAT", site_info['g4_vsat']); m3.metric("MW", site_info['g4_mw'])
        st.caption("⚡ Tipe Power: " + (", ".join([f"{k}:{v}" for k,v in power_4g.items()]) if power_4g else "-"))
        st.caption("🗼 Tipe Tower: " + (", ".join([f"{k}:{v}" for k,v in tower_4g.items()]) if tower_4g else "-"))
    with col_uso:
        st.markdown("#### **BTS USO**")
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Site", site_info['n_uso']); m2.metric("VSAT", site_info['uso_vsat']); m3.metric("MW", site_info['uso_mw'])
        st.caption("⚡ Tipe Power: " + (", ".join([f"{k}:{v}" for k,v in power_uso.items()]) if power_uso else "-"))
        st.caption("🗼 Tipe Tower: " + (", ".join([f"{k}:{v}" for k,v in tower_uso.items()]) if tower_uso else "-"))
    st.markdown("---")

    kab_col = temukan_kolom(result_df, ['kabupaten', 'regency', 'kab', 'kab/kota'])
    if kab_col and kab_col in result_df.columns:
        df_calc = result_df.copy()
        df_calc['is_4g'] = df_calc[cat_col].astype(str).str.upper().str.contains('4G', na=False) if cat_col else False
        df_calc['is_uso'] = df_calc[cat_col].astype(str).str.upper().str.contains('USO', na=False) if cat_col else False
        df_calc['is_up'], df_calc['is_down'] = df_calc['status'].astype(str).str.lower() == 'up', df_calc['status'].astype(str).str.lower() == 'down'

        kab_summary = df_calc.groupby(kab_col).apply(lambda g: pd.Series({
            'Total Site': len(g), 'BTS 4G': len(g[g['is_4g']]), 'BTS USO': len(g[g['is_uso']]),
            '4G_Up': len(g[g['is_4g'] & g['is_up']]), '4G_Down': len(g[g['is_4g'] & g['is_down']]),
            'USO_Up': len(g[g['is_uso'] & g['is_up']]), 'USO_Down': len(g[g['is_uso'] & g['is_down']])
        })).reset_index().rename(columns={kab_col: 'Kabupaten'}).sort_values(by=['4G_Down', 'USO_Down', 'Total Site'], ascending=[False, False, False]).reset_index(drop=True)

        st.markdown("### 🏛️ Breakdown Detail per Kabupaten")
        html_table = """
        <style>
            .c-table th, .c-table td { border: 1px solid #dcdcdc; padding: 8px; text-align: center; vertical-align: middle; }
            .c-table th { background-color: #009999; color: white; }
            .c-table tr:nth-child(even) { background-color: #f8f9fa; }
        </style>
        <table class="c-table" style="width:100%; border-collapse:collapse; margin-bottom:20px;">
            <tr><th rowspan="2">No</th><th rowspan="2">Kabupaten</th><th rowspan="2">Total Site</th><th rowspan="2">BTS 4G</th><th rowspan="2">BTS USO</th><th colspan="2">Status BTS 4G</th><th colspan="2">Status BTS USO</th></tr>
            <tr><th>Site Up</th><th>Site Down</th><th>Site Up</th><th>Site Down</th></tr>
        """
        for i, r in kab_summary.iterrows():
            c4g = f"color: #d32f2f; font-weight:bold;" if r['4G_Down']>0 else ""
            cuso = f"color: #d32f2f; font-weight:bold;" if r['USO_Down']>0 else ""
            html_table += f"<tr><td>{i+1}</td><td><b>{r['Kabupaten']}</b></td><td>{r['Total Site']}</td><td>{r['BTS 4G']}</td><td>{r['BTS USO']}</td><td>{r['4G_Up']}</td><td style='{c4g}'>{r['4G_Down']}</td><td>{r['USO_Up']}</td><td style='{cuso}'>{r['USO_Down']}</td></tr>"
        st.markdown(html_table + "</table>", unsafe_allow_html=True)

    total_period_mb = result_df['total_traffic_mb_period'].sum() if 'total_traffic_mb_period' in result_df.columns else result_df.get('traffic_mb', pd.Series([0])).sum()
    timeline_df = st.session_state.get("timeline_df", pd.DataFrame())
    rata_avail, label_avail = 0.0, "Rata-rata Availability"
    
    if not timeline_df.empty:
        trend_df_front = timeline_df[timeline_df['site_id_clean'].isin(result_df['site_id_clean'].unique())].copy()
        if not trend_df_front.empty:
            trend_df_front = trend_df_front[trend_df_front['parsed_time'] < trend_df_front['parsed_time'].max()]
            trend_df_front['hours_count'] = trend_df_front.groupby(['site_id_clean', trend_df_front['parsed_time'].dt.date])['parsed_time'].transform(lambda x: x.dt.hour.nunique())
            valid_avail_df = trend_df_front[trend_df_front['hours_count'] >= 24]
            rata_avail = valid_avail_df['avail_num'].mean() if not valid_avail_df.empty else trend_df_front['avail_num'].mean()
            label_avail = "Availability (Data 24 Jam)" if not valid_avail_df.empty else "Availability (Total Raw)"

    st.subheader("📊 Ringkasan Performa & Traffic (Periode Terdampak)")
    ac1, ac2, ac3 = st.columns(3)
    ac1.metric(label_avail, f"{rata_avail:.2f} %")
    ac2.metric("Total Traffic", f"{(total_period_mb/1024.0):,.2f} GB")
    ac3.metric("Total Active User", f"{int(result_df.get('max_user', pd.Series([0])).sum()):,}")
    st.divider()

    t1, t2, t3 = st.tabs(["🗺 Visual Peta Sebaran", "📈 Tren Layanan & Traffic", "📋 Data Lengkap"])
    
    fig_map = fig_avail = fig_traffic = fig_user = None
    with t1:
        map_display_df = df_master[df_master['jarak_km'] <= (st.session_state["active_radius"] + 20)].copy() if (st.session_state.get("filter_mode") != "Batas Administrasi Murni (Eksak)" and st.session_state.get("active_radius") and st.session_state.get("center_coords") and 'jarak_km' in df_master.columns) else result_df.copy()
        
        # Pemetakan ulang status aman dari reset
        if stored_api_dict:
            map_display_df['status'] = map_display_df['site_id_clean'].map(stored_api_dict).fillna('Unmonitor')
        else:
            map_display_df['status'] = 'Unmonitor'
            
        lat_col = temukan_kolom(map_display_df, ['latitude', 'lat', 'y'])
        lon_col = temukan_kolom(map_display_df, ['longitude', 'long', 'lon', 'x'])
        
        if lat_col and lon_col:
            valid_map_df = map_display_df.dropna(subset=[lat_col, lon_col]).copy()
            if not valid_map_df.empty:
                valid_map_df['Status_Site'] = valid_map_df['status'].astype(str).str.capitalize()
                
                # Mengubah label status untuk site di luar radius (+20km) agar tidak dilabeli "Unmonitor"
                if st.session_state.get("active_radius") and 'jarak_km' in valid_map_df.columns and st.session_state.get("filter_mode") != "Batas Administrasi Murni (Eksak)":
                    is_outside = valid_map_df['jarak_km'] > st.session_state["active_radius"]
                    valid_map_df.loc[is_outside, 'Status_Site'] = 'Lainnya' 
                    
                    cat_col_map = temukan_kolom(valid_map_df, ['program', 'kategori', 'category', 'jenis_bts', 'bts_type', 'tipe_bts', 'tipe site'])
                    if cat_col_map:
                        is_4g_out = is_outside & valid_map_df[cat_col_map].astype(str).str.upper().str.contains('4G', na=False)
                        is_uso_out = is_outside & valid_map_df[cat_col_map].astype(str).str.upper().str.contains('USO', na=False)
                        valid_map_df.loc[is_4g_out, 'Status_Site'] = 'BTS 4G'
                        valid_map_df.loc[is_uso_out, 'Status_Site'] = 'BTS USO'

                color_map = {
                    'Up': '#00BFFF', 
                    'Down': '#FF5252', 
                    'Unmonitor': '#FFC107',
                    'BTS 4G': '#BDBDBD',
                    'BTS USO': '#9E9E9E',
                    'Lainnya': '#E0E0E0'
                }
                
                fig_map = px.scatter_map(valid_map_df, lat=lat_col, lon=lon_col, color='Status_Site', color_discrete_map=color_map, hover_name='site_id_clean', hover_data={'status': True, 'traffic_mb': ':.2f', 'max_user': True, lat_col: False, lon_col: False}, zoom=hitung_auto_zoom(valid_map_df[lat_col], valid_map_df[lon_col], st.session_state.get("active_radius")), center=dict(lat=float(valid_map_df[lat_col].mean()), lon=float(valid_map_df[lon_col].mean())), height=600)
                map_layers = [{"below": 'traces', "sourcetype": "raster", "source": ["https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}"]}]
                
                if os.path.exists('indonesia.geojson'):
                    try:
                        with open('indonesia.geojson', 'r', encoding='utf-8') as f: map_layers.append({"sourcetype": "geojson", "source": json.load(f), "type": "line", "color": "rgba(255, 255, 255, 0.7)", "line": {"width": 1.5}, "below": "traces"})
                    except Exception: pass
                
                if st.session_state["center_coords"] and st.session_state.get("active_radius"):
                    ring_geojson = buat_polygon_lingkaran(st.session_state["center_coords"][0], st.session_state["center_coords"][1], st.session_state["active_radius"])
                    map_layers.append({"sourcetype": "geojson", "source": ring_geojson, "type": "fill", "color": "rgba(255, 0, 255, 0.15)", "below": "traces"})
                    map_layers.append({"sourcetype": "geojson", "source": ring_geojson, "type": "line", "color": "#FF00FF", "line": {"width": 2}, "below": "traces"})
                    fig_map.add_scattermap(lat=[st.session_state["center_coords"][0]], lon=[st.session_state["center_coords"][1]], mode='markers', marker=dict(size=25, color='#FF00FF', opacity=0.9), name="Pusat Area")
                
                fig_map.update_layout(map_style="white-bg", map_layers=map_layers, margin={"r":0,"t":0,"l":0,"b":0}, legend=dict(yanchor="bottom", y=0.03, xanchor="left", x=0.03, bgcolor="rgba(255, 255, 255, 0.8)"))
                st.plotly_chart(fig_map, use_container_width=True)

    with t2:
        if not timeline_df.empty:
            trend_df = timeline_df[timeline_df['site_id_clean'].isin(result_df['site_id_clean'].unique())].copy()
            if not trend_df.empty:
                unique_times = sorted(trend_df['parsed_time'].unique())
                if len(unique_times) >= 2: trend_df = trend_df[trend_df['parsed_time'] <= unique_times[-2]]
                if not trend_df.empty:
                    tg = trend_df.groupby('parsed_time').agg(a=('avail_num', 'mean'), t=('traffic_num', 'sum'), u=('user_num', 'sum')).reset_index().sort_values('parsed_time')
                    tg['t'] = tg['t'] / 1024.0
                    
                    fig_avail = px.line(tg, x='parsed_time', y='a', title='📈 Tren Rata-rata Availability (%)', markers=True)
                    fig_avail.update_traces(line_color='#2E7D32', marker=dict(size=6))
                    fig_avail.update_layout(plot_bgcolor='white', margin=dict(l=40, r=40, t=50, b=80), xaxis=dict(tickangle=-45, tickformat='%d %b %Y\n%H:%M', gridcolor='lightgray', title='Waktu'), yaxis=dict(gridcolor='lightgray', range=[0, 100], title='Availability (%)'))

                    fig_traffic = px.line(tg, x='parsed_time', y='t', title='📊 Tren Total Traffic (GB)', markers=True)
                    fig_traffic.update_traces(line_color='#0277BD', marker=dict(size=6))
                    fig_traffic.update_layout(plot_bgcolor='white', margin=dict(l=40, r=40, t=50, b=80), xaxis=dict(tickangle=-45, tickformat='%d %b %Y\n%H:%M', gridcolor='lightgray', title='Waktu'), yaxis=dict(gridcolor='lightgray', rangemode='tozero', title='Traffic (GB)'))

                    fig_user = px.line(tg, x='parsed_time', y='u', title='👥 Tren Total Active Users', markers=True)
                    fig_user.update_traces(line_color='#F9A825', marker=dict(size=6))
                    fig_user.update_layout(plot_bgcolor='white', margin=dict(l=40, r=40, t=50, b=80), xaxis=dict(tickangle=-45, tickformat='%d %b %Y\n%H:%M', gridcolor='lightgray', title='Waktu'), yaxis=dict(gridcolor='lightgray', rangemode='tozero', title='Users'))
                    
                    cc1, cc2 = st.columns(2)
                    with cc1: st.plotly_chart(fig_avail, use_container_width=True); st.plotly_chart(fig_user, use_container_width=True)
                    with cc2: st.plotly_chart(fig_traffic, use_container_width=True)

    with t3: st.dataframe(result_df, use_container_width=True)

    st.markdown("---")
    cg1, cg2 = st.columns(2)
    with cg1:
        if st.button("🚀 Generate PPT Flash Report", type="primary", use_container_width=True):
            with st.spinner("Memproses PPTX beserta Auto-Capture..."):
                st.session_state["last_map_df"] = valid_map_df if 'valid_map_df' in locals() else result_df
                capture_df = st.session_state.get("last_map_df", result_df).copy()
                
                fig_map_capture = None
                if not capture_df.empty and lat_col and lon_col:
                    min_lat_cap, max_lat_cap = capture_df[lat_col].min(), capture_df[lat_col].max()
                    min_lon_cap, max_lon_cap = capture_df[lon_col].min(), capture_df[lon_col].max()
                    
                    fig_map_capture = px.scatter_map(capture_df, lat=lat_col, lon=lon_col, color='Status_Site', color_discrete_map=color_map, height=600)
                    map_layers_capture = [{"below": 'traces', "sourcetype": "raster", "source": ["https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}"]}]
                    
                    if st.session_state["center_coords"] and st.session_state.get("active_radius"):
                        ring_geo_cap = buat_polygon_lingkaran(st.session_state["center_coords"][0], st.session_state["center_coords"][1], st.session_state["active_radius"])
                        map_layers_capture.append({"sourcetype": "geojson", "source": ring_geo_cap, "type": "fill", "color": "rgba(255, 0, 255, 0.15)", "below": "traces"})
                        map_layers_capture.append({"sourcetype": "geojson", "source": ring_geo_cap, "type": "line", "color": "#FF00FF", "line": {"width": 2}, "below": "traces"})
                        fig_map_capture.add_scattermap(lat=[st.session_state["center_coords"][0]], lon=[st.session_state["center_coords"][1]], mode='markers', marker=dict(size=25, color='#FF00FF', opacity=0.9))
                        
                    fig_map_capture.update_layout(
                        map=dict(bounds=dict(west=min_lon_cap - 0.05, east=max_lon_cap + 0.05, south=min_lat_cap - 0.05, north=max_lat_cap + 0.05)),
                        map_style="white-bg", map_layers=map_layers_capture, margin={"r":0,"t":0,"l":0,"b":0}, showlegend=False
                    )

                st.session_state["generated_pptx_bytes"] = generate_standard_pptx_report(j_kejadian, d_kejadian, w_kejadian, st.session_state["loc_label"], st.session_state["mode_label"], result_df, fig_map=fig_map_capture, fig_avail=fig_avail, fig_traffic=fig_traffic, fig_user=fig_user)
            st.success("File PPTX Berhasil Dibuat!")
    with cg2:
        if st.session_state.get("generated_pptx_bytes"):
            safe_lokasi = str(st.session_state.get("loc_label", "Area")).replace('/', '_').replace(':', '_').replace(' ', '_')
            st.download_button(label="📥 Download File PPTX", data=st.session_state["generated_pptx_bytes"], file_name=f"Flash_Report_{safe_lokasi}.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", use_container_width=True)
        else: st.info("Klik tombol **🚀 Generate PPT Flash Report** di kiri terlebih dahulu.")