import io
import os
import json
import re
import math
import requests
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

# ==========================================
# 1. FUNGSI UTILITAS & PERHITUNGAN
# ==========================================
def hitung_jarak_haversine_vec(lat1, lon1, lat2_series, lon2_series):
    R = 6371.0
    lat1_rad, lon1_rad = np.radians(lat1), np.radians(lon1)
    
    lat2_clean = lat2_series.fillna(0)
    lon2_clean = lon2_series.fillna(0)
    
    lat2_rad, lon2_rad = np.radians(lat2_clean.to_numpy()), np.radians(lon2_clean.to_numpy())

    dlat = lat2_rad - lat1_rad
    dlon = lon2_rad - lon1_rad

    a = np.sin(dlat / 2.0) ** 2 + np.cos(lat1_rad) * np.cos(lat2_rad) * np.sin(dlon / 2.0) ** 2
    c = 2 * np.arctan2(np.sqrt(a), np.sqrt(1 - a))
    return R * c

def buat_polygon_lingkaran(center_lat, center_lon, radius_km):
    R = 6371.0
    points = []
    for angle in range(0, 360, 5):
        theta = math.radians(angle)
        lat_rad = math.radians(center_lat)
        lon_rad = math.radians(center_lon)
        
        dist = radius_km / R
        new_lat = math.asin(math.sin(lat_rad) * math.cos(dist) + math.cos(lat_rad) * math.sin(dist) * math.cos(theta))
        new_lon = lon_rad + math.atan2(math.sin(theta) * math.sin(dist) * math.cos(lat_rad), math.cos(dist) - math.sin(lat_rad) * math.sin(new_lat))
        
        points.append([math.degrees(new_lon), math.degrees(new_lat)])
    points.append(points[0]) 
    
    return {
        "type": "Feature",
        "geometry": {
            "type": "Polygon",
            "coordinates": [points]
        }
    }

def hitung_auto_zoom(lat_series, lon_series, active_radius=None):
    if lat_series.empty or lon_series.empty:
        return 8.0
    lat_min, lat_max = lat_series.min(), lat_series.max()
    lon_min, lon_max = lon_series.min(), lon_series.max()
    
    max_diff = max(abs(lat_max - lat_min), abs(lon_max - lon_min))
    
    if active_radius is not None:
        min_deg = ((active_radius + 20) * 2) / 111.0
        max_diff = max(max_diff, min_deg)
    
    if max_diff == 0:
        max_diff = 0.5
        
    zoom = 8.2 - math.log2(max_diff)
    return max(4.0, min(zoom, 12.5))

def temukan_kolom(df, keywords):
    for col in df.columns:
        col_clean = re.sub(r'\s+', ' ', str(col).strip().lower())
        if col_clean in keywords:
            return col
    for col in df.columns:
        col_clean = str(col).strip().lower()
        if any(kw in col_clean for kw in keywords):
            return col
    return None

def clean_to_numeric(series):
    return pd.to_numeric(
        series.astype(str)
        .str.replace('%', '', regex=False)
        .str.replace(',', '.', regex=False)
        .str.strip(),
        errors='coerce'
    ).fillna(0)

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
    cat_col = temukan_kolom(df, ['program', 'kategori', 'category', 'jenis_bts', 'bts_type', 'tipe_bts', 'sumber_dana', 'tipe site', 'tipe_site', 'tipe'])
    trans_col = temukan_kolom(df, ['transmisi', 'transmission', 'transport', 'backhaul', 'tipe_transmisi', 'sistem_transmisi'])

    uso_sites = pd.DataFrame()
    g4_sites = pd.DataFrame()

    if cat_col and cat_col in df.columns:
        cat_str = df[cat_col].astype(str).str.upper()
        uso_sites = df[cat_str.str.contains('USO', na=False)]
        g4_sites = df[cat_str.str.contains('4G', na=False)]
    else:
        uso_mask = pd.Series(False, index=df.index)
        g4_mask = pd.Series(False, index=df.index)
        for c in df.select_dtypes(include=['object']):
            s = df[c].astype(str).str.upper()
            uso_mask |= s.str.contains('USO', na=False)
            g4_mask |= s.str.contains('4G', na=False)
        uso_sites = df[uso_mask]
        g4_sites = df[g4_mask]

    uso_vsat = uso_mw = g4_vsat = g4_mw = 0
    if trans_col and trans_col in df.columns:
        if not uso_sites.empty:
            uso_tr = uso_sites[trans_col].astype(str).str.upper()
            uso_vsat = len(uso_sites[uso_tr.str.contains('VSAT', na=False)])
            uso_mw = len(uso_sites[uso_tr.str.contains('MW|MICROWAVE', regex=True, na=False)])
        if not g4_sites.empty:
            g4_tr = g4_sites[trans_col].astype(str).str.upper()
            g4_vsat = len(g4_sites[g4_tr.str.contains('VSAT', na=False)])
            g4_mw = len(g4_sites[g4_tr.str.contains('MW|MICROWAVE', regex=True, na=False)])

    return {
        'n_uso': len(uso_sites), 'n_4g': len(g4_sites),
        'uso_vsat': uso_vsat, 'uso_mw': uso_mw,
        'g4_vsat': g4_vsat, 'g4_mw': g4_mw
    }

def hitung_breakdown_power_tower(df):
    power_summary, tower_summary = {}, {}
    power_col = df.columns[17] if len(df.columns) > 17 else temukan_kolom(df, ['tipe power', 'power', 'sumber_daya', 'catergory_power', 'tipe_power', 'catu_daya'])
    tower_col = df.columns[18] if len(df.columns) > 18 else temukan_kolom(df, ['tipe tower', 'tower', 'tipe_tower', 'tower_type', 'jenis_tower', 'structure'])
    if power_col and power_col in df.columns:
        power_summary = df[power_col].astype(str).str.strip().str.title().value_counts().to_dict()
    if tower_col and tower_col in df.columns:
        tower_summary = df[tower_col].astype(str).str.strip().str.upper().value_counts().to_dict()
    return power_summary, tower_summary

def replace_placeholder_shape_with_image(slide, img_bytes, target_keys, force_width=None, force_height=None):
    target_keys_upper = [k.upper() for k in target_keys]
    for shape in list(slide.shapes):
        shape_name = shape.name.strip().upper()
        if any(k in shape_name for k in target_keys_upper):
            left, top = shape.left, shape.top
            width = force_width if force_width else shape.width
            height = force_height if force_height else shape.height
            slide.shapes.add_picture(img_bytes, left, top, width=width, height=height)
            sp = shape._element
            sp.getparent().remove(sp)
            return True
    return False

def insert_kabupaten_table_to_pptx(slide, kab_summary, left=None, top=None, width=None):
    rows = len(kab_summary) + 2
    cols = 9
    
    pos_left = left if (left is not None and left < Inches(5.0)) else Inches(0.8)
    pos_top = top if top is not None else Inches(1.5)
    
    if width is not None and width >= Inches(6.5):
        pos_width = width
    else:
        pos_width = Inches(7.5)

    pos_height = Inches(0.28 * rows)

    table_shape = slide.shapes.add_table(rows, cols, pos_left, pos_top, pos_width, pos_height)
    table = table_shape.table

    col_weights = [0.06, 0.30, 0.09, 0.09, 0.09, 0.09, 0.09, 0.09, 0.09]
    for i, w_ratio in enumerate(col_weights):
        table.columns[i].width = int(pos_width * w_ratio)

    header_bg_color = RGBColor(0, 153, 153)  # #009999
    header_text_color = RGBColor(255, 255, 255)

    def format_cell(cell, text, bold=False, size_pt=8, color_rgb=None, bg_rgb=None):
        cell.text = text
        cell.margin_left = Inches(0.04)
        cell.margin_right = Inches(0.04)
        cell.margin_top = Inches(0.02)
        cell.margin_bottom = Inches(0.02)
        
        if bg_rgb:
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_rgb
            
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.font.bold = bold
        p.font.size = Pt(size_pt)
        if color_rgb:
            p.font.color.rgb = color_rgb

    # Header Baris 1
    h1 = ["No", "Kabupaten", "Total Site", "BTS 4G", "BTS USO", "Status BTS 4G", "", "Status BTS USO", ""]
    for c_idx, text in enumerate(h1):
        format_cell(table.cell(0, c_idx), text, bold=True, size_pt=9, color_rgb=header_text_color, bg_rgb=header_bg_color)

    table.cell(0, 5).merge(table.cell(0, 6))
    table.cell(0, 7).merge(table.cell(0, 8))

    # Header Baris 2
    h2 = ["", "", "", "", "", "Site Up", "Site Down", "Site Up", "Site Down"]
    for c_idx, text in enumerate(h2):
        format_cell(table.cell(1, c_idx), text, bold=True, size_pt=8, color_rgb=header_text_color, bg_rgb=header_bg_color)

    for c_idx in range(5):
        table.cell(0, c_idx).merge(table.cell(1, c_idx))

    # Isi Baris Data
    for r_idx, row in kab_summary.iterrows():
        row_pos = r_idx + 2
        vals = [
            str(r_idx + 1),
            str(row['Kabupaten']),
            str(int(row['Total Site'])),
            str(int(row['BTS 4G'])),
            str(int(row['BTS USO'])),
            str(int(row['4G_Up'])),
            str(int(row['4G_Down'])),
            str(int(row['USO_Up'])),
            str(int(row['USO_Down']))
        ]
        
        row_bg = RGBColor(248, 250, 251) if (r_idx % 2 == 1) else RGBColor(255, 255, 255)
        
        for c_idx, val in enumerate(vals):
            c_color = RGBColor(211, 47, 47) if (c_idx in [6, 8] and int(val) > 0) else None
            is_bold = True if (c_idx in [1, 6, 8] and (c_idx == 1 or int(val) > 0)) else False
            format_cell(table.cell(row_pos, c_idx), val, bold=is_bold, size_pt=8, color_rgb=c_color, bg_rgb=row_bg)

def generate_standard_pptx_report(jenis_kejadian, detail_kejadian, waktu_kejadian, loc_label, mode_label, df_result, fig_map=None, fig_avail=None, fig_traffic=None, fig_user=None):
    template_path = "template_baku.pptx"
    if not os.path.exists(template_path) and os.path.exists("template_baku.pptx.pptx"):
        template_path = "template_baku.pptx.pptx"
        
    if os.path.exists(template_path):
        prs = Presentation(template_path)
    else:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        for _ in range(4):
            prs.slides.add_slide(prs.slide_layouts[6])

    status_counts = df_result['status'].astype(str).str.lower().value_counts()
    bts_down = status_counts.get('down', 0)
    bts_up = status_counts.get('up', 0)
    unmonitor = len(df_result) - bts_down - bts_up

    total_traffic_mb = df_result['traffic_mb'].sum()
    total_traffic_gb = total_traffic_mb / 1024.0
    total_period_mb = df_result['total_traffic_mb_period'].sum() if 'total_traffic_mb_period' in df_result.columns else total_traffic_mb
    total_period_gb = total_period_mb / 1024.0
    total_active_users = int(df_result['max_user'].sum())

    wilayah_info = hitung_ringkasan_wilayah(df_result)
    site_info = hitung_breakdown_site_transmisi(df_result)
    power_summary, tower_summary = hitung_breakdown_power_tower(df_result)
    
    df_up = df_result[df_result['status'].str.lower() == 'up']
    df_down = df_result[df_result['status'].str.lower() == 'down']
    up_stats = hitung_breakdown_site_transmisi(df_up)
    down_stats = hitung_breakdown_site_transmisi(df_down)
    unmon_stats = hitung_breakdown_site_transmisi(df_result[~df_result['status'].str.lower().isin(['up', 'down'])])
    
    cat_col = temukan_kolom(df_result, ['program', 'kategori', 'category', 'jenis_bts', 'bts_type', 'tipe_bts', 'tipe site'])
    df_uso = df_result[df_result[cat_col].astype(str).str.upper().str.contains('USO', na=False)] if cat_col else pd.DataFrame()
    df_4g = df_result[df_result[cat_col].astype(str).str.upper().str.contains('4G', na=False)] if cat_col else pd.DataFrame()
    p_4g, t_4g = hitung_breakdown_power_tower(df_4g)
    p_uso, t_uso = hitung_breakdown_power_tower(df_uso)

    total_monitored = bts_up + bts_down
    avg_avail_val = (bts_up / total_monitored * 100.0) if total_monitored > 0 else 0.0

    kab_col_summary = temukan_kolom(df_result, ['kabupaten', 'regency', 'kab', 'kab/kota'])
    top_kab_name = "-"
    top_kab_down_total = 0
    top_kab_4g_down = 0
    top_kab_uso_down = 0

    if kab_col_summary and kab_col_summary in df_result.columns and bts_down > 0:
        if not df_down.empty:
            cat_col_detect = temukan_kolom(df_down, ['program', 'kategori', 'category', 'jenis_bts', 'bts_type'])
            top_kab_series = df_down[kab_col_summary].value_counts()
            if not top_kab_series.empty:
                top_kab_name = str(top_kab_series.index[0])
                top_kab_down_total = int(top_kab_series.iloc[0])
                df_top_kab = df_down[df_down[kab_col_summary] == top_kab_name]
                if cat_col_detect and cat_col_detect in df_top_kab.columns:
                    cat_str = df_top_kab[cat_col_detect].astype(str).str.upper()
                    top_kab_4g_down = len(df_top_kab[cat_str.str.contains('4G', na=False)])
                    top_kab_uso_down = len(df_top_kab[cat_str.str.contains('USO', na=False)])

    # --- AMBIL TIMESTAMP RAW DATA TERAKHIR ---
    timeline_df = st.session_state.get("timeline_df", pd.DataFrame())
    str_timestamp_raw = "-"
    str_tgl_raw = "-"
    str_jam_raw = "-"

    if not timeline_df.empty:
        unique_times = sorted(timeline_df['parsed_time'].dropna().unique())
        if len(unique_times) >= 2:
            latest_raw_time = unique_times[-2]
        elif len(unique_times) == 1:
            latest_raw_time = unique_times[-1]
        else:
            latest_raw_time = None

        if latest_raw_time is not None:
            str_timestamp_raw = latest_raw_time.strftime("%d/%m/%Y %H:%M WIB")
            str_tgl_raw = latest_raw_time.strftime("%d/%m/%Y")
            str_jam_raw = latest_raw_time.strftime("%H:%M WIB")

    replacement_dict = {
        "{{jenis_kejadian}}": str(jenis_kejadian),
        "{{detail_kejadian}}": str(detail_kejadian),
        "{{waktu_kejadian}}": str(waktu_kejadian),
        "{{lokasi}}": str(loc_label),
        "{{mode}}": str(mode_label),
        "{{timestamp_data}}": str_timestamp_raw,
        "{{timestamp_raw}}": str_timestamp_raw,
        "{{tgl_raw}}": str_tgl_raw,
        "{{jam_raw}}": str_jam_raw,
        "{{total_site}}": str(len(df_result)),
        "{{site_up}}": str(bts_up),
        "{{site_down}}": str(bts_down),
        "{{unmonitor}}": str(unmonitor),
        "{{total_traffic_gb}}": f"{total_period_gb:,.2f}",
        "{{daily_traffic_mb}}": f"{total_traffic_mb:,.2f}",
        "{{daily_traffic_gb}}": f"{total_traffic_gb:,.2f}",
        "{{active_users}}": f"{total_active_users:,}",
        "{{avg_avail}}": f"{avg_avail_val:.2f}%",
        "{{n_prov}}": str(wilayah_info['n_prov']),
        "{{n_kab}}": str(wilayah_info['n_kab']),
        "{{n_kec}}": str(wilayah_info['n_kec']),
        "{{n_desa}}": str(wilayah_info['n_desa']),
        "{{n_4g}}": str(site_info['n_4g']),
        "{{g4_vsat}}": str(site_info['g4_vsat']),
        "{{g4_mw}}": str(site_info['g4_mw']),
        "{{n_uso}}": str(site_info['n_uso']),
        "{{uso_vsat}}": str(site_info['uso_vsat']),
        "{{uso_mw}}": str(site_info['uso_mw']),
        "{{4g_up}}": str(up_stats['n_4g']),
        "{{uso_up}}": str(up_stats['n_uso']),
        "{{4g_up_vsat}}": str(up_stats['g4_vsat']),
        "{{4g_up_mw}}": str(up_stats['g4_mw']),
        "{{uso_up_vsat}}": str(up_stats['uso_vsat']),
        "{{uso_up_mw}}": str(up_stats['uso_mw']),
        "{{4g_down_vsat}}": str(down_stats['g4_vsat']),
        "{{4g_down_mw}}": str(down_stats['g4_mw']),
        "{{uso_down_vsat}}": str(down_stats['uso_vsat']),
        "{{uso_down_mw}}": str(down_stats['uso_mw']),
        "{{top_kab_down}}": str(top_kab_name),
        "{{top_kab_down_count}}": str(top_kab_down_total),
        "{{top_kab_down_4g}}": str(top_kab_4g_down),
        "{{top_kab_down_uso}}": str(top_kab_uso_down),
        "{{tipe_power}}": ", ".join([f"{k}: {v}" for k, v in power_summary.items()]) if power_summary else "-",
        "{{tipe_tower}}": ", ".join([f"{k}: {v}" for k, v in tower_summary.items()]) if tower_summary else "-",
        "{{tipe_power_4g}}": ", ".join([f"{k}: {v}" for k, v in p_4g.items()]) if p_4g else "-",
        "{{tipe_power_uso}}": ", ".join([f"{k}: {v}" for k, v in p_uso.items()]) if p_uso else "-",
        "{{tipe_tower_4g}}": ", ".join([f"{k}: {v}" for k, v in t_4g.items()]) if t_4g else "-",
        "{{tipe_tower_uso}}": ", ".join([f"{k}: {v}" for k, v in t_uso.items()]) if t_uso else "-",
        "{{4g_down}}": str(down_stats['n_4g']),
        "{{uso_down}}": str(down_stats['n_uso']),
        "{{4g_unmon_vsat}}": str(unmon_stats['g4_vsat']),
        "{{4g_unmon_mw}}": str(unmon_stats['g4_mw']),
        "{{uso_unmon_vsat}}": str(unmon_stats['uso_vsat']),
        "{{uso_unmon_mw}}": str(unmon_stats['uso_mw'])
    }

    def proses_shape(shape):
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                p_text = paragraph.text.replace('\u200b', '').replace('\u200d', '')
                if any(key in p_text for key in replacement_dict.keys()):
                    for key, val in replacement_dict.items():
                        p_text = p_text.replace(key, val)
                    if len(paragraph.runs) > 0:
                        for i in range(len(paragraph.runs)):
                            paragraph.runs[i].text = "" 
                        paragraph.runs[0].text = p_text 
                    else:
                        paragraph.text = p_text
        elif shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    proses_shape(cell)
        elif shape.shape_type == 6: 
            for sub_shape in shape.shapes:
                proses_shape(sub_shape)

    for slide in prs.slides:
        for shape in slide.shapes:
            proses_shape(shape)

    try:
        if len(prs.slides) >= 2 and fig_map is not None:
            slide_map = prs.slides[1]
            map_bytes = BytesIO(pio.to_image(fig_map, format="png", width=800, height=676, scale=2))
            if not replace_placeholder_shape_with_image(slide_map, map_bytes, ["PH_MAP", "PETA", "MAP"], force_width=Inches(5.88), force_height=Inches(4.97)):
                slide_map.shapes.add_picture(map_bytes, Inches(0.5), Inches(1.5), width=Inches(5.88), height=Inches(4.97))

        if len(prs.slides) >= 3:
            slide_chart = prs.slides[2]
            if fig_avail is not None:
                avail_bytes = BytesIO(pio.to_image(fig_avail, format="png", width=800, height=450, scale=2))
                replace_placeholder_shape_with_image(slide_chart, avail_bytes, ["PH_AVAIL"])
            if fig_traffic is not None:
                traffic_bytes = BytesIO(pio.to_image(fig_traffic, format="png", width=800, height=450, scale=2))
                replace_placeholder_shape_with_image(slide_chart, traffic_bytes, ["PH_TRAFFIC"])
            if fig_user is not None:
                user_bytes = BytesIO(pio.to_image(fig_user, format="png", width=800, height=450, scale=2))
                replace_placeholder_shape_with_image(slide_chart, user_bytes, ["PH_USER"])
    except Exception as e:
        print(f"Warning: Render grafik ke PPTX dilewati: {e}")

    # ==========================================
    # SLIDE 4: INSERT TABEL KE PPT
    # ==========================================
    kab_col_tbl = temukan_kolom(df_result, ['kabupaten', 'regency', 'kab', 'kab/kota'])
    cat_col_tbl = temukan_kolom(df_result, ['program', 'kategori', 'category', 'jenis_bts', 'bts_type', 'tipe_bts', 'tipe site', 'tipe_site'])
    
    if kab_col_tbl and kab_col_tbl in df_result.columns:
        df_calc_ppt = df_result.copy()
        if cat_col_tbl and cat_col_tbl in df_calc_ppt.columns:
            cat_upper = df_calc_ppt[cat_col_tbl].astype(str).str.upper()
            df_calc_ppt['is_4g'] = cat_upper.str.contains('4G', na=False)
            df_calc_ppt['is_uso'] = cat_upper.str.contains('USO', na=False)
        else:
            df_calc_ppt['is_4g'] = False
            df_calc_ppt['is_uso'] = False

        st_lower = df_calc_ppt['status'].astype(str).str.lower()
        df_calc_ppt['is_up'] = st_lower == 'up'
        df_calc_ppt['is_down'] = st_lower == 'down'

        kab_summary_ppt = df_calc_ppt.groupby(kab_col_tbl).apply(lambda g: pd.Series({
            'Total Site': len(g),
            'BTS 4G': len(g[g['is_4g']]),
            'BTS USO': len(g[g['is_uso']]),
            '4G_Up': len(g[g['is_4g'] & g['is_up']]),
            '4G_Down': len(g[g['is_4g'] & g['is_down']]),
            'USO_Up': len(g[g['is_uso'] & g['is_up']]),
            'USO_Down': len(g[g['is_uso'] & g['is_down']]),
        })).reset_index().rename(columns={kab_col_tbl: 'Kabupaten'})

        kab_summary_ppt = kab_summary_ppt.sort_values(by=['4G_Down', 'USO_Down', 'Total Site'], ascending=[False, False, False]).reset_index(drop=True)

        target_slide = prs.slides[3] if len(prs.slides) >= 4 else prs.slides.add_slide(prs.slide_layouts[6])
        target_shape = None

        for slide in prs.slides:
            for shape in slide.shapes:
                s_name = shape.name.strip().upper()
                s_text = shape.text_frame.text.strip().upper() if shape.has_text_frame else ""
                if "PH_TABEL" in s_name or "PH_TABEL" in s_text or "PH_TABLE" in s_name:
                    target_shape = shape
                    target_slide = slide
                    break
            if target_shape:
                break

        if target_shape:
            pos_left, pos_top, pos_width = target_shape.left, target_shape.top, target_shape.width
            sp = target_shape._element
            sp.getparent().remove(sp)
            insert_kabupaten_table_to_pptx(target_slide, kab_summary_ppt, left=pos_left, top=pos_top, width=pos_width)
        else:
            insert_kabupaten_table_to_pptx(target_slide, kab_summary_ppt)

    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer


# ==========================================
# 2. INISIALISASI SESSION STATE
# ==========================================
if "result_df" not in st.session_state: st.session_state["result_df"] = pd.DataFrame()
if "timeline_df" not in st.session_state: st.session_state["timeline_df"] = pd.DataFrame()
if "center_coords" not in st.session_state: st.session_state["center_coords"] = None
if "loc_label" not in st.session_state: st.session_state["loc_label"] = ""
if "mode_label" not in st.session_state: st.session_state["mode_label"] = ""
if "active_radius" not in st.session_state: st.session_state["active_radius"] = None
if "filter_mode" not in st.session_state: st.session_state["filter_mode"] = "Batas Administrasi Murni (Eksak)"
if "jenis_kejadian" not in st.session_state: st.session_state["jenis_kejadian"] = "Gempa Bumi"
if "detail_kejadian" not in st.session_state: st.session_state["detail_kejadian"] = "6.2 Mag"
if "waktu_kejadian" not in st.session_state: st.session_state["waktu_kejadian"] = "Oktober 2026"
if "generated_pptx_bytes" not in st.session_state: st.session_state["generated_pptx_bytes"] = None


# ==========================================
# 3. STREAMLIT UI & LOGIC
# ==========================================
st.set_page_config(page_title="Flash Report Bencana", layout="wide")
st.title("Dashboard Laporan Dampak Kejadian")
st.caption("Kementerian Komunikasi dan Digital — BAKTI")

st.sidebar.header("1. Upload Data Set")
file_master = st.sidebar.file_uploader("1. Master Site (Koordinat):", type=["xlsx", "xls", "csv"])
file_status = st.sidebar.file_uploader("2. Laporan OSS (Raw Data):", type=["xlsx", "xls", "csv"])

df_master = pd.DataFrame()
if file_master is not None:
    df_master = pd.read_csv(file_master) if file_master.name.endswith('.csv') else pd.read_excel(file_master)
    site_col_master = temukan_kolom(df_master, ['site id', 'site_id', 'id site', 'id_site', 'site', 'siteid'])
    
    if site_col_master:
        df_master['site_id_clean'] = df_master[site_col_master].astype(str).str.strip().str.upper()
        if file_status is not None:
            df_status = pd.read_csv(file_status) if file_status.name.endswith('.csv') else pd.read_excel(file_status)
            site_col_status = temukan_kolom(df_status, ['site id', 'site_id', 'id site', 'id_site', 'site', 'siteid'])
            avail_col = temukan_kolom(df_status, ['availability (%)', 'availability(%)', 'availability', 'avail', 'avail (%)'])
            traffic_col = temukan_kolom(df_status, ['traffic/payload (mb)', 'traffic/payload(mb)', 'payload (mb)', 'traffic (mb)', 'payload', 'traffic', 'traffic_mb'])
            user_col = temukan_kolom(df_status, ['max active user', 'max_active_user', 'active user', 'user', 'max active users', 'rrc', 'max_user'])

            if site_col_status and avail_col and len(df_status.columns) >= 2:
                df_status['site_id_clean'] = df_status[site_col_status].astype(str).str.strip().str.upper()
                for col_dup in ['status', 'traffic_mb', 'total_traffic_mb_period', 'max_user']:
                    if col_dup in df_master.columns: 
                        df_master = df_master.drop(columns=[col_dup])

                df_status['avail_num'] = clean_to_numeric(df_status[avail_col])
                df_status['traffic_num'] = clean_to_numeric(df_status[traffic_col]) if traffic_col else 0.0
                df_status['user_num'] = clean_to_numeric(df_status[user_col]) if user_col else 0.0

                combined_datetime = df_status.iloc[:, 0].astype(str).str.strip() + " " + df_status.iloc[:, 1].astype(str).str.strip()
                df_status['parsed_time'] = pd.to_datetime(combined_datetime, errors='coerce', dayfirst=False)
                
                if df_status['parsed_time'].isna().all():
                    df_status['parsed_time'] = pd.to_datetime(combined_datetime, errors='coerce')
                
                df_status['date_str'] = df_status['parsed_time'].dt.strftime('%Y-%m-%d').fillna(df_status.iloc[:, 0].astype(str))

                st.session_state["timeline_df"] = df_status[['site_id_clean', 'parsed_time', 'avail_num', 'traffic_num', 'user_num']].dropna(subset=['parsed_time']).copy()

                df_last = df_status.sort_values('parsed_time').groupby('site_id_clean').last().reset_index()
                df_last['status'] = df_last['avail_num'].apply(lambda x: 'Down' if x == 0 else 'Up')
                
                df_status['hours_count'] = df_status.groupby(['site_id_clean', 'date_str'])[df_status.columns[1]].transform('nunique')
                df_valid = df_status[df_status['hours_count'] >= 24].copy()
                if df_valid.empty: 
                    df_valid = df_status.copy()

                df_metrics = df_valid.groupby(['site_id_clean', 'date_str']).agg(
                    total_traffic_day=('traffic_num', 'sum'),
                    peak_user_day=('user_num', 'max')
                ).reset_index().groupby('site_id_clean').agg(
                    traffic_mb=('total_traffic_day', 'mean'),
                    total_traffic_mb_period=('total_traffic_day', 'sum'),
                    max_user=('peak_user_day', 'mean')
                ).reset_index()

                df_processed = pd.merge(df_last[['site_id_clean', 'status']], df_metrics, on='site_id_clean', how='left')
                df_master = pd.merge(df_master, df_processed, on='site_id_clean', how='left')
                df_master['status'] = df_master['status'].fillna('Unmonitor')
                for col in ['traffic_mb', 'total_traffic_mb_period', 'max_user']: 
                    df_master[col] = df_master[col].fillna(0.0)
                st.sidebar.success("Berhasil! Data siap.")
        else:
            df_master['status'] = 'Unmonitor'

st.sidebar.markdown("---")
st.sidebar.header("3. Parameter Area Terdampak")

main_category = st.sidebar.radio(
    "Metode Filter Titik Kejadian:",
    ["Pilih Berdasarkan Area / Identitas Site", "Input Koordinat Manual", "Data Live Gempa BMKG"]
)

if not df_master.empty and 'status' in df_master.columns:
    lat_col = temukan_kolom(df_master, ['latitude', 'lat', 'y'])
    lon_col = temukan_kolom(df_master, ['longitude', 'long', 'lon', 'x'])

    if lat_col and lon_col:
        df_master[lat_col] = pd.to_numeric(df_master[lat_col], errors='coerce')
        df_master[lon_col] = pd.to_numeric(df_master[lon_col], errors='coerce')

        if main_category == "Data Live Gempa BMKG":
            st.sidebar.markdown("📡 **Data BMKG Terkini (Riwayat Terbaru)**")
            
            if st.sidebar.button("⬇️ Tarik Riwayat Gempa dari BMKG"):
                try:
                    resp = requests.get("https://data.bmkg.go.id/DataMKG/TEWS/gempaterkini.json", timeout=10)
                    if resp.status_code == 200:
                        data_bmkg_list = resp.json()['Infogempa']['gempa']
                        st.session_state["bmkg_list"] = data_bmkg_list
                        st.sidebar.success("Riwayat data berhasil ditarik!")
                    else:
                        st.sidebar.error("Gagal membaca API BMKG.")
                except Exception as e:
                    st.sidebar.error("Gagal terhubung ke server BMKG.")

            if "bmkg_list" in st.session_state:
                gempa_list = st.session_state["bmkg_list"]
                gempa_options = {f"{g['Tanggal']} {g['Jam']} | {g['Magnitude']} M - {g['Wilayah']}": g for g in gempa_list}
                
                selected_gempa_key = st.sidebar.selectbox("Pilih Kejadian Gempa:", list(gempa_options.keys()))
                data_bmkg = gempa_options[selected_gempa_key]
                
                bmkg_tgl_jam = f"{data_bmkg['Tanggal']} {data_bmkg['Jam']}"
                bmkg_mag = data_bmkg['Magnitude']
                bmkg_wilayah = data_bmkg['Wilayah']
                bmkg_kedalaman = data_bmkg['Kedalaman']
                coords = data_bmkg['Coordinates'].split(",")
                bmkg_lat, bmkg_lon = float(coords[0]), float(coords[1])
                
                st.sidebar.info(
                    f"📍 **{bmkg_wilayah}**\n\n"
                    f"🕒 {bmkg_tgl_jam}\n\n"
                    f"🧲 {bmkg_mag} Mag (Kedalaman {bmkg_kedalaman})\n\n"
                    f"📌 {bmkg_lat}, {bmkg_lon}"
                )
                
                radius_km = st.sidebar.slider("Radius Terdampak (KM):", 1, 100, 10)
                
                if st.sidebar.button("Hitung Area Dampak BMKG"):
                    df_master['jarak_km'] = hitung_jarak_haversine_vec(bmkg_lat, bmkg_lon, df_master[lat_col], df_master[lon_col])
                    
                    st.session_state.update({
                        "center_coords": (bmkg_lat, bmkg_lon), 
                        "result_df": df_master[df_master['jarak_km'] <= radius_km].copy(),
                        "loc_label": bmkg_wilayah, 
                        "mode_label": f"Radius {radius_km} KM", 
                        "active_radius": radius_km, 
                        "filter_mode": "Radius Jarak (KM) dari Pusat Area",
                        "jenis_kejadian": "Gempa Bumi",
                        "detail_kejadian": f"{bmkg_mag} Mag, Kedalaman {bmkg_kedalaman}",
                        "waktu_kejadian": bmkg_tgl_jam,
                        "generated_pptx_bytes": None
                    })
        else:
            in_jenis = st.sidebar.text_input("Kejadian:", st.session_state.get("jenis_kejadian", "Gempa Bumi"))
            in_detail = st.sidebar.text_input("Detail:", st.session_state.get("detail_kejadian", "6.2 Mag"))
            in_waktu = st.sidebar.text_input("Waktu:", st.session_state.get("waktu_kejadian", "Oktober 2026"))

            if main_category == "Pilih Berdasarkan Area / Identitas Site":
                kategori_area = st.sidebar.selectbox("Pilih Kategori Pencarian:", ["Kabupaten", "Kecamatan", "Desa", "Provinsi", "Site ID"])
                key_mapping = {
                    "Provinsi": temukan_kolom(df_master, ['provinsi', 'province', 'prov']),
                    "Kabupaten": temukan_kolom(df_master, ['kabupaten', 'regency', 'kab']),
                    "Kecamatan": temukan_kolom(df_master, ['kecamatan', 'district', 'kec']),
                    "Desa": temukan_kolom(df_master, ['desa', 'kelurahan', 'village']),
                    "Site ID": 'site_id_clean'
                }
                target_col = key_mapping.get(kategori_area)

                if target_col and target_col in df_master.columns:
                    list_options = sorted(df_master[target_col].dropna().astype(str).unique().tolist())
                    selected_val = st.sidebar.selectbox(f"Cari/Pilih {kategori_area}:", options=["-- Pilih --"] + list_options)

                    filter_mode = st.sidebar.radio("Metode Cakupan Area:", ["Batas Administrasi Murni (Eksak)", "Radius Jarak (KM) dari Pusat Area"])
                    radius_km = 10
                    if filter_mode == "Radius Jarak (KM) dari Pusat Area":
                        radius_km = st.sidebar.slider("Radius Terdampak (KM):", 1, 100, 10)

                    if st.sidebar.button("Cari & Hitung Dampak") and selected_val != "-- Pilih --":
                        matched_df = df_master[df_master[target_col].astype(str).str.lower() == selected_val.lower()]
                        if not matched_df.empty:
                            center_lat = matched_df[lat_col].mean()
                            center_lon = matched_df[lon_col].mean()
                            df_master['jarak_km'] = hitung_jarak_haversine_vec(center_lat, center_lon, df_master[lat_col], df_master[lon_col])

                            if filter_mode == "Batas Administrasi Murni (Eksak)":
                                final_result = matched_df.copy()
                                mode_desc = f"Batas Administrasi Murni ({kategori_area}: {selected_val})"
                                rad = None 
                            else:
                                final_result = df_master[df_master['jarak_km'] <= radius_km].copy()
                                mode_desc = f"Radius {radius_km} KM dari Pusat {selected_val}"
                                rad = radius_km

                            st.session_state.update({
                                "center_coords": (center_lat, center_lon),
                                "result_df": final_result,
                                "loc_label": f"{kategori_area} {selected_val}",
                                "mode_label": mode_desc,
                                "active_radius": rad,
                                "filter_mode": filter_mode,
                                "jenis_kejadian": in_jenis, "detail_kejadian": in_detail, "waktu_kejadian": in_waktu,
                                "generated_pptx_bytes": None
                            })
                        else:
                            st.error("Data tidak ditemukan.")

            elif main_category == "Input Koordinat Manual":
                col_l, col_r = st.sidebar.columns(2)
                with col_l: input_lat = st.number_input("Latitude:", format="%.6f", value=-4.512300)
                with col_r: input_lon = st.number_input("Longitude:", format="%.6f", value=140.412300)
                radius_km = st.sidebar.slider("Radius Terdampak (KM):", 1, 100, 10)

                if st.sidebar.button("Hitung Area Terdampak"):
                    df_master['jarak_km'] = hitung_jarak_haversine_vec(input_lat, input_lon, df_master[lat_col], df_master[lon_col])
                    st.session_state.update({
                        "center_coords": (input_lat, input_lon),
                        "result_df": df_master[df_master['jarak_km'] <= radius_km].copy(),
                        "loc_label": f"Koordinat ({input_lat}, {input_lon})",
                        "mode_label": f"Radius {radius_km} KM",
                        "active_radius": radius_km,
                        "filter_mode": "Radius Jarak (KM) dari Pusat Area",
                        "jenis_kejadian": in_jenis, "detail_kejadian": in_detail, "waktu_kejadian": in_waktu,
                        "generated_pptx_bytes": None
                    })

# ==========================================
# 4. TAMPILAN DASHBOARD UTAMA
# ==========================================
result_df = st.session_state["result_df"]
center_coords = st.session_state["center_coords"]
active_radius = st.session_state.get("active_radius", None)
filter_mode_state = st.session_state.get("filter_mode", "Batas Administrasi Murni (Eksak)")

fig_map = fig_avail = fig_traffic = fig_user = None

if not result_df.empty:
    j_kejadian = st.session_state["jenis_kejadian"]
    d_kejadian = st.session_state["detail_kejadian"]
    w_kejadian = st.session_state["waktu_kejadian"]

    st.success(f"📍 **Lokasi Fokus:** {st.session_state['loc_label']} | **Metode:** {st.session_state['mode_label']}\n\n**Info Kejadian:** {j_kejadian} - {d_kejadian} | {w_kejadian}")

    # =========================================================================
    # A. GAMBARAN OVERALL
    # =========================================================================
    st.markdown("### 📊 1. Ringkasan Keseluruhan (Overall Status)")

    df_up = result_df[result_df['status'].str.lower() == 'up']
    df_down = result_df[result_df['status'].str.lower() == 'down']
    df_unmon = result_df[~result_df['status'].str.lower().isin(['up', 'down'])]

    total_stats = hitung_breakdown_site_transmisi(result_df)
    up_stats = hitung_breakdown_site_transmisi(df_up)
    down_stats = hitung_breakdown_site_transmisi(df_down)
    unmon_stats = hitung_breakdown_site_transmisi(df_unmon)

    st.caption("TOTAL KESELURUHAN SITE")
    r1_c1, r1_c2, r1_c3 = st.columns([1, 1.5, 1.5])
    r1_c1.metric("Total Site", len(result_df))
    r1_c2.metric("Total BTS 4G", f"{total_stats['n_4g']} Site", f"VSAT: {total_stats['g4_vsat']} | MW: {total_stats['g4_mw']}")
    r1_c3.metric("Total BTS USO", f"{total_stats['n_uso']} Site", f"VSAT: {total_stats['uso_vsat']} | MW: {total_stats['uso_mw']}")

    st.caption("STATUS SITE UP (NORMAL)")
    r2_c1, r2_c2, r2_c3 = st.columns([1, 1.5, 1.5])
    r2_c1.metric("Site UP (Total)", len(df_up))
    r2_c2.metric("4G UP", f"{up_stats['n_4g']} Site", f"VSAT: {up_stats['g4_vsat']} | MW: {up_stats['g4_mw']}")
    r2_c3.metric("USO UP", f"{up_stats['n_uso']} Site", f"VSAT: {up_stats['uso_vsat']} | MW: {up_stats['uso_mw']}")

    st.caption("STATUS SITE DOWN (KRITIS)")
    r3_c1, r3_c2, r3_c3 = st.columns([1, 1.5, 1.5])
    r3_c1.metric("Site DOWN (Total)", len(df_down), delta="- Down" if len(df_down) > 0 else "Normal", delta_color="inverse")
    r3_c2.metric("4G DOWN", f"{down_stats['n_4g']} Site", f"VSAT: {down_stats['g4_vsat']} | MW: {down_stats['g4_mw']}", delta_color="inverse")
    r3_c3.metric("USO DOWN", f"{down_stats['n_uso']} Site", f"VSAT: {down_stats['uso_vsat']} | MW: {down_stats['uso_mw']}", delta_color="inverse")

    if len(df_unmon) > 0:
        st.caption("STATUS SITE UNMONITOR")
        r4_c1, r4_c2, r4_c3 = st.columns([1, 1.5, 1.5])
        r4_c1.metric("Unmonitor (Total)", len(df_unmon))
        r4_c2.metric("4G Unmonitor", f"{unmon_stats['n_4g']} Site", f"VSAT: {unmon_stats['g4_vsat']} | MW: {unmon_stats['g4_mw']}")
        r4_c3.metric("USO Unmonitor", f"{unmon_stats['n_uso']} Site", f"VSAT: {unmon_stats['uso_vsat']} | MW: {unmon_stats['uso_mw']}")

    st.markdown("---")

    # =========================================================================
    # B. BREAKDOWN TABEL STATUS PER KABUPATEN
    # =========================================================================
    st.markdown("### 🏛️ 2. Breakdown Detail per Kabupaten")
    kab_col = temukan_kolom(result_df, ['kabupaten', 'regency', 'kab', 'kab/kota'])
    cat_col = temukan_kolom(result_df, ['program', 'kategori', 'category', 'jenis_bts', 'bts_type', 'tipe_bts', 'tipe site', 'tipe_site'])

    if kab_col and kab_col in result_df.columns:
        df_calc = result_df.copy()

        if cat_col and cat_col in df_calc.columns:
            cat_upper = df_calc[cat_col].astype(str).str.upper()
            df_calc['is_4g'] = cat_upper.str.contains('4G', na=False)
            df_calc['is_uso'] = cat_upper.str.contains('USO', na=False)
        else:
            df_calc['is_4g'] = False
            df_calc['is_uso'] = False

        st_lower = df_calc['status'].astype(str).str.lower()
        df_calc['is_up'] = st_lower == 'up'
        df_calc['is_down'] = st_lower == 'down'

        def aggregate_kabupaten(g):
            return pd.Series({
                'Total Site': len(g),
                'BTS 4G': len(g[g['is_4g']]),
                'BTS USO': len(g[g['is_uso']]),
                '4G_Up': len(g[g['is_4g'] & g['is_up']]),
                '4G_Down': len(g[g['is_4g'] & g['is_down']]),
                'USO_Up': len(g[g['is_uso'] & g['is_up']]),
                'USO_Down': len(g[g['is_uso'] & g['is_down']]),
            })

        kab_summary = df_calc.groupby(kab_col).apply(aggregate_kabupaten).reset_index()
        kab_summary.rename(columns={kab_col: 'Kabupaten'}, inplace=True)
        kab_summary = kab_summary.sort_values(by=['4G_Down', 'USO_Down', 'Total Site'], ascending=[False, False, False]).reset_index(drop=True)
        kab_summary.insert(0, 'No', range(1, len(kab_summary) + 1))

        columns_multi = pd.MultiIndex.from_tuples([
            ('No', ''),
            ('Kabupaten', ''),
            ('Total Site', ''),
            ('BTS 4G', ''),
            ('BTS USO', ''),
            ('Status BTS 4G', 'Site Up'),
            ('Status BTS 4G', 'Site Down'),
            ('Status BTS USO', 'Site Up'),
            ('Status BTS USO', 'Site Down'),
        ])
        
        kab_table_display = kab_summary.copy()
        kab_table_display.columns = columns_multi

        st.dataframe(kab_table_display, use_container_width=True, hide_index=True)
    else:
        st.info("Kolom Kabupaten tidak ditemukan di master data.")

    st.divider()

    # =========================================================================
    # C. TABS NAVIGASI (PETA, TREN, DATA LENGKAP)
    # =========================================================================
    t1, t2, t3 = st.tabs(["🗺 Peta Sebaran", "📈 Tren Layanan", "📋 Data Lengkap"])
    with t1:
        if filter_mode_state == "Batas Administrasi Murni (Eksak)" or active_radius is None:
            map_display_df = result_df.copy()
        else:
            if center_coords and 'jarak_km' in df_master.columns:
                map_display_df = df_master[df_master['jarak_km'] <= (active_radius + 20)].copy()
            else:
                map_display_df = result_df.copy()

        valid_map_df = map_display_df.dropna(subset=[lat_col, lon_col])
        if not valid_map_df.empty:
            auto_zoom = hitung_auto_zoom(valid_map_df[lat_col], valid_map_df[lon_col], active_radius)
            map_display_df['Status_Site'] = map_display_df['status'].astype(str).str.capitalize()
            
            fig_map = px.scatter_map(
                map_display_df, lat=lat_col, lon=lon_col, color='Status_Site',
                color_discrete_map={'Up': '#00BFFF', 'Down': '#FF5252', 'Unmonitor': '#FFC107'},
                zoom=auto_zoom, center=dict(lat=float(valid_map_df[lat_col].mean()), lon=float(valid_map_df[lon_col].mean())), height=600
            )
            
            map_layers = [{"below": 'traces', "sourcetype": "raster", "source": ["https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}"]}]
            
            geojson_path = 'indonesia.geojson'
            if os.path.exists(geojson_path):
                try:
                    with open(geojson_path, 'r', encoding='utf-8') as f:
                        geojson_kab = json.load(f)
                    map_layers.append({
                        "sourcetype": "geojson", "source": geojson_kab, "type": "line",
                        "color": "rgba(255, 255, 255, 0.7)", "line": {"width": 1.5}, "below": "traces"
                    })
                except Exception:
                    pass

            if center_coords and active_radius is not None:
                ring_geojson = buat_polygon_lingkaran(center_coords[0], center_coords[1], active_radius)
                map_layers.append({
                    "sourcetype": "geojson", "source": ring_geojson, "type": "fill",
                    "color": "rgba(255, 0, 255, 0.15)", "below": "traces"
                })
                map_layers.append({
                    "sourcetype": "geojson", "source": ring_geojson, "type": "line",
                    "color": "#FF00FF", "line": {"width": 2}, "below": "traces"
                })
                fig_map.add_scattermap(
                    lat=[center_coords[0]], lon=[center_coords[1]], 
                    mode='markers', marker=dict(size=25, color='#FF00FF', opacity=0.9), 
                    name="Pusat Area"
                )

            fig_map.update_layout(
                map_style="white-bg", map_layers=map_layers, margin={"r":0,"t":0,"l":0,"b":0},
                legend=dict(yanchor="bottom", y=0.03, xanchor="left", x=0.03, bgcolor="rgba(255, 255, 255, 0.8)", bordercolor="gray", borderwidth=1, title_text=None)
            )
            st.plotly_chart(fig_map, use_container_width=True)

    with t2:
        st.markdown("#### **Tren Layanan Historis pada Area Terdampak**")
        timeline_df = st.session_state.get("timeline_df", pd.DataFrame())
        
        if not timeline_df.empty:
            affected_sites = result_df['site_id_clean'].unique()
            trend_df = timeline_df[timeline_df['site_id_clean'].isin(affected_sites)]
            
            if not trend_df.empty:
                st.markdown("##### 📌 **Ringkasan Agregasi Periode Ini**")
                agg_c1, agg_c2, agg_c3 = st.columns(3)
                
                rata_avail = trend_df['avail_num'].mean()
                total_traffic_all_gb = trend_df['traffic_num'].sum() / 1024.0
                rata_user = trend_df['user_num'].mean()
                
                agg_c1.metric("Rata-rata Availability", f"{rata_avail:.2f} %")
                agg_c2.metric("Total Traffic Keseluruhan", f"{total_traffic_all_gb:,.2f} GB")
                agg_c3.metric("Rata-rata Active User", f"{int(rata_user):,}")
                
                st.divider()

                # TAMPILKAN SAMPAI 2 JAM TERAKHIR
                unique_times = sorted(trend_df['parsed_time'].unique())
                if len(unique_times) >= 2:
                    cutoff_time = unique_times[-2]
                    trend_df = trend_df[trend_df['parsed_time'] <= cutoff_time]

                if not trend_df.empty:
                    tg = trend_df.groupby('parsed_time').agg(a=('avail_num', 'mean'), t=('traffic_num', 'sum'), u=('user_num', 'sum')).reset_index().sort_values('parsed_time')
                    tg['t'] = tg['t'] / 1024.0
                    
                    fig_avail = px.line(tg, x='parsed_time', y='a', title='Tren Availability (%)', markers=True)
                    fig_avail.update_traces(line_color='#2E7D32', marker=dict(size=4)).update_layout(plot_bgcolor='white', margin=dict(l=40, r=40, t=50, b=80), xaxis=dict(tickangle=-45, tickformat='%m/%d/%Y\n%H:%M', gridcolor='lightgray'), yaxis=dict(gridcolor='lightgray'))

                    fig_traffic = px.line(tg, x='parsed_time', y='t', title='Tren Traffic (GB)', markers=True)
                    fig_traffic.update_traces(line_color='#0277BD', marker=dict(size=4)).update_layout(plot_bgcolor='white', margin=dict(l=40, r=40, t=50, b=80), xaxis=dict(tickangle=-45, tickformat='%m/%d/%Y\n%H:%M', gridcolor='lightgray'), yaxis=dict(gridcolor='lightgray'))

                    fig_user = px.line(tg, x='parsed_time', y='u', title='Tren Active Users', markers=True)
                    fig_user.update_traces(line_color='#F9A825', marker=dict(size=4)).update_layout(plot_bgcolor='white', margin=dict(l=40, r=40, t=50, b=80), xaxis=dict(tickangle=-45, tickformat='%m/%d/%Y\n%H:%M', gridcolor='lightgray'), yaxis=dict(gridcolor='lightgray'))
                    
                    c_l, c_r = st.columns(2)
                    with c_l:
                        st.plotly_chart(fig_avail, use_container_width=True)
                        st.plotly_chart(fig_user, use_container_width=True)
                    with c_r:
                        st.plotly_chart(fig_traffic, use_container_width=True)

    with t3:
        st.dataframe(result_df, use_container_width=True)

    st.markdown("---")
    col_gen, col_down = st.columns(2)
    with col_gen:
        if st.button("🚀 Generate PPT Flash Report", type="primary", use_container_width=True):
            with st.spinner("Memproses pembuatan slide & Auto-Capture..."):
                st.session_state["generated_pptx_bytes"] = generate_standard_pptx_report(
                    j_kejadian, d_kejadian, w_kejadian, st.session_state["loc_label"], st.session_state["mode_label"], result_df,
                    fig_map=fig_map, fig_avail=fig_avail, fig_traffic=fig_traffic, fig_user=fig_user
                )
            st.success("File PPTX Berhasil Dibuat!")

    with col_down:
        current_loc = st.session_state.get("loc_label", "Area_Terdampak")
        safe_lokasi = str(current_loc).replace('/', '_').replace(':', '_').replace(' ', '_')
        
        if st.session_state["generated_pptx_bytes"] is not None:
            st.download_button(
                label="📥 Download File PPTX",
                data=st.session_state["generated_pptx_bytes"],
                file_name=f"Flash Performance Report {safe_lokasi}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )
        else:
            st.info("Klik tombol **🚀 Generate PPT Flash Report** di sebelah kiri terlebih dahulu.")