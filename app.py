import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import re
from datetime import datetime
import os
import sys
import plotly.express as px
import plotly.graph_objects as go
from streamlit_calendar import calendar
import os
import sys

# Tambahkan path utils agar bisa diimport
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from utils.styles import get_main_css, get_sidebar_css, get_metric_card_css, get_page_header_css

# ===============================
# KONFIGURASI HALAMAN
# ===============================
st.set_page_config(
    page_title="Pusat Kendali Kehumasan | BPS Sultra",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ===============================
# PREMIUM UI STYLING (v3.0)
# ===============================
st.markdown(get_main_css(), unsafe_allow_html=True)

# ===============================
# CORE LOGIC: DATA LOADING
# ===============================
def render_metric_card(label, value, subtext="", icon="📊"):
    st.markdown(f"""
    <div class="gcard-wrap">
        <div class="gcard-label"><span>{icon}</span> {label}</div>
        <div class="gcard-value">{value}</div>
        <div class="gcard-sub">{subtext}</div>
    </div>
    """, unsafe_allow_html=True)

@st.cache_resource(ttl=3600)
def get_gspread_session():
    """Membangun sesi koneksi ke Google Sheets (Cached Resource)."""
    try:
        s = st.secrets["connections"]["gsheets"]
        sa = s["service_account"]
        pk = sa["private_key"].replace("\\n", "\n")
        
        creds = Credentials.from_service_account_info({
            "type": sa["type"],
            "project_id": sa["project_id"],
            "private_key": pk,
            "client_email": sa["client_email"],
            "token_uri": sa["token_uri"],
        }, scopes=["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"])
        
        client = gspread.authorize(creds)
        url = s["spreadsheet"]
        match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
        if not match: return None
        return client.open_by_key(match.group(1))
    except Exception as e:
        st.error(f"Koneksi GSheets Gagal: {e}")
        return None

@st.cache_data(ttl=600)
def load_single_sheet_data(sheet_name):
    """Memuat data dari satu sheet spesifik."""
    def clean_indo_month_string(date_str):
        if not isinstance(date_str, str): return date_str
        indo_to_eng = {
            'januari': 'January', 'februari': 'February', 'maret': 'March',
            'april': 'April', 'mei': 'May', 'juni': 'June',
            'juli': 'July', 'agustus': 'August', 'september': 'September',
            'oktober': 'October', 'november': 'November', 'desember': 'December'
        }
        
        original_str = str(date_str).strip()
        if not original_str or original_str.lower() == 'nan': return ""
        
        low_str = original_str.lower()
        for indo, eng in indo_to_eng.items():
            if indo in low_str:
                original_str = re.sub(re.escape(indo), eng, original_str, flags=re.IGNORECASE)
                break
        
        # Jika hasil tidak mengandung tahun (4 digit), tambahkan 2026
        if original_str and not re.search(r'\d{4}', original_str):
            original_str = f"{original_str} 2026"
            
        return original_str

    try:
        sh = get_gspread_session()
        if not sh: return pd.DataFrame()
        
        ws = sh.worksheet(sheet_name)
        data = ws.get_all_values()
        if not data: return pd.DataFrame()
        
        # Cari baris header otomatis (lebih cerdas)
        h_idx = 0
        header_keywords = ['no', 'bidang', 'nama', 'konten', 'bulan', 'jadwal', 'tanggal', 'tema', 'kegiatan']
        
        for i, row in enumerate(data):
            row_low = [str(c).strip().lower() for c in row if c]
            if any(any(k in c for k in header_keywords) for c in row_low):
                h_idx = i
                break
        else:
            # Fallback jika tidak ketemu keyword: baris pertama yang ada isinya
            for i, row in enumerate(data):
                if any(c.strip() != "" for c in row):
                    h_idx = i
                    break
        
        headers = [h if h.strip() != "" else f"Col_{i}" for i, h in enumerate(data[h_idx])]
        df = pd.DataFrame(data[h_idx + 1:], columns=headers)
        
        # Fast Whitespace Cleanup
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].str.strip()
        
        # Filter Baris Kosong
        key_patterns = ['Bidang', 'Nama', 'Konten', 'Bulan', 'Jadwal', 'Tema', 'Kegiatan', 'Agenda', 'Judul', 'Petugas', 'Tanggal', 'Materi', 'Rilis', 'Detail']
        available_keys = [c for c in df.columns if any(p.lower() in c.lower() for p in key_patterns)]
        
        if available_keys:
            df = df.replace('', pd.NA).dropna(subset=available_keys, how='all').fillna('')
        
        # Deteksi Tanggal Otomatis (Prioritaskan kolom spesifik)
        date_patterns = ["jadwal", "tanggal", "waktu", "tayang", "post", "rilis", "bulan"]
        date_cols = []
        for p in date_patterns:
            cols = [c for c in df.columns if p in c.lower()]
            date_cols.extend(cols)
        
        # Fallback: Cari kolom yang isinya terlihat seperti tanggal (dd/mm/yyyy)
        if not date_cols:
            for col in df.columns:
                sample = df[col].astype(str).head(20).tolist()
                if any(re.search(r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}', str(s)) for s in sample):
                    date_cols.append(col)

        if date_cols:
            # Cari kolom terbaik (yang punya data paling banyak)
            best_col = date_cols[0]
            max_filled = -1
            for col in date_cols:
                filled_count = df[col].astype(str).str.strip().replace(['', 'nan', 'None'], pd.NA).dropna().count()
                if filled_count > max_filled:
                    max_filled = filled_count
                    best_col = col
            
            def safe_parse_date(val):
                cleaned = clean_indo_month_string(val)
                # Hilangkan simbol-simbol aneh
                cleaned = re.sub(r'["\'\>\[\]]', '', cleaned).strip()
                try:
                    return pd.to_datetime(cleaned, errors='coerce', dayfirst=True)
                except:
                    return pd.NaT

            # Coba parse kolom utama
            df['dt_ref'] = df[best_col].apply(safe_parse_date)
            
            # Jika ada NaT tapi ada kolom tanggal kedua, lakukan pengisian (coalesce)
            if df['dt_ref'].isna().any() and len(date_cols) > 1:
                col2 = [c for c in date_cols if c != best_col][0]
                df['dt_ref'] = df['dt_ref'].fillna(df[col2].apply(safe_parse_date))
            
        return df.dropna(how="all")
    except Exception:
        return pd.DataFrame()

@st.cache_data(ttl=600)
def load_all_dashboard_data():
    """Memuat SEMUA sheet dari workbook secara total untuk menghindari data terlewat."""
    try:
        sh = get_gspread_session()
        if not sh: return pd.DataFrame()
        
        # Ambil seluruh daftar nama sheet tanpa kecuali
        all_worksheets = [ws.title for ws in sh.worksheets()]
        
        all_dfs = []
        loaded_sheets = []
        
        for s in all_worksheets:
            # Skip sheet sistem atau yang jelas kosong (opsional)
            if s.lower() in ['pilih', 'config', 'referensi', 'hidden']: continue
            
            tdf = load_single_sheet_data(s)
            if not tdf.empty and ('dt_ref' in tdf.columns or len(tdf.columns) > 3):
                tdf['Sumber'] = s
                all_dfs.append(tdf)
                loaded_sheets.append(s)
        
        # Simpan daftar sheet yang berhasil dimuat di session state untuk UI
        st.session_state.loaded_sheets_info = loaded_sheets
        
        if not all_dfs: return pd.DataFrame()
        return pd.concat(all_dfs, ignore_index=True)
    except Exception as e:
        st.error(f"Gagal memuat data total: {e}")
        return pd.DataFrame()

# ===============================
# RENDER SIDEBAR
# ===============================
if "active_menu" not in st.session_state:
    st.session_state.active_menu = "Dashboard"

with st.sidebar:
    # Load and display the official BPS Sultra logo
    logo_path = os.path.join("assets", "[Color] Logo BPS 7400.png")
    if os.path.exists(logo_path):
        st.image(logo_path, use_container_width=True)
    else:
        st.markdown("""
        <div class="sb-logo-wrap">
            <span class="sb-logo-icon">📡</span>
            <div class="sb-logo-title">Pusat Kendali Kehumasan</div>
            <div class="sb-logo-sub">BPS Provinsi Sulawesi Tenggara</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown('<div class="sb-divider" style="margin-top: 20px;"></div>', unsafe_allow_html=True)
    
    # Menu items
    menu_items = [
        {"id": "Dashboard", "label": "Dashboard Utama", "icon": "🏠"},
        {"id": "Kalender", "label": "Kalender Agenda", "icon": "📅"},
    ]
    
    for item in menu_items:
        is_active = st.session_state.active_menu == item["id"]
        
        # Clickable menu simulation with session state
        if st.button(f"{item['icon']} {item['label']}", 
                     key=f"btn_{item['id']}", 
                     use_container_width=True,
                     type="primary" if is_active else "secondary"):
            st.session_state.active_menu = item["id"]
            st.rerun()

    st.markdown('<div class="sb-divider"></div>', unsafe_allow_html=True)
    
    if st.button("⟳  Refresh Data", key="refresh_btn", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
        
    st.markdown(f"""
    <div class="sb-footer">
        🕐 Data Teraktif:<br>
        <b>{datetime.now().strftime('%d %b %Y, %H:%M')} WITA</b>
    </div>
    """, unsafe_allow_html=True)

# ===============================
# CONTENT SELECTOR
# ===============================
if st.session_state.active_menu == "Dashboard":
    # Inisialisasi state untuk full view
    if 'fv_source' not in st.session_state: st.session_state.fv_source = False
    if 'fv_tema' not in st.session_state: st.session_state.fv_tema = False
    if 'fv_leaderboard' not in st.session_state: st.session_state.fv_leaderboard = False

    st.markdown(f"""
    <div class="pg-header">
        <span class="header-tag">Executive Summary</span>
        <h1>⏰ SI-PEKAT</h1>
        <p>Ringkasan performa dan agenda kehumasan lintas channel.</p>
        <a href="https://docs.google.com/spreadsheets/d/1vliisXAUpSXAugCj78j8mc8q2kXX3JQ9T49kD5myCL8/edit?hl=id&gid=2087508271#gid=2087508271" target="_blank" class="drive-link">
            <span>📂</span> Buka Source Data (Google Sheets)
        </a>
    </div>
    """, unsafe_allow_html=True)
    
    # Silent Loading (Data Terkoneksi via Cache)
    df_all = load_all_dashboard_data()
    
    if not df_all.empty:
        # Terapkan filter tambahan: Buang baris yang benar-benar kosong 
        df_dash = df_all.copy()
        test_cols = [c for c in df_dash.columns if c not in ['dt_ref', 'Sumber', 'Unnamed: 0']]
        df_dash = df_dash.dropna(subset=test_cols, how='all').fillna('')
        
        # Kalkulasi Waktu
        now = datetime.now()
        cur_month = now.month
        cur_year = now.year
        
        # 1. Total Semua Agenda (Seluruh Tahun)
        total_all = len(df_all)
        
        # 2. Agenda Bulan Ini
        if 'dt_ref' in df_all.columns:
            mask_this_month = (df_all['dt_ref'].dt.month == cur_month) & (df_all['dt_ref'].dt.year == cur_year)
            df_this_month = df_all[mask_this_month]
            
            # 3. Selesai (Bulan Ini)
            status_cols = [c for c in df_dash.columns if any(p in c.lower() for p in ['status', 'keterangan', 'progres', 'done'])]
            done_this_month = 0
            if status_cols:
                # Cek keyword variatif: Selesai, ✅, Done, Sudah, Clear, Tuntas
                masks = [df_this_month[c].astype(str).str.contains("Selesai|✅|Done|Sudah|Clear|Tuntas|Archive", case=False, na=False) for c in status_cols]
                if masks:
                    mask_done = masks[0]
                    for m in masks[1:]:
                        mask_done = mask_done | m
                    done_this_month = len(df_this_month[mask_done])
                
            # 4. Bulan Depan
            next_month = (cur_month % 12) + 1
            next_month_year = cur_year if next_month > 1 else cur_year + 1
            mask_next_month = (df_dash['dt_ref'].dt.month == next_month) & (df_dash['dt_ref'].dt.year == next_month_year)
            total_next_month = len(df_dash[mask_next_month])
        else:
            df_this_month = pd.DataFrame()
            done_this_month = 0
            total_next_month = 0

        # RENDER METRICS Dashboard
        c1, c2, c3, c4 = st.columns(4)
        with c1: render_metric_card("Total Agenda", total_all, "Seluruh Tahun 2026", "📚")
        with c2: render_metric_card("Bulan Ini", len(df_this_month), f"{datetime.now().strftime('%B %Y')}", "📅")
        with c3: render_metric_card("Selesai", done_this_month, "Bulan Berjalan", "✅")
        with c4: render_metric_card("Bulan Depan", total_next_month, "Proyeksi Agenda", "🚀")
        
        st.markdown('<div class="sb-divider"></div>', unsafe_allow_html=True)
        
        # INSIGHTS
        # Render Charts dengan toggle full size
        def render_source_chart(expanded=False):
            with st.container():
                c1, c2 = st.columns([0.96, 0.04])
                with c1: st.markdown('<span class="chart-header">📊 Distribusi Channel</span>', unsafe_allow_html=True)
                with c2: 
                    if st.button("⛶" if not expanded else "⩵", key="btn_fv_source", help="Lihat Full-Size" if not expanded else "Kembali"):
                        st.session_state.fv_source = not st.session_state.fv_source
                        st.rerun()
                
                source_counts = df_all['Sumber'].value_counts().reset_index()
                source_counts.columns = ['Sheet', 'Jumlah']
                
                # Nexus Slate Chart Color Palette
                color_map = {
                    "📊Promosi Statistik 2026": "#f26522", # Primary Orange
                    "✨Hari Penting": "#ffb599",          # Light Orange
                    "📣Press Release": "#8cceff",         # Tertiary Blue
                    "🖼️Konten Medsos": "#009ade",         # Professional Blue
                    "Sosialisasi Publikasi 📢": "#581e03", # Dark Brown/Orange
                    "🕵🏻‍♂️Keprotokolan": "#e4e1ed",          # Neutral
                    "📸 Peliputan": "#f26522",
                    "📝 Media Massa (non rilis)": "#ffb599",
                    "🎙️Sosialisasi Kegiatan": "#8cceff",
                    "🤝🏻 Kelembagaan": "#009ade",
                    "📚 Pengembangan Kompetensi": "#581e03"
                }
                
                fig = px.bar(source_counts, x='Sheet', y='Jumlah', 
                             color='Sheet', 
                             color_discrete_map=color_map,
                             template='plotly_dark',
                             text='Jumlah')
                
                fig.update_traces(textposition='outside')
                fig.update_layout(showlegend=False, 
                                  plot_bgcolor='rgba(0,0,0,0)', 
                                  paper_bgcolor='rgba(0,0,0,0)',
                                  height=550 if expanded else 480,
                                  margin=dict(t=50, b=10, l=10, r=10),
                                  xaxis=dict(title=""),
                                  font=dict(color='#888aaa'))
                st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False}, theme=None)

        def render_trend_chart(expanded=False):
            with st.container():
                c1, c2 = st.columns([0.96, 0.04])
                with c1: st.markdown('<span class="chart-header">📈 Tren Aktivitas Bulanan</span>', unsafe_allow_html=True)
                with c2: 
                    if st.button("⛶" if not expanded else "⩵", key="btn_fv_tema", help="Lihat Full-Size" if not expanded else "Kembali"):
                        st.session_state.fv_tema = not st.session_state.fv_tema
                        st.rerun()
                
                if 'dt_ref' in df_dash.columns:
                    # Persiapkan data Tren Bulanan
                    df_trend = df_dash.copy()
                    df_trend['Bulan_Tahun'] = df_trend['dt_ref'].dt.to_period('M').astype(str)
                    
                    monthly_data = df_trend.groupby('Bulan_Tahun').size().reset_index(name='Jumlah')
                    monthly_data = monthly_data.sort_values('Bulan_Tahun')
                    
                    # Konversi Bulan_Tahun ke nama bulan yang lebih cantik
                    def format_month(period_str):
                        try:
                            dt = datetime.strptime(period_str, '%Y-%m')
                            return dt.strftime('%b %Y')
                        except: return period_str
                    
                    monthly_data['Label_Bulan'] = monthly_data['Bulan_Tahun'].apply(format_month)

                    fig_trend = px.area(monthly_data, x='Label_Bulan', y='Jumlah',
                                        template='plotly_dark',
                                        color_discrete_sequence=['#f26522'])
                    
                    fig_trend.update_traces(
                        line_color='#f26522',
                        line_width=4,
                        fillcolor='rgba(242, 101, 34, 0.2)',
                        mode='lines+markers',
                        marker=dict(size=10, color='#FFFFFF', line=dict(width=2, color='#f26522'))
                    )
                    # Filter data valid untuk trendline (Hanya Tahun 2026)
                    df_trend = df_all.dropna(subset=['dt_ref']).copy()
                    df_trend = df_trend[df_trend['dt_ref'].dt.year == 2026]
                    
                    # Buat index 12 bulan penuh agar grafik tidak terputus
                    all_months = pd.date_range(start='2026-01-01', end='2026-12-01', freq='MS')
                    base_df = pd.DataFrame({'dt_ref_period': all_months.to_period('M')})
                    
                    if not df_trend.empty:
                        df_trend['dt_ref_period'] = df_trend['dt_ref'].dt.to_period('M')
                        counts = df_trend.groupby('dt_ref_period').size().reset_index(name='Jumlah')
                        
                        # Gabungkan dengan base_df agar ada 12 bulan
                        trend_data = pd.merge(base_df, counts, on='dt_ref_period', how='left').fillna(0)
                        trend_data['Bulan_Teks'] = trend_data['dt_ref_period'].dt.strftime('%b %Y')
                        trend_data = trend_data.sort_values('dt_ref_period')
                        
                        fig_trend = px.area(trend_data, x='Bulan_Teks', y='Jumlah',
                                            markers=True, color_discrete_sequence=['#F26522'])
                        
                        # --- Perbaikan Visual Chart (Jan - Des Only) ---
                        fig_trend.update_layout(
                            plot_bgcolor='rgba(0,0,0,0)',
                            paper_bgcolor='rgba(0,0,0,0)',
                            xaxis=dict(showgrid=False, title='', tickfont=dict(color='#888aaa'), type='category'),
                            yaxis=dict(showgrid=True, gridcolor='rgba(255,255,255,0.05)', title='Jumlah Agenda', tickfont=dict(color='#888aaa')),
                            margin=dict(l=0, r=0, t=20, b=0),
                            height=350,
                            hovermode='x unified'
                        )
                        st.plotly_chart(fig_trend, use_container_width=True, config={'displayModeBar': False})
                    else:
                        st.info("Belum ada data di tahun 2026 untuk ditampilkan di tren.")
                else:
                    st.info("Data waktu tidak tersedia untuk menampilkan tren.")

        def render_leaderboard_chart(expanded=False):
            with st.container():
                c1, c2 = st.columns([0.96, 0.04])
                with c1: st.markdown('<span class="chart-header">🏆 Top 10 Produktivitas Pegawai</span>', unsafe_allow_html=True)
                with c2: 
                    if st.button("⛶", key="btn_fv_leader", help="Lihat Full-Size" if not expanded else "Kembali"):
                        st.session_state.fv_leaderboard = not st.session_state.fv_leaderboard
                        st.rerun()
                
                # Ekstrak Nama dari berbagai kolom potensial
                name_cols = ['Nama', 'Petugas', 'Writer', 'Designer', 'Kameramen', 'Cast/VO', 'Operator', 'Kontributor']
                found_names = []
                
                for col in df_dash.columns:
                    if any(p.lower() in col.lower() for p in name_cols):
                        # Ambil data, bersihkan nilai kosong/non-nama
                        valid_names = df_dash[col].astype(str).str.strip()
                        valid_names = valid_names[~valid_names.isin(['', 'nan', '-', 'None', 'NULL', 'nan nan'])]
                        found_names.extend(valid_names.tolist())
                
                if found_names:
                    df_names = pd.DataFrame(found_names, columns=['Pegawai'])
                    top_stats = df_names['Pegawai'].value_counts().head(10).reset_index()
                    top_stats.columns = ['Nama', 'Jumlah']
                    
                    fig_lead = px.bar(top_stats, x='Jumlah', y='Nama', 
                                      orientation='h',
                                      template='plotly_dark',
                                      color='Jumlah',
                                      color_continuous_scale=['#1a1a2e', '#f26522'])
                    
                    fig_lead.update_traces(marker_line_width=0)
                    fig_lead.update_layout(
                        coloraxis_showscale=False,
                        plot_bgcolor='rgba(0,0,0,0)', 
                        paper_bgcolor='rgba(0,0,0,0)',
                        height=550 if expanded else 400,
                        margin=dict(t=20, b=20, l=20, r=20),
                        xaxis=dict(title="Total Agenda", showgrid=True, gridcolor='rgba(255,255,255,0.05)', color='#888aaa'),
                        yaxis=dict(title="", autorange="reversed", color='#888aaa'),
                        font=dict(color='#888aaa', family='Inter')
                    )
                    st.plotly_chart(fig_lead, use_container_width=True, config={'displayModeBar': False}, theme=None)
                else:
                    st.info("Tidak ditemukan data nama pegawai pada sheet ini.")

        # Logic Layout Dinamis (Vertical Stack + Wide Leaderboard)
        if st.session_state.fv_source:
            render_source_chart(expanded=True)
        elif st.session_state.fv_tema:
            render_trend_chart(expanded=True)
        elif st.session_state.fv_leaderboard:
            render_leaderboard_chart(expanded=True)
        else:
            # Baris 1: Distribusi Channel (Single Column)
            render_source_chart(expanded=False)
            st.markdown('<div style="margin: 20px 0;"></div>', unsafe_allow_html=True)
            
            # Baris 2: Tren Aktivitas (Single Column)
            render_trend_chart(expanded=False)
            st.markdown('<div style="margin: 20px 0;"></div>', unsafe_allow_html=True)
            
            # Baris 3: Leaderboard (Marge/Wide)
            render_leaderboard_chart(expanded=False)

        st.markdown('<div class="sb-divider"></div>', unsafe_allow_html=True)
        
        c_i1, c_i2 = st.columns(2)
        with c_i1:
            if 'dt_ref' in df_all.columns:
                day_map = {
                    'Monday': 'Senin', 'Tuesday': 'Selasa', 'Wednesday': 'Rabu',
                    'Thursday': 'Kamis', 'Friday': 'Jumat', 'Saturday': 'Sabtu', 'Sunday': 'Minggu'
                }
                day_name_en = df_all['dt_ref'].dt.day_name().value_counts().idxmax()
                active_day = day_map.get(day_name_en, day_name_en)
                st.markdown(f"""
                <div style="background: #1E1E2F; border: 1px solid rgba(255,255,255,0.05); padding: 20px; border-radius: 18px; text-align: center; box-shadow: 0 10px 25px rgba(0,0,0,0.15);">
                    <div style="font-size: 0.75rem; color: #888aaa; text-transform: uppercase; margin-bottom: 8px;">Hari Teraktif</div>
                    <div style="font-size: 1.8rem; font-weight: 800; color: #F26522;">📅 {active_day}</div>
                </div>
                """, unsafe_allow_html=True)
        with c_i2:
            top_source = df_all['Sumber'].value_counts().idxmax().replace("📣", "").replace("📊", "").replace("🖼️", "").strip()
            st.markdown(f"""
                <div style="background: #1E1E2F; border: 1px solid rgba(255,255,255,0.05); padding: 20px; border-radius: 18px; text-align: center; box-shadow: 0 10px 25px rgba(0,0,0,0.15);">
                    <div style="font-size: 0.75rem; color: #888aaa; text-transform: uppercase; margin-bottom: 8px;">Kontributor Utama</div>
                    <div style="font-size: 1.8rem; font-weight: 800; color: #FFFFFF;">🚀 {top_source}</div>
                </div>
                """, unsafe_allow_html=True)

        st.markdown('<br>', unsafe_allow_html=True)
        st.markdown('<div class="section-label">📅 Agenda Mendatang Terdekat</div>', unsafe_allow_html=True)
        
        if 'dt_ref' in df_all.columns:
            # Filter yang belum lewat
            future_df = df_all[df_all['dt_ref'] >= now].sort_values('dt_ref').head(5)
            
            # Bangun HTML Table kustom dengan deteksi kolom cerdas
            html_table = f'<div class="table-wrap"><table class="modern-table">'
            html_table += '<thead><tr><th>Agenda</th><th>Jadwal</th><th>Bidang/Petugas</th><th>Sumber</th></tr></thead>'
            html_table += '<tbody>'
            
            for _, row in future_df.iterrows():
                # --- ENHANCED OFFICER DETECTOR (SMARTER) ---
                def get_all_officers(r_data):
                    # Kata kunci petugas, tapi hindari kata kunci judul
                    officer_keys = ["petugas", "bidang", "writer", "designer", "editor", "kameramen", "pic", "operator", "cast", "support"]
                    # Tambahkan "nama" hanya jika bukan "nama konten/agenda"
                    found = []
                    officer_cols = []
                    for col in r_data.index:
                        c_low = col.lower()
                        # Lewati jika terlihat seperti kolom Judul
                        if any(x in c_low for x in ["konten", "agenda", "tema", "judul", "kegiatan"]): continue
                        
                        if any(k in c_low for k in officer_keys) or ( "nama" in c_low and "petugas" in c_low ) or c_low == "nama":
                            val = str(r_data[col]).strip()
                            if val and val.lower() not in ['nan', '-', ''] and not val.isdigit():
                                if val not in found: found.append(val)
                                officer_cols.append(col)
                    return ", ".join(found) if found else "-", officer_cols

                petugas_txt, officer_cols = get_all_officers(row)

                # --- REFINED TITLE DETECTOR ---
                def get_best_title(r_data, k_patterns, skip_cols):
                    source_str = str(r_data.get('Sumber',''))
                    
                    # 1. Keyword Prioritas
                    for col in r_data.index:
                        if col in skip_cols: continue
                        if any(p in col.lower() for p in k_patterns):
                            v = r_data[col]
                            if pd.notna(v) and str(v).strip().lower() not in ['nan', '', '-']:
                                return str(v), col
                    
                    # 2. Smart Fallback
                    candidates = []
                    exclude = ['dt_ref', 'Sumber', 'no', 'urut', 'unnamed', 'col_', 's'] + skip_cols
                    for col in r_data.index:
                        if any(x in col.lower() for x in exclude): continue
                        val = str(r_data[col]).strip()
                        if val and val.lower() not in ['nan', '-', ''] and not val.isdigit() and not re.search(r'\d{1,2}[/-]\d{1,2}', val):
                            candidates.append(val)
                    
                    if candidates:
                        if "Publikasi" in source_str:
                             # Jika cuma ada 1 candidate (karena yang lain jadi petugas), tetap kasih prefix
                             main_part = " - ".join(candidates[:2]) if len(candidates) > 1 else candidates[0]
                             return f"Publikasi Bulan {main_part}", None
                        return max(candidates, key=len), None
                    return source_str.split('📢')[0].strip(), None

                # Deteksi Judul
                title_keys = ["nama konten", "tema", "kegiatan", "agenda", "judul", "hari penting", "materi", "topik"]
                agenda_txt, agenda_col = get_best_title(row, title_keys, officer_cols)
                
                # Deteksi Jadwal
                if pd.notna(row['dt_ref']):
                    jadwal_txt = row['dt_ref'].strftime("%d/%m/%Y")
                else:
                    jadwal_txt = str(row.get('Jadwal', row.get('Tanggal', '-')))
                
                source_tag = str(row['Sumber'])
                
                html_table += f"<tr>"
                html_table += f"<td><b>{agenda_txt}</b></td>"
                html_table += f"<td>{jadwal_txt}</td>"
                html_table += f"<td>{petugas_txt}</td>"
                html_table += f"<td><span class='badge-source'>{source_tag}</span></td>"
                html_table += f"</tr>"
            html_table += '</tbody></table></div>'

            st.markdown(html_table, unsafe_allow_html=True)
        
        st.stop()
    else:
        st.warning("Data gabungan belum tersedia.")
        st.stop()

elif st.session_state.active_menu == "Kalender":
    st.markdown("""
    <div class="pg-header">
        <span class="header-tag">Operational Schedule</span>
        <h1>📅 Kalender Agenda</h1>
        <p>Visualisasi jadwal terintegrasi lintas channel publikasi.</p>
        <a href="https://docs.google.com/spreadsheets/d/1vliisXAUpSXAugCj78j8mc8q2kXX3JQ9T49kD5myCL8/edit?hl=id&gid=2087508271#gid=2087508271" target="_blank" class="drive-link">
            <span>📅</span> Buka Source Data (Google Sheets)
        </a>
    </div>
    """, unsafe_allow_html=True)
    
    df_all = load_all_dashboard_data()
    
    if not df_all.empty and 'dt_ref' in df_all.columns:
        df_cal = df_all.dropna(subset=['dt_ref']).copy()
        
        events = []
        source_colors = {
            "📊Promosi Statistik 2026": "#36A2EB",
            "✨Hari Penting": "#FFCE56",
            "📣Press Release": "#F26522",
            "🖼️Konten Medsos": "#4BC0C0",
            "Sosialisasi Publikasi 📢": "#9966FF",
            "🕵🏻‍♂️Keprotokolan": "#C9CBCF",
            "📸 Peliputan": "#FF9F40",
            "📝 Media Massa (non rilis)": "#FF6384",
            "🎙️Sosialisasi Kegiatan": "#36A2EB",
            "🤝🏻 Kelembagaan": "#4BC0C0",
            "📚 Pengembangan Kompetensi": "#9966FF"
        }
        
        for idx, (_, row) in enumerate(df_cal.iterrows()):
            # --- SAME REFINED RADAR FOR CALENDAR ---
            title_keys = ["nama konten", "tema", "kegiatan", "agenda", "judul", "hari penting", "materi", "topik", "keterangan"]
            # Re-implementing simplified inline for speed/safety
            title_val = None
            for p in title_keys:
                match = [c for c in row.index if p in c.lower()]
                if match and str(row[match[0]]).strip().lower() not in ['nan','','-']:
                    title_val = str(row[match[0]])
                    break
            
            if not title_val:
                cands = [str(row[c]).strip() for c in row.index if c not in ['dt_ref','Sumber'] and not str(row[c]).isdigit() and len(str(row[c])) > 2]
                if cands:
                    if "Publikasi" in str(row.get('Sumber','')): 
                        title_val = f"Publikasi Bulan {cands[0]} ({cands[1]})" if len(cands) > 1 else cands[0]
                    else: title_val = max(cands, key=len)
            
            if not title_val or title_val.lower() == 'nan': 
                title_val = str(row.get('Sumber','Agenda')).split('📢')[0].strip()
            
            sumber = row.get('Sumber', 'Lainnya')
            color = source_colors.get(sumber, "#888aaa")
            s_tag = sumber.replace('📣','').replace('📊','').replace('🖼️','').replace('✨','').strip()[:5]
            
            # Formating tanggal (Jadwal)
            raw_sched = str(row.get('Jadwal Posting', row.get('Jadwal', row.get('Tanggal', ''))))
            time_str = ""
            if ":" in raw_sched and any(char.isdigit() for char in raw_sched):
                import re
                t_match = re.search(r'(\d{1,2}:\d{2})', raw_sched)
                if t_match: time_str = f"({t_match.group(1)}) "
            
            display_title = f"{time_str}[{s_tag}] {title_val}"
            sanitized_data = {str(k): str(v) for k, v in row.to_dict().items()}

            events.append({
                "id": str(idx),
                "title": display_title,
                "start": row['dt_ref'].strftime("%Y-%m-%d") if pd.notna(row['dt_ref']) else "",
                "backgroundColor": color,
                "borderColor": color,
                "allDay": True,
                "extendedProps": {
                    "sumber": sumber,
                    "row_data": sanitized_data
                }
            })

        calendar_options = {
            "headerToolbar": {"left": "today prev,next", "center": "title", "right": "dayGridMonth,dayGridWeek,listWeek"},
            "initialView": "dayGridMonth",
            "initialDate": datetime.now().strftime("%Y-%m-%d"),
            "height": 700,
        }
        
        custom_css = """
            .fc { background: rgba(30,33,50,0.4); border-radius: 15px; padding: 20px; border: 1px solid rgba(255,255,255,0.05); }
            .fc-event-title { font-weight: 600 !important; font-size: 0.8rem !important; }
        """
        st.markdown(f"<style>{custom_css}</style>", unsafe_allow_html=True)
        
        @st.dialog("📋 Detail Agenda Kehumasan")
        def show_agenda_detail(title, row_data):
            sumber = row_data.get('Sumber', '')
            if "Medsos" in sumber:
                def get_m(col):
                    v = row_data.get(col)
                    return str(v) if pd.notna(v) and str(v).strip().lower() != 'nan' and str(v).strip() != '' else "-"
                st.markdown(f'<div style="background:rgba(75,192,192,0.1); border-left:5px solid #4bc0c0; padding:15px; border-radius:8px; margin-bottom:20px;"><h4 style="margin:0; color:#FFFFFF;">{get_m("Nama Konten")}</h4><p style="margin:5px 0 0 0; color:#888aaa; font-size:0.9rem;">📍 Konten Media Sosial</p></div>', unsafe_allow_html=True)
                c1, c2 = st.columns(2)
                with c1:
                    st.write("**🎭 Jenis:**", get_m('Jenis'))
                    st.write("**📌 Rubrik:**", get_m('Rubrik'))
                    st.write("**✅ Status:**", get_m('Status'))
                with c2:
                    st.write("**✍️ Writer:**", get_m('Writer'))
                    st.write("**🎨 Designer:**", get_m('Designer/Editor'))
                    st.write("**⏰ Tayang:**", get_m('Jadwal Tayang'))
                sc1, sc2, sc3 = st.columns(3)
                with sc1: st.write("**📹 Kameramen:**", get_m('Kameramen'))
                with sc2: st.write("**🎙️ Cast/VO:**", get_m('Cast/VO'))
                with sc3: st.write("**🔧 Support:**", get_m('Support'))
            elif "Promosi" in sumber:
                def get_p(col):
                    v = row_data.get(col)
                    return str(v) if pd.notna(v) and str(v).strip().lower() != 'nan' and str(v).strip() != '' else "-"
                st.markdown(f'<div style="background:rgba(54,162,235,0.1); border-left:5px solid #36a2eb; padding:15px; border-radius:8px; margin-bottom:20px;"><h4 style="margin:0; color:#FFFFFF;">{get_p("Nama Konten")}</h4><p style="margin:5px 0 0 0; color:#888aaa; font-size:0.9rem;">📍 Promosi Statistik 2026</p></div>', unsafe_allow_html=True)
                c1, c2 = st.columns(2)
                with c1:
                    st.write("**🏢 Bidang:**", get_p('Bidang yang Bertugas'))
                    st.write("**📌 Jadwal:**", get_p('Jadwal Posting'))
                with c2:
                    st.write("**✅ Status:**", get_p('Status'))
                    st.write("**📈 Progress:**", get_p('Progress'))
            else:
                excluded = ['dt_ref', 'Sumber', 'Unnamed: 0', 'No', 'Column 2', 's']
                available = [c for c in row_data.index if c not in excluded and pd.notna(row_data[c]) and str(row_data[c]).strip().lower() != 'nan']
                st.markdown(f'<div style="background:rgba(255,255,255,0.05); border-left:5px solid #F26522; padding:15px; border-radius:8px; margin-bottom:20px;"><h4 style="margin:0; color:#FFFFFF;">{title}</h4><p style="margin:5px 0 0 0; color:#888aaa; font-size:0.9rem;">📍 {sumber}</p></div>', unsafe_allow_html=True)
                c1, c2 = st.columns(2)
                for i, col in enumerate(available):
                    with (c1 if i % 2 == 0 else c2): st.write(f"**{col}:**", str(row_data[col]))

        cal_result = calendar(events=events, options=calendar_options, key="hms_calendar")
        
        if cal_result and "eventClick" in cal_result:
            try:
                event_id = int(cal_result["eventClick"]["event"]["id"])
                df_cal_reset = df_cal.reset_index(drop=True)
                if event_id < len(df_cal_reset):
                    row = df_cal_reset.iloc[event_id]
                    show_agenda_detail(cal_result["eventClick"]["event"]["title"], row)
            except Exception as e: st.error(f"Gagal memuat detail: {e}")
            
        st.markdown("---")
        st.markdown("**Legenda Agenda:**")
        cols = st.columns(4)
        for i, (src, clr) in enumerate(source_colors.items()):
            with cols[i % 4]: st.markdown(f'<span style="color:{clr};">●</span> {src.replace("2026","")}', unsafe_allow_html=True)
        st.stop()
    else:
        st.warning("Belum ada data jadwal yang valid untuk ditampilkan di kalender.")
        st.stop()

# ===============================
# END OF APP
# ===============================
