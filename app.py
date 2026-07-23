import streamlit as st
import pandas as pd
import plotly.express as px
import os
import re

st.set_page_config(page_title="Verimlilik ve Performans Paneli", layout="wide")

EXCEL_FILE = "veri.xlsx"

# Güvenli Önbellek Fonksiyonu (Veriyi GitHub'daki veri.xlsx'ten çeker)
@st.cache_data(ttl=300)
def load_data_from_excel(file_path):
    if os.path.exists(file_path):
        excel = pd.ExcelFile(file_path)
        sheets = {s: excel.parse(s) for s in excel.sheet_names}
        return sheets, excel.sheet_names
    return None, []

# Tüm Tasarım ve CSS Ayarları
st.markdown("""
    <style>
    /* --- TABLO HÜCRE VE BAŞLIK HİZALAMALARI --- */
    th, td { text-align: center !important; }
    div[data-testid="stTable"] table, div[data-testid="stDataFrame"] table { width: 100%; text-align: center !important; }
    div[data-testid="stTable"] th, div[data-testid="stDataFrame"] th { text-align: center !important; }
    div[data-testid="stTable"] td, div[data-testid="stDataFrame"] td { text-align: center !important; }
    [data-testid="stDataFrame"] [role="gridcell"] { justify-content: center !important; text-align: center !important; }
    [data-testid="stDataFrame"] [role="columnheader"] { justify-content: center !important; text-align: center !important; }
    div[data-testid="stDataFrame"] > div { max-height: none !important; }

    /* --- TÜM TABLOLAR İÇİN FUŞYA ÇERÇEVE KURALI --- */
    div[data-testid="stDataFrame"], div[data-testid="stTable"] {
        border: 2px solid #FF007F !important;
        border-radius: 12px !important;
        padding: 4px !important;
        box-shadow: 0 0 10px rgba(255, 0, 127, 0.2) !important;
        overflow: hidden !important;
    }

    /* --- SEKME BAŞLIKLARI (RENKLİ ÇERÇEVELER & 2 SATIR METİN) --- */
    div[data-testid="stTabs"] > div[data-baseweb="tab-list"] {
        gap: 12px !important;
        margin-bottom: 15px !important;
    }
    
    div[data-testid="stTabs"] button[role="tab"] {
        white-space: normal !important; /* Metni 2 satıra kırmak için */
        word-break: break-word !important;
        width: 140px !important; /* Genişliği kısıtlayıp alt satıra geçmeye zorlar */
        min-height: 75px !important; /* 2 satır için yeterli alan */
        padding: 8px 5px !important;
        border-radius: 12px !important;
        border: 2px solid transparent !important;
        background-color: #131722 !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        text-align: center !important;
        transition: all 0.3s ease !important;
    }
    
    /* Sekme İçerisindeki Metinleri Ortala */
    div[data-testid="stTabs"] button[role="tab"] p {
        white-space: normal !important;
        text-align: center !important;
        line-height: 1.3 !important;
        font-size: 13px !important;
        font-weight: 700 !important;
        margin: 0 !important;
        padding: 0 !important;
    }

    /* Sekme Çerçeve ve Metin Renkleri */
    div[data-testid="stTabs"] button[role="tab"]:nth-child(1) { border-color: #FF4081 !important; background-color: rgba(255, 64, 129, 0.12) !important; }
    div[data-testid="stTabs"] button[role="tab"]:nth-child(1) p { color: #FF4081 !important; }
    
    div[data-testid="stTabs"] button[role="tab"]:nth-child(2) { border-color: #00B0FF !important; background-color: rgba(0, 176, 255, 0.12) !important; }
    div[data-testid="stTabs"] button[role="tab"]:nth-child(2) p { color: #00B0FF !important; }
    
    div[data-testid="stTabs"] button[role="tab"]:nth-child(3) { border-color: #00E676 !important; background-color: rgba(0, 230, 118, 0.12) !important; }
    div[data-testid="stTabs"] button[role="tab"]:nth-child(3) p { color: #00E676 !important; }
    
    div[data-testid="stTabs"] button[role="tab"]:nth-child(4) { border-color: #FF9100 !important; background-color: rgba(255, 145, 0, 0.12) !important; }
    div[data-testid="stTabs"] button[role="tab"]:nth-child(4) p { color: #FF9100 !important; }
    
    div[data-testid="stTabs"] button[role="tab"]:nth-child(5) { border-color: #FF1744 !important; background-color: rgba(255, 23, 68, 0.12) !important; }
    div[data-testid="stTabs"] button[role="tab"]:nth-child(5) p { color: #FF1744 !important; }
    
    div[data-testid="stTabs"] button[role="tab"]:nth-child(6) { border-color: #FFEA00 !important; background-color: rgba(255, 234, 0, 0.12) !important; }
    div[data-testid="stTabs"] button[role="tab"]:nth-child(6) p { color: #FFEA00 !important; }
    
    div[data-testid="stTabs"] button[role="tab"]:nth-child(7) { border-color: #00E5FF !important; background-color: rgba(0, 229, 255, 0.12) !important; }
    div[data-testid="stTabs"] button[role="tab"]:nth-child(7) p { color: #00E5FF !important; }
    
    div[data-testid="stTabs"] button[role="tab"]:nth-child(8) { border-color: #FF3D00 !important; background-color: rgba(255, 61, 0, 0.12) !important; }
    div[data-testid="stTabs"] button[role="tab"]:nth-child(8) p { color: #FF3D00 !important; }
    
    div[data-testid="stTabs"] button[role="tab"]:nth-child(9) { border-color: #D500F9 !important; background-color: rgba(213, 0, 249, 0.12) !important; }
    div[data-testid="stTabs"] button[role="tab"]:nth-child(9) p { color: #D500F9 !important; }

    /* Hover ve Active (Seçili) Efektleri */
    div[data-testid="stTabs"] button[role="tab"]:hover {
        transform: translateY(-4px) scale(1.03);
        box-shadow: 0 6px 15px rgba(255, 255, 255, 0.15) !important;
    }
    div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
        box-shadow: 0 0 12px currentColor !important;
        filter: brightness(1.2);
    }

    /* --- PERFORMANS MATRİSİ KART TASARIMI (ORTADAKİ 5 KART) --- */
    .matrix-card {
        background-color: #131722;
        border: 1px solid #1e222d;
        border-radius: 10px;
        padding: 15px;
        text-align: left;
    }
    .matrix-title { color: #00b0ff; font-size: 13px; font-weight: bold; letter-spacing: 0.5px; }
    .matrix-target { color: #8a8d93; font-size: 11px; margin-top: 2px; }
    .matrix-value { color: #ffffff; font-size: 26px; font-weight: 800; margin: 10px 0; }
    .matrix-badge { display: inline-block; background-color: #0e3a2f; color: #00e676; padding: 3px 10px; border-radius: 15px; font-size: 12px; font-weight: 600; }

    /* --- RENKLİ ÇERÇEVELİ KPI KUTULARI TASARIMI (ÜSTTEKİ 3 KUTU) --- */
    .kpi-card {
        background-color: #131722;
        border-radius: 12px;
        padding: 16px 20px;
        text-align: center;
        box-shadow: 0 4px 10px rgba(0,0,0,0.3);
        margin-bottom: 10px;
    }
    .kpi-title {
        font-size: 14px;
        font-weight: 600;
        color: #a0a5b5;
        margin-bottom: 8px;
    }
    .kpi-value {
        font-size: 32px;
        font-weight: 800;
        color: #ffffff;
    }

    /* Farklı Çerçeve Renkleri */
    .kpi-blue { border: 2px solid #00B0FF !important; box-shadow: 0 0 12px rgba(0, 176, 255, 0.25) !important; }
    .kpi-green { border: 2px solid #00E676 !important; box-shadow: 0 0 12px rgba(0, 230, 118, 0.25) !important; }
    .kpi-cyan { border: 2px solid #00E5FF !important; box-shadow: 0 0 12px rgba(0, 229, 255, 0.25) !important; }
    </style>
""", unsafe_allow_html=True)

sheets_dict, sheet_names = load_data_from_excel(EXCEL_FILE)

if sheets_dict is not None and len(sheet_names) > 0:
    df_raw = sheets_dict[sheet_names[0]].copy()

    haric_personel = ['CRM Admin', 'Luron AI API']
    
    if 'Görevi Alan' in df_raw.columns:
        df_raw = df_raw[~df_raw['Görevi Alan'].astype(str).str.strip().isin(haric_personel)].copy()
        temsilci_listesi = ["Tümü"] + sorted([str(x).strip() for x in df_raw['Görevi Alan'].dropna().unique()])
    else:
        temsilci_listesi = ["Tümü"]

    # -------------------------------------------------------------
    # KRİTER DIŞI SEKMEDEN REZERVASYON ALMA ORANINI ÇEKME
    # -------------------------------------------------------------
    rez_oran_dict = {}
    kd_sheet = None
    for s in sheet_names:
        s_norm = str(s).strip().lower().replace('ş', 's').replace('ı', 'i').replace('ğ', 'g').replace('ü', 'u').replace('ö', 'o').replace('ç', 'c')
        if "kriter" in s_norm and "disi" in s_norm and "hedef" not in s_norm:
            kd_sheet = s
            break
    
    if kd_sheet is not None:
        kd_df = sheets_dict[kd_sheet].copy()
        first_col = kd_df.columns[0]
        target_col = None
        for col in kd_df.columns:
            col_norm = str(col).lower()
            if "rez" in col_norm and "alma" in col_norm and "oran" in col_norm and "tarihi" in col_norm and "randevu" not in col_norm:
                target_col = col
                break
        if not target_col:
            for col in kd_df.columns:
                col_norm = str(col).lower()
                if "rez" in col_norm and "alma" in col_norm and "oran" in col_norm:
                    target_col = col
                    break
        
        if target_col is not None:
            for _, r in kd_df.iterrows():
                rep = str(r[first_col]).strip()
                v = r[target_col]
                if pd.notnull(v):
                    try:
                        rez_oran_dict[rep] = float(v)
                    except:
                        pass

    # -------------------------------------------------------------
    # ⚙️ SOL MENÜ
    # -------------------------------------------------------------
    st.sidebar.markdown("### ⚙️ Veri Kontrol Paneli")
    secilen_temsilci = st.sidebar.selectbox(
        "👤 Temsilci Ara (Dinamik)",
        options=temsilci_listesi,
        index=0
    )

    if st.sidebar.button("🔄 Verileri Yenile / Sıfırla", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    # Filtreleme Mantığı
    if secilen_temsilci != "Tümü" and 'Görevi Alan' in df_raw.columns:
        df = df_raw[df_raw['Görevi Alan'].astype(str).str.strip() == secilen_temsilci].copy()
    else:
        df = df_raw.copy()

    # --- ANA SAYFA BAŞLIK VE HESAPLAMALAR ---
    st.title("📊 Performans ve Özet Tablolar Paneli")

    # Genel Hesaplamalar
    toplam_gorev = len(df)
    tamamlanan = len(df[df['Görev Durumu'] == 'Tamamlandı']) if 'Görev Durumu' in df.columns else 0
    tamamlanmayan = toplam_gorev - tamamlanan
    tamamlanma_orani = (tamamlanan / toplam_gorev * 100) if toplam_gorev > 0 else 0

    # SOLDA METRİKLER (FARKLI RENKTE ÇERÇEVELER) - SAĞDA PASTA GRAFİĞİ
    kpi_col, pie_col = st.columns([1.2, 1])

    with kpi_col:
        m1, m2, m3 = st.columns(3)
        with m1:
            st.markdown(f"""
                <div class="kpi-card kpi-blue">
                    <div class="kpi-title">Toplam Görev Sayısı</div>
                    <div class="kpi-value">{toplam_gorev:,}</div>
                </div>
            """, unsafe_allow_html=True)
            
        with m2:
            st.markdown(f"""
                <div class="kpi-card kpi-green">
                    <div class="kpi-title">Tamamlanan Görev</div>
                    <div class="kpi-value">{tamamlanan:,}</div>
                </div>
            """, unsafe_allow_html=True)
            
        with m3:
            st.markdown(f"""
                <div class="kpi-card kpi-cyan">
                    <div class="kpi-title">Genel Tamamlanma Oranı</div>
                    <div class="kpi-value">%{tamamlanma_orani:.1f}</div>
                </div>
            """, unsafe_allow_html=True)

    with pie_col:
        pie_data = pd.DataFrame({
            'Durum': ['Tamamlandı', 'Tamamlanmadı'],
            'Adet': [tamamlanan, tamamlanmayan]
        })
        
        fig_pie = px.pie(
            pie_data, 
            values='Adet', 
            names='Durum', 
            hole=0.4,
            color='Durum',
            color_discrete_map={'Tamamlandı': '#00E676', 'Tamamlanmadı': '#FF1744'},
            title="Genel Görev Tamamlanma Dağılımı"
        )
        fig_pie.update_layout(
            template="plotly_dark",
            margin=dict(l=20, r=20, t=40, b=20),
            height=200,
            font=dict(size=12)
        )
        fig_pie.update_traces(textinfo='percent+label')
        st.plotly_chart(fig_pie, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # -------------------------------------------------------------
    # ⚡ ŞİRKET GENEL PERFORMANS MATRİSİ
    # -------------------------------------------------------------
    st.markdown("### ⚡ Şirket Genel Performans Matrisi")

    m_lead_h, m_lead_g, m_lead_o = "38,200", "14,283", "%37.4"
    m_rez_h, m_rez_g, m_rez_o = "808", "646", "%80.0"
    m_satis_h, m_satis_g, m_satis_o = "132", "49", "%37.1"
    m_kriter_h, m_kriter_g = "%20.0", "%34.4"
    m_gelme_h, m_gelme_g = "%40.0", "%41.3"

    genel_sheet = None
    for s in sheet_names:
        name_norm = str(s).strip().lower().replace('ş', 's').replace('ı', 'i').replace('ğ', 'g').replace('ü', 'u').replace('ö', 'o').replace('ç', 'c')
        if "genel" in name_norm and "hedef" in name_norm:
            genel_sheet = s
            break
    
    if genel_sheet is not None:
        gh_df = sheets_dict[genel_sheet].copy()
        for idx, r in gh_df.iterrows():
            row_label = str(r.iloc[0]).strip().lower()
            th = r.iloc[1] if pd.notnull(r.iloc[1]) else 0
            tg = r.iloc[2] if pd.notnull(r.iloc[2]) else 0
            
            if "lead" in row_label:
                m_lead_h = f"{int(th):,}"
                m_lead_g = f"{int(tg):,}"
                m_lead_o = f"%{(tg/th*100):.1f}" if th > 0 else "%0.0"
            elif "gelen" in row_label and "rezervasyon" in row_label:
                m_rez_h = f"{int(th):,}"
                m_rez_g = f"{int(tg):,}"
                m_rez_o = f"%{(tg/th*100):.1f}" if th > 0 else "%0.0"
            elif "satis" in row_label or "satış" in row_label:
                m_satis_h = f"{int(th):,}"
                m_satis_g = f"{int(tg):,}"
                m_satis_o = f"%{(tg/th*100):.1f}" if th > 0 else "%0.0"
            elif "kriter" in row_label:
                m_kriter_h = f"%{th:.1f}" if isinstance(th, (int, float)) else str(th)
                m_kriter_g = f"%{tg:.1f}"
            elif "gelme" in row_label:
                m_gelme_h = f"%{th:.1f}" if isinstance(th, (int, float)) else str(th)
                m_gelme_g = f"%{tg:.1f}"

    mc1, mc2, mc3, mc4, mc5 = st.columns(5)

    with mc1:
        st.markdown(f"""
            <div class="matrix-card">
                <div class="matrix-title">🔹 LEAD</div>
                <div class="matrix-target">Hedef: {m_lead_h}</div>
                <div class="matrix-value">{m_lead_g}</div>
                <div class="matrix-badge">↑ Başarı: {m_lead_o}</div>
            </div>
        """, unsafe_allow_html=True)

    with mc2:
        st.markdown(f"""
            <div class="matrix-card">
                <div class="matrix-title">🔹 GELEN REZERVASYON</div>
                <div class="matrix-target">Hedef: {m_rez_h}</div>
                <div class="matrix-value">{m_rez_g}</div>
                <div class="matrix-badge">↑ Başarı: {m_rez_o}</div>
            </div>
        """, unsafe_allow_html=True)

    with mc3:
        st.markdown(f"""
            <div class="matrix-card">
                <div class="matrix-title">🔹 SATIŞ</div>
                <div class="matrix-target">Hedef: {m_satis_h}</div>
                <div class="matrix-value">{m_satis_g}</div>
                <div class="matrix-badge">↑ Başarı: {m_satis_o}</div>
            </div>
        """, unsafe_allow_html=True)

    with mc4:
        st.markdown(f"""
            <div class="matrix-card">
                <div class="matrix-title">🔹 KRİTER DIŞI</div>
                <div class="matrix-target">Hedef: {m_kriter_h}</div>
                <div class="matrix-value">{m_kriter_g}</div>
                <div class="matrix-badge">↑ Gerçekleşen</div>
            </div>
        """, unsafe_allow_html=True)

    with mc5:
        st.markdown(f"""
            <div class="matrix-card">
                <div class="matrix-title">🔹 GELME ORANI</div>
                <div class="matrix-target">Hedef: {m_gelme_h}</div>
                <div class="matrix-value">{m_gelme_g}</div>
                <div class="matrix-badge">↑ Gerçekleşen</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("---")

    # SEKMELER (9 Sekme)
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9 = st.tabs([
        "👥 Temsilci Özet Tablosu", 
        "💬 Temsilci Yorumu", 
        "📞 Aksiyon Özet Tablosu", 
        "📍 İl & Marka Analizi",
        "🎯 Rezervasyon Hedef",
        "💰 Satış Hedef",
        "🚶‍♂️ Gelme Oranı Hedef",
        "🚫 Kriter Dışı Hedef",
        "📈 Data Analiz"
    ])

    # 1. TABLO & GRAFİK: TEMSİLCİ PERFORMANSI
    with tab1:
        st.subheader("Temsilci Bazlı Performans Özet Tablosu")
        if 'Görevi Alan' in df.columns and 'Görev Durumu' in df.columns:
            temsilci_ozet = df.groupby('Görevi Alan').agg(
                Toplam_Gorev=('Görev ID', 'count'),
                Tamamlanan=('Görev Durumu', lambda x: (x == 'Tamamlandı').sum()),
                Ort_Aksiyon=('Aksiyon Sayısı', 'mean')
            ).reset_index()

            # Dönüşüm Oranını Kriter Dışı sekmesindeki Rez. Alma Oranı (Rez. Alma Tarihi) kısmından al
            temsilci_ozet['Donusum_Orani_Val'] = temsilci_ozet['Görevi Alan'].astype(str).str.strip().map(rez_oran_dict).fillna(0)

            temsilci_ozet = temsilci_ozet.sort_values(by='Donusum_Orani_Val', ascending=False).reset_index(drop=True)
            temsilci_ozet['Ort_Aksiyon'] = temsilci_ozet['Ort_Aksiyon'].round(2)
            
            temsilci_ozet_gosterim = temsilci_ozet[[
                'Görevi Alan', 'Toplam_Gorev', 'Tamamlanan', 'Ort_Aksiyon', 'Donusum_Orani_Val'
            ]].copy()

            temsilci_ozet_gosterim.columns = [
                'Temsilci', 
                'Toplam Görev', 
                'Tamamlanan Görev', 
                'Ort. Aksiyon Sayısı', 
                'Rezervasyon Dönüşüm Oranı (%)'
            ]

            raw_donusum = temsilci_ozet_gosterim['Rezervasyon Dönüşüm Oranı (%)'].copy()
            temsilci_ozet_gosterim['Rezervasyon Dönüşüm Oranı (%)'] = raw_donusum.apply(lambda x: f"%{x:.1f}")

            def color_donusum(val):
                if val >= 6.0: return 'color: #00E676; font-weight: bold; text-align: center;'
                elif val >= 3.5: return 'color: #FFEA00; font-weight: bold; text-align: center;'
                else: return 'color: #FF1744; font-weight: bold; text-align: center;'

            styled_temsilci = temsilci_ozet_gosterim.style.apply(
                lambda col: [color_donusum(val) for val in raw_donusum] if col.name == 'Rezervasyon Dönüşüm Oranı (%)' else ['text-align: center;'] * len(col),
                axis=0
            ).hide(axis="index")

            col_config1 = {c: st.column_config.Column(alignment="center") for c in temsilci_ozet_gosterim.columns}
            
            calc_height = (len(temsilci_ozet_gosterim) + 1) * 35 + 38
            st.dataframe(styled_temsilci, use_container_width=True, height=calc_height, column_config=col_config1)

            st.markdown("---")
            st.markdown("### 📊 Temsilci Bazlı Rezervasyon Dönüşüm Oranı Grafiği")

            def get_donusum_status(val):
                if val >= 6.0: return 'Yüksek (>=%6.0)'
                elif val >= 3.5: return 'Orta (%3.5-%5.9)'
                else: return 'Düşük (<%3.5)'

            chart_temsilci = temsilci_ozet.copy()
            chart_temsilci['Performans Durumu'] = chart_temsilci['Donusum_Orani_Val'].apply(get_donusum_status)

            color_map_donusum = {'Yüksek (>=%6.0)': '#00E676', 'Orta (%3.5-%5.9)': '#FFEA00', 'Düşük (<%3.5)': '#FF1744'}

            fig_temsilci = px.bar(
                chart_temsilci, x='Görevi Alan', y='Donusum_Orani_Val',
                color='Performans Durumu', color_discrete_map=color_map_donusum,
                text=chart_temsilci['Donusum_Orani_Val'].apply(lambda x: f"%{x:.1f}"),
                title="Temsilcilere Göre Dönüşüm Oranları (%)"
            )
            fig_temsilci.update_layout(template="plotly_dark", xaxis_title="Temsilci", yaxis_title="Dönüşüm Oranı (%)")
            fig_temsilci.update_traces(textposition='outside')
            st.plotly_chart(fig_temsilci, use_container_width=True)

    # 2. TABLO: TEMSİLCİ YORUMU
    with tab2:
        st.subheader("Temsilci Yorumları ve Performans Analizi")
        ty_sheet = None
        for s in sheet_names:
            name_norm = str(s).strip().lower().replace('ş', 's').replace('ı', 'i').replace('ğ', 'g').replace('ü', 'u').replace('ö', 'o').replace('ç', 'c')
            if "temsilci" in name_norm and "yorum" in name_norm:
                ty_sheet = s
                break

        if ty_sheet is not None:
            ty_df = sheets_dict[ty_sheet].copy()
            if len(ty_df.columns) > 0:
                ty_df = ty_df[~ty_df.iloc[:, 0].astype(str).isin(haric_personel)].copy()
                if secilen_temsilci != "Tümü":
                    ty_df = ty_df[ty_df.iloc[:, 0].astype(str).str.strip() == secilen_temsilci]

            for idx, row in ty_df.iterrows():
                temsilci_adi = str(row.iloc[0]).strip()
                yorum_metni = str(row.iloc[1]) if len(row) > 1 and pd.notnull(row.iloc[1]) else "Yorum bulunamadı."
                
                # Rezervasyon Dönüşüm Oranını Kriter Dışı sekmesindeki Rez. Alma Oranı ile dinamik değiştir
                if temsilci_adi in rez_oran_dict:
                    val = rez_oran_dict[temsilci_adi]
                    formatted_val = f"%{val:.1f}"
                    yorum_metni = re.sub(r'🔄 Rezervasyon Dönüşümü:\s*[\d,\.]*%?', f'🔄 Rezervasyon Dönüşümü: {formatted_val}', yorum_metni)

                with st.expander(f"👤 **{temsilci_adi}** — Performans Yorumu", expanded=True):
                    for emoji in ['✅', '⚠️', '🎯', '🔄', '🚫', '📌']:
                        yorum_metni = yorum_metni.replace(emoji, f"\n{emoji}")
                    
                    satirlar = [s.strip() for s in yorum_metni.split('\n') if s.strip()]
                    for satir in satirlar:
                        st.markdown(satir)
        else:
            st.warning("⚠️ Excel dosyanızda 'Temsilci Yorumu' sekmesi bulunamadı.")

    # 3. TABLO: AKSİYON SONUÇLARI
    with tab3:
        st.subheader("Son Aksiyon ve Son Arama Dağılım Tablosu")
        col_t1, col_t2 = st.columns(2)

        with col_t1:
            if 'Son Aksiyon' in df.columns:
                st.markdown("**Son Aksiyon Türleri Özeti**")
                aksiyon_ozet = df['Son Aksiyon'].value_counts().reset_index()
                aksiyon_ozet.columns = ['Aksiyon Nedeni', 'Adet']
                toplam_aksiyon = aksiyon_ozet['Adet'].sum()
                aksiyon_ozet['Oran (%)'] = ((aksiyon_ozet['Adet'] / toplam_aksiyon) * 100).map(lambda x: f"%{x:.1f}")
                st.table(aksiyon_ozet)

        with col_t2:
            if 'Son Arama' in df.columns:
                st.markdown("**Son Arama Durumu Özeti**")
                arama_ozet = df['Son Arama'].value_counts().reset_index()
                arama_ozet.columns = ['Arama Durumu', 'Adet']
                toplam_arama = arama_ozet['Adet'].sum()
                arama_ozet['Oran (%)'] = ((arama_ozet['Adet'] / toplam_arama) * 100).map(lambda x: f"%{x:.1f}")
                st.table(arama_ozet)

    # 4. TABLO: İL & MARKA KIRILIMI
    with tab4:
        st.subheader("Bölgesel ve Marka Bazlı Görev Tablosu")
        if 'İl' in df.columns and 'Marka' in df.columns:
            il_marka_ozet = pd.crosstab(df['İl'], df['Marka'], margins=True, margins_name="TOPLAM")
            st.table(il_marka_ozet)

    # 5. TABLO & GRAFİK: REZERVASYON HEDEF
    with tab5:
        st.subheader("🎯 Rezervasyon Hedef Tablosu ve Performans Grafiği")
        target_sheet = None
        for s in sheet_names:
            name_clean = str(s).strip().lower().replace('ı', 'i').replace('ş', 's')
            if "rezervasyon" in name_clean and "hedef" in name_clean:
                target_sheet = s
                break

        if target_sheet is not None:
            rez_df = sheets_dict[target_sheet].copy()
            if len(rez_df.columns) > 0:
                rez_df = rez_df[~rez_df.iloc[:, 0].astype(str).isin(haric_personel)].copy()
                if secilen_temsilci != "Tümü":
                    filtered_rez = rez_df[rez_df.iloc[:, 0].astype(str).str.strip() == secilen_temsilci]
                    if not filtered_rez.empty:
                        rez_df = filtered_rez

            for c in rez_df.columns:
                if "oran" not in str(c).lower() and "%" not in str(c) and c != rez_df.columns[0]:
                    rez_df[c] = pd.to_numeric(rez_df[c], errors='coerce').fillna(0).round().astype(int)

            oran_col = [c for c in rez_df.columns if "oran" in str(c).lower() or "%" in str(c)]
            if oran_col:
                target_oran_col = oran_col[0]
                raw_values = rez_df[target_oran_col].apply(
                    lambda x: x * 100 if isinstance(x, (int, float)) and x <= 2 else (x if isinstance(x, (int, float)) else 0)
                )
                
                def get_status(val):
                    if val >= 100: return 'Yüksek (>=%100)'
                    elif val >= 80: return 'Orta (%80-%99)'
                    else: return 'Düşük (<%80)'

                def color_cell(val):
                    if val >= 100: return 'color: #00E676; font-weight: bold; text-align: center;'
                    elif val >= 80: return 'color: #FFEA00; font-weight: bold; text-align: center;'
                    else: return 'color: #FF1744; font-weight: bold; text-align: center;'

                display_df = rez_df.copy()
                display_df[target_oran_col] = raw_values.apply(lambda x: f"%{x:.1f}")

                styled_df = display_df.style.apply(
                    lambda col: [color_cell(val) for val in raw_values] if col.name == target_oran_col else ['text-align: center;'] * len(col),
                    axis=0
                ).hide(axis="index")
                
                col_config_rez = {c: st.column_config.Column(alignment="center") for c in display_df.columns}
                calc_height_rez = (len(display_df) + 1) * 35 + 38
                st.dataframe(styled_df, use_container_width=True, height=calc_height_rez, column_config=col_config_rez)

                st.markdown("### 📊 Temsilci Bazlı Gerçekleşen Rezervasyon Performansı")
                chart_df = rez_df[rez_df.iloc[:, 0].astype(str).str.lower() != 'toplam'].copy()
                chart_df['Oran_Val'] = raw_values
                chart_df['Performans Durumu'] = chart_df['Oran_Val'].apply(get_status)

                color_discrete_map = {'Yüksek (>=%100)': '#00E676', 'Orta (%80-%99)': '#FFEA00', 'Düşük (<%80)': '#FF1744'}
                y_col = chart_df.columns[3] if len(chart_df.columns) > 3 else chart_df.columns[1]

                fig = px.bar(
                    chart_df, x=chart_df.columns[0], y=y_col,
                    color='Performans Durumu', color_discrete_map=color_discrete_map,
                    text=y_col, title="Temsilcilere Göre Gerçekleşen Rezervasyon ve Performans Kategorisi"
                )
                fig.update_layout(template="plotly_dark", xaxis_title="Temsilci", yaxis_title="Gerçekleşen Adet")
                fig.update_traces(textposition='outside')
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.table(rez_df)

    # 6. TABLO & GRAFİK: SATIŞ HEDEF
    with tab6:
        st.subheader("💰 Satış Hedef Tablosu ve Performans Grafiği")
        satis_sheet = None
        for s in sheet_names:
            name_norm = str(s).strip().lower().replace('ş', 's').replace('ı', 'i').replace('ğ', 'g').replace('ü', 'u').replace('ö', 'o').replace('ç', 'c')
            if "satis" in name_norm and "hedef" in name_norm:
                satis_sheet = s
                break

        if satis_sheet is not None:
            satis_df = sheets_dict[satis_sheet].copy()
            if len(satis_df.columns) > 0:
                satis_df = satis_df[~satis_df.iloc[:, 0].astype(str).isin(haric_personel)].copy()
                if secilen_temsilci != "Tümü":
                    filtered_satis = satis_df[satis_df.iloc[:, 0].astype(str).str.strip() == secilen_temsilci]
                    if not filtered_satis.empty:
                        satis_df = filtered_satis

            for c in satis_df.columns:
                if "oran" not in str(c).lower() and "%" not in str(c) and c != satis_df.columns[0]:
                    satis_df[c] = pd.to_numeric(satis_df[c], errors='coerce').fillna(0).round().astype(int)

            satis_oran_col = [c for c in satis_df.columns if "oran" in str(c).lower() or "%" in str(c)]
            if satis_oran_col:
                target_satis_oran_col = satis_oran_col[0]
                raw_satis_values = satis_df[target_satis_oran_col].apply(
                    lambda x: x * 100 if isinstance(x, (int, float)) and x <= 2 else (x if isinstance(x, (int, float)) else 0)
                )
                
                def get_satis_status(val):
                    if val >= 50: return 'Yüksek (>=%50)'
                    elif val >= 35: return 'Orta (%35-%49)'
                    else: return 'Düşük (<%35)'

                def color_satis_cell(val):
                    if val >= 50: return 'color: #00E676; font-weight: bold; text-align: center;'
                    elif val >= 35: return 'color: #FFEA00; font-weight: bold; text-align: center;'
                    else: return 'color: #FF1744; font-weight: bold; text-align: center;'

                display_satis_df = satis_df.copy()
                display_satis_df[target_satis_oran_col] = raw_satis_values.apply(lambda x: f"%{x:.0f}" if x.is_integer() else f"%{x:.1f}")

                styled_satis_df = display_satis_df.style.apply(
                    lambda col: [color_satis_cell(val) for val in raw_satis_values] if col.name == target_satis_oran_col else ['text-align: center;'] * len(col),
                    axis=0
                ).hide(axis="index")
                
                col_config_satis = {c: st.column_config.Column(alignment="center") for c in display_satis_df.columns}
                calc_height_satis = (len(display_satis_df) + 1) * 35 + 38
                st.dataframe(styled_satis_df, use_container_width=True, height=calc_height_satis, column_config=col_config_satis)

                st.markdown("### 📊 Temsilci Bazlı Gerçekleşen Satış Performansı")
                chart_satis_df = satis_df[satis_df.iloc[:, 0].astype(str).str.lower() != 'toplam'].copy()
                chart_satis_df['Oran_Val'] = raw_satis_values
                chart_satis_df['Satış Durumu'] = chart_satis_df['Oran_Val'].apply(get_satis_status)

                color_satis_map = {'Yüksek (>=%50)': '#00E676', 'Orta (%35-%49)': '#FFEA00', 'Düşük (<%35)': '#FF1744'}
                y_satis_col = chart_satis_df.columns[2] if len(chart_satis_df.columns) > 2 else chart_satis_df.columns[1]

                fig_satis = px.bar(
                    chart_satis_df, x=chart_satis_df.columns[0], y=y_satis_col,
                    color='Satış Durumu', color_discrete_map=color_satis_map,
                    text=y_satis_col, title="Temsilcilere Göre Gerçekleşen Satış Adetleri ve Performans Kategorisi"
                )
                fig_satis.update_layout(template="plotly_dark", xaxis_title="Temsilci", yaxis_title="Gerçekleşen Satış Adedi")
                fig_satis.update_traces(textposition='outside')
                st.plotly_chart(fig_satis, use_container_width=True)
            else:
                st.table(satis_df)

    # 7. TABLO & GRAFİK: GELME ORANI HEDEF
    with tab7:
        st.subheader("🚶‍♂️ Gelme Oranı Hedef Tablosu ve Performans Grafiği")
        gelme_sheet = None
        for s in sheet_names:
            name_norm = str(s).strip().lower().replace('ş', 's').replace('ı', 'i').replace('ğ', 'g').replace('ü', 'u').replace('ö', 'o').replace('ç', 'c')
            if "gelme" in name_norm and "oran" in name_norm:
                gelme_sheet = s
                break

        if gelme_sheet is not None:
            gelme_df = sheets_dict[gelme_sheet].copy()
            if len(gelme_df.columns) > 0:
                gelme_df = gelme_df[~gelme_df.iloc[:, 0].astype(str).isin(haric_personel)].copy()
                if secilen_temsilci != "Tümü":
                    filtered_gelme = gelme_df[gelme_df.iloc[:, 0].astype(str).str.strip() == secilen_temsilci]
                    if not filtered_gelme.empty:
                        gelme_df = filtered_gelme

            hedef_col = None
            gerceklesen_col = None

            for c in gelme_df.columns:
                c_norm = str(c).lower()
                if "hedef" in c_norm and "gercekles" not in c_norm and "gerçekleş" not in c_norm: hedef_col = c
                elif "gercekles" in c_norm or "gerçekleş" in c_norm or "oran" in c_norm: gerceklesen_col = c

            if not hedef_col and len(gelme_df.columns) > 1: hedef_col = gelme_df.columns[1]
            if not gerceklesen_col and len(gelme_df.columns) > 2: gerceklesen_col = gelme_df.columns[2]

            display_gelme_df = gelme_df.copy()

            if hedef_col:
                display_gelme_df[hedef_col] = display_gelme_df[hedef_col].apply(
                    lambda x: f"%{int(round(x*100))}" if isinstance(x, (int, float)) and x <= 2 else (f"%{int(round(x))}" if isinstance(x, (int, float)) else str(x))
                )

            if gerceklesen_col:
                raw_gerceklesen_values = gelme_df[gerceklesen_col].apply(
                    lambda x: x * 100 if isinstance(x, (int, float)) and x <= 2 else (x if isinstance(x, (int, float)) else 0)
                )

                display_gelme_df[gerceklesen_col] = raw_gerceklesen_values.apply(lambda x: f"%{x:.0f}" if x.is_integer() else f"%{x:.1f}")

                def get_gelme_status(val): return 'Başarılı (>=%40)' if val >= 40 else 'Düşük (<%40)'
                def color_gerceklesen_cell(val): return 'color: #00E676; font-weight: bold; text-align: center;' if val >= 40 else 'color: #FF1744; font-weight: bold; text-align: center;'

                styled_gelme_df = display_gelme_df.style.apply(
                    lambda col: [color_gerceklesen_cell(val) for val in raw_gerceklesen_values] if col.name == gerceklesen_col else ['text-align: center;'] * len(col),
                    axis=0
                ).hide(axis="index")

                col_config_gelme = {c: st.column_config.Column(alignment="center") for c in display_gelme_df.columns}
                calc_height_gelme = (len(display_gelme_df) + 1) * 35 + 38
                st.dataframe(styled_gelme_df, use_container_width=True, height=calc_height_gelme, column_config=col_config_gelme)

                st.markdown("### 📊 Temsilci Bazlı Gelme Oranı Performansı")
                chart_gelme_df = gelme_df[gelme_df.iloc[:, 0].astype(str).str.lower() != 'toplam'].copy()
                chart_gelme_df['Oran_Val'] = raw_gerceklesen_values
                chart_gelme_df['Gelme Durumu'] = chart_gelme_df['Oran_Val'].apply(get_gelme_status)

                color_gelme_map = {'Başarılı (>=%40)': '#00E676', 'Düşük (<%40)': '#FF1744'}
                fig_gelme = px.bar(
                    chart_gelme_df, x=chart_gelme_df.columns[0], y='Oran_Val',
                    color='Gelme Durumu', color_discrete_map=color_gelme_map,
                    text=chart_gelme_df['Oran_Val'].apply(lambda x: f"%{x:.1f}"),
                    title="Temsilcilere Göre Gelme Oranı (%) ve Performans Durumu"
                )
                fig_gelme.update_layout(template="plotly_dark", xaxis_title="Temsilci", yaxis_title="Gerçekleşen Gelme Oranı (%)")
                fig_gelme.update_traces(textposition='outside')
                st.plotly_chart(fig_gelme, use_container_width=True)
            else:
                st.table(gelme_df)

    # 8. TABLO & GRAFİK: KRİTER DIŞI HEDEF
    with tab8:
        st.subheader("🚫 Kriter Dışı Hedef Tablosu ve Performans Grafiği")
        kriter_sheet = None
        for s in sheet_names:
            name_norm = str(s).strip().lower().replace('ş', 's').replace('ı', 'i').replace('ğ', 'g').replace('ü', 'u').replace('ö', 'o').replace('ç', 'c')
            if "kriter" in name_norm and "hedef" in name_norm:
                kriter_sheet = s
                break

        if kriter_sheet is not None:
            kriter_df = sheets_dict[kriter_sheet].copy()
            if len(kriter_df.columns) > 0:
                kriter_df = kriter_df[~kriter_df.iloc[:, 0].astype(str).isin(haric_personel)].copy()
                if secilen_temsilci != "Tümü":
                    filtered_kriter = kriter_df[kriter_df.iloc[:, 0].astype(str).str.strip() == secilen_temsilci]
                    if not filtered_kriter.empty:
                        kriter_df = filtered_kriter

            kriter_hedef_col = None
            kriter_gerceklesen_col = None

            for c in kriter_df.columns:
                c_norm = str(c).lower()
                if "hedef" in c_norm and "gercekles" not in c_norm and "gerçekleş" not in c_norm: kriter_hedef_col = c
                elif "gercekles" in c_norm or "gerçekleş" in c_norm or "oran" in c_norm: kriter_gerceklesen_col = c

            if not kriter_hedef_col and len(kriter_df.columns) > 1: kriter_hedef_col = kriter_df.columns[1]
            if not kriter_gerceklesen_col and len(kriter_df.columns) > 2: kriter_gerceklesen_col = kriter_df.columns[2]

            display_kriter_df = kriter_df.copy()

            if kriter_hedef_col:
                display_kriter_df[kriter_hedef_col] = display_kriter_df[kriter_hedef_col].apply(
                    lambda x: f"%{int(round(x*100))}" if isinstance(x, (int, float)) and x <= 2 else (f"%{int(round(x))}" if isinstance(x, (int, float)) else str(x))
                )

            if kriter_gerceklesen_col:
                raw_kriter_values = kriter_df[kriter_gerceklesen_col].apply(
                    lambda x: x * 100 if isinstance(x, (int, float)) and x <= 2 else (x if isinstance(x, (int, float)) else 0)
                )

                display_kriter_df[kriter_gerceklesen_col] = raw_kriter_values.apply(lambda x: f"%{x:.0f}" if x.is_integer() else f"%{x:.1f}")

                def get_kriter_status(val): return 'Yüksek/Riskli (>%20)' if val > 20 else 'İdeal (<=%20)'
                def color_kriter_cell(val): return 'color: #FF1744; font-weight: bold; text-align: center;' if val > 20 else 'color: #00E676; font-weight: bold; text-align: center;'

                styled_kriter_df = display_kriter_df.style.apply(
                    lambda col: [color_kriter_cell(val) for val in raw_kriter_values] if col.name == kriter_gerceklesen_col else ['text-align: center;'] * len(col),
                    axis=0
                ).hide(axis="index")

                col_config_kriter = {c: st.column_config.Column(alignment="center") for c in display_kriter_df.columns}
                calc_height_kriter = (len(display_kriter_df) + 1) * 35 + 38
                st.dataframe(styled_kriter_df, use_container_width=True, height=calc_height_kriter, column_config=col_config_kriter)

                st.markdown("### 📊 Temsilci Bazlı Kriter Dışı Oranı Performansı")
                chart_kriter_df = kriter_df[kriter_df.iloc[:, 0].astype(str).str.lower() != 'toplam'].copy()
                chart_kriter_df['Oran_Val'] = raw_kriter_values
                chart_kriter_df['Kriter Durumu'] = chart_kriter_df['Oran_Val'].apply(get_kriter_status)

                color_kriter_map = {'İdeal (<=%20)': '#00E676', 'Yüksek/Riskli (>%20)': '#FF1744'}
                fig_kriter = px.bar(
                    chart_kriter_df, x=chart_kriter_df.columns[0], y='Oran_Val',
                    color='Kriter Durumu', color_discrete_map=color_kriter_map,
                    text=chart_kriter_df['Oran_Val'].apply(lambda x: f"%{x:.1f}"),
                    title="Temsilcilere Göre Kriter Dışı Oranı (%) ve Performans Durumu"
                )
                fig_kriter.update_layout(template="plotly_dark", xaxis_title="Temsilci", yaxis_title="Gerçekleşen Kriter Dışı Oran (%)")
                fig_kriter.update_traces(textposition='outside')
                st.plotly_chart(fig_kriter, use_container_width=True)
            else:
                st.table(kriter_df)

    # 9. TABLO & GRAFİK: EXCEL'DEKİ "DATA ANALİZ" SEKME
    with tab9:
        st.subheader("📈 Data Analiz Tablosu ve Arama Sonuçları Grafiği")
        da_sheet = None
        for s in sheet_names:
            name_norm = str(s).strip().lower().replace('ş', 's').replace('ı', 'i').replace('ğ', 'g').replace('ü', 'u').replace('ö', 'o').replace('ç', 'c')
            if "data" in name_norm and "analiz" in name_norm:
                da_sheet = s
                break

        if da_sheet is not None:
            da_df = sheets_dict[da_sheet].copy()
            if len(da_df.columns) > 0:
                da_df = da_df[~da_df.iloc[:, 0].astype(str).isin(haric_personel)].copy()
                if secilen_temsilci != "Tümü":
                    filtered_da = da_df[da_df.iloc[:, 0].astype(str).str.strip() == secilen_temsilci]
                    if not filtered_da.empty:
                        da_df = filtered_da

            if len(da_df.columns) == 12:
                da_df.columns = [
                    'Görevi Alan', 'Konuşuldu', '% Oran', 
                    'Cevapsız (Açmadı)', '% Oran ', 
                    'Meşgul', '% Oran  ', 
                    'Cevapsız', '% Oran   ', 
                    'Toplam Gelen Data', 'Açık Görev', 'Toplam'
                ]

            display_da = da_df.copy()
            for col in display_da.columns:
                if col != display_da.columns[0]:
                    display_da[col] = pd.to_numeric(display_da[col], errors='coerce').fillna(0)
                    if "%" in str(col) or "oran" in str(col).lower():
                        display_da[col] = display_da[col].apply(
                            lambda x: f"%{x*100:.2f}" if 0 < x <= 1 else (f"%{x:.2f}" if x > 1 else "%0.00")
                        )
                    else:
                        display_da[col] = display_da[col].round().astype(int)

            col_config_da = {c: st.column_config.Column(alignment="center") for c in display_da.columns}
            calc_height_da = (len(display_da) + 1) * 35 + 38
            styled_da = display_da.style.hide(axis="index")
            st.dataframe(styled_da, use_container_width=True, height=calc_height_da, column_config=col_config_da)

            st.markdown("---")
            st.markdown("### 📊 Temsilci Bazlı Arama ve Görüşme Durumları Dağılımı")

            chart_da = da_df[da_df.iloc[:, 0].astype(str).str.lower() != 'toplam'].copy()
            durum_sutunlari = ['Konuşuldu', 'Cevapsız (Açmadı)', 'Meşgul', 'Cevapsız']
            for col in durum_sutunlari:
                if col in chart_da.columns:
                    chart_da[col] = pd.to_numeric(chart_da[col], errors='coerce').fillna(0)

            durum_sutunlari = [c for c in durum_sutunlari if c in chart_da.columns]

            if durum_sutunlari:
                chart_melted = chart_da.melt(
                    id_vars=[chart_da.columns[0]], 
                    value_vars=durum_sutunlari,
                    var_name='Arama Durumu', 
                    value_name='Adet'
                )

                color_map = {
                    'Konuşuldu': '#00E676',
                    'Cevapsız (Açmadı)': '#FFEA00',
                    'Meşgul': '#FF9100',
                    'Cevapsız': '#FF1744'
                }

                fig_da = px.bar(
                    chart_melted, x=chart_da.columns[0], y='Adet',
                    color='Arama Durumu', barmode='group', color_discrete_map=color_map,
                    text='Adet', title="Temsilcilere Göre Görüşme ve Arama Sonuçları Hacmi"
                )
                fig_da.update_layout(template="plotly_dark", xaxis_title="Temsilci", yaxis_title="Arama / Görüşme Adedi")
                fig_da.update_traces(textposition='outside')
                st.plotly_chart(fig_da, use_container_width=True)
        else:
            st.warning("⚠️ Excel dosyanızda 'Data Analiz' sekmesi bulunamadı.")
else:
    st.error("⚠️ Proje klasöründe 'veri.xlsx' dosyası bulunamadı. Lütfen GitHub'a 'veri.xlsx' dosyanızı yüklediğinizden emin olun.")
