import streamlit as st
import pandas as pd
import plotly.express as px
import os

# Sayfa Yapılandırması
st.set_page_config(
    page_title="Hedef ve Verimlilik Paneli", 
    layout="wide", 
    page_icon="📊",
    initial_sidebar_state="expanded"
)

EXCEL_FILE = "veri.xlsx"

# Önbellek Temizleme Fonksiyonu
@st.cache_data(ttl=300)
def load_data(file_path):
    if os.path.exists(file_path):
        return pd.ExcelFile(file_path)
    return None

# Custom CSS - Koyu Tasarım + Fuşya Çerçeveler + Sol Menü Tasarımı
st.markdown("""
    <style>
    .stApp { background-color: #0e1117; }
    
    /* Sol Menü (Sidebar) Özel Tasarımı */
    [data-testid="stSidebar"] {
        background-color: #131722 !important;
        border-right: 1px solid #1e222d !important;
    }
    
    th, td { text-align: center !important; }
    div[data-testid="stTable"] table, div[data-testid="stDataFrame"] table { width: 100%; text-align: center !important; }
    div[data-testid="stTable"] th, div[data-testid="stDataFrame"] th { text-align: center !important; background-color: #1a1c23 !important; color: #00e5ff !important; }
    div[data-testid="stTable"] td, div[data-testid="stDataFrame"] td { text-align: center !important; }
    [data-testid="stDataFrame"] [role="gridcell"], [data-testid="stDataFrame"] [role="columnheader"] { justify-content: center !important; text-align: center !important; }
    div[data-testid="stDataFrame"] > div { max-height: none !important; }

    /* Fuşya Renkli Tablo Çerçeveleri */
    div[data-testid="stDataFrame"], div[data-testid="stTable"] {
        border: 2px solid #FF007F !important;
        border-radius: 12px !important;
        padding: 4px !important;
        box-shadow: 0 0 12px rgba(255, 0, 127, 0.25) !important;
        background-color: #131722 !important;
        overflow: hidden !important;
    }

    /* Renkli Sekme (Tab) Butonları */
    button[data-baseweb="tab-tab-list"] { flex-wrap: wrap !important; gap: 10px !important; justify-content: flex-start !important; }
    button[data-baseweb="tab"] { white-space: normal !important; height: auto !important; padding: 10px 16px !important; border-radius: 8px !important; border-style: solid !important; border-width: 1.5px !important; font-weight: 600 !important; transition: all 0.3s ease !important; }

    button[data-baseweb="tab"]:nth-child(1) { border-color: #FF4081 !important; background-color: rgba(255, 64, 129, 0.1) !important; }
    button[data-baseweb="tab"]:nth-child(2) { border-color: #00B0FF !important; background-color: rgba(0, 176, 255, 0.1) !important; }
    button[data-baseweb="tab"]:nth-child(3) { border-color: #00E676 !important; background-color: rgba(0, 230, 118, 0.1) !important; }
    button[data-baseweb="tab"]:nth-child(4) { border-color: #FF9100 !important; background-color: rgba(255, 145, 0, 0.1) !important; }
    button[data-baseweb="tab"]:nth-child(5) { border-color: #FF1744 !important; background-color: rgba(255, 23, 68, 0.1) !important; }
    button[data-baseweb="tab"]:nth-child(6) { border-color: #FFEA00 !important; background-color: rgba(255, 234, 0, 0.1) !important; }
    button[data-baseweb="tab"]:nth-child(7) { border-color: #00E5FF !important; background-color: rgba(0, 229, 255, 0.1) !important; }
    button[data-baseweb="tab"]:nth-child(8) { border-color: #FF3D00 !important; background-color: rgba(255, 61, 0, 0.1) !important; }
    button[data-baseweb="tab"]:nth-child(9) { border-color: #D500F9 !important; background-color: rgba(213, 0, 249, 0.1) !important; }

    button[data-baseweb="tab"]:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(255, 255, 255, 0.15); }

    /* Performans Matrisi Kart Tasarımı */
    .matrix-card { background-color: #131722; border: 1px solid #1e222d; border-radius: 12px; padding: 16px; text-align: left; box-shadow: 0 4px 6px rgba(0,0,0,0.3); }
    .matrix-title { color: #00b0ff; font-size: 13px; font-weight: bold; letter-spacing: 0.5px; }
    .matrix-target { color: #8a8d93; font-size: 12px; margin-top: 3px; }
    .matrix-value { color: #ffffff; font-size: 26px; font-weight: 800; margin: 10px 0; }
    .matrix-badge { display: inline-block; background-color: #0e3a2f; color: #00e676; padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: 600; }
    </style>
""", unsafe_allow_html=True)

excel_file = load_data(EXCEL_FILE)

if excel_file is not None:
    sheet_names = excel_file.sheet_names
    df_raw = pd.read_excel(excel_file, sheet_name=0)

    # HARİÇ TUTULACAK PERSONEL
    haric_personel = ['CRM Admin', 'Luron AI API', 'Aleyna Daşdemir', 'Zeynep Güzel']
    
    if 'Görevi Alan' in df_raw.columns:
        df_raw = df_raw[~df_raw['Görevi Alan'].astype(str).isin(haric_personel)].copy()
        temsilci_listesi = ["Tümü"] + sorted([str(x) for x in df_raw['Görevi Alan'].dropna().unique()])
    else:
        temsilci_listesi = ["Tümü"]

    # --- SOL MENÜ (SIDEBAR) ---
    st.sidebar.header("👤 Temsilci Paneli")
    secilen_temsilci = st.sidebar.selectbox(
        "Temsilci İsmi Yazın veya Seçin:",
        options=temsilci_listesi,
        index=0
    )

    # Güncelleme Butonu
    if st.sidebar.button("🔄 Güncelle", use_container_width=True, type="primary"):
        st.cache_data.clear()
        st.rerun()

    st.sidebar.markdown("---")
    if secilen_temsilci != "Tümü":
        st.sidebar.info(f"🎯 Aktif Temsilci:\n**{secilen_temsilci}**")
    else:
        st.sidebar.caption("🌐 Tüm şirket verileri gösteriliyor.")

    # Temsilci Filtreleme Uygula
    if secilen_temsilci != "Tümü" and 'Görevi Alan' in df_raw.columns:
        df = df_raw[df_raw['Görevi Alan'] == secilen_temsilci].copy()
    else:
        df = df_raw.copy()

    # --- ANA SAYFA ---
    st.title("📊 Performans ve Özet Tablolar Paneli")

    # Genel KPI Hesaplamaları
    toplam_gorev = len(df)
    tamamlanan = len(df[df['Görev Durumu'] == 'Tamamlandı']) if 'Görev Durumu' in df.columns else 0
    tamamlanmayan = toplam_gorev - tamamlanan
    tamamlanma_orani = (tamamlanan / toplam_gorev * 100) if toplam_gorev > 0 else 0

    kpi_col, pie_col = st.columns([1.2, 1])

    with kpi_col:
        m1, m2, m3 = st.columns(3)
        m1.metric("Toplam Görev Sayısı", f"{toplam_gorev:,}")
        m2.metric("Tamamlanan Görev", f"{tamamlanan:,}")
        m3.metric("Genel Tamamlanma Oranı", f"%{tamamlanma_orani:.1f}")

    with pie_col:
        pie_data = pd.DataFrame({'Durum': ['Tamamlandı', 'Tamamlanmadı'], 'Adet': [tamamlanan, tamamlanmayan]})
        fig_pie = px.pie(pie_data, values='Adet', names='Durum', hole=0.4, color='Durum',
                         color_discrete_map={'Tamamlandı': '#00E676', 'Tamamlanmadı': '#FF1744'},
                         title="Görev Tamamlanma Dağılımı")
        fig_pie.update_layout(template="plotly_dark", margin=dict(l=20, r=20, t=40, b=20), height=180, font=dict(size=12))
        fig_pie.update_traces(textinfo='percent+label')
        st.plotly_chart(fig_pie, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ⚡ PERFORMANS MATRİSİ
    st.markdown("### ⚡ Şirket / Temsilci Performans Matrisi")

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
        gh_df = pd.read_excel(excel_file, sheet_name=genel_sheet)
        for idx, r in gh_df.iterrows():
            row_label = str(r.iloc[0]).strip().lower()
            th = r.iloc[1] if pd.notnull(r.iloc[1]) else 0
            tg = r.iloc[2] if pd.notnull(r.iloc[2]) else 0
            
            if "lead" in row_label:
                m_lead_h, m_lead_g, m_lead_o = f"{int(th):,}", f"{int(tg):,}", f"%{(tg/th*100):.1f}" if th > 0 else "%0.0"
            elif "gelen" in row_label and "rezervasyon" in row_label:
                m_rez_h, m_rez_g, m_rez_o = f"{int(th):,}", f"{int(tg):,}", f"%{(tg/th*100):.1f}" if th > 0 else "%0.0"
            elif "satis" in row_label or "satış" in row_label:
                m_satis_h, m_satis_g, m_satis_o = f"{int(th):,}", f"{int(tg):,}", f"%{(tg/th*100):.1f}" if th > 0 else "%0.0"
            elif "kriter" in row_label:
                m_kriter_h, m_kriter_g = f"%{th:.1f}" if isinstance(th, (int, float)) else str(th), f"%{tg:.1f}"
            elif "gelme" in row_label:
                m_gelme_h, m_gelme_g = f"%{th:.1f}" if isinstance(th, (int, float)) else str(th), f"%{tg:.1f}"

    mc1, mc2, mc3, mc4, mc5 = st.columns(5)
    with mc1: st.markdown(f'<div class="matrix-card"><div class="matrix-title">💎 LEAD</div><div class="matrix-target">Hedef: {m_lead_h}</div><div class="matrix-value">{m_lead_g}</div><div class="matrix-badge">↑ Başarı: {m_lead_o}</div></div>', unsafe_allow_html=True)
    with mc2: st.markdown(f'<div class="matrix-card"><div class="matrix-title">💎 GELEN REZERVASYON</div><div class="matrix-target">Hedef: {m_rez_h}</div><div class="matrix-value">{m_rez_g}</div><div class="matrix-badge">↑ Başarı: {m_rez_o}</div></div>', unsafe_allow_html=True)
    with mc3: st.markdown(f'<div class="matrix-card"><div class="matrix-title">💎 SATIŞ</div><div class="matrix-target">Hedef: {m_satis_h}</div><div class="matrix-value">{m_satis_g}</div><div class="matrix-badge">↑ Başarı: {m_satis_o}</div></div>', unsafe_allow_html=True)
    with mc4: st.markdown(f'<div class="matrix-card"><div class="matrix-title">💎 KRİTER DIŞI</div><div class="matrix-target">Hedef: {m_kriter_h}</div><div class="matrix-value">{m_kriter_g}</div><div class="matrix-badge">↑ Gerçekleşen</div></div>', unsafe_allow_html=True)
    with mc5: st.markdown(f'<div class="matrix-card"><div class="matrix-title">💎 GELME ORANI</div><div class="matrix-target">Hedef: {m_gelme_h}</div><div class="matrix-value">{m_gelme_g}</div><div class="matrix-badge">↑ Gerçekleşen</div></div>', unsafe_allow_html=True)

    st.markdown("---")

    # SEKMELER
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9 = st.tabs([
        "👥 Temsilci Özet Tablosu", "💬 Temsilci Yorumu", "📞 Aksiyon Özet Tablosu", 
        "📍 İl & Marka Analizi", "🎯 Rezervasyon Hedef", "💰 Satış Hedef", 
        "🚶‍♂️ Gelme Oranı Hedef", "🚫 Kriter Dışı Hedef", "📈 Data Analiz"
    ])

    # 1. TEMSİLCİ ÖZET
    with tab1:
        st.subheader("Temsilci Bazlı Performans Özet Tablosu")
        if 'Görevi Alan' in df.columns and 'Görev Durumu' in df.columns:
            temsilci_ozet = df.groupby('Görevi Alan').agg(
                Toplam_Gorev=('Görev ID', 'count'),
                Tamamlanan=('Görev Durumu', lambda x: (x == 'Tamamlandı').sum()),
                Ort_Aksiyon=('Aksiyon Sayısı', 'mean')
            ).reset_index()

            if 'Son Aksiyon' in df.columns:
                rezervasyon_sayilari = df[df['Son Aksiyon'] == 'Rezervasyon yapıldı'].groupby('Görevi Alan').size()
                temsilci_ozet['Rezervasyon_Sayisi'] = temsilci_ozet['Görevi Alan'].map(rezervasyon_sayilari).fillna(0)
                temsilci_ozet['Donusum_Orani_Val'] = (temsilci_ozet['Rezervasyon_Sayisi'] / temsilci_ozet['Toplam_Gorev']) * 100
            else:
                temsilci_ozet['Donusum_Orani_Val'] = (temsilci_ozet['Tamamlanan'] / temsilci_ozet['Toplam_Gorev']) * 100

            temsilci_ozet = temsilci_ozet.sort_values(by='Donusum_Orani_Val', ascending=False).reset_index(drop=True)
            temsilci_ozet['Ort_Aksiyon'] = temsilci_ozet['Ort_Aksiyon'].round(2)
            
            temsilci_ozet_gosterim = temsilci_ozet[['Görevi Alan', 'Toplam_Gorev', 'Tamamlanan', 'Ort_Aksiyon', 'Donusum_Orani_Val']].copy()
            temsilci_ozet_gosterim.columns = ['Temsilci', 'Toplam Görev', 'Tamamlanan Görev', 'Ort. Aksiyon Sayısı', 'Rezervasyon Dönüşüm Oranı (%)']

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

            st.dataframe(styled_temsilci, use_container_width=True, height=(len(temsilci_ozet_gosterim) + 1) * 35 + 38)

            st.markdown("---")
            st.markdown("### 📊 Temsilci Bazlı Rezervasyon Dönüşüm Oranı Grafiği")
            chart_temsilci = temsilci_ozet.copy()
            chart_temsilci['Performans Durumu'] = chart_temsilci['Donusum_Orani_Val'].apply(lambda v: 'Yüksek (>=%6.0)' if v>=6.0 else ('Orta (%3.5-%5.9)' if v>=3.5 else 'Düşük (<%3.5)'))

            fig_temsilci = px.bar(
                chart_temsilci, x='Görevi Alan', y='Donusum_Orani_Val',
                color='Performans Durumu', color_discrete_map={'Yüksek (>=%6.0)': '#00E676', 'Orta (%3.5-%5.9)': '#FFEA00', 'Düşük (<%3.5)': '#FF1744'},
                text=chart_temsilci['Donusum_Orani_Val'].apply(lambda x: f"%{x:.1f}"), title="Temsilcilere Göre Dönüşüm Oranları (%)"
            )
            fig_temsilci.update_layout(template="plotly_dark", xaxis_title="Temsilci", yaxis_title="Dönüşüm Oranı (%)")
            fig_temsilci.update_traces(textposition='outside')
            st.plotly_chart(fig_temsilci, use_container_width=True)

    # 2. TEMSİLCİ YORUMU
    with tab2:
        st.subheader("Temsilci Yorumları ve Performans Analizi")
        ty_sheet = None
        for s in sheet_names:
            name_norm = str(s).strip().lower().replace('ş', 's').replace('ı', 'i').replace('ğ', 'g').replace('ü', 'u').replace('ö', 'o').replace('ç', 'c')
            if "temsilci" in name_norm and "yorum" in name_norm:
                ty_sheet = s
                break

        if ty_sheet is not None:
            ty_df = pd.read_excel(excel_file, sheet_name=ty_sheet)
            ty_df = ty_df[~ty_df.iloc[:, 0].astype(str).isin(haric_personel)].copy()
            if secilen_temsilci != "Tümü":
                ty_df = ty_df[ty_df.iloc[:, 0].astype(str) == secilen_temsilci]

            for idx, row in ty_df.iterrows():
                temsilci_adi = str(row.iloc[0])
                yorum_metni = str(row.iloc[1]) if len(row) > 1 and pd.notnull(row.iloc[1]) else "Yorum bulunamadı."
                with st.expander(f"👤 **{temsilci_adi}** — Performans Yorumu", expanded=True):
                    for emoji in ['✅', '⚠️', '🎯', '🔄', '🚫', '📌']:
                        yorum_metni = yorum_metni.replace(emoji, f"\n{emoji}")
                    satirlar = [s.strip() for s in yorum_metni.split('\n') if s.strip()]
                    for satir in satirlar:
                        st.markdown(satir)

    # 3. AKSİYON ÖZET
    with tab3:
        st.subheader("Son Aksiyon ve Son Arama Dağılım Tablosu")
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            if 'Son Aksiyon' in df.columns:
                st.markdown("**Son Aksiyon Türleri Özeti**")
                aksiyon_ozet = df['Son Aksiyon'].value_counts().reset_index()
                aksiyon_ozet.columns = ['Aksiyon Nedeni', 'Adet']
                aksiyon_ozet['Oran (%)'] = ((aksiyon_ozet['Adet'] / aksiyon_ozet['Adet'].sum()) * 100).map(lambda x: f"%{x:.1f}")
                st.table(aksiyon_ozet)
        with col_t2:
            if 'Son Arama' in df.columns:
                st.markdown("**Son Arama Durumu Özeti**")
                arama_ozet = df['Son Arama'].value_counts().reset_index()
                arama_ozet.columns = ['Arama Durumu', 'Adet']
                arama_ozet['Oran (%)'] = ((arama_ozet['Adet'] / arama_ozet['Adet'].sum()) * 100).map(lambda x: f"%{x:.1f}")
                st.table(arama_ozet)

    # 4. İL & MARKA ANALİZİ
    with tab4:
        st.subheader("Bölgesel ve Marka Bazlı Görev Tablosu")
        if 'İl' in df.columns and 'Marka' in df.columns:
            st.table(pd.crosstab(df['İl'], df['Marka'], margins=True, margins_name="TOPLAM"))

    # Helper function for sheet filtering
    def filter_sheet_df(sheet_keyword):
        target_s = None
        for s in sheet_names:
            name_clean = str(s).strip().lower().replace('ı', 'i').replace('ş', 's').replace('ğ', 'g').replace('ü', 'u').replace('ö', 'o').replace('ç', 'c')
            if sheet_keyword in name_clean:
                target_s = s
                break
        if target_s is not None:
            s_df = pd.read_excel(excel_file, sheet_name=target_s)
            s_df = s_df[~s_df.iloc[:, 0].astype(str).isin(haric_personel)].copy()
            if secilen_temsilci != "Tümü":
                s_df = s_df[s_df.iloc[:, 0].astype(str) == secilen_temsilci]
            return s_df
        return None

    # 5. REZERVASYON HEDEF
    with tab5:
        st.subheader("🎯 Rezervasyon Hedef Tablosu ve Performans Grafiği")
        rez_df = filter_sheet_df("rezervasyon")
        if rez_df is not None and not rez_df.empty:
            oran_col = [c for c in rez_df.columns if "oran" in str(c).lower() or "%" in str(c)]
            if oran_col:
                target_oran_col = oran_col[0]
                raw_values = rez_df[target_oran_col].apply(lambda x: x * 100 if isinstance(x, (int, float)) and x <= 2 else (x if isinstance(x, (int, float)) else 0))
                display_df = rez_df.copy()
                display_df[target_oran_col] = raw_values.apply(lambda x: f"%{x:.1f}")
                styled_df = display_df.style.apply(lambda col: ['color: #00E676; font-weight: bold;' if val>=100 else ('color: #FFEA00; font-weight: bold;' if val>=80 else 'color: #FF1744; font-weight: bold;') for val in raw_values] if col.name == target_oran_col else ['text-align: center;'] * len(col), axis=0).hide(axis="index")
                st.dataframe(styled_df, use_container_width=True, height=(len(display_df) + 1) * 35 + 38)

    # 6. SATIŞ HEDEF
    with tab6:
        st.subheader("💰 Satış Hedef Tablosu ve Performans Grafiği")
        satis_df = filter_sheet_df("satis")
        if satis_df is not None and not satis_df.empty:
            satis_oran_col = [c for c in satis_df.columns if "oran" in str(c).lower() or "%" in str(c)]
            if satis_oran_col:
                target_satis_oran_col = satis_oran_col[0]
                raw_satis_values = satis_df[target_satis_oran_col].apply(lambda x: x * 100 if isinstance(x, (int, float)) and x <= 2 else (x if isinstance(x, (int, float)) else 0))
                display_satis_df = satis_df.copy()
                display_satis_df[target_satis_oran_col] = raw_satis_values.apply(lambda x: f"%{x:.1f}")
                styled_satis_df = display_satis_df.style.apply(lambda col: ['color: #00E676; font-weight: bold;' if val>=50 else ('color: #FFEA00; font-weight: bold;' if val>=35 else 'color: #FF1744; font-weight: bold;') for val in raw_satis_values] if col.name == target_satis_oran_col else ['text-align: center;'] * len(col), axis=0).hide(axis="index")
                st.dataframe(styled_satis_df, use_container_width=True, height=(len(display_satis_df) + 1) * 35 + 38)

    # 7. GELME ORANI HEDEF
    with tab7:
        st.subheader("🚶‍♂️ Gelme Oranı Hedef Tablosu ve Performans Grafiği")
        gelme_df = filter_sheet_df("gelme")
        if gelme_df is not None and not gelme_df.empty:
            if len(gelme_df.columns) > 2:
                gerceklesen_col = gelme_df.columns[2]
                raw_gerceklesen = gelme_df[gerceklesen_col].apply(lambda x: x * 100 if isinstance(x, (int, float)) and x <= 2 else (x if isinstance(x, (int, float)) else 0))
                display_gelme = gelme_df.copy()
                display_gelme[gerceklesen_col] = raw_gerceklesen.apply(lambda x: f"%{x:.1f}")
                styled_gelme = display_gelme.style.apply(lambda col: ['color: #00E676; font-weight: bold;' if val>=40 else 'color: #FF1744; font-weight: bold;' for val in raw_gerceklesen] if col.name == gerceklesen_col else ['text-align: center;'] * len(col), axis=0).hide(axis="index")
                st.dataframe(styled_gelme, use_container_width=True, height=(len(display_gelme) + 1) * 35 + 38)

    # 8. KRİTER DIŞI HEDEF
    with tab8:
        st.subheader("🚫 Kriter Dışı Hedef Tablosu ve Performans Grafiği")
        kriter_df = filter_sheet_df("kriter")
        if kriter_df is not None and not kriter_df.empty:
            if len(kriter_df.columns) > 2:
                k_gerceklesen_col = kriter_df.columns[2]
                raw_kriter = kriter_df[k_gerceklesen_col].apply(lambda x: x * 100 if isinstance(x, (int, float)) and x <= 2 else (x if isinstance(x, (int, float)) else 0))
                display_kriter = kriter_df.copy()
                display_kriter[k_gerceklesen_col] = raw_kriter.apply(lambda x: f"%{x:.1f}")
                styled_kriter = display_kriter.style.apply(lambda col: ['color: #FF1744; font-weight: bold;' if val>20 else 'color: #00E676; font-weight: bold;' for val in raw_kriter] if col.name == k_gerceklesen_col else ['text-align: center;'] * len(col), axis=0).hide(axis="index")
                st.dataframe(styled_kriter, use_container_width=True, height=(len(display_kriter) + 1) * 35 + 38)

    # 9. DATA ANALİZ
    with tab9:
        st.subheader("📈 Data Analiz Tablosu ve Arama Sonuçları Grafiği")
        da_df = filter_sheet_df("data")
        if da_df is not None and not da_df.empty:
            st.dataframe(da_df.style.hide(axis="index"), use_container_width=True, height=(len(da_df) + 1) * 35 + 38)

else:
    st.error("⚠️ Proje klasöründe 'veri.xlsx' bulunamadı.")
