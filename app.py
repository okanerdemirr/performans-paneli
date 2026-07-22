import streamlit as st
import pandas as pd
import io
import plotly.express as px

st.set_page_config(page_title="Verimlilik ve Performans Paneli", layout="wide")

# Tüm Tablo Hücrelerini, Başlıklarını ve İçeriklerini Ortalayan + Kaydırma Çubuğunu Engelleyen CSS
st.markdown("""
    <style>
    th, td {
        text-align: center !important;
    }
    div[data-testid="stTable"] table, div[data-testid="stDataFrame"] table {
        width: 100%;
        text-align: center !important;
    }
    div[data-testid="stTable"] th, div[data-testid="stDataFrame"] th {
        text-align: center !important;
    }
    div[data-testid="stTable"] td, div[data-testid="stDataFrame"] td {
        text-align: center !important;
    }
    [data-testid="stDataFrame"] [role="gridcell"] {
        justify-content: center !important;
        text-align: center !important;
    }
    [data-testid="stDataFrame"] [role="columnheader"] {
        justify-content: center !important;
        text-align: center !important;
    }
    div[data-testid="stDataFrame"] > div {
        max-height: none !important;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📊 Performans ve Özet Tablolar Paneli")

# Sol Menü
st.sidebar.header("📁 Veri Kaynağı")
uploaded_file = st.sidebar.file_uploader("Excel veya CSV Dosyanızı Yükleyin", type=["xlsx", "csv", "xls"])

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()

    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(io.BytesIO(file_bytes))
        sheet_names = []
    else:
        excel_file = pd.ExcelFile(io.BytesIO(file_bytes))
        sheet_names = excel_file.sheet_names
        df = pd.read_excel(excel_file, sheet_name=0)

    # HARIÇ TUTULACAK / SILINMIŞ KİŞİLER
    haric_personel = ['CRM Admin', 'Luron AI API', 'Aleyna Daşdemir', 'Zeynep Güzel']
    if 'Görevi Alan' in df.columns:
        df = df[~df['Görevi Alan'].astype(str).isin(haric_personel)].copy()

    # Genel KPI Kartları
    col1, col2, col3 = st.columns(3)
    toplam_gorev = len(df)
    tamamlanan = len(df[df['Görev Durumu'] == 'Tamamlandı']) if 'Görev Durumu' in df.columns else 0
    tamamlanma_orani = (tamamlanan / toplam_gorev * 100) if toplam_gorev > 0 else 0

    col1.metric("Toplam Görev Sayısı", f"{toplam_gorev:,}")
    col2.metric("Tamamlanan Görev", f"{tamamlanan:,}")
    col3.metric("Genel Tamamlanma Oranı", f"%{tamamlanma_orani:.1f}")

    st.markdown("---")

    # SEKMELER
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "👥 Temsilci Özet Tablosu", 
        "💬 Temsilci Yorumu", 
        "📞 Aksiyon Özet Tablosu", 
        "📍 İl & Marka Analizi",
        "🎯 Rezervasyon Hedef",
        "💰 Satış Hedef"
    ])

    # 1. TABLO: TEMSİLCİ PERFORMANSI
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
                if val >= 6.0:
                    return 'color: #00E676; font-weight: bold; text-align: center;'
                elif val >= 3.5:
                    return 'color: #FFEA00; font-weight: bold; text-align: center;'
                else:
                    return 'color: #FF1744; font-weight: bold; text-align: center;'

            styled_temsilci = temsilci_ozet_gosterim.style.apply(
                lambda col: [color_donusum(val) for val in raw_donusum] if col.name == 'Rezervasyon Dönüşüm Oranı (%)' else ['text-align: center;'] * len(col),
                axis=0
            ).hide(axis="index")

            col_config1 = {c: st.column_config.Column(alignment="center") for c in temsilci_ozet_gosterim.columns}
            
            calc_height = (len(temsilci_ozet_gosterim) + 1) * 35 + 38
            st.dataframe(styled_temsilci, use_container_width=True, height=calc_height, column_config=col_config1)

    # 2. TABLO: TEMSİLCİ YORUMU
    with tab2:
        st.subheader("Temsilci Yorumları ve Performans Analizi")
        if 'Görevi Alan' in df.columns:
            grup = df.groupby('Görevi Alan')

            for temsilci, veri in grup:
                toplam = len(veri)
                tamamlanan_sayisi = (veri['Görev Durumu'] == 'Tamamlandı').sum() if 'Görev Durumu' in veri.columns else 0
                
                if 'Son Aksiyon' in veri.columns:
                    rezervasyon_sayisi = (veri['Son Aksiyon'] == 'Rezervasyon yapıldı').sum()
                else:
                    rezervasyon_sayisi = tamamlanan_sayisi
                
                rezervasyon_orani = (rezervasyon_sayisi / toplam * 100) if toplam > 0 else 0
                
                guclu_yon = f"Satış kapatma rakamları ({tamamlanan_sayisi}) çok başarılı, ikna kabiliyeti yüksek."
                gelisim_alani = "Mevcut metrikleri dengeli, mevcut satış oranlarını artırmaya odaklanılabilir."
                
                if rezervasyon_orani >= 50:
                    donusum_yorumu = f"%{rezervasyon_orani:.1f} - Çok başarılı, müşteri eksperti çekmede harika iş çıkarıyor."
                elif rezervasyon_orani >= 20:
                    donusum_yorumu = f"%{rezervasyon_orani:.1f} - Ortalama seviyede, itiraz karşılama taktikleriyle daha da artırılabilir."
                else:
                    donusum_yorumu = f"%{rezervasyon_orani:.1f} - Düşük kalmış, randevu oluşturma ve ikna argümanları güçlendirilmeli."

                with st.expander(f"👤 **{temsilci}** — Performans Özeti ve Yorumu", expanded=True):
                    c1, c2 = st.columns([1, 4])
                    with c1:
                        st.write(f"**Toplam Görev:** {toplam}")
                        st.write(f"**Tamamlanan:** {tamamlanan_sayisi}")
                        st.write(f"**Rezervasyon Sayısı:** {rezervasyon_sayisi}")
                        st.write(f"**Dönüşüm Oranı:** %{rezervasyon_orani:.1f}")
                    with c2:
                        st.markdown(f"✅ **Güçlü Yönü:** {guclu_yon}")
                        st.markdown(f"⚠️ **Gelişim Alanı:** {gelisim_alani}")
                        st.markdown(f"🎯 **Rezervasyon Dönüşümü:** {donusum_yorumu}")

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

    # 5. TABLO & GRAFİK: RENKLENDİRİLMİŞ REZERVASYON HEDEF
    with tab5:
        st.subheader("🎯 Rezervasyon Hedef Tablosu ve Performans Grafiği")
        
        if sheet_names:
            target_sheet = None
            for s in sheet_names:
                name_clean = str(s).strip().lower().replace('ı', 'i').replace('ş', 's')
                if "rezervasyon" in name_clean and "hedef" in name_clean:
                    target_sheet = s
                    break

            if target_sheet is not None:
                rez_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=target_sheet)
                
                if len(rez_df.columns) > 0:
                    rez_df = rez_df[~rez_df.iloc[:, 0].astype(str).isin(haric_personel)].copy()

                # Sayısal sütunların küsüratlarını temizleme (Tam sayıya çevirme)
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
                        if val >= 100:
                            return 'Yüksek (>=%100)'
                        elif val >= 80:
                            return 'Orta (%80-%99)'
                        else:
                            return 'Düşük (<%80)'

                    def color_cell(val):
                        if val >= 100:
                            return 'color: #00E676; font-weight: bold; text-align: center;'
                        elif val >= 80:
                            return 'color: #FFEA00; font-weight: bold; text-align: center;'
                        else:
                            return 'color: #FF1744; font-weight: bold; text-align: center;'

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

                    color_discrete_map = {
                        'Yüksek (>=%100)': '#00E676',
                        'Orta (%80-%99)': '#FFEA00',
                        'Düşük (<%80)': '#FF1744'
                    }

                    y_col = chart_df.columns[3] if len(chart_df.columns) > 3 else chart_df.columns[1]

                    fig = px.bar(
                        chart_df,
                        x=chart_df.columns[0],
                        y=y_col,
                        color='Performans Durumu',
                        color_discrete_map=color_discrete_map,
                        text=y_col,
                        title="Temsilcilere Göre Gerçekleşen Rezervasyon ve Performans Kategorisi"
                    )

                    fig.update_layout(
                        template="plotly_dark",
                        xaxis_title="Temsilci",
                        yaxis_title="Gerçekleşen Adet",
                        legend_title="Performans Durumu",
                        font=dict(size=12)
                    )
                    fig.update_traces(textposition='outside')

                    st.plotly_chart(fig, use_container_width=True)

                else:
                    st.table(rez_df)
            else:
                st.warning("⚠️ Excel dosyanızda 'Rezervasyon Hedef' sekmesi bulunamadı.")
        else:
            st.info("CSV dosyalarında birden fazla sekme bulunmaz. Lütfen Excel (.xlsx) yükleyin.")

    # 6. TABLO & GRAFİK: SATIŞ HEDEF SEKME
    with tab6:
        st.subheader("💰 Satış Hedef Tablosu ve Performans Grafiği")
        
        if sheet_names:
            satis_sheet = None
            for s in sheet_names:
                name_norm = str(s).strip().lower().replace('ş', 's').replace('ı', 'i').replace('ğ', 'g').replace('ü', 'u').replace('ö', 'o').replace('ç', 'c')
                if "satis" in name_norm and "hedef" in name_norm:
                    satis_sheet = s
                    break

            if satis_sheet is not None:
                satis_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=satis_sheet)
                
                if len(satis_df.columns) > 0:
                    satis_df = satis_df[~satis_df.iloc[:, 0].astype(str).isin(haric_personel)].copy()

                # Sayısal sütunların küsüratlarını temizleme (Tam sayıya çevirme)
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
                        if val >= 50:
                            return 'Yüksek (>=%50)'
                        elif val >= 35:
                            return 'Orta (%35-%49)'
                        else:
                            return 'Düşük (<%35)'

                    def color_satis_cell(val):
                        if val >= 50:
                            return 'color: #00E676; font-weight: bold; text-align: center;'
                        elif val >= 35:
                            return 'color: #FFEA00; font-weight: bold; text-align: center;'
                        else:
                            return 'color: #FF1744; font-weight: bold; text-align: center;'

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

                    color_satis_map = {
                        'Yüksek (>=%50)': '#00E676',
                        'Orta (%35-%49)': '#FFEA00',
                        'Düşük (<%35)': '#FF1744'
                    }

                    y_satis_col = chart_satis_df.columns[2] if len(chart_satis_df.columns) > 2 else chart_satis_df.columns[1]

                    fig_satis = px.bar(
                        chart_satis_df,
                        x=chart_satis_df.columns[0],
                        y=y_satis_col,
                        color='Satış Durumu',
                        color_discrete_map=color_satis_map,
                        text=y_satis_col,
                        title="Temsilcilere Göre Gerçekleşen Satış Adetleri ve Performans Kategorisi"
                    )

                    fig_satis.update_layout(
                        template="plotly_dark",
                        xaxis_title="Temsilci",
                        yaxis_title="Gerçekleşen Satış Adedi",
                        legend_title="Satış Durumu",
                        font=dict(size=12)
                    )
                    fig_satis.update_traces(textposition='outside')

                    st.plotly_chart(fig_satis, use_container_width=True)

                else:
                    st.table(satis_df)
            else:
                st.warning("⚠️ Excel dosyanızda 'Satış Hedef' sekmesi bulunamadı.")
        else:
            st.info("CSV dosyalarında birden fazla sekme bulunmaz. Lütfen Excel (.xlsx) yükleyin.")

else:
    st.info("👈 Başlamak için sol menüden Excel dosyanızı yükleyin.")
