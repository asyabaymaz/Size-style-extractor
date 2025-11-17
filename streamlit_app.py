import io
import tempfile

import streamlit as st
import pandas as pd

from size_style_core import extract_from_path, to_dataframe, category_rank  

def summarize_by_category(agg: dict) -> dict:
    total_items = int(sum(agg.values()))

    sweatshirt_total = 0
    hoodie_total = 0
    vneck_total = 0
    longsleeve_total = 0
    adult_tee_total = 0
    youth_total = 0
    toddler_total = 0
    onesie_total = 0
    apron_total = 0
    tote_total = 0

    for style, q in agg.items():
        cat = category_rank(style)
        s_lower = style.lower()

    
        if "sweatshirt" in s_lower and "hoodie" not in s_lower and "hooded" not in s_lower:
            sweatshirt_total += q
        if "hoodie" in s_lower or "hooded sweatshirt" in s_lower or "unisex hoodie" in s_lower:
            hoodie_total += q

        if cat == 1:
            longsleeve_total += q
        elif cat == 3:
            adult_tee_total += q
        elif cat in (4, 5):
            youth_total += q
        elif cat in (6, 7):
            toddler_total += q
        elif cat == 8:
            onesie_total += q
        elif cat == 10:
            vneck_total += q

        if "apron" in s_lower:
            apron_total += q
        if "tote" in s_lower:
            tote_total += q

    return {
        "total_items": total_items,
        "sweatshirts": sweatshirt_total,
        "hoodies": hoodie_total,
        "adult_tees": adult_tee_total,
        "vnecks": vneck_total,
        "longsleeves": longsleeve_total,
        "youth": youth_total,
        "toddler": toddler_total,
        "onesie": onesie_total,
        "apron": apron_total,
        "tote": tote_total,
    }


def main():
    st.set_page_config(page_title="Size-Style Extractor", layout="wide")
    st.title("Size-Style Extractor")

    st.write("PDF veya TXT yükle, tüm Style / Size adetlerini saydır.")


    uploaded_file = st.file_uploader("UPLOAD PDF (PDF veya .txt)", type=["pdf", "txt"])

    if uploaded_file is not None:
        # Uploaded file'ı geçici dosyaya yazıp extract_from_path'e path veriyoruz
        suffix = ".pdf" if uploaded_file.type == "application/pdf" else ".txt"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(uploaded_file.getbuffer())
            tmp_path = tmp.name

        with st.spinner("Çözümleniyor..."):
            result = extract_from_path(tmp_path, max_blank=int(max_blank))
            df = to_dataframe(result.agg)

        # --- SUMMARY ---
        st.subheader("Özet")
        cats = summarize_by_category(result.agg)

        col1, col2 = st.columns(2)
        with col1:
            st.write(f"**Total unique combos:** {result.unique_count}")
            st.write(f"**Total items:** {cats['total_items']}")

        with col2:
            st.write("**By Category:**")
            st.write(f"- Sweatshirts (non-hoodie): {cats['sweatshirts']}")
            st.write(f"- Hoodies: {cats['hoodies']}")
            st.write(f"- Adult Short Sleeve / Tees: {cats['adult_tees']}")
            st.write(f"- V-Neck: {cats['vnecks']}")
            st.write(f"- Long Sleeve: {cats['longsleeves']}")
            st.write(f"- Youth (all): {cats['youth']}")
            st.write(f"- Toddler (all): {cats['toddler']}")
            st.write(f"- Onesie / Baby Bodysuit: {cats['onesie']}")
            st.write(f"- Apron: {cats['apron']}")
            st.write(f"- Tote Bag: {cats['tote']}")

        # --- DETAILS TABLE ---
        st.subheader("Detaylar (Style / Size)")
        st.dataframe(df, use_container_width=True)

        # --- Download buttons ---
        csv = df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Download CSV",
            data=csv,
            file_name="style_size_summary.csv",
            mime="text/csv",
        )

        # XLSX için in-memory buffer
        xlsx_buffer = io.BytesIO()
        with pd.ExcelWriter(xlsx_buffer, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="All", index=False)
        xlsx_buffer.seek(0)

        st.download_button(
            "Download XLSX",
            data=xlsx_buffer,
            file_name="style_size_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()




