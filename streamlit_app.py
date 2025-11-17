import tempfile
import os

import streamlit as st

from size_style_core import extract_from_path, to_dataframe, category_rank


def main():
    st.title("Style–Size Extractor")

    st.write("Dosyayı yükle, sonuçları tabloda ve özet olarak gör")

    uploaded_file = st.file_uploader("PDF yükle",type=["pdf", "txt"], )
    if uploaded_file is not None:
        suffix = "." + uploaded_file.name.split(".")[-1].lower()
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(uploaded_file.read())
            tmp_path = tmp.name

        try:
            result = extract_from_path(tmp_path)

        finally:
           
            try:
                os.remove(tmp_path)
            except OSError:
                pass

        df = to_dataframe(result.agg)

        # --- SUMMARY ---
        st.subheader("Özet")

        total_items = int(sum(result.agg.values()))

        sweatshirt_total = result.sweatshirt_nonhoodie_total
        hoodie_total = result.hoodie_total
        vneck_total = 0
        longsleeve_total = 0
        adult_tee_total = 0
        youth_total = 0
        toddler_total = 0
        onesie_total = 0
        apron_total = 0
        tote_total = 0

        for style, q in result.agg.items():
            cat = category_rank(style)
            s_lower = style.lower()

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

        st.markdown(
            f"""
**Total unique combos:** {result.unique_count}  
**Total items:** {total_items}  

**By Category**  
- Sweatshirts (non-hoodie): **{sweatshirt_total}**  
- Hoodies: **{hoodie_total}**  
- Adult Short Sleeve / Tees: **{adult_tee_total}**  
- V-Neck: **{vneck_total}**  
- Long Sleeve: **{longsleeve_total}**  
- Youth (all): **{youth_total}**  
- Toddler (all): **{toddler_total}**  
- Onesie / Baby Bodysuit: **{onesie_total}**  
- Apron: **{apron_total}**  
- Tote Bag: **{tote_total}**  
"""
        )

        # --- TABLE ---
        st.subheader("Detaylı tablo (Style / Size)")

        st.dataframe(df)

        # --- Download buttons ---
        st.subheader("Dışa aktarım")

        csv_data = df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="CSV indir",
            data=csv_data,
            file_name="style_size_summary.csv",
            mime="text/csv",
    


if __name__ == "__main__":
    main()
