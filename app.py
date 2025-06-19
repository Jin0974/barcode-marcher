import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz
import io

st.set_page_config(layout="wide")

st.title('商品名称与条码智能匹配工具')

with st.expander('使用说明', expanded=False):
    st.markdown('''
    1. 上传“准备的表格”（只含商品名称）
    2. 上传“商品明细表”（含商品名称和条码）
    3. 系统自动匹配，可手动修正
    4. 导出结果为Excel
    ''')

col1, col2 = st.columns(2)
with col1:
    file_names = st.file_uploader('上传准备的表格（商品名称）', type=['xls', 'xlsx'], key='names')
with col2:
    file_detail = st.file_uploader('上传商品明细表（商品名称+条码）', type=['xls', 'xlsx'], key='detail')

if file_names and file_detail:
    df_names = pd.read_excel(file_names)
    df_detail = pd.read_excel(file_detail)
    name_col = '商品名称'
    barcode_col = '条码'

    def option_str(row):
        return f"{row[barcode_col]} | {row[name_col]}"
    df_detail['option'] = df_detail.apply(option_str, axis=1)

    match_rows = []
    for i, name in enumerate(df_names[name_col]):
        matches = process.extract(
            name,
            df_detail[name_col],
            scorer=fuzz.token_sort_ratio,
            limit=20
        )
        filtered = [(m[0], m[1]) for m in matches if m[1] >= 20]
        if not filtered:
            filtered = [(matches[0][0], matches[0][1])]  # 至少保留一个
        # 按相似度降序排列
        filtered = sorted(filtered, key=lambda x: -x[1])
        options = []
        for m_name, score in filtered:
            # 兼容明细表中商品名称重复的情况
            rows = df_detail[df_detail[name_col] == m_name]
            for _, row in rows.iterrows():
                options.append(option_str(row))
        # 去重，保持顺序
        seen = set()
        options = [x for x in options if not (x in seen or seen.add(x))]
        default_idx = 0
        selected = st.selectbox(
            f"{name}",
            options,
            index=default_idx,
            key=f'sel_{i}'
        )
        # 拆分条码和商品名称
        barcode, match_name = selected.split(' | ', 1)
        match_rows.append({
            '原商品名称': name,
            '条码': barcode,
            '匹配商品名称': match_name
        })

    export_df = pd.DataFrame(match_rows)
    st.dataframe(export_df, use_container_width=True)

    towrite = io.BytesIO()
    export_df.to_excel(towrite, index=False, engine='openpyxl')
    st.download_button(
        label='导出匹配结果Excel',
        data=towrite.getvalue(),
        file_name='匹配结果.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
else:
    st.info('请上传“准备的表格（商品名称）”和“商品明细表（商品名称+条码）”，系统会自动进行匹配。')
