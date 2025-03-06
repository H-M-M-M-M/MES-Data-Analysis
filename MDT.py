import streamlit as st
import pandas as pd

st.title("🍧MES BI下载后的数据弄成你熟悉的样子( •̀ ω •́ )✧")

uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # 数据预览
    st.write("### 原始数据预览")
    st.dataframe(df.head())

    # 显示列名供用户选择
    #st.write("### 可用列名")
    #st.write(df.columns)

    # 让用户自定义列映射
    def get_default_index(col_name):
        return df.columns.tolist().index(col_name) if col_name in df.columns else 0

    sfc_col = st.selectbox("选择 SFC 列-SN", df.columns, index=get_default_index("SFC"))
    desc_col = st.selectbox("选择 MEASURE_NAME - Test Items列", [None] + df.columns.tolist(), index=(get_default_index("MEASURE_NAME") + 1 if "MEASURE_NAME" in df.columns else 0))
    actual_col = st.selectbox("选择 ACTUAL - 值 列", [None] + df.columns.tolist(), index=(get_default_index("ACTUAL") + 1 if "ACTUAL" in df.columns else 0))
    status_col = st.selectbox("选择 MEASURE_STATUS - PASS|FAIL😊 列", [None] + df.columns.tolist(), index=(get_default_index("MEASURE_STATUS") + 1 if "MEASURE_STATUS" in df.columns else 0))
    resource_col = st.selectbox("选择 RESOURCE - Station 列", [None] + df.columns.tolist(), index=(get_default_index("RESOURCE") + 1 if "RESOURCE" in df.columns else 0))
    date_col = st.selectbox("选择 TEST_DATE_TIME 列", [None] + df.columns.tolist(), index=(get_default_index("TEST_DATE_TIME") + 1 if "TEST_DATE_TIME" in df.columns else 0))
    part_number_col = st.selectbox("选择 PART_NUMBER 列", [None] + df.columns.tolist(), index=(get_default_index("PART_NUMBER") + 1 if "PART_NUMBER" in df.columns else 0))
    pn2desc_col = st.selectbox("选择 PN2MEASURE_NAME.ktext - 探头型号 列", [None] + df.columns.tolist(), index=(get_default_index("PN2MEASURE_NAME.ktext") + 1 if "PN2MEASURE_NAME.ktext" in df.columns else 0))

    # 仅保留 SFC 列非空的数据
    df = df.dropna(subset=[sfc_col])
    
    # 筛选列
    selected_columns = [col for col in [sfc_col, desc_col, actual_col, status_col, resource_col, date_col, part_number_col, pn2desc_col] if col is not None]
    filtered_df = df[selected_columns] if selected_columns else df[[sfc_col]]
    
    # 数据透视表，以 SFC 和 RESOURCE 作为索引
    if desc_col and actual_col:
        pivot_df = filtered_df.pivot_table(index=[sfc_col, resource_col] if resource_col else [sfc_col], columns=desc_col, values=actual_col, aggfunc="first").reset_index()
    else:
        pivot_df = df[[sfc_col, resource_col]] if resource_col else df[[sfc_col]]
    
    # 计算 MEASURE_STATUS
    if status_col:
        status_df = df.groupby([sfc_col, resource_col] if resource_col else [sfc_col])[status_col].apply(lambda x: "FAIL" if "FAIL" in x.values else "PASS").reset_index()
        
        # 查找 SFC 下的 FAIL 项，并提取所有 FAIL 子项
        fail_items_df = df[df[status_col] == "FAIL"]
        fail_items_df = fail_items_df.groupby([sfc_col, resource_col])[desc_col].apply(lambda x: ", ".join(x.unique())).reset_index(name="Fail_Items")
    else:
        status_df = pd.DataFrame()
        fail_items_df = pd.DataFrame()
    
    # 处理 TEST_DATE_TIME，拆分为 Date 和 Time 两列
    if date_col:
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')  # 转换为 datetime 类型，并处理无效日期
    #拆分    
        date_df = df.groupby([sfc_col, resource_col] if resource_col else [sfc_col])[date_col].min().reset_index()
        date_df["Date"] = date_df[date_col].dt.date
        date_df["Time"] = date_df[date_col].dt.time
        date_df = date_df.drop(columns=[date_col])
    else:
        date_df = pd.DataFrame()
    
    # 处理 PART_NUMBER
    if part_number_col:
        part_number_df = df.groupby([sfc_col, resource_col] if resource_col else [sfc_col])[part_number_col].first().reset_index()
    else:
        part_number_df = pd.DataFrame()
    
    # 处理 PN2MEASURE_NAME.ktext
    if pn2desc_col:
        pn2desc_df = df.groupby([sfc_col, resource_col] if resource_col else [sfc_col])[pn2desc_col].first().reset_index()
    else:
        pn2desc_df = pd.DataFrame()
    
    # 合并数据
    final_df = pivot_df
    for df_part in [status_df, date_df, part_number_df, pn2desc_df, fail_items_df]:
        if not df_part.empty:
            final_df = final_df.merge(df_part, on=[sfc_col, resource_col] if resource_col else [sfc_col], how='left')

    st.write("### 处理后数据预览")
    st.dataframe(final_df.head())

    # 导出按钮
    if st.button("导出数据"):
        output_path = "output.xlsx"
        final_df.to_excel(output_path, index=False)
        with open(output_path, "rb") as file:
            st.download_button("下载 Excel 文件", file, file_name="Processed_Data.xlsx")
