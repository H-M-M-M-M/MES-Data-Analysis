import streamlit as st
import pandas as pd

st.title("🍧MES BI下载后的数据弄成你熟悉的样子( •̀ ω •́ )✧")

uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # 数据预览
    st.write("### 原始数据预览")
    st.dataframe(df.head())

    def get_default_index(col_name):
        return df.columns.tolist().index(col_name) if col_name in df.columns else 0

    sfc_col = st.selectbox("选择 SFC 列-SN", df.columns, index=get_default_index("SFC"))
    desc_col = st.selectbox("选择 MEASURE_NAME - Test Items列", [None] + df.columns.tolist(), index=(get_default_index("MEASURE_NAME") + 1 if "MEASURE_NAME" in df.columns else 0))
    actual_col = st.selectbox("选择 ACTUAL - 值 列", [None] + df.columns.tolist(), index=(get_default_index("ACTUAL") + 1 if "ACTUAL" in df.columns else 0))
    status_col = st.selectbox("选择 MEASURE_STATUS - PASS|FAIL😊 列", [None] + df.columns.tolist(), index=(get_default_index("MEASURE_STATUS") + 1 if "MEASURE_STATUS" in df.columns else 0))
    resource_col = st.selectbox("选择 RESOURCE - Station 列", [None] + df.columns.tolist(), index=(get_default_index("RESOURCE") + 1 if "RESOURCE" in df.columns else 0))
    date_col = st.selectbox("选择 TEST_DATE_TIME 列", [None] + df.columns.tolist(), index=(get_default_index("TEST_DATE_TIME") + 1 if "TEST_DATE_TIME" in df.columns else 0))
    part_number_col = st.selectbox("选择 PART_NUMBER 列", [None] + df.columns.tolist(), index=(get_default_index("PART_NUMBER") + 1 if "PART_NUMBER" in df.columns else 0))
    operation_col = st.selectbox("选择 OPERATION -Test Porcess列", [None] + df.columns.tolist(), index=(get_default_index("OPERATION") + 1 if "OPERATION" in df.columns else 0))
    pn2desc_col = st.selectbox("选择 PN2DESCRIPTION.ktext - 探头型号 列", [None] + df.columns.tolist(), index=(get_default_index("PN2DESCRIPTION.ktext") + 1 if "PN2DESCRIPTION.ktext" in df.columns else 0))

    if not operation_col:
        st.error("请先选择 OPERATION 列")
    else:
        df = df.dropna(subset=[sfc_col])

        selected_columns = [col for col in [sfc_col, desc_col, actual_col, status_col, resource_col, date_col, part_number_col, operation_col, pn2desc_col] if col is not None]
        filtered_df = df[selected_columns] if selected_columns else df[[sfc_col]]

        output_path = "output.xlsx"
        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            for operation, group_df in filtered_df.groupby(operation_col):
                pivot_index = [sfc_col] + ([resource_col] if resource_col else [])
                if desc_col and actual_col:
                    pivot_df = group_df.pivot_table(index=pivot_index, columns=[desc_col] if desc_col else None, values=actual_col, aggfunc="first").reset_index()
                else:
                    pivot_df = group_df[pivot_index]
                
                if status_col:
                    status_df = group_df.groupby(pivot_index)[status_col].apply(lambda x: "FAIL" if "FAIL" in x.values else "PASS").reset_index()
                    pivot_df = pivot_df.merge(status_df, on=pivot_index, how='left')
                
                if date_col:
                    group_df[date_col] = pd.to_datetime(group_df[date_col], errors='coerce')
                    date_df = group_df.groupby(pivot_index)[date_col].min().reset_index()
                    date_df["Date"] = date_df[date_col].dt.date
                    date_df["Time"] = date_df[date_col].dt.time
                    date_df = date_df.drop(columns=[date_col])
                    pivot_df = pivot_df.merge(date_df, on=pivot_index, how='left')
                
                if part_number_col:
                    part_number_df = group_df.groupby(pivot_index)[part_number_col].first().reset_index()
                    pivot_df = pivot_df.merge(part_number_df, on=pivot_index, how='left')
                
                if pn2desc_col:
                    pn2desc_df = group_df.groupby(pivot_index)[pn2desc_col].first().reset_index()
                    pivot_df = pivot_df.merge(pn2desc_df, on=pivot_index, how='left')
                
                pivot_df.to_excel(writer, sheet_name=str(operation)[:31], index=False)  # Excel sheet 名称最多 31 个字符

        with open(output_path, "rb") as file:
            st.download_button("下载处理后的Excel文件", file, file_name="Processed_Data.xlsx")
