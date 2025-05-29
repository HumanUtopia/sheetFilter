import streamlit as st
import pandas as pd
from io import BytesIO

def fix_header(header_row):
    new_header = []
    used = {}
    for i, col in enumerate(header_row):
        name = str(col) if pd.notna(col) and str(col).strip() != '' and str(col).lower() != 'nan' else f"Unnamed: {i+1}"
        if name in used:
            used[name] += 1
            name = f"{name}_{used[name]}"
        else:
            used[name] = 1
        new_header.append(name)
    return new_header

def check_brackets(filters):
    # 括号完全配对返回True，否则False
    bracket_count = 0
    for f in filters:
        prefix = f.get('prefix', '')
        suffix = f.get('suffix', '')
        bracket_count += prefix.count('(')
        bracket_count -= suffix.count(')')
    return bracket_count == 0

def build_logic_expression(df, filters):
    # 括号配对检测
    bracket_count = 0
    for f in filters:
        prefix = f.get('prefix', '')
        suffix = f.get('suffix', '')
        bracket_count += prefix.count('(')
        bracket_count -= suffix.count(')')
    if bracket_count != 0:
        st.error("括号不配对，请调整括号数使左右括号匹配！")
        return pd.Series([False] * len(df))

    expressions = []
    for i, f in enumerate(filters):
        col_data = df[f['col']].astype(str).fillna("")
        mask = col_data.str.contains(f['keyword'], case=False, na=False)
        if not f['include']:
            mask = ~mask
        exp = f"{f.get('prefix', '')}mask_{i}{f.get('suffix', '')}"
        expressions.append((exp, mask))

    expr_str = expressions[0][0]
    for i in range(1, len(filters)):
        logic_op = '&' if filters[i]['logic'].upper() == 'AND' else '|'
        expr_str += f" {logic_op} {expressions[i][0]}"

    context = {f"mask_{i}": m for i, (_, m) in enumerate(expressions)}

    try:
        result = eval(expr_str, {}, context)
    except Exception as e:
        st.error(f"逻辑表达式错误：{e}")
        return pd.Series([False] * len(df))

    return result

def apply_advanced_filter(df, filters):
    if not filters:
        return df
    result_mask = build_logic_expression(df, filters)
    return df[result_mask]

def build_readable_logic(filters):
    if not filters:
        return ""
    expressions = []
    for i, f in enumerate(filters):
        col = f['col']
        keyword = f['keyword']
        include = f['include']
        prefix = f.get('prefix', '')
        suffix = f.get('suffix', '')
        if include:
            exp = f"{prefix}{col} like '*{keyword}*'{suffix}"
        else:
            exp = f"{prefix}{col} not like '*{keyword}*'{suffix}"
        expressions.append(exp)
    logic_exp = expressions[0]
    for i in range(1, len(filters)):
        logic = filters[i]['logic']
        logic_exp += f" {logic} {expressions[i]}"
    return logic_exp

st.title("Excel 筛选工具")

uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx"])
if uploaded_file:
    raw_df = pd.read_excel(uploaded_file, header=None)
    st.subheader("原始数据预览")
    st.dataframe(raw_df.head(10))

    header_mode = st.radio("表头模式", options=["单行表头", "两行合并表头"])

    if header_mode == "单行表头":
        header_row = st.number_input("表头所在行", min_value=0, max_value=len(raw_df) - 1, value=0)
        df = pd.read_excel(uploaded_file, header=None, skiprows=header_row + 1)
        fixed_header = fix_header(raw_df.iloc[header_row])
        df.columns = fixed_header
    else:
        header_row1 = st.number_input("上层表头行", min_value=0, max_value=len(raw_df) - 1, value=0)
        header_row2 = st.number_input("下层表头行（优先使用）", min_value=0, max_value=len(raw_df) - 1, value=1)
        h1 = raw_df.iloc[header_row1]
        h2 = raw_df.iloc[header_row2]
        mixed_header = [
            str(h2[i]) if pd.notna(h2[i]) and str(h2[i]).strip() != '' and str(h2[i]).lower() != 'nan'
            else str(h1[i]) if pd.notna(h1[i]) and str(h1[i]).strip() != '' and str(h1[i]).lower() != 'nan'
            else f"Unnamed: {i+1}"
            for i in range(len(h2))
        ]
        fixed_header = fix_header(mixed_header)
        df = pd.read_excel(uploaded_file, header=None, skiprows=max(header_row1, header_row2) + 1)
        df.columns = fixed_header

    st.success("✅ 当前使用的字段名：")
    st.write(list(df.columns))
    st.dataframe(df.head())

    # 初始化逻辑状态
    if 'filters' not in st.session_state:
        st.session_state.filters = []

    columns = df.columns.tolist()

    st.markdown("### ➕ 添加过滤条件")
    with st.form("add_filter"):
        col = st.selectbox("选择字段", columns, key="new_col")
        keyword = st.text_input("关键词（模糊匹配）", key="new_kw")
        include = st.checkbox("包含该关键词（取消为排除）", value=True, key="new_include")
        logic = st.selectbox("逻辑类型", ["AND", "OR"], key="new_logic")
        prefix = st.text_input("前缀括号（如 ( 或 ((，可留空）", value="", key="new_prefix")
        suffix = st.text_input("后缀括号（如 ) 或 ))，可留空）", value="", key="new_suffix")
        add_btn = st.form_submit_button("添加条件")

        if add_btn and keyword:
            st.session_state.filters.append({
                'col': col,
                'keyword': keyword,
                'include': include,
                'logic': logic,
                'prefix': prefix.strip(),
                'suffix': suffix.strip()
            })

    st.markdown("### ✅ 当前构建的筛选条件")
    if st.session_state.filters:
        for i, f in enumerate(st.session_state.filters):
            prefix = f.get('prefix', '')
            suffix = f.get('suffix', '')
            logic = f['logic'] if i != 0 else ""
            col = f['col']
            include = f['include']
            keyword = f['keyword']

            # 6列：左括号+、左括号-、逻辑+条件、右括号+、右括号-、删除
            cols = st.columns([0.055, 0.055, 0.69, 0.055, 0.055, 0.09])
            # 左括号 +
            with cols[0]:
                if st.button("＋(", key=f"add_prefix_{i}"):
                    st.session_state.filters[i]['prefix'] = prefix + '('
                    st.rerun()
            # 左括号 -
            with cols[1]:
                if prefix:
                    if st.button("－(", key=f"remove_prefix_{i}"):
                        st.session_state.filters[i]['prefix'] = prefix[:-1] if prefix.endswith('(') else prefix
                        st.rerun()
                else:
                    st.markdown("&nbsp;", unsafe_allow_html=True)  # 占位
            # 逻辑+条件描述
            with cols[2]:
                desc = (
                        f"**{i + 1}.** "
                        + (f"[{logic}] " if logic else "")
                        + f"{prefix}"
                        + f"`{col}` "
                        + ("包含" if include else "不包含")
                        + f" **{keyword}**"
                        + f"{suffix}"
                )
                st.markdown(desc, unsafe_allow_html=True)
            # 右括号 +
            with cols[3]:
                if st.button("＋)", key=f"add_suffix_{i}"):
                    st.session_state.filters[i]['suffix'] = suffix + ')'
                    st.rerun()
            # 右括号 -
            with cols[4]:
                if suffix:
                    if st.button("－)", key=f"remove_suffix_{i}"):
                        st.session_state.filters[i]['suffix'] = suffix[:-1] if suffix.endswith(')') else suffix
                        st.rerun()
                else:
                    st.markdown("&nbsp;", unsafe_allow_html=True)
            # 删除按钮
            with cols[5]:
                if st.button("🗑", key=f"del_{i}"):
                    st.session_state.filters.pop(i)
                    st.rerun()
    else:
        st.info("请添加筛选条件")

    # 括号配对提示
    if not check_brackets(st.session_state.filters):
        st.warning("当前筛选条件的括号不配对，请调整括号数使左右括号匹配，否则无法执行筛选！")

    # 可读逻辑表达式
    if st.session_state.filters:
        st.markdown("#### 🧩 匹配逻辑")
        st.code(build_readable_logic(st.session_state.filters), language="sql")

    # 禁用筛选按钮（括号不配对时禁用）
    filter_btn_disabled = not check_brackets(st.session_state.filters)
    if st.button("执行筛选", disabled=filter_btn_disabled):
        result_df = apply_advanced_filter(df, st.session_state.filters)
        st.success(f"筛选后剩余 {len(result_df)} 条记录")
        st.dataframe(result_df)

        # 导出 Excel 文件
        output = BytesIO()
        result_df.to_excel(output, index=False, engine='openpyxl')
        st.download_button(
            label="📥 下载结果 Excel",
            data=output.getvalue(),
            file_name="筛选结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )