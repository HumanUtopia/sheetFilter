# Streamlit Excel 智能筛选工具

这是一个基于 [Streamlit](https://streamlit.io/) 的 Excel 智能筛选工具，支持多条件组合、括号嵌套和可视化编辑，帮助你轻松筛选和导出 Excel 数据。

## ✨ 功能特性

- 支持自定义 Excel 表头（单行或多行）
- 多条件筛选，支持“包含/不包含”逻辑
- 支持 AND/OR 逻辑及括号嵌套，实现复杂筛选
- 直观的筛选条件可视化编辑
- 自动检查括号配对，防止语法错误
- 筛选结果可直接导出为新的 Excel 文件

## 🚀 快速开始

1. 安装依赖

   ```bash
   pip install streamlit pandas openpyxl
   ```

2. 运行应用

   ```bash
   streamlit run main.py
   ```

3. 打开浏览器访问 [http://localhost:8501](http://localhost:8501)

## 🖼️ 使用方法

1. 上传你的 Excel 文件（支持 `.xlsx` 格式）
2. 选择表头模式（单行或两行合并）
3. 添加筛选条件，可设置逻辑类型（AND/OR）、括号、包含/不包含等
4. 可视化编辑筛选条件，括号增减按钮和删除按钮均支持一键操作
5. 检查下方的人可读逻辑表达式，确认筛选逻辑无误
6. 点击“执行筛选”，在页面预览并下载结果 Excel

## 📷 示例界面

![界面示例](./screenshot.png)

## 📝 备注

- 本工具仅在本地运行，不上传任何数据，请放心使用。
- 如果遇到问题或有功能建议，欢迎提 Issue！

## 📄 License

MIT License