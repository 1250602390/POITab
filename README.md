# 百度地图 Bug 数据看板

基于 Python + Tkinter 构建的桌面数据看板，用于可视化分析百度地图 POI 标签分类及 Bug 数据。

## 功能特性

- 导入 Excel 数据文件（`.xlsx`）
- POI 标签分类统计与展示
- Bug 数据多维度看板（超时筛选、状态分布等）
- 高端深蓝渐变 UI 风格

## 环境要求

- Python 3.8+
- 依赖见 `requirements.txt`

## 快速开始

```bash
pip install -r requirements.txt
python map_dashboard.py
```

## 打包为可执行文件

```bash
pyinstaller POI标签分类.spec
```

打包产物位于 `dist/` 目录。

## 许可证

本项目基于 [MIT License](LICENSE) 开源。
