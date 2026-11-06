# 求职材料（私有）

针对「平价现制饮品 · 市场分析」岗位准备的一整套求职材料，以及生成它们的 Python 脚本。

> 仓库中的简历 PDF/Word 含手机号、邮箱等个人信息，投递前请自行确认可公开的范围。

## 交付物

| 文件 | 说明 |
|---|---|
| `简历-市场分析-幸运咖.docx / .pdf` | 简历成品（Word 可编辑版 + PDF 投递版） |
| `平价现制饮品市场分析报告-幸运咖.pptx / .pdf` | 市场分析报告（PPT + PDF） |
| `面试自我介绍-幸运咖市场分析.pptx / .pdf` | 面试自我介绍演示稿 |
| `charts/` | 报告与简历中使用的图表 |
| `png_*` | 导出为图片的成品页 |

## 生成脚本

| 脚本 | 作用 |
|---|---|
| `make_resume.py` | 用 python-docx 生成简历 |
| `make_charts.py` | 用 matplotlib 生成分析图表 |
| `make_report_ppt.py` | 用 python-pptx 生成市场分析报告 |
| `make_intro_ppt.py` | 生成面试自我介绍 PPT |
| `convert_pdf.ps1` | 调用 Office 批量转 PDF |
| `export_png.ps1` | 把 PPT/PDF 页导出为 PNG |

## 重新生成

```bash
pip install python-docx python-pptx matplotlib pandas openpyxl
python make_charts.py
python make_resume.py
python make_report_ppt.py
python make_intro_ppt.py
```

PDF 转换依赖本机安装的 Microsoft Office：

```powershell
powershell -ExecutionPolicy Bypass -File convert_pdf.ps1
```

## 说明

数据与分析结论为面试准备用途。
