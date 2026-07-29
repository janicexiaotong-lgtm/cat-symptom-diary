# 猫咪症状日记 + 长期病例可视化汇总

在原有的「猫咪症状日记」(cat-symptom-diary) 基础上，新增 **「病例汇总」** 入口，
点击后跳转到 `case-summary.html` —— 由 `猫咪病例数据2023-2026.xlsx` 生成的历年趋势可视化
（肌酐/尿素氮/血磷/血钙/HCT/离子钙/血钾 + 肾脏超声，含 IRIS 分期色带与临床事件时间线）。

## 目录结构
```
index.html                  猫咪症状日记（已加入「病例汇总」Tab 与概览卡片）
case-summary.html           自动生成的可视化病例（请勿手改，由脚本生成）
build_case_summary.py       生成 case-summary.html 的脚本
case-data/
  cat-case-2023-2026.xlsx  病例源数据（两级表头病历 Excel）
.github/workflows/
  build-case-summary.yml    GitHub Actions：数据更新后自动重建并部署
excel-to-json.html          原有：小程序 Excel → 日记 JSON 转换工具
```

## 如何部署（首次）
1. 克隆你的仓库：`git clone https://github.com/janicexiaotong-lgtm/cat-symptom-diary`
2. 把本目录全部文件复制进仓库根目录（覆盖/新增）。
3. 提交并推送：`git add -A && git commit -m "add 病例汇总" && git push`
4. 仓库 Settings → Pages 已指向 `main` 分支根目录，推送后约 1 分钟生效。
5. 访问 https://janicexiaotong-lgtm.github.io/cat-symptom-diary/ ，
   底部出现「📋 病例汇总」Tab，概览页也有「查看病例汇总 →」卡片。

## 后续新增病例如何「自动同步」
**只需更新数据文件并推送，其余全自动：**

- **方式 A（推荐，零代码）：**
  1. 用 Excel 打开 `case-data/cat-case-2023-2026.xlsx`，在「详情」表追加一行新就诊记录。
  2. 提交推送：`git add case-data/cat-case-2023-2026.xlsx && git commit -m "新增就诊" && git push`
  3. GitHub Actions 会自动重跑 `build_case_summary.py` → 重新生成 `case-summary.html` → 提交并部署。
  4. 日记里的「病例汇总」入口内容随之更新，无需再动 `index.html`。

- **方式 B（本地预览后再推）：**
  1. 本地运行：`python build_case_summary.py`（需 `pip install pandas openpyxl plotly`）。
  2. 双击 `case-summary.html` 本地核对效果。
  3. 提交推送，Pages 更新。

## 注意
- `case-summary.html` 由脚本生成，请**不要手动编辑**，否则下次自动构建会被覆盖。
- 若病例 Excel 改名或换路径，请同步修改 `build_case_summary.py` 中的 `SRC`。
- 症状日记本身的数据仍存于浏览器 localStorage（与病例可视化相互独立）；
  本「病例汇总」展示的是病历 Excel 中的临床检验/超声趋势。
