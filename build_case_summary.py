# -*- coding: utf-8 -*-
"""猫咪病例数据 2023-2026 可视化汇总 v3（医生视角 · 美观 + 易读）
 - 肌酐头条图：IRIS 分期色带 + 临床事件"时间线式"标注（按发生日期的竖线 + 双行无重叠标签）
 - 7 项指标网格：每图标题直接标注猫正常参考区间；增加"图示说明"面板
 - 删除体重趋势图（体重仅保留顶部卡片与年度表）
 - 顶部"病例概要速览"便于零背景医生快速建立认知
"""
import os
import re
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots

SRC = os.path.join(os.path.dirname(os.path.abspath(__file__)), "case-data", "cat-case-2023-2026.xlsx")
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "case-summary.html")

# ---------- 1. 读取并清洗 ----------
raw = pd.read_excel(SRC, sheet_name="详情", header=[0, 1])
newcols = []
for a, b in raw.columns:
    a = str(a).replace("\n", ""); b = str(b).replace("\n", "")
    newcols.append(b if not b.startswith("Unnamed") else a)
raw.columns = newcols
keep = ["病史", "日期", "肌酐(μmol/L)", "尿素氮(mg/dL)", "总钙(mmol/L)", "总磷(mmol/L)",
        "其他异常项", "HCT(%)", "钙离子(mmol/L)", "钾离子(mmol/L)",
        "左肾", "右肾", "其他指征", "体重(kg)", "补充检查", "用药", "医院"]
df = raw[keep].copy()
df = df.rename(columns={
    "肌酐(μmol/L)": "肌酐", "尿素氮(mg/dL)": "尿素氮", "总钙(mmol/L)": "总钙",
    "总磷(mmol/L)": "总磷", "HCT(%)": "HCT", "钙离子(mmol/L)": "钙离子",
    "钾离子(mmol/L)": "钾离子", "体重(kg)": "体重", "其他指征": "超声其他"})
for c in ["肌酐", "尿素氮", "总钙", "总磷", "HCT", "钙离子", "钾离子", "体重"]:
    df[c] = pd.to_numeric(df[c], errors="coerce")
df["HCT"] = df["HCT"] * 100.0

def parse_kidney(val):
    if pd.isna(val): return (np.nan, np.nan)
    s = str(val)
    m_len = re.match(r"\s*([\d.]+)", s); m_pel = re.search(r"肾盂([\d.]+)", s)
    return (float(m_len.group(1)) if m_len else np.nan, float(m_pel.group(1)) if m_pel else np.nan)
df["左肾长度"], df["左肾盂"] = zip(*df["左肾"].map(parse_kidney))
df["右肾长度"], df["右肾盂"] = zip(*df["右肾"].map(parse_kidney))

df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
df = df.dropna(subset=["日期"]).sort_values("日期").reset_index(drop=True)
df["年份"] = df["日期"].dt.year

events = df[df["病史"].notna() & (df["病史"].astype(str).str.strip() != "")][["日期", "病史"]].copy()
SHORT = {
    "2次桡尺骨骨折手术，2次拆板": "骨折术①②",
    "第3次桡尺骨骨折手术\n第3次拆板\nCKD": "CKD确诊·骨折③",
    "第4次桡尺骨骨折手术，骨板保留": "骨折④",
    "疑似FIC": "疑似FIC",
    "输尿管梗阻": "输尿管梗阻",
    "出血性肠炎": "出血性肠炎",
    "常规CKD复查\n慢性呕吐": "慢性呕吐",
    "FIC\n反复出现症状": "FIC复发",
}
events["标签"] = events["病史"].map(lambda x: SHORT.get(str(x).strip(), str(x).strip()[:6]))
events = events.sort_values("日期").reset_index(drop=True)

# ---------- 2. 指标配置（7 项，不含体重） ----------
metrics = [
    ("尿素氮", "尿素氮 BUN", "mg/dL", (15, 34), "#e76f51"),
    ("总磷", "总磷 P", "mmol/L", (1.0, 2.4), "#2a9d8f"),
    ("总钙", "总钙 Ca", "mmol/L", (2.0, 2.8), "#264653"),
    ("HCT", "红细胞压积 HCT", "%", (30, 45), "#e63946"),
    ("钙离子", "离子钙 iCa", "mmol/L", (1.2, 1.4), "#457b9d"),
    ("钾离子", "血钾 K", "mmol/L", (3.5, 5.8), "#8338ec"),
]
IRIS = [(0, 140, "#1a9850", "IRIS 1期 (<140)"),
        (140, 250, "#fee08b", "IRIS 2期 (140–250)"),
        (250, 440, "#fc8d59", "IRIS 3期 (250–440)"),
        (440, 600, "#d73027", "IRIS 4期 (>440)")]

crea = df["肌酐"].dropna(); crea_peak = crea.max(); crea_latest = crea.iloc[-1]
latest_w = df["体重"].dropna().iloc[-1]
n_visits = len(df); n_years = df["年份"].nunique()
span = f"{df['日期'].min():%Y-%m-%d} ~ {df['日期'].max():%Y-%m-%d}"

def iris_stage(v):
    if pd.isna(v): return "-"
    return "1" if v < 140 else "2" if v < 250 else "3" if v < 440 else "4"

# ---------- 3. 病例概要（零背景医生速览） ----------
# 骨折手术次数：源表按"就诊行"计数会漏算——首行"2次桡尺骨骨折手术"含 ①② 两次，
# 加第3次(③)、第4次(④) 共 4 次；麻醉次数来自临床记录汇总（源表未单列麻醉列）。
FRACTURE_SURGERIES = 4
ANESTHESIA_TIMES = 7
has_uro = bool(events["病史"].str.contains("输尿管梗阻", na=False).any())
has_fic = bool(events["病史"].str.contains("FIC", na=False).any())
has_ent = bool(events["病史"].str.contains("肠炎", na=False).any())
facts = [f"主线诊断：慢性肾病 <b>CKD</b>（以 IRIS 3 期为主，最新肌酐 {crea_latest:.0f} μmol/L）"]
if FRACTURE_SURGERIES: facts.append(f"骨骼：反复桡尺骨骨折手术共 <b>{FRACTURE_SURGERIES}</b> 次")
if ANESTHESIA_TIMES: facts.append(f"麻醉：累计 <b>{ANESTHESIA_TIMES}</b> 次")
if has_uro: facts.append("急症：<b>2025-01 输尿管梗阻</b>（肌酐冲至峰值 {:.0f}）".format(crea_peak))
if has_fic: facts.append("泌尿系：反复 <b>FIC</b>（猫特发性膀胱炎）")
if has_ent: facts.append("消化：<b>2025-05 出血性肠炎</b>")
facts.append(f"随访：{span}，共 <b>{n_visits}</b> 次就诊、{n_years} 个年份")
facts.append("长期关注：肌酐趋势、慢性呕吐、肾盂扩张")
snapshot = "".join(f"<li>{f}</li>" for f in facts)

# ---------- 4. 图例面板文本 ----------
legend_text = ("<b>图示说明</b><br><br>"
               '<span style="color:#9aa0a6">■</span> 阴影 ＝ 猫正常参考区间<br>'
               '<span style="color:#2b2d42">┊</span> 虚线 ＝ 临床事件（按发生当日）<br>'
               '<span style="color:#6a1b9a">━</span> 折线 ＝ 该指标随时间变化<br><br>'
               "曲线越出阴影即提示异常，<br>可据此做长期随访对照。")

# ---------- 5. 构建仪表板 ----------
fig = make_subplots(
    rows=6, cols=2, row_heights=[1.55, 1, 1, 1, 1, 0.7], column_widths=[0.5, 0.5],
    specs=[[{"colspan": 2}, None],
           [{}, {}], [{}, {}], [{}, {}],
           [{}, {}], [{"colspan": 2}, None]],
    subplot_titles=(
        "",
        "尿素氮 BUN（mg/dL，正常 15–34）", "总磷 P（mmol/L，正常 1.0–2.4）",
        "总钙 Ca（mmol/L，正常 2.0–2.8）", "红细胞压积 HCT（%，正常 30–45）",
        "离子钙 iCa（mmol/L，正常 1.2–1.4）", "血钾 K（mmol/L，正常 3.5–5.8）",
        "肾脏超声 · 左/右肾长度（cm）", "肾盂扩张 · 左/右（cm）",
        "图示说明",
    ),
    vertical_spacing=0.07, horizontal_spacing=0.09,
)

# 头条：肌酐 + IRIS 色带
row_crea = df.dropna(subset=["肌酐"])
fig.add_trace(go.Scatter(
    x=row_crea["日期"], y=row_crea["肌酐"], mode="lines+markers", name="肌酐",
    line=dict(color="#6a1b9a", width=3), marker=dict(size=8, color="#6a1b9a", line=dict(width=1, color="white")),
    fill="tozeroy", fillcolor="rgba(106,27,154,0.06)",
    hovertemplate="%{x|%Y-%m-%d}<br>肌酐: %{y:.1f} μmol/L<extra></extra>"),
    row=1, col=1)
for lo, hi, color, lab in IRIS:
    fig.add_hrect(y0=lo, y1=hi, line_width=0, fillcolor=color, opacity=0.12,
                  annotation_text=lab, annotation_position="right",
                  annotation=dict(font=dict(size=10, color="#555")), row=1, col=1)
fig.update_yaxes(range=[60, 520], row=1, col=1)

# 事件：按日期竖线 + 双行无重叠标签（lane）
LANE_GAP = 50
last = None; lane = 0; lanes = []
for _, ev in events.iterrows():
    if last is not None and (ev["日期"] - last).days < LANE_GAP:
        lane += 1
    else:
        lane = 0
    lanes.append(lane); last = ev["日期"]
events["lane"] = lanes
for _, ev in events.iterrows():
    fig.add_vline(x=ev["日期"], line_width=1.6, line_dash="dot", line_color="#2b2d42", row=1, col=1)
    fig.add_annotation(
        x=ev["日期"], y=1.045 + ev["lane"] * 0.085, yref="y domain", xref="x",
        text=ev["标签"], showarrow=False, textangle=0,
        yanchor="bottom", xanchor="center",
        font=dict(size=10.5, color="#2b2d42", family="Microsoft YaHei"),
        bgcolor="rgba(255,255,255,0.8)", bordercolor="#2b2d42", borderwidth=0.5, borderpad=1,
        row=1, col=1)

# 7 项网格
grid = [("尿素氮", 2, 1), ("总磷", 2, 2), ("总钙", 3, 1), ("HCT", 3, 2),
        ("钙离子", 4, 1), ("钾离子", 4, 2)]
for key, r, c in grid:
    m = next(x for x in metrics if x[0] == key)
    sub = df.dropna(subset=[key])
    fig.add_trace(go.Scatter(
        x=sub["日期"], y=sub[key], mode="lines+markers", name=m[1],
        line=dict(color=m[4], width=2.5), marker=dict(size=7, color=m[4]),
        hovertemplate="%{x|%Y-%m-%d}<br>" + m[1] + ": %{y:.2f} " + m[2] + "<extra></extra>"),
        row=r, col=c)
    if m[3]:
        fig.add_hrect(y0=m[3][0], y1=m[3][1], line_width=0, fillcolor=m[4], opacity=0.07, row=r, col=c)

# 超声：左/右肾长度（row5 c1）
for key, name, color in [("左肾长度", "左肾长度", "#006d77"), ("右肾长度", "右肾长度", "#f4a261")]:
    sub = df.dropna(subset=[key])
    fig.add_trace(go.Scatter(
        x=sub["日期"], y=sub[key], mode="lines+markers", name=name,
        line=dict(color=color, width=2.5), marker=dict(size=7, color=color),
        hovertemplate="%{x|%Y-%m-%d}<br>" + name + ": %{y:.2f} cm<extra></extra>"),
        row=5, col=1)
# 超声：左/右肾盂扩张（row5 c2）
for key, name, color in [("左肾盂", "左肾盂", "#118ab2"), ("右肾盂", "右肾盂", "#ef476f")]:
    sub = df.dropna(subset=[key])
    fig.add_trace(go.Scatter(
        x=sub["日期"], y=sub[key], mode="lines+markers", name=name,
        line=dict(color=color, width=2.5), marker=dict(size=7, color=color),
        hovertemplate="%{x|%Y-%m-%d}<br>" + name + ": %{y:.2f} cm<extra></extra>"),
        row=5, col=2)

# 图例说明面板（底部整行）
fig.add_trace(go.Scatter(
    x=[0], y=[0], mode="text", text=[legend_text], textposition="middle center",
    hoverinfo="skip", showlegend=False), row=6, col=1)
fig.update_xaxes(showticklabels=False, showgrid=False, zeroline=False, row=6, col=1)
fig.update_yaxes(showticklabels=False, showgrid=False, zeroline=False, row=6, col=1)

# 布局
fig.update_layout(
    template="plotly_white", height=2000, width=1280,
    title=dict(
        text="猫咪 CKD 长期随访可视化（2023–2026）",
        subtitle=dict(text="肾功能趋势 · IRIS 临床分期 · 关键临床事件时间线；各指标含猫正常参考区间", font=dict(size=14)),
        font=dict(size=22, color="#1d3557"), x=0.5, xanchor="center"),
    font=dict(family="Microsoft YaHei, PingFang SC, sans-serif", size=12, color="#222"),
    showlegend=False, hovermode="closest", margin=dict(l=55, r=40, t=120, b=40))
fig.update_xaxes(showgrid=True, gridcolor="#eee", tickformat="%Y-%m")
fig.update_yaxes(showgrid=True, gridcolor="#eee")

# ---------- 6. 表格 ----------
def clean(x):
    return "" if pd.isna(x) else str(x).replace("\n", " ").strip()
log_rows = ""
for _, r in df.iterrows():
    cell = " / ".join([v for v in [clean(r["病史"]), clean(r["其他异常项"]), clean(r["补充检查"]), clean(r["用药"])] if v])
    log_rows += f"<tr><td>{r['日期']:%Y-%m-%d}</td><td>{int(r['年份'])}</td><td>{clean(r['医院'])}</td><td>{cell}</td></tr>"

yearly = df.groupby("年份").agg(就诊次数=("日期", "count"), 肌酐均值=("肌酐", "mean"),
                                肌酐峰值=("肌酐", "max"), 体重均值=("体重", "mean")).reset_index()
yearly_html = ""
for _, r in yearly.iterrows():
    cm = f"{r['肌酐均值']:.0f}" if pd.notna(r["肌酐均值"]) else "-"
    cx = f"{r['肌酐峰值']:.0f}" if pd.notna(r["肌酐峰值"]) else "-"
    wm = f"{r['体重均值']:.2f}" if pd.notna(r["体重均值"]) else "-"
    yearly_html += f"<tr><td>{int(r['年份'])}</td><td>{int(r['就诊次数'])}</td><td>{cm}</td><td>{cx}</td><td>{wm}</td><td>{iris_stage(r['肌酐均值'])}</td></tr>"

# ---------- 7. 合并为自包含 HTML ----------
plotly_div = fig.to_html(full_html=False, include_plotlyjs=True, config={"displayModeBar": True})
html = f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>猫咪病例可视化汇总 2023-2026</title>
<style>
 body{{font-family:"Microsoft YaHei","PingFang SC",sans-serif;background:#f7f9fc;color:#222;margin:0;}}
 .wrap{{max-width:1320px;margin:0 auto;padding:24px;}}
 .cards{{display:flex;flex-wrap:wrap;gap:14px;margin:18px 0 6px;}}
 .card{{flex:1;min-width:150px;background:#fff;border-radius:12px;padding:16px 18px;box-shadow:0 2px 10px rgba(20,40,80,.08);border-left:4px solid #1d3557;}}
 .card .v{{font-size:24px;font-weight:700;color:#1d3557;}} .card .l{{font-size:13px;color:#666;margin-top:4px;}}
 .snap{{background:linear-gradient(135deg,#1d3557,#335c81);color:#fff;border-radius:12px;padding:16px 20px;margin:18px 0;box-shadow:0 4px 14px rgba(20,40,80,.18);}}
 .snap h2{{margin:0 0 8px;font-size:17px;color:#fff;}} .snap ul{{margin:0;padding-left:20px;line-height:1.7;font-size:13.5px;}}
 .panel{{background:#fff;border-radius:12px;padding:10px 14px 4px;margin:18px 0;box-shadow:0 2px 10px rgba(20,40,80,.08);}}
 h2{{font-size:18px;color:#1d3557;margin:6px 4px 0;}}
 table{{width:100%;border-collapse:collapse;font-size:13px;margin-top:8px;}}
 th,td{{border-bottom:1px solid #eef;padding:7px 8px;text-align:left;vertical-align:top;}}
 th{{background:#1d3557;color:#fff;position:sticky;top:0;}}
 tr:nth-child(even){{background:#f4f7fb;}}
 .note{{font-size:12px;color:#888;margin:4px 4px 14px;}}
</style></head><body><div class="wrap">
  <div class="cards">
    <div class="card"><div class="v">{n_visits}</div><div class="l">总就诊次数</div></div>
    <div class="card"><div class="v">{n_years}</div><div class="l">覆盖年份</div></div>
    <div class="card"><div class="v">{crea_latest:.0f}</div><div class="l">最新肌酐 μmol/L</div></div>
    <div class="card"><div class="v">{crea_peak:.0f}</div><div class="l">肌酐峰值 μmol/L</div></div>
    <div class="card"><div class="v">{latest_w:.2f}</div><div class="l">最新体重 kg</div></div>
  </div>

  <div class="snap"><h2>🩺 病例概要速览</h2><ul>{snapshot}</ul></div>

  <div class="panel">
    {plotly_div}
    <p class="note">肌酐图：背景色带为 IRIS 慢性肾病分期（绿=1期 / 黄=2期 / 橙=3期 / 红=4期）；竖向虚线为关键临床事件，严格对应发生当日，标签分两行避免重叠。其余各图阴影为猫正常参考区间，曲线越出阴影即异常。肾脏超声含左/右肾长度与肾盂扩张两条曲线，肾盂越粗（数值越大）表示梗阻 / 积水风险越高。</p>
  </div>

  <h2>📊 年度汇总（含 IRIS 主要分期）</h2>
  <div class="panel"><table>
    <tr><th>年份</th><th>就诊次数</th><th>肌酐均值(μmol/L)</th><th>肌酐峰值(μmol/L)</th><th>体重均值(kg)</th><th>IRIS 主要分期</th></tr>
    {yearly_html}
  </table></div>

  <h2>📋 完整就诊 / 用药 / 检查日志</h2>
  <div class="panel" style="max-height:520px;overflow:auto;"><table>
    <tr><th>日期</th><th>年份</th><th>医院</th><th>病史 / 异常项 / 补充检查 / 用药</th></tr>
    {log_rows}
  </table></div>
  <p class="note">数据来源：猫咪病例数据2023-2026.xlsx（数值占位如"只做B超"按缺失处理；超声已解析肾长度）。本页仅供长期趋势参考，不替代兽医诊断。</p>
</div></body></html>"""
with open(OUT, "w", encoding="utf-8") as f:
    f.write(html)

print("OK ->", OUT)
print("就诊:", n_visits, "| 年份:", sorted(df["年份"].unique().tolist()))
print("事件 lanes:", events["lane"].tolist(), "标签:", list(events["标签"]))
print("超声解析: 左肾长", df["左肾长度"].notna().sum(), "右肾长", df["右肾长度"].notna().sum(),
      "| 左肾盂", df["左肾盂"].notna().sum(), "右肾盂", df["右肾盂"].notna().sum())
print("肌酐最新/峰值:", round(crea_latest, 1), round(crea_peak, 1))
