import io
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="SKU 价格毛利分析", layout="wide")

# ── 页眉 ──────────────────────────────────────────────────────────────────────
h1, h2, h3 = st.columns([1, 3, 2])
with h1:
    try:
        st.image("carzone_logo.png", width=110)
    except Exception:
        pass
with h2:
    st.markdown("## SKU 价格毛利分析")
    st.caption("底盘件价格体系优化 · 核算价 × 毛利率模拟工具")
with h3:
    st.markdown(
        "<div style='text-align:right;line-height:1.7;font-size:13px;color:#888;padding-top:8px'>"
        "<b>南京新康众 · 供应链底盘组</b><br>制作人：李宇凡<br>"
        "<span style='color:#e74c3c;font-size:12px'>⚠ 仅供底盘组内部使用，注意保护敏感数据安全</span>"
        "</div>",
        unsafe_allow_html=True,
    )
st.divider()

# ── 侧边栏 ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("数据源")
    uploaded = st.file_uploader("上传 Excel 文件（.xlsx）", type=["xlsx"])

# ── 未上传时：使用说明页 ──────────────────────────────────────────────────────
if not uploaded:
    st.markdown("### 欢迎使用 SKU 价格毛利分析看板")
    st.markdown("上传 Excel 文件后，工具将自动识别列结构并开始分析。请先阅读以下使用说明。")

    with st.expander("📂 支持的文件格式", expanded=True):
        st.markdown("""
**推荐格式：SKU 价格导出标准文件**
- 必须包含：`加权平均采购成本`（供应商进货价）和 `核算价`（总部对门店定价）
- 工具会自动识别列名，无需手动调整

**不支持的格式：门店采购明细（如法雷奥/菲罗多采购流水）**
- 此类文件只记录门店向总部的拿货价，不含供应商进货成本，无法计算毛利率
- 上传后工具会自动检测并给出提示
        """)

    with st.expander("🔧 操作步骤", expanded=True):
        st.markdown("""
1. **上传文件**：点击左侧「上传 Excel 文件」，拖拽或选择 SKU 价格导出文件
2. **设置筛选**：在左侧选择品牌和品类（默认显示全部）
3. **调整调价幅度**：拖动「调价幅度」滑块，正数=涨价，负数=降价
4. **设置基准线**：拖动「毛利率基准线」滑块（默认 10%）
5. **调价范围筛选**：勾选「仅调整高于基准线的 SKU」（默认勾选），则当前毛利率 ≤ 基准线的 SKU 不参与调价，核算价维持不变；取消勾选则对所有 SKU 应用调价幅度
6. **读取分析结果**：
   - 顶部摘要卡：一句话总结调价影响（含排除 SKU 数）
   - 三张图：SKU 健康分布 / 毛利率分布对比 / 品类对比
   - 多幅度对比表：同时查看 5 种调价幅度的结果
   - 品类汇总表：各品类均值毛利率和风险 SKU 数
   - SKU 明细表：含最低达标价格和异常标记
7. **导出**：点击「导出带格式 Excel」，下载带颜色的 .xlsx 文件
        """)

    with st.expander("🎨 颜色说明"):
        st.markdown("""
| 颜色 | 含义 |
|------|------|
| 🟢 绿色 | 毛利率 ≥ 基准线，健康 |
| 🟡 黄色 | 毛利率在 8% 到基准线之间，预警 |
| 🔴 红色 | 毛利率 < 8%，危险 |
| ⚠️ 异常标记 | 该 SKU 毛利率偏离所在品类均值超过 1.5 个标准差 |
        """)

    with st.expander("❓ 常见问题"):
        st.markdown("""
**Q：上传后数据显示为空？**
A：检查文件是否含有加权平均采购成本列，且该列有有效数值（非 0 非空）。

**Q：列识别失败怎么办？**
A：工具会弹出手动选列面板，从下拉框中选择对应列即可。

**Q：调价后某些 SKU 毛利率变为负数？**
A：说明该 SKU 当前核算价已低于成本，需要涨价而非降价。

**Q：「最低达标价格」是怎么算的？**
A：最低达标价格 = 采购成本 ÷ (1 - 毛利率基准线)，即刚好达到基准线所需的最低核算价。
        """)

    st.stop()

# ── 文件类型检测 ──────────────────────────────────────────────────────────────
@st.cache_data
def detect_file_type(source):
    probe = pd.read_excel(source, engine="calamine", nrows=2)
    first_col = str(probe.columns[0]).lower()
    if "超过1w" in first_col or "出库查询" in first_col or "超过" in first_col:
        return "store"
    row1 = " ".join(str(v) for v in probe.iloc[0].values).lower()
    if "sku编码" in row1 and "折后单价" in row1:
        return "store"
    return "hq"

file_mode = detect_file_type(uploaded)

# ── 门店定价合理性分析 ────────────────────────────────────────────────────────
if file_mode == "store":
    @st.cache_data
    def load_store(source):
        COLS = ["sku编码","产品名称","产品俗称","小区","大区","品牌名称",
                "促销核算价","成本价","成本金额","核算价","核算金额",
                "折后单价","销售金额","坏件数量","成本毛利额QB","核算毛利额QB",
                "B端销售单价","B端销售金额","B端成本毛利"]
        df = pd.read_excel(source, engine="calamine", skiprows=1, header=0)
        df = df.iloc[:, :len(COLS)]
        df.columns = COLS
        df = df[df["sku编码"] != "sku编码"].copy()
        for c in ["核算价","折后单价","销售金额","成本价"]:
            df[c] = pd.to_numeric(df[c], errors="coerce")
        return df

    raw_s = load_store(uploaded)
    _has_price = raw_s[
        raw_s["折后单价"].notna() & raw_s["核算价"].notna() &
        (raw_s["核算价"] > 0) & (raw_s["折后单价"] > 0)
    ].copy()
    _has_price["门店毛利率"] = ((_has_price["折后单价"] - _has_price["核算价"]) / _has_price["折后单价"])
    _has_price = _has_price[np.isfinite(_has_price["门店毛利率"])].copy()
    # 折后单价 < 1元 = 测试/录入错误，剔出分析
    anomaly_s = _has_price[_has_price["折后单价"] < 1].copy()
    valid_s   = _has_price[_has_price["折后单价"] >= 1].copy()
    valid_s["是否亏本"] = valid_s["折后单价"] < valid_s["核算价"]
    n_skip = len(raw_s) - len(valid_s) - len(anomaly_s)

    with st.sidebar:
        st.divider()
        st.header("筛选")
        brands_s = ["全部"] + sorted(valid_s["品牌名称"].dropna().unique().tolist())
        regions_s = ["全部"] + sorted(valid_s["大区"].dropna().unique().tolist())
        sel_brand_s = st.selectbox("品牌", brands_s)
        sel_region_s = st.selectbox("大区", regions_s)
        st.divider()
        thr_store = st.slider("门店毛利率基准线（%）", 0, 30, 10, 1)

    view_s = valid_s.copy()
    if sel_brand_s != "全部":
        view_s = view_s[view_s["品牌名称"] == sel_brand_s]
    if sel_region_s != "全部":
        view_s = view_s[view_s["大区"] == sel_region_s]
    thr_s = thr_store / 100

    n_loss = int((view_s["门店毛利率"] < 0).sum())
    n_low  = int(((view_s["门店毛利率"] >= 0) & (view_s["门店毛利率"] < thr_s)).sum())
    n_ok   = int((view_s["门店毛利率"] >= thr_s).sum())
    avg_sm = view_s["门店毛利率"].mean() if len(view_s) else 0

    st.markdown("## 门店定价合理性分析")
    st.caption("出库敏感数据 · 以核算价为门店成本，折后单价为门店售价，计算门店实际毛利率")
    anomaly_note = f"  ⚠️ 另有 **{len(anomaly_s)}** 条折后单价 < 1元的疑似录入错误记录已剔除（不计入分析）。" if len(anomaly_s) > 0 else ""
    st.info(
        f"📊 共 **{len(view_s):,}** 条有效 B 端销售记录 · "
        f"亏本销售 **{n_loss}** 条 · 平均门店毛利率 **{avg_sm:.1%}** · "
        f"已过滤仓间调拨记录 **{n_skip}** 条（无折后单价）。{anomaly_note}"
    )

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("有效销售记录", f"{len(view_s):,}")
    c2.metric("亏本销售", f"{n_loss}",
              f"{n_loss/len(view_s)*100:.1f}%" if len(view_s) else "—", delta_color="inverse")
    c3.metric("低于基准线", f"{n_low}",
              f"{n_low/len(view_s)*100:.1f}%" if len(view_s) else "—", delta_color="inverse")
    c4.metric("平均门店毛利率", f"{avg_sm:.1%}")

    st.divider()

    v1, v2, v3 = st.columns([1, 1.6, 1.4])
    with v1:
        st.markdown("#### 门店定价健康分布")
        donut_s = go.Figure(go.Pie(
            labels=["健康（≥基准线）", f"微利（0%~{thr_store}%）", "亏本（<0%）"],
            values=[n_ok, n_low, n_loss],
            hole=0.62,
            marker_colors=["#2ecc71", "#f39c12", "#e74c3c"],
            textinfo="percent+value",
            hovertemplate="%{label}<br>%{value} 条<br>%{percent}<extra></extra>",
        ))
        donut_s.update_layout(
            margin=dict(t=10,b=10,l=10,r=10), height=260,
            legend=dict(orientation="h", yanchor="bottom", y=-0.25, font_size=11),
            annotations=[dict(text=f"<b>{n_ok}</b><br>健康", x=0.5, y=0.5,
                              font_size=15, showarrow=False)],
        )
        st.plotly_chart(donut_s, use_container_width=True, config={"displayModeBar": False})

    with v2:
        st.markdown("#### 门店毛利率分布")
        hist_s = go.Figure()
        hist_s.add_trace(go.Histogram(
            x=view_s["门店毛利率"]*100, nbinsx=40,
            marker_color="rgba(52,152,219,0.65)",
            hovertemplate="毛利率 %{x:.1f}%<br>记录数 %{y}<extra></extra>",
        ))
        hist_s.add_vline(x=thr_store, line_dash="dash", line_color="#e74c3c", line_width=1.5,
                         annotation_text=f"基准线 {thr_store}%", annotation_position="top right",
                         annotation_font_color="#e74c3c")
        hist_s.add_vline(x=0, line_dash="dot", line_color="#888", line_width=1,
                         annotation_text="零毛利", annotation_position="top left")
        hist_s.update_layout(
            margin=dict(t=10,b=10,l=10,r=10), height=260,
            xaxis_title="门店毛利率（%）", yaxis_title="记录数", bargap=0.05,
            showlegend=False,
        )
        st.plotly_chart(hist_s, use_container_width=True, config={"displayModeBar": False})

    with v3:
        st.markdown("#### 各大区平均门店毛利率")
        rgn_agg = view_s.groupby("大区")["门店毛利率"].mean().sort_values() * 100
        bar_s = go.Figure(go.Bar(
            x=rgn_agg.values, y=rgn_agg.index, orientation="h",
            marker_color=["#e74c3c" if v < thr_store else "#2ecc71" for v in rgn_agg.values],
            hovertemplate="%{y}<br>均值 %{x:.1f}%<extra></extra>",
        ))
        bar_s.add_vline(x=thr_store, line_dash="dash", line_color="#e74c3c", line_width=1.5)
        bar_s.update_layout(margin=dict(t=10,b=10,l=10,r=10), height=260,
                             xaxis_title="平均门店毛利率（%）")
        st.plotly_chart(bar_s, use_container_width=True, config={"displayModeBar": False})

    st.divider()

    # 品牌汇总表
    st.markdown("#### 各品牌门店定价汇总")
    brand_agg = (
        view_s.groupby("品牌名称")
        .agg(记录数=("门店毛利率","count"), 平均门店毛利率=("门店毛利率","mean"),
             亏本记录数=("是否亏本","sum"), 平均折后单价=("折后单价","mean"),
             平均核算价=("核算价","mean"))
        .reset_index()
    )
    brand_agg["亏本占比"] = brand_agg["亏本记录数"] / brand_agg["记录数"]
    brand_agg = brand_agg.sort_values("平均门店毛利率")

    def style_sm(v):
        try:
            f = float(v)
            if f < 0: return "background-color:#f8d7da;color:#1a1a1a"
            elif f < thr_s: return "background-color:#fff3cd;color:#1a1a1a"
            return "background-color:#d4edda;color:#1a1a1a"
        except: return ""

    st.dataframe(
        brand_agg.style
        .map(style_sm, subset=["平均门店毛利率"])
        .format({"平均门店毛利率":"{:.1%}","亏本占比":"{:.1%}",
                 "平均折后单价":"{:.2f}","平均核算价":"{:.2f}"}),
        use_container_width=True, height=300, hide_index=True,
    )

    st.divider()

    # 明细：低于基准线
    st.markdown(f"#### 低于基准线明细（门店毛利率 < {thr_store}%）")
    prob = (
        view_s[view_s["门店毛利率"] < thr_s]
        .sort_values("门店毛利率")[
            ["sku编码","产品俗称","品牌名称","大区","小区",
             "核算价","折后单价","销售金额","门店毛利率","是否亏本"]
        ].copy()
    )

    def urgent_s(row):
        bg = "#fff0f0" if row["是否亏本"] else "#fffbea"
        return [f"background-color:{bg};color:#1a1a1a"] * len(row)

    st.dataframe(
        prob.style.apply(urgent_s, axis=1)
        .format({"核算价":"{:.2f}","折后单价":"{:.2f}",
                 "销售金额":"{:.2f}","门店毛利率":"{:.1%}"}),
        use_container_width=True, height=400, hide_index=True,
    )
    st.download_button(
        "📋 下载低于基准线明细（CSV）",
        data=prob.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"门店定价异常_{thr_store}pct基准线.csv",
        mime="text/csv",
    )
    st.stop()

# ── 数据加载 ──────────────────────────────────────────────────────────────────
@st.cache_data
def load_raw(source):
    return pd.read_excel(source)

def find_col(cols, keywords, exclude=None):
    exclude = exclude or []
    for kw in keywords:
        for c in cols:
            c_lower = str(c).lower()
            if kw.lower() in c_lower and not any(ex.lower() in c_lower for ex in exclude):
                return c
    return None

raw_df = load_raw(uploaded)
cols = list(raw_df.columns)

auto = {
    "SKU编码":  find_col(cols, ["sku编码","sku","康众编码","商品编码","库内编码","货号"]),
    "品牌":     find_col(cols, ["品牌"]),
    "品类":     find_col(cols, ["品类","产品分类","一级分类","商品目录","粗称","分类","类目名称","类目"]),
    "采购成本": find_col(cols, ["加权平均","加权成本","进货价","供应商价","采购成本"],
                         exclude=["核算价","金额"]),
    "核算价":   find_col(cols, ["核算价","采购核算","结算价","售价","零售价","销售价"],
                         exclude=["金额"]),
}

missing = [k for k, v in auto.items() if v is None]
if missing:
    with st.expander("⚙️ 自动识别列失败，请手动选择（点击展开）", expanded=True):
        st.caption(f"未能自动识别：{', '.join(missing)}，请从下拉框选择对应列。")
        for field in missing:
            auto[field] = st.selectbox(f"{field} 对应哪一列？",
                                       ["（跳过）"] + cols, key=f"sel_{field}")
            if auto[field] == "（跳过）":
                auto[field] = None

if not auto["采购成本"] or not auto["核算价"]:
    st.warning("至少需要指定「采购成本」和「核算价」两列才能开始分析。")
    st.stop()

keep = {v: k for k, v in auto.items() if v}
raw = raw_df[list(keep.keys())].rename(columns=keep).copy()
for col in ["SKU编码", "品牌", "品类"]:
    if col not in raw.columns:
        raw[col] = "—"
raw["采购成本"] = pd.to_numeric(raw["采购成本"], errors="coerce")
raw["核算价"]   = pd.to_numeric(raw["核算价"],   errors="coerce")

valid = raw[raw["采购成本"].notna() & raw["核算价"].notna() &
            (raw["采购成本"] > 0) & (raw["核算价"] > 0)]
if len(valid) > 0:
    cm, pm = valid["采购成本"].mean(), valid["核算价"].mean()
    if abs(cm - pm) / max(pm, 1) < 0.02:
        st.warning(
            f"⚠️ **检测到成本列与核算价列数值几乎相等（差异 < 2%）**\n\n"
            f"成本列 `{auto['采购成本']}` 均值 {cm:.2f}，核算价列 `{auto['核算价']}` 均值 {pm:.2f}。\n\n"
            "此文件可能是门店采购流水，采购核算价为门店从总部的拿货价，不含供应商进货成本，无法计算毛利率。"
        )
        st.stop()

df = raw[raw["采购成本"].notna() & raw["核算价"].notna()].copy()
df = df[(df["采购成本"] > 0) & (df["核算价"] > 0)].reset_index(drop=True)
df["当前毛利率"] = (df["核算价"] - df["采购成本"]) / df["核算价"]

# ── 侧边栏：筛选 + 调价参数 ───────────────────────────────────────────────────
with st.sidebar:
    st.divider()
    st.header("筛选")
    brands = ["全部"] + sorted(df["品牌"].dropna().unique().tolist())
    cat_options = sorted(df["品类"].dropna().unique().tolist())
    sel_brand = st.selectbox("品牌", brands)
    sel_cat   = st.multiselect("品类", cat_options, placeholder="全部（不限）")
    st.divider()
    st.header("调价参数")
    threshold = st.slider("毛利率基准线（%）", 5, 25, 10, 1)
    only_above_thr = st.checkbox(
        "仅调整高于基准线的 SKU",
        value=True,
        help="勾选后：当前毛利率 ≤ 基准线的 SKU 不参与调价，核算价维持不变",
    )
    use_tiered = st.checkbox(
        "启用分层定价（目标毛利率）",
        value=False,
        help="按当前毛利率分两层，分别设置目标毛利率，自动计算每个 SKU 的精确落地价；只降不升",
    )
    if use_tiered and only_above_thr:
        tier_high_floor = st.slider("高毛利层起点（%）", 15, 50, 20, 1,
                                    help="高于此值为高毛利层，基准线到此值之间为中毛利层")
        tier_high_target = st.slider(
            f"高毛利层（≥{tier_high_floor}%）目标毛利率（%）",
            threshold, 40, 15, 1,
            help="调价后的目标落点，系统自动计算每个 SKU 所需的降价幅度")
        tier_mid_target = st.slider(
            f"中毛利层（{threshold}%~{tier_high_floor}%）目标毛利率（%）",
            threshold, 40, threshold + 2, 1,
            help="调价后的目标落点，只降不升")
        adj_pct = 0
    else:
        adj_pct = st.slider("调价幅度（%）", -30, 30, 0, 1, help="正数=涨价，负数=降价")
        tier_high_floor, tier_high_target, tier_mid_target = 20, 15, 12

# ── 筛选 + 计算 ───────────────────────────────────────────────────────────────
view = df.copy()
if sel_brand != "全部":
    view = view[view["品牌"] == sel_brand]
if sel_cat:
    view = view[view["品类"].isin(sel_cat)]
view = view.copy()
thr = threshold / 100
high_floor = tier_high_floor / 100
if only_above_thr and use_tiered:
    high_tgt = tier_high_target / 100
    mid_tgt  = tier_mid_target  / 100
    view["调整后核算价"] = np.select(
        [
            view["当前毛利率"] >= high_floor,
            (view["当前毛利率"] > thr) & (view["当前毛利率"] < high_floor),
        ],
        [
            np.minimum(view["采购成本"] / (1 - high_tgt), view["核算价"]),
            np.minimum(view["采购成本"] / (1 - mid_tgt),  view["核算价"]),
        ],
        default=view["核算价"],
    )
elif only_above_thr:
    view["调整后核算价"] = np.where(
        view["当前毛利率"] > thr,
        view["核算价"] * (1 + adj_pct / 100),
        view["核算价"],
    )
else:
    view["调整后核算价"] = view["核算价"] * (1 + adj_pct / 100)
view["调整后毛利率"] = (view["调整后核算价"] - view["采购成本"]) / view["调整后核算价"]
view["单位利润变化"] = view["调整后核算价"] - view["核算价"]

n_total    = len(view)
n_excluded = int((view["当前毛利率"] <= thr).sum()) if only_above_thr else 0
n_red    = int((view["调整后毛利率"] < 0.08).sum())
n_yellow = int(((view["调整后毛利率"] >= 0.08) & (view["调整后毛利率"] < thr)).sum())
n_green  = int((view["调整后毛利率"] >= thr).sum())
n_below  = n_red + n_yellow
n_below0 = int((view["当前毛利率"] < thr).sum())
avg_before = view["当前毛利率"].mean()
avg_after  = view["调整后毛利率"].mean()

# ── 功能1：调价影响摘要卡 ─────────────────────────────────────────────────────
cat_improvement = (
    view.groupby("品类")
    .agg(改善量=("调整后毛利率", "mean"))
    .assign(调价前均值=view.groupby("品类")["当前毛利率"].mean())
    .assign(改善pp=lambda x: (x["改善量"] - x["调价前均值"]) * 100)
    .sort_values("改善pp", ascending=False)
)
best_cat = cat_improvement.index[0] if len(cat_improvement) > 0 else "—"
best_pp  = cat_improvement["改善pp"].iloc[0] if len(cat_improvement) > 0 else 0

delta_below = n_below - n_below0
delta_sign  = "减少" if delta_below < 0 else ("增加" if delta_below > 0 else "持平")
delta_abs   = abs(delta_below)
pp_change   = (avg_after - avg_before) * 100

if use_tiered and only_above_thr:
    n_high_tier = int((view["当前毛利率"] >= high_floor).sum())
    n_mid_tier  = int(((view["当前毛利率"] > thr) & (view["当前毛利率"] < high_floor)).sum())
    summary_text = (
        f"**分层定价** · 高毛利层（≥{tier_high_floor}%）**{n_high_tier}** 个 SKU → 目标毛利率 **{tier_high_target}%**，"
        f"中毛利层（{threshold}%~{tier_high_floor}%）**{n_mid_tier}** 个 SKU → 目标毛利率 **{tier_mid_target}%**，"
        f"基准线以下 **{n_excluded}** 个 SKU 不动。"
        f"均值毛利率 **{avg_before:.1%}** → **{avg_after:.1%}**（{pp_change:+.1f}pp）。"
    )
else:
    filter_note = (
        f"，其中 **{n_excluded}** 个低于基准线的 SKU 已排除调价（筛选规则：仅调整高于 {threshold}% 的 SKU）"
        if only_above_thr and n_excluded > 0 else ""
    )
    if adj_pct == 0:
        summary_text = (
            f"当前未调价。共 **{n_total:,}** 个 SKU，其中 **{n_below0}** 个低于基准线 {threshold}%，"
            f"均值毛利率 **{avg_before:.1%}**。拖动左侧滑块模拟调价效果。"
        )
    else:
        summary_text = (
            f"调价 **{adj_pct:+d}%** 后，低于基准线的 SKU 从 **{n_below0}** 个{delta_sign}至 **{n_below}** 个"
            f"（{delta_sign} {delta_abs} 个），均值毛利率从 **{avg_before:.1%}** → **{avg_after:.1%}**"
            f"（{pp_change:+.1f}pp）。**{best_cat}** 品类改善最显著（{best_pp:+.1f}pp）{filter_note}。"
        )

st.info(f"📊 **调价影响摘要** · {summary_text}")

# ── 概览指标卡 ────────────────────────────────────────────────────────────────
c1, c2, c3, c4 = st.columns(4)
c1.metric("SKU 总数", f"{n_total:,}")
c2.metric("低于基准线", f"{n_below:,}",
          f"{n_below/n_total*100:.1f}%" if n_total else "—", delta_color="inverse")
c3.metric("当前均值毛利率", f"{avg_before:.1%}")
c4.metric("调整后均值毛利率", f"{avg_after:.1%}",
          f"{avg_after - avg_before:+.1%}", delta_color="normal")

st.divider()

# ── 可视化三图 ────────────────────────────────────────────────────────────────
v1, v2, v3 = st.columns([1, 1.6, 1.4])

with v1:
    st.markdown("#### SKU 健康分布")
    donut = go.Figure(go.Pie(
        labels=["健康（≥基准线）", f"预警（8%~{threshold}%）", "危险（<8%）"],
        values=[n_green, n_yellow, n_red],
        hole=0.62,
        marker_colors=["#2ecc71", "#f39c12", "#e74c3c"],
        textinfo="percent+value",
        hovertemplate="%{label}<br>%{value} 个 SKU<br>占比 %{percent}<extra></extra>",
    ))
    donut.update_layout(
        margin=dict(t=10, b=10, l=10, r=10), height=260,
        legend=dict(orientation="h", yanchor="bottom", y=-0.25, font_size=11),
        annotations=[dict(text=f"<b>{n_green}</b><br>健康", x=0.5, y=0.5,
                          font_size=15, showarrow=False)],
    )
    st.plotly_chart(donut, use_container_width=True, config={"displayModeBar": False})

with v2:
    st.markdown("#### 毛利率分布（调价前 vs 调价后）")
    hist = go.Figure()
    hist.add_trace(go.Histogram(x=view["当前毛利率"]*100, name="调价前", nbinsx=40,
                                marker_color="rgba(52,152,219,0.55)",
                                hovertemplate="毛利率 %{x:.1f}%<br>SKU 数 %{y}<extra>调价前</extra>"))
    hist.add_trace(go.Histogram(x=view["调整后毛利率"]*100, name="调价后", nbinsx=40,
                                marker_color="rgba(46,204,113,0.55)",
                                hovertemplate="毛利率 %{x:.1f}%<br>SKU 数 %{y}<extra>调价后</extra>"))
    hist.add_vline(x=threshold, line_dash="dash", line_color="#e74c3c", line_width=1.5,
                   annotation_text=f"基准线 {threshold}%", annotation_position="top right",
                   annotation_font_color="#e74c3c")
    hist.update_layout(barmode="overlay", margin=dict(t=10, b=30, l=40, r=10), height=260,
                       xaxis_title="毛利率（%）", yaxis_title="SKU 数",
                       legend=dict(orientation="h", yanchor="bottom", y=1.02),
                       plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
    st.plotly_chart(hist, use_container_width=True, config={"displayModeBar": False})

with v3:
    st.markdown("#### 品类毛利率对比")
    cat_sum = (view.groupby("品类")
               .agg(调价前=("当前毛利率","mean"), 调价后=("调整后毛利率","mean"))
               .reset_index().sort_values("调价后", ascending=True))
    bar = go.Figure()
    bar.add_trace(go.Bar(y=cat_sum["品类"], x=cat_sum["调价前"]*100, name="调价前",
                         orientation="h", marker_color="rgba(52,152,219,0.7)",
                         hovertemplate="%{y}<br>调价前 %{x:.1f}%<extra></extra>"))
    bar.add_trace(go.Bar(y=cat_sum["品类"], x=cat_sum["调价后"]*100, name="调价后",
                         orientation="h", marker_color="rgba(46,204,113,0.7)",
                         hovertemplate="%{y}<br>调价后 %{x:.1f}%<extra></extra>"))
    bar.add_vline(x=threshold, line_dash="dash", line_color="#e74c3c", line_width=1.5)
    bar.update_layout(barmode="group", margin=dict(t=10, b=30, l=10, r=10), height=260,
                      xaxis_title="均值毛利率（%）",
                      legend=dict(orientation="h", yanchor="bottom", y=1.02),
                      plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
    st.plotly_chart(bar, use_container_width=True, config={"displayModeBar": False})

st.divider()

# ── 功能2：多幅度并排对比表（分层模式下隐藏）───────────────────────────────────
if not (use_tiered and only_above_thr):
    st.subheader("多幅度调价对比")
    scenarios = [-10, -5, 0, 5, 10]
    scenario_rows = []
    for s in scenarios:
        adj_price = view["核算价"] * (1 + s / 100)
        if only_above_thr:
            adj_price = np.where(view["当前毛利率"] > thr, adj_price, view["核算价"])
        adj_margin = (adj_price - view["采购成本"]) / adj_price
        scenario_rows.append({
            "调价幅度": f"{s:+d}%",
            "均值毛利率":            adj_margin.mean(),
            f"低于{threshold}%的SKU数": int((adj_margin < thr).sum()),
            "危险SKU数(<8%)":        int((adj_margin < 0.08).sum()),
            "健康SKU数":             int((adj_margin >= thr).sum()),
        })
    scenario_df = pd.DataFrame(scenario_rows)

    def highlight_scenario(row):
        if row["调价幅度"] == f"{adj_pct:+d}%":
            return ["background-color: #1a4a2e; color: #fff; font-weight:bold"] * len(row)
        return [""] * len(row)

    fmt_s = {"均值毛利率": "{:.1%}"}
    st.dataframe(
        scenario_df.style.apply(highlight_scenario, axis=1).format(fmt_s),
        use_container_width=True, hide_index=True
    )
    st.divider()

# ── 品类汇总表 ────────────────────────────────────────────────────────────────
st.subheader("品类汇总")
summary = (
    view.groupby("品类", sort=False)
    .agg(SKU数=("SKU编码","count"), 均值成本=("采购成本","mean"),
         均值核算价=("核算价","mean"), 当前均值毛利率=("当前毛利率","mean"),
         调整后均值毛利率=("调整后毛利率","mean"),
         低于基准=("调整后毛利率", lambda x: (x < thr).sum()))
    .reset_index()
)

def style_margin(val):
    if isinstance(val, float):
        if val < 0.08:   return "color:#c0392b;font-weight:bold"
        elif val < thr:  return "color:#e67e22;font-weight:bold"
        else:            return "color:#27ae60"
    return ""

fmt = {"均值成本":"{:.2f}","均值核算价":"{:.2f}",
       "当前均值毛利率":"{:.1%}","调整后均值毛利率":"{:.1%}"}
st.dataframe(
    summary.style.map(style_margin, subset=["当前均值毛利率","调整后均值毛利率"]).format(fmt),
    use_container_width=True, hide_index=True
)

# ── 品牌分析面板 ──────────────────────────────────────────────────────────────
st.subheader("品牌分析")

brand_sum = (
    view.groupby("品牌")
    .agg(
        SKU数=("SKU编码", "count"),
        当前均值毛利率=("当前毛利率", "mean"),
        调整后均值毛利率=("调整后毛利率", "mean"),
        风险SKU数=("调整后毛利率", lambda x: (x < thr).sum()),
        危险SKU数=("调整后毛利率", lambda x: (x < 0.08).sum()),
    )
    .reset_index()
    .sort_values("调整后均值毛利率", ascending=True)
)
brand_sum["风险占比%"] = (brand_sum["风险SKU数"] / brand_sum["SKU数"] * 100).round(1)

n_brands = brand_sum["品牌"].nunique()
n_cats   = view["品类"].nunique()
ba1, ba2 = st.columns([1, 1.5]) if (n_brands > 1 and n_cats > 1) else st.columns(2)

with ba1:
    st.markdown("#### 各品牌均值毛利率")
    bar_colors = [
        "#e74c3c" if m < 0.08 else ("#f39c12" if m < thr else "#2ecc71")
        for m in brand_sum["调整后均值毛利率"]
    ]
    brand_bar = go.Figure(go.Bar(
        y=brand_sum["品牌"],
        x=brand_sum["调整后均值毛利率"] * 100,
        orientation="h",
        marker_color=bar_colors,
        text=[f"{m:.1%}" for m in brand_sum["调整后均值毛利率"]],
        textposition="outside",
        hovertemplate="%{y}<br>均值毛利率 %{x:.1f}%<extra></extra>",
    ))
    brand_bar.add_vline(x=threshold, line_dash="dash", line_color="#e74c3c",
                        line_width=1.5,
                        annotation_text=f"基准线 {threshold}%",
                        annotation_position="top right",
                        annotation_font_color="#e74c3c")
    brand_bar.update_layout(
        margin=dict(t=10, b=30, l=10, r=70),
        height=max(200, len(brand_sum) * 44),
        xaxis_title="均值毛利率（%）",
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
    )
    st.plotly_chart(brand_bar, use_container_width=True, config={"displayModeBar": False})

with ba2:
    if n_brands >= 2 and n_cats >= 2:
        st.markdown("#### 品牌 × 品类 毛利率热力图")
        pivot = (
            view.groupby(["品类", "品牌"])["调整后毛利率"]
            .mean()
            .unstack(fill_value=np.nan)
        )
        z_vals   = pivot.values * 100
        txt_vals = [[f"{v:.1f}%" if not np.isnan(v) else "—" for v in row]
                    for row in z_vals]
        hm = go.Figure(go.Heatmap(
            z=z_vals,
            x=list(pivot.columns),
            y=list(pivot.index),
            colorscale="RdYlGn",
            zmin=0, zmax=30,
            text=txt_vals,
            texttemplate="%{text}",
            hovertemplate="品类：%{y}<br>品牌：%{x}<br>均值毛利率：%{z:.1f}%<extra></extra>",
            colorbar=dict(title="毛利率%"),
        ))
        hm.update_layout(
            margin=dict(t=10, b=10, l=10, r=10),
            height=max(260, len(pivot.index) * 32 + 80),
            xaxis=dict(tickangle=-20, side="top"),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(hm, use_container_width=True, config={"displayModeBar": False})
    else:
        st.info(f"当前视图含 {n_brands} 个品牌 / {n_cats} 个品类，需至少各 2 个才能生成热力图。")

fmt_b = {"当前均值毛利率": "{:.1%}", "调整后均值毛利率": "{:.1%}", "风险占比%": "{:.1f}%"}
st.dataframe(
    brand_sum.style.map(style_margin, subset=["当前均值毛利率", "调整后均值毛利率"]).format(fmt_b),
    use_container_width=True, hide_index=True
)

st.divider()

# ── SKU 明细（含功能3异常标记 + 功能4最低达标价格）────────────────────────────
st.subheader("SKU 明细")

# 功能3：品类内异常标记
cat_stats = view.groupby("品类")["调整后毛利率"].agg(["mean","std"]).reset_index()
cat_stats.columns = ["品类","品类均值","品类标准差"]
view2 = view.merge(cat_stats, on="品类", how="left")
view2["异常"] = np.where(
    (view2["品类标准差"] > 0) &
    (abs(view2["调整后毛利率"] - view2["品类均值"]) > 1.5 * view2["品类标准差"]),
    "⚠️", ""
)

# 功能4：最低达标价格
view2["最低达标价格"] = view2["采购成本"] / (1 - thr)
view2["价格缺口"]     = view2["最低达标价格"] - view2["核算价"]

col_f, col_p = st.columns([2, 1])
with col_f:
    show_only_below = st.checkbox("只显示低于基准线的 SKU", value=False)
with col_p:
    page_size = st.selectbox("每页显示", [50, 100, 200], index=1)

detail = view2 if not show_only_below else view2[view2["调整后毛利率"] < thr]
detail = detail[["SKU编码","品类","品牌","采购成本","核算价","当前毛利率",
                  "调整后核算价","调整后毛利率","最低达标价格","价格缺口",
                  "单位利润变化","异常"]].reset_index(drop=True)

total_pages = max(1, (len(detail) - 1) // page_size + 1)
page = st.number_input(f"页码（共 {total_pages} 页，{len(detail)} 条）",
                       min_value=1, max_value=total_pages, value=1, step=1) - 1
page_df = detail.iloc[page*page_size:(page+1)*page_size].copy()

fmt2 = {"采购成本":"{:.2f}","核算价":"{:.2f}","当前毛利率":"{:.1%}",
        "调整后核算价":"{:.2f}","调整后毛利率":"{:.1%}",
        "最低达标价格":"{:.2f}","价格缺口":"{:+.2f}","单位利润变化":"{:+.2f}"}

def row_color(row):
    m = row["调整后毛利率"]
    bg = "#fff0f0" if m < 0.08 else ("#fffbea" if m < thr else "#f0fff4")
    return [f"background-color:{bg};color:#1a1a1a"] * len(row)

st.dataframe(
    page_df.style.apply(row_color, axis=1)
    .map(style_margin, subset=["当前毛利率","调整后毛利率"])
    .format(fmt2),
    use_container_width=True, height=420, hide_index=True
)

# ── 调价建议清单 ──────────────────────────────────────────────────────────────
st.subheader("调价建议清单")

if only_above_thr:
    above_df = view2[view2["当前毛利率"] > thr].copy()

    is_no_adj = False if use_tiered else (adj_pct == 0)

    if len(above_df) == 0:
        st.warning(f"当前筛选范围内无毛利率高于 {threshold}% 的 SKU。")
    elif is_no_adj:
        st.info(
            f"{'分层调价模式已开启。' if use_tiered else ''}"
            f"仅对高于基准线（{threshold}%）的 SKU 调价，当前共 **{len(above_df)}** 个符合条件，"
            f"均值毛利率 **{above_df['当前毛利率'].mean():.1%}**。调整左侧参数查看方案。"
        )
    else:
        if use_tiered:
            n_high_show = int((above_df["当前毛利率"] >= high_floor).sum())
            n_mid_show  = int((above_df["当前毛利率"] < high_floor).sum())
            n_high_adj = int(
                ((above_df["当前毛利率"] >= high_floor) &
                 (above_df["当前毛利率"] > tier_high_target / 100)).sum()
            )
            n_mid_adj = int(
                ((above_df["当前毛利率"] < high_floor) &
                 (above_df["当前毛利率"] > tier_mid_target / 100)).sum()
            )
            st.caption(
                f"**分层定价**：高毛利层（≥{tier_high_floor}%）**{n_high_show}** 个 SKU → 目标毛利率 **{tier_high_target}%**"
                f"（**{n_high_adj}** 个将实际降价）；"
                f"中毛利层（{threshold}%~{tier_high_floor}%）**{n_mid_show}** 个 SKU → 目标毛利率 **{tier_mid_target}%**"
                f"（**{n_mid_adj}** 个将实际降价）。"
            )
            above_df["调价层级"] = np.where(
                above_df["当前毛利率"] >= high_floor,
                f"高毛利层 → {tier_high_target}%",
                f"中毛利层 → {tier_mid_target}%",
            )
        else:
            n_still_above = int((above_df["调整后毛利率"] >= thr).sum())
            n_drop_below  = int((above_df["调整后毛利率"] < thr).sum())
            st.caption(
                f"共 **{len(above_df)}** 个高于基准线的 SKU 参与此次 **{adj_pct:+d}%** 调价。"
                f"调后仍高于基准线：**{n_still_above}** 个；跌破基准线：**{n_drop_below}** 个（需关注）。"
            )

        above_df["实际调幅%"] = (
            (above_df["调整后核算价"] - above_df["核算价"]) / above_df["核算价"] * 100
        )
        above_df["调价建议"] = above_df.apply(
            lambda r: f"¥{r['核算价']:.2f} → ¥{r['调整后核算价']:.2f}（{r['实际调幅%']:+.1f}%）",
            axis=1,
        )
        above_df = above_df.sort_values("调整后毛利率", ascending=True)

        base_cols = ["SKU编码", "品牌", "品类", "采购成本", "核算价",
                     "当前毛利率", "调整后核算价", "调整后毛利率", "调价建议", "异常"]
        if use_tiered:
            base_cols.insert(3, "调价层级")
        suggest_df = above_df[base_cols].reset_index(drop=True)

        fmt_sug = {
            "采购成本": "{:.2f}", "核算价": "{:.2f}",
            "当前毛利率": "{:.1%}", "调整后核算价": "{:.2f}", "调整后毛利率": "{:.1%}",
        }

        def above_row_color(row):
            m = row["调整后毛利率"]
            bg = "#fff0f0" if m < 0.08 else ("#fffbea" if m < thr else "#f0fff4")
            return [f"background-color:{bg};color:#1a1a1a"] * len(row)

        st.dataframe(
            suggest_df.style.apply(above_row_color, axis=1)
            .map(style_margin, subset=["当前毛利率", "调整后毛利率"])
            .format(fmt_sug),
            use_container_width=True, height=380, hide_index=True
        )
        csv_bytes = suggest_df.to_csv(index=False).encode("utf-8-sig")
        fname = (
            f"分层调价方案_{tier_high_floor}pct分界_{threshold}pct基准线.csv"
            if use_tiered
            else f"调价方案_{adj_pct:+d}pct_{threshold}pct基准线.csv"
        )
        st.download_button(
            "📋 下载调价方案清单（CSV）",
            data=csv_bytes,
            file_name=fname,
            mime="text/csv",
        )

else:
    # 筛选规则关闭：原逻辑，展示低于基准线需涨价的 SKU
    below_thr = view2[view2["调整后毛利率"] < thr].copy()

    if len(below_thr) == 0:
        st.success(f"✅ 当前筛选范围内所有 SKU 均已达到 {threshold}% 毛利基准，无需调价。")
    else:
        n_danger = int((below_thr["调整后毛利率"] < 0.08).sum())
        st.caption(
            f"共 **{len(below_thr)}** 个 SKU 低于 {threshold}% 基准线"
            f"（其中 **{n_danger}** 个危险级 <8%）。按紧迫程度排序，红色行优先处理。"
        )
        below_thr = below_thr.sort_values("调整后毛利率", ascending=True)
        below_thr["建议核算价"]   = below_thr["最低达标价格"]
        below_thr["建议调价幅度"] = (
            (below_thr["建议核算价"] - below_thr["核算价"]) / below_thr["核算价"]
        )
        below_thr["调价建议"] = below_thr.apply(
            lambda r: f"¥{r['核算价']:.2f} → ¥{r['建议核算价']:.2f}（{r['建议调价幅度']*100:+.1f}%）",
            axis=1,
        )
        suggest_df = below_thr[["SKU编码", "品牌", "品类", "采购成本", "核算价",
                                 "当前毛利率", "建议核算价", "建议调价幅度", "调价建议", "异常"]
                               ].reset_index(drop=True)

        fmt_sug = {
            "采购成本": "{:.2f}", "核算价": "{:.2f}",
            "当前毛利率": "{:.1%}", "建议核算价": "{:.2f}", "建议调价幅度": "{:+.1%}",
        }

        def urgent_row(row):
            bg = "#fff0f0" if row["当前毛利率"] < 0.08 else "#fffbea"
            return [f"background-color:{bg};color:#1a1a1a"] * len(row)

        st.dataframe(
            suggest_df.style.apply(urgent_row, axis=1)
            .map(style_margin, subset=["当前毛利率"])
            .format(fmt_sug),
            use_container_width=True, height=380, hide_index=True
        )
        csv_bytes = suggest_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "📋 下载调价建议清单（CSV）",
            data=csv_bytes,
            file_name=f"调价建议_{threshold}pct基准线.csv",
            mime="text/csv",
        )

st.divider()

# ── 功能5：带颜色的 Excel 导出 ────────────────────────────────────────────────
st.divider()

def make_excel(df_export, thr):
    wb = Workbook()
    ws = wb.active
    ws.title = "价格毛利分析"

    fill_green  = PatternFill("solid", fgColor="D4EDDA")
    fill_yellow = PatternFill("solid", fgColor="FFF3CD")
    fill_red    = PatternFill("solid", fgColor="F8D7DA")
    fill_header = PatternFill("solid", fgColor="2C3E50")
    font_header = Font(bold=True, color="FFFFFF", size=10)
    font_dark   = Font(color="1a1a1a", size=10)
    thin        = Side(style="thin", color="CCCCCC")
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)
    center      = Alignment(horizontal="center", vertical="center")

    headers = list(df_export.columns)
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = center
        cell.border = border

    margin_col = headers.index("调整后毛利率") + 1 if "调整后毛利率" in headers else None

    for ri, row in enumerate(df_export.itertuples(index=False), 2):
        for ci, val in enumerate(row, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.font = font_dark
            cell.border = border
            cell.alignment = Alignment(vertical="center")

        if margin_col:
            m_raw = df_export.iloc[ri-2]["调整后毛利率"]
            try:
                m = float(str(m_raw).strip('%')) / 100 if isinstance(m_raw, str) else float(m_raw)
            except Exception:
                m = 0
            fill = fill_red if m < 0.08 else (fill_yellow if m < thr else fill_green)
            for ci in range(1, len(headers)+1):
                ws.cell(row=ri, column=ci).fill = fill

    col_widths = {"SKU编码":18,"品类":16,"品牌":18,"采购成本":12,"核算价":12,
                  "当前毛利率":12,"调整后核算价":14,"调整后毛利率":14,
                  "最低达标价格":14,"价格缺口":12,"单位利润变化":12,"异常":8}
    for ci, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(ci)].width = col_widths.get(h, 14)
    ws.row_dimensions[1].height = 20

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

export_src = detail[["SKU编码","品类","品牌","采购成本","核算价","当前毛利率",
                      "调整后核算价","调整后毛利率","最低达标价格","价格缺口",
                      "单位利润变化","异常"]].copy()
# 格式化百分比列为数字方便 Excel 识别
for c in ["当前毛利率","调整后毛利率"]:
    export_src[c] = export_src[c].round(4)

excel_buf = make_excel(export_src, thr)
st.download_button(
    "📥 导出带格式 Excel（.xlsx）",
    data=excel_buf,
    file_name=f"价格毛利分析_调价{adj_pct:+d}pct.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
