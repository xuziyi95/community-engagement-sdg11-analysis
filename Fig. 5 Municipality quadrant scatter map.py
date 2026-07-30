# ============================================================
# Municipal-level engagement–integration quadrant analysis
#
# Code availability version
# - Uses repository-relative paths only.
# - Aggregates initiative-level records to municipalities.
# - Divides municipalities into four quadrants using the
#   municipal-level medians of engagement and integration.
# - Produces a labeled scatter plot, quadrant composition plots,
#   and an Excel workbook containing the underlying results.
# ============================================================

from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib import rcParams
from adjustText import adjust_text
import matplotlib.patches as patches
import matplotlib.lines as mlines


# ============================================================
# 0. Portable paths and font settings
# ============================================================

# Repository root: the directory containing this script
PROJECT_DIR = Path(__file__).resolve().parent

# Expected repository structure:
# project/
# ├── municipal_quadrant_scatter_code_availability.py
# ├── data/
# │   └── Source Data.xlsx
# └── results/
FILE_PATH = PROJECT_DIR / "data" / "Source Data.xlsx"
SHEET_NAME = "Initiative-level data"
OUTPUT_DIR = PROJECT_DIR / "results"

SCATTER_OUTPUT_PATH = (
    OUTPUT_DIR
    / "municipal_quadrant_engagement_integration.png"
)

PIE_OUTPUT_PATH = (
    OUTPUT_DIR
    / "municipal_type_composition_by_quadrant.png"
)

EXCEL_OUTPUT_PATH = (
    OUTPUT_DIR
    / "municipal_quadrant_results.xlsx"
)

# Use commonly available fonts. Matplotlib will use the first installed font.
rcParams["font.family"] = "sans-serif"
rcParams["font.sans-serif"] = [
    "Arial",
    "DejaVu Sans",
    "Noto Sans CJK JP",
    "Noto Sans CJK SC",
]


# ============================================================
# 2. 读取initiative-level数据
# ============================================================
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

if not FILE_PATH.exists():
    raise FileNotFoundError(
        "Input data file not found:\n"
        f"{FILE_PATH}\n\n"
        "Place the dataset at 'data/Source Data.xlsx' before running "
        "the analysis."
    )

df = pd.read_excel(
    FILE_PATH,
    sheet_name=SHEET_NAME,
)

# 清理列名中的前后空格和换行
df.columns = (
    df.columns
    .astype(str)
    .str.strip()
    .str.replace("\n", " ", regex=False)
)


# ============================================================
# 3. 设置原始变量
# ============================================================
city_no_col = "City No."

# 使用新的城市等级编码列：
# 1 = Towns and villages
# 2 = Ordinary city
# 3 = Core city
# 4 = Ordinance-designated city
city_level_col = "City_Level"

# initiative层面的参与值
engagement_col = "Engagement_Total"

# initiative层面的整合值
integration_col = "Total integration index"

# 城市英文名称列
possible_city_name_cols = [
    "City Name"
]

city_name_col = next(
    (
        col
        for col in possible_city_name_cols
        if col in df.columns
    ),
    None
)

print(f"识别到的城市名称列：{city_name_col}")


# ============================================================
# 4. 检查必要列
# ============================================================
required_cols = [
    city_no_col,
    city_level_col,
    engagement_col,
    integration_col
]

if city_name_col is not None:
    required_cols.append(city_name_col)

missing_cols = [
    col
    for col in required_cols
    if col not in df.columns
]

if city_name_col is None:
    missing_cols.append("City Name")

if missing_cols:
    raise ValueError(
        "The worksheet is missing the following required columns：\n"
        + "\n".join(missing_cols)
        + "\n\n当前工作表列名为：\n"
        + "\n".join(df.columns.tolist())
    )


# ============================================================
# 5. 数据清洗
# ============================================================
numeric_cols = [
    city_no_col,
    city_level_col,
    engagement_col,
    integration_col
]

for col in numeric_cols:
    df[col] = pd.to_numeric(
        df[col],
        errors="coerce"
    )

df[city_name_col] = (
    df[city_name_col]
    .astype("string")
    .str.strip()
)

# 只删除City No.和City_Level缺失的行
# Engagement或integration缺失时，groupby mean会自动忽略
df = df.dropna(
    subset=[
        city_no_col,
        city_level_col
    ]
).copy()

# City_Level的有效编码为1–4
levels_order = [1, 2, 3, 4]

df = df[
    df[city_level_col].isin(levels_order)
].copy()

df[city_no_col] = (
    df[city_no_col]
    .astype(int)
)

df[city_level_col] = (
    df[city_level_col]
    .astype(int)
)

print("\n===== Initiative-level city type check =====")

print(
    df[city_level_col]
    .value_counts()
    .reindex(levels_order)
    .fillna(0)
    .astype(int)
)


# ============================================================
# 6. 按City No.聚合到城市层面
# ============================================================
municipal_df = (
    df.groupby(
        city_no_col,
        as_index=False
    )
    .agg(
        City_Name=(
            city_name_col,
            "first"
        ),

        City_Level=(
            city_level_col,
            "first"
        ),

        N_initiatives=(
            city_no_col,
            "size"
        ),

        N_valid_engagement=(
            engagement_col,
            "count"
        ),

        N_valid_integration=(
            integration_col,
            "count"
        ),

        Mean_Engagement_Total=(
            engagement_col,
            "mean"
        ),

        Mean_Total_Integration_Index=(
            integration_col,
            "mean"
        )
    )
)

# 恢复原来的城市名称列名
municipal_df = municipal_df.rename(
    columns={
        "City_Name": city_name_col
    }
)

# 删除无法计算横轴或纵轴的城市
municipal_df = municipal_df.dropna(
    subset=[
        "Mean_Engagement_Total",
        "Mean_Total_Integration_Index"
    ]
).copy()

municipal_df["City_Level"] = (
    municipal_df["City_Level"]
    .astype(int)
)

municipal_df = (
    municipal_df
    .sort_values(city_no_col)
    .reset_index(drop=True)
)


# ============================================================
# 7. 设置城市层绘图变量
# ============================================================
x_col = "Mean_Engagement_Total"
y_col = "Mean_Total_Integration_Index"
type_col = "City_Level"

print("\n===== 城市层聚合结果 =====")
print(f"原始initiative数量：{len(df)}")
print(f"聚合后的有效城市数量：{len(municipal_df)}")

print("\n城市层各类型数量：")

print(
    municipal_df[type_col]
    .value_counts()
    .reindex(levels_order)
    .fillna(0)
    .astype(int)
)

print("\n前10个城市：")

print(
    municipal_df[
        [
            city_no_col,
            city_name_col,
            type_col,
            "N_initiatives",
            "N_valid_engagement",
            "N_valid_integration",
            x_col,
            y_col
        ]
    ]
    .head(10)
    .round(4)
    .to_string(index=False)
)


# ============================================================
# 8. Calculate municipal-level quadrant thresholds using medians
# ============================================================
x_median = municipal_df[x_col].median()
y_median = municipal_df[y_col].median()

print("\n===== 象限分割值 =====")

print(
    f"Median municipal Engagement_Total：{x_median:.4f}"
)

print(
    f"Median municipal Total integration index：{y_median:.4f}"
)


# ============================================================
# 9. 象限分类
# ============================================================
def get_quadrant(row):
    """
    Q1: 高参与、高整合
    Q2: 低参与、高整合
    Q3: 低参与、低整合
    Q4: 高参与、低整合

    等于中位数的城市归入“高”组。
    """

    if (
        row[x_col] >= x_median
        and row[y_col] >= y_median
    ):
        return "Q1"

    elif (
        row[x_col] < x_median
        and row[y_col] >= y_median
    ):
        return "Q2"

    elif (
        row[x_col] < x_median
        and row[y_col] < y_median
    ):
        return "Q3"

    else:
        return "Q4"


municipal_df["quadrant"] = municipal_df.apply(
    get_quadrant,
    axis=1
)

quad_colors = {
    "Q1": "#859CD1",
    "Q2": "#8BBE96",
    "Q3": "#D39ECB",
    "Q4": "#EFB86B"
}


# ============================================================
# 10. City_Level形状、大小和名称
# ============================================================
markers = {
    1: "o",
    2: "s",
    3: "D",
    4: "*"
}

marker_size = {
    1: 90,
    2: 65,
    3: 80,
    4: 120
}

type_labels = {
    1: "Towns and villages",
    2: "Ordinary cities",
    3: "Core cities",
    4: "Ordinance-designated cities"
}


# ============================================================
# 11. 图一：城市层面四象限散点图
# ============================================================
fig, ax = plt.subplots(
    figsize=(10, 8)
)

texts = []

for quadrant, color in quad_colors.items():

    quadrant_data = municipal_df[
        municipal_df["quadrant"] == quadrant
    ]

    for level in levels_order:

        subset = quadrant_data[
            quadrant_data[type_col] == level
        ]

        if subset.empty:
            continue

        ax.scatter(
            subset[x_col],
            subset[y_col],
            s=marker_size[level],
            color=color,
            marker=markers[level],
            alpha=0.82,
            edgecolor="black",
            linewidth=0.6,
            zorder=3
        )

        # 添加城市名称
        for _, row in subset.iterrows():

            text = ax.text(
                row[x_col],
                row[y_col],
                str(row[city_name_col]),
                fontsize=8,
                ha="left",
                va="bottom",
                zorder=4
            )

            texts.append(text)


# ============================================================
# 12. 象限分割线：使用城市中位数
# ============================================================
ax.axvline(
    x_median,
    color="gray",
    linestyle="--",
    linewidth=1.0,
    zorder=1
)

ax.axhline(
    y_median,
    color="gray",
    linestyle="--",
    linewidth=1.0,
    zorder=1
)


# ============================================================
# 13. 城市名称自动避让
# ============================================================
adjust_text(
    texts,
    ax=ax,

    # 增大文字之间、文字与数据点之间的避让范围
    expand_text=(1.35, 1.55),
    expand_points=(1.45, 1.65),

    # 增强文字之间和文字与点之间的排斥力
    force_text=(0.8, 1.0),
    force_points=(0.7, 0.9),

    # 允许文字被拉得更开
    force_pull=(0.01, 0.01),

    # 防止单次移动太小
    max_move=(20, 25),

    # 增加迭代次数
    iter_lim=1500,

    # 允许上下左右移动
    only_move={
        "text": "xy",
        "static": "xy",
        "explode": "xy",
        "pull": "xy"
    },

    arrowprops=dict(
        arrowstyle="-",
        color="gray",
        lw=0.45,
        alpha=0.65
    )
)


# ============================================================
# 14. 坐标轴范围
# ============================================================
# 横轴最大值扩展至2.2
# 在0.0左侧额外保留少量空间
ax.set_xlim(
    0.0,
    2.8
)

# 纵轴范围扩展至8.5
ax.set_ylim(
    0.6,
    8.5
)


# ============================================================
# 15. 象限文字说明
# ============================================================
ax.text(
    0.97,
    0.97,
    "Q1: High engagement\nHigh integration",
    transform=ax.transAxes,
    ha="right",
    va="top",
    fontsize=9
)

ax.text(
    0.03,
    0.97,
    "Q2: Low engagement\nHigh integration",
    transform=ax.transAxes,
    ha="left",
    va="top",
    fontsize=9
)

ax.text(
    0.03,
    0.03,
    "Q3: Low engagement\nLow integration",
    transform=ax.transAxes,
    ha="left",
    va="bottom",
    fontsize=9
)

ax.text(
    0.97,
    0.03,
    "Q4: High engagement\nLow integration",
    transform=ax.transAxes,
    ha="right",
    va="bottom",
    fontsize=9
)


# ============================================================
# 16. 图例
# ============================================================
city_level_handles = [
    mlines.Line2D(
        [],
        [],
        color="white",
        marker=markers[level],
        markerfacecolor="#BDBDBD",
        markeredgecolor="black",
        markeredgewidth=0.6,
        markersize=9,
        linestyle="None",
        label=type_labels[level]
    )
    for level in levels_order
]

quadrant_handles = [
    patches.Patch(
        facecolor=quad_colors[quadrant],
        edgecolor="black",
        linewidth=0.5,
        label=quadrant
    )
    for quadrant in [
        "Q1",
        "Q2",
        "Q3",
        "Q4"
    ]
]


ax.legend(
    handles=quadrant_handles,
    title="Quadrant",
    loc="upper left",
    bbox_to_anchor=(1.02, 0.58),
    frameon=False,
    fontsize=9,
    title_fontsize=10
)
fig.subplots_adjust(
    left=0.10,
    right=0.72,
    bottom=0.11,
    top=0.96
)

# ============================================================
# 17. 坐标轴和样式
# ============================================================
ax.set_xlabel(
    "Mean Community Engagement Score",
    fontsize=12
)

ax.set_ylabel(
    "Mean Total Integration Index",
    fontsize=12
)

ax.grid(
    alpha=0.25,
    zorder=0
)

for spine in ax.spines.values():
    spine.set_visible(True)
    spine.set_color("black")
    spine.set_linewidth(0.8)

plt.tight_layout()

plt.savefig(
    SCATTER_OUTPUT_PATH,
    dpi=600,
    bbox_inches="tight"
)

plt.close(fig)

print("\n象限散点图已保存：")
print(SCATTER_OUTPUT_PATH)


# ============================================================
# 18. 图二：四象限城市类型构成饼图
# ============================================================
fig = plt.figure(
    figsize=(10, 8)
)

# 4种城市等级对应4种灰度
pie_colors = {
    1: "#DFDFDF",
    2: "#C5C5C5",
    3: "#A5A5A5",
    4: "#808080"
}

positions = {
    "Q1": (0.70, 0.70),
    "Q2": (0.20, 0.70),
    "Q3": (0.20, 0.20),
    "Q4": (0.70, 0.20)
}

for quadrant, (px, py) in positions.items():

    subset = municipal_df[
        municipal_df["quadrant"] == quadrant
    ]

    inset_ax = fig.add_axes(
        [px, py, 0.20, 0.20]
    )

    inset_ax.set_axis_off()

    if not subset.empty:

        counts = (
            subset[type_col]
            .value_counts()
            .reindex(levels_order)
            .fillna(0)
            .astype(int)
        )

        sizes = counts.values

        total = counts.sum()

        labels = [
            (
                f"{value / total:.0%}"
                if value > 0 and total > 0
                else ""
            )
            for value in sizes
        ]

        inset_ax.pie(
            sizes,
            labels=labels,
            colors=[
                pie_colors[level]
                for level in levels_order
            ],
            textprops={
                "fontsize": 8
            },
            wedgeprops={
                "edgecolor": "black",
                "linewidth": 0.5
            },
            labeldistance=1.05
        )

    inset_ax.set_title(
        f"{quadrant} (n={len(subset)})",
        fontsize=12,
        pad=2
    )


legend_elements = [
    patches.Patch(
        facecolor=pie_colors[level],
        edgecolor="gray",
        label=type_labels[level]
    )
    for level in levels_order
]

fig.legend(
    handles=legend_elements,
    loc="center",
    fontsize=10,
    title="Municipal type",
    title_fontsize=11,
    frameon=True,
    edgecolor="black",
    bbox_to_anchor=(0.5, 0.50)
)

fig.suptitle(
    "Municipal-type composition within each quadrant",
    fontsize=14
)

plt.savefig(
    PIE_OUTPUT_PATH,
    dpi=600,
    bbox_inches="tight"
)

plt.close(fig)

print("\n象限城市类型构成图已保存：")
print(PIE_OUTPUT_PATH)


# ============================================================
# 19. 象限统计
# ============================================================
quadrant_counts = (
    municipal_df["quadrant"]
    .value_counts()
    .reindex(
        ["Q1", "Q2", "Q3", "Q4"]
    )
    .fillna(0)
    .astype(int)
)

quadrant_level_table = (
    pd.crosstab(
        municipal_df["quadrant"],
        municipal_df[type_col]
    )
    .reindex(
        index=[
            "Q1",
            "Q2",
            "Q3",
            "Q4"
        ],
        columns=levels_order,
        fill_value=0
    )
)

# 将城市等级列改为英文名称，方便Excel查看
quadrant_level_table_named = (
    quadrant_level_table.rename(
        columns=type_labels
    )
)

print("\n===== 各象限城市数量 =====")
print(quadrant_counts)

print("\n===== 象限 × City_Level =====")
print(quadrant_level_table)

print("\n===== 象限 × Municipal type =====")
print(quadrant_level_table_named)


# ============================================================
# 20. 生成阈值表
# ============================================================
threshold_df = pd.DataFrame({
    "Threshold": [
        "Median of municipal Engagement_Total means",
        "Median of municipal Total integration index means"
    ],
    "Value": [
        x_median,
        y_median
    ]
})


# ============================================================
# 21. 生成城市等级对应表
# ============================================================
city_level_mapping_df = pd.DataFrame({
    "City_Level": levels_order,
    "Municipal_type": [
        type_labels[level]
        for level in levels_order
    ]
})


# ============================================================
# 22. 将城市聚合结果合并回initiative数据，便于核对
# ============================================================
initiative_check_df = df.merge(
    municipal_df[
        [
            city_no_col,
            "N_initiatives",
            "N_valid_engagement",
            "N_valid_integration",
            x_col,
            y_col,
            "quadrant"
        ]
    ],
    on=city_no_col,
    how="left"
)


# ============================================================
# 23. 保存Excel结果
# ============================================================
with pd.ExcelWriter(
    EXCEL_OUTPUT_PATH,
    engine="openpyxl"
) as writer:

    municipal_df.to_excel(
        writer,
        sheet_name="Municipal aggregation",
        index=False
    )

    initiative_check_df.to_excel(
        writer,
        sheet_name="Initiative check",
        index=False
    )

    threshold_df.to_excel(
        writer,
        sheet_name="Thresholds",
        index=False
    )

    city_level_mapping_df.to_excel(
        writer,
        sheet_name="City level mapping",
        index=False
    )

    quadrant_level_table_named.to_excel(
        writer,
        sheet_name="Quadrant by City Level"
    )

print("\n结果Excel已保存：")
print(EXCEL_OUTPUT_PATH)