# ============================================================
# HORIZONTAL 100% STACKED BAR:
# MUNICIPAL-LEVEL QUADRANT SHARE BY MUNICIPAL TYPE
#
# Each municipality is counted once.
# Initiatives are first aggregated by City No.
# Quadrants are defined using municipal-level medians.
# ============================================================

from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


# ============================================================
# 1. 读取initiative-level数据
# ============================================================
PROJECT_DIR = Path(__file__).resolve().parent

file_path = (
    PROJECT_DIR
    / "data"
    / "Source Data.xlsx"
)

sheet_name = "Initiative-level data"

if not file_path.exists():
    raise FileNotFoundError(
        "Input data file not found:\n"
        f"{file_path}\n\n"
        "Place the dataset at 'data/Source Data.xlsx'."
    )

df = pd.read_excel(
    file_path,
    sheet_name=sheet_name
)

# 清理列名前后的空格和换行
df.columns = (
    df.columns
    .astype(str)
    .str.strip()
    .str.replace("\n", " ", regex=False)
)


# ============================================================
# 2. 设置变量
# ============================================================
city_no_col = "City No."
city_name_col = "City Name"

# 城市等级编码
# 1 = Towns and villages
# 2 = Ordinary cities
# 3 = Core cities
# 4 = Ordinance-designated cities
type_col = "City_Level"

# Initiative-level engagement indicator
engagement_col = "Engagement_Total"

# Initiative-level integration indicator
integration_col = "Total integration index"

required_cols = [
    city_no_col,
    city_name_col,
    type_col,
    engagement_col,
    integration_col
]

missing_cols = [
    col
    for col in required_cols
    if col not in df.columns
]

if missing_cols:
    raise ValueError(
        "Initiative-level data-new 中缺少以下必要列：\n"
        + "\n".join(missing_cols)
        + "\n\n当前工作表列名为：\n"
        + "\n".join(df.columns.tolist())
    )


# ============================================================
# 3. 输出路径
# ============================================================
out_dir = PROJECT_DIR / "results"
out_dir.mkdir(parents=True, exist_ok=True)

out_path = (
    out_dir
    / "municipal_quadrant_share_by_city_level_horizontal_median.png"
)

excel_out_path = (
    out_dir
    / "municipal_quadrant_share_by_city_level_median.xlsx"
)


# ============================================================
# 4. Initiative层数据清洗
# ============================================================
numeric_cols = [
    city_no_col,
    type_col,
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

# 只保留具有城市编号和城市类型的行
# Engagement或integration缺失时，groupby mean会自动忽略
df = df.dropna(
    subset=[
        city_no_col,
        type_col
    ]
).copy()

levels_fixed = [1, 2, 3, 4]

df = df[
    df[type_col].isin(levels_fixed)
].copy()

df[city_no_col] = (
    df[city_no_col]
    .astype(int)
)

df[type_col] = (
    df[type_col]
    .astype(int)
)

print("===== Initiative-level data check =====")
print(f"清洗后的initiative数量：{len(df)}")

print("\n各自治体类型的initiative数量：")
print(
    df[type_col]
    .value_counts()
    .reindex(levels_fixed)
    .fillna(0)
    .astype(int)
)


# ============================================================
# 5. 按City No.聚合到自治体层面
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
            type_col,
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

        Mean_Total_Integration_Index_2=(
            integration_col,
            "mean"
        )
    )
)

# 恢复城市名称列名
municipal_df = municipal_df.rename(
    columns={
        "City_Name": city_name_col
    }
)

# 删除无法计算横轴或纵轴的自治体
municipal_df = municipal_df.dropna(
    subset=[
        "Mean_Engagement_Total",
        "Mean_Total_Integration_Index_2"
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

x_col = "Mean_Engagement_Total"
y_col = "Mean_Total_Integration_Index_2"
municipal_type_col = "City_Level"

print("\n===== Municipal-level data check =====")
print(f"有效自治体数量：{len(municipal_df)}")

print("\n各自治体类型的自治体数量：")
print(
    municipal_df[municipal_type_col]
    .value_counts()
    .reindex(levels_fixed)
    .fillna(0)
    .astype(int)
)


# ============================================================
# 6. 使用自治体层总体中位数划分象限
# ============================================================
# 注意：
# 每个自治体的参与和整合值仍然是该自治体内部initiatives的均值；
# 四象限分割线使用所有自治体点的总体中位数。
x_median = municipal_df[x_col].median()
y_median = municipal_df[y_col].median()

print("\n===== Municipal-level quadrant thresholds =====")
print(f"Median municipal engagement = {x_median:.4f}")
print(f"Median municipal integration = {y_median:.4f}")


def get_quadrant(row):
    """
    Q1: 高参与、高整合
    Q2: 低参与、高整合
    Q3: 低参与、低整合
    Q4: 高参与、低整合

    等于总体中位数时归入“高”组。
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


# ============================================================
# 7. 按自治体数量统计各象限
# ============================================================
quad_order = [
    "Q1",
    "Q2",
    "Q3",
    "Q4"
]

# 每个自治体只计算一次
count_table = (
    municipal_df
    .groupby(
        [
            municipal_type_col,
            "quadrant"
        ]
    )
    .size()
    .unstack(fill_value=0)
    .reindex(
        index=levels_fixed,
        columns=quad_order,
        fill_value=0
    )
)

count_table.index.name = "City_Level"

# 每种自治体类型中，四个象限的自治体比例
share_table = (
    count_table
    .div(
        count_table
        .sum(axis=1)
        .replace(0, np.nan),
        axis=0
    )
    .fillna(0)
)

percentage_table = share_table * 100

print("\n===== Municipal counts by quadrant =====")
print(count_table)

print("\n===== Municipal share by quadrant (%) =====")
print(
    percentage_table.round(2)
)


# ============================================================
# 8. 自治体类型标签和颜色
# ============================================================
LEVEL_LABELS = {
    1: "Towns and\nvillages",
    2: "Ordinary\ncities",
    3: "Core\ncities",
    4: "Ordinance-designated\ncities"
}

quad_colors = {
    "Q1": "#859CD1",
    "Q2": "#8BBE96",
    "Q3": "#D39ECB",
    "Q4": "#EFB86B"
}


# ============================================================
# 9. 绘制横向100%堆叠柱状图
# ============================================================
fig, ax = plt.subplots(
    figsize=(5.2, 7.2)
)

y_pos = np.arange(
    len(levels_fixed)
)

left = np.zeros(
    len(levels_fixed)
)

for quadrant in quad_order:

    values = share_table[
        quadrant
    ].values

    bars = ax.barh(
        y_pos,
        values,
        left=left,
        height=0.75,
        color=quad_colors[quadrant],
        edgecolor="none",
        label=quadrant
    )

    # 柱内百分比标签
    # 超过3%才显示，避免过度拥挤
    for i, (value, bar) in enumerate(
        zip(values, bars)
    ):
        if value > 0.03:
            ax.text(
                left[i] + value / 2,
                bar.get_y()
                + bar.get_height() / 2,
                f"{value * 100:.0f}%",
                ha="center",
                va="center",
                fontsize=9,
                color="black"
            )

    left += values


# ============================================================
# 10. 坐标轴设置
# ============================================================
ax.set_yticks(
    y_pos
)

ax.set_yticklabels(
    [
        LEVEL_LABELS.get(
            level,
            str(level)
        )
        for level in levels_fixed
    ],
    fontsize=10
)

ax.set_xlim(
    0,
    1.0
)

ax.set_xticks(
    [
        0,
        0.25,
        0.5,
        0.75,
        1.0
    ]
)

ax.set_xticklabels(
    [
        "0%",
        "25%",
        "50%",
        "75%",
        "100%"
    ],
    fontsize=10
)

ax.set_xlabel(
    "Share of municipalities within municipal type",
    fontsize=11
)

ax.set_ylabel(
    "Municipal type",
    fontsize=11
)

# Level 1位于最下方
# barh默认第一个值位于底部，因此不需要invert_yaxis()


# ============================================================
# 11. 添加每种类型的自治体数量
# ============================================================
level_total_counts = (
    count_table
    .sum(axis=1)
)

for i, level in enumerate(levels_fixed):

    total_n = int(
        level_total_counts.loc[level]
    )

    ax.text(
        1.10,
        y_pos[i],
        f"n={total_n}",
        transform=ax.get_yaxis_transform(),
        ha="left",
        va="center",
        fontsize=9,
        clip_on=False
    )


# ============================================================
# 12. 图框、网格和图例
# ============================================================
for spine in ax.spines.values():
    spine.set_visible(True)
    spine.set_linewidth(0.8)
    spine.set_color("black")

ax.grid(
    axis="x",
    linestyle="--",
    linewidth=0.5,
    alpha=0.25
)

ax.set_axisbelow(True)

ax.legend(
    title="Quadrant",
    frameon=False,
    loc="center left",
    bbox_to_anchor=(1.28, 0.5),
    fontsize=9,
    title_fontsize=9
)

fig.subplots_adjust(
    left=0.30,
    right=0.69,
    bottom=0.12,
    top=0.97
)


# ============================================================
# 13. 保存图片
# ============================================================
fig.savefig(
    out_path,
    dpi=600,
    bbox_inches="tight",
    pad_inches=0.20
)

plt.show()

print("\n图片已保存：")
print(out_path)


# ============================================================
# 14. 生成综合汇总表
# ============================================================
summary_table = pd.DataFrame(
    index=levels_fixed
)

summary_table.index.name = "City_Level"

summary_table["Municipal_type"] = [
    LEVEL_LABELS[level].replace("\n", " ")
    for level in levels_fixed
]

summary_table["N_municipalities"] = (
    count_table.sum(axis=1)
)

for quadrant in quad_order:

    summary_table[
        f"{quadrant}_municipality_count"
    ] = count_table[quadrant]

    summary_table[
        f"{quadrant}_share"
    ] = share_table[quadrant]

    summary_table[
        f"{quadrant}_percent"
    ] = percentage_table[quadrant]


# ============================================================
# 15. 生成象限中位数阈值表
# ============================================================
threshold_table = pd.DataFrame({
    "Variable": [
        x_col,
        y_col
    ],
    "Median_threshold": [
        x_median,
        y_median
    ]
})


# ============================================================
# 16. 保存Excel结果
# ============================================================
with pd.ExcelWriter(
    excel_out_path,
    engine="openpyxl"
) as writer:

    # 每个自治体一行及其象限
    municipal_df.to_excel(
        writer,
        sheet_name="Municipal classification",
        index=False
    )

    # 自治体类型 × 象限数量
    count_table.to_excel(
        writer,
        sheet_name="Municipal counts"
    )

    # 自治体类型 × 象限比例
    share_table.to_excel(
        writer,
        sheet_name="Municipal shares"
    )

    # 自治体类型 × 象限百分比
    percentage_table.to_excel(
        writer,
        sheet_name="Municipal percentages"
    )

    # 汇总表
    summary_table.to_excel(
        writer,
        sheet_name="Summary"
    )

    # 象限分割中位数
    threshold_table.to_excel(
        writer,
        sheet_name="Thresholds",
        index=False
    )

print("\nExcel结果已保存：")
print(excel_out_path)