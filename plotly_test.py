import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import matplotlib.cm as cm
import matplotlib.colors as mcolors

# ========== 1. 读取 Excel 数据 ==========
excel_path = r"C:\\Users\\Lihanfei\\Desktop\\data.xlsx"
data = pd.read_excel(excel_path)

# ========== 2. 处理桑基图数据 ==========
columns = data.columns
all_nodes = []
node_indices = {}
current_index = 0

for col in columns:
    for node in data[col].dropna():
        if node not in node_indices:
            node_indices[node] = current_index
            all_nodes.append(node)
            current_index += 1

# 计算每列节点的位置
node_y_positions = [0] * len(all_nodes)
node_x_positions = [0] * len(all_nodes)

for col_idx, col in enumerate(columns):
    col_nodes = data[col].dropna().unique().tolist()
    num_nodes = len(col_nodes)

    # **让 y 轴以 0.5 为中心扩展**
    min_y, max_y = 0.3, 0.7
    if num_nodes > 10:
        min_y, max_y = 0.2, 0.8
    if num_nodes > 20:
        min_y, max_y = 0.1, 0.9

    # **均匀分布节点，加入随机扰动**
    y_positions = np.linspace(min_y, max_y, num_nodes) + np.random.uniform(0, 0.04, num_nodes)

    # **计算 x 轴**
    x_position = col_idx / (len(columns) - 1)

    for i, node in enumerate(col_nodes):
        if node in node_indices:
            node_y_positions[node_indices[node]] = y_positions[i]
            node_x_positions[node_indices[node]] = x_position

# 生成桑基图的 source 和 target
source_indices = []
target_indices = []
for i in range(len(columns) - 1):
    source_col = data[columns[i]].dropna()
    target_col = data[columns[i + 1]].dropna()
    for source, target in zip(source_col, target_col):
        if source in node_indices and target in node_indices:
            source_indices.append(node_indices[source])
            target_indices.append(node_indices[target])

# ========== 3. 创建图表 ==========
fig = make_subplots(
    rows=1, cols=2,
    column_widths=[0.7, 0.3],
    specs=[[{"type": "sankey"}, {"type": "scatter"}]],
    horizontal_spacing=0.02
)

# ========== 4. 添加桑基图 ==========
fig.add_trace(
    go.Sankey(
        domain=dict(x=[0.05, 0.55], y=[0.0, 1.0]),
        node=dict(
            pad=5,
            thickness=20,
            label=all_nodes,
            x=node_x_positions,
            y=node_y_positions,
        ),
        link=dict(
            source=source_indices,
            target=target_indices,
            value=[1] * len(source_indices),
            color="rgba(128, 128, 128, 0.5)"
        )
    ),
    row=1, col=1
)

# **🚀 关键修正：强制 Plotly 不调整 y 轴**
fig.update_yaxes(range=[0, 1], constrain="domain", row=1, col=1)

# ========== 5. 右侧半圆泡泡图 ==========
num_bubbles = 50
angles = np.linspace(-np.pi / 2, np.pi / 2, num_bubbles)
radii = np.random.uniform(1.3, 1.8, num_bubbles)

bubble_x = 1.5 + radii * np.cos(angles) + np.random.uniform(0, 0.1, num_bubbles)
bubble_y = 0.5 + radii * np.sin(angles) + np.random.uniform(0, 0.1, num_bubbles)

# 生成渐变颜色
colormap = cm.get_cmap('viridis')
normalize = mcolors.Normalize(vmin=min(bubble_y), vmax=max(bubble_y))
bubble_sizes = np.random.randint(15, 50, size=num_bubbles)
bubble_labels = [chr(65 + i % 5) for i in range(num_bubbles)]
bubble_colors = [mcolors.rgb2hex(colormap(normalize(y))) for y in bubble_y]

# ========== 6. 添加右侧泡泡图 ==========
fig.add_trace(
    go.Scatter(
        x=bubble_x,
        y=bubble_y,
        mode='markers+text',
        marker=dict(size=bubble_sizes, color=bubble_colors, opacity=0.8, line=dict(width=0)),
        text=bubble_labels,
        textposition='middle center',
    ),
    row=1, col=2
)

# ========== 7. 更新布局 ==========
fig.update_layout(
    title="融合图示例（优化 Y 轴分布 + 0.5 为中心扩展）",
    height=1000,
    width=2000,
    plot_bgcolor="white",
    paper_bgcolor="white",
)

fig.update_xaxes(visible=False, row=1, col=2)
fig.update_yaxes(visible=False, row=1, col=2)

# 显示图表
fig.show()
