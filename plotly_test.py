import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import matplotlib.cm as cm
import matplotlib.colors as mcolors

# ========== 1. 读取 Excel 数据 ==========
excel_path = r"C:\Users\Lihanfei\Desktop\data.xlsx"
data = pd.read_excel(excel_path)

# ========== 2. 处理桑基图数据 ==========
columns = data.columns  # 获取所有阶段的列名

# 目标最后一列的正确顺序（ABCDE）
desired_order_last_column = [
    "A 赋予“车联功能”以未来的“自然生态美感”", 
    "B 从保时捷“个性化配置”发展为未来的“智感科技融合”", 
    "C 从“信息系统”发展为未来的“全方位可访问性”",
    "E 从“多感官体验”发展为未来的“身心愉悦关怀”", 
    "D 从“操控界面”发展为未来的“数据驱动个性化”"
]

# 构建所有节点及索引
all_nodes = []
node_indices = {}
current_index = 0

for col in columns:
    for node in data[col].dropna():
        if node not in node_indices:
            node_indices[node] = current_index
            all_nodes.append(node)
            current_index += 1

# 重新调整 all_nodes 的顺序，把 ABCDE 放到最后
last_column_values = list(set(data[columns[-1]].dropna().tolist()))
ordered_last_column_nodes = [node for node in desired_order_last_column if node in last_column_values]
remaining_nodes = [node for node in all_nodes if node not in ordered_last_column_nodes]
all_nodes = remaining_nodes + ordered_last_column_nodes

# 重新构建索引
node_indices = {node: idx for idx, node in enumerate(all_nodes)}

# 重新生成 source_indices 和 target_indices
source_indices = []
target_indices = []
for i in range(len(columns) - 1):
    source_col = data[columns[i]].dropna()
    target_col = data[columns[i + 1]].dropna()
    for source, target in zip(source_col, target_col):
        if source in node_indices and target in node_indices:
            source_indices.append(node_indices[source])
            target_indices.append(node_indices[target])

# ========== 3. 创建图表布局 ==========
fig = make_subplots(
    rows=1, cols=2,
    column_widths=[0.7, 0.3],
    specs=[[{"type": "sankey"}, {"type": "scatter"}]],
    horizontal_spacing=0.02
)

# ========== 4. 添加桑基图 ==========
fig.add_trace(
    go.Sankey(
        domain=dict(x=[0.05, 0.55], y=[0.3, 1.0]),  # 调整桑基图整体位置，使其上移
        node=dict(
            pad=5,
            thickness=20,
            label=[
                "" if i < 75 or i >= 96 else node
                for i, node in enumerate(all_nodes)
            ],
            customdata=all_nodes,
            hovertemplate="%{customdata}<extra></extra>",
            y=np.linspace(0.3, 1.0, len(all_nodes))  # 让节点整体往上移
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

# ========== 5. 生成右侧半圆泡泡数据 ==========
num_bubbles = 50  # 泡泡数量
angles = np.linspace(-np.pi / 2, np.pi / 2, num_bubbles)  # 角度分布在右半圆
radii = np.random.uniform(1.3, 1.8, num_bubbles)  # 让泡泡分散得更开

# 计算泡泡的 x, y 坐标，使其排列成半圆形
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
    title="融合图表示例",
    height=1000,
    width=2000,
    plot_bgcolor="white",
    paper_bgcolor="white",
)

# 隐藏右侧图的坐标轴
fig.update_xaxes(visible=False, row=1, col=2)
fig.update_yaxes(visible=False, row=1, col=2)

# 显示图表
fig.show()