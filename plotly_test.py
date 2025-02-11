import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import matplotlib.cm as cm
import matplotlib.colors as mcolors

# ========== 1. 读取 Excel 数据 ==========
excel_path = r"C:\Users\Lihanfei\Desktop\data.xlsx"
data = pd.read_excel(excel_path)

bubble_data_path = r"C:\Users\Lihanfei\Desktop\bubble_data.xlsx"
bubble_data = pd.read_excel(bubble_data_path)

# 读取表格中的文本，并转换为字符串，去除空格
bubble_texts = bubble_data.astype(str).applymap(lambda x: x.strip() if isinstance(x, str) else x).values.flatten()

# 删除所有 'nan' 和 空字符串
bubble_texts = np.array([t for t in bubble_texts if t and t.lower() != "nan"])

# 计算唯一标签及其出现次数
unique_labels, label_counts = np.unique(bubble_texts, return_counts=True)

# **修改 num_bubbles，让它等于唯一的标签数量**
num_bubbles = len(unique_labels)  # **改为使用唯一标签数量**

# ========== 2. 处理桑基图数据 ==========
columns = data.columns

# 手动定义某些节点的位置
node_positions = {
    "底盘(Chassis)": (0.4, 0.3),
    "安全系统(Safety System)": (0.4, 0.31),
    "电气系统(Electric System)": (0.4, 0.33),
    "车身(Body)": (0.4, 0.36),
    "动力系统(Power System)": (0.4, 0.4),
    "内饰系统(Interior System)": (0.4, 0.6),
    "服务(Service)": (0.4, 0.6),
    "操控界面": (0.6, 0.3),
    "信息系统": (0.6, 0.35),
    "车联功能": (0.6, 0.4),
    "多感官体验": (0.6, 0.45),
    "个性化配置": (0.6, 0.55),
    "品牌文化": (0.6, 0.7),
    "自然生态美感": (0.8, 0.3),
    "全方位可访问性": (0.8, 0.4),
    "数据驱动个性化": (0.8, 0.5),
    "智感科技融合": (0.8, 0.6),
    "身心愉悦关怀": (0.8, 0.7),
    "A 赋予“车联功能”以未来的“自然生态美感”": (1, 0.3),
    "B 从保时捷“个性化配置”发展为未来的“智感科技融合”": (1, 0.4),
    "C 从“信息系统”发展为未来的“全方位可访问性”": (1, 0.5),
    "D 从“操控界面”发展为未来的“数据驱动个性化”": (1, 0.6),
    "E 从“多感官体验”发展为未来的“身心愉悦关怀”": (1, 0.7),
}

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

# ========== 3. 计算所有节点的位置 ==========
node_x_positions = []
node_y_positions = []

for col_idx, col in enumerate(columns):
    col_nodes = data[col].dropna().unique().tolist()
    for node in col_nodes:
        if node in node_positions:
            node_x_positions.append(node_positions[node][0])
            node_y_positions.append(node_positions[node][1])
        else:
            node_x_positions.append(np.linspace(0.05, 0.95, len(columns))[col_idx])
            node_y_positions.append(None)  # 让 Plotly 自动调整 y 轴

# ========== 4. 计算气泡图的 x, y 坐标 ==========
angles = np.linspace(-np.pi / 2, np.pi / 2, num_bubbles)  # 角度分布在右半圆

# **修改 radii，让泡泡更加分散**
radii = np.random.uniform(1.5, 2.5, num_bubbles)  # **让泡泡扩散得更开**

# 计算泡泡的 x, y 坐标，并加入**随机扰动**
bubble_x = 1.5 + radii * np.cos(angles) + np.random.uniform(-0.1, 0.1, num_bubbles)
bubble_y = 0.5 + radii * np.sin(angles) + np.random.uniform(-0.1, 0.1, num_bubbles)

# 生成渐变颜色（顶部到底部）
colormap = cm.get_cmap('viridis')  # 选择渐变色方案
normalize = mcolors.Normalize(vmin=min(bubble_y), vmax=max(bubble_y))  # 归一化 y 值
bubble_colors = [mcolors.rgb2hex(colormap(normalize(y))) for y in bubble_y]  # 转换为 HEX 颜色

# 计算泡泡大小，避免 ZeroDivisionError
bubble_min, bubble_max = label_counts.min(), label_counts.max()
if bubble_max == bubble_min:
    bubble_sizes = np.full_like(label_counts, 30)  # 统一大小
else:
    bubble_sizes = (label_counts - bubble_min) / (bubble_max - bubble_min) * 50 + 10

# ========== 5. 创建融合图表 ==========
fig = make_subplots(
    rows=1, cols=2,
    column_widths=[0.7, 0.3],
    specs=[[{"type": "sankey"}, {"type": "scatter"}]],
    horizontal_spacing=0.02
)

# ========== 6. 添加桑基图 ==========
fig.add_trace(
    go.Sankey(
        node=dict(
            pad=5,
            thickness=20,
            label=[node if 75 <= i < 96 else "" for i, node in enumerate(all_nodes)],
            x=node_x_positions,
            y=node_y_positions,
            customdata=all_nodes,
            hovertemplate="%{customdata}<extra></extra>",
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

# ========== 7. 添加气泡图 ==========
fig.add_trace(
    go.Scatter(
        x=bubble_x,
        y=bubble_y,
        mode='markers',
        marker=dict(
            size=bubble_sizes,
            color=bubble_colors,
            opacity=0.8,
            line=dict(width=0)
        ),
        hovertext=unique_labels,  # 文字只在悬停时显示
        hoverinfo="text"  # 仅在 hover 时显示文本
    ),
    row=1, col=2
)

# ========== 8. 调整布局 ==========
fig.update_layout(
    title="融合图示例（桑基图 + 气泡图）",
    height=1000,
    width=2000,
    plot_bgcolor="white",
    paper_bgcolor="white",
)

# **隐藏气泡图坐标轴**
fig.update_xaxes(visible=False, row=1, col=2)
fig.update_yaxes(visible=False, row=1, col=2)

fig.update_yaxes(range=[0, 1], constrain="domain", row=1, col=1)

fig.show()

