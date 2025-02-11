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

# ========== 2. 处理气泡图数据（保留文件中的顺序） ==========
# 将 bubble_data 转换为字符串，并去除空格
bubble_df = bubble_data.astype(str).applymap(lambda x: x.strip() if isinstance(x, str) else "")

# 获取表格尺寸（行数、列数）
num_rows, num_cols = bubble_df.shape

# 遍历表格（从上到下、从左到右），提取唯一标签、记录第一次出现的位置以及出现次数
unique_labels = []         # 按文件顺序存放每个唯一标签
label_counts_dict = {}     # 用于统计每个标签出现的次数
label_positions = {}       # 用于记录每个唯一标签第一次出现时的 (行, 列)

for i in range(num_rows):
    for j in range(num_cols):
        text = bubble_df.iloc[i, j]
        if text and text.lower() != "nan":
            if text not in label_counts_dict:
                unique_labels.append(text)
                label_counts_dict[text] = 0
                label_positions[text] = (i, j)  # 记录第一次出现的位置
            label_counts_dict[text] += 1

# 统计各标签出现次数，并确定气泡总数
label_counts = np.array([label_counts_dict[label] for label in unique_labels])
num_bubbles = len(unique_labels)

# ========== 3. 计算气泡图的 x, y 坐标 ==========
# 根据表格列数决定“半径”：第一列对应内层，最后一列对应外层
# 这里用 np.linspace 在 [1.0, 2.5] 之间生成 num_cols 个半径值
col_radii = np.linspace(1.5, 2.5, num_cols)

# 根据表格行数决定角度：第一行对应最上方（角度 = π/2），最后一行对应最下方（角度 = -π/2）
row_angles = np.linspace(np.pi/2, -np.pi/2, num_rows)

# 为每个唯一标签（气泡）确定其半径和角度，使用它在表格中的第一次出现位置
bubble_radii = []
bubble_angles = []
for label in unique_labels:
    i, j = label_positions[label]  # i:行号, j:列号
    bubble_radii.append(col_radii[j])
    bubble_angles.append(row_angles[i])
bubble_radii = np.array(bubble_radii)
bubble_angles = np.array(bubble_angles)

# 根据极坐标计算气泡图的 x, y 坐标，并加入轻微随机扰动（避免完全重合）
bubble_x = 1.5 + bubble_radii * np.cos(bubble_angles) + np.random.uniform(-0.1, 0.1, num_bubbles)
bubble_y = 0.5 + bubble_radii * np.sin(bubble_angles) + np.random.uniform(-0.1, 0.1, num_bubbles)

# 生成渐变颜色（依据 y 值从顶部到底部变化）
colormap = cm.get_cmap('viridis')
normalize = mcolors.Normalize(vmin=min(bubble_y), vmax=max(bubble_y))
bubble_colors = [mcolors.rgb2hex(colormap(normalize(y))) for y in bubble_y]

# 根据标签出现次数计算气泡大小（出现次数越多，气泡越大）
bubble_min, bubble_max = label_counts.min(), label_counts.max()
if bubble_max == bubble_min:
    bubble_sizes = np.full_like(label_counts, 30)
else:
    bubble_sizes = (label_counts - bubble_min) / (bubble_max - bubble_min) * 50 + 15

# ========== 4. 处理桑基图数据 ==========
columns = data.columns
# 手动定义某些节点的位置
node_positions = {
    "非物质环境 (Non-material context)": (0.2, 0.33),
    "人类行动者 (Actor)": (0.2, 0.45),
    "物质元素 (Material)": (0.2, 0.67),
    "底盘(Chassis)": (0.4, 0.3),
    "安全系统(Safety System)": (0.4, 0.31),
    "电气系统(Electric System)": (0.4, 0.33),
    "车身(Body)": (0.4, 0.36),
    "动力系统(Power System)": (0.4, 0.4),
    "内饰系统(Interior System)": (0.4, 0.5),
    "服务(Service)": (0.4, 0.55),
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

# 构建所有节点及其索引
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

# 重新生成 source 和 target 的索引
source_indices = []
target_indices = []
for i in range(len(columns) - 1):
    source_col = data[columns[i]].dropna()
    target_col = data[columns[i + 1]].dropna()
    for source, target in zip(source_col, target_col):
        if source in node_indices and target in node_indices:
            source_indices.append(node_indices[source])
            target_indices.append(node_indices[target])

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
            label=["" if i < 75 or i >= 96 else node for i, node in enumerate(all_nodes)],
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
        hovertext=unique_labels,  # 鼠标悬停时显示文本
        hoverinfo="text"
    ),
    row=1, col=2
)

# ========== 8. 调整布局 ==========
fig.update_layout(
    title="融合图示例",
    height=1000,
    width=2000,
    plot_bgcolor="white",
    paper_bgcolor="white",
)

# 隐藏气泡图坐标轴
fig.update_xaxes(visible=False, row=1, col=2)
fig.update_yaxes(visible=False, row=1, col=2)
fig.update_yaxes(range=[0, 1], constrain="domain", row=1, col=1)

fig.show()



