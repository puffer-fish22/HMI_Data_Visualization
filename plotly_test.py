import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np

# 创建布局
fig = make_subplots(
    rows=1, cols=2,  
    column_widths=[0.7, 0.3],  
    specs=[[{"type": "sankey"}, {"type": "scatter"}]],  
    horizontal_spacing=0.03  
)

# 读取 Excel 数据
excel_path = r"C:\Users\Lihanfei\Desktop\data.xlsx"  
data = pd.read_excel(excel_path)

# 获取所有阶段的列名
columns = data.columns  

# 目标最后一列的正确顺序（ABCDE）
desired_order_last_column = [
    "A 赋予“车联功能”以未来的“自然生态美感”", 
    "B 从保时捷“个性化配置”发展为未来的“智感科技融合”", 
    "C 从“信息系统”发展为未来的“全方位可访问性”", 
    "D 从“操控界面”发展为未来的“数据驱动个性化”", 
    "E 从“多感官体验”发展为未来的“身心愉悦关怀”"
]

# 构建所有节点及索引
all_nodes = []
node_indices = {}
current_index = 0

for col in columns:
    for node in data[col].dropna():  # 确保不会遗漏 NaN
        if node not in node_indices:  
            node_indices[node] = current_index
            all_nodes.append(node)
            current_index += 1

# 重新获取最后一列所有唯一节点，并按照 ABCDE 排序
last_column_values = list(set(data[columns[-1]].dropna().tolist()))
ordered_last_column_nodes = [node for node in desired_order_last_column if node in last_column_values]

# 重新调整 all_nodes 的顺序，把 ABCDE 放到最后
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

# 计算所有节点数量
total_nodes = len(all_nodes)
x_positions = np.full(total_nodes, np.nan)  
y_positions = np.full(total_nodes, np.nan)

# 强制 ABCDE 在 y 轴上按顺序排列
for i, node in enumerate(ordered_last_column_nodes):
    if node in node_indices:
        x_positions[node_indices[node]] = 0.9  
        y_positions[node_indices[node]] = i / (len(ordered_last_column_nodes) - 1) if len(ordered_last_column_nodes) > 1 else 0.5

# 创建桑基图
fig.add_trace(
    go.Sankey(
        arrangement="snap",  # 关键，防止 Plotly 自动调整顺序
        node=dict(
            pad=5,  
            thickness=20,  
            label=[
                "" if i < 75 or i >= 96 else node  # 控制显示哪些标签
                for i, node in enumerate(all_nodes)
            ],
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

# 右侧气泡图
np.random.seed(0)
bubble_x = [1, 2, 3, 1, 2, 3, 1, 2, 3, 1, 2, 3]
bubble_y = np.random.uniform(-3, 3, 12)
bubble_size = np.random.randint(10, 50, 12)
bubble_color = np.random.uniform(10, 50, 12)

fig.add_trace(
    go.Scatter(
        x=bubble_x,
        y=bubble_y,
        mode="markers",
        marker=dict(
            size=bubble_size,
            color=bubble_color,
            colorscale="Viridis",
            showscale=False,
        ),
        showlegend=False
    ),
    row=1, col=2  
)

# 更新布局
fig.update_layout(
    title="融合图表示例",
    height=800,
    width=2000,
    plot_bgcolor="black",
    paper_bgcolor="black"
)

# 隐藏中间图和右侧图的坐标轴
fig.update_xaxes(visible=False, row=1, col=2)
fig.update_yaxes(visible=False, row=1, col=2)

# 显示图表
fig.show()

