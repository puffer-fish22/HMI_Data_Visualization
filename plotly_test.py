import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np

#创建布局
fig = make_subplots(
    rows=1, cols=2,  # 设置1行3列布局
    column_widths=[0.7, 0.3],  # 每列宽度比例
    specs=[[{"type": "sankey"}, {"type": "scatter"}]],  # 定义每列的图表类型
    horizontal_spacing=0.03  # 控制子图之间的间距
)

# 读取 Excel 数据
excel_path = r"C:\Users\Lihanfei\Desktop\data.xlsx"  # 替换为你的 Excel 文件路径
data = pd.read_excel(excel_path)



# 获取所有阶段的列名
columns = data.columns  # 所有阶段列名

# 构建所有节点及索引（按顺序）
all_nodes = []
node_indices = {}
current_index = 0

for col in columns:
    for node in data[col]:
        if node not in node_indices:  # 确保节点唯一性
            node_indices[node] = current_index
            all_nodes.append(node)
            current_index += 1

# 创建 source 和 target 索引
source_indices = []
target_indices = []

for i in range(len(columns) - 1):  # 遍历相邻阶段
    source_col = data[columns[i]]
    target_col = data[columns[i + 1]]
    for source, target in zip(source_col, target_col):
        source_indices.append(node_indices[source])
        target_indices.append(node_indices[target])

# 定义节点颜色（可按需求调整）
node_colors = ["rgba(70, 130, 180, 0.8)" for _ in all_nodes]

# 定义连接的权重（可以统一为 1 或自定义）
values = [1 for _ in source_indices]

# 计算所有节点数量
total_nodes = len(all_nodes)

# 先创建所有节点的默认位置，未指定的让 Plotly 自动计算
x_positions = np.full(total_nodes, np.nan)  # 让 Plotly 处理未指定的位置
y_positions = np.full(total_nodes, np.nan)

# 获取数据中最后一列所有出现的节点
last_column_values = data[columns[-1]].unique().tolist()

desired_order_last_column = [
    "A 赋予“车联功能”以未来的“自然生态美感”", 
    "B 从保时捷“个性化配置”发展为未来的“智感科技融合”", 
    "C 从“信息系统”发展为未来的“全方位可访问性”", 
    "D 从“操控界面”发展为未来的“数据驱动个性化”", 
    "E 从“多感官体验”发展为未来的“身心愉悦关怀”"
]

# 确保最后一列节点按 ABCDE 的顺序排列
ordered_last_column_nodes = [node for node in desired_order_last_column if node in last_column_values]

# 重新调整 all_nodes 的顺序，把 ABCDE 放到最后
remaining_nodes = [node for node in all_nodes if node not in ordered_last_column_nodes]
all_nodes = remaining_nodes + ordered_last_column_nodes

# 重新构建索引
node_indices = {node: idx for idx, node in enumerate(all_nodes)}


# 重新映射 source 和 target 索引
source_indices = []
target_indices = []
for i in range(len(columns) - 1):  # 遍历相邻阶段
    source_col = data[columns[i]]
    target_col = data[columns[i + 1]]
    for source, target in zip(source_col, target_col):
        if source in node_indices and target in node_indices:
            source_indices.append(node_indices[source])
            target_indices.append(node_indices[target])


# 让 ABCDE 按照顺序在 y 轴均匀分布
for i, node in enumerate(ordered_last_column_nodes):  # ABCDE 按顺序分配 y 值
    if node in node_indices:
        x_positions[node_indices[node]] = 0.9  # 最后一列靠右
        y_positions[node_indices[node]] = i / (len(ordered_last_column_nodes) - 1) if len(ordered_last_column_nodes) > 1 else 0.5


# 创建桑基图
fig.add_trace(
    go.Sankey(
        arrangement="snap",
        node=dict(
            pad=5,  # 节点之间的间距
            thickness=20,  # 节点厚度 
            label = [
                "" if i<75 or i >=96 else node
                for i, node in enumerate(all_nodes)
            ],

            customdata=all_nodes,  # 节点名称存储在 customdata 中
            hovertemplate="%{customdata}<extra></extra>",  # 悬停时显示名称
            
        ),
        link=dict(
            source=source_indices,
            target=target_indices,
            value=values,
            color="rgba(128, 128, 128, 0.5)"  # 设置连接线颜色
        )
    ),
    row=1, col=1  # 指定位置为第1行第1列
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
    row=1, col=2  # 指定位置为第1行第3列
)

fig.update_layout(
    title_text="Sankey Diagram with Fixed Last Nodes",
    font_size=10
)


# 隐藏中间图和右侧图的坐标轴
fig.update_xaxes(visible=False, row=1, col=2)
fig.update_yaxes(visible=False, row=1, col=2)


# 更新布局
fig.update_layout(
    title="融合图表示例",
    height=800,
    width=2000,
    plot_bgcolor="black",
    paper_bgcolor="black"
)

