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

# 只为最后几个节点指定位置，前面的节点留空
x_positions = [None] * total_nodes  # 先初始化为 None，让 Plotly 自动布局
y_positions = [None] * total_nodes  

# 例如：最后两个节点（第 total_nodes-2 和 total_nodes-1 个）手动设置位置
x_positions[97] = 0.9  # 让倒数第二个节点靠右
y_positions[97] = 0.3  # 控制垂直方向

x_positions[96] = 0.9  # 让最后一个节点更靠右
y_positions[96] = 0.7  # 控制垂直方向

# 创建桑基图
fig.add_trace(
    go.Sankey(
        arrangement="snap",
        node=dict(
            pad=5,  # 节点之间的间距
            thickness=20,  # 节点厚度 
            label = [
                "" if i<74 or i >=95 else node
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

fig.show()
