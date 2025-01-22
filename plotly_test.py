import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np

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

# 创建桑基图
fig = make_subplots(
    rows=1, cols=3,  # 设置1行3列布局
    column_widths=[0.5, 0.3, 0.2],  # 每列宽度比例
    specs=[[{"type": "sankey"}, {"type": "scatter"}, {"type": "scatter"}]],  # 定义每列的图表类型
    horizontal_spacing=0.03  # 控制子图之间的间距
)

fig.add_trace(
    go.Sankey(
        node=dict(
            pad=5,  # 节点之间的间距
            thickness=20,  # 节点厚度
            label=["" if i < len(data[columns[0]]) else node for i, node in enumerate(all_nodes)],  # 隐藏第一列节点
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




# 定义分支数据
branches = [
    {"x": [0, 1], "y": [0, 0], "color": "teal", "hoverinfo": ["C 从“信息系统”发展为未来的“全方位可访问性”", ""]},   # 起始分支
    {"x": [1, 1.5, 2], "y": [0, -0.4, -0.45], "color": "teal", "hoverinfo": ["", "", None]},   # 第一条分支
    {"x": [1, 1.5, 2], "y": [0, -0.3, -0.35], "color": "teal", "hoverinfo": ["", "", None]},  # 第二条分支
    {"x": [1, 1.5, 2], "y": [0, -0.2, -0.25], "color": "teal", "hoverinfo": ["", "", None]},  # 第三条分支
    {"x": [1, 1.5, 2], "y": [0, -0.1, -0.15], "color": "teal", "hoverinfo": ["", "", None]},  # 第四条分支
    {"x": [1, 1.5, 2], "y": [0, 0, 0], "color": "teal", "hoverinfo": ["", "", None]},  # 第五条分支
    {"x": [1, 1.5, 2], "y": [0, 0.1, 0.15], "color": "teal", "hoverinfo": ["", "", None]},  # 第六条分支
    {"x": [1, 1.5, 2], "y": [0, 0.2, 0.25], "color": "teal", "hoverinfo": ["", "", None]},  # 第七条分支
    {"x": [1, 1.5, 2], "y": [0, 0.3, 0.35], "color": "teal", "hoverinfo": ["", "", None]},  # 第八条分支
    {"x": [1, 1.5, 2], "y": [0, 0.4, 0.45], "color": "teal", "hoverinfo": ["", "", None]},  # 第九条分支
    
    {"x": [0, 1], "y": [1, 1], "color": "green", "hoverinfo": ["B 从保时捷“个性化配置”发展为未来的“智感科技融合”", ""]},   # 起始分支
    {"x": [1, 1.5, 2], "y": [1, 0.6, 0.55], "color": "green", "hoverinfo": ["", "", None]},   # 第一条分支
    {"x": [1, 1.5, 2], "y": [1, 0.7, 0.65], "color": "green", "hoverinfo": ["", "", None]},  # 第二条分支
    {"x": [1, 1.5, 2], "y": [1, 0.8, 0.75], "color": "green", "hoverinfo": ["", "", None]},  # 第三条分支
    {"x": [1, 1.5, 2], "y": [1, 0.9, 0.85], "color": "green", "hoverinfo": ["", "", None]},  # 第四条分支
    {"x": [1, 1.5, 2], "y": [1, 1, 1], "color": "green", "hoverinfo": ["", "", None]},  # 第五条分支
    {"x": [1, 1.5, 2], "y": [1, 1.1, 1.15], "color": "green", "hoverinfo": ["", "", None]},  # 第六条分支
    {"x": [1, 1.5, 2], "y": [1, 1.2, 1.25], "color": "green", "hoverinfo": ["", "", None]},  # 第七条分支
    {"x": [1, 1.5, 2], "y": [1, 1.3, 1.35], "color": "green", "hoverinfo": ["", "", None]},  # 第八条分支
    {"x": [1, 1.5, 2], "y": [1, 1.4, 1.45], "color": "green", "hoverinfo": ["", "", None]},  # 第九条分支
    
    {"x": [0, 1], "y": [2,2], "color": "yellow", "hoverinfo": ["A 赋予“车联功能”以未来的“自然生态美感”", ""]},   # 起始分支
    {"x": [1, 1.5, 2], "y": [2, 1.6, 1.55], "color": "yellow", "hoverinfo": ["", "", None]},   # 第一条分支
    {"x": [1, 1.5, 2], "y": [2, 1.7, 1.65], "color": "yellow", "hoverinfo": ["", "", None]},  # 第二条分支
    {"x": [1, 1.5, 2], "y": [2, 1.8, 1.75], "color": "yellow", "hoverinfo": ["", "", None]},  # 第三条分支
    {"x": [1, 1.5, 2], "y": [2, 1.9, 1.85], "color": "yellow", "hoverinfo": ["", "", None]},  # 第四条分支
    {"x": [1, 1.5, 2], "y": [2, 2, 2], "color": "yellow", "hoverinfo": ["", "", None]},  # 第五条分支
    {"x": [1, 1.5, 2], "y": [2, 2.1, 2.15], "color": "yellow", "hoverinfo": ["", "", None]},  # 第六条分支
    {"x": [1, 1.5, 2], "y": [2, 2.2, 2.25], "color": "yellow", "hoverinfo": ["", "", None]},  # 第七条分支
    {"x": [1, 1.5, 2], "y": [2, 2.3, 2.35], "color": "yellow", "hoverinfo": ["", "", None]},  # 第八条分支
    {"x": [1, 1.5, 2], "y": [2, 2.4, 2.45], "color": "yellow", "hoverinfo": ["", "", None]},  # 第九条分支

    {"x": [0, 1], "y": [-1, -1], "color": "blue", "hoverinfo": ["D 从“操控界面”发展为未来的“数据驱动个性化”", ""]},   # 起始分支
    {"x": [1, 1.5, 2], "y": [-1, -1.4, -1.45], "color": "blue", "hoverinfo": ["", "", None]},   # 第一条分支
    {"x": [1, 1.5, 2], "y": [-1, -1.3, -1.35], "color": "blue", "hoverinfo": ["", "", None]},  # 第二条分支
    {"x": [1, 1.5, 2], "y": [-1, -1.2, -1.25], "color": "blue", "hoverinfo": ["", "", None]},  # 第三条分支
    {"x": [1, 1.5, 2], "y": [-1, -1.1, -1.15], "color": "blue", "hoverinfo": ["", "", None]},  # 第四条分支
    {"x": [1, 1.5, 2], "y": [-1, -1, -1], "color": "blue", "hoverinfo": ["", "", None]},  # 第五条分支
    {"x": [1, 1.5, 2], "y": [-1, -0.9, -0.85], "color": "blue", "hoverinfo": ["", "", None]},  # 第六条分支
    {"x": [1, 1.5, 2], "y": [-1, -0.8, -0.75], "color": "blue", "hoverinfo": ["", "", None]},  # 第七条分支
    {"x": [1, 1.5, 2], "y": [-1, -0.7, -0.65], "color": "blue", "hoverinfo": ["", "", None]},  # 第八条分支
    {"x": [1, 1.5, 2], "y": [-1, -0.6, -0.55], "color": "blue", "hoverinfo": ["", "", None]},  # 第九条分支

    {"x": [0, 1], "y": [-2, -2], "color": "purple", "hoverinfo": ["E 从“多感官体验”发展为未来的“身心愉悦关怀”", ""]},   # 起始分支
    {"x": [1, 1.5, 2], "y": [-2, -2.4, -2.45], "color": "purple", "hoverinfo": ["", "", None]},   # 第一条分支
    {"x": [1, 1.5, 2], "y": [-2, -2.3, -2.35], "color": "purple", "hoverinfo": ["", "", None]},  # 第二条分支
    {"x": [1, 1.5, 2], "y": [-2, -2.2, -2.25], "color": "purple", "hoverinfo": ["", "", None]},  # 第三条分支
    {"x": [1, 1.5, 2], "y": [-2, -2.1, -2.15], "color": "purple", "hoverinfo": ["", "", None]},  # 第四条分支
    {"x": [1, 1.5, 2], "y": [-2, -2, -2], "color": "purple", "hoverinfo": ["", "", None]},  # 第五条分支
    {"x": [1, 1.5, 2], "y": [-2, -1.9, -1.85], "color": "purple", "hoverinfo": ["", "", None]},  # 第六条分支
    {"x": [1, 1.5, 2], "y": [-2, -1.8, -1.75], "color": "purple", "hoverinfo": ["", "", None]},  # 第七条分支
    {"x": [1, 1.5, 2], "y": [-2, -1.7, -1.65], "color": "purple", "hoverinfo": ["", "", None]},  # 第八条分支
    {"x": [1, 1.5, 2], "y": [-2, -1.6, -1.55], "color": "purple", "hoverinfo": ["", "", None]},  # 第九条分支
]

# 添加分支曲线
for branch in branches:
    fig.add_trace(
        go.Scatter(
            x=branch["x"],
            y=branch["y"],
            mode="lines+markers",  # 同时显示线和点
            line=dict(
                color=branch["color"],
                shape="spline",  # 设置线为样条曲线
                width=1,
                
            ),
            marker=dict(size=0.001),  # 设置点的大小
            hovertemplate=[f"{info}<extra></extra>" if info else "<extra></extra>" for info in branch["hoverinfo"]],
            showlegend=False  # 隐藏图例
        
        ),

    row=1, col=2 # 指定位置为第1行第2列
        
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
            showscale=True,
            colorbar=dict(
                x=0.48,  # 设置颜色比例尺的位置
                y=0.5,  # 设置颜色比例尺的y位置
                len=1,  # 设置颜色比例尺的长度
                thickness=5,  # 设置颜色比例尺的宽度
                outlinewidth=0,  # 去掉描边
                tickvals=[],  # 不显示刻度值
                ticktext=[]  # 不显示刻度文本
            )
        ),
        showlegend=False
    ),
    row=1, col=3  # 指定位置为第1行第3列
)

# 隐藏中间图和右侧图的坐标轴
fig.update_xaxes(visible=False, row=1, col=2)
fig.update_yaxes(visible=False, row=1, col=2)
fig.update_xaxes(visible=False, row=1, col=3)
fig.update_yaxes(visible=False, row=1, col=3)

# 更新布局
fig.update_layout(
    title="融合图表示例",
    height=800,
    width=2000,
    plot_bgcolor="black",
    paper_bgcolor="black"
)

fig.show()
