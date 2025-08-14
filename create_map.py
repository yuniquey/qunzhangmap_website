import pandas as pd
import folium

# --- 配置 ---
EXCEL_FILE_PATH = '小区数据.xlsx'
SHEET_NAME = '小区经纬度'

# GPS 和名称列
NORTH_LAT_COL = '北边GPS纬度'
SOUTH_LAT_COL = '南边GPS纬度'
EAST_LON_COL = '东边GPS经度'
WEST_LON_COL = '西边GPS经度'
NAME_COLUMN = '小区名称'

# 新增的数据列 (已更正表头)
USER_COUNT_COL = '实装数'
INCIDENT_COUNT_COL = '群障次数'

# 新的输出文件名
OUTPUT_MAP_FILE = '小区群障地图.html'

# --- 主程序 ---
try:
    print(f"正在读取 Excel 文件: {EXCEL_FILE_PATH}...")
    df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=SHEET_NAME)
    print("文件读取成功！")

    # 定义绘图所需的基本列
    required_cols = [NORTH_LAT_COL, SOUTH_LAT_COL, EAST_LON_COL, WEST_LON_COL, NAME_COLUMN]
    df.dropna(subset=required_cols, inplace=True)
    
    # 将群障次数和用户数的空白格填充为0，便于处理
    df[INCIDENT_COUNT_COL] = df[INCIDENT_COUNT_COL].fillna(0)
    df[USER_COUNT_COL] = df[USER_COUNT_COL].fillna(0)

    if df.empty:
        print("错误：找不到任何有效的GPS数据。")
    else:
        # 计算中心点
        print("正在计算小区中心点...")
        df['center_lat'] = (df[SOUTH_LAT_COL] + df[NORTH_LAT_COL]) / 2
        df['center_lon'] = (df[WEST_LON_COL] + df[EAST_LON_COL]) / 2

        map_center = [df['center_lat'].iloc[0], df['center_lon'].iloc[0]]
        m = folium.Map(location=map_center, zoom_start=12) # 稍微缩小初始视野

        print("正在向地图添加带有颜色和详细信息的小区标记...")
        for idx, row in df.iterrows():
            # 1. 根据群障次数决定标记颜色
            incidents = row[INCIDENT_COUNT_COL]
            if incidents >= 2:
                marker_color = 'red'
            elif incidents == 1:
                marker_color = 'orange'  # 橙色比黄色更显眼
            else:
                marker_color = 'blue'

            # 2. 创建包含详细信息的弹出窗口内容 (使用HTML)
            popup_html = f"""
            <b>小区名称:</b> {row[NAME_COLUMN]}<br>
            <b>用户数:</b> {int(row[USER_COUNT_COL])}<br>
            <b>群障次数:</b> {int(incidents)}
            """
            popup = folium.Popup(popup_html, max_width=300)

            # 3. 创建带有自定义颜色图标和弹出信息的标记
            folium.Marker(
                location=[row['center_lat'], row['center_lon']],
                popup=popup,
                icon=folium.Icon(color=marker_color, icon='info-sign')
            ).add_to(m)

        # 保存地图
        m.save(OUTPUT_MAP_FILE)
        print(f"地图创建成功！已保存为: {OUTPUT_MAP_FILE}")
        print("请用浏览器打开该文件查看。")

except FileNotFoundError:
    print(f"错误：找不到文件 '{EXCEL_FILE_PATH}'。")
except KeyError as e:
    print(f"错误：找不到列 {e}。请检查 Excel 中是否包含所有必需的列名。")
except Exception as e:
    print(f"发生未知错误: {e}")