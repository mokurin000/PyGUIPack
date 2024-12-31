import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog

import pandas as pd


###------排放因子------###
# 燃料碳排放因子字典
fuel_factors = {
    "电": {"factor": 6.546, "unit": "kwh"},
    "柴油": {"factor": {"kg": 3.094, "L": 2.63}, "units": ["kg", "L"]},
    "汽油": {"factor": {"kg": 3.172, "L": 2.3}, "units": ["kg", "L"]},
}

transport_factors = {
    "汽油货车": {
        "2t": 0.334,
        "8t": 0.115,
        "10t": 0.104,
        "18t": 0.104,
    },
    "柴油货车运输": {
        "2t": 0.286,
        "8t": 0.179,
        "10t": 0.162,
        "18t": 0.129,
        "30t": 0.078,
        "46t": 0.057,
    },
    "电力机车": {"na": 0.01},
    "内燃机车": {"na": 0.011},
    "铁路": {"na": 0.01},
    "液货船运输": {"2000t": 0.019},
    "干散货船": {"2500t": 0.015},
    "集装箱船": {"200TEU": 0.012},
}

# 材料排放因子字典
material_factors = {
    "普通硅酸盐水泥（市场平均）": {"factor": 735, "unit": "kg/t", "units": ["t"]},
    "混凝土C30": {"factor": 295, "unit": "kg/m3", "units": ["m3"]},
    "混凝土C50": {"factor": 385, "unit": "kg/m3", "units": ["m3"]},
    "石灰生产（市场平均）": {"factor": 1190, "unit": "kg/t", "units": ["t"]},
    "消石灰（熟石灰、氢氧化钙）": {"factor": 747, "unit": "kg/t", "units": ["t"]},
    "天然石膏": {"factor": 32.8, "unit": "kg/t", "units": ["t"]},
    "砂(f/=1.6-3.0)": {"factor": 2.51, "unit": "kg/t", "units": ["t"]},
    "碎石(d=10mm-30mm)": {"factor": 2.18, "unit": "kg/t", "units": ["t"]},
    "页岩石": {"factor": 5.08, "unit": "kg/t", "units": ["t"]},
    "黏土": {"factor": 2.69, "unit": "kg/t", "units": ["t"]},
    "混凝土砖(240mm*115mm*90mm)": {"factor": 336, "unit": "kg/m3", "units": ["m3"]},
    "蒸压粉煤灰砖(240mm*115mm*53mm)": {"factor": 341, "unit": "kg/m3", "units": ["m3"]},
    "烧结粉煤灰实心砖（240mmX115mmX53mm,掺入量为50%）": {
        "factor": 134,
        "unit": "kg/m3",
        "units": ["m3"],
    },
    "页岩实心砖(240mm*115mm*53mm)": {"factor": 292, "unit": "kg/m3", "units": ["m3"]},
    "页岩空心砖(240mm*115mm*53mm)": {"factor": 204, "unit": "kg/m3", "units": ["m3"]},
    "黏土空心砖(240mm*115mm*53mm)": {"factor": 250, "unit": "kg/m3", "units": ["m3"]},
    "煤肝石实心砖（240mmX115mmX53mm,90%掺入量）": {
        "factor": 22.8,
        "unit": "kg/m3",
        "units": ["m3"],
    },
    "煤肝石空心砖（240mm*115mm*53mm,90%掺入量）": {
        "factor": 16,
        "unit": "kg/m3",
        "units": ["m3"],
    },
    "炼钢生铁": {"factor": 1700, "unit": "kg/t", "units": ["t"]},
    "铸造生铁": {"factor": 2280, "unit": "kg/t", "units": ["t"]},
    "炼钢用铁合金": {"factor": 9530, "unit": "kg/t", "units": ["t"]},
    "转炉碳钢": {"factor": 1990, "unit": "kg/t", "units": ["t"]},
    "电炉碳钢": {"factor": 3030, "unit": "kg/t", "units": ["t"]},
    "普通碳钢": {"factor": 2050, "unit": "kg/t", "units": ["t"]},
    "热轧碳钢小型型钢": {"factor": 2310, "unit": "kg/t", "units": ["t"]},
    "热轧碳钢中型型钢": {"factor": 2365, "unit": "kg/t", "units": ["t"]},
    "热轧碳钢大型轨梁（方圆坯、管坯）": {
        "factor": 2340,
        "unit": "kg/t",
        "units": ["t"],
    },
    "热轧碳钢大型轨梁（重轨、普通型钢）": {
        "factor": 2380,
        "unit": "kg/t",
        "units": ["t"],
    },
    "热轧碳钢中厚板": {"factor": 2400, "unit": "kg/t", "units": ["t"]},
    "热轧碳钢H钢": {"factor": 2350, "unit": "kg/t", "units": ["t"]},
    "热轧碳钢宽带钢": {"factor": 2310, "unit": "kg/t", "units": ["t"]},
    "热轧碳钢钢筋": {"factor": 2340, "unit": "kg/t", "units": ["t"]},
    "热轧碳钢高线材": {"factor": 2375, "unit": "kg/t", "units": ["t"]},
    "热轧碳钢棒材": {"factor": 2340, "unit": "kg/t", "units": ["t"]},
    "螺旋埋弧焊管": {"factor": 2520, "unit": "kg/t", "units": ["t"]},
    "大口径埋弧焊直缝钢管": {"factor": 2430, "unit": "kg/t", "units": ["t"]},
    "焊接直缝钢管": {"factor": 2530, "unit": "kg/t", "units": ["t"]},
    "热轧碳钢无缝钢管": {"factor": 3150, "unit": "kg/t", "units": ["t"]},
    "冷轧冷拔碳钢无缝钢管": {"factor": 3680, "unit": "kg/t", "units": ["t"]},
    "碳钢热镀锌板卷": {"factor": 3110, "unit": "kg/t", "units": ["t"]},
    "碳钢电镀锌板卷": {"factor": 3020, "unit": "kg/t", "units": ["t"]},
    "碳钢电镀锡板卷": {"factor": 2870, "unit": "kg/t", "units": ["t"]},
    "酸洗板卷": {"factor": 1730, "unit": "kg/t", "units": ["t"]},
    "冷轧碳钢板卷": {"factor": 2530, "unit": "kg/t", "units": ["t"]},
    "冷硬碳钢板卷": {"factor": 2410, "unit": "kg/t", "units": ["t"]},
    "平板玻璃": {"factor": 1130, "unit": "kg/t", "units": ["t"]},
    "电解铝": {"factor": 20300, "unit": "kg/t", "units": ["t"]},
    "铝板带": {"factor": 28500, "unit": "kg/t", "units": ["t"]},
    "断桥铝合金窗（100%原生铝型材）": {"factor": 254, "unit": "kg/m2", "units": ["m2"]},
    "断桥铝合金窗（原生铝：再生铝=7：3）": {
        "factor": 194,
        "unit": "kg/m2",
        "units": ["m2"],
    },
    "铝木复合窗（100%原生铝型材）": {"factor": 147, "unit": "kg/m2", "units": ["m2"]},
    "铝木复合窗（原生铝：再生铝=7：3）": {
        "factor": 1225,
        "unit": "kg/m2",
        "units": ["m2"],
    },
    "铝塑共挤窗": {"factor": 1295, "unit": "kg/m2", "units": ["m2"]},
    "塑钢窗": {"factor": 121, "unit": "kg/m2", "units": ["m2"]},
    "无规共聚聚丙烯管": {"factor": 3.72, "unit": "kg/kg", "units": ["kg"]},
    "聚乙烯管": {"factor": 3.6, "unit": "kg/kg", "units": ["kg"]},
    "硬聚氯乙烯管": {"factor": 7.93, "unit": "kg/kg", "units": ["kg"]},
    "聚苯乙烯泡沫板": {"factor": 5020, "unit": "kg/t", "units": ["t"]},
    "岩棉板": {"factor": 1980, "unit": "kg/t", "units": ["t"]},
    "硬泡聚氨酯板": {"factor": 5220, "unit": "kg/t", "units": ["t"]},
    "铝塑复合板": {"factor": 8.06, "unit": "kg/m2", "units": ["m2"]},
    "铜塑复合板": {"factor": 37.1, "unit": "kg/m2", "units": ["m2"]},
    "铜单板": {"factor": 218, "unit": "kg/m2", "units": ["m2"]},
    "普通聚苯乙烯": {"factor": 4620, "unit": "kg/t", "units": ["t"]},
    "线性低密度聚乙烯": {"factor": 1990, "unit": "kg/t", "units": ["t"]},
    "高密度聚乙烯": {"factor": 2620, "unit": "kg/t", "units": ["t"]},
    "低密度聚乙烯": {"factor": 2810, "unit": "kg/t", "units": ["t"]},
    "聚氯乙烯": {"factor": 7300, "unit": "kg/t", "units": ["t"]},
    "水泥砂浆（水泥：砂=1：2）": {
        "factor": None,
        "unit": None,
        "units": [],
    },  # 未提供数据
    "水": {"factor": 0.168, "unit": "kg/t", "units": ["t"]},
}


# VOCs排放因子字典
vocs_factors = {
    "内墙涂料（水性）": {"factors": {"L": 0.03423, "kg": 0.02463}},
    "外墙涂料（水性）": {"factors": {"L": 0.02363, "kg": 0.0175}},
    "墙体涂料（溶剂型）": {"factors": {"L": 0.3735, "kg": 0.2988}},
    "防水涂料（水性）": {"factors": {"L": 0.00431, "kg": 0.00275}},
    "防水涂料（反应固化型）": {"factors": {"L": 0.10983, "kg": 0.08786}},
    "防水涂料（溶剂型）": {"factors": {"L": 0.5, "kg": 0.4}},
    "地坪涂料（水性）": {"factors": {"L": 0.10775, "kg": 0.0862}},
    "地坪涂料（无溶剂型）": {"factors": {"L": 0.03155, "kg": 0.02524}},
    "地坪涂料（溶剂型）": {"factors": {"L": 0.39625, "kg": 0.317}},
    "防腐涂料（水性）": {"factors": {"L": 0.03994, "kg": 0.03195}},
    "防腐涂料（溶剂型）": {"factors": {"L": 0.52036, "kg": 0.46461}},
    "防火涂料（水性）": {"factors": {"L": 0.08, "kg": 0.0597}},
    "防火涂料（溶剂型）": {"factors": {"L": 0.46525, "kg": 0.3472}},
    "氯丁橡胶胶黏剂（溶剂型）": {"factors": {"L": 0.621}},
    "SBS胶黏剂（溶剂型）": {"factors": {"L": 0.593}},
    "聚氨酯类胶黏剂（溶剂型）": {"factors": {"L": 0.55}},
    "丙烯酸酯类胶黏剂（溶剂型）": {"factors": {"L": 0.539}},
    "一般溶剂型胶黏剂": {"factors": {"L": 0.597}},
    "聚乙酸乙烯酯类胶黏剂（水基型）": {"factors": {"L": 0.043}},
    "缩甲醛类（水基型）": {"factors": {"L": 0.107}},
    "聚氨酯类胶黏剂（水基型）": {"factors": {"L": 0.037}},
    "丙烯酸酯类胶黏剂（水基型）": {"factors": {"L": 0.225}},
    "橡胶类胶黏剂（水基型）": {"factors": {"L": 0.2}},
    "VAE乳液类胶黏剂（水基型）": {"factors": {"L": 0.225}},
    "一般水基型胶黏剂": {"factors": {"L": 0.068}},
    "有机硅类(含MS)胶黏剂（本体型）": {"factors": {"L": 0.074}},
    "聚氨酯类胶黏剂（本体型）": {"factors": {"L": 0.047}},
    "聚硫类胶黏剂（本体型）": {"factors": {"L": 0.075}},
    "环氧类胶黏剂（本体型）": {"factors": {"L": 0.075}},
    "一般本体型胶黏剂": {"factors": {"L": 0.063}},
}

# 人员排放因子字典
BREATHING_CO2_COEFFICIENT = 0.000308

# 施工设备排放因子字典
devices_fuel_factors = {
    "电": 6.546,
    "柴油": 3.094,
    "汽油": 3.172,
}

equipment_data = {
    "履带式推土机": {
        "specs": ["功率75kW", "功率105kW", "功率135kW"],
        "fuels": ["柴油"],
    },
    "履带式单斗液压挖掘机": {"specs": ["斗容量0.6m3", "斗容量1m3"], "fuels": ["柴油"]},
    "轮胎式装载机": {"specs": ["斗容量1m3", "斗容量1.5m3"], "fuels": ["柴油"]},
    "钢轮内燃压路机": {"specs": ["工作质量8t", "工作质量15t"], "fuels": ["柴油"]},
    "电动夯实机": {
        "specs": [
            "夯击能量250N•m",
            "夯击能量1200kN•m",
            "夯击能量2000kN•m",
            "夯击能量3000kN•m",
            "夯击能量4000kN•m",
            "夯击能量5000kN•m",
        ],
        "fuels": ["电"],
    },
    "锚杆钻孔机": {"specs": ["锚杆直径32mm"], "fuels": ["柴油"]},
    "履带式柴油打桩机": {
        "specs": [
            "冲击质量2.5t",
            "冲击质量3.5t",
            "冲击质量5t",
            "冲击质量7t",
            "冲击质量8t",
        ],
        "fuels": ["柴油"],
    },
    "轨道式柴油打桩机": {"specs": ["冲击质量3.5t", "冲击质量4t"], "fuels": ["柴油"]},
    "步履式柴油打桩机": {"specs": ["功率60kW"], "fuels": ["电"]},
    "振动沉拔桩机": {"specs": ["激振力300kN", "激振力400kN"], "fuels": ["柴油"]},
    "静力压桩机": {
        "specs": ["压力900kN", "压力2000kN", "压力3000kN", "压力4000kN"],
        "fuels": ["电"],
    },
    "汽车式钻机": {"specs": ["孔径1000mm"], "fuels": ["柴油"]},
    "回旋钻机": {"specs": ["孔径800mm", "孔径1000mm", "孔径1500mm"], "fuels": ["电"]},
    "螺旋钻机": {"specs": ["孔径600mm"], "fuels": ["电"]},
    "冲孔钻机": {"specs": ["孔径1000mm"], "fuels": ["电"]},
    "履带式旋挖钻机": {
        "specs": ["孔径1000mm", "孔径1500mm", "孔径2000mm"],
        "fuels": ["柴油"],
    },
    "三轴搅拌桩基": {"specs": ["轴径650mm", "轴径850mm"], "fuels": ["电"]},
    "电动灌浆机": {"specs": [None], "fuels": ["电"]},
    "履带式起重机": {
        "specs": [
            "提升质量5t",
            "提升质量10t",
            "提升质量15t",
            "提升质量20t",
            "提升质量25t",
            "提升质量30t",
            "提升质量40t",
            "提升质量50t",
            "提升质量60t",
        ],
        "fuels": ["柴油"],
    },
    "轮胎式起重机": {
        "specs": ["提升质量25t", "提升质量53", "提升质量54"],
        "fuels": ["柴油"],
    },
    "汽车式起重机": {
        "specs": [
            "提升质量8t",
            "提升质量12t",
            "提升质量16t",
            "提升质量20t",
            "提升质量30t",
            "提升质量40t",
        ],
        "fuels": ["柴油"],
    },
    "叉式起重机": {"specs": ["提升质量3t"], "fuels": ["汽油"]},
    "自升式塔式起重机": {
        "specs": [
            "提升质量400t",
            "提升质量600t",
            "提升质量800t",
            "提升质量l000t",
            "提升质量2500t",
            "提升质量3000t",
        ],
        "fuels": ["电"],
    },
    "门式起重机": {"specs": ["提升质量10t"], "fuels": ["电"]},
    "载重汽车": {
        "specs": [
            "装载质量4t",
            "装载质量6t",
            "装载质量8t",
            "装载质量12t",
            "装载质量15t",
            "装载质量20t",
        ],
        "fuels": ["柴油", "汽油"],
    },
    "自卸汽车": {"specs": ["装载质量5t", "装载质量15t"], "fuels": ["汽油"]},
    "平板拖车组": {"specs": ["装载质量20t"], "fuels": ["柴油"]},
    "机动翻斗车": {"specs": ["装载质量1t"], "fuels": ["柴油"]},
    "洒水车": {"specs": ["灌容量4000L"], "fuels": ["汽油"]},
    "泥浆罐车": {"specs": ["灌容量5000L"], "fuels": ["汽油"]},
    "电动单筒快速卷扬机": {"specs": ["牵引力10kN"], "fuels": ["电"]},
    "电动单筒慢速卷扬机": {"specs": ["牵引力10kN", "牵引力30kN"], "fuels": ["电"]},
    "单笼施工电梯": {
        "specs": ["提升质量1t提升高度75m", "提升质量1t提升高度100m"],
        "fuels": ["电"],
    },
    "双笼施工电梯": {
        "specs": ["提升质量2t提升高度100m", "提升质量2t提升高度200m"],
        "fuels": ["电"],
    },
    "平台作业升降车": {"specs": ["提升高度20m"], "fuels": ["汽油"]},
    "涡桨式混凝土搅拌机": {"specs": ["出料容量250L", "出料容量500L"], "fuels": ["电"]},
    "双锥反转出料混凝土搅拌机": {"specs": ["出料容量500L"], "fuels": ["电"]},
    "混凝土输送泵": {"specs": ["输送量45m3/h", "输送量75m3/h"], "fuels": ["电"]},
    "混凝土湿喷机": {"specs": ["生产率5m3/h"], "fuels": ["电"]},
    "灰浆搅拌机": {"specs": ["拌筒容量200L"], "fuels": ["电"]},
    "干混砂浆罐式搅拌机": {"specs": ["公称储量20000L"], "fuels": ["电"]},
    "挤压式灰浆输送泵": {"specs": ["输送量3m3/h"], "fuels": ["电"]},
    "偏心振动筛": {"specs": ["生产率16m3/h"], "fuels": ["电"]},
    "混凝土抹平机": {"specs": ["功率5.5kW"], "fuels": ["电"]},
    "钢筋切断机": {"specs": ["直径40mm"], "fuels": ["电"]},
    "钢筋弯曲机": {"specs": ["直径40mm"], "fuels": ["电"]},
    "预应力钢筋拉伸机": {"specs": ["拉伸力650kN", "拉伸力900kN"], "fuels": ["电"]},
    "木工圆锯机": {"specs": ["直径500mm"], "fuels": ["电"]},
    "木工平刨床": {"specs": ["刨削宽度500mm"], "fuels": ["电"]},
    "木工三面压刨床": {"specs": ["刨削宽度400mm"], "fuels": ["电"]},
    "木工梯机": {"specs": ["样头长度160mm"], "fuels": ["电"]},
    "木工打眼机": {"specs": [None], "fuels": ["电"]},
    "普通车床": {"specs": ["工件直径400mm工件长度2000mm"], "fuels": ["电"]},
    "摇臂钻床": {"specs": ["钻孔直径50mm", "钻孔直径63mm"], "fuels": ["电"]},
    "锥形螺纹车丝机": {"specs": ["直径45mm"], "fuels": ["电"]},
    "螺栓套丝机": {"specs": [None], "fuels": ["电"]},
    "板料校平机": {"specs": ["厚度16mm宽度2000mm"], "fuels": ["电"]},
    "刨边机": {"specs": ["加工长度12000mm"], "fuels": ["电"]},
    "半自动切割机": {"specs": ["厚度100mm"], "fuels": ["电"]},
    "自动仿形切割机": {"specs": ["厚度60mm"], "fuels": ["电"]},
    "管子切断机": {"specs": ["管径150mm", "管径250mm"], "fuels": ["电"]},
    "型钢剪断机": {"specs": ["剪断宽度500mm"], "fuels": ["电"]},
    "型钢矫正机": {"specs": ["厚度60mm宽度800mm"], "fuels": ["电"]},
    "电动弯管机": {"specs": ["管径108mm"], "fuels": ["电"]},
    "液压弯管机": {"specs": ["管径60mm"], "fuels": ["电"]},
    "空气锤": {"specs": ["锤体质量75kg"], "fuels": ["电"]},
    "摩擦压力机": {"specs": ["压力3000kN"], "fuels": ["电"]},
    "开式可倾压力机": {"specs": ["压力1250kN"], "fuels": ["电"]},
    "钢筋挤压连接机": {"specs": [None], "fuels": ["电"]},
    "电动修钎机": {"specs": [None], "fuels": ["电"]},
    "岩石切割机": {"specs": ["功率3kW"], "fuels": ["电"]},
    "平面水磨机": {"specs": ["功率3kW"], "fuels": ["电"]},
    "喷砂除锈机": {"specs": ["能力3m3/min"], "fuels": ["电"]},
    "抛丸除锈机": {"specs": ["直径219mm"], "fuels": ["电"]},
    "内燃单级离心清水泵": {"specs": ["出口直径50mm"], "fuels": ["柴油"]},
    "电动多级离心清水泵": {
        "specs": [
            "出口直径100mm扬程120m以下",
            "出口直径150mm扬程180m以下",
            "出口直径200mm扬程280m以下",
        ],
        "fuels": ["电"],
    },
    "泥浆泵": {"specs": ["出口直径50mm", "出口直径100mm"], "fuels": ["电"]},
    "潜水泵": {"specs": ["出口直径50mm", "出口直径100mm"], "fuels": ["电"]},
    "高压油泵": {"specs": ["压力80Mpa"], "fuels": ["电"]},
    "交流弧焊机": {
        "specs": ["容量21kV•A", "容量32kV•A", "容量40kV•A"],
        "fuels": ["电"],
    },
    "点焊机": {"specs": ["容量75kV•A"], "fuels": ["电"]},
    "对焊机": {"specs": ["容量75kV•A"], "fuels": ["电"]},
    "量弧焊机": {"specs": ["电流500A"], "fuels": ["电"]},
    "二氧化碳气体保护焊机": {"specs": ["电流250A"], "fuels": ["电"]},
    "电渣焊机": {"specs": ["电流1000A"], "fuels": ["电"]},
    "电焊条烘干箱": {"specs": ["容量45*35*45cm3"], "fuels": ["电"]},
    "电动空气压缩机": {
        "specs": [
            "排气量0.3m3/min",
            "排气量0.6m3/min",
            "排气量1m3/min",
            "排气量3m3/min",
            "排气量6m3/min",
            "排气量9m3/min",
            "排气量10m3/min",
        ],
        "fuels": ["电"],
    },
    "导杆式液压抓斗成槽机": {"specs": [None], "fuels": ["电"]},
    "超声波侧壁机": {"specs": [None], "fuels": ["电"]},
    "泥浆制作循环设备": {"specs": [None], "fuels": ["电"]},
    "锁扣管顶升机": {"specs": [None], "fuels": ["电"]},
    "工程地质液压钻机": {"specs": [None], "fuels": ["电"]},
    "轴流通风机": {"specs": ["功率7.5kW"], "fuels": ["电"]},
    "吹风机": {"specs": ["能力4m3/min"], "fuels": ["电"]},
    "井点降水钻机": {"specs": [None], "fuels": ["电"]},
}


###------处理------###
######workers######
def on_calculate_workers():
    try:
        h = float(hours_entry.get()) if hours_entry.get() else 0.0
        worker_CO2 = h * BREATHING_CO2_COEFFICIENT
        result_text.insert(tk.END, f"{worker_CO2}\n")
    except ValueError:
        # messagebox.showerror("错误", "请输入有效的数值。")
        result_text.insert(tk.END, "请输入有效的数值。\n")


######fuel######
def calculate_fuel_emission(fuel_type, usage, unit):
    if fuel_type in fuel_factors:
        factor_data = fuel_factors[fuel_type]
        if fuel_type == "电":
            factor = factor_data["factor"]
            emission = usage * factor
            return f"{fuel_type}排放量：{emission:.2f} kg"
        else:  # 柴油和汽油
            factor = factor_data["factor"][unit]
            emission = usage * factor
            return f"{fuel_type}排放量：{emission:.2f} kg"
    else:
        return "不支持的燃料类型。"


def on_calculate_fuel():
    fuel_type = fuel_type_var.get()
    try:
        usage = float(usage_entry.get())
        if fuel_type in fuel_factors:
            available_units = fuel_factors[fuel_type].get("units", ["kwh"])
            unit = unit_var.get()
            if fuel_type != "电" and unit not in available_units:
                messagebox.showerror(
                    "错误",
                    f"{fuel_type}不支持单位 {unit}，请选择 {', '.join(available_units)}",
                )
                return
            result = calculate_fuel_emission(fuel_type, usage, unit)
            # messagebox.showinfo("结果", result)
            result_text.insert(tk.END, f"{result}\n")
        else:
            # messagebox.showerror("错误", "不支持的燃料类型。")
            result_text.insert(tk.END, "不支持的燃料类型。\n")
    except ValueError:
        # messagebox.showerror("错误", "请输入有效的使用量。")
        result_text.insert(tk.END, "输入无效\n")


def update_unit_options(event):
    fuel_type = fuel_type_var.get()
    if fuel_type in fuel_factors:
        available_units = fuel_factors[fuel_type].get("units", ["kwh"])
        unit_var["values"] = available_units
        if available_units:
            unit_var.set(available_units[0])  # 默认值设置为第一个单位


def upload_csv_and_calculate_fuel():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if not file_path:
        return

    try:
        df = pd.read_csv(file_path)

        if (
            "燃料类型" not in df.columns
            or "使用量" not in df.columns
            or "单位" not in df.columns
        ):
            messagebox.showerror(
                "错误", "CSV 文件必须包含 '燃料类型'、'使用量' 和 '单位' 列。"
            )
            return

        results = []
        total_emission = 0.0

        # 为 DataFrame 添加新的列
        df["排放量"] = 0.0

        for index, row in df.iterrows():
            fuel_type = row["燃料类型"]
            usage = row["使用量"]
            unit = row["单位"]
            if fuel_type in fuel_factors and unit in fuel_factors[fuel_type].get(
                "units", ["kwh"]
            ):
                emission_result = calculate_fuel_emission(fuel_type, usage, unit)
                emission_value = float(emission_result.split("：")[1].split(" ")[0])
                df.at[index, "排放量"] = emission_value
                results.append(emission_result)
                total_emission += emission_value
            else:
                results.append(f"不支持的燃料类型或单位：{fuel_type} - {unit}")

        result_summary = "\n".join(results)
        result_summary += f"\n总排放量：{total_emission:.2f} kg"

        # 添加总排放量为新行
        total_row = pd.DataFrame([["总计", "", "", total_emission]], columns=df.columns)
        df = pd.concat([df, total_row], ignore_index=True)

        # 生成新的 CSV 文件
        new_file_path = file_path.replace(".csv", "_计算结果.csv")
        df.to_csv(new_file_path, index=False, encoding="utf-8-sig")

        # messagebox.showinfo("计算结果", result_summary)
        result_text.insert(tk.END, f"{result_summary}\n")
        messagebox.showinfo("CSV 生成", f"已生成新的 CSV 文件：{new_file_path}")

    except Exception as e:
        messagebox.showerror("错误", f"处理文件时发生错误：{str(e)}")


######transport######
def calculate_transport_emission(vehicle_type, spec, distance):
    if vehicle_type in transport_factors:
        if (
            spec in transport_factors[vehicle_type] or spec == "自定义"
        ):  # 允许自定义规格
            factor = transport_factors[vehicle_type].get(
                spec, 0.01
            )  # 如果自定义规格不存在，默认值设置为0.01
            emission = distance * factor
            return f"{vehicle_type}（{spec}）排放量：{emission:.2f} kg"
        else:
            return "不支持的规格。"
    else:
        return "不支持的运输设备类型。"


def on_calculate_transport():
    vehicle_type = vehicle_type_var.get()
    spec = spec_entry.get()
    try:
        distance = float(distance_entry.get())
        result = calculate_transport_emission(vehicle_type, spec, distance)
        # messagebox.showinfo("结果", result)
        result_text.insert(tk.END, f"{result}\n")
    except ValueError:
        # messagebox.showerror("错误", "请输入有效的运输量（t*km）。")
        result_text.insert(tk.END, "请输入有效的运输量（t*km）。\n")


def update_spec_options(event):
    vehicle_type = vehicle_type_var.get()
    if vehicle_type in transport_factors:
        available_specs = sorted(list(transport_factors[vehicle_type].keys()))
        spec_entry["values"] = available_specs
        if available_specs:
            spec_entry.set(available_specs[0])  # 设置默认值为第一个规格


def upload_csv_and_calculate_transport():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if not file_path:
        return

    try:
        df = pd.read_csv(file_path)

        if (
            "运输方式" not in df.columns
            or "距离" not in df.columns
            or "单位" not in df.columns
        ):
            messagebox.showerror(
                "错误", "CSV 文件必须包含 '运输方式'、'距离' 和 '单位' 列。"
            )
            return

        results = []
        total_emission = 0.0

        # 为 DataFrame 添加新的列
        df["排放量"] = 0.0

        for index, row in df.iterrows():
            transport_mode = row["运输方式"]
            distance = row["距离"]
            unit = row["单位"]
            if transport_mode in transport_factors and unit in transport_factors[
                transport_mode
            ].get("units", ["km"]):
                emission_result = calculate_transport_emission(
                    transport_mode, distance, unit
                )
                emission_value = float(emission_result.split("：")[1].split(" ")[0])
                df.at[index, "排放量"] = emission_value
                results.append(emission_result)
                total_emission += emission_value
            else:
                results.append(f"不支持的运输方式或单位：{transport_mode} - {unit}")

        result_summary = "\n".join(results)
        result_summary += f"\n总排放量：{total_emission:.2f} kg"

        # 添加总排放量为新行
        total_row = pd.DataFrame([["总计", "", "", total_emission]], columns=df.columns)
        df = pd.concat([df, total_row], ignore_index=True)

        # 生成新的 CSV 文件
        new_file_path = file_path.replace(".csv", "_计算结果.csv")
        df.to_csv(new_file_path, index=False, encoding="utf-8-sig")

        # messagebox.showinfo("计算结果", result_summary)
        result_text.insert(tk.END, f"{result_summary}\n")
        messagebox.showinfo("CSV 生成", f"已生成新的 CSV 文件：{new_file_path}")

    except Exception as e:
        messagebox.showerror("错误", f"处理文件时发生错误：{str(e)}")


######material######
def calculate_material_emission(material_name, quantity, unit):
    if material_name in material_factors:
        factor = material_factors[material_name]["factor"]
        emission = quantity * factor
        return f"{material_name} 排放量：{emission:.2f} kg"
    else:
        return "不支持的材料类型。"


def on_calculate_material():
    material_name = material_name_var.get()
    try:
        quantity_direct = quantity_entry.get()
        input_amount = input_amount_entry.get()
        loss_rate = loss_rate_entry.get()

        quantity = 0.0
        if quantity_direct:
            quantity = float(quantity_direct)
        elif input_amount and loss_rate:
            quantity = float(input_amount) * float(loss_rate)
        else:
            messagebox.showerror("错误", "请输入损耗量 或 投入量和损耗率。")
            return

        unit = unit_var_material.get()
        if material_name in material_factors:
            if unit not in material_factors[material_name]["units"]:
                messagebox.showerror(
                    "错误",
                    f"{material_name}不支持单位 {unit}，请选择 {', '.join(material_factors[material_name]['units'])}",
                )
                return
            result = calculate_material_emission(material_name, quantity, unit)
            # messagebox.showinfo("结果", result)
            result_text.insert(tk.END, f"{result}\n")
        else:
            # messagebox.showerror("错误", "不支持的材料类型。")
            result_text.insert(tk.END, "不支持的材料类型。\n")
    except ValueError:
        # messagebox.showerror("错误", "请输入有效的数值。")
        result_text.insert(tk.END, "请输入有效的数值。\n")


def update_material_unit_options(event):
    material_name = material_name_var.get()
    if material_name in material_factors:
        available_units = material_factors[material_name]["units"]
        unit_var_material["values"] = available_units
        if available_units:
            unit_var_material.set(available_units[0])  # 设置默认值为第一个单位


def upload_csv_and_calculate_material():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if not file_path:
        return

    try:
        df = pd.read_csv(file_path)

        if (
            "材料类型" not in df.columns
            or "用量" not in df.columns
            or "单位" not in df.columns
        ):
            messagebox.showerror(
                "错误", "CSV 文件必须包含 '依据材料'、'用量' 和 '单位' 列。"
            )
            return

        results = []
        total_emission = 0.0

        # 为 DataFrame 添加新的'排放量'列，初始化为 0
        df["排放量"] = 0.0

        for index, row in df.iterrows():
            material_type = row["依据材料"]
            amount = row["用量"]
            unit = row["单位"]
            if material_type in material_factors and unit in material_factors[
                material_type
            ].get("units", ["kg"]):
                emission_result = calculate_material_emission(
                    material_type, amount, unit
                )
                emission_value = float(emission_result.split("：")[1].split(" ")[0])
                df.at[index, "排放量"] = emission_value
                results.append(emission_result)
                total_emission += emission_value
            else:
                results.append(f"不支持的材料类型或单位：{material_type} - {unit}")

        result_summary = "\n".join(results)
        result_summary += f"\n总排放量：{total_emission:.2f} kg"

        # 创建总计行的数据
        total_row = pd.DataFrame([["总计", "", "", total_emission]], columns=df.columns)

        # 将总计行添加到 DataFrame
        df = pd.concat([df, total_row], ignore_index=True)

        # 生成新的 CSV 文件
        new_file_path = file_path.replace(".csv", "_计算结果.csv")
        df.to_csv(new_file_path, index=False, encoding="utf-8-sig")

        # messagebox.showinfo("计算结果", result_summary)
        result_text.insert(tk.END, f"{result_summary}\n")
        messagebox.showinfo("CSV 生成", f"已生成新的 CSV 文件：{new_file_path}")

    except Exception as e:
        messagebox.showerror("错误", f"处理文件时发生错误：{str(e)}")


######vocs######
def calculate_vocs_emission(material_name, usage, unit):
    if material_name in vocs_factors and unit in vocs_factors[material_name]["factors"]:
        factor = vocs_factors[material_name]["factors"][unit]
        emission = usage * factor
        return f"{material_name} VOCs 排放量：{emission:.2f} kg"
    else:
        return "不支持的VOCs材料类型或单位。"


def on_calculate_vocs():
    material_name = vocs_material_name_var.get()
    try:
        usage = float(vocs_usage_entry.get())
        unit = vocs_unit_var.get()
        result = calculate_vocs_emission(material_name, usage, unit)
        # messagebox.showinfo("结果", result)
        result_text.insert(tk.END, f"{result}\n")
    except ValueError:
        # messagebox.showerror("错误", "请输入有效的使用量。")
        result_text.insert(tk.END, "请输入有效的使用量。\n")


def update_vocs_unit_options(event):
    material_name = vocs_material_name_var.get()
    if material_name in vocs_factors:
        available_units = list(vocs_factors[material_name]["factors"].keys())
        vocs_unit_var["values"] = available_units
        if available_units:
            vocs_unit_var.set(available_units[0])  # 设置默认值为第一个单位


def upload_csv_and_calculate_vocs():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if not file_path:
        return

    try:
        df = pd.read_csv(file_path)

        if (
            "VOCs 类型" not in df.columns
            or "用量" not in df.columns
            or "单位" not in df.columns
        ):
            messagebox.showerror(
                "错误", "CSV 文件必须包含 'VOCs 类型'、'用量' 和 '单位' 列。"
            )
            return

        results = []
        total_emission = 0.0

        # 为 DataFrame 添加新的'排放量'列，初始化为 0
        df["排放量"] = 0.0

        for index, row in df.iterrows():
            vocs_type = row["VOCs 类型"]
            amount = row["用量"]
            unit = row["单位"]
            if vocs_type in vocs_factors and unit in vocs_factors[vocs_type].get(
                "units", ["kg"]
            ):
                emission_result = calculate_vocs_emission(vocs_type, amount, unit)
                emission_value = float(emission_result.split("：")[1].split(" ")[0])
                df.at[index, "排放量"] = emission_value
                results.append(emission_result)
                total_emission += emission_value
            else:
                results.append(f"不支持的 VOCs 类型或单位：{vocs_type} - {unit}")

        # 结果摘要
        result_summary = "\n".join(results)  # 使用换行符
        result_summary += f"\n总排放量：{total_emission:.2f} kg"

        # 重新初始化总排放量
        total_row = pd.DataFrame([["总计", "", "", total_emission]], columns=df.columns)

        # 将总计行添加到 DataFrame
        df = pd.concat([df, total_row], ignore_index=True)

        # 生成新的 CSV 文件
        new_file_path = file_path.replace(".csv", "_计算结果.csv")
        df.to_csv(new_file_path, index=False, encoding="utf-8-sig")

        # messagebox.showinfo("计算结果", result_summary)
        result_text.insert(tk.END, f"{result_summary}\n")
        messagebox.showinfo("CSV 生成", f"已生成新的 CSV 文件：{new_file_path}")

    except Exception as e:
        messagebox.showerror("错误", f"处理文件时发生错误：{str(e)}")


######devices######
def calculate_emission(fuel_type, shifts):
    if fuel_type in devices_fuel_factors:
        return devices_fuel_factors[fuel_type] * shifts
    else:
        return None


def on_calculate_emission():
    equipment_name = equipment_name_var.get()
    spec = spec_var.get()
    fuel_type = fuel_var.get()

    try:
        shifts = int(shifts_entry.get())

        if shifts < 0:
            messagebox.showerror("错误", "台班数不能为负数。")
            return

        if (
            equipment_name in equipment_data
            and spec in equipment_data[equipment_name]["specs"]
            and fuel_type in equipment_data[equipment_name]["fuels"]
        ):
            result = calculate_emission(fuel_type, shifts)
            # messagebox.showinfo("结果", f"{equipment_name}({spec}) 使用 {fuel_type} 的碳排放量：{emission:.2f} kg")
            result_text.insert(
                tk.END,
                f"{equipment_name}({spec}) 使用 {fuel_type} 的碳排放量：{result} kg\n",
            )
        else:
            # messagebox.showerror("错误", "不支持的施工设备、规格或燃料类型。")
            result_text.insert(tk.END, "不支持的施工设备、规格或燃料类型。\n")
    except ValueError:
        # messagebox.showerror("错误", "请输入有效的台班数。")
        result_text.insert(tk.END, "请输入有效的台班数。\n")


def update_spec_and_fuel_options(event):
    equipment_name = equipment_name_var.get()
    if equipment_name in equipment_data:
        available_specs = equipment_data[equipment_name]["specs"]
        available_fuels = equipment_data[equipment_name]["fuels"]
        spec_var["values"] = available_specs
        fuel_var["values"] = available_fuels

        if available_specs:
            spec_var.set(available_specs[0])  # Set default to the first option
        if available_fuels:
            fuel_var.set(available_fuels[0])  # Set default to the first option


def upload_csv_and_calculate_devices():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if not file_path:
        return

    try:
        df = pd.read_csv(file_path)

        # 检查必需的列是否存在
        if (
            "设备类型" not in df.columns
            or "规格" not in df.columns
            or "燃料类型" not in df.columns
            or "台班数" not in df.columns
        ):
            messagebox.showerror(
                "错误",
                "CSV 文件必须包含 '设备类型'、'规格'、'燃料类型' 和 '台班数' 列。",
            )
            return

        results = []
        total_emission = 0.0

        # 为 DataFrame 添加新的'排放量'列，初始化为 0
        df["排放量"] = 0.0

        for index, row in df.iterrows():
            equipment_name = row["设备类型"]
            spec = row["规格"]
            fuel_type = row["燃料类型"]
            shifts = row["台班数"]

            if (
                equipment_name in equipment_data
                and spec in equipment_data[equipment_name]["specs"]
                and fuel_type in equipment_data[equipment_name]["fuels"]
            ):
                emission_result = calculate_emission(fuel_type, shifts)
                df.at[index, "排放量"] = emission_result
                results.append(
                    f"{equipment_name}({spec}) 使用 {fuel_type} 的碳排放量：{emission_result:.2f} kg"
                )
                total_emission += emission_result
            else:
                results.append(
                    f"不支持的设备类型、规格或燃料类型：{equipment_name} - {spec} - {fuel_type}"
                )

        # 结果摘要
        result_summary = "\n".join(results)  # 使用换行符
        result_summary += f"\n总排放量：{total_emission:.2f} kg"

        # 重新初始化总排放量
        total_row = pd.DataFrame(
            [["总计", "", "", "", total_emission]], columns=df.columns
        )

        # 将总计行添加到 DataFrame
        df = pd.concat([df, total_row], ignore_index=True)

        # 生成新的 CSV 文件
        new_file_path = file_path.replace(".csv", "_计算结果.csv")
        df.to_csv(new_file_path, index=False, encoding="utf-8-sig")

        # 将结果显示在文本框中
        result_text.insert(tk.END, f"{result_summary}\\n")
        messagebox.showinfo("CSV 生成", f"已生成新的 CSV 文件：{new_file_path}")

    except Exception as e:
        messagebox.showerror("错误", f"处理文件时发生错误：{str(e)}")


###------GUI-----###
######root######
# 创建主窗口
root = tk.Tk()
root.title("碳排放计算器")

# 水平分隔线
separator1 = ttk.Separator(root, orient="horizontal")
separator1.grid(row=1, column=0, columnspan=4, sticky="ew")

separator4 = ttk.Separator(root, orient="horizontal")
separator4.grid(row=3, column=0, columnspan=4, sticky="ew")

# 垂直分隔线
separator2 = ttk.Separator(root, orient="vertical")
separator2.grid(row=0, column=1, rowspan=5, sticky="ns")

# separator3 = ttk.Separator(root, orient='vertical')
# separator3.grid(row=0, column=3, rowspan=3, sticky='ns')

# 设置 3x2 布局
frame_fuel = ttk.Frame(root)
frame_fuel.grid(row=0, column=0, padx=10, pady=10, sticky="n")

frame_material = ttk.Frame(root)
frame_material.grid(row=4, column=0, padx=10, pady=10, sticky="n")

frame_transport = ttk.Frame(root)
frame_transport.grid(row=2, column=0, padx=10, pady=10, sticky="n")

frame_vocs = ttk.Frame(root)
frame_vocs.grid(row=4, column=2, padx=10, pady=10, sticky="n")

frame_workers = ttk.Frame(root)
frame_workers.grid(row=0, column=2, padx=10, pady=10, sticky="n")

frame_devices = ttk.Frame(root)
frame_devices.grid(row=2, column=2, padx=10, pady=10, sticky="n")

frame_result = ttk.Frame(root)
frame_result.grid(row=0, column=5, rowspan=5, padx=10, pady=10, sticky="n")

# 结果显示区域 (单一区域)
result_text = tk.Text(frame_result, width=30, height=36, wrap=tk.WORD)
result_text.grid(row=0, column=5, padx=10, sticky="nsew")


def clear_results():
    result_text.delete("1.0", tk.END)


clear_button = tk.Button(frame_result, text="清空结果", command=clear_results)
clear_button.grid(row=13, column=5, pady=10)


######devices######
tk.Label(frame_devices, text="请选择施工设备：").grid(row=0, column=0)
equipment_name_var = tk.StringVar()
equipment_combobox = ttk.Combobox(frame_devices, textvariable=equipment_name_var)
equipment_combobox["values"] = sorted(equipment_data.keys())
equipment_combobox.bind("<<ComboboxSelected>>", update_spec_and_fuel_options)
equipment_combobox.grid(row=0, column=1)

tk.Label(frame_devices, text="请选择规格：").grid(row=1, column=0)
spec_var = ttk.Combobox(frame_devices)
spec_var.grid(row=1, column=1)

tk.Label(frame_devices, text="请选择燃料：").grid(row=2, column=0)
fuel_var = ttk.Combobox(frame_devices)
fuel_var.grid(row=2, column=1)


tk.Label(frame_devices, text="请输入台班数：").grid(row=3, column=0)
shifts_entry = tk.Entry(frame_devices)
shifts_entry.grid(row=3, column=1)

calculate_emission_button = tk.Button(
    frame_devices, text="计算碳排放量", command=on_calculate_emission
)
calculate_emission_button.grid(row=4, column=1)

upload_csv_button_fuel = tk.Button(
    frame_devices, text="上传CSV计算排放", command=upload_csv_and_calculate_devices
)
upload_csv_button_fuel.grid(row=4, column=0, pady=5)


######workers######
tk.Label(frame_workers, text="总投入人员工时（人·小时）:").grid(row=0, column=0, pady=5)
hours_entry = tk.Entry(frame_workers)
hours_entry.grid(row=0, column=1, pady=5)

calculate_button = tk.Button(
    frame_workers, text="计算排放量", command=on_calculate_fuel
)
calculate_button.grid(row=3, columnspan=2, pady=5)


######fuel######
fuel_type_var = tk.StringVar(value="电")
fuel_type_combobox = ttk.Combobox(frame_fuel, textvariable=fuel_type_var)
fuel_type_combobox["values"] = sorted(fuel_factors.keys())
fuel_type_combobox.bind("<<ComboboxSelected>>", update_unit_options)
fuel_type_combobox.grid(row=0, column=1, pady=5)
tk.Label(frame_fuel, text="燃料类型：").grid(row=0, column=0, pady=5)

unit_var = ttk.Combobox(frame_fuel)
unit_var.grid(row=1, column=1, pady=5)
tk.Label(frame_fuel, text="单位：").grid(row=1, column=0, pady=5)

tk.Label(frame_fuel, text="使用量：").grid(row=2, column=0, pady=5)
usage_entry = tk.Entry(frame_fuel)
usage_entry.grid(row=2, column=1, pady=5)

calculate_button = tk.Button(frame_fuel, text="计算排放量", command=on_calculate_fuel)
calculate_button.grid(row=3, column=1, pady=5)

upload_csv_button_fuel = tk.Button(
    frame_fuel, text="上传CSV计算排放", command=upload_csv_and_calculate_fuel
)
upload_csv_button_fuel.grid(row=3, column=0, pady=5)


######transport######
tk.Label(frame_transport, text="运输设备：").grid(row=0, column=0, pady=5)
vehicle_type_var = tk.StringVar(value="汽油货车")
vehicle_type_combobox = ttk.Combobox(frame_transport, textvariable=vehicle_type_var)
vehicle_type_combobox["values"] = sorted(transport_factors.keys())
vehicle_type_combobox.bind("<<ComboboxSelected>>", update_spec_options)
vehicle_type_combobox.grid(row=0, column=1, pady=5)

tk.Label(frame_transport, text="设备规格：").grid(row=1, column=0, pady=5)
spec_entry = ttk.Combobox(frame_transport)
spec_entry.grid(row=1, column=1, pady=5)

tk.Label(frame_transport, text="运输量（t*km）：").grid(row=2, column=0, pady=5)
distance_entry = tk.Entry(frame_transport)
distance_entry.grid(row=2, column=1, pady=5)

calculate_transport_button = tk.Button(
    frame_transport, text="计算排放量", command=on_calculate_transport
)
calculate_transport_button.grid(row=3, column=1, pady=5)

upload_csv_button_transport = tk.Button(
    frame_transport, text="上传CSV计算排放", command=upload_csv_and_calculate_transport
)
upload_csv_button_transport.grid(row=3, column=0, pady=5)


######material######
tk.Label(frame_material, text="材料名称：").grid(row=0, column=0, pady=5)
material_name_var = tk.StringVar()
material_name_entry = ttk.Combobox(frame_material, textvariable=material_name_var)
material_name_entry["values"] = sorted(material_factors.keys())
material_name_entry.bind("<<ComboboxSelected>>", update_material_unit_options)
material_name_entry.grid(row=0, column=1, pady=5)

unit_var_material = ttk.Combobox(frame_material)
unit_var_material.grid(row=1, column=1, pady=5)
tk.Label(frame_material, text="单位：").grid(row=1, column=0, pady=5)

tk.Label(frame_material, text="损耗量：").grid(row=2, column=0, pady=5)
quantity_entry = tk.Entry(frame_material)
quantity_entry.grid(row=2, column=1, pady=5)

tk.Label(frame_material, text="投入量(需同时输入损耗率)：").grid(
    row=3, column=0, pady=5
)
input_amount_entry = tk.Entry(frame_material)
input_amount_entry.grid(row=3, column=1, pady=5)

tk.Label(frame_material, text="损耗率(需同时输入投入量)：").grid(
    row=4, column=0, pady=5
)
loss_rate_entry = tk.Entry(frame_material)
loss_rate_entry.grid(row=4, column=1, pady=5)

calculate_material_button = tk.Button(
    frame_material, text="计算排放量", command=on_calculate_material
)
calculate_material_button.grid(row=5, column=1, pady=5)

upload_csv_button_material = tk.Button(
    frame_material, text="上传CSV计算排放", command=upload_csv_and_calculate_material
)
upload_csv_button_material.grid(row=5, column=0, pady=5)


######vocs######
tk.Label(frame_vocs, text="含VOCs的材料：").grid(row=0, column=0, pady=5)
vocs_material_name_var = tk.StringVar()
vocs_material_name_entry = ttk.Combobox(frame_vocs, textvariable=vocs_material_name_var)
vocs_material_name_entry["values"] = sorted(vocs_factors.keys())
vocs_material_name_entry.bind("<<ComboboxSelected>>", update_vocs_unit_options)
vocs_material_name_entry.grid(row=0, column=1, pady=5)

vocs_unit_var = ttk.Combobox(frame_vocs)
vocs_unit_var.grid(row=1, column=1, pady=5)
tk.Label(frame_vocs, text="单位：").grid(row=1, column=0, pady=5)

tk.Label(frame_vocs, text="使用量：").grid(row=2, column=0, pady=5)
vocs_usage_entry = tk.Entry(frame_vocs)
vocs_usage_entry.grid(row=2, column=1, pady=5)

calculate_vocs_button = tk.Button(
    frame_vocs, text="计算排放量", command=on_calculate_vocs
)
calculate_vocs_button.grid(row=3, column=1, pady=5)

upload_csv_button_vocs = tk.Button(
    frame_vocs, text="上传CSV计算排放", command=upload_csv_and_calculate_vocs
)
upload_csv_button_vocs.grid(row=3, column=0, pady=5)


######初始化和启动######
# 初始化单位和规格选项
update_unit_options(None)
update_spec_options(None)
update_material_unit_options(None)
update_vocs_unit_options(None)

root.mainloop()
