import io
import os
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

def _try_register_chinese_font():
    """
    尝试在 Windows 上注册中文字体，避免 PDF 里中文变方块。
    找不到就返回 False（仍可生成 PDF，但中文可能显示异常）。
    """
    candidates = [
        r"C:\Windows\Fonts\msyh.ttc",    # 微软雅黑
        r"C:\Windows\Fonts\msyh.ttf",
        r"C:\Windows\Fonts\simsun.ttc",  # 宋体
        r"C:\Windows\Fonts\simhei.ttf",  # 黑体
    ]
    for p in candidates:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont("CN", p))
                return True
            except Exception:
                # 有些机器/版本对 .ttc 支持差异，失败就继续尝试下一个
                continue
    return False

def build_pdf(result: dict, meta: dict) -> bytes:
    """
    result: calc_plan 的计算结果 dict
    meta: 你希望写进报告的口径信息（例如 400kW 等）
    """
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    # 页面设置
    width, height = A4
    left = 20 * mm
    top = height - 20 * mm
    line_h = 8 * mm

    # 字体（尽量中文）
    has_cn = _try_register_chinese_font()
    if has_cn:
        c.setFont("CN", 16)
    else:
        # fallback：仍能生成 PDF，但中文可能显示为方块
        c.setFont("Helvetica", 16)

    # 标题
    y = top
    c.drawString(left, y, "场站建设初设建议和回报评估")
    y -= 2 * line_h

    # 计算关键字段（按你的表达口径）
    n = result["n_recommend"]
    invest_wan = result["invest_total_yuan"] / 10000.0
    revenue_wan = result["revenue_year_yuan"] / 10000.0
    payback = result["payback_years"]

    # 报告摘要
    if has_cn:
        c.setFont("CN", 12)
    else:
        c.setFont("Helvetica", 12)

    def draw_kv(k, v):
        nonlocal y
        c.drawString(left, y, f"{k}：{v}")
        y -= line_h

    draw_kv("充电设备配置", f"{n} 台 {meta.get('pile_kw', 400)}kW 一机双枪充电桩")
    draw_kv("预计投资", f"{invest_wan:.1f} 万元")
    draw_kv("预计年收入", f"{revenue_wan:.1f} 万元（按服务费口径）")
    draw_kv("预计回报周期", f"{payback:.1f} 年" if payback else "N/A")
    draw_kv("运营口径", f"{meta.get('days', 330)} 天/年，单枪 {meta.get('kwh_per_gun_per_day', 1000)} kWh/天，服务费 {meta.get('service_fee', 0.3)} 元/度")


    y -= line_h

    # 详细数据区
    if has_cn:
        c.setFont("CN", 11)
    else:
        c.setFont("Helvetica", 11)

    c.drawString(left, y, "— 关键计算结果（V1口径）—")
    y -= line_h

    details = [
        ("场地面积", f"{result['site_area_sqm']:.0f} ㎡"),
        ("电力约束桩数 n_power", str(result["n_power"])),
        ("面积约束桩数 n_area", str(result["n_area"])),
        ("推荐桩数 N", str(result["n_recommend"])),
        ("车位数", str(result["stalls"])),
        ("模块占用面积", f"{result['used_area_sqm']:.0f} ㎡"),
        ("电力投资", f"{result['invest_power_yuan']:.0f} 元"),
        ("土建投资", f"{result['invest_civil_yuan']:.0f} 元"),
        ("设备投资", f"{result['invest_pile_yuan']:.0f} 元"),
        ("总投资", f"{result['invest_total_yuan']:.0f} 元"),
        ("年充电量", f"{result['energy_year_kwh']:.0f} kWh"),
        ("年收入（服务费口径）", f"{result['revenue_year_yuan']:.0f} 元"),
        ("静态回收期", f"{result['payback_years']:.2f} 年" if result["payback_years"] else "N/A"),
    ]

    for k, v in details:
        c.drawString(left, y, f"{k}：{v}")
        y -= 7 * mm
        if y < 20 * mm:
            c.showPage()
            y = top
            if has_cn:
                c.setFont("CN", 11)
            else:
                c.setFont("Helvetica", 11)

    # 风险/提示
    if result.get("notes"):
        y -= 2 * mm
        c.drawString(left, y, "— 风险/提示 —")
        y -= 7 * mm
        for nline in result["notes"]:
            c.drawString(left, y, f"• {nline}")
            y -= 7 * mm

    c.showPage()
    c.save()
    return buf.getvalue()
