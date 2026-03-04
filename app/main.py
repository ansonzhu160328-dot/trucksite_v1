from fastapi import FastAPI, Request
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from docx import Document
import tempfile
import os
from pathlib import Path

import base64
from datetime import datetime
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH


from app.schemas import CalcRequest
from app.calc import calc_plan

app = FastAPI(title="Truck Charging Site V1", version="0.3.0")

app.mount("/static", StaticFiles(directory="static"), name="static")

PRODUCT_ASSETS_DIR = Path("/root/trucksite_v1/assets/product")
ALLOWED_IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp"}
FINANCE_TEXT = """
金融合作方案（示例）
1. 可根据项目现金流测算结果，设计分期投放与回款计划。
2. 可提供设备融资租赁、项目贷款等多种金融工具组合。
3. 可结合运营数据设置动态授信与风险预警机制。
4. 具体合作条款以双方签署的正式协议为准。
""".strip()


@app.get("/")
def home():
    return FileResponse("static/index.html")


@app.post("/api/calculate")
def calculate(req: CalcRequest):
    data = req.model_dump()
    print("DEBUG /api/calculate keys:", sorted(list(data.keys())))
    result = calc_plan(data)
    return result


@app.post("/api/report_word")
async def report_word(request: Request):
    raw_data = await request.json()
    if not isinstance(raw_data, dict):
        raw_data = {}

    req = CalcRequest.model_validate(raw_data)
    data = req.model_dump()
    print("DEBUG /api/report_word keys:", sorted(list(raw_data.keys())))
    result = calc_plan(data)

    # ===== 生成 Word =====
    doc = Document()

    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.section import WD_SECTION_START
    from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    # ===== 封皮标题 =====
    site_location = data.get("site_location", "").strip()

    if not site_location:
        site_location = "未填写场站位置"

    title_text = f"{site_location}重卡充电站初步设计方案"

    # =========================
    # 通用：字体 + 段落格式
    # =========================
    INDENT_2CH = Pt(28)  # 首行缩进 2 字符（宋体14号下约等于 28pt）
    LINE_SPACING = 1.5   # 1.5倍行距
    BULLET = "■"         # 统一条目标识符号

    def set_cn_font(run, size_pt=14, bold=False, font_name="宋体"):
        """
        中文：宋体
        英文/数字：Times New Roman
        """
        run.font.size = Pt(size_pt)
        run.bold = bold

        r = run._element
        rPr = r.get_or_add_rPr()
        rFonts = rPr.get_or_add_rFonts()

        # 中文字体
        rFonts.set(qn('w:eastAsia'), font_name)

        # 英文 & 数字字体
        rFonts.set(qn('w:ascii'), 'Times New Roman')
        rFonts.set(qn('w:hAnsi'), 'Times New Roman')


    def format_para(p, first_line_indent=False):
        p.paragraph_format.line_spacing = LINE_SPACING
        if first_line_indent:
            p.paragraph_format.first_line_indent = INDENT_2CH

    def add_cover_line(text, size_pt=14, bold=False, align_center=True):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if align_center else WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(text)
        set_cn_font(run, size_pt=size_pt, bold=bold, font_name="宋体")
        format_para(p, first_line_indent=False)

    def add_title(text):
        # 一级标题：宋体14pt加粗，不缩进，1.5倍行距
        p = doc.add_paragraph()
        run = p.add_run(text)
        set_cn_font(run, size_pt=14, bold=True, font_name="宋体")
        format_para(p, first_line_indent=False)

    def add_body(text):
        # 正文：宋体14pt不加粗，首行缩进2字符，1.5倍行距
        p = doc.add_paragraph()
        run = p.add_run(text)
        set_cn_font(run, size_pt=14, bold=False, font_name="宋体")
        format_para(p, first_line_indent=True)

    def add_item(text):
        # 条目：用 ■ 符号；不做首行缩进（避免符号被挤歪），1.5倍行距
        p = doc.add_paragraph()
        run = p.add_run(f"{BULLET} {text}")
        set_cn_font(run, size_pt=14, bold=False, font_name="宋体")
        format_para(p, first_line_indent=False)

    def add_blank_line():
        # 章节结束空一行
        p = doc.add_paragraph("")
        p.paragraph_format.line_spacing = LINE_SPACING

    def _hide_table_borders(table):
        tbl = table._tbl
        tbl_pr = tbl.tblPr
        borders = OxmlElement('w:tblBorders')
        for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            elem = OxmlElement(f'w:{edge}')
            elem.set(qn('w:val'), 'nil')
            borders.append(elem)
        tbl_pr.append(borders)

    def add_cover_page():
        section1 = doc.sections[0]
        usable_height = section1.page_height - section1.top_margin - section1.bottom_margin

        table = doc.add_table(rows=3, cols=1)
        _hide_table_borders(table)

        row_heights = [int(usable_height * 0.3), int(usable_height * 0.4), int(usable_height * 0.3)]
        for idx, row in enumerate(table.rows):
            row.height = row_heights[idx]
            row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

        # 中间：标题（水平+垂直居中）
        mid_cell = table.cell(1, 0)
        mid_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        mid_para = mid_cell.paragraphs[0]
        mid_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = mid_para.add_run(title_text)
        set_cn_font(run, size_pt=22, bold=True, font_name="宋体")
        format_para(mid_para, first_line_indent=False)

        # 底部：编制单位/日期（左对齐+底部对齐）
        bottom_cell = table.cell(2, 0)
        bottom_cell.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM

        p1 = bottom_cell.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r1 = p1.add_run("编制单位：广东盈通智联数字技术有限公司")
        set_cn_font(r1, size_pt=14, bold=False, font_name="宋体")
        format_para(p1, first_line_indent=False)

        p2 = bottom_cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r2 = p2.add_run(f"编制日期：{datetime.now().strftime('%Y年%m月%d日')}")
        set_cn_font(r2, size_pt=14, bold=False, font_name="宋体")
        format_para(p2, first_line_indent=False)

    def _set_section_page_start(section, start_num=1):
        sect_pr = section._sectPr
        for child in list(sect_pr):
            if child.tag == qn('w:pgNumType'):
                sect_pr.remove(child)
        pg_num_type = OxmlElement('w:pgNumType')
        pg_num_type.set(qn('w:start'), str(start_num))
        sect_pr.append(pg_num_type)

    def _add_footer_page_field(section):
        footer = section.footer
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        run = p.add_run()
        fld_begin = OxmlElement('w:fldChar')
        fld_begin.set(qn('w:fldCharType'), 'begin')

        instr = OxmlElement('w:instrText')
        instr.set(qn('xml:space'), 'preserve')
        instr.text = ' PAGE '

        fld_separate = OxmlElement('w:fldChar')
        fld_separate.set(qn('w:fldCharType'), 'separate')

        fld_end = OxmlElement('w:fldChar')
        fld_end.set(qn('w:fldCharType'), 'end')

        run._r.append(fld_begin)
        run._r.append(instr)
        run._r.append(fld_separate)
        run._r.append(fld_end)
    
    def normalize_attachments_selected(value):
        if isinstance(value, list):
            items = value
        elif isinstance(value, str):
            items = [value]
        else:
            items = []

        normalized = []
        for item in items:
            if not isinstance(item, str):
                continue
            v = item.strip().lower()
            if v in {"layout", "product", "finance"} and v not in normalized:
                normalized.append(v)
        return normalized

    def add_attach_title(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(text)
        set_cn_font(run, size_pt=14, bold=True, font_name="宋体")
        format_para(p, first_line_indent=False)

    def add_attach_hint(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(text)
        set_cn_font(run, size_pt=14, bold=False, font_name="宋体")
        format_para(p, first_line_indent=False)

    def parse_layout_png_data_url(layout_png_data_url: str):
        if not layout_png_data_url:
            return None
        b64 = layout_png_data_url
        if "base64," in layout_png_data_url:
            b64 = layout_png_data_url.split("base64,", 1)[1]
        try:
            return base64.b64decode(b64)
        except Exception as e:
            print("WARN layout_png_data_url decode failed:", e)
            return None

    def append_layout_attachment(data_dict: dict):
        add_attach_title((data_dict.get("layout_title") or "附件1：场站布局示意图").strip())

        layout_png_data_url = (data_dict.get("layout_png_data_url") or "").strip()
        img_bytes = parse_layout_png_data_url(layout_png_data_url)
        if not img_bytes:
            add_attach_hint("（未获取到布局图）")
            return

        tmp_png = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        try:
            tmp_png.write(img_bytes)
            tmp_png.close()
            doc.add_picture(tmp_png.name, width=Cm(15))
        except Exception as e:
            print("WARN append layout image failed:", e)
            add_attach_hint("（未获取到布局图）")
        finally:
            try:
                os.remove(tmp_png.name)
            except Exception:
                pass

    def append_product_attachment():
        add_attach_title("附件2：产品及典型案例")

        if not PRODUCT_ASSETS_DIR.exists() or not PRODUCT_ASSETS_DIR.is_dir():
            add_attach_hint("（未配置产品图片）")
            return

        image_files = [
            p for p in PRODUCT_ASSETS_DIR.iterdir()
            if p.is_file() and p.suffix.lower() in ALLOWED_IMAGE_EXTS
        ]
        image_files.sort(key=lambda x: x.name)

        if not image_files:
            add_attach_hint("（未配置产品图片）")
            return

        inserted = False
        for image_path in image_files:
            try:
                if image_path.suffix.lower() == ".webp":
                    print("INFO skip unsupported webp for python-docx:", str(image_path))
                    continue
                doc.add_picture(str(image_path), width=Cm(15))
                doc.add_paragraph("")
                inserted = True
            except Exception as e:
                print("WARN append product image failed:", str(image_path), e)

        if not inserted:
            add_attach_hint("（未配置产品图片）")

    def append_finance_attachment():
        add_attach_title("附件3：金融合作方案")

        text = (FINANCE_TEXT or "").strip()
        if not text:
            add_attach_hint("（未配置金融方案文本）")
            return

        for line in [x.strip() for x in text.splitlines() if x.strip()]:
            add_body(line)


    # =========================
    # 封皮页（第1页）+ 分节（正文从第2页开始）
    # =========================
    add_cover_page()

    section2 = doc.add_section(WD_SECTION_START.NEW_PAGE)

    section2.footer.is_linked_to_previous = False
    section2.header.is_linked_to_previous = False
    _set_section_page_start(section2, start_num=1)
    _add_footer_page_field(section2)


    # =========================
    # 一、场站基本信息
    # =========================
    add_title("一、场站基本信息")

    add_body(
        f"充电站位于{data.get('site_location','')}，"
        f"长度约{data.get('site_length_m',0)}米、"
        f"宽度约{data.get('site_width_m',0)}米、"
        f"租金{data.get('rent_yuan_per_sqm_month',0)}元/月。"
    )

    add_blank_line()


    # =========================
    # 二、充电站初步设计
    # =========================
    add_title("二、充电站初步设计")

    add_body("结合场地信息、附近车队情况、电力容量，初步建议：")

    add_item(f"电力增容：{int(result.get('power_capacity_kva',0))}kVA。")
    add_item(f"充电设备配置：{result.get('n_recommend',0)}台400kW一体机。")

    add_blank_line()


    # =========================
    # 三、投资估算
    # =========================
    add_title("三、投资估算")

    add_body(
        f"根据充电站初步设计方案，预计总投资"
        f"{round(result.get('invest_total_yuan',0)/10000,2)}万元，其中："
    )

    add_item(f"电力增容：{round(result.get('invest_power_yuan',0)/10000,2)}万元。")
    add_item(f"场地平整：{round(result.get('invest_civil_yuan',0)/10000,2)}万元。")
    add_item(f"充电设备：{round(result.get('invest_pile_yuan',0)/10000,2)}万元。")

    add_blank_line()



    # =========================
    # 编号行（用于 1、2、3 这种）
    # =========================
    def add_numbered(text):
        # 编号行：宋体14，不加粗，1.5倍行距，不做首行缩进（避免“1、”被缩进挤歪）
        p = doc.add_paragraph()
        run = p.add_run(text)
        set_cn_font(run, size_pt=14, bold=False, font_name="宋体")
        format_para(p, first_line_indent=False)

    # =========================
    # 后端重算敏感性分析（27组）
    # =========================
    def calc_sensitivity_27(base_data: dict):
        """
        复刻前端的 27 组：kwh(0.6/1.0/1.2) × fee(0.8/1.0/1.2) × rent(0/1.0/1.5)
        输出：baseline/best/worst 三个情景（都带条件与净收益、回本期）
        """
        import copy

        base_kwh = float(base_data.get("kwh_per_gun_per_day", 1000.0))
        base_fee = float(base_data.get("service_fee_yuan_per_kwh", 0.3))
        base_rent = float(base_data.get("rent_yuan_per_sqm_month", 0.0))

        kwh_levels = [round(base_kwh * 0.6), round(base_kwh * 1.0), round(base_kwh * 1.2)]
        fee_levels = [round(base_fee * 0.8, 2), round(base_fee * 1.0, 2), round(base_fee * 1.2, 2)]
        rent_levels = [0.0, round(base_rent * 1.0, 2), round(base_rent * 1.5, 2)]

        rows = []
        idx = 0
        for kwh in kwh_levels:
            for fee in fee_levels:
                for rent in rent_levels:
                    idx += 1
                    p = copy.deepcopy(base_data)
                    p["kwh_per_gun_per_day"] = kwh
                    p["service_fee_yuan_per_kwh"] = fee
                    p["rent_yuan_per_sqm_month"] = rent

                    r = calc_plan(p)

                    net_yuan = float(r.get("revenue_net_year_yuan", 0.0) or 0.0)
                    pb = r.get("payback_net_years", None)
                    pb_val = float(pb) if pb is not None else None

                    # 跟你前端一致的状态规则
                    status = "🟢"
                    if net_yuan <= 0:
                        status = "🔴"
                    elif pb_val is not None and pb_val > 3:
                        status = "🔴"
                    elif pb_val is not None and pb_val > 2:
                        status = "🟡"

                    rows.append({
                        "idx": idx,
                        "kwh": kwh,
                        "fee": fee,
                        "rent": rent,
                        "net_wan": net_yuan / 10000.0,
                        "pb": pb_val,
                        "status": status
                    })

        # baseline：中档组合
        def is_same(a, b, eps=1e-9):
            return abs(float(a) - float(b)) < eps

        baseline = None
        for x in rows:
            if x["kwh"] == kwh_levels[1] and is_same(x["fee"], fee_levels[1]) and is_same(x["rent"], rent_levels[1]):
                baseline = x
                break
        if baseline is None:
            baseline = rows[0]

        # best：在可回收（不红）里回本期最小；如果全红，则取净收益最大的那个
        good = [x for x in rows if x["status"] != "🔴" and x["pb"] is not None]
        if good:
            best = sorted(good, key=lambda x: x["pb"])[0]
        else:
            best = sorted(rows, key=lambda x: x["net_wan"], reverse=True)[0]

        # worst：优先找红里最差（净收益最小），否则回本期最大的
        bad = [x for x in rows if x["status"] == "🔴"]
        if bad:
            worst = sorted(bad, key=lambda x: x["net_wan"])[0]
        else:
            # 全部非红就取 pb 最大
            tmp = [x for x in rows if x["pb"] is not None]
            worst = sorted(tmp, key=lambda x: x["pb"], reverse=True)[0] if tmp else rows[-1]

        return baseline, best, worst


    # =========================
    # 四、投资回报
    # =========================
    add_title("四、投资回报")

    net_income_wan = round(float(result.get("revenue_net_year_yuan", 0.0) or 0.0) / 10000.0, 2)
    payback_net = result.get("payback_net_years", None)
    payback_net_text = f"{round(float(payback_net), 2)}" if payback_net is not None else "N/A"

    add_body(
        f"根据附近车流量信息，初步估算，每年净收入约{net_income_wan}万元/年，"
        f"投资回报期{payback_net_text}年。估算的主要边界条件包括："
    )

    add_item(f"单枪充电量：{int(data.get('kwh_per_gun_per_day', 0))}度/枪/天。")
    add_item(f"充电服务费：{round(float(data.get('service_fee_yuan_per_kwh', 0.0)), 2)}元/度。")
    add_item(f"运行天数：{int(data.get('days_per_year', 0))}天/年。")
    add_item(f"运营人员：{int(data.get('staff_count', 0))}人。")
    add_item(f"人员工资：{int(data.get('salary_yuan_per_month', 0))}元/月。")

    add_blank_line()


    # =========================
    # 五、敏感性分析
    # =========================
    add_title("五、敏感性分析")

    add_body("对充电量、充电服务费等关键影响因素进行了敏感性分析（合计27组，详见附件），关键结论如下：")

    baseline, best, worst = calc_sensitivity_27(data)

    def fmt_sens(x):
        kwh = x["kwh"]
        fee = x["fee"]
        rent = x["rent"]
        net = round(x["net_wan"], 2)
        pb = x["pb"]
        pb_text = f"{round(pb, 2)}" if pb is not None else "N/A"
        return f"充电量{kwh}度/枪/天、服务费{fee}元/度、场地租金{rent}元/㎡·月条件下，年净收入{net}万元，回本期{pb_text}年"

    add_numbered(f"1、最佳情况：{fmt_sens(best)}；")
    add_numbered(f"2、最差情况：{fmt_sens(worst)}；")
    add_numbered(f"3、常规情况：{fmt_sens(baseline)}；")

    add_blank_line()


    # =========================
    # 六、结论与建议
    # =========================
    add_title("六、结论与建议")

    site_loc = data.get("site_location", "") or ""

    invest_total_wan = round(float(result.get("invest_total_yuan", 0.0) or 0.0) / 10000.0, 2)
    pb_norm = baseline.get("pb", None)
    net_norm = baseline.get("net_wan", 0.0)

    # 投资收益评价（你可以后面再精修口径）
    if net_norm <= 0 or pb_norm is None:
        level_text = "不太理想"
    elif pb_norm <= 3:
        level_text = "较好"
    elif pb_norm <= 4:
        level_text = "一般"
    else:
        level_text = "不太理想"

    pb_norm_text = f"{round(pb_norm, 2)}" if pb_norm is not None else "N/A"

    # 1）常规结论
    add_numbered(
        f"1、{site_loc}重卡充电站预计总投资{invest_total_wan}万元、投资回报期{pb_norm_text}年，"
        f"投资收益{level_text}（常规情况：{fmt_sens(baseline)}）；"
    )

    # 2）最差情景提醒
    add_numbered(
        f"2、在充电量{worst['kwh']}度/枪/天、服务费{worst['fee']}元/度、场地租金{worst['rent']}元/㎡·月条件下，"
        f"投资收益最差，需重点关注风险。"
    )

    add_blank_line()


    # ===== 文末附件：按前端选择动态插入 =====
    attachments_selected = normalize_attachments_selected(raw_data.get("attachments_selected", []))

    if "layout" in attachments_selected:
        append_layout_attachment(raw_data)

    if "product" in attachments_selected:
        append_product_attachment()

    if "finance" in attachments_selected:
        append_finance_attachment()

    # ===== 保存为临时文件并下载 =====
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(tmp.name)
    tmp.close()

    return FileResponse(
        tmp.name,
        filename="trucksite_preliminary_design.docx",  # 用 ASCII 文件名，避免 Windows/Header 坑
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
