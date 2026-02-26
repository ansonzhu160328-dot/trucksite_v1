import math

def _f(x, default=0.0):
    """安全取 float（None/缺失/NaN 都兜住）"""
    try:
        if x is None:
            return float(default)
        return float(x)
    except Exception:
        return float(default)

def _i(x, default=0):
    """安全取 int"""
    try:
        if x is None:
            return int(default)
        return int(x)
    except Exception:
        return int(default)

def calc_plan(d: dict) -> dict:
    # --- 输入（全部兜底，避免 KeyError） ---
    site_length = _f(d.get("site_length_m"), 0)
    site_width  = _f(d.get("site_width_m"), 0)

    pile_kva_per = _f(d.get("pile_kva_per"), 400)
    guns_per_pile = _i(d.get("guns_per_pile"), 2)
    kwh_per_gun_per_day = _f(d.get("kwh_per_gun_per_day"), 1000)
    service_fee = _f(d.get("service_fee_yuan_per_kwh"), 0.3)
    days_per_year = _i(d.get("days_per_year"), 330)

    power_cost = _f(d.get("power_cost_yuan_per_kva"), 600)     # 电力投资单价（元/kVA）
    civil_cost = _f(d.get("civil_cost_yuan_per_sqm"), 200)
    pile_cost = _f(d.get("pile_cost_yuan_each"), 45000)

   
    rent_yuan_per_sqm_month = _f(d.get("rent_yuan_per_sqm_month"), 0)
    staff_count = _i(d.get("staff_count"), 0)
    salary_yuan_per_month = _f(d.get("salary_yuan_per_month"), 0)

    # --- 核心约束（按场地布置→车位→桩数→电力） ---
    site_area = site_length * site_width

    # =========================
    # 口径（你定义的工程经验）
    # =========================
    STALL_WIDTH_M = 6.0                  # 单车位宽（重卡车宽口径）
    REQ_LEN_MIN_M = STALL_WIDTH_M * 2    # 长度<12：不具备建站（按2个车位宽）
    REQ_WIDTH_MIN_M = 30.0               # 宽度<30：转弯半径不足，不具备建站

    # 电力口径（新）
    PILE_KVA_RULE = 200.0                # 仍保留：仅用于“旧字段 transformer_required_kva”的兼容（见输出）
    PILE_KVA_POWER = 400.0               # 电力容量=桩数*400kVA（你要求）

    # 1) 长度决定：每排可布置车位数 = floor(长度/STALL_WIDTH_M)
    stalls_per_row_raw = int(site_length // STALL_WIDTH_M) if site_length >= REQ_LEN_MIN_M else 0

    # 2) 宽度决定：可布置几排（数据驱动：分段表 + 查表）
    row_count = 0
    layout_note = ""

    # 宽度分段口径（延伸到 500m；>500 提示人工评估）
    # 规则：从 30m 开始，区间宽度按 +15 / +30 交替增长，对应排数逐段 +1
    WIDTH_BANDS = [
        (30, 45, 1),
        (45, 75, 2),
        (75, 90, 3),
        (90, 120, 4),
        (120, 135, 5),
        (135, 165, 6),
        (165, 180, 7),
        (180, 210, 8),
        (210, 225, 9),
        (225, 255, 10),
        (255, 270, 11),
        (270, 300, 12),
        (300, 315, 13),
        (315, 345, 14),
        (345, 360, 15),
        (360, 390, 16),
        (390, 405, 17),
        (405, 435, 18),
        (435, 450, 19),
        (450, 480, 20),
        (480, 495, 21),
        (495, 500, 22),  # 到 500m 为止（含 500）
    ]

    def _rows_for_width(w: float):
        """返回 (rows, note)；w>500 返回(0,人工评估提示)"""
        if w < REQ_WIDTH_MIN_M:
            return 0, f"场地宽度{w:.1f}m<{REQ_WIDTH_MIN_M:.0f}m：转弯半径不足，不具备建站条件。"
        if w > 500:
            return 0, f"场地宽度{w:.1f}m>500m：超出当前口径范围，请人工评估。"
        for a, b, rows in WIDTH_BANDS:
            if (w >= a and w < b) or (b == 500 and w == 500):
                return rows, f"{a}m≤宽度{w:.1f}m<{b}m：可布置{rows}排车位。"
        return 0, f"场地宽度{w:.1f}m：未命中宽度分段口径，请人工评估。"

    row_count, layout_note = _rows_for_width(site_width)

    # 3) 单排绘图口径修正：变压器左右两侧车位数都要求为偶数
    stalls_per_row_draw = stalls_per_row_raw
    stalls_left = 0
    stalls_right = 0

    if row_count == 1:
        s = stalls_per_row_raw
        if s < 2:
            stalls_per_row_draw = 0
            layout_note = (layout_note + "；" if layout_note else "") + "单排可用车位数不足2，无法按变压器居中口径布置。"
        else:
            if s % 2 == 1:
                old_s = s
                s -= 1
                layout_note = (
                    layout_note + "；" if layout_note else ""
                ) + f"单排要求车位总数为偶数（2车位/桩），已从{old_s}调整为{s}。"

            left = s // 2
            right = s - left

            if left % 2 == 1:
                left -= 1
                right += 1

            stalls_left = left
            stalls_right = right
            stalls_per_row_draw = s
            layout_note = (
                layout_note + "；" if layout_note else ""
            ) + f"单排变压器左右两侧车位数需为偶数（避免3+3），已拆分为{left}+{right}。"

    # 4) 车位数量（单排按绘图口径，多排按原口径）
    if row_count == 1:
        stalls_total = stalls_per_row_draw
    else:
        stalls_total = stalls_per_row_raw * row_count

    # 5) 桩数量（布局口径）：桩 = 车位/2（取整）
    n_layout = stalls_total // 2

    # 5) 电力约束：现在“变压器容量不再输入”，先不做上限约束（无限大）
    #    如果未来你要引入“电网接入上限/客户可获批容量”，再把 n_power 改为 floor(可获批kVA/400)。
    n_power = 10**10

   
    # 6) 推荐桩数：二者取最小
    n_recommend = max(0, min(n_layout, n_power))

    # 7) 电力容量（kVA）：按你新口径 = 桩数 * 400kVA
    power_capacity_kva = n_recommend * PILE_KVA_POWER

    
    # --- CAPEX（你要求：桩=0 → 投资=0） ---
    if n_recommend <= 0:
        invest_power = 0.0
        invest_civil = 0.0
        invest_pile = 0.0
        invest_total = 0.0
    else:
        invest_power = power_cost * power_capacity_kva
        invest_civil = civil_cost * site_area
        invest_pile = pile_cost * n_recommend
        invest_total = invest_power + invest_civil + invest_pile

    # --- 收入（服务费口径） ---
    energy_year = n_recommend * guns_per_pile * kwh_per_gun_per_day * days_per_year
    revenue_year = service_fee * energy_year

    payback_years = None
    if revenue_year > 0 and invest_total > 0:
        payback_years = invest_total / revenue_year

    # --- OPEX: 租金、人工、净现金流（你要求：桩=0 → 全部0） ---
    if n_recommend <= 0:
        rent_year_yuan = 0.0
        labor_year_yuan = 0.0
        revenue_net_year_yuan = 0.0
    else:
        rent_year_yuan = site_area * rent_yuan_per_sqm_month * 12
        labor_year_yuan = staff_count * salary_yuan_per_month * 12
        revenue_net_year_yuan = revenue_year - rent_year_yuan - labor_year_yuan

    payback_net_years = None
    if revenue_net_year_yuan > 0 and invest_total > 0:
        payback_net_years = invest_total / revenue_net_year_yuan

    # --- notes / 提示（边界&口径解释） ---
    notes = []

    if site_length < REQ_LEN_MIN_M:
        notes.append(f"场地长度{site_length:.1f}m<{REQ_LEN_MIN_M:.0f}m：场地不足，不具备建站条件。")
    if site_width < REQ_WIDTH_MIN_M:
        notes.append(f"场地宽度{site_width:.1f}m<{REQ_WIDTH_MIN_M:.0f}m：转弯半径不足，不具备建站条件。")

    if layout_note:
        notes.append(layout_note)

    notes.append(
        f"布置口径：每排原始车位数=floor(长度/{STALL_WIDTH_M:.0f})={stalls_per_row_raw}；绘图每排车位数={stalls_per_row_draw}；排数={row_count}；车位={stalls_total}；桩(布局)=车位/2={n_layout}。"
    )

    notes.append(
        f"电力口径：电力容量=桩数×400kVA={n_recommend}×400={power_capacity_kva:.0f}kVA；电力投资=单价×电力容量={power_cost:.0f}×{power_capacity_kva:.0f}。"
    )

    
    if n_recommend <= 0:
        notes.append("推荐桩数为0：不建议硬化场地/投资建设（CAPEX按0处理）。")
    else:
        if n_recommend < n_layout:
            notes.append("受电力或面积约束：推荐桩数小于布局可布置桩数。")
        if n_recommend < n_power:
            notes.append("受面积或布局约束：推荐桩数小于电力可支持桩数。")        
        if revenue_net_year_yuan <= 0:
            notes.append("经营口径净现金流<=0：租金/人工假设较高或服务费较低，项目可能不具备回收性。")

    # --- 输出（字段永远存在，前端不会 NaN） ---
    return {
        "site_area_sqm": site_area,

        "stalls_per_row": stalls_per_row_draw,
        "stalls_per_row_raw": stalls_per_row_raw,
        "stalls_per_row_draw": stalls_per_row_draw,
        "row_count": row_count,
        "layout_note": layout_note,
        "stalls": stalls_total,
        "stalls_total": stalls_total,
        "stalls_left": stalls_left,
        "stalls_right": stalls_right,
        "n_layout": n_layout,

        "n_power": n_power,        
        "n_recommend": n_recommend,

        # 新增：给前端“推荐配置”展示电力容量
        "power_capacity_kva": power_capacity_kva,

        # 旧字段兼容：你之前用过 transformer_required_kva，这里保留但不再作为输入
        "transformer_required_kva": n_recommend * PILE_KVA_RULE,

        
        "invest_power_yuan": invest_power,
        "invest_civil_yuan": invest_civil,
        "invest_pile_yuan": invest_pile,
        "invest_total_yuan": invest_total,

        "energy_year_kwh": energy_year,
        "revenue_year_yuan": revenue_year,
        "payback_years": payback_years,

        "rent_year_yuan": rent_year_yuan,
        "labor_year_yuan": labor_year_yuan,
        "revenue_net_year_yuan": revenue_net_year_yuan,
        "payback_net_years": payback_net_years,

        "notes": notes,
    }
