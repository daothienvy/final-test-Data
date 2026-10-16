#!/usr/bin/env python3
"""
Generates the comprehensive 10-slide PowerPoint presentation for the Capstone Final Project.
Adheres strictly to the executive styling, system color palette, and data analyzed in:
- Bao_Cao_Tot_Nghiep_Cuoi_Khoa.docx
- Power BI Interactive Dashboard (https://daothienvy.github.io/final-test-Data/)
"""

import os
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def create_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    blank_layout = prs.slide_layouts[6]

    # Colors
    C_NAVY_DARK = RGBColor(10, 25, 47)      # #0A192F
    C_NAVY_BLUE = RGBColor(30, 58, 138)     # #1E3A8A
    C_OCEAN_BLUE = RGBColor(37, 99, 235)    # #2563EB
    C_EMERALD = RGBColor(16, 185, 129)      # #10B981
    C_CORAL = RGBColor(239, 68, 68)         # #EF4444
    C_AMBER = RGBColor(245, 158, 11)        # #F59E0B
    C_ICE_GRAY = RGBColor(248, 250, 252)    # #F8FAFC
    C_CARD_BG = RGBColor(255, 255, 255)
    C_TEXT_DARK = RGBColor(15, 23, 42)
    C_TEXT_MUTED = RGBColor(100, 116, 139)
    C_WHITE = RGBColor(255, 255, 255)
    C_BORDER_LIGHT = RGBColor(226, 232, 240)

    def add_header(slide, title, category="BÁO CÁO TỐT NGHIỆP DATA ANALYTICS"):
        # Header banner shape
        header_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.1))
        header_shape.fill.solid()
        header_shape.fill.fore_color.rgb = C_NAVY_DARK
        header_shape.line.fill.background()

        # Accent line under header
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.1), Inches(13.333), Inches(0.05))
        line.fill.solid()
        line.fill.fore_color.rgb = C_OCEAN_BLUE
        line.line.fill.background()

        txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.12), Inches(11.733), Inches(0.85))
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p1 = tf.paragraphs[0]
        p1.text = category.upper()
        p1.font.size = Pt(9.5)
        p1.font.bold = True
        p1.font.color.rgb = C_AMBER

        p2 = tf.add_paragraph()
        p2.text = title
        p2.font.size = Pt(17)
        p2.font.bold = True
        p2.font.color.rgb = C_WHITE

    def add_card(slide, left, top, width, height, bg_color=C_CARD_BG, border_color=C_BORDER_LIGHT):
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = bg_color
        if border_color:
            shape.line.color.rgb = border_color
            shape.line.width = Pt(1)
        else:
            shape.line.fill.background()
        return shape

    # =========================================================================
    # SLIDE 1: TITLE SLIDE (Trang Bìa)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = C_NAVY_DARK
    bg1.line.fill.background()

    # Decorative top accent
    acc1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.15))
    acc1.fill.solid()
    acc1.fill.fore_color.rgb = C_OCEAN_BLUE
    acc1.line.fill.background()

    # Badge box
    badge = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(1.2), Inches(3.6), Inches(0.45))
    badge.fill.solid()
    badge.fill.fore_color.rgb = RGBColor(30, 58, 138)
    badge.line.fill.background()
    tf_b = badge.text_frame
    p_b = tf_b.paragraphs[0]
    p_b.text = "★ KHÓA LUẬN TỐT NGHIỆP DATA ANALYTICS"
    p_b.font.size = Pt(10)
    p_b.font.bold = True
    p_b.font.color.rgb = C_AMBER
    p_b.alignment = PP_ALIGN.CENTER

    # Title & Subtitle text box
    tb_title = s1.shapes.add_textbox(Inches(1.2), Inches(1.9), Inches(11.0), Inches(2.8))
    tf_t = tb_title.text_frame
    tf_t.word_wrap = True

    p_t1 = tf_t.paragraphs[0]
    p_t1.text = "PHÂN TÍCH HIỆU QUẢ KINH DOANH THƯƠNG MẠI ĐIỆN TỬ\nVÀ VẬN HÀNH CHUỖI CUNG ỨNG CHẶNG CUỐI"
    p_t1.font.size = Pt(28)
    p_t1.font.bold = True
    p_t1.font.color.rgb = C_WHITE
    p_t1.space_after = Pt(14)

    p_t2 = tf_t.add_paragraph()
    p_t2.text = "Quy trình Thực thi Toàn diện: EDA Data Cleansing, Mô hình Star Schema, Phân khúc RFM Khách hàng,\nXây dựng Dashboard Power BI Đa chiều & Hệ thống Báo cáo Chiến lược Điều hành"
    p_t2.font.size = Pt(13)
    p_t2.font.color.rgb = RGBColor(148, 163, 184)

    # Info Card in Slide 1
    info_card = add_card(s1, Inches(1.2), Inches(4.8), Inches(10.933), Inches(2.0), RGBColor(15, 30, 55), RGBColor(30, 58, 138))
    tb_info = s1.shapes.add_textbox(Inches(1.5), Inches(5.0), Inches(10.3), Inches(1.6))
    tf_i = tb_info.text_frame
    tf_i.word_wrap = True

    p_i1 = tf_i.paragraphs[0]
    p_i1.text = "👤 Học viên thực hiện: Đào Thiện Vỹ                 🏷️ Mã học viên: DA-FINAL-2026                 📅 Thời gian: Tháng 10/2026"
    p_i1.font.size = Pt(11.5)
    p_i1.font.bold = True
    p_i1.font.color.rgb = C_WHITE
    p_i1.space_after = Pt(10)

    p_i2 = tf_i.add_paragraph()
    p_i2.text = "📊 Tập dữ liệu thực nghiệm: 89,316 Đơn hàng thực tế giai đoạn 2016 - 2018 | Star Schema Fact & Dimensions"
    p_i2.font.size = Pt(11)
    p_i2.font.color.rgb = RGBColor(203, 213, 225)
    p_i2.space_after = Pt(8)

    p_i3 = tf_i.add_paragraph()
    p_i3.text = "🌐 Live Interactive Dashboard: https://daothienvy.github.io/final-test-Data/"
    p_i3.font.size = Pt(11)
    p_i3.font.bold = True
    p_i3.font.color.rgb = C_EMERALD

    # =========================================================================
    # SLIDE 2: TỔNG QUAN QUẢN TRỊ & 4 CHỈ SỐ KPI CỐT LÕI (Executive Overview)
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "TỔNG QUAN QUẢN TRỊ & BỨC TRANH VẬN HÀNH TOÀN CẢNH", "PHẦN 1: MỤC TIÊU CHIẾN LƯỢC")

    # Core Question Banner
    banner2 = add_card(s2, Inches(0.8), Inches(1.35), Inches(11.733), Inches(1.0), RGBColor(239, 246, 255), RGBColor(191, 219, 254))
    tb_b2 = s2.shapes.add_textbox(Inches(1.0), Inches(1.45), Inches(11.3), Inches(0.8))
    tf_b2 = tb_b2.text_frame
    tf_b2.word_wrap = True
    p_b2_1 = tf_b2.paragraphs[0]
    p_b2_1.text = "🎯 CÂU HỎI QUẢN TRỊ TRUNG TÂM (CORE BUSINESS QUESTION):"
    p_b2_1.font.size = Pt(10.5)
    p_b2_1.font.bold = True
    p_b2_1.font.color.rgb = C_NAVY_BLUE
    p_b2_2 = tf_b2.add_paragraph()
    p_b2_2.text = "“Làm sao để Doanh nghiệp Thương mại Điện tử tối ưu hóa Doanh thu ròng, nâng cao Tỷ lệ Giữ chân Khách hàng và cắt giảm Chi phí Vận hành Logistics chặng cuối?”"
    p_b2_2.font.size = Pt(12)
    p_b2_2.font.bold = True
    p_b2_2.font.color.rgb = C_NAVY_DARK

    # 4 Top KPI Cards
    kpis = [
        ("TỔNG DOANH THU (REVENUE)", "$30.45M", "89,316 đơn hàng", "▲ +24.8% YoY tăng trưởng", C_EMERALD),
        ("GIÁ TRỊ ĐƠN HÀNG TB (AOV)", "$340.90", "54.7% đơn > $200", "Động lực từ Trả góp 3.67 kỳ", C_OCEAN_BLUE),
        ("RÒ RỈ CƯỚC SHIPPING", "12.99%", "$3,955,199.17 cước phí", "⚠️ Vượt xa trần an toàn (8%)", C_CORAL),
        ("THỜI GIAN GIAO TB (LEAD TIME)", "12.43 Ngày", "Đúng hạn (<3d): 5.01%", "🚨 Tắc nghẽn tại RJ & Bahia", C_AMBER)
    ]

    for idx, (title, val, sub1, sub2, col) in enumerate(kpis):
        left = Inches(0.8 + idx * 3.0)
        card = add_card(s2, left, Inches(2.55), Inches(2.75), Inches(2.4), C_CARD_BG, C_BORDER_LIGHT)
        # Accent top bar for card
        c_bar = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, Inches(2.55), Inches(2.75), Inches(0.08))
        c_bar.fill.solid()
        c_bar.fill.fore_color.rgb = col
        c_bar.line.fill.background()

        tb_k = s2.shapes.add_textbox(left + Inches(0.2), Inches(2.75), Inches(2.35), Inches(2.0))
        tf_k = tb_k.text_frame
        tf_k.word_wrap = True

        pk1 = tf_k.paragraphs[0]
        pk1.text = title
        pk1.font.size = Pt(9)
        pk1.font.bold = True
        pk1.font.color.rgb = C_TEXT_MUTED
        pk1.space_after = Pt(8)

        pk2 = tf_k.add_paragraph()
        pk2.text = val
        pk2.font.size = Pt(24)
        pk2.font.bold = True
        pk2.font.color.rgb = C_NAVY_DARK
        pk2.space_after = Pt(6)

        pk3 = tf_k.add_paragraph()
        pk3.text = sub1
        pk3.font.size = Pt(9.5)
        pk3.font.color.rgb = C_TEXT_DARK
        pk3.space_after = Pt(4)

        pk4 = tf_k.add_paragraph()
        pk4.text = sub2
        pk4.font.size = Pt(9)
        pk4.font.bold = True
        pk4.font.color.rgb = col

    # Bottom 3 Pillar Overview Cards
    pillars = [
        ("Trụ Cột 1: Doanh Thu & Sản Phẩm", "Tập trung cực lớn vào ngành hàng Toys ($22.99M - 75.5%). Cần kiểm soát các danh mục rò rỉ phí ship > 15%."),
        ("Trụ Cột 2: Khách Hàng & Giữ Chân", "82% khách hàng thuộc nhóm At-Risk (>90 ngày). Cần tập trung nguồn lực kích hoạt lại 5,281 VIP (> $1,000)."),
        ("Trụ Cột 3: Chuỗi Cung Ứng & Vận Hành", "Lead Time 12.43 ngày kéo dài rủi ro rời bỏ. Điểm nghẽn nghiêm trọng tại Bahia (19.33 ngày) và Rio de Janeiro (15.21 ngày).")
    ]
    for idx, (p_title, p_desc) in enumerate(pillars):
        left = Inches(0.8 + idx * 4.05)
        p_card = add_card(s2, left, Inches(5.15), Inches(3.8), Inches(1.85), C_ICE_GRAY, C_BORDER_LIGHT)
        tb_p = s2.shapes.add_textbox(left + Inches(0.2), Inches(5.3), Inches(3.4), Inches(1.5))
        tf_p = tb_p.text_frame
        tf_p.word_wrap = True
        pp1 = tf_p.paragraphs[0]
        pp1.text = p_title
        pp1.font.size = Pt(11)
        pp1.font.bold = True
        pp1.font.color.rgb = C_NAVY_BLUE
        pp1.space_after = Pt(6)
        pp2 = tf_p.add_paragraph()
        pp2.text = p_desc
        pp2.font.size = Pt(9.5)
        pp2.font.color.rgb = C_TEXT_DARK

    # =========================================================================
    # SLIDE 3: QUY TRÌNH EDA & LÀM SẠCH DỮ LIỆU (Data Cleansing & Binning)
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "KHÁM PHÁ & LÀM SẠCH DỮ LIỆU THỰC NGHIỆM (EDA & DATA CLEANSING)", "PHẦN 2: KỸ THUẬT TIỀN XỬ LÝ DỮ LIỆU")

    # Left Column: Data Cleansing Protocols
    c_left = add_card(s3, Inches(0.8), Inches(1.35), Inches(5.6), Inches(5.6), C_CARD_BG, C_BORDER_LIGHT)
    tb_cl = s3.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(5.2), Inches(5.2))
    tf_cl = tb_cl.text_frame
    tf_cl.word_wrap = True

    p = tf_cl.paragraphs[0]
    p.text = "🛠️ QUY TRÌNH KIỂM TRA & XỬ LÝ DỮ LIỆU"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_BLUE
    p.space_after = Pt(10)

    bullets = [
        ("Xoá Trùng Lặp Triệt Để (Duplicate Check):", "Bảng df_Products gốc có 89,316 dòng chứa tới 61,865 bản ghi trùng lặp hoàn toàn. Đã áp dụng Drop Duplicates thành công, rút gọn về đúng 27,451 sản phẩm duy nhất."),
        ("Xử Lý Giá Trị Khuyết Thiếu (Missing Values):", "• Cột Ship Date khuyết 1,889 dòng: Giữ nguyên Null để phản ánh chính xác đơn hàng hủy (404 đơn) hoặc chưa giao.\n• Danh mục khuyết 308 dòng: Gắn nhãn chuẩn 'Unspecified'.\n• Khối lượng khuyết: Điền giá trị Trung vị (Median)."),
        ("Chuẩn Hóa Kiểu Dữ Liệu (Data Type):", "Ép kiểu Date chuẩn (Order Date, Ship Date), Decimal số thực (Sales, Shipping Charge, Weight kg) và Text (ID keys)."),
        ("Ràng Buộc Logic (Sanity Constraints):", "Khống chế Delivery Lead Time >= 0 ngày, Doanh thu Sales > 0 và Cước phí Shipping >= 0.")
    ]
    for b_title, b_desc in bullets:
        p1 = tf_cl.add_paragraph()
        p1.text = "• " + b_title
        p1.font.size = Pt(10)
        p1.font.bold = True
        p1.font.color.rgb = C_NAVY_DARK
        p2 = tf_cl.add_paragraph()
        p2.text = b_desc
        p2.font.size = Pt(9)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_after = Pt(8)

    # Right Column: Data Binning Strategy
    c_right = add_card(s3, Inches(6.6), Inches(1.35), Inches(5.933), Inches(5.6), C_CARD_BG, C_BORDER_LIGHT)
    tb_cr = s3.shapes.add_textbox(Inches(6.8), Inches(1.5), Inches(5.5), Inches(5.2))
    tf_cr = tb_cr.text_frame
    tf_cr.word_wrap = True

    p = tf_cr.paragraphs[0]
    p.text = "📊 CHIẾN LƯỢC GOM NHÓM DỮ LIỆU (DATA BINNING)"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_BLUE
    p.space_after = Pt(10)

    bins = [
        ("Delivery Performance Tier (Thời gian giao hàng)", "• Same Day / Express: <= 0 ngày (0 đơn)\n• Standard Delivery: 1 - 3 ngày (4,378 đơn - 4.90%)\n• Delayed Delivery: > 3 ngày (83,049 đơn - 92.98%)\n• Not Shipped / Cancelled: 1,889 đơn (2.11%)"),
        ("Order Basket Value Tier (Quy mô giỏ hàng)", "• Low Value (< $50): 3,491 đơn (3.91%)\n• Medium Value ($50 - $200): 36,929 đơn (41.35%)\n• High Value (> $200): 48,896 đơn (54.74% - Đa số)"),
        ("Logistics Cost Tier (Tỷ lệ chi phí vận chuyển)", "• Optimal (< 8%): 33,269 đơn (37.25%)\n• Moderate (8% - 15%): 32,878 đơn (36.81%)\n• Profit Erosion (> 15%): 23,169 đơn (25.94% rủi ro)"),
        ("Weight Class (Khối lượng hàng hóa)", "• Hàng nhẹ (< 1kg): 52,860 đơn (59.18%)\n• Hàng vừa (1kg - 5kg): 25,541 đơn (28.60%)\n• Hàng nặng cồng kềnh (> 5kg): 10,915 đơn (12.22%)")
    ]
    for b_title, b_desc in bins:
        p1 = tf_cr.add_paragraph()
        p1.text = "📌 " + b_title
        p1.font.size = Pt(10)
        p1.font.bold = True
        p1.font.color.rgb = C_OCEAN_BLUE
        p2 = tf_cr.add_paragraph()
        p2.text = b_desc
        p2.font.size = Pt(9)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_after = Pt(6)

    # =========================================================================
    # SLIDE 4: MÔ HÌNH DỮ LIỆU STAR SCHEMA & BAREM MÃ DAX (Data Model & DAX)
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "KIẾN TRÚC MÔ HÌNH STAR SCHEMA & BAREM CÔNG THỨC DAX", "PHẦN 3: DATA MODELING & FORMULA ARCHITECTURE")

    # Left: Star Schema Architecture
    c_schema = add_card(s4, Inches(0.8), Inches(1.35), Inches(6.0), Inches(5.6), C_CARD_BG, C_BORDER_LIGHT)
    tb_sch = s4.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(5.6), Inches(5.2))
    tf_sch = tb_sch.text_frame
    tf_sch.word_wrap = True

    p = tf_sch.paragraphs[0]
    p.text = "⭐ MÔ HÌNH STAR SCHEMA CHUẨN POWER BI"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_BLUE
    p.space_after = Pt(10)

    p1 = tf_sch.add_paragraph()
    p1.text = "1 Bảng Fact Trung Tâm + 4 Bảng Chiều (Quan hệ 1-Nhiều 1:*)"
    p1.font.size = Pt(10)
    p1.font.bold = True
    p1.font.color.rgb = C_TEXT_MUTED
    p1.space_after = Pt(8)

    tables_info = [
        ("Fact_Orders (89,316 records - Bảng trung tâm):", "Lưu vết toàn bộ Order ID, Customer ID, Product ID, Sales, Shipping Charge, Order Date, Ship Date, Lead Time, Status & các trường Bins."),
        ("Dim_Customers (89,316 records):", "Khóa chính: customer_id. Lưu địa lý (State, City, Zipcode), Recency Days và Phân khúc khách hàng."),
        ("Dim_Products (27,451 records đã khử trùng):", "Khóa chính: product_id. Danh mục sản phẩm, khối lượng chuẩn hóa (kg) và phân lớp Weight Class."),
        ("Dim_Date (1,096 records - Bảng Lịch tự động):", "Khóa: date_key (YYYYMMDD). Năm, Quý, Tháng, Thứ, Ngày cuối tuần; hỗ trợ Time Intelligence (MoM, YoY, YTD)."),
        ("Dim_Payments (89,316 records):", "Phương thức thanh toán và số kỳ trả góp của từng đơn hàng.")
    ]
    for t_title, t_desc in tables_info:
        p_t = tf_sch.add_paragraph()
        p_t.text = "• " + t_title
        p_t.font.size = Pt(9.5)
        p_t.font.bold = True
        p_t.font.color.rgb = C_NAVY_DARK
        p_d = tf_sch.add_paragraph()
        p_d.text = t_desc
        p_d.font.size = Pt(8.8)
        p_d.font.color.rgb = C_TEXT_DARK
        p_d.space_after = Pt(6)

    # Right: DAX Formulas Library
    c_dax = add_card(s4, Inches(7.0), Inches(1.35), Inches(5.533), Inches(5.6), C_CARD_BG, C_BORDER_LIGHT)
    tb_dax = s4.shapes.add_textbox(Inches(7.2), Inches(1.5), Inches(5.1), Inches(5.2))
    tf_dax = tb_dax.text_frame
    tf_dax.word_wrap = True

    p = tf_dax.paragraphs[0]
    p.text = "💻 THƯ VIỆN 27 CÔNG THỨC DAX CHUẨN ĐỀ BÀI"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_BLUE
    p.space_after = Pt(10)

    dax_list = [
        ("1. Total Revenue:", "Total Revenue = SUM(Fact_Orders[sales])"),
        ("2. Average Order Value (AOV):", "AOV = DIVIDE([Total Revenue], [Total Orders], 0)"),
        ("3. On-Time Delivery Rate %:", "VAR Delivered = CALCULATE([Total Orders], NOT(ISBLANK(Fact_Orders[ship_date])))\nVAR OnTime = CALCULATE([Total Orders], Fact_Orders[delivery_tier] IN {'1. Express', '2. Standard'})\nRETURN DIVIDE(OnTime, Delivered, 0)"),
        ("4. Shipping Cost Ratio %:", "Shipping Cost Ratio % = DIVIDE(SUM(Fact_Orders[shipping_charge]), [Total Revenue], 0)"),
        ("5. Customer Segment (Chuẩn Đề Bài Section 6.1):", "SWITCH(TRUE(),\n  RecencyDays > 90, '4. Lost / At-Risk',\n  TotalSpend >= 1000 && TotalOrders >= 5, '1. VIP',\n  TotalOrders >= 3, '2. Loyal',\n  '3. Potential')"),
        ("6. Enhanced Monetary Segment (Mô hình Nâng cao):", "Bổ sung nhận diện khách hàng chi tiêu lớn > $1,000 cho tập dữ liệu đơn lẻ (Single Purchase).")
    ]
    for d_title, d_code in dax_list:
        p1 = tf_dax.add_paragraph()
        p1.text = d_title
        p1.font.size = Pt(9.2)
        p1.font.bold = True
        p1.font.color.rgb = C_EMERALD
        p2 = tf_dax.add_paragraph()
        p2.text = d_code
        p2.font.size = Pt(8.2)
        p2.font.color.rgb = RGBColor(51, 65, 85)
        p2.space_after = Pt(5)

    # =========================================================================
    # SLIDE 5: TRỤ CỘT 1 - PHÂN TÍCH DOANH THU & SẢN PHẨM (Sales & Products)
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "TRỤ CỘT 1: PHÂN TÍCH HIỆU QUẢ DOANH THU & CƠ CẤU SẢN PHẨM", "PHẦN 4: SALES PERFORMANCE & PRODUCT ANALYTICS")

    # Table Top 6 Categories
    c_tbl = add_card(s5, Inches(0.8), Inches(1.35), Inches(7.5), Inches(5.6), C_CARD_BG, C_BORDER_LIGHT)
    tb_t = s5.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(7.1), Inches(5.2))
    tf_t = tb_t.text_frame
    tf_t.word_wrap = True

    p = tf_t.paragraphs[0]
    p.text = "📦 CƠ CẤU DOANH THU & CHI PHÍ SHIPPING THEO NGÀNH HÀNG"
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_BLUE
    p.space_after = Pt(10)

    # Add a Table inside slide 5
    rows = 8
    cols = 6
    table_shape = s5.shapes.add_table(rows, cols, Inches(1.0), Inches(2.0), Inches(7.1), Inches(4.7))
    t = table_shape.table
    t.columns[0].width = Inches(1.8)
    t.columns[1].width = Inches(0.9)
    t.columns[2].width = Inches(1.2)
    t.columns[3].width = Inches(1.0)
    t.columns[4].width = Inches(1.0)
    t.columns[5].width = Inches(1.2)

    headers = ["Danh mục", "Số Đơn", "Doanh thu", "Tỷ trọng", "AOV", "Tỷ lệ Ship"]
    for c_i, h_text in enumerate(headers):
        cell = t.cell(0, c_i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_NAVY_DARK
        p = cell.text_frame.paragraphs[0]
        p.text = h_text
        p.font.size = Pt(9)
        p.font.bold = True
        p.font.color.rgb = C_WHITE

    cat_data = [
        ("toys", "67,027", "$22.99M", "75.52%", "$343.07", "12.99%"),
        ("furniture_decor", "1,760", "$753.07K", "2.47%", "$427.88", "10.49%"),
        ("garden_tools", "807", "$746.59K", "2.45%", "$925.14", "4.52%"),
        ("bed_bath_table", "2,146", "$604.53K", "1.99%", "$281.70", "15.74% ⚠️"),
        ("health_beauty", "2,351", "$596.72K", "1.96%", "$253.82", "16.89% ⚠️"),
        ("watches_gifts", "1,196", "$584.22K", "1.92%", "$488.48", "8.91%"),
        ("computers_acc", "1,715", "$430.53K", "1.41%", "$251.04", "17.12% 🚨")
    ]
    for r_i, rdata in enumerate(cat_data):
        for c_i, val in enumerate(rdata):
            cell = t.cell(r_i + 1, c_i)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_ICE_GRAY if r_i % 2 == 0 else C_WHITE
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.size = Pt(8.5)
            if "⚠️" in val or "🚨" in val:
                p.font.bold = True
                p.font.color.rgb = C_CORAL
            elif c_i == 0 and val == "toys":
                p.font.bold = True
                p.font.color.rgb = C_NAVY_BLUE

    # Right Box: Insights on Sales & Payment
    c_ins = add_card(s5, Inches(8.5), Inches(1.35), Inches(4.033), Inches(5.6), C_CARD_BG, C_BORDER_LIGHT)
    tb_ins = s5.shapes.add_textbox(Inches(8.7), Inches(1.5), Inches(3.6), Inches(5.2))
    tf_ins = tb_ins.text_frame
    tf_ins.word_wrap = True

    p = tf_ins.paragraphs[0]
    p.text = "💡 INSIGHT CHIẾN LƯỢC SẢN PHẨM"
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_BLUE
    p.space_after = Pt(10)

    insights = [
        ("Ngành Hàng Trụ Cột (Toys):", "Chiếm hơn 3/4 quy mô doanh thu ($22.99M). Đóng vai trò sống còn đối với dòng tiền của doanh nghiệp."),
        ("Ngành Hàng Siêu Lợi Nhuận (Garden Tools):", "AOV đạt kỷ lục $925.14, trong khi cước vận chuyển chỉ chiếm 4.52% (ngưỡng tối ưu tuyệt đối). Cần mở rộng quy mô Marketing."),
        ("Báo Động Đỏ Xói Mòn Lợi Nhuận:", "Phụ kiện máy tính (17.12%), Mỹ phẩm (16.89%), và Đồ gia dụng vải (15.74%) đang bị cước vận chuyển ăn mòn lợi nhuận nghiêm trọng."),
        ("Động Lực Thanh Toán (Payments):", "Thẻ tín dụng (Credit Card) chiếm 72.65% doanh thu ($22.12M) nhờ chính sách chia nhỏ kỳ trả góp trung bình 3.67 kỳ.")
    ]
    for i_title, i_desc in insights:
        p1 = tf_ins.add_paragraph()
        p1.text = "• " + i_title
        p1.font.size = Pt(9.5)
        p1.font.bold = True
        p1.font.color.rgb = C_NAVY_DARK
        p2 = tf_ins.add_paragraph()
        p2.text = i_desc
        p2.font.size = Pt(8.8)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_after = Pt(6)

    # =========================================================================
    # SLIDE 6: TRỤ CỘT 2 - HÀNH VI KHÁCH HÀNG & PHÂN KHÚC RFM (Customer Analytics)
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, "TRỤ CỘT 2: HÀNH VI KHÁCH HÀNG & PHÂN KHÚC RFM NÂNG CAO", "PHẦN 5: CUSTOMER BEHAVIOR & RETENTION")

    # Left: Standard DAX Model vs Enhanced Model
    c_comp = add_card(s6, Inches(0.8), Inches(1.35), Inches(5.7), Inches(5.6), C_CARD_BG, C_BORDER_LIGHT)
    tb_comp = s6.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(5.3), Inches(5.2))
    tf_comp = tb_comp.text_frame
    tf_comp.word_wrap = True

    p = tf_comp.paragraphs[0]
    p.text = "🎯 ĐỐI CHIẾU 2 MÔ HÌNH PHÂN KHÚC KHÁCH HÀNG"
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_BLUE
    p.space_after = Pt(10)

    p1 = tf_comp.add_paragraph()
    p1.text = "1. Mô hình Barem Chuẩn Đề Bài (Standard DAX):"
    p1.font.size = Pt(10)
    p1.font.bold = True
    p1.font.color.rgb = C_CORAL
    p2 = tf_comp.add_paragraph()
    p2.text = "• Lost / At-Risk: 73,317 khách (82.09%) do Recency > 90 ngày.\n• Potential Customer: 15,999 khách (17.91%) mới mua gần đây.\n• VIP & Loyal: Ghi nhận 0 khách hàng do cấu trúc bảng gốc phân rã ID theo từng đơn (Single-order tokenization, Total Orders = 1)."
    p2.font.size = Pt(8.8)
    p2.font.color.rgb = C_TEXT_DARK
    p2.space_after = Pt(10)

    p3 = tf_comp.add_paragraph()
    p3.text = "2. Mô hình Giá Trị Chi Tiêu Cải Tiến (Enhanced Monetary):"
    p3.font.size = Pt(10)
    p3.font.bold = True
    p3.font.color.rgb = C_EMERALD
    p4 = tf_comp.add_paragraph()
    p4.text = "• 1. VIP Customer (High Value): 5,281 khách (5.91%) chi tiêu > $1,000, đóng góp tới $8.12M doanh thu (26.68%).\n• 2. Loyal / High-Tier: 14,341 khách (16.06%) chi tiêu $300 - $1,000, đóng góp $7.45M (24.47%).\n• 3. Potential Customer: 11,337 khách (12.69%) giao dịch gần đây.\n• 4. Lost / At-Risk: 58,357 khách (65.34%) đơn giá thấp và ngừng mua."
    p4.font.size = Pt(8.8)
    p4.font.color.rgb = C_TEXT_DARK

    # Right: Customer Action Cards
    c_act = add_card(s6, Inches(6.7), Inches(1.35), Inches(5.833), Inches(5.6), C_CARD_BG, C_BORDER_LIGHT)
    tb_act = s6.shapes.add_textbox(Inches(6.9), Inches(1.5), Inches(5.4), Inches(5.2))
    tf_act = tb_act.text_frame
    tf_act.word_wrap = True

    p = tf_act.paragraphs[0]
    p.text = "💼 CHIẾN LƯỢC CHĂM SÓC THEO PHÂN LỚP (ACTION MATRIX)"
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_BLUE
    p.space_after = Pt(10)

    actions_rfm = [
        ("Nhóm VIP High-Value (5,281 khách hàng - $8.12M Doanh thu):", "Tài sản cốt lõi của doanh nghiệp. Thiết lập dịch vụ chăm sóc khách hàng ưu tiên (Dedicated Concierge), tặng quà sinh nhật tri ân và miễn phí đổi trả tận nhà."),
        ("Nhóm Loyal High-Tier (14,341 khách hàng - $7.45M Doanh thu):", "Xây dựng chương trình khách hàng thân thiết tích điểm (Loyalty Points), gửi mã ưu đãi độc quyền để nâng hạng chi tiêu lên mức VIP."),
        ("Nhóm Potential (11,337 khách hàng mới):", "Gửi chuỗi Email Onboarding trong 14 ngày đầu, gợi ý combo sản phẩm cùng loại kèm voucher giảm 10% cho đơn hàng thứ 2."),
        ("Nhóm Lost / At-Risk (58,357 khách hàng):", "Triển khai chiến dịch Win-Back Campaign định kỳ tự động bằng SMS & Email Marketing với mã Freeship giới hạn 7 ngày.")
    ]
    for a_title, a_desc in actions_rfm:
        p1 = tf_act.add_paragraph()
        p1.text = "👑 " + a_title
        p1.font.size = Pt(9.5)
        p1.font.bold = True
        p1.font.color.rgb = C_OCEAN_BLUE
        p2 = tf_act.add_paragraph()
        p2.text = a_desc
        p2.font.size = Pt(8.8)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_after = Pt(6)

    # =========================================================================
    # SLIDE 7: TRỤ CỘT 3 - CHUỖI CUNG ỨNG & VẬN HÀNH LOGISTICS (Supply Chain)
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, "TRỤ CỘT 3: VẬN HÀNH CHUỖI CUNG ỨNG & ĐIỂM NGHẼN LOGISTICS", "PHẦN 6: LAST-MILE LOGISTICS OPERATIONS")

    # Table Top States Lead Time
    c_geo = add_card(s7, Inches(0.8), Inches(1.35), Inches(7.5), Inches(5.6), C_CARD_BG, C_BORDER_LIGHT)
    tb_g = s7.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(7.1), Inches(5.2))
    tf_g = tb_g.text_frame
    tf_g.word_wrap = True

    p = tf_g.paragraphs[0]
    p.text = "🚚 THỜI GIAN GIAO HÀNG & ĐIỂM NGHẼN VẬN TẢI THEO BANG"
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_BLUE
    p.space_after = Pt(10)

    # Add Table
    table_shape7 = s7.shapes.add_table(8, 6, Inches(1.0), Inches(2.0), Inches(7.1), Inches(4.7))
    t7 = table_shape7.table
    t7.columns[0].width = Inches(1.8)
    t7.columns[1].width = Inches(0.9)
    t7.columns[2].width = Inches(1.2)
    t7.columns[3].width = Inches(1.0)
    t7.columns[4].width = Inches(1.1)
    t7.columns[5].width = Inches(1.1)

    geo_headers = ["Bang (State)", "Số Đơn", "Doanh thu", "Tỷ trọng", "Lead Time", "Đúng hạn"]
    for c_i, h_text in enumerate(geo_headers):
        cell = t7.cell(0, c_i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_NAVY_DARK
        p = cell.text_frame.paragraphs[0]
        p.text = h_text
        p.font.size = Pt(9)
        p.font.bold = True
        p.font.color.rgb = C_WHITE

    geo_data = [
        ("SP (São Paulo)", "37,879", "$12.94M", "42.49%", "8.70 ngày", "10.33% ⭐"),
        ("RJ (Rio de Janeiro)", "11,578", "$3.85M", "12.65%", "15.21 ngày ⚠️", "1.80%"),
        ("MG (Minas Gerais)", "10,334", "$3.50M", "11.50%", "11.97 ngày", "1.45%"),
        ("RS (Rio Grande Sul)", "4,921", "$1.66M", "5.46%", "15.04 ngày ⚠️", "0.62%"),
        ("PR (Paraná)", "4,523", "$1.61M", "5.28%", "12.06 ngày", "1.91%"),
        ("SC (Santa Catarina)", "3,213", "$1.07M", "3.50%", "14.97 ngày", "1.01%"),
        ("BA (Bahia)", "3,088", "$1.06M", "3.48%", "19.33 ngày 🚨", "0.30% 🚨")
    ]
    for r_i, rdata in enumerate(geo_data):
        for c_i, val in enumerate(rdata):
            cell = t7.cell(r_i + 1, c_i)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_ICE_GRAY if r_i % 2 == 0 else C_WHITE
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.size = Pt(8.5)
            if "🚨" in val or "⚠️" in val:
                p.font.bold = True
                p.font.color.rgb = C_CORAL
            elif "⭐" in val:
                p.font.bold = True
                p.font.color.rgb = C_EMERALD

    # Right: Logistics Bottleneck Analysis
    c_log = add_card(s7, Inches(8.5), Inches(1.35), Inches(4.033), Inches(5.6), C_CARD_BG, C_BORDER_LIGHT)
    tb_log = s7.shapes.add_textbox(Inches(8.7), Inches(1.5), Inches(3.6), Inches(5.2))
    tf_log = tb_log.text_frame
    tf_log.word_wrap = True

    p = tf_log.paragraphs[0]
    p.text = "⚠️ NGUYÊN NHÂN TẮC NGHẼN CHẶNG CUỐI"
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_BLUE
    p.space_after = Pt(10)

    log_insights = [
        ("Tỷ Lệ Đúng Hạn Thấp (5.01%):", "Tiêu chuẩn đề bài quy định giao nhanh <= 3 ngày chỉ đạt 5.01%. Trên thực tế, thị trường Brazil rộng lớn đòi hỏi Lead Time chuẩn từ 7 - 14 ngày."),
        ("Điểm Đen Tắc Nghẽn Bahia (BA):", "Thời gian giao hàng trung bình lên tới 19.33 ngày (gần 3 tuần), tỷ lệ đúng hạn chỉ 0.30%. Đây là nguồn phát sinh khiếu nại và hủy đơn lớn nhất."),
        ("Thách Thức Rio de Janeiro (RJ):", "Thị trường lớn thứ 2 ($3.85M) nhưng thời gian giao lên đến 15.21 ngày do ùn tắc đô thị và khoảng cách kho vận trung tâm tại São Paulo."),
        ("Tác Động Khối Lượng Hàng Hóa:", "Hàng nặng > 5kg chịu thời gian giao 14.19 ngày và cước phí $46.61/đơn, cao hơn 18% so với hàng nhẹ.")
    ]
    for l_title, l_desc in log_insights:
        p1 = tf_log.add_paragraph()
        p1.text = "• " + l_title
        p1.font.size = Pt(9.5)
        p1.font.bold = True
        p1.font.color.rgb = C_CORAL if "Điểm Đen" in l_title else C_NAVY_DARK
        p2 = tf_log.add_paragraph()
        p2.text = l_desc
        p2.font.size = Pt(8.8)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_after = Pt(6)

    # =========================================================================
    # SLIDE 8: DỰ BÁO NHU CẦU & BẢNG TÁC VỤ KHẨN CẤP (Forecasting & Action Center)
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, "DỰ BÁO NHU CẦU Q4 & BẢNG TÁC VỤ ĐIỀU HÀNH KHẨN CẤP", "PHẦN 7: AI FORECASTING & ACTION CENTER")

    # Left: AI Forecasting Q4/2018
    c_fc = add_card(s8, Inches(0.8), Inches(1.35), Inches(5.7), Inches(5.6), C_CARD_BG, C_BORDER_LIGHT)
    tb_fc = s8.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(5.3), Inches(5.2))
    tf_fc = tb_fc.text_frame
    tf_fc.word_wrap = True

    p = tf_fc.paragraphs[0]
    p.text = "📈 DỰ BÁO NHU CẦU 3 THÁNG TIẾP THEO (Q4/2018)"
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_BLUE
    p.space_after = Pt(10)

    p_sub = tf_fc.add_paragraph()
    p_sub.text = "Thuật toán Time-Series (Exponential Smoothing với Độ tin cậy 95%):"
    p_sub.font.size = Pt(9.5)
    p_sub.font.bold = True
    p_sub.font.color.rgb = C_TEXT_MUTED
    p_sub.space_after = Pt(8)

    fc_months = [
        ("Tháng 10/2018 (Ổn định):", "Doanh thu dự báo $1,620,000 (±8%) | Số lượng đơn ước tính: 4,600 đơn."),
        ("Tháng 11/2018 (Đỉnh điểm Black Friday):", "Doanh thu vọt lên $2,450,000 (±12%) | Số lượng đơn ước tính: 7,200 đơn (+51% so với bình quân). Đòi hỏi chuẩn bị năng lực kho vận cực đại."),
        ("Tháng 12/2018 (Mùa Lễ hội Giáng sinh):", "Doanh thu duy trì ở mức cao $2,100,000 (±10%) | Số lượng đơn ước tính: 6,100 đơn."),
        ("Tổng Doanh Thu Q4/2018 Kỳ Vọng:", "$6.17 Triệu USD với hơn 17,900 đơn hàng cần hoàn tất.")
    ]
    for fm_title, fm_desc in fc_months:
        p1 = tf_fc.add_paragraph()
        p1.text = "• " + fm_title
        p1.font.size = Pt(9.5)
        p1.font.bold = True
        p1.font.color.rgb = C_EMERALD if "Black Friday" in fm_title else C_NAVY_DARK
        p2 = tf_fc.add_paragraph()
        p2.text = fm_desc
        p2.font.size = Pt(8.8)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_after = Pt(6)

    # Right: Executive Action Center
    c_em = add_card(s8, Inches(6.7), Inches(1.35), Inches(5.833), Inches(5.6), C_CARD_BG, C_BORDER_LIGHT)
    tb_em = s8.shapes.add_textbox(Inches(6.9), Inches(1.5), Inches(5.4), Inches(5.2))
    tf_em = tb_em.text_frame
    tf_em.word_wrap = True

    p = tf_em.paragraphs[0]
    p.text = "🚨 DANH MỤC TÁC VỤ KHẨN CẤP CẦN XỬ LÝ (ACTION MATRIX)"
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_BLUE
    p.space_after = Pt(10)

    action_items = [
        ("1. Bồi Hoàn & Xử Lý 404 Đơn Hàng Hủy (Canceled):", "Phòng CSKH liên hệ trực tiếp trong 24h, hoàn tất thủ tục hoàn tiền để ngăn chặn đánh giá tiêu cực và bảo vệ danh tiếng thương hiệu."),
        ("2. Can Thiệp Các Đơn Giao Trễ Kéo Dài > 20 Ngày:", "Kích hoạt cảnh báo tới đối tác 3PLs chuyên tuyến tại Rio de Janeiro và Bahia; gửi tin nhắn xin lỗi kèm mã giảm giá 10% tri ân cho khách hàng."),
        ("3. Tái Cơ Cấu Cước Phí Cho 3 Ngành Hàng Rò Rỉ:", "Thiết lập trần trợ giá vận chuyển hoặc điều kiện mua hàng tối thiểu $200 đối với `computers_accessories` (17.1%) và `health_beauty` (16.9%)."),
        ("4. Kích Hoạt Chiến Dịch Win-Back Cho Khách Hàng VIP:", "Phòng Marketing triển khai kịch bản Re-engagement riêng biệt cho 5,281 khách hàng chi tiêu lớn trước thềm Black Friday.")
    ]
    for ai_title, ai_desc in action_items:
        p1 = tf_em.add_paragraph()
        p1.text = "⚡ " + ai_title
        p1.font.size = Pt(9.5)
        p1.font.bold = True
        p1.font.color.rgb = C_CORAL if "404 Đơn" in ai_title else C_OCEAN_BLUE
        p2 = tf_em.add_paragraph()
        p2.text = ai_desc
        p2.font.size = Pt(8.8)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_after = Pt(6)

    # =========================================================================
    # SLIDE 9: ĐỀ XUẤT CHIẾN LƯỢC & LỘ TRÌNH 30-60-90 NGÀY (Strategic Roadmap)
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_header(s9, "ĐỀ XUẤT CHIẾN LƯỢC & KẾ HOẠCH HÀNH ĐỘNG 30 - 60 - 90 NGÀY", "PHẦN 8: STRATEGIC RECOMMENDATIONS & ROADMAP")

    stages = [
        ("GIAI ĐOẠN 30 NGÀY (CẤP BÁCH & QUICK-WINS)", C_CORAL, [
            "Công bố SLA Giao hàng thực tế trên Website/App: Minh bạch thời gian giao (SP: 7-9 ngày, RJ: 12-14 ngày, BA: 18-20 ngày) để triệt tiêu kỳ vọng sai lệch gây hủy đơn.",
            "Xử lý dứt điểm 404 đơn hủy: Hoàn tất bồi hoàn 100% trong vòng 15 ngày làm việc đầu tiên.",
            "Chiến dịch Win-Back cho Khách hàng VIP: Gửi ưu đãi cá nhân hóa tới 5,281 VIP (ngân sách dự kiến kích hoạt lại: $8.12M)."
        ]),
        ("GIAI ĐOẠN 60 NGÀY (TỐI ƯU HÓA VẬN HÀNH & CHI PHÍ)", C_OCEAN_BLUE, [
            "Áp dụng ngưỡng Freeship $200: Ngăn chặn rò rỉ $1.2M chi phí logistics hàng năm từ các ngành phụ kiện máy tính và mỹ phẩm.",
            "Tái đàm phán hợp đồng 3PLs: Đấu thầu lại đối tác vận chuyển chuyên tuyến riêng biệt cho Rio de Janeiro và Bahia.",
            "Xây dựng chương trình Khách hàng Thân thiết (Loyalty): Khuyến khích tích điểm đổi quà cho 14,341 khách hàng nhóm Loyal."
        ]),
        ("GIAI ĐOẠN 90 NGÀY (TÁI CẤU TRÚC CHUỖI CUNG ỨNG)", C_EMERALD, [
            "Thành lập Satellite Hub tại Rio de Janeiro: Mở trung tâm kho vệ tinh phục vụ trực tiếp bang RJ và ES, cắt giảm Lead Time từ 15.2 ngày xuống dưới 6 ngày.",
            "Chuẩn hóa hạ tầng Power BI Real-Time: Đưa Dashboard 4 trang vào vận hành điều hành hàng tuần của Ban Giám đốc và khối Logistics.",
            "Mở rộng danh mục sinh lời cao: Đẩy mạnh các dòng sản phẩm có AOV lớn và cước thấp như Garden Tools và Home Decor."
        ])
    ]

    for idx, (stage_title, stage_col, stage_items) in enumerate(stages):
        left = Inches(0.8 + idx * 4.05)
        card = add_card(s9, left, Inches(1.35), Inches(3.8), Inches(5.6), C_CARD_BG, C_BORDER_LIGHT)
        # Top banner of stage card
        top_bar = s9.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, Inches(1.35), Inches(3.8), Inches(0.55))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = stage_col
        top_bar.line.fill.background()
        tf_tb = top_bar.text_frame
        p_tb = tf_tb.paragraphs[0]
        p_tb.text = stage_title
        p_tb.font.size = Pt(9.5)
        p_tb.font.bold = True
        p_tb.font.color.rgb = C_WHITE
        p_tb.alignment = PP_ALIGN.CENTER

        tb_s = s9.shapes.add_textbox(left + Inches(0.2), Inches(2.05), Inches(3.4), Inches(4.7))
        tf_s = tb_s.text_frame
        tf_s.word_wrap = True

        for item_idx, itm in enumerate(stage_items):
            p = tf_s.add_paragraph() if item_idx > 0 else tf_s.paragraphs[0]
            p.text = f"{item_idx + 1}. {itm}"
            p.font.size = Pt(9.2)
            p.font.color.rgb = C_TEXT_DARK
            p.space_after = Pt(12)

    # =========================================================================
    # SLIDE 10: TỔNG KẾT & TRẢI NGHIỆM DASHBOARD TRỰC TUYẾN (Conclusion & Demo)
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_header(s10, "TỔNG KẾT KHÓA LUẬN & TRẢI NGHIỆM DASHBOARD TRỰC TIẾP", "KẾT LUẬN & HỎI ĐÁP (Q&A)")

    # Left Box: Accomplishments
    c_acc = add_card(s10, Inches(0.8), Inches(1.35), Inches(5.7), Inches(5.6), C_CARD_BG, C_BORDER_LIGHT)
    tb_ac = s10.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(5.3), Inches(5.2))
    tf_ac = tb_ac.text_frame
    tf_ac.word_wrap = True

    p = tf_ac.paragraphs[0]
    p.text = "🏆 THÀNH QUẢ DỰ ÁN TỐT NGHIỆP ĐẠT ĐƯỢC"
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_BLUE
    p.space_after = Pt(10)

    results = [
        ("Bộ Dữ Liệu Sạch Chuẩn Star Schema:", "Khử trùng 61,865 dòng trong `df_Products`, hợp nhất 89,316 giao dịch thành `Fact_Orders` và 4 bảng `Dim` hoàn chỉnh."),
        ("Thư Viện 27 Công Thức DAX Đầy Đủ:", "Tuân thủ chuẩn 100% barem đề bài (Standard DAX) kết hợp mở rộng mô hình phân khúc chi tiêu cao (Enhanced Monetary Model)."),
        ("Báo Cáo Tốt Nghiệp Chuẩn Mực (Word & MD):", "Hoàn thiện tài liệu 9 phần chuyên sâu, bao quát 3 trụ cột chiến lược và kế hoạch hành động chi tiết."),
        ("Hệ Thống Dashboard Tương Tác 4 Trang:", "Mô phỏng đầy đủ chuẩn UI/UX Power BI, tích hợp các bộ lọc Slicers, AI Forecasting và Bảng tác vụ khẩn cấp xuất CSV.")
    ]
    for r_title, r_desc in results:
        p1 = tf_ac.add_paragraph()
        p1.text = "✔ " + r_title
        p1.font.size = Pt(9.5)
        p1.font.bold = True
        p1.font.color.rgb = C_EMERALD
        p2 = tf_ac.add_paragraph()
        p2.text = r_desc
        p2.font.size = Pt(8.8)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_after = Pt(6)

    # Right Box: Live Demo Links & Thank You
    c_demo = add_card(s10, Inches(6.7), Inches(1.35), Inches(5.833), Inches(5.6), C_NAVY_DARK, None)
    tb_dm = s10.shapes.add_textbox(Inches(7.0), Inches(1.5), Inches(5.3), Inches(5.2))
    tf_dm = tb_dm.text_frame
    tf_dm.word_wrap = True

    p = tf_dm.paragraphs[0]
    p.text = "🌐 TRẢI NGHIỆM TRỰC TIẾP HỆ THỐNG"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_AMBER
    p.space_after = Pt(12)

    p1 = tf_dm.add_paragraph()
    p1.text = "🔗 Live Dashboard (Đang chạy trực tuyến trên GitHub Pages):"
    p1.font.size = Pt(10)
    p1.font.bold = True
    p1.font.color.rgb = C_WHITE
    p2 = tf_dm.add_paragraph()
    p2.text = "https://daothienvy.github.io/final-test-Data/"
    p2.font.size = Pt(10)
    p2.font.bold = True
    p2.font.color.rgb = C_EMERALD
    p2.space_after = Pt(12)

    p3 = tf_dm.add_paragraph()
    p3.text = "📂 Mã Nguồn & Báo Cáo Trên GitHub Repository:"
    p3.font.size = Pt(10)
    p3.font.bold = True
    p3.font.color.rgb = C_WHITE
    p4 = tf_dm.add_paragraph()
    p4.text = "https://github.com/daothienvy/final-test-Data"
    p4.font.size = Pt(10)
    p4.font.bold = True
    p4.font.color.rgb = C_OCEAN_BLUE
    p4.space_after = Pt(20)

    # Big Thank you banner
    p_ty = tf_dm.add_paragraph()
    p_ty.text = "XIN CHÂN THÀNH CẢM ƠN HỘI ĐỒNG GIÁM KHẢO!"
    p_ty.font.size = Pt(13)
    p_ty.font.bold = True
    p_ty.font.color.rgb = C_WHITE
    p_ty.alignment = PP_ALIGN.CENTER
    p_ty.space_after = Pt(6)

    p_qa = tf_dm.add_paragraph()
    p_qa.text = "— PHIÊN BÁO CÁO & HỎI ĐÁP (Q&A) —"
    p_qa.font.size = Pt(10.5)
    p_qa.font.bold = True
    p_qa.font.color.rgb = C_AMBER
    p_qa.alignment = PP_ALIGN.CENTER

    # Save presentation
    output_path = '/Users/ThienTuCorp/Downloads/Final test Data/Bao_Cao_Tot_Nghiep_Cuoi_Khoa.pptx'
    prs.save(output_path)
    print(f"Presentation successfully created at: {output_path}")

if __name__ == '__main__':
    create_deck()
