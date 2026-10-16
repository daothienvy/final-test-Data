# 📊 Phân Tích Kinh Doanh E-Commerce & Chuỗi Cung Ứng (Power BI Capstone Project)

Dự án Khóa luận / Báo cáo Tốt nghiệp Phân tích Dữ liệu Kinh doanh Thương mại Điện tử và Vận hành Chuỗi Cung ứng trên tập dữ liệu thực tế 89,316 đơn hàng.

---

## 🌟 Trực quan hóa Dashboard Trực tuyến (Live Dashboard)
Dashboard 4 trang mô phỏng chuẩn xác giao diện Power BI Report:
* **Trang 1: Executive Sales & Revenue Overview** (Doanh thu $30.45M, AOV $340.90, Xu hướng MoM/YoY).
* **Trang 2: Customer Behavior & Loyalty Segmentation** (Phân khúc RFM và Khách hàng Giá trị cao VIP).
* **Trang 3: Supply Chain & Last-Mile Logistics Operations** (Lead Time 12.43 ngày, Tỷ lệ giao đúng hạn, Tỷ lệ chi phí Shipping 12.99%).
* **Trang 4: Demand Forecasting & Executive Action Center** (Dự báo AI Q4/2018, Cây phân rã Decomposition Tree và Bảng tác vụ khẩn cấp xuất CSV).

---

## 📁 Cấu trúc Thư mục Dự án

```
.
├── index.html                   # Giao diện chính của Multi-Page Power BI Dashboard
├── app.css                      # Hệ thống Design & Bảng màu chuẩn Navy Blue & Light Ice Gray
├── app.js                       # Logic điều khiển, tương tác biểu đồ Chart.js & bộ lọc Slicers
├── dashboard_data.js            # Dữ liệu tổng hợp phục vụ hiển thị Dashboard
├── Bao_Cao_Tot_Nghiep_Cuoi_Khoa.docx # Báo cáo tốt nghiệp hoàn chỉnh định dạng Word
├── Bao_Cao_Tot_Nghiep_Cuoi_Khoa.md   # Báo cáo tốt nghiệp định dạng Markdown
├── dax_measures_library.dax     # Thư viện 27 công thức DAX đầy đủ chú thích
├── power_query_scripts.m        # Toàn bộ mã nguồn ngôn ngữ M cho Power Query Advanced Editor
├── power_bi_modeling_guide.md   # Hướng dẫn thiết lập Mô hình Star Schema và Visuals
├── etl_process.py               # Script Python ETL làm sạch dữ liệu và tạo Star Schema
├── generate_capstone_report.py  # Script tự động tạo Báo cáo Word & Markdown
├── cleaned_data/                # Thư mục dữ liệu sạch chuẩn hóa (Fact & Dimensions)
│   ├── Fact_Orders.csv
│   ├── Dim_Customers.csv
│   ├── Dim_Products.csv
│   ├── Dim_Payments.csv
│   ├── Dim_Date.csv
│   └── kpi_metrics.json
└── train/                       # Dữ liệu thô gốc
```

---

## 🚀 Hướng dẫn Triển khai Trên Power BI Desktop

1. **Nạp dữ liệu**: Mở Power BI Desktop, sử dụng các file trong thư mục `cleaned_data/`.
2. **Power Query**: Dán mã nguồn từ `power_query_scripts.m` vào Advanced Editor.
3. **Mô hình Star Schema**: Nối khóa `customer_id`, `product_id`, `order_date_key`, `order_id` giữa `Fact_Orders` và các bảng `Dim` tương ứng (Quan hệ 1-nhiều).
4. **Mã DAX**: Tạo bảng `_Measures` và copy công thức từ `dax_measures_library.dax`.
