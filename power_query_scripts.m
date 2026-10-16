/*
========================================================================================
POWER QUERY (M LANGUAGE) AUTOMATION SCRIPTS
Dự án: Báo cáo Tốt nghiệp Phân tích E-Commerce & Chuỗi cung ứng
Hướng dẫn sử dụng trong Power BI Desktop:
1. Vào Home -> Transform Data (Power Query Editor).
2. Tạo nguồn dữ liệu (Source) hoặc New Source -> Blank Query -> Advanced Editor.
3. Dán đoạn mã M tương ứng cho từng bảng vào Advanced Editor và nhấn Done.
========================================================================================
*/

// --------------------------------------------------------------------------------------
// 1. DIM_DATE (TỰ ĐỘNG TẠO BẢNG LỊCH CALENDAR CHO MÔ HÌNH STAR SCHEMA)
// --------------------------------------------------------------------------------------
let
    // Thiết lập dải ngày bao phủ toàn bộ dữ liệu giao dịch
    StartDate = #date(2016, 1, 1),
    EndDate = #date(2018, 12, 31),
    DayCount = Duration.Days(Duration.From(EndDate - StartDate)) + 1,
    Source = List.Dates(StartDate, DayCount, #duration(1, 0, 0, 0)),
    #"Converted to Table" = Table.FromList(Source, Splitter.SplitByNothing(), {"date"}, null, ExtraValues.Error),
    #"Changed Type" = Table.TransformColumnTypes(#"Converted to Table",{{"date", type date}}),
    #"Added DateKey" = Table.AddColumn(#"Changed Type", "date_key", each Date.ToText([date], "yyyyMMdd"), type text),
    #"Added FullDate" = Table.AddColumn(#"Added DateKey", "full_date", each Date.ToText([date], "yyyy-MM-dd"), type text),
    #"Added Year" = Table.AddColumn(#"Added FullDate", "year", each Date.Year([date]), Int64.Type),
    #"Added Quarter" = Table.AddColumn(#"Added Year", "quarter", each Date.QuarterOfYear([date]), Int64.Type),
    #"Added Quarter Name" = Table.AddColumn(#"Added Quarter", "quarter_name", each "Q" & Text.From([quarter]), type text),
    #"Added Month" = Table.AddColumn(#"Added Quarter Name", "month", each Date.Month([date]), Int64.Type),
    #"Added Month Name" = Table.AddColumn(#"Added Month", "month_name", each Date.ToText([date], "MMM"), type text),
    #"Added Month Year" = Table.AddColumn(#"Added Month Name", "month_year", each Date.ToText([date], "yyyy-MM"), type text),
    #"Added Day" = Table.AddColumn(#"Added Month Year", "day", each Date.Day([date]), Int64.Type),
    #"Added DayOfWeek" = Table.AddColumn(#"Added Day", "day_of_week", each Date.DayOfWeek([date], Day.Monday) + 1, Int64.Type),
    #"Added Day Name" = Table.AddColumn(#"Added DayOfWeek", "day_name", each Date.ToText([date], "dddd"), type text),
    #"Added Weekend Flag" = Table.AddColumn(#"Added Day Name", "is_weekend", each if [day_of_week] = 6 or [day_of_week] = 7 then 1 else 0, Int64.Type)
in
    #"Added Weekend Flag"


// --------------------------------------------------------------------------------------
// 2. DIM_PRODUCTS (LÀM SẠCH VÀ LOẠI BỎ TRÙNG LẶP EXACT DUPLICATES)
// --------------------------------------------------------------------------------------
let
    // Trỏ đường dẫn đến file Dim_Products.csv trong thư mục cleaned_data
    Source = Csv.Document(File.Contents("D:\Final test Data\cleaned_data\Dim_Products.csv"), [Delimiter=",", Columns=8, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{
        {"product_id", type text},
        {"product_category_name", type text},
        {"product_weight_g", type number},
        {"product_length_cm", type number},
        {"product_height_cm", type number},
        {"product_width_cm", type number},
        {"product_weight_kg", type number},
        {"weight_class", type text}
    }),
    // Xoá triệt để trùng lặp bản ghi theo đúng chỉ đạo kỹ thuật của Đề tài
    #"Removed Duplicates" = Table.Distinct(#"Changed Type", {"product_id"})
in
    #"Removed Duplicates"


// --------------------------------------------------------------------------------------
// 3. DIM_CUSTOMERS (CHUẨN HOÁ DANH MỤC KHÁCH HÀNG & PHÂN VÙNG ĐỊA LÝ)
// --------------------------------------------------------------------------------------
let
    Source = Csv.Document(File.Contents("D:\Final test Data\cleaned_data\Dim_Customers.csv"), [Delimiter=",", Columns=9, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{
        {"customer_id", type text},
        {"customer_zip_code_prefix", Int64.Type},
        {"customer_city", type text},
        {"customer_state", type text},
        {"customer_recency_days", Int64.Type},
        {"customer_recency_status", type text},
        {"total_orders", Int64.Type},
        {"total_spend", type number},
        {"customer_segment", type text}
    }),
    #"Capitalized City" = Table.TransformColumns(#"Changed Type", {{"customer_city", Text.Proper, type text}})
in
    #"Capitalized City"


// --------------------------------------------------------------------------------------
// 4. DIM_PAYMENTS (CHUẨN HOÁ PHƯƠNG THỨC THANH TOÁN VÀ SỐ KỲ TRẢ GÓP)
// --------------------------------------------------------------------------------------
let
    Source = Csv.Document(File.Contents("D:\Final test Data\cleaned_data\Dim_Payments.csv"), [Delimiter=",", Columns=5, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{
        {"order_id", type text},
        {"payment_sequential", Int64.Type},
        {"payment_type", type text},
        {"payment_installments", Int64.Type},
        {"payment_value", type number}
    })
in
    #"Changed Type"


// --------------------------------------------------------------------------------------
// 5. FACT_ORDERS (BẢNG TRUNG TÂM FACT CHỨA TOÀN BỘ SỐ LIỆU ĐƠN HÀNG, DOANH THU & CHUỖI CUNG ỨNG)
// --------------------------------------------------------------------------------------
let
    Source = Csv.Document(File.Contents("D:\Final test Data\cleaned_data\Fact_Orders.csv"), [Delimiter=",", Columns=27, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{
        {"order_id", type text},
        {"customer_id", type text},
        {"product_id", type text},
        {"seller_id", type text},
        {"order_date_key", type text},
        {"order_date", type date},
        {"ship_date", type date},
        {"estimated_delivery_date", type date},
        {"order_status", type text},
        {"sales", type number},
        {"shipping_charge", type number},
        {"shipping_cost_ratio", type number},
        {"delivery_lead_time_days", type number},
        {"delivery_performance_tier", type text},
        {"order_basket_value_tier", type text},
        {"logistics_cost_tier", type text},
        {"customer_recency_days", Int64.Type},
        {"customer_recency_status", type text},
        {"customer_segment", type text},
        {"customer_segment_enhanced", type text},
        {"payment_type", type text},
        {"payment_installments", Int64.Type},
        {"payment_value", type number},
        {"customer_city", type text},
        {"customer_state", type text},
        {"product_category_name", type text},
        {"weight_class", type text}
    })
in
    #"Changed Type"
