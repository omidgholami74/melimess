import pandas as pd

def excel_row_to_dict(file_path, row_index=1, sheet_name=0):
    """
    یک سطر از اکسل رو می‌خونه و فقط ستون‌هایی که مقدار دارند رو به دیکشنری تبدیل می‌کنه.
    
    پارامترها:
    - file_path: مسیر فایل اکسل
    - row_index: شماره سطر داده‌ها (پیش‌فرض: 1 = سطر دوم)
    - sheet_name: نام یا ایندکس شیت (پیش‌فرض: شیت اول)
    
    خروجی: دیکشنری فقط با مقادیر غیرخالی
    """
    # خواندن اکسل بدون هدر
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    # گرفتن سطر هدر (نام عناصر) و سطر داده
    header_row = df.iloc[0]        # سطر اول: نام عناصر
    data_row = df.iloc[row_index]  # سطر داده (مثلاً OREAS 903)
    
    # ساخت دیکشنری فقط با مقادیر غیرخالی
    data_dict = {}
    for col_idx in data_row.index:
        value = data_row[col_idx]
        key = header_row[col_idx]
        
        # فقط اگر مقدار غیرخالی باشد
        if pd.notna(value) and value != "":
            try:
                data_dict[key] = float(value)
            except (ValueError, TypeError):
                data_dict[key] = value
    
    return data_dict


# استفاده:
oreas_903 = excel_row_to_dict('Book1.xlsx', row_index=1)
print(oreas_903)