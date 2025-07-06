import pandas as pd
import matplotlib
matplotlib.use('TkAgg')  # Ép dùng backend TkAgg để hiển thị đồ thị trên Windows
import matplotlib.pyplot as plt
from matplotlib.dates import DateFormatter  # Import DateFormatter

# Đường dẫn tới file Excel
file_path = "D:/TMPs/To0606-1006CB01.xlsx"

# Đọc file Excel
df = pd.read_excel(file_path, header=0)

# Loại bỏ khoảng trắng thừa trong tên cột
df.columns = df.columns.str.strip()

# In danh sách cột để kiểm tra
print("Các cột trong file:", df.columns.tolist())

# Kiểm tra dữ liệu
print("Kiểm tra dữ liệu:")
print(df.head())  # In 5 dòng đầu tiên
print("Giá trị thiếu (NaN):")
print(df.isnull().sum())  # Kiểm tra số lượng giá trị thiếu

# Chuyển đổi cột Timestamp sang định dạng datetime
df['Timestamp'] = pd.to_datetime(df['Timestamp'], format='%m/%d/%Y %H:%M')

# Điền giá trị cho Serial Number và Logger Name
df['Serial Number'] = df['Serial Number'].fillna(df['Serial Number'].iloc[0])
df['Logger Name'] = df['Logger Name'].fillna(df['Logger Name'].iloc[0])

# Lấy số Serial từ dữ liệu
serial_number = df['Serial Number'].iloc[0]

# Dữ liệu để vẽ
timestamps = df['Timestamp']
temp = df['Temp.(°C)']
temp_ll = df['Temp. LL(°C)']
temp_hl = df['Temp. HL(°C)']
rh = df['RH(%rh)']
rh_hl = df['RH HL(%rh)']
dew_point = df['Dew Point(°C)']

# Tạo figure và axes
fig, ax1 = plt.subplots(figsize=(12, 6))

try:
    # Vẽ đường nhiệt độ (Temp. °C) trên trục Y bên trái
    ax1.plot(timestamps, temp, color='red', label='Temp (°C)', linewidth=1)
    # Vẽ đường ngưỡng dưới (Temp. LL °C)
    ax1.plot(timestamps, temp_ll, color='red', linestyle='--', label='Temp. LL (°C)', linewidth=1)
    # Vẽ đường ngưỡng trên (Temp. HL °C)
    ax1.plot(timestamps, temp_hl, color='red', linestyle='--', label='Temp. HL (°C)', linewidth=1)
    # Vẽ đường điểm sương (DewPoint °C)
    ax1.plot(timestamps, dew_point, color='black', label='DewPoint (°C)', linewidth=1)

    # Thiết lập trục Y bên trái
    ax1.set_xlabel('Timestamp')
    ax1.set_ylabel('Temperature (°C)', color='red')
    ax1.tick_params(axis='y', labelcolor='red')
    ax1.set_ylim(-30, 60)  # Phạm vi nhiệt độ theo hình mẫu
    ax1.grid(True)

    # Tạo trục Y bên phải cho độ ẩm (RH %rh)
    ax2 = ax1.twinx()
    ax2.plot(timestamps, rh, color='green', label='RH (%rh)', linewidth=1)
    ax2.plot(timestamps, rh_hl, color='green', linestyle='--', label='RH HL (%rh)', linewidth=1)
    ax2.set_ylabel('Relative Humidity (%rh)', color='green')
    ax2.tick_params(axis='y', labelcolor='green')
    ax2.set_ylim(10, 100)  # Phạm vi độ ẩm theo hình mẫu

    # Định dạng trục thời gian (chỉ hiển thị ngày)
    ax1.xaxis.set_major_formatter(DateFormatter('%d-%b-%y'))  # Định dạng ngày: DD-MMM-YY
    ax1.xaxis.set_major_locator(plt.MaxNLocator(10))  # Giảm số lượng nhãn ngày
    plt.xticks(rotation=45)

    # Thêm tiêu đề với số Serial
    plt.title(f'{serial_number}\nTemp. and Humi. Logger', pad=20)

    # Thêm chú thích (legend) với bố cục ngang, đặt ở dưới cùng
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    all_lines = lines1 + lines2
    all_labels = labels1 + labels2
    # Đảm bảo thứ tự giống hình mẫu
    legend_order = [0, 1, 2, 3, 4, 5]  # Temp (°C), Temp. LL (°C), Temp. HL (°C), DewPoint (°C), RH (%rh), RH HL (%rh)
    ax1.legend([all_lines[i] for i in legend_order], [all_labels[i] for i in legend_order],
               loc='lower center', ncol=6, bbox_to_anchor=(0.5, -0.2), frameon=False)

    # Hiển thị đồ thị
    plt.tight_layout()
    plt.show()
    print("Đồ thị đã được vẽ thành công!")

except Exception as e:
    print(f"Lỗi khi vẽ đồ thị: {e}")