import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import MultipleLocator

# Đường dẫn tới file Excel
file_path = "D:/TMPs/fr1006_1606cb01.xlsx"

# Đọc file
xlsx = pd.ExcelFile(file_path)
df = pd.read_excel(xlsx, sheet_name=0)

# Làm sạch tên cột và dữ liệu
df.columns = df.columns.str.strip()
df["Timestamp"] = pd.to_datetime(df["Timestamp"], errors='coerce')
df = df.dropna(subset=["Timestamp", "Temp.(°C)", "RH(%rh)"])

# Trích Serial Number và Logger Name (lấy giá trị đầu tiên không rỗng)
serial_number = df["Serial Number"].dropna().iloc[0] if "Serial Number" in df else "UNKNOWN"
logger_name = df["Logger Name"].dropna().iloc[0] if "Logger Name" in df else "Logger"

# Tạo đồ thị
fig, ax1 = plt.subplots(figsize=(14, 7))

# Vẽ các đường nhiệt độ
ax1.plot(df["Timestamp"], df["Temp.(°C)"], color='red', label='Temp.(°C)')
if "Temp. LL(°C)" in df.columns:
    ax1.plot(df["Timestamp"], df["Temp. LL(°C)"], 'r--', label='Temp. LL(°C)')
if "Temp. HL(°C)" in df.columns:
    ax1.plot(df["Timestamp"], df["Temp. HL(°C)"], 'r:', label='Temp. HL(°C)')
if "Dew Point(°C)" in df.columns:
    ax1.plot(df["Timestamp"], df["Dew Point(°C)"], color='blue', label='Dew Point(°C)')

ax1.set_ylabel("Temperature (°C)", color='black')
ax1.set_ylim(-30, 60)
ax1.yaxis.set_major_locator(MultipleLocator(9))

# Trục RH bên phải
ax2 = ax1.twinx()
ax2.plot(df["Timestamp"], df["RH(%rh)"], color='green', label='RH(%rh)')
if "RH HL(%rh)" in df.columns:
    ax2.plot(df["Timestamp"], df["RH HL(%rh)"], 'g--', label='RH HL(%rh)')
ax2.set_ylabel("Relative Humidity (%rh)", color='black')
ax2.set_ylim(10, 104)

# Định dạng trục thời gian
ax1.xaxis.set_major_locator(mdates.DayLocator())
ax1.xaxis.set_major_formatter(mdates.DateFormatter('%d-%b-%y'))
ax1.xaxis.set_minor_locator(mdates.HourLocator(interval=6))
ax1.xaxis.set_minor_formatter(mdates.DateFormatter('%H:%M'))
plt.setp(ax1.xaxis.get_minorticklabels(), rotation=90, fontsize=8)

# Thêm tiêu đề
plt.title(f"{serial_number} - {logger_name}", fontsize=14, fontweight='bold')

# Gộp chú thích
lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper center', ncol=4, frameon=True)

# Ghi chú khoảng thời gian
start_time = df["Timestamp"].min().strftime('%d-%b-%y %H:%M:%S')
end_time = df["Timestamp"].max().strftime('%d-%b-%y %H:%M:%S')
plt.figtext(0.5, 0.01, f'From: {start_time}  To: {end_time}', ha='center', fontsize=10)

plt.tight_layout()
plt.grid(True)
plt.show()
