
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import MultipleLocator
from matplotlib.backends.backend_pdf import PdfPages

# === Đường dẫn tới file Excel ===
file_path = "D:/TMPs/To0206-0606.xlsx"  # Đổi tại đây nếu cần
df = pd.read_excel(file_path, header=0)
df.columns = df.columns.str.strip()

# Làm sạch dữ liệu
df["Timestamp"] = pd.to_datetime(df["Timestamp"], errors="coerce")
df = df.dropna(subset=["Timestamp", "Temp.(°C)", "RH(%rh)"])

# Serial và logger
serial_number = df["Serial Number"].dropna().iloc[0] if "Serial Number" in df else "UNKNOWN"
logger_name = df["Logger Name"].dropna().iloc[0] if "Logger Name" in df else "Logger"

# === Vẽ biểu đồ ===
fig, ax1 = plt.subplots(figsize=(11.7, 8.3))  # A4 landscape
fig.patch.set_facecolor('white')

# Nhiệt độ
ax1.plot(df["Timestamp"], df["Temp.(°C)"], color='red', label='Temp.(°C)')
if "Temp. LL(°C)" in df:
    ax1.plot(df["Timestamp"], df["Temp. LL(°C)"], 'brown', linestyle='-.', label='Temp. LL(°C)')
if "Temp. HL(°C)" in df:
    ax1.plot(df["Timestamp"], df["Temp. HL(°C)"], 'darkred', linestyle='--', label='Temp. HL(°C)')
if "Dew Point(°C)" in df:
    ax1.plot(df["Timestamp"], df["Dew Point(°C)"], color='navy', label='Dew Point(°C)')

ax1.set_ylabel("Temperature (°C)", color='black')
ax1.set_ylim(-30, 60)
ax1.yaxis.set_major_locator(MultipleLocator(9))

# RH
ax2 = ax1.twinx()
ax2.plot(df["Timestamp"], df["RH(%rh)"], color='green', label='RH(%rh)')
if "RH HL(%rh)" in df:
    ax2.plot(df["Timestamp"], df["RH HL(%rh)"], 'lime', linestyle='--', label='RH HL(%rh)')
ax2.set_ylabel("Relative Humidity (%rh)", color='black')
ax2.set_ylim(10, 104)

# Trục thời gian
ax1.xaxis.set_major_locator(mdates.DayLocator())
ax1.xaxis.set_major_formatter(mdates.DateFormatter('%d-%b-%y'))
ax1.xaxis.set_minor_locator(mdates.HourLocator(interval=6))
ax1.xaxis.set_minor_formatter(mdates.DateFormatter('%H:%M'))
plt.setp(ax1.xaxis.get_minorticklabels(), rotation=90, fontsize=8)

# === Tiêu đề gạch chân ===
ax1.text(0.5, 1.15, "Temp. and Humi. Logger", transform=ax1.transAxes,
         fontsize=15, fontweight='bold', ha='center')
ax1.plot([0.35, 0.65], [1.13, 1.13], transform=ax1.transAxes,
         color='black', linewidth=1)

# Serial info bên trái
ax1.text(0.01, 1.09, f"{serial_number} - {logger_name}",
         transform=ax1.transAxes, fontsize=9, ha='left')

# Chú thích
lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper center', ncol=4, frameon=True)

# Ghi chú thời gian
start_time = df["Timestamp"].min().strftime('%d-%b-%y %H:%M:%S')
end_time = df["Timestamp"].max().strftime('%d-%b-%y %H:%M:%S')
plt.figtext(0.5, 0.01, f'From: {start_time}  To: {end_time}', ha='center', fontsize=10)

# === Căn giữa nội dung khi in ===
plt.subplots_adjust(left=0.07, right=0.93, top=0.85, bottom=0.10)

# Xuất PDF
output_path = file_path.replace(".xlsx", "_centered_plot.pdf")
with PdfPages(output_path) as pdf:
    pdf.savefig(fig, dpi=300)

print(f"✅ Xuất PDF thành công: {output_path}")
