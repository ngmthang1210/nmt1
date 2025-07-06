
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import MultipleLocator
from matplotlib.backends.backend_pdf import PdfPages
from tkinter import Tk, filedialog, messagebox

def run_plot_from_file():
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Chọn file Excel (.xlsx)",
        filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        messagebox.showinfo("Huỷ bỏ", "Không có file nào được chọn.")
        return

    try:
        df = pd.read_excel(file_path, header=0)
        df.columns = df.columns.str.strip()
        df["Timestamp"] = pd.to_datetime(df["Timestamp"], errors="coerce")
        df = df.dropna(subset=["Timestamp", "Temp.(°C)", "RH(%rh)"])

        serial_number = df["Serial Number"].dropna().iloc[0] if "Serial Number" in df else "UNKNOWN"
        logger_name = df["Logger Name"].dropna().iloc[0] if "Logger Name" in df else "Logger"

        # Lấy thời điểm gần 0h hoặc 12h mỗi ngày
        df["DateOnly"] = df["Timestamp"].dt.date
        most_common_minute = df["Timestamp"].dt.minute.value_counts().idxmax()
        df_filtered = df[(df["Timestamp"].dt.hour.isin([0, 12])) &
                         (df["Timestamp"].dt.minute == most_common_minute)]
        time_per_day = df_filtered.groupby("DateOnly").first()["Timestamp"]
        xtick_times = time_per_day
        xtick_dates = pd.to_datetime(time_per_day.dt.normalize())

        # Vẽ biểu đồ
        fig, ax1 = plt.subplots(figsize=(11.7, 8.3))
        fig.patch.set_facecolor('white')
        ax1.plot(df["Timestamp"], df["Temp.(°C)"], color='red', linewidth=1)
        if "Temp. LL(°C)" in df:
            ax1.plot(df["Timestamp"], df["Temp. LL(°C)"], color='brown', linestyle='-.', linewidth=1)
        if "Temp. HL(°C)" in df:
            ax1.plot(df["Timestamp"], df["Temp. HL(°C)"], color='darkred', linestyle='--', linewidth=1)
        if "Dew Point(°C)" in df:
            ax1.plot(df["Timestamp"], df["Dew Point(°C)"], color='navy', linewidth=1)
        ax2 = ax1.twinx()
        ax2.plot(df["Timestamp"], df["RH(%rh)"], color='green', linewidth=1)
        if "RH HL(%rh)" in df:
            ax2.plot(df["Timestamp"], df["RH HL(%rh)"], color='lime', linestyle='--', linewidth=1)

        ax1.set_ylabel("Temperature (°C)")
        ax2.set_ylabel("Relative Humidity (%rh)")
        ax1.set_ylim(-30, 60)
        ax2.set_ylim(10, 104)
        ax1.yaxis.set_major_locator(MultipleLocator(9))
        ax1.grid(True, linestyle=':', linewidth=0.5)

        ax1.set_xticks(xtick_times)
        ax1.set_xticklabels([""] * len(xtick_times))
        ax1.tick_params(axis='x', which='major', length=0)

        for d, t in zip(xtick_dates, xtick_times):
            ax1.text(t, -0.07, t.strftime('%H:%M'), transform=ax1.get_xaxis_transform(),
                     fontsize=8, ha='center', va='top')
            ax1.text(t, -0.13, d.strftime('%d-%b-%y'), transform=ax1.get_xaxis_transform(),
                     fontsize=9, ha='center', va='top')

        ax1.text(0.5, 1.17, "Temp. and Humi. Logger", transform=ax1.transAxes,
                 fontsize=15, fontweight='bold', ha='center')
        ax1.annotate('', xy=(0.33, 1.15), xytext=(0.67, 1.15),
                     xycoords='axes fraction',
                     arrowprops=dict(arrowstyle='-', color='black', linewidth=1))
        ax1.text(0.01, 1.11, f"{serial_number} - {logger_name}",
                 transform=ax1.transAxes, fontsize=9, ha='left')

        # Chú thích
        legend_entries = [
            [("Temp.(°C)", 'red', '-'), ("Temp. LL(°C)", 'brown', '-.'), ("Temp. HL(°C)", 'darkred', '--')],
            [("RH(%rh)", 'green', '-'), ("RH HL(%rh)", 'lime', '--'), ("Dew Point(°C)", 'navy', '-')]
        ]
        x0, y0, dx, dy = 0.53, 1.11, 0.14, 0.035
        for row_idx, row in enumerate(legend_entries):
            for col_idx, (label, color, style) in enumerate(row):
                xpos = x0 + col_idx * dx
                ypos = y0 - row_idx * dy
                ax1.plot([xpos, xpos + 0.025], [ypos + 0.008, ypos + 0.008],
                         transform=ax1.transAxes, color=color, linestyle=style,
                         linewidth=2, clip_on=False)
                ax1.text(xpos + 0.03, ypos, label, transform=ax1.transAxes,
                         fontsize=9, ha='left', va='bottom', color='black')

        start_time = df["Timestamp"].min().strftime('%d-%b-%y %H:%M:%S')
        end_time = df["Timestamp"].max().strftime('%d-%b-%y %H:%M:%S')
        plt.figtext(0.5, 0.01, f'From: {start_time}  To: {end_time}', ha='center', fontsize=10)
        plt.subplots_adjust(left=0.07, right=0.93, top=0.83, bottom=0.18)

        output_path = file_path.replace(".xlsx", "_dlg_style.pdf")
        with PdfPages(output_path) as pdf:
            pdf.savefig(fig, dpi=300)

        messagebox.showinfo("Thành công", f"✅ Đã xuất PDF:
{output_path}")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Lỗi khi xử lý file:
                             {e}")

if __name__ == "__main__":
    run_plot_from_file()
