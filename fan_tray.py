import os
import sys
import json
import time
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import pystray
from pystray import MenuItem as item
from PIL import Image, ImageDraw
import winreg
import ctypes
import win32com.client

CONFIG_PATH = 'config.json'
DEFAULT_INTERVAL = 5
MIN_INTERVAL = 2
MAX_INTERVAL = 15

def ensure_single_instance():
    mutex = ctypes.windll.kernel32.CreateMutexW(None, False, "ECToolTrayUniqueMutexName")
    last_error = ctypes.windll.kernel32.GetLastError()
    if last_error == 183:
        messagebox.showerror("错误", "程序已经在运行中。")
        sys.exit(0)

def save_config(path, interval):
    with open(CONFIG_PATH, 'w') as f:
        json.dump({'ectool_path': path, 'interval': interval}, f)

def load_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, 'r') as f:
            data = json.load(f)
            return data.get('ectool_path', ''), data.get('interval', DEFAULT_INTERVAL)
    return '', DEFAULT_INTERVAL

def get_startup_shortcut_path():
    startup_dir = os.path.join(os.environ['APPDATA'], r"Microsoft\Windows\Start Menu\Programs\Startup")
    exe_path = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)
    name = os.path.splitext(os.path.basename(exe_path))[0]
    shortcut_path = os.path.join(startup_dir, f"{name}.lnk")
    return shortcut_path, exe_path

def add_startup_shortcut():
    shortcut_path, exe_path = get_startup_shortcut_path()
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = exe_path
        shortcut.WorkingDirectory = os.path.dirname(exe_path)
        shortcut.IconLocation = exe_path
        shortcut.save()
        messagebox.showinfo("成功", "已添加到启动项（快捷方式）")
    except Exception as e:
        messagebox.showerror("错误", f"添加启动项失败：{e}")

def remove_startup_shortcut():
    shortcut_path, _ = get_startup_shortcut_path()
    try:
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)
            messagebox.showinfo("成功", "已从启动项中移除")
        else:
            messagebox.showinfo("提示", "启动项不存在")
    except Exception as e:
        messagebox.showerror("错误", f"删除启动项失败：{e}")

def parse_max_temperature(path):
    try:
        result = subprocess.run(
            [path, 'temps', 'all'],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        lines = result.stdout.strip().splitlines()
        temps = []
        for line in lines:
            if "Temp" in line and "=" in line:
                parts = line.split('=')
                if len(parts) > 1:
                    temp_c_str = parts[1].split('C')[0].strip()
                    try:
                        temp_c = int(temp_c_str)
                        temps.append(temp_c)
                    except ValueError:
                        continue
        return max(temps) if temps else None
    except:
        return None

def set_fan_speed(path, speed_percent):
    try:
        subprocess.run(
            [path, 'fanduty', str(speed_percent)],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
    except:
        pass

def loop_command(shared_path_ref, interval_ref):
    prev_duty = None
    smooth_temp = None
    alpha = 0.3

    while True:
        path = shared_path_ref[0]
        if not path or not os.path.exists(path):
            time.sleep(interval_ref[0])
            continue

        current_temp = parse_max_temperature(path)
        if current_temp is None:
            time.sleep(interval_ref[0])
            continue

        if smooth_temp is None:
            smooth_temp = current_temp
        else:
            smooth_temp = alpha * current_temp + (1 - alpha) * smooth_temp

        target_duty = max(30, min(100, int(2.5 * smooth_temp - 50)))

        if prev_duty is None:
            duty = target_duty
        else:
            if target_duty > prev_duty:
                duty = min(prev_duty + 20, target_duty)
            elif target_duty < prev_duty:
                duty = max(prev_duty - 20, target_duty)
            else:
                duty = prev_duty

        set_fan_speed(path, duty)
        prev_duty = duty
        time.sleep(interval_ref[0])

def create_image():
    img = Image.new('RGB', (64, 64), color='white')
    d = ImageDraw.Draw(img)
    d.rectangle((16, 16, 48, 48), fill='blue')
    return img

def on_quit(icon, item):
    icon.stop()
    os._exit(0)

def main():
    ensure_single_instance()

    root = tk.Tk()
    root.title("ECTool 风扇控制")
    root.geometry('420x300')

    initial_path, saved_interval = load_config()
    shared_path = [initial_path]
    interval_ref = [saved_interval]
    path_var = tk.StringVar(value=initial_path)
    interval_var = tk.StringVar(value=str(saved_interval))

    tk.Label(root, text="ectool.exe 路径：").pack(pady=5)
    tk.Entry(root, textvariable=path_var, width=50).pack()
    tk.Button(root, text="选择路径", command=lambda: select_path()).pack(pady=5)

    tk.Label(root, text="检测间隔（秒）：建议 2~15 秒").pack(pady=5)
    tk.Entry(root, textvariable=interval_var, width=10).pack()
    tk.Button(root, text="应用检测间隔", command=lambda: apply_interval()).pack(pady=5)

    tk.Label(root, text="启动项管理：").pack(pady=5)
    tk.Button(root, text="加入开机启动", command=add_startup_shortcut).pack(pady=2)
    tk.Button(root, text="移除开机启动", command=remove_startup_shortcut).pack(pady=2)

    def select_path():
        path = filedialog.askopenfilename(title="选择 ectool.exe", filetypes=[("可执行文件", "*.exe")])
        if path:
            path_var.set(path)
            shared_path[0] = path
            save_config(path, interval_ref[0])

    def apply_interval():
        try:
            val = int(interval_var.get())
            if val < MIN_INTERVAL or val > MAX_INTERVAL:
                messagebox.showwarning("无效输入", f"建议检测间隔为 {MIN_INTERVAL} ~ {MAX_INTERVAL} 秒。")
                interval_var.set(str(DEFAULT_INTERVAL))
                interval_ref[0] = DEFAULT_INTERVAL
            else:
                interval_ref[0] = val
                save_config(shared_path[0], val)
        except:
            messagebox.showwarning("错误", "请输入有效的数字")
            interval_var.set(str(DEFAULT_INTERVAL))
            interval_ref[0] = DEFAULT_INTERVAL

    def show_window():
        root.after(0, root.deiconify)

    def hide_window():
        root.withdraw()

    def on_closing():
        hide_window()

    root.protocol("WM_DELETE_WINDOW", on_closing)

    tray_icon = pystray.Icon("ECTool", create_image(), "ECTool 风扇控制", menu=pystray.Menu(
        item('显示窗口', lambda icon, item: show_window()),
        item('退出', on_quit)
    ))

    threading.Thread(target=loop_command, args=(shared_path, interval_ref), daemon=True).start()
    threading.Thread(target=tray_icon.run, daemon=True).start()

    if initial_path and os.path.exists(initial_path):
        root.after(3000, hide_window)

    root.mainloop()

if __name__ == "__main__":
    main()
