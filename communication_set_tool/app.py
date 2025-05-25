import tkinter as tk
from tkinter import ttk, messagebox
import configparser
import net_setting
import config_util

EXCLUDE_SECTIONS = ("COMMON", "DISABLE_INTERFACE")

class IPChangerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("IPアドレス変更ツール")
        self.config = configparser.ConfigParser()
        self.prev_config = configparser.ConfigParser()
        self.if_list = net_setting.get_interface_names()
        self.profile_list = []

        self.create_widgets()
        self.load_ini()

    def create_widgets(self):
        frame = ttk.Frame(self.root, padding=10)
        frame.grid(row=0, column=0, sticky='nsew')

        # ラベル・入力欄
        ttk.Label(frame, text="設定プロファイル:").grid(row=0, column=0, sticky='e', padx=5, pady=3)
        self.profile_name = tk.StringVar()
        self.profile_dropdown = ttk.Combobox(frame, textvariable=self.profile_name, state="readonly", width=30)
        self.profile_dropdown.grid(row=0, column=1, padx=5, pady=3)
        self.profile_dropdown.bind("<<ComboboxSelected>>", self.on_profile_selected)

        ttk.Label(frame, text="IP設定対象IF:").grid(row=1, column=0, sticky='e', padx=5, pady=3)
        self.ip_if_name = tk.StringVar()
        self.ip_if_dropdown = ttk.Combobox(frame, textvariable=self.ip_if_name, values=self.if_list, width=30)
        self.ip_if_dropdown.grid(row=1, column=1, padx=5, pady=3)

        ttk.Label(frame, text="Wi-Fi無効対象IF:").grid(row=2, column=0, sticky='e', padx=5, pady=3)
        self.wifi_if_name = tk.StringVar()
        self.wifi_if_dropdown = ttk.Combobox(frame, textvariable=self.wifi_if_name, values=self.if_list, width=30)
        self.wifi_if_dropdown.grid(row=2, column=1, padx=5, pady=3)

        ttk.Label(frame, text="IPアドレス:").grid(row=3, column=0, sticky='e', padx=5, pady=3)
        self.ip_entry = ttk.Entry(frame, width=30)
        self.ip_entry.grid(row=3, column=1, padx=5, pady=3)

        ttk.Label(frame, text="サブネット:").grid(row=4, column=0, sticky='e', padx=5, pady=3)
        self.subnet_entry = ttk.Entry(frame, width=30)
        self.subnet_entry.grid(row=4, column=1, padx=5, pady=3)

        ttk.Label(frame, text="ゲートウェイ:").grid(row=5, column=0, sticky='e', padx=5, pady=3)
        self.gateway_entry = ttk.Entry(frame, width=30)
        self.gateway_entry.grid(row=5, column=1, padx=5, pady=3)

        # 横並びボタン配置用フレーム
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=6, column=0, columnspan=2, pady=10)

        apply_btn = ttk.Button(button_frame, text="IP変更 + Wi-Fiオフ", command=self.apply_changes)
        apply_btn.grid(row=0, column=0, padx=10)

        restore_btn = ttk.Button(button_frame, text="元に戻す", command=self.restore_previous)
        restore_btn.grid(row=0, column=1, padx=10)

    def load_ini(self):
        try:
            self.config = config_util.load_ini("ipconfig.ini")

            self.profile_list = [sec for sec in self.config.sections() if sec not in EXCLUDE_SECTIONS]
            self.profile_dropdown["values"] = self.profile_list
            if self.profile_list:
                self.profile_name.set(self.profile_list[0])
                self.load_profile(self.profile_list[0])

            self.ip_if_name.set(self.config.get("COMMON", "ip_interface", fallback=""))
            self.wifi_if_name.set(self.config.get("DISABLE_INTERFACE", "wifi_interface", fallback=""))
        except Exception as e:
            messagebox.showerror("エラー", f"INI読込失敗: {e}")

    def load_profile(self, section):
        try:
            self.ip_entry.delete(0, tk.END)
            self.ip_entry.insert(0, self.config.get(section, "ip", fallback=""))
            self.subnet_entry.delete(0, tk.END)
            self.subnet_entry.insert(0, self.config.get(section, "subnet", fallback=""))
            self.gateway_entry.delete(0, tk.END)
            self.gateway_entry.insert(0, self.config.get(section, "gateway", fallback=""))
        except Exception as e:
            messagebox.showerror("プロファイル読み込みエラー", str(e))

    def on_profile_selected(self, event):
        self.load_profile(self.profile_name.get())

    def apply_changes(self):
        ip_if = self.ip_if_name.get()
        wifi_if = self.wifi_if_name.get()
        new_ip = self.ip_entry.get()
        new_subnet = self.subnet_entry.get()
        new_gateway = self.gateway_entry.get()

        current_ip, current_subnet, current_gateway, current_mode = net_setting.get_interface_config_from_netsh(ip_if)

        prev = configparser.ConfigParser()
        prev["NETWORK"] = {
            "ip_interface": ip_if,
            "ip": current_ip,
            "subnet": current_subnet,
            "gateway": current_gateway,
            "mode": current_mode
        }
        prev["DISABLE_INTERFACE"] = {
            "wifi_interface": wifi_if
        }
        prev["APPLIED"] = {
            "ip": new_ip
        }

        try:
            config_util.save_ini(prev, "prev_config.ini")
        except Exception as e:
            messagebox.showerror("保存エラー", str(e))
            return

        rc1 = net_setting.set_ip_address(ip_if, new_ip, new_subnet, new_gateway)
        rc2 = net_setting.set_network_off(wifi_if)

        messagebox.showinfo("完了", f"IP設定: {rc1}\nWi-Fi無効化: {rc2}")

    def restore_previous(self):
        try:
            self.prev_config = config_util.load_ini("prev_config.ini")

            ip_if = self.prev_config.get("NETWORK", "ip_interface")
            ip = self.prev_config.get("NETWORK", "ip", fallback="")
            subnet = self.prev_config.get("NETWORK", "subnet", fallback="")
            gateway = self.prev_config.get("NETWORK", "gateway", fallback="")
            mode = self.prev_config.get("NETWORK", "mode", fallback="")
            ip_applied = self.prev_config.get("APPLIED", "ip", fallback="")
            wifi_if = self.prev_config.get("DISABLE_INTERFACE", "wifi_interface")

            net_setting.set_network_on(wifi_if)

            if not mode:
                mode = "dhcp" if not ip or not subnet else "static"

            if mode == "dhcp":
                rc = net_setting.set_dhcp(ip_if, known_static_ip=ip_applied)
                if rc:
                    messagebox.showinfo("復元完了", f"{ip_if} を DHCP に復元しました")
                else:
                    messagebox.showerror("復元失敗", f"{ip_if} の DHCP 設定に失敗しました")
            else:
                rc = net_setting.set_ip_address(ip_if, ip, subnet, gateway)
                messagebox.showinfo("復元完了", f"{ip_if} を静的設定に復元しました")
        except Exception as e:
            messagebox.showerror("復元エラー", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = IPChangerApp(root)
    root.mainloop()
