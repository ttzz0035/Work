from ctypes import windll
import subprocess
import re

def get_interface_names():
    """
    使用可能なインターフェース名の一覧を返す（日本語対応）
    """
    try:
        result = subprocess.run(
            ["netsh", "interface", "show", "interface"],
            capture_output=True, text=True, encoding="cp932"
        )
        output = result.stdout
        if not output:
            raise RuntimeError("netsh 出力が空です")

        interfaces = []
        lines = output.splitlines()
        for line in lines:
            if re.search(r'(有効|無効)', line):  # 日本語環境対応
                parts = line.strip().split()
                if len(parts) >= 4:
                    name = " ".join(parts[3:])
                    interfaces.append(name)
        return interfaces
    except Exception as e:
        print(f"[ERROR] インターフェース取得失敗: {e}")
        return []

def set_ip_address(if_name, ip, subnet, gateway):
    """
    静的IPを設定
    """
    shell = windll.shell32.ShellExecuteW
    cmd = f'netsh interface ip set address name="{if_name}" static {ip} {subnet} {gateway} 1'
    rc = shell(None, 'runas', 'cmd.exe', f'/c {cmd}', None, 0)
    return rc

def set_network_off(if_name):
    """
    ネットワークインターフェースを無効化
    """
    shell = windll.shell32.ShellExecuteW
    rc = shell(None, 'runas', 'cmd.exe',
               f'/c netsh interface set interface name="{if_name}" admin=disable', None, 0)
    return rc

def set_network_on(if_name):
    """
    ネットワークインターフェースを有効化
    """
    shell = windll.shell32.ShellExecuteW
    rc = shell(None, 'runas', 'cmd.exe',
               f'/c netsh interface set interface name="{if_name}" admin=enable', None, 0)
    return rc

def get_interface_config_from_netsh(if_name):
    """
    現在のIP, サブネット, ゲートウェイ, DHCPモードを取得
    """
    try:
        result = subprocess.run(
            ["netsh", "interface", "ip", "show", "config", f'name={if_name}'],
            capture_output=True, text=True, encoding="cp932"
        )
        output = result.stdout
        print(f"[DEBUG] netsh output:\n{output}")

        ip = subnet = gateway = ""
        mode = "static"

        for line in output.splitlines():
            if "DHCP" in line:
                if "はい" in line or "Yes" in line:
                    mode = "dhcp"
            if "IP アドレス" in line or "IP Address" in line:
                ip_match = re.search(r"(\d{1,3}(?:\.\d{1,3}){3})", line)
                if ip_match:
                    ip = ip_match.group(1)
            if "マスク" in line or "mask" in line:
                subnet_match = re.search(r"マスク\s+(\d{1,3}(?:\.\d{1,3}){3})", line)
                if not subnet_match:
                    subnet_match = re.search(r"mask\s+(\d{1,3}(?:\.\d{1,3}){3})", line)
                if subnet_match:
                    subnet = subnet_match.group(1)
            if "デフォルト ゲートウェイ" in line or "Default Gateway" in line:
                gw_match = re.search(r"(\d{1,3}(?:\.\d{1,3}){3})", line)
                if gw_match:
                    gateway = gw_match.group(1)

        return ip, subnet, gateway, mode
    except Exception as e:
        print(f"[ERROR] netsh解析失敗: {e}")
        return "", "", "", "static"

def set_dhcp(if_name, known_static_ip=None):
    """
    DHCPに切り替える（INIに記載されたIPを元に削除対象を特定）
    known_static_ip: INIから取得したIP。指定がない場合のみ自動取得
    """
    shell = windll.shell32.ShellExecuteW

    # --- IP削除処理（INIベース） ---
    if known_static_ip:
        print(f"[DEBUG] delete static IP + gateway from INI: {known_static_ip}")
        shell(None, 'runas', 'cmd.exe',
              f'/c netsh interface ip delete address name="{if_name}" address={known_static_ip} gateway=all',
              None, 0)
    else:
        ip, _, _, _ = get_interface_config_from_netsh(if_name)
        if ip:
            print(f"[DEBUG] delete static IP + gateway from live config: {ip}")
            shell(None, 'runas', 'cmd.exe',
                  f'/c netsh interface ip delete address name="{if_name}" address={ip} gateway=all',
                  None, 0)

    # --- DHCP再設定（IP + DNS） ---
    rc_ip = shell(None, 'runas', 'cmd.exe',
                  f'/c netsh interface ip set address name="{if_name}" source=dhcp', None, 0)
    rc_dns = shell(None, 'runas', 'cmd.exe',
                   f'/c netsh interface ip set dns name="{if_name}" source=dhcp', None, 0)

    success = (rc_ip > 32 and rc_dns > 32)
    print(f"[DEBUG] set_dhcp: ip_rc={rc_ip}, dns_rc={rc_dns}, success={success}")
    return success
