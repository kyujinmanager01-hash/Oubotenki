"""
Kết nối VPN Gate (server Nhật) trên GitHub Actions, chỉ cho trình duyệt đi qua VPN.

Cách hoạt động:
  1. Tải danh sách server từ API VPN Gate, lọc server Nhật (JP).
  2. Thử lần lượt từng server bằng OpenVPN (không nhận route từ server => runner
     và Google API vẫn đi thẳng, KHÔNG đi qua VPN).
  3. Mở tinyproxy (127.0.0.1:18888) gửi traffic ra qua cổng tunnel của VPN.
  4. Ghi PROXY_SERVER=http://127.0.0.1:18888 vào $GITHUB_ENV để aw.py / kbx.py dùng.

Chỉ dùng thư viện chuẩn Python. Cần sẵn: openvpn, tinyproxy, curl, sudo.
"""
import base64
import csv
import io
import os
import re
import subprocess
import sys
import time
import urllib.request

API_URL = "https://www.vpngate.net/api/iphone/"
PROXY_PORT = 18888
TUN_DEV = "tunjp"
TABLE = "100"
MAX_SERVERS = 15          # số server thử tối đa
CONNECT_TIMEOUT = 30      # giây chờ mỗi server
OVPN_PATH = "/tmp/vpngate.ovpn"
OVPN_LOG = "/tmp/vpngate.log"
TINY_CONF = "/tmp/tinyproxy.conf"
PROXY_URL = f"http://127.0.0.1:{PROXY_PORT}"

# Dòng trong file .ovpn bị loại bỏ (an toàn + không nhận route/DNS từ server lạ)
DROP_PREFIXES = (
    "redirect-gateway", "dhcp-option", "route ", "route-", "up ", "down ",
    "plugin", "script-security", "tls-verify", "ipchange", "learn-address",
    "client-connect", "client-disconnect", "setenv", "dev ", "dev-type",
    "daemon", "log", "writepid",
)


def log(msg):
    print(f"[vpngate] {msg}", flush=True)


def sh_bg(cmd):
    """Chạy lệnh tự chạy nền (daemon): KHÔNG giữ ống output, nếu không sẽ bị treo."""
    return subprocess.run(cmd, shell=True, check=False, timeout=30,
                          stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                          stderr=subprocess.DEVNULL)


def sh(cmd, check=True, **kw):
    return subprocess.run(cmd, shell=isinstance(cmd, str), check=check,
                          capture_output=True, text=True, **kw)


def fetch_servers():
    """Trả về list (hostname, ip, score, config_text) của server Nhật, điểm cao trước."""
    last_err = None
    for attempt in range(3):
        try:
            with urllib.request.urlopen(API_URL, timeout=30) as r:
                raw = r.read().decode("utf-8", "replace")
            return parse_servers(raw)
        except Exception as e:  # noqa: BLE001
            last_err = e
            log(f"Tải danh sách lần {attempt + 1} lỗi: {e}")
            time.sleep(3)
    raise SystemExit(f"Không tải được danh sách VPN Gate: {last_err}")


def parse_servers(raw):
    lines = [l for l in raw.splitlines() if l and not l.startswith("*")]
    if lines and lines[0].startswith("#"):
        lines[0] = lines[0][1:]  # bỏ dấu # ở dòng header
    servers = []
    for row in csv.DictReader(io.StringIO("\n".join(lines))):
        if row.get("CountryShort") != "JP":
            continue
        b64 = row.get("OpenVPN_ConfigData_Base64")
        if not b64:
            continue
        try:
            cfg = base64.b64decode(b64).decode("utf-8", "replace")
            score = int(row.get("Score") or 0)
        except Exception:  # noqa: BLE001
            continue
        servers.append((row.get("HostName"), row.get("IP"), score, cfg))
    servers.sort(key=lambda s: s[2], reverse=True)
    return servers


def build_config(cfg):
    kept = []
    for line in cfg.splitlines():
        s = line.strip()
        if any(s.startswith(p) for p in DROP_PREFIXES):
            continue
        kept.append(line)
    kept += [
        f"dev-type tun",
        f"dev {TUN_DEV}",
        "route-nopull",                 # KHÔNG nhận route => chỉ proxy đi qua VPN
        "script-security 1",             # chặn mọi script người dùng trong file cấu hình
        "providers legacy default",     # VPN Gate dùng thuật toán cũ
        "tls-cipher DEFAULT:@SECLEVEL=0",
        "data-ciphers AES-256-GCM:AES-128-GCM:AES-256-CBC:AES-128-CBC",
        "data-ciphers-fallback AES-128-CBC",
        "connect-retry-max 1",
        "verb 3",
    ]
    return "\n".join(kept) + "\n"


def stop_all():
    sh("sudo pkill -x openvpn", check=False)
    sh("pkill -x tinyproxy", check=False)
    sh(f"sudo ip rule del table {TABLE}", check=False)
    sh(f"sudo ip route flush table {TABLE}", check=False)
    time.sleep(1)


def tun_ip():
    out = sh(f"ip -4 -o addr show dev {TUN_DEV}", check=False).stdout
    m = re.search(r"inet (\d+\.\d+\.\d+\.\d+)", out)
    return m.group(1) if m else None


def start_openvpn(cfg):
    """Chạy OpenVPN ở foreground (Popen), ghi toàn bộ output vào OVPN_LOG."""
    with open(OVPN_PATH, "w") as f:
        f.write(build_config(cfg))
    logf = open(OVPN_LOG, "w")
    proc = subprocess.Popen(
        ["sudo", "openvpn", "--config", OVPN_PATH],
        stdin=subprocess.DEVNULL, stdout=logf, stderr=subprocess.STDOUT,
        start_new_session=True,
    )
    deadline = time.time() + CONNECT_TIMEOUT
    ok = False
    while time.time() < deadline:
        if "Initialization Sequence Completed" in open(OVPN_LOG).read():
            ok = True
            break
        if proc.poll() is not None:  # OpenVPN đã thoát => lỗi
            break
        time.sleep(1)
    if ok:
        return tun_ip()
    code = proc.poll()
    log(f"    openvpn exit code: {code if code is not None else 'còn chạy (hết thời gian chờ)'}")
    lines = [l for l in open(OVPN_LOG).read().splitlines() if l.strip()]
    for l in lines[-8:]:
        log(f"    openvpn: {l}")
    if not lines:
        log("    openvpn: (log trống)")
    return None


def start_proxy(ip):
    # Gói tin có source = IP của tunnel => đi ra cổng tunnel (bảng route riêng)
    sh(f"sudo ip rule add from {ip} table {TABLE}")
    sh(f"sudo ip route add default dev {TUN_DEV} table {TABLE}")
    with open(TINY_CONF, "w") as f:
        f.write(
            f"Port {PROXY_PORT}\nListen 127.0.0.1\nBind {ip}\n"
            "Allow 127.0.0.1\nTimeout 600\nMaxClients 50\n"
            "ConnectPort 443\nConnectPort 80\nDisableViaHeader Yes\n"
            "LogFile \"/tmp/tinyproxy.log\"\nPidFile \"/tmp/tinyproxy.pid\"\n"
        )
    sh_bg(f"tinyproxy -c {TINY_CONF}")
    time.sleep(1)


def curl(url, extra=""):
    r = sh(f"curl -s -m 20 {extra} -x {PROXY_URL} {url}", check=False)
    return r.stdout.strip()


def verify():
    """Xác nhận IP ra ngoài là Nhật và vào được trang đích."""
    country = ""
    for u in ("https://ipinfo.io/country", "https://ifconfig.co/country-iso"):
        country = curl(u)
        if country:
            break
    log(f"IP ra ngoài thuộc nước: {country or '?'}")
    if country != "JP":
        return False
    for u in ("https://ats.rct.airwork.net/entries", "https://saiyo.kyujinbox.com/"):
        code = curl(u, "-o /dev/null -w %{http_code}")
        log(f"  {u} -> HTTP {code}")
        if code in ("", "000"):
            return False
    return True


def main():
    log("phiên bản script: v3")
    log("openvpn: " + (sh("openvpn --version | head -1", check=False).stdout.strip() or "KHÔNG TÌM THẤY"))
    log("/dev/net/tun: " + ("có" if os.path.exists("/dev/net/tun") else "KHÔNG CÓ"))
    servers = fetch_servers()
    log(f"Có {len(servers)} server Nhật, thử tối đa {MAX_SERVERS}")
    for host, ip, score, cfg in servers[:MAX_SERVERS]:
        log(f"Thử {host} ({ip}) score={score}")
        stop_all()
        tip = start_openvpn(cfg)
        if not tip:
            log("  Không kết nối được, bỏ qua")
            continue
        try:
            start_proxy(tip)
            if verify():
                log(f"OK: dùng {host} ({ip}); proxy {PROXY_URL}")
                env_file = os.environ.get("GITHUB_ENV")
                if env_file:
                    with open(env_file, "a") as f:
                        f.write(f"PROXY_SERVER={PROXY_URL}\n")
                return 0
        except subprocess.CalledProcessError as e:
            log(f"  Lỗi cấu hình proxy: {e.stderr}")
        log("  Kiểm tra thất bại, thử server tiếp theo")
    stop_all()
    log("Không có server VPN Gate nào dùng được. Thử chạy lại sau vài phút.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
