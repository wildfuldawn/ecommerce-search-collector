"""把 Chrome 官方离线安装包下载到 vendor/browser/。

只在打包时用一次；分享出去的包里已经带好安装包，对方不需要运行这个脚本。
"""

import os
import shutil
import sys

import requests

try:
    import urllib3
    urllib3.disable_warnings()
except Exception:
    pass

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # 项目根
TARGET_DIR = os.path.join(BASE_DIR, "vendor", "browser")
TARGET_NAME = "ChromeStandaloneSetup64.exe"

# 依次尝试，第一个能用的就下
SOURCES = [
    "https://dl.google.com/chrome/install/ChromeStandaloneSetup64.exe",
    "https://dl.google.com/tag/s/dl/chrome/install/googlechromestandaloneenterprise64.msi",
]

EXPECTED_MIN_BYTES = 100 * 1024 * 1024      # 小于 100 MB 视为下载不完整


def download(url, destination):
    with requests.get(url, stream=True, timeout=60, verify=False) as response:
        response.raise_for_status()
        total = int(response.headers.get("Content-Length") or 0)

        done = 0
        last_report = -1
        with open(destination, "wb") as handle:
            for chunk in response.iter_content(chunk_size=1024 * 256):
                if not chunk:
                    continue
                handle.write(chunk)
                done += len(chunk)
                if total:
                    percent = done * 100 // total
                    if percent >= last_report + 5:
                        last_report = percent
                        print(f"  {percent:3d}%  {done/1024/1024:6.1f} / "
                              f"{total/1024/1024:.1f} MB", flush=True)
        return done


def main():
    os.makedirs(TARGET_DIR, exist_ok=True)
    destination = os.path.join(TARGET_DIR, TARGET_NAME)

    if os.path.isfile(destination) and os.path.getsize(destination) >= EXPECTED_MIN_BYTES:
        size = os.path.getsize(destination) / 1024 / 1024
        print(f"已存在，跳过下载：{destination}（{size:.1f} MB）")
        return 0

    for url in SOURCES:
        print(f"尝试：{url}")
        temporary = destination + ".part"
        try:
            written = download(url, temporary)
        except Exception as error:
            print(f"  失败：{type(error).__name__}: {error}")
            if os.path.isfile(temporary):
                os.remove(temporary)
            continue

        if written < EXPECTED_MIN_BYTES:
            print(f"  文件过小（{written} 字节），换下一个源")
            os.remove(temporary)
            continue

        shutil.move(temporary, destination)
        print(f"完成：{destination}（{written/1024/1024:.1f} MB）")
        return 0

    print("所有源都失败。可以改用 playwright 自带内核：")
    print("  python -m playwright install chromium")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
