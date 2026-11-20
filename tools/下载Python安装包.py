"""把 Python 官方安装包下载到 vendor/python/。

打包时用一次；分享的包里已经带好，对方直接双击 1-安装环境.bat 即可。
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
TARGET_DIR = os.path.join(BASE_DIR, "vendor", "python")

# 3.12 是当前最稳的版本：库支持齐全，又不像 3.13/3.14 那样可能遇到兼容问题
VERSION = "3.12.10"
TARGET_NAME = f"python-{VERSION}-amd64.exe"

SOURCES = [
    f"https://www.python.org/ftp/python/{VERSION}/{TARGET_NAME}",
    f"https://mirrors.huaweicloud.com/python/{VERSION}/{TARGET_NAME}",
    f"https://npm.taobao.org/mirrors/python/{VERSION}/{TARGET_NAME}",
]

EXPECTED_MIN_BYTES = 20 * 1024 * 1024


def download(url, destination):
    with requests.get(url, stream=True, timeout=60, verify=False) as response:
        response.raise_for_status()
        total = int(response.headers.get("Content-Length") or 0)

        done = 0
        last = -1
        with open(destination, "wb") as handle:
            for chunk in response.iter_content(chunk_size=1024 * 128):
                if not chunk:
                    continue
                handle.write(chunk)
                done += len(chunk)
                if total:
                    percent = done * 100 // total
                    if percent >= last + 10:
                        last = percent
                        print(f"  {percent:3d}%  {done/1024/1024:5.1f} / "
                              f"{total/1024/1024:.1f} MB", flush=True)
        return done


def main():
    os.makedirs(TARGET_DIR, exist_ok=True)
    destination = os.path.join(TARGET_DIR, TARGET_NAME)

    if os.path.isfile(destination) and os.path.getsize(destination) >= EXPECTED_MIN_BYTES:
        print(f"已存在，跳过：{destination}")
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

    print("所有源都失败，请手动下载：")
    print(f"  https://www.python.org/ftp/python/{VERSION}/{TARGET_NAME}")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
