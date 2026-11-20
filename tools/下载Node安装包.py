"""把 Node.js 安装包下载到 vendor/node/。

京东（test_fixed.js）和淘宝（execjs / aef.js）都要靠 node 跑 JS 签名。
"""

import os
import shutil

import requests

try:
    import urllib3
    urllib3.disable_warnings()
except Exception:
    pass

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # 项目根
TARGET_DIR = os.path.join(BASE_DIR, "vendor", "node")

VERSION = "20.18.1"
TARGET_NAME = f"node-v{VERSION}-x64.msi"

SOURCES = [
    # 国内镜像优先，快且稳
    f"https://cdn.npmmirror.com/binaries/node/v{VERSION}/{TARGET_NAME}",
    f"https://npmmirror.com/mirrors/node/v{VERSION}/{TARGET_NAME}",
    f"https://nodejs.org/dist/v{VERSION}/{TARGET_NAME}",
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
            print(f"  文件过小，换下一个源")
            os.remove(temporary)
            continue

        shutil.move(temporary, destination)
        print(f"完成：{destination}（{written/1024/1024:.1f} MB）")
        return 0

    print("所有源都失败，请手动下载：")
    print(f"  {SOURCES[0]}")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
