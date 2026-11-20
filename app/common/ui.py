"""控制台交互：菜单、输入校验、进度提示。

单独拆出来是为了让启动.py 保持精简，也方便以后加平台。
"""

import sys


def setup_console():
    """让 Windows 控制台正确显示中文。"""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


def line(char="=", width=72):
    print(char * width)


def title(text):
    print()
    line()
    print(text)
    line()


def ask(prompt, default="", allow_empty=True):
    """读一行输入，带默认值。"""
    suffix = f"（默认 {default}）" if default not in ("", None) else ""
    try:
        value = input(f"{prompt}{suffix}：").strip()
    except EOFError:
        return default

    # 清掉可能的代理项（终端编码不对时会出现，写文件会报错）
    value = value.encode("utf-8", "replace").decode("utf-8")

    if not value:
        return default
    return value


def ask_int(prompt, default, minimum=0, maximum=None):
    """读一个整数，越界或非法就重新问。"""
    while True:
        raw = ask(prompt, str(default))
        if raw == "":
            return default
        if raw.isascii() and raw.lstrip("-").isdigit():
            value = int(raw)
            if value < minimum:
                print(f"  不能小于 {minimum}，请重新输入")
                continue
            if maximum is not None and value > maximum:
                print(f"  不能大于 {maximum}，请重新输入")
                continue
            return value
        print("  请输入整数")


def ask_yes_no(prompt, default=False):
    hint = "Y/n" if default else "y/N"
    try:
        raw = input(f"{prompt} [{hint}]：").strip().lower()
    except EOFError:
        return default
    if not raw:
        return default
    return raw in ("y", "yes", "是", "1")


def choose(prompt, options, default_index=0):
    """options: [(key, label, description), ...]，返回选中的 key。"""
    print()
    print(prompt)
    print()
    for index, (key, label, description) in enumerate(options, 1):
        mark = " (当前默认)" if index - 1 == default_index else ""
        print(f"  {index}. {label}{mark}")
        if description:
            print(f"     {description}")
    print()

    while True:
        raw = ask("请输入序号", str(default_index + 1))
        if raw.isascii() and raw.isdigit():
            picked = int(raw)
            if 1 <= picked <= len(options):
                return options[picked - 1][0]
        print(f"  请输入 1~{len(options)}")


def pause(message="按回车返回菜单"):
    try:
        input(f"\n{message}...")
    except EOFError:
        pass
