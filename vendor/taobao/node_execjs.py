"""用本地 Node.js 执行 JS 签名，替代 PyExecJS。

为什么不用 PyExecJS：
  · PyPI 上的包名是 PyExecJS（不是 execjs），新手很容易装错
  · 它 2018 年后没再更新，在 Python 3.13+ 上有兼容风险
  · 而 node 本来就是京东平台的硬依赖，已经随包提供
  · 新版本 Python 往往还没有对应的二进制 wheel，能少一个依赖就少一个坑

对外接口和 PyExecJS 的最小可用子集保持一致：
    ctx = compile(js_source)
    ctx.call("函数名", 参数1, 参数2, ...)
"""

import json
import os
import shutil
import subprocess
import tempfile


def find_node():
    """返回 node 可执行文件路径；找不到返回空。"""
    found = shutil.which("node")
    if found:
        return found

    candidates = [
        os.path.expandvars(r"%ProgramFiles%\nodejs\node.exe"),
        os.path.expandvars(r"%ProgramFiles(x86)%\nodejs\node.exe"),
        os.path.expandvars(r"%LOCALAPPDATA%\Programs\nodejs\node.exe"),
        r"D:\Program Files\nodejs\node.exe",
    ]
    for path in candidates:
        if os.path.isfile(path):
            return path
    return ""


NODE_PATH = find_node()
NODE_AVAILABLE = bool(NODE_PATH)


class JsExecutionError(RuntimeError):
    pass


class Context:
    """一份已加载的 JS 源码，可以反复调用其中的函数。"""

    def __init__(self, source, timeout=60):
        if not NODE_AVAILABLE:
            raise JsExecutionError(
                "找不到 node 可执行文件。淘宝签名需要 Node.js。\n"
                "  处理：双击「安装环境.bat」，或手动安装 nodejs"
            )
        self.source = source
        self.timeout = timeout

    def call(self, name, *args):
        """调用 JS 里的函数并返回结果（支持字符串/数字/布尔/数组/对象）。"""
        arguments = json.dumps(list(args), ensure_ascii=False)

        # 用标记把真正的返回值圈出来 —— 被测脚本自己可能也有 console.log
        marker = "__EE_JS_RESULT__"

        # 把源码和调用语句拼在一起，这样 aef.js 里 var 声明的函数也能访问到
        runner = (
            self.source
            + "\n;(function () {\n"
            + f"  var MARKER = {marker!r};\n"
            + "  var fs = require('fs');\n"
            + "  var payload = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));\n"
            + f"  var result = ({name}).apply(null, payload);\n"
            + "  process.stdout.write(MARKER + JSON.stringify(result === undefined ? null : result));\n"
            + "})();\n"
        )

        workdir = tempfile.mkdtemp(prefix="ee_js_")
        try:
            script_path = os.path.join(workdir, "run.js")
            args_path = os.path.join(workdir, "args.json")

            with open(script_path, "w", encoding="utf-8") as handle:
                handle.write(runner)
            with open(args_path, "w", encoding="utf-8") as handle:
                handle.write(arguments)

            try:
                done = subprocess.run(
                    [NODE_PATH, script_path, args_path],
                    capture_output=True, timeout=self.timeout,
                )
            except subprocess.TimeoutExpired:
                raise JsExecutionError(
                    f"node 执行超时（{self.timeout} 秒）") from None
            except OSError as error:
                raise JsExecutionError(f"node 启动失败：{error}") from None

            if done.returncode != 0:
                stderr = done.stderr.decode("utf-8", errors="replace").strip()
                raise JsExecutionError(
                    f"node 返回码 {done.returncode}：{stderr[:400] or '（无错误输出）'}")

            output = done.stdout.decode("utf-8", errors="replace")

            # 只取标记之后的部分，前面的 console.log 噪音全部丢掉
            if marker not in output:
                return None
            output = output.rsplit(marker, 1)[-1].strip()
            if not output:
                return None

            try:
                return json.loads(output)
            except json.JSONDecodeError:
                # 有些脚本直接输出原始字符串
                return output
        finally:
            shutil.rmtree(workdir, ignore_errors=True)

    def eval(self, expression):
        return self.call("(function(){return (" + expression + ");})")


def compile(source, *args, **kwargs):
    """兼容 PyExecJS 的 compile(source)。"""
    return Context(source, **kwargs)
