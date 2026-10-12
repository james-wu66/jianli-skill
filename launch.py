# -*- coding: utf-8 -*-
"""
一键启动器：双击 启动.bat 会跑这个。
职责：检查依赖 -> 启动 Flask -> 等服务就绪 -> 自动打开浏览器 -> 阻塞到服务退出。
"""
import glob
import os
import socket
import subprocess
import sys
import time
import webbrowser

HERE = os.path.dirname(os.path.abspath(__file__))
PORT = 5050

BANNER = """
================================================
  简历定制器
  地址：http://127.0.0.1:{port}
  服务就绪后会自动打开浏览器
  关闭本窗口 = 停止服务
================================================
"""

MIRROR = "https://pypi.tuna.tsinghua.edu.cn/simple"

# 找 Python 的候选：PATH 里的命令 + 各台电脑都一样的常见安装位置。
# 机器专属路径（绿色版 / 某个工具自带的 Python）不要写在这里，
# 放项目根目录的 python_local.txt（一行一个绝对路径）——那个文件不入库，
# 否则仓库里就带上了你的 Windows 用户名。
_CANDIDATES = [
    "python", "python3", "py",
    r"C:\Python313\python.exe",
    r"C:\Python312\python.exe",
    r"C:\Program Files\Python313\python.exe",
    r"C:\Program Files\Python312\python.exe",
    r"C:\ProgramData\anaconda3\python.exe",
]

LOCAL_PY_FILE = "python_local.txt"

_DEP_CHECK = "import flask,docx,requests"


def _local_candidates():
    """用户在 python_local.txt 里自己填的本机 Python 路径。"""
    path = os.path.join(HERE, LOCAL_PY_FILE)
    try:
        with open(path, encoding="utf-8-sig", errors="replace") as f:
            return [ln.strip().strip('"') for ln in f if ln.strip()]
    except OSError:
        return []


def _extra_candidates():
    """用环境变量拼常见位置，这样不会把某个具体用户名写进代码。"""
    out = []
    for env in ("LOCALAPPDATA", "USERPROFILE"):
        base = os.environ.get(env)
        if not base:
            continue
        for pat in (os.path.join("Programs", "Python", "*", "python.exe"),
                    "*conda*\\python.exe",
                    os.path.join("*conda*", "python.exe")):
            out.extend(sorted(glob.glob(os.path.join(base, pat))))
    return out


def _can_run(py):
    try:
        return subprocess.run([py, "-c", "print(1)"],
                              stdout=subprocess.DEVNULL,
                              stderr=subprocess.DEVNULL,
                              timeout=20).returncode == 0
    except Exception:
        return False


def _has_deps(py):
    try:
        return subprocess.run([py, "-c", _DEP_CHECK],
                              stdout=subprocess.DEVNULL,
                              stderr=subprocess.DEVNULL,
                              timeout=30).returncode == 0
    except Exception:
        return False


def find_python():
    """
    挑一个能用的 Python：优先挑依赖齐全的。
    返回 (路径, 缺依赖的解释器 or None)
      - 找到依赖齐全的 → (路径, None)
      - 只能跑但缺包   → (路径, 该路径)
      - 一个都跑不起来 → (None, None)
    """
    tried = [sys.executable] + _local_candidates() + _CANDIDATES + _extra_candidates()
    seen = set()
    fallback = None
    for cand in tried:
        if not cand or cand.lower() in seen:
            continue
        seen.add(cand.lower())
        if not _can_run(cand):
            continue
        if _has_deps(cand):
            return cand, None
        if fallback is None:
            fallback = cand
    return fallback, fallback


def check_deps(py):
    """返回 py 上缺失的模块列表。"""
    missing = []
    for mod in ("flask", "docx", "requests"):
        try:
            subprocess.run([py, "-c", "import " + mod],
                           stdout=subprocess.DEVNULL,
                           stderr=subprocess.DEVNULL,
                           timeout=30, check=True)
        except Exception:
            missing.append(mod)
    return missing


def port_open(port):
    s = None
    try:
        s = socket.socket()
        s.settimeout(0.3)
        s.connect(("127.0.0.1", port))
        return True
    except Exception:
        return False
    finally:
        if s:
            s.close()


def service_alive():
    """5050 上是否已经有"我们的"服务在跑（认首页标题，防止被别的程序误认）。
    防的是双实例抢端口（9/14 一天复发两次）：Windows 下两个 Flask 都能同时
    bind 5050（SO_REUSEADDR，第二个不报错），请求在两实例间被抢占 → 间歇性失败。
    所以起服务前必须先探，活着就绝不起新的。"""
    try:
        import urllib.request
        with urllib.request.urlopen(f"http://127.0.0.1:{PORT}/", timeout=2) as r:
            body = r.read(4096).decode("utf-8", "ignore")
            return "简历定制器" in body
    except Exception:
        return False


def auto_open_browser():
    """
    只有从 启动.bat 进来才自动开浏览器，否则会和 .bat 里的 start 命令
    各开一个窗口（两个浏览器标签）。
    判断依据：双击 .bat 时 Python 的 stdout 被重定向成管道（不是真终端）。
    """
    try:
        import stat as _st
        return not _st.S_ISCHR(os.fstat(1).st_mode)
    except Exception:
        return True


DEP_PKGS = {"flask": "flask", "docx": "python-docx", "requests": "requests"}


def _pip_cmd(py, missing):
    cmd = [py, "-m", "pip", "install", "--disable-pip-version-check"]
    cmd += [DEP_PKGS[m] for m in missing]
    return cmd + ["-i", MIRROR, "--trusted-host", "pypi.tuna.tsinghua.edu.cn"]


def install_deps(py, missing, yes=False):
    """缺依赖时当场装上，装完继续启动。

    yes=False 时问一句；拿不到输入（非交互）就默认装 —— 双击这个入口的人
    本来就不该被要求手抄 pip 命令。
    """
    print("\n[缺少依赖] " + "、".join(missing))
    if not yes:
        try:
            ans = input("现在自动安装吗？（直接回车 = 装，输入 n 回车 = 跳过）").strip().lower()
        except EOFError:
            print("（没拿到输入，按装处理）")
            ans = ""
        if ans.startswith("n"):
            print("已跳过。需要时手动执行：\n  " + " ".join(_pip_cmd(py, missing)))
            return False
    print("开始安装，可能要一两分钟，请保持这个窗口开着…\n")
    try:
        rc = subprocess.call(_pip_cmd(py, missing))
    except Exception as e:
        print(f"\n安装命令没能执行：{type(e).__name__}: {e}")
        return False
    if rc != 0:
        print(f"\npip 退出码 {rc}，安装没成功。上面是它的原始报错。")
        print("也可以自己复制这行去执行：\n  " + " ".join(_pip_cmd(py, missing)))
        return False
    still = check_deps(py)
    if still:
        print("\n装完了，但这几个模块仍然导入不了：" + "、".join(still))
        print("多半是这台机器上有多个 Python，包装到了另一个里。")
        return False
    print("\n依赖装好了。")
    return True


def main():
    os.chdir(HERE)
    print(BANNER.format(port=PORT))

    # 防双实例：服务已经活着就绝不再起第二个，直接开页面收工。
    # （要在找 Python / 查依赖之前做 —— 这一步只靠标准库，是最快路径。）
    if service_alive():
        print("\n检测到服务已经在运行，本次不重复启动，直接打开页面。")
        print("（这个窗口可以直接关掉，已在运行的服务不受影响）")
        if auto_open_browser():
            try:
                webbrowser.open(f"http://127.0.0.1:{PORT}")
            except Exception:
                pass
        return

    py, broken = find_python()
    if not py:
        print("\n[没找到能用的 Python]")
        print("请先安装 Python 3.10 以上版本：https://www.python.org/downloads/")
        input("按回车退出")
        sys.exit(1)

    print(f"使用 Python：{py}")

    missing = check_deps(py)
    if missing and not install_deps(py, missing):
        print("\n依赖没装上，服务起不来。装好之后再双击一次 启动.bat。")
        input("按回车退出")
        sys.exit(1)

    try:
        proc = subprocess.Popen([py, os.path.join(HERE, "app.py")], cwd=HERE)
    except Exception as e:
        print(f"\n启动失败：{type(e).__name__}: {e}")
        input("按回车退出")
        sys.exit(1)

    # 等服务真正就绪再开浏览器，避免用户打开一个连不上的页面
    for _ in range(60):
        if port_open(PORT):
            break
        if proc.poll() is not None:
            print(f"\n服务启动失败，退出码 {proc.returncode}")
            print("常见原因：端口被占用，或代码有语法错误（可先运行 python app.py 看报错）")
            input("按回车退出")
            sys.exit(1)
        time.sleep(0.5)
    else:
        print("\n等待服务超时（30 秒），浏览器先试着打开…")

    if auto_open_browser():
        try:
            webbrowser.open(f"http://127.0.0.1:{PORT}")
        except Exception:
            print(f"\n自动打开浏览器失败，请手动访问 http://127.0.0.1:{PORT}")
    else:
        print(f"\n服务就绪，访问 http://127.0.0.1:{PORT} 使用（如需自动开浏览器：{PORT} 已就绪）")

    try:
        proc.wait()
    except KeyboardInterrupt:
        print("\n\n正在停止服务…")
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except Exception:
            proc.kill()


if __name__ == "__main__":
    main()
