#!/bin/bash
# macOS 双击启动（等价于 Windows 的 run.bat）
# 提示：如果双击提示「无法打开」，先在「终端」里执行一次：
#   chmod +x "$(dirname "$0")/启动.command"
cd "$(dirname "$0")" || exit 1
export PYTHONIOENCODING=utf-8

if ! command -v python3 >/dev/null 2>&1; then
    echo ""
    echo "[错误] 没找到 python3。请先装 Python 3.10 以上版本："
    echo "       https://www.python.org/downloads/macos/"
    echo ""
    read -r -p "按回车退出…"
    exit 1
fi

if ! python3 -c "import websocket" >/dev/null 2>&1; then
    echo "缺少依赖 websocket-client，正在安装 …"
    python3 -m pip install --user websocket-client || {
        echo ""
        echo "[错误] 依赖安装失败。手动执行一次： python3 -m pip install --user websocket-client"
        read -r -p "按回车退出…"
        exit 1
    }
fi

echo "启动 SHEIN 选品工具 …（这个窗口不要关，关掉工具就退出了）"
exec python3 ebay_tool.py
