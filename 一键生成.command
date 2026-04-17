#!/bin/bash
# 双击此文件会提示你选一个放月结单 PDF 的文件夹；
# 也可以把文件夹直接拖到终端里跑。

set -e
cd "$(dirname "$0")"

# 依赖检查
python3 -c "import pdfplumber, openpyxl" 2>/dev/null || {
    echo "首次使用，正在安装依赖…"
    pip3 install pdfplumber openpyxl
}

# 文件夹：命令行参数 > osascript 图形选择
if [ -n "$1" ] && [ -d "$1" ]; then
    DIR="$1"
else
    DIR=$(osascript -e 'tell app "System Events" to POSIX path of (choose folder with prompt "选择装有月结单 PDF 的文件夹")' 2>/dev/null || true)
fi

if [ -z "$DIR" ] || [ ! -d "$DIR" ]; then
    echo "未选择有效目录，退出。"
    exit 1
fi

python3 bank_statement_analyzer.py "$DIR"

echo
echo "完成。按任意键退出…"
read -n 1 -s
