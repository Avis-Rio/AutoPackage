#!/bin/bash
# 配分表自动转换工具 - 启动脚本

echo "======================================"
echo "  配分表自动转换工具"
echo "======================================"
echo ""

# 检查Python
if ! command -v python3 &> /dev/null
then
    echo "❌ 错误：未找到 Python 3"
    echo "请安装 Python 3.7 或更高版本"
    exit 1
fi

echo "✅ Python 版本："
python3 --version
echo ""

# 检查依赖
echo "检查依赖包..."
python3 -c "import xlrd, xlwt, xlutils" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "⚠️  缺少依赖包，正在安装..."
    python3 -m pip install -r requirements.txt
    if [ $? -ne 0 ]; then
        echo "❌ 依赖安装失败"
        exit 1
    fi
    echo "✅ 依赖安装完成"
else
    echo "✅ 依赖包已就绪"
fi

echo ""
echo "启动程序..."
echo "======================================"
echo ""

# 运行主程序
python3 main.py
