#!/bin/bash

# 確認參數是否正確
if [ $# -ne 1 ]; then
    echo "用法: $0 <輸入文字檔>"
    exit 1
fi

# 獲取輸入檔案路徑
input_file="$1"

# 檢查檔案是否存在
if [ ! -f "$input_file" ]; then
    echo "錯誤: 檔案 '$input_file' 不存在"
    exit 1
fi

# 獲取檔案目錄和檔案名稱（不含副檔名）
dir_name=$(dirname "$input_file")
base_name=$(basename "$input_file" | cut -d. -f1)
output_file="${dir_name}/${base_name}.xlsx"

# 檢查 Python 是否已安裝
if ! command -v python3 &> /dev/null; then
    echo "錯誤: 需要 Python 3 但未安裝"
    echo "請安裝 Python 3 後再試"
    exit 1
fi

# 建立臨時 Python 檔案來檢查和安裝必要的模組
check_modules_py=$(mktemp)

cat > "$check_modules_py" << 'EOL'
import sys
import subprocess
import importlib.util

def check_and_install(package):
    spec = importlib.util.find_spec(package)
    if spec is None:
        print(f"正在安裝必要的 Python 模組: {package}")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        print(f"{package} 安裝完成")
    else:
        print(f"{package} 已安裝")

if __name__ == "__main__":
    check_and_install("openpyxl")
EOL

# 執行模組檢查和安裝
python3 "$check_modules_py"

# 移除臨時檢查模組的 Python 檔案
rm "$check_modules_py"

# 建立臨時 Python 檔案
temp_py_file=$(mktemp)

# 寫入 Python 腳本內容
cat > "$temp_py_file" << 'EOL'
import sys
import os
import re
from openpyxl import Workbook
from openpyxl.styles import Alignment

def process_qa_file(input_file, output_file):
    # 讀取整個檔案內容
    with open(input_file, 'r', encoding='utf-8') as file:
        content = file.read()
    
    # 使用正則表達式匹配 Q: 和 A: 的模式
    qa_pattern = r'(Q:.+?)\n(A:.+?)(?=\nQ:|$)'
    qa_matches = re.findall(qa_pattern, content, re.DOTALL)
    
    # 創建一個新的工作簿
    wb = Workbook()
    ws = wb.active
    ws.title = "QA Data"
    
    # 添加標題行
    ws['A1'] = "問答對 (Q&A Pair)"
    
    # 將每個 QA 對寫入到工作表中
    for i, (question, answer) in enumerate(qa_matches, 1):
        # 組合問題和回答為一個完整的文本
        qa_text = f"{question.strip()}\n{answer.strip()}"
        
        # 寫入資料到儲存格
        cell = ws.cell(row=i+1, column=1, value=qa_text)
        
        # 設置儲存格的格式，支援換行
        cell.alignment = Alignment(wrap_text=True, vertical='top')
    
    # 調整欄寬
    ws.column_dimensions['A'].width = 100
    
    # 調整列高以適應內容
    for i in range(2, len(qa_matches) + 2):
        # 根據內容設置適當的列高
        text_length = len(str(ws.cell(row=i, column=1).value))
        # 估算需要的列高 (每80個字元假設需要一行，每行約15點高度)
        estimated_lines = text_length / 80
        row_height = max(15, min(estimated_lines * 15, 409))  # 最小15，最大409
        ws.row_dimensions[i].height = row_height
    
    # 儲存工作簿
    wb.save(output_file)
    print(f"已成功將 QA 資料轉換並儲存到: {output_file}")
    print(f"共處理了 {len(qa_matches)} 組問答對")

# 主程式
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(f"用法: {sys.argv[0]} <輸入文字檔> <輸出 XLSX 檔案>")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    
    process_qa_file(input_file, output_file)
EOL

# 執行 Python 腳本
python3 "$temp_py_file" "$input_file" "$output_file"

# 移除臨時 Python 檔案
rm "$temp_py_file"

echo "處理完成。輸出檔案: $output_file"