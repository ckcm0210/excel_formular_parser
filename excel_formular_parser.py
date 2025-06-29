import openpyxl
import formulas
import os
import re
from urllib.parse import unquote
import json
from openpyxl.worksheet.formula import ArrayFormula

# --- 登錄資訊 ---
print("--- 最終三重視角報告 (V20) ---")
print("Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): 2025-06-29 16:26:32")
print("Current User's Login: ckcm0210")
print("-" * 60)
print("目標：在 V19 的基礎上，重新加入『程式讀取公式』的顯示，形成三重視角報告。")
print("-" * 60)

# --- 設定檔案路徑與目標 ---
working_path = r"C:\Users\user\Desktop\pytest\Formula Difference Analyzer"
file_c_path = os.path.join(working_path, "File_C_Manufacturing.xlsx")
target_sheet_name = 'Cost_Analysis' 
target_cell_address = 'C3'
# ----------------------------------------

try:
    # =================================================================
    #  Part 1: 外部視角 (用戶所見) - 保持不變
    # =================================================================
    print("--- Part 1: 外部視角 (用戶所見) ---")
    print("1a. [openpyxl] 正在提取並重構權威的『用戶可見公式』...")

    workbook_openpyxl = openpyxl.load_workbook(filename=file_c_path, data_only=False)
    worksheet_openpyxl = workbook_openpyxl[target_sheet_name]
    cell_content = worksheet_openpyxl[target_cell_address].value
    
    raw_formula_with_indices = str(cell_content.text) if isinstance(cell_content, ArrayFormula) else str(cell_content)

    reconstructed_formula = raw_formula_with_indices
    if workbook_openpyxl._external_links:
        for i, link in enumerate(workbook_openpyxl._external_links):
            placeholder = f"[{i+1}]"
            full_external_path = os.path.abspath(os.path.join(working_path, os.path.basename(link.path)))
            replacement_text = f"'{full_external_path}'"
            reconstructed_formula = reconstructed_formula.replace(placeholder, replacement_text)

    print("\n--- 🏆 1b. 權威的『用戶可見公式』(已重構) 🏆 ---")
    print("=" * 45)
    print(reconstructed_formula)
    print("=" * 45)

    # =================================================================
    #  Part 2: 內部視角 (程式所見) - 增加內容
    # =================================================================
    print("\n\n--- Part 2: 內部視角 (程式所見) ---")
    print("2a. [formulas] 正在載入模型以進行深度剖析...")
    
    excel_model = formulas.ExcelModel().load(file_c_path)
    target_cell_key = f"'[{os.path.basename(file_c_path)}]COST_ANALYSIS'!{target_cell_address}"
    
    # --- 【全新增補】顯示 formulas 讀取的原始公式 ---
    program_read_formula = excel_model.to_dict()[target_cell_key]
    print("\n--- 2b. 程式讀取公式 (Formulas Library View) ---")
    print(program_read_formula)
    print("------------------------------------------------")
    # --- 增補結束 ---

    print("\n--- 2c. 依賴項剖析報告 ---")
    compiled_cell_object = excel_model.cells[target_cell_key]
    raw_references = []
    
    if hasattr(compiled_cell_object, 'inputs') and compiled_cell_object.inputs:
        raw_references = list(compiled_cell_object.inputs.keys())
        print(f"在公式中發現 {len(raw_references)} 個原始內部依賴項。")
    else:
        print("未在公式中發現任何內部依賴項。")

    # =================================================================
    #  Part 3: 可用零件拆解 (參照路徑正規化) - 保持不變
    # =================================================================
    print("\n\n--- Part 3: 可用零件拆解 (參照路徑正規化) ---")
    
    normalized_parts = []
    if raw_references:
        print("3a. 正在啟動『參照路徑正規化處理器』...")
        ref_pattern = re.compile(r"'(.*)\[(.*?)\](.*?)'!(.*)")

        for ref in raw_references:
            print(f"\n   處理中: {ref}")
            match = ref_pattern.match(ref)
            
            if match:
                relative_path_part, filename_part, sheetname_part, cell_address_part = match.groups()
                
                full_relative_path = os.path.join(relative_path_part, filename_part)
                decoded_path = unquote(full_relative_path)
                absolute_path = os.path.abspath(os.path.join(os.path.dirname(file_c_path), decoded_path))
                
                part = { "absolute_path": absolute_path, "sheet_name": sheetname_part, "cell_address": cell_address_part, "original_reference": ref }
                normalized_parts.append(part)
                
                print(f"     ✅ 拆解成功!")
                print(f"        -> 絕對路徑: {part['absolute_path']}")
                print(f"        -> 工作表名: {part['sheet_name']}")
                print(f"        -> 單元格: {part['cell_address']}")
            else:
                print(f"     ⚠️  注意: 參照 '{ref}' 為內部參照，非外部檔案。")
                part = { "absolute_path": file_c_path, "sheet_name": ref.split('!')[0].strip("'"), "cell_address": ref.split('!')[1], "original_reference": ref }
                normalized_parts.append(part)

        print("\n\n--- 🏆 3b. 『可用零件』清單 (已正規化) 🏆 ---")
        print("=" * 50)
        print(json.dumps(normalized_parts, indent=2, ensure_ascii=False))
        print("=" * 50)
    
    print(f"\n\n✅✅✅ **最終報告完成！** ✅✅✅")

except Exception as e:
    print(f"\n❌ 發生未知錯誤：{type(e).__name__} - {e}")
    import traceback
    traceback.print_exc()
