import pandas as pd

# ====================== 你只需要改这里 ======================
input_file = r"C:\Users\Admin\Downloads\填充tos.xlsx"  # 改成你的.xlsx路径
output_file = r"C:\Users\Admin\Downloads\填充tos_output.xlsx"
# ==========================================================

# 读取所有sheet
excel_file = pd.ExcelFile(input_file)
all_sheets = excel_file.sheet_names

# 创建一个 writer，用于输出多Sheet Excel
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    for sheet_name in all_sheets:
        print(f"正在处理 Sheet：{sheet_name}")

        # 读取当前sheet
        df = pd.read_excel(input_file, sheet_name=sheet_name)

        # 清理列名空格
        df.columns = df.columns.str.strip()

        # 类型统一
        df["session_id"] = df["session_id"].astype(str).str.strip()
        df["round_id"] = pd.to_numeric(df["round_id"], errors="coerce")

        # 取每个 session round=0 的 tos_url 作为基准
        base_tos = (
            df[df["round_id"] == 0][["session_id", "tos_url"]]
            .drop_duplicates("session_id")
            .rename(columns={"tos_url": "base_tos"})
        )

        # 合并
        df = df.merge(base_tos, on="session_id", how="left")

        # 填充
        df["tos补全"] = df["tos_url"].fillna(df["base_tos"])

        # 删除临时列
        df = df.drop(columns=["base_tos"])

        # 写回新的 Excel，保持原来的Sheet名
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print("✅ 全部Sheet处理完成！文件已保存到：", output_file)

