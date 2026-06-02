import pdfplumber
import sys

pdf_path = r"c:\Users\aa\Desktop\造价清单文件标准化\施工成本管理\施工成本管理.pdf"

try:
    with pdfplumber.open(pdf_path) as pdf:
        print(f"总页数: {len(pdf.pages)}")
        print("\n" + "="*80)

        full_text = []

        for i, page in enumerate(pdf.pages):
            print(f"\n--- 第 {i+1} 页 ---\n")
            text = page.extract_text()

            if text:
                full_text.append(text)
                print(text)
            else:
                print("[此页无法提取文本]")

            tables = page.extract_tables()
            if tables:
                print(f"\n[本页包含 {len(tables)} 个表格]")
                for j, table in enumerate(tables):
                    print(f"\n表格 {j+1}:")
                    for row in table:
                        print(" | ".join(str(cell) if cell else "" for cell in row))

            print("\n" + "="*80)

except Exception as e:
    print(f"读取PDF时出错: {e}", file=sys.stderr)
    import traceback
    traceback.print_exc()
