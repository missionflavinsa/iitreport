"""
Analyze the Word template structure to understand exact cell positions
"""
from docx import Document
import os

template_path = r'c:\AusKar\IIT Report\Report Card Format updated.docx'
doc = Document(template_path)

print(f"Number of tables: {len(doc.tables)}")
print("=" * 80)

for table_idx, table in enumerate(doc.tables):
    print(f"\n{'='*80}")
    print(f"TABLE {table_idx}: {len(table.rows)} rows x {len(table.columns)} columns")
    print(f"{'='*80}")
    
    for row_idx, row in enumerate(table.rows):
        print(f"\n--- Row {row_idx} ---")
        for col_idx, cell in enumerate(row.cells):
            text = cell.text.strip()[:50] if cell.text.strip() else "(empty)"
            has_runs = len(cell.paragraphs[0].runs) > 0 if cell.paragraphs else False
            print(f"  Col {col_idx}: [{text}] runs={has_runs}")
