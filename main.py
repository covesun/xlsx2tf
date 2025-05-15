import hcl2
import openpyxl
from pathlib import Path

# ===== パス設定 =====
tf_path = Path("main.tf")
input_excel = Path("param_sheet.xlsx")
output_excel = Path("output/param_sheet_reflected.xlsx")
output_excel.parent.mkdir(parents=True, exist_ok=True)

# ===== ネスト＋配列対応 flatten関数 =====
def flatten_split_keys(value, parent_key=""):
    items = {}
    if isinstance(value, dict):
        for k, v in value.items():
            new_key = f"{parent_key}.{k}" if parent_key else k
            items.update(flatten_split_keys(v, new_key))
    elif isinstance(value, list):
        for i, elem in enumerate(value):
            new_key = f"{parent_key}[{i}]" if parent_key else f"[{i}]"
            items.update(flatten_split_keys(elem, new_key))
    else:
        items[parent_key] = value
    return items

# ===== Terraform読み込みと辞書化（resource名, 親キー, 子キー）=====
terraform_data = {}

with tf_path.open("r") as tf_file:
    parsed = hcl2.load(tf_file)

for block in parsed.get("resource", []):
    for resource_type, resource_instances in block.items():
        for res_name, res_body in resource_instances.items():
            flat_attrs = flatten_split_keys(res_body)
            for full_key, val in flat_attrs.items():
                if "." in full_key:
                    parent_key, child_key = full_key.rsplit(".", 1)
                else:
                    parent_key, child_key = "", full_key
                terraform_data[(res_name, parent_key, child_key)] = val

# ===== Excel反映処理（A:resource, B:親キー, C:子キー, D:値）=====
wb = openpyxl.load_workbook(input_excel)
ws = wb.active

for row in ws.iter_rows(min_row=2):
    res_cell, parent_cell, child_cell, val_cell = row[:4]
    key = (
        str(res_cell.value).strip(),
        str(parent_cell.value).strip(),
        str(child_cell.value).strip()
    )
    if key in terraform_data:
        val_cell.value = terraform_data[key]

wb.save(output_excel)
print(f"反映完了：{output_excel}")
