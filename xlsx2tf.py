import sys
import re
import openpyxl
import hcl2
from pathlib import Path

# ===== flatten_split_keys =====
def flatten_split_keys(value, parent_key=""):
    """
    任意のネストdict/listを 'foo.bar[0].baz' のような連結キー:値 のdictに展開する。
    """
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

# ===== export_hcl_to_excel =====
def export_hcl_to_excel(tf_path, output_excel):
    """
    Terraform HCLファイルをパースし、リソース属性情報をExcelファイルへ出力する。
    - resource_type/res_name/parent_key/child_key/値/連結キー の形式で書き出し
    """
    with Path(tf_path).open("r") as tf_file:
        parsed = hcl2.load(tf_file)
    terraform_data = {}
    for block in parsed.get("resource", []):
        for resource_type, resource_instances in block.items():
            for res_name, res_body in resource_instances.items():
                flat_attrs = flatten_split_keys(res_body)
                for full_key, val in flat_attrs.items():
                    if "." in full_key:
                        parent_key, child_key = full_key.rsplit(".", 1)
                    else:
                        parent_key, child_key = "", full_key
                    terraform_data[(resource_type, res_name, parent_key, child_key)] = val

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "tf_key_list"
    ws.cell(row=1, column=1, value="resource_type")
    ws.cell(row=1, column=2, value="res_name")
    ws.cell(row=1, column=3, value="parent_key")
    ws.cell(row=1, column=4, value="child_key")
    ws.cell(row=1, column=5, value="値")
    ws.cell(row=1, column=6, value="連結キー")

    for idx, ((resource_type, res_name, parent_key, child_key), value) in enumerate(terraform_data.items(), start=2):
        concat_key = ".".join([x for x in [res_name, parent_key, child_key] if x])
        ws.cell(row=idx, column=1, value=resource_type)
        ws.cell(row=idx, column=2, value=res_name)
        ws.cell(row=idx, column=3, value=parent_key)
        ws.cell(row=idx, column=4, value=child_key)
        ws.cell(row=idx, column=5, value=str(value))
        ws.cell(row=idx, column=6, value=concat_key)

    wb.save(output_excel)
    print(f"出力完了：{output_excel}")

# ===== set_nested_dict_from_concat_key =====
def set_nested_dict_from_concat_key(data, keys, value):
    """
    連結キー 'foo.bar[0].baz' 形式のリストを元に、dict/list構造を再帰的に構築して値をセットする。
    """
    key = keys[0]
    m = re.match(r'(\w+)\[(\d+)\]', key)
    if m:
        k, idx = m.group(1), int(m.group(2))
        if k not in data:
            data[k] = []
        while len(data[k]) <= idx:
            data[k].append({})
        if len(keys) == 1:
            data[k][idx] = value
        else:
            if not isinstance(data[k][idx], dict):
                data[k][idx] = {}
            set_nested_dict_from_concat_key(data[k][idx], keys[1:], value)
    else:
        if len(keys) == 1:
            data[key] = value
        else:
            if key not in data or not isinstance(data[key], dict):
                data[key] = {}
            set_nested_dict_from_concat_key(data[key], keys[1:], value)

# ===== format_hcl_value =====
def format_hcl_value(name, val, indent, eqpad="", is_map=False):
    """
    HCLの値をインデント付きのHCL文形式に整形して返す。
    - 文字列、数値、bool、null、式（${...}）などに対応
    """
    ind = '  ' * indent
    eqpad = eqpad or ''
    if isinstance(val, str):
        m = re.fullmatch(r"\$\{([^}]+)\}", val)
        if m:
            expr = m.group(1)
            return f"{ind}{name}{eqpad} = {expr}\n"
        else:
            return f"{ind}{name}{eqpad} = \"{val}\"\n"
    elif isinstance(val, bool):
        return f"{ind}{name}{eqpad} = {'true' if val else 'false'}\n"
    elif val is None:
        return f"{ind}{name}{eqpad} = null\n"
    else:
        return f"{ind}{name}{eqpad} = {val}\n"

# ===== dict_to_hcl_block =====
def dict_to_hcl_block(name, val, indent=0):
    """
    ネストdict/listをHCLブロックとして再帰的に出力（module, map, 配列, blockすべて対応）
    """
    ind = '  ' * indent
    if isinstance(val, dict):
        # 完全なmap（ネストなし）は "name = { ... }" 形式
        if all(not isinstance(v, (dict, list)) for v in val.values()):
            keys = list(val.keys())
            maxlen = max(len(k) for k in keys) if keys else 0
            hcl = f"{ind}{name} = {{\n"
            for k in keys:
                v = val[k]
                eqpad = ' ' * (maxlen - len(k))
                hcl += format_hcl_value(k, v, indent+1, eqpad, is_map=True)
            hcl += f"{ind}}}\n"
            return hcl
        # 通常ブロック（ネストあり）
        kvs = []
        blocks = []
        for k, v in val.items():
            if isinstance(v, dict) or isinstance(v, list):
                blocks.append((k, v))
            else:
                kvs.append((k, v))
        maxlen = max((len(k) for k, _ in kvs), default=0)
        hcl = f"{ind}{name} {{\n"
        for k, v in kvs:
            eqpad = ' ' * (maxlen - len(k))
            hcl += format_hcl_value(k, v, indent+1, eqpad)
        for k, v in blocks:
            hcl += dict_to_hcl_block(k, v, indent+1)
        hcl += f"{ind}}}\n"
        return hcl
    elif isinstance(val, list):
        # listがプリミティブなら配列、dictならblock
        if all(isinstance(i, str) for i in val):
            arr = ', '.join([f"\"{x}\"" for x in val])
            return f"{ind}{name} = [{arr}]\n"
        else:
            blocks = ""
            for v in val:
                blocks += dict_to_hcl_block(name, v, indent)
            return blocks
    else:
        return format_hcl_value(name, val, indent)

# ===== dict_to_resource_hcl =====
def dict_to_resource_hcl(d):
    """
    最上位dict（リソースタイプ別dict）をHCLファイル形式のテキストに変換する。
    """
    hcl = ""
    for res_type, res_objs in d.items():
        for res_name, content in res_objs.items():
            hcl += f'resource "{res_type}" "{res_name}" {{\n'
            kvs = []
            blocks = []
            for k, v in content.items():
                if isinstance(v, dict) or isinstance(v, list):
                    blocks.append((k, v))
                else:
                    kvs.append((k, v))
            maxlen = max((len(k) for k, _ in kvs), default=0)
            for k, v in kvs:
                eqpad = ' ' * (maxlen - len(k))
                hcl += format_hcl_value(k, v, 1, eqpad)
            for k, v in blocks:
                hcl += dict_to_hcl_block(k, v, 1)
            hcl += '}\n\n'
    return hcl

# ===== find_header_row =====
def find_header_row(ws, concat_key_header="連結キー", value_header="設定値"):
    """
    指定したシートで '連結キー' '値' が並ぶヘッダー行を探し、その行番号・該当列indexを返す。
    """
    for idx, row in enumerate(ws.iter_rows(values_only=True), 1):
        if concat_key_header in row and value_header in row:
            key_col = row.index(concat_key_header)
            val_col = row.index(value_header)
            return idx, key_col, val_col
    return None

# ===== read_excel_concatkey_to_dict =====
def read_excel_concatkey_to_dict(input_excel, concat_key_header="連結キー", value_header="設定値"):
    """
    Excel全シートを走査し、見出し行を自動特定して連結キー:値 のdictを集める。
    """
    wb = openpyxl.load_workbook(input_excel, data_only=True)
    params = {}
    for ws in wb.worksheets:
        res = find_header_row(ws, concat_key_header, value_header)
        if not res:
            continue
        header_row_idx, key_col, val_col = res
        for row in ws.iter_rows(min_row=header_row_idx+1, values_only=True):
            concat_key = row[key_col]
            value = row[val_col]
            if concat_key:
                params[concat_key] = value
    return params

# ===== export_excel_to_hcl =====
def export_excel_to_hcl(input_excel, output_hcl):
    """
    Excelパラシートから連結キー＆値をdict化して、HCLファイルとして書き出す。
    """
    params = read_excel_concatkey_to_dict(input_excel)
    data = {}
    for k, v in params.items():
        set_nested_dict_from_concat_key(data, k.split("."), v)
    hcl = dict_to_resource_hcl(data)
    with open(output_hcl, "w", encoding="utf-8") as f:
        f.write(hcl)
    print(f"HCL出力完了：{output_hcl}")

# ===== reflect_tf_to_excel =====
def reflect_tf_to_excel(tf_path, input_excel, output_excel):
    """
    Terraform HCLファイルの値を、Excel（全シート）の対応行に反映し、保存する。
    （リソース・親キー・子キーを突合。値のみ上書き）
    """
    with Path(tf_path).open("r") as tf_file:
        parsed = hcl2.load(tf_file)
    terraform_data = {}
    for block in parsed.get("resource", []):
        for resource_type, resource_instances in block.items():
            for res_name, res_body in resource_instances.items():
                flat_attrs = flatten_split_keys(res_body)
                for full_key, val in flat_attrs.items():
                    if "." in full_key:
                        parent_key, child_key = full_key.rsplit(".", 1)
                    else:
                        parent_key, child_key = "", full_key
                    terraform_data[(resource_type, res_name, parent_key, child_key)] = val
    wb = openpyxl.load_workbook(input_excel)
    for ws in wb.worksheets:
        res = find_header_row(ws, "resource_type", "値")  # resource_type, 値で判定
        if not res:
            continue
        header_row_idx, res_type_col, val_col = res
        res_name_col = res_type_col + 1
        parent_col = res_type_col + 2
        child_col = res_type_col + 3
        for row in ws.iter_rows(min_row=header_row_idx+1):
            key = (
                str(row[res_type_col].value).strip() if row[res_type_col].value else "",
                str(row[res_name_col].value).strip() if row[res_name_col].value else "",
                str(row[parent_col].value).strip() if row[parent_col].value else "",
                str(row[child_col].value).strip() if row[child_col].value else ""
            )
            if key in terraform_data:
                row[val_col].value = terraform_data[key]
    wb.save(output_excel)
    print(f"Excel反映完了：{output_excel}")

# ===== usage =====
def usage():
    """
    各コマンドの使い方を表示し、異常終了する。
    """
    print("使い方:")
    print(" python xlsx2tf.py export_hcl_to_excel <tf_path> <output_excel>")
    print(" python xlsx2tf.py export_excel_to_hcl <input_excel> <output_hcl>")
    print(" python xlsx2tf.py reflect_tf_to_excel <tf_path> <input_excel> <output_excel>")
    sys.exit(1)

# ===== CLIディスパッチ =====
funcs = {
    "export_hcl_to_excel": export_hcl_to_excel,
    "export_excel_to_hcl": export_excel_to_hcl,
    "reflect_tf_to_excel": reflect_tf_to_excel,
}

if __name__ == "__main__":
    if len(sys.argv) < 2:
        usage()
    func_name = sys.argv[1]
    args = sys.argv[2:]
    if func_name in funcs:
        try:
            funcs[func_name](*args)
        except Exception as e:
            print(f"エラー: {e}")
            usage()
    else:
        usage()
