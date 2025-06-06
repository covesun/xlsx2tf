import re

def set_nested(data, keys, value):
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
            set_nested(data[k][idx], keys[1:], value)
    else:
        if len(keys) == 1:
            data[key] = value
        else:
            if key not in data or not isinstance(data[key], dict):
                data[key] = {}
            set_nested(data[key], keys[1:], value)

params = {
    "azurerm_subnet.snet-conswkld-vnet-int.name": "snet-conswkld-vnet-int",
    "azurerm_subnet.snet-conswkld-vnet-int.resource_group_name": "${var.resource_group_name}",
    "azurerm_subnet.snet-conswkld-vnet-int.virtual_network_name": "${azurerm_virtual_network.vnet-container.name}",
    "azurerm_subnet.snet-conswkld-vnet-int.address_prefixes[0]": "10.4.8.32/27",
    "azurerm_subnet.snet-conswkld-vnet-int.tags.env1": "dev1",
    "azurerm_subnet.snet-conswkld-vnet-int.tags.env2": "dev2",
    "azurerm_subnet.snet-conswkld-vnet-int.delegation[0].name": "delegation",
    "azurerm_subnet.snet-conswkld-vnet-int.delegation[0].service_delegation[0].name": "Microsoft.Network/privateEndpoints",
    "azurerm_subnet.snet-conswkld-vnet-int.delegation[0].service_delegation[0].actions[0]": "Microsoft.Network/virtualNetworks/subnets/action",
    "azurerm_subnet.snet-pe.name": "snet-pe",
    "azurerm_subnet.snet-pe.resource_group_name": "${var.resource_group_name}",
    "azurerm_subnet.snet-pe.virtual_network_name": "${azurerm_virtual_network.vnet-container.name}",
    "azurerm_subnet.snet-pe.address_prefixes[0]": "10.4.16.0/29"
}

data = {}
for k, v in params.items():
    set_nested(data, k.split("."), v)

def dict_to_hcl_block(name, val, indent=0):
    ind = '  ' * indent
    if isinstance(val, dict):
        # tagsなどmap判定：全要素がプリミティブ
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
        # それ以外は通常ブロック
        kvs = []
        blocks = []
        for k, v in val.items():
            # ここでプリミティブリストもblocks側へ
            if isinstance(v, dict) or (isinstance(v, list)):
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
        # **必ずここを通す（プリミティブリストも）**
        if all(isinstance(i, str) for i in val):
            arr = ', '.join([f"\"{x}\"" for x in val])
            return f"{ind}{name} = [{arr}]\n"
        blocks = ""
        for v in val:
            blocks += dict_to_hcl_block(name, v, indent)
        return blocks
    else:
        return format_hcl_value(name, val, indent)


def format_hcl_value(name, val, indent, eqpad="", is_map=False):
    ind = '  ' * indent
    eqpad = eqpad or ''
    if isinstance(val, str):
        return f"{ind}{name}{eqpad} = \"{val}\"\n"
    elif isinstance(val, bool):
        return f"{ind}{name}{eqpad} = {'true' if val else 'false'}\n"
    elif val is None:
        return f"{ind}{name}{eqpad} = null\n"
    else:
        return f"{ind}{name}{eqpad} = {val}\n"

def dict_to_resource_hcl(d):
    hcl = ""
    for res_type, res_objs in d.items():
        for res_name, content in res_objs.items():
            hcl += f'resource "{res_type}" "{res_name}" {{\n'
            # 1階層目もプリミティブだけmaxlenで揃え
            kvs = []
            blocks = []
            for k in content.keys():
                v = content[k]
                if isinstance(v, (dict, list)) and not (isinstance(v, list) and all(isinstance(i, str) for i in v)):
                    blocks.append((k, v))
                else:
                    kvs.append((k, v))
            if kvs:
                maxlen = max(len(k) for k, _ in kvs)
            else:
                maxlen = 0
            for k, v in kvs:
                eqpad = ' ' * (maxlen - len(k))
                hcl += format_hcl_value(k, v, 1, eqpad)
            for k, v in blocks:
                hcl += dict_to_hcl_block(k, v, 1)
            hcl += '}\n\n'
    return hcl

# ---（set_nestedなど前処理はそのまま）---
# --- ここでtfファイルへ出力 ---
hcl_text = dict_to_resource_hcl(data)
with open("output.tf", "w", encoding="utf-8") as f:
    f.write(hcl_text)
print(dict_to_resource_hcl(data))
