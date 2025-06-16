' Microsoft Scripting Runtimeが必要（参照設定で「Microsoft Scripting Runtime」をチェック）

' dict: Scripting.Dictionary
' list: Collection
' Return: Scripting.Dictionary (Key: "foo.bar[0].baz", Value: actual value)
Function flatten_split_keys(value As Variant, Optional parent_key As String = "") As Scripting.Dictionary
    Dim items As Scripting.Dictionary
    Set items = New Scripting.Dictionary

    If TypeName(value) = "Dictionary" Then
        Dim k As Variant
        For Each k In value.Keys
            Dim v As Variant
            v = value(k)
            Dim new_key As String
            If parent_key <> "" Then
                new_key = parent_key & "." & k
            Else
                new_key = k
            End If
            Dim subItems As Scripting.Dictionary
            Set subItems = flatten_split_keys(v, new_key)
            Dim subK As Variant
            For Each subK In subItems.Keys
                items.Add subK, subItems(subK)
            Next
        Next
    ElseIf TypeName(value) = "Collection" Then
        Dim i As Long
        For i = 1 To value.Count
            Dim elem As Variant
            elem = value(i)
            Dim new_key As String
            If parent_key <> "" Then
                new_key = parent_key & "[" & (i - 1) & "]"
            Else
                new_key = "[" & (i - 1) & "]"
            End If
            Dim subItems As Scripting.Dictionary
            Set subItems = flatten_split_keys(elem, new_key)
            Dim subK As Variant
            For Each subK In subItems.Keys
                items.Add subK, subItems(subK)
            Next
        Next
    Else
        items.Add parent_key, value
    End If
    Set flatten_split_keys = items
End Function

' targetDict: Scripting.Dictionary
' keys: 配列（Splitで.区切りにする）
' value: セットする値
Sub set_nested_dict_from_concat_key(targetDict As Scripting.Dictionary, keys As Variant, value As Variant)
    Dim key As String
    key = keys(0)
    
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(\w+)\[(\d+)\]"
    
    If re.Test(key) Then
        Dim m As Object
        Set m = re.Execute(key)(0)
        Dim k As String, idx As Integer
        k = m.SubMatches(0)
        idx = CInt(m.SubMatches(1))
        
        If Not targetDict.Exists(k) Then
            targetDict.Add k, New Collection
        End If
        
        Dim col As Collection
        Set col = targetDict(k)
        ' Collectionは1-basedなのでpadding
        Do While col.Count <= idx
            col.Add New Scripting.Dictionary
        Loop
        If UBound(keys) = 0 Then
            col(idx + 1) = value
        Else
            If TypeName(col(idx + 1)) <> "Dictionary" Then
                Set col(idx + 1) = New Scripting.Dictionary
            End If
            set_nested_dict_from_concat_key col(idx + 1), SliceArray(keys, 1), value
        End If
    Else
        If UBound(keys) = 0 Then
            targetDict(key) = value
        Else
            If Not targetDict.Exists(key) Or TypeName(targetDict(key)) <> "Dictionary" Then
                targetDict(key) = New Scripting.Dictionary
            End If
            set_nested_dict_from_concat_key targetDict(key), SliceArray(keys, 1), value
        End If
    End If
End Sub

' 配列の2番目以降を返す（Zero-based Slice用ヘルパ）
Function SliceArray(arr As Variant, startIdx As Integer) As Variant
    Dim n As Integer: n = UBound(arr) - startIdx + 1
    Dim res() As Variant
    ReDim res(0 To n - 1)
    Dim i As Integer
    For i = 0 To n - 1
        res(i) = arr(i + startIdx)
    Next
    SliceArray = res
End Function

' 文字列を返す。パディング処理も再現
Function format_hcl_value(name As String, val As Variant, indent As Integer, Optional eqpad As String = "", Optional is_map As Boolean = False) As String
    Dim ind As String
    ind = String(indent * 2, " ")
    If IsMissing(eqpad) Then eqpad = ""
    
    If IsEmpty(val) Or (IsArray(val) And UBound(val) = -1) Or (TypeName(val) = "Collection" And val.Count = 0) Then
        format_hcl_value = ""
        Exit Function
    End If
    If VarType(val) = vbString Then
        Dim re As Object: Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "(\$\{|\{\$)([^}]+)\}"
        If re.Test(val) Then
            Dim expr As String
            expr = re.Execute(val)(0).SubMatches(1)
            format_hcl_value = ind & name & eqpad & " = " & expr & vbCrLf
        Else
            format_hcl_value = ind & name & eqpad & " = """ & val & """" & vbCrLf
        End If
    ElseIf VarType(val) = vbBoolean Then
        format_hcl_value = ind & name & eqpad & " = " & LCase(CStr(val)) & vbCrLf
    ElseIf VarType(val) = vbInteger Or VarType(val) = vbLong Or VarType(val) = vbSingle Or VarType(val) = vbDouble Then
        format_hcl_value = ind & name & eqpad & " = " & val & vbCrLf
    ElseIf TypeName(val) = "Collection" Then
        If val.Count = 0 Then
            format_hcl_value = ind & name & eqpad & " = []" & vbCrLf
        Else
            Dim arr() As String
            ReDim arr(1 To val.Count)
            Dim i As Integer
            For i = 1 To val.Count
                arr(i) = """" & val(i) & """"
            Next
            format_hcl_value = ind & name & eqpad & " = [" & Join(arr, ", ") & "]" & vbCrLf
        End If
    Else
        format_hcl_value = ind & name & eqpad & " = " & val & vbCrLf
    End If
End Function

' name: 文字列
' val: Scripting.Dictionary or Collection or primitive
' indent: Integer
' 戻り値: 文字列（HCLブロック）
Function dict_to_hcl_block(name As String, val As Variant, Optional indent As Integer = 0) As String
    Dim ind As String
    ind = String(indent * 2, " ")
    Dim hcl As String
    hcl = ""

    If TypeName(val) = "Dictionary" Then
        Dim allPrimitive As Boolean
        allPrimitive = True
        Dim k As Variant
        For Each k In val.Keys
            If TypeName(val(k)) = "Dictionary" Or TypeName(val(k)) = "Collection" Then
                allPrimitive = False
                Exit For
            End If
        Next
        If allPrimitive Then
            Dim keys As Variant
            keys = val.Keys
            Dim maxlen As Integer
            maxlen = 0
            For Each k In keys
                If Len(k) > maxlen Then maxlen = Len(k)
            Next
            hcl = ind & name & " = {" & vbCrLf
            For Each k In keys
                Dim v As Variant
                v = val(k)
                Dim eqpad As String
                eqpad = String(maxlen - Len(k), " ")
                Dim out As String
                out = format_hcl_value(k, v, indent + 1, eqpad, True)
                If out <> "" Then
                    hcl = hcl & out
                End If
            Next
            hcl = hcl & ind & "}" & vbCrLf
            dict_to_hcl_block = hcl
            Exit Function
        End If
        ' kvs: primitive、blocks: dict or list of dict
        Dim kvs As New Collection
        Dim blocks As New Collection
        For Each k In val.Keys
            If TypeName(val(k)) = "Dictionary" Or (TypeName(val(k)) = "Collection" And val(k).Count > 0 And TypeName(val(k)(1)) = "Dictionary") Then
                blocks.Add Array(k, val(k))
            Else
                kvs.Add Array(k, val(k))
            End If
        Next
        maxlen = 0
        For i = 1 To kvs.Count
            If Len(kvs(i)(0)) > maxlen Then maxlen = Len(kvs(i)(0))
        Next
        hcl = ind & name & " {" & vbCrLf
        For i = 1 To kvs.Count
            eqpad = String(maxlen - Len(kvs(i)(0)), " ")
            out = format_hcl_value(kvs(i)(0), kvs(i)(1), indent + 1, eqpad)
            If out <> "" Then
                hcl = hcl & out
            End If
        Next
        For i = 1 To blocks.Count
            hcl = hcl & dict_to_hcl_block(blocks(i)(0), blocks(i)(1), indent + 1)
        Next
        hcl = hcl & ind & "}" & vbCrLf
        dict_to_hcl_block = hcl
        Exit Function
    ElseIf TypeName(val) = "Collection" Then
        Dim allDict As Boolean
        allDict = True
        For i = 1 To val.Count
            If TypeName(val(i)) <> "Dictionary" Then allDict = False
        Next
        If allDict Then
            Dim blocksStr As String
            blocksStr = ""
            For i = 1 To val.Count
                blocksStr = blocksStr & dict_to_hcl_block(name, val(i), indent)
            Next
            dict_to_hcl_block = blocksStr
            Exit Function
        End If
        If val.Count = 0 Then
            dict_to_hcl_block = ind & name & " = []" & vbCrLf
            Exit Function
        Else
            Dim arr() As String
            ReDim arr(1 To val.Count)
            For i = 1 To val.Count
                arr(i) = """" & val(i) & """"
            Next
            dict_to_hcl_block = ind & name & " = [" & Join(arr, ", ") & "]" & vbCrLf
            Exit Function
        End If
    Else
        dict_to_hcl_block = format_hcl_value(name, val, indent)
        Exit Function
    End If
End Function

' d: Scripting.Dictionary
' Return: 文字列（全リソースHCL）
Function dict_to_resource_hcl(d As Scripting.Dictionary) As String
    Dim hcl As String
    hcl = ""
    Dim res_type As Variant, res_objs As Variant
    For Each res_type In d.Keys
        Set res_objs = d(res_type)
        Dim res_name As Variant, content As Variant
        For Each res_name In res_objs.Keys
            content = res_objs(res_name)
            hcl = hcl & "resource """ & res_type & """ """ & res_name & """ {" & vbCrLf
            Dim keys As Variant: keys = content.Keys
            Dim maxlen As Integer: maxlen = 0
            Dim k As Variant
            For Each k In keys
                If Len(k) > maxlen Then maxlen = Len(k)
            Next
            For Each k In keys
                Dim v As Variant: v = content(k)
                Dim eqpad As String: eqpad = String(maxlen - Len(k), " ")
                If TypeName(v) = "Dictionary" Or (TypeName(v) = "Collection" And v.Count > 0 And TypeName(v(1)) = "Dictionary") Then
                    ' skip, handled after
                Else
                    Dim out As String: out = format_hcl_value(k, v, 1, eqpad)
                    If out <> "" Then hcl = hcl & out
                End If
            Next
            For Each k In content.Keys
                v = content(k)
                If TypeName(v) = "Dictionary" Or (TypeName(v) = "Collection" And v.Count > 0 And TypeName(v(1)) = "Dictionary") Then
                    hcl = hcl & dict_to_hcl_block(k, v, 1)
                End If
            Next
            hcl = hcl & "}" & vbCrLf & vbCrLf
        Next
    Next
    dict_to_resource_hcl = hcl
End Function

' ws: Worksheet
' headerWords: 配列
' 戻り値: 配列 (行番号, 各列index) or Nothing
Function find_header_row(ws As Worksheet, headerWords As Variant) As Variant
    Dim r As Range
    Dim idx As Long
    idx = 1
    For Each r In ws.UsedRange.Rows
        Dim foundAll As Boolean
        foundAll = True
        Dim i As Integer
        For i = LBound(headerWords) To UBound(headerWords)
            Dim word As String: word = headerWords(i)
            Dim found As Boolean: found = False
            Dim c As Range
            For Each c In r.Cells
                If c.Value = word Then found = True
            Next
            If Not found Then
                foundAll = False
                Exit For
            End If
        Next
        If foundAll Then
            Dim indices() As Integer
            ReDim indices(0 To UBound(headerWords))
            For i = LBound(headerWords) To UBound(headerWords)
                word = headerWords(i)
                For Each c In r.Cells
                    If c.Value = word Then
                        indices(i) = c.Column - 1
                        Exit For
                    End If
                Next
            Next
            Dim result() As Variant
            ReDim result(0 To UBound(headerWords) + 1)
            result(0) = idx
            For i = 1 To UBound(headerWords) + 1
                result(i) = indices(i - 1)
            Next
            find_header_row = result
            Exit Function
        End If
        idx = idx + 1
    Next
    find_header_row = Nothing
End Function

' tf_path: ファイルパス（JSONやCSVなど。HCLはパース不可）
' output_excel: 出力Excelファイルパス
Sub export_hcl_to_excel(tf_path As String, output_excel As String)
    MsgBox "この処理（HCL→Excel）はVBA単体では難しいです。" & vbCrLf & _
           "→Pythonや外部ツールで事前にJSON/CSVに変換してください。", vbCritical
End Sub

' input_excel: Excelファイルパス
' output_hcl: HCLファイルパス
' concat_key_header, tf_col_header: ヘッダ名
Sub export_excel_to_hcl(input_excel As String, output_hcl As String, _
    Optional concat_key_header As String = "連結キー", _
    Optional tf_col_header As String = "tf設定値")

    Dim wb As Workbook, ws As Worksheet
    Set wb = Workbooks.Open(input_excel)
    Dim tf_data As Object
    Set tf_data = CreateObject("Scripting.Dictionary")

    For Each ws In wb.Worksheets
        Dim headRow As Variant
        headRow = find_header_row(ws, Array(concat_key_header, tf_col_header))
        If IsEmpty(headRow) Then GoTo ContinueNextSheet
        Dim header_row_idx As Long, concat_idx As Long, tf_idx As Long
        header_row_idx = headRow(0): concat_idx = headRow(1) + 1: tf_idx = headRow(2) + 1
        Dim r As Range
        For Each r In ws.UsedRange.Rows
            If r.Row <= header_row_idx Then GoTo NextRow
            Dim concat_key As String, tf_val As Variant
            concat_key = r.Cells(concat_idx).Value
            tf_val = r.Cells(tf_idx).Value
            If concat_key = "" Or tf_val = "" Then GoTo NextRow
            Dim parts() As String: parts = Split(concat_key, ".")
            If UBound(parts) < 2 Then GoTo NextRow
            Dim resource_type As String, res_name As String
            resource_type = parts(0): res_name = parts(1)
            Dim attr_path() As String
            If UBound(parts) > 1 Then
                attr_path = Split(Mid(concat_key, Len(resource_type & "." & res_name) + 2), ".")
            End If
            If Not tf_data.Exists(resource_type) Then tf_data.Add resource_type, CreateObject("Scripting.Dictionary")
            If Not tf_data(resource_type).Exists(res_name) Then tf_data(resource_type).Add res_name, CreateObject("Scripting.Dictionary")
            set_nested_dict_from_concat_key_excel tf_data(resource_type)(res_name), attr_path, tf_val
NextRow:
        Next
ContinueNextSheet:
    Next

    Dim hcl As String
    hcl = dict_to_resource_hcl(tf_data)
    Dim f As Integer
    f = FreeFile
    Open output_hcl For Output As #f
    Print #f, hcl
    Close #f
    wb.Close False
    MsgBox "HCL出力完了: " & output_hcl
End Sub

' VBA用 set_nested_dict_from_concat_key（attr_pathは配列）
Sub set_nested_dict_from_concat_key_excel(data As Object, keys As Variant, value As Variant)
    If UBound(keys) < 0 Then Exit Sub
    Dim key As String: key = keys(0)
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(\w+)\[(\d+)\]"
    If re.Test(key) Then
        Dim matches As Object: Set matches = re.Execute(key)
        Dim k As String: k = matches(0).SubMatches(0)
        Dim idx As Long: idx = CLng(matches(0).SubMatches(1))
        If Not data.Exists(k) Then data.Add k, New Collection
        Do While data(k).Count <= idx
            data(k).Add CreateObject("Scripting.Dictionary")
        Loop
        If UBound(keys) = 0 Then
            data(k)(idx + 1) = value
        Else
            If TypeName(data(k)(idx + 1)) <> "Dictionary" Then
                Set data(k)(idx + 1) = CreateObject("Scripting.Dictionary")
            End If
            set_nested_dict_from_concat_key_excel data(k)(idx + 1), SliceArray(keys, 1), value
        End If
    Else
        If UBound(keys) = 0 Then
            data(key) = value
        Else
            If Not data.Exists(key) Or TypeName(data(key)) <> "Dictionary" Then
                data(key) = CreateObject("Scripting.Dictionary")
            End If
            set_nested_dict_from_concat_key_excel data(key), SliceArray(keys, 1), value
        End If
    End If
End Sub

' 配列を先頭indexから1個スライスして返す
Function SliceArray(arr As Variant, startIndex As Long) As Variant
    Dim n As Long: n = UBound(arr) - startIndex + 1
    If n < 1 Then
        SliceArray = Array()
        Exit Function
    End If
    Dim newArr() As String
    ReDim newArr(0 To n - 1)
    Dim i As Long
    For i = 0 To n - 1
        newArr(i) = arr(i + startIndex)
    Next
    SliceArray = newArr
End Function

Sub import_tf_to_excel(tf_path As String, input_excel As String, output_excel As String)
    MsgBox "この処理（HCL→Excel値反映）はVBA単体では難しいです。" & vbCrLf & _
           "→Pythonや外部ツールで事前にJSON/CSVに変換してください。", vbCritical
End Sub

' d: Dictionary
' 戻り値: Collection（重複なしvar名）
Function extract_vars_from_dict(d As Object) As Collection
    Dim vars_found As New Collection
    extract_vars_from_dict_rec d, vars_found
    Set extract_vars_from_dict = vars_found
End Function

Sub extract_vars_from_dict_rec(val As Variant, vars_found As Collection)
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "(\$\{var\.([a-zA-Z0-9_]+)\}|\{\$var\.([a-zA-Z0-9_]+)\}|var\.([a-zA-Z0-9_]+))"
    If TypeName(val) = "Dictionary" Then
        Dim v As Variant
        For Each v In val.Items
            extract_vars_from_dict_rec v, vars_found
        Next
    ElseIf TypeName(val) = "Collection" Then
        For i = 1 To val.Count
            extract_vars_from_dict_rec val(i), vars_found
        Next
    ElseIf VarType(val) = vbString Then
        Dim matches As Object: Set matches = re.Execute(val)
        Dim i As Long
        For i = 0 To matches.Count - 1
            Dim g As String
            g = matches(i).SubMatches(1)
            If g <> "" Then
                On Error Resume Next
                vars_found.Add g, g
                On Error GoTo 0
            End If
        Next
    End If
End Sub

Sub export_excel_to_tf(input_excel As String, output_dir As String, _
    Optional concat_key_header As String = "連結キー", _
    Optional tf_col_header As String = "tf設定値", _
    Optional sheet_name As String = "")

    Dim wb As Workbook, ws As Worksheet
    Set wb = Workbooks.Open(input_excel)
    If Dir(output_dir, vbDirectory) = "" Then MkDir output_dir

    Dim worksheets As Collection: Set worksheets = New Collection
    If sheet_name <> "" Then
        On Error Resume Next
        worksheets.Add wb.Worksheets(sheet_name)
        On Error GoTo 0
        If worksheets.Count = 0 Then
            MsgBox "シート '" & sheet_name & "' が見つかりません。"
            wb.Close False
            Exit Sub
        End If
    Else
        Dim ws2 As Worksheet
        For Each ws2 In wb.Worksheets
            worksheets.Add ws2
        Next
    End If

    For Each ws In worksheets
        Dim res As Variant
        res = find_header_row(ws, Array(concat_key_header, tf_col_header))
        If IsEmpty(res) Then GoTo ContinueNextSheet
        Dim header_row_idx As Long, concat_idx As Long, tf_idx As Long
        header_row_idx = res(0): concat_idx = res(1) + 1: tf_idx = res(2) + 1

        Dim tf_data As Object: Set tf_data = CreateObject("Scripting.Dictionary")
        Dim r As Range

        For Each r In ws.UsedRange.Rows
            If r.Row <= header_row_idx Then GoTo NextRow
            Dim concat_key As String, tf_val As Variant
            concat_key = r.Cells(concat_idx).Value
            tf_val = r.Cells(tf_idx).Value
            If concat_key = "" Then GoTo NextRow
            Dim parts() As String: parts = Split(concat_key, ".")
            If UBound(parts) < 2 Then GoTo NextRow
            Dim resource_type As String, res_name As String
            resource_type = parts(0): res_name = parts(1)
            Dim attr_path() As String
            If UBound(parts) > 1 Then attr_path = Split(Mid(concat_key, Len(resource_type & "." & res_name) + 2), ".")
            ' main.tf用
            If tf_val <> "" Then
                If Not tf_data.Exists(resource_type) Then tf_data.Add resource_type, CreateObject("Scripting.Dictionary")
                If Not tf_data(resource_type).Exists(res_name) Then tf_data(resource_type).Add res_name, CreateObject("Scripting.Dictionary")
                set_nested_dict_from_concat_key_excel tf_data(resource_type)(res_name), attr_path, tf_val
            End If
NextRow:
        Next

        ' main.tf, variables.tf 出力
        Dim resource_type As Variant
        For Each resource_type In tf_data.Keys
            Dim tf_file As String: tf_file = output_dir & "\" & resource_type & ".tf"
            Dim hcl As String: hcl = dict_to_resource_hcl(CreateObject("Scripting.Dictionary"): Add resource_type, tf_data(resource_type))
            Dim f As Integer: f = FreeFile
            Open tf_file For Output As #f: Print #f, hcl: Close #f

            ' variables.tf
            Dim used_vars As Collection: Set used_vars = New Collection
            Dim content As Variant
            For Each content In tf_data(resource_type).Items
                Dim col As Collection: Set col = extract_vars_from_dict(content)
                Dim v As Variant
                For Each v In col
                    On Error Resume Next
                    used_vars.Add v, v
                    On Error GoTo 0
                Next
            Next
            If used_vars.Count > 0 Then
                Dim vars_file As String: vars_file = output_dir & "\" & resource_type & ".variables.tf"
                f = FreeFile
                Open vars_file For Output As #f
                For Each v In used_vars
                    Print #f, "variable """ & v & """ {" & vbCrLf & "  default = ""undefined""" & vbCrLf & "}" & vbCrLf
                Next
                Close #f
            End If
        Next
ContinueNextSheet:
    Next
    wb.Close False
    MsgBox "出力完了: " & output_dir & "\*.tf, *.variables.tf"
End Sub

Sub export_excel_to_tfvars(input_excel As String, output_dir As String, _
    Optional concat_key_header As String = "連結キー", _
    Optional tf_col_header As String = "tf設定値", _
    Optional tfvars_col_header As String = "tfvars設定値", _
    Optional sheet_name As String = "")

    Dim wb As Workbook, ws As Worksheet
    Set wb = Workbooks.Open(input_excel)
    If Dir(output_dir, vbDirectory) = "" Then MkDir output_dir

    Dim worksheets As Collection: Set worksheets = New Collection
    If sheet_name <> "" Then
        On Error Resume Next
        worksheets.Add wb.Worksheets(sheet_name)
        On Error GoTo 0
        If worksheets.Count = 0 Then
            MsgBox "シート '" & sheet_name & "' が見つかりません。"
            wb.Close False
            Exit Sub
        End If
    Else
        Dim ws2 As Worksheet
        For Each ws2 In wb.Worksheets
            worksheets.Add ws2
        Next
    End If

    For Each ws In worksheets
        Dim res As Variant
        res = find_header_row(ws, Array(concat_key_header, tf_col_header, tfvars_col_header))
        If IsEmpty(res) Then GoTo ContinueNextSheet
        Dim header_row_idx As Long, concat_idx As Long, tf_idx As Long, tfvars_idx As Long
        header_row_idx = res(0): concat_idx = res(1) + 1: tf_idx = res(2) + 1: tfvars_idx = res(3) + 1

        Dim tfvars_data As Object: Set tfvars_data = CreateObject("Scripting.Dictionary")
        Dim r As Range

        For Each r In ws.UsedRange.Rows
            If r.Row <= header_row_idx Then GoTo NextRow
            Dim concat_key As String, tf_val As Variant, tfvars_val As Variant
            concat_key = r.Cells(concat_idx).Value
            tf_val = r.Cells(tf_idx).Value
            tfvars_val = r.Cells(tfvars_idx).Value
            If concat_key = "" Then GoTo NextRow
            ' tfvars用（var名のみ抽出）
            If tfvars_val <> "" And tf_val <> "" Then
                Dim re As Object: Set re = CreateObject("VBScript.RegExp")
                re.Pattern = "\{\$var\.([a-zA-Z0-9_]+)\}"
                If re.Test(tf_val) Then
                    Dim matches As Object: Set matches = re.Execute(tf_val)
                    Dim varname As String: varname = matches(0).SubMatches(0)
                    tfvars_data(varname) = tfvars_val
                End If
            End If
NextRow:
        Next

        ' tfvars出力（=位置パディング。リソースタイプごとに1ファイル）
        If tfvars_data.Count > 0 Then
            Dim tfvars_keys As Variant: tfvars_keys = tfvars_data.Keys
            Dim maxlen As Integer: maxlen = 0
            For Each v In tfvars_keys
                If Len(v) > maxlen Then maxlen = Len(v)
            Next
            ' 連結キーからresource_typeを抽出（1つ目のみ）
            Dim resource_type As String
            For Each r In ws.UsedRange.Rows
                Dim concat_key As String
                concat_key = r.Cells(concat_idx).Value
                If concat_key <> "" And InStr(concat_key, ".") > 0 Then
                    resource_type = Split(concat_key, ".")(0)
                    Exit For
                End If
            Next
            If resource_type <> "" Then
                Dim tfvars_file As String: tfvars_file = output_dir & "\" & resource_type & ".tfvars"
                Dim f As Integer: f = FreeFile
                Open tfvars_file For Output As #f
                For Each v In tfvars_keys
                    Dim eqpad As String: eqpad = String(maxlen - Len(v), " ")
                    Print #f, v & eqpad & " = """ & tfvars_data(v) & """"
                Next
                Close #f
            End If
        End If
ContinueNextSheet:
    Next
    wb.Close False
    MsgBox "出力完了: " & output_dir & "\*.tfvars"
End Sub

