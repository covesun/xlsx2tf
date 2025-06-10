'--- flatten_split_keys: ネストdict/listを "foo.bar[0].baz" 連結キー化
Function FlattenSplitKeys(value As Variant, Optional parentKey As String = "") As Object
    Dim items As Object
    Set items = CreateObject("Scripting.Dictionary")
    
    If IsObject(value) Then
        If TypeName(value) = "Dictionary" Then
            Dim k As Variant
            For Each k In value.Keys
                Dim v As Variant
                v = value(k)
                Dim newKey As String
                If parentKey <> "" Then
                    newKey = parentKey & "." & k
                Else
                    newKey = k
                End If
                Dim tmp As Object
                Set tmp = FlattenSplitKeys(v, newKey)
                Dim subk As Variant
                For Each subk In tmp.Keys
                    items(subk) = tmp(subk)
                Next
            Next
        ElseIf TypeName(value) = "Collection" Or IsArray(value) Then
            Dim i As Long
            Dim arr As Variant
            If TypeName(value) = "Collection" Then
                arr = CollectionToArray(value)
            Else
                arr = value
            End If
            For i = LBound(arr) To UBound(arr)
                Dim elem As Variant
                elem = arr(i)
                Dim newKey As String
                If parentKey <> "" Then
                    newKey = parentKey & "[" & i & "]"
                Else
                    newKey = "[" & i & "]"
                End If
                Dim tmp As Object
                Set tmp = FlattenSplitKeys(elem, newKey)
                Dim subk As Variant
                For Each subk In tmp.Keys
                    items(subk) = tmp(subk)
                Next
            Next
        End If
    Else
        items(parentKey) = value
    End If
    
    Set FlattenSplitKeys = items
End Function

Function CollectionToArray(col As Collection) As Variant
    Dim arr() As Variant
    Dim i As Long
    ReDim arr(0 To col.Count - 1)
    For i = 1 To col.Count
        arr(i - 1) = col(i)
    Next
    CollectionToArray = arr
End Function

'--- set_nested_dict_from_concat_key: "foo.bar[0].baz"連結キー→ネストdict/list構築
Sub SetNestedDictFromConcatKey(ByRef data As Object, keys As Variant, value As Variant)
    Dim key As String
    key = keys(0)
    Dim m As Object
    Set m = GetRegexMatch(key, "(\w+)\[(\d+)\]")
    If Not m Is Nothing Then
        Dim k As String, idx As Long
        k = m.SubMatches(0)
        idx = CLng(m.SubMatches(1))
        If Not data.Exists(k) Then
            data(k) = CreateObject("Scripting.Dictionary")
            data(k)("__array__") = New Collection
        End If
        Dim arr As Collection
        Set arr = data(k)("__array__")
        While arr.Count < idx + 1
            arr.Add CreateObject("Scripting.Dictionary")
        Wend
        If UBound(keys) = 0 Then
            arr.Remove idx + 1
            arr.Add value, , idx + 1
        Else
            SetNestedDictFromConcatKey arr(idx + 1), SliceArray(keys, 1), value
        End If
        data(k)("__array__") = arr
    Else
        If UBound(keys) = 0 Then
            data(key) = value
        Else
            If Not data.Exists(key) Or Not IsObject(data(key)) Then
                Set data(key) = CreateObject("Scripting.Dictionary")
            End If
            SetNestedDictFromConcatKey data(key), SliceArray(keys, 1), value
        End If
    End If
End Sub

Function SliceArray(arr As Variant, Optional fromIndex As Long = 1) As Variant
    Dim i As Long, newArr() As Variant
    ReDim newArr(LBound(arr) To UBound(arr) - fromIndex)
    For i = LBound(newArr) To UBound(newArr)
        newArr(i) = arr(i + fromIndex)
    Next
    SliceArray = newArr
End Function

Function GetRegexMatch(str As String, pattern As String) As Object
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = pattern
    regex.Global = False
    If regex.Test(str) Then
        Set GetRegexMatch = regex.Execute(str)(0)
    Else
        Set GetRegexMatch = Nothing
    End If
End Function

'--- format_hcl_value: HCL値フォーマット
Function FormatHclValue(name As String, val As Variant, indent As Integer, Optional eqpad As String = "", Optional isMap As Boolean = False) As String
    Dim ind As String
    ind = String(indent * 2, " ")
    If IsMissing(eqpad) Then eqpad = ""
    If IsEmpty(val) Or IsNull(val) Or val = "" Then
        FormatHclValue = ""
        Exit Function
    End If
    If VarType(val) = vbString Then
        Dim m As Object
        Set m = GetRegexMatch(val, "(\$\{|\{\$)([^}]+)\}")
        If Not m Is Nothing Then
            FormatHclValue = ind & name & eqpad & " = " & m.SubMatches(1) & vbCrLf
        Else
            FormatHclValue = ind & name & eqpad & " = """ & val & """" & vbCrLf
        End If
    ElseIf VarType(val) = vbBoolean Then
        If val Then
            FormatHclValue = ind & name & eqpad & " = true" & vbCrLf
        Else
            FormatHclValue = ind & name & eqpad & " = false" & vbCrLf
        End If
    ElseIf VarType(val) = vbInteger Or VarType(val) = vbLong Or VarType(val) = vbSingle Or VarType(val) = vbDouble Then
        FormatHclValue = ind & name & eqpad & " = " & val & vbCrLf
    ElseIf IsObject(val) Then
        If TypeName(val) = "Collection" Then
            If val.Count = 0 Then
                FormatHclValue = ind & name & eqpad & " = []" & vbCrLf
            Else
                Dim arrStr As String, v As Variant
                arrStr = ""
                For Each v In val
                    arrStr = arrStr & """" & v & """, "
                Next
                arrStr = Left(arrStr, Len(arrStr) - 2)
                FormatHclValue = ind & name & eqpad & " = [" & arrStr & "]" & vbCrLf
            End If
        End If
    End If
End Function

'--- dict_to_hcl_block: 再帰でHCLブロック出力
Function DictToHclBlock(name As String, val As Variant, Optional indent As Integer = 0) As String
    Dim ind As String
    ind = String(indent * 2, " ")
    Dim hcl As String
    hcl = ""
    If IsObject(val) Then
        If TypeName(val) = "Dictionary" Then
            Dim keys As Variant
            keys = val.Keys
            Dim kvs As String, k As Variant
            kvs = ""
            For Each k In keys
                Dim v As Variant
                v = val(k)
                kvs = kvs & FormatHclValue(k, v, indent + 1)
            Next
            hcl = ind & name & " {" & vbCrLf & kvs & ind & "}" & vbCrLf
        ElseIf TypeName(val) = "Collection" Then
            Dim blockStr As String, item As Variant
            blockStr = ""
            For Each item In val
                blockStr = blockStr & DictToHclBlock(name, item, indent)
            Next
            hcl = blockStr
        End If
    Else
        hcl = FormatHclValue(name, val, indent)
    End If
    DictToHclBlock = hcl
End Function

'--- dict_to_resource_hcl: リソース辞書 → HCL出力
Function DictToResourceHcl(d As Object) As String
    Dim hcl As String
    hcl = ""
    Dim resType As Variant, resObjs As Object
    For Each resType In d.Keys
        Set resObjs = d(resType)
        Dim resName As Variant, content As Object
        For Each resName In resObjs.Keys
            Set content = resObjs(resName)
            hcl = hcl & "resource """ & resType & """ """ & resName & """ {" & vbCrLf
            Dim keys As Variant
            keys = content.Keys
            Dim k As Variant
            For Each k In keys
                hcl = hcl & FormatHclValue(k, content(k), 1)
            Next
            hcl = hcl & "}" & vbCrLf & vbCrLf
        Next
    Next
    DictToResourceHcl = hcl
End Function

'--- find_header_row: ヘッダワードで行番号・列番号返す
Function FindHeaderRow(ws As Worksheet, headerWords As Variant) As Variant
    Dim i As Long, row As Range, j As Long
    For i = 1 To ws.UsedRange.Rows.Count
        Set row = ws.Rows(i)
        Dim found As Boolean
        found = True
        For j = LBound(headerWords) To UBound(headerWords)
            If IsError(Application.Match(headerWords(j), row, 0)) Then
                found = False
                Exit For
            End If
        Next
        If found Then
            Dim indices() As Variant
            ReDim indices(LBound(headerWords) To UBound(headerWords))
            For j = LBound(headerWords) To UBound(headerWords)
                indices(j) = Application.Match(headerWords(j), row, 0)
            Next
            FindHeaderRow = Array(i, indices)
            Exit Function
        End If
    Next
    FindHeaderRow = Null
End Function

'--- extract_vars_from_dict: {$var.xxx} 形式からxxx抽出（単純化）
Function ExtractVarNameFromCell(val As Variant) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\{\$var\.([a-zA-Z0-9_]+)\}"
    re.Global = False
    If re.Test(val) Then
        ExtractVarNameFromCell = re.Execute(val)(0).SubMatches(0)
    Else
        ExtractVarNameFromCell = ""
    End If
End Function
