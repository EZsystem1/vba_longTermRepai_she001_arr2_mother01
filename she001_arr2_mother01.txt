Option Explicit

Sub she001_arr2_mother01()
    Dim arr1 As Variant
    Dim arr2 As Variant
    Dim temp1 As Variant, temp2 As Variant, temp3 As Variant, temp4 As Variant, temp5 As Variant

    ' データ取得
    Call she001_arr2In(arr1, arr2, temp1, temp2, temp3, temp4, temp5)
    arr2 = she001_arr2_1(arr2)
    '内訳No枝番を記入する
    Call she001_arr2Val1(arr2, arr1)
    '内訳No枝番出力
    Call she001_arr2Out_1(arr2)

    ' ここで arr1, arr2 をさらに処理する場合は追加
End Sub

'内訳No枝番出力
Sub she001_arr2Out_1(arr2 As Variant)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim outputRange As Range
    Dim dataOut As Variant, cleanDataOut As Variant
    Dim i As Long, rowCount As Long
    Dim countNonEmpty As Long

    ' 出力先のシートを取得
    Set ws = ThisWorkbook.Sheets("計画表補助2")

    ' AG列の最終行を取得（10行目から下のデータがある最終行）
    lastRow = ws.Cells(ws.Rows.count, "AG").End(xlUp).Row

    ' 出力先のセル範囲をクリア（AG10行目から最終行まで）
    If lastRow >= 10 Then
        ws.Range("AG10:AG" & lastRow).ClearContents
    End If

    ' 配列の行数を取得（タイトル行を除外）
    rowCount = UBound(arr2, 1) - 1 ' **タイトル行を除外**

    ' **エラー9対策: rowCountが0以下の場合は処理をスキップ**
    If rowCount <= 0 Then
        MsgBox "出力対象のデータがありません。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' **出力用の配列を作成（1列分）**
    ReDim dataOut(1 To rowCount, 1 To 1)

    ' **arr2 の 6列目を dataOut に格納（2行目以降）**
    For i = 2 To rowCount + 1
        dataOut(i - 1, 1) = arr2(i, 6) ' **タイトル行をスキップ**
    Next i

    ' **デバッグ: 配列の中身を確認**
    Debug.Print "=== dataOut の内容 ==="
    For i = 1 To rowCount
        Debug.Print "dataOut(" & i & ",1) = [" & dataOut(i, 1) & "]"
    Next i

    ' **空白を削除した cleanDataOut を作成**
    countNonEmpty = 0
    For i = 1 To rowCount
        If Trim(dataOut(i, 1)) <> "" Then countNonEmpty = countNonEmpty + 1
    Next i

    ' **データが全て空白なら終了**
    If countNonEmpty = 0 Then
        MsgBox "出力対象のデータがすべて空白です。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' **空白を除いた配列を作成**
    ReDim cleanDataOut(1 To countNonEmpty, 1 To 1)
    countNonEmpty = 0
    For i = 1 To rowCount
        If Trim(dataOut(i, 1)) <> "" Then
            countNonEmpty = countNonEmpty + 1
            cleanDataOut(countNonEmpty, 1) = dataOut(i, 1)
        End If
    Next i

    ' **出力範囲を設定（AG11 から）**
    Set outputRange = ws.Range("AG11").Resize(UBound(cleanDataOut, 1), 1)



    ' **一括貼り付け**
    outputRange.Value = cleanDataOut

    ' **オブジェクトの解放**
    Set ws = Nothing
    Set outputRange = Nothing
End Sub



'修繕周期の値が-の時は修繕と更新周期の値を入れ替える
Function she001_arr2_1(arr2 As Variant) As Variant
    Dim i As Long
    Dim rowCount As Long
    Dim temp As Variant
    
    ' 配列の行数を取得
    rowCount = UBound(arr2, 1)

    ' 1行目はタイトルなので対象外にゃ！
    If rowCount <= 1 Then
        she001_arr2_1 = arr2
        Exit Function
    End If

    ' **4列目が "-" のとき、4列目と5列目を入れ替え**
    For i = 2 To rowCount
        If arr2(i, 4) = "-" Then
            ' **値を入れ替え**
            temp = arr2(i, 4)
            arr2(i, 4) = arr2(i, 5)
            arr2(i, 5) = temp
        End If
    Next i
    
    ' **変更後の配列を返す**
    she001_arr2_1 = arr2
End Function


'内訳No枝番を記入する
Sub she001_arr2Val1(arr2 As Variant, arr1 As Variant)
    Dim i As Long, j As Long
    Dim lastRowArr2 As Long, lastRowArr1 As Long
    Dim keyArr2 As String, keyArr1 As String

    ' 配列の行数を取得（タイトル行を除く）
    lastRowArr2 = UBound(arr2, 1)
    lastRowArr1 = UBound(arr1, 1)

    ' データ元配列arr2のループ（2行目から）
    For i = 2 To lastRowArr2
        ' 1列目がブランクならスキップ
        If Trim(arr2(i, 1)) <> "" Then
            ' arr2のi行目の1～5列を連結してキーを作成
            keyArr2 = arr2(i, 1) & "|" & arr2(i, 2) & "|" & arr2(i, 3) & "|" & arr2(i, 4) & "|" & arr2(i, 5)

            ' 参照配列arr1のループ
            For j = 2 To lastRowArr1
                ' arr1のj行目の1～5列を連結してキーを作成
                keyArr1 = arr1(j, 1) & "|" & arr1(j, 2) & "|" & arr1(j, 3) & "|" & arr1(j, 4) & "|" & arr1(j, 5)

                ' キーが一致した場合、arr2の6列目にarr1の8列目の値を転写
                If keyArr2 = keyArr1 Then
                    arr2(i, 6) = arr1(j, 8)
                    Exit For ' 一致したら探索を終了し、次の `arr2` の行へ
                End If
            Next j
        End If
    Next i
End Sub



Sub she001_arr2In(arr1 As Variant, arr2 As Variant, temp1 As Variant, temp2 As Variant, temp3 As Variant, temp4 As Variant, temp5 As Variant)
    Dim wsPlan As Worksheet, wsDetail As Worksheet
    Dim lastRowPlan As Long, lastRowDetail As Long
    Dim i As Long

    ' シートのセット
    Set wsPlan = ThisWorkbook.Sheets("計画表補助2")
    Set wsDetail = ThisWorkbook.Sheets("内訳書")

    ' 計画表補助2の最終行（B列）
    lastRowPlan = wsPlan.Cells(wsPlan.Rows.count, "B").End(xlUp).Row
    
    ' 内訳書の最終行（A列）
    lastRowDetail = wsDetail.Cells(wsDetail.Rows.count, "A").End(xlUp).Row

    ' arr1のデータ格納（B6:I最終行 ）
    arr1 = wsPlan.Range("B6:I" & lastRowPlan).Value
    
    ' temp1～temp5のデータ格納（B, D, C, R, X列の10行目から最終行）
    temp1 = wsDetail.Range("B10:B" & lastRowDetail).Value
    temp2 = wsDetail.Range("D10:D" & lastRowDetail).Value
    temp3 = wsDetail.Range("C10:C" & lastRowDetail).Value
    temp4 = wsDetail.Range("R10:R" & lastRowDetail).Value
    temp5 = wsDetail.Range("X10:X" & lastRowDetail).Value

    ' 配列の行数を取得
    Dim tempRowCount As Long
    tempRowCount = UBound(temp1, 1)

    ' arr2のサイズを定義（temp1～temp5のデータを保持）
    ReDim arr2(1 To tempRowCount, 1 To 6)

    ' arr2 の 1～5列目に temp1～temp5 のデータを格納
    For i = 1 To tempRowCount
        arr2(i, 1) = temp1(i, 1) ' 1列目に temp1
        arr2(i, 2) = temp2(i, 1) ' 2列目に temp2
        arr2(i, 3) = temp3(i, 1) ' 3列目に temp3
        arr2(i, 4) = temp4(i, 1) ' 4列目に temp4
        arr2(i, 5) = temp5(i, 1) ' 5列目に temp5
    Next i
    
    '配列タイトル追記
    arr2(1, 6) = "内訳No枝番"
    
    ' シートのオブジェクトを解放
    Set wsPlan = Nothing
    Set wsDetail = Nothing
End Sub


