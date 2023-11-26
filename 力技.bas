Attribute VB_Name = "Main"
#If VBA7 And Win64 Then
  ' 64Bit 版
  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
'sleep
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#Else
  ' 32Bit 版
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
#End If
'*********************
'変数
'Dim sumcell_list(0, 0)
Dim filter_list As New Collection
'コード
Dim code1
Dim div1

Dim code2
Dim div2

Dim code3
Dim div3

Dim code4
Dim div4

Dim code5
Dim div5

Dim code6
Dim div6

Dim code7
Dim div7

Dim code8
Dim div8

Dim code9
Dim div9

Dim code10
Dim div10

Dim code11
Dim div11

'code12 = "82J" Or x = "82L" Or x = "820"
Dim code12_1
Dim code12_2
Dim code12_3
Dim div12

Dim code13
Dim div13

Dim code14
Dim div14

'code15
Dim code15_1
Dim code15_2
Dim div15

'判定用受注元セル
Dim jyutyuu_cell
'csv
Dim csvfile
'保存場所
Dim save_dir
'ファイル名
Dim save_filename


'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'メイン処理
'*********************
'*********************
'1.csvデータ処理(不要列削除とコピー)
'2.既存シート削除
'3.受注元リスト作成し
'4.フィルタした受注元ごとに
'5.シート作成とデータコピーし
'6.受注元が複数一致するパターンのシートを統合
'7.シート名を変更
'7-2.受注元名称(シート名と同様)を追加
'8.各シートと元シートに合計行の追加
'9.色付け
'10.フィルタ解除と保存
'*********************
Sub Main()
Application.ScreenUpdating = False
'*********************
'変数
'ReDim sumcell_list(0, 0)
'Dim filter_list As New Collection
Set filter_list = New Collection
'コード
code1 = "870"
div1 = "札幌"

code2 = "880"
div2 = "東北"

code3 = "825"
div3 = "関東"

code4 = "824"
div4 = "横浜"

code5 = "850"
div5 = "中部"

code6 = "851"
div6 = "北陸"

code7 = "830"
div7 = "関西"

code8 = "860"
div8 = "中四"

code9 = "840"
div9 = "九州"

code10 = "82B"
div10 = "(本)(F)"

code11 = "82K"
div11 = "(本)(B)"

'code12 = "82J" Or x = "82L" Or x = "820"
code12_1 = "82J"
code12_2 = "82L"
code12_3 = "820"
div12 = "(本)(D)"

code13 = "82M"
div13 = "(本)(耐震)"

code14 = "81A"
div14 = "(本)(M)"

'code15 = "890" Or "891"
code15_1 = "890"
code15_2 = "891"
div15 = "海外"

'受注元コード参照セル
jyutyuu_cell = "G2"


'@rem 202002移動対応--
'csvデータ
csvfile = "C:\データ.csv"

'保存場所
save_dir = "C:\"
'@rem 202002移動対応--

'ファイル名
save_filename = "送付先リスト.xlsx"
'*********************

'*********************
'1.csvデータ処理(不要列削除とコピー)
'csv開く
Workbooks.OpenText csvfile, Comma:=True
'csv移動
Call move_csv
'列削除
Call del_col
'コピー
Call csv_copy

'*********************
'2.既存シート削除
'Call read_data
Call del_sh

'*********************
'色付け
Call sh_format

'*********************
'3.フィルタ用リスト作成
Call create_filter_list

'---------------------
'---------------------
'繰り返し
'For i = LBound(filter_list) To UBound(filter_list)
For i = 1 To filter_list.Count
'*********************
'4.フィルタ
Call jyutyuu_filter(i)

'*********************
'5.シート追加、シート名変更
Call sh_copy

Next i
'---------------------
'---------------------

'*********************
'6.受注元が複数一致するパターンのシートを統合
'条件は固定とする
Call sh_merge

'*********************
'7.シート名を変更
Call sh_name_set
'7-2.受注元名称(シート名と同様)を追加
Call jyutyuuname_set

'*********************
'8.各シートに合計行追加、元シートにも受注元追加
Call add_sum_each


'*********************
'保存
Call save

End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'1.csvデータ処理(不要列削除とコピー)
'*********************
'csv移動
'*********************
Sub move_csv()
'
' Macro1 Macro
'
'*********************
'変数
'シート
Dim basebook As Workbook
Set basebook = Workbooks("送付先リスト.xlsm")
Dim base_sh As Worksheet
Set base_sh = basebook.Sheets("全社")
'*********************

'
    Windows("残確データ.csv").Activate
    Sheets("残確データ").Select
'    Sheets("残確データ").Move After:=Workbooks("送付先リスト.xlsm").Sheets(15)
    Sheets("残確データ").Move After:=base_sh
End Sub
'*********************
'シート追加と名称変更
'*********************
Public Sub csv_copy()

'*********************
'変数
'シート
Dim target_sh As Worksheet
Set target_sh = Sheets("全社")
Dim base_sh As Worksheet
'*********************

'*********************
'シート追加
Set base_sh = ActiveSheet
'値貼り付け
target_sh.UsedRange.ClearContents
target_sh.UsedRange.Value = Empty
target_sh.UsedRange.Interior.ColorIndex = 0
base_sh.Range("A1").CurrentRegion.Copy
target_sh.Range("A1").PasteSpecial Paste:=xlPasteValues

End Sub
'*********************
'列を検索して削除
'*********************
Public Sub del_col()
'*********************
'変数
col1 = "残高確認書指定区分"
col2 = "注文主会社コード"
col3 = "受注元名称"
col4 = "カンパニー名称"
'*********************


'*********************
'削除
'残高確認書指定区分
If ((Cells.Find(col1, LookAt:=xlWhole) Is Nothing) = False) Then
zandaka_col = Cells.Find(col1, LookAt:=xlWhole).Column
Columns(zandaka_col).Delete
End If

'注文主会社コード
If ((Cells.Find(col2, LookAt:=xlWhole) Is Nothing) = False) Then
tyuumon_col = Cells.Find(col2, LookAt:=xlWhole).Column
Columns(tyuumon_col).Delete
End If

'受注元名称
If ((Cells.Find(col3, LookAt:=xlWhole) Is Nothing) = False) Then
jyutyuumoto_col = Cells.Find(col3, LookAt:=xlWhole).Column
Columns(jyutyuumoto_col).Delete
End If

'カンパニー名称
If ((Cells.Find(col4, LookAt:=xlWhole) Is Nothing) = False) Then
company_col = Cells.Find(col4, LookAt:=xlWhole).Column
Columns(company_col).Delete
End If


End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'2.既存シート削除
'*********************
'シート削除
'*********************
Public Sub del_sh()
'*********************
'変数
'シート
Dim base_sh As Worksheet
Set base_sh = Sheets("全社")
    Dim mySht As Worksheet
'*********************
base_sh.Activate
    
    With Application
        '警告や確認のメッセージを非表示に設定
        .DisplayAlerts = False
        'シート名をチェックして、アクティブシートでなければ削除
        For Each mySht In Worksheets
            If mySht.Name <> ActiveSheet.Name Then mySht.Delete
        Next
        '設定を元に戻す
        .DisplayAlerts = True
    End With
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'3.フィルタ用リスト作成
'*********************
'フィルタ用リスト作成
'*********************
Public Function create_filter_list()
'*********************
'変数
'col1 = "受注元"
col1 = "受注元コード"
col2 = "当月末合計残高"
Dim i As Long
startrow = 2
'最終行
If ((Cells.Find(col2, LookAt:=xlWhole) Is Nothing) = False) Then
zandaka_col = Cells.Find(col2, LookAt:=xlWhole).Column
End If
zandaka_col_char = Chr(zandaka_col + 64)
maxrow = Cells(Rows.Count, zandaka_col).End(xlUp).Row

'作業列
'最終列プラス1
maxcol = Cells(1, Columns.Count).End(xlToLeft).Column + 1
maxcol_char = Chr(maxcol + 64)

'リスト取得用
If ((Cells.Find(col1, LookAt:=xlWhole) Is Nothing) = False) Then
jyutyuumoto_col = Cells.Find(col1, LookAt:=xlWhole).Column
jyutyuumoto_col_char = Chr(jyutyuumoto_col + 64)
End If

'Set filter_list = New Collection
'*********************
'unique作成
ActiveSheet.Range("$" & jyutyuumoto_col_char & ":$" & jyutyuumoto_col_char & "").AdvancedFilter Action:=xlFilterCopy, _
    CopyToRange:=Range("$" & maxcol_char & "1"), _
    Unique:=True

'リスト化
'データを登録する間、エラーを無視する
maxrow_list = Cells(Rows.Count, maxcol).End(xlUp).Row
On Error Resume Next

For i = startrow To maxrow_list
filter_list.Add Cells(i, maxcol).Value
Next i
On Error GoTo 0

'作業列削除
ActiveSheet.Range("$" & maxcol_char & ":$" & maxcol_char & "").Delete
'*********************


End Function
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'4.フィルタ
'*********************
'列を検索してフィルタ
'*********************
Public Sub jyutyuu_filter(i)
'*********************
'変数
'シート
Dim base_sh As Worksheet
Set base_sh = Sheets("全社")

'col1 = "受注元"
col1 = "受注元コード"
col2 = "当月末合計残高"
startrow = 2

'列取得
If ((Cells.Find(col1, LookAt:=xlWhole) Is Nothing) = False) Then
jyutyuumoto_col = Cells.Find(col1, LookAt:=xlWhole).Column
End If
jyutyuumoto_col_char = Chr(jyutyuumoto_col + 64)

'Set filter_list = New Collection
'*********************


'*********************
'フィルタ
'受注元
base_sh.Range("$" & jyutyuumoto_col_char & ":$" & jyutyuumoto_col_char & "").AutoFilter
base_sh.Range("$" & jyutyuumoto_col_char & ":$" & jyutyuumoto_col_char & "").AutoFilter Field:=1, Criteria1:=filter_list(i)


End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'5.シート追加、シート名変更
'*********************
'シート追加と名称変更
'*********************
Public Sub sh_copy()

'*********************
'変数
'シート
Dim base_sh As Worksheet
Set base_sh = Sheets("全社")
Dim act_sh As Worksheet

col1 = "受注元コード"
If ((Cells.Find(col1, LookAt:=xlWhole) Is Nothing) = False) Then
jyutyuumoto_col = Cells.Find(col1, LookAt:=xlWhole).Column
End If
jyutyuumoto_col_char = Chr(jyutyuumoto_col + 64)
'*********************

'*********************
'シート追加
Worksheets.Add Before:=ActiveSheet
Set act_sh = ActiveSheet
base_sh.Range("A1").CurrentRegion.Copy
act_sh.Range("A1").PasteSpecial Paste:=xlPasteColumnWidths
base_sh.Range("A1").CurrentRegion.Copy
base_sh.Range("A1").CurrentRegion.Copy act_sh.Range("A1")
ActiveSheet.Name = Replace(Replace(Trim(act_sh.Range(jyutyuumoto_col_char & "2").Value), " ", ""), "　", "")

End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'6.受注元が複数一致するパターンのシートを統合
'*********************
'シート統合
'どれか1つのシートに統合する
'*********************
Public Sub sh_merge()
'*********************
'変数
col1 = "受注元コード"
col2 = "当月末合計残高"
'シート
Dim base_sh As Worksheet
Dim merge_sh As Worksheet
Dim sheetName As String
'最終行
If ((Cells.Find(col2, LookAt:=xlWhole) Is Nothing) = False) Then
zandaka_col = Cells.Find(col2, LookAt:=xlWhole).Column
End If
zandaka_col_char = Chr(zandaka_col + 64)
maxrow = Cells(Rows.Count, zandaka_col).End(xlUp).Row


'受注元コード参照セル
'jyutyuu_cell = "G2"
'*********************

'---------------------
'---------------------
'繰り返し
'For i = 1 To Worksheets.Count - 1
For i = Worksheets.Count To 1 Step -1
'*********************
sheetName = Worksheets(i).Name

Set merge_sh = Sheets(i)
merge_maxrow = merge_sh.Cells(Rows.Count, zandaka_col).End(xlUp).Row

'値取得と空白削除
'x = Replace(Replace(Trim(merge_sh.Range(jyutyuu_cell).Value), " ", ""), "　", "")
x = Replace(Replace(Trim(sheetName), " ", ""), "　", "")
If x = "82L" Or x = "820" Then
'x = "82J"
Set base_sh = Sheets("82J")
maxrow = base_sh.Cells(Rows.Count, zandaka_col).End(xlUp).Row

merge_sh.Activate
merge_sh.Range("A1").CurrentRegion.Select
Selection.Offset(1, 0).Select
Selection.Resize(Selection.Rows.Count - 1).Select
Selection.Copy base_sh.Range("A" & maxrow + 1&)

'シート削除
'i = i - 1
Application.DisplayAlerts = False
merge_sh.Delete
Application.DisplayAlerts = True

ElseIf x = "891" Then
'x = "890"
Set base_sh = Sheets("890")
maxrow = base_sh.Cells(Rows.Count, zandaka_col).End(xlUp).Row

merge_sh.Activate
merge_sh.Range("A1").CurrentRegion.Select
Selection.Offset(1, 0).Select
Selection.Resize(Selection.Rows.Count - 1).Select
Selection.Copy base_sh.Range("A" & maxrow + 1&)

'シート削除
'i = i - 1
Application.DisplayAlerts = False
merge_sh.Delete
Application.DisplayAlerts = True

End If
Next i
'---------------------
'---------------------

End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'7.シート名を変更
'*********************
'シート名称をコードから名称へ変更
'*********************
Public Sub sh_name_set()

'*********************
'変数
'シート
Dim base_sh As Worksheet
Set base_sh = Sheets("全社")
Dim rename_sh As Worksheet
Dim sheetName As String
'受注元コード参照セル
'jyutyuu_cell = "G2"
'*********************

'---------------------
'---------------------
'繰り返し
For i = 1 To Worksheets.Count
'*********************
Set rename_sh = Worksheets(i)
sheetName = Worksheets(i).Name
'値取得と空白削除
'x = Replace(Replace(Trim(rename_sh.Range(jyutyuu_cell).Value), " ", ""), "　", "")
x = Replace(Replace(Trim(sheetName), " ", ""), "　", "")

'シート名
'div1 = "札幌"
If x = code1 Then
rename_sh.Name = div1

'div2 = "東北"
ElseIf x = code2 Then
rename_sh.Name = div2

'div3 = "関東"
ElseIf x = code3 Then
rename_sh.Name = div3

'div4 = "横浜"
ElseIf x = code4 Then
rename_sh.Name = div4

'div5 = "中部"
ElseIf x = code5 Then
rename_sh.Name = div5

'div6 = "北陸"
ElseIf x = code6 Then
rename_sh.Name = div6

'div7 = "関西"
ElseIf x = code7 Then
rename_sh.Name = div7

'div8 = "中四"
ElseIf x = code8 Then
rename_sh.Name = div8

'div9 = "九州"
ElseIf x = code9 Then
rename_sh.Name = div9

'div10 = "(本)(F)"
ElseIf x = code10 Then
rename_sh.Name = div10

'div11 = "(本)(B)"
ElseIf x = code11 Then
rename_sh.Name = div11

'code12_1 = "82J"
'code12_2 = "82L"
'code12_3 = "820"
'div12 = "(本)(D)"
ElseIf x = code12_1 Or x = code12_2 Or x = code12_3 Then
rename_sh.Name = div12

'div13 = "(本)(耐震)"
ElseIf x = code13 Then
rename_sh.Name = div13

'div14 = "(本)(M)"
ElseIf x = code14 Then
rename_sh.Name = div14

'code15_1 = "890"
'code15_2 = "891"
'div15 = "海外"
ElseIf x = code15_1 Or x = code15_2 Then
rename_sh.Name = div15

End If

Next i
'---------------------
'---------------------

End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'7-2.受注元名称(シート名と同様)を追加
'*********************
'シート名称をコードから名称へ変更
'*********************
Public Sub jyutyuuname_set()

'*********************
'変数
'シート
Dim sheetName As String
Dim target_sheet As Worksheet
'変数
col1 = "受注元コード"
col2 = "当月末合計残高"
startrow = 2
'最終行
If ((Cells.Find(col2, LookAt:=xlWhole) Is Nothing) = False) Then
zandaka_col = Cells.Find(col2, LookAt:=xlWhole).Column
End If
zandaka_col_char = Chr(zandaka_col + 64)
maxrow = Cells(Rows.Count, zandaka_col).End(xlUp).Row

'*********************

'---------------------
'---------------------
'繰り返し
For i = 1 To Worksheets.Count
'*********************
Set target_sheet = Worksheets(i)
sheetName = Worksheets(i).Name
target_sheet.Activate
maxrow = target_sheet.Cells(Rows.Count, zandaka_col).End(xlUp).Row

'フィルタ解除
If target_sheet.AutoFilterMode = True Then
target_sheet.Cells.AutoFilter
End If

'値取得と空白削除
x = Replace(Replace(Trim(sheetName), " ", ""), "　", "")

'列追加
target_sheet.Columns(1).Insert

'項目名(受注先名)をセット
target_sheet.Cells(1, 1) = "受注先名"

'シート名(受注先名)をセット
target_sheet.Range(Cells(startrow, 1), Cells(maxrow, 1)).Value = sheetName

'書式コピペ
target_sheet.Range(Cells(1, 2), Cells(maxrow, 2)).Copy
target_sheet.Range(Cells(1, 1), Cells(maxrow, 1)).PasteSpecial Paste:=xlPasteFormats

Next i
'---------------------
'---------------------

End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'8.各シートに合計行追加、元シートにも受注元追加
'*********************
'最終行を取得して合計追加
'*********************
Public Sub add_sum_each()
'*********************
'*********************
'シート
Dim base_sh As Worksheet
Set base_sh = Sheets("全社")
Dim target_sheet As Worksheet
'変数
'col1 = "受注元"
col1 = "受注元コード"
col2 = "当月末合計残高"
startrow = 2
'最終行
If ((Cells.Find(col2, LookAt:=xlWhole) Is Nothing) = False) Then
zandaka_col = Cells.Find(col2, LookAt:=xlWhole).Column
End If
zandaka_col_char = Chr(zandaka_col + 64)
maxrow = Cells(Rows.Count, zandaka_col).End(xlUp).Row

'元シート
base_maxrow = base_sh.Cells(Rows.Count, zandaka_col).End(xlUp).Row

'*********************


'---------------------
'---------------------
'繰り返し
For i = 1 To Worksheets.Count
'*********************
Set rename_sh = Worksheets(i)
sheetName = Worksheets(i).Name
Set target_sheet = Worksheets(i)
maxrow = target_sheet.Cells(Rows.Count, zandaka_col).End(xlUp).Row

If (sheetName <> base_sh.Name) Then
'合計追加
target_sheet.Cells(maxrow + 1, zandaka_col).Value = "=SUM(" & zandaka_col_char & startrow & ":" & zandaka_col_char & maxrow & ")"
base_sh.Cells((base_maxrow + 2) + i, zandaka_col - 1).Value = sheetName
base_sh.Cells((base_maxrow + 2) + i, zandaka_col).Value = "='" & sheetName & "'!" & zandaka_col_char & maxrow + 1
Else
base_sh.Cells(base_maxrow + 1, zandaka_col).Value = "=SUM(" & zandaka_col_char & startrow & ":" & zandaka_col_char & base_maxrow & ")"
base_sh.Cells(base_maxrow + 2, zandaka_col).Value = "=SUM(" & zandaka_col_char & base_maxrow + 3 & ":" & zandaka_col_char & base_maxrow + ((i - 1) + 2) & ")"

End If

Next i
'---------------------
'---------------------

End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'9.色付け
'*********************
'*********************
Public Sub sh_format()
'*********************
'変数
col1 = "確認書№"
'シート
Dim target_sh As Worksheet
'最終行
If ((Cells.Find(col1, LookAt:=xlWhole) Is Nothing) = False) Then
target_col = Cells.Find(col1, LookAt:=xlWhole).Column
End If
target_col_char = Chr(target_col + 64)
maxrow = Cells(Rows.Count, target_col).End(xlUp).Row
maxcol = Cells(1, Columns.Count).End(xlToLeft).Column

target = "-1-"

'受注元コード参照セル
'jyutyuu_cell = "G2"
'*********************

'---------------------
'---------------------
'繰り返し
For i = Worksheets.Count To 1 Step -1
'*********************
Set target_sh = Worksheets(i)
maxrow = target_sh.Cells(Rows.Count, target_col).End(xlUp).Row

For j = 1 To maxrow
'色変更
If Cells(j, target_col) Like "*" & target & "*" Then
Range(Cells(j, 1), Cells(j, maxcol)).Interior.Color = 11389944
End If
Next j

Next i
'---------------------
'---------------------

End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'10.フィルタ解除と保存
'*********************
'*********************
Public Sub save()
'*********************
'*********************
'シート
Dim base_sh As Worksheet
Set base_sh = Sheets("全社")
'変数
'ファイル名
save_filename = "送付先リスト.xlsx"
'*********************

'フィルタ解除
base_sh.Cells.AutoFilter

'保存
Application.DisplayAlerts = False
ActiveWorkbook.SaveAs Filename:=save_dir & save_filename, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
Application.DisplayAlerts = True

End Sub
