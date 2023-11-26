Attribute VB_Name = "Main"
#If VBA7 And Win64 Then
  ' 64Bit ��
  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
'sleep
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#Else
  ' 32Bit ��
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
#End If
'*********************
'�ϐ�
'Dim sumcell_list(0, 0)
Dim filter_list As New Collection
'�R�[�h
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

'����p�󒍌��Z��
Dim jyutyuu_cell
'csv
Dim csvfile
'�ۑ��ꏊ
Dim save_dir
'�t�@�C����
Dim save_filename


'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'���C������
'*********************
'*********************
'1.csv�f�[�^����(�s�v��폜�ƃR�s�[)
'2.�����V�[�g�폜
'3.�󒍌����X�g�쐬��
'4.�t�B���^�����󒍌����Ƃ�
'5.�V�[�g�쐬�ƃf�[�^�R�s�[��
'6.�󒍌���������v����p�^�[���̃V�[�g�𓝍�
'7.�V�[�g����ύX
'7-2.�󒍌�����(�V�[�g���Ɠ��l)��ǉ�
'8.�e�V�[�g�ƌ��V�[�g�ɍ��v�s�̒ǉ�
'9.�F�t��
'10.�t�B���^�����ƕۑ�
'*********************
Sub Main()
Application.ScreenUpdating = False
'*********************
'�ϐ�
'ReDim sumcell_list(0, 0)
'Dim filter_list As New Collection
Set filter_list = New Collection
'�R�[�h
code1 = "870"
div1 = "�D�y"

code2 = "880"
div2 = "���k"

code3 = "825"
div3 = "�֓�"

code4 = "824"
div4 = "���l"

code5 = "850"
div5 = "����"

code6 = "851"
div6 = "�k��"

code7 = "830"
div7 = "�֐�"

code8 = "860"
div8 = "���l"

code9 = "840"
div9 = "��B"

code10 = "82B"
div10 = "(�{)(F)"

code11 = "82K"
div11 = "(�{)(B)"

'code12 = "82J" Or x = "82L" Or x = "820"
code12_1 = "82J"
code12_2 = "82L"
code12_3 = "820"
div12 = "(�{)(D)"

code13 = "82M"
div13 = "(�{)(�ϐk)"

code14 = "81A"
div14 = "(�{)(M)"

'code15 = "890" Or "891"
code15_1 = "890"
code15_2 = "891"
div15 = "�C�O"

'�󒍌��R�[�h�Q�ƃZ��
jyutyuu_cell = "G2"


'@rem 202002�ړ��Ή�--
'csv�f�[�^
csvfile = "C:\�f�[�^.csv"

'�ۑ��ꏊ
save_dir = "C:\"
'@rem 202002�ړ��Ή�--

'�t�@�C����
save_filename = "���t�惊�X�g.xlsx"
'*********************

'*********************
'1.csv�f�[�^����(�s�v��폜�ƃR�s�[)
'csv�J��
Workbooks.OpenText csvfile, Comma:=True
'csv�ړ�
Call move_csv
'��폜
Call del_col
'�R�s�[
Call csv_copy

'*********************
'2.�����V�[�g�폜
'Call read_data
Call del_sh

'*********************
'�F�t��
Call sh_format

'*********************
'3.�t�B���^�p���X�g�쐬
Call create_filter_list

'---------------------
'---------------------
'�J��Ԃ�
'For i = LBound(filter_list) To UBound(filter_list)
For i = 1 To filter_list.Count
'*********************
'4.�t�B���^
Call jyutyuu_filter(i)

'*********************
'5.�V�[�g�ǉ��A�V�[�g���ύX
Call sh_copy

Next i
'---------------------
'---------------------

'*********************
'6.�󒍌���������v����p�^�[���̃V�[�g�𓝍�
'�����͌Œ�Ƃ���
Call sh_merge

'*********************
'7.�V�[�g����ύX
Call sh_name_set
'7-2.�󒍌�����(�V�[�g���Ɠ��l)��ǉ�
Call jyutyuuname_set

'*********************
'8.�e�V�[�g�ɍ��v�s�ǉ��A���V�[�g�ɂ��󒍌��ǉ�
Call add_sum_each


'*********************
'�ۑ�
Call save

End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'1.csv�f�[�^����(�s�v��폜�ƃR�s�[)
'*********************
'csv�ړ�
'*********************
Sub move_csv()
'
' Macro1 Macro
'
'*********************
'�ϐ�
'�V�[�g
Dim basebook As Workbook
Set basebook = Workbooks("���t�惊�X�g.xlsm")
Dim base_sh As Worksheet
Set base_sh = basebook.Sheets("�S��")
'*********************

'
    Windows("�c�m�f�[�^.csv").Activate
    Sheets("�c�m�f�[�^").Select
'    Sheets("�c�m�f�[�^").Move After:=Workbooks("���t�惊�X�g.xlsm").Sheets(15)
    Sheets("�c�m�f�[�^").Move After:=base_sh
End Sub
'*********************
'�V�[�g�ǉ��Ɩ��̕ύX
'*********************
Public Sub csv_copy()

'*********************
'�ϐ�
'�V�[�g
Dim target_sh As Worksheet
Set target_sh = Sheets("�S��")
Dim base_sh As Worksheet
'*********************

'*********************
'�V�[�g�ǉ�
Set base_sh = ActiveSheet
'�l�\��t��
target_sh.UsedRange.ClearContents
target_sh.UsedRange.Value = Empty
target_sh.UsedRange.Interior.ColorIndex = 0
base_sh.Range("A1").CurrentRegion.Copy
target_sh.Range("A1").PasteSpecial Paste:=xlPasteValues

End Sub
'*********************
'����������č폜
'*********************
Public Sub del_col()
'*********************
'�ϐ�
col1 = "�c���m�F���w��敪"
col2 = "�������ЃR�[�h"
col3 = "�󒍌�����"
col4 = "�J���p�j�[����"
'*********************


'*********************
'�폜
'�c���m�F���w��敪
If ((Cells.Find(col1, LookAt:=xlWhole) Is Nothing) = False) Then
zandaka_col = Cells.Find(col1, LookAt:=xlWhole).Column
Columns(zandaka_col).Delete
End If

'�������ЃR�[�h
If ((Cells.Find(col2, LookAt:=xlWhole) Is Nothing) = False) Then
tyuumon_col = Cells.Find(col2, LookAt:=xlWhole).Column
Columns(tyuumon_col).Delete
End If

'�󒍌�����
If ((Cells.Find(col3, LookAt:=xlWhole) Is Nothing) = False) Then
jyutyuumoto_col = Cells.Find(col3, LookAt:=xlWhole).Column
Columns(jyutyuumoto_col).Delete
End If

'�J���p�j�[����
If ((Cells.Find(col4, LookAt:=xlWhole) Is Nothing) = False) Then
company_col = Cells.Find(col4, LookAt:=xlWhole).Column
Columns(company_col).Delete
End If


End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'2.�����V�[�g�폜
'*********************
'�V�[�g�폜
'*********************
Public Sub del_sh()
'*********************
'�ϐ�
'�V�[�g
Dim base_sh As Worksheet
Set base_sh = Sheets("�S��")
    Dim mySht As Worksheet
'*********************
base_sh.Activate
    
    With Application
        '�x����m�F�̃��b�Z�[�W���\���ɐݒ�
        .DisplayAlerts = False
        '�V�[�g�����`�F�b�N���āA�A�N�e�B�u�V�[�g�łȂ���΍폜
        For Each mySht In Worksheets
            If mySht.Name <> ActiveSheet.Name Then mySht.Delete
        Next
        '�ݒ�����ɖ߂�
        .DisplayAlerts = True
    End With
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'3.�t�B���^�p���X�g�쐬
'*********************
'�t�B���^�p���X�g�쐬
'*********************
Public Function create_filter_list()
'*********************
'�ϐ�
'col1 = "�󒍌�"
col1 = "�󒍌��R�[�h"
col2 = "���������v�c��"
Dim i As Long
startrow = 2
'�ŏI�s
If ((Cells.Find(col2, LookAt:=xlWhole) Is Nothing) = False) Then
zandaka_col = Cells.Find(col2, LookAt:=xlWhole).Column
End If
zandaka_col_char = Chr(zandaka_col + 64)
maxrow = Cells(Rows.Count, zandaka_col).End(xlUp).Row

'��Ɨ�
'�ŏI��v���X1
maxcol = Cells(1, Columns.Count).End(xlToLeft).Column + 1
maxcol_char = Chr(maxcol + 64)

'���X�g�擾�p
If ((Cells.Find(col1, LookAt:=xlWhole) Is Nothing) = False) Then
jyutyuumoto_col = Cells.Find(col1, LookAt:=xlWhole).Column
jyutyuumoto_col_char = Chr(jyutyuumoto_col + 64)
End If

'Set filter_list = New Collection
'*********************
'unique�쐬
ActiveSheet.Range("$" & jyutyuumoto_col_char & ":$" & jyutyuumoto_col_char & "").AdvancedFilter Action:=xlFilterCopy, _
    CopyToRange:=Range("$" & maxcol_char & "1"), _
    Unique:=True

'���X�g��
'�f�[�^��o�^����ԁA�G���[�𖳎�����
maxrow_list = Cells(Rows.Count, maxcol).End(xlUp).Row
On Error Resume Next

For i = startrow To maxrow_list
filter_list.Add Cells(i, maxcol).Value
Next i
On Error GoTo 0

'��Ɨ�폜
ActiveSheet.Range("$" & maxcol_char & ":$" & maxcol_char & "").Delete
'*********************


End Function
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'4.�t�B���^
'*********************
'����������ăt�B���^
'*********************
Public Sub jyutyuu_filter(i)
'*********************
'�ϐ�
'�V�[�g
Dim base_sh As Worksheet
Set base_sh = Sheets("�S��")

'col1 = "�󒍌�"
col1 = "�󒍌��R�[�h"
col2 = "���������v�c��"
startrow = 2

'��擾
If ((Cells.Find(col1, LookAt:=xlWhole) Is Nothing) = False) Then
jyutyuumoto_col = Cells.Find(col1, LookAt:=xlWhole).Column
End If
jyutyuumoto_col_char = Chr(jyutyuumoto_col + 64)

'Set filter_list = New Collection
'*********************


'*********************
'�t�B���^
'�󒍌�
base_sh.Range("$" & jyutyuumoto_col_char & ":$" & jyutyuumoto_col_char & "").AutoFilter
base_sh.Range("$" & jyutyuumoto_col_char & ":$" & jyutyuumoto_col_char & "").AutoFilter Field:=1, Criteria1:=filter_list(i)


End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'5.�V�[�g�ǉ��A�V�[�g���ύX
'*********************
'�V�[�g�ǉ��Ɩ��̕ύX
'*********************
Public Sub sh_copy()

'*********************
'�ϐ�
'�V�[�g
Dim base_sh As Worksheet
Set base_sh = Sheets("�S��")
Dim act_sh As Worksheet

col1 = "�󒍌��R�[�h"
If ((Cells.Find(col1, LookAt:=xlWhole) Is Nothing) = False) Then
jyutyuumoto_col = Cells.Find(col1, LookAt:=xlWhole).Column
End If
jyutyuumoto_col_char = Chr(jyutyuumoto_col + 64)
'*********************

'*********************
'�V�[�g�ǉ�
Worksheets.Add Before:=ActiveSheet
Set act_sh = ActiveSheet
base_sh.Range("A1").CurrentRegion.Copy
act_sh.Range("A1").PasteSpecial Paste:=xlPasteColumnWidths
base_sh.Range("A1").CurrentRegion.Copy
base_sh.Range("A1").CurrentRegion.Copy act_sh.Range("A1")
ActiveSheet.Name = Replace(Replace(Trim(act_sh.Range(jyutyuumoto_col_char & "2").Value), " ", ""), "�@", "")

End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'6.�󒍌���������v����p�^�[���̃V�[�g�𓝍�
'*********************
'�V�[�g����
'�ǂꂩ1�̃V�[�g�ɓ�������
'*********************
Public Sub sh_merge()
'*********************
'�ϐ�
col1 = "�󒍌��R�[�h"
col2 = "���������v�c��"
'�V�[�g
Dim base_sh As Worksheet
Dim merge_sh As Worksheet
Dim sheetName As String
'�ŏI�s
If ((Cells.Find(col2, LookAt:=xlWhole) Is Nothing) = False) Then
zandaka_col = Cells.Find(col2, LookAt:=xlWhole).Column
End If
zandaka_col_char = Chr(zandaka_col + 64)
maxrow = Cells(Rows.Count, zandaka_col).End(xlUp).Row


'�󒍌��R�[�h�Q�ƃZ��
'jyutyuu_cell = "G2"
'*********************

'---------------------
'---------------------
'�J��Ԃ�
'For i = 1 To Worksheets.Count - 1
For i = Worksheets.Count To 1 Step -1
'*********************
sheetName = Worksheets(i).Name

Set merge_sh = Sheets(i)
merge_maxrow = merge_sh.Cells(Rows.Count, zandaka_col).End(xlUp).Row

'�l�擾�Ƌ󔒍폜
'x = Replace(Replace(Trim(merge_sh.Range(jyutyuu_cell).Value), " ", ""), "�@", "")
x = Replace(Replace(Trim(sheetName), " ", ""), "�@", "")
If x = "82L" Or x = "820" Then
'x = "82J"
Set base_sh = Sheets("82J")
maxrow = base_sh.Cells(Rows.Count, zandaka_col).End(xlUp).Row

merge_sh.Activate
merge_sh.Range("A1").CurrentRegion.Select
Selection.Offset(1, 0).Select
Selection.Resize(Selection.Rows.Count - 1).Select
Selection.Copy base_sh.Range("A" & maxrow + 1&)

'�V�[�g�폜
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

'�V�[�g�폜
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
'7.�V�[�g����ύX
'*********************
'�V�[�g���̂��R�[�h���疼�̂֕ύX
'*********************
Public Sub sh_name_set()

'*********************
'�ϐ�
'�V�[�g
Dim base_sh As Worksheet
Set base_sh = Sheets("�S��")
Dim rename_sh As Worksheet
Dim sheetName As String
'�󒍌��R�[�h�Q�ƃZ��
'jyutyuu_cell = "G2"
'*********************

'---------------------
'---------------------
'�J��Ԃ�
For i = 1 To Worksheets.Count
'*********************
Set rename_sh = Worksheets(i)
sheetName = Worksheets(i).Name
'�l�擾�Ƌ󔒍폜
'x = Replace(Replace(Trim(rename_sh.Range(jyutyuu_cell).Value), " ", ""), "�@", "")
x = Replace(Replace(Trim(sheetName), " ", ""), "�@", "")

'�V�[�g��
'div1 = "�D�y"
If x = code1 Then
rename_sh.Name = div1

'div2 = "���k"
ElseIf x = code2 Then
rename_sh.Name = div2

'div3 = "�֓�"
ElseIf x = code3 Then
rename_sh.Name = div3

'div4 = "���l"
ElseIf x = code4 Then
rename_sh.Name = div4

'div5 = "����"
ElseIf x = code5 Then
rename_sh.Name = div5

'div6 = "�k��"
ElseIf x = code6 Then
rename_sh.Name = div6

'div7 = "�֐�"
ElseIf x = code7 Then
rename_sh.Name = div7

'div8 = "���l"
ElseIf x = code8 Then
rename_sh.Name = div8

'div9 = "��B"
ElseIf x = code9 Then
rename_sh.Name = div9

'div10 = "(�{)(F)"
ElseIf x = code10 Then
rename_sh.Name = div10

'div11 = "(�{)(B)"
ElseIf x = code11 Then
rename_sh.Name = div11

'code12_1 = "82J"
'code12_2 = "82L"
'code12_3 = "820"
'div12 = "(�{)(D)"
ElseIf x = code12_1 Or x = code12_2 Or x = code12_3 Then
rename_sh.Name = div12

'div13 = "(�{)(�ϐk)"
ElseIf x = code13 Then
rename_sh.Name = div13

'div14 = "(�{)(M)"
ElseIf x = code14 Then
rename_sh.Name = div14

'code15_1 = "890"
'code15_2 = "891"
'div15 = "�C�O"
ElseIf x = code15_1 Or x = code15_2 Then
rename_sh.Name = div15

End If

Next i
'---------------------
'---------------------

End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'7-2.�󒍌�����(�V�[�g���Ɠ��l)��ǉ�
'*********************
'�V�[�g���̂��R�[�h���疼�̂֕ύX
'*********************
Public Sub jyutyuuname_set()

'*********************
'�ϐ�
'�V�[�g
Dim sheetName As String
Dim target_sheet As Worksheet
'�ϐ�
col1 = "�󒍌��R�[�h"
col2 = "���������v�c��"
startrow = 2
'�ŏI�s
If ((Cells.Find(col2, LookAt:=xlWhole) Is Nothing) = False) Then
zandaka_col = Cells.Find(col2, LookAt:=xlWhole).Column
End If
zandaka_col_char = Chr(zandaka_col + 64)
maxrow = Cells(Rows.Count, zandaka_col).End(xlUp).Row

'*********************

'---------------------
'---------------------
'�J��Ԃ�
For i = 1 To Worksheets.Count
'*********************
Set target_sheet = Worksheets(i)
sheetName = Worksheets(i).Name
target_sheet.Activate
maxrow = target_sheet.Cells(Rows.Count, zandaka_col).End(xlUp).Row

'�t�B���^����
If target_sheet.AutoFilterMode = True Then
target_sheet.Cells.AutoFilter
End If

'�l�擾�Ƌ󔒍폜
x = Replace(Replace(Trim(sheetName), " ", ""), "�@", "")

'��ǉ�
target_sheet.Columns(1).Insert

'���ږ�(�󒍐於)���Z�b�g
target_sheet.Cells(1, 1) = "�󒍐於"

'�V�[�g��(�󒍐於)���Z�b�g
target_sheet.Range(Cells(startrow, 1), Cells(maxrow, 1)).Value = sheetName

'�����R�s�y
target_sheet.Range(Cells(1, 2), Cells(maxrow, 2)).Copy
target_sheet.Range(Cells(1, 1), Cells(maxrow, 1)).PasteSpecial Paste:=xlPasteFormats

Next i
'---------------------
'---------------------

End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'8.�e�V�[�g�ɍ��v�s�ǉ��A���V�[�g�ɂ��󒍌��ǉ�
'*********************
'�ŏI�s���擾���č��v�ǉ�
'*********************
Public Sub add_sum_each()
'*********************
'*********************
'�V�[�g
Dim base_sh As Worksheet
Set base_sh = Sheets("�S��")
Dim target_sheet As Worksheet
'�ϐ�
'col1 = "�󒍌�"
col1 = "�󒍌��R�[�h"
col2 = "���������v�c��"
startrow = 2
'�ŏI�s
If ((Cells.Find(col2, LookAt:=xlWhole) Is Nothing) = False) Then
zandaka_col = Cells.Find(col2, LookAt:=xlWhole).Column
End If
zandaka_col_char = Chr(zandaka_col + 64)
maxrow = Cells(Rows.Count, zandaka_col).End(xlUp).Row

'���V�[�g
base_maxrow = base_sh.Cells(Rows.Count, zandaka_col).End(xlUp).Row

'*********************


'---------------------
'---------------------
'�J��Ԃ�
For i = 1 To Worksheets.Count
'*********************
Set rename_sh = Worksheets(i)
sheetName = Worksheets(i).Name
Set target_sheet = Worksheets(i)
maxrow = target_sheet.Cells(Rows.Count, zandaka_col).End(xlUp).Row

If (sheetName <> base_sh.Name) Then
'���v�ǉ�
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
'9.�F�t��
'*********************
'*********************
Public Sub sh_format()
'*********************
'�ϐ�
col1 = "�m�F����"
'�V�[�g
Dim target_sh As Worksheet
'�ŏI�s
If ((Cells.Find(col1, LookAt:=xlWhole) Is Nothing) = False) Then
target_col = Cells.Find(col1, LookAt:=xlWhole).Column
End If
target_col_char = Chr(target_col + 64)
maxrow = Cells(Rows.Count, target_col).End(xlUp).Row
maxcol = Cells(1, Columns.Count).End(xlToLeft).Column

target = "-1-"

'�󒍌��R�[�h�Q�ƃZ��
'jyutyuu_cell = "G2"
'*********************

'---------------------
'---------------------
'�J��Ԃ�
For i = Worksheets.Count To 1 Step -1
'*********************
Set target_sh = Worksheets(i)
maxrow = target_sh.Cells(Rows.Count, target_col).End(xlUp).Row

For j = 1 To maxrow
'�F�ύX
If Cells(j, target_col) Like "*" & target & "*" Then
Range(Cells(j, 1), Cells(j, maxcol)).Interior.Color = 11389944
End If
Next j

Next i
'---------------------
'---------------------

End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'10.�t�B���^�����ƕۑ�
'*********************
'*********************
Public Sub save()
'*********************
'*********************
'�V�[�g
Dim base_sh As Worksheet
Set base_sh = Sheets("�S��")
'�ϐ�
'�t�@�C����
save_filename = "���t�惊�X�g.xlsx"
'*********************

'�t�B���^����
base_sh.Cells.AutoFilter

'�ۑ�
Application.DisplayAlerts = False
ActiveWorkbook.SaveAs Filename:=save_dir & save_filename, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
Application.DisplayAlerts = True

End Sub
