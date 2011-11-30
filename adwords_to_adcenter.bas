Attribute VB_Name = "adwords_to_adcenter"
Option Explicit

Sub adwords_to_adcenter()
    Dim row_count As Long
    Dim i As Long
    Dim Prefix As Long
    Dim Suffix As Long
    Dim Slash As Long
    Dim w_url As String
    Dim arr() As String
    Workbooks.Open (ThisWorkbook.Path & "\原始数据.xlsx")
    With Workbooks("原始数据.xlsx").Sheets(1) '删除不导入的列
        .Columns("ah:al").Delete
        .Columns("w:aa").Delete
        .Columns("q:s").Delete
        .Columns("o:o").Delete
        .Columns("h:l").Delete
        .Columns("e:f").Delete
        .Columns("b:c").Delete
        row_count = .[a1048576].End(xlUp).Row
        Application.ScreenUpdating = False
        ReDim arr(2 To row_count)
        
        For i = row_count To 2 Step -1
            arr(i) = .Cells(i, 11) '目标网址数组

            If arr(i) = "" Then
                If .Cells(i, 13) <> "" Then '删除不导入的行
                    .Rows(i).Delete
                End If
            Else
                If .Cells(i, 3) = "" Then '删除不导入的行
                    .Rows(i).Delete
                Else
                    Prefix = InStr(1, arr(i), "?u") '前缀
                    Suffix = InStr(1, arr(i), "?g") '后缀
                    If Prefix <> 0 Then
                        .Cells(i, 11) = Mid(arr(i), 37, Suffix - 37) '去除网址参数
                    Else
                        .Cells(i, 11) = Left(arr(i), Suffix - 1) '去除网址参数
                    End If
                 End If
            End If
        Next i
        
        row_count = .[a1048576].End(xlUp).Row
        For i = row_count To 2 Step -1
            w_url = .Cells(i, 5)
            If w_url <> "" Then '去除url中的路径
                Slash = InStr(1, w_url, "/")
                If Slash <> 0 Then
                    .Cells(i, 5) = Left(w_url, Slash - 1)
                End If
            End If
        Next i
        
        Application.ScreenUpdating = False
    End With
    
    Workbooks("原始数据.xlsx").Save
    Workbooks("原始数据.xlsx").Close True '关闭提示
    MsgBox ("操作完成")
End Sub

