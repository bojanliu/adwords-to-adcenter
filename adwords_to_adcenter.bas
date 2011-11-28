Attribute VB_Name = "adwords_to_adcenter"
Option Explicit

Sub adwords_to_adcenter()
    Dim row_count As Long
    Dim i As Long
    Dim Prefix As Long
    Dim Suffix As Long
    Dim Slash As Long
    Dim d_url As String
    Dim w_url As String
    Workbooks.Open ("原始数据.csv")
    With Workbooks("原始数据.csv").Sheets(1) '删除不导入的列
        .Range("b:b,c:c,e:e,f:f,h:h,i:i,j:j,k:k,l:l,o:o,q:q,r:r,v:v,w:w,x:x,y:y,z:z,ag:ag,ah:ah,ai:ai,aj:aj,ak:ak").Delete
        row_count = .[a1048576].End(xlUp).Row
        Application.ScreenUpdating = False
        For i = row_count To 2 Step -1 '删除不导入的行
            If .Cells(i, 11) <> "" And .Cells(i, 3) = "" Or _
                .Cells(i, 13) <> "" And .Cells(i, 14) = "" And .Cells(i, 4) = "" Then
                .Rows(i).Delete
            End If
        Next i
        row_count = .[a1048576].End(xlUp).Row
        For i = row_count To 2 Step -1
            d_url = .Cells(i, 11)
            If d_url <> "" Then '清楚目标网址参数
                Prefix = InStr(1, d_url, "l=h") '前缀
                Suffix = InStr(1, d_url, "?g") '后缀
                If Prefix <> 0 Then
                    .Cells(i, 11) = Mid(d_url, 37, Suffix - 37)
                Else
                    .Cells(i, 11) = Left(d_url, Suffix - 1)
                End If
            End If
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
    Workbooks("原始数据.csv").Save
    Workbooks("原始数据.csv").Close True '关闭提示
    MsgBox ("操作完成")
End Sub

