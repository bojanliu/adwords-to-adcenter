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
    Workbooks.Open (ThisWorkbook.Path & "\ԭʼ����.xlsx")
    With Workbooks("ԭʼ����.xlsx").Sheets(1) 'ɾ�����������
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
            arr(i) = .Cells(i, 11) 'Ŀ����ַ����

            If arr(i) = "" Then
                If .Cells(i, 13) <> "" Then 'ɾ�����������
                    .Rows(i).Delete
                End If
            Else
                If .Cells(i, 3) = "" Then 'ɾ�����������
                    .Rows(i).Delete
                Else
                    Prefix = InStr(1, arr(i), "?u") 'ǰ׺
                    Suffix = InStr(1, arr(i), "?g") '��׺
                    If Prefix <> 0 Then
                        .Cells(i, 11) = Mid(arr(i), 37, Suffix - 37) 'ȥ����ַ����
                    Else
                        .Cells(i, 11) = Left(arr(i), Suffix - 1) 'ȥ����ַ����
                    End If
                 End If
            End If
        Next i
        
        row_count = .[a1048576].End(xlUp).Row
        For i = row_count To 2 Step -1
            w_url = .Cells(i, 5)
            If w_url <> "" Then 'ȥ��url�е�·��
                Slash = InStr(1, w_url, "/")
                If Slash <> 0 Then
                    .Cells(i, 5) = Left(w_url, Slash - 1)
                End If
            End If
        Next i
        
        Application.ScreenUpdating = False
    End With
    
    Workbooks("ԭʼ����.xlsx").Save
    Workbooks("ԭʼ����.xlsx").Close True '�ر���ʾ
    MsgBox ("�������")
End Sub

