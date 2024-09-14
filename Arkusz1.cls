VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arkusz1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub listDataAdd()

Dim dataRange As Range
Dim dataList1, dataList2 As ListBox
Dim i As Integer
Dim j As Integer

Dim check As Boolean

Set dataList1 = ThisWorkbook.Sheets("FORM").ListBoxes("currListPick")
Set dataList2 = ThisWorkbook.Sheets("FORM").ListBoxes("currListPicked")


For i = 1 To dataList1.ListCount

    If dataList1.Selected(i) Then
        check = True
        For j = 1 To dataList2.ListCount
            If dataList2.List(j) = dataList1.List(i) Then
                check = False
            End If
        Next j
        
        If check = True Then
            dataList2.AddItem (dataList1.List(i))
        End If
    End If
Next i

End Sub


Sub listDataRemove()
Dim dataList2 As ListBox
Dim i As Integer

Set dataList2 = ThisWorkbook.Sheets("FORM").ListBoxes("currListPicked")

For i = 1 To dataList2.ListCount
    If i <= dataList2.ListCount Then
    If dataList2.Selected(i) Then
        dataList2.RemoveItem (i)
        i = i - 1
    End If
    End If
    
Next i

End Sub

