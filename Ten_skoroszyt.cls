VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ten_skoroszyt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public cn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rs2 As New ADODB.Recordset
Public strsQL$
Sub DwnldData()

    Dim fso As Scripting.FileSystemObject
    Dim fsoFolder As Object
    Dim fsoFile As Object
    Dim edgeURL As String
    Dim downloadsPath As String
    Dim edgePath As String
    Dim destFolderName
    Dim i As Integer
    Dim j As Integer
    
    Dim currYear As Integer
    Dim ws As Worksheet
    
    Dim cmbBox As ComboBox
    Set cmbBox = ThisWorkbook.Sheets("FORM").ComboBox1
    
    ' Get selected years ago value
    yearsAgo = CInt(cmbBox.Value)
    
    ' Remove old reports
    If ThisWorkbook.Sheets.Count > 1 Then
        For j = ThisWorkbook.Sheets.Count To 2 Step -1
        Application.DisplayAlerts = False
        
            ThisWorkbook.Sheets(j).Delete
            
        Application.DisplayAlerts = True
        Next j
    End If
    ' Set destination folder name
    destFolderName = Format(DateSerial(Year(Now), Month(Now), Day(Now)), "ddMMyyyy") & Format(TimeSerial(Hour(Now), Minute(Now), Second(Now)), "hhmmss")
    
    Set fso = New Scripting.FileSystemObject
  
    ' Download files for all years
    For j = 0 To yearsAgo
    
    currYear = Year(Date) - j
    
    ' Set the URL from which you want to download file
    edgeURL = "https://static.nbp.pl/dane/kursy/Archiwum/archiwum_tab_a_" & CStr(currYear) & ".csv"
         
    ' Set the destination path

    destinationPath = "C:\Users\magda\Documents\testy\" & destFolderName
    
    If Not fso.FolderExists(destinationPath) Then
    
    fso.CreateFolder destinationPath
    
    End If
    
    destinationFilePath = destinationPath & "\" & archiwum_tab_a_ & CStr(currYear) & ".csv"
    
    If URLDownloadToFile(0, edgeURL, destinationFilePath, 0, 0) Then
    Debug.Print "downloaded " & destinationFilePath
    End If
    
    Set ws = ThisWorkbook.Sheets.Add(after:=ThisWorkbook.Sheets(1))
    ws.Name = CStr(currYear)
    
    Dim csv As String
    Dim line As String
    
    i = 1
    csv = destinationFilePath
    
    Open csv For Input As #1
    
    ' Read downloaded file
    Do Until EOF(1)
        Line Input #1, line
        ws.Cells(i, 1).Value = line
        i = i + 1
    Loop
    
    Close #1
    
    ' Split csv data to columns
    ws.Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
    Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar:=Application.International(xlListSeparator)
    ws.Visible = xlSheetHidden
    
    Next j

End Sub

Sub prprReport()

Dim i, j, k, l As Integer
Dim Rng As Range
Dim n As Name

Application.ScreenUpdating = False

Call DwnldData

' Delete names from the workbook(name manager)

For Each n In ThisWorkbook.Names
    n.Delete
Next

' Prepare data
For i = 2 To ThisWorkbook.Sheets.Count

' Remove 2nd row
    With ThisWorkbook.Sheets(i)
    .Rows("2:2").Delete
    
' Count columns
    j = 1
    Do While IsEmpty(.Cells(1, j)) = False
        j = j + 1
    Loop
    
' Count rows
    k = 1
    Do While IsEmpty(.Cells(k, 1)) = False
        k = k + 1
    Loop

' Remove last few rows
    .Rows(CStr(k) + ":" + CStr(k + 5)).Delete
    
    k = k - 1
    j = j - 1

' Prepare column 'data' with date values
    For l = 2 To k
        .Cells(l, 1).Value = DateSerial(CInt(Left(CStr(.Cells(l, 1).Value), 4)), CInt(Mid(CStr(.Cells(l, 1).Value), 5, 2)), CInt(Right(CStr(.Cells(l, 1).Value), 2)))
    Next l

' Remove numbers from column headers(for SQL queries)
    For l = 1 To j
    .Cells(1, l).Value = Replace(.Cells(1, l).Value, "1", "")
    .Cells(1, l).Value = Replace(.Cells(1, l).Value, "0", "")
    Next l
    
' Prepare datatables (for SQL queries)
    Set Rng = .Range("A1:" & .Cells(k, j).Address)
    nm = "tab" & CStr(i) & "sql"

    ThisWorkbook.Names.Add Name:=nm, RefersTo:=Rng

End With
Next i

Call prpr_charts

Application.ScreenUpdating = True

End Sub

Sub prpr_charts()
Dim strFile As String
Dim i As Integer
Dim lbx As ListBox
Dim curr, t As String
Dim sc As Integer
Dim ch As Object
Dim r As Integer

Set lbx = ThisWorkbook.Sheets("FORM").ListBoxes("currListPicked")
strFile = ThisWorkbook.Path & "\" & ThisWorkbook.Name
Set cn = New Connection
Set rs = New Recordset
cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFile & ";Extended Properties=""Excel 12.0 macro;HDR=Yes;IMEX=1"";"
cn.CursorLocation = adUseClient
sc = ThisWorkbook.Sheets.Count

For j = 1 To lbx.ListCount

    curr = lbx.List(j)

    For i = 2 To sc
        
        t = "tab" & CStr(i) & "sql"
        
        If i = 2 Then
            strsQL = "SELECT data," & curr & " FROM " & t
        Else
            strsQL = strsQL & " UNION ALL SELECT data," & curr & " FROM " & t
        End If
        
    Next i
    
    cn.Open
    Set rs = New Recordset
    rs.Open strsQL, cn
    
    Set rs2 = rs.Clone
    Set ws = ThisWorkbook.Sheets.Add(after:=ThisWorkbook.Sheets(1))
    ws.Name = curr
    ws.Cells(1, 1).Value = rs2.Fields(0).Name
    ws.Cells(1, 2).Value = rs2.Fields(1).Name
    ws.Range("A2").CopyFromRecordset rs2
    
    rs.Close
    rs2.Close
    cn.Close
    Close Connection
    
    
    With ws
    r = 1
    Do While IsEmpty(.Cells(r, 1)) = False
    r = r + 1
    Loop
    r = r - 1
    .Range("A2:A" & r).NumberFormat = "dd.mm.yyyy"
    
   Set ch = .Shapes.AddChart2(227, xlLineStacked)
'    ch.SetSourceData Source:=Range(ws.Name & "!$A$1:$B$" & r)
    End With
Next j



End Sub


Private Sub Workbook_Open()
Dim i As Integer
For i = 1 To 12
    ThisWorkbook.Sheets("FORM").ComboBox1.AddItem CStr(i)
Next i
End Sub
