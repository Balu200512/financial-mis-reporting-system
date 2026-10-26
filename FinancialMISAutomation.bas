\
Attribute VB_Name = "FinancialMISAutomation"
Option Explicit

' Save this module in a macro-enabled Excel workbook (.xlsm).
' Expected worksheets: Transactions, MIS Summary
' Transactions headers in row 1: transaction_id, transaction_date, account_code,
' cost_center_id, description, reference_no, amount, currency, status, prepared_by.
' This is a workbook-side refresh/export helper; configure Power Query connections separately.

Public Sub RefreshAllMIS()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Refreshing Financial MIS data..."

    ThisWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone

    On Error Resume Next
    ThisWorkbook.Worksheets("MIS Summary").Range("B2").Value = Now
    ThisWorkbook.Worksheets("MIS Summary").Range("B2").NumberFormat = "dd-mmm-yyyy hh:mm"
    On Error GoTo ErrHandler

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Financial MIS refresh completed.", vbInformation, "Financial MIS"
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "MIS refresh failed: " & Err.Description, vbExclamation, "Financial MIS"
End Sub

Public Sub ExportMISSummaryToPDF()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Dim filePath As String

    Set ws = ThisWorkbook.Worksheets("MIS Summary")
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Save the workbook before exporting the PDF.", vbExclamation, "Financial MIS"
        Exit Sub
    End If

    filePath = ThisWorkbook.Path & Application.PathSeparator & _
               "Financial_MIS_Summary_" & Format(Now, "yyyymmdd_hhnn") & ".pdf"

    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=filePath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True

    MsgBox "PDF exported to:" & vbCrLf & filePath, vbInformation, "Financial MIS"
    Exit Sub

ErrHandler:
    MsgBox "PDF export failed: " & Err.Description, vbExclamation, "Financial MIS"
End Sub

Public Sub CreateMISSummarySheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("MIS Summary")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "MIS Summary"
    End If

    With ws
        .Range("A1").Value = "Financial MIS Reporting System"
        .Range("A2").Value = "Last refresh:"
        .Range("A4").Value = "KPI"
        .Range("B4").Value = "Value"
        .Range("A5").Value = "Posted Transaction Count"
        .Range("B5").Formula = "=COUNTIF(Transactions!I:I,""Posted"")"
        .Range("A6").Value = "Posted Transaction Amount (INR)"
        .Range("B6").Formula = "=SUMIFS(Transactions!G:G,Transactions!I:I,""Posted"")"
        .Range("A7").Value = "Pending Transaction Count"
        .Range("B7").Formula = "=COUNTIF(Transactions!I:I,""Pending"")"
        .Range("A1:B1").Font.Bold = True
        .Range("A1").Font.Size = 16
        .Range("A4:B4").Font.Bold = True
        .Columns("A:B").AutoFit
    End With

    MsgBox "MIS Summary sheet is ready. Confirm column positions if your source table differs.", vbInformation
End Sub
