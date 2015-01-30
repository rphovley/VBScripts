Sub EmailReps(ByVal repName As String, ByVal repEmail As String)
'Working in Excel 2000-2013
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim Sourcewb As Workbook
    Dim Destwb As Workbook
    Dim repList As Worksheet
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim OutApp As Object
    Dim OutMail As Object

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    Set Sourcewb = ActiveWorkbook
    Set repList = Sourcewb.Sheets("Reps")
    'Copy the ActiveSheet to a new workbook
    ActiveSheet.Copy
    Set Destwb = ActiveWorkbook

    'Determine the Excel version and file extension/format
    With Destwb
        If Val(Application.Version) < 12 Then
            'You use Excel 97-2003
            FileExtStr = ".xls": FileFormatNum = -4143
        Else
            'You use Excel 2007-2013
            Select Case Sourcewb.FileFormat
            Case 51: FileExtStr = ".xlsx": FileFormatNum = 51
            Case 52:
                If .HasVBProject Then
                    FileExtStr = ".xlsm": FileFormatNum = 52
                Else
                    FileExtStr = ".xlsx": FileFormatNum = 51
                End If
            Case 56: FileExtStr = ".xls": FileFormatNum = 56
            Case Else: FileExtStr = ".xlsb": FileFormatNum = 50
            End Select
        End If
    End With

    '    'Change all cells in the worksheet to values if you want
    '    With Destwb.Sheets(1).UsedRange
    '        .Cells.Copy
    '        .Cells.PasteSpecial xlPasteValues
    '        .Cells(1).Select
    '    End With
    '    Application.CutCopyMode = False

    'Save the new workbook/Mail it/Delete it
    TempFilePath = Environ$("temp") & "\"
    TempFileName = repName & " " & Format(Now, "dd-mmm-yy")

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With Destwb
        .SaveAs TempFilePath & TempFileName & FileExtStr, FileFormat:=FileFormatNum
        On Error Resume Next
        With OutMail
            .To = repEmail
            .CC = ""
            .BCC = ""
            .Subject = "Overrides for new pay structure"
            .Body = "Evolve Team," & vbNewLine & vbNewLine & _
            "This email contains a breakdown of jobs that have now reached a 100% payable status for the reps you receive overrides on (according to the new pay structure). " & _
            "Any pay shown in the breakdown will be factored into this week's paycheck. The process of paying out the overrides related to the new pay structure will be paid out over the course of the next 3 months.  " & _
            "So only a third of what is shown in the breakdown will affect this week's paycheck. " & vbNewLine & vbNewLine & _
            "Let us know if you have any quesitons," & vbNewLine & vbNewLine & _
            "Thanks!" & vbNewLine & vbNewLine & _
            "Evolve Finance Team" & vbNewLine & vbNewLine & _
            "http://goo.gl/forms/j0fQXnf5ov"


            .Attachments.Add Destwb.FullName
            'You can add other files also like this
            '.Attachments.Add ("C:\test.txt")
            .Send   'or use .Display
        End With
        On Error GoTo 0
        .Close savechanges:=False
    End With

    'Delete the file you have send
    Kill TempFilePath & TempFileName & FileExtStr

    Set OutMail = Nothing
    Set OutApp = Nothing

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub
