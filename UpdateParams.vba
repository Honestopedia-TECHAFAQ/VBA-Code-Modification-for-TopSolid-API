Sub UpdateParams()
    On Error GoTo ErrorHandler

    Dim myApp As Object
    Dim myDoc As Object
    Dim ws As Worksheet
    Dim ParaNameTS As String
    Dim ParaValueTS As Variant
    Dim i As Integer
    Dim param As Object
    Dim logMsg As String
    logMsg = "Updating parameters:" & vbCrLf
    Set myApp = CreateObject("TopSolid.Application")
    Set myDoc = myApp.ActiveDocument

    If myDoc Is Nothing Then
        MsgBox "No document is open or TopSolid is not running."
        Exit Sub
    End If
    myApp.SynchronizeAuto = True 
    myApp.Visible = True 
    myDoc.ProgramUnitLength = "mm"
    Set ws = ThisWorkbook.Sheets(1) 
    i = 2 

    Do While Not IsEmpty(ws.Cells(i, "A").Value)
        ParaNameTS = CStr(ws.Cells(i, "A").Value) 
        ParaValueTS = ws.Cells(i, "B").Value

        If ParaValueTS <> "" Then
            Set param = Nothing 
            On Error Resume Next 
            Set param = myDoc.Parameters.Item(ParaNameTS)
            On Error GoTo ErrorHandler 

            If Not param Is Nothing Then 
                Do
                    On Error Resume Next
                    Set myDoc = myApp.CurrentDocument
                    If IsNumeric(ParaValueTS) Then
                        param.NominalValue = ParaValueTS
                    Else
                        param.NominalValueExpression = ParaValueTS
                    End If
                    param.Element.Basify
                    On Error GoTo ErrorHandler
                    Application.Wait (Now + TimeValue("0:00:01"))
                Loop Until Not myDoc Is Nothing

                myApp.SynchronizeAuto = True 
                logMsg = logMsg & "Parameter " & ParaNameTS & " updated successfully." & vbCrLf
            Else
                logMsg = logMsg & "Parameter " & ParaNameTS & " not found." & vbCrLf
            End If
        End If

        i = i + 1
    Loop
    MsgBox logMsg, vbInformation, "Parameter Update Summary"
    Set myApp = Nothing
    Set myDoc = Nothing
    Set ws = Nothing

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbExclamation
    Resume Next
End Sub
