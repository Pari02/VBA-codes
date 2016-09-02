' Author = Parikshita Tripathi
' Date = 07/27/2016
' Comments = code to hide/unhide specific worksheets based on their current visibility

Sub HideUnhide_Click()

' define the workseet variable
Dim ws As Worksheet

' Run a loop for each sheet name
' check the current visbility and change it to other
For Each ws In ActiveWorkbook.Sheets(Array("Data", "Wo Data"))
    If ws.Visible = True Then
        ws.Visible = False
    Else
        ws.Visible = True
    End If
Next

End Sub