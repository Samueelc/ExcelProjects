Attribute VB_Name = "AutomaticTableofContents"
Option Explicit

Sub Auto_Table_Contents()

    Dim StartCell As Range 'for inputbox to select range
    Dim EndCell As Range 'for message box as info
    Dim Sh As Worksheet
    Dim ShName As String
    Dim MsgConfirm As VBA.VbMsgBoxResult 'for message box to confirm
    

    On Error Resume Next
     
    Set StartCell = Excel.Application.InputBox("Where do you want to insert the Table of Contents?" _
    & vbNewLine & "Please select a cell:", "Insert table of contents", , , , , , 8)
    
    If Err.Number = 424 Then Exit Sub
    On Error GoTo Handle
    Set StartCell = StartCell.Cells(1, 1)
    Set EndCell = StartCell.Offset(Worksheets.Count - 2, 1)
    
    'get confirmation
    MsgConfirm = VBA.MsgBox("The value in cells:" & vbNewLine & StartCell.Address & " to " & EndCell.Address & _
    " could be overwritten." & vbNewLine & " Would you like to continue?", vbOKCancel + vbDefaultButton2, "Confirmation required!")
    If MsgConfirm = vbCancel Then Exit Sub
    
    For Each Sh In Worksheets
        ShName = Sh.Name
        If ActiveSheet.Name <> ShName Then
            If Sh.Visible = xlSheetVisible Then
                ActiveSheet.Hyperlinks.Add Anchor:=StartCell, Address:="", SubAddress:= _
                     "'" & ShName & "'!A1", TextToDisplay:=ShName
                StartCell.Offset(0, 1).Value = Sh.Range("a1").Value
                Set StartCell = StartCell.Offset(1, 0)
            End If 'sheet is invisible
        End If 'sheet is not de activesheet
    Next Sh
    Exit Sub
    
Handle:
MsgBox "Unfortunately, an error has ocurred!"

    
End Sub
 

