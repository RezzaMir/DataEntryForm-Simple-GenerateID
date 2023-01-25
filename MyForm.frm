VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MyForm 
   Caption         =   "Data Entry Form"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9810
   OleObjectBlob   =   "MyForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Data")
Dim last_Row As Long
last_Row = Application.WorksheetFunction.CountA(sh.Range("A:A"))

'Validation--------------------------------------------------------------
If Me.ComboBox1.Value = "" Then
MsgBox "Please select the Title", vbCritical
Exit Sub
End If

If Me.TextBox1.Value = "" Then
MsgBox "Please enter the Name", vbCritical
Exit Sub
End If

If Me.TextBox2.Value = "" Then
MsgBox "Please enter the Email", vbCritical
Exit Sub
End If

If Me.TextBox3.Value = "" Then
MsgBox "Please enter the Phone", vbCritical
Exit Sub
End If

'------------------------------------------------------------------------
sh.Range("A" & last_Row + 1).Value = "=Row()+5088"
sh.Range("B" & last_Row + 1).Value = Me.ComboBox1.Value
sh.Range("C" & last_Row + 1).Value = Me.TextBox1.Value
sh.Range("D" & last_Row + 1).Value = Me.TextBox2.Value
sh.Range("E" & last_Row + 1).Value = Me.TextBox3.Value
sh.Range("F" & last_Row + 1).Value = Now
'------------------------------------------------------------------------

Me.ComboBox1.Value = ""
Me.TextBox1.Value = ""
Me.TextBox2.Value = ""
Me.TextBox3.Value = ""

Call Refresh_Data

End Sub

Private Sub cmdClear_Click()
Me.ComboBox1.Value = ""
Me.TextBox1.Value = ""
Me.TextBox2.Value = ""
Me.TextBox3.Value = ""
Me.TextBox4.Value = ""

End Sub

Private Sub cmdDelete_Click()

If Me.TextBox4.Value = "" Then
MsgBox "Select the record to delete"
Exit Sub
End If

Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Data")
Dim Selected_Row As Long
Selected_Row = Application.WorksheetFunction.Match(CLng(Me.TextBox4.Value), sh.Range("A:A"), 0)
'----------------------------------------------
sh.Range("A" & Selected_Row).EntireRow.Delete
'----------------------------------------------
Me.ComboBox1.Value = ""
Me.TextBox1.Value = ""
Me.TextBox2.Value = ""
Me.TextBox3.Value = ""
Me.TextBox4.Value = ""

Call Refresh_Data

End Sub

Private Sub cmdSave_Click()
ThisWorkbook.Save
MsgBox " Data Saved"
End Sub

Private Sub cmdUpdate_Click()
If Me.TextBox4.Value = "" Then
MsgBox "Select the record to update"
Exit Sub
End If

Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Data")
Dim Selected_Row As Long
Selected_Row = Application.WorksheetFunction.Match(CLng(Me.TextBox4.Value), sh.Range("A:A"), 0)

'Validation--------------------------------------------------------------
If Me.ComboBox1.Value = "" Then
MsgBox "Please select the Title", vbCritical
Exit Sub
End If

If Me.TextBox1.Value = "" Then
MsgBox "Please enter the Name", vbCritical
Exit Sub
End If

If Me.TextBox2.Value = "" Then
MsgBox "Please enter the Email", vbCritical
Exit Sub
End If

If Me.TextBox3.Value = "" Then
MsgBox "Please enter the Phone", vbCritical
Exit Sub
End If

'------------------------------------------------------------------------
sh.Range("B" & Selected_Row).Value = Me.ComboBox1.Value
sh.Range("C" & Selected_Row).Value = Me.TextBox1.Value
sh.Range("D" & Selected_Row).Value = Me.TextBox2.Value
sh.Range("E" & Selected_Row).Value = Me.TextBox3.Value
sh.Range("F" & Selected_Row).Value = Now
'------------------------------------------------------------------------

Me.ComboBox1.Value = ""
Me.TextBox1.Value = ""
Me.TextBox2.Value = ""
Me.TextBox3.Value = ""
Me.TextBox4.Value = ""

Call Refresh_Data

End Sub

'--------------By doble clock on a row, it will bring it up ----------------

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.TextBox4.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 0)
Me.ComboBox1.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 1)
Me.TextBox1.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 2)
Me.TextBox2.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 3)
Me.TextBox3.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 4)

'---------------------------------------------------------------------------


End Sub

Private Sub UserForm_Activate()
With Me.ComboBox1
        .Clear
        .AddItem ""
        .AddItem "Mr."
        .AddItem "Mrs."
End With
Call Refresh_Data
End Sub

Sub Refresh_Data()

Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Data")
Dim last_Row As Long
last_Row = Application.WorksheetFunction.CountA(sh.Range("A:A"))

With Me.ListBox1
        .ColumnHeads = True
        .ColumnCount = 6
        .ColumnWidths = "30,40,100,110,70,90"
        
        If last_Row = 1 Then
        .RowSource = "Data!A2:F2"
        Else
        .RowSource = "Data!A2:F" & last_Row
        End If
        
End With
        
End Sub
