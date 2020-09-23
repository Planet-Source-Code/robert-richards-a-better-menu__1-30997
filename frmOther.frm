VERSION 5.00
Begin VB.Form frmOther 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Other form"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResult 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label lblFile 
      Caption         =   " &File"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents MenuForm As frmMenu
Attribute MenuForm.VB_VarHelpID = -1

Private Sub Form_Unload(Cancel As Integer)
    TerminateMenu
End Sub

Private Sub Form_Click()
    TerminateMenu
End Sub

Private Sub txtResult_Click()
    TerminateMenu
End Sub

Private Sub lblFile_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
    'align it when it pops up
    Set MenuForm = New frmMenu
    Call MenuForm.EnableItems(True, True, False, False, False)
    lblFile.BorderStyle = 1
    MenuForm.Show , Me
    MenuForm.Left = Me.Left + 60
    MenuForm.Top = Me.Top + 600
    Me.SetFocus
End Sub

'****************************************** Menu form event handlers
'the code to handle the menu items selected is here, where it should be
Private Sub MenuForm_NewClick()
    txtResult.Text = "Open a new file and unload the menuform"
    TerminateMenu
End Sub

Private Sub MenuForm_OpenClick()
    txtResult.Text = "Open a file and unload the menuform"
    TerminateMenu
End Sub

Private Sub MenuForm_PrintClick()
    txtResult.Text = "Print and unload the menuform"
    TerminateMenu
End Sub

Private Sub MenuForm_SaveAsClick()
    txtResult.Text = "Save file As and unload the menuform"
    TerminateMenu
End Sub

Private Sub MenuForm_SaveClick()
    txtResult.Text = "Save file and unload the menuform"
    TerminateMenu
End Sub

Private Sub TerminateMenu()
    If Not MenuForm Is Nothing Then
        Unload MenuForm
        Set MenuForm = Nothing
    End If
    lblFile.BorderStyle = 0
End Sub

