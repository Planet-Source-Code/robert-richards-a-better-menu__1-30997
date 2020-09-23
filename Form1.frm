VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1815
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   2955
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image4 
      Height          =   255
      Left            =   2520
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   2520
      Picture         =   "Form1.frx":030A
      Stretch         =   -1  'True
      Top             =   1480
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   2520
      Picture         =   "Form1.frx":0614
      Stretch         =   -1  'True
      Top             =   760
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   2520
      Picture         =   "Form1.frx":0EDE
      Stretch         =   -1  'True
      Top             =   360
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   2520
      Picture         =   "Form1.frx":2BA8
      Stretch         =   -1  'True
      Top             =   60
      Width           =   255
   End
   Begin VB.Line Line 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   360
      X2              =   2880
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      Index           =   3
      X1              =   2880
      X2              =   360
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   360
      X2              =   2880
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      Index           =   0
      X1              =   2880
      X2              =   360
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lblPrint 
      Caption         =   "  Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   1480
      Width           =   2535
   End
   Begin VB.Label lblSaveAs 
      Caption         =   "  Save As"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label lblSave 
      Caption         =   "  Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   760
      Width           =   2535
   End
   Begin VB.Image Image 
      Height          =   1800
      Left            =   0
      Picture         =   "Form1.frx":3472
      Stretch         =   -1  'True
      Top             =   0
      Width           =   300
   End
   Begin VB.Label lblOpen 
      Caption         =   "  Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label lblNew 
      BackColor       =   &H8000000A&
      Caption         =   "  New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   360
      TabIndex        =   0
      Top             =   40
      Width           =   2535
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'adapted and modified code from Joel CinCity e-m: linilsen@c2i.net
'Ecapsulated menu form by using as object with events.
'This makes the menu form reusable throughout a project
'by removing references to the main form that calls it.
'Bob Richards 1-20-02

Option Explicit

'declare an event for every menu selection
Public Event NewClick()
Public Event OpenClick()
Public Event SaveClick()
Public Event SaveAsClick()
Public Event PrintClick()


'***************************************************** New
Private Sub lblNew_Click()
    Me.Hide
    RaiseEvent NewClick
    Unload Me
End Sub

Private Sub lblNew_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LabelDefault
    lblNew.BackColor = &H80&
    lblNew.ForeColor = &HFFFF&
End Sub

'***************************************************** Open
Private Sub lblOpen_Click()
    Me.Hide
    RaiseEvent OpenClick
    Unload Me
End Sub

Private Sub lblOpen_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LabelDefault
    lblOpen.BackColor = &H80&
    lblOpen.ForeColor = &HFFFF&
End Sub

'***************************************************** Save
Private Sub lblSave_Click()
    Me.Hide
    RaiseEvent SaveClick
    Unload Me
End Sub

Private Sub lblSave_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LabelDefault
    lblSave.BackColor = &H80&
    lblSave.ForeColor = &HFFFF&
End Sub

'***************************************************** SaveAs
Private Sub lblSaveAs_Click()
    Me.Hide
    RaiseEvent SaveAsClick
    Unload Me
End Sub

Private Sub lblSaveAs_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LabelDefault
    lblSaveAs.BackColor = &H80&
    lblSaveAs.ForeColor = &HFFFF&
End Sub

'***************************************************** Print
Private Sub lblPrint_Click()
    Me.Hide
    RaiseEvent PrintClick
    Unload Me
End Sub

Private Sub lblPrint_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LabelDefault
    lblPrint.BackColor = &H80&
    lblPrint.ForeColor = &HFFFF&
    
End Sub

'*****************************************************
Private Sub LabelDefault()  'sets labels to default colors
    lblPrint.BackColor = &H8000000F
    lblPrint.ForeColor = vbBlack
    lblOpen.BackColor = &H8000000F
    lblOpen.ForeColor = vbBlack
    lblSave.BackColor = &H8000000F
    lblSave.ForeColor = vbBlack
    lblNew.BackColor = &H8000000F
    lblNew.ForeColor = vbBlack
    lblSaveAs.BackColor = &H8000000F
    lblSaveAs.ForeColor = vbBlack
End Sub

Public Sub EnableItems(Optional NewFile As Boolean = True, _
                       Optional OpenFile As Boolean = True, _
                       Optional SaveFile As Boolean = True, _
                       Optional SaveAs As Boolean = True, _
                       Optional PrintFile As Boolean = True)

    lblNew.Enabled = NewFile
    lblOpen.Enabled = OpenFile
    lblSave.Enabled = SaveFile
    lblSaveAs.Enabled = SaveAs
    lblPrint.Enabled = PrintFile
    
End Sub

