VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9855
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   9480
      Picture         =   "frmAbout.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   5250
      Left            =   0
      Picture         =   "frmAbout.frx":5B8D
      Top             =   0
      Width           =   9705
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MoveScreen As Boolean, color As Long, flag As Byte
Dim MousX, MousY, CurrX, CurrY As Integer

Private Sub Command3_Click()
    End
End Sub

Private Sub Form_Load()
    On Error GoTo errRtn
    color = RGB(0, 0, 255): flag = 0
    flag = flag Or LWA_COLORKEY: frmAbout.Show
    SetTranslucent frmAbout.hwnd, color, 0, flag
    Exit Sub
errRtn:
    MsgBox Err.Description & " Source : " & Err.Description
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Image1_Click()
    Unload Me
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveScreen = True: MousX = X: MousY = Y
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MoveScreen Then
        CurrX = Me.Left - MousX + X
        CurrY = Me.Top - MousY + Y
        Me.Move CurrX, CurrY
    End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveScreen = False
End Sub
