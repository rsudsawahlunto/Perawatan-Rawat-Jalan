VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDaftarCetakLapDataBumil 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Cetak Laporan Data Ibu Hamil"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   Icon            =   "frmDaftarCetakLapDataBumil.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdtutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdcetak 
      Caption         =   "&Cetak"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periode"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   4935
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   330
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   61014019
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   330
         Left            =   2760
         TabIndex        =   1
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   61014019
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "s/d"
         Height          =   210
         Left            =   2400
         TabIndex        =   5
         Top             =   525
         Width           =   255
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1800
      _cx             =   3175
      _cy             =   1720
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   3240
      Picture         =   "frmDaftarCetakLapDataBumil.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1755
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarCetakLapDataBumil.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3135
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarCetakLapDataBumil.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmDaftarCetakLapDataBumil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCetak_Click()
    On Error GoTo hell
    mdTglAwal = dtpAwal.Value 'TglAwal
    mdTglAkhir = dtpAkhir.Value 'TglAkhir
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmCetakLapDataBumil.Show
    Exit Sub
hell:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCetak.SetFocus
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    dtpAwal.Value = Now
    dtpAkhir.Value = Now
    Call PlayFlashMovie(Me)
End Sub
