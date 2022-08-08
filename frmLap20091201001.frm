VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLap20091201001 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Kunjungan Pasien Berdasarkan Diagnosa"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8070
   Icon            =   "frmLap20091201001.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   8070
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   2040
      Width           =   8055
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   8055
      Begin VB.Frame Frame4 
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
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   5055
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   2640
            TabIndex        =   1
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
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
            Format          =   61079555
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
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
            CustomFormat    =   "ddMMMMyyyy"
            Format          =   61079555
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   2320
            TabIndex        =   10
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   1680
            TabIndex        =   7
            Top             =   360
            Visible         =   0   'False
            Width           =   375
         End
      End
      Begin MSDataListLib.DataCombo dcRuangan 
         Height          =   360
         Left            =   5280
         TabIndex        =   2
         Top             =   360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Ruang Rawat"
         Height          =   255
         Left            =   5280
         TabIndex        =   9
         Top             =   120
         Width           =   1215
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   11
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
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmLap20091201001.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4455
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   6240
      Picture         =   "frmLap20091201001.frx":2328
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
End
Attribute VB_Name = "frmLap20091201001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCetak_Click()
    If Periksa("datacombo", dcRuangan, "Nama Ruangan kosong") = False Then Exit Sub
    strNNamaRuangan = dcRuangan.Text
    If ValidasiTanggal(dtpAwal, dtpAkhir) = False Then Exit Sub
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then MsgBox "Tidak Ada Data", vbExclamation, "Validasi": Exit Sub
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmctkcrLap20091201001.Show
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcRuangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcRuangan.MatchedWithList = True Then cmdCetak.SetFocus
        strSQL = "select kdRuangan, NamaRuangan from Ruangan where StatusEnabled='1' and (NamaRuangan LIKE '%" & dcRuangan.Text & "%')order by NamaRuangan"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcRuangan.Text = ""
            cmdCetak.SetFocus
            Exit Sub
        End If
        dcRuangan.BoundText = rs(0).Value
        dcRuangan.Text = rs(1).Value
    End If
End Sub

Private Sub dtpAkhir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcRuangan.SetFocus
End Sub

Private Sub dtpAwal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subLoadDcSource
    dtpAwal.Value = Format(Now, "dd MMMM yyyy 00:00:00")
    dtpAkhir.Value = Format(Now, "dd MMMM yyyy 23:59:59")
End Sub

Private Sub subLoadDcSource()
    Call msubDcSource(dcRuangan, rs, "select kdRuangan, NamaRuangan from Ruangan where StatusEnabled='1' order by NamaRuangan")
    dcRuangan.BoundText = mstrKdRuangan
End Sub
