VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDaftarTindakan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Kunjungan Pasien "
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarTindakan.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5670
   Begin VB.Frame fraPeriode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   5685
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
         Height          =   855
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   5415
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   240
            TabIndex        =   0
            Top             =   360
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
            OLEDropMode     =   1
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   123600899
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3000
            TabIndex        =   1
            Top             =   360
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
            Format          =   123600899
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2520
            TabIndex        =   9
            Top             =   480
            Width           =   255
         End
      End
   End
   Begin VB.Frame fraButton 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   2160
      Width           =   5685
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   495
         Left            =   2280
         TabIndex        =   4
         Top             =   240
         Width           =   1545
      End
      Begin VB.OptionButton optpoli 
         Caption         =   "Pasien Poli"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.OptionButton optkonsul 
         Caption         =   "Pasien Konsul"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   3960
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   10
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
      Left            =   3840
      Picture         =   "frmDaftarTindakan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarTindakan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarTindakan.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "frmDaftarTindakan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCetak_Click()
    On Error GoTo hell
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value
    strSQL = "Select KdRuangan, KdPelayananRS,NoPendaftaran, TglPendaftaran, TglPelayanan, NoCM, NamaPasien, Umur, JenisKelamin, JenisPasien,NamaDiagnosa,StatusKasus, DokterPemeriksa, TindakanPelayanan, Tarif, TarifCito, JmlPelayanan, TotalBiaya, JenisPelayanan from V_DaftarTindakanPasienperDokter " & _
    "WHERE (TglPelayanan BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
    "and  KdRuangan='" & mstrKdRuangan & "' and KdPelayananRS <>'001001' and KdPelayananRS<>'202001' and KdPelayananRS <>'202002' and KdPelayananRS <>'301004' and KdPelayananRS <>'206001'"
    If ValidasiTanggal(dtpAwal, dtpAkhir) = False Then Exit Sub
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then MsgBox "Data Tidak Ada", vbExclamation, "Validasi": Exit Sub
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmCetakDaftarTindakan.Show
    Exit Sub
hell:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCetak.SetFocus
End Sub

Private Sub dtpAkhir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetak.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub dtpAwal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    With Me
        .dtpAwal.Value = Now
        .dtpAkhir.Value = Now
    End With
    Call PlayFlashMovie(Me)
End Sub

Private Sub optGroupBy_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAwal.SetFocus
End Sub

Private Sub optkonsul_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetak.SetFocus
End Sub

Private Sub optpoli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetak.SetFocus
End Sub

