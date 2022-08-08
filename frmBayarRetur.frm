VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmBayarRetur 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pembayaran Retur Struk Pasien"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13620
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBayarRetur.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   13620
   Begin VB.CheckBox chkBiayaAdministrasiAwal 
      Caption         =   "Biaya Administrasi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6960
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Frame fraDebug 
      Caption         =   "Debug"
      Height          =   1095
      Left            =   0
      TabIndex        =   22
      Top             =   -240
      Visible         =   0   'False
      Width           =   6975
      Begin VB.TextBox txtNoStruk 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4080
         TabIndex        =   46
         Text            =   "txtNoStruk"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtNoBKM 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2760
         TabIndex        =   45
         Text            =   "txtNoBKM"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtNoRetur 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1440
         TabIndex        =   44
         Text            =   "txtNoRetur"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtNoBKK 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Text            =   "txtNOBKK"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtNamaFormPengirim 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         TabIndex        =   32
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Nama Form Pengirim"
         Height          =   210
         Index           =   9
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1665
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   35
      Top             =   2040
      Width           =   13575
      Begin MSComCtl2.DTPicker dtpTglRetur 
         Height          =   330
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   49807363
         UpDown          =   -1  'True
         CurrentDate     =   38295
      End
      Begin MSDataListLib.DataCombo dcCaraBayar 
         Height          =   330
         Left            =   2280
         TabIndex        =   8
         Top             =   480
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. BKK/Retur"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cara Bayar"
         Height          =   210
         Index           =   1
         Left            =   2280
         TabIndex        =   39
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Retur ->"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7950
         TabIndex        =   38
         Top             =   360
         Width           =   2250
      End
      Begin VB.Label lblTotalTagihan 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Rp. 800.000.000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10320
         TabIndex        =   36
         Top             =   360
         Width           =   2970
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   0
      TabIndex        =   37
      Top             =   5040
      Width           =   13575
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   430
         Left            =   11640
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   430
         Left            =   9840
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   23
      Top             =   4080
      Width           =   13575
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7320
         TabIndex        =   15
         Top             =   480
         Width           =   6015
      End
      Begin VB.TextBox txtJmlUang 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2520
         MaxLength       =   15
         TabIndex        =   13
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtKembalian 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4920
         TabIndex        =   14
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtBiayaAdministrasi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         MaxLength       =   15
         TabIndex        =   12
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         Height          =   210
         Index           =   0
         Left            =   7320
         TabIndex        =   43
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Jumlah Bayar"
         Height          =   210
         Left            =   2520
         TabIndex        =   26
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Uang Kembalian"
         Height          =   210
         Left            =   4920
         TabIndex        =   25
         Top             =   240
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Biaya Administrasi"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   1410
      End
   End
   Begin VB.Frame fraAlatPembayaran 
      Caption         =   "Alat Pembayaran Retur"
      Enabled         =   0   'False
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
      TabIndex        =   27
      Top             =   3000
      Width           =   13575
      Begin VB.TextBox txtNamaPemilik 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   10080
         MaxLength       =   50
         TabIndex        =   11
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox txtNoKartu 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6360
         MaxLength       =   50
         TabIndex        =   10
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txtBankPenyediaKartu 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   9
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pemilik Kartu / Rekening"
         Height          =   210
         Index           =   3
         Left            =   10080
         TabIndex        =   30
         Top             =   360
         Width           =   2490
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Kartu / Rekening"
         Height          =   210
         Index           =   2
         Left            =   6360
         TabIndex        =   29
         Top             =   360
         Width           =   1725
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank / Penyedia Kartu"
         Height          =   210
         Index           =   1
         Left            =   2160
         TabIndex        =   28
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   0
      TabIndex        =   18
      Top             =   960
      Width           =   13575
      Begin VB.TextBox txtTotalHarusRetur 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   11280
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtPembebasanAwal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtBiayaAdministrasiAwal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox TxtTotalBiaya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtTanggunganPenjamin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtTanggunganRS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Total Harus Retur"
         Height          =   210
         Index           =   1
         Left            =   11280
         TabIndex        =   42
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Pembebasan"
         Height          =   210
         Index           =   0
         Left            =   9120
         TabIndex        =   41
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Biaya Pelayanan"
         Height          =   210
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1755
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tanggungan Penjamin"
         Height          =   210
         Left            =   2520
         TabIndex        =   20
         Top             =   240
         Width           =   1830
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tanggungan RS"
         Height          =   210
         Left            =   4800
         TabIndex        =   19
         Top             =   240
         Width           =   1305
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   31
      Top             =   5925
      Width           =   13620
      _ExtentX        =   24024
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   11959
            Text            =   "Cetak Detail Kuitansi (F1)"
            TextSave        =   "Cetak Detail Kuitansi (F1)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   11959
            Text            =   "Cetak Kuitansi (F9)"
            TextSave        =   "Cetak Kuitansi (F9)"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   47
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
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmBayarRetur.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   11760
      Picture         =   "frmBayarRetur.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmBayarRetur.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "frmBayarRetur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnCmdSimpan As Boolean

Private Sub subCetakDetailKuitansi()
On Error GoTo hell
    vLaporan = "Print"
    mstrNoBKK = txtNoBKK.Text
    mstrNoBKM = txtNoBKM.Text
    strSQL = "SELECT NoStruk FROM PembayaranTagihanPasien WHERE (NoBKM = '" & mstrNoBKM & "')"
    Call msubRecFO(rs, strSQL)
    mstrNoStruk = rs("NoStruk").Value
    frmCetakRetur.CetakUlangRetur
hell:
End Sub

Private Sub subCetakKuitansi()
On Error GoTo hell
    vLaporan = "Print"
    mstrNoBKK = txtNoBKK.Text
    mstrNoBKM = txtNoBKM.Text
    strSQL = "SELECT NoStruk FROM PembayaranTagihanPasien WHERE (NoBKM = '" & mstrNoBKM & "')"
    Call msubRecFO(rs, strSQL)
    mstrNoStruk = rs("NoStruk").Value
    frmCetakRetur.CetakUlangJenisKuitansiRetur
hell:
End Sub

Private Sub chkBiayaAdministrasiAwal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtBiayaAdministrasi.SetFocus
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo errLoad
Dim i As Integer
    
    If sp_AddStrukBuktiKasKeluar() = False Then Exit Sub
    If sp_Retur = False Then Exit Sub
    If sp_PembayaranReturStrukPelayananPasien() = False Then Exit Sub
        
    MsgBox "Pembayaran Tagihan Pasien Sukses", vbInformation, "Validasi"
    cmdSimpan.Enabled = False
    mcurPembebasan = 0
    blnCmdSimpan = True
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    If txtNamaFormPengirim.Text = "frmReturStrukPelayananPasien" Then frmReturStrukPelayananPasien.Enabled = True
    Unload Me
    blnCmdSimpan = False
End Sub

Private Sub dcCaraBayar_Change()
On Error GoTo errLoad
    strSQL = "SELECT NamaBank, NoKartu, AtasNama FROM dbo.StrukBuktiKasMasuk WHERE (NoBKM = '" & txtNoBKM.Text & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        txtBankPenyediaKartu.Text = rs("NamaBank")
        txtNoKartu.Text = rs("NoKartu")
        txtNamaPemilik.Text = rs("AtasNama")
    Else
        txtBankPenyediaKartu.Text = ""
        txtNoKartu.Text = ""
        txtNamaPemilik.Text = ""
    End If
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dtpTglRetur_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcCaraBayar.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errLoad
    
    Select Case KeyCode
        Case vbKeyF1
            If cmdSimpan.Enabled = True Then Exit Sub
            Call subCetakDetailKuitansi
        Case vbKeyF9
            If cmdSimpan.Enabled = True Then Exit Sub
            Call subCetakKuitansi
    End Select

Exit Sub
errLoad:
'    Call msubPesanError
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error GoTo errLoad

    Call centerForm(Me, MDIUtama)
    
    dtpTglRetur.Value = Now
    Call subLoadDCSource
    Call PlayFlashMovie(Me)
Exit Sub
errLoad:
    Call msubPesanError
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    If txtNamaFormPengirim.Text = "frmTagihanPasien" Then frmTagihanPasien.Enabled = True
'End Sub

Private Sub txtBiayaAdministrasi_Change()
On Error GoTo errLoad
        
    If txtBiayaAdministrasi.Text = "" Or txtBiayaAdministrasi.Text = Format(0, "#,###.00") Then
        If chkBiayaAdministrasiAwal.Value = vbChecked Then
            lblTotalTagihan.Caption = CCur(txtTotalBiaya.Text) - CCur(txtTanggunganPenjamin.Text) - CCur(txtTanggunganRS.Text) - CCur(txtPembebasanAwal.Text) + CCur(txtBiayaAdministrasiAwal.Text)
        Else
            lblTotalTagihan.Caption = CCur(txtTotalBiaya.Text) - CCur(txtTanggunganPenjamin.Text) - CCur(txtTanggunganRS.Text) - CCur(txtPembebasanAwal.Text)
        End If
    Else
        If chkBiayaAdministrasiAwal.Value = vbChecked Then
            lblTotalTagihan.Caption = CCur(txtTotalBiaya.Text) - CCur(txtTanggunganPenjamin.Text) - CCur(txtTanggunganRS.Text) - CCur(txtPembebasanAwal.Text) + CCur(txtBiayaAdministrasiAwal.Text) - CCur(txtBiayaAdministrasi.Text)
        Else
            lblTotalTagihan.Caption = CCur(txtTotalBiaya.Text) - CCur(txtTanggunganPenjamin.Text) - CCur(txtTanggunganRS.Text) - CCur(txtPembebasanAwal.Text) - CCur(txtBiayaAdministrasi.Text)
        End If
    End If
    lblTotalTagihan.Caption = Format(lblTotalTagihan.Caption, "#,###.00")
    txtJmlUang.Text = lblTotalTagihan.Caption
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtBiayaAdministrasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJmlUang.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtBiayaAdministrasi_LostFocus()
On Error GoTo errLoad
    If Len(Trim(txtBiayaAdministrasi.Text)) < 1 Then txtBiayaAdministrasi.Text = 0
    txtBiayaAdministrasi.Text = Format(txtBiayaAdministrasi, "#,###.00")

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtJmlUang_Change()
    If txtJmlUang.Text = "" Then txtJmlUang.Text = 0
    txtJmlUang = Format(txtJmlUang, "###,##0")
    txtJmlUang.SelStart = Len(txtJmlUang.Text)
    
    If CCur(lblTotalTagihan.Caption) - CCur(txtJmlUang.Text) >= 0 Then
        txtKembalian.Text = Format(0, "#,###.00")
    Else
        txtKembalian.Text = Format(CCur(txtJmlUang.Text) - CCur(lblTotalTagihan.Caption), "#,###.00")
    End If
End Sub

Private Sub txtJmlUang_GotFocus()
    txtJmlUang.SelStart = 0
    txtJmlUang.SelLength = Len(txtJmlUang.Text)
End Sub

Private Sub txtJmlUang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
    If Not (KeyAscii >= 0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

'Store procedure untuk mengisi struk kas keluar
Private Function sp_AddStrukBuktiKasKeluar() As Boolean
On Error GoTo errLoad
    
    sp_AddStrukBuktiKasKeluar = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("TglBKK", adDate, adParamInput, , Format(dtpTglRetur.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdCaraBayar", adChar, adParamInput, 2, dcCaraBayar.BoundText)
        .Parameters.Append .CreateParameter("NamaBank", adVarChar, adParamInput, 100, IIf(Len(Trim(txtBankPenyediaKartu.Text)) > 0, Trim(txtBankPenyediaKartu.Text), Null))
        .Parameters.Append .CreateParameter("NoAccount", adVarChar, adParamInput, 50, IIf(Len(Trim(txtNoKartu.Text)) > 0, Trim(txtNoKartu.Text), Null))
        .Parameters.Append .CreateParameter("AtasNama", adVarChar, adParamInput, 50, IIf(Len(Trim(txtNamaPemilik.Text)) > 0, Trim(txtNamaPemilik.Text), Null))
        
        If CCur(txtJmlUang.Text) > CCur(lblTotalTagihan.Caption) Then
            .Parameters.Append .CreateParameter("JmlBayar", adCurrency, adParamInput, , CCur(lblTotalTagihan.Caption))
        Else
            .Parameters.Append .CreateParameter("JmlBayar", adCurrency, adParamInput, , CCur(txtJmlUang.Text))
        End If
        
        .Parameters.Append .CreateParameter("Administrasi", adCurrency, adParamInput, , CCur(txtBiayaAdministrasi.Text))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 100, IIf(Len(Trim(txtKeterangan.Text)) > 0, Trim(txtKeterangan.Text), Null))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, noidpegawai)
        .Parameters.Append .CreateParameter("OutputNoBKK", adChar, adParamOutput, 10, Null)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_StrukBuktiKasKeluar"
        .CommandType = adCmdStoredProc
        .Execute
    
        If .Parameters("RETURN_VALUE").Value <> 0 Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Struk Kas Keluar", vbCritical, "Validasi"
            sp_AddStrukBuktiKasKeluar = False
        Else
            If Not IsNull(.Parameters("OutputNoBKK").Value) Then txtNoBKK = .Parameters("OutputNoBKK").Value
        End If
        Set dbcmd = Nothing
    End With
    
Exit Function
errLoad:
    sp_AddStrukBuktiKasKeluar = False
    Call msubPesanError
End Function

Private Function sp_Retur() As Boolean
On Error GoTo errLoad

    sp_Retur = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRetur", adChar, adParamInput, 10, txtNoRetur.Text)
        .Parameters.Append .CreateParameter("TglRetur", adDate, adParamInput, , Format(dtpTglRetur.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 50, IIf(Len(Trim(txtKeterangan.Text)) = 0, Null, Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("OutputNoRetur", adChar, adParamOutput, 10, Null)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_Retur"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam retur struk", vbCritical, "Informasi"
            sp_Retur = False
        Else
            txtNoRetur.Text = .Parameters("OutputNoRetur").Value
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

Exit Function
errLoad:
    sp_Retur = False
    Call msubPesanError
End Function

Private Function sp_PembayaranReturStrukPelayananPasien() As Boolean
On Error GoTo errLoad

    sp_PembayaranReturStrukPelayananPasien = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRetur", adChar, adParamInput, 10, txtNoRetur.Text)
        .Parameters.Append .CreateParameter("NoBKM", adChar, adParamInput, 10, txtNoBKM.Text)
        .Parameters.Append .CreateParameter("NoStruk", adChar, adParamInput, 10, txtNoStruk.Text)
        .Parameters.Append .CreateParameter("TotalBiaya", adCurrency, adParamInput, , CCur(txtTotalBiaya.Text))
        .Parameters.Append .CreateParameter("TotalPpn", adCurrency, adParamInput, , 0)
        .Parameters.Append .CreateParameter("TotalDiscount", adCurrency, adParamInput, , 0)
        .Parameters.Append .CreateParameter("JmlHutangPenjamin", adCurrency, adParamInput, , CCur(txtTanggunganPenjamin.Text))
        .Parameters.Append .CreateParameter("JmlTanggunganRS", adCurrency, adParamInput, , CCur(txtTanggunganRS.Text))
'        If CCur(txtJmlUang.Text) > CCur(lblTotalTagihan.Caption) Then
'            .Parameters.Append .CreateParameter("JmlHarusDiretur", adCurrency, adParamInput, , CCur(lblTotalTagihan.Caption))
'        Else
            .Parameters.Append .CreateParameter("JmlHarusDiretur", adCurrency, adParamInput, , CCur(txtTotalHarusRetur.Text))
'        End If
        .Parameters.Append .CreateParameter("NoBKK", adChar, adParamInput, 10, txtNoBKK.Text)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PembayaranReturStrukPelayananPasien"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam retur struk", vbCritical, "Informasi"
            sp_PembayaranReturStrukPelayananPasien = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

Exit Function
errLoad:
    sp_PembayaranReturStrukPelayananPasien = False
    Call msubPesanError
End Function

Private Sub subLoadDCSource()
On Error GoTo errLoad
    
    Call msubDcSource(dcCaraBayar, rs, "SELECT KdCaraBayar, CaraBayar FROM CaraBayar")
    If rs.EOF = False Then dcCaraBayar.BoundText = rs(0).Value

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtKeterangan_LostFocus()
    txtKeterangan.Text = StrConv(txtKeterangan.Text, vbProperCase)
End Sub
