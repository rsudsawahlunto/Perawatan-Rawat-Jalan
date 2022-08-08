VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmPasienBaru 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Data Pasien"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPasienBaru.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   8790
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   0
      TabIndex        =   45
      Top             =   7320
      Width           =   8775
      Begin VB.CommandButton cmdDetailPasien 
         Caption         =   "&Detail Pasien"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   20
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   6840
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   5040
         TabIndex        =   21
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame fraAlamatPas 
      Caption         =   "Alamat Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   35
      Top             =   4680
      Width           =   8775
      Begin VB.TextBox txtTelepon 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5760
         MaxLength       =   15
         TabIndex        =   14
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtKodePos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7680
         MaxLength       =   5
         TabIndex        =   19
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtAlamat 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         MaxLength       =   100
         TabIndex        =   12
         Top             =   600
         Width           =   4575
      End
      Begin MSDataListLib.DataCombo dcKota 
         Height          =   330
         Left            =   4200
         TabIndex        =   16
         Top             =   1320
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcKecamatan 
         Height          =   330
         Left            =   240
         TabIndex        =   17
         Top             =   2040
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcKelurahan 
         Height          =   330
         Left            =   4200
         TabIndex        =   18
         Top             =   2040
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcPropinsi 
         Height          =   330
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox meRTRW 
         Height          =   330
         Left            =   4920
         TabIndex        =   13
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         Mask            =   "##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Telepon"
         Height          =   210
         Left            =   5760
         TabIndex        =   43
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Kode Pos"
         Height          =   210
         Left            =   7680
         TabIndex        =   42
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "RT/RW"
         Height          =   210
         Left            =   4920
         TabIndex        =   41
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Kelurahan (Desa)"
         Height          =   210
         Left            =   4200
         TabIndex        =   40
         Top             =   1800
         Width           =   1395
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Kecamatan"
         Height          =   210
         Left            =   240
         TabIndex        =   39
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Propinsi"
         Height          =   210
         Left            =   240
         TabIndex        =   38
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Kota (Kabupaten)"
         Height          =   210
         Left            =   4200
         TabIndex        =   37
         Top             =   1080
         Width           =   1470
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Alamat Lengkap"
         Height          =   210
         Left            =   240
         TabIndex        =   36
         Top             =   360
         Width           =   1305
      End
   End
   Begin VB.Frame fraPasien 
      Caption         =   "Data Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   26
      Top             =   2040
      Width           =   8775
      Begin VB.TextBox txtNamaPanggilan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         MaxLength       =   25
         TabIndex        =   47
         Top             =   2040
         Width           =   2895
      End
      Begin VB.ComboBox cboNamaDepan 
         Appearance      =   0  'Flat
         Height          =   330
         ItemData        =   "frmPasienBaru.frx":0CCA
         Left            =   240
         List            =   "frmPasienBaru.frx":0CE0
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNoIdentitas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5760
         MaxLength       =   20
         TabIndex        =   5
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtHari 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7920
         MaxLength       =   2
         TabIndex        =   11
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtBulan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7200
         MaxLength       =   2
         TabIndex        =   10
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   4
         Top             =   600
         Width           =   3975
      End
      Begin VB.ComboBox cboJnsKelaminPasien 
         Appearance      =   0  'Flat
         Height          =   330
         ItemData        =   "frmPasienBaru.frx":0D04
         Left            =   240
         List            =   "frmPasienBaru.frx":0D0E
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtTempatLahir 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1680
         MaxLength       =   25
         TabIndex        =   7
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox txtTahun 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6480
         MaxLength       =   3
         TabIndex        =   9
         Top             =   1320
         Width           =   615
      End
      Begin MSMask.MaskEdBox meTglLahir 
         Height          =   330
         Left            =   5040
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         HideSelection   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mm-yy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nama Panggilan"
         Height          =   210
         Left            =   240
         TabIndex        =   48
         Top             =   1800
         Width           =   1275
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Nama Depan"
         Height          =   210
         Left            =   240
         TabIndex        =   44
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Bulan"
         Height          =   210
         Left            =   7200
         TabIndex        =   34
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Hari"
         Height          =   210
         Left            =   7920
         TabIndex        =   33
         Top             =   1080
         Width           =   300
      End
      Begin VB.Label lblNamaPasien 
         AutoSize        =   -1  'True
         Caption         =   "Nama Lengkap"
         Height          =   210
         Left            =   1680
         TabIndex        =   32
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label lblJnsKlm 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   240
         TabIndex        =   31
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label lblTmpLhr 
         AutoSize        =   -1  'True
         Caption         =   "Tempat Lahir"
         Height          =   210
         Left            =   1680
         TabIndex        =   30
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label lblTglLhr 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Lahir"
         Height          =   210
         Left            =   5040
         TabIndex        =   29
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label lblumur 
         AutoSize        =   -1  'True
         Caption         =   "Tahun"
         Height          =   210
         Left            =   6480
         TabIndex        =   28
         Top             =   1080
         Width           =   525
      End
      Begin VB.Label lblGolDrh 
         AutoSize        =   -1  'True
         Caption         =   "No. Identitas"
         Height          =   210
         Left            =   5760
         TabIndex        =   27
         Top             =   360
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   23
      Top             =   1080
      Width           =   8775
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "No. CM Otomatis"
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
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6600
         TabIndex        =   2
         Top             =   450
         Width           =   1935
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3840
         MaxLength       =   12
         TabIndex        =   1
         Top             =   480
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtpTglPendaftaran 
         Height          =   330
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dddd,dd MMMM yyyy HH:mm"
         Format          =   140378115
         CurrentDate     =   38061
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   3840
         TabIndex        =   25
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Pendaftaran"
         Height          =   210
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   1365
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   46
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
      Left            =   6960
      Picture         =   "frmPasienBaru.frx":0D28
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPasienBaru.frx":1AB0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPasienBaru.frx":4471
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmPasienBaru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim j As Integer

Private Sub cboJnsKelaminPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTempatLahir.SetFocus
End Sub

Private Sub cboNamaDepan_Click()
    If cboNamaDepan.Text = "Bayi" Then
        cboJnsKelaminPasien.Enabled = True
'        txtNamaPasien.SetFocus
    ElseIf cboNamaDepan.Text = "Tn." Then
        cboJnsKelaminPasien.Text = "Laki-Laki"
        cboJnsKelaminPasien.Enabled = False
    ElseIf cboNamaDepan.Text = "Ny." Or cboNamaDepan.Text = "Nn." Then
        cboJnsKelaminPasien.Text = "Perempuan"
        cboJnsKelaminPasien.Enabled = False
    Else
        cboJnsKelaminPasien.Enabled = True
    End If
End Sub

Private Sub cboNamaDepan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaPasien.SetFocus
End Sub

Private Sub cmdDetailPasien_Click()
    Load frmDetailPasien
    With frmDetailPasien
        .Show
        .txtnocm.Text = mstrNoCM
        .txtNamaPasien.Text = txtNamaPasien.Text
        .txtJK.Text = cboJnsKelaminPasien.Text
        .txtThn.Text = txtTahun.Text
        .txtBln.Text = txtBulan.Text
        .txtHr.Text = txtHari.Text
    End With
End Sub

Private Sub cmdSimpan_Click()
    If funcCekValidasi = False Then Exit Sub
    If txtBulan.Text = "" Then txtBulan.Text = 0
    If txtHari.Text = "" Then txtHari.Text = 0
    If sp_IdentitasPasien() = False Then Exit Sub
    Call subEnableButtonReg(True)
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcKecamatan_Click(Area As Integer)
   dcKelurahan.Text = ""
    txtKodePos = ""
    CekPilihanWilayah "dcKecamatan", "Click"

End Sub

Private Sub dcKecamatan_GotFocus()
On Error GoTo hell
    strSQL = "SELECT DISTINCT ISNULL(NamaKecamatan, '-') AS Alias,ISNULL(NamaKecamatan, '-') AS NamaKecamatan" & _
    " FROM V_Wilayah WHERE (NamaKotaKabupaten LIKE '%" & dcKota.Text & "%') AND (NamaPropinsi LIKE '%" & dcPropinsi.Text & "%')"
    Call msubDcSource(dcKecamatan, rs, strSQL)
    If rs.EOF = False Then dcKecamatan.Text = rs(1).Value
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcKecamatan_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 39 Then KeyAscii = 0
     If KeyAscii = 13 Then
        If dcKecamatan.MatchedWithList = True Then dcKelurahan.SetFocus
        strSQL = "SELECT DISTINCT ISNULL(NamaKecamatan, '-') AS Alias,ISNULL(NamaKecamatan, '-') AS NamaKecamatan" & _
        " FROM V_Wilayah WHERE (NamaPropinsi LIKE '%" & dcPropinsi.Text & "%') AND (NamaKotaKabupaten LIKE '%" & dcKota.Text & "%') AND (NamaKecamatan LIKE '%" & dcKecamatan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcKecamatan.BoundText = rs(0).Value
        dcKecamatan.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
'        j = 3
'        dcKelurahan.Enabled = True
'        'Call subLoadDataWilayah("kecamatan")
'        If dcKelurahan.Enabled = True Then
'            dcKelurahan.SetFocus
'        End If
'    End If
'    If KeyAscii = 13 Then
'        Call subLoadDataWilayah("kecamatan")
'        dcKelurahan.SetFocus
'    End If
End Sub

Private Sub dcKecamatan_LostFocus()
    If dcKecamatan.Text = "" Then Exit Sub
    If dcKecamatan.MatchedWithList = False Then dcKecamatan.Text = "": dcKecamatan.SetFocus: Exit Sub
    dcKecamatan = Trim(StrConv(dcKecamatan, vbProperCase))
End Sub

Private Sub dcKelurahan_Click(Area As Integer)
    txtKodePos = ""
    CekPilihanWilayah "dcKelurahan", "Click"

End Sub

Private Sub dcKelurahan_GotFocus()
 On Error GoTo hell
    strSQL = "SELECT DISTINCT ISNULL(NamaKelurahan, '-') AS Alias,ISNULL(NamaKelurahan, '-') AS NamaKelurahan" & _
    " FROM V_Wilayah WHERE (NamaKecamatan LIKE '%" & dcKecamatan.Text & "%') AND (NamaKotaKabupaten LIKE '%" & dcKota.Text & "%') AND (NamaPropinsi LIKE '%" & dcPropinsi.Text & "%')"
    Call msubDcSource(dcKelurahan, rs, strSQL)
    If rs.EOF = False Then dcKelurahan.Text = rs(1).Value
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcKelurahan_KeyPress(KeyAscii As Integer)
On Error GoTo hell
Call SetKeyPressToChar(KeyAscii)
If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcKelurahan.MatchedWithList = True Then txtKodePos.SetFocus
        strSQL = "SELECT DISTINCT ISNULL(NamaKelurahan, '-') AS Alias,ISNULL(NamaKelurahan, '-') AS NamaKelurahan" & _
        " FROM V_Wilayah WHERE (NamaPropinsi LIKE '%" & dcPropinsi.Text & "%') AND (NamaKotaKabupaten LIKE '%" & dcKota.Text & "%') AND (NamaKecamatan LIKE '%" & dcKecamatan.Text & "%') AND (NamaKelurahan LIKE '%" & dcKelurahan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcKelurahan.BoundText = rs(0).Value
        dcKelurahan.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcKelurahan_LostFocus()
    If dcKelurahan.Text = "" Then Exit Sub
    If dcKelurahan.MatchedWithList = False Then dcKelurahan.Text = "": dcKelurahan.SetFocus: Exit Sub
    dcKelurahan = Trim(StrConv(dcKelurahan, vbProperCase))
End Sub

Private Sub dcKota_Click(Area As Integer)
    dcKecamatan.Text = ""
    dcKelurahan.Text = ""
    txtKodePos = ""
    CekPilihanWilayah "dcKota", "Click"
End Sub

Private Sub dcKota_GotFocus()
 On Error GoTo hell
    strSQL = "SELECT DISTINCT ISNULL(NamaKotaKabupaten, '-') AS Alias,ISNULL(NamaKotaKabupaten, '-') AS NamaKotaKabupaten" & _
    " FROM V_Wilayah WHERE (NamaPropinsi LIKE '%" & dcPropinsi.Text & "%')"
    Call msubDcSource(dcKota, rs, strSQL)
    If rs.EOF = False Then dcKota.Text = rs(1).Value
    Exit Sub
hell:
    Call msubPesanError

End Sub

Private Sub dcKota_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcKota.MatchedWithList = True Then dcKecamatan.SetFocus
        strSQL = "SELECT DISTINCT ISNULL(NamaKotaKabupaten, '-') AS Alias,ISNULL(NamaKotaKabupaten, '-') AS NamaKotaKabupaten" & _
        " FROM V_Wilayah WHERE (NamaPropinsi LIKE '%" & dcPropinsi.Text & "%') AND (NamaKotaKabupaten LIKE '%" & dcKota.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcKota.BoundText = rs(0).Value
        dcKota.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
'       If KeyAscii = 13 Then
'        j = 2
'        dcKecamatan.Enabled = True
'        dcKelurahan.Enabled = True
'        'Call subLoadDataWilayah("kota")
'        If dcKecamatan.Enabled = True Then
'            dcKecamatan.SetFocus
'        End If
'    End If
'
''    If KeyAscii = 13 Then
''        Call subLoadDataWilayah("kota")
''        dcKecamatan.SetFocus
''    End If
End Sub

Private Sub dcKota_LostFocus()
    If dcKota.Text = "" Then Exit Sub
    If dcKota.MatchedWithList = False Then dcKota.Text = "": dcKota.SetFocus: Exit Sub
    dcKota = Trim(StrConv(dcKota, vbProperCase))
End Sub

Private Sub dcPropinsi_Click(Area As Integer)
    dcKota.Text = ""
    dcKecamatan.Text = ""
    dcKelurahan.Text = ""
    txtKodePos = ""
    CekPilihanWilayah "dcPropinsi", "Click"

End Sub

Private Sub dcPropinsi_KeyPress(KeyAscii As Integer)
On Error GoTo hell
Call SetKeyPressToChar(KeyAscii)
If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcPropinsi.MatchedWithList = True Then dcKota.SetFocus
        strSQL = "SELECT DISTINCT ISNULL(NamaPropinsi, '-') AS Alias,ISNULL(NamaPropinsi, '-') AS NamaPropinsi" & _
        " FROM V_Wilayah WHERE NamaPropinsi LIKE '%" & dcPropinsi.BoundText & "%'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcPropinsi.BoundText = rs(0).Value
        dcPropinsi.Text = rs(1).Value
    End If
    Exit Sub
hell:
    msubPesanError
    
'    Call SetKeyPressToChar(KeyAscii)
'    If KeyAscii = 39 Then KeyAscii = 0
''    If KeyAscii = 13 Then
''        Call subLoadDataWilayah("propinsi")
''        dcKota.SetFocus
''    End If
'     If KeyAscii = 13 Then
'        j = 1
'        dcKota.Enabled = True
'        dcKecamatan.Enabled = True
'        dcKelurahan.Enabled = True
'        'Call subLoadDataWilayah("propinsi")
'        If dcKota.Enabled = True Then
'            dcKota.SetFocus
'        End If
'    End If

End Sub

Private Sub dcPropinsi_LostFocus()
    If dcPropinsi.Text = "" Then dcPropinsi.BoundText = "": Exit Sub
    If dcPropinsi.MatchedWithList = False Then dcPropinsi.Text = "": dcPropinsi.SetFocus: Exit Sub
    dcPropinsi = Trim(StrConv(dcPropinsi, vbProperCase))
End Sub

Private Sub dtpTglPendaftaran_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cboNamaDepan.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    dtpTglPendaftaran.MaxDate = Now
    dtpTglPendaftaran.Value = Now

'    Call subDcSource
   subDcSource "Propinsi"
   subDcSource "Kota"
   subDcSource "Kecamatan"
   subDcSource "Kelurahan"

    If strPasien = "Lama" Then
        Call subEnableButtonReg(True)
        Call subVisibleButtonReg(True)
        cmdSimpan.Enabled = True
        Call subLoadDataPasien(mstrNoCM)
    ElseIf strPasien = "View" Then
        Call subEnableButtonReg(True)
        Call subVisibleButtonReg(False)
        cmdSimpan.Enabled = True
        Call subLoadDataPasien(mstrNoCM)
    End If
    Call PlayFlashMovie(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnForm = True Then Call frmDaftarPasienRJ.cmdCari_Click
End Sub

Private Sub meRTRW_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtTelepon.SetFocus
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub meRTRW_LostFocus()
    If meRTRW.Text <> "__/__" Then
        If InStr(1, meRTRW.Text, "_") <> 0 Then
            MsgBox "Format RT/RW adalah 00/00, harap isi RT/RW dengan benar", vbCritical, "Validasi"
            meRTRW.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub meTglLahir_Change()
'    If meTglLahir.Text = "__/__/____" Then
'        txtTahun.SetFocus
'        Exit Sub
'    End If
'    If funcCekValidasiTgl("TglLahir", meTglLahir) = "NoErr" Then
'        Call subYearOldCount(Format(meTglLahir.Text, "yyyy/mm/dd"))
'        txtTahun.Text = YOC_intYear
'        txtBulan.Text = YOC_intMonth
'        txtHari.Text = YOC_intDay
'        If strPasien = "Lama" Or strPasien = "View" Then Exit Sub
'        txtalamat.SetFocus
'    Else
'        txtTahun.Text = ""
'        txtBulan.Text = ""
'        txtHari.Text = ""
'    End If
End Sub

Private Sub meTglLahir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If meTglLahir.Text = "__/__/____" Then
            txtTahun.SetFocus
            Exit Sub
        End If
        If funcCekValidasiTgl("TglLahir", meTglLahir) = "NoErr" Then
            Call subYearOldCount(Format(meTglLahir.Text, "yyyy/mm/dd"))
            txtTahun.Text = YOC_intYear
            txtBulan.Text = YOC_intMonth
            txtHari.Text = YOC_intDay
            If strPasien = "Lama" Or strPasien = "View" Then Exit Sub
            txtalamat.SetFocus
        Else
            txtTahun.Text = ""
            txtBulan.Text = ""
            txtHari.Text = ""
        End If
    End If
End Sub

Private Sub meTglLahir_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub meTglLahir_LostFocus()
    If meTglLahir.Text = "__/__/____" Then Exit Sub
    If funcCekValidasiTgl("TglLahir", meTglLahir) = "NoErr" Then
        Call subYearOldCount(Format(meTglLahir.Text, "yyyy/mm/dd"))
        txtTahun.Text = YOC_intYear
        txtBulan.Text = YOC_intMonth
        txtHari.Text = YOC_intDay
    Else
        txtTahun.Text = ""
        txtBulan.Text = ""
        txtHari.Text = ""
    End If
End Sub

Private Sub txtAlamat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then meRTRW.SetFocus
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtBulan_Change()
    Dim dTglLahir As Date
    If txtBulan.Text = "" And txtTahun.Text = "" Then txtHari.SetFocus: Exit Sub
    If txtBulan.Text = "" Then txtBulan.Text = 0
    If txtTahun.Text = "" And txtHari.Text = "" Then
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
    ElseIf txtTahun.Text <> "" And txtHari.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    ElseIf txtTahun.Text = "" And txtHari.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
    ElseIf txtTahun.Text <> "" And txtHari.Text = "" Then
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    End If
End Sub

Private Sub txtBulan_KeyPress(KeyAscii As Integer)
    Dim dTglLahir As Date
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
        If txtBulan.Text = "" And txtTahun.Text = "" Then txtHari.SetFocus: Exit Sub
        If txtBulan.Text = "" Then txtBulan.Text = 0
        If txtTahun.Text = "" And txtHari.Text = "" Then
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
        ElseIf txtTahun.Text <> "" And txtHari.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        ElseIf txtTahun.Text = "" And txtHari.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
        ElseIf txtTahun.Text <> "" And txtHari.Text = "" Then
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        End If
        txtHari.SetFocus
    End If
End Sub

Private Sub txtHari_Change()
    Dim dTglLahir As Date
    If txtHari.Text = "" And txtBulan.Text = "" And txtTahun.Text = "" Then txtalamat.SetFocus: Exit Sub
    If txtHari.Text = "" Then txtHari.Text = 0
    If txtTahun.Text = "" And txtBulan.Text = "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
    ElseIf txtTahun.Text <> "" And txtBulan.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    ElseIf txtTahun.Text = "" And txtBulan.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
    ElseIf txtTahun.Text <> "" And txtBulan.Text = "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    End If
End Sub

Private Sub txtHari_KeyPress(KeyAscii As Integer)
    Dim dTglLahir As Date
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
        If txtHari.Text = "" And txtBulan.Text = "" And txtTahun.Text = "" Then txtalamat.SetFocus: Exit Sub
        If txtHari.Text = "" Then txtHari.Text = 0
        If txtTahun.Text = "" And txtBulan.Text = "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        ElseIf txtTahun.Text <> "" And txtBulan.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        ElseIf txtTahun.Text = "" And txtBulan.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
        ElseIf txtTahun.Text <> "" And txtBulan.Text = "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        End If
        txtalamat.SetFocus
    End If
End Sub

Private Sub txtKodePos_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
'        Call subLoadDataWilayah("kodepos")
        If cmdSimpan.Enabled = True Then cmdSimpan.SetFocus Else cmdTutup.SetFocus
    End If
End Sub


Private Sub txtNamaPanggilan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtalamat.SetFocus
End Sub

Private Sub txtNamaPasien_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtNoIdentitas.SetFocus
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtNoCM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtNoIdentitas_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
        If cboJnsKelaminPasien.Enabled = False Then
            txtTempatLahir.SetFocus
        Else
            cboJnsKelaminPasien.SetFocus
        End If
    End If
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtTahun_Change()
    Dim dTglLahir As Date
    If txtTahun = "" Then txtBulan.SetFocus: Exit Sub
    If txtBulan.Text = "" And txtHari.Text = "" Then
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), Date)
    ElseIf txtBulan.Text <> "" And txtHari.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    ElseIf txtBulan.Text = "" And txtHari.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    ElseIf txtBulan.Text <> "" And txtHari.Text = "" Then
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    End If
End Sub

Private Sub txtTahun_KeyPress(KeyAscii As Integer)
    Dim dTglLahir As Date
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
        If txtTahun = "" Then txtBulan.SetFocus: Exit Sub
        If txtBulan.Text = "" And txtHari.Text = "" Then
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), Date)
        ElseIf txtBulan.Text <> "" And txtHari.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        ElseIf txtBulan.Text = "" And txtHari.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        ElseIf txtBulan.Text <> "" And txtHari.Text = "" Then
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        End If
        txtBulan.SetFocus
    End If
End Sub

Private Sub txtTelepon_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then dcPropinsi.SetFocus
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtTempatLahir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then meTglLahir.SetFocus
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'untuk cek validasi
Private Function funcCekValidasi() As Boolean
    If cboNamaDepan.Text = "" Then
        MsgBox "Titel Pasien harus diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        cboNamaDepan.SetFocus
        Exit Function
    End If
    If txtNamaPasien.Text = "" Then
        MsgBox "Nama Pasien harus diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        txtNamaPasien.SetFocus
        Exit Function
    End If
    If meTglLahir.Text = "__/__/____" Then
        MsgBox "Tanggal Lahir Pasien harus diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        meTglLahir.SetFocus
        Exit Function
    End If
    If cboJnsKelaminPasien.Text = "" Then
        MsgBox "Jenis Kelamin Pasien harus diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        cboJnsKelaminPasien.SetFocus
        Exit Function
    End If
    funcCekValidasi = True
End Function

'Store procedure untuk mengisi identitas pasien
Private Function sp_IdentitasPasien() As Boolean
    On Error GoTo errsp
    sp_IdentitasPasien = True
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtnocm.Text)
        .Parameters.Append .CreateParameter("NoIdentitas", adVarChar, adParamInput, 20, IIf(txtNoIdentitas.Text = "", Null, Trim(txtNoIdentitas.Text)))
        .Parameters.Append .CreateParameter("TglDaftarMembership", adDate, adParamInput, , Format(dtpTglPendaftaran.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TitlePasien", adVarChar, adParamInput, 4, cboNamaDepan.Text)
        .Parameters.Append .CreateParameter("NamaLengkap", adVarChar, adParamInput, 50, Trim(txtNamaPasien.Text))
        .Parameters.Append .CreateParameter("NamaPanggilan", adVarChar, adParamInput, 50, Trim(txtNamaPanggilan.Text))
        .Parameters.Append .CreateParameter("TempatLahir", adVarChar, adParamInput, 25, IIf(txtTempatLahir.Text = "", Null, Trim(txtTempatLahir.Text)))
        .Parameters.Append .CreateParameter("TglLahir", adDate, adParamInput, , Format(meTglLahir.Text, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("JenisKelamin", adChar, adParamInput, 1, Left(cboJnsKelaminPasien.Text, 1))
        .Parameters.Append .CreateParameter("Alamat", adVarChar, adParamInput, 100, IIf(txtalamat.Text = "", Null, Trim(txtalamat.Text)))
        .Parameters.Append .CreateParameter("Telepon", adVarChar, adParamInput, 15, IIf(txtTelepon.Text = "", Null, Trim(txtTelepon.Text)))
        .Parameters.Append .CreateParameter("Propinsi", adVarChar, adParamInput, 30, IIf(dcPropinsi.Text = "", Null, Trim(dcPropinsi.Text)))
        .Parameters.Append .CreateParameter("Kota", adVarChar, adParamInput, 50, IIf(dcKota.Text = "", Null, Trim(dcKota.Text)))
        .Parameters.Append .CreateParameter("Kecamatan", adVarChar, adParamInput, 50, IIf(dcKecamatan.Text = "", Null, Trim(dcKecamatan.Text)))
        .Parameters.Append .CreateParameter("Kelurahan", adVarChar, adParamInput, 50, IIf(dcKelurahan.Text = "", Null, Trim(dcKelurahan.Text)))
        .Parameters.Append .CreateParameter("RTRW", adVarChar, adParamInput, 5, IIf(meRTRW.Text = "__/__", Null, meRTRW.Text))
        .Parameters.Append .CreateParameter("KodePos", adChar, adParamInput, 5, IIf(txtKodePos.Text = "", Null, Trim(txtKodePos.Text)))
        .Parameters.Append .CreateParameter("OutputNoCM", adChar, adParamOutput, 6, Null)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AU_Pasien"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan data Pasien", vbCritical, "Validasi"
        Else
            MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
            If Not IsNull(.Parameters("OutputNoCM").Value) Then mstrNoCM = .Parameters("OutputNoCM").Value
            txtnocm.Text = mstrNoCM
            Call Add_HistoryLoginActivity("AU_Pasien")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
errsp:
    Call msubPesanError("sp_IdentitasPasien")
    Set dbcmd = Nothing
    sp_IdentitasPasien = False
End Function

'untuk membersihkan data pasien
Private Sub subClearData()
    txtnocm.Text = ""
    cboNamaDepan.ListIndex = -1
    txtNamaPasien.Text = ""
    txtNoIdentitas.Text = ""
    cboJnsKelaminPasien.ListIndex = -1
    txtTempatLahir.Text = ""
    meTglLahir.Text = "__/__/____"
    txtTahun.Text = ""
    txtBulan.Text = ""
    txtHari.Text = ""
    txtalamat.Text = ""
    meRTRW.Text = "__/__"
    txtTelepon.Text = ""
    dcPropinsi.Text = ""
    dcKota.Text = ""
    dcKecamatan.Text = ""
    dcKelurahan.Text = ""
    txtKodePos.Text = ""
End Sub

'untuk enable/disable button reg
Private Sub subEnableButtonReg(blnStatus As Boolean)
    cmdDetailPasien.Enabled = blnStatus
    cmdSimpan.Enabled = Not blnStatus
End Sub

'untuk enable/disable button reg
Private Sub subVisibleButtonReg(blnStatus As Boolean)
    cmdSimpan.Visible = blnStatus
End Sub

'untuk load data pasien yg sudah pernah didaftarkan
Private Sub subLoadDataPasien(strInput As String)
    Dim strSQLLoadPasien As String
    Dim rsLoadPasien As New ADODB.recordset
    strSQLLoadPasien = "SELECT * FROM Pasien WHERE NoCM = '" & strInput & "'"
    Set rsLoadPasien = Nothing
    rsLoadPasien.Open strSQLLoadPasien, dbConn, adOpenForwardOnly, adLockReadOnly
    txtnocm.Text = mstrNoCM
    dtpTglPendaftaran.MaxDate = Now
    dtpTglPendaftaran.Value = rsLoadPasien.Fields("TglDaftarMembership").Value
    cboNamaDepan.Text = rsLoadPasien.Fields("Title").Value
    txtNamaPasien.Text = rsLoadPasien.Fields("NamaLengkap").Value
    If IsNull(rsLoadPasien.Fields("NamaPanggilan")) Then txtNamaPanggilan.Text = "" Else txtNamaPanggilan.Text = rsLoadPasien.Fields("NamaPanggilan")
    If Not IsNull(rsLoadPasien.Fields("NoIdentitas").Value) Then txtNoIdentitas.Text = rsLoadPasien.Fields("NoIdentitas").Value
    If rsLoadPasien.Fields("JenisKelamin").Value = "L" Then
        cboJnsKelaminPasien.ListIndex = 0
    ElseIf rsLoadPasien.Fields("JenisKelamin").Value = "P" Then
        cboJnsKelaminPasien.ListIndex = 1
    End If
    If Not IsNull(rsLoadPasien.Fields("TempatLahir").Value) Then txtTempatLahir.Text = rsLoadPasien.Fields("TempatLahir").Value
    meTglLahir.Text = Format(rsLoadPasien.Fields("TglLahir").Value, "dd/mm/yyyy")
    If Not IsNull(rsLoadPasien.Fields("Alamat").Value) Then txtalamat.Text = rsLoadPasien.Fields("Alamat").Value
    If Not IsNull(rsLoadPasien.Fields("RTRW").Value) Then
        If Len(rsLoadPasien.Fields("RTRW").Value) = 5 And InStr(1, rsLoadPasien.Fields("RTRW").Value, "/") = 3 Then
            meRTRW.Text = rsLoadPasien.Fields("RTRW").Value
        Else
            If InStr(1, rsLoadPasien.Fields("RTRW").Value, "/") = 0 Then
                meRTRW.Text = Format(Left(rsLoadPasien.Fields("RTRW").Value, Len(rsLoadPasien.Fields("RTRW").Value) / 2), "00") & "/" & Format(Right(rsLoadPasien.Fields("RTRW").Value, Len(rsLoadPasien.Fields("RTRW").Value) / 2), "00")
            Else
                meRTRW.Text = Format(Left(rsLoadPasien.Fields("RTRW").Value, InStr(1, rsLoadPasien.Fields("RTRW").Value, "/") - 1), "00") & "/" & Format(Right(rsLoadPasien.Fields("RTRW").Value, Len(rsLoadPasien.Fields("RTRW").Value) - InStr(1, rsLoadPasien.Fields("RTRW").Value, "/")), "00")
            End If
        End If
    End If
    If Not IsNull(rsLoadPasien.Fields("Telepon").Value) Then txtTelepon.Text = rsLoadPasien.Fields("Telepon").Value
    If Not IsNull(rsLoadPasien.Fields("Propinsi").Value) Then dcPropinsi.Text = rsLoadPasien.Fields("Propinsi").Value
    If Not IsNull(rsLoadPasien.Fields("Kota").Value) Then dcKota.Text = rsLoadPasien.Fields("Kota").Value
    If Not IsNull(rsLoadPasien.Fields("Kecamatan").Value) Then dcKecamatan.Text = rsLoadPasien.Fields("Kecamatan").Value
    If Not IsNull(rsLoadPasien.Fields("Kelurahan").Value) Then dcKelurahan.Text = rsLoadPasien.Fields("Kelurahan").Value
    If Not IsNull(rsLoadPasien.Fields("KodePos").Value) Then txtKodePos.Text = rsLoadPasien.Fields("KodePos").Value
    Call meTglLahir_KeyPress(13)
    Set rsLoadPasien = Nothing
End Sub

Private Sub subLoadDataWilayah(strPencarian As String)
    On Error GoTo errLoad
    Dim strTempSql As String

    Select Case strPencarian
        Case "propinsi"
            If Len(Trim(dcPropinsi.Text)) = 0 Then Exit Sub
            strTempSql = " WHERE (NamaPropinsi LIKE '" & dcPropinsi.Text & "%') and StatusEnabled=1"

        Case "kota"
            If Len(Trim(dcKota.Text)) = 0 Then Exit Sub
            strTempSql = " WHERE (NamaKotaKabupaten LIKE '" & dcKota.Text & "%')and Expr1=1"

        Case "kecamatan"
            If Len(Trim(dcKecamatan.Text)) = 0 Then Exit Sub
            strTempSql = " WHERE (NamaKecamatan LIKE '" & dcKecamatan.Text & "%')and Expr2=1"

        Case "desa"
            If Len(Trim(dcKelurahan.Text)) = 0 Then Exit Sub
            strTempSql = " WHERE (NamaKelurahan LIKE '" & dcKelurahan.Text & "%')and  Expr3=1"

        Case "kodepos"
            If Len(Trim(txtKodePos.Text)) = 0 Then Exit Sub
            strTempSql = " WHERE (KodePos LIKE '" & txtKodePos.Text & "%')"

    End Select

    strSQL = "SELECT DISTINCT ISNULL(NamaPropinsi, '-') AS NamaPropinsi, ISNULL(NamaKotaKabupaten, '-') AS NamaKotaKabupaten, ISNULL(NamaKecamatan, '-')  AS NamaKecamatan, ISNULL(NamaKelurahan, '-') AS NamaKelurahan, ISNULL(KodePos, '-') AS KodePos" & _
    " FROM V_Wilayah" & _
    " " & strTempSql

    Call msubRecFO(rs, strSQL)
    If rs.EOF Then
    Else
        dcPropinsi.BoundText = rs("NamaPropinsi")
        dcKota.BoundText = rs("NamaKotaKabupaten")
        dcKecamatan.BoundText = rs("NamaKecamatan")
        dcKelurahan.BoundText = rs("NamaKelurahan")
        txtKodePos.Text = rs("KodePos")
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

'Private Sub subDcSource(varstrPilihan As String, Optional varStrSQL As String)
'On Error GoTo errLoad
''
''    strSQL = "SELECT DISTINCT NamaPropinsi, NamaPropinsi AS alias FROM V_Wilayah where StatusEnabled=1"
''    Call msubDcSource(dcPropinsi, rs, strSQL)
''
''    strSQL = "SELECT DISTINCT NamaKotaKabupaten, NamaKotaKabupaten AS alias FROM V_Wilayah where Expr1=1"
''    Call msubDcSource(dcKota, rs, strSQL)
''
''    strSQL = "SELECT DISTINCT NamaKecamatan, NamaKecamatan AS alias FROM V_Wilayah where Expr2=1"
''    Call msubDcSource(dcKecamatan, rs, strSQL)
''
''    strSQL = "SELECT DISTINCT NamaKelurahan, NamaKelurahan AS alias FROM V_Wilayah where Expr3=1"
''    Call msubDcSource(dcKelurahan, rs, strSQL)
''
''Exit Sub
' Select Case varstrPilihan
'
'        Case "Propinsi"
'            strSQL = "SELECT DISTINCT KdPropinsi, NamaPropinsi AS alias FROM V_Wilayah where StatusEnabled=1 order by NamaPropinsi"
'            Set rsPropinsi = Nothing
'            rsPropinsi.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
'            Set dcPropinsi.RowSource = rsPropinsi
'            dcPropinsi.BoundColumn = rsPropinsi(0).Name
'            dcPropinsi.ListField = rsPropinsi(1).Name
'        Case "Kota"
'            strSQL = "SELECT DISTINCT KdKotaKabupaten, NamaKotaKabupaten AS alias FROM V_Wilayah " & varStrSQL & ""
'            Set rsKota = Nothing
'            rsKota.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
'            Set dcKota.RowSource = rsKota
'            dcKota.BoundColumn = rsKota(0).Name
'            dcKota.ListField = rsKota(1).Name
'
'            If rsKota.EOF = False Then
'            dcKota.BoundText = rsKota(0)
'            End If
'        Case "Kecamatan"
'           If dcKota.BoundText = "" Then Exit Sub
'            strSQL = "SELECT DISTINCT KdKecamatan, NamaKecamatan AS alias FROM V_Wilayah " & varStrSQL & ""
'            Set rsKecamatan = Nothing
'            rsKecamatan.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
'            Set dcKecamatan.RowSource = rsKecamatan
'            dcKecamatan.BoundColumn = rsKecamatan(0).Name
'            dcKecamatan.ListField = rsKecamatan(1).Name
'            If rsKecamatan(0) <> Null Or rsKecamatan(0) <> "" Then dcKecamatan.BoundText = rsKecamatan(0)
'
'        Case "Kelurahan"
'            strSQL = "SELECT DISTINCT KdKelurahan, NamaKelurahan AS alias FROM V_Wilayah " & varStrSQL & ""
'            Set rsKelurahan = Nothing
'            rsKelurahan.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
'            Set dcKelurahan.RowSource = rsKelurahan
'            dcKelurahan.BoundColumn = rsKelurahan(0).Name
'            dcKelurahan.ListField = rsKelurahan(1).Name
'            If rsKelurahan.EOF = True Then
'                If rsKelurahan(0) <> Null Or rsKelurahan(0) <> "" Then dcKelurahan.BoundText = rsKelurahan(0)
'                Else
'            End If
'    End Select
'
'    Exit Sub
'errLoad:
'    Call msubPesanError
'
'End Sub
Private Sub subDcSource(varstrPilihan As String, Optional varStrSQL As String)
    On Error GoTo errLoad
    
    Select Case varstrPilihan
    
        Case "Propinsi"
            strSQL = "SELECT DISTINCT NamaPropinsi, NamaPropinsi AS alias FROM V_Wilayah"
            Call msubDcSource(dcPropinsi, rs, strSQL)

        Case "Kota"
            strSQL = "SELECT DISTINCT NamaKotaKabupaten, NamaKotaKabupaten AS alias FROM V_Wilayah"
            Call msubDcSource(dcKota, rs, strSQL)

        Case "Kecamatan"
            strSQL = "SELECT DISTINCT NamaKecamatan, NamaKecamatan AS alias FROM V_Wilayah"
            Call msubDcSource(dcKecamatan, rs, strSQL)

        Case "Kelurahan"
            strSQL = "SELECT DISTINCT NamaKelurahan, NamaKelurahan AS alias FROM V_Wilayah"
            Call msubDcSource(dcKelurahan, rs, strSQL)
    
    End Select
    
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub CekPilihanWilayah(strItem As String, Optional strEvent As String)
    Dim X As Integer
    Dim y

    X = 0
    Select Case strItem
        Case "dcPropinsi"
            Set dcKota.RowSource = Nothing
            Set dcKecamatan.RowSource = Nothing
            Set dcKelurahan.RowSource = Nothing
            dcKota.Text = ""
            dcKecamatan.Text = ""
            dcKelurahan.Text = ""
            txtKodePos = ""
            Select Case strEvent
                Case "Click"
                    subDcSource "Kota", " where NamaPropinsi = '" & dcPropinsi.Text & "' order by NamaKotaKabupaten"
                Case "KeyPress"
                    If dcPropinsi.MatchedWithList = False Then
                        MsgBox "Pilih Propinsi"
                        X = 1
                        GoTo kosong
                        dcPropinsi.SetFocus
                    Else
                        subDcSource "Kota", " where kdPropinsi = '" & dcPropinsi.BoundText & "' order by NamaKotaKabupaten"
                        dcKota.SetFocus
                    End If
                Case "LostFocus"
                    If dcPropinsi.MatchedWithList = False Then
                        MsgBox "Pilih Propinsi"
                        X = 1
                        GoTo kosong
                        dcPropinsi.SetFocus
                    Else
                        subDcSource "Kota", " where kdPropinsi = '" & dcPropinsi.BoundText & "' order by NamaKotaKabupaten"
                        dcKota.SetFocus
                    End If
                   
            End Select
        Case "dcKota"
            Set dcKecamatan.RowSource = Nothing
            Set dcKelurahan.RowSource = Nothing
            dcKecamatan.Text = ""
            dcKelurahan.Text = ""
            txtKodePos = ""
            If dcPropinsi.MatchedWithList = True Then
                Select Case strEvent
                    Case "Click"
                        subDcSource "Kecamatan", " where KdKotaKabupaten = '" & dcKota.BoundText & "' order by NamaKecamatan"
                    Case "KeyPress"
                        If dcKota.MatchedWithList = False Then
                           MsgBox "Pilih Kota"
                            X = 2
                            GoTo kosong
                            dcKota.SetFocus
                        Else
                            subDcSource "Kecamatan", " where kdKotaKabupaten = '" & dcKota.BoundText & "' order by NamaKecamatan"
                            dcKecamatan.SetFocus
                        End If
                    Case "LostFocus"
                        If dcKota.MatchedWithList = False Then
                            MsgBox "Pilih Kota"
                            X = 2
                            GoTo kosong
                            dcKota.SetFocus
                        Else
                            subDcSource "Kecamatan", " where kdKotaKabupaten = '" & dcKota.BoundText & "' order by NamaKecamatan"
                            dcKecamatan.SetFocus
                        End If
                End Select
            End If
        Case "dcKecamatan"
            Set dcKelurahan.RowSource = Nothing
            dcKelurahan.Text = ""
            txtKodePos = ""
            If dcKota.MatchedWithList = True Then
                Select Case strEvent
                    Case "Click"
                        If dcKecamatan.Text = "" Then Exit Sub
                        subDcSource "Kelurahan", " where kdkecamatan = '" & dcKecamatan.BoundText & "' order by NamaKelurahan"
                    Case "KeyPress"
                        If dcKecamatan.MatchedWithList = False Then
                            MsgBox "Pilih Kecamatan"
                            X = 3
                            GoTo kosong
                            dcKecamatan.SetFocus
                        Else
                            subDcSource "Kelurahan", " where kdkecamatan = '" & dcKecamatan.BoundText & "' order by NamaKelurahan"
                            dcKelurahan.SetFocus
                        End If
                    Case "LostFocus"
                        If dcKecamatan.MatchedWithList = False Then
                            MsgBox "Pilih Kecamatan"
                            X = 3
                            GoTo kosong
                            dcKecamatan.SetFocus
                        Else
                            subDcSource "Kelurahan", " where kdkecamatan = '" & dcKecamatan.BoundText & "' order by NamaKelurahan"
                            dcKelurahan.SetFocus
                        End If
                End Select
            End If
        Case "dcKelurahan"
            txtKodePos = ""
            If dcKecamatan.MatchedWithList = True Then
                Select Case strEvent
                    Case "KeyPress"
                        If dcKelurahan.MatchedWithList = False Then
                            MsgBox "Pilih Desa/Kelurahan"
                            X = 4
                            GoTo kosong
                            dcKelurahan.Text = ""
                            dcKelurahan.SetFocus
                        Else
                            txtKodePos.SetFocus
                        End If
                    Case "LostFocus"
                        If dcKelurahan.MatchedWithList = False Then
                            MsgBox "Pilih Desa/Kelurahan"
                            X = 4
                            GoTo kosong
                            dcKelurahan.SetFocus
                        End If
                End Select
            End If
    End Select

    Exit Sub

kosong:
    y = MsgBox("Mulai lagi dari awal", vbYesNo, "Wilayah") ' vbYesNoCancel
    Select Case y
        Case vbYes
            dcPropinsi.Text = ""
            dcKota.Text = ""
            dcKecamatan.Text = ""
            dcKelurahan.Text = ""
            dcPropinsi.SetFocus
        Case vbNo
            Exit Sub
'            Select Case X
'                Case 1
'                    dcPropinsi.SetFocus
'                Case 2
'                    dcKota.SetFocus
'                Case 3
'                    dcKecamatan.SetFocus
'                Case 4
'                    dcKelurahan.SetFocus
'            End Select
'        Case vbCancel
'            Exit Sub
    End Select
End Sub
