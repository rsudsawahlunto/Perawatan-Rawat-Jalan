VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmRegistrasiRJ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Registrasi Poliklinik"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10350
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegistrasiRJ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   10350
   Begin VB.Frame fraDokter 
      Caption         =   "Data Dokter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   2400
      TabIndex        =   21
      Top             =   1200
      Visible         =   0   'False
      Width           =   7695
      Begin MSDataGridLib.DataGrid dgDokter 
         Height          =   1815
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   3201
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   16
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   30
      Top             =   3840
      Width           =   10335
      Begin VB.CommandButton cmTutup2 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   8445
         TabIndex        =   18
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdRujukan 
         Caption         =   "&Data Rujukan"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   5055
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "&Lanjutkan"
         Height          =   375
         Left            =   6750
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame fraDataRegistrasiRJ 
      Caption         =   "Data Registrasi Rawat Jalan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   22
      Top             =   2040
      Width           =   10335
      Begin VB.TextBox txtTglMasuk 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtKelompok 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtRuangan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   5640
         TabIndex        =   11
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtKelas 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   5520
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtJenisKelas 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1920
         TabIndex        =   9
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtDokter 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6120
         TabIndex        =   14
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox txtKdDokter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7440
         TabIndex        =   37
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSDataListLib.DataCombo dcRujukanAsal 
         Height          =   330
         Left            =   3000
         TabIndex        =   13
         Top             =   1320
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSDataListLib.DataCombo DcKelasPelayanan 
         Height          =   330
         Left            =   4080
         TabIndex        =   41
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
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
      Begin MSDataListLib.DataCombo DcKasusPenyakit 
         Height          =   330
         Left            =   7800
         TabIndex        =   42
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "SMF (Kasus Penyakit)"
         Height          =   210
         Left            =   7800
         TabIndex        =   43
         Top             =   360
         Width           =   1740
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Rujukan Dari"
         Height          =   210
         Left            =   3000
         TabIndex        =   39
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelas Pelayanan"
         Height          =   210
         Left            =   1920
         TabIndex        =   38
         Top             =   360
         Width           =   1725
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Kelompok Pasien"
         Height          =   210
         Left            =   240
         TabIndex        =   32
         Top             =   1080
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Pendaftaran"
         Height          =   210
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Kelas Pelayanan"
         Height          =   210
         Left            =   4080
         TabIndex        =   25
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Dokter Pemeriksa"
         Height          =   210
         Left            =   6120
         TabIndex        =   24
         Top             =   1080
         Width           =   1425
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Ruang Pemeriksaan"
         Height          =   210
         Left            =   5640
         TabIndex        =   23
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   1095
      Left            =   0
      TabIndex        =   26
      Top             =   960
      Width           =   10335
      Begin VB.CheckBox chkDetailPasien 
         Caption         =   "Detail Pasien"
         Enabled         =   0   'False
         Height          =   255
         Left            =   8880
         TabIndex        =   7
         Top             =   120
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         Caption         =   "Umur"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7680
         TabIndex        =   19
         Top             =   320
         Width           =   2535
         Begin VB.TextBox txtHr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            MaxLength       =   6
            TabIndex        =   6
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtBln 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   960
            MaxLength       =   6
            TabIndex        =   5
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtThn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            MaxLength       =   6
            TabIndex        =   4
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            Height          =   210
            Left            =   2280
            TabIndex        =   35
            Top             =   292
            Width           =   165
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            Height          =   210
            Left            =   1440
            TabIndex        =   34
            Top             =   292
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            Height          =   210
            Left            =   600
            TabIndex        =   33
            Top             =   292
            Width           =   285
         End
      End
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   6480
         MaxLength       =   9
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         MaxLength       =   10
         TabIndex        =   0
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label lblJnsKlm 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   6480
         TabIndex        =   36
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "No. Registrasi"
         Height          =   210
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   1800
         TabIndex        =   28
         Top             =   360
         Width           =   585
      End
      Begin VB.Label lblNamaPasien 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   3600
         TabIndex        =   27
         Top             =   360
         Width           =   1020
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   40
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
      Left            =   8520
      Picture         =   "frmRegistrasiRJ.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRegistrasiRJ.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRegistrasiRJ.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmRegistrasiRJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFilter As String
Dim intRowNow As Integer
Dim strSubInstalasi As String
Dim strNoAntrian As String
Dim booldokter As Boolean

Private Sub chkDetailPasien_Click()
    If chkDetailPasien.Value = 1 Then
        strPasien = "View"
        Load frmPasienBaru
        frmPasienBaru.Show
    Else
        Unload frmPasienBaru
        Unload frmDetailPasien
    End If
End Sub

Private Sub chkDetailPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTglMasuk.SetFocus
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo hell
    If funcCekValidasi = False Then Exit Sub
    
    If DcKelasPelayanan.BoundText <> "" Then
    dbConn.Execute "Update PasienMasukRumahSakit set StatusPeriksa='Y', KdKelas='" & DcKelasPelayanan.BoundText & "' where NoPendaftaran='" & txtNoPendaftaran.Text & "'"
    dbConn.Execute "Update RegistrasiRJ set KdKelas='" & DcKelasPelayanan.BoundText & "' where NoPendaftaran='" & txtNoPendaftaran.Text & "'"
    dbConn.Execute "Update PasienDaftar set KdKelasAkhir='" & DcKelasPelayanan.BoundText & "' where NoPendaftaran='" & txtNoPendaftaran.Text & "'"
    
    End If
    If dcKasusPenyakit.BoundText <> "" Then
    dbConn.Execute "Update RegistrasiRJ set KdSubInstalasi='" & dcKasusPenyakit.BoundText & "'  where NoPendaftaran='" & txtNoPendaftaran.Text & "'"
    dbConn.Execute "Update PasienMasukRumahSakit set KdSubInstalasi='" & dcKasusPenyakit.BoundText & "' where NoPendaftaran='" & txtNoPendaftaran.Text & "'"
    End If
    If sp_RegistrasiRJ() = False Then Exit Sub
    Call subEnableButtonReg(True)
   Dim Path As String
    
       strSQL = "select Value from SettingGlobal where Prefix='PathSdkAntrian'"
    Call msubRecFO(rs, strSQL)
      
    If Not rs.EOF Then
        If rs(0).Value <> "" Then
            Path = rs(0).Value
        End If
    End If
    
    strSQL = "select StatusAntrian from SettingDataUmum"
    Call msubRecFO(rs, strSQL)
    Dim coba As Long
    
    If Not rs.EOF Then
        If rs(0).Value = "1" Then
           If Dir(Path) <> "" Then
                strSQL = "select * from settingglobal where Prefix like 'KdRuanganAntrian%' and Value='" & mstrKdRuangan & "'"
                Dim prefix As String
                prefix = ""
                Call msubRecFO(rs, strSQL)
                If (rs.EOF = False) Then
                    If (rs("Prefix").Value = "KdRuanganAntrianKanan") Then
                         prefix = "Kanan-"
                    ElseIf (rs("Prefix").Value = "KdRuanganAntrianKiri") Then
                        prefix = "Kiri-"
                    Else
                    End If
                End If
                Path = Path + " Type:" & Chr(34) & "Update Patient" & Chr(34) & " NoAntrian:" & prefix & strNoAntrian & " loket:" & mstrKdRuangan
                coba = Shell(Path, vbNormalFocus)
           End If
        End If
    End If
Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    On Error GoTo errLoad
    If Periksa("datacombo", dcRujukanAsal, "Data rujukan asal kosong") = False Then Exit Sub
    If Periksa("text", txtDokter, "Nama dokter kosong") = False Then Exit Sub
    If cmdSimpan.Enabled = True Then
        If MsgBox("Simpan data registrasi Rawat Jalan?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
        Else
            Exit Sub
        End If
    End If

    Call subLoadFormTP
    Unload Me
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadFormTP()
    On Error GoTo hell

    mstrNoPen = txtNoPendaftaran.Text
    mstrNoCM = txtNoCM.Text
    mstrKdSubInstalasi = mstrKdSubInstalasi
    With frmTransaksiPasien
        .Show
        .txtNoPendaftaran.Text = mstrNoPen
        .txtNoCM.Text = mstrNoCM
        .txtNamaPasien.Text = txtNamaPasien
        .txtSex.Text = txtJK.Text
        .txtThn.Text = txtThn.Text
        .txtBln.Text = txtBln.Text
        .txtHr.Text = txtHr.Text
        .txtKls.Text = DcKelasPelayanan.Text 'txtKelas.Text
        .txtJenisPasien.Text = txtKelompok.Text

        .txtTglDaftar.Text = txtTglMasuk.Text
    End With
    Exit Sub
hell:
End Sub

Private Sub cmTutup2_Click()
    If cmdSimpan.Enabled = True Then
        If MsgBox("Simpan data registrasi Rawat Jalan", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub dcRujukanAsal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcRujukanAsal.MatchedWithList = True Then txtDokter.SetFocus
        strSQL = "SELECT KdRujukanAsal,RujukanAsal FROM RujukanAsal WHERE (RujukanAsal LIKE '%" & dcRujukanAsal.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcRujukanAsal.Text = ""
            txtDokter.SetFocus
            Exit Sub
        End If
        dcRujukanAsal.BoundText = rs(0).Value
        dcRujukanAsal.Text = rs(1).Value
    End If
End Sub

Private Sub dgDokter_DblClick()
    Call dgDokter_KeyPress(13)
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlDokter = 0 Then Exit Sub
        txtDokter.Text = dgDokter.Columns(0).Value
        txtKdDokter.Text = dgDokter.Columns(1).Value
        mstrKdDokter = dgDokter.Columns(1).Value
        If txtKdDokter.Text = "" Then
            MsgBox "Pilih dulu Dokter yang akan menangani Pasien", vbCritical, "Validasi"
            txtDokter.Text = ""
            dgDokter.SetFocus
            Exit Sub
        End If
        fraDokter.Visible = False
        cmdSimpan.SetFocus
        Me.Height = 5010
        Call centerForm(Me, MDIUtama)
    End If
End Sub

Private Sub dtpTglPendaftaran_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcRujukanAsal.SetFocus
End Sub

Private Sub dgDokter_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If dgDokter.Visible = False Then Exit Sub
        txtDokter.SetFocus
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo hell
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    strRegistrasi = "RJ"
    strSQL = "SELECT distinct KdKelas, Kelas FROM V_KelasPelayanan WHERE KdInstalasi='" & mstrKdInstalasiLogin & "' and KdKelas<>04 and Expr2='1'"
    Call msubDcSource(DcKelasPelayanan, rs, strSQL)
      
    strSQL = "SELECT DISTINCT KdSubInstalasi, NamaSubInstalasi FROM V_SubInstalasiRuangan WHERE KdInstalasi='" & mstrKdInstalasiLogin & "' and StatusEnabled='1' and KdRuangan='" & mstrKdRuangan & "'  ORDER BY NamaSubInstalasi"
    Call msubDcSource(dcKasusPenyakit, rs, strSQL)
    
    strSQL = "SELECT KdRujukanAsal,RujukanAsal FROM RujukanAsal "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcRujukanAsal.RowSource = rs
    dcRujukanAsal.ListField = rs.Fields(1).Name
    dcRujukanAsal.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
    Call subTampilData(mstrNoPen)
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmDaftarAntrianPasien.Enabled = True
    strSQL = "select StatusPeriksa from PasienMasukRumahSakit  where NoPendaftaran='" + mstrNoPen + "'"
    Call msubRecFO(rs, strSQL)
    If (rs.EOF = False) Then
        If (rs(0) = "S") Then dbConn.Execute ("update PasienMasukRumahSakit set StatusPeriksa='T' where NoPendaftaran='" + mstrNoPen + "'")
        Dim Path As String
       strSQL = "select Value from SettingGlobal where Prefix='PathSdkAntrian'"
    Call msubRecFO(rs, strSQL)
    Dim pathtemp As String
    If Not rs.EOF Then
        If rs(0).Value <> "" Then
            pathtemp = rs(0).Value
        End If
    End If
    
    strSQL = "select StatusAntrian from SettingDataUmum"
    Call msubRecFO(rs, strSQL)
    Dim coba As Long
    
    If Not rs.EOF Then
        If rs(0).Value = "1" Then
        
        If Dir(Path) <> "" Then
                strSQL = "select * from settingglobal where Prefix like 'KdRuanganAntrian%' and Value='" & mstrKdRuangan & "'"
                Dim prefix As String
                 prefix = ""
                Call msubRecFO(rs, strSQL)
                If (rs.EOF = False) Then
                    If (rs("Prefix").Value = "KdRuanganAntrianKanan") Then
                         prefix = "Kanan-"
                    ElseIf (rs("Prefix").Value = "KdRuanganAntrianKiri") Then
                        prefix = "Kiri-"
                    Else
                    End If
                End If
                      
            Path = pathtemp + "  Type:" & Chr(34) & "Update Patient" & Chr(34) & " loket:" & mstrKdRuangan
            coba = Shell(Path, vbNormalFocus)
            Call frmDaftarAntrianPasien.cmdCari_Click
        End If
        End If
    End If
    End If
End Sub

Private Sub txtDokter_Change()
If booldokter = False Then
        Call subLoadDokter
End If
End Sub

Private Sub txtDokter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = 13 Then
        If dgDokter.Visible = False Then cmdSimpan.SetFocus: Exit Sub
        dgDokter.SetFocus
    End If
End Sub

Private Sub txtDokter_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then
        If intJmlDokter = 0 Then Exit Sub
        dgDokter.SetFocus
    End If
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 27 Then
        fraDokter.Visible = False
        Me.Height = 5010
    End If
    Exit Sub
errLoad:
End Sub

Private Sub txtNoCM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkDetailPasien.Enabled = True Then chkDetailPasien.SetFocus
    End If
End Sub

Private Sub txtNoPendaftaran_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then Call subTampilData(txtNoPendaftaran)
End Sub
Public Sub subNoAntrian(noAntrian As String)
strNoAntrian = noAntrian
End Sub
Public Sub subTampilData(strNoPenndaftaran As String)
    On Error GoTo errLoad

    Call subClearData
    Call subEnableButtonReg(False)
    strSQL = "Select * from V_DaftarAntrianPasienMRS WHERE NoPendaftaran ='" & mstrNoPen & "' AND Ruangan = '" & strNNamaRuangan & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        mstrNoCM = ""
        mstrNoPen = ""
        chkDetailPasien.Enabled = False
        cmdSimpan.Enabled = False
        Exit Sub
    End If
    txtNoCM.Text = rs("NoCM")
    mstrNoCM = txtNoCM.Text
    txtNamaPasien.Text = rs.Fields("Nama Pasien").Value
    If rs.Fields("JK").Value = "P" Then
        txtJK.Text = "Perempuan"
    ElseIf rs.Fields("JK").Value = "L" Then
        txtJK.Text = "Laki-laki"
    End If
    DcKelasPelayanan.BoundText = rs.Fields("Kdkelas").Value
    dcKasusPenyakit.BoundText = rs.Fields("KdSubInstalasi").Value
    txtThn.Text = rs.Fields("UmurTahun").Value
    txtBln.Text = rs.Fields("UmurBulan").Value
    txtHr.Text = rs.Fields("UmurHari").Value

    txtTglMasuk.Text = rs("TglMasuk")
    txtJenisKelas.Text = ""
    txtKelas.Text = rs("Kelas")
    txtRuangan.Text = rs("Ruangan")
    dcRujukanAsal.BoundText = ""
    txtKelompok.Text = ""

    mdTglMasuk = txtTglMasuk.Text
    mstrKdKelas = rs("KdKelas")
    mstrKelas = rs("Kelas")
    
    
    If Not IsNull(rs.Fields("NamaDokter")) Or rs.Fields("NamaDokter").Value <> "" Then
        booldokter = True
        txtKdDokter.Text = rs.Fields("IdDokter").Value
        txtDokter.Text = rs.Fields("NamaDokter").Value
        fraDokter.Visible = False
        booldokter = False
    End If
   
    strSQL = "SELECT dbo.PasienDaftar.KdKelompokPasien, dbo.KelompokPasien.JenisPasien, dbo.DetailJenisJasaPelayanan.DetailJenisJasaPelayanan, dbo.JenisJasaPelayanan.JenisJasaPelayanan, dbo.PasienDaftar.NoPendaftaran" & _
    " FROM dbo.PasienDaftar INNER JOIN dbo.DetailJenisJasaPelayanan ON dbo.PasienDaftar.KdDetailJenisJasaPelayanan = dbo.DetailJenisJasaPelayanan.KdDetailJenisJasaPelayanan INNER JOIN dbo.KelompokPasien ON dbo.PasienDaftar.KdKelompokPasien = dbo.KelompokPasien.KdKelompokPasien INNER JOIN dbo.JenisJasaPelayanan ON dbo.DetailJenisJasaPelayanan.KdJenisJasaPelayanan = dbo.JenisJasaPelayanan.KdJenisJasaPelayanan" & _
    " WHERE (dbo.PasienDaftar.NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rs, strSQL)
    If Not rs.EOF Then
        txtJenisKelas.Text = rs("JenisJasaPelayanan")
        txtKelompok.Text = rs("JenisPasien")
    End If

    chkDetailPasien.Enabled = True
    strSQL = "SELECT KdRujukanAsal FROM RegistrasiRJ WHERE (NoPendaftaran = '" & strNoPenndaftaran & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then dcRujukanAsal.BoundText = rs(0) Else dcRujukanAsal.BoundText = "01"
'    txtDokter.SetFocus

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

'untuk meload data dokter di grid
Private Sub subLoadDokter()
    On Error GoTo errLoad
    strSQL = "SELECT NamaDokter AS [Nama Dokter],KodeDokter AS [Kode Dokter],JK,Jabatan FROM V_DaftarDokter where namadokter like '%" & txtDokter.Text & "%' "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlDokter = rs.RecordCount
    Set dgDokter.DataSource = rs
    With dgDokter
        .Columns(0).Width = 3000 'nama dokter
        .Columns(1).Width = 0 'kode dokter
    End With
    fraDokter.Visible = True
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subEnableButtonReg(blnStatus As Boolean)
    cmdRujukan.Enabled = blnStatus
    cmdSimpan.Enabled = Not blnStatus
    txtDokter.Enabled = Not blnStatus
    DcKelasPelayanan.Enabled = Not blnStatus
    dcKasusPenyakit.Enabled = Not blnStatus
End Sub

'Store procedure untuk mengisi registrasi pasien
Private Function sp_RegistrasiRJ() As Boolean
    On Error GoTo hell
    sp_RegistrasiRJ = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(txtTglMasuk, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, txtKdDokter.Text)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, noidpegawai)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_RegistrasiPasienMasukRJ"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Pendaftaran Pasien ke Instalasi Rawat Jalan", vbCritical, "Validasi"
            sp_RegistrasiRJ = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
hell:
    Call msubPesanError("sp_RegistrasiRJ")
    sp_RegistrasiRJ = False
    Set dbcmd = Nothing
End Function

'untuk cek validasi
Private Function funcCekValidasi() As Boolean
    If txtNamaPasien.Text = "" Then
        MsgBox "No. CM Harus Diisi", vbExclamation, "Validasi"
        funcCekValidasi = False
        txtNoCM.SetFocus
        Exit Function
    End If
    If txtKdDokter.Text = "" Then
        MsgBox "Pilihan Dokter harus diisi sesuai data daftar dokter", vbExclamation, "Validasi"
        funcCekValidasi = False
        txtDokter.SetFocus
        Exit Function
    End If
    funcCekValidasi = True
End Function

'untuk membersihkan data pasien registrasi
Private Sub subClearData()
    txtNoCM.Text = ""
    txtNamaPasien.Text = ""
    txtJK.Text = ""
    txtThn.Text = ""
    txtBln.Text = ""
    txtHr.Text = ""
    dcRujukanAsal.Text = ""
    txtDokter.Text = ""
    txtKdDokter.Text = ""
End Sub

