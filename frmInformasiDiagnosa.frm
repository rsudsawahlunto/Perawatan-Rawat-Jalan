VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmInformasiDiagnosa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informasi Diagnosa Pasien"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInformasiDiagnosa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   14715
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6195
      Left            =   0
      TabIndex        =   12
      Top             =   960
      Width           =   14715
      Begin VB.CheckBox chkStatusPeriksa 
         Caption         =   "Status Periksa"
         Height          =   255
         Left            =   6840
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox combDiagnosa 
         Enabled         =   0   'False
         Height          =   330
         ItemData        =   "frmInformasiDiagnosa.frx":0CCA
         Left            =   6840
         List            =   "frmInformasiDiagnosa.frx":0CD4
         TabIndex        =   6
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton optPasienKonsul 
         Caption         =   "Pasien Konsul"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton optPasienPoli 
         Caption         =   "Pasien Poliklinik"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   4935
      End
      Begin VB.Frame Frame3 
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
         Left            =   8760
         TabIndex        =   13
         Top             =   150
         Width           =   5775
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker DTPickerAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   0
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   103219203
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin MSComCtl2.DTPicker DTPickerAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   1
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   103219203
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   14
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgData 
         Height          =   4935
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   8705
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   2
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
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
      Begin VB.Label lblJumData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data 0/0"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cari Data Pasien"
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
      TabIndex        =   11
      Top             =   7200
      Width           =   14685
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   450
         Width           =   2655
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "Ceta&k"
         Height          =   495
         Left            =   11040
         TabIndex        =   9
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   12840
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan  Nama Pasien / No.CM"
         Height          =   210
         Left            =   1560
         TabIndex        =   15
         Top             =   200
         Width           =   2640
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   17
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
      Picture         =   "frmInformasiDiagnosa.frx":0CE6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   12840
      Picture         =   "frmInformasiDiagnosa.frx":36A7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmInformasiDiagnosa.frx":442F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12975
   End
End
Attribute VB_Name = "FrmInformasiDiagnosa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkStatusPeriksa_Click()
    If chkStatusPeriksa.Value = vbChecked Then
        combDiagnosa.Enabled = True
    Else
        combDiagnosa.Enabled = False
    End If
End Sub

Private Sub chkStatusPeriksa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkStatusPeriksa.Value = vbChecked Then
            combDiagnosa.Enabled = True
            combDiagnosa.SetFocus
        Else
            combDiagnosa.Enabled = False
            cmdCari.SetFocus
        End If
    End If
End Sub

Private Sub cmdCari_Click()
    On Error GoTo errLoad

    lblJumData.Caption = "Data 0/0"
    'Jika kondisi pasien poli belum di diagnosa
    If (optPasienPoli.Value = True And combDiagnosa.Enabled = True And combDiagnosa.Text = "BELUM") Then
        Set rs = Nothing
        cetak = "PasienPoliBelum"
        strSQL = "select TOP 100 TglMasuk, NoPendaftaran, NoCM, [Nama Pasien], Umur, JK, Diagnosa, [Dokter Pemeriksa], Ruangan " _
        & "from V_DaftarPasienLamaRJ " _
        & "WHERE ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%') and Ruangan='" & strNNamaRuangan & "' and TglMasuk BETWEEN '" & Format(DTPickerAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd HH:mm:59") & "'and  Diagnosa is Null   " 'and  Diagnosa is Null

        rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
        Set dgData.DataSource = rs
        Call subSetGrid
        'jika kodisi pasien poli sudah di diagnosa
    ElseIf (optPasienPoli.Value = True And combDiagnosa.Enabled = True And combDiagnosa.Text = "SUDAH") Then
        Set rs = Nothing
        cetak = "PasienPoliSudah"
        strSQL = "select TOP 100 TglMasuk, NoPendaftaran, NoCM, [Nama Pasien], Umur, JK, Diagnosa, [Dokter Pemeriksa], Ruangan " _
        & "from V_DaftarPasienLamaRJ " _
        & "WHERE ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%') and Ruangan='" & strNNamaRuangan & "' and TglMasuk BETWEEN '" & Format(DTPickerAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' and  Diagnosa is Not Null  "

        rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
        Set dgData.DataSource = rs
        Call subSetGrid
        'jika combo diagnosa =""
    ElseIf (optPasienPoli.Value = True And combDiagnosa.Enabled = True And combDiagnosa.Text = "") Then
        Set rs = Nothing
        cetak = "PasienPoli1"
        strSQL = "select TOP 100 TglMasuk, NoPendaftaran, NoCM, [Nama Pasien], Umur, JK, Diagnosa, [Dokter Pemeriksa], Ruangan " _
        & "from V_DaftarPasienLamaRJ " _
        & "WHERE ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%') and Ruangan='" & strNNamaRuangan & "' and TglMasuk BETWEEN '" & Format(DTPickerAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' "

        rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
        Set dgData.DataSource = rs
        Call subSetGrid
        'jika semua pasien poli yang diinginkan
    ElseIf optPasienPoli.Value = True And combDiagnosa.Enabled = False Then
        Set rs = Nothing
        cetak = "PasienPoli2"
        strSQL = "select TOP 100 TglMasuk, NoPendaftaran, NoCM, [Nama Pasien], Umur, JK, Diagnosa, [Dokter Pemeriksa], Ruangan " _
        & "from V_DaftarPasienLamaRJ " _
        & "WHERE ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%') and Ruangan='" & strNNamaRuangan & "' and TglMasuk BETWEEN '" & Format(DTPickerAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' "

        rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
        Set dgData.DataSource = rs
        Call subSetGrid
        'Jika kondisi pasien konsul belum di diagnosa
    ElseIf (optPasienKonsul.Value = True And combDiagnosa.Enabled = True And combDiagnosa.Text = "BELUM") Then
        Set rs = Nothing
        cetak = "PasienKonsulBelum"
        strSQL = "select TOP 100 TglDirujuk, NoPendaftaran, NoCM, [Nama Pasien], Umur, JK, Diagnosa, [Dokter Perujuk], [Ruangan Tujuan] " _
        & "from V_DaftarPasienKonsul " _
        & "WHERE ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%') and [Ruangan Tujuan]='" & strNNamaRuangan & "' and TglDirujuk BETWEEN '" & Format(DTPickerAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' And statusPeriksa like '%" & combDiagnosa.Text & "%'" ' and  Diagnosa is Null  "

        rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
        Set dgData.DataSource = rs
        Call subsetgrid_1
        'jika kodisi pasien konsul sudah di diagnosa
    ElseIf (optPasienKonsul.Value = True And combDiagnosa.Enabled = True And combDiagnosa.Text = "SUDAH") Then
        Set rs = Nothing
        cetak = "PasienKonsulSudah"
        strSQL = "select TOP 100 TglDirujuk, NoPendaftaran, NoCM, [Nama Pasien], Umur, JK, Diagnosa, [Dokter Perujuk], [Ruangan Tujuan] " _
        & "from V_DaftarPasienKonsul " _
        & "WHERE ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%') and [Ruangan Tujuan]='" & strNNamaRuangan & "' and TglDirujuk BETWEEN '" & Format(DTPickerAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' And statusPeriksa like '%" & combDiagnosa.Text & "%'" ' and  Diagnosa is Not Null  "
'V_DaftarPasienKonsul
        rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
        Set dgData.DataSource = rs
        Call subsetgrid_1
        'jika combo diagnosa =""
    ElseIf (optPasienKonsul.Value = True And combDiagnosa.Enabled = True And combDiagnosa.Text = "") Then
        Set rs = Nothing
        cetak = "PasienKonsul1"
        strSQL = "select TOP 100 TglDirujuk, NoPendaftaran, NoCM, [Nama Pasien], Umur, JK, Diagnosa, [Dokter Perujuk], [Ruangan Tujuan] " _
        & "from V_DaftarPasienKonsul " _
        & "WHERE ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%') and [Ruangan Tujuan]='" & strNNamaRuangan & "' and TglDirujuk BETWEEN '" & Format(DTPickerAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd HH:mm:59") & "'"

        rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
        Set dgData.DataSource = rs
        Call subsetgrid_1
        'jika semua pasien konsul yang diinginkan
    ElseIf optPasienKonsul.Value = True And combDiagnosa.Enabled = False Then
        Set rs = Nothing
        cetak = "PasienKonsul2"
        strSQL = "select TOP 100 TglDirujuk, NoPendaftaran, NoCM, [Nama Pasien], Umur, JK, Diagnosa, [Dokter Perujuk], [Ruangan Tujuan] " _
        & "from V_DaftarPasienKonsul " _
        & "WHERE ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%') and [Ruangan Tujuan]='" & strNNamaRuangan & "' and TglDirujuk BETWEEN '" & Format(DTPickerAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' "

        rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
        Set dgData.DataSource = rs
        Call subsetgrid_1
    End If
    lblJumData.Caption = "Data 0/" & rs.RecordCount
errLoad:
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo sea
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    FrmCetakInformasiDiagnosaPasien.Show
sea:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub combDiagnosa_Change()
    Call cmdCari_Click
End Sub

Private Sub combDiagnosa_Click()
    Call cmdCari_Click
End Sub

Private Sub combDiagnosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCari.SetFocus
End Sub

Private Sub dgData_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgData
    WheelHook.WheelHook dgData
End Sub

Private Sub dgData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetak.SetFocus
End Sub

Private Sub dgData_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    lblJumData.Caption = "Data " & dgData.Bookmark & "/" & dgData.ApproxCount
End Sub

Private Sub DTPickerAkhir_Change()
    DTPickerAkhir.MaxDate = Now
End Sub

Private Sub DTPickerAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub DTPickerAwal_Change()
    DTPickerAwal.MaxDate = Now
End Sub

Private Sub DTPickerAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then DTPickerAkhir.SetFocus
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    DTPickerAwal.Value = Format(Now, "dd MMM yyyy 00:00:00")
    DTPickerAkhir.Value = Now
    optPasienPoli.Caption = "Pasien " + strNNamaRuangan
    optPasienPoli.Value = True
    If optPasienPoli.Value = True Then
        Set rs = Nothing
        rs.Open "select TOP 100 TglMasuk, NoPendaftaran, NoCM, [Nama Pasien], Umur, JK, Diagnosa, [Dokter Pemeriksa], Ruangan " _
        & "from V_DaftarPasienLamaRJ " _
        & "WHERE ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%') and Ruangan='" & strNNamaRuangan & "' and TglMasuk BETWEEN '" & Format(DTPickerAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' ", dbConn, adOpenStatic, adLockOptimistic
        Set dgData.DataSource = rs
        Call subSetGrid
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subsetgrid_1()
    With dgData
        .Columns(0).Caption = "Tgl. Dirujuk"
        .Columns(0).Width = 1900
        .Columns(1).Caption = "No.Registrasi"
        .Columns(1).Width = 1200
        .Columns(2).Caption = "No.CM"
        .Columns(2).Width = 950
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Caption = "Nama Pasien"
        .Columns(3).Width = 2500
        .Columns(4).Caption = "Umur"
        .Columns(4).Width = 1500
        .Columns(5).Caption = "JK"
        .Columns(5).Width = 400
        .Columns(5).Alignment = dbgCenter
        .Columns(6).Caption = "Diagnosa"
        .Columns(6).Width = 5000
        .Columns(7).Caption = "Dokter Perujuk"
        .Columns(7).Width = 2400
        .Columns(8).Width = 0
    End With
End Sub

Private Sub subSetGrid()
    With dgData
        .Columns(0).Caption = "Tgl. Masuk"
        .Columns(0).Width = 1900
        .Columns(1).Caption = "No.Registrasi"
        .Columns(1).Width = 1200
        .Columns(2).Caption = "No.CM"
        .Columns(2).Width = 950
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Caption = "Nama Pasien"
        .Columns(3).Width = 2500
        .Columns(4).Caption = "Umur"
        .Columns(4).Width = 1500
        .Columns(5).Caption = "JK"
        .Columns(5).Width = 400
        .Columns(5).Alignment = dbgCenter
        .Columns(6).Caption = "Diagnosa"
        .Columns(6).Width = 5000
        .Columns(7).Caption = "Dokter Pemeriksa"
        .Columns(7).Width = 2400
        .Columns(8).Width = 0
    End With
End Sub

Private Sub optPasienKonsul_Click()
    Call cmdCari_Click
End Sub

Private Sub optPasienKonsul_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkStatusPeriksa.SetFocus
End Sub

Private Sub optPasienPoli_Click()
    Call cmdCari_Click
End Sub

Private Sub optPasienPoli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkStatusPeriksa.SetFocus
End Sub

Private Sub txtParameter_Change()
    Call cmdCari_Click
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        Call cmdCari_Click
        txtParameter.SetFocus
    End If
End Sub

