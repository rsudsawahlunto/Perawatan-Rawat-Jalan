VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDaftarPasienRJRIIGD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Daftar Pasien"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPasienRJRIIGD.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   14910
   Begin VB.Frame fraCari 
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
      Height          =   960
      Left            =   0
      TabIndex        =   8
      Top             =   7320
      Width           =   14895
      Begin VB.CommandButton cmdKonsul 
         Caption         =   "&Konsul ke Unit Lain"
         Height          =   495
         Left            =   8760
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdTP 
         Caption         =   "&Transaksi Pelayanan"
         Height          =   495
         Left            =   10800
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   12840
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   480
         TabIndex        =   4
         Top             =   440
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Masukan Nama Pasien /  No.CM / Ruangan"
         Height          =   210
         Left            =   480
         TabIndex        =   11
         Top             =   195
         Width           =   3450
      End
   End
   Begin VB.Frame fraDaftar 
      Caption         =   "Daftar Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   14895
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
         Height          =   735
         Left            =   9000
         TabIndex        =   10
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
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   0
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   252248067
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   1
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   252248067
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   12
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgDaftarPasienRJ 
         Height          =   5295
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   9340
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   2
         RowHeight       =   16
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
      Begin VB.Label LblJumData 
         AutoSize        =   -1  'True
         Caption         =   "10 / 100 Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1155
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   14
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   7860
      Width           =   14910
      _ExtentX        =   26300
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Visible         =   0   'False
            Object.Width           =   26239
            MinWidth        =   2294
            Text            =   "Cetak Daftar Pasien (F9)"
            TextSave        =   "Cetak Daftar Pasien (F9)"
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
   Begin VB.Image Image2 
      Height          =   945
      Left            =   13080
      Picture         =   "frmDaftarPasienRJRIIGD.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarPasienRJRIIGD.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarPasienRJRIIGD.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13215
   End
End
Attribute VB_Name = "frmDaftarPasienRJRIIGD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dTglMasuk As Date

Public Sub cmdCari_Click()
    On Error GoTo errLoad
    dgDaftarPasienRJ.SetFocus
    lblJumData.Caption = ""
    If dtpAwal.Day <> dtpAkhir.Day Or dtpAwal.Month <> dtpAkhir.Month Or dtpAwal.Year <> dtpAkhir.Year Then
        strSQL = "SELECT TOP 100 RuanganPerawatan, NoPendaftaran, NoCM, NamaPasien, JK, Umur, JenisPasien, Kelas, TglMasuk, TglKeluar, StatusKeluar, KondisiPulang, KasusPenyakit, NoKamar, NoBed, Alamat, DokterPemeriksa, UmurTahun, UmurBulan, UmurHari, KdKelas, KdJenisTarif, KdSubInstalasi, KdRuangan, IdDokter" & _
        " FROM V_DaftarInfoPasienAll " & _
        " WHERE (NamaPasien like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%' OR RuanganPerawatan like '%" & txtParameter.Text & "%' OR NoKamar like '%" & txtParameter.Text & "%') and ((TglMasuk between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "') OR TglKeluar is NULL)"
    Else
        strSQL = "SELECT Top 100 RuanganPerawatan, NoPendaftaran, NoCM, NamaPasien, JK, Umur, JenisPasien, Kelas, TglMasuk, TglKeluar, StatusKeluar, KondisiPulang, KasusPenyakit, NoKamar, NoBed, Alamat, DokterPemeriksa, UmurTahun, UmurBulan, UmurHari, KdKelas, KdJenisTarif, KdSubInstalasi, KdRuangan, IdDokter" & _
        " FROM V_DaftarInfoPasienAll " & _
        " WHERE (NamaPasien like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%' OR RuanganPerawatan like '%" & txtParameter.Text & "%' OR NoKamar like '%" & txtParameter.Text & "%') and ((TglMasuk between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "') OR TglKeluar is NULL)"
    End If
    Call msubRecFO(rs, strSQL)
    Set dgDaftarPasienRJ.DataSource = rs
    Call SetGridPasienRJ
    lblJumData.Caption = "Data 0/" & rs.RecordCount
    Exit Sub
errLoad:
End Sub

Private Sub cmdKonsul_Click()
    Call subLoadFormKonsul
End Sub

Private Sub cmdTP_Click()
    On Error GoTo hell
    Call subLoadFormTP
hell:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgDaftarPasienRJ_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDaftarPasienRJ
    WheelHook.WheelHook dgDaftarPasienRJ
End Sub

Private Sub dgDaftarPasienRJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdKonsul.SetFocus
End Sub

Private Sub dgDaftarPasienRJ_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    lblJumData.Caption = dgDaftarPasienRJ.Bookmark & " / " & dgDaftarPasienRJ.ApproxCount & " Data"
End Sub

Private Sub dtpAkhir_Change()
    On Error Resume Next
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpAkhir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpAwal_Change()
    On Error Resume Next
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub dtpAwal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF9
            lblJumData.Caption = ""
            Set rs = Nothing
            If dtpAwal.Day <> dtpAkhir.Day Or dtpAwal.Month <> dtpAkhir.Month Or dtpAwal.Year <> dtpAkhir.Year Then
                strSQL = "SELECT RuanganPerawatan, NoPendaftaran, NoCM, NamaPasien, JK, Umur, JenisPasien, Kelas, TglMasuk, TglKeluar, StatusKeluar, KondisiPulang, KasusPenyakit, NoKamar, NoBed, Alamat, DokterPemeriksa, UmurTahun, UmurBulan, UmurHari, KdKelas, KdJenisTarif, KdSubInstalasi, KdRuangan, IdDokter" & _
                " FROM V_DaftarInfoPasienAll " & _
                " WHERE (NamaPasien like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%' OR RuanganPerawatan like '%" & txtParameter.Text & "%' OR NoKamar like '%" & txtParameter.Text & "%') and ((TglMasuk between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "') OR TglKeluar is NULL)"
            Else
                strSQL = "SELECT RuanganPerawatan, NoPendaftaran, NoCM, NamaPasien, JK, Umur, JenisPasien, Kelas, TglMasuk, TglKeluar, StatusKeluar, KondisiPulang, KasusPenyakit, NoKamar, NoBed, Alamat, DokterPemeriksa, UmurTahun, UmurBulan, UmurHari, KdKelas, KdJenisTarif, KdSubInstalasi, KdRuangan, IdDokter" & _
                " FROM V_DaftarInfoPasienAll " & _
                " WHERE (NamaPasien like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%' OR RuanganPerawatan like '%" & txtParameter.Text & "%' OR NoKamar like '%" & txtParameter.Text & "%') and ((TglMasuk between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "') OR TglKeluar is NULL)"
            End If
            Call msubRecFO(rs, strSQL)
            frmCtkDaftarPasienRS.Show
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    dtpAwal.Value = Format(Now, "dd MMM yyyy 00:00:00")
    dtpAkhir.Value = Now

    Call cmdCari_Click
    mblnForm = True
End Sub

Sub SetGridPasienRJ()
    Dim i As Integer

    With dgDaftarPasienRJ
        For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next i
        .Columns("RuanganPerawatan").Width = 1700
        .Columns("NoPendaftaran").Width = 1200
        .Columns("NoPendaftaran").Caption = "No. Registrasi"
        .Columns("NoCM").Width = 800
        .Columns("NamaPasien").Width = 1800
        .Columns("JK").Width = 300
        .Columns("Umur").Width = 1500

        .Columns("JenisPasien").Width = 1100
        .Columns("Kelas").Width = 1200
        .Columns("TglMasuk").Width = 1590
        .Columns("TglKeluar").Width = 1590
        .Columns("StatusKeluar").Width = 2250
        .Columns("KondisiPulang").Width = 1800
        .Columns("KasusPenyakit").Width = 2100
        .Columns("NoKamar").Width = 1000
        .Columns("NoBed").Width = 700
        .Columns("Alamat").Width = 6000
        .Columns("DokterPemeriksa").Width = 2000
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnForm = False
End Sub

Private Sub txtParameter_Change()
    On Error GoTo errLoad
    Call cmdCari_Click
    txtParameter.SetFocus
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        Call cmdCari_Click
        txtParameter.SetFocus
    End If
End Sub

Private Sub subLoadFormTP()
    On Error GoTo hell
    mstrNoPen = dgDaftarPasienRJ.Columns("No. Registrasi").Value
    mstrNoCM = dgDaftarPasienRJ.Columns("NoCM").Value

    mstrKdRuanganPasien = dgDaftarPasienRJ.Columns("KdRuangan").Value 'Kode Ruangan Pasien
    mstrNamaRuanganPasien = dgDaftarPasienRJ.Columns("RuanganPerawatan").Value 'Nama Ruangan Pasien

    With frmTransaksiPasien
        .Show
        .txtNoPendaftaran.Text = dgDaftarPasienRJ.Columns("No. Registrasi").Value
        .txtNoCM.Text = dgDaftarPasienRJ.Columns("NoCM").Value
        .txtNamaPasien.Text = dgDaftarPasienRJ.Columns("NamaPasien").Value
        If dgDaftarPasienRJ.Columns("JK").Value = "L" Then
            .txtSex.Text = "Laki-Laki"
        Else
            .txtSex.Text = "Perempuan"
        End If
        .txtThn.Text = dgDaftarPasienRJ.Columns("UmurTahun").Value
        .txtBln.Text = dgDaftarPasienRJ.Columns("UmurBulan").Value
        .txtHr.Text = dgDaftarPasienRJ.Columns("UmurHari").Value
        .txtKls.Text = dgDaftarPasienRJ.Columns("Kelas").Value
        .txtJenisPasien.Text = dgDaftarPasienRJ.Columns("JenisPasien").Value
        .txtTglDaftar.Text = dgDaftarPasienRJ.Columns("TglMasuk").Value
    End With

    mdTglMasuk = dgDaftarPasienRJ.Columns("TglMasuk").Value
    mstrKdKelas = dgDaftarPasienRJ.Columns("KdKelas").Value
    mstrKelas = dgDaftarPasienRJ.Columns("Kelas").Value
    mstrKdSubInstalasi = dgDaftarPasienRJ.Columns("KdSubInstalasi").Value
    mstrKdDokter = dgDaftarPasienRJ.Columns("IdDokter").Value
    mstrNamaDokter = dgDaftarPasienRJ.Columns("DokterPemeriksa").Value

    strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdJenisPasien = rs("KdKelompokPasien").Value
        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
    End If
hell:
End Sub

Private Sub subLoadFormKonsul()
    On Error GoTo hell
    mstrNoPen = dgDaftarPasienRJ.Columns("No. Registrasi").Value
    mstrNoCM = dgDaftarPasienRJ.Columns("NoCM").Value

    mstrKdRuanganPasien = dgDaftarPasienRJ.Columns("KdRuangan").Value 'Kode Ruangan Pasien
    mstrNamaRuanganPasien = dgDaftarPasienRJ.Columns("RuanganPerawatan").Value 'Nama Ruangan Pasien
    mstrKdSubInstalasi = dgDaftarPasienRJ.Columns("KdSubInstalasi").Value
    mstrKdDokter = IIf(IsNull(dgDaftarPasienRJ.Columns("IdDokter")), Null, dgDaftarPasienRJ.Columns("IdDokter"))
    mstrNamaDokter = IIf(IsNull(dgDaftarPasienRJ.Columns("DokterPemeriksa")), Null, dgDaftarPasienRJ.Columns("DokterPemeriksa"))
    mstrFormPengirim = Me.Name

    With frmPasienRujukan
        .Show
        .txtNoPendaftaran.Text = dgDaftarPasienRJ.Columns("No. Registrasi").Value
        .txtNoCM.Text = dgDaftarPasienRJ.Columns("NoCM").Value
        .txtNamaPasien.Text = dgDaftarPasienRJ.Columns("NamaPasien").Value
        If dgDaftarPasienRJ.Columns("JK").Value = "L" Then
            .txtSex.Text = "Laki-Laki"
        Else
            .txtSex.Text = "Perempuan"
        End If
        .txtThn.Text = dgDaftarPasienRJ.Columns("UmurTahun").Value
        .txtBln.Text = dgDaftarPasienRJ.Columns("UmurBulan").Value
        .txtHari.Text = dgDaftarPasienRJ.Columns("UmurHari").Value
    End With

    mdTglMasuk = dgDaftarPasienRJ.Columns("TglMasuk").Value
    mstrKdKelas = dgDaftarPasienRJ.Columns("KdKelas").Value
    mstrKelas = dgDaftarPasienRJ.Columns("Kelas").Value

    strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdJenisPasien = rs("KdKelompokPasien").Value
        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
    End If
    Exit Sub
hell:
End Sub

