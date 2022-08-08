VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBukuRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Buku Register Pasien"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14700
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBukuRegister.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   14700
   Begin VB.Frame frInstalasi 
      Caption         =   "&Instalasi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5520
      TabIndex        =   19
      Top             =   6600
      Visible         =   0   'False
      Width           =   5265
      Begin VB.CheckBox chkSemua 
         Caption         =   "&Cek Semua"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   7
         Top             =   20
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvInstalasi 
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   1720
         View            =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
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
      Height          =   5595
      Left            =   0
      TabIndex        =   13
      Top             =   960
      Width           =   14715
      Begin VB.OptionButton optPasienKonsul 
         Caption         =   "Pasien Konsul"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5520
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton optPasienPoli 
         Caption         =   "Pasien Poliklinik"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5295
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
         TabIndex        =   14
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
            Format          =   147128323
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
            Format          =   147128323
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   15
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgData 
         Height          =   4455
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   7858
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
         TabIndex        =   17
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
      Height          =   1455
      Left            =   0
      TabIndex        =   12
      Top             =   6600
      Width           =   14685
      Begin VB.CommandButton cmdCetak2 
         Caption         =   "C&etak Versi 2"
         Height          =   495
         Left            =   10920
         TabIndex        =   10
         Top             =   840
         Width           =   1665
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   450
         Width           =   2775
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "Ceta&k"
         Height          =   495
         Left            =   10920
         TabIndex        =   9
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   12720
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan  Nama Pasien / No.CM"
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   195
         Width           =   2640
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   18
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
      Picture         =   "FrmBukuRegister.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   12840
      Picture         =   "FrmBukuRegister.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "FrmBukuRegister.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12975
   End
End
Attribute VB_Name = "FrmBukuRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkSemua_Click()
    If chkSemua.Value = 1 Then
        LoadDataCombo "Kabeh", True
    Else
        LoadDataCombo "Kabeh"
    End If
End Sub

Private Sub LoadDataCombo(NCek As String, Optional BoolCek As Boolean)
    On Error GoTo errLoad

    Select Case NCek
        Case "Kabeh"
            lvInstalasi.ListItems.Clear
            dbrs.MoveFirst
            While Not dbrs.EOF
                lvInstalasi.ListItems.Add , "A" & dbrs(1).Value, dbrs(0).Value
                If BoolCek = True Then
                    lvInstalasi.ListItems("A" & dbrs(1)).Checked = True
                    lvInstalasi.ListItems("A" & dbrs(1)).ForeColor = vbBlue
                End If
                dbrs.MoveNext
            Wend
            lvInstalasi.Sorted = True
    End Select

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdCari_Click()
    On Error GoTo errLoad

    Call LoadDataGrid

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo errLoad
    Dim i As Integer

    If optPasienPoli.Value = True Then
        strSQL = "select TglMasuk,NoRegister,NoCm,NamaPasien,Alamat,Agama,Umur,JK," _
        & "AsalRujukan,JenisPasien,Diagnosa,Null as Keterangan " _
        & "from V_BukuRegisterPasienRJ where (NoCM like '%" & txtParameter.Text & "%' OR NamaPasien like '%" & txtParameter.Text & "%') AND TglMasuk BETWEEN " _
        & "'" & Format(DTPickerAwal.Value, "yyyy/MM/dd HH:mm:00") & "' AND " _
        & "'" & Format(DTPickerAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' AND " _
        & "KdRuangan='" & mstrKdRuangan & "'"

        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount = 0 Then
            MsgBox "Tidak ada data", vbInformation, "Informasi"
            cmdCetak.Enabled = True
            Exit Sub
        End If
        cetak = "BukuBesar"
        vLaporan = ""
        If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
        FrmViewerLaporan.Show
    Else
        xDaftarInstalasiA = ""
        xDaftarInstalasi = " and kdInstalasi in ("
        For i = 1 To lvInstalasi.ListItems.Count
            If lvInstalasi.ListItems.Item(i).Checked = True Then
                xDaftarInstalasi = xDaftarInstalasi & Right(lvInstalasi.ListItems(i).Key, Len(lvInstalasi.ListItems(i).Key) - 1) & ","
                xDaftarInstalasiA = xDaftarInstalasiA & lvInstalasi.ListItems(i).Text & ","
            End If
        Next i

        If xDaftarInstalasi = " and kdInstalasi in (" Then
            xDaftarInstalasi = ""
        Else
            xDaftarInstalasi = Left(xDaftarInstalasi, Len(xDaftarInstalasi) - 1) & ")"
        End If

        If xDaftarInstalasiA <> "" Then xDaftarInstalasiA = Left(xDaftarInstalasiA, Len(xDaftarInstalasiA) - 1)

        strSQL = "select distinct Tgldirujuk,NoPendaftaran,NoCM,NamaPasien,Alamat,Umur,JK," _
        & "RuangPerujuk,Diagnosa,JenisPasien,Null as Keterangan,KdRuanganTujuan " _
        & "from V_DaftarPasienKonsul_N where (NamaPasien like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%')and " _
        & "KdRuanganTujuan='" & strNKdRuangan & "' AND TglDirujuk BETWEEN " _
        & "'" & Format(DTPickerAwal.Value, "yyyy/MM/dd HH:mm:00") & "' AND " _
        & "'" & Format(DTPickerAkhir.Value, "yyyy/MM/dd HH:mm:00") & "' " & xDaftarInstalasi & ""
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount = 0 Then
            MsgBox "Tidak ada data", vbInformation, "Informasi"
            cmdCetak.Enabled = True
            Exit Sub
        End If

        cetak = "PasienKonsul"
        vLaporan = ""
        If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
        FrmViewerLaporan.Show
    End If
    Exit Sub
errLoad:
End Sub

Private Sub cmdCetak2_Click()
    On Error GoTo errLoad
    Dim i As Integer

    Call filter_kriteria
    If optPasienPoli.Value = True Then
        strSQL = "select dbo.FB_TakeBlnThn(TglMasuk) as TglMasuk,NoRegister,NoCm,NamaPasien,Alamat,Agama,Umur,JK," _
        & "AsalRujukan,JenisPasien,Diagnosa,Null as Keterangan " _
        & "from V_BukuRegisterPasienRJ where (NoCM like '%" & txtParameter.Text & "%' OR NamaPasien like '%" & txtParameter.Text & "%') AND TglMasuk BETWEEN " _
        & "'" & Format(DTPickerAwal.Value, "yyyy/MM/dd HH:mm:00") & "' AND " _
        & "'" & Format(DTPickerAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' AND " _
        & "KdRuangan='" & mstrKdRuangan & "'"

        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount = 0 Then
            MsgBox "Tidak ada data", vbInformation, "Informasi"
            cmdCetak.Enabled = True
            Exit Sub
        End If
        cetak = "BukuBesar"
        vLaporan = ""
        If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
        FrmViewerLaporan2.Show
    Else
        xDaftarInstalasiA = ""
        xDaftarInstalasi = " and kdInstalasi in ("
        For i = 1 To lvInstalasi.ListItems.Count
            If lvInstalasi.ListItems.Item(i).Checked = True Then
                xDaftarInstalasi = xDaftarInstalasi & Right(lvInstalasi.ListItems(i).Key, Len(lvInstalasi.ListItems(i).Key) - 1) & ","
                xDaftarInstalasiA = xDaftarInstalasiA & lvInstalasi.ListItems(i).Text & ","
            End If
        Next i

        If xDaftarInstalasi = " and kdInstalasi in (" Then
            xDaftarInstalasi = ""
        Else
            xDaftarInstalasi = Left(xDaftarInstalasi, Len(xDaftarInstalasi) - 1) & ")"
        End If

        If xDaftarInstalasiA <> "" Then xDaftarInstalasiA = Left(xDaftarInstalasiA, Len(xDaftarInstalasiA) - 1)

        strSQL = "select dbo.FB_TakeBlnThn(TglDirujuk) as TglDirujuk,NoPendaftaran,NoCM,NamaPasien,Alamat,Umur,JK," _
        & "RuangPerujuk,Diagnosa,JenisPasien,Null as Keterangan,KdRuanganTujuan " _
        & "from V_DaftarPasienKonsul_N where (NamaPasien like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%')and " _
        & "KdRuanganTujuan='" & strNKdRuangan & "' AND TglDirujuk BETWEEN " _
        & "'" & Format(DTPickerAwal.Value, "yyyy/MM/dd HH:mm:00") & "' AND " _
        & "'" & Format(DTPickerAkhir.Value, "yyyy/MM/dd HH:mm:00") & "' " & xDaftarInstalasi & "" ' AND " _

        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount = 0 Then
            MsgBox "Tidak ada data", vbInformation, "Informasi"
            cmdCetak.Enabled = True
            Exit Sub
        End If
        cetak = "PasienKonsul"
        vLaporan = ""
        If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
        FrmViewerLaporan2.Show
    End If
    Exit Sub
errLoad:
End Sub

Sub filter_kriteria()
    mdTglAwal = DTPickerAwal.Value
    mdTglAkhir = DTPickerAkhir.Value
    Dim mdtBulan As Integer
    Dim MdtTahun As Integer

    mdTglAwal = CDate(Format(DTPickerAwal.Value, "yyyy-mm ") & "-01 00:00:00") 'TglAwal
    mdtBulan = CStr(Format(DTPickerAkhir.Value, "mm"))
    MdtTahun = CStr(Format(DTPickerAkhir.Value, "yyyy"))
    mdTglAkhir = CDate(Format(DTPickerAkhir.Value, "yyyy-mm") & "-" & funcHitungHari(mdtBulan, MdtTahun) & " 23:59:59")
End Sub

Private Sub cmdTutup_Click()
    Unload Me
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
    On Error Resume Next
    DTPickerAkhir.MaxDate = Now
End Sub

Private Sub DTPickerAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub DTPickerAwal_Change()
    On Error Resume Next
    DTPickerAwal.MaxDate = Now
End Sub

Private Sub DTPickerAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then DTPickerAkhir.SetFocus
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    With Me
        .DTPickerAwal.Value = Format(Now, "dd MMM yyyy 00:00:00")
        .DTPickerAkhir.Value = Now
    End With
    optPasienPoli.Caption = "Pasien " + strNNamaRuangan
    optPasienPoli.Value = True

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subSetGrid()
    On Error Resume Next
    With dgData
        .Columns(0).Caption = "Tgl. Masuk"
        .Columns(0).Width = 1590
        .Columns(1).Caption = "No.Registrasi"
        .Columns(1).Width = 1200
        .Columns(2).Caption = "No.CM"
        .Columns(2).Width = 800
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Caption = "Nama Pasien"
        .Columns(3).Width = 2500
        .Columns(4).Caption = "Alamat"
        .Columns(4).Width = 2500
        .Columns(5).Caption = "Agama"
        .Columns(5).Width = 1500
        .Columns(6).Caption = "Umur"
        .Columns(6).Width = 1500
        .Columns(7).Caption = "JK"
        .Columns(7).Width = 400
        .Columns(7).Alignment = dbgCenter
        .Columns(8).Caption = "Status"
        .Columns(8).Width = 600
        .Columns(8).Alignment = dbgCenter
        .Columns(9).Caption = "Asal Rujukan"
        .Columns(9).Width = 1500
        .Columns(10).Caption = "Jenis Pasien"
        .Columns(10).Width = 1500
        .Columns(10).Alignment = dbgCenter
        .Columns(11).Caption = "Keterangan"
        .Columns(11).Width = 1500
    End With
End Sub

Private Sub subSetGridPasienKonsul()
    On Error Resume Next
    With dgData
        .Columns(0).Caption = "Tgl. DiRujuk"
        .Columns(0).Width = 1590
        .Columns(1).Caption = "No.Registrasi"
        .Columns(1).Width = 1200
        .Columns(2).Caption = "No.CM"
        .Columns(2).Width = 800
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Caption = "Nama Pasien"
        .Columns(3).Width = 2500
        .Columns(4).Caption = "Alamat"
        .Columns(4).Width = 2500
        .Columns(5).Caption = "Umur"
        .Columns(5).Width = 1500
        .Columns(6).Caption = "JK"
        .Columns(6).Width = 400
        .Columns(6).Alignment = dbgCenter
        .Columns(7).Caption = "Asal Rujukan"
        .Columns(7).Width = 1500
        .Columns(8).Caption = "Jenis Pasien"
        .Columns(8).Width = 1500
        .Columns(8).Alignment = dbgCenter
        .Columns(9).Caption = "Keterangan"
        .Columns(9).Width = 1500
        .Columns(10).Width = 0
    End With
End Sub

Private Sub optPasienKonsul_Click()
    Call cmdCari_Click
    frInstalasi.Enabled = True
    frInstalasi.Visible = True
    cmdCetak2.Visible = True
    cmdCetak2.Enabled = True
    Call msubRecFO(dbrs, "Select NamaInstalasi, kdInstalasi from Instalasi where kdInstalasi in ('01', '02', '03', '04', '06', '09', '10', '11', '12', '13', '16', '19', '22', '24')")
    lvInstalasi.ListItems.Clear
    While Not dbrs.EOF
        lvInstalasi.ListItems.Add , "A" & dbrs(1).Value, dbrs(0).Value
        lvInstalasi.ListItems("A" & dbrs(1)).Checked = False
        lvInstalasi.ListItems("A" & dbrs(1)).ForeColor = vbBlack
        dbrs.MoveNext
    Wend
End Sub

Private Sub optPasienKonsul_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCari.SetFocus
End Sub

Private Sub optPasienPoli_Click()
    Call cmdCari_Click
    frInstalasi.Enabled = False
    frInstalasi.Visible = False
    chkSemua.Value = 0
    cmdCetak2.Visible = False
    lvInstalasi.ListItems.Clear
End Sub

Private Sub LoadDataGrid()
    On Error GoTo errLoad
    Dim i As Integer
    lblJumData.Caption = "Data 0/0"
    If optPasienPoli.Value = True Then
        Set rs = Nothing
        rs.Open "select TglMasuk,NoRegister,NoCM,NamaPasien,Alamat,Agama,Umur,JK," _
        & "StatusPasien,AsalRujukan,JenisPasien,Null as Keterangan " _
        & "from V_BukuRegisterPasienRJ_new where (NoCM like '%" & txtParameter.Text & "%' OR NamaPasien like '%" & txtParameter.Text & "%') AND TglMasuk BETWEEN " _
        & "'" & Format(DTPickerAwal.Value, "yyyy/MM/dd HH:mm:00") & "' AND " _
        & "'" & Format(DTPickerAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' AND KdRuangan='" & mstrKdRuangan & "'", dbConn, adOpenStatic, adLockOptimistic
        Set dgData.DataSource = rs
        Call subSetGrid
    Else
        xDaftarInstalasi = " and kdInstalasi in ("
        For i = 1 To lvInstalasi.ListItems.Count
            If lvInstalasi.ListItems.Item(i).Checked = True Then
                xDaftarInstalasi = xDaftarInstalasi & Right(lvInstalasi.ListItems(i).Key, Len(lvInstalasi.ListItems(i).Key) - 1) & ","
            End If
        Next i

        If xDaftarInstalasi = " and kdInstalasi in (" Then
            xDaftarInstalasi = ""
        Else
            xDaftarInstalasi = Left(xDaftarInstalasi, Len(xDaftarInstalasi) - 1) & ")"
        End If

        Set rs = Nothing
        strSQL = "select Tgldirujuk,NoPendaftaran,NoCM,[Nama Pasien],Alamat,Umur,JK," _
        & "[Ruangan Perujuk],JenisPasien,Null as Keterangan,KdRuanganTujuan " _
        & "from V_DaftarPasienKonsul where ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%')and KdRuanganTujuan='" & strNKdRuangan & "' AND TglDirujuk BETWEEN " _
        & "'" & Format(DTPickerAwal.Value, "yyyy/MM/dd HH:mm:00") & "' AND " _
        & "'" & Format(DTPickerAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' " & xDaftarInstalasi & ""

        rs.Open "select Tgldirujuk,NoPendaftaran,NoCM,[Nama Pasien],Alamat,Umur,JK," _
        & "[Ruangan Perujuk],JenisPasien,Null as Keterangan,KdRuanganTujuan " _
        & "from V_DaftarPasienKonsul where ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%')and KdRuanganTujuan='" & strNKdRuangan & "' AND TglDirujuk BETWEEN " _
        & "'" & Format(DTPickerAwal.Value, "yyyy/MM/dd HH:mm:00") & "' AND " _
        & "'" & Format(DTPickerAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' " & xDaftarInstalasi & "", dbConn, adOpenStatic, adLockOptimistic

        Set dgData.DataSource = rs
        Call subSetGridPasienKonsul
    End If
    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub optPasienPoli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCari.SetFocus
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        Call LoadDataGrid
        txtParameter.SetFocus
    End If
End Sub

