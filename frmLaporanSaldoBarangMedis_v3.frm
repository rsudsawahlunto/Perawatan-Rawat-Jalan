VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLaporanSaldoBarangMedis_v3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Laporan Saldo Barang"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLaporanSaldoBarangMedis_v3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   12795
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   10
      Top             =   7560
      Width           =   12735
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   495
         Left            =   9720
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   11280
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin MSComctlLib.ProgressBar pbData 
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   873
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Min             =   1e-4
         Scrolling       =   1
      End
      Begin VB.Label lblPersen 
         Caption         =   "0 %"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   7200
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
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
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   12735
      Begin VB.Frame Frame4 
         Caption         =   "Group By"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   5775
         Begin VB.OptionButton optStatusBrg 
            Caption         =   "Status Barang"
            Height          =   495
            Left            =   3600
            TabIndex        =   26
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optPabrik 
            Caption         =   "Pabrik"
            Height          =   495
            Left            =   4800
            TabIndex        =   25
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optGolBarang 
            Caption         =   "Golongan Barang"
            Height          =   495
            Left            =   2400
            TabIndex        =   24
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optKategoriBrg 
            Caption         =   "Kategory Barang"
            Height          =   495
            Left            =   1200
            TabIndex        =   23
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optJnsBarang 
            Caption         =   "Jenis Barang"
            Height          =   495
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkGroupBy 
         Caption         =   "Jenis Barang"
         Height          =   255
         Left            =   6000
         TabIndex        =   19
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CheckBox chkAsalBarang 
         Caption         =   "Asal Barang"
         Height          =   255
         Left            =   9480
         TabIndex        =   18
         Top             =   225
         Width           =   1935
      End
      Begin VB.CommandButton cmdProses 
         Caption         =   "&Proses"
         Height          =   495
         Left            =   9840
         TabIndex        =   16
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   495
         Left            =   10920
         TabIndex        =   14
         Top             =   1200
         Width           =   975
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4335
         Begin VB.OptionButton optHari 
            Caption         =   "Per Hari"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.OptionButton optBulan 
            Caption         =   "Per Bulan"
            Height          =   375
            Left            =   1440
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optTahun 
            Caption         =   "Per Tahun"
            Height          =   375
            Left            =   2880
            TabIndex        =   3
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   330
         Left            =   4680
         TabIndex        =   6
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
         Format          =   115933187
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   330
         Left            =   7200
         TabIndex        =   7
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
         Format          =   115933187
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSDataListLib.DataCombo dcAsalBarang 
         Height          =   330
         Left            =   9480
         TabIndex        =   17
         Top             =   480
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo dcGroupBy 
         Height          =   330
         Left            =   6000
         TabIndex        =   20
         Top             =   1335
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "s/d"
         Height          =   210
         Left            =   6840
         TabIndex        =   8
         Top             =   525
         Width           =   255
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
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
   Begin MSFlexGridLib.MSFlexGrid fgData 
      Height          =   4335
      Left            =   0
      TabIndex        =   9
      Top             =   3120
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   7646
      _Version        =   393216
      AllowUserResizing=   1
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
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmLaporanSaldoBarangMedis_v3.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   9120
      Picture         =   "frmLaporanSaldoBarangMedis_v3.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmLaporanSaldoBarangMedis_v3.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10935
   End
End
Attribute VB_Name = "frmLaporanSaldoBarangMedis_v3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim sKriteria As String

Private Sub chkAsalBarang_Click()
    If chkAsalBarang.Value = vbChecked Then
        dcAsalBarang.Enabled = True
        Call msubDcSource(dcAsalBarang, rs, "Select KdAsal,NamaAsal From AsalBarang Where KdInstalasi = '07'")
        dcAsalBarang.BoundText = rs(0)
        dcAsalBarang.Text = ""
    Else
        dcAsalBarang.Enabled = False
        dcAsalBarang.BoundColumn = ""
        dcAsalBarang.ListField = ""
        dcAsalBarang.BoundText = ""
        dcAsalBarang.Text = ""
    End If
End Sub

Private Sub chkGroupBy_Click()
    If chkGroupBy.Value = vbChecked Then
        dcGroupBy.Enabled = True
        Select Case chkGroupBy.Caption
            Case "Jenis Barang"
                Call msubDcSource(dcGroupBy, rs, "Select KdJenisBarang,JenisBarang From JenisBarang Where KdKelompokBarang = '02'")
                dcGroupBy.BoundText = rs(0)
                dcGroupBy.Text = ""
            Case "Kategory Barang"
                Call msubDcSource(dcGroupBy, rs, "Select KdKategoryBarang,KategoryBarang From KategoryBarang")
                dcGroupBy.BoundText = rs(0)
                dcGroupBy.Text = ""
            Case "Golongan Barang"
                Call msubDcSource(dcGroupBy, rs, "Select KdGolBarang,GolonganBarang From GolonganBarang")
                dcGroupBy.BoundText = rs(0)
                dcGroupBy.Text = ""
            Case "Status Barang"
                Call msubDcSource(dcGroupBy, rs, "Select KdStatusBarang,StatusBarang From StatusBarang")
                dcGroupBy.BoundText = rs(0)
                dcGroupBy.Text = ""
            Case "Pabrik"
                Call msubDcSource(dcGroupBy, rs, "Select KdPabrik,NamaPabrik From Pabrik")
                dcGroupBy.BoundText = rs(0)
                dcGroupBy.Text = ""
        End Select
    Else
        dcGroupBy.Enabled = False
        dcGroupBy.Text = ""
        dcGroupBy.ListField = ""
        dcGroupBy.BoundColumn = ""
    End If
End Sub

Private Sub cmdBatal_Click()
    MousePointer = vbDefault
    optHari.Value = False
    optBulan.Value = True
    Call optBulan_Click
    optTahun.Value = False
    dtpAwal.Value = Now
    dtpAkhir.Value = Now
    dcAsalBarang.Text = ""
    dcGroupBy.Text = ""
    chkAsalBarang.Value = Unchecked
    chkGroupBy.Value = Unchecked
    cmdProses.Enabled = True
    cmdCetak.Enabled = False
    optJnsBarang.Value = False
    optKategoriBrg.Value = False
    optGolBarang.Value = False
    optPabrik.Value = False
    optStatusBrg.Value = False
    pbData.Value = 0.0001
    lblPersen.Caption = "0%"
    Call setgrid
End Sub

Private Sub cmdCetak_Click()
    Dim m As Integer
    On Error GoTo hell
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    With fgData
        For m = 1 To .Rows - 2
            If sp_cetakLaporanSaldobarang(.TextMatrix(m, 1), .TextMatrix(m, 2), .TextMatrix(m, 3), .TextMatrix(m, 4), .TextMatrix(m, 5), _
                .TextMatrix(m, 6), .TextMatrix(m, 7), .TextMatrix(m, 8), .TextMatrix(m, 9), .TextMatrix(m, 10), _
                .TextMatrix(m, 11)) = False Then Exit Sub
            Next m
        End With
        strSQL = ""
        strSQL = "select * From LaporanSaldoBarangMedis_T where KdRuangan Like '%" & mstrKdRuangan & "%'"
        On Error GoTo errLoad
        frmCetakLapSaldoBarangFIFO.Show
        Exit Sub
hell:
errLoad:
End Sub

Private Sub cmdProses_Click()
    On Error GoTo hell
    Dim cekbFInsert As Boolean
    Dim cekbFInsert1 As Boolean
    Dim i, j, k, l, m, n, o, p, q As Integer
    Dim dtTgl As Date
    Dim intBln As Integer
    Dim intTgl As Integer
    Dim intThn As Integer
    Dim intTglLast As Integer
    Dim intTglNow As Integer
    Dim intHr As Integer
    Dim sFltrAsalBrg As String
    Dim sFltrGroupBy As String
    Dim sFltrSaldoBy As String
    Dim sQuery As String
    Dim sbQuery As String
    Dim sJmlKirim As Long
    Dim sJmlTotalpengiriman As Long
    Dim sJmlKirimRuangan As Long
    Dim sJmlKeluarPemakaianAlat As Long
    Dim sJmlPasienRetur As Long
    Dim sJmlKeluarBarangReturRuangan As Long
    Dim sJmlTotalPengirimanSetelahOpname As Long
    Dim sJmlTambahdariSupplier As Long

    Dim sJmlTambahdariReturSupplier As Long
    Dim sJmlTotalJumlahTerima As Long
    Dim sJmlPengirimanSetelahStokOpname As Long
    
    Dim sJmlKeluarPasien As Long
    
    
    Dim statusOpnameBaru As Long

    Dim sJmlKirimRuanganSetelahOpname As Long
    Dim sJmlKeluarPemakaianAlatSetelahOpname As Long
    Dim sJmlKeluarBarangReturRuanganSetelahOpname As Long

    Dim sJmlKeluarTotal As Long
    Dim sJmlTambahdariReturSupplierSetelahOpname As Long
    Dim sJmlKeluarTotalSetelahOpname As Long
    Dim sJmlTotalRetur As Long
    Dim waktu As Date
    Dim zz As Integer

    MousePointer = vbHourglass
    Call setgrid
    cmdProses.Enabled = False
    sJmlKirim = 0
    sJmlTotalpengiriman = 0
    sJmlKirimRuangan = 0
    sJmlKeluarPemakaianAlat = 0
    sJmlPasienRetur = 0
    sJmlTotalRetur = 0
    sJmlPengirimanSetelahStokOpname = 0
    sJmlKeluarBarangReturRuangan = 0
    sJmlTambahdariReturSupplier = 0
    sJmlTambahdariReturSupplierSetelahOpname = 0
    sJmlKeluarBarangReturRuanganSetelahOpname = 0
    sJmlTotalPengirimanSetelahOpname = 0
    sJmlTambahdariSupplier = 0
  
    
    sFltrAsalBrg = ""
    sFltrAsalBrg = ""
    If chkAsalBarang.Value = Checked Then
        sFltrAsalBrg = " AND KdAsal = '" & dcAsalBarang.BoundText & "'"
    Else
        sFltrAsalBrg = ""
    End If

    If optJnsBarang.Value = False And optKategoriBrg.Value = False And optGolBarang = False And optStatusBrg.Value = False And optPabrik.Value = False Then
        optJnsBarang.Value = True
    End If

    If chkGroupBy.Value = Checked Then
        If optJnsBarang.Value = True Then
            sFltrGroupBy = " AND KdJenisBarang Like '%" & dcGroupBy.BoundText & "%'"
            sFltrSaldoBy = "JenisBarang"
            sKriteria = "JenisBarang"
        ElseIf optKategoriBrg.Value = True Then
            sFltrGroupBy = " AND KdKategoryBarang Like '%" & dcGroupBy.BoundText & "%'"
            sFltrSaldoBy = "KategoryBarang"
            sKriteria = "KategoryBarang"
        ElseIf optGolBarang.Value = True Then
            sFltrGroupBy = " AND KdGolBarang Like '%" & dcGroupBy.BoundText & "%'"
            sFltrSaldoBy = "GolonganBarang"
            sKriteria = "GolonganBarang"
        ElseIf optStatusBrg.Value = True Then
            sFltrGroupBy = " AND KdStatusBarang Like '%" & dcGroupBy.BoundText & "%'"
            sFltrSaldoBy = "StatusBarang"
            sKriteria = "StatusBarang"
        ElseIf optPabrik.Value = True Then
            sFltrGroupBy = " AND KdPabrik Like '%" & dcGroupBy.BoundText & "%'"
            sFltrSaldoBy = "NamaPabrik"
            sKriteria = "NamaPabrik"
        End If
        fgData.TextMatrix(0, 10) = sFltrSaldoBy
    Else
        sFltrGroupBy = ""
        sFltrSaldoBy = ""
    End If

    fgData.TextMatrix(0, 10) = sFltrSaldoBy

    If optHari.Value = True Then
        dtTgl = dtpAwal.Value
    ElseIf optBulan.Value = True Then
        dtTgl = Format(dtpAwal.Value, "yyyy/MM/01 HH:mm:ss")
    ElseIf optTahun.Value = True Then
        dtTgl = Format(dtpAwal.Value, "yyyy/01/01 HH:mm:ss")
    End If

    intHr = CInt(Day(dtTgl)) - 1
    If intHr = 0 Then
        intBln = CInt(Month(dtTgl)) - 1
        If intBln = 0 Then
            intBln = 12
            intThn = CInt(Year(dtTgl)) - 1
            intTglLast = funcHitungHari(intBln, CInt(Year(dtTgl)) - 1)
        Else
            intThn = CInt(Year(Now))
            intTglLast = funcHitungHari(CInt(Month(dtTgl)) - 1, Year(dtTgl))
            intTglNow = funcHitungHari(CInt(Month(dtTgl)), Year(dtTgl))
        End If
    Else
        intTgl = CInt(Day(dtTgl))
        intBln = CInt(Month(dtTgl))
        intThn = CInt(Year(dtTgl))
    End If
    If intHr = 0 Then
        intTgl = intTglLast
    End If

    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value

    tgl = ""
    tgl = funcHitungHari(Month(mdTglAkhir), Year(mdTglAkhir))

    Set rs = Nothing
    strSQL = " Select *   From V_StokBarangMedisRuanganLengkap Where KdRuangan='" & mstrKdRuangan & "'" & sFltrAsalBrg & sFltrGroupBy
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
        MsgBox "Ruangan belum ada stoknya atau Master data barang belum lengkap", vbExclamation, "Validasi"
        MousePointer = vbDefault
        cmdProses.Enabled = True
        Exit Sub
    End If
    rs.MoveFirst
    k = 1
    For i = 1 To rs.RecordCount
        pbData.Max = rs.RecordCount
        DoEvents
        lblPersen.Caption = Int((i / rs.RecordCount) * 100) & "%"
           sJmlTotalJumlahTerima = 0
           sJmlKeluarPasien = 0

        With fgData
'            strSQL = ""
'
'            If optHari.Value = True Then
'                strSQL = "SELECT KdBarang, KdAsal,NamaBarang,NamaAsal, CASE (SELECT SIGN(SUM(JmlStok)) FROM LaporanSaldoBarangMedisApotikNRuangan_V WHERE TglKirim BETWEEN '" & Format(intThn & "/" & intBln & "/" & intTgl, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(intThn & "/" & intBln & "/" & intTgl, "yyyy/MM/dd 23:59:59") & "' AND KdBarang = '" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "') " & _
'                "WHEN -1 THEN 0 WHEN 0 THEN 0 WHEN 1 THEN (SELECT SUM(JmlStok) FROM LaporanSaldoBarangMedisApotikNRuangan_V WHERE TglKirim BETWEEN '" & Format(intThn & "/" & intBln & "/" & intTgl, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(intThn & "/" & intBln & "/" & intTgl, "yyyy/MM/dd 23:59:59") & "' AND KdBarang = '" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "') ELSE 0 END AS SaldoAwal, SUM(JmlKirim) AS JmlTerima, SUM(JmlKeluar) AS JmlKeluar, HargaJual, JenisBarang, KategoryBarang, GolonganBarang, StatusBarang, NamaPabrik,KdRuanganTujuan " & _
'                "FROM LaporanSaldoBarangMedisApotikNRuangan_V " & _
'                "WHERE (TglKirim BETWEEN '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') And KdBarang = '" & rs("KdBarang") & "' And KdAsal='" & rs("KdAsal") & "' AND KdRuanganTujuan Like '%" & mstrKdRuangan & "%'" & _
'                "GROUP BY KdBarang, KdAsal, HargaJual, NamaBarang, NamaAsal, JenisBarang, KategoryBarang, GolonganBarang, StatusBarang, NamaPabrik,KdRuanganTujuan"
'
'            ElseIf optBulan.Value = True Then
'                'View ini hanya mencakup penerimaan dan pengiriman antar ruangan saja (Untuk Supplier nya tidak mencakup) (bagian penerimaan 1)
'                strSQL = "SELECT KdBarang, KdAsal,NamaBarang,NamaAsal, CASE (SELECT SIGN(SUM(JmlStok)) FROM LaporanSaldoBarangMedisApotikNRuangan_V WHERE TglKirim BETWEEN '" & Format(intThn & "/" & intBln & "/01", "yyyy/MM/dd 00:00:00") & "' AND '" & Format(intThn & "/" & intBln & "/" & intTgl, "yyyy/MM/dd 23:59:59") & "' AND KdBarang = '" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "' and KdRuanganTujuan = '" & mstrKdRuangan & "') " & _
'                "WHEN -1 THEN 0 WHEN 0 THEN 0 WHEN 1 THEN (SELECT SUM(JmlStok) FROM LaporanSaldoBarangMedisApotikNRuangan_V WHERE TglKirim BETWEEN '" & Format(intThn & "/" & intBln & "/01", "yyyy/MM/dd 00:00:00") & "' AND '" & Format(intThn & "/" & intBln & "/" & intTgl, "yyyy/MM/dd 23:59:59") & "' AND KdBarang = '" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "') ELSE 0 END AS SaldoAwal, SUM(JmlKirim) AS JmlTerima, SUM(JmlKeluar) AS JmlKeluar, HargaJual, JenisBarang, KategoryBarang, GolonganBarang, StatusBarang, NamaPabrik,KdRuanganTujuan " & _
'                "FROM LaporanSaldoBarangMedisApotikNRuangan_V " & _
'                "WHERE (TglKirim BETWEEN '" & Format(mdTglAwal, "yyyy/MM/01 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "') And KdBarang = '" & rs("KdBarang") & "' And KdAsal='" & rs("KdAsal") & "' AND KdRuanganTujuan Like '%" & mstrKdRuangan & "%'" & _
'                "GROUP BY KdBarang, KdAsal, HargaJual, NamaBarang, NamaAsal, JenisBarang, KategoryBarang, GolonganBarang, StatusBarang, NamaPabrik,KdRuanganTujuan"
'
'                'PENERIMAAN DARI SUPPLIER DAN RUANGAN LAIN (BAGIAN PENERIMAAN KE 1 DAN 2)
''                 strSQL = "SELECT KdBarang, KdAsal,NamaBarang,NamaAsal, CASE (SELECT SIGN(SUM(JmlStok)) FROM V_LaporanSaldoBarangMedisFromSupplierdanRuangan WHERE TglKirim BETWEEN '" & Format(intThn & "/" & intBln & "/01", "yyyy/MM/dd 00:00:00") & "' AND '" & Format(intThn & "/" & intBln & "/" & intTgl, "yyyy/MM/dd 23:59:59") & "' AND KdBarang = '" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "' and KdRuanganTujuan = '" & mstrKdRuangan & "') " & _
'                "WHEN -1 THEN 0 WHEN 0 THEN 0 WHEN 1 THEN (SELECT SUM(JmlStok) FROM V_LaporanSaldoBarangMedisFromSupplierdanRuangan WHERE TglKirim BETWEEN '" & Format(intThn & "/" & intBln & "/01", "yyyy/MM/dd 00:00:00") & "' AND '" & Format(intThn & "/" & intBln & "/" & intTgl, "yyyy/MM/dd 23:59:59") & "' AND KdBarang = '" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "') ELSE 0 END AS SaldoAwal, SUM(JmlKirim) AS JmlTerima, SUM(JmlKeluar) AS JmlKeluar, HargaJual, JenisBarang, KategoryBarang, GolonganBarang, StatusBarang, NamaPabrik,KdRuanganTujuan " & _
'                "FROM V_LaporanSaldoBarangMedisFromSupplierdanRuangan " & _
'                "WHERE (TglKirim BETWEEN '" & Format(mdTglAwal, "yyyy/MM/01 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "') And KdBarang = '" & rs("KdBarang") & "' And KdAsal='" & rs("KdAsal") & "' AND KdRuanganTujuan Like '%" & mstrKdRuangan & "%'" & _
'                "GROUP BY KdBarang, KdAsal, HargaJual, NamaBarang, NamaAsal, JenisBarang, KategoryBarang, GolonganBarang, StatusBarang, NamaPabrik,KdRuanganTujuan"
'
'
'            ElseIf optTahun.Value = True Then
'
'                strSQL = "SELECT KdBarang, KdAsal,NamaBarang,NamaAsal, CASE (SELECT SIGN(SUM(JmlStok)) FROM LaporanSaldoBarangMedisApotikNRuangan_V WHERE TglKirim BETWEEN '" & Format(mdTglAwal, "yyyy/01/01 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/12/" & tgl & " 23:59:59") & "' AND KdBarang = '" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "') " & _
'                "WHEN -1 THEN 0 WHEN 0 THEN 0 WHEN 1 THEN (SELECT SUM(JmlStok) FROM LaporanSaldoBarangMedisApotikNRuangan_V WHERE TglKirim BETWEEN '" & Format(mdTglAwal, "yyyy/01/01 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/12/" & tgl & " 23:59:59") & "' AND KdBarang = '" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "') ELSE 0 END AS SaldoAwal, SUM(JmlKirim) AS JmlTerima, SUM(JmlKeluar) AS JmlKeluar, HargaJual, JenisBarang, KategoryBarang, GolonganBarang, StatusBarang, NamaPabrik,KdRuanganTujuan " & _
'                "FROM LaporanSaldoBarangMedisApotikNRuangan_V " & _
'                "WHERE (TglKirim BETWEEN '" & Format(mdTglAwal, "yyyy/01/01 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/12/" & tgl & " 23:59:59") & "') And KdBarang = '" & rs("KdBarang") & "' And KdAsal='" & rs("KdAsal") & "' AND KdRuanganTujuan Like '%" & mstrKdRuangan & "%'" & _
'                "GROUP BY KdBarang, KdAsal, HargaJual, NamaBarang, NamaAsal, JenisBarang, KategoryBarang, GolonganBarang, StatusBarang, NamaPabrik,KdRuanganTujuan"
'            End If

            'Set dbRst = Nothing
            'Call msubRecFO(rs, strSQL)

           If rs.RecordCount <> 0 Then
              
                'PENGECEKAN SUDAH STOK OPNAME ATAU BELUM
                'EDIT SIGIT 2013
                'PILIH TOP(1) DATASTOCKOPNAME
                strSQL2 = "select TOP 1 * from DataStokBarangMedisStokOpname where KdRuangan = '" & mstrKdRuangan & "' and  KdBarang = '" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "'  and Tglinputstok is not null  order by TglInputStok desc"
                Call msubRecFO(rsx, strSQL2)
                'JIKA DATA STOCK OPNAME YES(Clear)
                If rsx.RecordCount <> 0 Then
                           'PILIH TOP 1 Laporan Saldo Barang
                            strSQL9 = "Select top  1 * from LapSaldoBarangMedis " & _
                            " WHERE KdRuanganTujuan =  '" & mstrKdRuangan & "' AND KdBarang = '" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "' order by TanggalStok desc "
                            Call msubRecFO(rsI, strSQL9)
                
                            If rsI.RecordCount <> 0 Then
                            
                                'tanggal DATASTOCKOPNAME > tanggal LAPSALDOBARANG
                                If rsx("TglInputStok") > rsI("TanggalStok") Then
                                cekbFInsert = True
                               'Cek dulu apa ada transaksi Sebelum dari tanggal 1 sampai hari ini checkALL
                                GoTo checkALL
balik:
                                   sJmlTotalJumlahTerima = 0
                                   sJmlKeluarPasien = 0
                                   waktu = DateAdd("s", 1, Now)
                                   
                                    sbQuery = "insert into LapSaldoBarangMedis values('" & rs("KdBarang") & "','" & rs("KdAsal") & "','" & rs("NamaBarang") & "','" & rs("NamaAsal") & "','" & rs("JmlStok") & "', 0, " & _
                                     "'" & sJmlKeluarPasien & "', null, '" & rs("JenisBarang") & "', null, null, null, null, '" & mstrKdRuangan & "', '" & Format(waktu, "yyyy/MM/dd hh:mm:ss ") & "','','','T')"
                                     Call msubRecFO(rsB, sbQuery)
                                    cekbFInsert = False
                                   
                               
                                End If
                            Else
                            cekbFInsert1 = True
                            'Cek dulu apa ada transaksi Sebelum dari tanggal 1 sampai hari ini checkALL
                            GoTo checkALL
balik1:
                            'set penerimaan dan pengeluaran =0 karena stok opname
                             sJmlTotalJumlahTerima = 0
                             sJmlKeluarPasien = 0
                             'Insert Laporan Saldo Barang Berdasarkan data stock opname
                             
                             waktu = DateAdd("s", 1, Now)
                             sbQuery = "insert into LapSaldoBarangMedis values('" & rs("KdBarang") & "','" & rs("KdAsal") & "','" & rs("NamaBarang") & "','" & rs("NamaAsal") & "','" & rsx("JmlStokReal") & "', '" & sJmlTotalJumlahTerima & "', " & _
                                      "'" & sJmlKeluarPasien & "', null, '" & rs("JenisBarang") & "', null, null, null, null, '" & mstrKdRuangan & "', '" & Format(waktu, "yyyy/MM/dd hh:mm:ss") & "','','','T')"
                            'sbQuery = "insert into LapSaldoBarangMedis values('" & dbRst("KdBarang") & "','" & dbRst("KdAsal") & "','" & dbRst("NamaBarang") & "','" & dbRst("NamaAsal") & "','" & rsx("JmlStokReal") & "', '" & sJmlTotalJumlahTerima & "', " & _
                            '          "'" & sJmlKeluarPasien & "', '" & dbRst("HargaJual") & "', '" & dbRst("JenisBarang") & "', '" & dbRst("KategoryBarang") & "', '" & dbRst("GolonganBarang") & "', '" & dbRst("StatusBarang") & "', '" & dbRst("NamaPabrik") & "', '" & mstrKdRuangan & "', '" & Format(Now, "yyyy/MM/dd hh:mm:ss") & "','','','Y')"
                            cekbFInsert1 = False
                            Call msubRecFO(rsB, sbQuery)
                            End If
                        
                        
               'TIDAK ADA DATASTOCKOPNAME (Clear)
               Else
               
                 'Nothing
               End If
            End If
               'Ketika Sudah ada LaporanSaldoBarang
               '*************************************************************
checkALL:
              '
              'PILIH TOP 1 Laporan Saldo Barang
             ' If dbRst.EOF = False Then
                strSQL9 = "Select top  1 * from LapSaldoBarangMedis " & _
                " WHERE KdRuanganTujuan =  '" & mstrKdRuangan & "' AND KdBarang = '" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "' order by TanggalStok desc "
                Call msubRecFO(rsI, strSQL9)
                If rsI.EOF = True Then
                         '# PENGELUARAN # Jika Belum ada LaporanSaldoBarangMedis
                        'Cek Pemakaian alkes
                        strSQL10 = "Select SUM(JmlBarang) as JmlPakaiAlkes from PemakaianAlkes where KdRuangan = '" & mstrKdRuangan & "' and  KdBarang = '" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "' and TglPelayanan BETWEEN '" & Format(mdTglAwal, "yyyy/MM/01 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "'"
                        Call msubRecFO(rsJ, strSQL10)
                        If rsJ("JmlPakaiAlkes") <> Null Or rsJ("JmlPakaiAlkes") <> "" Then
                         sJmlKeluarPasien = sJmlKeluarPasien + rsJ("JmlPakaiAlkes")
                        End If
                        'Pengecekan pemakaian obat pada di pemakaian bahan dan alat
                        strSQL11 = "Select SUM(JmlBarang) as JmlPakaiOA  from PemakaianBahanAlat where KdRuangan = '" & mstrKdRuangan & "' and  KdBarang = '" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "' and TglPemakaian BETWEEN '" & Format(mdTglAwal, "yyyy/MM/01 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "'"
                        Call msubRecFO(rsK, strSQL11)
                        If rsK("JmlPakaiOA") <> Null Or rsK("JmlPakaiOA") <> "" Then
                         sJmlKeluarPasien = sJmlKeluarPasien + rsK("JmlPakaiOA")
                        End If
                        'Pengecekan Apkah ada Retur keruangan lain
                         sTRSQL13 = "SELECT   SUM(DetailReturStrukBarangKeluar.JmlRetur) as JmlRetur " & _
                         "   FROM         Retur INNER JOIN " & _
                         "   DetailReturStrukBarangKeluar ON Retur.NoRetur = DetailReturStrukBarangKeluar.NoRetur INNER JOIN " & _
                         "   StrukKirim ON DetailReturStrukBarangKeluar.NoKirim = StrukKirim.NoKirim " & _
                         "   where StrukKirim.KdRuanganTujuan = '" & mstrKdRuangan & "' and KdBarang = '" & rs("KdBarang") & "' and KdAsal = '" & rs("KdAsal") & "' and TglRetur Between '" & Format(mdTglAwal, "yyyy/MM/01 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "'"
                         Call msubRecFO(rsM, sTRSQL13)
                         If rsM("JmlRetur") <> Null Or rsM("JmlRetur") <> "" Then
                         sJmlKeluarPasien = sJmlKeluarPasien + rsM("JmlRetur")
                        End If
                        'Pengecekan Penjualan Pasien Bebas
                        sTRSQL13 = "SELECT    SUM( ApotikJual.JmlBarang) as JmlBarang " & _
                        " FROM         StrukPelayananPasien INNER JOIN " & _
                        " ApotikJual ON StrukPelayananPasien.NoStruk = ApotikJual.NoStruk " & _
                        " Where ApotikJual.KdBarang='" & rs("KdBarang") & "' and ApotikJual.KdAsal = '" & rs("KdAsal") & "'  And StrukPelayananPasien.TglStruk Between'" & Format(mdTglAwal, "yyyy/MM/01 00:00:00") & "'  AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "'"
                        Call msubRecFO(rsC, sTRSQL13)
                        If rsC("JmlBarang") <> Null Or rsC("JmlBarang") <> "" Then
                         sJmlKeluarPasien = sJmlKeluarPasien + rsC("JmlBarang")
                        End If
                        
                        'pENERIMAAN DARI RUANGAN LAIN
                          strSQL7 = "SELECT SUM(DetailBarangKeluar.JmlKirim)as jmlKirim " & _
                             " FROM StrukKirim INNER JOIN" & _
                             " DetailBarangKeluar ON StrukKirim.NoKirim = DetailBarangKeluar.NoKirim" & _
                             " WHERE StrukKirim.KdRuangan =  '" & mstrKdRuangan & "' AND DetailBarangKeluar.KdBarang = '" & rs("KdBarang") & "' AND DetailBarangKeluar.KdAsal='" & rs("KdAsal") & "'and TglKirim BETWEEN '" & Format(mdTglAwal, "yyyy/MM/01 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "' "
                             Call msubRecFO(rsG, strSQL7)
                            If rsG("jmlKirim") <> Null Or rsG("jmlKirim") <> "" Then
                            
                            sJmlKeluarPasien = sJmlKeluarPasien + rsG("jmlKirim")
                            End If
                            
                             'Retur Ke Supplier
                          strSQL7 = "SELECT     SUM(DetailReturStrukTerimaBarang.JmlRetur) as JmlRetur " & _
                         " FROM         DetailReturStrukTerimaBarang INNER JOIN " & _
                         " Retur ON DetailReturStrukTerimaBarang.NoRetur = Retur.NoRetur LEFT OUTER JOIN  " & _
                         " StrukTerima ON DetailReturStrukTerimaBarang.NoTerima = StrukTerima.NoTerima " & _
                         "  WHERE     StrukTerima.KdRuangan ='" & mstrKdRuangan & "' AND DetailReturStrukTerimaBarang.KdBarang = '" & rs("KdBarang") & "' AND DetailReturStrukTerimaBarang.KdAsal='" & rs("KdAsal") & "' and Retur.TglRetur BETWEEN '" & Format(mdTglAwal, "yyyy/MM/01 00:00:00") & "'  AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "'  "
                             Call msubRecFO(rsG, strSQL7)
                            If rsG("JmlRetur") <> Null Or rsG("JmlRetur") <> "" Then
                            
                            sJmlKeluarPasien = sJmlKeluarPasien + rsG("JmlRetur")
                            End If
                            
                        ''# END PENGELUARAN # Jika Belum ada LaporanSaldoBarangMedis
                        
                        '-----------------------------------------------------------------------
                        
                        '# PENERIMAAN # Jika Belum ada LaporanSaldoBarangMedis
                        
                        'pENERIMAAN DARI RUANGAN LAIN
                          strSQL7 = "SELECT SUM(DetailBarangKeluar.JmlKirim)as jmlKirim " & _
                             " FROM StrukKirim INNER JOIN" & _
                             " DetailBarangKeluar ON StrukKirim.NoKirim = DetailBarangKeluar.NoKirim" & _
                             " WHERE StrukKirim.KdRuanganTujuan =  '" & mstrKdRuangan & "' AND DetailBarangKeluar.KdBarang = '" & rs("KdBarang") & "' AND DetailBarangKeluar.KdAsal='" & rs("KdAsal") & "'and TglKirim BETWEEN '" & Format(mdTglAwal, "yyyy/MM/01 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "' "
                             Call msubRecFO(rsG, strSQL7)
                            If rsG("jmlKirim") <> Null Or rsG("jmlKirim") <> "" Then
                            
                            sJmlTotalJumlahTerima = sJmlTotalJumlahTerima + rsG("jmlKirim")
                            End If
                             
                        'Pengecekan Retur Obat dari Pasien, sehingga akan menambah ke Jumlah Terima Obat edit sigit
                            StrSQL12 = "Select  SUM(JmlRetur) as JmlRetur from v_returpelayananoOAPasienKasir where KdRuangan = '" & mstrKdRuangan & "' and  KdBarang = '" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "' and TglPelayanan BETWEEN '" & Format(mdTglAwal, "yyyy/MM/01 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "'"
                            Call msubRecFO(rsL, StrSQL12)
                             If rsL("JmlRetur") <> Null Or rsL("JmlRetur") <> "" Then
                            
                            sJmlTotalJumlahTerima = sJmlTotalJumlahTerima + rsL("JmlRetur")
                            End If
                            
                         'Retur Pasien Bebas
                            StrSQL12 = "SELECT     Sum(JmlBarangRetur) as JmlRetur " & _
                            " FROM  Retur INNER JOIN DetailReturStrukApotik ON Retur.NoRetur = DetailReturStrukApotik.NoRetur " & _
                            " WHERE  DetailReturStrukApotik.KdRuangan = '" & mstrKdRuangan & "' and  KdBarang = '" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "' and TglRetur BETWEEN '" & Format(mdTglAwal, "yyyy/MM/01 00:00:00") & "'  AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "' "
                            Call msubRecFO(rsL, StrSQL12)
                            If rsL("JmlRetur") <> Null Or rsL("JmlRetur") <> "" Then
                            
                            sJmlTotalJumlahTerima = sJmlTotalJumlahTerima + rsL("JmlRetur")
                            End If
                            
                             'Pengecekan Apkah ada Retur Dari Ruangan Lain
                             sTRSQL13 = "SELECT     SUM(JmlRetur) as JmlRetur " & _
                                " FROM         DetailReturStrukBarangKeluar INNER JOIN" & _
                                " Retur ON DetailReturStrukBarangKeluar.NoRetur = Retur.NoRetur INNER JOIN " & _
                                " StrukKirim ON DetailReturStrukBarangKeluar.NoKirim = StrukKirim.NoKirim" & _
                                " where StrukKirim.KdRuangan= '" & mstrKdRuangan & "' and DetailReturStrukBarangKeluar.KdBarang = '" & rs("KdBarang") & "' and KdAsal = '" & rs("KdAsal") & "' and TglRetur BETWEEN '" & Format(mdTglAwal, "yyyy/MM/01 00:00:00") & "'  AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "' "
                             Call msubRecFO(rsM, sTRSQL13)
                             If rsM("JmlRetur") <> Null Or rsM("JmlRetur") <> "" Then
                             sJmlTotalJumlahTerima = sJmlTotalJumlahTerima + rsM("JmlRetur")
                            End If
                            
                            'Penerimaan Dari Supplier
                            sTRSQL13 = "select SUM(JmlTerima) as JmlTerima  from DetailTerimaBarang " & _
                            "  DetailTerimaBarang INNER JOIN " & _
                            "  StrukTerima ON DetailTerimaBarang.NoTerima = StrukTerima.NoTerima " & _
                            " where  KdRuangan= '" & mstrKdRuangan & "' and KdBarang = '" & rs("KdBarang") & "' and KdAsal = '" & rs("KdAsal") & "' and StrukTerima.TglFaktur BETWEEN '" & Format(mdTglAwal, "yyyy/MM/01 00:00:00") & "'  AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "' "
                             Call msubRecFO(rsM, sTRSQL13)
                             If rsM("JmlTerima") <> Null Or rsM("JmlTerima") <> "" Then
                             sJmlTotalJumlahTerima = sJmlTotalJumlahTerima + rsM("JmlTerima")
                            End If
                            
                            
                         ''# END PENERIMAAN # Jika Belum ada LaporanSaldoBarangMedis
                         
                         ' Tambahkan Ke LaporanSaldoBarang
                          sbQuery = "insert into LapSaldoBarangMedis values('" & rs("KdBarang") & "','" & rs("KdAsal") & "','" & rs("NamaBarang") & "','" & rs("NamaAsal") & "','0', '" & sJmlTotalJumlahTerima & "', " & _
                                     "'" & sJmlKeluarPasien & "', null, '" & rs("JenisBarang") & "', null, null, null, null, '" & mstrKdRuangan & "', '" & Format(Now, "yyyy/MM/01 00:00:00") & "','','','Y')"
                          'Kosongin Lagi
                        
                          
                          Call msubRecFO(rsB, sbQuery)
                        If cekbFInsert = True Then GoTo balik
                        If cekbFInsert1 = True Then GoTo balik1
                 
               Else
               'Jika Laporan Saldo Barang SUDAH aDa (clear)
               '# PENGELUARAN #
                        
                        'Cek Pemakaian alkes
                        strSQL10 = "Select SUM(JmlBarang) as JmlPakaiAlkes from PemakaianAlkes where KdRuangan = '" & mstrKdRuangan & "' and  KdBarang = '" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "' and TglPelayanan BETWEEN  '" & Format(rsI("TanggalStok"), "yyyy/MM/dd hh:MM:ss") & "'   AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "'"
                        Call msubRecFO(rsJ, strSQL10)
                        If rsJ("JmlPakaiAlkes") <> Null Or rsJ("JmlPakaiAlkes") <> "" Then
                         sJmlKeluarPasien = sJmlKeluarPasien + rsJ("JmlPakaiAlkes")
                        End If
                        'Pengecekan pemakaian obat pada di pemakaian bahan dan alat
                        strSQL11 = "Select SUM(JmlBarang) as JmlPakaiOA  from PemakaianBahanAlat where KdRuangan = '" & mstrKdRuangan & "' and  KdBarang = '" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "' and TglPemakaian BETWEEN '" & Format(rsI("TanggalStok"), "yyyy/MM/dd hh:MM:ss") & "'   AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "'"
                        Call msubRecFO(rsK, strSQL11)
                         If rsK("JmlPakaiOA") <> Null Or rsK("JmlPakaiOA") <> "" Then
                         sJmlKeluarPasien = sJmlKeluarPasien + rsK("JmlPakaiOA")
                        End If
                        'Pengecekan Apkah ada Retur keruangan lain
                         sTRSQL13 = "SELECT     SUM(JmlRetur) as JmlRetur " & _
                            " FROM         DetailReturStrukBarangKeluar INNER JOIN" & _
                            " Retur ON DetailReturStrukBarangKeluar.NoRetur = Retur.NoRetur INNER JOIN " & _
                            " StrukKirim ON DetailReturStrukBarangKeluar.NoKirim = StrukKirim.NoKirim" & _
                            " where StrukKirim.KdRuanganTujuan = '" & mstrKdRuangan & "' and DetailReturStrukBarangKeluar.KdBarang = '" & rs("KdBarang") & "' and KdAsal = '" & rs("KdAsal") & "' and TglRetur Between'" & Format(rsI("TanggalStok"), "yyyy/MM/DD HH:MM:SS") & "'  AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "'"
                         Call msubRecFO(rsM, sTRSQL13)
                         If rsM("JmlRetur") <> Null Or rsM("JmlRetur") <> "" Then
                         sJmlKeluarPasien = sJmlKeluarPasien + rsM("JmlRetur")
                        End If
                        
                        
                         'Cek Penjualan Pasien Bebas
                         strSQL9 = "SELECT    SUM( ApotikJual.JmlBarang) as JmlBarang " & _
                        " FROM         StrukPelayananPasien INNER JOIN " & _
                        " ApotikJual ON StrukPelayananPasien.NoStruk = ApotikJual.NoStruk " & _
                        " Where ApotikJual.KdBarang='" & rs("KdBarang") & "' and ApotikJual.KdAsal = '" & rs("KdAsal") & "'  And StrukPelayananPasien.TglStruk Between'" & Format(rsI("TanggalStok"), "yyyy/MM/DD HH:MM:SS") & "'  AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "'"
                        Call msubRecFO(rsC, strSQL9)
                          If rsC("JmlBarang") <> Null Or rsC("JmlBarang") <> "" Then
                         sJmlKeluarPasien = sJmlKeluarPasien + rsC("JmlBarang")
                        End If
                        
                           strSQL7 = "SELECT SUM(DetailBarangKeluar.JmlKirim)as jmlKirim " & _
                             " FROM StrukKirim INNER JOIN" & _
                             " DetailBarangKeluar ON StrukKirim.NoKirim = DetailBarangKeluar.NoKirim" & _
                             " WHERE StrukKirim.KdRuangan =  '" & mstrKdRuangan & "' AND DetailBarangKeluar.KdBarang = '" & rs("KdBarang") & "' AND DetailBarangKeluar.KdAsal='" & rs("KdAsal") & "' and StrukKirim.TglKirim BETWEEN '" & Format(rsI("TanggalStok"), "yyyy/MM/dd hh:MM:ss") & "'  AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "'  "
                             Call msubRecFO(rsG, strSQL7)
                            If rsG("jmlKirim") <> Null Or rsG("jmlKirim") <> "" Then
                            
                            sJmlKeluarPasien = sJmlKeluarPasien + rsG("jmlKirim")
                            End If
                       'Retur Ke Supplier
                          strSQL7 = "SELECT     SUM(DetailReturStrukTerimaBarang.JmlRetur) as JmlRetur " & _
                         " FROM         DetailReturStrukTerimaBarang INNER JOIN " & _
                         " Retur ON DetailReturStrukTerimaBarang.NoRetur = Retur.NoRetur LEFT OUTER JOIN  " & _
                         " StrukTerima ON DetailReturStrukTerimaBarang.NoTerima = StrukTerima.NoTerima " & _
                         "  WHERE     StrukTerima.KdRuangan ='" & mstrKdRuangan & "' AND DetailReturStrukTerimaBarang.KdBarang = '" & rs("KdBarang") & "' AND DetailReturStrukTerimaBarang.KdAsal='" & rs("KdAsal") & "' and Retur.TglRetur BETWEEN '" & Format(rsI("TanggalStok"), "yyyy/MM/dd hh:MM:ss") & "'  AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "'  "
                             Call msubRecFO(rsG, strSQL7)
                            If rsG("JmlRetur") <> Null Or rsG("JmlRetur") <> "" Then
                            
                            sJmlKeluarPasien = sJmlKeluarPasien + rsG("JmlRetur")
                            End If

                      
                      
                        ''# END PENGELUARAN #
                        
                        '-----------------------------------------------------------------------
                        
                        '# PENERIMAAN #
                             ' Untuk Ruangan Pelayanan Dan Apotik ***********************************************
                            'Penerimaan Dari Apotik Atau Farmasi
                            strSQL7 = "SELECT SUM(DetailBarangKeluar.JmlKirim)as jmlKirim " & _
                             " FROM StrukKirim INNER JOIN" & _
                             " DetailBarangKeluar ON StrukKirim.NoKirim = DetailBarangKeluar.NoKirim" & _
                             " WHERE StrukKirim.KdRuanganTujuan =  '" & mstrKdRuangan & "' AND DetailBarangKeluar.KdBarang = '" & rs("KdBarang") & "' AND DetailBarangKeluar.KdAsal='" & rs("KdAsal") & "' and StrukKirim.TglKirim BETWEEN '" & Format(rsI("TanggalStok"), "yyyy/MM/dd hh:MM:ss") & "'  AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "'  "
                             Call msubRecFO(rsG, strSQL7)
                            If rsG("jmlKirim") <> Null Or rsG("jmlKirim") <> "" Then
                            
                            sJmlTotalJumlahTerima = sJmlTotalJumlahTerima + rsG("jmlKirim")
                            End If
                            ' End Untuk  Apotik ***********************************************
                            
                          'Pengecekan Retur Obat dari Pasien, sehingga akan menambah ke Jumlah Terima Obat edit sigit
                            StrSQL12 = "Select  SUM(JmlRetur) as JmlRetur from v_returpelayananoOAPasienKasir where KdRuangan = '" & mstrKdRuangan & "' and  KdBarang = '" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "' and TglPelayanan BETWEEN '" & Format(rsI("TanggalStok"), "yyyy/MM/dd hh:MM:ss") & "'  AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "' "
                            Call msubRecFO(rsL, StrSQL12)
                            If rsL("JmlRetur") <> Null Or rsL("JmlRetur") <> "" Then
                            
                            sJmlTotalJumlahTerima = sJmlTotalJumlahTerima + rsL("JmlRetur")
                            End If
                            ' End Untuk Ruangan Pelayanan"***********************************************
                            ' UNTUK  APOTIK"***********************************************
                            'Retur Pasien Bebas
                            StrSQL12 = "SELECT     Sum(JmlBarangRetur) as JmlRetur " & _
                            " FROM  Retur INNER JOIN DetailReturStrukApotik ON Retur.NoRetur = DetailReturStrukApotik.NoRetur " & _
                            " WHERE  DetailReturStrukApotik.KdRuangan = '" & mstrKdRuangan & "' and  KdBarang = '" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "' and TglRetur BETWEEN '" & Format(rsI("TanggalStok"), "yyyy/MM/dd hh:MM:ss") & "'  AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "' "
                            Call msubRecFO(rsL, StrSQL12)
                            If rsL("JmlRetur") <> Null Or rsL("JmlRetur") <> "" Then
                            sJmlTotalJumlahTerima = sJmlTotalJumlahTerima + rsL("JmlRetur")
                            End If
                            ' End Untuk APOTIK"***********************************************
                            ' UNTUK FARMASI DAN APOTIK"***********************************************
                             'Pengecekan Apkah ada Retur Dari Ruangan Lain
                             sTRSQL13 = "SELECT     SUM(JmlRetur) as JmlRetur " & _
                                " FROM         DetailReturStrukBarangKeluar INNER JOIN" & _
                                " Retur ON DetailReturStrukBarangKeluar.NoRetur = Retur.NoRetur INNER JOIN " & _
                                " StrukKirim ON DetailReturStrukBarangKeluar.NoKirim = StrukKirim.NoKirim" & _
                                " where StrukKirim.KdRuangan= '" & mstrKdRuangan & "' and DetailReturStrukBarangKeluar.KdBarang = '" & rs("KdBarang") & "' and KdAsal = '" & rs("KdAsal") & "' and TglRetur Between'" & Format(rsI("TanggalStok"), "yyyy/MM/DD HH:MM:SS") & "'  AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "'"
                             Call msubRecFO(rsM, sTRSQL13)
                             If rsM("JmlRetur") <> Null Or rsM("JmlRetur") <> "" Then
                             sJmlTotalJumlahTerima = sJmlTotalJumlahTerima + rsM("JmlRetur")
                            End If
               
                            'Penerimaan Dari Supplier
                            sTRSQL13 = "select SUM(JmlTerima) as JmlTerima  from DetailTerimaBarang " & _
                            "  DetailTerimaBarang INNER JOIN " & _
                            "  StrukTerima ON DetailTerimaBarang.NoTerima = StrukTerima.NoTerima " & _
                            " where  KdRuangan= '" & mstrKdRuangan & "' and KdBarang = '" & rs("KdBarang") & "' and KdAsal = '" & rs("KdAsal") & "' and StrukTerima.TglFaktur Between'" & Format(rsI("TanggalStok"), "yyyy/MM/DD HH:MM:SS") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "' "
                             Call msubRecFO(rsM, sTRSQL13)
                             If rsM("JmlTerima") <> Null Or rsM("JmlTerima") <> "" Then
                             sJmlTotalJumlahTerima = sJmlTotalJumlahTerima + rsM("JmlTerima")
                            End If
                            ' END UNTUK FARMASI DAN APOTIK"***********************************************
                         ''# END PENERIMAAN #
                         
                         'UPDATE LAPORAN SALDO BARANG
                            sbQuery = "Update LapSaldoBarangMedis set JmlTerima  = '" & sJmlTotalJumlahTerima & "', " & _
                                     " JmlKeluar = '" & sJmlKeluarPasien & "', StatusSaldoAwal = 'T' where KdRuanganTujuan = '" & mstrKdRuangan & "' and  KdBarang = '" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "' AND TanggalStok='" & Format(rsI("TanggalStok"), "yyyy/MM/DD HH:MM:SS") & "' AND SaldoAwal='" & rsI("SaldoAwal") & "'  "
                            dbConn.Execute sbQuery
                        
                        If cekbFInsert = True Then GoTo balik
                        If cekbFInsert1 = True Then GoTo balik1
         
            End If
           '
          
            strSQL3 = "Select * from LapSaldoBarangMedis where KdRuanganTujuan = '" & mstrKdRuangan & "' and  KdBarang = '" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "' and TanggalStok BETWEEN '" & Format(mdTglAwal, "yyyy/MM/01 00:00:00") & "'  AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "' order by TanggalStok"
            Call msubRecFO(rsC, strSQL3)
            'END EDIT SIGIT
            If rsC.EOF = False Then
           
                    For j = 1 To rsC.RecordCount
                        .TextMatrix(k, 1) = rsC("KdBarang")
                        .TextMatrix(k, 2) = rsC("KdAsal")
                        .TextMatrix(k, 3) = rsC("NamaBarang")
                        .TextMatrix(k, 4) = rsC("NamaAsal")
        
                        .TextMatrix(k, 5) = IIf(IsNull(rsC("SaldoAwal")), "0", rsC("SaldoAwal"))
                        .TextMatrix(k, 6) = IIf(IsNull(rsC("JmlTerima")), "0", rsC("JmlTerima"))
        
                        .TextMatrix(k, 7) = rsC("JmlKeluar")
                        .TextMatrix(k, 8) = (Val(.TextMatrix(k, 5)) + Val(.TextMatrix(k, 6))) - .TextMatrix(k, 7)
                        .TextMatrix(k, 9) = IIf(IsNull(rsC("HargaJual")), "0", rsC("HargaJual"))
                        If sFltrSaldoBy = "" Then
                            .TextMatrix(k, 10) = IIf(IsNull(rsC("JenisBarang")), "", rsC("JenisBarang"))
                        Else
                            .TextMatrix(k, 10) = IIf(IsNull(rsC("" & sFltrSaldoBy & "")), "", rsC("" & sFltrSaldoBy & ""))
                        End If
        
                        .TextMatrix(k, 11) = rsC("KdRuanganTujuan")
                        .TextMatrix(k, 12) = IIf(IsNull(rsC("TanggalStok")), "0", rsC("TanggalStok"))
                        rsC.MoveNext
                        .Rows = .Rows + 1
                        k = k + 1
                        
                    Next j
                End If
       End With

        rs.MoveNext
        pbData.Value = Int(pbData.Value) + 1
    Next i
  
            
    cmdCetak.Enabled = True
    MousePointer = vbDefault
    Exit Sub
hell:
    pbData.Value = 0.0001
    MousePointer = vbDefault
    Call msubPesanError
    Resume 0
End Sub

Private Sub cmdTutup_Click()
    On Error GoTo hell
    strSQL = ""
    strSQL = "Delete From LaporanSaldoBarangMedis_T where KdRuangan Like '%" & mstrKdRuangan & "%'"
    Call msubRecFO(rs, strSQL)
    sKriteria = ""
    Unload Me
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call cmdBatal_Click
End Sub

Private Sub LoadDcsource()
    On Error GoTo hell
    Call msubDcSource(dcAsalBarang, rs, "Select KdAsal,NamaAsal where KdInstalasi='07'")
    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub setgrid()
    With fgData
        .Clear
        .Rows = 2
        .Cols = 13
        .TextMatrix(0, 1) = "KdBarang"
        .TextMatrix(0, 2) = "KdAsal"
        .TextMatrix(0, 3) = "NamaBarang"
        .TextMatrix(0, 4) = "AsalBarang"
        .TextMatrix(0, 5) = "SaldoAwal"
        .TextMatrix(0, 6) = "JmlTerima"
        .TextMatrix(0, 7) = "JmlKeluar"
        .TextMatrix(0, 8) = "SaldoAkhir"
        .TextMatrix(0, 9) = "HargaNetto"
        .TextMatrix(0, 10) = "JenisBarang"
        .TextMatrix(0, 11) = "KdRuangan"
        .TextMatrix(0, 12) = "Tanggal Stok"

        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = 1500
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 800
        .ColWidth(7) = 800
        .ColWidth(8) = 1000
        .ColWidth(9) = 0
        .ColWidth(10) = 1000
        .ColWidth(11) = 0
        .ColWidth(12) = 1800
    End With
End Sub

Private Sub optBulan_Click()
    dtpAwal.CustomFormat = "MMMM yyyy"
    dtpAkhir.CustomFormat = "MMMM yyyy"
    strCetak = "Bulan"
End Sub

Private Sub optGolBarang_Click()
    chkGroupBy.Caption = "Golongan Barang"
    chkGroupBy.Value = vbChecked
    Call chkGroupBy_Click
End Sub

Private Sub optHari_Click()
    dtpAwal.CustomFormat = "dd MMMM yyyy"
    dtpAkhir.CustomFormat = "dd MMMM yyyy"
    strCetak = "Hari"
End Sub

Private Sub optJnsBarang_Click()
    chkGroupBy.Caption = "Jenis Barang"
    chkGroupBy.Value = vbChecked
    Call chkGroupBy_Click
End Sub

Private Sub optKategoriBrg_Click()
    chkGroupBy.Caption = "Kategory Barang"
    chkGroupBy.Value = vbChecked
    Call chkGroupBy_Click
End Sub

Private Sub optPabrik_Click()
    chkGroupBy.Caption = "Pabrik"
    chkGroupBy.Value = vbChecked
    Call chkGroupBy_Click
End Sub

Private Sub optStatusBrg_Click()
    chkGroupBy.Caption = "Status Barang"
    chkGroupBy.Value = vbChecked
    Call chkGroupBy_Click
End Sub

Private Sub optTahun_Click()
    dtpAwal.CustomFormat = "yyyy"
    dtpAkhir.CustomFormat = "yyyy"
    strCetak = "Tahun"
End Sub

Private Function sp_cetakLaporanSaldobarang(f_KdBarang As String, f_KdAsal As String, f_NamaBarang As String, f_AsalBarang As String, _
    f_SaldoAwal As Double, f_JmlTerima As Double, f_JmlKeluar As Double, f_SaldoAkhir As Double, _
    f_HargaNetto As Double, f_JenisBarang As String, f_KdRuangan As String) As Boolean

    sp_cetakLaporanSaldobarang = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdBarang", adChar, adParamInput, 10, f_KdBarang)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("AsalBarang", adChar, adParamInput, 30, f_AsalBarang)

        .Parameters.Append .CreateParameter("NamaBarang", adChar, adParamInput, 100, f_NamaBarang)
        .Parameters.Append .CreateParameter("SaldoAwal", adDouble, adParamInput, , f_SaldoAwal)
        .Parameters.Append .CreateParameter("JmlTerima", adDouble, adParamInput, , f_JmlTerima)
        .Parameters.Append .CreateParameter("JmlKeluar", adDouble, adParamInput, , f_JmlKeluar)
        .Parameters.Append .CreateParameter("SaldoAkhir", adDouble, adParamInput, , f_SaldoAkhir)
        .Parameters.Append .CreateParameter("HargaNetto", adDouble, adParamInput, , f_HargaNetto)
        If sKriteria = "JenisBarang" Then
            .Parameters.Append .CreateParameter("JenisBarang", adChar, adParamInput, 50, f_JenisBarang)
        Else
            .Parameters.Append .CreateParameter("JenisBarang", adChar, adParamInput, 50, Null)
        End If

        If sKriteria = "KategoryBarang" Then
            .Parameters.Append .CreateParameter("KategoryBarang", adChar, adParamInput, 50, f_JenisBarang)
        Else
            .Parameters.Append .CreateParameter("KategoryBarang", adChar, adParamInput, 50, Null)
        End If
        If sKriteria = "GolonganBarang" Then
            .Parameters.Append .CreateParameter("GolonganBarang", adChar, adParamInput, 50, f_JenisBarang)
        Else
            .Parameters.Append .CreateParameter("GolonganBarang", adChar, adParamInput, 50, Null)
        End If

        If sKriteria = "StatusBarang" Then
            .Parameters.Append .CreateParameter("StatusBarang", adChar, adParamInput, 50, f_JenisBarang)
        Else
            .Parameters.Append .CreateParameter("StatusBarang", adChar, adParamInput, 50, Null)
        End If

        If sKriteria = "NamaPabrik" Then
            .Parameters.Append .CreateParameter("Pabrik", adChar, adParamInput, 50, f_JenisBarang)
        Else
            .Parameters.Append .CreateParameter("Pabrik", adChar, adParamInput, 50, Null)
        End If
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .ActiveConnection = dbConn
        .CommandText = "LaporanSaldoBarangMedis_TAdd"
        .CommandType = adCmdStoredProc
        .Execute
        If .Parameters("RETURN_VALUE").Value <> 0 Then
            MsgBox "Ada Kesalahan dalam Tanggungan Penjamin", vbCritical, "Validasi"
            sp_cetakLaporanSaldobarang = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

