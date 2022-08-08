VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLaporanSaldoBarangNM_v3 
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
   Icon            =   "frmLaporanSaldoBarangNM_v3.frx":0000
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
      TabIndex        =   11
      Top             =   7560
      Width           =   12735
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   495
         Left            =   9720
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   11280
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin MSComctlLib.ProgressBar pbData 
         Height          =   495
         Left            =   120
         TabIndex        =   13
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
         TabIndex        =   14
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
         TabIndex        =   22
         Top             =   960
         Width           =   5775
         Begin VB.OptionButton optMerkBrg 
            Caption         =   "Merk Barang"
            Height          =   495
            Left            =   3600
            TabIndex        =   27
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optWarnaBarang 
            Caption         =   "Warna Barang"
            Height          =   495
            Left            =   4800
            TabIndex        =   26
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optTypeBarang 
            Caption         =   "Type Barang"
            Height          =   495
            Left            =   2400
            TabIndex        =   25
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optBahanBrg 
            Caption         =   "Bahan Barang"
            Height          =   495
            Left            =   1200
            TabIndex        =   24
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optJnsBarang 
            Caption         =   "Jenis Barang"
            Height          =   495
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkGroupBy 
         Caption         =   "Jenis Barang"
         Height          =   255
         Left            =   6000
         TabIndex        =   20
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CheckBox chkAsalBarang 
         Caption         =   "Asal Barang"
         Height          =   255
         Left            =   9480
         TabIndex        =   19
         Top             =   225
         Width           =   1935
      End
      Begin VB.CommandButton cmdProses 
         Caption         =   "&Proses"
         Height          =   495
         Left            =   9840
         TabIndex        =   17
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   495
         Left            =   10920
         TabIndex        =   15
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
            TabIndex        =   6
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.OptionButton optBulan 
            Caption         =   "Per Bulan"
            Height          =   375
            Left            =   1200
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optTotal 
            Caption         =   "Total"
            Enabled         =   0   'False
            Height          =   375
            Left            =   4080
            TabIndex        =   4
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.OptionButton optTahun 
            Caption         =   "Per Tahun"
            Height          =   375
            Left            =   2400
            TabIndex        =   3
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   330
         Left            =   4680
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
         Format          =   54198275
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   330
         Left            =   7200
         TabIndex        =   8
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
         Format          =   109445123
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSDataListLib.DataCombo dcAsalBarang 
         Height          =   330
         Left            =   9480
         TabIndex        =   18
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
         TabIndex        =   21
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
         TabIndex        =   9
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
      TabIndex        =   10
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
      Picture         =   "frmLaporanSaldoBarangNM_v3.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmLaporanSaldoBarangNM_v3.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10935
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   9120
      Picture         =   "frmLaporanSaldoBarangNM_v3.frx":4CE9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "frmLaporanSaldoBarangNM_v3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAsalBarang_Click()
    If chkAsalBarang.Value = vbChecked Then
        dcAsalBarang.Enabled = True
        Call msubDcSource(dcAsalBarang, rs, "Select KdAsal,NamaAsal From AsalBarang Where KdInstalasi = '05'")
        dcAsalBarang.BoundText = rs(0)
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
                Call msubDcSource(dcGroupBy, rs, "SELECT DetailJenisBarang.KdDetailJenisBarang, DetailJenisBarang.DetailJenisBarang, KelompokBarang.KdKelompokBarang " & _
                                                 "FROM DetailJenisBarang INNER JOIN " & _
                                                 "JenisBarang ON DetailJenisBarang.KdJenisBarang = JenisBarang.KdJenisBarang INNER JOIN " & _
                                                 "KelompokBarang ON JenisBarang.KdKelompokBarang = KelompokBarang.KdKelompokBarang " & _
                                                 "WHERE (KelompokBarang.KdKelompokBarang = '01') ")
            Case "Bahan Barang"
                Call msubDcSource(dcGroupBy, rs, "Select KdBahanBarang,NamaBahanBarang From BahanBarangNonMedis")
                dcGroupBy.BoundText = rs(0)
                
            Case "Type Barang"
                Call msubDcSource(dcGroupBy, rs, "Select KdType,NamaType From TypeBarangNonMedis")
                dcGroupBy.BoundText = rs(0)
                
            Case "Merk Barang"
                Call msubDcSource(dcGroupBy, rs, "Select KdMerk,NamaMerk From MerkBarangNonMedis")
                dcGroupBy.BoundText = rs(0)
                
            Case "Warna Barang"
                Call msubDcSource(dcGroupBy, rs, "Select KdWarnaBarang,WarnaBarang From WarnaBarang")
                dcGroupBy.BoundText = rs(0)
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
    optHari.Value = True
    optBulan.Value = False
    optTahun.Value = False
    optBulan.Value = True
    dtpAwal.Value = Now
    dtpAkhir.Value = Now
    dcAsalBarang.Text = ""
    dcGroupBy.Text = ""
    chkAsalBarang.Value = Unchecked
    chkGroupBy.Value = Unchecked
    cmdProses.Enabled = True
    cmdCetak.Enabled = False
    optJnsBarang.Value = True
    pbData.Value = 0.0001
    lblPersen.Caption = "0%"
    Call setgrid
End Sub

Private Sub cmdCetak_Click()
Dim sValues  As String
Dim sValuesMoney As String
Dim sQuery As String
Dim iCols As Integer
Dim i As Integer
Dim j As Integer
    
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    
    Call DeleteTable
    iCols = fgData.Cols
    pbData.Value = 0.0001
    If fgData.Rows = 2 Then
       MsgBox "Tidak ada Data..."
       Exit Sub
    Else
    pbData.Max = fgData.Rows - 2
    End If
    
    
    With fgData
        For i = 1 To .Rows - 2

            sValues = ""
            sValuesMoney = ""
            For j = 3 To iCols - 1
                If j = 3 Then
                    sValuesMoney = "'" & .TextMatrix(i, 3) & "'"
                End If
                 If j <> 3 And j <> 10 Then
                    sValues = sValues & "," & "'" & .TextMatrix(i, j) & "'"
                 ElseIf j = 11 Then
                    sValues = sValues & "," & "'" & .TextMatrix(i, j) & "'"
                End If
            Next j
        
        sQuery = "Insert into V_LapGudangTemp" & _
                " values (" & _
                " " & sValuesMoney & "" & _
                " " & sValues & "" & _
                " )"
        dbConn.Execute sQuery
        
        lblPersen.Caption = Int((i / (fgData.Rows - 2)) * 100) & "%"
        pbData.Value = Int(pbData.Value) + 1

        Next i
        pbData.Value = 0.0001
    End With
    
    strSQL = "Select * from V_LapGudangTemp Where NamaKomputer='" & strNamaHostLocal & "'"
    
    On Error GoTo hell
    mdTglAwal = dtpAwal.Value
    Set frmCetakLapSaldo = Nothing
    frmCetakLapSaldo.Show
Exit Sub
hell:
End Sub

Private Sub cmdProses_Click()
On Error GoTo hell
Dim i, j, k As Integer
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
        
    MousePointer = vbHourglass
    Call setgrid
    cmdProses.Enabled = False
    
    sFltrAsalBrg = ""
    sFltrAsalBrg = ""
    If chkAsalBarang.Value = Checked Then
        sFltrAsalBrg = " AND KdAsal = '" & dcAsalBarang.BoundText & "'"
    Else
        sFltrAsalBrg = ""
    End If
    
    If chkGroupBy.Value = Checked Then
        If optJnsBarang.Value = True Then
            sFltrGroupBy = " AND KdDetailJenisBarang = '" & dcGroupBy.BoundText & "' order by NamaBarang "
            sFltrSaldoBy = "JenisBarang"
        ElseIf optBahanBrg.Value = True Then
            sFltrGroupBy = " AND KdBahanBarang = '" & dcGroupBy.BoundText & "' order by NamaBarang "
            sFltrSaldoBy = "NamaBahanBarang"
        ElseIf optTypeBarang.Value = True Then
            sFltrGroupBy = " AND KdType = '" & dcGroupBy.BoundText & "' order by NamaBarang "
            sFltrSaldoBy = "NamaType"
        ElseIf optMerkBrg.Value = True Then
            sFltrGroupBy = " AND KdMerk = '" & dcGroupBy.BoundText & "' order by NamaBarang "
            sFltrSaldoBy = "NamaMerk"
        ElseIf optWarnaBarang.Value = True Then
            sFltrGroupBy = " AND KdWarnaBarang = '" & dcGroupBy.BoundText & "' order by NamaBarang "
            sFltrSaldoBy = "WarnaBarang"
        End If
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
        intTgl = intHr
        intBln = CInt(Month(dtTgl))
        intThn = CInt(Year(dtTgl))
    End If
    intTgl = intTglLast
    
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value
    
    tgl = ""
    tgl = funcHitungHari(Month(mdTglAkhir), Year(mdTglAkhir))
    
    Set rs = Nothing
    strSQL = "Select KdBarang,KdAsal,NamaBarang,NamaAsal,HargaNetto,KdDetailJenisBarang,KdBahanBarang,KdType,KdMerk,KdWarnaBarang,JenisBarang,NamaBahanBarang,NamaType,WarnaBarang From V_StokBarangNonMedisRuanganLengkap Where KdRuangan='" & mstrKdRuangan & "'" & sFltrAsalBrg & sFltrGroupBy
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
        With fgData
                .TextMatrix(i, 1) = rs("KdBarang") 'dbRst("KdBarang")
                .TextMatrix(i, 2) = rs("KdAsal") 'dbRst("KdAsal")
                .TextMatrix(i, 3) = rs("NamaBarang")
                .TextMatrix(i, 4) = rs("NamaAsal")
                
                strSQL = "SELECT MAX(dbo.Closing.TglClosing) AS TglClosing, dbo.DataStokBarangNonMedis.JmlStokReal AS SaldoAwal, dbo.Closing.KdRuangan, " & _
                         "dbo.DataStokBarangNonMedis.KdBarang , dbo.DataStokBarangNonMedis.KdAsal " & _
                         "FROM dbo.DataStokBarangNonMedis INNER JOIN " & _
                         "dbo.Closing ON dbo.DataStokBarangNonMedis.NoClosing = dbo.Closing.NoClosing " & _
                         "WHERE (TglClosing BETWEEN '" & Format(mdTglAwal, "yyyy/MM/01 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "') AND KdBarang = '" & rs("KdBarang") & "' And KdAsal='" & rs("KdAsal") & "' " & _
                         "GROUP BY dbo.Closing.KdRuangan, dbo.DataStokBarangNonMedis.KdBarang, dbo.DataStokBarangNonMedis.KdAsal, dbo.DataStokBarangNonMedis.JmlStokReal "
                Set dbRst = Nothing
                Call msubRecFO(dbRst, strSQL)
                If dbRst.EOF = True Then
                    .TextMatrix(i, 5) = 0
                Else
                    .TextMatrix(i, 5) = IIf(IsNull(dbRst("SaldoAwal")), 0, dbRst("SaldoAwal"))
                End If
                
                If optHari.Value = True Then
                    strSQL = "SELECT dbo.StrukTerima.TglTerima, SUM(dbo.DetailTerimaBarangNonMedis.JmlTerima) AS JmlTerima, dbo.StrukTerima.KdRuangan " & _
                             "FROM dbo.StrukTerima INNER JOIN " & _
                             "dbo.DetailTerimaBarangNonMedis ON dbo.StrukTerima.NoTerima = dbo.DetailTerimaBarangNonMedis.NoTerima " & _
                             "WHERE (TglTerima BETWEEN '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') And KdBarang = '" & rs("KdBarang") & "' And KdAsal='" & rs("KdAsal") & "' " & _
                             "GROUP BY dbo.StrukTerima.TglTerima, dbo.StrukTerima.KdRuangan "
                             
                ElseIf optBulan.Value = True Then
                    strSQL = "SELECT dbo.StrukTerima.TglTerima, SUM(dbo.DetailTerimaBarangNonMedis.JmlTerima) AS JmlTerima, dbo.StrukTerima.KdRuangan " & _
                             "FROM dbo.StrukTerima INNER JOIN " & _
                             "dbo.DetailTerimaBarangNonMedis ON dbo.StrukTerima.NoTerima = dbo.DetailTerimaBarangNonMedis.NoTerima " & _
                             "WHERE (TglTerima BETWEEN '" & Format(mdTglAwal, "yyyy/MM/01 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "') And KdBarang = '" & rs("KdBarang") & "' And KdAsal='" & rs("KdAsal") & "' " & _
                             "GROUP BY dbo.StrukTerima.TglTerima, dbo.StrukTerima.KdRuangan "
                ElseIf optTahun.Value = True Then
                    strSQL = "SELECT dbo.StrukTerima.TglTerima, SUM(dbo.DetailTerimaBarangNonMedis.JmlTerima) AS JmlTerima, dbo.StrukTerima.KdRuangan " & _
                             "FROM dbo.StrukTerima INNER JOIN " & _
                             "dbo.DetailTerimaBarangNonMedis ON dbo.StrukTerima.NoTerima = dbo.DetailTerimaBarangNonMedis.NoTerima " & _
                             "WHERE (TglTerima BETWEEN '" & Format(mdTglAwal, "yyyy/01/01 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/12/" & tgl & " 23:59:59") & "') And KdBarang = '" & rs("KdBarang") & "' And KdAsal='" & rs("KdAsal") & "' " & _
                             "GROUP BY dbo.StrukTerima.TglTerima, dbo.StrukTerima.KdRuangan "
                End If
                
                Set dbRst = Nothing
                Call msubRecFO(dbRst, strSQL)
                If dbRst.EOF = True Then
                    .TextMatrix(i, 6) = 0
                Else
                    .TextMatrix(i, 6) = dbRst("JmlTerima")
                End If
                
                If optHari.Value = True Then
                    strSQL = "SELECT dbo.StrukKirim.TglKirim, SUM(dbo.DetailBarangNonMedisKeluar.JmlKirim) AS JmlKirim, dbo.StrukKirim.KdRuangan " & _
                             "FROM dbo.StrukKirim INNER JOIN " & _
                             "dbo.DetailBarangNonMedisKeluar ON dbo.StrukKirim.NoKirim = dbo.DetailBarangNonMedisKeluar.NoKirim " & _
                             "WHERE (TglKirim BETWEEN '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') And KdBarang = '" & rs("KdBarang") & "' And KdAsal='" & rs("KdAsal") & "' " & _
                             "GROUP BY dbo.StrukKirim.TglKirim, dbo.StrukKirim.KdRuangan "
                             
                ElseIf optBulan.Value = True Then
                    strSQL = "SELECT dbo.StrukKirim.TglKirim, SUM(dbo.DetailBarangNonMedisKeluar.JmlKirim) AS JmlKirim, dbo.StrukKirim.KdRuangan " & _
                             "FROM dbo.StrukKirim INNER JOIN " & _
                             "dbo.DetailBarangNonMedisKeluar ON dbo.StrukKirim.NoKirim = dbo.DetailBarangNonMedisKeluar.NoKirim " & _
                             "WHERE (TglKirim BETWEEN '" & Format(mdTglAwal, "yyyy/MM/01 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/" & tgl & " 23:59:59") & "') And KdBarang = '" & rs("KdBarang") & "' And KdAsal='" & rs("KdAsal") & "' " & _
                             "GROUP BY dbo.StrukKirim.TglKirim, dbo.StrukKirim.KdRuangan "
                ElseIf optTahun.Value = True Then
                    strSQL = "SELECT dbo.StrukKirim.TglKirim, SUM(dbo.DetailBarangNonMedisKeluar.JmlKirim) AS JmlKirim, dbo.StrukKirim.KdRuangan " & _
                             "FROM dbo.StrukKirim INNER JOIN " & _
                             "dbo.DetailBarangNonMedisKeluar ON dbo.StrukKirim.NoKirim = dbo.DetailBarangNonMedisKeluar.NoKirim " & _
                             "WHERE (TglKirim BETWEEN '" & Format(mdTglAwal, "yyyy/01/01 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/12/" & tgl & " 23:59:59") & "') And KdBarang = '" & rs("KdBarang") & "' And KdAsal='" & rs("KdAsal") & "' " & _
                             "GROUP BY dbo.StrukKirim.TglKirim, dbo.StrukKirim.KdRuangan "
                End If
                
                Set dbRst = Nothing
                Call msubRecFO(dbRst, strSQL)
                If dbRst.EOF = True Then
                    .TextMatrix(i, 7) = 0
                Else
                    .TextMatrix(i, 7) = dbRst("JmlKirim")
                End If
                
                
                .TextMatrix(i, 8) = CDbl(.TextMatrix(i, 5)) + CDbl(.TextMatrix(i, 6)) - CDbl(.TextMatrix(i, 7))
                .TextMatrix(i, 9) = rs("HargaNetto")
                If sFltrSaldoBy = "" Then
                    .TextMatrix(i, 10) = ""
                Else
                    .TextMatrix(i, 10) = IIf(IsNull(rs("" & sFltrSaldoBy & "")), "", rs("" & sFltrSaldoBy & ""))
                End If
                .TextMatrix(i, 11) = strNamaHostLocal
                
                .Rows = .Rows + 1
        End With
        
        rs.MoveNext
        pbData.Value = Int(pbData.Value) + 1
    Next i
        If dbRst.RecordCount = 0 Then
       MsgBox "Saldo barang tidak ditemukan", vbInformation, "Info"
       cmdCetak.Enabled = False
       cmdProses.Enabled = True
       MousePointer = vbDefault
       Exit Sub
    Else
        MsgBox "Proses pencarian saldo barang berhasil", vbInformation, "Info"
    End If
    
    cmdCetak.Enabled = True
    MousePointer = vbDefault
Exit Sub
hell:
    pbData.Value = 0.0001
    MousePointer = vbDefault
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
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
        .Cols = 12
        
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
' ==================== 05 Januari 2010 By 703 (Menambahkan Nama Komputer untuk dimasukan dalam tabel temp) ===========
        .TextMatrix(0, 11) = "NamaKomputer"
' ==================== 05 Januari 2010 By 703 =====================
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = 1500
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 800
        .ColWidth(7) = 800
        .ColWidth(8) = 1000
        .ColWidth(9) = 1000
        .ColWidth(10) = 1000
' ==================== 05 Januari 2010 By 703 =====================
        .ColWidth(11) = 0
' ==================== 05 Januari 2010 By 703 =====================
    End With
End Sub

Private Sub optBulan_Click()
    dtpAwal.CustomFormat = "MMMM yyyy"
    dtpAkhir.CustomFormat = "MMMM yyyy"
End Sub

Private Sub optTypeBarang_Click()
    chkGroupBy.Caption = "Type Barang"
    Call chkGroupBy_Click
End Sub

Private Sub optHari_Click()
    dtpAwal.CustomFormat = "dd MMMM yyyy"
    dtpAkhir.CustomFormat = "dd MMMM yyyy"
End Sub

Private Sub optJnsBarang_Click()
    chkGroupBy.Caption = "Jenis Barang"
    Call chkGroupBy_Click
End Sub

Private Sub optBahanBrg_Click()
    chkGroupBy.Caption = "Bahan Barang"
    Call chkGroupBy_Click
End Sub

Private Sub optWarnaBarang_Click()
    chkGroupBy.Caption = "Warna Barang"
    Call chkGroupBy_Click
End Sub

Private Sub optMerkBrg_Click()
    chkGroupBy.Caption = "Merk Barang"
    Call chkGroupBy_Click
End Sub

Private Sub optTahun_Click()
    dtpAwal.CustomFormat = "yyyy"
    dtpAkhir.CustomFormat = "yyyy"
End Sub

Public Sub CreateTabel()
Dim sQuery As String

    sQuery = "create Table V_LapGudang_NewTemp" & "_" & strNamaHostLocal & _
             " (NamaBarang varchar(50)," & _
             " AsalBarang varchar(25)," & _
             " SaldoAwal real," & _
             " JmlTerima real ," & _
             " JmlKeluar real ," & _
             " SaldoAkhir real," & _
             " HargaNetto money)"
    dbConn.Execute sQuery
    
End Sub
    
Public Sub DeleteTable()
Dim sQuery As String
On Error Resume Next
    sQuery = "delete V_LapGudangTemp Where NamaKomputer= '" & strNamaHostLocal & "'"
    dbConn.Execute sQuery
End Sub


