VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmRekapLaporanKomp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Rekapituasi"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRekapLaporanKomp.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   14655
   Begin VB.Frame Frame1 
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
      TabIndex        =   8
      Top             =   7320
      Width           =   14655
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   495
         Left            =   11520
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   13080
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
   End
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
      Height          =   6255
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   14655
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
         TabIndex        =   10
         Top             =   150
         Width           =   5775
         Begin VB.CommandButton cmdTampilkanTemp 
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
            TabIndex        =   4
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   2
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
            Format          =   53805059
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   3
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
            Format          =   53805059
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   11
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataListLib.DataCombo dcPeriode 
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
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
      Begin MSDataListLib.DataCombo dcJenisLaporan 
         Height          =   360
         Left            =   2520
         TabIndex        =   1
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
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
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   5055
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   8916
         _Version        =   393216
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Laporan"
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
         Left            =   2520
         TabIndex        =   12
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Periode"
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
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1200
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
   Begin VB.Image Image2 
      Height          =   945
      Left            =   12840
      Picture         =   "frmRekapLaporanKomp.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRekapLaporanKomp.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRekapLaporanKomp.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmRekapLaporanKomp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i, j As Integer
Dim intJmlRow As Integer
Dim intRowNow As Integer
Dim rsa As New ADODB.recordset
Dim rsB As New ADODB.recordset
Dim rsC As New ADODB.recordset

Dim subTotalBiaya As Currency
Dim subTotalBayar As Currency
Dim subTotalPiutang As Currency
Dim subTotalTanggunganRS As Currency
Dim subTotalPembebasan As Currency
Dim subTotalSisaTagihan As Currency

Private Sub cmdCetak_Click()
    If dcPeriode.MatchedWithList = False Then Exit Sub
    If dcJenisLaporan.MatchedWithList = False Then Exit Sub
    
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value
    
    If dcJenisLaporan.BoundText = "01" Then
        mstrLaporan = "PenerimaanKasirPerda"
    ElseIf dcJenisLaporan.BoundText = "02" Then
        mstrLaporan = "PenerimaanKasirNonPerda"
    ElseIf dcJenisLaporan.BoundText = "03" Then
        mstrLaporan = "PenerimaanKasirTotal"
    End If
    Set frmCetakLaporanKomp = Nothing
    frmCetakLaporanKomp.Show
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdTampilkanTemp_Click()
On Error GoTo errTampilkan
        
    'validasi periode
    If dcPeriode.MatchedWithList = False Then
        dcPeriode.SetFocus
        Exit Sub
    End If
    
    'validasi jenis laporan
    If dcJenisLaporan.MatchedWithList = False Then
        dcJenisLaporan.SetFocus
        Exit Sub
    End If
    
    subSetGridAja
    intJmlRow = 0
    fgData.Visible = False: MousePointer = vbHourglass
    
    Call subLoadDataPerdaNonPerda
    
    fgData.Visible = True: MousePointer = vbNormal
    Exit Sub
errTampilkan:
    msubPesanError
    fgData.Visible = True: MousePointer = vbNormal
End Sub

Private Sub dcJenisLaporan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdTampilkanTemp.SetFocus
End Sub

Private Sub dcPeriode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAwal.SetFocus
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcJenisLaporan.SetFocus
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo errFormLoad
    Call centerForm(Me, MDIUtama)
    
    dtpAwal.Value = Now
    dtpAkhir.Value = Now
    
    Me.Caption = "Medifirst2000 - Laporan Pendapatan Ruangan"
    strSQL = "SELECT * FROM V_S_Lap_PenerimaanKasir_Perda WHERE 1=2"
    
    msubRecFO rs, strSQL
    Call subLoadDC
    Call subSetGridAja
    Call PlayFlashMovie(Me)
    Exit Sub
errFormLoad:
    msubPesanError
End Sub

Private Sub subLoadDC()
On Error GoTo errLoad
    strSQL = "SELECT KdJenisLaporan, JenisLaporan FROM JenisLaporan"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    
    Set dcPeriode.RowSource = rs
    dcPeriode.BoundColumn = rs(0).Name
    dcPeriode.ListField = rs(1).Name
    If rs.EOF = False Then dcPeriode.BoundText = rs(0).Value
    
    strSQL = "SELECT KdJenisLaporan,JenisLaporan FROM JenisLaporanKasir"
    msubRecFO rs, strSQL
    Set dcJenisLaporan.RowSource = rs
    dcJenisLaporan.BoundColumn = rs(0).Name
    dcJenisLaporan.ListField = rs(1).Name
    If Not rs.EOF Then dcJenisLaporan.Text = rs(1).Value
    
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

'untuk setting grid tanpa loading data
Private Sub subSetGridAja()
Dim i As Integer
    With fgData
        .Clear
        .Cols = 11
        .Rows = 2
        .ColWidth(0) = 150
        .ColWidth(1) = 1200
        .ColWidth(2) = 1100
        .ColWidth(3) = 1500
        .ColWidth(4) = 1500
        .ColWidth(5) = 1450
        .ColWidth(6) = 1450
        .ColWidth(7) = 1450
        .ColWidth(8) = 1450
        .ColWidth(9) = 1450
        .ColWidth(10) = 1450
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignRightCenter
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignRightCenter
        .ColAlignment(10) = flexAlignRightCenter
        .Row = 0
        .RowHeight(0) = 300
        For i = 1 To .Cols - 1
            .Col = i
            .CellAlignment = flexAlignCenterCenter
            .CellFontBold = True
        Next i
        .TextMatrix(0, 1) = "Ruangan"
        .TextMatrix(0, 2) = dcPeriode.Text
        .TextMatrix(0, 3) = "Jenis Pasien"
        .TextMatrix(0, 4) = "Komponen Tarif"
        .TextMatrix(0, 5) = "Total Biaya"
        .TextMatrix(0, 6) = "Bayar"
        .TextMatrix(0, 7) = "Piutang"
        .TextMatrix(0, 8) = "Tanggungan RS"
        .TextMatrix(0, 9) = "Pembebasan"
        .TextMatrix(0, 10) = "Sisa Tagihan"
        .MergeCells = 1
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
    End With
End Sub

Private Sub subLoadRuanganPeriodeJenisPasien(strCase As String, strJenisLaporan As String)
On Error GoTo errLoad

    Select Case strCase
        Case "01" 'per Hari
            If strJenisLaporan = "perda" Then
                strSQL = "SELECT [Ruang Pelayanan], dbo.S_AmbilTanggal(TglStruk) AS Alias, JenisPasien, SUM(Total) AS TotalBiaya, SUM(JmlBayar) AS TotalBayar, SUM(Piutang) AS TotalPiutang, SUM(TanggunganRS) AS TotalTanggunganRS, SUM(Pembebasan) AS TotalPembebasan, SUM(SisaTagihan) AS TotalSisaTagihan " & _
                    " From  V_RekapPendapatanRSPerKomponen " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " AND (KdKomponen IN ('01','02','03'))" & _
                    " GROUP BY [Ruang Pelayanan], dbo.S_AmbilTanggal(TglStruk), JenisPasien" & _
                    " ORDER BY [Ruang Pelayanan], dbo.S_AmbilTanggal(TglStruk), JenisPasien"
            ElseIf strJenisLaporan = "nonperda" Then
                strSQL = "SELECT [Ruang Pelayanan], dbo.S_AmbilTanggal(TglStruk) AS Alias, JenisPasien, SUM(Total) AS TotalBiaya, SUM(JmlBayar) AS TotalBayar, SUM(Piutang)  AS TotalPiutang, SUM(TanggunganRS) AS TotalTanggunganRS, SUM(Pembebasan) AS TotalPembebasan, SUM(SisaTagihan) AS TotalSisaTagihan " & _
                    " From  V_RekapPendapatanRSPerKomponenNonPerda " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " GROUP BY [Ruang Pelayanan], dbo.S_AmbilTanggal(TglStruk), JenisPasien" & _
                    " ORDER BY [Ruang Pelayanan], dbo.S_AmbilTanggal(TglStruk), JenisPasien"
            ElseIf strJenisLaporan = "total" Then
                strSQL = "SELECT [Ruang Pelayanan], dbo.S_AmbilTanggal(TglStruk) AS Alias, JenisPasien, SUM(Total) AS TotalBiaya, SUM(JmlBayar) AS TotalBayar, SUM(Piutang)  AS TotalPiutang, SUM(TanggunganRS) AS TotalTanggunganRS, SUM(Pembebasan) AS TotalPembebasan, SUM(SisaTagihan) AS TotalSisaTagihan " & _
                    " From  V_RekapPendapatanRSTotal " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " GROUP BY [Ruang Pelayanan], dbo.S_AmbilTanggal(TglStruk), JenisPasien" & _
                    " ORDER BY [Ruang Pelayanan], dbo.S_AmbilTanggal(TglStruk), JenisPasien"
            End If
                
        Case "02" 'per Bulan
            If strJenisLaporan = "perda" Then
                strSQL = "SELECT [Ruang Pelayanan], (cast(dbo.UbahBulan(MONTH(TglStruk)) as varchar) + '/' + cast(year(TglStruk) as varchar))AS Alias, JenisPasien, SUM(Total) AS TotalBiaya, SUM(JmlBayar) AS TotalBayar, SUM(Piutang) AS TotalPiutang, SUM(TanggunganRS) AS TotalTanggunganRS, SUM(Pembebasan) AS TotalPembebasan, SUM(SisaTagihan) AS TotalSisaTagihan " & _
                    " From  V_RekapPendapatanRSPerKomponen " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " AND (KdKomponen IN ('01','02','03'))" & _
                    " GROUP BY [Ruang Pelayanan], cast(dbo.UbahBulan(MONTH(TglStruk)) as varchar) + '/' + cast(year(TglStruk) as varchar), JenisPasien" & _
                    " ORDER BY [Ruang Pelayanan], cast(dbo.UbahBulan(MONTH(TglStruk)) as varchar) + '/' + cast(year(TglStruk) as varchar), JenisPasien"
            ElseIf strJenisLaporan = "nonperda" Then
                strSQL = "SELECT [Ruang Pelayanan], (cast(dbo.UbahBulan(MONTH(TglStruk)) as varchar) + '/' + cast(year(TglStruk) as varchar))AS Alias, JenisPasien, SUM(Total) AS TotalBiaya, SUM(JmlBayar) AS TotalBayar, SUM(Piutang)  AS TotalPiutang, SUM(TanggunganRS) AS TotalTanggunganRS, SUM(Pembebasan) AS TotalPembebasan, SUM(SisaTagihan) AS TotalSisaTagihan " & _
                    " From  V_RekapPendapatanRSPerKomponenNonPerda " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " GROUP BY [Ruang Pelayanan], cast(dbo.UbahBulan(MONTH(TglStruk)) as varchar) + '/' + cast(year(TglStruk) as varchar), JenisPasien" & _
                    " ORDER BY [Ruang Pelayanan], cast(dbo.UbahBulan(MONTH(TglStruk)) as varchar) + '/' + cast(year(TglStruk) as varchar), JenisPasien"
            ElseIf strJenisLaporan = "total" Then
                strSQL = "SELECT [Ruang Pelayanan], (cast(dbo.UbahBulan(MONTH(TglStruk)) as varchar) + '/' + cast(year(TglStruk) as varchar))AS Alias, JenisPasien, SUM(Total) AS TotalBiaya, SUM(JmlBayar) AS TotalBayar, SUM(Piutang)  AS TotalPiutang, SUM(TanggunganRS) AS TotalTanggunganRS, SUM(Pembebasan) AS TotalPembebasan, SUM(SisaTagihan) AS TotalSisaTagihan " & _
                    " From  V_RekapPendapatanRSTotal " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " GROUP BY [Ruang Pelayanan], cast(dbo.UbahBulan(MONTH(TglStruk)) as varchar) + '/' + cast(year(TglStruk) as varchar), JenisPasien" & _
                    " ORDER BY [Ruang Pelayanan], cast(dbo.UbahBulan(MONTH(TglStruk)) as varchar) + '/' + cast(year(TglStruk) as varchar), JenisPasien"
            End If
            
        Case "03" 'per Tahun
            If strJenisLaporan = "perda" Then
                strSQL = "SELECT [Ruang Pelayanan], YEAR(TglStruk) AS Alias, JenisPasien, SUM(Total) AS TotalBiaya, SUM(JmlBayar) AS TotalBayar, SUM(Piutang) AS TotalPiutang, SUM(TanggunganRS) AS TotalTanggunganRS, SUM(Pembebasan) AS TotalPembebasan, SUM(SisaTagihan) AS TotalSisaTagihan " & _
                    " From  V_RekapPendapatanRSPerKomponen " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " AND (KdKomponen IN ('01','02','03'))" & _
                    " GROUP BY [Ruang Pelayanan], YEAR(TglStruk), JenisPasien" & _
                    " ORDER BY [Ruang Pelayanan], YEAR(TglStruk), JenisPasien"
            ElseIf strJenisLaporan = "nonperda" Then
                strSQL = "SELECT [Ruang Pelayanan], YEAR(TglStruk) AS Alias, JenisPasien, SUM(Total) AS TotalBiaya, SUM(JmlBayar) AS TotalBayar, SUM(Piutang) AS TotalPiutang, SUM(TanggunganRS) AS TotalTanggunganRS, SUM(Pembebasan) AS TotalPembebasan, SUM(SisaTagihan) AS TotalSisaTagihan " & _
                    " From  V_RekapPendapatanRSPerKomponenNonPerda " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " GROUP BY [Ruang Pelayanan], YEAR(TglStruk), JenisPasien" & _
                    " ORDER BY [Ruang Pelayanan], YEAR(TglStruk), JenisPasien"
            ElseIf strJenisLaporan = "total" Then
                strSQL = "SELECT [Ruang Pelayanan], YEAR(TglStruk) AS Alias, JenisPasien, SUM(Total) AS TotalBiaya, SUM(JmlBayar) AS TotalBayar, SUM(Piutang)  AS TotalPiutang, SUM(TanggunganRS) AS TotalTanggunganRS, SUM(Pembebasan) AS TotalPembebasan, SUM(SisaTagihan) AS TotalSisaTagihan " & _
                    " From  V_RekapPendapatanRSTotal " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " GROUP BY [Ruang Pelayanan], YEAR(TglStruk), JenisPasien" & _
                    " ORDER BY [Ruang Pelayanan], YEAR(TglStruk), JenisPasien"
            End If
        
        Case "04" 'per Periode
            If strJenisLaporan = "perda" Then
                strSQL = "SELECT [Ruang Pelayanan], '-' As Alias, JenisPasien, SUM(Total) AS TotalBiaya, SUM(JmlBayar) AS TotalBayar, SUM(Piutang)  AS TotalPiutang, SUM(TanggunganRS) AS TotalTanggunganRS, SUM(Pembebasan) AS TotalPembebasan, SUM(SisaTagihan) AS TotalSisaTagihan " & _
                    " From  V_RekapPendapatanRSPerKomponen " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " AND (KdKomponen IN ('01','02','03'))" & _
                    " GROUP BY [Ruang Pelayanan], JenisPasien" & _
                    " ORDER BY [Ruang Pelayanan], JenisPasien"
            ElseIf strJenisLaporan = "nonperda" Then
                strSQL = "SELECT [Ruang Pelayanan], '-' AS Alias, JenisPasien, SUM(Total) AS TotalBiaya, SUM(JmlBayar) AS TotalBayar, SUM(Piutang)  AS TotalPiutang, SUM(TanggunganRS) AS TotalTanggunganRS, SUM(Pembebasan) AS TotalPembebasan, SUM(SisaTagihan) AS TotalSisaTagihan " & _
                    " From  V_RekapPendapatanRSPerKomponenNonPerda " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " GROUP BY [Ruang Pelayanan], JenisPasien" & _
                    " ORDER BY [Ruang Pelayanan], JenisPasien"
            ElseIf strJenisLaporan = "total" Then
                strSQL = "SELECT [Ruang Pelayanan], '-'  AS Alias, JenisPasien, SUM(Total) AS TotalBiaya, SUM(JmlBayar) AS TotalBayar, SUM(Piutang)  AS TotalPiutang, SUM(TanggunganRS) AS TotalTanggunganRS, SUM(Pembebasan) AS TotalPembebasan, SUM(SisaTagihan) AS TotalSisaTagihan " & _
                    " From  V_RekapPendapatanRSTotal " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " GROUP BY [Ruang Pelayanan], JenisPasien" & _
                    " ORDER BY [Ruang Pelayanan], JenisPasien"
            End If
        
        Case Else
            strSQL = ""
            
    End Select
    
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadDataPerdaNonPerda()
    'data yang group Ruangan, Periode, JenisPasien
    If dcJenisLaporan.BoundText = "01" Then 'perda
        Call subLoadRuanganPeriodeJenisPasien(dcPeriode.BoundText, "perda")
    ElseIf dcJenisLaporan.BoundText = "02" Then 'noon perda
        Call subLoadRuanganPeriodeJenisPasien(dcPeriode.BoundText, "nonperda")
    ElseIf dcJenisLaporan.BoundText = "03" Then 'total
        Call subLoadRuanganPeriodeJenisPasien(dcPeriode.BoundText, "total")
    End If
    If strSQL = "" Then Exit Sub
    msubRecFO rsa, strSQL
    intJmlRow = intJmlRow + rsa.RecordCount
    
    'semua data
    If dcJenisLaporan.BoundText = "01" Then 'perda
        Call subAmbilData(dcPeriode.BoundText, "perda")
    ElseIf dcJenisLaporan.BoundText = "02" Then 'noon perda
        Call subAmbilData(dcPeriode.BoundText, "nonperda")
    ElseIf dcJenisLaporan.BoundText = "03" Then 'total
        Call subAmbilData(dcPeriode.BoundText, "total")
    End If
    If strSQL = "" Then Exit Sub
    msubRecFO rs, strSQL
    
    'jumlah baris keseluruhan
    intJmlRow = intJmlRow + rs.RecordCount
    
    fgData.Rows = intJmlRow + 2
    intRowNow = 0
    subTotalBiaya = 0: subTotalBayar = 0: subTotalPiutang = 0: subTotalTanggunganRS = 0: subTotalPembebasan = 0: subTotalSisaTagihan = 0

    For i = 1 To rs.RecordCount
        intRowNow = intRowNow + 1
        For j = 1 To fgData.Cols - 1
            fgData.TextMatrix(intRowNow, j) = rs(j - 1).Value
        Next j
        rs.MoveNext
        'sub total per JenisPasien
        If rs.EOF = True Then GoTo stepSubTotalJenisPasien
        If rs("JenisPasien").Value <> rsa("JenisPasien").Value Then
stepSubTotalJenisPasien:
            intRowNow = intRowNow + 1
            fgData.TextMatrix(intRowNow, 1) = fgData.TextMatrix(intRowNow - 1, 1)
            fgData.TextMatrix(intRowNow, 2) = fgData.TextMatrix(intRowNow - 1, 2)
            fgData.TextMatrix(intRowNow, 3) = fgData.TextMatrix(intRowNow - 1, 3)
            fgData.TextMatrix(intRowNow, 4) = "Sub Total"
            fgData.TextMatrix(intRowNow, 5) = IIf(rsa("TotalBiaya").Value = 0, 0, Format(rsa("TotalBiaya").Value, "#,###"))
            fgData.TextMatrix(intRowNow, 6) = IIf(rsa("TotalBayar").Value = 0, 0, Format(rsa("TotalBayar").Value, "#,###"))
            fgData.TextMatrix(intRowNow, 7) = IIf(rsa("TotalPiutang").Value = 0, 0, Format(rsa("TotalPiutang").Value, "#,###"))
            fgData.TextMatrix(intRowNow, 8) = IIf(rsa("TotalTanggunganRS").Value = 0, 0, Format(rsa("TotalTanggunganRS").Value, "#,###"))
            fgData.TextMatrix(intRowNow, 9) = IIf(rsa("TotalPembebasan").Value = 0, 0, Format(rsa("TotalPembebasan").Value, "#,###"))
            fgData.TextMatrix(intRowNow, 10) = IIf(rsa("TotalSisaTagihan").Value = 0, 0, Format(rsa("TotalSisaTagihan").Value, "#,###"))
            
            subTotalBiaya = subTotalBiaya + rsa("TotalBiaya")
            subTotalBayar = subTotalBayar + rsa("TotalBayar")
            subTotalPiutang = subTotalPiutang + rsa("TotalPiutang")
            subTotalTanggunganRS = subTotalTanggunganRS + rsa("TotalTanggunganRS")
            subTotalPembebasan = subTotalPembebasan + rsa("TotalPembebasan")
            subTotalSisaTagihan = subTotalSisaTagihan + rsa("TotalSisaTagihan")
            
            subSetSubTotalRow intRowNow, 2, vbBlackness, vbWhite
            If rsa.EOF Then Exit Sub
            rsa.MoveNext
        ElseIf rs("JenisPasien").Value = rsa("JenisPasien").Value And rs("Alias").Value <> rsa("Alias").Value Then
            GoTo stepSubTotalJenisPasien
        End If
    Next i
    
    intRowNow = intRowNow + 1
    fgData.TextMatrix(intRowNow, 1) = "Total"
    fgData.TextMatrix(intRowNow, 5) = Format(subTotalBiaya, "#,###")
    fgData.TextMatrix(intRowNow, 6) = Format(subTotalBayar, "#,###")
    fgData.TextMatrix(intRowNow, 7) = Format(subTotalPiutang, "#,###")
    fgData.TextMatrix(intRowNow, 8) = Format(subTotalTanggunganRS, "#,###")
    fgData.TextMatrix(intRowNow, 9) = Format(subTotalPembebasan, "#,###")
    fgData.TextMatrix(intRowNow, 10) = Format(subTotalSisaTagihan, "#,###")
    subSetSubTotalRow intRowNow, 1, vbBlue, vbWhite
End Sub

Private Sub subSetSubTotalRow(iRowNow As Integer, iColBegin As Integer, vbBackColor, vbForeColor)
Dim i As Integer
    With fgData
        'tampilan Black & White
        For i = iColBegin To .Cols - 1
            .Row = iRowNow
            .Col = i
            .CellBackColor = vbBackColor
            .CellForeColor = vbForeColor
            .CellFontBold = True
        Next
    End With
End Sub

Private Sub subAmbilData(strCase As String, strJenisLaporan As String)
On Error GoTo errLoad

    Select Case strCase
        Case "01" 'per Hari
            If strJenisLaporan = "perda" Then
                strSQL = "SELECT [Ruang Pelayanan], dbo.S_AmbilTanggal(TglStruk) AS Alias, JenisPasien, [Komponen Tarif], SUM(Total) AS Total, SUM(JmlBayar) AS JmlBayar, SUM(Piutang)  AS Piutang, SUM(TanggunganRS) AS TanggunganRS, SUM(Pembebasan) AS Pembebasan, SUM(SisaTagihan) AS SisaTagihan " & _
                    " From  V_RekapPendapatanRSPerKomponen " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " AND (KdKomponen IN ('01','02','03'))" & _
                    " GROUP BY [Ruang Pelayanan], dbo.S_AmbilTanggal(TglStruk), JenisPasien, [Komponen Tarif] " & _
                    " ORDER BY [Ruang Pelayanan], dbo.S_AmbilTanggal(TglStruk), JenisPasien, [Komponen Tarif]"
            ElseIf strJenisLaporan = "nonperda" Then
                strSQL = "SELECT [Ruang Pelayanan], dbo.S_AmbilTanggal(TglStruk) AS Alias, JenisPasien, [Komponen Tarif], SUM(Total) AS Total, SUM(JmlBayar) AS JmlBayar, SUM(Piutang)  AS Piutang, SUM(TanggunganRS) AS TanggunganRS, SUM(Pembebasan) AS Pembebasan, SUM(SisaTagihan) AS SisaTagihan " & _
                    " From  V_RekapPendapatanRSPerKomponenNonPerda " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " GROUP BY [Ruang Pelayanan], dbo.S_AmbilTanggal(TglStruk), JenisPasien, [Komponen Tarif] " & _
                    " ORDER BY [Ruang Pelayanan], dbo.S_AmbilTanggal(TglStruk), JenisPasien, [Komponen Tarif]"
            ElseIf strJenisLaporan = "total" Then
                strSQL = "SELECT [Ruang Pelayanan], dbo.S_AmbilTanggal(TglStruk) AS Alias, JenisPasien, [Komponen Tarif], SUM(Total) AS Total, SUM(JmlBayar) AS JmlBayar, SUM(Piutang)  AS Piutang, SUM(TanggunganRS) AS TanggunganRS, SUM(Pembebasan) AS Pembebasan, SUM(SisaTagihan) AS SisaTagihan " & _
                    " From  V_RekapPendapatanRSTotal " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " GROUP BY [Ruang Pelayanan], dbo.S_AmbilTanggal(TglStruk), JenisPasien, [Komponen Tarif] " & _
                    " ORDER BY [Ruang Pelayanan], dbo.S_AmbilTanggal(TglStruk), JenisPasien, [Komponen Tarif]"
            End If
                
        Case "02" 'per Bulan
            If strJenisLaporan = "perda" Then
                strSQL = "SELECT [Ruang Pelayanan], (cast(dbo.UbahBulan(MONTH(TglStruk)) as varchar) + '/' + cast(year(TglStruk) as varchar))AS Alias, JenisPasien, [Komponen Tarif], SUM(Total) AS Total, SUM(JmlBayar) AS JmlBayar, SUM(Piutang)  AS Piutang, SUM(TanggunganRS) AS TanggunganRS, SUM(Pembebasan) AS Pembebasan, SUM(SisaTagihan) AS SisaTagihan " & _
                    " From  V_RekapPendapatanRSPerKomponen " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " AND (KdKomponen IN ('01','02','03'))" & _
                    " GROUP BY [Ruang Pelayanan], cast(dbo.UbahBulan(MONTH(TglStruk)) as varchar) + '/' + cast(year(TglStruk) as varchar), JenisPasien, [Komponen Tarif] " & _
                    " ORDER BY [Ruang Pelayanan], cast(dbo.UbahBulan(MONTH(TglStruk)) as varchar) + '/' + cast(year(TglStruk) as varchar), JenisPasien, [Komponen Tarif]"
            ElseIf strJenisLaporan = "nonperda" Then
                strSQL = "SELECT [Ruang Pelayanan], (cast(dbo.UbahBulan(MONTH(TglStruk)) as varchar) + '/' + cast(year(TglStruk) as varchar))AS Alias, JenisPasien, [Komponen Tarif], SUM(Total) AS Total, SUM(JmlBayar) AS JmlBayar, SUM(Piutang)  AS Piutang, SUM(TanggunganRS) AS TanggunganRS, SUM(Pembebasan) AS Pembebasan, SUM(SisaTagihan) AS SisaTagihan " & _
                    " From  V_RekapPendapatanRSPerKomponenNonPerda " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " GROUP BY [Ruang Pelayanan], cast(dbo.UbahBulan(MONTH(TglStruk)) as varchar) + '/' + cast(year(TglStruk) as varchar), JenisPasien, [Komponen Tarif] " & _
                    " ORDER BY [Ruang Pelayanan], cast(dbo.UbahBulan(MONTH(TglStruk)) as varchar) + '/' + cast(year(TglStruk) as varchar), JenisPasien, [Komponen Tarif]"
            ElseIf strJenisLaporan = "total" Then
                strSQL = "SELECT [Ruang Pelayanan], (cast(dbo.UbahBulan(MONTH(TglStruk)) as varchar) + '/' + cast(year(TglStruk) as varchar))AS Alias, JenisPasien, [Komponen Tarif], SUM(Total) AS Total, SUM(JmlBayar) AS JmlBayar, SUM(Piutang)  AS Piutang, SUM(TanggunganRS) AS TanggunganRS, SUM(Pembebasan) AS Pembebasan, SUM(SisaTagihan) AS SisaTagihan " & _
                    " From  V_RekapPendapatanRSTotal " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " GROUP BY [Ruang Pelayanan], cast(dbo.UbahBulan(MONTH(TglStruk)) as varchar) + '/' + cast(year(TglStruk) as varchar), JenisPasien, [Komponen Tarif] " & _
                    " ORDER BY [Ruang Pelayanan], cast(dbo.UbahBulan(MONTH(TglStruk)) as varchar) + '/' + cast(year(TglStruk) as varchar), JenisPasien, [Komponen Tarif]"
            End If
            
        Case "03" 'per Tahun
            If strJenisLaporan = "perda" Then
                strSQL = "SELECT [Ruang Pelayanan], YEAR(TglStruk) AS Alias, JenisPasien, [Komponen Tarif], SUM(Total) AS Total, SUM(JmlBayar) AS JmlBayar, SUM(Piutang)  AS Piutang, SUM(TanggunganRS) AS TanggunganRS, SUM(Pembebasan) AS Pembebasan, SUM(SisaTagihan) AS SisaTagihan " & _
                    " From  V_RekapPendapatanRSPerKomponen " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " AND (KdKomponen IN ('01','02','03'))" & _
                    " GROUP BY [Ruang Pelayanan], YEAR(TglStruk), JenisPasien, [Komponen Tarif] " & _
                    " ORDER BY [Ruang Pelayanan], YEAR(TglStruk), JenisPasien, [Komponen Tarif]"
            ElseIf strJenisLaporan = "nonperda" Then
                strSQL = "SELECT [Ruang Pelayanan], YEAR(TglStruk) AS Alias, JenisPasien, [Komponen Tarif], SUM(Total) AS Total, SUM(JmlBayar) AS JmlBayar, SUM(Piutang)  AS Piutang, SUM(TanggunganRS) AS TanggunganRS, SUM(Pembebasan) AS Pembebasan, SUM(SisaTagihan) AS SisaTagihan " & _
                    " From  V_RekapPendapatanRSPerKomponenNonPerda " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " GROUP BY [Ruang Pelayanan], YEAR(TglStruk), JenisPasien, [Komponen Tarif] " & _
                    " ORDER BY [Ruang Pelayanan], YEAR(TglStruk), JenisPasien, [Komponen Tarif]"
            ElseIf strJenisLaporan = "total" Then
                strSQL = "SELECT [Ruang Pelayanan], YEAR(TglStruk) AS Alias, JenisPasien, [Komponen Tarif], SUM(Total) AS Total, SUM(JmlBayar) AS JmlBayar, SUM(Piutang)  AS Piutang, SUM(TanggunganRS) AS TanggunganRS, SUM(Pembebasan) AS Pembebasan, SUM(SisaTagihan) AS SisaTagihan " & _
                    " From  V_RekapPendapatanRSTotal " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " GROUP BY [Ruang Pelayanan], YEAR(TglStruk), JenisPasien, [Komponen Tarif] " & _
                    " ORDER BY [Ruang Pelayanan], YEAR(TglStruk), JenisPasien, [Komponen Tarif]"
            End If
        
        Case "04" 'per Periode
            If strJenisLaporan = "perda" Then
                strSQL = "SELECT [Ruang Pelayanan], '-' AS Alias, JenisPasien, [Komponen Tarif], SUM(Total) AS Total, SUM(JmlBayar) AS JmlBayar, SUM(Piutang)  AS Piutang, SUM(TanggunganRS) AS TanggunganRS, SUM(Pembebasan) AS Pembebasan, SUM(SisaTagihan) AS SisaTagihan " & _
                    " From  V_RekapPendapatanRSPerKomponen " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " AND (KdKomponen IN ('01','02','03'))" & _
                    " GROUP BY [Ruang Pelayanan], JenisPasien, [Komponen Tarif] " & _
                    " ORDER BY [Ruang Pelayanan], JenisPasien, [Komponen Tarif]"
            ElseIf strJenisLaporan = "nonperda" Then
                strSQL = "SELECT [Ruang Pelayanan], '-' AS Alias, JenisPasien, [Komponen Tarif], SUM(Total) AS Total, SUM(JmlBayar) AS JmlBayar, SUM(Piutang)  AS Piutang, SUM(TanggunganRS) AS TanggunganRS, SUM(Pembebasan) AS Pembebasan, SUM(SisaTagihan) AS SisaTagihan " & _
                    " From  V_RekapPendapatanRSPerKomponenNonPerda " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " GROUP BY [Ruang Pelayanan], JenisPasien, [Komponen Tarif] " & _
                    " ORDER BY [Ruang Pelayanan], JenisPasien, [Komponen Tarif]"
            ElseIf strJenisLaporan = "total" Then
                strSQL = "SELECT [Ruang Pelayanan],'-'  AS Alias, JenisPasien, [Komponen Tarif], SUM(Total) AS Total, SUM(JmlBayar) AS JmlBayar, SUM(Piutang)  AS Piutang, SUM(TanggunganRS) AS TanggunganRS, SUM(Pembebasan) AS Pembebasan, SUM(SisaTagihan) AS SisaTagihan " & _
                    " From  V_RekapPendapatanRSTotal " & _
                    " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdruangan = '" & mstrKdRuangan & "'" & _
                    " GROUP BY [Ruang Pelayanan], JenisPasien, [Komponen Tarif] " & _
                    " ORDER BY [Ruang Pelayanan], JenisPasien, [Komponen Tarif]"
            End If
        
        Case Else
            strSQL = ""
            
    End Select
    
    Exit Sub
errLoad:
    Call msubPesanError
End Sub
