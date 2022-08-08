VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmRekapLaporan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Rekapituasi"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14685
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRekapLaporan.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   14685
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   11
      Top             =   6960
      Width           =   14655
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   495
         Left            =   11520
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   13080
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
   End
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
      Height          =   5895
      Left            =   0
      TabIndex        =   2
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
         TabIndex        =   5
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
            TabIndex        =   9
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   6
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
            OLEDropMode     =   1
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   53805059
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   7
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
            CurrentDate     =   38209
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   8
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataListLib.DataCombo dcPenjamin 
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   550
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         Text            =   ""
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
         Left            =   2880
         TabIndex        =   1
         Top             =   550
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
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   4695
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   8281
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
      Begin VB.Label lblJenisLaporan 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Laporan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2880
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pasien (Cara Bayar)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2490
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
      Picture         =   "frmRekapLaporan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRekapLaporan.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRekapLaporan.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12975
   End
End
Attribute VB_Name = "frmRekapLaporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim subTotalBiaya As Currency
Dim subTotalBayar As Currency
Dim subTotalPiutang As Currency
Dim subTotalCostSharing As Currency
Dim subTotalPembebasan As Currency
Dim subTotalSisaTagihan As Currency

Private Sub cmdCetak_Click()
    frmCetakLaporanKasir.Show
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdTampilkanTemp_Click()
Dim i, j As Integer
Dim intJmlRow As Integer
Dim intRowNow As Integer
    
    On Error GoTo errTampilkan
    
    If dcPenjamin.MatchedWithList = False Then dcPenjamin.SetFocus: Exit Sub
'    If dcJenisLaporan.MatchedWithList = False Then dcJenisLaporan.SetFocus: Exit Sub
    
    subSetGridAja
    intJmlRow = 0
    fgData.Visible = False
    Select Case mstrLaporan
    Case "Penerimaan"
        strSQL = "SELECT NamaRuangan,COUNT(NoStruk) AS JmlStruk," _
        & "SUM(Biaya) AS TotalBiaya,SUM(Bayar) AS TotalBayar," _
        & "SUM(Piutang) AS TotalPiutang,SUM(CostSharing) AS TotalCostSharing," _
        & "SUM(Pembebasan) AS TotalPembebasan," _
        & "SUM(SisaTagihan) AS TotalSisaTagihan " _
        & "FROM v_S_Lap_PenerimaanKasir " _
        & "WHERE JenisPasien='" & dcPenjamin.Text _
        & "' AND TglStruk BETWEEN '" _
        & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " _
        & "AND KdRuanganKasir='" & mstrKdRuangan _
        & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "' " _
        & "GROUP BY NamaRuangan ORDER BY NamaRuangan"
        msubRecFO rsB, strSQL
        intJmlRow = rsB.RecordCount
        strSQL = "SELECT * FROM v_S_Lap_PenerimaanKasir WHERE JenisPasien='" _
        & dcPenjamin.Text & "' AND TglStruk BETWEEN '" _
        & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " _
        & "AND KdRuanganKasir='" & mstrKdRuangan _
        & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "' " _
        & "ORDER BY NamaRuangan,NamaPenjamin"
        msubRecFO rs, strSQL
        intJmlRow = intJmlRow + rs.RecordCount
        fgData.Rows = intJmlRow + 1
        intRowNow = 0
        For i = 1 To rs.RecordCount
            intRowNow = intRowNow + 1
            For j = 1 To fgData.Cols - 1
                fgData.TextMatrix(intRowNow, j) = rs(j - 1).Value
            Next j
            rs.MoveNext
            If rs.EOF = True Then GoTo stepPenerimaanSet
            If rsB(0).Value <> rs("NamaRuangan").Value Then
stepPenerimaanSet:
                intRowNow = intRowNow + 1
                fgData.TextMatrix(intRowNow, 2) = fgData.TextMatrix(intRowNow - 1, 2)
                fgData.TextMatrix(intRowNow, 3) = "Sub Total"
                fgData.TextMatrix(intRowNow, 7) = rsB("TotalBiaya").Value
                fgData.TextMatrix(intRowNow, 8) = rsB("TotalBayar").Value
                fgData.TextMatrix(intRowNow, 9) = rsB("TotalPiutang").Value
                fgData.TextMatrix(intRowNow, 10) = rsB("TotalCostSharing").Value
                fgData.TextMatrix(intRowNow, 11) = rsB("TotalPembebasan").Value
                fgData.TextMatrix(intRowNow, 12) = rsB("TotalSisaTagihan").Value
                subSetSubTotalRow intRowNow
                If rsB.EOF Then Exit Sub
                rsB.MoveNext
            End If
        Next i
    
    Case "PenerimaanKomponen"
        strSQL = "SELECT NamaRuangan,COUNT(NoStruk) AS JmlStruk," _
        & "SUM(JasaRS) AS TotalJasaRS," _
        & "SUM(JasaPelayanan) AS TotalJasaPelayanan," _
        & "SUM(JasaManajemen) AS TotalJasaManajemen," _
        & "SUM(Alkes) AS TotalAlkes,SUM(Biaya) AS TotalBiaya," _
        & "SUM(Bayar) AS TotalBayar,SUM(Piutang) AS TotalPiutang," _
        & "SUM(CostSharing) AS TotalCostSharing," _
        & "SUM(Pembebasan) AS TotalPembebasan," _
        & "SUM(SisaTagihan) AS TotalSisaTagihan " _
        & "FROM v_S_Lap_PenerimaanKasir_Komponen " _
        & "WHERE JenisPasien='" & dcPenjamin.Text _
        & "' AND TglStruk BETWEEN '" _
        & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " _
        & "AND KdRuanganKasir='" & mstrKdRuangan _
        & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "' " _
        & "GROUP BY NamaRuangan ORDER BY NamaRuangan"
        msubRecFO rsB, strSQL
        intJmlRow = rsB.RecordCount
        strSQL = "SELECT * FROM v_S_Lap_PenerimaanKasir_Komponen WHERE JenisPasien='" _
        & dcPenjamin.Text & "' AND TglStruk BETWEEN '" _
        & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " _
        & "AND KdRuanganKasir='" & mstrKdRuangan _
        & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "' " _
        & "ORDER BY NamaRuangan"
        msubRecFO rs, strSQL
        intJmlRow = intJmlRow + rs.RecordCount
        fgData.Rows = intJmlRow + 1
        intRowNow = 0
        For i = 1 To rs.RecordCount
            intRowNow = intRowNow + 1
            For j = 1 To fgData.Cols - 1
                fgData.TextMatrix(intRowNow, j) = rs(j - 1).Value
            Next j
            rs.MoveNext
            If rs.EOF = True Then GoTo stepPenerimaanKomp
            If rsB(0).Value <> rs("NamaRuangan").Value Then
stepPenerimaanKomp:
                intRowNow = intRowNow + 1
                fgData.TextMatrix(intRowNow, 2) = fgData.TextMatrix(intRowNow - 1, 2)
                fgData.TextMatrix(intRowNow, 3) = "Sub Total"
                fgData.TextMatrix(intRowNow, 8) = rsB("TotalJasaRS").Value
                fgData.TextMatrix(intRowNow, 9) = rsB("TotalJasaPelayanan").Value
                fgData.TextMatrix(intRowNow, 10) = rsB("TotalJasaManajemen").Value
                fgData.TextMatrix(intRowNow, 11) = rsB("TotalAlkes").Value
                fgData.TextMatrix(intRowNow, 12) = rsB("TotalBiaya").Value
                fgData.TextMatrix(intRowNow, 13) = rsB("TotalBayar").Value
                fgData.TextMatrix(intRowNow, 14) = rsB("TotalPiutang").Value
                fgData.TextMatrix(intRowNow, 15) = rsB("TotalCostSharing").Value
                fgData.TextMatrix(intRowNow, 16) = rsB("TotalPembebasan").Value
                fgData.TextMatrix(intRowNow, 17) = rsB("TotalSisaTagihan").Value
                subSetSubTotalRow intRowNow
                If rsB.EOF Then Exit Sub
                rsB.MoveNext
            End If
        Next i
    Case "PenerimaanDokter"
        strSQL = "SELECT DokterPemeriksa,COUNT(NoStruk) AS JmlStruk," _
        & "SUM(JasaPelayanan) AS TotalJasaPelayanan," _
        & "SUM(Biaya) AS TotalBiaya,SUM(Bayar) AS TotalBayar," _
        & "SUM(Piutang) AS TotalPiutang,SUM(CostSharing) AS TotalCostSharing," _
        & "SUM(Pembebasan) AS TotalPembebasan," _
        & "SUM(SisaTagihan) AS TotalSisaTagihan " _
        & "FROM v_S_Lap_PenerimaanKasir_Dokter " _
        & "WHERE JenisPasien='" & dcPenjamin.Text _
        & "' AND TglStruk BETWEEN '" _
        & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " _
        & "AND KdRuanganKasir='" & mstrKdRuangan _
        & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "' " _
        & "GROUP BY DokterPemeriksa ORDER BY DokterPemeriksa"
        msubRecFO rsB, strSQL
        intJmlRow = rsB.RecordCount
        strSQL = "SELECT * FROM v_S_Lap_PenerimaanKasir_Dokter WHERE JenisPasien='" _
        & dcPenjamin.Text & "' AND TglStruk BETWEEN '" _
        & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " _
        & "AND KdRuanganKasir='" & mstrKdRuangan _
        & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "'"
        msubRecFO rs, strSQL
        intJmlRow = intJmlRow + rs.RecordCount
        fgData.Rows = intJmlRow + 1
        intRowNow = 0
        For i = 1 To rs.RecordCount
            intRowNow = intRowNow + 1
            For j = 1 To fgData.Cols - 1
                fgData.TextMatrix(intRowNow, j) = rs(j - 1).Value
            Next j
            rs.MoveNext
            If rs.EOF = True Then GoTo stepPenerimaanDok
            If rsB(0).Value <> rs("DokterPemeriksa").Value Then
stepPenerimaanDok:
                intRowNow = intRowNow + 1
                fgData.TextMatrix(intRowNow, 2) = fgData.TextMatrix(intRowNow - 1, 2)
                fgData.TextMatrix(intRowNow, 3) = "Sub Total"
                fgData.TextMatrix(intRowNow, 9) = rsB("TotalJasaPelayanan").Value
                fgData.TextMatrix(intRowNow, 10) = rsB("TotalBiaya").Value
                fgData.TextMatrix(intRowNow, 11) = rsB("TotalBayar").Value
                fgData.TextMatrix(intRowNow, 12) = rsB("TotalPiutang").Value
                fgData.TextMatrix(intRowNow, 13) = rsB("TotalCostSharing").Value
                fgData.TextMatrix(intRowNow, 14) = rsB("TotalPembebasan").Value
                fgData.TextMatrix(intRowNow, 15) = rsB("TotalSisaTagihan").Value
                subSetSubTotalRow intRowNow
                If rsB.EOF Then Exit Sub
                rsB.MoveNext
            End If
        Next i
    
    Case "PenerimaanPerda"
        If dcJenisLaporan.BoundText = "03" Then 'total
            strSQL = "SELECT NamaRuangan,COUNT(NoStruk) AS JmlStruk, SUM(Biaya) AS TotalBiaya,SUM(Bayar) AS TotalBayar, SUM(Piutang) AS TotalPiutang,SUM(CostSharing) AS TotalCostSharing, SUM(Pembebasan) AS TotalPembebasan, SUM(SisaTagihan) AS TotalSisaTagihan" & _
                " FROM v_S_Lap_PenerimaanKasir " & _
                " WHERE JenisPasien='" & dcPenjamin.Text & "' AND TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuanganKasir='" & mstrKdRuangan & " ' AND IdPegawaiKasir = '" & strIDPegawaiAktif & "' " & _
                " GROUP BY NamaRuangan ORDER BY NamaRuangan"
            msubOpenRecFO rsB, strSQL, dbConn
            intJmlRow = rsB.RecordCount
            strSQL = "SELECT * " & _
                " FROM v_S_Lap_PenerimaanKasir " & _
                " WHERE JenisPasien='" & dcPenjamin.Text & "' AND TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuanganKasir='" & mstrKdRuangan & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "' " & _
                " ORDER BY NamaRuangan,NamaPenjamin"
            msubOpenRecFO rs, strSQL, dbConn
        
        ElseIf dcJenisLaporan.BoundText = "01" Then 'perda a/ tindakan
            strSQL = "SELECT NamaRuangan,COUNT(NoStruk) AS JmlStruk, SUM(Biaya) AS TotalBiaya,SUM(Bayar) AS TotalBayar, SUM(Piutang) AS TotalPiutang,SUM(CostSharing) AS TotalCostSharing, SUM(Pembebasan) AS TotalPembebasan, SUM(SisaTagihan) AS TotalSisaTagihan " & _
                " FROM V_S_Lap_PenerimaanKasir_Perda " & _
                " WHERE JenisPasien='" & dcPenjamin.Text & "' AND TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuanganKasir='" & mstrKdRuangan & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "' " & _
                " GROUP BY NamaRuangan ORDER BY NamaRuangan"
            msubRecFO rsB, strSQL
            intJmlRow = rsB.RecordCount
            strSQL = "SELECT * " & _
                " FROM V_S_Lap_PenerimaanKasir_Perda " & _
                " WHERE JenisPasien='" & dcPenjamin.Text & "' AND TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuanganKasir='" & mstrKdRuangan & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "' " & _
                " ORDER BY NamaRuangan,NamaPenjamin"
            msubRecFO rs, strSQL
        
        ElseIf dcJenisLaporan.BoundText = "02" Then 'alkes a/ non perda
            strSQL = "SELECT NamaRuangan,COUNT(NoStruk) AS JmlStruk, SUM(Biaya) AS TotalBiaya,SUM(Bayar) AS TotalBayar, SUM(Piutang) AS TotalPiutang,SUM(CostSharing) AS TotalCostSharing, SUM(Pembebasan) AS TotalPembebasan, SUM(SisaTagihan) AS TotalSisaTagihan " & _
                " FROM V_S_Lap_PenerimaanKasir_NonPerda " & _
                " WHERE JenisPasien='" & dcPenjamin.Text & "' AND TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuanganKasir='" & mstrKdRuangan & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "' " & _
                " GROUP BY NamaRuangan ORDER BY NamaRuangan"
            msubRecFO rsB, strSQL
            intJmlRow = rsB.RecordCount
            strSQL = "SELECT * " & _
                " FROM V_S_Lap_PenerimaanKasir_NonPerda " & _
                " WHERE JenisPasien='" & dcPenjamin.Text & "' AND TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuanganKasir='" & mstrKdRuangan & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "' " & _
                " ORDER BY NamaRuangan,NamaPenjamin"
            msubRecFO rs, strSQL
        End If
        
        intJmlRow = intJmlRow + rs.RecordCount
        fgData.Rows = intJmlRow + 2
        intRowNow = 0
        subTotalBiaya = 0: subTotalBayar = 0: subTotalPiutang = 0: subTotalCostSharing = 0: subTotalPembebasan = 0: subTotalSisaTagihan = 0
        For i = 1 To rs.RecordCount
            intRowNow = intRowNow + 1
            For j = 1 To fgData.Cols - 1
                fgData.TextMatrix(intRowNow, j) = rs(j - 1).Value
            Next j
            rs.MoveNext
            If rs.EOF = True Then GoTo stepPenerimaanKasirPerdaNonPerdaSet
            If rsB(0).Value <> rs("NamaRuangan").Value Then
stepPenerimaanKasirPerdaNonPerdaSet:
                intRowNow = intRowNow + 1
                fgData.TextMatrix(intRowNow, 2) = fgData.TextMatrix(intRowNow - 1, 2)
                fgData.TextMatrix(intRowNow, 3) = "Sub Total"
                fgData.TextMatrix(intRowNow, 7) = IIf(rsB("TotalBiaya").Value = 0, 0, Format(rsB("TotalBiaya").Value, "#,###"))
                fgData.TextMatrix(intRowNow, 8) = IIf(rsB("TotalBayar").Value = 0, 0, Format(rsB("TotalBayar").Value, "#,###"))
                fgData.TextMatrix(intRowNow, 9) = IIf(rsB("TotalPiutang").Value = 0, 0, Format(rsB("TotalPiutang").Value, "#,###"))
                fgData.TextMatrix(intRowNow, 10) = IIf(rsB("TotalCostSharing").Value = 0, 0, Format(rsB("TotalCostSharing").Value, "#,###"))
                fgData.TextMatrix(intRowNow, 11) = IIf(rsB("TotalPembebasan").Value = 0, 0, Format(rsB("TotalPembebasan").Value, "#,###"))
                fgData.TextMatrix(intRowNow, 12) = IIf(rsB("TotalSisaTagihan").Value = 0, 0, Format(rsB("TotalSisaTagihan").Value, "#,###"))
                
                subTotalBiaya = subTotalBiaya + rsB("TotalBiaya")
                subTotalBayar = subTotalBayar + rsB("TotalBayar")
                subTotalPiutang = subTotalPiutang + rsB("TotalPiutang")
                subTotalCostSharing = subTotalCostSharing + rsB("TotalCostSharing")
                subTotalPembebasan = subTotalPembebasan + rsB("TotalPembebasan")
                subTotalSisaTagihan = subTotalSisaTagihan + rsB("TotalSisaTagihan")
                
                subSetSubTotalRow intRowNow
                If rsB.EOF Then Exit Sub
                rsB.MoveNext
            End If
        Next i
    
        intRowNow = intRowNow + 1
        fgData.TextMatrix(intRowNow, 1) = "Total"
        fgData.TextMatrix(intRowNow, 7) = IIf(subTotalBiaya = 0, 0, Format(subTotalBiaya, "#,###"))
        fgData.TextMatrix(intRowNow, 8) = IIf(subTotalBayar = 0, 0, Format(subTotalBayar, "#,###"))
        fgData.TextMatrix(intRowNow, 9) = IIf(subTotalPiutang = 0, 0, Format(subTotalPiutang, "#,###"))
        fgData.TextMatrix(intRowNow, 10) = IIf(subTotalCostSharing = 0, 0, Format(subTotalCostSharing, "#,###"))
        fgData.TextMatrix(intRowNow, 11) = IIf(subTotalPembebasan = 0, 0, Format(subTotalPembebasan, "#,###"))
        fgData.TextMatrix(intRowNow, 12) = IIf(subTotalSisaTagihan = 0, 0, Format(subTotalSisaTagihan, "#,###"))
        subSetTotalRow intRowNow, 1, vbBlue, vbWhite
        
    Case "PenerimaanDetail"
        If dcJenisLaporan.BoundText = "03" Then 'total
            strSQL = "SELECT NamaRuangan,COUNT(NoStruk) AS JmlStruk, SUM(Biaya) AS TotalBiaya,SUM(Bayar) AS TotalBayar, SUM(Piutang) AS TotalPiutang,SUM(CostSharing) AS TotalCostSharing, SUM(Pembebasan) AS TotalPembebasan, SUM(SisaTagihan) AS TotalSisaTagihan " & _
                " FROM v_S_Lap_PenerimaanKasir_Detail " & _
                " WHERE JenisPasien='" & dcPenjamin.Text & "' AND TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuanganKasir='" & mstrKdRuangan & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "' " & _
                " GROUP BY NamaRuangan ORDER BY NamaRuangan"
            msubRecFO rsB, strSQL
            intJmlRow = rsB.RecordCount
            strSQL = "SELECT * " & _
                " FROM v_S_Lap_PenerimaanKasir_Detail " & _
                " WHERE JenisPasien='" & dcPenjamin.Text & "' AND TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuanganKasir='" & mstrKdRuangan & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "' " & _
                " ORDER BY NamaRuangan"
            msubRecFO rs, strSQL
        ElseIf dcJenisLaporan.BoundText = "01" Then 'perda
            strSQL = "SELECT NamaRuangan,COUNT(NoStruk) AS JmlStruk, SUM(Biaya) AS TotalBiaya,SUM(Bayar) AS TotalBayar, SUM(Piutang) AS TotalPiutang,SUM(CostSharing) AS TotalCostSharing, SUM(Pembebasan) AS TotalPembebasan, SUM(SisaTagihan) AS TotalSisaTagihan " & _
                " FROM V_S_Lap_PenerimaanKasir_Perda_Detail " & _
                " WHERE JenisPasien='" & dcPenjamin.Text & "' AND TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuanganKasir='" & mstrKdRuangan & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "' " & _
                " GROUP BY NamaRuangan ORDER BY NamaRuangan"
            msubRecFO rsB, strSQL
            intJmlRow = rsB.RecordCount
            strSQL = "SELECT * " & _
                " FROM V_S_Lap_PenerimaanKasir_Perda_Detail " & _
                " WHERE JenisPasien='" & dcPenjamin.Text & "' AND TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuanganKasir='" & mstrKdRuangan & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "' " & _
                " ORDER BY NamaRuangan"
            msubRecFO rs, strSQL
        ElseIf dcJenisLaporan.BoundText = "02" Then 'non perda
            strSQL = "SELECT NamaRuangan,COUNT(NoStruk) AS JmlStruk, SUM(Biaya) AS TotalBiaya,SUM(Bayar) AS TotalBayar, SUM(Piutang) AS TotalPiutang,SUM(CostSharing) AS TotalCostSharing, SUM(Pembebasan) AS TotalPembebasan, SUM(SisaTagihan) AS TotalSisaTagihan " & _
                " FROM V_S_Lap_PenerimaanKasir_NonPerda_Detail " & _
                " WHERE JenisPasien='" & dcPenjamin.Text & "' AND TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuanganKasir='" & mstrKdRuangan & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "' " & _
                " GROUP BY NamaRuangan ORDER BY NamaRuangan"
            msubRecFO rsB, strSQL
            intJmlRow = rsB.RecordCount
            strSQL = "SELECT * " & _
                " FROM V_S_Lap_PenerimaanKasir_NonPerda_Detail " & _
                " WHERE JenisPasien='" & dcPenjamin.Text & "' AND TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuanganKasir='" & mstrKdRuangan & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "' " & _
                " ORDER BY NamaRuangan"
            msubRecFO rs, strSQL
        End If
        
        intJmlRow = intJmlRow + rs.RecordCount
        fgData.Rows = intJmlRow + 2
        intRowNow = 0
        subTotalBiaya = 0: subTotalBayar = 0: subTotalPiutang = 0: subTotalCostSharing = 0: subTotalPembebasan = 0: subTotalSisaTagihan = 0
        For i = 1 To rs.RecordCount
            intRowNow = intRowNow + 1
            For j = 1 To fgData.Cols - 1
                fgData.TextMatrix(intRowNow, j) = rs(j - 1).Value
            Next j
            rs.MoveNext
            If rs.EOF = True Then GoTo stepPenerimaanDet
            If rsB(0).Value <> rs("NamaRuangan").Value Then
stepPenerimaanDet:
                intRowNow = intRowNow + 1
                fgData.TextMatrix(intRowNow, 2) = fgData.TextMatrix(intRowNow - 1, 2)
                fgData.TextMatrix(intRowNow, 3) = "Sub Total"
                fgData.TextMatrix(intRowNow, 8) = rsB("TotalBiaya").Value
                fgData.TextMatrix(intRowNow, 9) = rsB("TotalBayar").Value
                fgData.TextMatrix(intRowNow, 10) = rsB("TotalPiutang").Value
                fgData.TextMatrix(intRowNow, 11) = rsB("TotalCostSharing").Value
                fgData.TextMatrix(intRowNow, 12) = rsB("TotalPembebasan").Value
                fgData.TextMatrix(intRowNow, 13) = rsB("TotalSisaTagihan").Value
                
                subTotalBiaya = subTotalBiaya + rsB("TotalBiaya")
                subTotalBayar = subTotalBayar + rsB("TotalBayar")
                subTotalPiutang = subTotalPiutang + rsB("TotalPiutang")
                subTotalCostSharing = subTotalCostSharing + rsB("TotalCostSharing")
                subTotalPembebasan = subTotalPembebasan + rsB("TotalPembebasan")
                subTotalSisaTagihan = subTotalSisaTagihan + rsB("TotalSisaTagihan")
                
                subSetSubTotalRow intRowNow
                If rsB.EOF Then Exit Sub
                rsB.MoveNext
            End If
        Next i
    
        intRowNow = intRowNow + 1
        fgData.TextMatrix(intRowNow, 1) = "Total"
        fgData.TextMatrix(intRowNow, 8) = IIf(subTotalBiaya = 0, 0, Format(subTotalBiaya, "#,###"))
        fgData.TextMatrix(intRowNow, 9) = IIf(subTotalBayar = 0, 0, Format(subTotalBayar, "#,###"))
        fgData.TextMatrix(intRowNow, 10) = IIf(subTotalPiutang = 0, 0, Format(subTotalPiutang, "#,###"))
        fgData.TextMatrix(intRowNow, 11) = IIf(subTotalCostSharing = 0, 0, Format(subTotalCostSharing, "#,###"))
        fgData.TextMatrix(intRowNow, 12) = IIf(subTotalPembebasan = 0, 0, Format(subTotalPembebasan, "#,###"))
        fgData.TextMatrix(intRowNow, 13) = IIf(subTotalSisaTagihan = 0, 0, Format(subTotalSisaTagihan, "#,###"))
        subSetTotalRow intRowNow, 1, vbBlue, vbWhite
        
    Case "PembatalanStruk"
        strSQL = "SELECT NamaRuangan,COUNT(NoStruk) AS JmlStruk, SUM(Biaya) AS TotalBiaya,SUM(Bayar) AS TotalBayar, SUM(Piutang) AS TotalPiutang,SUM(CostSharing) AS TotalCostSharing, SUM(Pembebasan) AS TotalPembebasan, SUM(SisaTagihan) AS TotalSisaTagihan" & _
            " FROM V_LaporanReturStrukPelayananRS " & _
            " WHERE JenisPasien='" & dcPenjamin.Text & "' AND TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuanganKasir='" & mstrKdRuangan & " ' AND IdPegawaiKasir = '" & strIDPegawaiAktif & "' " & _
            " GROUP BY NamaRuangan ORDER BY NamaRuangan"
        msubOpenRecFO rsB, strSQL, dbConn
        intJmlRow = rsB.RecordCount
        strSQL = "SELECT * " & _
            " FROM V_LaporanReturStrukPelayananRS " & _
            " WHERE JenisPasien='" & dcPenjamin.Text & "' AND TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuanganKasir='" & mstrKdRuangan & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "' " & _
            " ORDER BY NamaRuangan,NamaPenjamin"
        msubOpenRecFO rs, strSQL, dbConn
        
        intJmlRow = intJmlRow + rs.RecordCount
        fgData.Rows = intJmlRow + 2
        intRowNow = 0
        subTotalBiaya = 0: subTotalBayar = 0: subTotalPiutang = 0: subTotalPembebasan = 0: subTotalSisaTagihan = 0
        For i = 1 To rs.RecordCount
            intRowNow = intRowNow + 1
            For j = 1 To fgData.Cols - 1
                fgData.TextMatrix(intRowNow, j) = rs(j - 1).Value
            Next j
            rs.MoveNext
            If rs.EOF = True Then GoTo stepPembatalanStruk
            If rsB(0).Value <> rs("NamaRuangan").Value Then
stepPembatalanStruk:
                intRowNow = intRowNow + 1
                fgData.TextMatrix(intRowNow, 2) = fgData.TextMatrix(intRowNow - 1, 2)
                fgData.TextMatrix(intRowNow, 3) = "Sub Total"
                fgData.TextMatrix(intRowNow, 7) = IIf(rsB("TotalBiaya").Value = 0, 0, Format(rsB("TotalBiaya").Value, "#,###"))
                fgData.TextMatrix(intRowNow, 8) = IIf(rsB("TotalBayar").Value = 0, 0, Format(rsB("TotalBayar").Value, "#,###"))
                fgData.TextMatrix(intRowNow, 9) = IIf(rsB("TotalPiutang").Value = 0, 0, Format(rsB("TotalPiutang").Value, "#,###"))
                fgData.TextMatrix(intRowNow, 10) = IIf(rsB("TotalCostSharing").Value = 0, 0, Format(rsB("TotalCostSharing").Value, "#,###"))
                fgData.TextMatrix(intRowNow, 11) = IIf(rsB("TotalPembebasan").Value = 0, 0, Format(rsB("TotalPembebasan").Value, "#,###"))
                fgData.TextMatrix(intRowNow, 12) = IIf(rsB("TotalSisaTagihan").Value = 0, 0, Format(rsB("TotalSisaTagihan").Value, "#,###"))
                
                subTotalBiaya = subTotalBiaya + rsB("TotalBiaya")
                subTotalBayar = subTotalBayar + rsB("TotalBayar")
                subTotalPiutang = subTotalPiutang + rsB("TotalPiutang")
                subTotalCostSharing = subTotalCostSharing + rsB("TotalCostSharing")
                subTotalPembebasan = subTotalPembebasan + rsB("TotalPembebasan")
                subTotalSisaTagihan = subTotalSisaTagihan + rsB("TotalSisaTagihan")
                
                subSetSubTotalRow intRowNow
                If rsB.EOF Then Exit Sub
                rsB.MoveNext
            End If
        Next i
    
        intRowNow = intRowNow + 1
        fgData.TextMatrix(intRowNow, 1) = "Total"
        fgData.TextMatrix(intRowNow, 7) = IIf(subTotalBiaya = 0, 0, Format(subTotalBiaya, "#,###"))
        fgData.TextMatrix(intRowNow, 8) = IIf(subTotalBayar = 0, 0, Format(subTotalBayar, "#,###"))
        fgData.TextMatrix(intRowNow, 9) = IIf(subTotalPiutang = 0, 0, Format(subTotalPiutang, "#,###"))
        fgData.TextMatrix(intRowNow, 10) = IIf(subTotalCostSharing = 0, 0, Format(subTotalCostSharing, "#,###"))
        fgData.TextMatrix(intRowNow, 11) = IIf(subTotalPembebasan = 0, 0, Format(subTotalPembebasan, "#,###"))
        fgData.TextMatrix(intRowNow, 12) = IIf(subTotalSisaTagihan = 0, 0, Format(subTotalSisaTagihan, "#,###"))
        subSetTotalRow intRowNow, 1, vbBlue, vbWhite
        
    End Select
    fgData.Visible = True
    Exit Sub
errTampilkan:
    msubPesanError
End Sub

Private Sub subDcSource()
    strSQL = "SELECT KdKelompokPasien,JenisPasien FROM KelompokPasien"
    msubRecFO rs, strSQL
    Set dcPenjamin.RowSource = rs
    dcPenjamin.BoundColumn = rs(0).Name
    dcPenjamin.ListField = rs(1).Name
    If Not rs.EOF Then dcPenjamin.Text = rs(1).Value

'    strSQL = "SELECT KdJenisLaporan,JenisLaporan FROM JenisLaporanKasir"
'    msubRecFO rs, strSQL
'    Set dcJenisLaporan.RowSource = rs
'    dcJenisLaporan.BoundColumn = rs(0).Name
'    dcJenisLaporan.ListField = rs(1).Name
'    If Not rs.EOF Then dcJenisLaporan.Text = rs(1).Value
End Sub

Private Sub dcJenisLaporan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdTampilkanTemp.SetFocus
End Sub

Private Sub dcPenjamin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If dcJenisLaporan.Visible = True Then dcJenisLaporan.SetFocus Else cmdTampilkanTemp.SetFocus
    End If
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcPenjamin.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    On Error GoTo errFormLoad
    Select Case mstrLaporan
    Case "PenerimaanPerda", "PenerimaanDetail"
        lblJenisLaporan.Visible = True
        dcJenisLaporan.Visible = True
    Case Else
        lblJenisLaporan.Visible = False
        dcJenisLaporan.Visible = False
    End Select
    
    Call subDcSource
    
    Select Case mstrLaporan
    Case "Penerimaan"
        Me.Caption = "Medifirst2000 - Laporan Penerimaan Kasir"
        strSQL = "SELECT * FROM v_S_Lap_PenerimaanKasir WHERE 1=2"
    Case "PenerimaanDetail"
        Me.Caption = "Medifirst2000 - Laporan Detail Penerimaan Kasir"
        strSQL = "SELECT * FROM v_S_Lap_PenerimaanKasir_Detail WHERE 1=2"
    Case "PenerimaanKomponen"
        Me.Caption = "Medifirst2000 - Laporan Penerimaan Kasir Per Komponen Jasa"
        strSQL = "SELECT * FROM v_S_Lap_PenerimaanKasir_Komponen WHERE 1=2"
    Case "PenerimaanDokter"
        Me.Caption = "Medifirst2000 - Laporan Penerimaan Kasir Per Dokter"
        strSQL = "SELECT * FROM v_S_Lap_PenerimaanKasir_Dokter WHERE 1=2"
    End Select
    msubRecFO rs, strSQL
    subSetGridAja
    dtpAwal.Value = Now
    dtpAkhir.Value = Now
    Call PlayFlashMovie(Me)
    Exit Sub
errFormLoad:
    msubPesanError
End Sub

Private Sub subSetGrid()
Dim i, j As Integer
    Select Case mstrLaporan
    Case "Penerimaan"
        With fgData
            .Clear
            .Cols = 13
            .Rows = rs.RecordCount + 1
            .ColWidth(0) = 150
            .ColWidth(1) = 0
            .ColWidth(2) = 1900
            .ColWidth(3) = 1400
            .ColWidth(4) = 1150
            .ColWidth(5) = 2000
            .ColWidth(6) = 1050
            .ColWidth(7) = 1000
            .ColWidth(8) = 1000
            .ColWidth(9) = 1000
            .ColWidth(10) = 1150
            .ColWidth(11) = 1150
            .ColWidth(12) = 1150
            .ColAlignment(3) = flexAlignLeftCenter
            .ColAlignment(4) = flexAlignCenterCenter
            .Row = 0
            For i = 2 To 12
                .Col = i
                .CellAlignment = flexAlignCenterCenter
                .CellFontBold = True
            Next i
            .TextMatrix(0, 1) = "NamaPenjamin"
            .TextMatrix(0, 2) = "Ruang Pemeriksaan"
            .TextMatrix(0, 3) = "Tanggal Struk"
            .TextMatrix(0, 4) = "No. Register"
            .TextMatrix(0, 5) = "Nama Pasien"
            .TextMatrix(0, 6) = "No. Struk"
            .TextMatrix(0, 7) = "Biaya"
            .TextMatrix(0, 8) = "Bayar"
            .TextMatrix(0, 9) = "Piutang"
            .TextMatrix(0, 10) = "Tngngan RS"
            .TextMatrix(0, 11) = "Pembebasan"
            .TextMatrix(0, 12) = "Sisa Tagihan"
            For i = 1 To rs.RecordCount
                For j = 1 To 12
                    .TextMatrix(i, j) = rs(j - 1).Value
                Next j
                rs.MoveNext
            Next i
            .MergeCells = 1
            .MergeCol(2) = True
            .MergeCol(3) = True
            .MergeCol(4) = True
        End With
    Case "PenerimaanDetail"
        With fgData
            .Clear
            .Cols = 14
            .Rows = rs.RecordCount + 1
            .ColWidth(0) = 150
            .ColWidth(1) = 0
            .ColWidth(2) = 1900
            .ColWidth(3) = 1400
            .ColWidth(4) = 1150
            .ColWidth(5) = 2000
            .ColWidth(6) = 1050
            .ColWidth(7) = 2500
            .ColWidth(8) = 1000
            .ColWidth(9) = 1000
            .ColWidth(10) = 1000
            .ColWidth(11) = 1200
            .ColWidth(12) = 1200
            .ColWidth(13) = 1200
            .ColAlignment(3) = flexAlignLeftCenter
            .ColAlignment(4) = flexAlignCenterCenter
            .Row = 0
            For i = 2 To 13
                .Col = i
                .CellAlignment = flexAlignCenterCenter
                .CellFontBold = True
            Next i
            .TextMatrix(0, 1) = "NamaPenjamin"
            .TextMatrix(0, 2) = "Ruang Pemeriksaan"
            .TextMatrix(0, 3) = "Tanggal Struk"
            .TextMatrix(0, 4) = "No. Register"
            .TextMatrix(0, 5) = "Nama Pasien"
            .TextMatrix(0, 6) = "No. Struk"
            .TextMatrix(0, 7) = "Nama Pemeriksaan"
            .TextMatrix(0, 8) = "Biaya"
            .TextMatrix(0, 9) = "Bayar"
            .TextMatrix(0, 10) = "Piutang"
            .TextMatrix(0, 11) = "Tngngan RS"
            .TextMatrix(0, 12) = "Pembebasan"
            .TextMatrix(0, 13) = "Sisa Tagihan"
            For i = 1 To rs.RecordCount
                For j = 1 To 13
                    .TextMatrix(i, j) = rs(j - 1).Value
                Next j
                rs.MoveNext
            Next i
            .MergeCells = 1
            .MergeCol(2) = True
            .MergeCol(3) = True
            .MergeCol(4) = True
            .MergeCol(5) = True
            .MergeCol(6) = True
            .MergeCol(7) = True
        End With
    Case "PenerimaanKomponen"
        With fgData
            .Clear
            .Cols = 18
            .Rows = rs.RecordCount + 1
            .ColWidth(0) = 150
            .ColWidth(1) = 0
            .ColWidth(2) = 1900
            .ColWidth(3) = 1400
            .ColWidth(4) = 1150
            .ColWidth(5) = 2000
            .ColWidth(6) = 1050
            .ColWidth(7) = 2500
            .ColWidth(8) = 1000
            .ColWidth(9) = 1500
            .ColWidth(10) = 1500
            .ColWidth(11) = 1300
            .ColWidth(12) = 1000
            .ColWidth(13) = 1000
            .ColWidth(14) = 1000
            .ColWidth(15) = 1200
            .ColWidth(16) = 1200
            .ColWidth(17) = 1200
            .ColAlignment(3) = flexAlignLeftCenter
            .ColAlignment(4) = flexAlignCenterCenter
            .Row = 0
            For i = 2 To 17
                .Col = i
                .CellAlignment = flexAlignCenterCenter
                .CellFontBold = True
            Next i
            .TextMatrix(0, 1) = "NamaPenjamin"
            .TextMatrix(0, 2) = "Ruang Pemeriksaan"
            .TextMatrix(0, 3) = "Tanggal Struk"
            .TextMatrix(0, 4) = "No. Register"
            .TextMatrix(0, 5) = "Nama Pasien"
            .TextMatrix(0, 6) = "No. Struk"
            .TextMatrix(0, 7) = "Nama Pemeriksaan"
            .TextMatrix(0, 8) = "Jasa RS"
            .TextMatrix(0, 9) = "Jasa Pelayanan"
            .TextMatrix(0, 10) = "Jasa Manajemen"
            .TextMatrix(0, 11) = "Obat & Alkes"
            .TextMatrix(0, 12) = "Biaya"
            .TextMatrix(0, 13) = "Bayar"
            .TextMatrix(0, 14) = "Piutang"
            .TextMatrix(0, 15) = "Tngngan RS"
            .TextMatrix(0, 16) = "Pembebasan"
            .TextMatrix(0, 17) = "Sisa Tagihan"
            For i = 1 To rs.RecordCount
                For j = 1 To 17
                    .TextMatrix(i, j) = rs(j - 1).Value
                Next j
                rs.MoveNext
            Next i
            .MergeCells = 1
            .MergeCol(2) = True
            .MergeCol(3) = True
            .MergeCol(4) = True
            .MergeCol(5) = True
            .MergeCol(6) = True
        End With
    Case "PenerimaanDokter"
        With fgData
            .Clear
            .Cols = 16
            .Rows = rs.RecordCount + 1
            .ColWidth(0) = 150
            .ColWidth(1) = 0
            .ColWidth(2) = 2000
            .ColWidth(3) = 1400
            .ColWidth(4) = 1150
            .ColWidth(5) = 2000
            .ColWidth(6) = 1050
            .ColWidth(7) = 2500
            .ColWidth(8) = 1900
            .ColWidth(9) = 1500
            .ColWidth(10) = 0
            .ColWidth(11) = 1000
            .ColWidth(12) = 1000
            .ColWidth(13) = 1200
            .ColWidth(14) = 1200
            .ColWidth(15) = 1200
            .ColAlignment(3) = flexAlignLeftCenter
            .ColAlignment(4) = flexAlignCenterCenter
            .Row = 0
            For i = 2 To 15
                .Col = i
                .CellAlignment = flexAlignCenterCenter
                .CellFontBold = True
            Next i
            .TextMatrix(0, 1) = "NamaPenjamin"
            .TextMatrix(0, 2) = "Dokter Pemeriksa"
            .TextMatrix(0, 3) = "Tanggal Struk"
            .TextMatrix(0, 4) = "No. Register"
            .TextMatrix(0, 5) = "Nama Pasien"
            .TextMatrix(0, 6) = "No. Struk"
            .TextMatrix(0, 7) = "Nama Pemeriksaan"
            .TextMatrix(0, 8) = "Ruang Pemeriksaan"
            .TextMatrix(0, 9) = "Jasa Pelayanan"
            .TextMatrix(0, 10) = "Biaya"
            .TextMatrix(0, 11) = "Bayar"
            .TextMatrix(0, 12) = "Piutang"
            .TextMatrix(0, 13) = "Tngngan RS"
            .TextMatrix(0, 14) = "Pembebasan"
            .TextMatrix(0, 15) = "Sisa Tagihan"
            For i = 1 To rs.RecordCount
                For j = 1 To 15
                    .TextMatrix(i, j) = rs(j - 1).Value
                Next j
                rs.MoveNext
            Next i
            .MergeCells = 1
            .MergeCol(2) = True
            .MergeCol(3) = True
            .MergeCol(4) = True
            .MergeCol(5) = True
            .MergeCol(6) = True
            .MergeCol(8) = True
        End With
    End Select
End Sub

'untuk setting grid tanpa loading data
Private Sub subSetGridAja()
Dim i, j As Integer
    Select Case mstrLaporan
    Case "Penerimaan"
        With fgData
            .Clear
            .Cols = 13
            .Rows = 2
            .ColWidth(0) = 150
            .ColWidth(1) = 2000
            .ColWidth(2) = 1900
            .ColWidth(3) = 1400
            .ColWidth(4) = 1150
            .ColWidth(5) = 2000
            .ColWidth(6) = 1050
            .ColWidth(7) = 1000
            .ColWidth(8) = 1000
            .ColWidth(9) = 1000
            .ColWidth(10) = 1150
            .ColWidth(11) = 1150
            .ColWidth(12) = 1150
            .ColAlignment(3) = flexAlignLeftCenter
            .ColAlignment(4) = flexAlignCenterCenter
            .Row = 0
            For i = 1 To 12
                .Col = i
                .CellAlignment = flexAlignCenterCenter
                .CellFontBold = True
            Next i
            .TextMatrix(0, 1) = "NamaPenjamin"
            .TextMatrix(0, 2) = "Ruang Pemeriksaan"
            .TextMatrix(0, 3) = "Tanggal Struk"
            .TextMatrix(0, 4) = "No. Register"
            .TextMatrix(0, 5) = "Nama Pasien"
            .TextMatrix(0, 6) = "No. Struk"
            .TextMatrix(0, 7) = "Biaya"
            .TextMatrix(0, 8) = "Bayar"
            .TextMatrix(0, 9) = "Piutang"
            .TextMatrix(0, 10) = "Tngngan RS"
            .TextMatrix(0, 11) = "Pembebasan"
            .TextMatrix(0, 12) = "Sisa Tagihan"
            .MergeCells = 1
            .MergeCol(1) = True
            .MergeCol(2) = True
            .MergeCol(3) = True
        End With
    Case "PenerimaanDetail"
        With fgData
            .Clear
            .Cols = 14
            .Rows = 2
            .ColWidth(0) = 150
            .ColWidth(1) = 2000
            .ColWidth(2) = 1900
            .ColWidth(3) = 1400
            .ColWidth(4) = 1150
            .ColWidth(5) = 2000
            .ColWidth(6) = 1050
            .ColWidth(7) = 2500
            .ColWidth(8) = 1000
            .ColWidth(9) = 1000
            .ColWidth(10) = 1000
            .ColWidth(11) = 1200
            .ColWidth(12) = 1200
            .ColWidth(13) = 1200
            .ColAlignment(3) = flexAlignLeftCenter
            .ColAlignment(4) = flexAlignCenterCenter
            .Row = 0
            For i = 1 To 13
                .Col = i
                .CellAlignment = flexAlignCenterCenter
                .CellFontBold = True
            Next i
            .TextMatrix(0, 1) = "NamaPenjamin"
            .TextMatrix(0, 2) = "Ruang Pemeriksaan"
            .TextMatrix(0, 3) = "Tanggal Struk"
            .TextMatrix(0, 4) = "No. Register"
            .TextMatrix(0, 5) = "Nama Pasien"
            .TextMatrix(0, 6) = "No. Struk"
            .TextMatrix(0, 7) = "Nama Pemeriksaan"
            .TextMatrix(0, 8) = "Biaya"
            .TextMatrix(0, 9) = "Bayar"
            .TextMatrix(0, 10) = "Piutang"
            .TextMatrix(0, 11) = "Tnggngan RS"
            .TextMatrix(0, 12) = "Pembebasan"
            .TextMatrix(0, 13) = "Sisa Tagihan"
            .MergeCells = 1
            .MergeCol(1) = True
            .MergeCol(2) = True
            .MergeCol(3) = True
            .MergeCol(4) = True
            .MergeCol(5) = True
            .MergeCol(6) = True
        End With
    Case "PenerimaanKomponen"
        With fgData
            .Clear
            .Cols = 18
            .Rows = 2
            .ColWidth(0) = 150
            .ColWidth(1) = 0
            .ColWidth(2) = 1900
            .ColWidth(3) = 1400
            .ColWidth(4) = 1150
            .ColWidth(5) = 2000
            .ColWidth(6) = 1050
            .ColWidth(7) = 2500
            .ColWidth(8) = 1000
            .ColWidth(9) = 1500
            .ColWidth(10) = 1500
            .ColWidth(11) = 1300
            .ColWidth(12) = 1000
            .ColWidth(13) = 1000
            .ColWidth(14) = 1000
            .ColWidth(15) = 1200
            .ColWidth(16) = 1200
            .ColWidth(17) = 1200
            .ColAlignment(3) = flexAlignLeftCenter
            .ColAlignment(4) = flexAlignCenterCenter
            .Row = 0
            For i = 2 To 17
                .Col = i
                .CellAlignment = flexAlignCenterCenter
                .CellFontBold = True
            Next i
            .TextMatrix(0, 1) = "NamaPenjamin"
            .TextMatrix(0, 2) = "Ruang Pemeriksaan"
            .TextMatrix(0, 3) = "Tanggal Struk"
            .TextMatrix(0, 4) = "No. Register"
            .TextMatrix(0, 5) = "Nama Pasien"
            .TextMatrix(0, 6) = "No. Struk"
            .TextMatrix(0, 7) = "Nama Pemeriksaan"
            .TextMatrix(0, 8) = "Jasa RS"
            .TextMatrix(0, 9) = "Jasa Pelayanan"
            .TextMatrix(0, 10) = "Jasa Manajemen"
            .TextMatrix(0, 11) = "Obat & Alkes"
            .TextMatrix(0, 12) = "Biaya"
            .TextMatrix(0, 13) = "Bayar"
            .TextMatrix(0, 14) = "Piutang"
            .TextMatrix(0, 15) = "Tngngan RS"
            .TextMatrix(0, 16) = "Pembebasan"
            .TextMatrix(0, 17) = "Sisa Tagihan"
            .MergeCells = 1
            .MergeCol(1) = True
            .MergeCol(2) = True
            .MergeCol(3) = True
            .MergeCol(4) = True
            .MergeCol(5) = True
            .MergeCol(6) = True
        End With
    Case "PenerimaanDokter"
        With fgData
            .Clear
            .Cols = 16
            .Rows = 2
            .ColWidth(0) = 150
            .ColWidth(1) = 0
            .ColWidth(2) = 2000
            .ColWidth(3) = 1400
            .ColWidth(4) = 1150
            .ColWidth(5) = 2000
            .ColWidth(6) = 1050
            .ColWidth(7) = 2500
            .ColWidth(8) = 1900
            .ColWidth(9) = 1500
            .ColWidth(10) = 0
            .ColWidth(11) = 1000
            .ColWidth(12) = 1000
            .ColWidth(13) = 1200
            .ColWidth(14) = 1200
            .ColWidth(15) = 1200
            .ColAlignment(3) = flexAlignLeftCenter
            .ColAlignment(4) = flexAlignCenterCenter
            .Row = 0
            For i = 2 To 15
                .Col = i
                .CellAlignment = flexAlignCenterCenter
                .CellFontBold = True
            Next i
            .TextMatrix(0, 1) = "NamaPenjamin"
            .TextMatrix(0, 2) = "Dokter Pemeriksa"
            .TextMatrix(0, 3) = "Tanggal Struk"
            .TextMatrix(0, 4) = "No. Register"
            .TextMatrix(0, 5) = "Nama Pasien"
            .TextMatrix(0, 6) = "No. Struk"
            .TextMatrix(0, 7) = "Nama Pemeriksaan"
            .TextMatrix(0, 8) = "Ruang Pemeriksaan"
            .TextMatrix(0, 9) = "Jasa Pelayanan"
            .TextMatrix(0, 10) = ""
            .TextMatrix(0, 11) = "Bayar"
            .TextMatrix(0, 12) = "Piutang"
            .TextMatrix(0, 13) = "Tngngan RS"
            .TextMatrix(0, 14) = "Pembebasan"
            .TextMatrix(0, 15) = "Sisa Tagihan"
            .MergeCells = 1
            .MergeCol(1) = True
            .MergeCol(2) = True
            .MergeCol(3) = True
            .MergeCol(4) = True
            .MergeCol(5) = True
            .MergeCol(6) = True
            .MergeCol(8) = True
        End With
    Case "PenerimaanPerda", "PembatalanStruk"
        With fgData
            .Clear
            .Cols = 13
            .Rows = 2
            .ColWidth(0) = 150
            .ColWidth(1) = 1000
            .ColWidth(2) = 1500
            .ColWidth(3) = 1400
            .ColWidth(4) = 1150
            .ColWidth(5) = 2000
            .ColWidth(6) = 1050
            .ColWidth(7) = 1000
            .ColWidth(8) = 1000
            .ColWidth(9) = 1000
            .ColWidth(10) = 1150
            .ColWidth(11) = 1150
            .ColWidth(12) = 1150
            .ColAlignment(3) = flexAlignLeftCenter
            .ColAlignment(4) = flexAlignCenterCenter
            .Row = 0
            For i = 1 To 12
                .Col = i
                .CellAlignment = flexAlignCenterCenter
                .CellFontBold = True
            Next i
            .TextMatrix(0, 1) = "Penjamin"
            .TextMatrix(0, 2) = "Ruangan"
            .TextMatrix(0, 3) = "Tanggal Struk"
            .TextMatrix(0, 4) = "No. Register"
            .TextMatrix(0, 5) = "Nama Pasien"
            .TextMatrix(0, 6) = "No. Struk"
            .TextMatrix(0, 7) = "Biaya"
            .TextMatrix(0, 8) = "Bayar"
            .TextMatrix(0, 9) = "Piutang"
            .TextMatrix(0, 10) = "Tngngan RS"
            .TextMatrix(0, 11) = "Pembebasan"
            .TextMatrix(0, 12) = "Sisa Tagihan"
            .MergeCells = 1
            .MergeCol(1) = True
            .MergeCol(2) = True
            .MergeCol(3) = True
        End With
    End Select
End Sub

Private Sub subSetSubTotalRow(iRowNow As Integer)
Dim i As Integer
    With fgData
        'tampilan Black & White
        For i = 1 To .Cols - 1
            .Col = i
            .Row = iRowNow
            .CellBackColor = vbBlackness
            .CellForeColor = vbWhite
            If .Col = 1 Then
                .TextMatrix(.Row, 1) = .TextMatrix(.Row - 1, 1)
                .CellBackColor = vbWhite
                .CellForeColor = vbBlack
            End If
'            .RowHeight(.Row) = 300
            .CellFontBold = True
        Next
    End With
End Sub

Private Sub subSetTotalRow(iRowNow As Integer, iColBegin As Integer, vbBackColor, vbForeColor)
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
