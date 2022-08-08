VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInfoJasaPelDktr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Informasi Jasa Pelayanan Dokter"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInfoJasaPelDktr.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   13815
   Begin VB.Frame fraButton 
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   8040
      Width           =   13815
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   12000
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   10200
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraRiwayat 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   0
      TabIndex        =   6
      Top             =   3720
      Width           =   13815
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   3975
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   7011
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
   End
   Begin VB.Frame fraPeriode 
      Caption         =   "Periode Pelayanan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   13815
      Begin VB.TextBox txtDokter 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6600
         TabIndex        =   10
         Top             =   622
         Width           =   3255
      End
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
         Left            =   9960
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   375
         Left            =   1680
         TabIndex        =   0
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   48103427
         CurrentDate     =   38212
      End
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   375
         Left            =   4320
         TabIndex        =   1
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   48103427
         CurrentDate     =   38212
      End
      Begin MSDataGridLib.DataGrid dgDokter 
         Height          =   1455
         Left            =   1680
         TabIndex        =   12
         Top             =   1080
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   2566
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan Nama Dokter"
         Height          =   210
         Left            =   6600
         TabIndex        =   11
         Top             =   360
         Width           =   1905
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "s/d"
         Height          =   210
         Left            =   3960
         TabIndex        =   5
         Top             =   675
         Width           =   255
      End
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   4320
      Picture         =   "frmInfoJasaPelDktr.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   0
      Picture         =   "frmInfoJasaPelDktr.frx":431A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmInfoJasaPelDktr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strKdDokter As String
Dim intJmlDokter As Integer
Dim rsA As New ADODB.recordset
Dim rsB As New ADODB.recordset
Dim rsC As New ADODB.recordset
Public strNamaDokter As String

Private Sub cmdCari_Click()
    subLoadDataJasaPel "WHERE [Dokter Pemeriksa]='" & dgDokter.Columns(1).Value & "' AND TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "'" ' AND [Ruang Pelayanan]='" & mstrNamaRuangan & "'"
    strNamaDokter = dgDokter.Columns(1).Value
End Sub

Private Sub cmdCetak_Click()
    If strNamaDokter = "" Then Exit Sub
    Set frmCetakLaporanJPD = Nothing
    frmCetakLaporanJPD.subLoadNmDokter strNamaDokter
    frmCetakLaporanJPD.Show
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgDokter_DblClick()
'    Call dgDokter_KeyPress(13)
    Call cmdCari_Click
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlDokter = 0 Then Exit Sub
        Call cmdCari_Click
'        txtDokter.Text = dgDokter.Columns(1).Value
'        strKdDokter = dgDokter.Columns(0).Value
'        If strKdDokter = "" Then
'            MsgBox "Pilih dulu Dokter yang akan menangani Pasien", vbCritical, "Validasi"
'            txtDokter.Text = ""
'            dgDokter.SetFocus
'            Exit Sub
'        End If
'        subLoadDataJasaPel "WHERE [Dokter Pemeriksa]='" & txtDokter.Text & "' AND TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "'" ' AND [Ruang Pelayanan]='" & mstrNamaRuangan & "'"
    End If
End Sub

Private Sub Form_Load()
    centerForm Me, MDIUtama
    dtpAwal.Value = Date
    dtpAkhir.Value = Date
    subLoadDokter
    subSetGrid
End Sub

Private Sub txtDokter_Change()
    strKdDokter = ""
    subLoadDokter "WHERE NamaDokter like '%" & txtDokter.Text & "%'"
End Sub

Private Sub txtDokter_GotFocus()
    txtDokter.SelStart = 0
    txtDokter.SelLength = Len(txtDokter.Text)
    subLoadDokter "WHERE NamaDokter like '%" & txtDokter.Text & "%'"
End Sub

Private Sub txtDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlDokter = 0 Then Exit Sub
        dgDokter.SetFocus
    End If
End Sub

'untuk meload data dokter di grid
Private Sub subLoadDokter(Optional strFilter As String)
    On Error GoTo errLoad
    strSQL = "SELECT KodeDokter AS [Kode Dokter],NamaDokter AS [Nama Dokter],JK,Jabatan FROM V_DaftarDokter " & strFilter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlDokter = rs.RecordCount
    With dgDokter
        Set .DataSource = rs
        .Columns(0).Width = 1200
        .Columns(1).Width = 3000
        .Columns(2).Width = 400
        .Columns(3).Width = 3000
    End With
    Exit Sub
errLoad:
    msubPesanError
End Sub

'untuk meload jasa pelayanan dokter di grid
Private Sub subLoadDataJasaPel(strFilter As String)
Dim intJmlRow As Integer
Dim intNo As Integer
Dim curTotalBiaya As Currency
Dim curJmlBayar As Currency
Dim curPiutang As Currency
Dim curPembebasan As Currency
Dim curSisaTagihan As Currency
Dim intRowNow As Integer
    On Error GoTo errLoad
    intNo = 1
    curTotalBiaya = 0
    curJmlBayar = 0
    curPiutang = 0
    curPembebasan = 0
    curSisaTagihan = 0
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value
    fgData.Visible = False
    fgData.Clear
    subSetGrid
        
    'data yang group JenisPasien
    strSQL = "SELECT JenisPasien,SUM(TotalBiaya) AS TotalBiaya,SUM(JmlBayar) AS JmlBayar,SUM(Piutang) AS Piutang,SUM(Pembebasan) AS Pembebasan,SUM(SisaTagihan) AS SisaTagihan " _
        & "FROM V_LaporanJasaPelayananDokter " _
        & strFilter & " " _
        & "GROUP BY JenisPasien ORDER BY JenisPasien"
    msubRecFO rsA, strSQL
    intJmlRow = rsA.RecordCount
    
    'data yang group JenisPasien, Penjamin
    strSQL = "SELECT JenisPasien,Penjamin,SUM(TotalBiaya) AS TotalBiaya,SUM(JmlBayar) AS JmlBayar,SUM(Piutang) AS Piutang,SUM(Pembebasan) AS Pembebasan,SUM(SisaTagihan) AS SisaTagihan " _
        & "FROM V_LaporanJasaPelayananDokter " _
        & strFilter & " " _
        & "GROUP BY JenisPasien,Penjamin ORDER BY JenisPasien"
    msubRecFO rsB, strSQL
    intJmlRow = intJmlRow + rsB.RecordCount
    
    'data yang group JenisPasien, Penjamin, Ruang Pelayanan [Ruang Pelayanan]
    strSQL = "SELECT JenisPasien,Penjamin,[Ruang Pelayanan],SUM(TotalBiaya) AS TotalBiaya,SUM(JmlBayar) AS JmlBayar,SUM(Piutang) AS Piutang,SUM(Pembebasan) AS Pembebasan,SUM(SisaTagihan) AS SisaTagihan " _
        & "FROM V_LaporanJasaPelayananDokter " _
        & strFilter & " " _
        & "GROUP BY JenisPasien,Penjamin,[Ruang Pelayanan] ORDER BY JenisPasien"
    msubRecFO rsC, strSQL
    intJmlRow = intJmlRow + rsC.RecordCount
    
    'semua data
    strSQL = "SELECT JenisPasien,Penjamin,[Ruang Pelayanan],[Jenis Pelayanan],[Nama Pelayanan],[Komponen Jasa],SUM(TotalBiaya) AS TotalBiaya,SUM(JmlBayar) AS JmlBayar,SUM(Piutang) AS Piutang,SUM(Pembebasan) AS Pembebasan,SUM(SisaTagihan) AS SisaTagihan " _
        & "FROM V_LaporanJasaPelayananDokter " _
        & strFilter & " " _
        & "GROUP BY JenisPasien,Penjamin,[Ruang Pelayanan],[Jenis Pelayanan],[Nama Pelayanan],[Komponen Jasa] " _
        & "ORDER BY JenisPasien"
    msubRecFO rs, strSQL
    
    'jumlah baris keseluruhan
    intJmlRow = intJmlRow + rs.RecordCount
    fgData.Rows = intJmlRow + 2
    intRowNow = 0
    For i = 1 To rs.RecordCount
        intRowNow = intRowNow + 1
        For j = 1 To 6
            fgData.TextMatrix(intRowNow, j) = rs(j - 1).Value
        Next j
        For j = 7 To fgData.Cols - 1
            fgData.TextMatrix(intRowNow, j) = FormatCurrency(rs(j - 1).Value, 2)
        Next j
        curTotalBiaya = curTotalBiaya + rs("TotalBiaya").Value
        curJmlBayar = curJmlBayar + rs("JmlBayar").Value
        curPiutang = curPiutang + rs("Piutang").Value
        curPembebasan = curPembebasan + rs("Pembebasan").Value
        curSisaTagihan = curSisaTagihan + rs("SisaTagihan").Value
        fgData.TextMatrix(intRowNow, 0) = intNo
        rs.MoveNext
        intNo = intNo + 1
        'sub total per Ruang Pelayanan
        If rs.EOF = True Then GoTo stepRuangan
        If rs("Ruang Pelayanan").Value <> rsC("Ruang Pelayanan").Value Then
stepRuangan:
            intRowNow = intRowNow + 1
            fgData.TextMatrix(intRowNow, 1) = fgData.TextMatrix(intRowNow - 1, 1)
            fgData.TextMatrix(intRowNow, 2) = fgData.TextMatrix(intRowNow - 1, 2)
            fgData.TextMatrix(intRowNow, 3) = fgData.TextMatrix(intRowNow - 1, 3)
            fgData.TextMatrix(intRowNow, 4) = "Sub Total"
            fgData.TextMatrix(intRowNow, 7) = Format(rsC("TotalBiaya").Value, "#,###.00")
            fgData.TextMatrix(intRowNow, 8) = Format(rsC("JmlBayar").Value, "#,###.00")
            fgData.TextMatrix(intRowNow, 9) = Format(rsC("Piutang").Value, "#,###.00")
            fgData.TextMatrix(intRowNow, 10) = Format(rsC("Pembebasan").Value, "#,###.00")
            fgData.TextMatrix(intRowNow, 11) = Format(rsC("SisaTagihan").Value, "#,###.00")
            subSetSubTotalRow intRowNow, 4, vbBlackness, vbWhite
            If rsC.EOF Then Exit Sub
            rsC.MoveNext
        ElseIf rs("Ruang Pelayanan").Value = rsC("Ruang Pelayanan").Value And rs("JenisPasien").Value <> rsB("JenisPasien").Value Then
            intRowNow = intRowNow + 1
            fgData.TextMatrix(intRowNow, 1) = fgData.TextMatrix(intRowNow - 1, 1)
            fgData.TextMatrix(intRowNow, 2) = fgData.TextMatrix(intRowNow - 1, 2)
            fgData.TextMatrix(intRowNow, 3) = fgData.TextMatrix(intRowNow - 1, 3)
            fgData.TextMatrix(intRowNow, 4) = "Sub Total"
            fgData.TextMatrix(intRowNow, 7) = Format(rsC("TotalBiaya").Value, "#,###.00")
            fgData.TextMatrix(intRowNow, 8) = Format(rsC("JmlBayar").Value, "#,###.00")
            fgData.TextMatrix(intRowNow, 9) = Format(rsC("Piutang").Value, "#,###.00")
            fgData.TextMatrix(intRowNow, 10) = Format(rsC("Pembebasan").Value, "#,###.00")
            fgData.TextMatrix(intRowNow, 11) = Format(rsC("SisaTagihan").Value, "#,###.00")
            subSetSubTotalRow intRowNow, 4, vbBlackness, vbWhite
            If rsC.EOF Then Exit Sub
            rsC.MoveNext
        End If
        
        'sub total per Penjamin
        If rs.EOF = True Then GoTo stepPenjamin
        If rs("Penjamin").Value <> rsB("Penjamin").Value Then
stepPenjamin:
            intRowNow = intRowNow + 1
            fgData.TextMatrix(intRowNow, 1) = fgData.TextMatrix(intRowNow - 1, 1)
            fgData.TextMatrix(intRowNow, 2) = fgData.TextMatrix(intRowNow - 1, 2)
            fgData.TextMatrix(intRowNow, 3) = "Sub Total"
            fgData.TextMatrix(intRowNow, 7) = Format(rsB("TotalBiaya").Value, "#,###.00")
            fgData.TextMatrix(intRowNow, 8) = Format(rsB("JmlBayar").Value, "#,###.00")
            fgData.TextMatrix(intRowNow, 9) = Format(rsB("Piutang").Value, "#,###.00")
            fgData.TextMatrix(intRowNow, 10) = Format(rsB("Pembebasan").Value, "#,###.00")
            fgData.TextMatrix(intRowNow, 11) = Format(rsB("SisaTagihan").Value, "#,###.00")
            subSetSubTotalRow intRowNow, 3, vbYellow, vbBlack
            If rsB.EOF Then Exit Sub
            rsB.MoveNext
        ElseIf rs("Penjamin").Value = rsB("Penjamin").Value And rs("JenisPasien").Value <> rsB("JenisPasien").Value Then
            intRowNow = intRowNow + 1
            fgData.TextMatrix(intRowNow, 1) = fgData.TextMatrix(intRowNow - 1, 1)
            fgData.TextMatrix(intRowNow, 2) = fgData.TextMatrix(intRowNow - 1, 2)
            fgData.TextMatrix(intRowNow, 3) = "Sub Total"
            fgData.TextMatrix(intRowNow, 7) = Format(rsB("TotalBiaya").Value, "#,###.00")
            fgData.TextMatrix(intRowNow, 8) = Format(rsB("JmlBayar").Value, "#,###.00")
            fgData.TextMatrix(intRowNow, 9) = Format(rsB("Piutang").Value, "#,###.00")
            fgData.TextMatrix(intRowNow, 10) = Format(rsB("Pembebasan").Value, "#,###.00")
            fgData.TextMatrix(intRowNow, 11) = Format(rsB("SisaTagihan").Value, "#,###.00")
            subSetSubTotalRow intRowNow, 3, vbYellow, vbBlack
            If rsB.EOF Then Exit Sub
            rsB.MoveNext
        End If
        
        'sub total per JenisPasien
        If rs.EOF = True Then GoTo stepJenisPasien
        If rs("JenisPasien").Value <> rsA("JenisPasien").Value Then
stepJenisPasien:
            intRowNow = intRowNow + 1
            fgData.TextMatrix(intRowNow, 1) = fgData.TextMatrix(intRowNow - 1, 1)
            fgData.TextMatrix(intRowNow, 2) = "Sub Total"
            fgData.TextMatrix(intRowNow, 7) = Format(rsA("TotalBiaya").Value, "#,###.00")
            fgData.TextMatrix(intRowNow, 8) = Format(rsA("JmlBayar").Value, "#,###.00")
            fgData.TextMatrix(intRowNow, 9) = Format(rsA("Piutang").Value, "#,###.00")
            fgData.TextMatrix(intRowNow, 10) = Format(rsA("Pembebasan").Value, "#,###.00")
            fgData.TextMatrix(intRowNow, 11) = Format(rsA("SisaTagihan").Value, "#,###.00")
            subSetSubTotalRow intRowNow, 2, vbGreen, vbBlack
            If rsA.EOF Then Exit Sub
            rsA.MoveNext
        End If
    Next i
    
    intRowNow = intRowNow + 1
    fgData.TextMatrix(intRowNow, 1) = "Total"
    fgData.TextMatrix(intRowNow, 7) = Format(curTotalBiaya, "#,###.00")
    fgData.TextMatrix(intRowNow, 8) = Format(curJmlBayar, "#,###.00")
    fgData.TextMatrix(intRowNow, 9) = Format(curPiutang, "#,###.00")
    fgData.TextMatrix(intRowNow, 10) = Format(curPembebasan, "#,###.00")
    fgData.TextMatrix(intRowNow, 11) = Format(curSisaTagihan, "#,###.00")
    subSetSubTotalRow intRowNow, 1, vbRed, vbWhite
    
    fgData.Visible = True
    Exit Sub
errLoad:
    fgData.Visible = True
    msubPesanError
End Sub

'untuk setting row sub total
Private Sub subSetSubTotalRow(iRowNow As Integer, iColBegin As Integer, vbBackColor, vbForeColor)
Dim i As Integer
    With fgData
        'tampilan Black & White
        For i = iColBegin To .Cols - 1
            .Col = i
            .Row = iRowNow
            .CellBackColor = vbBackColor
            .CellForeColor = vbForeColor
            .RowHeight(.Row) = 300
            .CellFontBold = True
        Next
    End With
End Sub

'untuk setting grid jasa pelayanan
Private Sub subSetGrid()
Dim i As Integer
    With fgData
        .Clear
        .Cols = 12
        .Rows = 2
        .ColWidth(0) = 450
        .ColWidth(1) = 1200
        .ColWidth(2) = 1400
        .ColWidth(3) = 1800
        .ColWidth(4) = 2000
        .ColWidth(5) = 2000
        .ColWidth(6) = 2000
        .ColWidth(7) = 1700
        .ColWidth(8) = 1700
        .ColWidth(9) = 1700
        .ColWidth(10) = 1700
        .ColWidth(11) = 1700
        
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignRightCenter
        .ColAlignment(10) = flexAlignRightCenter
        .ColAlignment(11) = flexAlignRightCenter
        .Row = 0
        .RowHeight(0) = 300
        For i = 1 To .Cols - 1
            .Col = i
            .CellAlignment = flexAlignCenterCenter
            .CellFontBold = True
        Next i
        .TextMatrix(0, 1) = "Kel. Pasien"
        .TextMatrix(0, 2) = "Penjamin"
        .TextMatrix(0, 3) = "Ruangan Pelayanan"
        .TextMatrix(0, 4) = "Jenis Pelayanan"
        .TextMatrix(0, 5) = "Nama Pelayanan"
        .TextMatrix(0, 6) = "Komponen Jasa"
        .TextMatrix(0, 7) = "Total Biaya"
        .TextMatrix(0, 8) = "Bayar"
        .TextMatrix(0, 9) = "Piutang"
        .TextMatrix(0, 10) = "Pembebasan"
        .TextMatrix(0, 11) = "Sisa Tagihan"
        .MergeCells = 1
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .MergeCol(4) = True
    End With
End Sub
