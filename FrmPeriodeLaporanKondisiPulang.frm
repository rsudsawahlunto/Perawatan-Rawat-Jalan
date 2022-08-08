VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmPeriodeLaporanKondisiPulang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPeriodeLaporanKondisiPulang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   9975
   Begin MSFlexGridLib.MSFlexGrid fgdata 
      Height          =   4425
      Left            =   0
      TabIndex        =   9
      Top             =   2100
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   7805
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periode Laporan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   9915
      Begin VB.CommandButton cmdcari 
         Caption         =   "Cari"
         Height          =   375
         Left            =   6720
         TabIndex        =   14
         Top             =   480
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPickerAwal 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM, yyyy"
         Format          =   52232195
         CurrentDate     =   37956
      End
      Begin MSComCtl2.DTPicker DTPickerAkhir 
         Height          =   375
         Left            =   4320
         TabIndex        =   5
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM, yyyy"
         Format          =   52232195
         CurrentDate     =   37956
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "s/d"
         Height          =   210
         Index           =   1
         Left            =   3960
         TabIndex        =   15
         Top             =   562
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Awal"
         Height          =   210
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Akhir"
         Height          =   210
         Index           =   0
         Left            =   4320
         TabIndex        =   6
         Top             =   240
         Width           =   1110
      End
   End
   Begin VB.Frame Frame2 
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
      Left            =   0
      TabIndex        =   0
      Top             =   6510
      Width           =   5805
      Begin VB.CommandButton cmdgrafik 
         Caption         =   "&Grafik"
         Height          =   375
         Left            =   2070
         TabIndex        =   8
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Spreadsheet"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   210
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   3900
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox txtJmlTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   8610
      TabIndex        =   10
      Top             =   6810
      Width           =   1000
   End
   Begin VB.TextBox txtJmlWanita 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   7350
      TabIndex        =   11
      Top             =   6810
      Width           =   1000
   End
   Begin VB.TextBox txtJmlPria 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   6060
      TabIndex        =   12
      Top             =   6810
      Width           =   1000
   End
   Begin VB.Frame Frame4 
      Caption         =   "  Jml Pria            Jml Wanita      Jml Total"
      Height          =   735
      Left            =   5820
      TabIndex        =   13
      Top             =   6510
      Width           =   4125
   End
   Begin VB.Image Image2 
      Height          =   930
      Left            =   -210
      Picture         =   "FrmPeriodeLaporanKondisiPulang.frx":08CA
      Top             =   0
      Width           =   10200
   End
End
Attribute VB_Name = "FrmPeriodeLaporanKondisiPulang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iRowNow As Integer
Dim rswilayah As New ADODB.recordset
Dim iRowNow2 As Integer

Private Sub cmdcari_Click()
Dim intJmlRow As Integer
Dim intJmlPria As Integer
Dim intJmlWanita As Integer
Dim intJmlTotal As Integer

    Call subSetGrid
    'u/ mempercepat
    fgdata.Visible = False: MousePointer = vbHourglass
    
    'Hitung jumlah row dari data yang hendak ditampilkan
    strSQL = "SELECT COUNT(Tglkeluar) AS JmlRow " & _
        " FROM V_RekapitulasiBKondisiPulang " & _
        " WHERE Tglkeluar BETWEEN " & _
        " '" & Format(DTPickerAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " & _
        " AND kdruangan = '" & mstrKdRuangan & "'"
        
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    'jika tidak ada data
    If rs(0).Value = 0 Then
        fgdata.Visible = True: MousePointer = vbNormal
        MsgBox "Tidak ada Data"
'        MsgBox "Tidak ada data antara tanggal  '" & Format(DTPickerAwal.Value, "dd - MMMM - yyyy") & "' dan '" & Format(dtpTglAkhir.Value, "dd - MMMM - yyyy") & "' ", vbInformation, "Validasi"
        txtJmlPria = "0": txtJmlTotal = "0": txtJmlWanita = "0"
        Exit Sub
    End If
    
    intJmlRow = rs("JmlRow").Value
    
    strSQL = "SELECT namaruangan,COUNT(namaruangan) AS JmlkdsPlg, " & _
        " SUM(JmlPasienPria) AS JmlPria,SUM(JmlPasienWanita) AS JmlWanita," & _
        " SUM(Total) AS Total From V_RekapitulasiBKondisiPulang" & _
        " WHERE Tglkeluar BETWEEN " & _
        " '" & Format(DTPickerAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " & _
        " AND kdruangan = '" & mstrKdRuangan & "'" & _
        " GROUP BY namaruangan"
        
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    
    'Tambahkan jumlah row dengan jumlah subtotal
    intJmlRow = intJmlRow + rs.RecordCount
    
    'u/ menampilkan yang di group by
    Dim rsjenisoperasi As New ADODB.recordset
    With fgdata
        'jml baris akhir
        .Rows = intJmlRow + 2
        While rs.EOF = False
            strSQL = "SELECT namaruangan,kondisipulang,SUM(JmlPasienPria) AS TJmlPasienPria, " & _
                " SUM(JmlPasienWanita) AS TJmlPasienWanita, " & _
                " SUM(Total) AS TTotal From V_RekapitulasiBKondisiPulang " & _
                " WHERE (Tglkeluar BETWEEN " & _
                " '" & Format(DTPickerAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND " & _
                " '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd 23:59:59") & "') " & _
                " AND namaruangan='" & rs("namaruangan").Value & "' " & _
                " AND kdruangan = '" & mstrKdRuangan & "'" & _
                " GROUP BY namaruangan, kondisipulang" & _
                " ORDER BY kondisipulang"
            
            Set rskondisipulang = Nothing
            rsjenisoperasi.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            While rsjenisoperasi.EOF = False
                'baris u/ sub total
                iRowNow = iRowNow + 1
                .TextMatrix(iRowNow, 1) = rsjenisoperasi("namaruangan").Value
                .TextMatrix(iRowNow, 2) = rsjenisoperasi("kondisipulang").Value
                .TextMatrix(iRowNow, 3) = rsjenisoperasi("TJmlPasienPria").Value
                .TextMatrix(iRowNow, 4) = rsjenisoperasi("TJmlPasienWanita").Value
                .TextMatrix(iRowNow, 5) = rsjenisoperasi("TTotal").Value
                rsjenisoperasi.MoveNext
            Wend
            
            iRowNow = iRowNow + 1
            'isi sub total
            .TextMatrix(iRowNow, 2) = "Sub Total"
            .TextMatrix(iRowNow, 3) = rs("JmlPria").Value
            .TextMatrix(iRowNow, 4) = rs("JmlWanita").Value
            .TextMatrix(iRowNow, 5) = rs("Total").Value
            
            'tampilan Black & White
            For i = 1 To .Cols - 1
                .Col = i
                .Row = iRowNow
                .CellBackColor = vbBlackness
                .CellForeColor = vbWhite
                If .Col = 1 Then .TextMatrix(.Row, 1) = .TextMatrix(.Row - 1, 1): .CellBackColor = vbWhite: .CellForeColor = vbBlack
                .RowHeight(.Row) = 300
                .CellFontBold = True
            Next
            
            'disimpan u/ jml total
            intJmlPria = intJmlPria + rs("JmlPria").Value
            intJmlWanita = intJmlWanita + rs("JmlWanita").Value
            intJmlTotal = intJmlTotal + rs("Total").Value
            
            rs.MoveNext
        Wend
        'banyak baris berdasarkan irownow
        .Rows = iRowNow + 2
    
        'jml total
        txtJmlPria.Text = intJmlPria
        txtJmlWanita.Text = intJmlWanita
        txtJmlTotal.Text = intJmlTotal
        
        .Col = 1
        For i = 1 To .Rows - 1
            .Row = i
            .CellFontBold = True
        Next
        
        .Visible = True: MousePointer = vbNormal
    End With


End Sub

Private Sub cmdCetak_Click()
    cmdCetak.Enabled = False
'    mdTglAwal = dtpTglAwal.Value
'    mdTglAkhir = dtpTglAkhir.Value
    strSQL = "SELECT * FROM V_RekapitulasiBKondisiPulang " _
        & "WHERE (Tglkeluar BETWEEN '" _
        & Format(DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "') ORDER BY namaruangan,kondisipulang"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    cetak = "Rekapkondisipulang"
    FrmViewerLaporan.Show
    FrmViewerLaporan.Caption = "Medifirst2000 - Rekapitulasi Pasien Berdasarkan Kondisi Pulang Pasien"
    cmdCetak.Enabled = True

End Sub

Private Sub cmdgrafik_Click()
    cmdCetak.Enabled = False
    
    strSQL = "SELECT * FROM V_RekapitulasiBKondisiPulang " _
        & "WHERE (Tglkeluar BETWEEN '" _
        & Format(DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "') ORDER BY namaruangan,kondisipulang"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    
    cetak = "Rekapgrafikkondisipulang"
    FrmViewerLaporan.Show
    FrmViewerLaporan.Caption = "Medifirst2000 - Grafik Rekapitulasi Pasien Berdasarkan Kondisi Pulang Pasien"
    cmdCetak.Enabled = True

End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    txtJmlPria = "0": txtJmlTotal = "0": txtJmlWanita = "0"
    
    With Me
        .DTPickerAwal.Value = Now
        .DTPickerAkhir.Value = Now
    End With
    
'    Set dbRec = New ADODB.recordset
'    dbRec.Open " SELECT     KdInstalasi, NamaInstalasi " _
'             & " FROM         Instalasi where kdinstalasi <> '06' ", dbConn, adOpenDynamic, adLockOptimistic
'
'    While dbRec.EOF = False
'        CboInstalasi.AddItem dbRec.Fields(0).Value & " - " & dbRec.Fields(1).Value
'        dbRec.MoveNext
'    Wend

    Call subSetGrid
    Call SetText
End Sub

Private Sub subSetGrid()
    With fgdata
        .Visible = False
        .Clear
        .Cols = 6
        .Rows = 2
        .Row = 0
        
        For i = 1 To .Cols - 1
            .Col = i
            .CellFontBold = True
            .RowHeight(0) = 300
            .CellAlignment = flexAlignCenterCenter
        Next
        
        .MergeCells = 1
        .MergeCol(1) = True
        
        .TextMatrix(0, 1) = "Ruangan"
        .TextMatrix(0, 2) = "Kondisi Pulang"
        .TextMatrix(0, 3) = "Laki-Laki"
        .TextMatrix(0, 4) = "Perempuan"
        .TextMatrix(0, 5) = "Total"
        
        .ColWidth(0) = 500
        .ColWidth(1) = 2850
        .ColWidth(2) = 2850
        .ColWidth(3) = 1100
        .ColWidth(4) = 1100
        .ColWidth(5) = 1100
        
        .Visible = True
        iRowNow = 0
    End With
End Sub

Sub SetText()
    With fgdata
        .Col = 3
        txtJmlPria.Left = .Left + .CellLeft
        txtJmlPria.Width = .CellWidth
    
        .Col = 4
        txtJmlWanita.Left = .Left + .CellLeft
        txtJmlWanita.Width = .CellWidth
    
        .Col = 5
        txtJmlTotal.Left = .Left + .CellLeft
        txtJmlTotal.Width = .CellWidth
    End With
End Sub

