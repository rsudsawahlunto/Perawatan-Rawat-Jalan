VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLapMorbiditas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Data Morbiditas Pasien"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14115
   Icon            =   "frmLapMorbiditas.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   14115
   Begin VB.Frame fraButton 
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
      TabIndex        =   6
      Top             =   6960
      Width           =   14085
      Begin VB.CommandButton cmdCetak 
         Caption         =   "Ce&tak"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10440
         TabIndex        =   4
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12240
         TabIndex        =   5
         Top             =   240
         Width           =   1695
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
      Height          =   5955
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   14085
      Begin VB.Frame Frame4 
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
         Left            =   8160
         TabIndex        =   8
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
            Format          =   60686339
            UpDown          =   -1  'True
            CurrentDate     =   38209
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
            Format          =   60686339
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   9
            Top             =   315
            Width           =   255
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   4635
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   13845
         _ExtentX        =   24421
         _ExtentY        =   8176
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   10
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
      Left            =   12240
      Picture         =   "frmLapMorbiditas.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmLapMorbiditas.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmLapMorbiditas.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12375
   End
End
Attribute VB_Name = "frmLapMorbiditas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iRowNow As Integer
Dim rsTemp1 As ADODB.recordset
Dim rsTemp2 As ADODB.recordset
Dim i As Integer

Private Sub cmdCari_Click()
    On Error GoTo hell
    cmdCetak.Enabled = True
    Dim intJmlRow As Integer
    Dim intNo As Integer
    Dim intJmlPria As Integer
    Dim intJmlWanita As Integer
    Dim intJmlTotal As Integer
    Call subSetGrid
    fgData.Visible = False
    MousePointer = vbHourglass
    intNo = 0
    iRowNow = 0
    intJmlPria = 0
    intJmlWanita = 0
    intJmlTotal = 0
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value
    'Hitung jumlah row dari data yang hendak ditampilkan
    strSQL = "SELECT NoDTD,NoDTerperinci,NamaDTD,SUM(Kel_Umur1) AS Kel_Umur1,SUM(Kel_Umur2) AS Kel_Umur2,SUM(Kel_Umur3) AS Kel_Umur3,SUM(Kel_Umur4) AS Kel_Umur4,SUM(Kel_Umur5) AS Kel_Umur5,SUM(Kel_Umur6) AS Kel_Umur6,SUM(Kel_Umur7) AS Kel_Umur7,SUM(Kel_Umur8) AS Kel_Umur8,SUM(Kel_L) AS Kel_L,SUM(Kel_P) AS Kel_P,SUM(Kel_Kunj) AS Kel_Kunj " _
    & "FROM v_S_RekapMorbidBuatRJ " _
    & "WHERE TglPeriksa BETWEEN '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "' " _
    & " AND (KdRuangan = '" & mstrKdRuangan & "') " _
    & "GROUP BY NoDTD,NoDTerperinci,NamaDTD"
    Call msubRecFO(rs, strSQL)
    'jika tidak ada data
    If rs.EOF = True Then
        fgData.Visible = True
        MousePointer = vbNormal
        MsgBox "Tidak ada Data", vbExclamation, "Validasi"
        Exit Sub
    End If
    intJmlRow = rs.RecordCount + 2
    With fgData
        .Rows = intJmlRow
        While rs.EOF = False
            iRowNow = iRowNow + 1
            intNo = intNo + 1
            .TextMatrix(iRowNow, 0) = intNo
            .TextMatrix(iRowNow, 1) = rs("NoDTD").Value
            .TextMatrix(iRowNow, 2) = rs("NoDTerperinci").Value
            .TextMatrix(iRowNow, 3) = rs("NamaDTD").Value
            .TextMatrix(iRowNow, 4) = rs("Kel_Umur1").Value
            .TextMatrix(iRowNow, 5) = rs("Kel_Umur2").Value
            .TextMatrix(iRowNow, 6) = rs("Kel_Umur3").Value
            .TextMatrix(iRowNow, 7) = rs("Kel_Umur4").Value
            .TextMatrix(iRowNow, 8) = rs("Kel_Umur5").Value
            .TextMatrix(iRowNow, 9) = rs("Kel_Umur6").Value
            .TextMatrix(iRowNow, 10) = rs("Kel_Umur7").Value
            .TextMatrix(iRowNow, 11) = rs("Kel_Umur8").Value
            .TextMatrix(iRowNow, 12) = rs("Kel_L").Value
            .TextMatrix(iRowNow, 13) = rs("Kel_P").Value
            .TextMatrix(iRowNow, 14) = rs("Kel_Kunj").Value
            rs.MoveNext
        Wend
    End With
    fgData.Visible = True
    MousePointer = vbNormal
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value

    If ValidasiTanggal(dtpAwal, dtpAkhir) = False Then Exit Sub
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then MsgBox "Tidak Ada Data", vbExclamation, "Validasi": Exit Sub
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmMorbiditasRJ.Show
    Exit Sub
hell:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dtpAkhir_Change()
    On Error Resume Next
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpAwal_Change()
    On Error Resume Next
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub fgData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetak.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    With Me
        .dtpAwal.Value = Now
        .dtpAkhir.Value = Now
    End With
    Call subSetGrid
    Call PlayFlashMovie(Me)
    cmdCetak.Enabled = False
End Sub

'Untuk setting grid
Private Sub subSetGrid()
    With fgData
        .Visible = False
        .Clear
        .Cols = 15
        .Rows = 2
        .Row = 0

        For i = 0 To .Cols - 1
            .Col = i
            .CellFontBold = True
            .RowHeight(0) = 300
            .CellAlignment = flexAlignCenterCenter
        Next

        .TextMatrix(0, 0) = "No."
        .TextMatrix(0, 1) = "NoDTD"
        .TextMatrix(0, 2) = "NoDTerperinci"
        .TextMatrix(0, 3) = "NamaDTD"
        .TextMatrix(0, 4) = "0>28hr"
        .TextMatrix(0, 5) = ">=28hr<=1thn"
        .TextMatrix(0, 6) = "1-4thn"
        .TextMatrix(0, 7) = "5-14thn"
        .TextMatrix(0, 8) = "15-24thn"
        .TextMatrix(0, 9) = "25-44thn"
        .TextMatrix(0, 10) = "45-64thn"
        .TextMatrix(0, 11) = ">=65thn"
        .TextMatrix(0, 12) = "Laki-Laki"
        .TextMatrix(0, 13) = "Wanita"
        .TextMatrix(0, 14) = "Total"

        .ColWidth(0) = 500
        .ColWidth(1) = 900
        .ColWidth(2) = 3500
        .ColWidth(3) = 4000
        .ColWidth(4) = 1300
        .ColWidth(5) = 1300
        .ColWidth(6) = 1300
        .ColWidth(7) = 1300
        .ColWidth(8) = 1300
        .ColWidth(9) = 1300
        .ColWidth(10) = 1300
        .ColWidth(11) = 1300
        .ColWidth(12) = 1300
        .ColWidth(13) = 1300
        .ColWidth(14) = 1300

        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignCenterCenter
        .ColAlignment(7) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignCenterCenter
        .ColAlignment(9) = flexAlignCenterCenter
        .ColAlignment(10) = flexAlignCenterCenter
        .ColAlignment(11) = flexAlignCenterCenter
        .ColAlignment(12) = flexAlignCenterCenter
        .ColAlignment(13) = flexAlignCenterCenter
        .ColAlignment(14) = flexAlignCenterCenter
        .Visible = True
        iRowNow = 0
    End With
End Sub

'Untuk mensetting grid di row subtotal
Private Sub subSetSubTotalRow(iRowNow As Integer, iColMulai As Integer, vbBackColor, vbForeColor)
    Dim i As Integer
    With fgData
        'tampilan Black & White
        For i = iColMulai To .Cols - 1
            .Col = i
            .Row = iRowNow
            .CellBackColor = vbBackColor
            .CellForeColor = vbForeColor
            .RowHeight(.Row) = 300
            .CellFontBold = True
        Next
    End With
End Sub

