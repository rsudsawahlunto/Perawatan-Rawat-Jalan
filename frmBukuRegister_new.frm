VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBukuRegister_new 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buku Register"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBukuRegister_new.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   10200
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
      Height          =   5835
      Left            =   0
      TabIndex        =   8
      Top             =   930
      Width           =   10215
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
         Left            =   4320
         TabIndex        =   9
         Top             =   150
         Width           =   5775
         Begin VB.CommandButton cmdcari 
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
            Format          =   61079555
            UpDown          =   -1  'True
            CurrentDate     =   38209
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
            Format          =   61079555
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   10
            Top             =   315
            Width           =   255
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgdata 
         Height          =   4545
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   9945
         _ExtentX        =   17542
         _ExtentY        =   8017
         _Version        =   393216
         Appearance      =   0
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
      TabIndex        =   7
      Top             =   6840
      Width           =   10245
      Begin VB.CommandButton cmdgrafik 
         Caption         =   "&Grafik"
         Height          =   375
         Left            =   6510
         TabIndex        =   5
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Spredsheet"
         Height          =   375
         Left            =   4680
         TabIndex        =   4
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   8340
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   11
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
      Left            =   8400
      Picture         =   "frmBukuRegister_new.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmBukuRegister_new.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmBukuRegister_new.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmBukuRegister_new"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iRowNow As Integer
Dim rstopten As New ADODB.recordset
Dim iRowNow2 As Integer
Dim i As Integer

Private Sub cmdCari_Click()
    On Error GoTo hell
    fgData.SetFocus
    Dim intJmlRow As Integer
    Dim intJmlPria As Integer
    Dim intJmlWanita As Integer
    Dim intJmlTotal As Integer

    Call subSetGrid
    fgData.Visible = False: MousePointer = vbHourglass

    'Hitung jumlah row dari data yang hendak ditampilkan
    strSQL = "SELECT COUNT(TglPeriksa) AS JmlRow " & _
    " FROM V_RekapitulasiDiagnosaTopTen " & _
    " WHERE TglPeriksa BETWEEN " & _
    " '" & Format(DTPickerAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND " & _
    " '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " & _
    " AND NamaRuangan = '" & mstrNamaRuangan & "'"

    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    'jika tidak ada data
    If rs(0).Value = 0 Then
        fgData.Visible = True: MousePointer = vbNormal
        MsgBox "Tidak ada Data"
        Exit Sub
    End If

    intJmlRow = rs("JmlRow").Value

    strSQL = "SELECT Instalasi,COUNT(instalasi) AS Jmldiagnosa, " & _
    " SUM(jumlahpasien)as jumlah From V_RekapitulasiDiagnosaTopTen " & _
    " WHERE TglPeriksa BETWEEN " & _
    " '" & Format(DTPickerAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND " & _
    " '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " & _
    " AND NamaRuangan = '" & mstrNamaRuangan & "'" & _
    " GROUP BY instalasi"

    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    'u/ menampilkan yang di group by
    Dim rstopten As New ADODB.recordset
    With fgData
        'jml baris akhir
        .Rows = intJmlRow + 2
        While rs.EOF = False
            strSQL = "SELECT instalasi,diagnosa,sum(jumlahpasien) as TjumlahPasien " & _
            " From V_RekapitulasiDiagnosaTopTen " & _
            " WHERE (TglPeriksa BETWEEN " & _
            " '" & Format(DTPickerAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND " & _
            " '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd 23:59:59") & "') " & _
            " AND instalasi ='" & rs("instalasi").Value & "' " & _
            " AND NamaRuangan = '" & mstrNamaRuangan & "'" & _
            " GROUP BY instalasi, diagnosa" & _
            " ORDER BY diagnosa"

            Set rstopten = Nothing
            rstopten.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            While rstopten.EOF = False
                'baris u/ sub total
                iRowNow = iRowNow + 1
                .TextMatrix(iRowNow, 1) = rstopten("Instalasi").Value
                .TextMatrix(iRowNow, 2) = rstopten("Diagnosa").Value
                .TextMatrix(iRowNow, 3) = rstopten("TJumlahPasien").Value
                rstopten.MoveNext
            Wend

            'disimpan u/ jml total
            intJmlTotal = intJmlTotal + rs("jumlah").Value

            rs.MoveNext
        Wend
        'banyak baris berdasarkan irownow
        .Rows = iRowNow + 2

        .Col = 1
        For i = 1 To .Rows - 1
            .Row = i
            .CellFontBold = True
        Next

        .Visible = True: MousePointer = vbNormal
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell
    cmdCetak.Enabled = False
    strSQL = "SELECT * FROM V_RekapitulasiDiagnosaTopTen " _
    & "WHERE (TglPeriksa BETWEEN '" _
    & Format(DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
    & Format(DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "') and NamaRuangan = '" & mstrNamaRuangan & "' "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    cetak = "RekapTopten"
    FrmViewerLaporan.Show
    FrmViewerLaporan.Caption = "Medifirst2000 - Rekapitulasi 10 Besar Penyakit"
    cmdCetak.Enabled = True
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdgrafik_Click()
    On Error GoTo hell
    cmdCetak.Enabled = False
    strSQL = "SELECT * FROM V_RekapitulasiDiagnosaTopTen " _
    & "WHERE (TglPeriksa BETWEEN '" _
    & Format(DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
    & Format(DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "') and NamaRuangan = '" & mstrNamaRuangan & "' ORDER BY instalasi,diagnosa"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If

    cetak = "RekapToptenGrafik"
    FrmViewerLaporan.Show
    FrmViewerLaporan.Caption = "Medifirst2000 - Grafik Rekapitulasi 10 Besar Penyakit"
    cmdCetak.Enabled = True
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub DTPickerAkhir_Change()
    DTPickerAkhir.MaxDate = Now
End Sub

Private Sub DTPickerAkhir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCari.SetFocus
End Sub

Private Sub DTPickerAwal_Change()
    DTPickerAwal.MaxDate = Now
End Sub

Private Sub DTPickerAwal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then DTPickerAkhir.SetFocus
End Sub

Private Sub fgData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetak.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)

    With Me
        .DTPickerAwal.Value = Now
        .DTPickerAkhir.Value = Now
    End With

    Call subSetGrid
    Call PlayFlashMovie(Me)
End Sub

Private Sub subSetGrid()
    With fgData
        .Visible = False
        .Clear
        .Cols = 4
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

        .TextMatrix(0, 1) = "Nama Sub Instalasi"
        .TextMatrix(0, 2) = "Diagnosa Penyakit"
        .TextMatrix(0, 3) = "Jumlah Pasien"

        .ColWidth(0) = 500
        .ColWidth(1) = 2850
        .ColWidth(2) = 4000
        .ColWidth(3) = 2000

        .Visible = True
        iRowNow = 0
    End With
End Sub

