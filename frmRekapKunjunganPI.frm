VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRekapKunjunganPI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rekapitulasi Kunjungan Pasien Internal"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRekapKunjunganPI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   10380
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
      TabIndex        =   14
      Top             =   7320
      Width           =   10365
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   8520
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Spreadsheet"
         Height          =   375
         Left            =   4920
         TabIndex        =   3
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdgrafik 
         Caption         =   "&Grafik"
         Height          =   375
         Left            =   6720
         TabIndex        =   4
         Top             =   240
         Width           =   1665
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
      Height          =   6375
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   10335
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
         Left            =   4440
         TabIndex        =   15
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
            OLEDropMode     =   1
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   58064899
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
            Format          =   58064899
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   16
            Top             =   315
            Width           =   255
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgData 
         Height          =   5055
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   8916
         _Version        =   393216
         BackColorBkg    =   8421504
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame fraTotal 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   6480
      Visible         =   0   'False
      Width           =   7455
      Begin VB.TextBox txtJmlPria 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   440
         Width           =   1500
      End
      Begin VB.TextBox txtJmlWanita 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   11040
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   440
         Width           =   1500
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   12600
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   440
         Width           =   1500
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Jml Pria"
         Height          =   210
         Left            =   9930
         TabIndex        =   13
         Top             =   195
         Width           =   600
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Jml Wanita"
         Height          =   210
         Left            =   11340
         TabIndex        =   11
         Top             =   195
         Width           =   900
      End
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   210
         Left            =   12600
         TabIndex        =   8
         Top             =   195
         Width           =   1500
      End
   End
   Begin VB.Image Image1 
      Height          =   930
      Left            =   200
      Picture         =   "frmRekapKunjunganPI.frx":08CA
      Top             =   0
      Width           =   10200
   End
   Begin VB.Image Image2 
      Height          =   930
      Left            =   0
      Picture         =   "frmRekapKunjunganPI.frx":6012
      Top             =   0
      Width           =   10200
   End
End
Attribute VB_Name = "frmRekapKunjunganPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i, j As Long
Dim intJmlRow As Long
Dim intRowNow As Long
Dim rsa As New ADODB.recordset
Dim rsB As New ADODB.recordset
Dim rsC As New ADODB.recordset

Private Sub cmdCetak_Click()
    cmdCetak.Enabled = False
    strSQL = " SELECT TglPendaftaran, Ruangan, [Ruang Perujuk], JenisPasien, JmlPasienPria, JmlPasienWanita " & _
        " FROM V_RekapitulasiPasienBRujukanInternal " & _
        " WHERE TglPendaftaran BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "'" & _
        " AND KdRuangan = '" & mstrKdRuangan & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbExclamation, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If

    Set frmCetakRekapKunjunganPI = Nothing
    frmCetakRekapKunjunganPI.Show
    cmdCetak.Enabled = True
End Sub

Private Sub cmdgrafik_Click()
    cmdCetak.Enabled = False
    strSQL = "SELECT * FROM V_RekapitulasiPasienBRujukanInternal " _
        & "WHERE (TglPendaftaran BETWEEN '" _
        & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "')" _
        & " AND kdruangan = '" & mstrKdRuangan & "'"
    
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value
    
    cetak = "RekapKunjunganPIGrafik"
    FrmViewerLaporan.Show
    FrmViewerLaporan.Caption = "Medifirst2000 - Grafik Rekapitulasi Kunjungan Pasien Internal"
    cmdCetak.Enabled = True

End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdTampilkanTemp_Click()
    On Error GoTo errTampilkan
    subSetGridAja
    intJmlRow = 0
    fgData.Visible = False
    
    subLoadKunjunganPI
   
    fgData.Visible = True
    cmdCetak.SetFocus
    Exit Sub
errTampilkan:
    msubPesanError
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdTampilkanTemp.SetFocus
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
    
    Me.Caption = "Medifirst2000 - Laporan Rekapitulasi Kunjungan Pasien Internal"
    subSetGridAja
    dtpAwal.Value = Now
    dtpAkhir.Value = Now
    Exit Sub
errFormLoad:
    msubPesanError
End Sub

'untuk setting grid tanpa loading data
Private Sub subSetGridAja()
Dim i As Long
    With fgData
        .Clear
        .Cols = 6
        .Rows = 2
        .ColWidth(0) = 500
        .ColWidth(1) = 3000
        .ColWidth(2) = 3000
        .ColWidth(3) = 1100
        .ColWidth(4) = 1100
        .ColWidth(5) = 1100
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignment(5) = flexAlignRightCenter
        .Row = 0
        .RowHeight(0) = 300
        For i = 1 To .Cols - 1
            .Col = i
            .CellAlignment = flexAlignCenterCenter
            .CellFontBold = True
        Next i
        .TextMatrix(0, 1) = "Ruang Perujuk"
        .MergeCol(1) = True
        .TextMatrix(0, 2) = "Jenis Pasien"
        .MergeCol(2) = True
        .TextMatrix(0, 3) = "Laki-Laki"
        .TextMatrix(0, 4) = "Perempuan"
        .TextMatrix(0, 5) = "Total"
        .MergeCells = 1
    End With
End Sub

Private Sub subLoadKunjunganPI()
    strSQL = " SELECT [Ruang Perujuk], JenisPasien, JmlPasienPria, JmlPasienWanita FROM V_RekapitulasiPasienBRujukanInternal " & _
             " WHERE TglPendaftaran BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "'" & _
             " AND KdRuangan = '" & mstrKdRuangan & "'"
    msubOpenRecFO rs, strSQL, dbConn
    
    'jumlah baris keseluruhan
    intJmlRow = intJmlRow + rs.RecordCount
    fgData.Rows = intJmlRow + 2
    intRowNow = 0
    For i = 1 To rs.RecordCount
        intRowNow = intRowNow + 1
        For j = 1 To fgData.Cols - 1
            If j = fgData.Cols - 1 Then
                fgData.TextMatrix(intRowNow, j) = rs(j - 3).Value + rs(j - 2).Value
            Else
                If IsNull(rs(j - 1).Value) Then
                    fgData.TextMatrix(intRowNow, j) = ""
                Else
                    fgData.TextMatrix(intRowNow, j) = rs(j - 1).Value
                End If
            End If
        Next j
        rs.MoveNext
    Next i
    
    strSQL = " SELECT TglPendaftaran " & _
             " FROM V_RekapitulasiPasienBRujukanInternal " & _
             " WHERE TglPendaftaran BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "'" & _
             " AND KdRuangan = '" & mstrKdRuangan & "'"
    msubOpenRecFO rsB, strSQL, dbConn
    If rsB.RecordCount = 0 Then
        txtJmlPria.Text = 0
        txtJmlWanita.Text = 0
        txtTotal.Text = 0
    Else
        strSQL = " SELECT SUM(JmlPasienPria) as TotalPria, SUM(JmlPasienWanita) as TotalWanita FROM V_RekapitulasiPasienBRujukanInternal" & _
                 " WHERE TglPendaftaran BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "'" & _
                 " AND KdRuangan = '" & mstrKdRuangan & "'"
        msubOpenRecFO rsB, strSQL, dbConn
        intRowNow = intRowNow + 1
        fgData.TextMatrix(intRowNow, 1) = "Total"
        fgData.TextMatrix(intRowNow, 3) = rsB("TotalPria")
        fgData.TextMatrix(intRowNow, 4) = rsB("TotalWanita")
        fgData.TextMatrix(intRowNow, 5) = rsB("TotalPria") + rsB("TotalWanita")
        subSetSubTotalRow intRowNow, 1, vbBlack, vbWhite
        txtJmlPria.Text = rsB("TotalPria")
        txtJmlWanita.Text = rsB("TotalWanita")
        txtTotal.Text = rsB("TotalPria") + rsB("totalWanita")
    End If
End Sub

Private Sub subSetSubTotalRow(iRowNow As Long, iColBegin As Long, vbBackColor, vbForeColor)
Dim i As Long
    With fgData
        'tampilan Black & White
        For i = iColBegin To .Cols - 1
            .Col = i
            .Row = iRowNow
            .CellBackColor = vbBackColor
            .CellForeColor = vbForeColor
'            .RowHeight(.Row) = 300
            .CellFontBold = True
        Next
    End With
End Sub





