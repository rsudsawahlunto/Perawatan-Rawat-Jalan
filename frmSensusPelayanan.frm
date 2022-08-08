VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSensusPelayanan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sensus Pelayanan"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15015
   Icon            =   "frmSensusPelayanan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   15015
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   7320
      Width           =   15015
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11400
         TabIndex        =   0
         Top             =   240
         Width           =   1695
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
         Height          =   495
         Left            =   13200
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   15015
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
         Left            =   9120
         TabIndex        =   4
         Top             =   240
         Width           =   5775
         Begin VB.CommandButton cmdTampilkanTemp 
            Caption         =   "&Cari"
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
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   5
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
            Format          =   22740995
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
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
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   22740995
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   7
            Top             =   315
            Width           =   255
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgData 
         Height          =   4935
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8705
         _Version        =   393216
         BackColorBkg    =   8421504
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   5520
      Picture         =   "frmSensusPelayanan.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   0
      Picture         =   "frmSensusPelayanan.frx":431A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmSensusPelayanan"
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

Dim subTotal As Currency

Private Sub cmdCetak_Click()
    cmdCetak.Enabled = False
    strSQL = " SELECT * FROM V_D_LaporanSensusPelayanan " & _
             " WHERE TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "'" & _
             " AND KdRuangan = '" & mstrKdRuangan & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbExclamation, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If

    Set frmCetakSensusPelayanan = Nothing
    frmCetakSensusPelayanan.Show
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
    
    subLoadSensusPelayanan
   
    fgData.Visible = True
    Exit Sub
errTampilkan:
    msubPesanError
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdTampilkanTemp.SetFocus
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    On Error GoTo errFormLoad
    
    Me.Caption = "Medifirst2000 - Laporan Sensus Pelayanan"
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
        .Cols = 13
        .Rows = 2
        .ColWidth(0) = 150
        .ColWidth(1) = 1500
        .ColWidth(2) = 1150
        .ColWidth(3) = 800
        .ColWidth(4) = 1350
        .ColWidth(5) = 2050
        .ColWidth(6) = 2050
        .ColWidth(7) = 1200
        .ColWidth(8) = 500
        .ColWidth(9) = 600
        .ColWidth(10) = 600
        .ColWidth(11) = 1000
        .ColWidth(12) = 1450
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter
        .ColAlignment(7) = flexAlignLeftCenter
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
        
        .MergeCol(1) = True
        
        .TextMatrix(0, 1) = "Jenis Pasien"
        .TextMatrix(0, 2) = "No Registrasi"
        .TextMatrix(0, 3) = "No. CM"
        .TextMatrix(0, 4) = "Tgl. Pelayanan"
        .TextMatrix(0, 5) = "Jenis Pelayanan"
        .TextMatrix(0, 6) = "Nama Pelayanan"
        .TextMatrix(0, 7) = "Kelas"
        .TextMatrix(0, 8) = "Jml"
        .TextMatrix(0, 9) = "Tarif"
        .TextMatrix(0, 10) = "CITO"
        .TextMatrix(0, 11) = "Tarif CITO"
        .TextMatrix(0, 12) = "Total"
        .MergeCells = 1
    End With
End Sub

Private Sub subLoadSensusPelayanan()
    
    strSQL = " SELECT JenisPasien,SUM(Total) AS Total " & _
            " FROM V_D_LaporanSensusPelayanan " & _
             " WHERE TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuangan = '" & mstrKdRuangan & "'" & _
             " GROUP BY JenisPasien ORDER BY JenisPasien"
    msubOpenRecFO rsa, strSQL, dbConn
    intJmlRow = intJmlRow + rsa.RecordCount
    
    strSQL = " SELECT * FROM V_D_LaporanSensusPelayanan " & _
             " WHERE TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "'" & _
             " AND KdRuangan = '" & mstrKdRuangan & "' ORDER BY JenisPasien"
    msubOpenRecFO rs, strSQL, dbConn
    
    'jumlah baris keseluruhan
    intJmlRow = intJmlRow + rs.RecordCount
    fgData.Rows = intJmlRow + 2
    intRowNow = 0
    subTotal = 0
    For i = 1 To rs.RecordCount
        intRowNow = intRowNow + 1
        For j = 1 To fgData.Cols - 1
            If IsNull(rs(j - 1).Value) Then
                fgData.TextMatrix(intRowNow, j) = ""
            Else
                fgData.TextMatrix(intRowNow, j) = rs(j - 1).Value
            End If
        Next j
        rs.MoveNext
        'sub total per JenisPasien
        If rs.EOF = True Then GoTo stepSensusPelayanan
        If rs("JenisPasien").Value <> rsa("JenisPasien").Value Then
stepSensusPelayanan:
            intRowNow = intRowNow + 1
            fgData.TextMatrix(intRowNow, 1) = fgData.TextMatrix(intRowNow - 1, 1)
            fgData.TextMatrix(intRowNow, 11) = "Sub Total"
            fgData.TextMatrix(intRowNow, 12) = IIf(rsa("Total").Value = 0, 0, Format(rsa("Total").Value, "#,###"))
            
            subTotal = subTotal + rsa("Total")
            
            subSetSubTotalRow intRowNow, 2, vbBlackness, vbWhite
            If rsa.EOF Then Exit Sub
            rsa.MoveNext
        End If
    Next i
    
    intRowNow = intRowNow + 1
    fgData.TextMatrix(intRowNow, 12) = Format(subTotal, "#,###")
    subSetSubTotalRow intRowNow, 1, vbBlue, vbWhite
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







