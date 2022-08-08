VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRekapJenisPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Rekapitulasi Berdasarkan Jenis Pasien"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9945
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRekapJenisPasien.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   9945
   Begin VB.CommandButton cmdGrafik 
      Caption         =   "Tampilkan &Grafik"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   7320
      Width           =   1455
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
      Left            =   6840
      TabIndex        =   12
      Top             =   7320
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
      Left            =   7845
      TabIndex        =   11
      Top             =   7320
      Width           =   1000
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
      Left            =   8850
      TabIndex        =   10
      Top             =   7320
      Width           =   1000
   End
   Begin VB.Frame fraPerioda 
      Caption         =   "Perioda"
      Height          =   1095
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   9855
      Begin VB.CommandButton cmdTutup 
         Caption         =   "&Tutup"
         Height          =   375
         Left            =   8160
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6660
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdTampilkan 
         Caption         =   "&Tampilkan"
         Height          =   375
         Left            =   5160
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpTglAwal 
         Height          =   330
         Left            =   360
         TabIndex        =   0
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   22740995
         CurrentDate     =   38258
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker dtpTglAkhir 
         Height          =   330
         Left            =   3000
         TabIndex        =   1
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   22740995
         CurrentDate     =   38258
         MinDate         =   2
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "s/d"
         Height          =   195
         Left            =   2520
         TabIndex        =   9
         Top             =   555
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Akhir"
         Height          =   195
         Left            =   3000
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Awal"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   960
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fgData 
      Height          =   5055
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2160
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8916
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Image Image2 
      Height          =   930
      Left            =   -360
      Picture         =   "frmRekapJenisPasien.frx":08CA
      Top             =   0
      Width           =   10200
   End
End
Attribute VB_Name = "frmRekapJenisPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iRowNow As Integer
Dim rsJenisPasien As ADODB.recordset
Dim iRowNow2 As Integer

Private Sub cmdCetak_Click()
    cmdCetak.Enabled = False
    mdTglAwal = dtpTglAwal.Value
    mdTglAkhir = dtpTglAkhir.Value
    strSQL = "SELECT * FROM V_RekapitulasiPasienBJenis " _
        & "WHERE (TglPendaftaran BETWEEN '" _
        & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') ORDER BY Ruangan,JenisPasien"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    mstrCetak = "RekapJenis"
    frmCetak.Show
    frmCetak.Caption = "Medifirst2000 - Rekapitulasi Berdasarkan Jenis Pasien"
    cmdCetak.Enabled = True
End Sub

Private Sub cmdGrafik_Click()
    With frmGrafik
        .Show
        mstrGrafik = "JenisPasienPerJP"
        .Caption = "Medifirst2000 - Grafik Jenis Pasien"
        .dtpAwal.Value = dtpTglAwal.Value
        .dtpAkhir.Value = dtpTglAkhir.Value
        Call .cmdOK_Click
    End With
    Me.Enabled = False
End Sub

Private Sub cmdTampilkan_Click()
Dim intJmlRow As Integer
Dim intJmlPria As Integer
Dim intJmlWanita As Integer
Dim intJmlTotal As Integer
    
    Call subSetGrid
    fgData.Visible = False: MousePointer = vbHourglass

    'Hitung jumlah row dari data yang hendak ditampilkan
    strSQL = "SELECT COUNT(TglPendaftaran) AS JmlRow " & _
        " FROM V_RekapitulasiPasienBJenis " & _
        " WHERE TglPendaftaran BETWEEN " & _
        " '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " & _
        " AND (KdInstalasi = '02' OR KdInstalasi = '06') " & _
        " AND (KdRuangan = '" & mstrKdRuangan & "')"
        
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs(0).Value = 0 Then
        fgData.Visible = True: MousePointer = vbNormal
        MsgBox "Tidak ada data antara tanggal  " & vbCr & "'" & Format(dtpTglAwal.Value, "dd - MMMM - yyyy") & "' dan '" & Format(dtpTglAkhir.Value, "dd - MMMM - yyyy") & "' ", vbInformation, "Validasi"
        txtJmlPria = "0": txtJmlTotal = "0": txtJmlWanita = "0"
        Exit Sub
    End If
    
    intJmlRow = rs("JmlRow").Value
            
    strSQL = "SELECT Ruangan,COUNT(Ruangan) AS JmlJenis, " & _
        " SUM(JmlPasienPria) AS JmlPria,SUM(JmlPasienWanita) AS JmlWanita," & _
        " SUM(Total) AS Total From V_RekapitulasiPasienBJenis " & _
        " WHERE TglPendaftaran BETWEEN " & _
        " '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " & _
        " AND (KdInstalasi ='02' OR KdInstalasi = '06') " & _
        " AND (KdRuangan = '" & mstrKdRuangan & "')" & _
        " GROUP BY Ruangan"
        
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    'Tambahkan jumlah row dengan jumlah subtotal
    intJmlRow = intJmlRow + rs.RecordCount
    
    Dim rsJenisPasien As New ADODB.recordset
    With fgData
        .Rows = intJmlRow + 2
        While rs.EOF = False
            strSQL = "SELECT Ruangan,JenisPasien,SUM(JmlPasienPria) AS TJmlPasienPria, " & _
                " SUM(JmlPasienWanita) AS TJmlPasienWanita, " & _
                " SUM(Total) AS TTotal From V_RekapitulasiPasienBJenis " & _
                " WHERE (TglPendaftaran BETWEEN " & _
                " '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND " & _
                " '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "') " & _
                " AND Ruangan='" & rs("Ruangan").Value & "' " & _
                " AND (KdInstalasi = '02' OR KdInstalasi ='06')" & _
                " GROUP BY Ruangan, JenisPasien " & _
                " ORDER BY JenisPasien"
            
            Set rsJenisPasien = Nothing
            rsJenisPasien.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            While rsJenisPasien.EOF = False
                iRowNow = iRowNow + 1
                .TextMatrix(iRowNow, 1) = rsJenisPasien("Ruangan").Value
                .TextMatrix(iRowNow, 2) = rsJenisPasien("JenisPasien").Value
                .TextMatrix(iRowNow, 3) = rsJenisPasien("TJmlPasienPria").Value
                .TextMatrix(iRowNow, 4) = rsJenisPasien("TJmlPasienWanita").Value
                .TextMatrix(iRowNow, 5) = rsJenisPasien("TTotal").Value
                rsJenisPasien.MoveNext
            Wend
            iRowNow = iRowNow + 1
            For i = 1 To .Cols - 1
                .Col = i
                .Row = iRowNow
                .CellBackColor = vbBlackness
                .CellForeColor = vbWhite
                If .Col = 1 Then .TextMatrix(.Row, 1) = .TextMatrix(.Row - 1, 1): .CellBackColor = vbWhite: .CellForeColor = vbBlack
                .RowHeight(.Row) = 300
                .CellFontBold = True
            Next
            
            intJmlPria = intJmlPria + rs("JmlPria").Value
            intJmlWanita = intJmlWanita + rs("JmlWanita").Value
            intJmlTotal = intJmlTotal + rs("Total").Value
            .TextMatrix(iRowNow, 2) = "Sub Total"
            .TextMatrix(iRowNow, 3) = rs("JmlPria").Value
            .TextMatrix(iRowNow, 4) = rs("JmlWanita").Value
            .TextMatrix(iRowNow, 5) = rs("Total").Value
            .Font.Bold = False
            rs.MoveNext
        Wend
        .Rows = iRowNow + 2
    End With
    txtJmlPria.Text = intJmlPria
    txtJmlWanita.Text = intJmlWanita
    txtJmlTotal.Text = intJmlTotal
    
    fgData.Col = 1
    For i = 1 To fgData.Rows - 1
        fgData.Row = i
        fgData.CellFontBold = True
    Next

    fgData.Visible = True: MousePointer = vbNormal
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    txtJmlPria = "0": txtJmlTotal = "0": txtJmlWanita = "0"
    
    With dtpTglAwal
        .Month = Month(Now)
        .Day = 1
        .Year = Year(Now)
    End With
    dtpTglAkhir.Value = Now
    
    Call subSetGrid
    Call SetText
End Sub

Private Sub subSetGrid()
    With fgData
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
        .TextMatrix(0, 2) = "Jenis Pasien"
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
    With fgData
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
