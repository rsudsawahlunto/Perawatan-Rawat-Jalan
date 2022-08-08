VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmJumlahPasiendanJumlahPendapatan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Laporan Jumlah Pasien Dan Jumlah Pendapatan"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmJumlahPasiendanJumlahPendapatan.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   8805
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
      Height          =   1215
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   8775
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   840
         TabIndex        =   7
         Top             =   120
         Width           =   6855
         Begin MSComctlLib.ProgressBar pbData 
            Height          =   615
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   1085
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
            Left            =   6120
            TabIndex        =   9
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   7200
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   495
         Left            =   5520
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
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
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5295
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   240
            TabIndex        =   0
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
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
            CustomFormat    =   "dd MMM yyyy HH:mm "
            Format          =   17367043
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3000
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
            CustomFormat    =   "dd MMM yyyy HH:mm "
            Format          =   17367043
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   2640
            TabIndex        =   4
            Top             =   360
            Width           =   255
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
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
      Height          =   1695
      Left            =   720
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2990
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
      Picture         =   "frmJumlahPasiendanJumlahPendapatan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   6960
      Picture         =   "frmJumlahPasiendanJumlahPendapatan.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmJumlahPasiendanJumlahPendapatan.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmJumlahPasiendanJumlahPendapatan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public iRow As Integer
Public iRows As Integer
Public iCol As Integer
Public iCols As Integer
Dim xFilter As String
Dim rsTmp As New ADODB.recordset
Dim i As Integer
Dim j As Integer
Dim x As Integer
Dim xcount As Integer
Dim xError As String
Dim xCountTgl As Integer
Dim xCountTglKe As Integer
Dim xTglAwal As String
Dim xTglAkhir As String

Private Sub IsiTabelTemp()
Dim i As Integer
Dim sValues As String
Dim sValuesMoney As String
Dim j As Integer
Dim sQuery As String

    Call DeleteTable
    
    iCols = fgData.Cols
    pbData.Max = fgData.Rows - 3
    
        With fgData
        For i = 1 To .Rows - 3

            sValues = ""
            sValuesMoney = ""
            For j = 1 To iCols - 1
                If j = 1 Then
                    sValuesMoney = "'" & .TextMatrix(i, 1) & "'"
                End If
                 If j <> 1 Then
                    sValues = sValues & "," & "'" & .TextMatrix(i, j) & "'"
                    
                End If
            Next j
        
        sQuery = "Insert into V_LapRJTemp" & _
                " values (" & _
                " " & sValuesMoney & "" & _
                " " & sValues & "" & _
                " )"
        dbConn.Execute sQuery
        
        lblPersen.Caption = Int((i / (fgData.Rows - 3)) * 100) & "%"
        pbData.Value = Int(pbData.Value) + 1

        Next i
        pbData.Value = 0.0001
    End With

End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdCetak_Click()
    Frame1.Visible = True
    Call setgrid
    Call getData
    If xError = "" Then
       Call IsiTabelTemp
       Call setgrid
       Call getdatarekap
       Call IsiTabelTemp
       strSQL = "select * from V_LapRJTemp Where NamaKomputer='" & strNamaHostLocal & "'"
       Set frmCetakPendapatan = Nothing
       frmCetakPendapatan.Show
       Frame1.Visible = False
    End If
End Sub

Private Sub Form_Load()
    
    On Error GoTo errFormLoad
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    
    Frame1.Visible = False
    dtpAwal.Value = Format(Now, "dd MMM yyyy 00:00")
    dtpAkhir.Value = Format(Now, "dd MMM yyyy 23:59")
    xError = ""
    
Exit Sub
errFormLoad:
    msubPesanError
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Public Sub getData()
On Error GoTo errorLoad
Dim strFilter As String

    xFilter = " where kdRuangPelayanan = '" & strNKdRuangan & "' and" & _
              " (ThnTglPelayanan between '" & Format(dtpAwal.Value, "yyyy") & "' And '" & Format(dtpAkhir.Value, "yyyy") & "')And" & _
              " (BlnTglPelayanan between '" & Format(dtpAwal.Value, "MM") & "' And '" & Format(dtpAkhir.Value, "MM") & "')And" & _
              " (HariTglPelayanan between '" & Format(dtpAwal.Value, "dd") & "' And '" & Format(dtpAkhir.Value, "dd") & "')"

     Call getData1

Exit Sub
errorLoad:
    Call msubPesanError
End Sub

Public Sub getData1()
On Error GoTo errorLoad

Dim xfilter2 As String

    xcount = 1
    i = 1
    x = 0
    
    For xcount = 1 To 2
        
        xfilter2 = ""
        
        If xcount = 1 Then
        
           xfilter2 = " COUNT(DISTINCT NoCM + NoPendaftaran) AS JmlOS, 0 AS Pendapatan," & _
                      " '1' as JmlPelayanan " & _
                      " From V_StatRJDasarAll " & xFilter & _
                      " GROUP BY ThnTglPelayanan, BlnTglPelayanan, HariTglPelayanan, kdRuangPelayanan, " & _
                      " JmlPelayanan, KdKelompokPasien, IdPenjamin"
        ElseIf xcount = 2 Then
           xfilter2 = " 0 AS JmlOS, Sum (JmlBayar)  AS Pendapatan," & _
                      " '1' JmlPelayanan From V_StatRJDasarAll" & xFilter & _
                      " GROUP BY ThnTglPelayanan, BlnTglPelayanan, HariTglPelayanan, kdRuangPelayanan, " & _
                      " JmlPelayanan, KdKelompokPasien, IdPenjamin"
        End If
    
        strSQL = ""
        strSQL = "SELECT distinct ThnTglPelayanan AS Tahun, BlnTglPelayanan AS Bulan," & _
                 " HariTglPelayanan AS Tanggal, kdRuangPelayanan, IdPenjamin," & _
                 " KdKelompokPasien, " & xfilter2
                 
        Set rsTmp = Nothing
        Call msubRecFO(rsTmp, strSQL)

        If rsTmp.EOF = True Then
           MsgBox "Maaf Data tidak ada.."
           xError = "Error"
           Exit Sub
        Else
           If x = 0 Then
              fgData.Rows = 2 + i
           End If
           loaddata
        End If
      
    Next xcount

Exit Sub
errorLoad:
    Call msubPesanError
End Sub

Public Sub loaddata()
Dim z As Integer

    pbData.Max = rsTmp.RecordCount
    z = 0
    For i = (1 + x) To (rsTmp.RecordCount + x)
        fgData.TextMatrix(i, 1) = rsTmp.Fields("kdRuangPelayanan")
        fgData.TextMatrix(i, 2) = rsTmp.Fields("IdPenjamin")
        fgData.TextMatrix(i, 3) = rsTmp.Fields("kdKelompokPasien")
        fgData.TextMatrix(i, 4) = rsTmp.Fields("JmlOS") * rsTmp.Fields("JmlPelayanan")
        fgData.TextMatrix(i, 5) = rsTmp.Fields("Pendapatan") * rsTmp.Fields("JmlPelayanan")
        fgData.TextMatrix(i, 6) = strNamaHostLocal
        
        z = z + 1
        lblPersen.Caption = Int((z / rsTmp.RecordCount) * 100) & "%"
        pbData.Value = Int(pbData.Value) + 1

        fgData.Rows = fgData.Rows + 1

        rsTmp.MoveNext

    Next i
    x = i - 1
    pbData.Value = 0.0001


End Sub
Private Sub setgrid()
Dim i As Integer
Dim j As Integer
Dim TotalR

    With fgData
        .Clear
        .Rows = 2
        .Cols = 7
               
        .TextMatrix(0, 1) = "kdRuangPelayanan"
        .TextMatrix(0, 2) = "IdPenjamin"
        .TextMatrix(0, 3) = "kdKelompokPasien"
        .TextMatrix(0, 4) = "JmlOS"
        .TextMatrix(0, 5) = "Pendapatan"
        .TextMatrix(0, 6) = "NamaKomputer"
   
        
        .ColWidth(1) = 1500
        .ColWidth(2) = 1500
        .ColWidth(3) = 1500
        .ColWidth(4) = 1500
        .ColWidth(5) = 1500
        .ColWidth(6) = 0
        

    End With
End Sub

Public Sub DeleteTable()
Dim sQuery As String
On Error Resume Next
'    sQuery = "drop Table V_LapLabNo2_NewTemp" & "_" & strNamaHostLocal & ""
    sQuery = "delete V_LapRJTemp Where NamaKomputer= '" & strNamaHostLocal & "'"
    dbConn.Execute sQuery
End Sub

Public Sub getdatarekap()
On Error GoTo errorLoad

 i = 1
 x = 0
 xcount = 0

  For xcount = 1 To 6
    
    If xcount = 1 Then '========== UMUM =========='

       strSQL = ""
       strSQL = " select KdRuangPelayanan, sum(JmlOS) as JmlOS," & _
                " sum(Pendapatan) as Pendapatan" & _
                " from V_LapRJTemp" & _
                " WHERE (KdKelompokPasien = '01') AND NamaKomputer='" & strNamaHostLocal & "'" & _
                " Group by KdRuangPelayanan"
                
    ElseIf xcount = 2 Then '========== ASKES =========='
       strSQL = ""
       strSQL = " select KdRuangPelayanan,sum(JmlOS) as JmlOS," & _
                " sum(Pendapatan) as Pendapatan" & _
                " from V_LapRJTemp" & _
                " WHERE (KdKelompokPasien IN (02, 19, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 39, 40, 41, 42, 43, 44, 45, 46)) AND NamaKomputer='" & strNamaHostLocal & "'" & _
                " Group by KdRuangPelayanan"
    
    ElseIf xcount = 3 Then '========== GAKIN =========='
       strSQL = ""
       strSQL = " select KdRuangPelayanan,sum(JmlOS) as JmlOS," & _
                " sum(Pendapatan) as Pendapatan" & _
                " from V_LapRJTemp" & _
                " WHERE (KdKelompokPasien IN (05, 21, 15, 07)) AND NamaKomputer='" & strNamaHostLocal & "'" & _
                " Group by KdRuangPelayanan"
                
    ElseIf xcount = 4 Then '========== KARYAWAN RSUD KARAWANG =========='
       strSQL = ""
       strSQL = " select KdRuangPelayanan,sum(JmlOS) as JmlOS," & _
                " sum(Pendapatan) as Pendapatan" & _
                " from V_LapRJTemp" & _
                " WHERE (KdKelompokPasien IN (06, 08, 09)) AND NamaKomputer='" & strNamaHostLocal & "'" & _
                " Group by KdRuangPelayanan"
                
    ElseIf xcount = 5 Then '========== KONTRAK  =========='
       strSQL = ""
       strSQL = " select KdRuangPelayanan,sum(JmlOS) as JmlOS," & _
                " sum(Pendapatan) as Pendapatan" & _
                " from V_LapRJTemp" & _
                " WHERE (KdKelompokPasien NOT IN (01, 02, 19, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 39, 40, 41, 42, 43, 44, 45, 46, 05, 21, 15, 07, 06, 08, 09, 14)) AND NamaKomputer='" & strNamaHostLocal & "'" & _
                " Group by KdRuangPelayanan"
                
    Else '========== KOLEGA  =========='
       strSQL = ""
       strSQL = " select KdRuangPelayanan,sum(JmlOS) as JmlOS," & _
                " sum(Pendapatan) as Pendapatan" & _
                " from V_LapRJTemp" & _
                " WHERE (KdKelompokPasien = '14 ') AND NamaKomputer='" & strNamaHostLocal & "'" & _
                " Group by KdRuangPelayanan"
'                " from V_LapLabNo2_NewTemp" & "_" & strNamaHostLocal
    End If
    
    Set rsTmp = Nothing
    Call msubRecFO(rsTmp, strSQL)

'        MsgBox rsTmp.RecordCount
    If rsTmp.EOF = False Then
       Frame1.Visible = True
       lblPersen.Caption = "0 %"
       If x = 0 Then
          fgData.Rows = 2 + i
       End If
       LoadDataRekap
    End If
    
  Next xcount

Exit Sub
errorLoad:
    Call msubPesanError
End Sub

Private Sub LoadDataRekap()
Dim z As Integer

    pbData.Max = rsTmp.RecordCount
    z = 0
    
    For i = (1 + x) To (rsTmp.RecordCount + x)
        fgData.TextMatrix(i, 1) = rsTmp.Fields("kdRuangPelayanan")
        
        If xcount = 1 Then
           fgData.TextMatrix(i, 2) = "UMUM"
        ElseIf xcount = 2 Then
           fgData.TextMatrix(i, 2) = "ASKES"
        ElseIf xcount = 3 Then
           fgData.TextMatrix(i, 2) = "GAKIN"
        ElseIf xcount = 4 Then
           fgData.TextMatrix(i, 2) = "KARYAWAN RSUD"
        ElseIf xcount = 5 Then
           fgData.TextMatrix(i, 2) = "KONTRAK"
        Else
           fgData.TextMatrix(i, 2) = "KOLEGA"
        End If

        fgData.TextMatrix(i, 3) = "99"
        fgData.TextMatrix(i, 4) = rsTmp.Fields("JmlOS")
        fgData.TextMatrix(i, 5) = rsTmp.Fields("Pendapatan")
        fgData.TextMatrix(i, 6) = strNamaHostLocal
        
        z = z + 1
        lblPersen.Caption = Int((z / rsTmp.RecordCount) * 100) & "%"
        pbData.Value = Int(pbData.Value) + 1

        fgData.Rows = fgData.Rows + 1

        rsTmp.MoveNext

    Next i
    x = i - 1
    pbData.Value = 0.0001

End Sub


