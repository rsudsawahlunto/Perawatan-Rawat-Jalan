VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmGrafik 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Grafik"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12600
   Icon            =   "frmGrafik.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   12600
   Begin MSComctlLib.StatusBar sbarStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   8160
      Width           =   12600
      _ExtentX        =   22225
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22172
         EndProperty
      EndProperty
   End
   Begin MSChart20Lib.MSChart mscGrafik 
      Height          =   6015
      Left            =   0
      OleObjectBlob   =   "frmGrafik.frx":08CA
      TabIndex        =   13
      Top             =   960
      Width           =   12495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "X"
      Height          =   375
      Left            =   12960
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame fraControl 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   6960
      Width           =   12495
      Begin VB.ComboBox cboJnsGrafik 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmGrafik.frx":2D82
         Left            =   3600
         List            =   "frmGrafik.frx":2DAA
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "&Kirim ke Excel"
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
         Left            =   10680
         TabIndex        =   9
         Top             =   480
         Width           =   1575
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
         Left            =   9240
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Tampilkan"
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
         Left            =   7800
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   330
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
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
         Format          =   52625409
         CurrentDate     =   37459
      End
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   330
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
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
         Format          =   52625409
         CurrentDate     =   37459
      End
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
         Height          =   375
         Left            =   10800
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Awal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Akhir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1920
         TabIndex        =   6
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipe Grafik"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3600
         TabIndex        =   5
         Top             =   240
         Width           =   870
      End
   End
   Begin VB.Label lblDataPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select a data point to see its value. Double click to change it."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   12
      Top             =   720
      Width           =   6330
   End
   Begin VB.Image Image1 
      Height          =   930
      Left            =   2400
      Picture         =   "frmGrafik.frx":2DD4
      Top             =   0
      Width           =   10200
   End
   Begin VB.Image Image2 
      Height          =   930
      Left            =   0
      Picture         =   "frmGrafik.frx":851C
      Top             =   0
      Width           =   10200
   End
End
Attribute VB_Name = "frmGrafik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Baris As Integer
Dim Kolom As Integer

Private Sub cboJnsGrafik_Click()
    Select Case cboJnsGrafik.ListIndex
    Case 0 To 9
        mscGrafik.ChartType = cboJnsGrafik.ListIndex
    Case 10
        mscGrafik.ChartType = VtChChartType2dPie
    Case 11
        mscGrafik.ChartType = VtChChartType2dXY
    End Select
    If mscGrafik.Chart3d = True Then
        lblDataPoint.Caption = "Hold down the Ctrl key and mouse down to rotate the chart."
    Else
        lblDataPoint.Caption = "Select a point to see it's value"
    End If
End Sub

Private Sub cboJnsGrafik_LostFocus()
    lblDataPoint.Caption = "Select a point to see it's value"
End Sub

Private Sub cmdCetak_Click()
''      MSChart1.EditCopy
''      frmCetak.Picture1.Picture = Clipboard.GetData()
''      Printer.Print " "
''      Printer.PaintPicture frmCetak.Picture1.Picture, 0, 0
''      Printer.EndDoc
''      Unload frmCetak
''MsgBox "Maaf !!! Belum dapat mencetak langsung", vbInformation
'''   With MSChart1
'''      ' Displays a 3d chart with 8 columns and 8 rows
'''      ' data.
'''      .chartType = VtChChartType3dBar
'''      .ColumnCount = 8
'''      .RowCount = 8
'''      For Column = 1 To 8
'''         For Row = 1 To 8
'''            .Column = Column
'''            .Row = Row
'''            .Data = Row * 10
'''         Next Row
'''      Next Column
'''      ' Use the chart as the backdrop of the legend.
'''      .ShowLegend = True
'''      .SelectPart VtChPartTypePlot, index1, index2, _
'''      index3, index4
'''      .EditCopy
'''      .SelectPart VtChPartTypeLegend, index1, _
'''      index2, index3, index4
'''      .EditPaste
'''   End With
End Sub

Private Sub cmdExcel_Click()
    Dim ExcelApp As Object
    Dim ExcelChart As Object
    Dim ChartTypeVal As Integer
    Dim i As Integer
    '-4100 is the value for the MS Excel constant xl3DColumn. Visual
    'Basic does not understand MS Excel constants, so the value must be
    'used instead.
    ChartTypeVal = -4100
    Set ExcelApp = CreateObject("excel.application")
    ExcelApp.Visible = True
    ExcelApp.Workbooks.Add
Dim Baris, Kolom As Integer
Dim x
    Baris = 0
    Do Until Baris > mintJmlBarisGrafik
        Kolom = 1
        Do Until Kolom > mintJmlKolomGrafik + 1
            If Baris = 0 And Kolom <> 1 Then
                x = Chr(65 + Kolom - 1) + Trim(Str(Baris + 1))
                ExcelApp.Range(x).Value = JnsKriteria(Kolom - 1)
            End If
            If Baris <> 0 Then
                x = Chr(65 + Kolom - 1) + Trim(Str(Baris + 1))
                ExcelApp.Range(x).Value = arrGrafik(Baris, Kolom)
            End If
            Kolom = Kolom + 1
        Loop
        Baris = Baris + 1
    Loop
    x = "A1:" + Chr(65 + mintJmlKolomGrafik) + Trim(Str(mintJmlBarisGrafik + 1))
    ExcelApp.Range(x).Select
    Set ExcelChart = ExcelApp.Charts.Add()
    ExcelChart.Type = ChartTypeVal
    For i = 30 To 360 Step 10
        ExcelChart.Rotation = i
    Next
End Sub

Public Sub cmdOK_Click()
    'loading data for the chart
    Call msubLoadDataArray(dtpAwal.Value, dtpAkhir.Value)
    
    'Set MSChart
    Call msubSetChart(mscGrafik, cboJnsGrafik)
    cmdExcel.Enabled = True
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = dtpAkhir.Value
    dtpAkhir.MinDate = dtpAwal.Value
End Sub

Private Sub dtpAkhir_Change()
    dtpAwal.MaxDate = dtpAkhir.Value
    dtpAkhir.MinDate = dtpAwal.Value
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    dtpAwal.Value = Now
    dtpAkhir.Value = dtpAwal.Value
    mscGrafik.Visible = False
    cmdExcel.Enabled = False
    cboJnsGrafik.Clear
    
    ' Configure combobox with chart types.
    With cboJnsGrafik
        .AddItem "3dBar"    ' 0
        .AddItem "2dBar"    ' 1
        .AddItem "3dLine"   ' 2
        .AddItem "2dLine"   ' 3
        .AddItem "3dArea"   ' 4
        .AddItem "2dArea"   ' 5
        .AddItem "3dStep"   ' 6
        .AddItem "2dStep"   ' 7
        .AddItem "3dCombination"    ' 8
        .AddItem "2dCombination"    ' 9
        .AddItem "2dPie"    ' 14
        .AddItem "2dXY"     ' 16
        .ListIndex = 1
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Select Case mstrGrafik
    Case "JenisPasienPerRuangan"
        frmRekapJenisPasien.Enabled = True
    Case "JenisPasienPerJP"
        frmRekapJenisPasien.Enabled = True
    Case "StatusPasienPerRuangan"
        frmRekapStatusPasien.Enabled = True
    Case "StatusPasienPerSP"
        frmRekapStatusPasien.Enabled = True
    End Select
End Sub

Private Sub mscGrafik_LostFocus()
    lblDataPoint.Caption = "Select a point to see it's value. Double-click to change it."
End Sub

Private Sub mscGrafik_PointSelected(Series As Integer, DataPoint As Integer, MouseFlags As Integer, Cancel As Integer)
    ' This allows the user to see the value of any particular data point in a
    '   series by selecting it. The value of the data point is shown in the label
    '   named lblDatapoint.
    With mscGrafik
        .Column = Series
        .Row = DataPoint
        lblDataPoint.Caption = "Pasien " & arrGrafik(DataPoint, 1) & ", Jenis " & JnsKriteria(Series) & " , jumlah pasien = " & .Data
    End With
End Sub

Private Sub mscGrafik_SeriesSelected(Series As Integer, MouseFlags As Integer, Cancel As Integer)
    lblDataPoint.Caption = "Select a point to see it's value. Double-click to change it."
End Sub
