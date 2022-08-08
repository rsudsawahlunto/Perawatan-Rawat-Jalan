VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmDaftarPenerimaanKasir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Penerimaan Kasir"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPenerimaanKasir.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   14910
   Begin VB.Frame fraCariPasien 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   7320
      Width           =   14895
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   450
         Left            =   12720
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblJumlahData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data 0/0"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Frame fraDafPasien 
      Caption         =   "Daftar Penerimaan Kasir"
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
      TabIndex        =   8
      Top             =   960
      Width           =   14895
      Begin VB.CheckBox chkSemua 
         Caption         =   "Semua"
         Height          =   375
         Left            =   5640
         TabIndex        =   0
         Top             =   480
         Width           =   975
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
         Left            =   9000
         TabIndex        =   9
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
            TabIndex        =   4
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   2
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   64946179
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   3
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   64946179
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   10
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgData 
         Height          =   5295
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   9340
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   2
         RowHeight       =   16
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
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
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcJenisPasien 
         Height          =   330
         Left            =   6720
         TabIndex        =   1
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Pasien (Cara Bayar)"
         Height          =   210
         Index           =   1
         Left            =   6720
         TabIndex        =   12
         Top             =   240
         Width           =   2010
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   8175
      Width           =   14910
      _ExtentX        =   26300
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   13097
            Text            =   "Cetak Penerimaan Kasir per User (F1)"
            TextSave        =   "Cetak Penerimaan Kasir per User (F1)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   13097
            Text            =   "Cetak Penerimaan Kasir per Shift (Shift + F1)"
            TextSave        =   "Cetak Penerimaan Kasir per Shift (Shift + F1)"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   13080
      Picture         =   "frmDaftarPenerimaanKasir.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarPenerimaanKasir.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarPenerimaanKasir.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmDaftarPenerimaanKasir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkSemua_Click()
    If chkSemua.Value = vbChecked Then dcJenisPasien.Enabled = False Else dcJenisPasien.Enabled = True
End Sub

Private Sub chkSemua_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If chkSemua.Value = vbChecked Then dtpAwal.SetFocus Else dcJenisPasien.SetFocus
End Sub

Private Sub cmdCari_Click()
On Error GoTo errLoad
    
    Call subCariData
    If dgData.ApproxCount = 0 Then chkSemua.SetFocus Else dgData.SetFocus
    
Exit Sub
errLoad:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcJenisPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAwal.SetFocus
End Sub

Private Sub dgData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdTutup.SetFocus
End Sub

Private Sub dgData_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo errLoad
    lblJumlahData.Caption = "Data " & dgData.Bookmark & "/" & dgData.ApproxCount
Exit Sub
errLoad:
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errLoad
Dim strShiftKey As String
    strShiftKey = (Shift + vbShiftMask)
    
    Select Case KeyCode
        Case vbKeyF1
            If strShiftKey = 2 Then mLapPerParameter = "shift"
            Call subCetakPenerimaan
    End Select

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    Call centerForm(Me, MDIUtama)
    dtpAwal.Value = Now
    dtpAkhir.Value = Now
    Call subLoadDCSource
    Call subCariData
    Call PlayFlashMovie(Me)
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subCetakPenerimaan()
On Error GoTo hell
    If Len(dgData.Columns(0).Value) < 1 Then Exit Sub
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value
    FrmCetakPenerimaanKasir.Show
hell:
End Sub

Private Sub subLoadDCSource()
On Error GoTo errLoad

    Call msubDcSource(dcJenisPasien, rs, "SELECT KdKelompokPasien,JenisPasien FROM KelompokPasien")
    If rs.EOF = False Then dcJenisPasien.BoundText = rs(0).Value
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subCariData()
On Error GoTo errLoad

    If chkSemua.Value = vbChecked Then
        strSQL = "SELECT JenisPasien, NoBKM, TglBKM, NoStruk, NoCM, NamaPasien, JK, TotalBiaya, JmlHutangPenjamin, JmlTanggunganRS, JmlHarusDibayar, JmlPembebasan, Administrasi, JmlBayar, SisaTagihan, PembayaranKe, [User]" & _
            " FROM V_LaporanPenerimaanKasKasir" & _
            " WHERE KdRuangan = '" & mstrKdRuangan & "' AND TglBKM between '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "'" & _
            " ORDER BY JenisPasien"
    Else
        strSQL = "SELECT NoBKM, TglBKM, NoStruk, NoCM, NamaPasien, JK, TotalBiaya, JmlHutangPenjamin, JmlTanggunganRS, JmlHarusDibayar, JmlPembebasan, Administrasi, JmlBayar, SisaTagihan, PembayaranKe, [User]" & _
            " FROM V_LaporanPenerimaanKasKasir" & _
            " WHERE KdRuangan = '" & mstrKdRuangan & "' AND JenisPasien = '" & dcJenisPasien.Text & "' AND TglBKM between '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "'"
    End If
    Call msubRecFO(rs, strSQL)
    lblJumlahData.Caption = "Data 0/" & rs.RecordCount
    Set dgData.DataSource = rs
    With dgData
        If chkSemua.Value = vbChecked Then .Columns("JenisPasien").Width = 1150
        .Columns("NoBKM").Width = 1150
        .Columns("TglBKM").Width = 1569
        .Columns("NoStruk").Width = 1150
        .Columns("NoCM").Width = 750
        .Columns("NamaPasien").Width = 2000
        .Columns("JK").Width = 300
        
        .Columns("TotalBiaya").Width = 1500
        .Columns("JmlHutangPenjamin").Width = 1500
        .Columns("JmlTanggunganRS").Width = 1500
        .Columns("JmlHarusDibayar").Width = 1500
        .Columns("JmlPembebasan").Width = 1150
        .Columns("Administrasi").Width = 1150
        .Columns("JmlBayar").Width = 1500
        .Columns("SisaTagihan").Width = 1500
        .Columns("PembayaranKe").Width = 1500
        
        .Columns("NoBKM").Caption = "No. BKM"
        .Columns("TglBKM").Caption = "Tgl. BKM"
        .Columns("NoStruk").Caption = "No. Struk"
        .Columns("TotalBiaya").Caption = "Total Biaya"
        .Columns("JmlHutangPenjamin").Caption = "Hutang Penjamin"
        .Columns("JmlTanggunganRS").Caption = "Tanggungan RS"
        .Columns("JmlHarusDibayar").Caption = "Harus Dibayar"
        .Columns("JmlPembebasan").Caption = "Pembebasan"
        .Columns("Administrasi").Caption = "Administrasi"
        .Columns("JmlBayar").Caption = "Jumlah Bayar"
        .Columns("SisaTagihan").Caption = "Sisa Tagihan"
        .Columns("PembayaranKe").Caption = "Pembayaran Ke"
        
        .Columns("Total Biaya").Alignment = dbgRight
        .Columns("Hutang Penjamin").Alignment = dbgRight
        .Columns("Tanggungan RS").Alignment = dbgRight
        .Columns("Harus Dibayar").Alignment = dbgRight
        .Columns("Pembebasan").Alignment = dbgRight
        .Columns("Administrasi").Alignment = dbgRight
        .Columns("Jumlah Bayar").Alignment = dbgRight
        .Columns("Sisa Tagihan").Alignment = dbgRight
        
        .Columns("Total Biaya").NumberFormat = "#,###.00"
        .Columns("Hutang Penjamin").NumberFormat = "#,###.00"
        .Columns("Tanggungan RS").NumberFormat = "#,###.00"
        .Columns("Harus Dibayar").NumberFormat = "#,###.00"
        .Columns("Pembebasan").NumberFormat = "#,###.00"
        .Columns("Administrasi").NumberFormat = "#,###.00"
        .Columns("Jumlah Bayar").NumberFormat = "#,###.00"
        .Columns("Sisa Tagihan").NumberFormat = "#,###.00"
    End With
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub
