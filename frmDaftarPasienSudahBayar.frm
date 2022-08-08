VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmDaftarPasienSudahBayar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Pasien Sudah Bayar"
   ClientHeight    =   8520
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
   Icon            =   "frmDaftarPasienSudahBayar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   14910
   Begin VB.Frame fraCariPasien 
      Caption         =   "Cari Data Pasien"
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
      TabIndex        =   8
      Top             =   7320
      Width           =   14895
      Begin VB.CommandButton cmdBayarUlang 
         Caption         =   "&Bayar Ulang"
         Height          =   450
         Left            =   8400
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdBatalKuitansi 
         Caption         =   "&Retur Kuitansi"
         Height          =   450
         Left            =   10560
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   450
         Left            =   12720
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   400
         Width           =   2895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan Nama Pasien /  No.CM"
         Height          =   210
         Left            =   1560
         TabIndex        =   9
         Top             =   165
         Width           =   2640
      End
   End
   Begin VB.Frame fraDafPasien 
      Caption         =   "Daftar Pasien Sudah Bayar"
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
      TabIndex        =   10
      Top             =   960
      Width           =   14895
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
         TabIndex        =   11
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
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   20381699
            UpDown          =   -1  'True
            CurrentDate     =   38212
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
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   20381699
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   12
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgDaftarPasienSudahBayar 
         Height          =   5295
         Left            =   120
         TabIndex        =   3
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
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   8145
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
            Text            =   "Cetak Detail Kuitansi (F1)"
            TextSave        =   "Cetak Detail Kuitansi (F1)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   13097
            Text            =   "Cetak Kuitansi (F9)"
            TextSave        =   "Cetak Kuitansi (F9)"
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
      Picture         =   "frmDaftarPasienSudahBayar.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarPasienSudahBayar.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarPasienSudahBayar.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13215
   End
End
Attribute VB_Name = "frmDaftarPasienSudahBayar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dTglMasuk As Date

Private Sub cmdBatalKuitansi_Click()
On Error GoTo hell
    If dgDaftarPasienSudahBayar.ApproxCount = 0 Then Exit Sub
    
    strSQL = " SELECT NoBKM FROM V_JudulReturStrukPelayananPasien WHERE NoBKM = '" & dgDaftarPasienSudahBayar.Columns("No. BKM").Value & "'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        MsgBox "No. BKM " & dgDaftarPasienSudahBayar.Columns("No. BKM").Value & " sudah diretur", vbInformation, "Informasi"
        Exit Sub
    End If
    frmReturStrukPelayananPasien.Show
    frmReturStrukPelayananPasien.txtNoBKM.Text = dgDaftarPasienSudahBayar.Columns("No. BKM").Value
    Call frmReturStrukPelayananPasien.txtNoBKM_KeyPress(13)
hell:
End Sub

Private Sub cmdBayarUlang_Click()
On Error GoTo errload
Dim tempPembayaranKe As Integer

    If dgDaftarPasienSudahBayar.ApproxCount = 0 Then Exit Sub
    If MsgBox("Anda yakin akan membayar ulang", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "SELECT PembayaranKe FROM PembayaranTagihanPasien WHERE (NoStruk = '" & dgDaftarPasienSudahBayar.Columns("No. Struk") & "') AND (NoBKM = '" & dgDaftarPasienSudahBayar.Columns("No. BKM") & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then tempPembayaranKe = rs(0) Else tempPembayaranKe = 0

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoStruk", adChar, adParamInput, 10, dgDaftarPasienSudahBayar.Columns("No. Struk"))
        .Parameters.Append .CreateParameter("NoBKM", adChar, adParamInput, 10, dgDaftarPasienSudahBayar.Columns("No. BKM"))
        .Parameters.Append .CreateParameter("PembayaranKe", adTinyInt, adParamInput, , tempPembayaranKe)
            
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_BayarUlangKasir"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Add_BayarUlangKasir")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Call cmdCari_Click

Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub cmdCari_Click()
On Error GoTo errload

    mstrFilter = ""
    Set rs = Nothing
    If dtpAwal.Day <> dtpAkhir.Day Or dtpAwal.Month <> dtpAkhir.Month Or dtpAwal.Year <> dtpAkhir.Year Then
        rs.Open "select top 100 * from V_DaftarPasienYgSudahBayar where (NamaPasien like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%') AND KdInstalasi = '" & mstrKdInstalasiLogin & "' and TglBKM between '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " & mstrFilter & "", dbConn, adOpenStatic, adLockOptimistic
    Else
        rs.Open "select * from V_DaftarPasienYgSudahBayar where (NamaPasien like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%') AND KdInstalasi = '" & mstrKdInstalasiLogin & "' and TglBKM between '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " & mstrFilter & "", dbConn, adOpenStatic, adLockOptimistic
    End If
    Set dgDaftarPasienSudahBayar.DataSource = rs
    Call SetGridPasienSudahBayar
    If dgDaftarPasienSudahBayar.ApproxCount = 0 Then dtpAwal.SetFocus Else dgDaftarPasienSudahBayar.SetFocus
    
Exit Sub
errload:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgDaftarPasienSudahBayar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdBatalKuitansi.SetFocus
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

Private Sub Form_Activate()
    cmdCari_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errload
Dim strShiftKey As String

    strShiftKey = (Shift + vbShiftMask)

    Select Case KeyCode
        Case vbKeyF1
            If dgDaftarPasienSudahBayar.ApproxCount = 0 Then Exit Sub
            If strShiftKey = 2 Then
            Else
                Call subCetakDetailKuitansi
            End If
        Case vbKeyF9
            If dgDaftarPasienSudahBayar.ApproxCount = 0 Then Exit Sub
            Call subCetakKuitansi
    End Select

Exit Sub
errload:
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call centerForm(Me, MDIUtama)
    dtpAwal.Value = Now
    dtpAkhir.Value = Now
    Set rs = Nothing
    strQuery = "select * from V_DaftarPasienYgSudahBayar where TglBKM between '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "'"
    rs.Open strQuery, dbConn, adOpenStatic, adLockOptimistic
    Set dgDaftarPasienSudahBayar.DataSource = rs
    Call SetGridPasienSudahBayar
    
    If mblnAdmin = False Then
        cmdBatalKuitansi.Enabled = False
        cmdBayarUlang.Enabled = False
    Else
        cmdBatalKuitansi.Enabled = True
        cmdBayarUlang.Enabled = True
    End If
    Call PlayFlashMovie(Me)
End Sub

Sub SetGridPasienSudahBayar()
    With dgDaftarPasienSudahBayar
        .Columns(0).Width = 1150
        .Columns(0).Caption = "No. BKM"
        .Columns(1).Width = 1150
        .Columns(1).Caption = "No. Struk"
        .Columns(2).Width = 750
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Width = 2000
        .Columns(4).Width = 300
        .Columns(5).Width = 1600
        .Columns(6).Width = 1300
        .Columns(7).Width = 1569
        .Columns(8).Width = 1569
        .Columns(9).Width = 0
        .Columns(10).Width = 2700
        .Columns(11).Width = 0
        .Columns(12).Width = 0
        .Columns("KdInstalasi").Width = 0
        .Columns("KdRuangan").Width = 0
    End With
End Sub

Private Sub subCetakDetailKuitansi()
On Error GoTo hell
    If Len(dgDaftarPasienSudahBayar.Columns(0).Value) < 1 Then Exit Sub
    vLaporan = "" '"Print"
    mstrNoStruk = dgDaftarPasienSudahBayar.Columns("No. Struk").Value
    mstrNoBKM = dgDaftarPasienSudahBayar.Columns("No. BKM").Value
    frmCetak.CetakUlang
hell:
End Sub

Private Sub subCetakKuitansi()
On Error GoTo hell
    If Len(dgDaftarPasienSudahBayar.Columns(0).Value) < 1 Then Exit Sub
    vLaporan = "" '"Print"
    mstrNoStruk = dgDaftarPasienSudahBayar.Columns("No. Struk").Value
    mstrNoBKM = dgDaftarPasienSudahBayar.Columns("No. BKM").Value
    frmCetak.CetakUlangJenisKuitansi
hell:
End Sub

Private Sub txtParameter_Change()
    Call cmdCari_Click
    txtParameter.SetFocus
    txtParameter.SelStart = Len(txtParameter.Text)
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then If dgDaftarPasienSudahBayar.ApproxCount = 0 Then dtpAwal.SetFocus Else dgDaftarPasienSudahBayar.SetFocus
End Sub

Private Sub txtParameter_LostFocus()
    txtParameter.Text = StrConv(txtParameter.Text, vbProperCase)
End Sub
