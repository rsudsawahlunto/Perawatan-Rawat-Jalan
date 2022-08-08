VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmKondisiBarangNM2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Kondisi Barang Non Medis"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12990
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmKondisiBarangNM2.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   12990
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   495
      Left            =   7800
      TabIndex        =   2
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   0
      TabIndex        =   11
      Top             =   2640
      Width           =   12975
      Begin VB.TextBox txtCariBarang 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   1
         Top             =   4200
         Width           =   3240
      End
      Begin MSDataGridLib.DataGrid dgKondisiBarang 
         Height          =   3735
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   6588
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   2
         RowHeight       =   15
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
            Locked          =   -1  'True
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Cari Barang"
         Height          =   210
         Index           =   6
         Left            =   255
         TabIndex        =   13
         Top             =   4245
         Width           =   900
      End
      Begin VB.Label lblJmlData 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Jumlah Barang"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   11505
         TabIndex        =   12
         Top             =   4260
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   9480
      TabIndex        =   3
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   11160
      TabIndex        =   4
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Frame fraBarang 
      Height          =   1695
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   12975
      Begin VB.TextBox txtBahan 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6960
         MaxLength       =   50
         TabIndex        =   26
         Top             =   1080
         Width           =   3240
      End
      Begin VB.TextBox txtType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6960
         MaxLength       =   50
         TabIndex        =   24
         Top             =   720
         Width           =   3240
      End
      Begin VB.TextBox txtMerk 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6960
         MaxLength       =   50
         TabIndex        =   21
         Top             =   360
         Width           =   3240
      End
      Begin VB.TextBox txtDetailJenis 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   20
         Top             =   1080
         Width           =   3720
      End
      Begin VB.TextBox txtAsal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   18
         Top             =   720
         Width           =   3720
      End
      Begin VB.TextBox txtKondisi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   11400
         MaxLength       =   25
         TabIndex        =   16
         Top             =   360
         Width           =   1320
      End
      Begin VB.TextBox txtKdBarang 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   240
         MaxLength       =   50
         TabIndex        =   14
         Text            =   "txtkdbarang"
         Top             =   0
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.TextBox txtNamaBarang 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   5
         Top             =   360
         Width           =   3720
      End
      Begin VB.TextBox txtJmlBarang 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   11400
         MaxLength       =   25
         TabIndex        =   6
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Bahan"
         Height          =   210
         Index           =   7
         Left            =   6240
         TabIndex        =   25
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Type"
         Height          =   210
         Index           =   5
         Left            =   6240
         TabIndex        =   23
         Top             =   720
         Width           =   420
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Merk"
         Height          =   210
         Index           =   4
         Left            =   6240
         TabIndex        =   22
         Top             =   360
         Width           =   390
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Detail Jenis Barang"
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Kondisi"
         Height          =   210
         Index           =   8
         Left            =   10680
         TabIndex        =   15
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Nama Barang"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Asal Barang"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   930
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Jumlah"
         Height          =   210
         Index           =   3
         Left            =   10680
         TabIndex        =   8
         Top             =   720
         Width           =   555
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   17
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
      Left            =   11160
      Picture         =   "frmKondisiBarangNM2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmKondisiBarangNM2.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmKondisiBarangNM2.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11295
   End
End
Attribute VB_Name = "frmKondisiBarangNM2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tempbolTampil As Boolean

Private Sub cmdBatal_Click()
    On Error GoTo errLoad

'    Call subKosong
    Call subLoadGridSource
    cmdCetak.SetFocus

    Exit Sub
errLoad:
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell
    If dgKondisiBarang.ApproxCount = 0 Then Exit Sub
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmCetakKondisiBarang.Show
    Exit Sub
hell:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgKondisiBarang_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errLoad

    With dgKondisiBarang
        If .ApproxCount = 0 Then Exit Sub
        txtKdBarang.Text = .Columns("KdBarang")
        txtNamaBarang.Text = .Columns("NamaBarang")
        txtAsal.Text = .Columns("NamaAsal")
        txtDetailJenis.Text = .Columns("DetailJenisBarang")
        txtMerk.Text = .Columns("NamaMerk")
        txtType.Text = .Columns("NamaType")
        txtBahan.Text = .Columns("NamaBahan")
        txtJmlBarang.Text = .Columns("JmlBarang")
        txtKondisi.Text = .Columns("Kondisi")
    End With
    lblJmlData.Caption = dgKondisiBarang.Bookmark & " / " & dgKondisiBarang.ApproxCount & " Data"

    Exit Sub
errLoad:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
'    Call subKosong
    Call subLoadGridSource
    Exit Sub
errLoad:
End Sub

Private Sub txtCariBarang_Change()
    Call subLoadGridSource
End Sub

Private Sub subLoadGridSource()
    On Error GoTo errLoad
    Dim i As Integer

    tempbolTampil = True

    strSQL = "SELECT * " & _
    " FROM V_KondisiBarangNonMedis " & _
    " WHERE kdruangan='" & mstrKdRuangan & "' AND NamaBarang LIKE '%" & txtCariBarang & "%'"

    Call msubRecFO(rs, strSQL)
    Set dgKondisiBarang.DataSource = rs
    With dgKondisiBarang
        For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next i
        .Columns("NamaBarang").Width = 2500
        .Columns("NamaAsal").Width = 1500
        .Columns("DetailJenisBarang").Width = 1500
        .Columns("NamaMerk").Width = 1500
        .Columns("NamaType").Width = 1500
        .Columns("NamaBahanBarang").Width = 1800
        .Columns("Kondisi").Width = 1500
        .Columns("JmlBarang").Width = 1500
    End With
    lblJmlData.Caption = 0 & " / " & dgKondisiBarang.ApproxCount & " Data"
    tempbolTampil = False

    Exit Sub
errLoad:
    Call msubPesanError
End Sub


