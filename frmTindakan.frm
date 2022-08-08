VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTindakan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Pelayanan Tindakan"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTindakan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   11805
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9240
      Top             =   2520
   End
   Begin MSComctlLib.ListView lvPemeriksa 
      Height          =   1815
      Left            =   4080
      TabIndex        =   22
      Top             =   -840
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3201
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nama Pemeriksa"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame fraPelayanan 
      Caption         =   "Data Pelayanan Tindakan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   8880
      TabIndex        =   24
      Top             =   5400
      Visible         =   0   'False
      Width           =   9855
      Begin MSDataGridLib.DataGrid dgPelayanan 
         Height          =   2415
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   4260
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
   Begin VB.Frame fraDokter 
      Caption         =   "Data Dokter Pemeriksa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   8880
      TabIndex        =   21
      Top             =   3960
      Visible         =   0   'False
      Width           =   8895
      Begin MSDataGridLib.DataGrid dgDokter 
         Height          =   2295
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4048
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
   Begin VB.Frame fradoa 
      Caption         =   "Daftar Layanan Obat && Alkes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   25
      Top             =   5880
      Width           =   11775
      Begin MSFlexGridLib.MSFlexGrid fgDOA 
         Height          =   1335
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   2355
         _Version        =   393216
         Rows            =   50
         Cols            =   10
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   8577768
         ForeColorFixed  =   -2147483627
         ForeColorSel    =   -2147483628
         BackColorBkg    =   16777215
         FocusRect       =   0
         HighLight       =   2
         FillStyle       =   1
         GridLines       =   3
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Daftar Layanan Tindakan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   27
      Top             =   3840
      Width           =   11775
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgPelayanan 
         Height          =   1575
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   2778
         _Version        =   393216
         Rows            =   50
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   8577768
         BackColorBkg    =   16777215
         FocusRect       =   0
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Frame fraButton 
      Enabled         =   0   'False
      Height          =   735
      Left            =   0
      TabIndex        =   28
      Top             =   3120
      Width           =   11775
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   360
         Left            =   6960
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdTambah 
         Caption         =   "&Tambah"
         Height          =   360
         Left            =   5760
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   360
         Left            =   8160
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "Tutu&p"
         Height          =   360
         Left            =   9360
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraPPelayanan 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   29
      Top             =   2160
      Width           =   11775
      Begin VB.OptionButton optNonPaket 
         Caption         =   " Non Paket"
         Height          =   375
         Left            =   5880
         TabIndex        =   14
         Top             =   550
         Width           =   1215
      End
      Begin VB.OptionButton optPaket 
         Caption         =   " Paket"
         Height          =   375
         Left            =   5880
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtKuantitas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   5040
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "1"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtNamaPelayanan 
         Appearance      =   0  'Flat
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
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   4695
      End
      Begin VB.CheckBox chkAPBD 
         Caption         =   "Pos APBD"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   518
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo dcStatusCito 
         Height          =   360
         Left            =   7320
         TabIndex        =   36
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Jumlah"
         Height          =   240
         Left            =   5040
         TabIndex        =   31
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pelayanan"
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   1425
      End
   End
   Begin VB.Frame fraPDokter 
      Height          =   1095
      Left            =   0
      TabIndex        =   32
      Top             =   1080
      Width           =   11775
      Begin VB.TextBox txtDokter2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   4840
         TabIndex        =   5
         Top             =   525
         Width           =   2655
      End
      Begin VB.CheckBox chkDelegasi 
         Caption         =   "Di Delegasikan"
         Height          =   255
         Left            =   4800
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Caption         =   "Status CITO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9960
         TabIndex        =   33
         Top             =   240
         Width           =   1695
         Begin VB.OptionButton optCito 
            Caption         =   "Ya"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optCito 
            Caption         =   "Tidak"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.CheckBox chkPerawat 
         Caption         =   "Paramedis"
         Height          =   255
         Left            =   7600
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtDokter 
         Appearance      =   0  'Flat
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
         Left            =   2400
         TabIndex        =   2
         Top             =   525
         Width           =   2415
      End
      Begin VB.TextBox txtNamaPerawat 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   7600
         TabIndex        =   7
         Text            =   "txtNamaPerawat"
         Top             =   525
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker dtpTglPeriksa 
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   525
         Width           =   2175
         _ExtentX        =   3836
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
         CustomFormat    =   "dd/MM/yyyy HH:mm:ss"
         Format          =   182845443
         UpDown          =   -1  'True
         CurrentDate     =   37823
      End
      Begin VB.CheckBox chkDilayaniDokter 
         Caption         =   "Dokter Pemeriksa "
         Height          =   255
         Left            =   2400
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Periksa"
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1365
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fgPerawatPerPelayanan 
      Height          =   1215
      Left            =   5400
      TabIndex        =   26
      Top             =   4440
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2143
      _Version        =   393216
      FixedCols       =   0
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   35
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
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmTindakan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   9960
      Picture         =   "frmTindakan.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmTindakan.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "frmTindakan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BolPesanTindakan As Boolean
Dim strFilterDokter As String
Dim strFilterPelayanan As String
Dim strCito As String
Dim strKdKelas As String
Dim strKelas As String
Dim strJenisTarif As String
Dim strKdJenisTarif As String
Dim intJmlPelayanan As Integer
Dim strKodePelayananRS As String
Dim curBiaya As Currency
Dim curJP As Currency
Dim intRowNow As Integer
Dim penamwaktu As Date

Dim intBarang As Integer
Dim intJmlBarang As Integer
Dim intMaxJmlBarang As Integer
Dim strPilihGrid As String
Dim mstrKdDokter2 As String

Dim strStatusAPBD As String

Dim subKdPemeriksa() As String
Dim subJmlTotal As Integer
Dim i As Integer
Dim j As Integer
Dim curTarifCito As Currency
Dim subcurTarifCito As Currency
Dim subcurTarifBiayaSatuan As Currency
Dim subcurTarifHargaSatuan As Currency
Dim bolPasienKonsul As Boolean

Private Function sp_DelegasiBiayaPelayanan(f_NoPendaftaran As String, f_KdRuangan As String, f_KdPelayananRS As String, f_TglPelayanan As Date, f_StatusDelegasi As String) As Boolean
    On Error GoTo errLoad

    sp_DelegasiBiayaPelayanan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, f_KdPelayananRS)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(f_TglPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("StatusDelegasi", adChar, adParamInput, 1, f_StatusDelegasi)
        'RJ tidak ada dokter delegasi. asumsi dokter delegasi = dokter jaga/ pengganti bukan dokter yg seharusnya
        .Parameters.Append .CreateParameter("IdDokterDelegasi", adChar, adParamInput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_DelegasiBiayaPelayanan"
        .CommandType = adCmdStoredProc
        .Execute
    End With

    Exit Function
errLoad:
    sp_DelegasiBiayaPelayanan = False
    Call msubPesanError("sp_DelegasiBiayaPelayanan")
End Function

Private Sub chkAPBD_Click()
    If chkAPBD.Value = 1 Then
        strStatusAPBD = "01"
    Else
        strStatusAPBD = "02"
    End If
End Sub

Private Sub chkAPBD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaPelayanan.SetFocus
End Sub

Private Sub chkDelegasi_Click()
    If chkDelegasi.Value = vbChecked Then
        If MsgBox("Akan Didelegasikan Ke Dokter Atau Paramedis ?? " & vbCrLf & "Pilih YES Untuk DOKTER atau Pilih NO Untuk PARAMEDIS ", vbYesNo, "Validasi") = vbYes Then
            chkPerawat.Value = vbUnchecked
            chkPerawat.Enabled = False
            txtDokter2.Enabled = True
            lvPemeriksa.Enabled = False
        Else
            chkPerawat.Value = vbChecked
            chkPerawat.Enabled = True
            txtDokter2.Enabled = False
            lvPemeriksa.Enabled = True
        End If
    Else
        chkPerawat.Value = vbChecked
        chkPerawat.Enabled = True
        txtDokter2.Enabled = False
        lvPemeriksa.Enabled = True
    End If

End Sub

Private Sub chkDelegasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If chkPerawat.Enabled = True Then chkPerawat.SetFocus Else txtNamaPelayanan.SetFocus
     End If
End Sub

Private Sub chkDilayaniDokter_Click()
    On Error GoTo errLoad

    If chkDilayaniDokter.Value = 0 Then
        txtDokter.Enabled = False
        txtDokter.Text = ""

        If fraDokter.Visible = True Then fraDokter.Visible = False
    Else
        lvPemeriksa.Visible = False

        txtDokter.Enabled = True
        strSQL = "SELECT dbo.RegistrasiRJ.IdDokter, dbo.DataPegawai.NamaLengkap " & _
        " FROM dbo.RegistrasiRJ INNER JOIN dbo.DataPegawai ON dbo.RegistrasiRJ.IdDokter = dbo.DataPegawai.IdPegawai " & _
        " WHERE (dbo.RegistrasiRJ.NoPendaftaran = '" & mstrNoPen & "')"
        Call msubRecFO(rs, strSQL)

        If Not rs.EOF Then
            txtDokter.Text = rs(1).Value
            mstrKdDokter = rs(0).Value
            intJmlDokter = rs.RecordCount
            fraDokter.Visible = False
        End If
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub chkDilayaniDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkDilayaniDokter.Value = 0 Then
            chkPerawat.SetFocus
        Else
            txtDokter.SetFocus
        End If
    End If
End Sub

Private Sub chkPerawat_Click()
    If chkPerawat.Value = vbChecked Then
        strSQL = "SELECT IdPegawai FROM V_DaftarPemeriksaPasien WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
            txtNamaPerawat.Text = strNmPegawai
            If lvPemeriksa.ListItems.Count > 0 Then
                lvPemeriksa.ListItems.Item("key" & strIDPegawaiAktif).Checked = True
                Call lvPemeriksa_ItemCheck(lvPemeriksa.ListItems.Item("key" & strIDPegawaiAktif))
            End If
        Else
            txtNamaPerawat.Text = ""
        End If
    Else
        txtNamaPerawat.Text = ""
    End If
    lvPemeriksa.Visible = False
End Sub

Private Sub chkPerawat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkPerawat.Value = vbChecked Then
            txtNamaPerawat.SetFocus
        Else
            optCito(1).SetFocus
        End If
    End If
End Sub

Private Sub cmdBatal_Click()
    If cmdSimpan.Enabled = True Then
        If MsgBox("Simpan data tindakan pasien?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
            Exit Sub
        End If
    End If
    Unload Me
    bolPasienKonsul = False
    frmTransaksiPasien.Enabled = True
End Sub

Private Sub cmdHapus_Click()
    Dim h As Integer
    Dim j As Integer
    With fgPelayanan
        If .Row = .Rows Then Exit Sub
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        h = 1
        Do While h <= fgDOA.Rows - 2
            If fgDOA.TextMatrix(h, 9) = .TextMatrix(.Row, 0) Then
                For j = 1 To intMaxJmlBarang
                    If typBarang(j).strkdbarang = fgDOA.TextMatrix(h, 0) Then
                        If fgDOA.TextMatrix(h, 5) = "S" Then
                            typBarang(j).intJmlTempTotal = typBarang(j).intJmlTempTotal + (fgDOA.TextMatrix(h, 3) * typBarang(j).intJmlTerkecil)
                        Else
                            typBarang(j).intJmlTempTotal = typBarang(j).intJmlTempTotal + fgDOA.TextMatrix(h, 3)
                        End If
                    End If
                Next j
                Call msubRemoveItem(fgDOA, h)
                h = 0
            End If
            h = h + 1
        Loop
        For j = 1 To intMaxJmlBarang
            For h = 1 To fgDOA.Rows - 1
                If typBarang(j).strkdbarang = fgDOA.TextMatrix(h, 0) Then Exit For
                If h = fgDOA.Rows - 1 Then
                    intMaxJmlBarang = intMaxJmlBarang - 1
                    If intMaxJmlBarang < 0 Then intMaxJmlBarang = 0
                End If
            Next h
        Next j
        Call msubRemoveItem(fgPelayanan, .Row)
        strKodePelayananRS = ""
    End With
End Sub

Private Sub cmdSimpan_Click()
    Dim i As Integer
    Dim sisaKSO As String
    Dim TotKSO As String
    
    If funcCekValidasi = False Then Exit Sub
    Call subEnableButtonReg(False)
    For i = 1 To fgPelayanan.Rows - 2
        strSQL = "SELECT NoPendaftaran From BiayaPelayanan Where (KdRuangan = '" & mstrKdRuangan & "') And (KdPelayananRS = '" & fgPelayanan.TextMatrix(i, 0) & "') And (TglPelayanan = '" & Format(fgPelayanan.TextMatrix(i, 9), "yyyy/MM/dd HH:mm:ss") & "')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
            MsgBox "Tindakan tersebut sudah pernah diinputkan dengan waktu yang sama." & vbNewLine & " Rubah tanggal pelayanan", vbExclamation, "Validasi"
            dtpTglPeriksa.SetFocus
            Exit Sub
            
        End If
    Next i

    For i = 1 To fgPelayanan.Rows - 2
        
      If i = 1 Then
        penamwaktu = fgPelayanan.TextMatrix(i, 9)
      Else
        penamwaktu = penamwaktu
      End If
      
        'If sp_BiayaPelayanan(dbcmd, fgPelayanan.TextMatrix(i, 0), fgPelayanan.TextMatrix(i, 3), fgPelayanan.TextMatrix(i, 2), fgPelayanan.TextMatrix(i, 9), fgPelayanan.TextMatrix(i, 6), fgPelayanan.TextMatrix(i, 7), fgPelayanan.TextMatrix(i, 8)) = False Then Exit Sub
        If sp_BiayaPelayanan(dbcmd, fgPelayanan.TextMatrix(i, 0), fgPelayanan.TextMatrix(i, 3), fgPelayanan.TextMatrix(i, 2), penamwaktu, fgPelayanan.TextMatrix(i, 6), fgPelayanan.TextMatrix(i, 7), fgPelayanan.TextMatrix(i, 8)) = False Then Exit Sub
        
        '@dimas 2014-05-15
        '-------------------------------------------------------------
        strSQLx = "Select KdPelayananRS from SettingPelayananKSO where KdInstalasi='02' and StatusPelayanan='TM'"
        Call msubRecFO(rsx, strSQLx)
        
        strSQL = "Select sum((JmlPelayanan * Tarif)) from DetailBiayaPelayanan where NoPendaftaran='" & mstrNoPen & "' and KdPelayananRS= '" & rsx(0).Value & "' and NoStruk is null"
        Call msubRecFO(rs, strSQL)
        
        TotKSO = Val(IIf(IsNull(rs(0).Value), 0, rs(0).Value))
        sisaKSO = TotKSO - 35000
        If TotKSO > 35000 Then
            If update_BiayaPelayananKSO(dbcmd, rsx(0).Value) = False Then Exit Sub
        End If
        '-------------------------------------------------------------
        
        'jika 1 maka dr pesan pelayanan, jika selain 1 maka bukan dari pesan pelayanan
        If fgPelayanan.TextMatrix(i, 11) = "1" Then
            If chkDelegasi.Value = vbChecked Then
                If sp_DelegasiBiayaPelayanan(mstrNoPen, mstrKdRuangan, fgPelayanan.TextMatrix(i, 0), fgPelayanan.TextMatrix(i, 9), "Y") = False Then Exit Sub
            End If
            If update_DetailOrderTMOA(dbcmd, fgPelayanan.TextMatrix(i, 0), "TM") = False Then Exit Sub
        Else
            If chkDelegasi.Value = vbChecked Then
                If sp_DelegasiBiayaPelayanan(mstrNoPen, mstrKdRuangan, fgPelayanan.TextMatrix(i, 0), fgPelayanan.TextMatrix(i, 9), IIf(fgPelayanan.TextMatrix(i, 10) = "1", "Y", "T")) = False Then Exit Sub
            End If
        End If
        
    Next i
    
    If chkPerawat.Value = Checked Then
        For i = 1 To fgPerawatPerPelayanan.Rows - 1
            With fgPerawatPerPelayanan
                If sp_PetugasPemeriksaBP(.TextMatrix(i, 2), .TextMatrix(i, 3), .TextMatrix(i, 4), .TextMatrix(i, 5)) = False Then Exit Sub
            End With
        Next i
    End If

    Dim adoCommand As New ADODB.Command
    If fgDOA.Rows = 2 Then GoTo stepNonPaketSemua
    For i = 1 To fgDOA.Rows - 2
        With adoCommand

            .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
            .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, fgDOA.TextMatrix(i, 0))
            .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, fgDOA.TextMatrix(i, 2))
            .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
            .Parameters.Append .CreateParameter("Satuan", adChar, adParamInput, 1, fgDOA.TextMatrix(i, 5))
            .Parameters.Append .CreateParameter("JmlBrg", adDouble, adParamInput, , fgDOA.TextMatrix(i, 3))
            .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
            .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
            .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, mstrKdKelas)
            .Parameters.Append .CreateParameter("HargaSatuan", adCurrency, adParamInput, , CCur(fgDOA.TextMatrix(i, 4)))
            .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(fgDOA.TextMatrix(i, 7), "yyyy/MM/dd HH:mm:ss"))
            .Parameters.Append .CreateParameter("NoLabRad", adChar, adParamInput, 10, Null)
            .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, fgDOA.TextMatrix(i, 6))
            .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
            .Parameters.Append .CreateParameter("IdPegawai2", adChar, adParamInput, 10, Null)

            .Parameters.Append .CreateParameter("KdJenisObat", adChar, adParamInput, 2, Null)
            .Parameters.Append .CreateParameter("JmlService", adInteger, adParamInput, , 0)
            .Parameters.Append .CreateParameter("TarifService", adCurrency, adParamInput, , 0)
            .Parameters.Append .CreateParameter("NoResep", adVarChar, adParamInput, 15, Null)
            .Parameters.Append .CreateParameter("Rke", adInteger, adParamInput, , 0)
            .Parameters.Append .CreateParameter("StatusStok", adChar, adParamInput, 1, "1")

            strSQL = "Select KdRuangan from Pendaftaran Where NoPendaftaran = '" & mstrNoPen & "'"
            Set rs = Nothing
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            If rs.EOF = True Then
                .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, mstrKdRuangan)
            Else
                .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, IIf(IsNull(rs("KdRuangan")), mstrKdRuangan, Trim(rs("KdRuangan"))))
            End If

            .Parameters.Append .CreateParameter("KdPelayananRSUsed", adChar, adParamInput, 6, IIf(Trim(fgDOA.TextMatrix(i, 9)) = "", Null, Trim(fgDOA.TextMatrix(i, 9))))
            .Parameters.Append .CreateParameter("KdStatusHasil", adChar, adParamInput, 2, Null)
            .Parameters.Append .CreateParameter("JmlExpose", adInteger, adParamInput, , Null)
            .Parameters.Append .CreateParameter("KdStatusKontras", adInteger, adParamInput, , Null)
            .Parameters.Append .CreateParameter("IdPenanggungjawab", adChar, adParamInput, 1, Null)
            .Parameters.Append .CreateParameter("Keterangan", adChar, adParamInput, 3, Null)
            .Parameters.Append .CreateParameter("NoTerima", adChar, adParamInput, 10, fgDOA.TextMatrix(i, 10))

            .ActiveConnection = dbConn
            .CommandText = "dbo.Add_PemakaianObatAlkesResepNew"
            .CommandType = adCmdStoredProc
            .Execute

            If Not (.Parameters("RETURN_VALUE").Value = 0) Then
                MsgBox "Ada Kesalahan dalam Penyimpanan Biaya Pelayanan Pasien", vbCritical, "Validasi"
                Call deleteADOCommandParameters(adoCommand)
                Set adoCommand = Nothing
                GoTo stepErrorPaket
            End If
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
        End With
    Next i

    Call Add_HistoryLoginActivity("Add_BiayaPelayanan+Add_DelegasiBiayaPelayanan+Add_PetugasPemeriksaBP+Add_PemakaianObatAlkes")
stepNonPaketSemua:
stepErrorPaket:
    frmTransaksiPasien.subLoadPelayananDidapat
    frmTransaksiPasien.subPemakaianObatAlkes
    frmTransaksiPasien.subLoadRiwayatPemeriksaan False
End Sub

Private Sub cmdTambah_Click()
    Dim i As Integer
    Dim j As Integer
    Dim h As Integer
    Dim adocmd As New ADODB.Command
    'add utk FIFO
    Dim dblStokTemp As Double
    Dim strNoTerima As String
    Dim rsTemp As ADODB.recordset
    Dim curHargaBarang As Currency
    Dim dblSelisih As Double
    Dim dblJmlTerkecil As Double
    Dim intRowTemp As Integer
    Dim k, l As Integer

    If (mstrKdDokter = "") And (chkDilayaniDokter.Value = 1) Then
        MsgBox "Pilih dulu Dokter Pemeriksa Pasien", vbCritical, "Validasi"
        txtDokter.SetFocus
        Exit Sub
    End If

    If chkPerawat.Value = vbChecked And subJmlTotal = 0 Then
        MsgBox "Nama perawat kosong", vbCritical, "Validasi"
        lvPemeriksa.Visible = True
        txtNamaPerawat.SetFocus
        Exit Sub
    End If

    If strKodePelayananRS = "" Then Exit Sub
    If optNonPaket.Value = True Then GoTo stepNonPaket
    Dim dTglPlyn As Date
    dTglPlyn = Now
    strSQL = "Select * FROM V_PaketPelayananObatAlkes WHERE KdPelayananRS='" & strKodePelayananRS & "' AND KdKelompokPasien = '" & mstrKdJenisPasien & "' AND IdPenjamin = '" & mstrKdPenjaminPasien & "'"
    Call msubRecFO(rs, strSQL)
    For i = 1 To rs.RecordCount
        'cek data barang & asal barang di grid paket obat alkes
        For j = 1 To fgDOA.Rows - 1
            'barang dengan asal barang tersebut sudah ada di grid obat alkes
            If fgDOA.TextMatrix(j, 0) = rs("KdBarang").Value And fgDOA.TextMatrix(j, 2) = rs("KdAsal").Value Then
                For h = 1 To intMaxJmlBarang
                    If typBarang(h).strkdbarang = rs("KdBarang").Value And typBarang(h).strKdAsal = rs("KdAsal").Value Then
                        intJmlBarang = h
                        GoTo stepCekStokBarang
                    End If
                Next h
            End If
            'sampai data terakhir data barang tidak ada di grid obat alkes
            If j = fgDOA.Rows - 1 Then
                'tambahkan data total barang yang terpakai
                intMaxJmlBarang = intMaxJmlBarang + 1
                intJmlBarang = intMaxJmlBarang
                ReDim Preserve typBarang(intMaxJmlBarang)

                Set rsB = Nothing
                Call msubRecFO(rsB, "Select JmlStok as Stok From StokRuangan Where KdRuangan='" & mstrKdRuangan & "' and KdBarang='" & rs("KdBarang").Value & "' and KdAsal='" & rs("KdAsal").Value & "'")
                dblStokTemp = IIf(IsNull(rsB("Stok").Value), 0, rsB("Stok").Value)

                strSQL = "Select JmlTerkecil,JmlJualTerkecil From MasterBarang Where KdBarang='" & rs("KdBarang").Value & "'"
                Set rsB = Nothing
                rsB.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

                typBarang(intJmlBarang).strkdbarang = rs("KdBarang").Value
                typBarang(intJmlBarang).strNamaBarang = rs("NamaBarang").Value
                typBarang(intJmlBarang).strKdAsal = rs("KdAsal").Value
                typBarang(intJmlBarang).intJmlTerkecil = rsB("JmlTerkecil").Value
                typBarang(intJmlBarang).intJmlJualTerkecil = rsB("JmlJualTerkecil").Value
                typBarang(intJmlBarang).intJmlTempTotal = dblStokTemp 'rsB("JmlTotalBarangTemp").Value

            End If
        Next j
stepCekStokBarang:
        If funcCekStokBarang(intJmlBarang, rs("SatuanJml"), (CInt(txtKuantitas) * rs("JmlBarang").Value)) = False Then
            'hapus grid obat alkes dengan kode pelayanan tersebut
            h = 1
            Do While h <= fgDOA.Rows - 2
                If fgDOA.TextMatrix(h, 9) = strKodePelayananRS Then
                    For j = 1 To intMaxJmlBarang
                        If typBarang(j).strkdbarang = fgDOA.TextMatrix(h, 0) Then
                            If fgDOA.TextMatrix(h, 5) = "S" Then
                                typBarang(j).intJmlTempTotal = typBarang(j).intJmlTempTotal + fgDOA.TextMatrix(h, 3)
                            Else
                                typBarang(j).intJmlTempTotal = typBarang(j).intJmlTempTotal + ((fgDOA.TextMatrix(h, 3) * typBarang(j).intJmlJualTerkecil) / typBarang(j).intJmlTerkecil)
                            End If
                        End If
                    Next j
                    fgDOA.RemoveItem h
                    h = 0
                End If
                h = h + 1
            Loop
            h = 1
            For j = 1 To intMaxJmlBarang
                For h = 1 To fgDOA.Rows - 1
                    If typBarang(j).strkdbarang = fgDOA.TextMatrix(h, 0) Then Exit For
                    If h = fgDOA.Rows - 1 Then
                        intMaxJmlBarang = intMaxJmlBarang - 1
                        If intMaxJmlBarang < 0 Then intMaxJmlBarang = 0
                    End If
                Next h
            Next j
            Exit Sub
        End If
        With fgDOA
            intRowNow = .Rows - 1

            strNoTerima = ""
            strSQL = ""
            Set rsB = Nothing
            Call msubRecFO(rsB, "select dbo.TakeNoFIFO_F('" & rs("KdBarang") & "','" & rs("KdAsal") & "','" & mstrKdRuangan & "') as NoFIFO")
            strNoTerima = IIf(IsNull(rsB("NoFIFO")), "0000000000", rsB("NoFIFO"))

            .TextMatrix(intRowNow, 0) = rs("KdBarang").Value
            .TextMatrix(intRowNow, 1) = rs("NamaBarang").Value
            .TextMatrix(intRowNow, 2) = rs("KdAsal").Value
            .TextMatrix(intRowNow, 3) = CInt(txtKuantitas) * rs("JmlBarang").Value
            .TextMatrix(intRowNow, 10) = strNoTerima
            
            .TextMatrix(intRowNow, 5) = rs("SatuanJml").Value

            'add New FIFO blm selesai
            If bolStatusFIFO = True Then
                Set rsB = Nothing
                Call msubRecFO(rsB, "Select JmlTerkecil,JmlJualTerkecil From MasterBarang Where KdBarang='" & rs("KdBarang").Value & "'")
                If rsB.EOF = False Then
                    dblJmlTerkecil = rsB("JmlTerkecil")
                Else
                    dblJmlTerkecil = 1
                End If
                Set rsB = Nothing
                Call msubRecFO(rsB, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & rs("KdBarang") & "','" & rs("KdAsal") & "','" & strNoTerima & "') as stok")
                If .TextMatrix(intRowNow, 5) = "S" Then
                    dblSelisih = rsB(0) - CDbl(.TextMatrix(intRowNow, 3))
                Else
                    dblSelisih = (rsB(0) * dblJmlTerkecil) - CDbl(.TextMatrix(intRowNow, 3))
                End If

                If dblSelisih < 0 Then
                    If .TextMatrix(intRowNow, 5) = "S" Then
                        .TextMatrix(intRowNow, 3) = rsB(0)
                    Else
                        .TextMatrix(intRowNow, 3) = rsB(0) * dblJmlTerkecil
                    End If

                Else
                    Set rsB = Nothing
                    strSQL = "Select JmlStok as Stok From StokRuangan Where KdBarang='" & rs("KdBarang") & "' and KdAsal='" & rs("KdAsal") & "' and KdRuangan='" & mstrKdRuangan & "'"
                    Call msubRecFO(rsB, strSQL)
                    If rsB.EOF Then
                        .TextMatrix(intRowNow, 3) = 0
                    Else
                        .TextMatrix(intRowNow, 3) = IIf(IsNull(rsB("Stok")), 0, rsB("Stok"))
                    End If
                End If

                strSQL = ""
                Set rsTemp = Nothing
                curHargaBarang = 0
                strSQL = "SELECT dbo.FB_TakeHargaNettoOA('" & mstrKdPenjaminPasien & "','" & mstrKdJenisPasien & "','" & .TextMatrix(intRowNow, 0) & "','" & .TextMatrix(intRowNow, 2) & "','" & .TextMatrix(intRowNow, 5) & "', '" & mstrKdRuangan & "','" & .TextMatrix(intRowNow, 10) & "') AS HargaBarang"
                Call msubRecFO(rsTemp, strSQL)
                If rsTemp.EOF = True Then curHargaBarang = 0 Else curHargaBarang = rsTemp(0).Value

                strSQL = ""
                Set rsTemp = Nothing
                subcurTarifHargaSatuan = 0
                strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "','" & .TextMatrix(intRowNow, 2) & "', " & msubKonversiKomaTitik(CStr(curHargaBarang)) & ")  as HargaSatuan"
                Call msubRecFO(rsTemp, strSQL)
                If rsTemp.EOF = True Then subcurTarifHargaSatuan = 0 Else subcurTarifHargaSatuan = rsTemp(0).Value

            Else
                curHargaBarang = rs("HargaBarang")
                strSQL = ""
                Set rsTemp = Nothing
                subcurTarifHargaSatuan = 0
                strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "','" & .TextMatrix(intRowTemp, 2) & "', " & msubKonversiKomaTitik(CStr(curHargaBarang)) & ")  as HargaSatuan"
                Call msubRecFO(rsTemp, strSQL)
                If rsTemp.EOF = True Then subcurTarifHargaSatuan = 0 Else subcurTarifHargaSatuan = rsTemp(0).Value
            End If
            'end fifo

            .TextMatrix(intRowNow, 4) = subcurTarifHargaSatuan

            If chkDilayaniDokter.Value = 1 Then
                .TextMatrix(intRowNow, 6) = mstrKdDokter
            Else
                .TextMatrix(intRowNow, 6) = UserID
            End If
            .TextMatrix(intRowNow, 7) = Format(dTglPlyn, "dd/mm/yyyy HH:mm:ss")
            .TextMatrix(intRowNow, 8) = rs("AsalBarang").Value
            .TextMatrix(intRowNow, 9) = strKodePelayananRS

            'add for FIFO jika jml yg diinput melebihi stok penerimaan, mk otomatis muncul di row selanjutnya
            If dblSelisih < 0 Then
                strSQL = "select NoTerima As NoFIFO,JmlStokMax from V_StokRuanganFIFO where KdBarang='" & .TextMatrix(intRowNow, 0) & "' and KdAsal='" & .TextMatrix(intRowNow, 2) & "' and NoTerima<>'" & .TextMatrix(intRowNow, 10) & "' and JmlStok<>0 order by TglTerima asc"
                Set rsB = Nothing
                Call msubRecFO(rsB, strSQL)
                If rsB.EOF = False Then
                    rsB.MoveFirst
                    For k = 1 To rsB.RecordCount
                        .Rows = .Rows + 1
                        intRowTemp = intRowNow '.row
                        If .TextMatrix(.Rows - 2, 2) = "" Then
                            .Row = .Rows - 2
                        Else
                            .Row = .Rows - 1
                        End If
                        For l = 0 To .Cols - 1
                            .Col = l
                            .CellBackColor = vbRed
                            .CellForeColor = vbWhite
                        Next l

                        .Row = intRowTemp
                        intRowTemp = 0
                        If .TextMatrix(.Rows - 2, 0) = "" Then
                            intRowTemp = .Rows - 2
                        Else
                            intRowTemp = .Rows - 1
                        End If
                        .TextMatrix(intRowTemp, 0) = .TextMatrix(.Row, 0)
                        .TextMatrix(intRowTemp, 1) = .TextMatrix(.Row, 1)
                        .TextMatrix(intRowTemp, 2) = .TextMatrix(.Row, 2)
                        .TextMatrix(intRowTemp, 5) = .TextMatrix(.Row, 5)
                        .TextMatrix(intRowTemp, 6) = .TextMatrix(.Row, 6)
                        .TextMatrix(intRowTemp, 7) = .TextMatrix(.Row, 7)
                        .TextMatrix(intRowTemp, 8) = .TextMatrix(.Row, 8)
                        .TextMatrix(intRowTemp, 9) = .TextMatrix(.Row, 9)
                        .TextMatrix(intRowTemp, 10) = rsB("NoFIFO")

                        strSQL = ""
                        Set rsTemp = Nothing
                        curHargaBarang = 0
                        strSQL = "SELECT dbo.FB_TakeHargaNettoOA('" & mstrKdPenjaminPasien & "','" & mstrKdJenisPasien & "','" & .TextMatrix(intRowTemp, 0) & "','" & .TextMatrix(intRowTemp, 2) & "','" & .TextMatrix(intRowTemp, 5) & "', '" & mstrKdRuangan & "','" & .TextMatrix(intRowTemp, 10) & "') AS HargaBarang"
                        Call msubRecFO(rsTemp, strSQL)
                        If rsTemp.EOF = True Then curHargaBarang = 0 Else curHargaBarang = rsTemp(0).Value

                        strSQL = ""
                        Set rsTemp = Nothing
                        subcurTarifHargaSatuan = 0
                        strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "','" & .TextMatrix(intRowTemp, 2) & "', " & msubKonversiKomaTitik(CStr(curHargaBarang)) & ")  as HargaSatuan"
                        Call msubRecFO(rsTemp, strSQL)
                        If rsTemp.EOF = True Then subcurTarifHargaSatuan = 0 Else subcurTarifHargaSatuan = rsTemp(0).Value
                        'khusus OA harga tidak dikalikan lg Ppn krn OA termasuk pelayanan yg include ke tindakan (TM)
                        .TextMatrix(intRowTemp, 4) = subcurTarifHargaSatuan

                        .TextMatrix(intRowTemp, 3) = Abs(dblSelisih)
                        If .TextMatrix(intRowTemp, 5) = "S" Then
                            dblSelisih = Abs(dblSelisih) - CDbl(rsB("JmlStokMax"))
                        Else
                            dblSelisih = Abs(dblSelisih) - CDbl(rsB("JmlStokMax") * dblJmlTerkecil)
                        End If
                        If dblSelisih >= 0 Then
                            If .TextMatrix(intRowTemp, 5) = "S" Then
                                .TextMatrix(intRowTemp, 3) = rsB("JmlStokMax")
                            Else
                                .TextMatrix(intRowTemp, 3) = rsB("JmlStokMax") * dblJmlTerkecil
                            End If
                        End If

                        If dblSelisih <= 0 Then Exit For
                        rsB.MoveNext
                    Next k
                End If
            End If
            'end Fifo

            .Rows = .Rows + 1
            .SetFocus
        End With
        rs.MoveNext
    Next i
stepNonPaket:
    With fgPelayanan
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = strKodePelayananRS Then Exit Sub
        Next i
        intRowNow = .Rows - 1
        .TextMatrix(intRowNow, 0) = strKodePelayananRS
        .TextMatrix(intRowNow, 1) = txtNamaPelayanan.Text
        .TextMatrix(intRowNow, 2) = CInt(txtKuantitas.Text)

        subcurTarifCito = sp_Take_TarifBPT
        .TextMatrix(intRowNow, 3) = subcurTarifBiayaSatuan 'curBiaya
        .TextMatrix(intRowNow, 4) = funcRoundUp(CStr(subcurTarifBiayaSatuan + subcurTarifCito)) * CInt(txtKuantitas.Text)
        .TextMatrix(intRowNow, 8) = subcurTarifCito

        .TextMatrix(intRowNow, 5) = mdTglBerlaku
        If chkDilayaniDokter.Value = 1 Then
            .TextMatrix(intRowNow, 6) = mstrKdDokter
        Else
            .TextMatrix(intRowNow, 6) = UserID
        End If
        If optCito(0).Value = True Or strCito = "1" Then
            .TextMatrix(intRowNow, 7) = "Y"
        Else
            .TextMatrix(intRowNow, 7) = "T"
        End If
        .TextMatrix(intRowNow, 9) = dtpTglPeriksa.Value
        .TextMatrix(intRowNow, 10) = IIf(chkDelegasi.Value = vbChecked, "1", "0")

        .TextMatrix(intRowNow, 11) = "0"

        .Rows = .Rows + 1
'        .SetFocus
    End With

    If strKodePelayananRS = "117053" Then
        mdTglBerlaku = dtpTglPeriksa.Value
        frmDetailPemakaianDarah.Show
        frmTindakan.Visible = False

    End If

    If chkPerawat.Value = vbChecked Then Call subLoadPelayananPerPerawat
    txtNamaPelayanan.Text = ""
    txtKuantitas.Text = 1
    fraPelayanan.Visible = False
End Sub

Private Sub subLoadPelayananPerPerawat()
    With fgPerawatPerPelayanan
        For i = 1 To subJmlTotal
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = mstrNoPen
            .TextMatrix(.Rows - 1, 1) = mstrKdRuangan
            .TextMatrix(.Rows - 1, 2) = dtpTglPeriksa.Value
            .TextMatrix(.Rows - 1, 3) = strKodePelayananRS
            .TextMatrix(.Rows - 1, 4) = Mid(subKdPemeriksa(i), 4, Len(subKdPemeriksa(i)) - 3)
            .TextMatrix(.Rows - 1, 5) = strIDPegawaiAktif
        Next
    End With

    subJmlTotal = 0
    txtNamaPerawat.BackColor = &HFFFFFF
    ReDim Preserve subKdPemeriksa(subJmlTotal)
    chkPerawat.Value = vbUnchecked
End Sub

Private Sub dgDokter_DblClick()
    Call dgDokter_KeyPress(13)
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)

    If strPilihGrid = "Dokter" Then
        If KeyAscii = 13 Then
            If intJmlDokter = 0 Then Exit Sub
            txtDokter.Text = dgDokter.Columns(1).Value
            mstrKdDokter = dgDokter.Columns(0).Value
            If mstrKdDokter = "" Then
                MsgBox "Pilih dulu Dokter yang akan menangani Pasien", vbCritical, "Validasi"
                txtDokter.Text = ""
                dgDokter.SetFocus
                Exit Sub
            End If
            chkDilayaniDokter.Value = 1
            fraDokter.Visible = False
            If chkPerawat.Enabled = True Then chkPerawat.SetFocus Else txtNamaPelayanan.SetFocus

 '           chkPerawat.SetFocus
        End If
    ElseIf strPilihGrid = "Dokter2" Then
        If KeyAscii = 13 Then
            If mintJmlDokter = 0 Then Exit Sub
            txtDokter2.Text = dgDokter.Columns(1).Value
            mstrKdDokter2 = dgDokter.Columns(0).Value
            If mstrKdDokter2 = "" Then
                MsgBox "Pilih dulu Dokter yang akan menangani Pasien", vbCritical, "Validasi"
                txtDokter2.Text = ""
                dgDokter.SetFocus
                Exit Sub
            End If
            fraDokter.Visible = False
        End If

        If KeyAscii = 27 Then
            fraDokter.Visible = False
        End If
    End If
End Sub

Private Sub dgPelayanan_DblClick()
    Call dgPelayanan_KeyPress(13)
End Sub

Private Sub dgPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlPelayanan = 0 Then Exit Sub
        Dim strkd As String
        strkd = dgPelayanan.Columns(5).Value
        curBiaya = dgPelayanan.Columns(4).Value
        txtNamaPelayanan.Text = dgPelayanan.Columns(1).Value
        strKodePelayananRS = strkd
        optNonPaket.Value = True
        If strKodePelayananRS = "" Then
            MsgBox "Pilih dulu tindakan pelayanan Pasien", vbCritical, "Validasi"
            txtNamaPelayanan.Text = ""
            dgPelayanan.SetFocus
            Exit Sub
        End If
        fraPelayanan.Visible = False
        txtKuantitas.SetFocus
    End If
    If KeyAscii = 27 Then
        fraDokter.Visible = False
    End If
End Sub

Private Sub dtpTglPeriksa_Change()
    If dtpTglPeriksa.Value < mdTglMasuk Then dtpTglPeriksa = mdTglMasuk
    dtpTglPeriksa.MaxDate = Now
End Sub

Private Sub dtpTglPeriksa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then chkDelegasi.SetFocus
End Sub

Private Sub fgPelayanan_Click()

       intRowNow = fgPelayanan.Row
          
       If fgPelayanan.TextMatrix(intRowNow, 7) = "Y" Then optCito(0).Value = True Else optCito(1).Value = True
    
    
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    frmTransaksiPasien.Enabled = False
    strKdKelas = mstrKdKelas
    strKelas = mstrKelas
    Set rs = Nothing
    strSQL = "SELECT KdJenisTarif,JenisTarif " _
    & "FROM v_JenisTarifPasien " _
    & "WHERE NoPendaftaran='" & mstrNoPen & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockOptimistic
    strKdJenisTarif = rs.Fields(0).Value
    strJenisTarif = rs.Fields(1).Value
    Set rs = Nothing

    Call subSetGidPelayanan
    dtpTglPeriksa.Value = Now
    strSQL11 = ""
    strSQL11 = "SELECT    KdPelayananRS, NamaPelayanan, JmlPelayanan, BiayaSatuan, BiayaTotal, IdDokterOrder, StatusCito, NoRiwayat, KdKelas" & _
    " FROM V_DetailOrderTM where NoPendaftaran ='" & mstrNoPen & "' and KdRuanganTujuan ='" & mstrKdRuangan & "' and KdKelas ='" & mstrKdKelas & "'"
    Set rsD = Nothing
    Call msubRecFO(rsD, strSQL11)
    
    If rsD.EOF = False Then
        rsD.MoveFirst
        With fgPelayanan
            For i = 1 To rsD.RecordCount
                For j = 1 To fgPelayanan.Rows - 2
                     strKodePelayananRS = rsD.Fields("KdPelayananRS")
                     If (.TextMatrix(j, 0) = strKodePelayananRS) And _
                      (.TextMatrix(j, 9) = dtpTglPeriksa.Value) Then dtpTglPeriksa.Value = DateAdd("s", 1, dtpTglPeriksa.Value)
                 Next j

                intRowNow = i
                .TextMatrix(intRowNow, 0) = rsD.Fields("KdPelayananRS") 'strKodePelayananRS
                 strKodePelayananRS = rsD.Fields("KdPelayananRS")
                .TextMatrix(intRowNow, 1) = rsD.Fields("NamaPelayanan") 'txtNamaPelayanan.Text
                .TextMatrix(intRowNow, 2) = rsD.Fields("JmlPelayanan") 'CInt(txtKuantitas.Text)
                
                
                 If rsD.Fields("StatusCito") = "Y" Or rsD.Fields("StatusCito") = "1" Then strCito = 1
                If rsD.Fields("StatusCito") = "T" Or rsD.Fields("StatusCito") = "0" Then strCito = 0

                ' If rs.Fields("StatusCito") = 1 Then optCito(0).Value = True Else optCito(1).Value = True
                If rsD.Fields("StatusCito") = "Y" Or rsD.Fields("StatusCito") = "1" Then optCito(0).Value = True
                If rsD.Fields("StatusCito") = "T" Or rsD.Fields("StatusCito") = "0" Then optCito(1).Value = True
                subcurTarifCito = sp_Take_TarifBPT
                .TextMatrix(intRowNow, 3) = IIf(rsD.Fields("BiayaSatuan") = 0, 0, Format(rsD.Fields("BiayaSatuan"), "#,###")) 'curBiaya
                .TextMatrix(intRowNow, 4) = Format(CCur(rsD.Fields("BiayaTotal")) + subcurTarifCito, "#,###")
                .TextMatrix(intRowNow, 8) = subcurTarifCito

                .TextMatrix(intRowNow, 5) = mdTglBerlaku
                If IsNull(rsD.Fields("IdDokterOrder")) Then
                    .TextMatrix(intRowNow, 6) = UserID
                Else
                    .TextMatrix(intRowNow, 6) = rsD.Fields("IdDokterOrder")
                    mstrKdDokter = rsD.Fields("IdDokterOrder")
                End If
                

                If strCito = "" Then
                    .TextMatrix(intRowNow, 7) = "T"
                    strCito = 0
                End If
                
                If strCito = 1 Then
                   .TextMatrix(intRowNow, 7) = "Y"
                Else
                
                   .TextMatrix(intRowNow, 7) = "T"
                
                End If
                
                 bolPasienKonsul = True
'                .TextMatrix(intRowNow, 7) = strCito
                .TextMatrix(intRowNow, 9) = dtpTglPeriksa.Value ' Format(Now, "yyyy/MM/dd HH:mm:ss") '
                .TextMatrix(intRowNow, 11) = "1"
                
                .Rows = .Rows + 1
                rsD.MoveNext
            Next i
        End With
    End If

    strCito = "0"
    strStatusAPBD = "01"
    optNonPaket.Value = True
    Call subSetGridObatAlkes
    intBarang = 0
    intJmlBarang = 0
    intMaxJmlBarang = 0
    ReDim typBarang(0)

    subJmlTotal = 0
    Call subSetGridPerawatPerPelayanan
    Call subLoadListPemeriksa
    chkDilayaniDokter.Value = vbChecked
    chkPerawat.Value = vbChecked
    lvPemeriksa.Visible = False

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmTransaksiPasien.Enabled = True
    BolPesanTindakan = False
End Sub

Private Sub lvPemeriksa_DblClick()
    Call lvPemeriksa_KeyPress(13)
End Sub

Private Sub lvPemeriksa_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim blnSelected As Boolean
    If Item.Checked = True Then
        subJmlTotal = subJmlTotal + 1
        ReDim Preserve subKdPemeriksa(subJmlTotal)
        subKdPemeriksa(subJmlTotal) = Item.Key
    Else
        blnSelected = False
        For i = 1 To subJmlTotal
            If subKdPemeriksa(i) = Item.Key Then blnSelected = True
            If blnSelected = True Then
                If i = subJmlTotal Then
                    subKdPemeriksa(i) = ""
                Else
                    subKdPemeriksa(i) = subKdPemeriksa(i + 1)
                End If
            End If
        Next i
        subJmlTotal = subJmlTotal - 1
    End If

    If subJmlTotal = 0 Then
        txtNamaPerawat.BackColor = &HFFFFFF
    Else
        txtNamaPerawat.BackColor = &HC0FFFF
    End If
End Sub

Private Sub lvPemeriksa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lvPemeriksa.Visible = False: txtNamaPerawat.SetFocus
End Sub

Private Sub optCito_Click(Index As Integer)
  If bolPasienKonsul = False Then
    If Index = 0 Then
        strCito = "1"
    Else
        strCito = "0"
        
    End If
 Else
 
     strSQL = "SELECT    KdPelayananRS, NamaPelayanan, JmlPelayanan, BiayaSatuan, BiayaTotal, IdDokterOrder, StatusCito, NoRiwayat, KdKelas" & _
    " FROM V_DetailOrderTM where NoPendaftaran ='" & mstrNoPen & "' and KdRuanganTujuan ='" & mstrKdRuangan & "' and KdKelas ='" & mstrKdKelas & "' and KdPelayananRS='" & fgPelayanan.TextMatrix(fgPelayanan.Row, 0) & "' "

    Call msubRecFO(rs, strSQL)
    If rs.RecordCount <> 0 Then
'       For i = 1 To rs.RecordCount
         If Index = 0 Then
         
            strCito = "1"
            fgPelayanan.TextMatrix(fgPelayanan.Row, 7) = "Y"
            optCito(0).Value = True
            subcurTarifCito = sp_Take_TarifBPTUbahKonsulCito(fgPelayanan.TextMatrix(fgPelayanan.Row, 0))
             
'            fgPelayanan.TextMatrix(intRowNow, 3) = IIf(rs.Fields("BiayaSatuan") = 0, 0, Format(rs.Fields("BiayaSatuan"), "#,###")) 'curBiaya
            fgPelayanan.TextMatrix(fgPelayanan.Row, 4) = Format(CCur(rs.Fields("BiayaTotal")) + subcurTarifCito, "#,###") ' Format(CCur(fgPelayanan.TextMatrix(fgPelayanan.Row, 4)) + subcurTarifCito, "#,###") '
            fgPelayanan.TextMatrix(fgPelayanan.Row, 8) = subcurTarifCito
            
            'fgPelayanan.TextMatrix(intRowNow, 4) = fgPelayanan.TextMatrix(intRowNow, 3) + subcurTarifCito
            'fgPelayanan.TextMatrix(intRowNow, 8) = subcurTarifCito
        Else
            strCito = "0"
            
            fgPelayanan.TextMatrix(fgPelayanan.Row, 7) = "T"
            optCito(0).Value = False
            subcurTarifCito = sp_Take_TarifBPTUbahKonsulCito(fgPelayanan.TextMatrix(fgPelayanan.Row, 0))
            fgPelayanan.TextMatrix(fgPelayanan.Row, 8) = 0
            fgPelayanan.TextMatrix(fgPelayanan.Row, 4) = IIf(rs.Fields("BiayaSatuan") = 0, 0, Format(rs.Fields("BiayaSatuan"), "#,###")) + 0
            
            
        End If
'        rs.MoveNext
'       Next i
    End If
 End If
 
    
End Sub

Private Sub optCito_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkAPBD.Enabled = True Then
            chkAPBD.SetFocus
        Else
            txtNamaPelayanan.SetFocus
        End If
    End If
End Sub

Private Sub optNonPaket_Click()
    fraButton.Enabled = True
End Sub

Private Sub optNonPaket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdTambah.SetFocus
End Sub

Private Sub optPaket_Click()
    strSQL = "SELECT * FROM PaketLayanan WHERE KdPelayananRS='" & strKodePelayananRS _
    & "' AND KdRuangan='" & mstrKdRuangan & "'" 'mstrKdRuanganPasien & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada paket untuk pelayanan yang dipilih", vbCritical, "Validasi"
        optNonPaket.SetFocus
    Else
    End If
    fraButton.Enabled = True
End Sub

Private Sub optPaket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdTambah.SetFocus
End Sub

'Private Sub Timer1_Timer()
'    dtpTglPeriksa.Value = Now
'End Sub

Private Sub txtDokter_Change()
'    On Error GoTo errLoad
    strFilterDokter = "WHERE NamaDokter like '%" & txtDokter.Text & "%'"
    mstrKdDokter = ""
    strPilihGrid = "Dokter"
    fraDokter.Visible = True
    Call subLoadDokter
'    Exit Sub
'errLoad:
'    Call msubPesanError
End Sub

Private Sub txtDokter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If fraDokter.Visible = False Then Exit Sub
        dgDokter.SetFocus
    End If
End Sub

Private Sub txtDokter_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If intJmlDokter = 0 Then Exit Sub
        If fraDokter.Visible = True Then
            dgDokter.SetFocus
        Else
            chkPerawat.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        fraDokter.Visible = False
    End If
hell:
End Sub

Private Sub txtDokter2_Change()
    strPilihGrid = "Dokter2"
    fraDokter.Visible = True
    Call subLoadDokter2
End Sub

Private Sub txtDokter2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If fraDokter.Visible = True Then dgDokter.SetFocus
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtKuantitas_GotFocus()
    txtKuantitas.SelStart = 0
    txtKuantitas.SelLength = Len(txtKuantitas.Text)
End Sub

Private Sub txtKuantitas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then optNonPaket.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtKuantitas_LostFocus()
    If txtKuantitas.Text = "" Then txtKuantitas.Text = 1: Exit Sub
    If txtKuantitas.Text = 0 Then txtKuantitas.Text = 1
End Sub

Private Sub txtNamaPelayanan_Change()
    strFilterPelayanan = "WHERE [Nama Pelayanan] like '%" & txtNamaPelayanan.Text _
    & "%' AND KdKelas='" & strKdKelas & "' AND KdJenisTarif='" & strKdJenisTarif _
    & "' AND KdRuangan='" & mstrKdRuangan & "'"
    fraPelayanan.Visible = True
    Call subLoadPelayanan
End Sub

Private Sub txtNamaPelayanan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If fraPelayanan.Visible = False Then Exit Sub
        dgPelayanan.SetFocus
    End If
End Sub

Private Sub txtNamaPelayanan_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If intJmlPelayanan = 0 Then Exit Sub
        If fraPelayanan.Visible = True Then
            dgPelayanan.SetFocus
        Else
            txtKuantitas.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        fraPelayanan.Visible = False
    End If
hell:
End Sub

'untuk meload data dokter di grid
Private Sub subLoadDokter()
'If mstrloaddokter = False Then
    strSQL = "SELECT KodeDokter AS [Kode Dokter],NamaDokter AS [Nama Dokter],JK,Jabatan FROM V_DaftarDokter " & strFilterDokter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.EOF = False Then
    intJmlDokter = rs.RecordCount
    With dgDokter
        Set .DataSource = rs
        .Columns(0).Width = 1200
        .Columns(1).Width = 3000
        .Columns(2).Width = 400
        .Columns(3).Width = 3000
    End With
    fraDokter.Left = 0
    fraDokter.Top = 1920
    Else
       'txtDokter.Text = ""
    End If
'End If
End Sub

'untuk meload data pelayanan di grid
Private Sub subLoadPelayanan()
    On Error Resume Next
    strSQL = "SELECT [Jenis Pelayanan],[Nama Pelayanan],Kelas,JenisTarif,Tarif,KdPelayananRS FROM V_TarifPelayananTindakan " & strFilterPelayanan
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlPelayanan = rs.RecordCount
    With dgPelayanan
        Set .DataSource = rs
        .Columns(0).Width = 2100
        .Columns(1).Width = 3900
        .Columns(2).Width = 1000
        .Columns(3).Width = 1100
        .Columns(4).Width = 900
        .Columns(4).Alignment = dbgRight
        .Columns(5).Width = 0
    End With
    fraPelayanan.Left = 0
    fraPelayanan.Top = 3240
End Sub

'Store procedure untuk mengisi biaya pelayanan pasien
Private Function sp_BiayaPelayanan(ByVal adoCommand As ADODB.Command, strKdPelayananRS As String, curTarif As Currency, intJmlPel As Integer, dtTanggalPelayanan As Date, strkodedokter As String, strStatusCITO As String, f_TarifCito As Currency) As Boolean
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan) 'mstrKdRuanganPasien)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, strKdPelayananRS)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, strKdKelas)
        .Parameters.Append .CreateParameter("StatusCITO", adChar, adParamInput, 1, IIf(strStatusCITO = "Y", "1", "0"))
        .Parameters.Append .CreateParameter("Tarif", adInteger, adParamInput, , curTarif)
        .Parameters.Append .CreateParameter("JmlPelayanan", adInteger, adParamInput, , intJmlPel)
        '.Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(dtTanggalPelayanan, "yyyy/MM/dd HH:mm:ss"))
         .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(penamwaktu, "yyyy/MM/dd HH:mm:ss"))
         'untuk nambah 1 detik
         penamwaktu = DateAdd("s", 1, penamwaktu)
         
        .Parameters.Append .CreateParameter("NoLab_Rad", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, strkodedokter)
        .Parameters.Append .CreateParameter("StatusAPBD", adChar, adParamInput, 2, strStatusAPBD)
        .Parameters.Append .CreateParameter("KdJenisTarif", adChar, adParamInput, 2, strKdJenisTarif)
        .Parameters.Append .CreateParameter("TarifCito", adInteger, adParamInput, , f_TarifCito)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("IdPegawai2", adChar, adParamInput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_BiayaPelayanan"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Biaya Pelayanan Pasien", vbCritical, "Validasi"
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
            sp_BiayaPelayanan = False
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
        sp_BiayaPelayanan = True
    End With
End Function

'@Dimas 2014-05-16
'untuk update Biaya Pelayanan KSO Max 35000
Private Function update_BiayaPelayananKSO(ByVal adoCommand As ADODB.Command, f_KdPelayananKSO As String) As Boolean
    On Error GoTo errLoad
    update_BiayaPelayananKSO = True
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 10, f_KdPelayananKSO)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_BiayaPelayananKSOMax3500"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan data", vbCritical, "Validasi"
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
            update_BiayaPelayananKSO = False

        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With

    Exit Function
errLoad:
    update_BiayaPelayananKSO = False
    Call msubPesanError
End Function

'simpan data perawat
Private Function sp_PetugasPemeriksaBP(F_dtTanggalPelayanan As Date, F_strKodePelayanan As String, F_StrIdPerawat As String, F_IdUser As String) As Boolean
    On Error GoTo errLoad

    sp_PetugasPemeriksaBP = False

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan) 'mstrKdRuanganPasien)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(F_dtTanggalPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, F_strKodePelayanan)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, F_StrIdPerawat)  'kode perawat
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, F_IdUser)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PetugasPemeriksaBP"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data petugas pemeriksa BP", vbExclamation, "Validasi"
            sp_PetugasPemeriksaBP = False
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
        sp_PetugasPemeriksaBP = True
    End With

    Exit Function
errLoad:
    Call msubPesanError
    sp_PetugasPemeriksaBP = False
End Function

'untuk set grid pelayanan
Private Sub subSetGidPelayanan()
    With fgPelayanan
        .Clear
        .Rows = 2
        .Cols = 12
        .TextMatrix(0, 0) = "Kode Pelayanan"
        .TextMatrix(0, 1) = "Nama Pelayanan"
        .TextMatrix(0, 2) = "Jumlah"
        .TextMatrix(0, 3) = "Biaya Satuan"
        .TextMatrix(0, 4) = "Biaya Total"
        .TextMatrix(0, 5) = "Tgl Berlaku"
        .TextMatrix(0, 6) = "Kode Dokter"
        .TextMatrix(0, 7) = "Status CITO"
        .TextMatrix(0, 8) = "Biaya CITO"
        .TextMatrix(0, 9) = "Tanggal Pelayanan"
        .TextMatrix(0, 10) = "StatusDelegasi"
        .TextMatrix(0, 11) = "StatusOrder" 'for pesan pelayanan

        .ColWidth(0) = 0
        .ColWidth(1) = 4500
        .ColWidth(2) = 700
        .ColWidth(3) = 1200
        .ColWidth(4) = 1400
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 1200
        .ColWidth(8) = 1200
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColWidth(11) = 0
    End With
End Sub

'untuk set grid obat alkes
Private Sub subSetGridObatAlkes()
    With fgDOA
        .Clear
        .Rows = 2
        .Cols = 11
        .TextMatrix(0, 0) = "Kode Barang"
        .TextMatrix(0, 1) = "Nama Barang"
        .TextMatrix(0, 2) = "Kode Asal"
        .TextMatrix(0, 3) = "Jumlah"
        .TextMatrix(0, 4) = "Harga Satuan"
        .TextMatrix(0, 5) = "Satuan"
        .TextMatrix(0, 6) = "Kode Dokter"
        .TextMatrix(0, 7) = "tgl Pelayanan"
        .TextMatrix(0, 8) = "Asal Barang"
        .TextMatrix(0, 9) = "KdPelayananRS"
        .TextMatrix(0, 10) = "NoTerima"

        .ColWidth(0) = 0
        .ColWidth(1) = 4500
        .ColWidth(2) = 0
        .ColWidth(3) = 700
        .ColWidth(4) = 1200
        .ColWidth(5) = 700
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 1000
        .ColWidth(9) = 0
        .ColWidth(10) = 0
    End With
End Sub

'untuk cek validasi
Private Function funcCekValidasi() As Boolean
    If fgPelayanan.TextMatrix(1, 0) = "" Then
        MsgBox "Pilihan Pelayanan Pasien Harus Diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        txtNamaPelayanan.SetFocus
        Exit Function
    End If
    funcCekValidasi = True
End Function

'untuk enable/disable button reg
Private Sub subEnableButtonReg(blnStatus As Boolean)
    fraPDokter.Enabled = blnStatus
    fraPPelayanan.Enabled = blnStatus
    fgPelayanan.Enabled = blnStatus
    fgDOA.Enabled = blnStatus
    cmdSimpan.Enabled = blnStatus
End Sub

Private Function funcCekStokBarang(intBarang As Integer, strSatuanJml As String, dblJmlBrg As Double) As Boolean
    If strSatuanJml = "S" Then
        'jika stokRU - JmlBrg, stokRU = typBarang(intBarang).intJmlTempTotal
        If (typBarang(intBarang).intJmlTempTotal - dblJmlBrg) < 0 Then
            MsgBox "Stok Barang '" & typBarang(intBarang).strNamaBarang & "' Tidak Cukup !", vbCritical, "Validasi"
            funcCekStokBarang = False
            Exit Function
        Else
            typBarang(intBarang).intJmlTempTotal = typBarang(intBarang).intJmlTempTotal - (dblJmlBrg)
        End If
    Else
        'jika stokRU - JmlBrg, stokRU = typBarang(intBarang).intJmlTempTotal
        If (typBarang(intBarang).intJmlTempTotal - ((dblJmlBrg * typBarang(intBarang).intJmlJualTerkecil) / typBarang(intBarang).intJmlTerkecil)) < 0 Then
            MsgBox "Stok Barang '" & typBarang(intBarang).strNamaBarang & "' Tidak Cukup !", vbCritical, "Validasi"
            funcCekStokBarang = False
            Exit Function
        Else
            typBarang(intBarang).intJmlTempTotal = typBarang(intBarang).intJmlTempTotal - ((dblJmlBrg * typBarang(intBarang).intJmlJualTerkecil) / typBarang(intBarang).intJmlTerkecil)
        End If
    End If
    funcCekStokBarang = True
End Function

Private Sub txtNamaPerawat_Change()
    On Error GoTo errLoad

    Call subLoadListPemeriksa("where [Nama Pemeriksa] LIKE '%" & txtNamaPerawat.Text & "%'")
    lvPemeriksa.Visible = True

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtNamaPerawat_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If lvPemeriksa.Visible = True Then If lvPemeriksa.ListItems.Count > 0 Then lvPemeriksa.SetFocus
        Case vbKeyEscape
            lvPemeriksa.Visible = False
    End Select
End Sub

Private Sub txtNamaPerawat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If lvPemeriksa.Visible = True Then
            lvPemeriksa.SetFocus
        Else
            optCito(1).SetFocus
        End If
    End If
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub subSetGridPerawatPerPelayanan()
    With fgPerawatPerPelayanan
        .Cols = 6
        .Rows = 1

        .MergeCells = flexMergeFree

        .TextMatrix(0, 0) = "NoPendaftaran"
        .TextMatrix(0, 1) = "Kode Ruangan"
        .TextMatrix(0, 2) = "Tgl Pelayanan"
        .TextMatrix(0, 3) = "Kode Pelayanan"
        .TextMatrix(0, 4) = "IdPegawai"
        .TextMatrix(0, 5) = "IdUser"

    End With
End Sub

Private Sub subLoadListPemeriksa(Optional strKriteria As String)
    Dim strKey As String

    strSQL = "select * from v_daftarpemeriksapasien " & strKriteria & " order by [Nama Pemeriksa]"
    Call msubRecFO(rs, strSQL)

    With lvPemeriksa
        .ListItems.Clear
        For i = 0 To rs.RecordCount - 1
            strKey = "key" & rs(0).Value
            .ListItems.Add , strKey, rs(1).Value
            rs.MoveNext
        Next

        .Top = fraPDokter.Top + txtNamaPerawat.Top + txtNamaPerawat.Height
        .Left = fraPDokter.Left + txtNamaPerawat.Left
        .Height = 1815
        .ColumnHeaders.Item(1).Width = lvPemeriksa.Width - 500

        If subJmlTotal = 0 Then Exit Sub
        For i = 1 To .ListItems.Count
            For j = 1 To subJmlTotal
                If .ListItems(i).Key = subKdPemeriksa(j) Then .ListItems(i).Checked = True
            Next j
        Next i
    End With
End Sub

Private Function sp_Take_TarifBPT() As Currency
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, strKodePelayananRS)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, mstrKdKelas)
        .Parameters.Append .CreateParameter("KdJenisTarif", adChar, adParamInput, 2, strKdJenisTarif)
        .Parameters.Append .CreateParameter("TarifCito", adCurrency, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("TarifTotal", adCurrency, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("StatusCito", adChar, adParamInput, 1, IIf(optCito(0).Value = True, "Y", "T"))
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, IIf(chkDilayaniDokter.Value = vbChecked, mstrKdDokter, Null))
        .Parameters.Append .CreateParameter("IdDokter2", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("IdDokter3", adChar, adParamInput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Take_TarifBPT"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam Pengambilan biaya tarif", vbExclamation, "Validasi"
            sp_Take_TarifBPT = 0
            subcurTarifBiayaSatuan = 0
        Else
            sp_Take_TarifBPT = .Parameters("TarifCito").Value
            subcurTarifBiayaSatuan = .Parameters("TarifTotal").Value
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Function sp_Take_TarifOA(f_KdAsal As String, f_HargaSatuan As Currency) As Currency
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 6, f_KdAsal)
        .Parameters.Append .CreateParameter("HargaSatuan", adCurrency, adParamInput, , CCur(f_HargaSatuan))
        .Parameters.Append .CreateParameter("TarifTotal", adCurrency, adParamOutput, , Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Take_TarifOA"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam Pengambilan biaya tarif", vbExclamation, "Validasi"
            sp_Take_TarifOA = 0
        Else
            sp_Take_TarifOA = .Parameters("TarifTotal").Value
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Function update_DetailOrderTMOA(ByVal adoCommand As ADODB.Command, sItem As String, sStatus As String) As Boolean
    On Error GoTo errLoad
    update_DetailOrderTMOA = True
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdItem", adVarChar, adParamInput, 9, sItem)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 2, sStatus)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_DetailOrderTMOA"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan data", vbCritical, "Validasi"
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
            update_DetailOrderTMOA = False

        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With

    Exit Function
errLoad:
    update_DetailOrderTMOA = False
    Call msubPesanError
End Function

'untuk meload data dokter delegasi di grid
Private Sub subLoadDokter2()
    On Error GoTo errLoad
    strSQL = "SELECT KodeDokter AS [Kode Dokter],NamaDokter AS [Nama Dokter],JK,Jabatan FROM V_DaftarDokter WHERE NamaDokter like '%" & txtDokter2.Text & "%'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    mintJmlDokter = rs.RecordCount
    With dgDokter
        Set .DataSource = rs
        .Columns(0).Width = 1200
        .Columns(1).Width = 3000
        .Columns(2).Width = 400
        .Columns(3).Width = 3000
    End With
    fraDokter.Left = 4000
    fraDokter.Top = 1920
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Function sp_Take_TarifBPTUbahKonsulCito(strKodePelayananRSCito As String) As Currency
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, strKodePelayananRSCito)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, mstrKdKelas)
        .Parameters.Append .CreateParameter("KdJenisTarif", adChar, adParamInput, 2, strKdJenisTarif)
        .Parameters.Append .CreateParameter("TarifCito", adCurrency, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("TarifTotal", adCurrency, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("StatusCito", adChar, adParamInput, 1, IIf(optCito(0).Value = True, "Y", "T"))
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, IIf(chkDilayaniDokter.Value = vbChecked, mstrKdDokter, Null))
        .Parameters.Append .CreateParameter("IdDokter2", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("IdDokter3", adChar, adParamInput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Take_TarifBPT"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam Pengambilan biaya tarif", vbExclamation, "Validasi"
            sp_Take_TarifBPTUbahKonsulCito = 0
            subcurTarifBiayaSatuan = 0
        Else
            sp_Take_TarifBPTUbahKonsulCito = .Parameters("TarifCito").Value
            subcurTarifBiayaSatuan = .Parameters("TarifTotal").Value
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function
