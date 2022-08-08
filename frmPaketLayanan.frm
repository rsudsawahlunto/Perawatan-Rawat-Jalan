VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmPaketLayanan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Paket Pelayanan Tindakan"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPaketLayanan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   10230
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
      Height          =   2655
      Left            =   240
      TabIndex        =   18
      Top             =   2040
      Visible         =   0   'False
      Width           =   8775
      Begin MSDataGridLib.DataGrid dgPelayanan 
         Height          =   2175
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   3836
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   16
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraObatAlkes 
      Caption         =   "Data Barang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   1200
      TabIndex        =   27
      Top             =   2040
      Visible         =   0   'False
      Width           =   8775
      Begin MSDataGridLib.DataGrid dgObatAlkes 
         Height          =   2175
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   3836
         _Version        =   393216
         HeadLines       =   1
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
               LCID            =   1033
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
               LCID            =   1033
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
   Begin VB.Frame Frame1 
      Caption         =   "Data Paket Pelayanan Tindakan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      TabIndex        =   13
      Top             =   960
      Width           =   10215
      Begin VB.TextBox txtSatuanJml 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   2760
         TabIndex        =   31
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtHarga 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   7440
         TabIndex        =   29
         Top             =   2280
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtJmlJualTerkecil 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1320
         TabIndex        =   22
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtJmlTerkecil 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtNamaBrg 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4440
         TabIndex        =   2
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtNamaPelayanan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         TabIndex        =   0
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox txtjml 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4200
         TabIndex        =   12
         Top             =   1440
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo dcAsal 
         Height          =   330
         Left            =   8040
         TabIndex        =   4
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Satuan"
         Height          =   210
         Left            =   2760
         TabIndex        =   32
         Top             =   1200
         Width           =   570
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Harga Satuan"
         Height          =   210
         Left            =   7440
         TabIndex        =   28
         Top             =   2040
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblSatuanK2 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3600
         TabIndex        =   24
         Top             =   1500
         Width           =   60
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Jml.Jual Terkecil"
         Height          =   210
         Left            =   1320
         TabIndex        =   23
         Top             =   1200
         Width           =   1305
      End
      Begin VB.Label lblSatuanK 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1320
         TabIndex        =   21
         Top             =   1500
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Jml.Terkecil"
         Height          =   210
         Left            =   240
         TabIndex        =   20
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pelayanan"
         Height          =   210
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nama Barang"
         Height          =   210
         Left            =   4440
         TabIndex        =   16
         Top             =   480
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Asal Barang"
         Height          =   210
         Left            =   8040
         TabIndex        =   15
         Top             =   480
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Jml. Barang"
         Height          =   210
         Left            =   4200
         TabIndex        =   14
         Top             =   1200
         Width           =   930
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Daftar Paket Layanan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   0
      TabIndex        =   25
      Top             =   3120
      Width           =   10215
      Begin VB.CommandButton cmdCari 
         Caption         =   "&Cari"
         Height          =   375
         Left            =   3360
         TabIndex        =   7
         Top             =   4920
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   4960
         Width           =   3135
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   7320
         TabIndex        =   10
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CommandButton cmdBaru 
         Caption         =   "&Batal"
         Height          =   375
         Left            =   4440
         TabIndex        =   9
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   5880
         TabIndex        =   8
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   8760
         TabIndex        =   11
         Top             =   4920
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid dgPaketPelayanan 
         Height          =   4335
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   7646
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan Nama Pelayanan"
         Height          =   210
         Left            =   120
         TabIndex        =   26
         Top             =   4720
         Width           =   2160
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   30
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
      Picture         =   "frmPaketLayanan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPaketLayanan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPaketLayanan.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmPaketLayanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strKodePelayananRS As String
Dim strFilterPelayanan As String
Dim intJmlPelayanan As Integer
Dim strkdbarang As String
Dim intJmlObatAlkes As Integer
Dim strFilterObatAlkes As String

Private Sub cmdBaru_Click()
    subClearData
End Sub

Private Sub cmdBaru_GotFocus()
    subInvisibleFram
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errHapus
    If MsgBox("Yakin akan menghapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    If strKodePelayananRS = "" Then
        MsgBox "Nama pelayanan kosong", vbCritical, "Validasi"
        txtNamaPelayanan.SetFocus
        Exit Sub
    ElseIf strkdbarang = "" Then
        MsgBox "Nama barang kosong", vbCritical, "Validasi"
        txtNamaBrg.SetFocus
        Exit Sub
    ElseIf txtjml.Text = "" Then
        MsgBox "Jumlah kosong", vbCritical, "Validasi"
        txtjml.SetFocus
        Exit Sub
    End If
    If sp_PaketLayanan("D") = False Then GoTo errHapus

    MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
    subLoadPaketLayanan
    subClearData
    subInvisibleFram
    Exit Sub
errHapus:
    Call msubPesanError
End Sub

Private Sub cmdHapus_GotFocus()
    subInvisibleFram
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errSimpan
    If strKodePelayananRS = "" Then
        MsgBox "Isi dulu nama pelayanannya", vbExclamation, "Validasi"
        txtNamaPelayanan.SetFocus
        Exit Sub
    ElseIf strkdbarang = "" Then
        MsgBox "Isi dulu nama barangnya", vbExclamation, "Validasi"
        txtNamaBrg.SetFocus
        Exit Sub
    ElseIf txtjml.Text = "" Then
        MsgBox "Jumlah harus diisi", vbExclamation, "Validasi"
        txtjml.SetFocus
        Exit Sub
    End If

    If sp_PaketLayanan("A") = False Then Exit Sub

    MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
    subClearData
    subLoadPaketLayanan
    subInvisibleFram
    If txtNamaPelayanan.Enabled = True Then txtNamaPelayanan.SetFocus
    Exit Sub
errSimpan:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_GotFocus()
    subInvisibleFram
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdTutup_GotFocus()
    subInvisibleFram
End Sub

Private Sub dcAsal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcAsal.MatchedWithList = True Then cmdSimpan.SetFocus
        strSQL = "SELECT KdAsal,NamaAsal FROM AsalBarang where(NamaAsal LIKE '%" & dcAsal.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcAsal.Text = ""
            cmdSimpan.SetFocus
            Exit Sub
        End If
        dcAsal.BoundText = rs(0).Value
        dcAsal.Text = rs(1).Value
    End If
End Sub

Private Sub dgObatAlkes_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgObatAlkes
    WheelHook.WheelHook dgObatAlkes
End Sub

Private Sub dgObatAlkes_DblClick()
    Call dgObatAlkes_KeyPress(13)
End Sub

Private Sub dgObatAlkes_KeyPress(KeyAscii As Integer)
    On Error GoTo errError
    If KeyAscii = 13 Then
        If intJmlObatAlkes = 0 Then Exit Sub

        Dim strKdBrg As String
        Dim intHrgSat As Currency
        Dim strSatJml As String
        Dim strKdAsl As String
        Dim strNmAsl As String
        Dim strSatuanJml As String
        Dim strKdAsal As String
        Dim strAsalBarang As String

        strKdBrg = dgObatAlkes.Columns(5).Value
        intHrgSat = dgObatAlkes.Columns(4).Value
        txtHarga.Text = dgObatAlkes.Columns(4).Value
        strSatJml = dgObatAlkes.Columns(3).Value
        strKdAsl = dgObatAlkes.Columns(6).Value
        strNmAsl = dgObatAlkes.Columns(2).Value
        txtNamaBrg.Text = dgObatAlkes.Columns(1).Value
        fraObatAlkes.Visible = False
        strkdbarang = strKdBrg
        txtHarga = intHrgSat
        strSatuanJml = strSatJml
        strKdAsal = strKdAsl
        strAsalBarang = strNmAsl
        dcAsal.BoundText = strKdAsal
        txtSatuanJml.Text = Trim(strSatuanJml)

        strSQL = "SELECT MasterBarang.JmlTerkecil,MasterBarang.JmlJualTerkecil,SatuanJumlahK.SatuanJmlK FROM MasterBarang LEFT OUTER JOIN SatuanJumlahK ON MasterBarang.KdSatuanJmlK=SatuanJumlahK.KdSatuanJmlK WHERE MasterBarang.KdBarang='" & strkdbarang & "'"
        msubRecFO rs, strSQL
        txtJmlTerkecil.Text = rs("JmlTerkecil").Value
        txtJmlJualTerkecil.Text = rs("JmlJualTerkecil").Value
        If strkdbarang = "" Then
            MsgBox "Nama Barang belum dipilih", vbCritical, "Validasi"
            txtNamaBrg.Text = ""
            dgObatAlkes.SetFocus
            Exit Sub
        End If
        fraObatAlkes.Visible = False
        txtjml.SetFocus
    End If
    If KeyAscii = 27 Then
        fraObatAlkes.Visible = False
    End If

    Exit Sub
errError:
    msubPesanError
End Sub

Private Sub dgObatAlkes_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If dgObatAlkes.Visible = False Then Exit Sub
        txtNamaBrg.SetFocus
    End If
End Sub

Private Sub dgPaketPelayanan_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgPaketPelayanan
    WheelHook.WheelHook dgPaketPelayanan
End Sub

Private Sub dgPaketPelayanan_GotFocus()
    subInvisibleFram
End Sub

Private Sub dgPaketPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdHapus.SetFocus
End Sub

Private Sub dgPaketPelayanan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgPaketPelayanan.ApproxCount = 0 Then Exit Sub
    txtNamaPelayanan.Text = dgPaketPelayanan.Columns("Nama Pemeriksaan").Value
    strKodePelayananRS = dgPaketPelayanan.Columns("KdPelayananRS").Value
    txtNamaBrg.Text = dgPaketPelayanan.Columns("Nama Barang").Value
    strkdbarang = dgPaketPelayanan.Columns("KdBarang").Value
    dcAsal.BoundText = dgPaketPelayanan.Columns("KdAsal").Value
    txtSatuanJml.Text = dgPaketPelayanan.Columns("Satuan").Value
    txtjml.Text = dgPaketPelayanan.Columns("JmlBarang").Value
    strSQL = "SELECT MasterBarang.JmlTerkecil,MasterBarang.JmlJualTerkecil,SatuanJumlahK.SatuanJmlK FROM MasterBarang INNER JOIN SatuanJumlahK ON MasterBarang.KdSatuanJmlK=SatuanJumlahK.KdSatuanJmlK WHERE MasterBarang.KdBarang='" & strkdbarang & "'"
    msubRecFO rs, strSQL
    txtJmlTerkecil.Text = rs("JmlTerkecil").Value
    txtJmlJualTerkecil.Text = rs("JmlJualTerkecil").Value
    lblSatuanK.Caption = rs("SatuanJmlK").Value
    lblSatuanK2.Caption = rs("SatuanJmlK").Value
    strSQL = "SELECT distinct HargaBarang FROM V_HargaNetto1Barang WHERE KdBarang='" & strkdbarang & "' and kdAsal = '" & dcAsal.BoundText & "' and Satuan='" & dgPaketPelayanan.Columns(4).Value & "' "
    msubRecFO rs, strSQL
    txtHarga.Text = rs(0).Value
    txtNamaPelayanan.Enabled = False
    txtNamaBrg.Enabled = False
    subInvisibleFram
End Sub

Private Sub dgPelayanan_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgPelayanan
    WheelHook.WheelHook dgPelayanan
End Sub

Private Sub dgPelayanan_DblClick()
    Call dgPelayanan_KeyPress(13)
End Sub

Private Sub dgPelayanan_GotFocus()
    fraObatAlkes.Visible = False
End Sub

Private Sub dgPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlPelayanan = 0 Then Exit Sub
        Dim strkd As String
        strkd = dgPelayanan.Columns(2).Value
        txtNamaPelayanan.Text = dgPelayanan.Columns(1).Value
        strKodePelayananRS = strkd
        If strKodePelayananRS = "" Then
            MsgBox "Pilih dulu tindakan pelayanan Pasien", vbCritical, "Validasi"
            txtNamaPelayanan.Text = ""
            dgPelayanan.SetFocus
            Exit Sub
        End If
        fraPelayanan.Visible = False
    End If
    If KeyAscii = 27 Then
        fraPelayanan.Visible = False
    End If
End Sub

Private Sub dgPelayanan_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If dgPelayanan.Visible = False Then Exit Sub
        txtNamaPelayanan.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    On Error GoTo errLoad
    strSQL = "SELECT KdAsal,NamaAsal FROM AsalBarang  "
    msubRecFO rs, strSQL
    Set dcAsal.RowSource = rs
    dcAsal.BoundColumn = rs(0).Name
    dcAsal.ListField = rs(1).Name
    subLoadPaketLayanan
    Call PlayFlashMovie(Me)
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtjml_GotFocus()
    subInvisibleFram
End Sub

Private Sub txtjml_KeyPress(KeyAscii As Integer)
    SetKeyPressToNumber KeyAscii
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNamaBrg_Change()
    Call subLoadObatAlkes
End Sub

Private Sub txtNamaBrg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If dgObatAlkes.Visible = False Then Exit Sub
        dgObatAlkes.SetFocus
    End If
End Sub

Private Sub txtNamaBrg_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If intJmlObatAlkes = 0 Then Exit Sub
        dgObatAlkes.SetFocus
    End If
    If KeyAscii = 27 Then
        fraObatAlkes.Visible = False
    End If
End Sub

Private Sub txtNamaPelayanan_Change()
    strFilterPelayanan = "WHERE NamaPelayanan like '%" & txtNamaPelayanan.Text _
    & "%' AND KdRuangan='" & mstrKdRuangan & "'"
    strKodePelayananRS = ""
    fraPelayanan.Visible = True
    Call subLoadPelayanan
End Sub

Private Sub txtNamaPelayanan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If dgPelayanan.Visible = False Then Exit Sub
        dgPelayanan.SetFocus
    End If
End Sub

Private Sub txtNamaPelayanan_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If intJmlPelayanan = 0 Then Exit Sub
        dgPelayanan.SetFocus
    End If
    If KeyAscii = 27 Then
        fraPelayanan.Visible = False
    End If
End Sub

Private Sub txtParameter_Change()
    subLoadPaketLayanan "AND [Nama Pemeriksaan] LIKE '" & txtParameter.Text & "%'"
End Sub

Private Sub txtParameter_GotFocus()
    subInvisibleFram
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtParameter.SetFocus
    subLoadPaketLayanan "AND [Nama Pemeriksaan] LIKE '" & txtParameter.Text & "%'"

End Sub

Private Sub subLoadPelayanan()
    On Error GoTo errLoad
    strSQL = "SELECT JenisPelayanan,NamaPelayanan,KdPelayananRS FROM V_DetailPelayananMedis " & strFilterPelayanan
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlPelayanan = rs.RecordCount
    With dgPelayanan
        Set .DataSource = rs
        .Columns(0).Width = 2600
        .Columns(1).Width = 5000
        .Columns(2).Width = 0
    End With
    Exit Sub
errLoad:
    msubPesanError
    Set rs = Nothing
End Sub

Private Sub subLoadObatAlkes()
    On Error GoTo errLoad

    strSQL = "SELECT distinct JenisBarang,NamaBarang,AsalBarang,Satuan,HargaBarang,KdBarang,KdAsal FROM V_HargaNetto1Barang " & _
    " WHERE namabarang like '%" & txtNamaBrg.Text & "%' AND KdRuangan='" & mstrKdRuangan & "' "

    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    strkdbarang = ""
    intJmlObatAlkes = rs.RecordCount
    With dgObatAlkes
        Set .DataSource = rs
        .Columns(5).Width = 0
        .Columns(6).Width = 0
        .Columns(0).Width = 2000
        .Columns(1).Width = 2800
        .Columns(2).Width = 1000
        .Columns(3).Width = 800
        .Columns(3).Alignment = dbgCenter
        .Columns(4).Width = 0
        .Columns(4).Alignment = dbgRight
    End With
    fraObatAlkes.Visible = True
    Exit Sub
errLoad:
    msubPesanError
    Set rs = Nothing
End Sub

Private Sub subLoadPaketLayanan(Optional strKriteria As String)
    On Error GoTo errLoad
    Set rs = Nothing
    strSQL = "SELECT * FROM V_DaftarPaketLayananTindakan WHERE KdRuangan='" & mstrKdRuangan & "' " & strKriteria
    msubRecFO rs, strSQL
    Set dgPaketPelayanan.DataSource = rs
    With dgPaketPelayanan
        .Columns(0).Width = 3500
        .Columns(1).Width = 2500
        .Columns(2).Width = 1300
        .Columns(3).Width = 1000
        .Columns(4).Width = 1000
        .Columns(5).Width = 0
        .Columns(6).Width = 0
        .Columns(7).Width = 0
        .Columns(8).Width = 0
        .Columns(3).Alignment = dbgCenter
        .Columns(4).Alignment = dbgCenter
    End With
    Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub subInvisibleFram()
    fraObatAlkes.Visible = False
    fraPelayanan.Visible = False
End Sub

Private Sub subClearData()
    strKodePelayananRS = ""
    strkdbarang = ""
    txtNamaPelayanan.Text = ""
    txtNamaBrg.Text = ""
    txtNamaPelayanan.Enabled = True
    txtNamaBrg.Enabled = True
    dcAsal.Text = ""
    txtjml.Text = ""
    txtJmlTerkecil.Text = ""
    txtJmlJualTerkecil.Text = ""
    subInvisibleFram
    txtHarga.Text = ""
    txtNamaPelayanan.SetFocus
    txtSatuanJml.Text = ""
End Sub

Private Function sp_PaketLayanan(f_Status As String) As Boolean
    On Error GoTo errSp_PaketLayanan
    sp_PaketLayanan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, strKodePelayananRS)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, strkdbarang)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, dcAsal.BoundText)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("JmlBarang", adTinyInt, adParamInput, , txtjml.Text)
        .Parameters.Append .CreateParameter("SatuanJml", adChar, adParamInput, 1, Trim(txtSatuanJml.Text))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_PaketLayanan"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan data paket layanan", vbCritical, "Validasi"
            sp_PaketLayanan = False
            Set dbcmd = Nothing
        Else
            sp_PaketLayanan = True
            Call Add_HistoryLoginActivity("AUD_PaketLayanan")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
errSp_PaketLayanan:
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    msubPesanError
End Function

