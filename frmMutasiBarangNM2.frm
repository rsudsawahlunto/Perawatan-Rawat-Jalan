VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMutasiBarangNM2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Mutasi Barang Non Medis"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMutasiBarangNM2.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   10470
   Begin MSDataGridLib.DataGrid dgNamaPenerima 
      Height          =   2535
      Left            =   3120
      TabIndex        =   4
      Top             =   4680
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4471
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
   Begin VB.Frame Frame2 
      Caption         =   "Data Pengiriman"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   25
      Top             =   960
      Width           =   10455
      Begin VB.TextBox txtNoKirim 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         MaxLength       =   15
         TabIndex        =   0
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtNamaPenerima 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6960
         TabIndex        =   3
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txtKdUserPenerima 
         Height          =   315
         Left            =   9240
         TabIndex        =   26
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpTglKirim 
         Height          =   330
         Left            =   1680
         TabIndex        =   1
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
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
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   145162243
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSDataListLib.DataCombo dcRuanganPenerima 
         Height          =   330
         Left            =   3720
         TabIndex        =   2
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Kirim"
         Height          =   210
         Index           =   10
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Kirim"
         Height          =   210
         Index           =   9
         Left            =   1680
         TabIndex        =   29
         Top             =   240
         Width           =   750
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruangan Penerima"
         Height          =   210
         Index           =   11
         Left            =   3720
         TabIndex        =   28
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Penerima"
         Height          =   210
         Index           =   8
         Left            =   6960
         TabIndex        =   27
         Top             =   240
         Width           =   1260
      End
   End
   Begin MSDataGridLib.DataGrid dgCariBarang 
      Height          =   2535
      Left            =   2160
      TabIndex        =   6
      Top             =   4200
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4471
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
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   0
      TabIndex        =   20
      Top             =   3600
      Width           =   10455
      Begin VB.TextBox txtCariBarang 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   11
         Top             =   4200
         Width           =   3240
      End
      Begin MSDataGridLib.DataGrid dgMutasiBarang 
         Height          =   3735
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   9975
         _ExtentX        =   17595
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
         TabIndex        =   22
         Top             =   4245
         Width           =   900
      End
      Begin VB.Label lblJmlData 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Jumlah Barang"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   8985
         TabIndex        =   21
         Top             =   4260
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   3960
      TabIndex        =   14
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   5535
      TabIndex        =   13
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   7080
      TabIndex        =   12
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   8565
      TabIndex        =   15
      Top             =   8400
      Width           =   1695
   End
   Begin VB.Frame fraBarang 
      Height          =   1695
      Left            =   0
      TabIndex        =   16
      Top             =   1920
      Width           =   10455
      Begin VB.TextBox txtStok 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   9
         Top             =   1080
         Width           =   1320
      End
      Begin VB.TextBox txtKdBarang 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2880
         MaxLength       =   50
         TabIndex        =   23
         Text            =   "txtkdbarang"
         Top             =   0
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.TextBox txtNamaBarang 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   5
         Top             =   360
         Width           =   5520
      End
      Begin MSDataListLib.DataCombo dcAsalBarang 
         Height          =   330
         Left            =   1560
         TabIndex        =   7
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         Text            =   ""
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
      Begin VB.TextBox txtJmlBarang 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4440
         MaxLength       =   25
         TabIndex        =   8
         Top             =   1080
         Width           =   1320
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Stok Ruangan"
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Nama Barang"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Asal Barang"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   930
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Jumlah"
         Height          =   210
         Index           =   3
         Left            =   3720
         TabIndex        =   17
         Top             =   1080
         Width           =   555
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   31
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
      Left            =   8640
      Picture         =   "frmMutasiBarangNM2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmMutasiBarangNM2.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMutasiBarangNM2.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmMutasiBarangNM2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tempbolTampil As Boolean
Dim tempbolEdit As Boolean
Dim substrKdPegawai As String
Dim substrNoOrder As String

Private Sub cmdBatal_Click()
    On Error GoTo errLoad

    Call subKosong
    Call subLoadDcSource
    Call subLoadGridSource
    tempbolEdit = False
    dtpTglKirim.SetFocus

    Exit Sub
errLoad:
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errLoad

    If txtKdBarang.Text = "" Then
        MsgBox "Nama barang kosong", vbExclamation, "Validasi": txtNamaBarang.SetFocus: Exit Sub
    End If

    If MsgBox("Anda yakin akan menghapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub

    If sp_StockBarang(CDbl(txtStok.Text) + CDbl(txtJmlBarang.Text)) = False Then Exit Sub
    dbConn.Execute "DELETE MutasiBarangNonMedis WHERE NoKirim = '" & txtNoKirim.Text & "' AND KdBarang ='" & txtKdBarang.Text & "' AND KdAsal='" & dcAsalBarang.BoundText & "' "

    Call cmdBatal_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad

    If txtNamaBarang.Text = "" Then
        MsgBox "Nama barang kosong", vbExclamation, "Validasi": txtNamaBarang.SetFocus: Exit Sub
    End If
    If Periksa("datacombo", dcAsalBarang, "Asal barang kosong") = False Then Exit Sub

    If CDbl(txtJmlBarang.Text) > CDbl(txtStok.Text) Then
        MsgBox "Jumlah Barang tidak boleh melebihi Stok Ruangan", vbCritical, "Validasi"
        Exit Sub
    End If

    If sp_StrukKirim() = False Then Exit Sub
    If tempbolEdit = True Then
        If sp_StockBarang(CDbl(txtStok.Text + CDbl(dgMutasiBarang.Columns("JmlBarang")))) = False Then Exit Sub
    End If

    Call msubRecFO(rs, "select dbo.FB_TakeStokBrgNonMedis('" & mstrKdRuangan & "', '" & txtKdBarang & "','" & dcAsalBarang.BoundText & "') as stok")
    If rs.EOF = False Then txtStok.Text = rs(0).Value Else txtStok.Text = 0

    If sp_StockBarang(CDbl(txtStok.Text - txtJmlBarang.Text)) = False Then Exit Sub
    If sp_MutasiBarangNonMedis() = False Then Exit Sub

    Call cmdBatal_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcAsalBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(dcAsalBarang.Text)) = 0 Then txtJmlBarang.SetFocus: Exit Sub
        If dcAsalBarang.MatchedWithList = True Then txtJmlBarang.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "SELECT KdAsal, NamaAsal FROM AsalBarang where StatusEnabled='1' and NamaAsal LIKE '%" & dcAsalBarang.Text & "%'ORDER BY NamaAsal")
        If dbRst.EOF = True Then
            dcAsalBarang.Text = ""
            txtJmlBarang.SetFocus
            Exit Sub
        End If
        dcAsalBarang.BoundText = dbRst(0).Value
        dcAsalBarang.Text = dbRst(1).Value
    End If
End Sub

Private Sub dcRuanganPenerima_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    If KeyAscii = 13 Then
        If Len(Trim(dcRuanganPenerima.Text)) = 0 Then txtNamaPenerima.SetFocus: Exit Sub
        If dcRuanganPenerima.MatchedWithList = True Then txtNamaPenerima.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "SELECT KdRuangan, NamaRuangan FROM Ruangan WHERE NamaRuangan LIKE '%" & dcRuanganPenerima.Text & "%'")
        If dbRst.EOF = True Then
            dcRuanganPenerima.Text = ""
            txtNamaPenerima.SetFocus
            Exit Sub
        End If
        dcRuanganPenerima.BoundText = dbRst(0).Value
        dcRuanganPenerima.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dgCariBarang_DblClick()
    On Error GoTo errLoad

    With dgCariBarang
        If .ApproxCount = 0 Then Exit Sub
        txtKdBarang.Text = .Columns("KdBarang")
        txtNamaBarang.Text = .Columns("Nama Barang")
        dcAsalBarang.BoundText = .Columns("KdAsal")

        .Visible = False
    End With

    Call msubRecFO(rs, "select dbo.FB_TakeStokBrgNonMedis('" & mstrKdRuangan & "', '" & txtKdBarang & "','" & dcAsalBarang.BoundText & "') as stok")
    If rs.EOF = False Then txtStok.Text = rs(0).Value Else txtStok.Text = 0

    txtJmlBarang.SetFocus

    Exit Sub
errLoad:
End Sub

Private Sub dgCariBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call dgCariBarang_DblClick
End Sub

Private Sub dgMutasiBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaBarang.SetFocus
End Sub

Private Sub dgMutasiBarang_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errLoad

    With dgMutasiBarang
        If .ApproxCount = 0 Then Exit Sub
        tempbolEdit = True
        txtKdBarang.Text = .Columns("KdBarang")
        txtNamaBarang.Text = .Columns("Nama Barang")
        dcAsalBarang.BoundText = .Columns("KdAsal")
        txtJmlBarang.Text = .Columns("JmlBarang")
        txtNoKirim.Text = .Columns("NoKirim")
        dtpTglKirim.Value = .Columns("TglKirim")
        dcRuanganPenerima.BoundText = .Columns("KdRuanganTujuan")
        txtNamaPenerima.Text = .Columns("UserPenerima")
    End With
    Call msubRecFO(rs, "select dbo.FB_TakeStokBrgNonMedis('" & mstrKdRuangan & "', '" & txtKdBarang & "','" & dcAsalBarang.BoundText & "') as stok")
    If rs.EOF = False Then txtStok.Text = rs(0).Value Else txtStok.Text = CDbl(txtJmlBarang.Text)

    dgCariBarang.Visible = False
    dgNamaPenerima.Visible = False
    lblJmlData.Caption = dgMutasiBarang.Bookmark & " / " & dgMutasiBarang.ApproxCount & " Data"

    Exit Sub
errLoad:
End Sub

Private Sub dgNamaPenerima_DblClick()
    On Error GoTo errLoad
    If dgNamaPenerima.ApproxCount = 0 Then Exit Sub
    txtKdUserPenerima.Text = dgNamaPenerima.Columns("IdPegawai").Value
    txtNamaPenerima.Text = dgNamaPenerima.Columns("Nama Pemeriksa").Value
    substrKdPegawai = dgNamaPenerima.Columns("IdPegawai").Value
    dgNamaPenerima.Visible = False
    txtNamaBarang.SetFocus
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dgNamaPenerima_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call dgNamaPenerima_DblClick
End Sub

Private Sub dtpTglKirim_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcRuanganPenerima.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call subKosong
    Call subLoadDcSource
    Call subLoadGridSource
    Exit Sub
errLoad:
End Sub

Private Sub txtCariBarang_Change()
    On Error GoTo errLoad

    Call subLoadGridSource

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtLokasi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then dgMutasiBarang.SetFocus
End Sub

Private Sub txtJmlBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtJmlBarang_LostFocus()
    txtJmlBarang.Text = IIf(Val(txtJmlBarang) = 0, 0, Format(txtJmlBarang, "#,###"))
End Sub

Private Sub txtLokasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtNamaBarang_Change()
    On Error GoTo errLoad

    If tempbolTampil = True Then Exit Sub
    Call subCariBarang

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtNamaBarang_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyDown Then If dgCariBarang.Visible = True Then dgCariBarang.SetFocus
End Sub

Private Sub txtNamaBarang_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then If dgCariBarang.Visible = True Then dgCariBarang.SetFocus Else dcAsalBarang.SetFocus
    If KeyAscii = 27 Then If dgCariBarang.Visible = True Then dgCariBarang.Visible = False
End Sub

Private Sub subKosong()
    txtNoKirim.Text = ""
    dtpTglKirim.Value = Now
    dcRuanganPenerima.Text = ""
    txtNamaPenerima.Text = ""
    dgNamaPenerima.Visible = False

    txtKdBarang.Text = ""
    txtNamaBarang.Text = ""
    txtCariBarang.Text = ""
    dcAsalBarang.BoundText = ""
    txtJmlBarang.Text = 0
    txtStok.Text = 0
    dgCariBarang.Visible = False
End Sub

Private Sub subLoadDcSource()
    On Error GoTo errLoad

    Call msubDcSource(dcAsalBarang, rs, "SELECT KdAsal, NamaAsal FROM AsalBarang where StatusEnabled='1'ORDER BY NamaAsal")
    Call msubDcSource(dcRuanganPenerima, rs, "SELECT KdRuangan, NamaRuangan FROM Ruangan where StatusEnabled='1'ORDER BY NamaRuangan")

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subCariBarang()
    On Error GoTo errLoad

    strSQL = "SELECT  [Nama Barang], Satuan, DetailJenisBarang AS [Jenis Barang], KdBarang, KdAsal FROM V_CariBarangNonMedis " & _
    " WHERE [Nama Barang] LIKE '%" & txtNamaBarang.Text & "%' AND KdRuangan = '" & mstrKdRuangan & "'" & _
    " ORDER BY [Nama Barang]"
    Call msubRecFO(rs, strSQL)
    Set dgCariBarang.DataSource = rs
    With dgCariBarang
        .Columns("Nama Barang").Width = 2900
        .Columns("Satuan").Width = 1000
        .Columns("Jenis Barang").Width = 1440
        .Columns("KdBarang").Width = 0
        .Columns("KdAsal").Width = 0

        .Height = 2390
        .Top = 2640
        .Left = 1560
        .Visible = True
    End With

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadGridSource()
    On Error GoTo errLoad
    Dim i As Integer

    tempbolTampil = True
    strSQL = "SELECT * " & _
    " FROM V_MutasiBarangNonMedis " & _
    " WHERE [Nama Barang] LIKE '%" & txtCariBarang & "%' AND KdRuangan = '" & mstrKdRuangan & "'"
    Call msubRecFO(rs, strSQL)
    Set dgMutasiBarang.DataSource = rs
    With dgMutasiBarang
        For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next i
        .Columns("Nama Barang").Width = 2200
        .Columns("Asal").Width = 1000

        .Columns("JmlBarang").Width = 900
        .Columns("RuanganTujuan").Width = 2000
    End With
    lblJmlData.Caption = 0 & " / " & dgMutasiBarang.ApproxCount & " Data"
    tempbolTampil = False

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Function sp_StrukKirim() As Boolean
    On Error GoTo errLoad
    sp_StrukKirim = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoKirim", adChar, adParamInput, 10, txtNoKirim.Text)
        .Parameters.Append .CreateParameter("TglKirim", adDate, adParamInput, , Format(dtpTglKirim.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, IIf(substrNoOrder = "", Null, substrNoOrder))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, dcRuanganPenerima.BoundText)
        .Parameters.Append .CreateParameter("IdUserPenerima", adChar, adParamInput, 10, txtKdUserPenerima.Text)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("OutputNoKirim", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "Add_StrukKirim"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data struk kirim antar ruangan", vbCritical, "Validasi"
            sp_StrukKirim = False
        Else
            txtNoKirim.Text = .Parameters("OutputNoKirim").Value
        End If
    End With
    Exit Function
errLoad:
    Call msubPesanError
    sp_StrukKirim = False
End Function

Private Function sp_MutasiBarangNonMedis() As Boolean
    On Error GoTo errLoad

    sp_MutasiBarangNonMedis = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoKirim", adChar, adParamInput, 10, txtNoKirim.Text)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, txtKdBarang.Text)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, dcAsalBarang.BoundText)
        .Parameters.Append .CreateParameter("JmlBarang", adInteger, adParamInput, , txtJmlBarang.Text)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")

        .ActiveConnection = dbConn
        .CommandText = "AUD_MutasiBarangNonMedis"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_MutasiBarangNonMedis = False
        End If
    End With

    Exit Function
errLoad:
    Call msubPesanError
End Function

Private Function sp_StockBarang(f_JmlStok As Double) As Boolean
    On Error GoTo errLoad

    sp_StockBarang = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, txtKdBarang.Text)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, dcAsalBarang.BoundText)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("JmlMin", adDouble, adParamInput, , 1)
        .Parameters.Append .CreateParameter("JmlStok", adDouble, adParamInput, , f_JmlStok)
        .Parameters.Append .CreateParameter("Lokasi", adVarChar, adParamInput, 12, Null)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")

        .ActiveConnection = dbConn
        .CommandText = "AUD_StokBarangNonMedis"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_StockBarang = False
        End If
    End With

    Exit Function
errLoad:
    Call msubPesanError
End Function

Private Sub txtNamaPenerima_Change()
    On Error GoTo errLoad
    Dim i As Integer

    strSQL = " SELECT [Nama Pemeriksa], JK, [Jenis Pemeriksa], IdPegawai " & _
    " From V_DaftarPemeriksaPasien" & _
    " where [Nama Pemeriksa] like '" & txtNamaPenerima.Text & "%' " & _
    " ORDER BY [Nama Pemeriksa], [Jenis Pemeriksa]"
    Call msubRecFO(dbRst, strSQL)

    Set dgNamaPenerima.DataSource = dbRst
    With dgNamaPenerima
        .Columns("Nama Pemeriksa").Width = 2000
        .Columns("JK").Width = 360
        .Columns("Jenis Pemeriksa").Width = 1500
        .Columns("IdPegawai").Width = 0
        .Columns("JK").Alignment = dbgCenter

        .Top = 1800
        .Left = 5880
    End With
    dgNamaPenerima.Visible = True

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtNamaPenerima_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If dgNamaPenerima.Visible = True Then dgNamaPenerima.SetFocus Else txtNamaBarang.SetFocus
    If KeyAscii = 27 Then dgNamaPenerima.Visible = False
End Sub

