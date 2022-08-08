VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmPemakaianAlkesNonCharge 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Pemakaian Bahan"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10860
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPemakaianAlkesNonCharge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   10860
   Begin VB.Frame fraObatAlkes 
      Caption         =   "Data Harga Barang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   240
      TabIndex        =   39
      Top             =   3720
      Visible         =   0   'False
      Width           =   9375
      Begin MSDataGridLib.DataGrid dgObatAlkes 
         Height          =   2535
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
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
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   21
      Top             =   5880
      Width           =   10935
      Begin VB.CommandButton cmdtambah 
         Caption         =   "&Tambah"
         Height          =   465
         Left            =   3600
         TabIndex        =   43
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   465
         Left            =   5520
         TabIndex        =   42
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   7320
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   9120
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame8 
      Height          =   3855
      Left            =   0
      TabIndex        =   18
      Top             =   2040
      Width           =   10935
      Begin VB.TextBox txtsatuam 
         Height          =   315
         Left            =   6000
         TabIndex        =   33
         Top             =   2760
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtkdbarang 
         Height          =   315
         Left            =   4920
         TabIndex        =   32
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtasalbarang 
         Height          =   315
         Left            =   4200
         TabIndex        =   35
         Top             =   2760
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtHargaSatuan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1200
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo dcJnsPelayanan 
         Height          =   330
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.TextBox txtNamaBrg 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   4455
      End
      Begin VB.TextBox txtjml 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4800
         MaxLength       =   4
         TabIndex        =   12
         Top             =   1200
         Width           =   615
      End
      Begin MSFlexGridLib.MSFlexGrid fgAlkes 
         Height          =   1935
         Left            =   240
         TabIndex        =   15
         Top             =   1680
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   3413
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   0
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin VB.TextBox txtkdAsal 
         Height          =   315
         Left            =   2160
         TabIndex        =   34
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSDataListLib.DataCombo dcPelayanan 
         Height          =   330
         Left            =   3960
         TabIndex        =   8
         Top             =   480
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcTglPelayanan 
         Height          =   330
         Left            =   8400
         TabIndex        =   9
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Pelayanan "
         Height          =   210
         Index           =   1
         Left            =   8400
         TabIndex        =   40
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   210
         Index           =   2
         Left            =   6960
         TabIndex        =   38
         Top             =   960
         Width           =   420
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Harga Satuan"
         Height          =   210
         Index           =   1
         Left            =   5520
         TabIndex        =   37
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pemeriksaan"
         Height          =   210
         Index           =   0
         Left            =   3960
         TabIndex        =   36
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pemeriksaan"
         Height          =   210
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Nama Barang"
         Height          =   210
         Left            =   240
         TabIndex        =   20
         Top             =   960
         Width           =   1065
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Jumlah"
         Height          =   210
         Index           =   0
         Left            =   4800
         TabIndex        =   19
         Top             =   960
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Pasien"
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
      TabIndex        =   22
      Top             =   1080
      Width           =   10935
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   2880
         MaxLength       =   50
         TabIndex        =   2
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   7080
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         Caption         =   "Umur"
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
         Left            =   8280
         TabIndex        =   23
         Top             =   200
         Width           =   2535
         Begin VB.TextBox txtThn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            MaxLength       =   6
            TabIndex        =   4
            Top             =   250
            Width           =   375
         End
         Begin VB.TextBox txtBln 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   960
            MaxLength       =   6
            TabIndex        =   5
            Top             =   250
            Width           =   375
         End
         Begin VB.TextBox txtHr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            MaxLength       =   6
            TabIndex        =   6
            Top             =   250
            Width           =   375
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            Height          =   210
            Left            =   600
            TabIndex        =   26
            Top             =   300
            Width           =   285
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            Height          =   210
            Left            =   1440
            TabIndex        =   25
            Top             =   300
            Width           =   240
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            Height          =   210
            Left            =   2280
            TabIndex        =   24
            Top             =   300
            Width           =   165
         End
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         MaxLength       =   10
         TabIndex        =   0
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   1800
         TabIndex        =   30
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lblNamaPasien 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   2880
         TabIndex        =   29
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label lblJnsKlm 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   7080
         TabIndex        =   28
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   41
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
      Left            =   9120
      Picture         =   "frmPemakaianAlkesNonCharge.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPemakaianAlkesNonCharge.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPemakaianAlkesNonCharge.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmPemakaianAlkesNonCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim subarrKdBarang() As String
Dim subarrKdAsal() As String
Dim subarrSatuanJml() As String
Dim subintJmlArray As Integer
Dim subcurHargaSatuan As Currency

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad
    Dim i As Integer

    If fgAlkes.TextMatrix(1, 0) = "" Then MsgBox "Data barang harus diisi", vbExclamation, "Validasi": Exit Sub
    If Periksa("datacombo", dcTglPelayanan, "Tanggal pelayanan kosong") = False Then Exit Sub

    For i = 1 To fgAlkes.Rows - 2
        If sp_PemakaianAlkesNonCharge(fgAlkes.TextMatrix(i, 9), fgAlkes.TextMatrix(i, 7), _
            fgAlkes.TextMatrix(i, 8), fgAlkes.TextMatrix(i, 3), fgAlkes.TextMatrix(i, 4), _
            fgAlkes.TextMatrix(i, 5)) = False Then Exit Sub
        Next i
        Call Add_HistoryLoginActivity("AU_PemakaianObatAlkesNonCharge")
        cmdSimpan.Enabled = False

        Exit Sub
errLoad:
        Call msubPesanError
End Sub

Private Sub cmdTambah_Click()
    Dim i As Integer

    If Periksa("datacombo", dcPelayanan, "Nama pelayanan kosong") = False Then Exit Sub
    If Periksa("text", txtNamaBrg, "Nama barang kosong") = False Then Exit Sub
    If Periksa("nilai", txtjml, "Jumlah barang kosong") = False Then Exit Sub

    With fgAlkes
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 9) = dcPelayanan.BoundText And _
                .TextMatrix(i, 7) = txtKdBarang.Text And _
                .TextMatrix(i, 8) = txtKdAsal.Text And _
                .TextMatrix(i, 3) = txtsatuam.Text Then Exit Sub
            Next i
        End With

        ' cek stok barang
        Set dbcmd = New ADODB.Command
        Set dbcmd.ActiveConnection = dbConn
        With dbcmd
            .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
            .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, txtKdBarang)
            .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, txtKdAsal)
            .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
            .Parameters.Append .CreateParameter("Satuan", adChar, adParamInput, 1, txtsatuam)
            .Parameters.Append .CreateParameter("JmlBrg", adInteger, adParamInput, , CInt(txtjml))
            .Parameters.Append .CreateParameter("OutputPesan", adChar, adParamOutput, 1, Null)

            .CommandText = "dbo.Check_StokBarangRuangan"
            .CommandType = adCmdStoredProc
            .Execute

            If .Parameters("OutputPesan") = "T" Then
                deleteADOCommandParameters dbcmd
                MsgBox "Stok barang tidak cukup", vbExclamation, "Validasi"
                txtjml.SetFocus
                Exit Sub
            End If
            deleteADOCommandParameters dbcmd
        End With

        With fgAlkes
            .TextMatrix(.Rows - 1, 0) = dcPelayanan.Text
            .TextMatrix(.Rows - 1, 1) = txtNamaBrg.Text
            .TextMatrix(.Rows - 1, 2) = txtAsalBarang.Text
            .TextMatrix(.Rows - 1, 3) = txtsatuam.Text
            .TextMatrix(.Rows - 1, 4) = txtjml.Text

            strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "','" & txtKdAsal.Text & "', " & CCur(txthargasatuan) & ")  as HargaSatuan"
            Call msubRecFO(rs, strSQL)
            If rs.EOF = True Then subcurHargaSatuan = 0 Else subcurHargaSatuan = rs(0).Value
            .TextMatrix(.Rows - 1, 5) = subcurHargaSatuan

            .TextMatrix(.Rows - 1, 6) = txtTotal.Text
            .TextMatrix(.Rows - 1, 7) = txtKdBarang.Text
            .TextMatrix(.Rows - 1, 8) = txtKdAsal.Text
            .TextMatrix(.Rows - 1, 9) = dcPelayanan.BoundText
            .TextMatrix(.Rows - 1, 10) = mstrKdRuangan
            .Rows = .Rows + 1
        End With

        txtKdBarang.Text = ""
        txtKdAsal.Text = ""
        txtsatuam.Text = ""

        txtNamaBrg.Text = ""
        txtjml.Text = 0
        txthargasatuan.Text = 0
        txtTotal.Text = 0
        txtNamaBrg.SetFocus
End Sub

Private Sub Cmdtambah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dcJnsPelayanan.SetFocus
    End If
End Sub

Private Sub cmdTutup_Click()
    If cmdSimpan.Enabled = True Then
        If MsgBox("Simpan data Pemakaian Bahan", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub cmdHapus_Click()
    Dim i As Integer
    With fgAlkes
        If .Row = .Rows Then Exit Sub
        If .Row = 0 Then Exit Sub

        If .TextMatrix(.Row, 11) = LCase("ada") Then
            MsgBox "Data yang sudah pernah diinput, tidak bisa dihapus", vbExclamation, "Validasi"
            Exit Sub
        End If

        If .Rows = 2 Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
            Next i
            Exit Sub
        Else
            .RemoveItem .Row
        End If
    End With
End Sub

Private Sub dcJnsPelayanan_GotFocus()
    On Error GoTo errLoad
    Dim tempKode As String

    tempKode = dcJnsPelayanan.BoundText
    strSQL = "select distinct kdjnspelayanan, JenisPemeriksaan from V_RiwayatPemeriksaanPasien where NoPendaftaran = '" & txtNoPendaftaran.Text & "' AND KdRuangan = '" & mstrKdRuangan & "' and StatusEnabled='1'"
    Call msubDcSource(dcJnsPelayanan, dbRst, strSQL)
    dcJnsPelayanan.BoundText = tempKode

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcJnsPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcJnsPelayanan.MatchedWithList = True Then dcPelayanan.SetFocus
        strSQL = "select distinct kdjnspelayanan, JenisPemeriksaan from V_RiwayatPemeriksaanPasien where NoPendaftaran = '" & txtNoPendaftaran.Text & "' AND KdRuangan = '" & mstrKdRuangan & "' and StatusEnabled='1' and (JenisPemeriksaan LIKE '%" & dcJnsPelayanan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcJnsPelayanan.Text = ""
            dcPelayanan.SetFocus
            Exit Sub
        End If
        dcJnsPelayanan.BoundText = rs(0).Value
        dcJnsPelayanan.Text = rs(1).Value
    End If
End Sub

Private Sub dcPelayanan_GotFocus()
    On Error GoTo errLoad
    Dim tempKode As String

    tempKode = dcPelayanan.BoundText
    strSQL = "select distinct kdpelayananrs,NamaPemeriksaan from V_RiwayatPemeriksaanPasien where kdjnspelayanan = '" & dcJnsPelayanan.BoundText & "' AND NoPendaftaran = '" & txtNoPendaftaran.Text & "' AND KdRuangan = '" & mstrKdRuangan & "'and StatusEnabled='1'"
    Call msubDcSource(dcPelayanan, dbRst, strSQL)
    dcPelayanan.BoundText = tempKode

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcPelayanan.MatchedWithList = True Then dcTglPelayanan.SetFocus
        strSQL = "select distinct kdpelayananrs,NamaPemeriksaan from V_RiwayatPemeriksaanPasien where kdjnspelayanan = '" & dcJnsPelayanan.BoundText & "' AND NoPendaftaran = '" & txtNoPendaftaran.Text & "' AND KdRuangan = '" & mstrKdRuangan & "'and StatusEnabled='1' and (NamaPemeriksaan LIKE '%" & dcPelayanan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcPelayanan.Text = ""
            dcTglPelayanan.SetFocus
            Exit Sub
        End If
        dcPelayanan.BoundText = rs(0).Value
        dcPelayanan.Text = rs(1).Value
    End If
End Sub

Private Sub dcTglPelayananx()
    On Error GoTo errLoad
    Dim i As Integer

    If dcTglPelayanan.MatchedWithList = False Then Exit Sub
    strSQL = "SELECT NamaPemeriksaan, NamaBarang, AsalBarang, SatuanJml, JmlBarang, HargaSatuan, JmlBarang * HargaSatuan AS Total, KdBarang, KdAsal, KdPelayananRS, KdRuangan" & _
    " FROM V_PemakaianObatAlkesNonCharge" & _
    " WHERE (TglPelayanan = '" & Format(dcTglPelayanan.BoundText, "yyyy/MM/dd HH:mm:ss") & "') "
    Call msubRecFO(dbRst, strSQL)

    ReDim Preserve subarrKdBarang(dbRst.RecordCount)
    ReDim Preserve subarrKdAsal(dbRst.RecordCount)
    ReDim Preserve subarrSatuanJml(dbRst.RecordCount)
    subintJmlArray = dbRst.RecordCount

    If dbRst.RecordCount > 0 Then
        For i = 1 To dbRst.RecordCount
            fgAlkes.Rows = fgAlkes.Rows + 1
            With fgAlkes
                .TextMatrix(i, 0) = dbRst("NamaPemeriksaan").Value
                .TextMatrix(i, 1) = dbRst("NamaBarang").Value
                .TextMatrix(i, 2) = dbRst("AsalBarang").Value
                .TextMatrix(i, 3) = dbRst("SatuanJml").Value
                .TextMatrix(i, 4) = dbRst("JmlBarang").Value
                .TextMatrix(i, 5) = dbRst("HargaSatuan").Value
                .TextMatrix(i, 6) = dbRst("Total").Value
                .TextMatrix(i, 7) = dbRst("KdBarang").Value
                .TextMatrix(i, 8) = dbRst("KdAsal").Value
                .TextMatrix(i, 9) = dbRst("KdPelayananRS").Value
                .TextMatrix(i, 10) = dbRst("KdRuangan").Value
                .TextMatrix(i, 11) = "ada"

                subarrKdBarang(i) = dbRst("KdBarang").Value
                subarrKdAsal(i) = dbRst("KdAsal").Value
                subarrSatuanJml(i) = dbRst("SatuanJml").Value
            End With
            dbRst.MoveNext
        Next i
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcTglPelayanan_GotFocus()
    On Error GoTo errLoad
    Dim tempKode As String

    tempKode = dcTglPelayanan.BoundText
    strSQL = "select distinct [Tgl. Periksa], [Tgl. Periksa] as Alias " & _
    " from V_RiwayatPemeriksaanPasien " & _
    " where kdjnspelayanan = '" & dcJnsPelayanan.BoundText & "' AND NoPendaftaran = '" & txtNoPendaftaran.Text & "' AND KdRuangan = '" & mstrKdRuangan & "' AND KdPelayananRS = '" & dcPelayanan.BoundText & "'"
    Call msubDcSource(dcTglPelayanan, dbRst, strSQL)
    dcTglPelayanan.BoundText = tempKode

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcTglPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcTglPelayanan.MatchedWithList = True Then txtNamaBrg.SetFocus
        strSQL = "select distinct [Tgl. Periksa], [Tgl. Periksa] as Alias " & _
        " from V_RiwayatPemeriksaanPasien " & _
        " where kdjnspelayanan = '" & dcJnsPelayanan.BoundText & "' AND NoPendaftaran = '" & txtNoPendaftaran.Text & "' AND KdRuangan = '" & mstrKdRuangan & "' AND KdPelayananRS = '" & dcPelayanan.BoundText & "' and ([Tgl. Periksa] LIKE '%" & dcTglPelayanan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcTglPelayanan.Text = ""
            txtNamaBrg.SetFocus
            Exit Sub
        End If
        dcTglPelayanan.BoundText = rs(0).Value
        dcTglPelayanan.Text = rs(1).Value
        Call dcTglPelayananx
    End If
End Sub

Private Sub dgObatAlkes_DblClick()
    On Error GoTo a
    txtKdBarang.Text = dgObatAlkes.Columns("KdBarang").Value
    txtsatuam.Text = dgObatAlkes.Columns("Satuan").Value
    txtKdAsal.Text = dgObatAlkes.Columns("KdAsal").Value
    txtAsalBarang.Text = dgObatAlkes.Columns("AsalBarang").Value
    txthargasatuan.Text = dgObatAlkes.Columns("HargaBarang").Value
    txtNamaBrg.Text = dgObatAlkes.Columns("NamaBarang").Value
    fraObatAlkes.Visible = False
    txtjml.SetFocus
a:
    MsgBox "Maaf data obat alkes tidak ada", vbCritical, "Validasi"
End Sub

Private Sub dgObatAlkes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call dgObatAlkes_DblClick
    End If
End Sub

Private Sub dgObatAlkes_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If dgObatAlkes.Visible = False Then Exit Sub
        txtNamaBrg.SetFocus
    End If
End Sub

Private Sub fgAlkes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdHapus.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    Call centerForm(Me, MDIUtama)

    Call subLoadGridSource
    Call PlayFlashMovie(Me)
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtHargaSatuan_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtTotal.SetFocus
End Sub

Private Sub txtjml_Change()
    Dim Qty As Integer
    If txtjml = "" Then
        Qty = 0
        txtTotal.Text = CDbl(Qty) * Val(txthargasatuan.Text)
    Else
        Qty = txtjml
        txtTotal.Text = CDbl(Qty) * Val(txthargasatuan.Text)
    End If
End Sub

Private Sub txtjml_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then cmdTambah.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtjml_LostFocus()
    txtjml.Text = Val(txtjml.Text)
End Sub

Private Sub txtNamaBrg_Change()
    Dim s As String
    Dim i As Integer
    txtjml.Text = 0
    If txtNamaBrg = "" Then
        fraObatAlkes.Visible = False
        Exit Sub
    End If

    fraObatAlkes.Top = 3600
    fraObatAlkes.Left = 240
    fraObatAlkes.Visible = True

    strSQL = "select * from V_HargaNettoBarang where NamaBarang like '" & txtNamaBrg & "%' AND kdruangan = '" & mstrKdRuangan & "' AND KdKelompokPasien = '" & mstrKdJenisPasien & "' AND IdPenjamin = '" & mstrKdPenjaminPasien & "'"
    Call msubRecFO(dbRst, strSQL)

    Set dgObatAlkes.DataSource = dbRst
    With dgObatAlkes
        For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next i
        .Columns("JenisBarang").Width = 1200
        .Columns("NamaBarang").Width = 3165
        .Columns("AsalBarang").Width = 1000
        .Columns("JenisPasien").Width = 1100
        .Columns("Satuan").Width = 675
        .Columns("HargaBarang").Width = 1200
        .Columns("HargaBarang").NumberFormat = "#,###"
        .Columns("HargaBarang").Alignment = dbgRight
    End With
End Sub

Private Sub txtNamaBrg_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown Then
        If dgObatAlkes.Visible = False Then Exit Sub
        dgObatAlkes.SetFocus
    End If
End Sub

Private Sub txtNamaBrg_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then
        If fraObatAlkes.Visible = True Then
            dgObatAlkes.SetFocus
        Else
            txtjml.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        fraObatAlkes.Visible = False
    End If
End Sub

Private Sub subLoadGridSource()
    strSQL = "select * from V_HargaBarang where kdruangan = '" & mstrKdRuangan & "' "
    Call msubRecFO(dbRst, strSQL)
    Set dgObatAlkes.DataSource = dbRst
    dgObatAlkes.Columns(0).Width = 2000
    dgObatAlkes.Columns(1).Width = 2000
    dgObatAlkes.Columns(2).Width = 2000
    dgObatAlkes.Columns(3).Width = 2000
    dgObatAlkes.Columns(4).Width = 2000
    dgObatAlkes.Columns(5).Width = 2000
    dgObatAlkes.Columns(6).Width = 0
    dgObatAlkes.Columns(7).Width = 0
    dgObatAlkes.Columns(8).Width = 0

    With fgAlkes
        .Clear
        .Rows = 2
        .Cols = 12
        .TextMatrix(0, 0) = "Nama Pemeriksaan"
        .ColWidth(0) = 2500
        .TextMatrix(0, 1) = "Nama Barang"
        .ColWidth(1) = 2500
        .TextMatrix(0, 2) = "Asal Barang"
        .ColWidth(2) = 1200
        .TextMatrix(0, 3) = "Satuan"
        .ColWidth(3) = 800
        .TextMatrix(0, 4) = "Jumlah"
        .ColWidth(4) = 800
        .TextMatrix(0, 5) = "Harga Satuan"
        .ColWidth(5) = 1200
        .TextMatrix(0, 6) = "Total Harga"
        .ColWidth(6) = 1200

        .ColWidth(7) = 0 'KdBarang
        .ColWidth(8) = 0 'KdAsal
        .ColWidth(9) = 0 'KdPelayanRS
        .ColWidth(10) = 0 'KdRuangan
        .ColWidth(11) = 0 'jk udah ada diisi ada. jd ngak bisa dihapus
    End With
End Sub

Private Function sp_PemakaianAlkesNonCharge(f_KdPelayananRS As String, _
    f_KdBarang As String, f_KdAsal As String, f_Satuan As String, _
    f_JmlBarang As Integer, f_Harga As Currency) As Boolean

    sp_PemakaianAlkesNonCharge = True
    Set dbcmd = Nothing
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(dcTglPelayanan.BoundText, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, f_KdPelayananRS)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("Satuan", adChar, adParamInput, 1, f_Satuan)
        .Parameters.Append .CreateParameter("JmlBrg", adInteger, adParamInput, , CInt(f_JmlBarang))
        .Parameters.Append .CreateParameter("HargaSatuan", adCurrency, adParamInput, , CCur(f_Harga))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AU_PemakaianObatAlkesNonCharge"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_PemakaianAlkesNonCharge = False
        End If
    End With

    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
End Function

Private Sub txtTotal_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then cmdTambah.SetFocus
End Sub

