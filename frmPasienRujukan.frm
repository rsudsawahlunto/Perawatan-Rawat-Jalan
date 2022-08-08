VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPasienRujukan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Pasien Konsul ke Unit Lain"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10725
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPasienRujukan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   10725
   Begin VB.Frame fraDokter 
      Caption         =   "Data Dokter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   28
      Top             =   1080
      Visible         =   0   'False
      Width           =   10815
      Begin MSDataGridLib.DataGrid dgDokter 
         Height          =   1455
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   2566
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
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   29
      Top             =   3960
      Width           =   10815
      Begin VB.CommandButton cmdInputPermintaan 
         Caption         =   "&Input Permintaan Radiologi"
         Enabled         =   0   'False
         Height          =   340
         Left            =   3840
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   340
         Left            =   8640
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   340
         Left            =   6360
         TabIndex        =   13
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   0
      TabIndex        =   24
      Top             =   2400
      Width           =   10815
      Begin VB.TextBox txtDokter 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7680
         TabIndex        =   10
         Top             =   480
         Width           =   2895
      End
      Begin MSDataListLib.DataCombo dcInstalasi 
         Height          =   330
         Left            =   2280
         TabIndex        =   8
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcRuangan 
         Height          =   330
         Left            =   5160
         TabIndex        =   9
         Top             =   480
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker dtpTglDirujuk 
         Height          =   330
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   120717315
         UpDown          =   -1  'True
         CurrentDate     =   38077
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Dirujuk"
         Height          =   210
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dokter Pengirim"
         Height          =   210
         Left            =   7680
         TabIndex        =   27
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Instalasi Tujuan"
         Height          =   210
         Left            =   2280
         TabIndex        =   26
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ruangan Tujuan"
         Height          =   210
         Left            =   5160
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   1455
      Left            =   0
      TabIndex        =   15
      Top             =   960
      Width           =   10815
      Begin VB.Frame Frame5 
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
         Height          =   580
         Left            =   7320
         TabIndex        =   16
         Top             =   360
         Width           =   3255
         Begin VB.TextBox txtHari 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2400
            MaxLength       =   6
            TabIndex        =   6
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtBln 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1500
            MaxLength       =   6
            TabIndex        =   5
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtThn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   600
            MaxLength       =   6
            TabIndex        =   4
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2970
            TabIndex        =   19
            Top             =   270
            Width           =   150
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2070
            TabIndex        =   18
            Top             =   270
            Width           =   210
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1155
            TabIndex        =   17
            Top             =   270
            Width           =   240
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
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3240
         TabIndex        =   2
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5880
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   1920
         TabIndex        =   22
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   3240
         TabIndex        =   21
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   5880
         TabIndex        =   20
         Top             =   360
         Width           =   1065
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
      Left            =   9000
      Picture         =   "frmPasienRujukan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPasienRujukan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPasienRujukan.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmPasienRujukan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFilterDokter As String

Private Sub cmdInputPermintaan_Click()
    On Error GoTo hell
    Dim adoCommand As New ADODB.Command
    If dcInstalasi.Text = "" Then
        MsgBox "Instalasi Tujuan Belum Diisi !", vbCritical, "Konfirmasi"
        dcInstalasi.SetFocus
        Exit Sub
    End If
    If dcRuangan.Text = "" Then
        MsgBox "Ruangan Tujuan Belum Diisi !", vbCritical, "Konfirmasi"
        dcRuangan.SetFocus
        Exit Sub
    End If
    If mstrKdDokter = "" Then
        MsgBox "Nama Dokter Perujuk Belum Diisi !", vbCritical, "Konfirmasi"
        txtDokter.SetFocus
        Exit Sub
    End If
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtnocm)
        .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, IIf(mstrFormPengirim <> "frmDaftarPasienRJRIIGD", strNKdRuangan, mstrKdRuanganPasien))
        .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, dcRuangan.BoundText)
        .Parameters.Append .CreateParameter("IdDokterPerujuk", adChar, adParamInput, 10, mstrKdDokter)
        .Parameters.Append .CreateParameter("TglDirujuk", adDate, adParamInput, , Format(dtpTglDirujuk.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PasienRujukan"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Data", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Add_PasienRujukan")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    frmPasienRujukan2.Show
    Unload Me
    Exit Sub
hell:
    MsgBox "Maaf data sudah ada", vbCritical, "Validasi"
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell
    Dim adoCommand As New ADODB.Command
    If Periksa("datacombo", dcInstalasi, "Instalasi Tujuan Belum Diisi !") = False Then Exit Sub
    If Periksa("datacombo", dcRuangan, "Ruangan Tujuan Belum Diisi !") = False Then Exit Sub
    If Periksa("Text", txtDokter, "Nama Dokter Perujuk Belum Diisi !") = False Then Exit Sub
'    If dcInstalasi.BoundText = "" Then MsgBox "Instalasi Tujuan Belum Diisi !", vbExclamation, "Validasi": Exit Sub
'
'    If dcRuangan.BoundText = "" Then MsgBox "Ruangan Tujuan Belum Diisi !", vbExclamation, "Validasi": Exit Sub
'
'    If txtDokter.Text = "" Then MsgBox "Nama Dokter Perujuk Belum Diisi !", vbExclamation, "Validasi": Exit Sub
        
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtnocm)
        .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, IIf(mstrFormPengirim <> "frmDaftarPasienRJRIIGD", strNKdRuangan, mstrKdRuanganPasien))
        .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, dcRuangan.BoundText)
        .Parameters.Append .CreateParameter("IdDokterPerujuk", adChar, adParamInput, 10, mstrKdDokter)
        .Parameters.Append .CreateParameter("TglDirujuk", adDate, adParamInput, , Format(dtpTglDirujuk.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PasienRujukan"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Data", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Add_PasienRujukan")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With

    MsgBox "Penyimpanan Data Sukses !", vbInformation, "Informasi"
    dcRuangan.Enabled = False
    dcInstalasi.Enabled = False
    txtDokter.Enabled = False
    cmdSimpan.Enabled = False
    Exit Sub
hell:
    MsgBox "Maaf No.CM tidak ada", vbCritical, "Validasi"
End Sub

Private Sub cmdTutup_Click()
    If cmdSimpan.Enabled = True Then
        If MsgBox("Simpan data konsul", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
            Exit Sub
        End If
    End If
    Unload Me
End Sub

'/kodingan Lama
'Private Sub dcInstalasi_Change()
'    dcRuangan.Text = ""
'End Sub
'
'Private Sub dcInstalasi_GotFocus()
'    On Error GoTo errLoad
'    Dim tempKode As String
'
'    tempKode = dcInstalasi.BoundText
'
'    strSQL = "select distinct KdInstalasi,NamaInstalasi from V_RuanganTujuanRujukan "
'    Call msubDcSource(dcInstalasi, rs, strSQL)
'
'    dcInstalasi.BoundText = tempKode
'
'    Exit Sub
'errLoad:
'    Call msubPesanError
'End Sub
'
'Private Sub dcInstalasi_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 39 Then KeyAscii = 0
'    If KeyAscii = 13 Then
'        If dcInstalasi.MatchedWithList = True Then dcRuangan.SetFocus
'        strSQL = "select distinct KdInstalasi,NamaInstalasi from V_RuanganTujuanRujukan where kdInstalasi = '" & mstrKdInstalasiLogin & "' and (NamaInstalasi LIKE '%" & dcInstalasi.Text & "%')"
'        Call msubRecFO(rs, strSQL)
''        Call msubDcSource(dcInstalasi, rs, strSQL)
'        If rs.EOF = True Then
'            dcInstalasi.Text = ""
'            dcInstalasi.SetFocus
'            Exit Sub
'        End If
'        dcInstalasi.BoundText = rs(0).Value
'        dcInstalasi.Text = rs(1).Value
''        dcInstalasi.SetFocus
'    End If
'End Sub
'//kodingan Lama

Private Sub dcInstalasi_Change()
    dcRuangan.Text = ""
End Sub

Private Sub dcInstalasi_GotFocus()
    On Error GoTo errLoad
    Dim tempKode As String

    tempKode = dcInstalasi.BoundText
    Set rs = Nothing
    rs.Open "select distinct KdInstalasi,NamaInstalasi from V_RuanganTujuanRujukan where StatusEnabled='1' order by NamaInstalasi", dbConn, adOpenStatic, adLockReadOnly
    Set dcInstalasi.RowSource = rs
    dcInstalasi.ListField = rs(1).Name
    dcInstalasi.BoundColumn = rs(0).Name
    Set rs = Nothing
    dcInstalasi.BoundText = tempKode

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcInstalasi_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcInstalasi.Text)) = 0 Then dcRuangan.SetFocus: Exit Sub
        If dcInstalasi.MatchedWithList = True Then dcRuangan.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "select distinct KdInstalasi,NamaInstalasi from V_RuanganTujuanRujukan where NamaInstalasi LIKE '%" & dcInstalasi.Text & "%'")
        If dbRst.EOF = True Then dcInstalasi.Text = "": Exit Sub
        dcInstalasi.BoundText = dbRst(0).Value
        dcInstalasi.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub
'//kodingan Lama
'Private Sub dcRuangan_GotFocus()
'    On Error GoTo errLoad
'    Dim tempKode As String
'
'    If mstrFormPengirim <> "frmDaftarPasienRJRIIGD" Then
'        tempKode = dcRuangan.BoundText
'        strSQL = "select distinct KdRuangan,NamaRuangan from V_RuanganTujuanRujukan where KdInstalasi='" & dcInstalasi.BoundText & "' and KdRuangan <> '" & mstrKdRuangan & "'  order by NamaRuangan "
'        Call msubDcSource(dcRuangan, rs, strSQL)
'        dcRuangan.BoundText = tempKode
'    Else
'        tempKode = dcRuangan.BoundText
'        strSQL = "select distinct KdRuangan,NamaRuangan from V_RuanganTujuanRujukan where KdInstalasi='" & dcInstalasi.BoundText & "' order by NamaRuangan "
'        Call msubDcSource(dcRuangan, rs, strSQL)
'        dcRuangan.BoundText = tempKode
'    End If
'    Exit Sub
'errLoad:
'    Call msubPesanError
'End Sub
'
'Private Sub dcRuangan_KeyPress(KeyAscii As Integer)
'If KeyAscii = 39 Then KeyAscii = 0
'    If KeyAscii = 13 Then
'        If dcRuangan.MatchedWithList = True Then txtDokter.SetFocus
'        strSQL = "select distinct KdRuangan,NamaRuangan from V_RuanganTujuanRujukan where KdInstalasi='" & dcInstalasi.BoundText & "' and NamaRuangan = '" & mstrNamaRuangan & "' and (NamaRuangan LIKE '%" & dcRuangan.Text & "%') order by NamaRuangan"
'        Call msubRecFO(rs, strSQL)
''        Call msubDcSource(dcRuangan, rs, strSQL)
'        If rs.EOF = True Then
'            dcRuangan.Text = ""
'            dcRuangan.SetFocus
'            Exit Sub
'        End If
'        dcRuangan.BoundText = rs(0).Value
'        dcRuangan.Text = rs(1).Value
'    End If
'End Sub
'//kodingan Lama

Private Sub dcRuangan_GotFocus()
    On Error GoTo errLoad
    Dim tempKode As String

    tempKode = dcRuangan.BoundText
    Set rs = Nothing
    rs.Open "select distinct KdRuangan,NamaRuangan from V_RuanganTujuanRujukan where KdInstalasi='" & dcInstalasi.BoundText & "' and Expr1='1' order by NamaRuangan", dbConn, adOpenStatic, adLockReadOnly
    Set dcRuangan.RowSource = rs
    dcRuangan.ListField = rs(1).Name
    dcRuangan.BoundColumn = rs(0).Name
    Set rs = Nothing
    dcRuangan.BoundText = tempKode

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcRuangan_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
'        If Len(Trim(dcRuangan.Text)) <> 0 Then txtDokter.SetFocus: Exit Sub
        If dcRuangan.MatchedWithList = True Then txtDokter.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "select distinct KdRuangan,NamaRuangan from V_RuanganTujuanRujukan where KdInstalasi LIKE '%" & dcInstalasi.BoundText & "%' And NamaRuangan like'%" & dcRuangan.Text & "%'")
        If dbRst.EOF = True Then Exit Sub  'dcRuangan.Text = "":
        dcRuangan.BoundText = dbRst(0).Value
        dcRuangan.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dgDokter_DblClick()
    Call dgDokter_KeyPress(13)
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDokter.Text = dgDokter.Columns(1).Value
        mstrKdDokter = dgDokter.Columns(0).Value
        If mstrKdDokter = "" Then
            MsgBox "Pilih dulu Dokter yang akan menangani Pasien", vbCritical, "Validasi"
            txtDokter.Text = ""
            dgDokter.SetFocus
            Exit Sub
        End If
        fraDokter.Visible = False
    End If
    If KeyAscii = 27 Then
        fraDokter.Visible = False
    End If
End Sub

Private Sub dgDokter_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If dgDokter.Visible = False Then Exit Sub
        txtDokter.SetFocus
    End If
End Sub

Private Sub dtpTglDirujuk_Change()
    dtpTglDirujuk.MaxDate = Now
End Sub

Private Sub dtpTglDirujuk_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        dcInstalasi.SetFocus
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo hell

    Call centerForm(Me, MDIUtama)
    dtpTglDirujuk.Value = Now
    cmdInputPermintaan.Enabled = False
    If mstrFormPengirim = "frmDaftarPasienRJRIIGD" Then
        Call msubDcSource(dcInstalasi, rs, "select distinct KdInstalasi,NamaInstalasi from V_RuanganTujuanRujukan where kdInstalasi = '" & mstrKdInstalasiLogin & "'  ")
        dcInstalasi.BoundText = mstrKdInstalasiLogin
        Call msubDcSource(dcRuangan, rs, "select distinct KdRuangan,NamaRuangan from V_RuanganTujuanRujukan where KdInstalasi='" & dcInstalasi.BoundText & "' and NamaRuangan = '" & mstrNamaRuangan & "'  order by NamaRuangan")
        dcRuangan.Text = mstrNamaRuangan
        txtDokter.Text = IIf(IsNull(mstrNamaDokter), "", mstrNamaDokter)
    End If
    Call PlayFlashMovie(Me)
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mstrFormPengirim = "frmDaftarPasienRJRIIGD" Then
        Exit Sub
    Else
        frmTransaksiPasien.Enabled = True
        frmTransaksiPasien.subLoadRiwayatKonsul
    End If
End Sub

Private Sub txtDokter_Change()
    If mstrFormPengirim <> "frmDaftarPasienRJRIIGD" Then
        strFilterDokter = "WHERE NamaDokter like '%" & txtDokter.Text & "%'"
        fraDokter.Visible = True
        Call subLoadDokter
    Else
        Exit Sub
    End If
End Sub

Private Sub txtDokter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If dgDokter.Visible = False Then Exit Sub
        dgDokter.SetFocus
    End If
End Sub

Private Sub txtDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then
        strFilterDokter = "WHERE NamaDokter like '%" & txtDokter.Text & "%'"
        fraDokter.Visible = True
        Call subLoadDokter
        dgDokter.SetFocus
    End If
    If KeyAscii = 27 Then
        fraDokter.Visible = False
    End If
End Sub

Private Sub subLoadDokter()
    On Error GoTo hell
    strSQL = "SELECT KodeDokter AS [Kode Dokter],NamaDokter AS [Nama Dokter],JK,Jabatan FROM V_DaftarDokter " & strFilterDokter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgDokter.DataSource = rs
    With dgDokter
        .Columns(0).Width = 1300
        .Columns(1).Width = 3000
        .Columns(2).Width = 500
        .Columns(3).Width = 3000
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub
