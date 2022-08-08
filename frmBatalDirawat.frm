VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBatalDirawat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Batal Pemeriksaan"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBatalDirawat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmBatalDirawat.frx":0CCA
   ScaleHeight     =   4110
   ScaleWidth      =   9315
   Begin VB.TextBox txtTglMasuk 
      Height          =   495
      Left            =   4200
      TabIndex        =   15
      Text            =   "u/ simpan tgl masuk"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame frame6 
      Caption         =   "Invisible"
      Height          =   975
      Left            =   2640
      TabIndex        =   30
      Top             =   2040
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox txtKdSubInstalasi 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   13
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtKdRuangan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   120
         MaxLength       =   50
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpTglMasuk 
         Height          =   330
         Left            =   2520
         TabIndex        =   14
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy  hh:mm"
         Format          =   107610115
         CurrentDate     =   38081
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "TglMasuk"
         Height          =   210
         Left            =   2520
         TabIndex        =   35
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "KdSubInstalasi"
         Height          =   210
         Left            =   1200
         TabIndex        =   32
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "KdRuangan"
         Height          =   210
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame Frame5 
      Height          =   855
      Left            =   0
      TabIndex        =   29
      Top             =   3240
      Width           =   9255
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   7200
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   5040
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame fraBatalDirawat 
      Caption         =   "Ruang Pemeriksaan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   25
      Top             =   2160
      Width           =   9255
      Begin VB.TextBox txtRuanganLama 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   8
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5040
         MaxLength       =   100
         TabIndex        =   9
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox txtDokterLama 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   27
         Top             =   1080
         Visible         =   0   'False
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker dtpTglBatalRawat 
         Height          =   330
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy  hh:mm"
         Format          =   108199939
         UpDown          =   -1  'True
         CurrentDate     =   38081
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Batal Dirawat"
         Height          =   210
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   1770
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan"
         Height          =   210
         Left            =   5040
         TabIndex        =   33
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dokter Pemeriksa"
         Height          =   210
         Left            =   600
         TabIndex        =   28
         Top             =   1080
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ruang Pemeriksaan"
         Height          =   210
         Left            =   2400
         TabIndex        =   26
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   1215
      Left            =   0
      TabIndex        =   16
      Top             =   960
      Width           =   9255
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   3240
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   240
         MaxLength       =   10
         TabIndex        =   0
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5400
         MaxLength       =   9
         TabIndex        =   3
         Top             =   600
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
         Left            =   6600
         TabIndex        =   17
         Top             =   320
         Width           =   2535
         Begin VB.TextBox txtThn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            MaxLength       =   6
            TabIndex        =   4
            Top             =   240
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
            Top             =   240
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
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            Height          =   210
            Left            =   600
            TabIndex        =   20
            Top             =   292
            Width           =   285
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            Height          =   210
            Left            =   1440
            TabIndex        =   19
            Top             =   292
            Width           =   240
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            Height          =   210
            Left            =   2280
            TabIndex        =   18
            Top             =   292
            Width           =   165
         End
      End
      Begin VB.Label lblNamaPasien 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   3240
         TabIndex        =   24
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   1800
         TabIndex        =   23
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblJnsKlm 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   5400
         TabIndex        =   21
         Top             =   360
         Width           =   1065
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   36
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
      Left            =   7440
      Picture         =   "frmBatalDirawat.frx":1994
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmBatalDirawat.frx":271C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmBatalDirawat.frx":50DD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmBatalDirawat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public dTglBatal As Date

Private Sub cmdSimpan_Click()
    On Error GoTo errSimpan

    strSQL = "SELECT NoPendaftaran FROM PasienBatalDirawat WHERE (NoPendaftaran = '" & txtnopendaftaran.Text & "')"
    Call msubRecFO(rs, strSQL)
    If Not rs.EOF Then
        MsgBox "No pendaftaran " & txtnopendaftaran.Text & " tersebut pernah batal dirawat", vbExclamation, "Validasi"
        Exit Sub
    End If
    
    If mblnFormDaftarAntrian = True Then
        If bolStatusDelPelayanan = True Then
            Call sp_DelBiayaPelayanan(dbcmd)
            bolStatusDelPelayanan = False
        End If
    End If
    
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtnocm.Text)

        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, txtKdSubInstalasi.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(dtpTglMasuk.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglBatal", adDate, adParamInput, , Format(dtpTglBatalRawat.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 100, IIf(Len(Trim(txtKeterangan.Text)) = 0, Null, Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, noidpegawai)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PasienBatalDiPeriksa"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Pendaftaran Pasien ke Instalasi Rawat Jalan", vbCritical, "Validasi"

        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Call sp_UpdateStatusPasienMRS(dbcmd)

    MsgBox "Penyimpanan data pasien batal dirawat sukses", vbInformation, "Informasi"
    fraBatalDirawat.Enabled = False
    cmdSimpan.Enabled = False

    Call Add_HistoryLoginActivity("Add_PasienBatalDiPeriksa+Update_RegistrasiPasienMRS")
    Exit Sub
errSimpan:
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    MsgBox "Penyimpanan data pasien batal dirawat gagal", vbCritical, "Validasi"
End Sub

'Store procedure untuk Untuk Hapus Tindakan pasien yang Belum Memiliki NoStruk
Private Sub sp_DelBiayaPelayanan(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, KdPelayananRSBatalPeriksa)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(dtpTglMasuk.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Delete_BiayaPelayananNew"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penghapusan Biaya Pelayanan Pasien", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Delete_BiayaPelayananNew")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

'Store procedure untuk mengupdate status oasien masuk rs
Private Sub sp_UpdateStatusPasienMRS(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(dtpTglMasuk.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("StatusPeriksa", adChar, adParamInput, 1, "B") '@StatusPeriksa char(1) --Y=Sudah, T=Belum, S=Sedang, B=Batal

        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_RegistrasiPasienMRS"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam update status pasien masuk rumah sakit", vbCritical, "Validasi"

        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dtpTglBatalRawat_Change()
    If dtpTglBatalRawat.Value > Now Then
        MsgBox "Tanggal batal dirawat harus lebih kecil atau sama dengan tanggal sekarang", vbCritical, "Validasi"
        dtpTglBatalRawat.SetFocus
        Exit Sub
    ElseIf dtpTglBatalRawat.Value < dtpTglMasuk.Value Then
        MsgBox "Tanggal batal dirawat harus lebih besar dari tanggal masuk", vbCritical, "Validasi"
        dtpTglBatalRawat.SetFocus
        Exit Sub
    End If
End Sub

Private Sub dtpTglBatalRawat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtKeterangan.SetFocus
End Sub

Private Sub dtpTglBatalRawat_LostFocus()
    Call dtpTglBatalRawat_Change
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpTglBatalRawat.Value = Now
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnFormDaftarAntrian = True Then
        frmDaftarAntrianPasien.Enabled = True
        Call frmDaftarAntrianPasien.cmdCari_Click
    End If
    If mblnFormDaftarPasienRJ = True Then
        frmDaftarPasienRJ.Enabled = True
        Call frmDaftarPasienRJ.cmdCari_Click
    End If
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtKeterangan_LostFocus()
    txtKeterangan.Text = StrConv(txtKeterangan.Text, vbProperCase)
End Sub

Private Sub UpdateKeteranganBatal()
    On Error GoTo hell
    strSQL = "UPDATE PasienBatalDirawat set Keterangan = '" & txtKeterangan.Text & "' WHERE NoPendaftaran = '" & txtnopendaftaran.Text & "' "
    dbConn.Execute strSQL
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub txtRuanganLama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub
