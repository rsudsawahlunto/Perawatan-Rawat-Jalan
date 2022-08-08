VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmSuratKeterangan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Surat Keterangan"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSuratKeterangan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   6630
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   0
      TabIndex        =   36
      Top             =   4200
      Width           =   6615
      Begin VB.TextBox txtPengujian 
         Height          =   315
         Left            =   3240
         TabIndex        =   6
         Top             =   240
         Width           =   2895
      End
      Begin VB.CheckBox chkHepatitis 
         Caption         =   "Cek Hepatitis"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Pengujian ke :"
         Height          =   255
         Left            =   2040
         TabIndex        =   37
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox txtJenisKelamin 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   1560
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pengujian"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      TabIndex        =   27
      Top             =   4800
      Width           =   6615
      Begin VB.TextBox txtNIP 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4320
         MaxLength       =   200
         TabIndex        =   14
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox txtTekanan2 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2160
         MaxLength       =   200
         TabIndex        =   12
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtKeperluan 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1560
         MaxLength       =   200
         TabIndex        =   15
         Top             =   2400
         Width           =   4935
      End
      Begin VB.TextBox txtTekanan 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1560
         MaxLength       =   200
         TabIndex        =   11
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtBerat 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1560
         MaxLength       =   200
         TabIndex        =   9
         Top             =   650
         Width           =   975
      End
      Begin VB.TextBox txtTinggi 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1560
         MaxLength       =   200
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin MSDataListLib.DataCombo dcDokterPenguji 
         Height          =   315
         Left            =   1560
         TabIndex        =   13
         Top             =   1920
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcGolDarah 
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   1030
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   375
         Left            =   4320
         TabIndex        =   7
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
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
         OLEDropMode     =   1
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   112394243
         UpDown          =   -1  'True
         CurrentDate     =   38209
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "NIP"
         Height          =   210
         Left            =   3960
         TabIndex        =   38
         Top             =   1920
         Width           =   285
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal"
         Height          =   210
         Left            =   3600
         TabIndex        =   35
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label13 
         Caption         =   "/"
         Height          =   255
         Left            =   2040
         TabIndex        =   34
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Digunakan untuk"
         Height          =   210
         Left            =   120
         TabIndex        =   33
         Top             =   2400
         Width           =   1380
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Tekanan Darah"
         Height          =   210
         Left            =   120
         TabIndex        =   32
         Top             =   1440
         Width           =   1230
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Golongan Darah"
         Height          =   210
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   1290
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Berat Badan"
         Height          =   210
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   990
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tinggi Badan"
         Height          =   210
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Dokter Penguji"
         Height          =   210
         Left            =   120
         TabIndex        =   28
         Top             =   1920
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdPrint2 
      Caption         =   "Cetak Keterangan Syarat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   17
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton cmdOut 
      Caption         =   "Tutup"
      Height          =   495
      Left            =   5400
      TabIndex        =   18
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Cetak Keterangan Sehat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      TabIndex        =   19
      Top             =   1080
      Width           =   6615
      Begin VB.TextBox txtKeterangan 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1560
         MaxLength       =   200
         TabIndex        =   3
         Top             =   2160
         Width           =   4935
      End
      Begin VB.TextBox txtTempat 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1560
         TabIndex        =   2
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtNama 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   4695
      End
      Begin MSDataListLib.DataCombo dcPekerjaan 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   2640
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox txtTglLahir 
         Height          =   375
         Left            =   1560
         TabIndex        =   39
         Top             =   1680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         HideSelection   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Alamat"
         Height          =   210
         Left            =   240
         TabIndex        =   26
         Top             =   2180
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Pekerjaan"
         Height          =   210
         Left            =   240
         TabIndex        =   25
         Top             =   2640
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Lahir"
         Height          =   210
         Left            =   240
         TabIndex        =   24
         Top             =   1680
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tempat Lahir"
         Height          =   210
         Left            =   240
         TabIndex        =   23
         Top             =   1200
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   1020
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   21
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
      Left            =   4800
      Picture         =   "frmSuratKeterangan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmSuratKeterangan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4815
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmSuratKeterangan.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmSuratKeterangan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkHepatitis_Click()
    If chkHepatitis.Value = 1 Then
        Frame2.Enabled = False
        txtTinggi.Enabled = False
        txtBerat.Enabled = False
        dcGolDarah.Enabled = False
        txtTekanan.Enabled = False
        txtTekanan2.Enabled = False
        dcDokterPenguji.Enabled = False
        txtKeperluan.Enabled = False
        dtpAwal.Enabled = False
        cmdPrint2.Enabled = False
        txtPengujian.Enabled = True
        txtPengujian.SetFocus
    End If

    If chkHepatitis.Value = 0 Then
        Frame2.Enabled = True
        txtTinggi.Enabled = True
        txtBerat.Enabled = True
        dcGolDarah.Enabled = True
        txtTekanan.Enabled = True
        txtTekanan2.Enabled = True
        dcDokterPenguji.Enabled = True
        txtKeperluan.Enabled = True
        dtpAwal.Enabled = True
        cmdPrint2.Enabled = True
        txtPengujian.Enabled = False
        dtpAwal.SetFocus
    End If

End Sub

Private Sub chkHepatitis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkHepatitis.Value = vbUnchecked Then
            txtPengujian.Enabled = False
            dtpAwal.SetFocus
        End If
        If KeyAscii = 13 Then
            If chkHepatitis.Value = vbChecked Then txtPengujian.Enabled = True
        End If
    End If
End Sub

Private Sub cmdOut_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    If chkHepatitis.Value = 0 Then

        If txtNama.Text = "" Then
            MsgBox "Nama Pasien harap diisi", vbInformation, "Informasi"
            txtNama.SetFocus
            Exit Sub
        End If

        If txtTempat.Text = "" Then
            MsgBox "Tempat Lahir Pasien harap diisi", vbInformation, "Informasi"
            txtTempat.SetFocus
            Exit Sub
        End If

        If dcPekerjaan.Text = "" Then
            MsgBox "Pekerjaan harap diiisi", vbInformation, "Informasi"
            dcPekerjaan.SetFocus
            Exit Sub
        End If

        If txtKeperluan.Text = "" Then
            MsgBox "Keperluan pembuatan surat harus diisi", vbInformation, "Informasi"
            txtKeperluan.SetFocus
            Exit Sub
        End If

        frmCetakSuratKeterangan.Show
    End If

    If chkHepatitis.Value = 1 Then
        frmCetakSuratKeteranganHepatitis.Show
    End If

End Sub

Private Sub cmdPrint2_Click()
    If txtNama.Text = "" Then
        MsgBox "Nama Pasien harap diisi", vbInformation, "Informasi"
        txtNama.SetFocus
        Exit Sub
    End If
    If txtTempat.Text = "" Then
        MsgBox "Tempat Lahir Pasien harap diisi", vbInformation, "Informasi"
        txtTempat.SetFocus
        Exit Sub
    End If

    If dcPekerjaan.Text = "" Then
        MsgBox "Pekerjaan harap diiisi", vbInformation, "Informasi"
        dcPekerjaan.SetFocus
        Exit Sub
    End If

    If txtKeperluan.Text = "" Then
        MsgBox "Keperluan pembuatan surat harus diisi", vbInformation, "Informasi"
        txtKeperluan.SetFocus
        Exit Sub
    End If

    frmCetakSuratKeterangan2.Show
End Sub

Private Sub dcDokterPenguji_Click(Area As Integer)
        strSQL = "Select isnull(NIP,'-') as NIP from V_M_DataPegawaiNew where [Nama Lengkap] like '%" & dcDokterPenguji.Text & "%'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            txtNIP.Text = "-"
          Else
            txtNIP.Text = rs("NIP")
        End If

End Sub

Private Sub dcDokterPenguji_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcDokterPenguji.MatchedWithList = True Then txtKeperluan.SetFocus
        strSQL = "Select idpegawai, namalengkap from DataPegawai where KdJenisPegawai = '001' and (namalengkap LIKE '%" & dcDokterPenguji.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcDokterPenguji.Text = ""
            txtKeperluan.SetFocus
            Exit Sub
        End If
        dcDokterPenguji.BoundText = rs(0).Value
        dcDokterPenguji.Text = rs(1).Value
        
        strSQL = "Select isnull(NIP,'-') as NIP from V_M_DataPegawaiNew where [Nama Lengkap] like '%" & dcDokterPenguji.Text & "%'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            txtNIP.Text = "-"
          Else
            txtNIP.Text = rs("NIP")
        End If

    End If

End Sub

Private Sub dcDokterPenguji_LostFocus()
    If dcDokterPenguji.Text = "" Then Exit Sub
    If dcDokterPenguji.MatchedWithList = False Then dcDokterPenguji.Text = "": dcDokterPenguji.SetFocus: Exit Sub

        strSQL = "Select isnull(NIP,'-') as NIP from V_M_DataPegawaiNew where [Nama Lengkap] like '%" & dcDokterPenguji.Text & "%'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            txtNIP.Text = "-"
          Else
            txtNIP.Text = rs("NIP")
        End If
    
End Sub

Private Sub dcGolDarah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcGolDarah.MatchedWithList = True Then txtTekanan.SetFocus
        strSQL = "Select kdgolongandarah, golongandarah from GolonganDarah where (golongandarah LIKE '%" & dcGolDarah.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcGolDarah.Text = ""
            txtTekanan.SetFocus
            Exit Sub
        End If
        dcGolDarah.BoundText = rs(0).Value
        dcGolDarah.Text = rs(1).Value
    End If
End Sub

Private Sub dcGolDarah_LostFocus()
    If dcGolDarah.MatchedWithList = True Then txtTekanan.SetFocus
    strSQL = "Select kdgolongandarah, golongandarah from GolonganDarah where (golongandarah LIKE '%" & dcGolDarah.Text & "%')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
        dcGolDarah.Text = ""
        txtTekanan.SetFocus
        Exit Sub
    End If
    dcGolDarah.BoundText = rs(0).Value
    dcGolDarah.Text = rs(1).Value

End Sub

Private Sub dcPekerjaan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcPekerjaan.MatchedWithList = True Then chkHepatitis.SetFocus
        strSQL = "Select kdpekerjaan, pekerjaan from Pekerjaan where (pekerjaan LIKE '%" & dcPekerjaan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcPekerjaan.Text = ""
            dcPekerjaan.SetFocus
            Exit Sub
        End If
        dcPekerjaan.BoundText = rs(0).Value
        dcPekerjaan.Text = rs(1).Value
    End If
End Sub

Private Sub dcPekerjaan_LostFocus()
    If dcPekerjaan.Text = "" Then Exit Sub
    If dcPekerjaan.MatchedWithList = False Then dcPekerjaan.Text = "": dcPekerjaan.SetFocus

End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtTinggi.SetFocus

End Sub

Private Sub dtpAwal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTinggi.SetFocus
End Sub

Private Sub Form_Load()
    On Error GoTo hell
    chkHepatitis.Value = 0
    txtPengujian.Enabled = False
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call SetComboPekerjaan
    Call SetComboNamaDokter
    Call SetComboGolonganDarah

    dtpAwal.Value = Format(Now, "dd MMM yyyy 00:00:00")
    
    strSQL = "select Title , NamaLengkap,JenisKelamin,TempatLahir,TglLahir,Alamat  from Pasien WHERE NoCM = '" & frmDaftarPasienRJ.dgDaftarPasienRJ.Columns("NoCM").Value & "'"
    Call msubRecFO(rs, strSQL)

    txtNama.Text = rs("Title").Value & " " & rs("NamaLengkap").Value
    If rs("JenisKelamin").Value = "L" Then
        txtJenisKelamin.Text = "Laki - Laki"
    Else
        txtJenisKelamin.Text = "Perempuan"
    End If
    
    txtTempat.Text = rs("TempatLahir") & ""
    txtTglLahir.Text = Format(rs("TglLahir").Value, "dd/mm/yyyy")
'    txtTglLahir1.Text = rs("TglLahir").Value
    txtKeterangan.Text = rs("Alamat") & ""
    
'    dcDokterPenguji.Text = frmDaftarPasienRJ.dgDaftarPasienRJ.Columns("Dokter Pemeriksa").Value

    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub SetComboPekerjaan()
    Set rs = Nothing
    rs.Open "Select * from Pekerjaan ", dbConn, , adLockOptimistic
    Set dcPekerjaan.RowSource = rs
    dcPekerjaan.ListField = rs.Fields(1).Name
    dcPekerjaan.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
End Sub

Sub SetComboNamaDokter()
    Set rs = Nothing
    rs.Open "Select * from DataPegawai where KdJenisPegawai = '001' order by NamaLengkap", dbConn, , adLockOptimistic
    Set dcDokterPenguji.RowSource = rs
    dcDokterPenguji.ListField = rs.Fields(3).Name
    dcDokterPenguji.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
End Sub

Sub SetComboGolonganDarah()
    Set rs = Nothing
    rs.Open "Select * from GolonganDarah ", dbConn, , adLockOptimistic
    Set dcGolDarah.RowSource = rs
    dcGolDarah.ListField = rs.Fields(1).Name
    dcGolDarah.BoundColumn = rs.Fields(0).Name
End Sub

Private Sub txtJenisKelamin_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtTempat.SetFocus
End Sub

Private Sub txtKeperluan_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then cmdPrint.SetFocus
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcPekerjaan.SetFocus
End Sub

Private Sub txtNama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJenisKelamin.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtNIP_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtKeperluan.SetFocus
End Sub

Private Sub txtPengujian_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then cmdPrint.SetFocus
End Sub

Private Sub txtTekanan2_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then dcDokterPenguji.SetFocus
End Sub

Private Sub txtTempat_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtKeterangan.SetFocus 'txtTglLahir.SetFocus
End Sub

Private Sub txtTglLahir_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
End Sub

Private Sub txtTinggi_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtBerat.SetFocus
End Sub

Private Sub txtBerat_kEYpRESS(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then dcGolDarah.SetFocus
End Sub

Private Sub txtTekanan_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtTekanan2.SetFocus
End Sub
