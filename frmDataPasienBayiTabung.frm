VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDataPasienBayiTabung 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pasien Bayi Tabung"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDataPasienBayiTabung.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   11955
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   465
      Left            =   8400
      TabIndex        =   8
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   465
      Left            =   10200
      TabIndex        =   9
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   465
      Left            =   6600
      TabIndex        =   40
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Pemeriksaan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   29
      Top             =   2640
      Width           =   11895
      Begin VB.TextBox txtNoHasilLab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   8640
         MaxLength       =   10
         TabIndex        =   38
         Top             =   1320
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo dcSubInstalasi 
         Height          =   330
         Left            =   2400
         TabIndex        =   1
         Top             =   720
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
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
      Begin MSComCtl2.DTPicker dtpTglPeriksa 
         Height          =   330
         Left            =   2400
         TabIndex        =   0
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   122093571
         UpDown          =   -1  'True
         CurrentDate     =   38077
      End
      Begin MSDataListLib.DataCombo dcDokter 
         Height          =   330
         Left            =   8640
         TabIndex        =   5
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
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
      Begin MSDataListLib.DataCombo dcParamedis 
         Height          =   330
         Left            =   8640
         TabIndex        =   6
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
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
      Begin MSComCtl2.DTPicker dtpTglPelayananAwal 
         Height          =   330
         Left            =   2400
         TabIndex        =   4
         Top             =   1800
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   122093571
         UpDown          =   -1  'True
         CurrentDate     =   38077
      End
      Begin MSComCtl2.DTPicker dtpTglKehamilan 
         Height          =   330
         Left            =   8640
         TabIndex        =   7
         Top             =   960
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   122093571
         UpDown          =   -1  'True
         CurrentDate     =   38077
      End
      Begin MSDataListLib.DataCombo dcMetBayiTabung 
         Height          =   330
         Left            =   2400
         TabIndex        =   2
         Top             =   1080
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
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
      Begin MSDataListLib.DataCombo dcSiklusPengobatan 
         Height          =   330
         Left            =   2400
         TabIndex        =   3
         Top             =   1440
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No Hasil Laboratorium"
         Height          =   210
         Index           =   2
         Left            =   6720
         TabIndex        =   39
         Top             =   1365
         Width           =   1755
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Siklus Pengobatan"
         Height          =   210
         Left            =   240
         TabIndex        =   37
         Top             =   1500
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Metodologi Bayi Tabung"
         Height          =   210
         Left            =   240
         TabIndex        =   36
         Top             =   1140
         Width           =   1965
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Kehamilan"
         Height          =   210
         Left            =   6720
         TabIndex        =   35
         Top             =   1020
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Kasus Penyakit"
         Height          =   210
         Left            =   240
         TabIndex        =   34
         Top             =   780
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Periksa"
         Height          =   210
         Left            =   240
         TabIndex        =   33
         Top             =   420
         Width           =   870
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Dokter Pemeriksa"
         Height          =   210
         Left            =   6720
         TabIndex        =   32
         Top             =   300
         Width           =   1425
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Paramedis"
         Height          =   210
         Left            =   6720
         TabIndex        =   31
         Top             =   660
         Width           =   810
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Pelayanan Awal"
         Height          =   210
         Left            =   240
         TabIndex        =   30
         Top             =   1860
         Width           =   1575
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
      Height          =   1575
      Left            =   0
      TabIndex        =   11
      Top             =   1080
      Width           =   11895
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   19
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   17
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   8640
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtHr 
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
         Height          =   315
         Left            =   10200
         MaxLength       =   6
         TabIndex        =   15
         Top             =   720
         Width           =   375
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
         Height          =   315
         Left            =   9420
         MaxLength       =   6
         TabIndex        =   14
         Top             =   720
         Width           =   375
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
         Height          =   315
         Left            =   8640
         MaxLength       =   6
         TabIndex        =   13
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtSubInstalasi 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   8640
         TabIndex        =   12
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   6720
         TabIndex        =   25
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Umur"
         Height          =   210
         Left            =   6720
         TabIndex        =   24
         Top             =   720
         Width           =   435
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "hr"
         Height          =   210
         Left            =   10650
         TabIndex        =   23
         Top             =   750
         Width           =   165
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "bln"
         Height          =   210
         Left            =   9870
         TabIndex        =   22
         Top             =   750
         Width           =   240
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "thn"
         Height          =   210
         Left            =   9075
         TabIndex        =   21
         Top             =   750
         Width           =   285
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Kasus Penyakit"
         Height          =   210
         Left            =   6720
         TabIndex        =   20
         Top             =   1080
         Width           =   1200
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   10
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
      Left            =   10080
      Picture         =   "frmDataPasienBayiTabung.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDataPasienBayiTabung.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "frmDataPasienBayiTabung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub subKosong()
    On Error GoTo hell
    dtpTglPeriksa.Value = Now
    dtpTglPelayananAwal.Value = Now
    dtpTglKehamilan.Value = Now
    dcMetBayiTabung.Text = ""
    dcSubInstalasi.Text = ""
    dcSiklusPengobatan.Text = ""
    dcDokter.Text = "'"
    dcParamedis.Text = ""
    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub subLoadDcSource()
    On Error GoTo hell
    Call msubDcSource(dcSubInstalasi, rs, "Select KdSubInstalasi,NamaSubInstalasi From SubInstalasi Where StatusEnabled=1")
    Call msubDcSource(dcMetBayiTabung, rs, "Select KdMetodologiBayiTabung,MetodologiBayiTabung From MetodologiBayiTabung Where StatusEnabled=1")
    Call msubDcSource(dcSiklusPengobatan, rs, "Select KdSiklusPengobatanBT,SiklusPengobatanBT From SiklusPengobatanBayiTabung Where StatusEnabled=1")

    strSQL = "Select IdPegawai,NamaLengkap From DataPegawai"
    Call msubDcSource(dcDokter, rs, strSQL)
    Call msubDcSource(dcParamedis, rs, strSQL)

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdBatal_Click()
    Call subKosong
    Call subLoadDcSource
End Sub

Private Sub cmdSimpan_Click()
    If Periksa("datacombo", dcSubInstalasi, "Pilih Kasus Penyakit!!") = False Then Exit Sub
    If Periksa("datacombo", dcMetBayiTabung, "Pilih Metodology Bayi Tabung!") = False Then Exit Sub
    If Periksa("datacombo", dcSiklusPengobatan, "Pilih Siklus Pengobatan!!") = False Then Exit Sub
    If Periksa("datacombo", dcDokter, "Pilih Dokter Pemeriksa!!") = False Then Exit Sub
    If sp_AddBayiTabungPasien("A") = False Then Exit Sub
    MsgBox "Simpan Data Pasien Bayi Tabung Berhasil", vbInformation, "Informasi"
    Call cmdBatal_Click
End Sub

Private Function sp_AddBayiTabungPasien(strStatus As String) As Boolean
    On Error GoTo errLoad

    sp_AddBayiTabungPasien = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtnocm.Text)
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dtpTglPeriksa.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, dcSubInstalasi.BoundText)

        .Parameters.Append .CreateParameter("KdMetodologiBayiTabung", adTinyInt, adParamInput, , dcMetBayiTabung.BoundText)
        .Parameters.Append .CreateParameter("KdSiklusPengobatanBT", adTinyInt, adParamInput, , dcSiklusPengobatan.BoundText)

        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, dcDokter.BoundText)
        .Parameters.Append .CreateParameter("IdParamedis", adChar, adParamInput, 10, IIf(dcParamedis.Text = "", Null, dcParamedis.BoundText))
        .Parameters.Append .CreateParameter("TglKehamilan", adDate, adParamInput, , IIf(IsNull(dtpTglKehamilan.Value), Null, Format(dtpTglKehamilan.Value, "yyyy/MM/dd HH:mm")))
        .Parameters.Append .CreateParameter("NoHasilLaboratorium", adChar, adParamInput, 10, IIf(txtNoHasilLab.Text = "", Null, txtNoHasilLab.Text))
        .Parameters.Append .CreateParameter("TglPelayananAwal", adDate, adParamInput, , Format(dtpTglPelayananAwal.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, strStatus)

        .ActiveConnection = dbConn
        .CommandText = "AUD_BayiTabungPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_AddBayiTabungPasien = False
        End If
        Set dbcmd = Nothing
    End With

    Exit Function
errLoad:
    Call msubPesanError("AUD_BayiTabungPasien")
End Function

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcDokter_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 13 Then
        If dcDokter.MatchedWithList = True Then dcParamedis.SetFocus
        strSQL = "Select IdPegawai,NamaLengkap From V_DataPegawai WHERE (NamaLengkap LIKE '%" & dcDokter.Text & "%') and KdJenisPegawai='001'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcDokter.BoundText = rs(0).Value
        dcDokter.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcMetBayiTabung_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 13 Then
        If dcMetBayiTabung.MatchedWithList = True Then dcSiklusPengobatan.SetFocus
        strSQL = "Select KdMetodologiBayiTabung,MetodologiBayiTabung From MetodologiBayiTabung WHERE (MetodologiBayiTabung LIKE '%" & dcMetBayiTabung.Text & "%') And StatusEnabled=1"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcMetBayiTabung.BoundText = rs(0).Value
        dcMetBayiTabung.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcParamedis_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 13 Then
        If dcParamedis.MatchedWithList = True Then dtpTglKehamilan.SetFocus
        strSQL = "Select IdPegawai,NamaLengkap From V_DataPegawai WHERE (NamaLengkap LIKE '%" & dcParamedis.Text & "%') and KdJenisPegawai<>'001'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcParamedis.BoundText = rs(0).Value
        dcParamedis.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcSiklusPengobatan_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 13 Then
        If dcSiklusPengobatan.MatchedWithList = True Then dcDokter.SetFocus
        strSQL = "Select KdSiklusPengobatanBT,SiklusPengobatanBT From SiklusPengobatanBayiTabung WHERE (SiklusPengobatanBT LIKE '%" & dcSiklusPengobatan.Text & "%') And StatusEnabled=1"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcSiklusPengobatan.BoundText = rs(0).Value
        dcSiklusPengobatan.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcSubInstalasi_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 13 Then
        If dcSubInstalasi.MatchedWithList = True Then dcMetBayiTabung.SetFocus
        strSQL = "Select KdSubInstalasi,NamaSubInstalasi From SubInstalasi WHERE (NamaSubInstalasi LIKE '%" & dcSubInstalasi.Text & "%') And StatusEnabled=1"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcSubInstalasi.BoundText = rs(0).Value
        dcSubInstalasi.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dtpTglKehamilan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub dtpTglPelayananAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcDokter.SetFocus
End Sub

Private Sub dtpTglPeriksa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcSubInstalasi.SetFocus
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call cmdBatal_Click
    dcDokter.BoundText = mstrKdDokter
    dcSubInstalasi.BoundText = mstrKdSubInstalasi
End Sub
