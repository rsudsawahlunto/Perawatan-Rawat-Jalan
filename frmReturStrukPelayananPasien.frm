VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmReturStrukPelayananPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Retur Struk Pembayaran Pasien"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10245
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReturStrukPelayananPasien.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   10245
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   975
      Left            =   0
      TabIndex        =   36
      Top             =   -120
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox txtCaraBayar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   39
      Top             =   4440
      Width           =   10215
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   8400
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdBayar 
         Caption         =   "&Bayar"
         Height          =   495
         Left            =   6720
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraTransaksi 
      Caption         =   "Transaksi Pasien"
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
      TabIndex        =   15
      Top             =   2640
      Width           =   10215
      Begin VB.TextBox txtTotalTagihan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7320
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtBiayaAdministrasi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtPembebasan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtTanggunganRS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7320
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtTanggunganPenjamin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtTotalBiaya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Total Tagihan"
         Height          =   210
         Index           =   2
         Left            =   7320
         TabIndex        =   35
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Biaya Administrasi"
         Height          =   210
         Index           =   1
         Left            =   1680
         TabIndex        =   33
         Top             =   960
         Width           =   1410
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Pembebasan"
         Height          =   210
         Index           =   0
         Left            =   4560
         TabIndex        =   24
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Tanggungan Rumah Sakit"
         Height          =   210
         Left            =   7320
         TabIndex        =   22
         Top             =   240
         Width           =   2100
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Tanggungan Penjamin"
         Height          =   210
         Left            =   4440
         TabIndex        =   20
         Top             =   240
         Width           =   1830
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total Biaya"
         Height          =   210
         Left            =   1680
         TabIndex        =   18
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.Frame fraDataPasien 
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
      Height          =   1695
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   10215
      Begin VB.TextBox txtRuanganKasir 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtTglBKM 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtUmur 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtNoBKM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         MaxLength       =   10
         TabIndex        =   0
         ToolTipText     =   "Tekan tombol Enter untuk mencari struk pasien yang akan diretur"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtNoStruk 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   3
         ToolTipText     =   "Tekan tombol Enter untuk mencari struk pasien yang akan diretur"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtJenisPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtPenjamin 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Ruangan Kasir"
         Height          =   210
         Index           =   2
         Left            =   7320
         TabIndex        =   31
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. BKM"
         Height          =   210
         Index           =   1
         Left            =   5040
         TabIndex        =   29
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Umur"
         Height          =   210
         Index           =   1
         Left            =   8400
         TabIndex        =   27
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "No. BKM"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "No. Struk"
         Height          =   210
         Index           =   0
         Left            =   1800
         TabIndex        =   16
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   3360
         TabIndex        =   14
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Index           =   0
         Left            =   4320
         TabIndex        =   13
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Index           =   0
         Left            =   7200
         TabIndex        =   12
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pasien"
         Height          =   210
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Penjamin"
         Height          =   210
         Left            =   1800
         TabIndex        =   10
         Top             =   960
         Width           =   735
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   40
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
      Picture         =   "frmReturStrukPelayananPasien.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmReturStrukPelayananPasien.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmReturStrukPelayananPasien.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmReturStrukPelayananPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim subTglStruk As Date

Private Sub cmdBayar_Click()
On Error GoTo errLoad

    If Periksa("text", txtNoBKM, "No BKM kosong") = False Then Exit Sub
'    strSQL = "SELECT NoBKM FROM V_DaftarPasienYgSudahBayar WHERE NoBKM = '" & txtNoBKM.Text & "'"
'    Call msubRecFO(rs, strSQL)
'    If rs.EOF = False Then
'        MsgBox "No. BKM " & txtNoBKM.Text & " belum bayar", vbInformation, "Informasi"
'        Exit Sub
'    End If
    
    strSQL = " SELECT NoBKM FROM V_JudulReturStrukPelayananPasien WHERE NoBKM = '" & txtNoBKM.Text & "'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        MsgBox "No. BKM " & txtNoBKM.Text & " sudah diretur", vbInformation, "Informasi"
        Exit Sub
    End If
    
    With frmBayarRetur
        .Show
        Me.Enabled = False
        
        .txtNoBKM.Text = txtNoBKM.Text
        .txtNoStruk.Text = txtNoStruk.Text
        
        .txtTotalBiaya.Text = txtTotalBiaya.Text
        .txtTanggunganPenjamin.Text = txtTanggunganPenjamin.Text
        .txtTanggunganRS.Text = txtTanggunganRS.Text
        .chkBiayaAdministrasiAwal.Value = vbUnchecked
        .txtBiayaAdministrasiAwal.Text = txtBiayaAdministrasi.Text
        .txtPembebasanAwal.Text = txtPembebasan.Text
        .txtTotalHarusRetur.Text = txtTotalTagihan.Text
        
        .lblTotalTagihan.Caption = txtTotalTagihan.Text
        .dtpTglRetur.Value = txtTglBKM.Text
        .dcCaraBayar.Text = txtCaraBayar.Text
        
        .txtBiayaAdministrasi.Text = 0
        .txtJmlUang.Text = txtTotalTagihan.Text
        .txtKeterangan.Text = ""
        
        .txtNamaFormPengirim.Text = Me.Name
    End With

Exit Sub
errLoad:
    Call msubPesanError
End Sub

'Private Function sp_DetailReturPelayananPasien() As Boolean
'On Error GoTo errLoad
'
'    sp_DetailReturPelayananPasien = True
'    Set dbcmd = New ADODB.Command
'    With dbcmd
'        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
'        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
'        .Parameters.Append .CreateParameter("NoCM", adChar, adParamInput, 6, txtNoCM.Text)
'        .Parameters.Append .CreateParameter("NoRetur", adChar, adParamInput, 10, txtNoRetur.Text)
'        .Parameters.Append .CreateParameter("NoBKM", adChar, adParamInput, 10, txtNoBKM.Text)
'        .Parameters.Append .CreateParameter("NoStruk", adChar, adParamInput, 10, txtNoStruk.Text)
'        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 2, IIf(optJnsRetur(0).Value = True, "BU", "TB"))
'
'        .ActiveConnection = dbConn
'        .CommandText = "dbo.Add_DetailReturStrukPelayananPasien"
'        .CommandType = adCmdStoredProc
'        .Execute
'
'        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
'            MsgBox "Ada kesalahan dalam detail retur struk", vbCritical, "Informasi"
'            sp_DetailReturPelayananPasien = False
'        End If
'        Call deleteADOCommandParameters(dbcmd)
'        Set dbcmd = Nothing
'    End With
'
'Exit Function
'errLoad:
'    sp_DetailReturPelayananPasien = False
'    Call msubPesanError
'End Function
'Private Sub cmdSimpan_Click()
'On Error GoTo errLoad
'    If Periksa("text", txtNoStruk, "No Struk kosong") = False Then Exit Sub
'
'    If sp_Retur() = False Then Exit Sub
'    If sp_DetailReturPelayananPasien() = False Then Exit Sub
'    Call subClear
'    txtNoStruk.SetFocus
'
'Exit Sub
'errLoad:
'    Call msubPesanError
'End Sub


Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
End Sub

Private Sub subClear()
    txtNoStruk.Text = ""
    txtNoCM.Text = ""
    txtNoPendaftaran.Text = ""
    txtNamaPasien.Text = ""
    txtSex.Text = ""
    txtUmur.Text = ""
    
    txtJenisPasien.Text = ""
    txtPenjamin.Text = ""
    txtTglBKM.Text = ""
    txtRuanganKasir.Text = ""
    
    txtTotalBiaya.Text = 0
    txtTanggunganPenjamin.Text = 0
    txtTanggunganRS.Text = 0
    txtBiayaAdministrasi.Text = 0
    txtPembebasan.Text = 0
    txtTotalTagihan.Text = 0
End Sub

Public Sub txtNoBKM_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    If KeyAscii = 13 Then
        Call subClear
        strSQL = "SELECT NoStruk, NoPendaftaran, NoCM, NamaPasien, JK, Umur, JenisPasien, NamaPenjamin, TglBKM, RuanganKasir, CaraBayar, TotalBiaya, JmlHutangPenjamin, JmlTanggunganRS, Administrasi, JmlPembebasan, JmlBayar " & _
            " FROM V_JudulStrukPembayaranPasien " & _
            " WHERE NoBKM = '" & txtNoBKM.Text & "'"
        Call msubRecFO(rs, strSQL)
        
        If rs.EOF = True Then
            MsgBox "No. BKM belum ada", vbExclamation, "Validasi": Call subClear: Exit Sub
        Else
            Call msubRecFO(dbRst, "SELECT NoBKM FROM DetailReturStrukPelayananPasien WHERE (NoBKM = '" & txtNoBKM.Text & "')")
            If dbRst.EOF = False Then MsgBox "No. BKM " & txtNoBKM.Text & " sudah pernah diretur", vbExclamation, "Validasi": Call subClear: Exit Sub
        End If
        
        txtNoPendaftaran.Text = IIf(IsNull(rs("NoPendaftaran")), "", rs("NoPendaftaran"))
        txtCaraBayar.Text = rs("CaraBayar")
        
        txtNoStruk.Text = rs("NoStruk").Value
        txtNoCM.Text = IIf(IsNull(rs("NoCM")), "", rs("NoCM"))
        txtNamaPasien.Text = IIf(IsNull(rs("NamaPasien")), "", rs("NamaPasien"))
        txtSex.Text = IIf(IsNull(rs("JK")), "", IIf(rs("JK") = "L", "Laki-Laki", "Perempuan"))
        txtUmur.Text = IIf(IsNull(rs("Umur")), "", rs("Umur"))
        txtJenisPasien.Text = rs("JenisPasien").Value
        txtPenjamin.Text = rs("NamaPenjamin").Value
        txtTglBKM.Text = rs("TglBKM").Value
        txtRuanganKasir.Text = rs("RuanganKasir").Value
        
        txtTotalBiaya.Text = Format(rs("TotalBiaya").Value, "#,###.00")
        txtTanggunganPenjamin.Text = Format(rs("JmlHutangPenjamin").Value, "#,###.00")
        txtTanggunganRS.Text = Format(rs("JmlTanggunganRS").Value, "#,###.00")
        txtBiayaAdministrasi.Text = Format(rs("Administrasi").Value, "#,###.00")
        txtPembebasan.Text = Format(rs("JmlPembebasan").Value, "#,###.00")
        txtTotalTagihan.Text = Format(rs("JmlBayar").Value, "#,###.00")
        
    End If

Exit Sub
errLoad:
    Call msubPesanError
End Sub
