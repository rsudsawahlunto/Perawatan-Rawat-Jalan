VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReturStruk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retur Struk"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReturStruk.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8565
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   5160
      TabIndex        =   17
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   6840
      TabIndex        =   18
      Top             =   5400
      Width           =   1575
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
      Height          =   1215
      Left            =   0
      TabIndex        =   31
      Top             =   4080
      Width           =   8535
      Begin VB.TextBox txtPembebasan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   6840
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtTggRS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5160
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtHtgPjmn 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtJmlByr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtTotalBiaya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Pembebasan"
         Height          =   210
         Left            =   6840
         TabIndex        =   37
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Tanggungan RS"
         Height          =   210
         Left            =   5160
         TabIndex        =   36
         Top             =   360
         Width           =   1305
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Hutang Penjamin"
         Height          =   210
         Left            =   3480
         TabIndex        =   35
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Jumlah Bayar"
         Height          =   210
         Left            =   1800
         TabIndex        =   34
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total Biaya"
         Height          =   210
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.Frame fraRetur 
      Caption         =   "Retur Struk"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   21
      Top             =   2760
      Width           =   8535
      Begin VB.OptionButton optJnsRetur 
         Caption         =   "Pasien Tidak Bayar"
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optJnsRetur 
         Caption         =   "Bayar Ulang"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   2520
         MaxLength       =   100
         TabIndex        =   11
         Top             =   840
         Width           =   5775
      End
      Begin MSComCtl2.DTPicker dtpTglRetur 
         Height          =   330
         Left            =   240
         TabIndex        =   10
         Top             =   840
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
         CalendarBackColor=   -2147483624
         CalendarForeColor=   128
         CalendarTitleForeColor=   128
         CustomFormat    =   "dd/MM/yyyy HH:mm:ss"
         Format          =   120913923
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan"
         Height          =   210
         Left            =   2520
         TabIndex        =   23
         Top             =   600
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Retur"
         Height          =   210
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   1155
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
      Height          =   1815
      Left            =   0
      TabIndex        =   24
      Top             =   960
      Width           =   8535
      Begin VB.TextBox txtNoRetur 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   20
         ToolTipText     =   "Tekan tombol Enter untuk mencari struk pasien yang akan diretur"
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtUmur 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtNoBKM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   19
         ToolTipText     =   "Tekan tombol Enter untuk mencari struk pasien yang akan diretur"
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtNoStruk 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   240
         MaxLength       =   10
         TabIndex        =   0
         ToolTipText     =   "Tekan tombol Enter untuk mencari struk pasien yang akan diretur"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtJenisPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtPenjamin 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1200
         Width           =   4095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "No. retur"
         Height          =   210
         Index           =   2
         Left            =   6600
         TabIndex        =   40
         Top             =   120
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Umur"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   39
         Top             =   960
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "No. BKM"
         Height          =   210
         Index           =   1
         Left            =   5280
         TabIndex        =   38
         Top             =   120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "No. Struk"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Left            =   1560
         TabIndex        =   30
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   3000
         TabIndex        =   29
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   3960
         TabIndex        =   28
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Index           =   0
         Left            =   7200
         TabIndex        =   27
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pasien"
         Height          =   210
         Left            =   2760
         TabIndex        =   26
         Top             =   960
         Width           =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Penjamin"
         Height          =   210
         Left            =   4200
         TabIndex        =   25
         Top             =   960
         Width           =   735
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
      Left            =   6720
      Picture         =   "frmReturStruk.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmReturStruk.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmReturStruk.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmReturStruk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim subTglStruk As Date

Private Function sp_Retur() As Boolean
    On Error GoTo errLoad

    sp_Retur = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRetur", adChar, adParamInput, 10, txtNoRetur.Text)
        .Parameters.Append .CreateParameter("TglRetur", adDate, adParamInput, , Format(dtpTglRetur.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 50, IIf(Len(Trim(txtKeterangan.Text)) = 0, Null, Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("OutputNoRetur", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_Retur"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam retur struk", vbCritical, "Informasi"
            sp_Retur = False
        Else
            txtNoRetur.Text = .Parameters("OutputNoRetur").Value
            Call Add_HistoryLoginActivity("Add_Retur")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function
errLoad:
    sp_Retur = False
    Call msubPesanError
End Function

Private Function sp_DetailReturPelayananPasien() As Boolean
    On Error GoTo errLoad

    sp_DetailReturPelayananPasien = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtnocm.Text)
        .Parameters.Append .CreateParameter("NoRetur", adChar, adParamInput, 10, txtNoRetur.Text)
        .Parameters.Append .CreateParameter("NoBKM", adChar, adParamInput, 10, txtNoBKM.Text)
        .Parameters.Append .CreateParameter("NoStruk", adChar, adParamInput, 10, txtNoStruk.Text)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 2, IIf(optJnsRetur(0).Value = True, "BU", "TB"))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_DetailReturStrukPelayananPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam detail retur struk", vbCritical, "Informasi"
            sp_DetailReturPelayananPasien = False
        Else
            Call Add_HistoryLoginActivity("Add_DetailReturStrukPelayananPasien")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function
errLoad:
    sp_DetailReturPelayananPasien = False
    Call msubPesanError
End Function

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad
    If Periksa("text", txtNoStruk, "No Struk kosong") = False Then Exit Sub

    If sp_Retur() = False Then Exit Sub
    If sp_DetailReturPelayananPasien() = False Then Exit Sub
    Call subClear
    txtNoStruk.SetFocus

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dtpTglRetur_Change()
    dtpTglRetur.MaxDate = Now
End Sub

Private Sub dtpTglRetur_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If dtpTglRetur.Value < subTglStruk Then
            MsgBox "Tanggal retur tidak bisa lebih kecil dari tanggal struk", vbCritical
            dtpTglRetur.SetFocus
            dtpTglRetur.Value = subTglStruk
            Exit Sub
        End If
        txtKeterangan.SetFocus
    End If
End Sub

Private Sub dtpTglRetur_LostFocus()
    If dtpTglRetur.Value < subTglStruk Then
        MsgBox "Tanggal retur tidak bisa lebih kecil dari tanggal struk", vbCritical
        dtpTglRetur.SetFocus
        dtpTglRetur.Value = subTglStruk
        Exit Sub
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    dtpTglRetur.Value = Now
    Call PlayFlashMovie(Me)
End Sub

Private Sub optJnsRetur_Click(Index As Integer)
    If Index = 0 Then
        txtKeterangan.Text = "Bayar Ulang"
    Else
        txtKeterangan.Text = ""
    End If
End Sub

Private Sub optJnsRetur_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglRetur.SetFocus
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtKeterangan_LostFocus()
    txtKeterangan.Text = StrConv(txtKeterangan.Text, vbProperCase)
End Sub

Public Sub txtNoStruk_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
        strSQL = "SELECT NoBKM, NoPendaftaran, NoCM, NamaPasien, JK, Umur, JenisPasien, NamaPenjamin, TotalBiaya, JmlBayar, JmlHutangPenjamin, JmlTanggunganRS, JmlPembebasan" & _
        " FROM V_JudulStrukPembayaranPasien " & _
        " WHERE NoStruk = '" & txtNoStruk.Text & "'"
        Call msubRecFO(rs, strSQL)

        If rs.EOF = True Then
            MsgBox "No. Struk belum ada", vbExclamation, "Validasi": Call subClear: Exit Sub
        Else
            Call msubRecFO(dbRst, "SELECT NoBKM FROM DetailReturStrukPelayananPasien WHERE (NoBKM = '" & rs("NoBKM").Value & "')")
            If dbRst.EOF = False Then MsgBox "No. BKM " & rs("NoBKM").Value & " sudah pernah diretur", vbExclamation, "Validasi": Call subClear: Exit Sub
        End If

        txtNoBKM.Text = rs("NoBKM").Value
        txtnopendaftaran.Text = IIf(IsNull(rs("NoPendaftaran")), "", rs("NoPendaftaran"))
        txtnocm.Text = IIf(IsNull(rs("NoCM")), "", rs("NoCM"))
        txtNamaPasien.Text = IIf(IsNull(rs("NamaPasien")), "", rs("NamaPasien"))
        txtSex.Text = IIf(IsNull(rs("JK")), "", IIf(rs("JK") = "L", "Laki-Laki", "Perempuan"))
        txtumur.Text = IIf(IsNull(rs("Umur")), "", rs("Umur"))
        txtJenisPasien.Text = rs("JenisPasien").Value
        txtPenjamin.Text = rs("NamaPenjamin").Value

        txtTotalBiaya.Text = rs("TotalBiaya").Value
        txtJmlByr.Text = rs("JmlBayar").Value
        txtHtgPjmn.Text = rs("JmlHutangPenjamin").Value
        txtTggRS.Text = rs("JmlTanggunganRS").Value
        txtPembebasan.Text = rs("JmlPembebasan").Value

        optJnsRetur(0).SetFocus
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subClear()
    txtNoBKM.Text = ""
    txtNoRetur.Text = ""
    txtnopendaftaran.Text = ""
    txtnocm.Text = ""
    txtNamaPasien.Text = ""
    txtSex.Text = ""
    txtumur.Text = ""
    txtJenisPasien.Text = ""
    txtPenjamin.Text = ""
    dtpTglRetur.Value = Now
    txtKeterangan.Text = ""
    txtTotalBiaya.Text = ""
    txtJmlByr.Text = ""
    txtHtgPjmn.Text = ""
    txtTggRS.Text = ""
    txtPembebasan.Text = ""
End Sub

