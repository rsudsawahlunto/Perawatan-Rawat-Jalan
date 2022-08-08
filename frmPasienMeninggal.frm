VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPasienMeninggal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pasien Meninggal"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9525
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPasienMeninggal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   9525
   Begin VB.TextBox txtKdPenyebab 
      Height          =   375
      Left            =   8400
      TabIndex        =   29
      Text            =   "8"
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
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
      Height          =   2055
      Left            =   0
      TabIndex        =   27
      Top             =   2160
      Visible         =   0   'False
      Width           =   9495
      Begin MSDataGridLib.DataGrid dgDokter 
         Height          =   1695
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   2990
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
      TabIndex        =   26
      Top             =   3240
      Width           =   9495
      Begin VB.CommandButton cmdsimpan 
         Caption         =   "&Simpan"
         Height          =   340
         Left            =   5160
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   340
         Left            =   7320
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
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
      TabIndex        =   22
      Top             =   2160
      Width           =   9495
      Begin VB.TextBox txtDokter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6000
         TabIndex        =   9
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox txtPenyebab 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   0
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker dtpTglMeninggal 
         Height          =   345
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm:ss"
         Format          =   125239299
         UpDown          =   -1  'True
         CurrentDate     =   38076
      End
      Begin MSDataListLib.DataCombo dcPenyebabKematian 
         Height          =   330
         Left            =   2520
         TabIndex        =   30
         Top             =   480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         ListField       =   ""
         Text            =   ""
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Dokter Penanggung Jawab"
         Height          =   210
         Left            =   6000
         TabIndex        =   25
         Top             =   240
         Width           =   2220
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Penyebab Kematian"
         Height          =   210
         Left            =   2520
         TabIndex        =   24
         Top             =   240
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Meninggal"
         Height          =   210
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   1185
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
      Height          =   1215
      Left            =   0
      TabIndex        =   13
      Top             =   960
      Width           =   9495
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
         Left            =   6960
         TabIndex        =   14
         Top             =   360
         Width           =   2415
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
            Left            =   1680
            MaxLength       =   6
            TabIndex        =   6
            Top             =   240
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
            Height          =   285
            Left            =   900
            MaxLength       =   6
            TabIndex        =   5
            Top             =   240
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
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   4
            Top             =   240
            Width           =   375
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
            Left            =   2130
            TabIndex        =   17
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
            Left            =   1350
            TabIndex        =   16
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
            Left            =   555
            TabIndex        =   15
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
         Width           =   1335
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3360
         TabIndex        =   2
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5760
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   1680
         TabIndex        =   20
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   3360
         TabIndex        =   19
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   5760
         TabIndex        =   18
         Top             =   360
         Width           =   1065
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   28
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
      Left            =   7680
      Picture         =   "frmPasienMeninggal.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPasienMeninggal.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPasienMeninggal.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmPasienMeninggal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strkodedokter As String
Dim strFilterDokter As String

Private Sub cmdSimpan_Click()
On Error GoTo errLoad

'    If txtPenyebab.Text = "" Then
'        MsgBox "Penyebab kematian harus diisi", vbCritical, "Informasi"
'        txtPenyebab.SetFocus
'        Exit Sub
'    End If

 If dcPenyebabKematian.Text = "" Then
        MsgBox "Penyebab kematian harus diisi", vbCritical, "Informasi"
        dcPenyebabKematian.SetFocus
        Exit Sub
    End If
    
    'If strkodedokter = "" Then
     If txtDokter.Text = "" Then
     
        MsgBox "Dokter penanggung jawab pasien harus diisi", vbCritical, "Informasi"
        txtDokter.SetFocus
        Exit Sub
    End If
    
    Dim adoCommand As New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM)
        .Parameters.Append .CreateParameter("TglMeninggal", adDate, adParamInput, , Format(dtpTglMeninggal, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdPenyebab", adTinyInt, adParamInput, , dcPenyebabKematian.BoundText)
'        .Parameters.Append .CreateParameter("PenyebabKematian", adVarChar, adParamInput, 50, dcPenyebabKematian.Text)
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, strkodedokter)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PasienMeninggal"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Data", vbCritical, "Validasi"
            Exit Sub
        Else
            Call Add_HistoryLoginActivity("Add_PasienMeninggal")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
   
    MsgBox "Data berhasil disimpan", vbInformation, "Informasi"
    Call subDisableControl(False)
    Call frmPasienPulangRJ.subSavePsnPulang

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
 Unload Me
End Sub



Private Sub dcPenyebabKematian_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtDokter.SetFocus
End Sub

Private Sub dgDokter_DblClick()
     Call dgDokter_KeyPress(13)
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        If intJmlDokter = 0 Then Exit Sub
        txtDokter.Text = dgDokter.Columns(1).Value
        strkodedokter = dgDokter.Columns(0).Value
        If strkodedokter = "" Then
            MsgBox "Pilih dulu Dokter yang akan menangani Pasien", vbCritical, "Validasi"
            txtDokter.Text = ""
            dgDokter.SetFocus
            Exit Sub
        End If
        fraDokter.Visible = False
        Me.Height = 4470
    ElseIf KeyAscii = 27 Then
        fraDokter.Visible = False
        txtDokter.SetFocus
        Me.Height = 4350
    End If
End Sub

Private Sub dtpTglMeninggal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        dcPenyebabKematian.SetFocus
    End If
End Sub

Private Sub Form_Load()

    Call PlayFlashMovie(Me)
    dtpTglMeninggal.Value = Now
    Call centerForm(Me, MDIUtama)
    Me.Height = 4350
     
    dcPenyebabKematian.Enabled = True
    strSQL = "Select * from PenyebabKematian where StatusEnabled = '1'"
    Call msubDcSource(dcPenyebabKematian, rs, strSQL)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strFilterDokter = ""
    strkodedokter = ""
    frmPasienPulangRJ.Enabled = True
End Sub

Private Sub txtDokter_Change()
    strFilterDokter = "WHERE NamaDokter like '%" & txtDokter.Text & "%'"
    strkodedokter = ""
    fraDokter.Visible = True
    Me.Height = 5190
    Call subLoadDokter
End Sub

Private Sub txtDokter_GotFocus()
    txtDokter.SelStart = 0
    txtDokter.SelLength = Len(txtDokter.Text)
    fraDokter.Visible = True
    strFilterDokter = "WHERE NamaDokter like '%" & txtDokter.Text & "%'"
    Call subLoadDokter
End Sub

'untuk meload data dokter di grid
Private Sub subLoadDokter()
    On Error Resume Next
    strSQL = "SELECT KodeDokter AS [Kode Dokter],NamaDokter AS [Nama Dokter],JK,Jabatan FROM V_DaftarDokter " & strFilterDokter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlDokter = rs.RecordCount
    Set dgDokter.DataSource = rs
    With dgDokter
        .Columns(0).Width = 1200
        .Columns(1).Width = 3000
        .Columns(2).Width = 400
        .Columns(3).Width = 3000
    End With
End Sub

Private Sub txtDokter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If fraDokter.Visible = False Then Exit Sub
        dgDokter.SetFocus
    End If
End Sub

Private Sub txtDokter_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    If KeyAscii = 27 Then
        txtDokter = ""
        txtDokter.SetFocus
        fraDokter.Visible = False
        Me.Height = 4350
    End If
    If KeyAscii = 13 Then
        dgDokter.SetFocus
    End If
    If KeyAscii = 39 Then KeyAscii = 0
    Call SetKeyPressToChar(KeyAscii)
hell:
End Sub

Private Sub txtPenyebab_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDokter.SetFocus
    End If
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtPenyebab_LostFocus()
    txtPenyebab = StrConv(txtPenyebab, vbProperCase)
End Sub

'untuk enable/disable control2
Private Sub subDisableControl(blnStatus As Boolean)
    dtpTglMeninggal.Enabled = blnStatus
    txtPenyebab.Enabled = blnStatus
    txtDokter.Enabled = blnStatus
    cmdSimpan.Enabled = blnStatus
End Sub
