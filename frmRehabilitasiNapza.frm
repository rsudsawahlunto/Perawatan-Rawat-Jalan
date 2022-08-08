VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRehabilitasiNapza 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Rehabilitasi Napza"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRehabilitasiNapza.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   11475
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   7920
      TabIndex        =   5
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   6120
      TabIndex        =   7
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   4320
      TabIndex        =   8
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   9720
      TabIndex        =   6
      Top             =   6960
      Width           =   1635
   End
   Begin VB.TextBox txtNoHasilPeriksa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   0
      MaxLength       =   10
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detail Rehabilitasi Napza"
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
      TabIndex        =   26
      Top             =   2040
      Width           =   11415
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5760
         TabIndex        =   4
         Top             =   1080
         Width           =   5415
      End
      Begin VB.TextBox txtHasilRehabilitasi 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   5415
      End
      Begin MSDataListLib.DataCombo dcPelayananRS 
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcNapza 
         Height          =   330
         Left            =   4440
         TabIndex        =   1
         Top             =   480
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcModelPelayanan 
         Height          =   330
         Left            =   8760
         TabIndex        =   2
         Top             =   480
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan"
         Height          =   210
         Left            =   5760
         TabIndex        =   32
         Top             =   840
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Model Pelayanan"
         Height          =   210
         Left            =   8760
         TabIndex        =   31
         Top             =   240
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nama Napza"
         Height          =   210
         Left            =   4440
         TabIndex        =   30
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pelayanan "
         Height          =   210
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasil Rehabilitasi"
         Height          =   210
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   1275
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
      Height          =   975
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   11415
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
         Left            =   8880
         TabIndex        =   15
         Top             =   240
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            TabIndex        =   16
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   19
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
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         MaxLength       =   12
         TabIndex        =   13
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3600
         TabIndex        =   12
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   7440
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   1680
         TabIndex        =   24
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   3600
         TabIndex        =   23
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   7440
         TabIndex        =   22
         Top             =   240
         Width           =   1065
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   9
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
   Begin MSDataGridLib.DataGrid dgData 
      Height          =   3135
      Left            =   0
      TabIndex        =   33
      Top             =   3720
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   5530
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
   Begin VB.Image Image2 
      Height          =   945
      Left            =   9600
      Picture         =   "frmRehabilitasiNapza.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRehabilitasiNapza.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "frmRehabilitasiNapza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strStatusSP As String

Private Sub cmdBatal_Click()
    Call subKosong
    Call subLoadDcSource
    Call subLoadGrid
    dcNapza.Enabled = True
    dcPelayananRS.Enabled = True
    strStatusSP = ""
End Sub

Private Sub cmdHapus_Click()
Dim vbMsgboxRslt As VbMsgBoxResult
    On Error GoTo hell
    If Trim(txtNoHasilPeriksa.Text) = "" Then
        MsgBox "Pilih data yang akan dihapus", vbExclamation, "Validasi"
        Exit Sub
    End If

    If Trim(dcPelayananRS.BoundText) = "" Then
        MsgBox "Pilih data yang akan dihapus", vbExclamation, "Validasi"
        Exit Sub
    End If

    If strStatusSP <> "U" Then
        MsgBox "Pilih data yang akan dihapus", vbExclamation, "Validasi"
        Exit Sub
    End If
    
    vbMsgboxRslt = MsgBox("Yakin data akan dihapus", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbNo Then Exit Sub
    

    If sp_UDDetailRehabilitasiNapza("D") = False Then Exit Sub

    MsgBox "Data berhasil dihapus", vbInformation, "Sukses"
    strStatusSP = ""
    Call cmdBatal_Click
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell

    If Periksa("datacombo", dcPelayananRS, "Nama Pelayanan kosong") = False Then Exit Sub
    If Periksa("datacombo", dcNapza, "Nama napza kosong") = False Then Exit Sub
    If Periksa("datacombo", dcModelPelayanan, "Model pelayanan kosong") = False Then Exit Sub

    If strStatusSP = "" Then strStatusSP = "A"
    If strStatusSP = "A" Then
        If sp_AddDetailRehabilitasiNapza = False Then Exit Sub
    ElseIf strStatusSP = "U" Then
        If sp_UDDetailRehabilitasiNapza(strStatusSP) = False Then Exit Sub
    End If

    strStatusSP = ""
    MsgBox "Data berhasil disimpan", vbInformation, "Sukses"
    Call cmdBatal_Click

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcModelPelayanan_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcModelPelayanan.MatchedWithList = True Then txtHasilRehabilitasi.SetFocus
        strSQL = "Select KdModelPelayanan,ModelPelayanan From ModelPelayanan Where ModelPelayanan like '%" & dcModelPelayanan.Text & "%' and Statusenabled=1"
        Call msubRecFO(dbRst, strSQL)
        If dbRst.EOF = True Then Exit Sub
        dcModelPelayanan.BoundText = dbRst(0).Value
        dcModelPelayanan.Text = dbRst(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcModelPelayanan_LostFocus()
    If dcModelPelayanan.MatchedWithList = False Then dcModelPelayanan.Text = ""

End Sub

Private Sub dcNapza_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcNapza.MatchedWithList = True Then dcModelPelayanan.SetFocus
        strSQL = "Select KdPelayananRS,NamaPelayanan From ListPelayananRS Where NamaPelayanan like '%" & dcNapza.Text & "%' and Statusenabled=1"
        Call msubRecFO(dbRst, strSQL)
        If dbRst.EOF = True Then dcNapza = "": Exit Sub
        dcNapza.BoundText = dbRst(0).Value
        dcNapza.Text = dbRst(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcNapza_LostFocus()
'        strSQL = "Select KdPelayananRS,NamaPelayanan From ListPelayananRS Where NamaPelayanan like '%" & dcNapza.Text & "%' and Statusenabled=1"
'        Call msubRecFO(dbRst, strSQL)
'        If dbRst.EOF = True Then Exit Sub
'        dcNapza.BoundText = dbRst(0).Value
'        dcNapza.Text = dbRst(1).Value
    If dcNapza.MatchedWithList = False Then dcNapza.Text = ""

End Sub

Private Sub dcPelayananRS_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcPelayananRS.MatchedWithList = True Then dcNapza.SetFocus
        strSQL = "Select KdPelayananRS,NamaPelayanan From ListPelayananRS Where NamaPelayanan like '%" & dcPelayananRS.Text & "%' and Statusenabled=1"
        Call msubRecFO(dbRst, strSQL)
'        Call msubDcSource(dcPelayananRS, dbRst, strSQL)
        If dbRst.EOF = True Then dcPelayananRS = "": Exit Sub
        dcPelayananRS.BoundText = dbRst(0).Value
        dcPelayananRS.Text = dbRst(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcPelayananRS_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> 40 And KeyCode <> 38 Then
        strSQL = "Select KdPelayananRS,NamaPelayanan From ListPelayananRS Where NamaPelayanan like '%" & dcPelayananRS.Text & "%' and Statusenabled=1"
       ' Call msubRecFO(dbRst, strSQL)
        Call msubDcSource(dcPelayananRS, dbRst, strSQL)
End If
End Sub

Private Sub dcPelayananRS_LostFocus()
        strSQL = "Select KdPelayananRS,NamaPelayanan From ListPelayananRS Where NamaPelayanan like '%" & dcPelayananRS.Text & "%' and Statusenabled=1"
       ' Call msubRecFO(dbRst, strSQL)
        Call msubDcSource(dcPelayananRS, dbRst, strSQL)
        If dbRst.EOF = True Then Exit Sub
        dcPelayananRS.BoundText = dbRst(0).Value
        dcPelayananRS.Text = dbRst(1).Value

End Sub

Private Sub dgData_Click()
'    WheelHook.WheelUnHook
'    Set MyProperty = dgData
'    WheelHook.WheelHook dgData
On Error GoTo hell

    If dgData.ApproxCount <= 0 Then Exit Sub
    strStatusSP = "U"
    With dgData
        txtNoHasilPeriksa.Text = .Columns("NoHasilPeriksa")
        dcPelayananRS.BoundText = .Columns("KdPelayananRS")
        dcNapza.BoundText = .Columns("KdNapza")
        dcModelPelayanan.BoundText = .Columns("KdModelPelayanan")
        txtHasilRehabilitasi.Text = .Columns("HasilRehabilitasi")
        txtKeterangan.Text = .Columns("Keterangan")
        dcPelayananRS.Enabled = False
        dcNapza.Enabled = False
    End With

    Exit Sub
hell:

End Sub

Private Sub dgData_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo hell
    Call dgData_Click
'    If dgData.ApproxCount <= 0 Then Exit Sub
''    strStatusSP = "U"
'    With dgData
'        txtNoHasilPeriksa.Text = .Columns("NoHasilPeriksa")
'        dcPelayananRS.BoundText = .Columns("KdPelayananRS")
'        dcNapza.BoundText = .Columns("KdNapza")
'        dcModelPelayanan.BoundText = .Columns("KdModelPelayanan")
'        txtHasilRehabilitasi.Text = .Columns("HasilRehabilitasi")
'        txtKeterangan.Text = .Columns("Keterangan")
'        dcPelayananRS.Enabled = False
'        dcNapza.Enabled = False
'    End With
hell:
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subKosong
    Call subLoadDcSource
    Call subLoadGrid
End Sub

Sub subKosong()
'    txtNoHasilPeriksa.Text = ""
    dcPelayananRS.Text = ""
    dcNapza.Text = ""
    dcModelPelayanan.Text = ""
    txtHasilRehabilitasi.Text = ""
    txtKeterangan.Text = ""
End Sub

Sub subLoadDcSource()
    On Error GoTo hell

    Call msubDcSource(dcPelayananRS, rs, "SELECT KdPelayananRS,NamaPelayanan From ListPelayananRS Where KdPelayananRS='063008'")
    Call msubDcSource(dcNapza, rs, "SELECT KdPelayananRS,NamaPelayanan From ListPelayananRS Where KdJnsPelayanan='901'")
    Call msubDcSource(dcModelPelayanan, rs, "Select KdModelPelayanan,ModelPelayanan From ModelPelayanan Where StatusEnabled=1")

    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub subLoadGrid()
    On Error GoTo hell
    Dim i As Integer

    strSQL = "Select * From V_DetailRehabilitasiNapza Where NoPendaftaran='" & mstrNoPen & "'" 'txtNoPendaftaran.Text
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    Set dgData.DataSource = rs
    With dgData
        For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next i

        .Columns("NoHasilPeriksa").Width = 1200
        .Columns("NamaPelayanan").Width = 2000
        .Columns("NamaNapza").Width = 2000
        .Columns("ModelPelayanan").Width = 1500
        .Columns("HasilRehabilitasi").Width = 1500
        .Columns("Keterangan").Width = 1500
    End With

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Function sp_AddDetailRehabilitasiNapza() As Boolean
    On Error GoTo hell
    sp_AddDetailRehabilitasiNapza = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoHasilPeriksa", adVarChar, adParamInput, 10, IIf(Trim(txtNoHasilPeriksa.Text) = "", Null, Trim(txtNoHasilPeriksa.Text)))
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, Trim(txtnopendaftaran.Text))
        .Parameters.Append .CreateParameter("NoLab_Rad", adChar, adParamInput, 10, IIf(Trim(mstrNoLabRad) = "", Null, Trim(mstrNoLabRad)))
        .Parameters.Append .CreateParameter("NoPakai", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, dcPelayananRS.BoundText)
        .Parameters.Append .CreateParameter("KdNapza", adChar, adParamInput, 6, dcNapza.BoundText)
        .Parameters.Append .CreateParameter("KdModelPelayanan", adTinyInt, adParamInput, , dcModelPelayanan.BoundText)
        .Parameters.Append .CreateParameter("HasilRehabilitasi", adVarChar, adParamInput, 100, IIf(Trim(txtHasilRehabilitasi.Text) = "", Null, Trim(txtHasilRehabilitasi.Text)))
        .Parameters.Append .CreateParameter("KeteranganLainnya", adVarChar, adParamInput, 100, IIf(Trim(txtKeterangan.Text) = "", Null, Trim(txtKeterangan.Text)))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_DetailRehabilitasiNapza"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_AddDetailRehabilitasiNapza = False
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
hell:
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    sp_AddDetailRehabilitasiNapza = False
    MsgBox "Nama Pelayanan dan Nama Napsa sudah di input.", vbCritical, "Validasi"

'    Call msubPesanError
End Function

Private Function sp_UDDetailRehabilitasiNapza(f_status As String) As Boolean
    On Error GoTo hell
    sp_UDDetailRehabilitasiNapza = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoHasilPeriksa", adVarChar, adParamInput, 10, txtNoHasilPeriksa.Text)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, Trim(txtnopendaftaran.Text))
        .Parameters.Append .CreateParameter("NoLab_Rad", adChar, adParamInput, 10, IIf(Trim(mstrNoLabRad) = "", Null, Trim(mstrNoLabRad)))
        .Parameters.Append .CreateParameter("NoPakai", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, dcPelayananRS.BoundText)
        .Parameters.Append .CreateParameter("KdNapza", adChar, adParamInput, 6, dcNapza.BoundText)
        .Parameters.Append .CreateParameter("KdModelPelayanan", adTinyInt, adParamInput, , dcModelPelayanan.BoundText)
        .Parameters.Append .CreateParameter("HasilRehabilitasi", adVarChar, adParamInput, 100, IIf(Trim(txtHasilRehabilitasi.Text) = "", Null, Trim(txtHasilRehabilitasi.Text)))
        .Parameters.Append .CreateParameter("KeteranganLainnya", adVarChar, adParamInput, 100, IIf(Trim(txtKeterangan.Text) = "", Null, Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.UD_DetailRehabilitasiNapza"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_UDDetailRehabilitasiNapza = False
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
hell:
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    sp_UDDetailRehabilitasiNapza = False
    Call msubPesanError
End Function

Private Sub txtHasilRehabilitasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub
