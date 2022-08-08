VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTerimaDarahRuangan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Kirim Darah"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTerimaDarahRuangan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   10740
   Begin VB.TextBox txtNoKirim 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4440
      MaxLength       =   10
      TabIndex        =   43
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtKdRuanganTujuan 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      MaxLength       =   10
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   465
      Left            =   7200
      TabIndex        =   7
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   465
      Left            =   9000
      TabIndex        =   8
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   465
      Left            =   5400
      TabIndex        =   40
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Frame fraKirim 
      Caption         =   "Kirim Darah"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   0
      TabIndex        =   31
      Top             =   3000
      Width           =   10695
      Begin MSDataListLib.DataCombo dcAsalDarah 
         Height          =   330
         Left            =   2280
         TabIndex        =   41
         Top             =   2520
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.TextBox txtIsi 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   4200
         TabIndex        =   37
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo dcGolDarah 
         Height          =   330
         Left            =   2280
         TabIndex        =   38
         Top             =   2160
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcBentukDarah 
         Height          =   330
         Left            =   2280
         TabIndex        =   39
         Top             =   1800
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   960
         Width           =   8415
      End
      Begin MSDataListLib.DataCombo dcPJawab 
         Height          =   330
         Left            =   7560
         TabIndex        =   5
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
      Begin MSDataListLib.DataCombo dcPegawaiPemeriksa 
         Height          =   330
         Left            =   7560
         TabIndex        =   4
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
      Begin MSComCtl2.DTPicker dtpTglKirim 
         Height          =   330
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   116326403
         UpDown          =   -1  'True
         CurrentDate     =   38077
      End
      Begin MSComCtl2.DTPicker dtpTglPeriksa 
         Height          =   330
         Left            =   2160
         TabIndex        =   2
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   116326403
         UpDown          =   -1  'True
         CurrentDate     =   38077
      End
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   2295
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   4048
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   -2147483643
         FocusRect       =   2
         HighLight       =   2
         Appearance      =   0
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Kirim"
         Height          =   210
         Left            =   240
         TabIndex        =   36
         Top             =   300
         Width           =   690
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Pegawai Pemeriksa"
         Height          =   210
         Left            =   5640
         TabIndex        =   35
         Top             =   300
         Width           =   1545
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Penanggung Jawab"
         Height          =   210
         Left            =   5640
         TabIndex        =   34
         Top             =   640
         Width           =   1605
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Periksa"
         Height          =   210
         Left            =   240
         TabIndex        =   33
         Top             =   660
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan"
         Height          =   210
         Left            =   240
         TabIndex        =   32
         Top             =   1020
         Width           =   945
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Pemesan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   10695
      Begin VB.TextBox txtAsalRuangan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   7560
         MaxLength       =   50
         TabIndex        =   29
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox txtNoOrder 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   18
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   17
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   16
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   7560
         TabIndex        =   15
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
         Left            =   9120
         MaxLength       =   6
         TabIndex        =   14
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
         Left            =   8340
         MaxLength       =   6
         TabIndex        =   13
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
         Left            =   7560
         MaxLength       =   6
         TabIndex        =   12
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtSubInstalasi 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   7560
         TabIndex        =   11
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Asal Ruangan"
         Height          =   210
         Index           =   3
         Left            =   5640
         TabIndex        =   30
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No. Order"
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   27
         Top             =   1440
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   240
         TabIndex        =   25
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   5640
         TabIndex        =   24
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Umur"
         Height          =   210
         Left            =   5640
         TabIndex        =   23
         Top             =   720
         Width           =   435
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "hr"
         Height          =   210
         Left            =   9570
         TabIndex        =   22
         Top             =   750
         Width           =   165
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "bln"
         Height          =   210
         Left            =   8790
         TabIndex        =   21
         Top             =   750
         Width           =   240
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "thn"
         Height          =   210
         Left            =   7995
         TabIndex        =   20
         Top             =   750
         Width           =   285
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Kasus Penyakit"
         Height          =   210
         Left            =   5640
         TabIndex        =   19
         Top             =   1080
         Width           =   1200
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
   Begin VB.Image Image2 
      Height          =   945
      Left            =   8880
      Picture         =   "frmTerimaDarahRuangan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmTerimaDarahRuangan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmTerimaDarahRuangan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub subKosong()
    On Error GoTo hell
    txtNoOrder.Text = ""
    dtpTglKirim.Value = Now
    dtpTglPeriksa.Value = Now
    txtKeterangan.Text = ""
    dcPegawaiPemeriksa.Text = ""
    dcPJawab.Text = ""
    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub subSetGrid()
    On Error GoTo hell
    With fgData
        .Clear
        .Rows = 2
        .Cols = 8

        .RowHeight(0) = 400

        .TextMatrix(0, 0) = "Bentuk Darah"
        .TextMatrix(0, 1) = "Gol. Darah"
        .TextMatrix(0, 2) = "Asal Darah"
        .TextMatrix(0, 3) = "Qty Darah"
        .TextMatrix(0, 4) = "No Labu"
        .TextMatrix(0, 5) = "KdBentukDarah"
        .TextMatrix(0, 6) = "KdGolDarah"
        .TextMatrix(0, 7) = "KdAsalDarah"

        .ColWidth(0) = 4000
        .ColWidth(1) = 1500
        .ColWidth(2) = 1500
        .ColWidth(3) = 1500
        .ColWidth(4) = 1300
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0

        .ColAlignment(3) = flexAlignRightCenter
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Public Sub subLoadData(strNoKirim As String)

    On Error Resume Next
    strSQL = "SELECT     TOP (200) NoKirim, BentukDarah, GolonganDarah, NamaAsal, KdBarang, KdBentukDarah, KdGolonganDarah, KdAsal, JmlDarah, NoLabu, " & _
    " NoPendaftaran , NoCM " & _
    " FROM V_DetailTerimaDarahRuangan where NoKirim ='" & strNoKirim & "'"
    Call msubRecFO(rs, strSQL)

    For i = 1 To rs.RecordCount
        With fgData
            .TextMatrix(i, 0) = rs.Fields("BentukDarah")
            .TextMatrix(i, 1) = rs.Fields("GolonganDarah")
            .TextMatrix(i, 2) = rs.Fields("NamaAsal")
            .TextMatrix(i, 3) = rs.Fields("JmlDarah")
            .TextMatrix(i, 4) = rs.Fields("NoLabu")
            .TextMatrix(i, 5) = rs.Fields("KdBentukDarah")
            .TextMatrix(i, 6) = rs.Fields("KdGolonganDarah")
            .TextMatrix(i, 7) = rs.Fields("KdAsal")

            .Rows = .Rows + 1
        End With
    Next i
End Sub

Sub subLoadDcSource()
    On Error GoTo hell

    strSQL = "Select IdPegawai,NamaLengkap From DataPegawai"
    Call msubDcSource(dcPegawaiPemeriksa, rs, strSQL)
    Call msubDcSource(dcPJawab, rs, strSQL)

    Call msubDcSource(dcBentukDarah, rs, "Select KdBentukDarah,BentukDarah From BentukDarah Where StatusEnabled=1")
    Call msubDcSource(dcGolDarah, rs, "Select KdGolonganDarah,GolonganDarah From GolonganDarah Where StatusEnabled=1")
    Call msubDcSource(dcAsalDarah, rs, "Select KdAsal,NamaAsal From AsalBarang Where StatusEnabled=1")

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub subLoadText()
    Dim i As Integer
    txtIsi.Left = fgData.Left

    For i = 0 To fgData.Col - 1
        txtIsi.Left = txtIsi.Left + fgData.ColWidth(i)
    Next i
    txtIsi.Visible = True
    txtIsi.Top = fgData.Top - 7

    For i = 0 To fgData.Row - 1
        txtIsi.Top = txtIsi.Top + fgData.RowHeight(i)
    Next i

    If fgData.TopRow > 1 Then
        txtIsi.Top = txtIsi.Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
    End If

    txtIsi.Width = fgData.ColWidth(fgData.Col)

    txtIsi.Visible = True
    txtIsi.SelStart = Len(txtIsi.Text)
    txtIsi.SetFocus
End Sub

Private Sub subLoadDataCombo(s_DcName As Object)
    Dim i As Integer
    s_DcName.Left = fgData.Left
    For i = 0 To fgData.Col - 1
        s_DcName.Left = s_DcName.Left + fgData.ColWidth(i)
    Next i
    s_DcName.Visible = True
    s_DcName.Top = fgData.Top - 7

    For i = 0 To fgData.Row - 1
        s_DcName.Top = s_DcName.Top + fgData.RowHeight(i)
    Next i

    If fgData.TopRow > 1 Then
        s_DcName.Top = s_DcName.Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
    End If

    s_DcName.Width = fgData.ColWidth(fgData.Col)
    s_DcName.Height = fgData.RowHeight(fgData.Row)

    s_DcName.Visible = True
    s_DcName.SetFocus
End Sub

Private Sub cmdBatal_Click()
    Call subKosong
    Call subSetGrid
    Call subLoadDcSource
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad
    Dim i As Integer
    If txtNoOrder.Text = "" Then MsgBox "Darah tidak perlu melalui proses penerimaan karna tidak sebelumnya tidak di order": Exit Sub
    If fgData.TextMatrix(1, 0) = "" Then MsgBox "Data darah harus diisi", vbExclamation, "Validasi": Exit Sub

    If sp_StrukTerimaRuangan() = False Then Exit Sub

    MsgBox "No Terima : " & txtNoKirim.Text, vbInformation, "Informasi"

    Call cmdBatal_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Function sp_StrukTerimaRuangan() As Boolean
    On Error GoTo errLoad
    sp_StrukTerimaRuangan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoKirim", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("TglKirim", adDate, adParamInput, , Format(dtpTglKirim.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, IIf(txtNoOrder.Text = "", Null, txtNoOrder))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, txtKdRuanganTujuan.Text)
        .Parameters.Append .CreateParameter("IdUserPenerima", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("OutputNoTerima", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_StrukTerimaDarahRuangan"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data struk kirim", vbCritical, "Validasi"
            sp_StrukTerimaRuangan = False
        Else
            txtNoKirim.Text = .Parameters("OutputNoTerima").Value
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Exit Function
errLoad:
    Call msubPesanError("Add_StrukTerimaRuangan")
    sp_StrukTerimaRuangan = False
End Function

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcAsalDarah_Change()
    On Error GoTo errLoad
    fgData.TextMatrix(fgData.Row, 2) = dcAsalDarah.Text
    fgData.TextMatrix(fgData.Row, 7) = dcAsalDarah.BoundText
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcAsalDarah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call dcAsalDarah_Change
        dcAsalDarah.Visible = False
        fgData.Col = 3
        fgData.SetFocus
    End If
End Sub

Private Sub dcAsalDarah_LostFocus()
    dcAsalDarah.Visible = False
End Sub

Private Sub dcBentukDarah_Change()
    On Error GoTo errLoad
    fgData.TextMatrix(fgData.Row, 0) = dcBentukDarah.Text
    fgData.TextMatrix(fgData.Row, 5) = dcBentukDarah.BoundText
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcBentukDarah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call dcBentukDarah_Change
        dcBentukDarah.Visible = False
        fgData.Col = 1
        fgData.SetFocus
    End If
End Sub

Private Sub dcBentukDarah_LostFocus()
    dcBentukDarah.Visible = False
End Sub

Private Sub dcGolDarah_Change()
    On Error GoTo errLoad
    fgData.TextMatrix(fgData.Row, 1) = dcGolDarah.Text
    fgData.TextMatrix(fgData.Row, 6) = dcGolDarah.BoundText
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcGolDarah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call dcGolDarah_Change
        dcGolDarah.Visible = False
        fgData.Col = 2
        fgData.SetFocus
    End If
End Sub

Private Sub dcGolDarah_LostFocus()
    dcGolDarah.Visible = False
End Sub

Private Sub dcPegawaiPemeriksa_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 13 Then
        If dcPegawaiPemeriksa.MatchedWithList = True Then dcPJawab.SetFocus
        strSQL = "Select IdPegawai,NamaPegawai From DataPegawai WHERE (NamaPegawai LIKE '%" & dcPegawaiPemeriksa.Text & "%') "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcPegawaiPemeriksa.BoundText = rs(0).Value
        dcPegawaiPemeriksa.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcPJawab_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 13 Then
        If dcPJawab.MatchedWithList = True Then fgData.SetFocus
        strSQL = "Select IdPegawai,NamaPegawai From DataPegawai WHERE (NamaPegawai LIKE '%" & dcPJawab.Text & "%') "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcPJawab.BoundText = rs(0).Value
        dcPJawab.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dtpTglKirim_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpTglPeriksa.SetFocus
End Sub

Private Sub dtpTglPeriksa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtKeterangan.SetFocus
End Sub

Private Sub fgData_DblClick()
    Call fgData_KeyDown(13, 0)
End Sub

Private Sub fgData_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer

End Sub

Private Sub fgData_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad

    txtIsi.Text = ""

    Select Case fgData.Col
        Case 0 'bentuk darah
            Call subLoadDataCombo(dcBentukDarah)

        Case 1 'golonga  darah
            Call subLoadDataCombo(dcGolDarah)

        Case 2 'golonga  darah
            Call subLoadDataCombo(dcAsalDarah)

        Case 3 'Jumlah
            txtIsi.MaxLength = 4
            Call subLoadText
            ' txtIsi.Text = Chr(KeyAscii)
            txtIsi.SelStart = Len(txtIsi.Text)

        Case 4 'no labu

            Call subLoadText
            txtIsi.SelStart = Len(txtIsi.Text)
    End Select

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call cmdBatal_Click
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    On Error GoTo errLoad
    If KeyAscii = 13 Then
        Select Case fgData.Col
            Case 3

                fgData.TextMatrix(fgData.Row, 3) = txtIsi.Text
                fgData.Col = 4
                fgData.SetFocus

            Case 4

                fgData.TextMatrix(fgData.Row, 4) = txtIsi.Text
                fgData.Col = 0

                fgData.SetFocus
                SendKeys "{DOWN}"

        End Select

        txtIsi.Visible = False

    ElseIf KeyAscii = 27 Then
        txtIsi.Visible = False
        fgData.SetFocus
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtIsi_LostFocus()
    txtIsi.Visible = False
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcPegawaiPemeriksa.SetFocus
End Sub

