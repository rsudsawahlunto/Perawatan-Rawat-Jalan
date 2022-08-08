VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKegiatanPenyuluhan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Kegiatan Penyuluhan"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9075
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmKegiatanPenyuluhan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   9075
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   4440
      TabIndex        =   22
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   2880
      TabIndex        =   21
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   7560
      TabIndex        =   9
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detail Kegiatan Penyuluhan"
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
      TabIndex        =   13
      Top             =   3840
      Width           =   9015
      Begin MSComctlLib.ListView lvwCaraTindakanP 
         Height          =   1455
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cara Tindakan Penyuluhan"
            Object.Width           =   13229
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Kegiatan Penyuluhan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   0
      TabIndex        =   11
      Top             =   1080
      Width           =   9015
      Begin VB.TextBox txtInstitusiPembicara 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3720
         TabIndex        =   5
         Top             =   1680
         Width           =   5055
      End
      Begin VB.TextBox txtNarasumber 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox txtLokasiKegiatan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3720
         TabIndex        =   3
         Top             =   1080
         Width           =   5055
      End
      Begin VB.TextBox txtDeskKegiatan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3720
         TabIndex        =   1
         Top             =   480
         Width           =   5055
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3720
         TabIndex        =   6
         Top             =   2280
         Width           =   5055
      End
      Begin MSDataListLib.DataCombo dcJnsPenyuluhan 
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   1080
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcPJawab 
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker dtpTglPenyuluhan 
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy HH:mm"
         Format          =   251002883
         UpDown          =   -1  'True
         CurrentDate     =   38212
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Penyuluhan"
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Institusi Pembicara"
         Height          =   210
         Left            =   3720
         TabIndex        =   20
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Pembicara Narasumber"
         Height          =   210
         Left            =   120
         TabIndex        =   19
         Top             =   2040
         Width           =   1845
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Lokasi Kegiatan"
         Height          =   210
         Left            =   3720
         TabIndex        =   18
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Penanggung Jawab"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Deskripsi Kegiatan"
         Height          =   210
         Left            =   3720
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Penyuluhan"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan"
         Height          =   210
         Left            =   3720
         TabIndex        =   12
         Top             =   2040
         Width           =   945
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
   Begin MSDataGridLib.DataGrid dgKegiatanPenyuluhan 
      Height          =   2415
      Left            =   0
      TabIndex        =   25
      Top             =   5520
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   15
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
         Locked          =   -1  'True
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtNoRiwayat 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   2160
      TabIndex        =   14
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7200
      Picture         =   "frmKegiatanPenyuluhan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmKegiatanPenyuluhan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "frmKegiatanPenyuluhan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intJmlDataDipilih As Integer
Dim strKdCaraTindakanP() As String
Dim strStatusSP As String
Dim tempbolTampil As Boolean
Dim i As Integer

Private Sub cmdBatal_Click()
    Call subClear
    Call subLoadGrid
    Call subLoadDcSource
    Call subLoadLvw
    Call subLoadDetailTindakanpenyuluhan
    strStatusSP = ""
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo hell
    If Trim(txtNoRiwayat.Text) = "" Then
        MsgBox "Pilih data yang akan dihapus", vbExclamation, "Validasi"
        Exit Sub
    End If

    dbConn.Execute "DELETE FROM DetailKegiatanPenyuluhan WHERE NoRiwayat='" & txtNoRiwayat.Text & "'"
    If sp_UDRiwayatKegiatanPenyuluhan("D") = False Then Exit Sub

    MsgBox "Data berhasil dihapus", vbInformation, "Sukses"
    Call cmdBatal_Click
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
    Dim i As Integer
    If Periksa("datacombo", dcJnsPenyuluhan, "Jenis penyuluhan kosong") = False Then Exit Sub
    If Periksa("datacombo", dcPJawab, "Penanggung jawab kosong") = False Then Exit Sub
    If Periksa("text", txtLokasiKegiatan, "Lokasi kegiatan kosong") = False Then Exit Sub

    If txtNoRiwayat.Text = "" Then strStatusSP = "A"
    If txtNoRiwayat.Text <> "" Then strStatusSP = "U"
    If strStatusSP = "A" Then
        If sp_AddRiwayatKegiatanPenyuluhan = False Then Exit Sub
        For i = 1 To lvwCaraTindakanP.ListItems.Count
            If lvwCaraTindakanP.ListItems(i).Checked = True Then
                If sp_DetailCaraTindakan(Right(lvwCaraTindakanP.ListItems(i).Key, Len(lvwCaraTindakanP.ListItems(i).Key) - 3), "A") = False Then Exit Sub

            End If
        Next i
    ElseIf strStatusSP = "U" Then
        If sp_UDRiwayatKegiatanPenyuluhan(strStatusSP) = False Then Exit Sub

        For i = 1 To lvwCaraTindakanP.ListItems.Count
            If lvwCaraTindakanP.ListItems(i).Checked = True Then
                If sp_DetailCaraTindakan(Right(lvwCaraTindakanP.ListItems(i).Key, Len(lvwCaraTindakanP.ListItems(i).Key) - 3), "A") = False Then Exit Sub
            End If
        Next i
    End If
    strStatusSP = ""
    MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"
    Form_Load

    Exit Sub
hell:
    Call msubPesanError

End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgKegiatanPenyuluhan_DblClick()
    On Error GoTo hell

    If dgKegiatanPenyuluhan.ApproxCount = 0 Then Exit Sub
    strStatusSP = "U"
    With dgKegiatanPenyuluhan
        If IsNull(.Columns(0).Value) Then txtNoRiwayat.Text = "" Else txtNoRiwayat.Text = .Columns(0).Value
        If IsNull(.Columns(1).Value) Then dcJnsPenyuluhan.Text = "" Else dcJnsPenyuluhan.Text = .Columns(1).Value
        If IsNull(.Columns(2).Value) Then txtDeskKegiatan.Text = "" Else txtDeskKegiatan.Text = .Columns(2).Value
        If IsNull(.Columns(3).Value) Then dcPJawab.Text = "" Else dcPJawab.Text = .Columns(3).Value
        If IsNull(.Columns(4).Value) Then txtLokasiKegiatan.Text = "" Else txtLokasiKegiatan.Text = .Columns(4).Value
        If IsNull(.Columns(5).Value) Then txtNarasumber.Text = "" Else txtNarasumber.Text = .Columns(5).Value
        If IsNull(.Columns(6).Value) Then txtInstitusiPembicara.Text = "" Else txtInstitusiPembicara.Text = .Columns(6).Value
        If IsNull(.Columns(7).Value) Then txtKeterangan.Text = "" Else txtKeterangan.Text = .Columns(7).Value

        Call subLoadDetailTindakanpenyuluhan

    End With

    Exit Sub

hell:
    Call msubPesanError
End Sub

Sub subLoadDetailTindakanpenyuluhan()
    On Error GoTo hell
    Dim tempJumData As Integer
    If txtNoRiwayat.Text <> "" Then
        Call subLoadLvw
        strSQL = "SELECT * FROM DetailKegiatanPenyuluhan where NoRiwayat='" & txtNoRiwayat.Text & "'"

        Call msubRecFO(rs, strSQL)
        tempJumData = 0

        Do While rs.EOF = False
            lvwCaraTindakanP.ListItems("key" & rs!KdCaraTindakanP).Checked = True
            lvwCaraTindakanP.ListItems("key" & rs!KdCaraTindakanP).ForeColor = vbBlue
            lvwCaraTindakanP.ListItems("key" & rs!KdCaraTindakanP).Bold = True
            tempJumData = tempJumData + 1
            rs.MoveNext
        Loop

        Exit Sub
    Else
        Exit Sub
    End If
hell:
    msubPesanError
End Sub

Sub subClear()
    txtNoRiwayat.Text = ""
    dcJnsPenyuluhan.Text = ""
    txtDeskKegiatan.Text = ""
    dcPJawab.Text = ""
    txtLokasiKegiatan.Text = ""
    txtNarasumber.Text = ""
    txtInstitusiPembicara.Text = ""
    txtKeterangan.Text = ""
    dtpTglPenyuluhan.Value = Now
End Sub

Sub subLoadGrid()
    On Error GoTo hell
    tempbolTampil = True
    strSQL = "select * from V_RiwayatKegiatanPenyuluhan"

    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    Set dgKegiatanPenyuluhan.DataSource = rs
    With dgKegiatanPenyuluhan
        For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next i

        .Columns(0).Width = 1000
        .Columns(1).Width = 1200
        .Columns(2).Width = 1200
        .Columns(3).Width = 1200
        .Columns(4).Width = 1200
        .Columns(5).Width = 1200
        .Columns(6).Width = 1200
        .Columns(7).Width = 1200

    End With
    Call subLoadDetailTindakanpenyuluhan
    tempbolTampil = False
    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub subLoadDcSource()
    On Error GoTo hell
    Call msubDcSource(dcJnsPenyuluhan, rs, "Select KdJenisPenyuluhan,JenisPenyuluhan From JenisPenyuluhan Where StatusEnabled=1")
    Call msubDcSource(dcPJawab, rs, "Select IdPegawai,NamaLengkap From DataPegawai")
    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub subLoadLvw()
    On Error GoTo errLoad
    Dim strKey As String

    strSQL = "select KdCaraTindakanP,CaraTindakanP from CaraTindakanPenyuluhan where StatusEnabled=1"
    Call msubRecFO(rs, strSQL)
    lvwCaraTindakanP.ListItems.Clear
    lvwCaraTindakanP.Sorted = False
    Do While rs.EOF = False
        strKey = "key" & rs!KdCaraTindakanP
        lvwCaraTindakanP.ListItems.Add , strKey, rs!CaraTindakanP
        rs.MoveNext
    Loop
    lvwCaraTindakanP.Sorted = True

    Exit Sub

errLoad:
    Call msubPesanError
End Sub

Private Sub dcJnsPenyuluhan_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad

    If KeyAscii = 13 Then
        If dcJnsPenyuluhan.MatchedWithList = True Then txtDeskKegiatan.SetFocus
        strSQL = "Select KdJenisPenyuluhan,JenisPenyuluhan From JenisPenyuluhan Where (JenisPenyuluhan LIKE '%" & dcJnsPenyuluhan.Text & "%') And StatusEnabled=1"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcJnsPenyuluhan.BoundText = rs(0).Value
        dcJnsPenyuluhan.Text = rs(1).Value
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcPJawab_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad

    If KeyAscii = 13 Then
        If dcPJawab.MatchedWithList = True Then txtLokasiKegiatan.SetFocus
        strSQL = "Select IdPegawai,NamaLengkap From DataPegawai Where (NamaLengkap LIKE '%" & dcPJawab.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcPJawab.BoundText = rs(0).Value
        dcPJawab.Text = rs(1).Value
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dtpTglPenyuluhan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcJnsPenyuluhan.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    cmdBatal_Click

End Sub

Private Sub lvwCaraTindakanP_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If lvwCaraTindakanP.ListItems(Item.Key).Checked = True Then
        lvwCaraTindakanP.ListItems(Item.Key).ForeColor = vbBlue
    Else
        lvwCaraTindakanP.ListItems(Item.Key).ForeColor = vbBlack
    End If
End Sub

Private Function sp_AddRiwayatKegiatanPenyuluhan() As Boolean
    On Error GoTo hell
    sp_AddRiwayatKegiatanPenyuluhan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdJenisPenyuluhan", adTinyInt, adParamInput, , dcJnsPenyuluhan.BoundText)
        .Parameters.Append .CreateParameter("DeskripsiKegiatan", adVarChar, adParamInput, 150, IIf(Trim(txtDeskKegiatan.Text) = "", Null, Trim(txtDeskKegiatan.Text)))
        .Parameters.Append .CreateParameter("TglPenyuluhan", adDate, adParamInput, , Format(dtpTglPenyuluhan.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdPegawaiPJawab", adChar, adParamInput, 10, dcPJawab.BoundText)
        .Parameters.Append .CreateParameter("LokasiKegiatan", adVarChar, adParamInput, 150, Trim(txtLokasiKegiatan.Text))
        .Parameters.Append .CreateParameter("PembicaraNarasumber", adVarChar, adParamInput, 50, IIf(Trim(txtNarasumber.Text) = "", Null, Trim(txtNarasumber.Text)))
        .Parameters.Append .CreateParameter("InstitusiPembicara", adVarChar, adParamInput, 100, IIf(Trim(txtInstitusiPembicara.Text) = "", Null, Trim(txtInstitusiPembicara.Text)))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 150, IIf(Trim(txtKeterangan.Text) = "", Null, Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("KdCaraTindakanP", adTinyInt, adParamInput, , dcJnsPenyuluhan.BoundText)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("OutNoRiwayat", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_RiwayatKegiatanPenyuluhan"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_AddRiwayatKegiatanPenyuluhan = False
        Else
            txtNoRiwayat.Text = .Parameters("OutNoRiwayat")
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
hell:
    Call msubPesanError

End Function

Private Function sp_UDRiwayatKegiatanPenyuluhan(f_Status As String) As Boolean
    On Error GoTo hell
    sp_UDRiwayatKegiatanPenyuluhan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, txtNoRiwayat.Text)
        .Parameters.Append .CreateParameter("KdJenisPenyuluhan", adTinyInt, adParamInput, , dcJnsPenyuluhan.BoundText)
        .Parameters.Append .CreateParameter("DeskripsiKegiatan", adVarChar, adParamInput, 150, IIf(Trim(txtDeskKegiatan.Text) = "", Null, Trim(txtDeskKegiatan.Text)))
        .Parameters.Append .CreateParameter("IdPegawaiPJawab", adChar, adParamInput, 10, dcPJawab.BoundText)
        .Parameters.Append .CreateParameter("LokasiKegiatan", adVarChar, adParamInput, 150, Trim(txtLokasiKegiatan.Text))
        .Parameters.Append .CreateParameter("PembicaraNarasumber", adVarChar, adParamInput, 50, IIf(Trim(txtNarasumber.Text) = "", Null, Trim(txtNarasumber.Text)))
        .Parameters.Append .CreateParameter("InstitusiPembicara", adVarChar, adParamInput, 100, IIf(Trim(txtInstitusiPembicara.Text) = "", Null, Trim(txtInstitusiPembicara.Text)))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 150, IIf(Trim(txtKeterangan.Text) = "", Null, Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdCaraTindakanP", adTinyInt, adParamInput, , Null)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Ud_RiwayatKegiatanPenyuluhan"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_UDRiwayatKegiatanPenyuluhan = False

        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
hell:
    Call msubPesanError
End Function

Private Sub lvwCaraTindakanP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtDeskKegiatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcPJawab.SetFocus
End Sub

Private Sub txtInstitusiPembicara_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lvwCaraTindakanP.SetFocus
End Sub

Private Sub txtLokasiKegiatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNarasumber.SetFocus
End Sub

Private Sub txtNarasumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtInstitusiPembicara.SetFocus
End Sub

Private Function sp_DetailCaraTindakan(f_KdCaraTindakanP As String, f_Status As String) As Boolean

    sp_DetailCaraTindakan = True

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, txtNoRiwayat.Text)
        .Parameters.Append .CreateParameter("KdCaraTindakanP", adTinyInt, adParamInput, , Trim(f_KdCaraTindakanP))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandType = adCmdStoredProc
        .CommandText = "AUD_DetailCaraTindakan"
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value) = 0 Then
            MsgBox "Ada kesalahan dalam pemasukan data", vbCritical, "Validasi"
            sp_DetailCaraTindakan = False
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing

    End With
End Function
