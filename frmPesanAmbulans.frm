VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPesanAmbulans 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Medifirst2000 - Pesan Ambulans"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12630
   Icon            =   "frmPesanAmbulans.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   12630
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNoOrder 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1200
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtNoPendaftaran 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   0
      TabIndex        =   38
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtNoPakai 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtNoCM 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   0
      TabIndex        =   36
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid dgPasien 
      Height          =   2055
      Left            =   3600
      TabIndex        =   4
      Top             =   -1080
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3625
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
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   35
      Top             =   5040
      Width           =   12615
      Begin VB.CommandButton frmPesanAmbulans 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   10800
         TabIndex        =   18
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   9000
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   495
         Left            =   7200
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   495
         Left            =   5400
         TabIndex        =   15
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pesan Ambulans"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   0
      TabIndex        =   19
      Top             =   1080
      Width           =   12615
      Begin VB.TextBox txtQtyPelayanan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8760
         TabIndex        =   13
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtTempatTujuan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8760
         TabIndex        =   11
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox txtNoTelpHp2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2400
         TabIndex        =   8
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox txtNmPnggungJwbPsien 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2400
         TabIndex        =   5
         Top             =   1800
         Width           =   4095
      End
      Begin VB.TextBox txtAlamatTujuan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2400
         TabIndex        =   7
         Top             =   2760
         Width           =   10095
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2400
         TabIndex        =   9
         Top             =   3480
         Width           =   10095
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         Top             =   1440
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker dtpTglOrder 
         Height          =   330
         Left            =   2400
         TabIndex        =   0
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
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
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   120520707
         CurrentDate     =   37760
      End
      Begin MSDataListLib.DataCombo dcRuanganTujuan 
         Height          =   315
         Left            =   2400
         TabIndex        =   1
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcTujuanOrder 
         Height          =   315
         Left            =   8760
         TabIndex        =   10
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcStatusOrder 
         Height          =   315
         Left            =   2400
         TabIndex        =   2
         Top             =   1080
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcPelayananRS 
         Height          =   315
         Left            =   8760
         TabIndex        =   12
         Top             =   1080
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker dtpTglPelayanan 
         Height          =   330
         Left            =   8760
         TabIndex        =   14
         Top             =   1800
         Width           =   2295
         _ExtentX        =   4048
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
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   120520707
         CurrentDate     =   37760
      End
      Begin VB.TextBox txtNoTelpHp1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2400
         TabIndex        =   6
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label lblRed 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   105
         Index           =   7
         Left            =   12480
         TabIndex        =   47
         Top             =   2760
         Width           =   105
      End
      Begin VB.Label lblRed 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   105
         Index           =   6
         Left            =   4680
         TabIndex        =   46
         Top             =   720
         Width           =   105
      End
      Begin VB.Label lblRed 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   105
         Index           =   5
         Left            =   12480
         TabIndex        =   45
         Top             =   720
         Width           =   105
      End
      Begin VB.Label lblRed 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   105
         Index           =   4
         Left            =   9600
         TabIndex        =   44
         Top             =   1440
         Width           =   105
      End
      Begin VB.Label lblRed 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   105
         Index           =   2
         Left            =   12480
         TabIndex        =   43
         Top             =   1080
         Width           =   105
      End
      Begin VB.Label lblRed 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   105
         Index           =   1
         Left            =   11040
         TabIndex        =   42
         Top             =   360
         Width           =   105
      End
      Begin VB.Label lblRed 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   105
         Index           =   0
         Left            =   4680
         TabIndex        =   41
         Top             =   1080
         Width           =   105
      End
      Begin VB.Label lblRed 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   105
         Index           =   3
         Left            =   11400
         TabIndex        =   40
         Top             =   2280
         Width           =   105
      End
      Begin VB.Label Label4 
         Caption         =   "data harus diisi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11640
         TabIndex        =   39
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Pelayanan"
         Height          =   195
         Left            =   6960
         TabIndex        =   34
         Top             =   1800
         Width           =   1020
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Qty Pelayanan"
         Height          =   195
         Left            =   6960
         TabIndex        =   33
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pelayanan"
         Height          =   195
         Left            =   6960
         TabIndex        =   32
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Nama Tempat Tujuan"
         Height          =   195
         Left            =   6960
         TabIndex        =   31
         Top             =   720
         Width           =   1545
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Tujuan Order"
         Height          =   195
         Left            =   6960
         TabIndex        =   30
         Top             =   360
         Width           =   930
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "No Telp/HP Tempat Tujuan"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   3120
         Width           =   1995
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Alamat Lengkap Tujuan"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "No Telp/HP Penanggung Jawab Keluarga"
         Height          =   435
         Left            =   120
         TabIndex        =   27
         Top             =   2280
         Width           =   1935
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nama Penanggung Jawab Keluarga"
         Height          =   420
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   1485
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Order"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan Lainnya"
         Height          =   210
         Left            =   120
         TabIndex        =   23
         Top             =   3480
         Width           =   1605
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ruangan Tujuan"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Status Order"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   945
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   25
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
      Left            =   10800
      Picture         =   "frmPesanAmbulans.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPesanAmbulans.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "frmPesanAmbulans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tempStatusTampil As Boolean
Dim a As Boolean

Private Sub cmdCetak_Click()
    On Error GoTo hell
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmCetakStrukPesanAmbulans.Show
hell:
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell
    If Periksa("datacombo", dcStatusOrder, "Status Order Kosong") = False Then Exit Sub
    If Periksa("datacombo", dcTujuanOrder, "Tujuan Order Kosong") = False Then Exit Sub
    If Periksa("datacombo", dcPelayananRS, "Pelayanan RS Kosong") = False Then Exit Sub
    If Periksa("text", txtQtyPelayanan, "Qty Pelayanan Kosong") = False Then Exit Sub
    If Periksa("text", txtTempatTujuan, "Ruangan Tujuan Kosong") = False Then Exit Sub
    If Periksa("text", txtAlamatTujuan, "Alamat Lengkap Tujuan Kosong") = False Then Exit Sub

    If sp_PesanAmbulans() = False Then Exit Sub
    MsgBox "Data berhasil disimpan", vbInformation, "Informasi"

    Call Add_HistoryLoginActivity("Add_PesanAmbulans")

    Call cmdBatal_Click
    NoOrder = txtNoOrder.Text
    cmdCetak.Enabled = True
    Exit Sub
hell:
    msubPesanError
End Sub

Private Function sp_PesanAmbulans() As Boolean
    On Error GoTo errLoad

    sp_PesanAmbulans = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("TglOrder", adDate, adParamInput, , Format(dtpTglPelayanan.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, IIf(txtnopendaftaran.Text = "", Null, txtnopendaftaran.Text))
        .Parameters.Append .CreateParameter("NoPakai", adChar, adParamInput, 10, IIf(txtNoPakai.Text = "", Null, txtNoPakai.Text))
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, IIf(txtnocm.Text = "", Null, txtnocm.Text))
        .Parameters.Append .CreateParameter("NamaPasien", adVarChar, adParamInput, 40, txtNamaPasien.Text)
        .Parameters.Append .CreateParameter("NamaPJawabKeluarga", adVarChar, adParamInput, 10, txtNmPnggungJwbPsien.Text)
        .Parameters.Append .CreateParameter("NoTlpHP", adVarChar, adParamInput, 30, txtNoTelpHp1.Text)
        .Parameters.Append .CreateParameter("KdStatusOrder", adTinyInt, adParamInput, , dcStatusOrder.BoundText)
        .Parameters.Append .CreateParameter("KdTujuanOrder", adTinyInt, adParamInput, , dcTujuanOrder.BoundText)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, dcPelayananRS.BoundText)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(dtpTglPelayanan.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("QtyPelayanan", adDouble, adParamInput, , txtQtyPelayanan.Text)
        .Parameters.Append .CreateParameter("NamaTempatTujuan", adVarChar, adParamInput, 75, txtTempatTujuan.Text)
        .Parameters.Append .CreateParameter("AlamatLengkapTempatTujuan", adVarChar, adParamInput, 150, txtAlamatTujuan.Text)
        .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, dcRuanganTujuan.BoundText)
        .Parameters.Append .CreateParameter("NoTlpHPTempatTujuan", adVarChar, adParamInput, 30, IIf(txtNoTelpHp2.Text = "", Null, txtNoTelpHp2.Text))
        .Parameters.Append .CreateParameter("KeteranganLainnya", adVarChar, adParamInput, 150, IIf(txtKeterangan.Text = "", Null, txtKeterangan.Text))
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("NoRiSchedule", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("NoRetur", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("QtyRetur", adDouble, adParamInput, , 0)
        .Parameters.Append .CreateParameter("KeteranganReturLainnya", adVarChar, adParamInput, 150, Null)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("OutputNoOrder", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PesanAmbulans"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_PesanAmbulans = False
        Else
            txtNoOrder.Text = .Parameters("OutputNoOrder")

        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function
errLoad:
    Call msubPesanError
    sp_PesanAmbulans = False
End Function

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdBatal_Click()
    Call subKosong
    Call subLoadDcSource
End Sub

Sub subKosong()
    dtpTglOrder.Value = Now
    dcRuanganTujuan.Text = ""
    dcStatusOrder.Text = ""
    txtNamaPasien.Text = ""
    txtNmPnggungJwbPsien.Text = ""
    txtNoTelpHp1.Text = ""
    txtAlamatTujuan.Text = ""
    txtNoTelpHp2.Text = ""
    txtKeterangan.Text = ""
    dcTujuanOrder.Text = ""
    txtTempatTujuan.Text = ""
    dcPelayananRS.Text = ""
    txtQtyPelayanan.Text = ""
    dtpTglPelayanan.Value = Now
    dgPasien.Visible = False
    dtpTglOrder.SetFocus
    txtNoPakai.Text = ""
    txtnopendaftaran.Text = ""
End Sub

Private Sub subLoadDcSource()
    On Error GoTo errLoad

    strSQL = "SELECT KdRuangan, NamaRuangan From Ruangan WHERE StatusEnabled = 1"
    Call msubDcSource(dcRuanganTujuan, rs, strSQL)

    strSQL = "SELECT KdStatusOrder, StatusOrder From StatusOrder WHERE StatusEnabled = 1"
    Call msubDcSource(dcStatusOrder, rs, strSQL)

    strSQL = "SELECT KdTujuanOrder,TujuanOrder From TujuanOrder WHERE StatusEnabled = 1"
    Call msubDcSource(dcTujuanOrder, rs, strSQL)

    strSQL = "SELECT KdPelayananRS,NamaPelayanan From ListPelayananRS WHERE StatusEnabled = 1"
    Call msubDcSource(dcPelayananRS, rs, strSQL)

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcPelayananRS_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 13 Then
        If dcPelayananRS.MatchedWithList = True Then txtQtyPelayanan.SetFocus
        strSQL = "Select KdPelayananRS,NamaPelayanan From ListPelayananRS Where NamaPelayanan like '%" & dcPelayananRS.Text & "%' and Statusenabled=1"
        Call msubRecFO(dbRst, strSQL)
        If dbRst.EOF = True Then Exit Sub
        dcPelayananRS.BoundText = dbRst(0).Value
        dcPelayananRS.Text = dbRst(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dgPasien_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgPasien
    WheelHook.WheelHook dgPasien
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    dtpTglOrder.Value = Format(Now, "dd MMMM yyyy HH:mm:00")
    dtpTglPelayanan.Value = Format(Now, "dd MMMM yyyy HH:mm:00")
    Call subLoadDcSource
    cmdCetak.Enabled = False
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub frmAmbulansSchedule_Click()
    Unload Me
End Sub

Private Sub frmPesanAmbulans_Click()
    Unload Me
End Sub

Private Sub txtNamaPasien_Change()
    On Error GoTo hell
    If tempStatusTampil = True Then Exit Sub
    strSQL = "select NoPendaftaran,NoCM,[Nama Pasien] from V_DaftarPasienLamaRJ where [Nama Pasien] like '%" & txtNamaPasien.Text & "%' and Ruangan='" & strNNamaRuangan & "'"
    Call msubRecFO(rs, strSQL)
    Set dgPasien.DataSource = rs
    With dgPasien
        .Columns("NoCM").Width = 1200
        .Columns("Nama Pasien").Width = 4340
        .Columns("NoPendaftaran").Width = 1500

        .Top = 2880
        .Left = 2400
        .Visible = True
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub txtNamaPasien_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then If dgPasien.Visible = True Then dgPasien.SetFocus
End Sub

Private Sub txtNamaPasien_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            If dgPasien.Visible = True Then
                dgPasien.SetFocus
            Else
                txtNmPnggungJwbPsien.SetFocus
            End If
        Case 27
            dgPasien.Visible = False
    End Select
End Sub

Private Sub dgPasien_DblClick()
    On Error GoTo hell
    With dgPasien
        If .ApproxCount = 0 Then Exit Sub
        tempStatusTampil = True
        txtNamaPasien.Text = .Columns("Nama Pasien").Value
        txtnocm.Text = .Columns("NoCM").Value
        txtnopendaftaran.Text = .Columns("NoPendaftaran").Value
        txtNmPnggungJwbPsien.SetFocus
        tempStatusTampil = False
        .Visible = False
    End With
    txtNmPnggungJwbPsien.SetFocus
    Exit Sub
hell:
End Sub

Private Sub dgPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call dgPasien_DblClick
    If KeyAscii = 27 Then dgPasien.Visible = False
End Sub

Private Sub dtpTglOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcRuanganTujuan.SetFocus
End Sub

Private Sub dtpTglOrder_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcRuanganTujuan.SetFocus
End Sub

Private Sub dcRuanganTujuan_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 13 Then
        If dcPelayananRS.MatchedWithList = True Then dcStatusOrder.SetFocus
        strSQL = "Select KdRuangan, NamaRuangan From Ruangan Where NamaRuangan like '%" & dcRuanganTujuan.Text & "%' and StatusEnabled=1"
        Call msubRecFO(dbRst, strSQL)
        If dbRst.EOF = True Then Exit Sub
        dcRuanganTujuan.BoundText = dbRst(0).Value
        dcRuanganTujuan.Text = dbRst(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcStatusOrder_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 13 Then
        If dcPelayananRS.MatchedWithList = True Then txtNamaPasien.SetFocus
        strSQL = "Select KdStatusOrder, StatusOrder From StatusOrder Where StatusOrder like '%" & dcStatusOrder.Text & "%' and StatusEnabled=1"
        Call msubRecFO(dbRst, strSQL)
        If dbRst.EOF = True Then Exit Sub
        dcStatusOrder.BoundText = dbRst(0).Value
        dcStatusOrder.Text = dbRst(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub txtNmPnggungJwbPsien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNoTelpHp1.SetFocus
End Sub

Private Sub txtNoTelpHp1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAlamatTujuan.SetFocus
End Sub

Private Sub txtAlamatTujuan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNoTelpHp2.SetFocus
End Sub

Private Sub txtNoTelpHp2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcTujuanOrder.SetFocus
End Sub

Private Sub dcTujuanOrder_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTempatTujuan.SetFocus
End Sub

Private Sub txtTempatTujuan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcPelayananRS.SetFocus
End Sub

Private Sub txtQtyPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglPelayanan.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub dtpTglPelayanan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub dtpTglPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

