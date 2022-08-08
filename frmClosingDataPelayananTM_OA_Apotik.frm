VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClosingDataPelayananTM_OA_Apotik 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Pelayanan TM,OA,Apotik"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClosingDataPelayananTM_OA_Apotik.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   14895
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   0
      TabIndex        =   27
      Top             =   6600
      Width           =   14895
      Begin VB.Frame Frame6 
         Caption         =   "Nama Dokter"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4560
         TabIndex        =   31
         Top             =   0
         Width           =   2535
         Begin MSDataListLib.DataCombo dcNamaDokter 
            Height          =   330
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Kriteria Pelayanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   7200
         TabIndex        =   17
         Top             =   0
         Width           =   4335
         Begin VB.OptionButton optApotik 
            Caption         =   "Apotik"
            Height          =   255
            Left            =   3240
            TabIndex        =   14
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton optObatAlKes 
            Caption         =   "Obat AlKes"
            Height          =   255
            Left            =   1800
            TabIndex        =   13
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton optTindakanMedis 
            Caption         =   "Tindakan Medis"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdClosingData 
         Caption         =   "Cl&osing"
         Height          =   495
         Left            =   11760
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   13320
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "F9 = Cetak Rekapitulasi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Label4 
         Caption         =   "Shift = Cetak Singkat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   30
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "F1 = Cetak Detail"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   22
      Top             =   7680
      Width           =   14895
      Begin VB.TextBox txtNoClosing 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSComctlLib.ProgressBar pbData 
         Height          =   360
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Max             =   200
      End
   End
   Begin MSDataGridLib.DataGrid dgClosingDataPelayananTM_OA_Apotik 
      Height          =   3855
      Left            =   0
      TabIndex        =   10
      Top             =   2640
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   6800
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   18
      Top             =   1080
      Width           =   14895
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   360
         Width           =   5175
      End
      Begin VB.Frame Frame4 
         Caption         =   "Periode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8280
         TabIndex        =   19
         Top             =   120
         Width           =   6495
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            Height          =   375
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   855
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   1320
            TabIndex        =   0
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
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   60686339
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3960
            TabIndex        =   1
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
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   60686339
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3600
            TabIndex        =   20
            Top             =   315
            Width           =   255
         End
      End
      Begin MSComCtl2.DTPicker DTPTglClosing 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
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
         CustomFormat    =   "dd MMM yyyy HH:mm"
         Format          =   60686339
         UpDown          =   -1  'True
         CurrentDate     =   38373
      End
      Begin VB.Label Label2 
         Caption         =   "Tanggal Closing"
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Keterangan"
         Height          =   495
         Left            =   2760
         TabIndex        =   25
         Top             =   120
         Width           =   1455
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
   Begin MSDataListLib.DataCombo dcJenisPasien 
      Height          =   330
      Left            =   3540
      TabIndex        =   6
      ToolTipText     =   "Jenis Pasein"
      Top             =   2280
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcKelas 
      Height          =   330
      Left            =   5400
      TabIndex        =   7
      ToolTipText     =   "Kelas"
      Top             =   2280
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcJenisItem 
      Height          =   330
      Left            =   1185
      TabIndex        =   5
      ToolTipText     =   "Jenis Item"
      Top             =   2280
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcNamaItem 
      Height          =   330
      Left            =   7800
      TabIndex        =   8
      ToolTipText     =   "Nama Item"
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcAsalPasien 
      Height          =   330
      Left            =   13920
      TabIndex        =   9
      ToolTipText     =   "Asal Pasien"
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin VB.Label lblJumData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data 0/0"
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   120
      TabIndex        =   29
      Top             =   2280
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmClosingDataPelayananTM_OA_Apotik.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   13080
      Picture         =   "frmClosingDataPelayananTM_OA_Apotik.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmClosingDataPelayananTM_OA_Apotik.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13455
   End
End
Attribute VB_Name = "frmClosingDataPelayananTM_OA_Apotik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer

Private Sub subLoadDcSource()
    Call msubDcSource(dcNamaDokter, rs, "Select KdJenisPegawai, NamaLengkap from DataPegawai where KdJenisPegawai = '001'  order by NamaLengkap")
    Call msubDcSource(dcJenisPasien, rs, "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien where statusenabled='1'  order by JenisPasien")
    Call msubDcSource(dcKelas, rs, "SELECT KdKelas, DeskKelas FROM KelasPelayanan where statusenabled='1'")
    If optTindakanMedis.Value = True Then
        Call msubDcSource(dcJenisItem, rs, "SELECT KdJnsPelayanan, Deskripsi FROM JenisPelayanan where statusenabled='1'")
        Call msubDcSource(dcNamaItem, rs, "SELECT KdPelayananRS, NamaPelayanan FROM ListPelayananRS where statusenabled='1'")
    Else
        Call msubDcSource(dcJenisItem, rs, "SELECT KdDetailJenisBarang, DetailJenisBarang FROM DetailJenisBarang where statusenabled='1'")
        Call msubDcSource(dcNamaItem, rs, "SELECT KdBarang, NamaBarang From MasterBarang where statusenabled='1'")
    End If

    Call msubDcSource(dcAsalPasien, rs, "SELECT KdRuangan, NamaRuangan FROM Ruangan where statusenabled='1'")

End Sub

Private Sub subClearData()
    dcJenisPasien.Text = ""
    dcKelas.Text = ""
    dcJenisItem.Text = ""
    dcNamaItem.Text = ""
    dcAsalPasien.Text = ""
    txtKeterangan.Text = ""
    dcNamaDokter.Text = ""
End Sub

Private Sub subLoadAlignSizeDC()
    dcJenisPasien.Left = 2220
    dcJenisPasien.Width = dgClosingDataPelayananTM_OA_Apotik.Columns("JenisPasien").Width
    dcJenisPasien.Width = 1800

    If optApotik.Value = True Then
        dcKelas.Visible = False
    Else
        dcKelas.Visible = True
        dcKelas.Left = 2200 + dcJenisPasien.Width
        dcKelas.Width = dgClosingDataPelayananTM_OA_Apotik.Columns("Kelas").Width
        dcKelas.Left = 4100
    End If

    If optTindakanMedis.Value = True Then
        dcJenisItem.Left = dcKelas.Left + dgClosingDataPelayananTM_OA_Apotik.Columns("Kelas").Width + _
        dgClosingDataPelayananTM_OA_Apotik.Columns("Nopendaftaran").Width + _
        dgClosingDataPelayananTM_OA_Apotik.Columns("NoCm").Width + _
        dgClosingDataPelayananTM_OA_Apotik.Columns("NamaPasien").Width
        dcJenisItem.Width = dgClosingDataPelayananTM_OA_Apotik.Columns("JenisItem").Width

        dcNamaItem.Left = dcJenisItem.Left + dgClosingDataPelayananTM_OA_Apotik.Columns("JenisItem").Width
        dcNamaItem.Width = dgClosingDataPelayananTM_OA_Apotik.Columns("NamaItem").Width
    ElseIf optApotik.Value = True Then
        dcJenisItem.Left = 6750 + dcJenisPasien.Left
        dcJenisItem.Width = dgClosingDataPelayananTM_OA_Apotik.Columns("JenisItem").Width
        dcJenisItem.Left = 8400
        dcNamaItem.Left = dcJenisItem.Left + dgClosingDataPelayananTM_OA_Apotik.Columns("JenisItem").Width
        dcNamaItem.Width = dgClosingDataPelayananTM_OA_Apotik.Columns("NamaItem").Width
        dcNamaItem.Left = 10800
    ElseIf optObatAlKes.Value = True Then
        dcJenisItem.Left = 8650 + dcJenisPasien.Left
        dcJenisItem.Width = dgClosingDataPelayananTM_OA_Apotik.Columns("JenisItem").Width

        dcNamaItem.Left = dcJenisItem.Left + dgClosingDataPelayananTM_OA_Apotik.Columns("JenisItem").Width
        dcNamaItem.Width = dgClosingDataPelayananTM_OA_Apotik.Columns("NamaItem").Width
    End If

End Sub

Private Function sp_ClosingDataPelayanan()
    On Error GoTo errLoad
    sp_ClosingDataPelayanan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoClosing", adChar, adParamInput, 10, txtNoClosing.Text)
        .Parameters.Append .CreateParameter("TglClosing", adDate, adParamInput, , Format(DTPTglClosing.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 200, IIf(txtKeterangan.Text = "", Null, Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("OutputNoClosing", adChar, adParamInputOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_ClosingDataPelayanan"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam posting data", vbCritical, "Validasi"
            sp_ClosingDataPelayanan = False
        Else
            If Not IsNull(.Parameters("OutputNoClosing").Value) Then txtNoClosing.Text = .Parameters("OutputNoClosing").Value
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Exit Function
errLoad:
    msubPesanError
    sp_ClosingDataPelayanan = False
    cmdClosingData.Enabled = True
    pbData.Value = 0
End Function

Private Function sp_Loop_AddDataPelayananPasienTM_OA_ApotikPH(f_NoPendaftaran As String, f_tglPelayanan As Date, f_KdItem As String, f_KdAsal As String, f_SatuanJml As String, f_NoLab_Rad As String, f_Jenis As String) As Boolean
    On Error GoTo errLoad
    sp_Loop_AddDataPelayananPasienTM_OA_ApotikPH = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoClosing", adChar, adParamInput, 10, txtNoClosing.Text)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(f_tglPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdItem", adChar, adParamInput, 9, f_KdItem)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("SatuanJml", adChar, adParamInput, 1, IIf(f_SatuanJml = "", Null, f_SatuanJml))
        .Parameters.Append .CreateParameter("NoLab_Rad", adChar, adParamInput, 10, f_NoLab_Rad)
        .Parameters.Append .CreateParameter("Jenis", adChar, adParamInput, 2, f_Jenis)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Loop_AddDataPelayananPasienTMOAApotikPH"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan Asuransi Pasien", vbCritical, "Validasi"
            sp_Loop_AddDataPelayananPasienTM_OA_ApotikPH = False
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Exit Function
errLoad:
    msubPesanError
    sp_Loop_AddDataPelayananPasienTM_OA_ApotikPH = False
    cmdClosingData.Enabled = True
End Function

Public Sub subLoadGridSource()
    On Error GoTo errorLoad
    If optTindakanMedis.Value = True Then
        frmClosingDataPelayananTM_OA_Apotik.Caption = "Data Pelayanan Tindakan Medis"
        strSQL = "SELECT  TglPelayanan, JenisPasien, Kelas, NoPendaftaran, NoCM, NamaPasien, JenisItem, NamaItem, JmlItem, HargaSatuan, HargaCito, TotalBiaya, Ruangan," & _
        " AsalPasien , NoClosing, KdItem,  KdAsal, NoLab_Rad, DokterOperator, JmlHutangPenjamin, JmlTanggunganRS FROM V_DaftarDataPelayananTMForClosing" & _
        " WHERE DokterOperator like '%" & dcNamaDokter.Text & "%' AND TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd hh:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd hh:mm:59") & "' " & _
        " AND JenisPasien like '%" & dcJenisPasien.Text & "%' AND Kelas like '%" & dcKelas.Text & "%' AND JenisItem like '" & dcJenisItem.Text & "%' AND NamaItem like '%" & dcNamaItem.Text & "%'  AND AsalPasien LIKE '%" & dcAsalPasien.Text & "%' AND KdRuangan = '" & mstrKdRuangan & "'  "

    ElseIf optObatAlKes.Value = True Then
        frmClosingDataPelayananTM_OA_Apotik.Caption = "Data Pelayanan Obat AlKes"
        strSQL = "SELECT TglPelayanan, JenisPasien, Kelas, NoPendaftaran, NoCM, NamaPasien, AsalItem, JenisItem, NamaItem, JmlItem, HargaSatuan,SatuanJml, HargaService, " & _
        " Administrasi , TotalBiaya, JmlHutangPenjamin, JmlTanggunganRS, Ruangan, AsalPasien, NoClosing,KdItem,NoLab_Rad,  KdAsal" & _
        " FROM V_DaftarDataPelayananOAForClosing" & _
        " WHERE TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd hh:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd hh:mm:59") & "' " & _
        " AND JenisPasien like '%" & dcJenisPasien.Text & "%' AND Kelas like '%" & dcKelas.Text & "%' AND JenisItem like '" & dcJenisItem.Text & "%' AND NamaItem like '%" & dcNamaItem.Text & "%'  AND AsalPasien LIKE '%" & dcAsalPasien.Text & "%' AND KdRuangan = '" & mstrKdRuangan & "'  "

    ElseIf optApotik.Value = True Then
        frmClosingDataPelayananTM_OA_Apotik.Caption = "Data Pelayanan Apotik"
        strSQL = "SELECT TglStruk, JenisPasien, NoStruk, NoPendaftaran,NoCM, NamaPasien, JenisItem, NamaItem, JmlItem, HargaSatuan, HargaService, Administrasi, TotalBiaya, Ruangan, " & _
        " AsalPasien , NoClosing,KdItem,  KdAsal,SatuanJml  FROM         V_DaftarDataPelayananApotikForClosing" & _
        " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd hh:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd hh:mm:59") & "' " & _
        " AND JenisPasien like '%" & dcJenisPasien.Text & "%'  AND JenisItem like '%" & dcJenisItem.Text & "%' AND NamaItem like '%" & dcNamaItem.Text & "%'  AND AsalPasien LIKE '%" & dcAsalPasien.Text & "%' AND KdRuangan = '" & mstrKdRuangan & "'  "
    End If

    Call msubRecFO(rsB, strSQL)
    Set dgClosingDataPelayananTM_OA_Apotik.DataSource = rsB
    lblJumData.Caption = "Data " & dgClosingDataPelayananTM_OA_Apotik.Bookmark & "/" & dgClosingDataPelayananTM_OA_Apotik.ApproxCount
    With dgClosingDataPelayananTM_OA_Apotik
        If optApotik.Value = True Then
            .Columns("TglStruk").Width = 1900
        Else
            .Columns("TglPelayanan").Width = 1900
            .Columns("Kelas").Width = 1300
            .Columns("NoLab_Rad").Width = 0
        End If

        If optTindakanMedis.Value = True Then
            .Columns("HargaCito").Width = 900
            .Columns("HargaCito").Alignment = vbRightJustify
            .Columns("HargaCito").NumberFormat = "#,###.00"
            .Columns("JenisItem").Width = 2300
        Else
            .Columns("SatuanJml").Width = 0
            .Columns("HargaService").Width = 1000
            .Columns("HargaService").NumberFormat = "#,###.00"
            .Columns("HargaService").Alignment = vbRightJustify
            .Columns("Administrasi").Width = 1000
            .Columns("Administrasi").NumberFormat = "#,###.00"
            .Columns("Administrasi").Alignment = vbRightJustify
            .Columns("JenisItem").Width = 1000
        End If

        .Columns("KdItem").Width = 0
        .Columns("KdAsal").Width = 0

        .Columns("NoPendaftaran").Width = 1300
        .Columns("NoCM").Width = 800
        .Columns("JenisPasien").Width = 1500
        .Columns("NamaPasien").Width = 2000

        .Columns("NamaItem").Width = 2000
        .Columns("JmlItem").Width = 700
        .Columns("JmlItem").Alignment = vbCenter
        .Columns("HargaSatuan").Width = 1100
        .Columns("HargaSatuan").Alignment = vbRightJustify
        .Columns("HargaSatuan").NumberFormat = "#,###.00"
        .Columns("TotalBiaya").Width = 1200
        .Columns("TotalBiaya").NumberFormat = "#,###.00"
        .Columns("TotalBiaya").Alignment = vbRightJustify
        .Columns("Ruangan").Width = 2000
        .Columns("AsalPasien").Width = 2000
        .Columns("NoClosing").Width = 1500
    End With

    Exit Sub
errorLoad:
    Call msubPesanError
End Sub

Public Sub cmdCari_Click()
    On Error Resume Next
    MousePointer = vbHourglass
    pbData.Value = 0
    cmdClosingData.Enabled = True
    Call subLoadDcSource
    Call subLoadGridSource
    Call subLoadAlignSizeDC
    MousePointer = vbNormal
    If dgClosingDataPelayananTM_OA_Apotik.ApproxCount = 0 Then dtpAwal.SetFocus Else dgClosingDataPelayananTM_OA_Apotik.SetFocus
    Exit Sub
End Sub

Private Sub cmdClosingData_Click()
    On Error GoTo errLoad
    If SetJadwalClosingPosting(cmdClosingData) = False Then Exit Sub
    cmdClosingData.Enabled = False
    MousePointer = vbHourglass
    If dgClosingDataPelayananTM_OA_Apotik.ApproxCount = 0 Then Exit Sub

    If sp_ClosingDataPelayanan() = False Then Exit Sub
    pbData.Value = 0
    pbData.Max = rsB.RecordCount
    For i = 1 To dgClosingDataPelayananTM_OA_Apotik.ApproxCount
        dgClosingDataPelayananTM_OA_Apotik.Bookmark = i
        DoEvents
        pbData.Value = pbData.Value + 1
        If optTindakanMedis.Value = True Then
            If sp_Loop_AddDataPelayananPasienTM_OA_ApotikPH(dgClosingDataPelayananTM_OA_Apotik.Columns("NoPendaftaran"), dgClosingDataPelayananTM_OA_Apotik.Columns("TglPelayanan").Value, dgClosingDataPelayananTM_OA_Apotik.Columns("KdItem").Value, dgClosingDataPelayananTM_OA_Apotik.Columns("KdAsal").Value, "", dgClosingDataPelayananTM_OA_Apotik.Columns("NoLab_Rad"), "TM") = False Then Exit Sub
        ElseIf optObatAlKes.Value = True Then
            If sp_Loop_AddDataPelayananPasienTM_OA_ApotikPH(dgClosingDataPelayananTM_OA_Apotik.Columns("NoPendaftaran"), dgClosingDataPelayananTM_OA_Apotik.Columns("TglPelayanan").Value, dgClosingDataPelayananTM_OA_Apotik.Columns("KdItem").Value, dgClosingDataPelayananTM_OA_Apotik.Columns("KdAsal").Value, dgClosingDataPelayananTM_OA_Apotik.Columns("SatuanJml").Value, "", "OA") = False Then Exit Sub
        ElseIf optApotik.Value = True Then
            If sp_Loop_AddDataPelayananPasienTM_OA_ApotikPH(dgClosingDataPelayananTM_OA_Apotik.Columns("NoStruk"), dgClosingDataPelayananTM_OA_Apotik.Columns("TglStruk").Value, dgClosingDataPelayananTM_OA_Apotik.Columns("KdItem").Value, dgClosingDataPelayananTM_OA_Apotik.Columns("KdAsal").Value, dgClosingDataPelayananTM_OA_Apotik.Columns("SatuanJml"), "", "AP") = False Then Exit Sub
        End If

    Next i

    MsgBox "" & frmClosingDataPelayananTM_OA_Apotik.Caption & " Berhasil ", vbInformation, "Informasi"
    pbData.Value = 0
    cmdClosingData.Enabled = True
    MousePointer = vbDefault
    Call cmdCari_Click

    Exit Sub
errLoad:
    msubPesanError
    MousePointer = vbDefault
    cmdClosingData.Enabled = True
    pbData.Value = 0

End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcAsalPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcAsalPasien.MatchedWithList = True Then cmdCari.SetFocus
        strSQL = "SELECT KdRuangan, NamaRuangan FROM Ruangan WHERE NamaRuangan like '%" & dcAsalPasien.Text & "%' and StatusEnabled ='1'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcAsalPasien.Text = ""
            cmdCari.SetFocus
            Exit Sub
        End If
        dcAsalPasien.BoundText = rs(0).Value
        dcAsalPasien.Text = rs(1).Value
    End If

End Sub

Private Sub dcJenisItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcJenisItem.MatchedWithList = True Then cmdCari.SetFocus
        If optTindakanMedis.Value = True Then
            strSQL = "SELECT KdJnsPelayanan, Deskripsi FROM JenisPelayanan WHERE Deskripsi like '%" & dcJenisItem.Text & "%' and StatusEnabled ='1'"
        Else
            strSQL = "SELECT KdDetailJenisBarang, DetailJenisBarang FROM DetailJenisBarang WHERE DetailJenisBarang like '%" & dcJenisItem.Text & "%' and StatusEnabled ='1' "
        End If
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcJenisItem.Text = ""
            cmdCari.SetFocus
            Exit Sub
        End If
        dcJenisItem.BoundText = rs(0).Value
        dcJenisItem.Text = rs(1).Value
    End If

End Sub

Private Sub dcJenisPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcJenisPasien.MatchedWithList = True Then cmdCari.SetFocus
        strSQL = "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien WHERE JenisPasien like '%" & dcJenisPasien.Text & "%' and StatusEnabled ='1'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcJenisPasien.Text = ""
            cmdCari.SetFocus
            Exit Sub
        End If
        dcJenisPasien.BoundText = rs(0).Value
        dcJenisPasien.Text = rs(1).Value
    End If
End Sub

Private Sub dcNamaDokter_Change()
    Call cmdCari_Click
End Sub

Private Sub dcNamaDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcNamaDokter.MatchedWithList = True Then cmdCari.SetFocus
        strSQL = "Select KdJenisPegawai, NamaLengkap from DataPegawai where KdJenisPegawai = '001' and StatusEnabled ='1'and (NamaLengkap LIKE '%" & dcNamaDokter.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcNamaDokter.Text = ""
            cmdCari.SetFocus
            Exit Sub
        End If
        dcNamaDokter.BoundText = rs(0).Value
        dcNamaDokter.Text = rs(1).Value
    End If
End Sub

Private Sub dcKelas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcKelas.MatchedWithList = True Then cmdCari.SetFocus
        strSQL = "SELECT KdKelas, DeskKelas FROM KelasPelayanan WHERE DeskKelas like '%" & dcKelas.Text & "%' and StatusEnabled ='1'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcKelas.Text = ""
            cmdCari.SetFocus
            Exit Sub
        End If
        dcKelas.BoundText = rs(0).Value
        dcKelas.Text = rs(1).Value
    End If
End Sub

Private Sub dcNamaItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcNamaItem.MatchedWithList = True Then cmdCari.SetFocus
        If optTindakanMedis.Value = True Then
            strSQL = "SELECT KdPelayananRS, NamaPelayanan FROM ListPelayananRS WHERE NamaPelayanan like '%" & dcNamaItem.Text & "%' and StatusEnabled ='1'"
        Else
            strSQL = "SELECT KdBarang, NamaBarang From MasterBarang WHERE NamaBarang like '%" & dcNamaItem.Text & "%' and StatusEnabled ='1'"
        End If
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcNamaItem.Text = ""
            cmdCari.SetFocus
            Exit Sub
        End If
        dcNamaItem.BoundText = rs(0).Value
        dcNamaItem.Text = rs(1).Value
    End If

End Sub

Private Sub dgClosingDataPelayananTM_OA_Apotik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdClosingData.SetFocus
End Sub

Private Sub dgClosingDataPelayananTM_OA_Apotik_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    lblJumData.Caption = "Data " & dgClosingDataPelayananTM_OA_Apotik.Bookmark & "/" & dgClosingDataPelayananTM_OA_Apotik.ApproxCount
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Activate()
    cmdCari_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        If dgClosingDataPelayananTM_OA_Apotik.ApproxCount = 0 Then Exit Sub
        Call subLoadGridSource
        mdTglAwal = dtpAwal.Value
        mdTglAkhir = dtpAkhir
        If ValidasiTanggal(dtpAwal, dtpAkhir) = False Then Exit Sub

        frmCetakClosingDataPelayananTMOAApotik.Show
    End If

    If KeyCode = vbKeyShift Then
        If dgClosingDataPelayananTM_OA_Apotik.ApproxCount = 0 Then Exit Sub
        mdTglAwal = dtpAwal.Value
        mdTglAkhir = dtpAkhir
        If ValidasiTanggal(dtpAwal, dtpAkhir) = False Then Exit Sub
        frmCetakDataPelayananTMOAApotik.Show
    End If

    If KeyCode = vbKeyF9 Then
        If dcNamaDokter.Text = "" Then

            strSQL = "SELECT * FROM V_DaftarKunjunganPasienRJ" & _
            " Where Tanggal Between '" & Format(frmDaftarPasienRJ.dtpAwal.Value, "yyyy/MM/dd hh:mm:00") & "' AND '" & Format(frmDaftarPasienRJ.dtpAkhir.Value, "yyyy/MM/dd hh:mm:59") & "' " & _
            " AND JenisPasien like '%" & dcJenisPasien.Text & "%' AND KdRuangan = '" & mstrKdRuangan & "'  "

            Call msubRecFO(rsB, strSQL)

            mdTglAwal = dtpAwal.Value
            mdTglAkhir = dtpAkhir
            If ValidasiTanggal(dtpAwal, dtpAkhir) = False Then Exit Sub
            FrmCetakLapClosing.Show
        Else
            strSQL = "SELECT * FROM V_BiayaPelayananTindakan2 " _
            & "WHERE DokterPemeriksa like '%" & dcNamaDokter.Text & "%' AND TglPelayanan BETWEEN '" & Format(frmDaftarPasienRJ.dtpAwal.Value, "yyyy/MM/dd hh:mm:00") & "' AND '" & Format(frmDaftarPasienRJ.dtpAkhir.Value, "yyyy/MM/dd hh:mm:59") & "' " & _
            "AND JenisPasien like '%" & dcJenisPasien.Text & "%' and KdJnsPelayanan <> '001' AND KdRuangan = '" & mstrKdRuangan & "' "

            Call msubRecFO(rsB, strSQL)

            mdTglAwal = dtpAwal.Value
            mdTglAkhir = dtpAkhir
            If ValidasiTanggal(dtpAwal, dtpAkhir) = False Then Exit Sub
            FrmCetakLapClosing2.Show
        End If
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo errFormLoad
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)

    strSQL = "SELECT KdInstalasi FROM Ruangan WHERE KdRuangan ='" & mstrKdRuangan & "' "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    mstrKdInstalasi = rs(0).Value

    If mstrKdInstalasi = "07" Then
        optApotik.Value = True
        optApotik.Enabled = True
        optObatAlKes.Enabled = False
        optTindakanMedis.Enabled = False
    Else
        optTindakanMedis.Value = True
        optTindakanMedis.Enabled = True
        optObatAlKes.Enabled = True
        optApotik.Enabled = True
    End If

    dtpAwal.Value = Format(Now, "dd MMMM yyyy 00:00:00")
    dtpAkhir.Value = Format(Now, "dd MMMM yyyy 23:59:59")
    DTPTglClosing.Value = Now

    Exit Sub
errFormLoad:
    msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    blnFrmCariPasien = False
End Sub

Private Sub optApotik_Click()
    Call subClearData
    Call cmdCari_Click
End Sub

Private Sub optObatAlKes_Click()
    Call subClearData
    Call cmdCari_Click
End Sub

Private Sub optTindakanMedis_Click()
    Call subClearData
    Call cmdCari_Click
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
End Sub

