VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTindakanMedisPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Tindakan Medis Pasien"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11940
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTindakanMedisPasien.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   11940
   Begin VB.TextBox txtNoHasilPeriksa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   0
      MaxLength       =   10
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Hasil Tindakan Medis"
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
      TabIndex        =   10
      Top             =   2640
      Width           =   11895
      Begin VB.CheckBox chkStatusTindakan 
         Caption         =   "Dilakukan di RS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   44
         ToolTipText     =   "cek jika tindakan dilakukan di rumah sakit /tidak di cek jika tindakan di lakukan di luar rumah sakit"
         Top             =   360
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo dcJnsTindakanMedis 
         Height          =   330
         Left            =   2160
         TabIndex        =   0
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
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
      Begin MSComCtl2.DTPicker dtpTglAwalPeriksa 
         Height          =   330
         Left            =   2160
         TabIndex        =   1
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   128122883
         UpDown          =   -1  'True
         CurrentDate     =   38077
      End
      Begin MSComCtl2.DTPicker dtpTglAkhirPeriksa 
         Height          =   330
         Left            =   4560
         TabIndex        =   2
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   128122883
         UpDown          =   -1  'True
         CurrentDate     =   38077
      End
      Begin MSDataListLib.DataCombo dcDokterKepala 
         Height          =   330
         Left            =   8640
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
      Begin MSDataListLib.DataCombo dcDokterDelegasi 
         Height          =   330
         Left            =   8640
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
      Begin MSDataListLib.DataCombo dcKepalaParamedis 
         Height          =   330
         Left            =   8640
         TabIndex        =   6
         Top             =   960
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
      Begin MSComCtl2.DTPicker dtpTglHasilPeriksa 
         Height          =   330
         Left            =   2160
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   128122883
         UpDown          =   -1  'True
         CurrentDate     =   38077
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Hasil Periksa"
         Height          =   210
         Left            =   240
         TabIndex        =   27
         Top             =   1140
         Width           =   1275
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Kepala Paramedis"
         Height          =   210
         Left            =   6720
         TabIndex        =   16
         Top             =   1020
         Width           =   1395
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Dokter Delegasi"
         Height          =   210
         Left            =   6720
         TabIndex        =   15
         Top             =   660
         Width           =   1275
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Dokter Kepala"
         Height          =   210
         Left            =   6720
         TabIndex        =   14
         Top             =   300
         Width           =   1140
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "s/d"
         Height          =   210
         Left            =   4200
         TabIndex        =   13
         Top             =   780
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Mulai Periksa"
         Height          =   210
         Left            =   240
         TabIndex        =   12
         Top             =   780
         Width           =   1380
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Tindakan Medis"
         Height          =   210
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1695
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
      TabIndex        =   20
      Top             =   1080
      Width           =   11895
      Begin VB.TextBox txtSubInstalasi 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   8640
         TabIndex        =   37
         Top             =   1080
         Width           =   3015
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
         TabIndex        =   33
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
         TabIndex        =   32
         Top             =   720
         Width           =   375
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
         TabIndex        =   31
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   8640
         TabIndex        =   28
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         MaxLength       =   12
         TabIndex        =   25
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   22
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Kasus Penyakit"
         Height          =   210
         Left            =   6720
         TabIndex        =   38
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "thn"
         Height          =   210
         Left            =   9075
         TabIndex        =   36
         Top             =   750
         Width           =   285
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "bln"
         Height          =   210
         Left            =   9870
         TabIndex        =   35
         Top             =   750
         Width           =   240
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "hr"
         Height          =   210
         Left            =   10650
         TabIndex        =   34
         Top             =   750
         Width           =   165
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Umur"
         Height          =   210
         Left            =   6720
         TabIndex        =   30
         Top             =   720
         Width           =   435
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   6720
         TabIndex        =   29
         Top             =   360
         Width           =   1065
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   465
      Left            =   6600
      TabIndex        =   19
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   465
      Left            =   10200
      TabIndex        =   18
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   465
      Left            =   8400
      TabIndex        =   8
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Detail Tindakan Medis"
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
      TabIndex        =   17
      Top             =   4200
      Width           =   11895
      Begin VB.TextBox txtIsi 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   4440
         TabIndex        =   42
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo dcKualitasHasil 
         Height          =   330
         Left            =   2640
         TabIndex        =   41
         Top             =   1200
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcTindakanOperasi 
         Height          =   330
         Left            =   2640
         TabIndex        =   40
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcPelayananRS 
         Height          =   330
         Left            =   2640
         TabIndex        =   39
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   2415
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   4260
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   -2147483643
         FocusRect       =   2
         HighLight       =   2
         Appearance      =   0
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
      Left            =   10080
      Picture         =   "frmTindakanMedisPasien.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmTindakanMedisPasien.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmTindakanMedisPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub cmdBatal_Click()
    Call subKosong
    Call subSetGrid
    Call subLoadDcSource
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell
    Dim i As Integer
    
    If Periksa("datacombo", dcJnsTindakanMedis, "Jenis Tindakan Medis tidak boleh kosong") = False Then Exit Sub
    
    If Periksa("datacombo", dcDokterKepala, "Dokter kepala tidak boleh kosong") = False Then Exit Sub
    
    If (fgData.TextMatrix(1, 0) = "") And (fgData.TextMatrix(1, 1) = "") And (fgData.TextMatrix(1, 7) = "") Then
        MsgBox "Detail tindakan Operasi/anastesi masih kosong", vbCritical, "Validation"
        fgData.SetFocus
        Exit Sub
    End If
    With fgData
        For i = 1 To .Rows - 1
            If (.TextMatrix(i, 0) = "") Or (.TextMatrix(i, 1) = "") Or (.TextMatrix(i, 7) = "") Then
               MsgBox "Detail tindakan Operasi Harus lengkap", vbCritical, "Validation"
               Exit Sub
            End If
        Next i
'        Exit For
    End With

    If SP_AUDHasilTindakanMedis("A") = False Then Exit Sub
    With fgData
        For i = 1 To .Rows - 1
            If (.TextMatrix(i, 0) <> "") Or (.TextMatrix(i, 1) <> "") Or (.TextMatrix(i, 7) <> "") Then
                If SP_AUDDetailHasilTindakanMedisPasien(.TextMatrix(i, 5), .TextMatrix(i, 6), .TextMatrix(i, 7), .TextMatrix(i, 3), .TextMatrix(i, 4), "A") = False Then Exit Sub
            End If
        Next i
    End With
    MsgBox "Data berhasil disimpan", vbInformation, "Sukses"
    Call cmdBatal_Click
    cmdSimpan.Enabled = False
    frmTransaksiPasien.subLoadRiwayatOperasi

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Function SP_AUDDetailHasilTindakanMedisPasien(f_KdPelayananRS As String, f_KdTindakanMedis As String, f_KdKualitasHasil As Integer, f_HasilPeriksa As String, f_MemoHasilPeriksa As String, f_status As String) As Boolean
    On Error GoTo hell
    SP_AUDDetailHasilTindakanMedisPasien = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoHasilPeriksa", adVarChar, adParamInput, 10, Trim(txtNoHasilPeriksa.Text))
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, f_KdPelayananRS)
        .Parameters.Append .CreateParameter("KdTindakanMedis", adVarChar, adParamInput, 4, f_KdTindakanMedis)
        .Parameters.Append .CreateParameter("KdKualitasHasil", adTinyInt, adParamInput, , IIf(Trim(f_KdKualitasHasil) = "", Null, CInt(Trim(f_KdKualitasHasil))))
        .Parameters.Append .CreateParameter("HasilPeriksa", adVarChar, adParamInput, 100, IIf(Trim(f_HasilPeriksa) = "", Null, Trim(f_HasilPeriksa)))
        .Parameters.Append .CreateParameter("MemoHasilPeriksa", adVarChar, adParamInput, 200, IIf(Trim(f_MemoHasilPeriksa) = "", Null, Trim(f_MemoHasilPeriksa)))
        .Parameters.Append .CreateParameter("StatusOnSiteService", adTinyInt, adParamInput, , IIf(chkStatusTindakan.Value = vbChecked, 1, 0))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_DetailHasilTindakanMedisPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Error - Ada kesalahan dalam penyimpanan data detail, Hubungi administrator", vbCritical, "Error"
            SP_AUDDetailHasilTindakanMedisPasien = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function
hell:
    SP_AUDDetailHasilTindakanMedisPasien = False
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Call msubPesanError
End Function

Private Function SP_AUDHasilTindakanMedis(f_status As String) As Boolean
    On Error GoTo hell
    SP_AUDHasilTindakanMedis = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoHasilPeriksa", adVarChar, adParamInput, 10, IIf(Trim(txtNoHasilPeriksa.Text) = "", Null, Trim(txtNoHasilPeriksa.Text)))
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("NoLab_Rad", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("NoPakai", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("KdKelompokUmur", adChar, adParamInput, 2, Null)
        .Parameters.Append .CreateParameter("JenisKelamin", adChar, adParamInput, 1, Left(txtSex.Text, 1))
        .Parameters.Append .CreateParameter("KdJenisTindakanMedis", adVarChar, adParamInput, 2, IIf(Trim(dcJnsTindakanMedis.Text) = "", Null, dcJnsTindakanMedis.BoundText))
        .Parameters.Append .CreateParameter("KdKeadaanLahirBayi", adTinyInt, adParamInput, , Null)
        .Parameters.Append .CreateParameter("ParitasKe", adTinyInt, adParamInput, , Null)
        .Parameters.Append .CreateParameter("TglMulaiPeriksa", adDate, adParamInput, , Format(dtpTglAwalPeriksa.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglAkhirPeriksa", adDate, adParamInput, , Format(dtpTglAkhirPeriksa.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglHasilPeriksa", adDate, adParamInput, , Format(dtpTglHasilPeriksa.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdKamar", adChar, adParamInput, 4, Null)
        .Parameters.Append .CreateParameter("IdDokterKepala", adChar, adParamInput, 10, IIf(Trim(dcDokterKepala.Text) = "", Null, dcDokterKepala.BoundText))
        .Parameters.Append .CreateParameter("IdDokterOperator1", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("IdDokterOperator2", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("IdDokterAnastesi", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("IdDokterPendamping", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("IdDokterDelegasi", adChar, adParamInput, 10, IIf(Trim(dcDokterDelegasi.Text) = "", Null, dcDokterDelegasi.BoundText))
        .Parameters.Append .CreateParameter("IdParamedisKepala", adChar, adParamInput, 10, IIf(Trim(dcKepalaParamedis.Text) = "", Null, dcKepalaParamedis.BoundText))
        .Parameters.Append .CreateParameter("KetHasilPeriksa", adVarChar, adParamInput, 150, Null)
        .Parameters.Append .CreateParameter("TglMulaiAnastesi", adDate, adParamInput, , Null)
        .Parameters.Append .CreateParameter("TglSelesaiAnatesi", adDate, adParamInput, , Null)
        .Parameters.Append .CreateParameter("KdJenisAnastesi", adTinyInt, adParamInput, , Null)
        .Parameters.Append .CreateParameter("IdAnastesiKepala", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("KetHasilPeriksa", adVarChar, adParamInput, 150, Null)
        .Parameters.Append .CreateParameter("NoHasilPeriksaOutput", adVarChar, adParamOutput, 10, Null)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_HasilTindakanMedis"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Error - Ada kesalahan dalam penyimpanan data, Hubungi administrator", vbCritical, "Error"
            SP_AUDHasilTindakanMedis = False
        Else
            txtNoHasilPeriksa.Text = Trim(.Parameters("NoHasilPeriksaOutput").Value)
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
hell:
    SP_AUDHasilTindakanMedis = False
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Call msubPesanError
End Function

Private Sub dcDokterDelegasi_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcDokterDelegasi.MatchedWithList = True Then dcKepalaParamedis.SetFocus
        strSQL = "Select IdPegawai,NamaLengkap From V_DataPegawai WHERE (NamaLengkap LIKE '%" & dcDokterDelegasi.Text & "%') and KdJenisPegawai='001'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcDokterDelegasi.BoundText = rs(0).Value
        dcDokterDelegasi.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcDokterDelegasi_LostFocus()
    If dcDokterDelegasi.MatchedWithList = False Then dcDokterDelegasi.Text = "": Exit Sub

End Sub

Private Sub dcDokterKepala_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcDokterKepala.MatchedWithList = True Then dcDokterDelegasi.SetFocus
        strSQL = "Select IdPegawai,NamaLengkap From V_DataPegawai WHERE (NamaLengkap LIKE '%" & dcDokterKepala.Text & "%') and KdJenisPegawai='001'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcDokterKepala.BoundText = rs(0).Value
        dcDokterKepala.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcDokterKepala_LostFocus()
        strSQL = "Select IdPegawai,NamaLengkap From V_DataPegawai WHERE (NamaLengkap LIKE '%" & dcDokterKepala.Text & "%') and KdJenisPegawai='001'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then dcDokterKepala.Text = "": Exit Sub
        dcDokterKepala.BoundText = rs(0).Value
        dcDokterKepala.Text = rs(1).Value

End Sub

Private Sub dcJnsTindakanMedis_Change()
    dcTindakanOperasi.Text = ""
   For i = 1 To fgData.Rows - 1
    fgData.TextMatrix(i, 1) = ""
    fgData.TextMatrix(i, 6) = ""
    
    fgData.TextMatrix(i, 2) = ""
    fgData.TextMatrix(i, 7) = ""
        
   Next i

End Sub

Private Sub dcJnsTindakanMedis_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then
        If dcJnsTindakanMedis.MatchedWithList = True Then dtpTglAwalPeriksa.SetFocus
        strSQL = "SELECT JenisTindakanMedis.KdJenisTindakanMedis, JenisTindakanMedis.JenisTindakanMedis " & _
        "FROM JenisTindakanMedis INNER JOIN MapJenisTindakanMedisToRuangan ON JenisTindakanMedis.KdJenisTindakanMedis = MapJenisTindakanMedisToRuangan.KdJenisTindakanMedis " & _
        "WHERE (JenisTindakanMedis.JenisTindakanMedis LIKE '%" & dcJnsTindakanMedis.Text & "%') AND (MapJenisTindakanMedisToRuangan.KdRuangan = '" & mstrKdRuangan & "') AND (JenisTindakanMedis.StatusEnabled = 1)"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcJnsTindakanMedis.BoundText = rs(0).Value
        dcJnsTindakanMedis.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcJnsTindakanMedis_LostFocus()
    If dcJnsTindakanMedis.MatchedWithList = False Then dcJnsTindakanMedis.Text = "": Exit Sub

End Sub

Private Sub dcKepalaParamedis_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then
        If dcKepalaParamedis.MatchedWithList = True Then fgData.SetFocus
        strSQL = "Select IdPegawai,NamaLengkap From V_DataPegawai WHERE (NamaLengkap LIKE '%" & dcKepalaParamedis.Text & "%') and KdJenisPegawai<>'001'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcKepalaParamedis.BoundText = rs(0).Value
        dcKepalaParamedis.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub


Private Sub dcKepalaParamedis_LostFocus()
    If dcKepalaParamedis.MatchedWithList = False Then dcKepalaParamedis.Text = "": Exit Sub
End Sub

Private Sub dcKualitasHasil_Change()
    If dcKualitasHasil.Text = "" Then fgData.TextMatrix(fgData.Row, 2) = "": fgData.TextMatrix(fgData.Row, 7) = "": Exit Sub
    fgData.TextMatrix(fgData.Row, 2) = dcKualitasHasil.Text
    fgData.TextMatrix(fgData.Row, 7) = dcKualitasHasil.BoundText
End Sub

Private Sub dcKualitasHasil_GotFocus()
    Call msubDcSource(dcKualitasHasil, rs, "Select KdKualitasHasil,KualitasHasil From V_TindakanMedisToHasil WHERE TindakanMedis='" & fgData.TextMatrix(fgData.Row, 1) & "'")
End Sub

Private Sub dcKualitasHasil_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then dcKualitasHasil.Visible = False: fgData.SetFocus
End Sub

Private Sub dcKualitasHasil_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Or KeyAscii = 27 Then
        If dcKualitasHasil.MatchedWithList = True Then fgData.Col = 3: fgData.SetFocus
        strSQL = "Select KdKualitasHasil,KualitasHasil From KualitasHasil Where (KualitasHasil LIKE '%" & dcKualitasHasil.Text & "%')"
        Call msubRecFO(dbRst, strSQL)
        If dbRst.EOF = True Then dcKualitasHasil.Text = "": Exit Sub
        dcKualitasHasil.BoundText = dbRst(0).Value
        dcKualitasHasil.Text = dbRst(1).Value
        Call dcKualitasHasil_Change
        dcKualitasHasil.Visible = False
        fgData.Col = 2
        fgData.SetFocus
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcKualitasHasil_LostFocus()
    If dcKualitasHasil.MatchedWithList = False Then dcKualitasHasil.Text = "": dcKualitasHasil.BoundText = "": dcKualitasHasil.Visible = False: Exit Sub
    dcKualitasHasil.Visible = False
End Sub

Private Sub dcTindakanOperasi_Change()
    If dcTindakanOperasi.BoundText = "" Then Exit Sub
    dcKualitasHasil.Text = ""
    fgData.TextMatrix(fgData.Row, 1) = dcTindakanOperasi.Text
    fgData.TextMatrix(fgData.Row, 6) = dcTindakanOperasi.BoundText
End Sub

Private Sub dcTindakanOperasi_GotFocus()
    Call msubDcSource(dcTindakanOperasi, rs, "SELECT Distinct KdTindakanMedis,TindakanMedis From V_JenisTindakanBersalinx WHERE KdJenisTindakanMedis = '" & dcJnsTindakanMedis.BoundText & "' and KdRuangan='" & mstrKdRuangan & "' ")
'     strSQL = "Select KdTindakanMedis,TindakanMedis From TindakanMedis Where KdJenisTindakanMedis like '%" & dcJnsTindakanMedis.BoundText & "%' And KdInstalasi='" & mstrKdInstalasiLogin & "' And StatusEnabled=1" ' And (TindakanMedis LIKE '%" & dcTindakanOperasi.Text & "%')
'     Call msubDcSource(dcTindakanOperasi, dbRst, strSQL)
End Sub

Private Sub dcTindakanOperasi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then dcTindakanOperasi.Visible = False: fgData.SetFocus
End Sub

Private Sub dcTindakanOperasi_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Or KeyAscii = 27 Then
        If dcTindakanOperasi.MatchedWithList = True Then fgData.Col = 2: fgData.SetFocus
        strSQL = "Select KdTindakanMedis,TindakanMedis From TindakanMedis Where KdTindakanMedis like '%" & dcJnsTindakanMedis.BoundText & "%' And (TindakanMedis LIKE '%" & dcTindakanOperasi.Text & "%') And KdInstalasi='" & mstrKdInstalasiLogin & "' And StatusEnabled=1"
        Call msubRecFO(dbRst, strSQL)
        If dbRst.EOF = True Then Exit Sub
        dcTindakanOperasi.BoundText = dbRst(0).Value
        dcTindakanOperasi.Text = dbRst(1).Value
        Call dcTindakanOperasi_Change
        dcTindakanOperasi.Visible = False
        fgData.Col = 2
        fgData.SetFocus
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcTindakanOperasi_LostFocus()
    If dcTindakanOperasi.MatchedWithList = False Then dcTindakanOperasi.Text = "": dcTindakanOperasi.Visible = False: Exit Sub
    
    dcTindakanOperasi.Visible = False
End Sub

Private Sub dtpTglAkhirPeriksa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpTglHasilPeriksa.SetFocus
End Sub

Private Sub dtpTglAwalPeriksa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpTglAkhirPeriksa.SetFocus
End Sub

Private Sub dtpTglHasilPeriksa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcDokterKepala.SetFocus
End Sub

Private Sub fgData_DblClick()
'    Call fgData_KeyDown(13, 0)
End Sub

Private Sub dcPelayananRS_Change()
    If dcPelayananRS.BoundText = "" Then Exit Sub
    fgData.TextMatrix(fgData.Row, 0) = dcPelayananRS.Text
    fgData.TextMatrix(fgData.Row, 5) = dcPelayananRS.BoundText
End Sub

Private Sub dcPelayananRS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then dcPelayananRS.Visible = False: fgData.SetFocus
End Sub

Private Sub dcPelayananRS_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Or KeyAscii = 27 Then
        If dcPelayananRS.MatchedWithList = True Then fgData.Col = 1: fgData.SetFocus
'        V_BiayaPelayananTindakan
'        strSQL = "Select KdPelayananRS,NamaPelayanan From ListPelayananRS Where (NamaPelayanan LIKE '%" & dcPelayananRS.Text & "%') And StatusEnabled=1"
        strSQL = "Select Distinct KdPelayananRS,NamaPelayanan From V_BiayaPelayananTindakan Where (NamaPelayanan LIKE '%" & dcPelayananRS.Text & "%') and NoPendaftaran='" & mstrNoPen & "'"
        Call msubRecFO(dbRst, strSQL)
        If dbRst.EOF = True Then dcPelayananRS = "": dcPelayananRS.Visible = True: Exit Sub
        dcPelayananRS.BoundText = dbRst(0).Value
        dcPelayananRS.Text = dbRst(1).Value
        Call dcPelayananRS_Change
        dcPelayananRS.Visible = False
        fgData.Col = 1
        fgData.SetFocus
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcPelayananRS_LostFocus()
    If dcPelayananRS.MatchedWithList = False Then dcPelayananRS.Text = "": dcPelayananRS.Visible = False: fgData.TextMatrix(fgData.Row, 0) = "": fgData.TextMatrix(fgData.Row, 5) = "": Exit Sub
    dcPelayananRS.Visible = False
End Sub

Private Sub fgData_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo hell
    Dim i As Integer
    Dim tempKode As String

    Select Case KeyCode
        Case 13
            If fgData.Col = fgData.Cols - 4 Then
                If fgData.TextMatrix(fgData.Row, 5) <> "" Then
                    If fgData.TextMatrix(fgData.Rows - 1, 0) <> "" Then
                        fgData.Rows = fgData.Rows + 1
                    End If
                    fgData.Row = fgData.Rows - 1
                    fgData.Col = 0
                Else
                    fgData.Col = 0
                End If
            Else
                For i = 0 To fgData.Cols - 3
                    If fgData.Col = fgData.Cols - 1 Then Exit For
                    fgData.Col = fgData.Col + 1
                    If fgData.ColWidth(fgData.Col) > 0 Then Exit For
                Next i
            End If
            fgData.SetFocus
            If fgData.Col = 1 Then
                tempKode = fgData.TextMatrix(fgData.Row, 6)
                dcTindakanOperasi.BoundText = ""
                dcTindakanOperasi.BoundText = tempKode
                Call subLoadDataCombo(dcTindakanOperasi)
            End If
            If fgData.Col = 2 Then
                tempKode = fgData.TextMatrix(fgData.Row, 7)
                dcKualitasHasil.BoundText = ""
                dcKualitasHasil.BoundText = tempKode
                Call subLoadDataCombo(dcKualitasHasil)
            End If
        Case 27
            txtIsi.Visible = False
            dcPelayananRS.Visible = False
            dcTindakanOperasi.Visible = False
            dcKualitasHasil.Visible = False
        Case vbKeyDelete
            With fgData
                If .Row = .Rows Then Exit Sub
                If .Row = 0 Then Exit Sub

                If .Rows = 2 Then
                    For i = 0 To .Cols - 1
                        .TextMatrix(1, i) = ""
                    Next i
                Else
                    .RemoveItem .Row
                End If
            End With
    End Select
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub fgData_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    txtIsi.Text = ""
    If Not (KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ Or KeyAscii = 32 Or KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace Or KeyAscii = Asc(".")) Then
        KeyAscii = 0
        Exit Sub
    End If

    Select Case fgData.Col
        Case 0 'Nama Pelayanan
            fgData.Col = 0
            Call subLoadDataCombo(dcPelayananRS)
        Case 1 'Nama Tindakan Operasi
            If dcJnsTindakanMedis.Text = "" Then MsgBox "Jenis Tindakan Medis Masih Kosong.", vbCritical, "Validasi": dcJnsTindakanMedis.SetFocus: Exit Sub
                fgData.Col = 1
                Call subLoadDataCombo(dcTindakanOperasi)
                        
        Case 2 'Kualitas Hasil
            If fgData.TextMatrix(fgData.Rows - 1, 1) = "" Then MsgBox "Nama Tindakan Medis Masih Kosong.", vbCritical, "Validasi": fgData.SetFocus: fgData.Col = 1: Exit Sub
            fgData.Col = 2
            Call subLoadDataCombo(dcKualitasHasil)
        Case 3
            Call SubLoadText
            txtIsi.Text = Chr(KeyAscii)
            txtIsi.SelStart = Len(txtIsi.Text)
        Case 4
            Call SubLoadText
            txtIsi.Text = Chr(KeyAscii)
            txtIsi.SelStart = Len(txtIsi.Text)
    End Select

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call cmdBatal_Click
End Sub

Sub subKosong()
    On Error GoTo hell
    dcJnsTindakanMedis.Text = ""
    dtpTglAwalPeriksa.Value = Now
    dtpTglAkhirPeriksa.Value = Now
    dtpTglHasilPeriksa.Value = Now
    dcDokterKepala.Text = ""
    dcDokterDelegasi.Text = ""
    dcKepalaParamedis.Text = ""
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

        .TextMatrix(0, 0) = "Nama Pelayanan"
        .TextMatrix(0, 1) = "Nama Tindakan Medis"
        .TextMatrix(0, 2) = "Kualitas Hasil"
        .TextMatrix(0, 3) = "Hasil Pemeriksaan"
        .TextMatrix(0, 4) = "Memo Hasil Pemeriksaan"
        .TextMatrix(0, 5) = "KdPelayananRS"
        .TextMatrix(0, 6) = "KdTindakanMedis"
        .TextMatrix(0, 7) = "KdKualitasHasil"

        .ColWidth(0) = 2500
        .ColWidth(1) = 2200
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
        .ColWidth(4) = 2900
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub subLoadDcSource()
    On Error GoTo hell

    strSQL = "SELECT JenisTindakanMedis.KdJenisTindakanMedis, JenisTindakanMedis.JenisTindakanMedis " & _
    "FROM JenisTindakanMedis INNER JOIN MapJenisTindakanMedisToRuangan ON JenisTindakanMedis.KdJenisTindakanMedis = MapJenisTindakanMedisToRuangan.KdJenisTindakanMedis " & _
    "WHERE (MapJenisTindakanMedisToRuangan.KdRuangan = '" & mstrKdRuangan & "') AND (JenisTindakanMedis.StatusEnabled = 1)"

    Call msubDcSource(dcJnsTindakanMedis, rs, strSQL)
    strSQL = "Select IdPegawai,NamaLengkap From V_DataPegawai WHERE  KdJenisPegawai='001'"
    Call msubDcSource(dcDokterKepala, rs, strSQL)
    Call msubDcSource(dcDokterDelegasi, rs, strSQL)
    strSQL = "Select IdPegawai,NamaLengkap From V_DataPegawai WHERE  KdJenisPegawai<>'001'"
    Call msubDcSource(dcKepalaParamedis, rs, strSQL)

'    Call msubDcSource(dcPelayananRS, rs, "Select KdPelayananRS,NamaPelayanan From ListPelayananRS Where StatusEnabled=1")
'menampilkan pelayanan yang di dapat
    Call msubDcSource(dcPelayananRS, rs, "Select Distinct KdPelayananRS,NamaPelayanan From V_BiayaPelayananTindakan Where NoPendaftaran='" & mstrNoPen & "'") 'AND NoLab_Rad='" & mstrNoIBS & "'  ORDER BY TglPelayanan

'    Call msubDcSource(dcTindakanOperasi, rs, "Select KdTindakanMedis,TindakanMedis From TindakanMedis Where KdInstalasi='" & mstrKdInstalasiLogin & "' And StatusEnabled=1")
'    Call msubDcSource(dcKualitasHasil, rs, "Select KdKualitasHasil,KualitasHasil From KualitasHasil")
    Exit Sub
hell:
    Call msubPesanError
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

Private Sub SubLoadText()
    Dim i As Integer
    txtIsi.Left = fgData.Left

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

Private Sub Form_Unload(Cancel As Integer)
    frmTransaksiPasien.Enabled = True
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    With fgData
        If KeyAscii = 13 Then
            Select Case .Col
                Case 3
                    .TextMatrix(.Row, .Col) = txtIsi.Text
                    .SetFocus: .Col = 4
                Case 4
                    .TextMatrix(.Row, .Col) = txtIsi.Text
                    If fgData.RowPos(fgData.Row) >= fgData.Height - 360 Then
                        fgData.SetFocus
                        SendKeys "{DOWN}"
                        Exit Sub
                    End If
                    fgData.SetFocus
            End Select

        ElseIf KeyAscii = 27 Then
            txtIsi.Visible = False
        End If
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub txtIsi_LostFocus()
    txtIsi.Visible = False
End Sub

