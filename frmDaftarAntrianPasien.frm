VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDaftarAntrianPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Daftar Antrian Pasien"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14280
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarAntrianPasien.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   14280
   Begin VB.Frame fraCari 
      Caption         =   "Cari Data Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   0
      TabIndex        =   10
      Top             =   7200
      Width           =   14295
      Begin VB.CommandButton Command1 
         Caption         =   "Panggil Antrian"
         Height          =   450
         Left            =   5640
         TabIndex        =   18
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdHapusRegistrasi 
         Caption         =   "&Hapus Data"
         Height          =   450
         Left            =   5640
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1560
         TabIndex        =   5
         Top             =   400
         Width           =   2895
      End
      Begin VB.CommandButton cmdTP 
         Caption         =   "&Batal DiPeriksa"
         Height          =   450
         Left            =   9960
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdPasienDirujuk 
         Caption         =   "&Masuk Poliklinik"
         Height          =   450
         Left            =   7800
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   450
         Left            =   12120
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan Nama Pasien /  No.CM"
         Height          =   240
         Index           =   0
         Left            =   1560
         TabIndex        =   11
         Top             =   150
         Width           =   2820
      End
   End
   Begin VB.Frame fraDaftar 
      Caption         =   "Daftar Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   0
      TabIndex        =   12
      Top             =   960
      Width           =   14295
      Begin VB.Frame Frame1 
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
         Left            =   8400
         TabIndex        =   13
         Top             =   150
         Width           =   5775
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   0
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   184025091
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   1
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   184090627
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   14
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgDaftarAntrianPasien 
         Height          =   5175
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   9128
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
      Begin MSDataListLib.DataCombo dcStatusPeriksa 
         Height          =   360
         Left            =   6480
         TabIndex        =   3
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin VB.Label lblJumData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data 0/0"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Status Periksa"
         Height          =   240
         Index           =   1
         Left            =   6480
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   17
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
      Left            =   12480
      Picture         =   "frmDaftarAntrianPasien.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarAntrianPasien.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarAntrianPasien.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "frmDaftarAntrianPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub cmdCari_Click()
    On Error GoTo errLoad
   ' If dcStatusPeriksa.Text = "Y" Then
    If dtpAwal.Day <> dtpAkhir.Day Or dtpAwal.Month <> dtpAkhir.Month Or dtpAwal.Year <> dtpAkhir.Year Then
        strSQL = "select TOP 100 * " & _
        " from V_DaftarAntrianPasienMRS " & _
        " where ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%') and [status periksa] = '" & dcStatusPeriksa.Text & "' and Ruangan='" & strNNamaRuangan & "' and TglMasuk between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "'"
    Else
        strSQL = "select TOP 100 * " & _
        " from V_DaftarAntrianPasienMRS " & _
        " where ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%') and [status periksa] = '" & dcStatusPeriksa.Text & "' and Ruangan='" & strNNamaRuangan & "' and TglMasuk between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "'"
        
    'End If
    
    End If
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    Set dgDaftarAntrianPasien.DataSource = rs
    Call SetGridAntrianPasien

    lblJumData.Caption = "Data 0/" & rs.RecordCount

    Exit Sub
errLoad:
    msubPesanError
    Set rs = Nothing
End Sub

Private Sub cmdCari_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dgDaftarAntrianPasien.SetFocus
End Sub

Private Sub cmdHapusRegistrasi_Click()
    On Error GoTo errLoad

    If dgDaftarAntrianPasien.ApproxCount = 0 Then Exit Sub
    If MsgBox("Anda yakin akan menghapus data registrasi pasien", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, dgDaftarAntrianPasien.Columns("No. Registrasi"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 10, dgDaftarAntrianPasien.Columns("KdRuangan"))
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , dgDaftarAntrianPasien.Columns("TglMasuk"))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_DeleteRegistrasiPasienMRS"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Add_DeleteRegistrasiPasienMRS")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Call cmdCari_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdPasienDirujuk_Click()
    On Error GoTo errLoad
    If dgDaftarAntrianPasien.ApproxCount < 1 Then Exit Sub

    strSQL = "SELECT  RegistrasiRJ.NoPendaftaran, Ruangan.NamaRuangan " & _
    " FROM   RegistrasiRJ INNER JOIN Ruangan ON  RegistrasiRJ.KdRuangan = Ruangan.KdRuangan " & _
    " WHERE  RegistrasiRJ.NoPendaftaran = '" & dgDaftarAntrianPasien.Columns(1) & "' AND IdDokter IS NOT NULL"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        MsgBox "Pasien tersebut sudah terdaftar di " & rs(1) & "", vbExclamation, "Validasi"
        Exit Sub
    End If

    'validasi status periksa
    If dgDaftarAntrianPasien.Columns(17) = "" Then Exit Sub 'status periksa
    If LCase(dgDaftarAntrianPasien.Columns(17)) = "sedang" Then
        MsgBox "Pasien sedang dalam proses", vbExclamation, "Validasi"
        Exit Sub
    ElseIf LCase(dgDaftarAntrianPasien.Columns(17)) = "sudah" Then
        MsgBox "Pasien sudah selesai diproses", vbExclamation, "Validasi"
        Exit Sub
    End If

    mstrNoPen = dgDaftarAntrianPasien.Columns(1) 'no pendaftaran
    mstrKdJenisPasien = dgDaftarAntrianPasien.Columns("KdKelompokPasien") '
    mstrKdSubInstalasi = dgDaftarAntrianPasien.Columns("KdSubInstalasi")
    frmRegistrasiRJ.txtnopendaftaran = dgDaftarAntrianPasien.Columns(1)
    frmRegistrasiRJ.subTampilData (dgDaftarAntrianPasien.Columns(1))
    frmRegistrasiRJ.subNoAntrian (dgDaftarAntrianPasien.Columns(0))
    frmRegistrasiRJ.Show
    frmDaftarAntrianPasien.Enabled = False
     dbConn.Execute ("update PasienMasukRumahSakit set StatusPeriksa='S' where NoPendaftaran='" + dgDaftarAntrianPasien.Columns(1) + "'")
'    Dim Path As String
'       strSQL = "select Value from SettingGlobal where Prefix='PathSdkAntrian'"
'    Call msubRecFO(rs, strSQL)
'    Dim pathtemp As String
'    If Not rs.EOF Then
'        If rs(0).Value <> "" Then
'            pathtemp = rs(0).Value
'        End If
'    End If
'
'    strSQL = "select StatusAntrian from SettingDataUmum"
'    Call msubRecFO(rs, strSQL)
'    Dim coba As Long
'
'    If Not rs.EOF Then
'        If rs(0).Value = "1" Then
'
'        If Dir(Path) <> "" Then
'                strSQL = "select * from settingglobal where Prefix like 'KdRuanganAntrian%' and Value='" & mstrKdRuangan & "'"
'                Dim prefix As String
'                prefix = ""
'                Call msubRecFO(rs, strSQL)
'                If (rs.EOF = False) Then
'                    If (rs("Prefix").Value = "KdRuanganAntrianKanan") Then
'                         prefix = "Kanan-"
'                    ElseIf (rs("Prefix").Value = "KdRuanganAntrianKiri") Then
'                        prefix = "Kiri-"
'                    Else
'                    End If
'                End If
'            Path = pathtemp + " Type:" & Chr(34) & "Counting Patient" & Chr(34) & " NoAntrian:" & prefix & dgDaftarAntrianPasien.Columns(0).Value & " loket:" & mstrKdRuangan
'            coba = Shell(Path, vbNormalFocus)
'
'            Path = pathtemp + "  Type:" & Chr(34) & "Update Patient" & Chr(34) & " loket:" & mstrKdRuangan
'            coba = Shell(Path, vbNormalFocus)
'            Call cmdCari_Click
'        End If
'        End If
'    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTP_Click()
    On Error GoTo errLoad
    If dgDaftarAntrianPasien.ApproxCount = 0 Then Exit Sub
    If LCase(dgDaftarAntrianPasien.Columns(17)) <> "belum" Then
        MsgBox "Hanya pasien yang belum registrasi dapat dibatalkan", vbExclamation, "Validasi"
        Exit Sub
    End If
    Set rs = Nothing
    strSQL = "SELECT NoPendaftaran,NoStruk,KdPelayananRS FROM BiayaPelayanan WHERE NoPendaftaran ='" & dgDaftarAntrianPasien.Columns("No. Registrasi") & "' AND KdRuangan ='" & dgDaftarAntrianPasien.Columns("KdRuangan") & "'  AND TglPelayanan ='" & Format(dgDaftarAntrianPasien.Columns("TglMasuk"), "yyyy/MM/dd HH:mm:ss") & "'" 'AND NoStruk is null "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        KdPelayananRSBatalPeriksa = rs("KdPelayananRS").Value
        NoStrukBatalPeriksa = rs("NoStruk").Value & ""
        If NoStrukBatalPeriksa = "" Then
            bolStatusDelPelayanan = True
            GoTo BATAL_
        Else
            strSQL = "SELECT NoStruk from ReturStrukPembayaranPasien WHERE NoStruk ='" & NoStrukBatalPeriksa & "'"
            Call msubRecFO(rs, strSQL)
            If rs.EOF = False Then
                GoTo BATAL_
            Else
                'rehab medis gk bayar
                If mstrKdRuangan = "601" Or mstrKdRuangan = "222" Then GoTo BATAL_
                MsgBox "Pasien " & dgDaftarAntrianPasien.Columns("Nama Pasien").Value & " Sudah Bayar - Silahkan Retur pembayaran pada kasir ", vbExclamation, "Validasi"
                Exit Sub
            End If
        End If
    End If

BATAL_:
    frmDaftarAntrianPasien.Enabled = False
    With frmBatalDirawat
        .Show
        .txtnocm.Text = dgDaftarAntrianPasien.Columns(2).Value
        .txtNamaPasien.Text = dgDaftarAntrianPasien.Columns("Nama Pasien").Value
        If dgDaftarAntrianPasien.Columns("JK").Value = "P" Then
            .txtJK.Text = "Perempuan"
        Else
            .txtJK.Text = "Laki-Laki"
        End If
        .txtThn.Text = dgDaftarAntrianPasien.Columns("UmurTahun").Value
        .txtBln.Text = dgDaftarAntrianPasien.Columns("UmurBulan").Value
        .txtHr.Text = dgDaftarAntrianPasien.Columns("UmurHari").Value
        .txtnopendaftaran.Text = dgDaftarAntrianPasien.Columns(1).Value
        .txtDokterLama.Text = ""
        .txtRuanganLama.Text = dgDaftarAntrianPasien.Columns("Ruangan").Value
        .txtKdRuangan.Text = dgDaftarAntrianPasien.Columns("KdRuangan").Value
        .txtKdSubInstalasi.Text = dgDaftarAntrianPasien.Columns("KdSubInstalasi").Value
        .dtpTglMasuk.Value = dgDaftarAntrianPasien.Columns("TglMasuk").Value
    End With

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdtutup_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    On Error GoTo errLoad
    
    Dim Path As String
    Dim pathtemp As String
    Dim coba As Long
    Dim prefix As String
    
    strSQL = "select Value from SettingGlobal where Prefix='PathSdkAntrian'"
    
    Call msubRecFO(rs, strSQL)
    
    If Not rs.EOF Then
        If rs(0).Value <> "" Then
            pathtemp = rs(0).Value
        End If
    End If
    
    strSQL = "select StatusAntrian from SettingDataUmum"
    Call msubRecFO(rs, strSQL)
    
    If Not rs.EOF Then
        If rs(0).Value = "1" Then
            If Dir(Path) <> "" Then
                strSQL = "select * from settingglobal where Prefix like 'KdRuanganAntrian%' and Value='" & mstrKdRuangan & "'"
                prefix = ""
                Call msubRecFO(rs, strSQL)
                If (rs.EOF = False) Then
                    If (rs("Prefix").Value = "KdRuanganAntrianKanan") Then
                        prefix = "Kanan-"
                    ElseIf (rs("Prefix").Value = "KdRuanganAntrianKiri") Then
                        prefix = "Kiri-"
                    Else
                    
                    End If
                End If
                Path = pathtemp + " Type:" & Chr(34) & "Counting Patient" & Chr(34) & " NoAntrian:" & prefix & dgDaftarAntrianPasien.Columns(0).Value & " loket:" & mstrKdRuangan
                coba = Shell(Path, vbNormalFocus)
                
                Path = pathtemp + "  Type:" & Chr(34) & "Update Patient" & Chr(34) & " loket:" & mstrKdRuangan
                coba = Shell(Path, vbNormalFocus)
'                Call cmdCari_Click
            End If
        End If
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcStatusPeriksa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcStatusPeriksa.MatchedWithList = True Then dtpAwal.SetFocus
        strSQL = "Select kdstatusperiksa, statusperiksa From StatusPeriksaPasien Where StatusEnabled='1' and (statusperiksa LIKE '%" & dcStatusPeriksa.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcStatusPeriksa.Text = ""
            dtpAwal.SetFocus
            Exit Sub
        End If
        dcStatusPeriksa.BoundText = rs(0).Value
        dcStatusPeriksa.Text = rs(1).Value
    End If
End Sub

Private Sub dgDaftarAntrianPasien_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDaftarAntrianPasien
    WheelHook.WheelHook dgDaftarAntrianPasien
End Sub

Private Sub dgDaftarAntrianPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdPasienDirujuk.SetFocus
End Sub

Private Sub dgDaftarAntrianPasien_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    lblJumData.Caption = "Data " & dgDaftarAntrianPasien.Bookmark & "/" & dgDaftarAntrianPasien.ApproxCount
End Sub

Private Sub dtpAkhir_Change()
    On Error Resume Next
    'dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpAwal_Change()
    On Error Resume Next
    'dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    mblnFormDaftarAntrian = True
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    dtpAwal.Value = Format(Now, "dd MMM yyyy 00:00:00")
    dtpAkhir.Value = Now
    dcStatusPeriksa.BoundText = ""

    If mblnAdmin = True Then
        cmdHapusRegistrasi.Enabled = True
    Else
        cmdHapusRegistrasi.Enabled = False
    End If

    Call subLoadDcSource
    Call cmdCari_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadDcSource()
    On Error GoTo errLoad

    strSQL = "Select * From StatusPeriksaPasien where statusEnabled='1'"
    Call msubDcSource(dcStatusPeriksa, rs, strSQL)
    If rs.EOF = False Then dcStatusPeriksa.BoundText = rs(0).Value

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub SetGridAntrianPasien()
    On Error Resume Next
    With dgDaftarAntrianPasien
        .Columns(0).Width = 800
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 1150
        .Columns(1).Caption = "No. Registrasi"
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 750
        .Columns(2).Caption = "No. CM"
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Width = 2600
        .Columns(4).Width = 300
        .Columns(4).Alignment = dbgCenter
        .Columns(5).Width = 1500
        .Columns(6).Width = 1590
        .Columns(7).Width = 1700
        .Columns(7).Alignment = dbgCenter
        .Columns(8).Width = 0
        .Columns(8).Alignment = dbgCenter
        .Columns(9).Width = 0
        .Columns(10).Width = 0
        .Columns(11).Width = 0
        .Columns(12).Width = 0
        .Columns(13).Width = 0
        .Columns(14).Width = 0
        .Columns(15).Width = 0
        .Columns(16).Width = 1700
        .Columns(17).Width = 1300
        .Columns(18).Width = 3000
        .Columns(19).Width = 0
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnFormDaftarAntrian = False
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        Call cmdCari_Click
        txtParameter.SetFocus
    End If
End Sub

