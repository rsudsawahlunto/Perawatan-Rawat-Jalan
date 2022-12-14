VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDaftarAntrianRegistrasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Antrian Pasien"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarAntrianRegistrasi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   10455
   Begin VB.Frame fraCari 
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
      TabIndex        =   6
      Top             =   7200
      Width           =   10335
      Begin VB.CommandButton cmdPanggilAntrian 
         Caption         =   "Panggil &Antrian"
         Height          =   450
         Left            =   6120
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   450
         Left            =   8160
         TabIndex        =   5
         Top             =   240
         Width           =   1935
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
      TabIndex        =   7
      Top             =   960
      Width           =   10335
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   4920
         Top             =   2400
      End
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
         Left            =   4320
         TabIndex        =   8
         Top             =   120
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
            Format          =   123142147
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
            Format          =   123076611
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   9
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgDaftarAntrianPasien 
         Height          =   5175
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   9975
         _ExtentX        =   17595
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
      Begin MSDataListLib.DataCombo cbLoket 
         Height          =   360
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loket"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lblJumData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data 0/0"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   11
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
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarAntrianRegistrasi.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarAntrianRegistrasi.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "frmDaftarAntrianRegistrasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbLoket_Change()
    Call cmdCari_Click
End Sub

Public Sub cmdCari_Click()
    On Error GoTo Errload
    
    If dtpAwal.Day <> dtpAkhir.Day Or dtpAwal.Month <> dtpAkhir.Month Or dtpAwal.Year <> dtpAkhir.Year Then
        strSQL = "select TOP 100 * " & _
        " from V_AntrianPasienRegistrasi " & _
        " where ( NoPendaftaran = '0000000000' or NoPendaftaran is null) and TglAntrian between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "' and JenisPasien='" & cbLoket.Text & "'"
    Else
        strSQL = "select TOP 100 * " & _
        " from V_AntrianPasienRegistrasi " & _
        " where ( NoPendaftaran = '0000000000' or NoPendaftaran is null)  and  TglAntrian between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'  and JenisPasien='" & cbLoket.Text & "' and KdInstalasi<>'07'"
    End If
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    Set dgDaftarAntrianPasien.DataSource = rs
    Call SetGridAntrianPasien

    lblJumData.Caption = "Data 0/" & rs.RecordCount

    Exit Sub
Errload:
    msubPesanError
    Set rs = Nothing
End Sub

Private Sub cmdcari_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dgDaftarAntrianPasien.SetFocus
End Sub

Private Sub cmdPanggilAntrian_Click()
 On Error GoTo Errload
  If dgDaftarAntrianPasien.ApproxCount <> 0 Then
  
   strSQL = "select NoCounter from Ruangan where KdRuangan='" & mstrKdRuangan & "'"
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    Dim NoCounter As String
    If Not rs.EOF Then
        If rs(0).value <> "" Then
            NoCounter = rs(0).value
        End If
    End If

    dbConn.Execute ("update AntrianPasienRegistrasi set NoPendaftaran='0000000000',NoLoketCounter='" + NoCounter + "',iduser='" + strIDPegawaiAktif + "' where kdAntrian='" + dgDaftarAntrianPasien.Columns(1).value + "'")
    tempkodeAntrian = dgDaftarAntrianPasien.Columns(1).value
    Set rs = Nothing
    Dim path As String
    
    strNoAntrian = dgDaftarAntrianPasien.Columns(0).value
    
    strSQL = "select Value from SettingGlobal where Prefix='PathSdkAntrian'"
    Call msubRecFO(rs, strSQL)
      
    If Not rs.EOF Then
        If rs(0).value <> "" Then
            path = rs(0).value
        End If
    End If
    
    strSQL = "select StatusAntrian from SettingDataUmum"
    Call msubRecFO(rs, strSQL)
    Dim coba As Long
    If Not rs.EOF Then
        If rs(0).value = "1" Then
            If Dir(path) <> "" Then
                path = path + " Type:" & Chr(34) & "Counting Patient" & Chr(34) & " NoAntrian:" & dgDaftarAntrianPasien.Columns(0).value
                coba = Shell(path, vbNormalFocus)
                path = path + " Type:" & Chr(34) & "Update Patient" & Chr(34)
                coba = Shell(path, vbNormalFocus)
            End If
        End If
    End If
    
   
       bolStatusFrmAntrian = True
       MstrKdRuanganAntrian = dgDaftarAntrianPasien.Columns("KdRuangan").value
       strSQL = "select * from ruangan where kdRuangan='" & MstrKdRuanganAntrian & "'"
       Call msubRecFO(rs, strSQL)
       MstrKdIstalasiAntrian = rs("KdInstalasi").value
'        Dim MstrNamaJenisAntrian As String
'        strSQL = "SELECT KDDetailJenisJasaPelayanan,DetailJenisJasaPelayanan,Kelas,NamaRuangan,NamaInstalasi  FROM V_KelasPelayanan " _
'                 & " Where KdInstalasi='" & dgDaftarAntrianPasien.Columns("KdInstalasi").value & "' and " _
'                 & " KdRuangan='" & dgDaftarAntrianPasien.Columns("KdRuangan").value & "' and " _
'                 & " StatusEnabled ='1' And Expr1 ='1' and Expr2 ='1' and Expr3 ='1'"
'        Call msubRecFO(rs, strSQL)
'        If rs.EOF = False Then
'            MstrKdIstalasiAntrian = rs("NamaInstalasi").value
'            MstrJenisKelasAntrian = rs("KDDetailJenisJasaPelayanan").value
'            MstrNamaJenisAntrian = rs("DetailJenisJasaPelayanan").value
'            MstrKelasAntrian = rs("Kelas").value
'            MstrKdRuanganAntrian = rs("NamaRuangan").value ' dgDaftarAntrianPasien.Columns("KdRuangan").value 'rs("NamaRuangan").value
'        End If

        If dgDaftarAntrianPasien.Columns(3).value = "Pasien" Then
            With frmPasienBaru
                strPasien = "Baru"
                .Show
                .txtKdAntrian = dgDaftarAntrianPasien.Columns(0).value
                bolAntrian = False
                boltampil = True
            End With
        Else
            With frmRegistrasiAll
                .Show
                .txtFormPengirim.Text = Me.Name
                .txtKdAntrian = dgDaftarAntrianPasien.Columns(0).value
                .txtNoCM = dgDaftarAntrianPasien.Columns(3).value
                .CariData
                
'                .dcInstalasi.Text = MstrKdIstalasiAntrian
'                .dcJenisKelas.Text = MstrNamaJenisAntrian
'                .dcJenisKelas.BoundText = MstrJenisKelasAntrian
'                .dcKelas.BoundText = MstrKelasAntrian
                Call .dcRuangan_GotFocus
                .dcRuangan.BoundText = dgDaftarAntrianPasien.Columns("KdRuangan").value
                Call .dcRuangan_LostFocus
                
            End With
        End If
    
    
 
    Call cmdCari_Click
    'Me.Enabled = False
    
    End If
    Exit Sub
Errload:
    msubPesanError
    Set rs = Nothing
   
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgDaftarAntrianPasien_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDaftarAntrianPasien
    WheelHook.WheelHook dgDaftarAntrianPasien
End Sub


Private Sub dgDaftarAntrianPasien_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    lblJumData.Caption = "Data " & dgDaftarAntrianPasien.Bookmark & "/" & dgDaftarAntrianPasien.ApproxCount
End Sub

Private Sub dtpAkhir_Change()
    On Error Resume Next
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpAwal_Change()
    On Error Resume Next
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo Errload
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    dtpAwal.value = Format(Now, "dd MMM yyyy 00:00:00")
    dtpAkhir.value = Format(Now, "dd MMM yyyy 23:59:59") 'Now
    Call cmdCari_Click
    Call msubDcSource(cbLoket, rs, "SELECT NamaLoket,NamaLoket FROM dbo.AntrianLoket where kdruangan='" & mstrKdRuangan & "' ")
    Exit Sub
Errload:
    Call msubPesanError
End Sub

Sub SetGridAntrianPasien()
    On Error Resume Next
    With dgDaftarAntrianPasien
        .Columns(0).Caption = "No Antrian"
        .Columns(0).Width = 1050
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 0
        .Columns(2).Width = 2600
        .Columns(2).Caption = "Tgl Antrian"
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Width = 1500
        .Columns(3).Caption = "No.CM"
        .Columns(3).Alignment = dbgCenter
        .Columns(4).Width = 0
        .Columns(5).Width = 1500
        .Columns(5).Alignment = dbgCenter
        .Columns(6).Width = 0
        .Columns(7).Width = 0
        .Columns(8).Width = 0
        .Columns("Kdinstalasi").Width = 0
        .Columns("Namainstalasi").Width = 0
        .Columns("StatusPasien").Width = 0
        .Columns("NoPendaftaran").Width = 0
        .Columns("IdDokter").Width = 0

    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnFormDaftarAntrian = False
End Sub

Private Sub Timer1_Timer()
'dtpAkhir.value = Now
'dtpAkhir.Refresh
End Sub
