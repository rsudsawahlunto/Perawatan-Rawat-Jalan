VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMonitoringPembayaran 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Monitoring Pembayaran Pasien"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14325
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMonitoringPembayaran.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   14325
   Begin VB.Frame fraTindakan 
      Caption         =   "Detail Pelayanan Tindakan Dan Obat Alkes"
      Height          =   5895
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   14055
      Begin VB.CommandButton cmdTutupTindakan 
         Caption         =   "Tutup"
         Height          =   495
         Left            =   12120
         TabIndex        =   19
         Top             =   5280
         Width           =   1815
      End
      Begin TabDlg.SSTab sstDetail 
         Height          =   4695
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   8281
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Detail Pelayanaan Tindakan"
         TabPicture(0)   =   "frmMonitoringPembayaran.frx":0CCA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "dgTindakan"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Detail Pelayanaan Obat Alkes"
         TabPicture(1)   =   "frmMonitoringPembayaran.frx":0CE6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "dgDetaiObatAlkes"
         Tab(1).ControlCount=   1
         Begin MSDataGridLib.DataGrid dgTindakan 
            Height          =   3975
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   13455
            _ExtentX        =   23733
            _ExtentY        =   7011
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
         Begin MSDataGridLib.DataGrid dgDetaiObatAlkes 
            Height          =   3975
            Left            =   -74880
            TabIndex        =   18
            Top             =   480
            Width           =   13455
            _ExtentX        =   23733
            _ExtentY        =   7011
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
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   7200
      Width           =   14295
      Begin VB.CommandButton cmdDetail 
         Caption         =   "Detail Pelayanan"
         Height          =   495
         Left            =   9600
         TabIndex        =   10
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   12240
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
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
      Height          =   6135
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   14295
      Begin VB.Frame Frame4 
         Caption         =   "Status Bayar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2880
         TabIndex        =   11
         Top             =   120
         Width           =   5415
         Begin VB.OptionButton optSudah 
            Caption         =   "Sudah Bayar"
            Height          =   375
            Left            =   3600
            TabIndex        =   13
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optBelum 
            Caption         =   "Belum Bayar"
            Height          =   375
            Left            =   1560
            TabIndex        =   12
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin VB.Frame Frame3 
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
         TabIndex        =   4
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
            TabIndex        =   5
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   6
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   130744323
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   7
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   130744323
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3120
            TabIndex        =   8
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgPembayaranPasien 
         Height          =   5055
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   8916
         _Version        =   393216
         HeadLines       =   2
         RowHeight       =   16
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label lblJumData 
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   2655
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
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
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMonitoringPembayaran.frx":0D02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12495
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   12480
      Picture         =   "frmMonitoringPembayaran.frx":2360
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
End
Attribute VB_Name = "frmMonitoringPembayaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCari_Click()
    On Error GoTo Errload
    If optBelum.Value = True Then
        strSQL = "Select Distinct TglMasuk,NoPendaftaran,NoCm,Title,NamaLengkap,Umur,JenisPasien,NamaPenjamin,NamaRuangan from V_MonitoringPembayaranPasien" & _
        " WHERE KdRuangan = '" & mstrKdRuangan & "' and TglMasuk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' and NoStruk Is Null"
    ElseIf optSudah.Value = True Then
        strSQL = "Select Distinct TglMasuk,NoPendaftaran,NoCm,Title,NamaLengkap,Umur,JenisPasien,NamaPenjamin,NamaRuangan from V_MonitoringPembayaranPasien" & _
        " WHERE KdRuangan = '" & mstrKdRuangan & "' and TglMasuk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' and NoStruk Is Not Null"
    End If

    Call msubRecFO(rsB, strSQL)
    Set dgPembayaranPasien.DataSource = rsB
    With dgPembayaranPasien
        .Columns(0).Width = 1900
        .Columns(0).Caption = "Tgl.Pendaftaran"
        .Columns(1).Width = 1200 'NoPendaftaran
        .Columns(2).Width = 900 'No CM
        .Columns(3).Width = 900 'title
        .Columns(4).Width = 2600 'nama
        .Columns(5).Width = 1750 'Umur
        .Columns(6).Width = 1590 'Nama Penjamin
        .Columns(7).Width = 1590 'TglPendaftaran
        .Columns(8).Width = 2210 'Ruangan
'        .Columns(9).Width = 1590 'TglPulang
    End With

    lblJumData.Caption = "Data 0/" & rsB.RecordCount

    Exit Sub
Errload:
    msubPesanError
    Set rsB = Nothing
End Sub

Private Sub cmdDetail_Click()
On Error GoTo Errload
    
    If optBelum.Value = True Then
        strSQL = "Select TglMasuk,NoPendaftaran,NoCm,Title,NamaLengkap,Umur,JenisPasien,NamaPenjamin,NamaRuangan,NoStruk from V_MonitoringPembayaranPasien" & _
        " WHERE KdRuangan = '" & mstrKdRuangan & "' and TglMasuk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' and NoStruk Is Null"
    ElseIf optSudah.Value = True Then
        strSQL = "Select TglMasuk,NoPendaftaran,NoCm,Title,NamaLengkap,Umur,JenisPasien,NamaPenjamin,NamaRuangan,NoStruk from V_MonitoringPembayaranPasien" & _
        " WHERE KdRuangan = '" & mstrKdRuangan & "' and TglMasuk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' and NoStruk <> 'null'"
    End If
        Call msubRecFO(rsC, strSQL)

If rsC.EOF = True Then
        MsgBox "Data tidak ada", vbInformation, "Informasi"
Else

    fraTindakan.Visible = True
    sstDetail.Tab = 0
    Call subLoadDetailPelayananTindakan
    Call subLoadDetailPelayananObatAlkes
    
    optBelum.Enabled = False
    optSudah.Enabled = False
    cmdDetail.Enabled = False
    cmdTutup.Enabled = False
End If
Exit Sub
Errload:
        Call msubPesanError
End Sub

Private Sub cmdTutupTindakan_Click()
    fraTindakan.Visible = False
    optBelum.Enabled = True
    optSudah.Enabled = True
    cmdDetail.Enabled = True
    cmdTutup.Enabled = True
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo Errload

    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    dtpAwal.Value = Format(Now, "dd MMM yyyy 00:00:00")
    dtpAkhir.Value = Now
    Call cmdCari_Click
    
Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub optBelum_Click()
    Call cmdCari_Click
End Sub

Private Sub optSudah_Click()
    Call cmdCari_Click
End Sub

Private Sub subLoadDetailPelayananTindakan()
    If optBelum.Value = True Then
        strSQL = "SELECT TglPelayanan,JenisPelayanan,NamaPelayanan,NamaRuangan AS [Ruang Pelayanan]," _
        & "Kelas,JenisTarif,CITO,JmlPelayanan as Jml,Total as Tarif,BiayaTotal," _
        & "DokterPemeriksa,[Status Bayar],KdPelayananRS,KdRuangan,Operator FROM V_BiayaPelayananTindakan WHERE " _
        & "NoPendaftaran='" & dgPembayaranPasien.Columns(1).Value & "' and [Status Bayar] = 'Belum DiBayar' ORDER BY TglPelayanan"
    Else
        strSQL = "SELECT TglPelayanan,JenisPelayanan,NamaPelayanan,NamaRuangan AS [Ruang Pelayanan]," _
        & "Kelas,JenisTarif,CITO,JmlPelayanan as Jml,Total as Tarif,BiayaTotal," _
        & "DokterPemeriksa,[Status Bayar],KdPelayananRS,KdRuangan,Operator FROM V_BiayaPelayananTindakan WHERE " _
        & "NoPendaftaran='" & dgPembayaranPasien.Columns(1) & "' and [Status Bayar] = 'Sudah DiBayar' ORDER BY TglPelayanan"
    End If
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        
    If rs.EOF = False Then
        Set dgTindakan.DataSource = rs
        With dgTindakan
            .Columns(0).Width = 1600
            .Columns(1).Width = 2000
            .Columns(2).Width = 2000
            .Columns(3).Width = 1600
            .Columns(4).Width = 900
            .Columns(5).Width = 1000
            .Columns(6).Width = 0
            .Columns(7).Width = 400
            .Columns(7).Alignment = dbgRight
            .Columns(8).Width = 0
            .Columns(8).Alignment = dbgRight
            .Columns(9).Width = 0
            .Columns(9).Alignment = dbgRight
            .Columns(10).Width = 2400
            .Columns(11).Width = 1200
            .Columns(12).Width = 0 'KdPelayananRS
            .Columns(13).Width = 0 'KdRuangan
            .Columns(14).Width = 2000
    
            .Columns(8).NumberFormat = "#,###"
            .Columns(9).NumberFormat = "#,###"
        End With
    End If

End Sub

Private Sub subLoadDetailPelayananObatAlkes()
'    strSQL = "SELECT TglPelayanan,[Detail Jenis Brg],NamaBarang," _
'    & "NamaRuangan AS [Ruang Pelayanan],Kelas,JenisTarif,SatuanJml as Sat," _
'    & "JmlBarang as Jml,HargaSatuan as Tarif,BiayaTotal,DokterPemeriksa," _
'    & "[Status Bayar],KdBarang,KdAsal,Operator, KdRuangan, NoTerima, ResepKe " _
'    & "FROM V_BiayaPemakaianObatAlkesLab WHERE NoLaboratorium='" _
'    & mstrNoLab & "' AND NoPendaftaran ='" & mstrNoPen & "' ORDER BY TglPelayanan"

    If optBelum.Value = True Then
        strSQL = "SELECT TglPelayanan,NamaRuangan AS [Ruang Pelayanan],[Detail Jenis Brg]," _
        & "NamaBarang,NamaAsal as [Asal Barang],Kelas,JmlBarang as Jml,Hargasatuan as Tarif ,BiayaTotal," _
        & "DokterPemeriksa,[Status Bayar],KdRuangan FROM V_BiayaPemakaianObatAlkes WHERE " _
        & "NoPendaftaran='" & dgPembayaranPasien.Columns(1).Value & "' and [Status Bayar] = 'Belum DiBayar' ORDER BY TglPelayanan"
    Else
        strSQL = "SELECT TglPelayanan,NamaRuangan AS [Ruang Pelayanan],[Detail Jenis Brg]," _
        & "NamaBarang,NamaAsal as [Asal Barang],Kelas,JmlBarang as Jml,Hargasatuan as Tarif ,BiayaTotal," _
        & "DokterPemeriksa,[Status Bayar],KdRuangan FROM V_BiayaPemakaianObatAlkes WHERE " _
        & "NoPendaftaran='" & dgPembayaranPasien.Columns(1) & "' and [Status Bayar] = 'Sudah DiBayar' ORDER BY TglPelayanan"
    End If
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        
    If rs.EOF = False Then
        Set dgDetaiObatAlkes.DataSource = rs
        With dgDetaiObatAlkes
            .Columns(0).Width = 1600 'TglPelayanan
            .Columns(1).Width = 2500 'NamaRuangan
            .Columns(2).Width = 2000 'DetailJenisBarang
            .Columns(3).Width = 3000 'NamaBarang
            .Columns(4).Width = 1000 'NamaAsal
            .Columns(4).Alignment = dbgRight
            .Columns(5).Width = 1500 'Kelas
            .Columns(6).Width = 700 'JmlBarang
            .Columns(7).Width = 1000 'Hargasatuan
            .Columns(7).Alignment = dbgRight
            .Columns(8).Width = 2000 'TotalBiaya
            .Columns(8).Alignment = dbgRight
            .Columns(9).Width = 2500 'DokterPemeriksa
            .Columns(9).Alignment = dbgRight
            .Columns(10).Width = 2400 'Status Bayar
            .Columns(11).Width = 0 'KdRuangan

    
            .Columns(8).NumberFormat = "#,###"
            .Columns(9).NumberFormat = "#,###"
        End With
    End If

End Sub
Private Sub sstDetail_Click(PreviousTab As Integer)
    Select Case sstDetail.Tab
        Case 0
            Call subLoadDetailPelayananTindakan
        Case 1
            Call subLoadDetailPelayananObatAlkes

    End Select
End Sub

Private Sub sstDetail_KeyPress(KeyAscii As Integer)
    On Error GoTo Errload

    If KeyAscii = 13 Then
        Select Case sstDetail.Tab
            Case 0
                dgTindakan.SetFocus
            Case 1
                dgDetaiObatAlkes.SetFocus
        End Select
    End If

    Exit Sub
Errload:
    Call msubPesanError
End Sub
