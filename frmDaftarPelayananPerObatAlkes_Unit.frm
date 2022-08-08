VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmDaftarPelayananPerObatAlkes_Unit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Laporan Pelayanan Obat Alkes-->Unit"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPelayananPerObatAlkes_Unit.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   8925
   Begin VB.CommandButton cmdCetak 
      Caption         =   "C&etak"
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   7200
      TabIndex        =   9
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Frame fraPeriode 
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
      Top             =   960
      Width           =   8895
      Begin VB.CheckBox ChkPilihan 
         Caption         =   "Group By"
         Height          =   210
         Left            =   5520
         TabIndex        =   6
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2025
      End
      Begin VB.Frame Frame2 
         Caption         =   "Group By"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   5055
         Begin VB.OptionButton optInstalasiAwal 
            Caption         =   "Instalasi Asal"
            Height          =   375
            Left            =   2760
            TabIndex        =   5
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optJenisPasien 
            Caption         =   "Jenis Pasien"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.CheckBox chkKelasPelayanan 
         Caption         =   "Jenis Kelas"
         Height          =   210
         Left            =   5520
         TabIndex        =   2
         Top             =   240
         Value           =   1  'Checked
         Width           =   2025
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
         Height          =   855
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   5055
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   120
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
            Format          =   22675459
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   2760
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
            Format          =   22675459
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   2400
            TabIndex        =   12
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataListLib.DataCombo dcKelasPelayanan 
         Height          =   360
         Left            =   5520
         TabIndex        =   3
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
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
      Begin MSDataListLib.DataCombo DcGroupBy 
         Height          =   360
         Left            =   5520
         TabIndex        =   7
         Top             =   1320
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
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
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   14
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
      Left            =   7080
      Picture         =   "frmDaftarPelayananPerObatAlkes_Unit.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarPelayananPerObatAlkes_Unit.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarPelayananPerObatAlkes_Unit.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmDaftarPelayananPerObatAlkes_Unit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkKelasPelayanan_Click()
    If chkKelasPelayanan.Value = Checked Then
        dcKelasPelayanan.Enabled = True
    Else
        dcKelasPelayanan.Text = ""
        dcKelasPelayanan.Enabled = False
    End If
End Sub

Private Sub chkKelasPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If chkKelasPelayanan.Value = vbChecked Then dcKelasPelayanan.SetFocus Else optJenisPasien.SetFocus
End Sub

Private Sub ChkPilihan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If ChkPilihan.Value = Checked Then DcGroupBy.SetFocus Else cmdCetak.SetFocus
End Sub

Private Sub cmdCetak_Click()
On Error GoTo hell
If optJenisPasien.Value = False And optInstalasiAwal.Value = False Then MsgBox "Pilih Dulu group By nya", vbExclamation, "Validasi": Exit Sub
mdTglAwal = dtpAwal.Value
mdTglAkhir = dtpAkhir.Value
If ValidasiTanggal(dtpAwal, dtpAkhir) = False Then Exit Sub
Call getData
Set rs = Nothing
Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mdTglAwal = dtpAwal.Value
        mdTglAkhir = dtpAkhir.Value
        Set frmCetakPelayananPerObatAlkes_Ruangan = Nothing
        frmCetakPelayananPerObatAlkes_Ruangan.Show
    Else
        MsgBox "Data Tidak Ada", vbInformation, "Medifirst2000-Validasi"
    End If
    
Set rs = Nothing
Exit Sub
hell:
    msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub ChkPilihan_Click()
    If ChkPilihan.Value = Checked Then
        DcGroupBy.Enabled = True
        If optJenisPasien.Value = True Then
            ChkPilihan.Caption = "Jenis Pasien"
        ElseIf optInstalasiAwal.Value = True Then
            ChkPilihan.Caption = "Instalasi Asal"
        End If
    ElseIf ChkPilihan.Value = Unchecked Then
        DcGroupBy.Enabled = False
        If optJenisPasien.Value = True Then
            ChkPilihan.Caption = "Jenis Pasien ALL"
        ElseIf optInstalasiAwal.Value = True Then
            ChkPilihan.Caption = "Instalasi Asal ALL"
        End If
    End If
End Sub

Private Sub DcGroupBy_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetak.SetFocus
End Sub

Private Sub dcKelasPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then optJenisPasien.SetFocus
End Sub

Private Sub dgData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetak.SetFocus
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then chkKelasPelayanan.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error GoTo errFormLoad
    Call centerForm(Me, MDIUtama)
    
    dtpAwal.Value = Format(Now, "dd MMMM yyyy 00:00:00")
    dtpAkhir.Value = Format(Now, "dd MMMM yyyy 23:59:59")
    
    If optJenisPasien.Value = True Then
    Call subloadDCGroupby
    End If
    If optInstalasiAwal.Value = True Then
    Call subloadDCGroupby
    End If
    
    Call subLoadDC
    Call subloadDCGroupby
    Call getData
    Call PlayFlashMovie(Me)
    
Exit Sub
errFormLoad:
    msubPesanError
End Sub
Private Sub subloadDCGroupby()
Set rs = Nothing
    If optJenisPasien.Value = True Then
        strSQL = "SELECT KdKelompokPasien,JenisPasien FROM KelompokPasien order by JenisPasien"
        Call msubDcSource(DcGroupBy, rs, strSQL)
            If Not rs.EOF Then DcGroupBy.BoundText = rs(0).Value
    ElseIf optInstalasiAwal.Value = True Then
        strSQL = " SELECT     KdInstalasi, NamaInstalasi From dbo.instalasi" & _
                 " WHERE     (NOT (KdInstalasi IN ('05', '06', '07', '13', '14', '15', '17', '18', '19', '20', '21', '22')))"
        Call msubDcSource(DcGroupBy, rs, strSQL)
            If Not rs.EOF Then DcGroupBy.BoundText = rs(0).Value
End If
Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub subLoadDC()
On Error GoTo errload
    
    strSQL = "SELECT KdDetailJenisJasaPelayanan,DetailJenisJasaPelayanan FROM DetailJenisJasaPelayanan"
    Call msubDcSource(dcKelasPelayanan, rs, strSQL)
    If Not rs.EOF Then dcKelasPelayanan.BoundText = rs(0).Value
    
Exit Sub
errload:
    Call msubPesanError
End Sub

Public Sub getData()
On Error GoTo errorLoad
Dim strFilter As String
mstrFilter = ""

    If chkKelasPelayanan.Value = vbChecked Then
        If Periksa("datacombo", dcKelasPelayanan, "Kelas Pelayanan kosong") = False Then Exit Sub
        mstrFilter = mstrFilter & " And KdJenisKelas = '" & dcKelasPelayanan.BoundText & "'"
    End If
    
    If ChkPilihan.Value = vbChecked Then
        If optJenisPasien.Value = True Then
        mstrFilter = mstrFilter & " And KdKelompokPasien= '" & DcGroupBy.BoundText & "' "
        ElseIf optInstalasiAwal.Value = True Then
        mstrFilter = mstrFilter & " And KdInstalasiAsal = '" & DcGroupBy.BoundText & "'"
        End If
    End If
    
    Call sub_KriteriaLaporan
    Call msubRecFO(rs, strSQL)
    
Exit Sub
errorLoad:
    Call msubPesanError
End Sub

Private Sub optInstalasiAwal_Click()
Call subloadDCGroupby
    If ChkPilihan.Value = Unchecked Then
    ChkPilihan.Caption = "Instalasi Asal ALL"
    Else
    ChkPilihan.Caption = "Instalasi Asal"
    End If
End Sub

Private Sub sub_KriteriaLaporan()
On Error GoTo errload
    
    If optJenisPasien.Value = True Then
        strSQL = " SELECT JenisKelas, JenisPasien,Penjamin, RuanganAsal, KomponenUnit, KomponenTarif, SUM(TotalBiaya) AS TotalBayar " & _
            " FROM V_SensusPelayananStrukAllOARekap " & _
            " WHERE TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd hh:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd hh:mm:59") & "'" & mstrFilter & " " & _
            " and RuanganPelayanan = '" & mstrNamaRuangan & "' GROUP BY JenisKelas, JenisPasien,Penjamin, RuanganAsal,  KomponenUnit, KomponenTarif"
    ElseIf optInstalasiAwal.Value = True Then
        strSQL = " SELECT JenisKelas, InstalasiAsal,Penjamin, RuanganAsal, KomponenUnit, KomponenTarif, SUM(TotalBiaya) AS TotalBayar " & _
            " FROM V_SensusPelayananStrukAllOARekap " & _
            " WHERE TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd hh:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd hh:mm:59") & "'" & mstrFilter & " " & _
            " and RuanganPelayanan = '" & mstrNamaRuangan & "' GROUP BY JenisKelas, InstalasiAsal,Penjamin, RuanganAsal,  KomponenUnit, KomponenTarif"
    End If

Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub optInstalasiAwal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call subloadDCGroupby
        ChkPilihan.SetFocus
    End If
End Sub

Private Sub optJenisPasien_Click()
Call subloadDCGroupby
    If ChkPilihan.Value = Unchecked Then
       ChkPilihan.Caption = "Jenis Pasien ALL"
    Else
       ChkPilihan.Caption = "Jenis Pasien"
    End If
End Sub

Private Sub optJenisPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call subloadDCGroupby
       ChkPilihan.SetFocus
    End If
End Sub
