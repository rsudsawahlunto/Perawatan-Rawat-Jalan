VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmDaftarSensusPelayanan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Laporan Sensus Pelayanan"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarSensusPelayanan.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   8910
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
      Height          =   855
      Left            =   0
      TabIndex        =   11
      Top             =   2880
      Width           =   8895
      Begin VB.CommandButton cmdCetak 
         Caption         =   "C&etak"
         Height          =   495
         Left            =   5760
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   7320
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
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
         Height          =   735
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   8655
         Begin VB.CheckBox ChkGroupby 
            Caption         =   "Group By"
            Height          =   330
            Left            =   3240
            TabIndex        =   4
            Top             =   240
            Value           =   1  'Checked
            Width           =   1995
         End
         Begin VB.OptionButton optInstalasiAwal 
            Caption         =   "Instalasi Asal"
            Height          =   375
            Left            =   1560
            TabIndex        =   3
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optJenisPasien 
            Caption         =   "Jenis Pasien"
            Height          =   375
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo DCGroupBy 
            Height          =   360
            Left            =   5280
            TabIndex        =   5
            Top             =   240
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
      Begin VB.CheckBox chkKelasPelayanan 
         Caption         =   "Kelas Pelayanan"
         Height          =   210
         Left            =   120
         TabIndex        =   0
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
         Height          =   735
         Left            =   3480
         TabIndex        =   12
         Top             =   240
         Width           =   5295
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   240
            TabIndex        =   6
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
            Format          =   63307779
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   2880
            TabIndex        =   7
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
            Format          =   63307779
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   2520
            TabIndex        =   13
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataListLib.DataCombo dcKelasPelayanan 
         Height          =   360
         Left            =   120
         TabIndex        =   1
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
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   15
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
      Picture         =   "frmDaftarSensusPelayanan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarSensusPelayanan.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarSensusPelayanan.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmDaftarSensusPelayanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ChkGroupby_Click()
    Call subloaddcgrouby
    If ChkGroupby.Value = vbChecked Then DcGroupBy.Enabled = True Else DcGroupBy.Enabled = False
End Sub

Private Sub ChkGroupby_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Call subloaddcgrouby
    If ChkGroupby.Value = vbChecked Then
    DcGroupBy.Enabled = True
    DcGroupBy.SetFocus
    Else: DcGroupBy.Enabled = False
    cmdCetak.SetFocus
End If
End If
End Sub

Private Sub chkKelasPelayanan_Click()
    If chkKelasPelayanan.Value = vbChecked Then dcKelasPelayanan.Enabled = True Else dcKelasPelayanan.Enabled = False
End Sub

Private Sub chkKelasPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If chkKelasPelayanan.Value = vbChecked Then dcKelasPelayanan.SetFocus Else optJenisPasien.SetFocus
End Sub

Private Sub cmdCetak_Click()
On Error GoTo hell
If optInstalasiAwal.Value = False And optJenisPasien.Value = False Then Exit Sub

 Call getData
    Call sub_KriteriaLaporan
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value
    If ValidasiTanggal(dtpAwal, dtpAkhir) = False Then Exit Sub
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then MsgBox "Tidak Ada Data", vbExclamation, "Validasi": Exit Sub
    Set frmCetakSensusPelayanan = Nothing
    frmCetakSensusPelayanan.Show
Exit Sub
hell:
    msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub DcGroupBy_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call subloaddcgrouby
cmdCetak.SetFocus
End If
End Sub

Private Sub dcKelasPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAwal.SetFocus
End Sub

Private Sub dgData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetak.SetFocus
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then optJenisPasien.SetFocus
End Sub

Private Sub dtpAkhir_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then optJenisPasien.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

'Private Sub Form_Activate()
'ChkGroupby_Click
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error GoTo errFormLoad
    Call centerForm(Me, MDIUtama)
    
    dtpAwal.Value = Format(Now, "dd MMMM yyyy 00:00:00")
    dtpAkhir.Value = Now
    
    If ChkGroupby.Value = Checked Then
    DcGroupBy.Enabled = True
        Call subloaddcgrouby
    Else
        DcGroupBy.Enabled = False
    End If
    
    Call subLoadDC
    Call subloaddcgrouby
    Call getData
'    Call cmdTampilkanTemp_Click
    Call PlayFlashMovie(Me)
Exit Sub
errFormLoad:
    msubPesanError
End Sub

Private Sub subLoadDC()
On Error GoTo errload
 Set rs = Nothing
    strSQL = "SELECT KdDetailJenisJasaPelayanan,DetailJenisJasaPelayanan FROM DetailJenisJasaPelayanan"
    Call msubDcSource(dcKelasPelayanan, rs, strSQL)
    If Not rs.EOF Then dcKelasPelayanan.BoundText = rs(0).Value
Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub subloaddcgrouby()
On Error GoTo hell
Set rs = Nothing
    If optJenisPasien.Value = True Then
        strSQL = "SELECT KdKelompokPasien,JenisPasien FROM KelompokPasien order by JenisPasien"
    ElseIf optInstalasiAwal.Value = True Then
         strSQL = "select * from v_InstalasiRujukan "
    End If
    Call msubDcSource(DcGroupBy, rs, strSQL)
      If Not rs.EOF Then DcGroupBy.BoundText = rs(0).Value
'    Call getDataExit Sub
Exit Sub
hell:
    Call msubPesanError
End Sub

Public Sub getData()
On Error GoTo errorLoad
Dim strFilter As String

    mstrFilter = ""
    If ChkGroupby.Value = vbChecked Then
        If optJenisPasien.Value = True Then
            If chkKelasPelayanan.Value = vbChecked Then
                If Periksa("datacombo", dcKelasPelayanan, "Kelas Pelayanan kosong") = False Then Exit Sub
                mstrFilter = mstrFilter & " And KdJenisKelas = '" & dcKelasPelayanan.BoundText & "'And KdKelompokPasien= '" & DcGroupBy.BoundText & "' "
            End If
        ElseIf optInstalasiAwal.Value = True Then
            If chkKelasPelayanan.Value = vbChecked Then
                mstrFilter = mstrFilter & " And KdJenisKelas = '" & dcKelasPelayanan.BoundText & "'And InstalasiAsal= '" & DcGroupBy.Text & "' "
            End If
        End If
    Else
        mstrFilter = mstrFilter & " And KdJenisKelas = '" & dcKelasPelayanan.BoundText & "'"
        DcGroupBy.Text = ""
    End If

    Call sub_KriteriaLaporan
    Call msubRecFO(rs, strSQL)
Exit Sub
errorLoad:
    Call msubPesanError
End Sub

Private Sub sub_KriteriaLaporan()
On Error GoTo errload
    Set rs = Nothing
    If optJenisPasien.Value = True Then
        strSQL = " SELECT JenisKelas, JenisPasien,Penjamin,RuanganAsal, NamaPelayanan, RuanganPelayanan, SUM(JmlPelayanan) AS Jumlah, KomponenTarif, SUM(TotalBiaya) AS TotalBayar " & _
            " FROM V_SensusPelayananStrukAllTMRekap " & _
            " WHERE (KdRuanganPelayanan = '" & mstrKdRuangan & "') AND TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "'" & mstrFilter & " " & _
            " GROUP BY JenisKelas, JenisPasien,Penjamin,RuanganAsal, NamaPelayanan, RuanganPelayanan, KomponenTarif"
    ElseIf optInstalasiAwal.Value = True Then
        strSQL = " SELECT JenisKelas, InstalasiAsal,Penjamin,RuanganAsal, NamaPelayanan, RuanganPelayanan, SUM(JmlPelayanan) AS Jumlah, KomponenTarif, SUM(TotalBiaya) AS TotalBayar " & _
            " FROM V_SensusPelayananStrukAllTMRekap " & _
            " WHERE ( KdRuanganPelayanan = '" & mstrKdRuangan & "') AND TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "'" & mstrFilter & " " & _
            " GROUP BY JenisKelas, InstalasiAsal,Penjamin,RuanganAsal, NamaPelayanan, RuanganPelayanan, KomponenTarif"
    End If
Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs = Nothing
End Sub

Private Sub optInstalasiAwal_Click()
    ChkGroupby.Caption = "Instalasi Asal"
    Call subloaddcgrouby
End Sub

Private Sub optInstalasiAwal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Call subloaddcgrouby
    dtpAwal.SetFocus
    End If
End Sub

Private Sub optJenisPasien_Click()
    ChkGroupby.Caption = "Jenis Pasien"
    Call subloaddcgrouby
End Sub

Private Sub optJenisPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Call subloaddcgrouby
    ChkGroupby.SetFocus
    End If
End Sub

