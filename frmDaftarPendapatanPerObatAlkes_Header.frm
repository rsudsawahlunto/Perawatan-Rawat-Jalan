VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmDaftarPendapatanPerObatAlkes_Header 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Laporan Pendapatan Obat Alkes-->Unit"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPendapatanPerObatAlkes_Header.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   8445
   Begin VB.CommandButton cmdCetak 
      Caption         =   "C&etak"
      Height          =   495
      Left            =   5280
      TabIndex        =   14
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   6840
      TabIndex        =   13
      Top             =   3120
      Width           =   1335
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
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   2160
      Width           =   8415
      Begin VB.OptionButton optSisaTagihan 
         Caption         =   "Sisa Tagihan"
         Height          =   495
         Left            =   6840
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optTanggunganRS 
         Caption         =   "Tanggungan RS"
         Height          =   495
         Left            =   5160
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optPembebasan 
         Caption         =   "Pembebasan"
         Height          =   495
         Left            =   3480
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optHutangPenjamin 
         Caption         =   "Hutang Penjamin"
         Height          =   495
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optTotalBayar 
         Caption         =   "Total Bayar"
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
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
      Height          =   1215
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   8415
      Begin VB.CheckBox chkjenisKelas 
         Caption         =   "Jenis Kelas"
         Height          =   210
         Left            =   240
         TabIndex        =   5
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
         Left            =   3120
         TabIndex        =   3
         Top             =   240
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
            Format          =   62914563
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
            Format          =   62914563
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   2400
            TabIndex        =   4
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataListLib.DataCombo dcJenisKelas 
         Height          =   360
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   2415
         _ExtentX        =   4260
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
      Left            =   6600
      Picture         =   "frmDaftarPendapatanPerObatAlkes_Header.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarPendapatanPerObatAlkes_Header.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarPendapatanPerObatAlkes_Header.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmDaftarPendapatanPerObatAlkes_Header"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkjenisKelas_Click()
    If chkjenisKelas.Value = Checked Then
        dcJenisKelas.Enabled = True
    Else
        dcJenisKelas.Text = ""
        dcJenisKelas.Enabled = False
    End If
End Sub

Private Sub chkjenisKelas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If chkjenisKelas.Value = Checked Then dcJenisKelas.SetFocus Else dtpAwal.SetFocus
End Sub

Private Sub cmdCetak_Click()
On Error GoTo hell
mdTglAwal = dtpAwal.Value
mdTglAkhir = dtpAkhir.Value
If ValidasiTanggal(dtpAwal, dtpAkhir) = False Then Exit Sub
Call getData

    If optTotalBayar.Value = True Then
        mstrKriteria = "TotalBayar"
    ElseIf optHutangPenjamin.Value = True Then
        mstrKriteria = "TotalHutangPenjamin"
    ElseIf optPembebasan.Value = True Then
        mstrKriteria = "TotalPembebasan"
    ElseIf optTanggunganRS.Value = True Then
        mstrKriteria = "TotalTanggunganRS"
    ElseIf optSisaTagihan.Value = True Then
        mstrKriteria = "TotalSisaTagihan"
    End If
    
Call sub_KriteriaLaporan(mstrKriteria)
Set rs = Nothing
Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mdTglAwal = dtpAwal.Value
        mdTglAkhir = dtpAkhir.Value
        Set frmCetakPendapatanPerObatAlkes_Header = Nothing
        frmCetakPendapatanPerObatAlkes_Header.Show
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

Private Sub dcJenisKelas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAwal.SetFocus
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then optTotalBayar.SetFocus
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
    
    mstrKriteria = "TotalBayar"
    Call subLoadDC
    Call getData
'    Call cmdTampilkanTemp_Click
    Call PlayFlashMovie(Me)
    
Exit Sub
errFormLoad:
    msubPesanError
End Sub

Private Sub subLoadDC()
On Error GoTo errload
    
    strSQL = "SELECT KdDetailJenisJasaPelayanan,DetailJenisJasaPelayanan FROM DetailJenisJasaPelayanan"
    Call msubDcSource(dcJenisKelas, rs, strSQL)
    If Not rs.EOF Then dcJenisKelas.BoundText = rs(0).Value
    
Exit Sub
errload:
    Call msubPesanError
End Sub

Public Sub getData()
On Error GoTo errorLoad
Dim strFilter As String
    
    mstrFilter = ""
    If chkjenisKelas.Value = Checked Then
        mstrFilter = mstrFilter & " AND KdJenisKelas ='" & dcJenisKelas.BoundText & "'"
    End If
    
    Call sub_KriteriaLaporan(mstrKriteria)
    Call msubRecFO(rs, strSQL)
    
Exit Sub
errorLoad:
    Call msubPesanError
End Sub

Private Sub optHutangPenjamin_Click()
    If optHutangPenjamin.Value = True Then optHutangPenjamin.ForeColor = vbBlue
    optTotalBayar.ForeColor = vbBlack
    optPembebasan.ForeColor = vbBlack
    optTanggunganRS.ForeColor = vbBlack
    optSisaTagihan.ForeColor = vbBlack
End Sub

Private Sub optJenisPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAwal.SetFocus
End Sub

Private Sub optRuanganKasir_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAwal.SetFocus
End Sub

Private Sub optInstalasiAwal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAwal.SetFocus
End Sub

Private Sub optHutangPenjamin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdcetak.SetFocus
End Sub

Private Sub optPembebasan_Click()
    If optPembebasan.Value = True Then optPembebasan.ForeColor = vbBlue
    optHutangPenjamin.ForeColor = vbBlack
    optTotalBayar.ForeColor = vbBlack
    optTanggunganRS.ForeColor = vbBlack
    optSisaTagihan.ForeColor = vbBlack
End Sub

Private Sub optPembebasan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdcetak.SetFocus
End Sub

Private Sub optSisaTagihan_Click()
    If optSisaTagihan.Value = True Then optSisaTagihan.ForeColor = vbBlue
    optHutangPenjamin.ForeColor = vbBlack
    optTotalBayar.ForeColor = vbBlack
    optPembebasan.ForeColor = vbBlack
    optTanggunganRS.ForeColor = vbBlack
End Sub

Private Sub optSisaTagihan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdcetak.SetFocus
End Sub

Private Sub optTanggunganRS_Click()
    If optTanggunganRS.Value = True Then optTanggunganRS.ForeColor = vbBlue
    optHutangPenjamin.ForeColor = vbBlack
    optTotalBayar.ForeColor = vbBlack
    optPembebasan.ForeColor = vbBlack
    optSisaTagihan.ForeColor = vbBlack
End Sub

Private Sub optTanggunganRS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdcetak.SetFocus
End Sub

Private Sub optTotalBayar_Click()
    If optTotalBayar.Value = True Then optTotalBayar.ForeColor = vbBlue
    optHutangPenjamin.ForeColor = vbBlack
    optPembebasan.ForeColor = vbBlack
    optTanggunganRS.ForeColor = vbBlack
    optSisaTagihan.ForeColor = vbBlack
End Sub

Private Sub sub_KriteriaLaporan(s_Kriteria As String)
On Error GoTo errload
    
    Select Case s_Kriteria
        Case "TotalBayar"
            If chkjenisKelas.Value = Checked Then
                strSQL = " SELECT JenisKelas, JenisPasien,Penjamin, RuanganAsal, KomponenUnit, KomponenTarif, SUM(TotalBayar) AS TotalBayar " & _
                    " FROM V_LaporanPendapatanPerUnitLaporanOARekap " & _
                    " WHERE TglBKM BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd hh:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd hh:mm:59") & "'" & mstrFilter & " " & _
                    " And RuanganPelayanan = '" & mstrNamaRuangan & "' GROUP BY JenisKelas, JenisPasien,Penjamin, RuanganAsal, KomponenUnit, KomponenTarif"
            Else
                strSQL = " SELECT JenisKelas, JenisPasien,Penjamin, RuanganAsal, KomponenUnit, KomponenTarif, SUM(TotalBayar) AS TotalBayar " & _
                    " FROM V_LaporanPendapatanPerUnitLaporanOARekap " & _
                    " WHERE TglBKM BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd hh:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd hh:mm:59") & "' " & _
                    " And RuanganPelayanan = '" & mstrNamaRuangan & "' GROUP BY JenisKelas, JenisPasien,Penjamin, RuanganAsal, KomponenUnit, KomponenTarif"
            End If
        
        Case "TotalHutangPenjamin"
            If chkjenisKelas.Value = Checked Then
               strSQL = " SELECT JenisKelas, JenisPasien,Penjamin, RuanganAsal, KomponenUnit, KomponenTarif, SUM(TotalHutangPenjamin) AS TotalHutangPenjamin " & _
                    " FROM V_LaporanPendapatanPerUnitLaporanOARekap " & _
                    " WHERE TglBKM BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd hh:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd hh:mm:59") & "'" & mstrFilter & " " & _
                    " And RuanganPelayanan = '" & mstrNamaRuangan & "' GROUP BY JenisKelas, JenisPasien,Penjamin, RuanganAsal, KomponenUnit, KomponenTarif"
            Else
                strSQL = " SELECT JenisKelas, JenisPasien,Penjamin, RuanganAsal, KomponenUnit, KomponenTarif, SUM(TotalHutangPenjamin) AS TotalHutangPenjamin " & _
                    " FROM V_LaporanPendapatanPerUnitLaporanOARekap " & _
                    " WHERE TglBKM BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd hh:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd hh:mm:59") & "'" & _
                    " And RuanganPelayanan = '" & mstrNamaRuangan & "' GROUP BY JenisKelas, JenisPasien,Penjamin, RuanganAsal, KomponenUnit, KomponenTarif"
            End If
            
        Case "TotalPembebasan"
            If chkjenisKelas.Value = Checked Then
                strSQL = " SELECT JenisKelas, JenisPasien,Penjamin, RuanganAsal, KomponenUnit, KomponenTarif, SUM(TotalPembebasan) AS TotalPembebasan " & _
                    " FROM V_LaporanPendapatanPerUnitLaporanOARekap " & _
                    " WHERE TglBKM BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd hh:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd hh:mm:59") & "'" & mstrFilter & " " & _
                    " And RuanganPelayanan = '" & mstrNamaRuangan & "' GROUP BY JenisKelas, JenisPasien,Penjamin, RuanganAsal, KomponenUnit, KomponenTarif"
            Else
                strSQL = " SELECT JenisKelas, JenisPasien,Penjamin, RuanganAsal, KomponenUnit, KomponenTarif, SUM(TotalPembebasan) AS TotalPembebasan " & _
                    " FROM V_LaporanPendapatanPerUnitLaporanOARekap " & _
                    " WHERE TglBKM BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd hh:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd hh:mm:59") & "'" & _
                    " And RuanganPelayanan = '" & mstrNamaRuangan & "' GROUP BY JenisKelas, JenisPasien,Penjamin, RuanganAsal, KomponenUnit, KomponenTarif"
            End If
        
        Case "TotalTanggunganRS"
            If chkjenisKelas.Value = Checked Then
                strSQL = " SELECT JenisKelas, JenisPasien,Penjamin, RuanganAsal, KomponenUnit, KomponenTarif, SUM(TotalTanggunganRS) AS TotalTanggunganRS " & _
                    " FROM V_LaporanPendapatanPerUnitLaporanOARekap " & _
                    " WHERE TglBKM BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd hh:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd hh:mm:59") & "'" & mstrFilter & " " & _
                    " And RuanganPelayanan = '" & mstrNamaRuangan & "' GROUP BY JenisKelas, JenisPasien,Penjamin, RuanganAsal, KomponenUnit, KomponenTarif"
            Else
                strSQL = " SELECT JenisKelas, JenisPasien,Penjamin, RuanganAsal, KomponenUnit, KomponenTarif, SUM(TotalTanggunganRS) AS TotalTanggunganRS " & _
                    " FROM V_LaporanPendapatanPerUnitLaporanOARekap " & _
                    " WHERE TglBKM BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd hh:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd hh:mm:59") & "'" & _
                    " And RuanganPelayanan = '" & mstrNamaRuangan & "' GROUP BY JenisKelas, JenisPasien,Penjamin, RuanganAsal, KomponenUnit, KomponenTarif"
            End If
        
        Case "TotalSisaTagihan"
            If chkjenisKelas.Value = Checked Then
               strSQL = " SELECT JenisKelas, JenisPasien,Penjamin, RuanganAsal, KomponenUnit, KomponenTarif, SUM(TotalSisaTagihan) AS TotalSisaTagihan " & _
                    " FROM V_LaporanPendapatanPerUnitLaporanOARekap " & _
                    " WHERE TglBKM BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd hh:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd hh:mm:59") & "'" & mstrFilter & " " & _
                    " And RuanganPelayanan = '" & mstrNamaRuangan & "' GROUP BY JenisKelas, JenisPasien,Penjamin, RuanganAsal, KomponenUnit, KomponenTarif"
            Else
                strSQL = " SELECT JenisKelas, JenisPasien,Penjamin, RuanganAsal, KomponenUnit, KomponenTarif, SUM(TotalSisaTagihan) AS TotalSisaTagihan " & _
                    " FROM V_LaporanPendapatanPerUnitLaporanOARekap " & _
                    " WHERE TglBKM BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd hh:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd hh:mm:59") & "' " & _
                    " And RuanganPelayanan = '" & mstrNamaRuangan & "' GROUP BY JenisKelas, JenisPasien,Penjamin, RuanganAsal, KomponenUnit, KomponenTarif"
            End If
    End Select
    
Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub optTotalBayar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdcetak.SetFocus
End Sub
