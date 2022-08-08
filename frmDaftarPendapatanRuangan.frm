VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDaftarPendapatanRuangan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Laporan Pendapatan Ruangan"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPendapatanRuangan.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   8805
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   7200
      TabIndex        =   14
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "S&preadsheet"
      Height          =   495
      Left            =   5640
      TabIndex        =   13
      Top             =   3720
      Width           =   1455
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
      TabIndex        =   16
      Top             =   2760
      Width           =   8775
      Begin VB.OptionButton optTotalBayar 
         Caption         =   "Total Bayar"
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optHutangPenjamin 
         Caption         =   "Hutang Penjamin"
         Height          =   495
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optPembebasan 
         Caption         =   "Pembebasan"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3480
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optTanggunganRS 
         Caption         =   "Tanggungan RS"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5040
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optSisaTagihan 
         Caption         =   "Sisa Tagihan"
         Enabled         =   0   'False
         Height          =   495
         Left            =   6960
         TabIndex        =   12
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
      Height          =   1815
      Left            =   0
      TabIndex        =   15
      Top             =   960
      Width           =   8775
      Begin VB.CheckBox ChkGroupby 
         Caption         =   "Group By"
         Height          =   210
         Left            =   5400
         TabIndex        =   6
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1995
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
         Height          =   735
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   5055
         Begin VB.OptionButton optInstalasiAwal 
            Caption         =   "Instalasi Asal"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3000
            TabIndex        =   5
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optJenisPasien 
            Caption         =   "Jenis Pasien"
            Height          =   375
            Left            =   360
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.CheckBox chkKelasPelayanan 
         Caption         =   "Jenis Kelas Pelayanan"
         Enabled         =   0   'False
         Height          =   210
         Left            =   5400
         TabIndex        =   2
         Top             =   360
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
         Left            =   120
         TabIndex        =   17
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
            Format          =   54394883
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
            Format          =   116916227
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   2400
            TabIndex        =   18
            Top             =   360
            Width           =   255
         End
      End
      Begin MSDataListLib.DataCombo dcKelasPelayanan 
         Height          =   360
         Left            =   5400
         TabIndex        =   3
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSDataListLib.DataCombo DCGroupBy 
         Height          =   360
         Left            =   5400
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
      TabIndex        =   20
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
      Picture         =   "frmDaftarPendapatanRuangan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   6960
      Picture         =   "frmDaftarPendapatanRuangan.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarPendapatanRuangan.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmDaftarPendapatanRuangan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim kriteria As String

Private Sub ChkGroupby_Click()
    If chkGroupBy.Value = Checked Then dcGroupBy.Enabled = True Else dcGroupBy.Enabled = False
End Sub

Private Sub ChkGroupby_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If chkGroupBy.Value = Checked Then dcGroupBy.SetFocus Else optTotalBayar.SetFocus
End Sub

Private Sub chkKelasPelayanan_Click()
    If chkKelasPelayanan.Value = vbChecked Then dcKelasPelayanan.Enabled = True Else dcKelasPelayanan.Enabled = False
End Sub

Private Sub chkKelasPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If chkKelasPelayanan.Value = vbChecked Then dcKelasPelayanan.SetFocus Else optJenisPasien.SetFocus
End Sub

Private Sub cmdCetak_Click()
    mstrFilter = ""
   
    
'    If optHutangPenjamin.Value = True Then
'        'add arief, validasi IdPenjamin utk mengetahui Hut. Penjamin/ TRS & add hut. Penjamin/ TRS
'        strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
'        Call msubRecFO(rs, strSQL)
'        If rs.EOF = False Then
'            mstrKdJenisPasien = rs("KdKelompokPasien").Value
'            mstrKdPenjamin = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
'        End If
'
'        If mstrKdPenjamin <> "2222222222" Then
'            Screen.MousePointer = vbHourglass
'            If sp_PostingHutangPenjaminPasien_AU(mstrNoPen, "A") = False Then Exit Sub
'            Screen.MousePointer = vbDefault
'        End If
'        'end arief
'    End If
    Call getData

    
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value
    If ValidasiTanggal(dtpAwal, dtpAkhir) = False Then Exit Sub
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then MsgBox "Data Tidak Ada", vbExclamation, "Validasi": Exit Sub
    
    If optTotalBayar.Value = True Then
        Set frmCetakPendapatanRuangan = Nothing
        frmCetakPendapatanRuangan.Show
    ElseIf optHutangPenjamin.Value = True Then
        Set frmCetakPendapatanRuangan1 = Nothing
        frmCetakPendapatanRuangan1.Show
    End If
    
End Sub

Private Sub cmdtutup_Click()
    Unload Me
End Sub

Private Sub DcGroupBy_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then optTotalBayar.SetFocus
End Sub

Private Sub dcKelasPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then optJenisPasien.SetFocus
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = 13 Then chkKelasPelayanan.SetFocus
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
    Call PlayFlashMovie(Me)
    
    dtpAwal.Value = Format(Now, "dd MMMM yyyy 00:00:00")
    dtpAkhir.Value = Now
    
    If chkGroupBy.Value = vbChecked Then
       Call subloadDCGroupby
    End If
    Call subLoadDC
    Call getData
'    Call cmdTampilkanTemp_Click
    
Exit Sub
errFormLoad:
    msubPesanError
End Sub

Private Sub subLoadDC()
On Error GoTo errLoad
    
    strSQL = "SELECT KdDetailJenisJasaPelayanan,DetailJenisJasaPelayanan FROM DetailJenisJasaPelayanan"
    Call msubDcSource(dcKelasPelayanan, rs, strSQL)
    If Not rs.EOF Then dcKelasPelayanan.BoundText = rs(0).Value
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Public Sub getData()
On Error GoTo errorLoad
Dim strFilter As String
    
    mstrFilter = ""
    If chkGroupBy.Value = vbChecked Then
        If optJenisPasien.Value = True Then
            If chkKelasPelayanan.Value = vbChecked Then
                If Periksa("datacombo", dcKelasPelayanan, "Kelas Pelayanan kosong") = False Then Exit Sub
                mstrFilter = mstrFilter & " And KdJenisKelas = '" & dcKelasPelayanan.BoundText & "'And KdKelompokPasien= '" & dcGroupBy.BoundText & "' "
            End If
        ElseIf optInstalasiAwal.Value = True Then
            If chkKelasPelayanan.Value = vbChecked Then
                mstrFilter = mstrFilter & " And KdJenisKelas = '" & dcKelasPelayanan.BoundText & "'And InstalasiPerujuk like '%" & dcGroupBy.Text & "%' "
            End If
            mstrFilter = mstrFilter & " And KdJenisKelas = '" & dcKelasPelayanan.BoundText & "'And InstalasiPerujuk like '%" & dcGroupBy.Text & "%' "
        End If
    Else
        mstrFilter = mstrFilter & " And KdJenisKelas = '" & dcKelasPelayanan.BoundText & "'"
        dcGroupBy.Text = ""
    End If
    
    If optTotalBayar.Value = True Then
        kriteria = "And TotalBayar <> 0"
        Call sub_KriteriaLaporan("TotalBayar")
        
    ElseIf optHutangPenjamin.Value = True Then
        kriteria = "And JmlHutangPenjamin <> 0"
        Call sub_KriteriaLaporan("TotalHutangPenjamin")
        
    ElseIf optPembebasan.Value = True Then
        kriteria = "And TotalPembebasan <> 0"
        Call sub_KriteriaLaporan("TotalPembebasan")
    ElseIf optTanggunganRS.Value = True Then
        kriteria = "And TotalTanggunganRS <> 0"
        Call sub_KriteriaLaporan("TotalTanggunganRS")
    ElseIf optSisaTagihan.Value = True Then
        kriteria = "And TotalSisaTagihan <> 0"
        Call sub_KriteriaLaporan("TotalSisaTagihan")
    End If
    Call msubRecFO(rs, strSQL)

Exit Sub
errorLoad:
    Call msubPesanError
End Sub

Private Sub sub_KriteriaLaporan(s_KriteriaBayar As String)
Dim mstrKdJenisPasien1 As String
Dim mstrKdPenjamin1 As String
Dim i As Integer

On Error GoTo errLoad
    
     If optJenisPasien.Value = True Then
'        strSQL = " SELECT Penjamin, JenisKelas, JenisPasien, NamaPelayanan,  SUM(DISTINCT Jumlah) AS Jumlah, KomponenTarif , SUM(" & s_KriteriaBayar & ") AS " & s_KriteriaBayar & " , RuanganPelayanan as RuanganPerujuk, InstalasiPelayanan as InstalasiPerujuk" & _
'            " FROM V_LaporanPendapatanPerunitLaporanTMRekap2 " & _
'            " WHERE (KdRuanganPelayanan = '" & mstrKdRuangan & "') AND TglBKM BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "'" & mstrFilter & " " & _
'            " GROUP BY Penjamin, JenisKelas, JenisPasien, NamaPelayanan, KomponenTarif, RuanganPelayanan, InstalasiPelayanan"

'        strSQL = " SELECT Penjamin, JenisKelas, JenisPasien, (Select DISTINCT NamaPelayanan from V_DaftarTindakanPasienperDokter Where TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' AND KdRuangan = '" & mstrKdRuangan & "') AS NamaPelayanan," & _
'            " (Select SUM(DISTINCT JmlPelayanan) from V_DaftarTindakanPasienperDokter Where TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' AND KdRuangan = '" & mstrKdRuangan & "'" & _
'            " AND TindakanPelayanan = V_LaporanPendapatanPerunitLaporanTMRekap2.NamaPelayanan AND JenisPasien = V_LaporanPendapatanPerunitLaporanTMRekap2.JenisPasien AND NamaPenjamin = V_LaporanPendapatanPerunitLaporanTMRekap2.Penjamin)AS Jumlah  , " & _
'            " KomponenTarif , SUM(" & s_KriteriaBayar & ") AS " & s_KriteriaBayar & " , RuanganPelayanan as RuanganPerujuk, InstalasiPelayanan as InstalasiPerujuk" & _
'            " FROM V_LaporanPendapatanPerunitLaporanTMRekap2 " & _
'            " WHERE (KdRuanganPelayanan = '" & mstrKdRuangan & "') AND TglBKM BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "'" & mstrFilter & " " & _
'            " GROUP BY Penjamin, JenisKelas, JenisPasien, NamaPelayanan, KomponenTarif, RuanganPelayanan, InstalasiPelayanan"

    'add dimas 25/10/2011
    '**********************************************************************
        If optHutangPenjamin.Value = True Then
            strSQL = " SELECT DISTINCT NoPendaftaran from V_DaftarTindakanPasienperDokterNew Where TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' AND KdRuangan = '" & mstrKdRuangan & "'" & mstrFilter & " "
            Call msubRecFO(rs, strSQL)
            If rs.EOF = True Or rs.BOF = True Then Exit Sub
            For i = 1 To rs.RecordCount
                'add arief, validasi IdPenjamin utk mengetahui Hut. Penjamin/ TRS & add hut. Penjamin/ TRS
                strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & rs.Fields(0).Value & "')"
                Call msubRecFO(rsB, strSQL)
                If rsB.EOF = False Then
                    mstrKdJenisPasien1 = rsB("KdKelompokPasien").Value
                    mstrKdPenjamin1 = IIf(IsNull(rsB("IdPenjamin")), "2222222222", rsB("IdPenjamin"))
                End If

                If mstrKdPenjamin1 <> "2222222222" Then
                    Screen.MousePointer = vbHourglass
                    If sp_PostingHutangPenjaminPasien_AU(rs.Fields(0).Value, "A") = False Then Exit Sub
                    Screen.MousePointer = vbDefault
                End If
                If i = rs.RecordCount Then Exit For
                rs.MoveNext
            Next i
        End If
    '**********************************************************************
    'end
        
'        strSQL = " SELECT DISTINCT NamaRuangan, JenisPasien, NamaPenjamin, TindakanPelayanan,  SUM(JmlPelayanan) AS Jumlah, Tarif, NamaKomponen , SUM (JmlHutangPenjamin) AS JmlHutangPenjamin, SUM(isnull (JmlBayar,0)) AS " & s_KriteriaBayar & "  " & _
'                 " from V_DaftarTindakanPasienperDokter Where TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' AND KdRuangan = '" & mstrKdRuangan & "'" & mstrFilter & " " & _
'                 " Group By NamaRuangan, JenisPasien, NamaPenjamin, TindakanPelayanan, NamaKomponen, Tarif "

        strSQL = " SELECT DISTINCT NamaRuangan, JenisPasien, NamaPenjamin, TindakanPelayanan,  SUM(JmlPelayanan) AS Jumlah, Tarif, DeskKelas, NamaKomponen , SUM (JmlHutangPenjamin) AS JmlHutangPenjamin, SUM(isnull (JmlBayar,0)) AS " & s_KriteriaBayar & " , RuanganPerujuk " & _
                 " from V_DaftarTindakanPasienRumahSakit Where KdRuangan = '" & mstrKdRuangan & "' AND TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' AND KdRuangan = '" & mstrKdRuangan & "'" & mstrFilter & " " & _
                 " Group By NamaRuangan, JenisPasien, NamaPenjamin, TindakanPelayanan, NamaKomponen, Tarif, DeskKelas, RuanganPerujuk "

                 

    ElseIf optInstalasiAwal.Value = True Then
        strSQL = " SELECT Penjamin, JenisKelas, case InstalasiPerujuk when 'Instalasi Laboratorium Klinik' then 'Instalasi Rawat Jalan' end as InstalasiPerujuk, NamaPelayanan, RuanganPerujuk, (Select SUM(JmlPelayanan) from V_DaftarTindakanPasienperDokter Where TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' AND KdRuangan = '" & mstrKdRuangan & "' AND TindakanPelayanan = V_LaporanPendapatanPerunitLaporanTMRekap2.NamaPelayanan )AS Jumlah, KomponenTarif, SUM(" & s_KriteriaBayar & ") AS " & s_KriteriaBayar & " " & _
            " FROM V_LaporanPendapatanPerunitLaporanTMRekap " & _
            " WHERE (KdRuanganPerujuk = '" & mstrKdRuangan & "' OR KdRuanganPelayanan = '" & mstrKdRuangan & "') And KdInstalasiPerujuk like '%" & dcGroupBy.BoundText & "%'  AND " & s_KriteriaBayar & " <> 0 AND TglBKM BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "'" & mstrFilter & " " & _
            " GROUP BY Penjamin, JenisKelas, NamaPelayanan, KomponenTarif, KdRuanganPerujuk, RuanganPerujuk, KdInstalasiPerujuk, InstalasiPerujuk"
    End If

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub optInstalasiAwal_Click()
    chkGroupBy.Caption = "Instalasi Asal"
    Call subloadDCGroupby
End Sub

Private Sub optInstalasiAwal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkGroupBy.SetFocus
End Sub

Private Sub optJenisPasien_Click()
    chkGroupBy.Caption = "Jenis Pasien"
    Call subloadDCGroupby
End Sub

Private Sub optJenisPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkGroupBy.SetFocus
End Sub

Private Sub optHutangPenjamin_Click()
    If optHutangPenjamin.Value = True Then optHutangPenjamin.ForeColor = vbBlue
    optTotalBayar.ForeColor = vbBlack
    optPembebasan.ForeColor = vbBlack
    optTanggunganRS.ForeColor = vbBlack
    optSisaTagihan.ForeColor = vbBlack
End Sub

Private Sub optPembebasan_Click()
    If optPembebasan.Value = True Then optPembebasan.ForeColor = vbBlue
    optHutangPenjamin.ForeColor = vbBlack
    optTotalBayar.ForeColor = vbBlack
    optTanggunganRS.ForeColor = vbBlack
    optSisaTagihan.ForeColor = vbBlack
End Sub

Private Sub optPembebasan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetak.SetFocus
End Sub

Private Sub optSisaTagihan_Click()
    If optSisaTagihan.Value = True Then optSisaTagihan.ForeColor = vbBlue
    optHutangPenjamin.ForeColor = vbBlack
    optTotalBayar.ForeColor = vbBlack
    optPembebasan.ForeColor = vbBlack
    optTanggunganRS.ForeColor = vbBlack
End Sub

Private Sub optSisaTagihan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetak.SetFocus
End Sub

Private Sub optTanggunganRS_Click()
    If optTanggunganRS.Value = True Then optTanggunganRS.ForeColor = vbBlue
    optHutangPenjamin.ForeColor = vbBlack
    optTotalBayar.ForeColor = vbBlack
    optPembebasan.ForeColor = vbBlack
    optSisaTagihan.ForeColor = vbBlack
End Sub

Private Sub optTanggunganRS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetak.SetFocus
End Sub

Private Sub optTotalBayar_Click()
    If optTotalBayar.Value = True Then optTotalBayar.ForeColor = vbBlue
    optHutangPenjamin.ForeColor = vbBlack
    optPembebasan.ForeColor = vbBlack
    optTanggunganRS.ForeColor = vbBlack
    optSisaTagihan.ForeColor = vbBlack
End Sub

Private Sub optTotalBayar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetak.SetFocus
End Sub

Private Sub subloadDCGroupby()
Set rs = Nothing
    If optJenisPasien.Value = True Then
        strSQL = "SELECT KdKelompokPasien,JenisPasien FROM KelompokPasien"
    ElseIf optInstalasiAwal.Value = True Then
        strSQL = "select * from v_InstalasiRujukan "
    End If
    Call msubDcSource(dcGroupBy, rs, strSQL)
    If Not rs.EOF Then dcGroupBy.BoundText = rs(0).Value
    'Call getData
End Sub

