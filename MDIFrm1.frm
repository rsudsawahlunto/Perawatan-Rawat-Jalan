VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIUtama 
   BackColor       =   &H8000000C&
   Caption         =   "Medifirst2000 - Perawatan Rawat Jalan (Outpatient)"
   ClientHeight    =   8100
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   15195
   Icon            =   "MDIFrm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIFrm1.frx":0CCA
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CDPrinter 
      Left            =   960
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7845
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7832
            MinWidth        =   7832
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   9596
            MinWidth        =   9596
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "11/12/2019"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "15:10"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   6059
            MinWidth        =   6068
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnberkas 
      Caption         =   "&Berkas"
      Begin VB.Menu mnudata 
         Caption         =   "Data"
         Begin VB.Menu mnucdp 
            Caption         =   "Daftar Pasien Poliklinik"
            Shortcut        =   {F2}
         End
         Begin VB.Menu mnusepdpl 
            Caption         =   "-"
         End
         Begin VB.Menu mnudpa 
            Caption         =   "Daftar Antrian Pasien"
            Shortcut        =   {F6}
         End
         Begin VB.Menu MDaftarPasienRumahSakit 
            Caption         =   "Daftar Pasien Rumah Sakit"
            Shortcut        =   {F7}
            Visible         =   0   'False
         End
         Begin VB.Menu mnusepdap 
            Caption         =   "-"
         End
         Begin VB.Menu MDaftarDokumenRekamMedis 
            Caption         =   "Daftar Dokumen Rekam Medis"
            Shortcut        =   {F12}
         End
         Begin VB.Menu LDaftarPasienSudahBayar 
            Caption         =   "-"
         End
         Begin VB.Menu mnuKegiatanPenyuluhan 
            Caption         =   "Kegiatan Penyuluhan"
            Visible         =   0   'False
         End
         Begin VB.Menu mnupeny 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuBDataDiag 
            Caption         =   "Diagnosa Ruangan"
            Shortcut        =   ^R
         End
         Begin VB.Menu mnusepdr 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnupp 
            Caption         =   "Paket Pelayanan"
            Shortcut        =   ^D
            Visible         =   0   'False
         End
         Begin VB.Menu mnuClosingDataPelayananTMOAApotik 
            Caption         =   "Informasi Data Pelayanan TMOAApotik"
            Visible         =   0   'False
         End
         Begin VB.Menu linee 
            Caption         =   "-"
         End
         Begin VB.Menu MInformasiTarifPelayanan 
            Caption         =   "Informasi Tarif Pelayanan"
         End
         Begin VB.Menu mnPanAbulance 
            Caption         =   "Pesan Ambulance"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuGaris 
         Caption         =   "-"
      End
      Begin VB.Menu mSettingPrinter 
         Caption         =   "Setting Printer"
         Shortcut        =   ^P
      End
      Begin VB.Menu mGantiKataKunci 
         Caption         =   "Ganti Kata Kunci"
         Shortcut        =   ^G
      End
      Begin VB.Menu mspace3 
         Caption         =   "-"
      End
      Begin VB.Menu mnlogout 
         Caption         =   "Log Off"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnSelesai 
         Caption         =   "Keluar"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuInformasi 
      Caption         =   "&Informasi"
      Begin VB.Menu mnPembayaran 
         Caption         =   "Monitoring Pembayaran"
      End
      Begin VB.Menu MnDiagnosa 
         Caption         =   "Diagnosa Pasien"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuedge 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPesanPelayananTMOA 
         Caption         =   "Daftar Pesan Pelayanan Dan Resep"
      End
      Begin VB.Menu mnuDaftarPengirimanDarah 
         Caption         =   "Daftar Pengiriman Darah"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuivt 
      Caption         =   "In&ventory"
      Begin VB.Menu mnupb 
         Caption         =   "Pesan Barang"
      End
      Begin VB.Menu mnuPemakaianBahandanAlat 
         Caption         =   "Pemakaian Bahan dan Alat"
         Visible         =   0   'False
      End
      Begin VB.Menu batasinv 
         Caption         =   "-"
      End
      Begin VB.Menu mnBrgM 
         Caption         =   "Barang Medis"
         Begin VB.Menu mnStokBarang 
            Caption         =   "Stok Barang"
         End
         Begin VB.Menu MStokOpname 
            Caption         =   "Closing Stok"
            Begin VB.Menu mnCetakLembarInput 
               Caption         =   "Cetak Lembar Input"
            End
            Begin VB.Menu mnInputStokOpname 
               Caption         =   "Input Stok Opname"
            End
            Begin VB.Menu MNilaiPersediaan 
               Caption         =   "Nilai Persediaan"
            End
         End
         Begin VB.Menu mnSO 
            Caption         =   "-"
         End
         Begin VB.Menu mnuipb 
            Caption         =   "Informasi Pemesanan && Penerimaan Barang"
         End
         Begin VB.Menu mnuipoal 
            Caption         =   "Informasi Pemakaian Barang"
            Visible         =   0   'False
         End
         Begin VB.Menu mnLapSaldoBarang 
            Caption         =   "Laporan Saldo Barang"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnBrgNM 
         Caption         =   "Barang Non Medis"
         Begin VB.Menu mnStokBarangNM 
            Caption         =   "Stok Barang"
         End
         Begin VB.Menu mnKondisiBarang 
            Caption         =   "Kondisi Barang"
         End
         Begin VB.Menu mMutasiBarangNM 
            Caption         =   "Mutasi Barang"
         End
         Begin VB.Menu mnClosingStokNM 
            Caption         =   "Closing Stok"
            Begin VB.Menu mnCetakFormStokNM 
               Caption         =   "Cetak Lembar Input"
            End
            Begin VB.Menu mnInputStokOpnameNM 
               Caption         =   "Input Stok Opname"
            End
            Begin VB.Menu mnNilaiPersediaanNM 
               Caption         =   "Nilai Persediaan"
            End
         End
         Begin VB.Menu mnusepsb 
            Caption         =   "-"
         End
         Begin VB.Menu mnInfoPesanBrgNM 
            Caption         =   "Informasi Pemesanan && Penerimaan Barang"
         End
         Begin VB.Menu mnLapSaldoBarangNM 
            Caption         =   "Laporan Saldo Barang"
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu mnulap 
      Caption         =   "&Laporan"
      Begin VB.Menu mnubrp 
         Caption         =   "Buku Register Pasien"
      End
      Begin VB.Menu MnLPRP 
         Caption         =   "Laporan Buku Register Pelayanan"
      End
      Begin VB.Menu mnltp 
         Caption         =   "Laporan Tindakan Pasien per Dokter"
         Visible         =   0   'False
      End
      Begin VB.Menu mnusepbrp 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu RPBJS 
         Caption         =   "Rekap Kunjungan Berdasarkan Status dan Jenis "
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu SR 
         Caption         =   "Rekap Kunjungan Berdasarkan Status dan Rujukan"
         Visible         =   0   'False
      End
      Begin VB.Menu SDJP 
         Caption         =   "Rekap Kunjungan Pasien Berdasarkan Status dan Kasus Penyakit"
         Visible         =   0   'False
      End
      Begin VB.Menu JOP 
         Caption         =   "Rekap Kunjungan Pasien Berdasarkan Status dan Kelas"
         Visible         =   0   'False
      End
      Begin VB.Menu RKPBW 
         Caption         =   "Rekap Kunjungan Berdasarkan Wilayah"
         Visible         =   0   'False
      End
      Begin VB.Menu aaaaaaaa 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu RPBD 
         Caption         =   "Rekapitulasi Pasien Berdasarkan Diagnosa"
         Visible         =   0   'False
      End
      Begin VB.Menu RPBWD 
         Caption         =   "Rekapitulasi Pasien Berdasarkan Wilayah Diagnosa"
         Visible         =   0   'False
      End
      Begin VB.Menu RPBT 
         Caption         =   "Rekapitulasi Pasien Berdasarkan Jenis Tindakan"
         Visible         =   0   'False
      End
      Begin VB.Menu MnRekPPerDokter 
         Caption         =   "Rekapitulasi Pasien per Dokter Berdasarkan Jenis Pasien"
         Visible         =   0   'False
      End
      Begin VB.Menu ssssssss 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnur10bp 
         Caption         =   "Rekapitulasi 10 Besar Penyakit"
         Visible         =   0   'False
      End
      Begin VB.Menu mnudsmp 
         Caption         =   "Data Surveilens Morbiditas Pasien"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnsprtr 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MnLapDIH 
         Caption         =   "Laporan Data Ibu Hamil"
         Visible         =   0   'False
      End
      Begin VB.Menu MnLapDIK 
         Caption         =   "Laporan Data Ibu KB"
         Visible         =   0   'False
      End
      Begin VB.Menu MnLapDB 
         Caption         =   "Laporan Data Bayi"
         Visible         =   0   'False
      End
      Begin VB.Menu mnLab 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu LapKegKeu 
         Caption         =   "Laporan Kegiatan Keuangan Poliklinik Rawat Jalan"
         Visible         =   0   'False
      End
      Begin VB.Menu LapKunjPasienBdsrDiagStPasien 
         Caption         =   "Laporan Kunjungan Pasien bdsr Diagnosa Status Pasien"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLapPendapatan 
         Caption         =   "Laporan Pendapatan Ruangan"
      End
   End
   Begin VB.Menu mnuw 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuc 
         Caption         =   "&Cascade"
      End
   End
   Begin VB.Menu mbantuan 
      Caption         =   "Ban&tuan"
      Begin VB.Menu mTentang 
         Caption         =   "Tentang Medifirst2000"
         Shortcut        =   ^T
      End
   End
End
Attribute VB_Name = "MDIUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sepuh As Boolean

Private Sub JOP_Click()
    strCetak = "LapKunjunganKelasStatus"
    frmLapRKP_KPSK.Show
End Sub

Private Sub LapKegKeu_Click()
    frmLapKegiatanKeuangan.Show
End Sub

Private Sub LapKunjPasienBdsrDiagStPasien_Click()
    frmLap20091201001.Show
End Sub

Private Sub MDaftarDokumenRekamMedis_Click()
    frmDaftarDokumenRekamMedisPasien.Show
End Sub

Private Sub MDaftarPasienRumahSakit_Click()
    frmDaftarPasienRJRIIGD.Show
End Sub

Private Sub MDIForm_Load()
'    Call openConnection
'    strSQL = "SELECT * FROM DataPegawai WHERE IdPegawai = '" & strIDPegawaiAktif & "'"
'    Set rs = Nothing
'    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
'    strNmPegawai = rs.Fields("NamaLengkap").Value
'    Set rs = Nothing
'
'    StatusBar1.Panels(1).Text = "Nama User : " & strNmPegawai
'    StatusBar1.Panels(2).Text = "Nama Ruangan : " & mstrNamaRuangan
'    StatusBar1.Panels(5).Text = "Nama Komputer : " & strNamaHostLocal
'    mnlogout.Caption = "Log Off..." & strNmPegawai
'    blnFrmCariPasien = False
'
'    strSQL = "SELECT * FROM StatusObject WHERE KdAplikasi='006' AND NamaForm='MDIUtama' AND NamaObject='mnupkpk' AND StatusEnable='T'"
'    msubRecFO rs, strSQL
'    If mstrKdRuangan = "204" Or mstrKdRuangan = "113" Then
'        MnLapDIH.Enabled = True
'        MnLapDIK.Enabled = True
'    Else
'        MnLapDIH.Enabled = False
'        MnLapDIK.Enabled = False
'    End If
'
'    If mstrKdRuangan = "112" Or mstrKdRuangan = "111" Then
'        MnLapDB.Enabled = True
'    Else
'        MnLapDB.Enabled = False
'    End If
'
'     strSQL = "Select MetodeStokBarang From SuratKeputusanRuleRS where statusenabled=1"
'    Call msubRecFO(dbRst, strSQL)
'    If dbRst.EOF = True Then
'        bolStatusFIFO = False
'    Else
'        If dbRst("MetodeStokBarang") = 0 Then
'            bolStatusFIFO = False
'        Else
'            bolStatusFIFO = True
'        End If
'    End If
    On Error GoTo errLoad
    Call openConnection
    strSQL = "SELECT * FROM DataPegawai WHERE IdPegawai = '" & strIDPegawaiAktif & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    strNmPegawai = rs.Fields("NamaLengkap").Value
    Set rs = Nothing
    StatusBar1.Panels(1).Text = "Nama User : " & strNmPegawai
    StatusBar1.Panels(2).Text = "Nama Ruangan : " & mstrNamaRuangan
    StatusBar1.Panels(5).Text = "Nama Komputer : " & strNamaHostLocal
    StatusBar1.Panels(6).Text = "Server : " & strServerName & " (" & strDatabaseName & ")"
    mnlogout.Caption = "Log Off..." & strNmPegawai

    
'    If mblnAdmin = False Then
'        MTransaksi.Enabled = False
'    Else
'        MTransaksi.Enabled = True
'    End If

    strSQL = "SELECT TerminBayarFakturSupplier, PersentasePpn, PersentaseLimitDiscount, PersentaseJasaPenulisResep, BiayaAdministrasi " & _
    " From SettingDataPendukung" & _
    " WHERE (KdInstalasi = '" & mstrKdInstalasiLogin & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
        typSettingDataPendukung.intTerminBayarFaktur = 0
        typSettingDataPendukung.realJasaPenulisResep = 0
        typSettingDataPendukung.realLimitDiscount = 0
        typSettingDataPendukung.realPPn = 0
        typSettingDataPendukung.curBiayaAdministrasi = 0
    Else
        typSettingDataPendukung.intTerminBayarFaktur = rs("TerminBayarFakturSupplier").Value
        typSettingDataPendukung.realJasaPenulisResep = rs("PersentaseJasaPenulisResep").Value
        typSettingDataPendukung.realLimitDiscount = rs("PersentaseLimitDiscount").Value
        typSettingDataPendukung.realPPn = rs("PersentasePpn").Value
        typSettingDataPendukung.curBiayaAdministrasi = rs("BiayaAdministrasi").Value
    End If

    strSQL = "SELECT JmlPembulatanHarga, JumlahBAdminOAPerBaris, JumlahBAdminTMPerHari FROM MasterDataPendukung"
    Call msubRecFO(dbRst, strSQL)
    If dbRst.EOF = True Then
        typSettingDataPendukung.intJmlPembulatanHarga = 0
        typSettingDataPendukung.intJumlahBAdminOAPerBaris = 0
        typSettingDataPendukung.intJumlahBAdminTMPerHari = 0
    Else
        typSettingDataPendukung.intJmlPembulatanHarga = dbRst(0)
        typSettingDataPendukung.intJumlahBAdminOAPerBaris = dbRst(1)
        typSettingDataPendukung.intJumlahBAdminTMPerHari = dbRst(2)
    End If
    
    strSQL = "SELECT JmlBarisOAPerTarifAdminOA from SettingBiayaAdministrasi"
    Call msubRecFO(dbRst, strSQL)
    If dbRst.EOF = True Then
        typSettingDataPendukung.intJumlahBAdminOAPerBaris = 0
    Else
        
        typSettingDataPendukung.intJumlahBAdminOAPerBaris = dbRst(0)
        
    End If

    strSQL = "Select MetodeStokBarang From SuratKeputusanRuleRS where statusenabled=1"
    Call msubRecFO(dbRst, strSQL)
    If dbRst.EOF = True Then
        bolStatusFIFO = False
    Else
        If dbRst("MetodeStokBarang") = 0 Then
            bolStatusFIFO = False
        Else
            bolStatusFIFO = True
        End If
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub MDIForm_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbLeftButton Then Exit Sub
    PopupMenu mnudata
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim q As String
    If sepuh = True Then
        q = MsgBox("Log Off user " & strNmPegawai & " ", vbQuestion + vbOKCancel, "Konfirmasi")
        If q = 2 Then
            Unload frmLogin
            Cancel = 1
        Else
            Cancel = 0
            frmLogin.Show
        End If
        sepuh = False
    Else
        q = MsgBox("Tutup aplikasi ", vbQuestion + vbOKCancel, "Konfirmasi")
        If q = 2 Then

            Unload frmLogin
            Cancel = 1
        Else
            dTglLogout = Now
            Call subSp_HistoryLoginAplikasi("U")
            Cancel = 0
        End If
    End If
End Sub

Private Sub mGantiKataKunci_Click()
    frmLoginEditAccount.Show
End Sub

Private Sub MInformasiTarifPelayanan_Click()
    frmInformasiTarifPelayanan.Show
End Sub

Private Sub mMutasiBarangNM_Click()
    frmMutasiBarangNM.Show
End Sub

Private Sub mnCetakFormStokNM_Click()
    mstrKdKelompokBarang = "01" 'non medis
    frmDaftarCetakInputStokOpnameNM.Show
End Sub

Private Sub mnCetakLembarInput_Click()
    mstrKdKelompokBarang = "02" 'medis
    frmDaftarCetakInputStokOpname.Show
End Sub

Private Sub MnDiagnosa_Click()
    FrmInformasiDiagnosa.Show
End Sub

Private Sub MNilaiPersediaan_Click()
    mstrKdKelompokBarang = "02"
    frmNilaiPersediaan.Show
End Sub

Private Sub mnInfoPesanBrgNM_Click()
    mstrKdKelompokBarang = "01"     'non medis
    frmInfoPesanBarangNM.Show
End Sub

Private Sub mnInputStokOpname_Click()
    mstrKdKelompokBarang = "02" 'medis
    frmStokOpname.Show
End Sub

Private Sub mnInputStokOpnameNM_Click()
    mstrKdKelompokBarang = "01"     'non medis
    frmStokOpnameNM.Show
End Sub

Private Sub mnKondisiBarang_Click()
    frmKondisiBarangNM.Show
End Sub

Private Sub MnLapDB_Click()
    frmDaftarCetakLapDataBayiBulan.Show
End Sub

Private Sub MnLapDIH_Click()
    frmDaftarCetakLapDataBumil.Show
End Sub

Private Sub MnLapDIK_Click()
    frmDaftarCetakLapDataKB.Show
End Sub

Private Sub mnLapSaldoBarang_Click()
    mstrKdKelompokBarang = "02" 'medis
    frmLaporanSaldoBarangMedis_v3.Show
End Sub

Private Sub mnLapSaldoBarangNM_Click()
    mstrKdKelompokBarang = "01" 'non medis
    frmLaporanSaldoBarangNM_v3.Show
End Sub

Private Sub mnlogout_Click()
    Dim adoCommand As New ADODB.Command
    openConnection
    sepuh = True
    strQuery = "UPDATE Login SET Status = '0' " & _
    "WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
    adoCommand.ActiveConnection = dbConn
    adoCommand.CommandText = strQuery
    adoCommand.CommandType = adCmdText
    adoCommand.Execute
    dTglLogout = Now
    Call subSp_HistoryLoginAplikasi("U")

    Unload Me
End Sub

Private Sub MnLPRP_Click()
    FrmBukuRegisterPelayanan.Show
End Sub

Private Sub mnltp_Click()
    frmDaftarTindakan.Show
End Sub

Private Sub mnMutasiBarang_Click()
    frmMutasiBarangNM.Show
End Sub

Private Sub mnNilaiPersediaanNM_Click()
    mstrKdKelompokBarang = "01"
    frmNilaiPersediaanNM.Show
End Sub

Private Sub mnRekapTransBrg_Click()
    mstrKdKelompokBarang = "02"     'medis
    frmDataTransaksiBarang.Show
End Sub

Private Sub mnRekapTransBrgNM_Click()
    mstrKdKelompokBarang = "01" 'non medis
    frmDataTransaksiBarangNM.Show
End Sub

Private Sub mnPanAbulance_Click()
    frmPesanAmbulans.Show
End Sub

Private Sub mnPembayaran_Click()
    frmMonitoringPembayaran.Show
End Sub

Private Sub MnRekPPerDokter_Click()
    frmLapRekPerDokter.Show
End Sub

Private Sub mnSelesai_Click()
    Dim pesan As VbMsgBoxResult
    Dim adoCommand As New ADODB.Command
    pesan = MsgBox("Tutup aplikasi ", vbQuestion + vbYesNo, "Konfirmasi")
    If pesan = vbYes Then

        openConnection
        strQuery = "UPDATE Login SET Status = '0' " & _
        "WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
        adoCommand.ActiveConnection = dbConn
        adoCommand.CommandText = strQuery
        adoCommand.CommandType = adCmdText
        adoCommand.Execute

        dTglLogout = Now
        Call subSp_HistoryLoginAplikasi("U")
        End
    End If
End Sub

Private Sub mnStokBarang_Click()
    frmStokBrg.Show
End Sub

Private Sub mnStokBarangNM_Click()
    frmStokBarangNonMedis.Show
End Sub

Private Sub mnuBDataDiag_Click()
    frmDataDiagnosa.Show
End Sub

Private Sub mnubrp_Click()
   FrmBukuRegisterPasien.Show
'    FrmBukuRegister.Show
End Sub

Private Sub mnuc_Click()
    MDIUtama.Arrange vbCascade
End Sub

Private Sub mnucdp_Click()
    frmDaftarPasienRJ.Show
End Sub

Private Sub mnuClosingDataPelayananTMOAApotik_Click()
    frmClosingDataPelayananTM_OA_Apotik.Show
End Sub

Private Sub mnuDaftarPengirimanDarah_Click()
    frmDaftarPengirimanDarah.Show
End Sub

Private Sub mnudpa_Click()
    frmDaftarAntrianPasien.Show
End Sub

Private Sub mnudsmp_Click()
    frmLapMorbiditas.Show
    frmLapMorbiditas.Caption = "Medifirst2000 - Data Keadaan Morbiditas Pasien"
End Sub

Private Sub mnuipb_Click()
    mstrKdKelompokBarang = "02"
    frmInfoPesanBarang.Show
End Sub

Private Sub mnuipoal_Click()
    frmDaftarPakaiAlkesKaryawan.Show
End Sub

Private Sub mnuKegiatanPenyuluhan_Click()
    frmKegiatanPenyuluhan.Show
End Sub

Private Sub mnuLapPendapatan_Click()
    frmDaftarPendapatanRuangan.Show
End Sub

Private Sub mnupb_Click()
    frmPemesananBarang.Show
End Sub

Private Sub mnuPemakaianBahandanAlat_Click()
    frmPemakaianBahanAlat.Show
End Sub

Private Sub mnuPesanPelayananTMOA_Click()
    frmInfoPesanPelayananTMOA.Show
End Sub

Private Sub mnupp_Click()
    frmPaketLayanan.Show
End Sub

Private Sub mnur10bp_Click()
    FrmPeriodeLaporanTopTen.Show
    FrmPeriodeLaporanTopTen.Caption = "Medifirst2000 - Rekapitulasi 10 Besar Penyakit"
End Sub

Private Sub mnusb_Click()
    frmStokBrg.Show
End Sub

Private Sub mSettingPrinter_Click()
'    frmSetupPrinter.Show
    frmSetupPrinter2.Show
End Sub

Private Sub mTentang_Click()
    frmAbout.Show
End Sub

Private Sub RKPBW_Click()
    strCetak = "LapKunjunganBwilayah"
    frmLapRKP_KPSK.Show
End Sub

Private Sub RPBD_Click()
    strCetak = "LapKunjunganBDiagnosa"
    frmLapRKP_KPSK.Show
End Sub

Private Sub RPBJS_Click()
    strCetak = "LapKunjunganJenisStatus"
    frmLapRKP_KPSK.Show
End Sub

Private Sub RPBT_Click()
    strCetak = "lapkunjunganjenistindakan"
    frmFilterJenisPeriksa.Show
End Sub

Private Sub RPBWD_Click()
    strCetak = "LapKunjunganPasienBDiagnosaWilayah"
    frmLapRKP_KPSK.Show
End Sub

Private Sub SDJP_Click()
    strCetak = "LapKunjunganSt_PnyktPsn"
    frmLapRKP_KPSK.Show
End Sub

Private Sub SdKP_Click()
    strCetak = "LapKunjunganKonPulang_Status"
    frmLapRKP_KPSK.Show
End Sub

Private Sub SR_Click()
    strCetak = "LapKunjunganRujukanBStatus"
    frmLapRKP_KPSK.Show
End Sub

