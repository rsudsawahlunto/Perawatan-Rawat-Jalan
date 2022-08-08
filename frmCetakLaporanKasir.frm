VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakLaporanKasir 
   Caption         =   "Medifirst2000 - Laporan"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakLaporanKasir.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCetakLaporanKasir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rptPenerimaan As crPenerimaan
Dim rptPembatalanStruk As crPembatalanStruk
Dim rptPenerimaanDetail As crPenerimaanDetail

Private Sub Form_Load()
    Me.WindowState = 2
    Screen.MousePointer = vbHourglass
    Set dbcmd = New ADODB.Command
    Set dbcmd.ActiveConnection = dbConn
    
    Select Case mstrLaporan
    Case "PenerimaanPerda"
        If frmRekapLaporan.dcJenisLaporan.BoundText = "01" Then 'perda a/ tindakan
            Me.Caption = "Medifirst2000 - Laporan Penerimaan Perda"
            Set rptPenerimaan = New crPenerimaan
            strSQL = "SELECT * " & _
                " FROM V_S_Lap_PenerimaanKasir_Perda " & _
                " WHERE JenisPasien='" & frmRekapLaporan.dcPenjamin.Text & "' AND TglStruk BETWEEN '" & Format(frmRekapLaporan.dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(frmRekapLaporan.dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuanganKasir='" & mstrKdRuangan & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "'"
            dbcmd.CommandText = strSQL
            dbcmd.CommandType = adCmdText
        ElseIf frmRekapLaporan.dcJenisLaporan.BoundText = "02" Then 'non perda / alkes
            Me.Caption = "Medifirst2000 - Laporan Penerimaan Kasir Non Perda"
            Set rptPenerimaan = New crPenerimaan
            strSQL = "SELECT * " & _
                " FROM V_S_Lap_PenerimaanKasir_NonPerda " & _
                " WHERE JenisPasien='" & frmRekapLaporan.dcPenjamin.Text & "' AND TglStruk BETWEEN '" & Format(frmRekapLaporan.dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(frmRekapLaporan.dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuanganKasir='" & mstrKdRuangan & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "'"
            dbcmd.CommandText = strSQL
            dbcmd.CommandType = adCmdText
        ElseIf frmRekapLaporan.dcJenisLaporan.BoundText = "03" Then 'total
            Me.Caption = "Medifirst2000 - Laporan Penerimaan Kasir"
            Set rptPenerimaan = New crPenerimaan
            strSQL = "SELECT * " & _
                " FROM v_S_Lap_PenerimaanKasir " & _
                " WHERE JenisPasien='" & frmRekapLaporan.dcPenjamin.Text & "' AND TglStruk BETWEEN '" & Format(frmRekapLaporan.dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(frmRekapLaporan.dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuanganKasir='" & mstrKdRuangan & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "'"
            dbcmd.CommandText = strSQL
            dbcmd.CommandType = adCmdText
        End If
        With rptPenerimaan
            .Database.AddADOCommand dbConn, dbcmd
            .txtNamaRS.SetText strNNamaRS
            .txtAlamat.SetText strNAlamatRS
            .txtAlamat2.SetText strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
            .txtPeriode.SetText Format(frmRekapLaporan.dtpAwal.Value, "dd MMMM yyyy") & " s/d " & Format(frmRekapLaporan.dtpAkhir.Value, "dd MMMM yyyy")
            .txtKota.SetText strNKotaRS
            .txtKasir.SetText strNmPegawai
            .UsPenjamin.SetUnboundFieldSource ("{ado.NamaPenjamin}")
            .usKelPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
            .usRuangan.SetUnboundFieldSource ("{ado.NamaRuangan}")
            .udTglStruk.SetUnboundFieldSource ("{ado.TglStruk}")
            .usNoPen.SetUnboundFieldSource ("{ado.NoPendaftaran}")
            .usNamaPasien.SetUnboundFieldSource ("{ado.NamaLengkap}")
            .usNoStruk.SetUnboundFieldSource ("{ado.NoStruk}")
            .ucBiaya.SetUnboundFieldSource ("{ado.Biaya}")
            .ucBayar.SetUnboundFieldSource ("{ado.Bayar}")
            .ucPiutang.SetUnboundFieldSource ("{ado.Piutang}")
            .ucCostCharing.SetUnboundFieldSource ("{ado.CostSharing}")
            .ucPembebasan.SetUnboundFieldSource ("{ado.Pembebasan}")
            .ucSisaTagihan.SetUnboundFieldSource ("{ado.SisaTagihan}")
        End With
        CRViewer1.ReportSource = rptPenerimaan
    
    Case "PenerimaanDetail"
        If frmRekapLaporan.dcJenisLaporan.BoundText = "01" Then 'perda a/ tindakan
            Me.Caption = "Medifirst2000 - Laporan Penerimaan Detail Perda"
            Set rptPenerimaanDetail = crPenerimaanDetail
            strSQL = "SELECT * " & _
                " FROM V_S_Lap_PenerimaanKasir_Perda_Detail " & _
                " WHERE JenisPasien='" & frmRekapLaporan.dcPenjamin.Text & "' AND TglStruk BETWEEN '" & Format(frmRekapLaporan.dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(frmRekapLaporan.dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuanganKasir='" & mstrKdRuangan & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "'"
            dbcmd.CommandText = strSQL
            dbcmd.CommandType = adCmdText
        ElseIf frmRekapLaporan.dcJenisLaporan.BoundText = "02" Then 'non perda / alkes
            Me.Caption = "Medifirst2000 - Laporan Penerimaan Detail Kasir Non Perda"
            Set rptPenerimaanDetail = New crPenerimaanDetail
            strSQL = "SELECT * " & _
                " FROM V_S_Lap_PenerimaanKasir_NonPerda_Detail " & _
                " WHERE JenisPasien='" & frmRekapLaporan.dcPenjamin.Text & "' AND TglStruk BETWEEN '" & Format(frmRekapLaporan.dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(frmRekapLaporan.dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuanganKasir='" & mstrKdRuangan & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "'"
            dbcmd.CommandText = strSQL
            dbcmd.CommandType = adCmdText
        ElseIf frmRekapLaporan.dcJenisLaporan.BoundText = "03" Then 'total
            Me.Caption = "Medifirst2000 - Laporan Detail Penerimaan Kasir"
            Set rptPenerimaanDetail = New crPenerimaanDetail
            strSQL = "SELECT * " & _
                " FROM v_S_Lap_PenerimaanKasir_Detail " & _
                " WHERE JenisPasien='" & frmRekapLaporan.dcPenjamin.Text & "' AND TglStruk BETWEEN '" & Format(frmRekapLaporan.dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(frmRekapLaporan.dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuanganKasir='" & mstrKdRuangan & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "'"
            dbcmd.CommandText = strSQL
            dbcmd.CommandType = adCmdText
        End If
        With rptPenerimaanDetail
            .Database.AddADOCommand dbConn, dbcmd
            .txtNamaRS.SetText strNNamaRS
            .txtAlamat.SetText strNAlamatRS
            .txtAlamat2.SetText strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
            .txtPeriode.SetText Format(frmRekapLaporan.dtpAwal.Value, "dd MMMM yyyy") & " s/d " & Format(frmRekapLaporan.dtpAkhir.Value, "dd MMMM yyyy")
            .txtKota.SetText strNKotaRS
            .txtKasir.SetText strNmPegawai
            .UsPenjamin.SetUnboundFieldSource ("{ado.NamaPenjamin}")
            .usKelPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
            .usRuangan.SetUnboundFieldSource ("{ado.NamaRuangan}")
            .udTglStruk.SetUnboundFieldSource ("{ado.TglStruk}")
            .usNoPen.SetUnboundFieldSource ("{ado.NoPendaftaran}")
            .usNamaPasien.SetUnboundFieldSource ("{ado.NamaLengkap}")
            .usNoStruk.SetUnboundFieldSource ("{ado.NoStruk}")
            .usPemeriksaan.SetUnboundFieldSource ("{ado.NamaPelayanan}")
            .ucBiaya.SetUnboundFieldSource ("{ado.Biaya}")
            .ucBayar.SetUnboundFieldSource ("{ado.Bayar}")
            .ucPiutang.SetUnboundFieldSource ("{ado.Piutang}")
            .ucCostCharing.SetUnboundFieldSource ("{ado.CostSharing}")
            .ucPembebasan.SetUnboundFieldSource ("{ado.Pembebasan}")
            .ucSisaTagihan.SetUnboundFieldSource ("{ado.SisaTagihan}")
        End With
        CRViewer1.ReportSource = rptPenerimaanDetail
        
    Case "PembatalanStruk"
        Me.Caption = "Medifirst2000 - Laporan Pembatalan (Retur) Struk"
        Set rptPembatalanStruk = New crPembatalanStruk
        strSQL = "SELECT * " & _
            " FROM V_LaporanReturStrukPelayananRS " & _
            " WHERE JenisPasien='" & frmRekapLaporan.dcPenjamin.Text & "' AND TglStruk BETWEEN '" & Format(frmRekapLaporan.dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(frmRekapLaporan.dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuanganKasir='" & mstrKdRuangan & "' AND IdPegawaiKasir='" & strIDPegawaiAktif & "'"
        dbcmd.CommandText = strSQL
        dbcmd.CommandType = adCmdText
        With rptPembatalanStruk
            .Database.AddADOCommand dbConn, dbcmd
            .txtNamaRS.SetText strNNamaRS
            .txtAlamat.SetText strNAlamatRS
            .txtAlamat2.SetText strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
            .txtPeriode.SetText Format(frmRekapLaporan.dtpAwal.Value, "dd MMMM yyyy") & " s/d " & Format(frmRekapLaporan.dtpAkhir.Value, "dd MMMM yyyy")
            .txtKota.SetText strNKotaRS
            .txtKasir.SetText strNmPegawai
            .UsPenjamin.SetUnboundFieldSource ("{ado.NamaPenjamin}")
            .usKelPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
            .usRuangan.SetUnboundFieldSource ("{ado.NamaRuangan}")
            .udTglStruk.SetUnboundFieldSource ("{ado.TglStruk}")
            .usNoPen.SetUnboundFieldSource ("{ado.NoPendaftaran}")
            .usNamaPasien.SetUnboundFieldSource ("{ado.NamaLengkap}")
            .usNoStruk.SetUnboundFieldSource ("{ado.NoStruk}")
            .ucBiaya.SetUnboundFieldSource ("{ado.Biaya}")
            .ucBayar.SetUnboundFieldSource ("{ado.Bayar}")
            .ucPiutang.SetUnboundFieldSource ("{ado.Piutang}")
            .ucCostCharing.SetUnboundFieldSource ("{ado.CostSharing}")
            .ucPembebasan.SetUnboundFieldSource ("{ado.Pembebasan}")
            .ucSisaTagihan.SetUnboundFieldSource ("{ado.SisaTagihan}")
        End With
        CRViewer1.ReportSource = rptPembatalanStruk
    
    End Select
    
    With CRViewer1
        .EnableGroupTree = False
        .ViewReport
        .Zoom 1
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    With CRViewer1
        .Top = 0
        .Left = 0
        .Height = ScaleHeight
        .Width = ScaleWidth
    End With
End Sub

