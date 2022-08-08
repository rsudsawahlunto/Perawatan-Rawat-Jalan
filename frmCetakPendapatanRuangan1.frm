VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakPendapatanRuangan1 
   Caption         =   "Medifirst2000 - Laporan"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakPendapatanRuangan1.frx":0000
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
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCetakPendapatanRuangan1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rptPendapatanRuangan As crPendapatanRuangan1

Private Sub Form_Load()
    Me.WindowState = 2
    Screen.MousePointer = vbHourglass
    Set dbcmd = New ADODB.Command
    Set dbcmd.ActiveConnection = dbConn
    
    Me.Caption = "Medifirst2000 - Cetak Laporan Pendapatan Ruangan"
    Set rptPendapatanRuangan = New crPendapatanRuangan1
    
    'sqlnya dari frmdaftar
    dbcmd.CommandText = strSQL
    dbcmd.CommandType = adCmdText
    With rptPendapatanRuangan
        .Database.AddADOCommand dbConn, dbcmd
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strWebsite & ", " & strEmail
        .txtPeriode.SetText Format(mdTglAwal, "dd MMMM yyyy HH:mm") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy HH:mm")
        .txtPeriodeFooter.SetText "Periode : " & Format(mdTglAwal, "dd MMMM yyyy HH:mm") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy HH:mm")
        .Text1.SetText .Text1.Text
        
        'jenis kelas pav atau non pav
        .unJumlah.SetUnboundFieldSource ("{ado.Jumlah}")
'        .usJenisKasir.SetUnboundFieldSource ("{ado.JenisKelas}")
            
        If frmDaftarPendapatanRuangan.optJenisPasien.Value = True Then
            .usJenisPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
            .txtJenisKasir.SetText "Group By    : Jenis Pasien"
        '    .txtkriteria.SetText "Jenis Pasien"
        ElseIf frmDaftarPendapatanRuangan.optInstalasiAwal.Value = True Then
            .usJenisPasien.SetUnboundFieldSource ("{ado.InstalasiPerujuk}")
           .txtJenisKasir.SetText "Group By     : Instalasi Awal"
        '   .txtkriteria.SetText "Instalasi Asal"
        End If
            .txtRuangan.SetText "Ruangan Pelayanan  : " & mstrNamaRuangan
            .usRuanganPerujuk.SetUnboundFieldSource ("{ado.RuanganPerujuk}")

        If frmDaftarPendapatanRuangan.optTotalBayar.Value = True Then
            .txtNamaKasir.SetText "Kriteria Bayar : Total Bayar"
            .ucBayar.SetUnboundFieldSource ("{ado.TotalBayar}")
        ElseIf frmDaftarPendapatanRuangan.optHutangPenjamin.Value = True Then
            .txtNamaKasir.SetText "Kriteria Bayar : Hutang Penjamin"
            .ucBayar.SetUnboundFieldSource ("{ado.TotalHutangPenjamin}")
            .ucHutangPenjamin.SetUnboundFieldSource ("{ado.JmlHutangPenjamin}")
        ElseIf frmDaftarPendapatanRuangan.optPembebasan.Value = True Then
            .txtNamaKasir.SetText "Kriteria Bayar : Pembebasan"
            .ucBayar.SetUnboundFieldSource ("{ado.TotalPembebasan}")
        ElseIf frmDaftarPendapatanRuangan.optTanggunganRS.Value = True Then
            .txtNamaKasir.SetText "Kriteria Bayar : Tanggungan RS"
            .ucBayar.SetUnboundFieldSource ("{ado.TotalTanggunganRS}")
        ElseIf frmDaftarPendapatanRuangan.optSisaTagihan.Value = True Then
            .txtNamaKasir.SetText "Kriteria Bayar : Sisa Tagihan"
            .ucBayar.SetUnboundFieldSource ("{ado.TotalSisaTagihan}")
        End If
        
        .usPenjamin.SetUnboundFieldSource ("{ado.NamaPenjamin}")
'        .usKelasPelayanan.SetUnboundFieldSource ("{ado.JenisKelas}")
        .usRuangan.SetUnboundFieldSource ("{ado.NamaRuangan}")
        .usKomponenUnit.SetUnboundFieldSource ("{ado.TindakanPelayanan}")
        .ucTarif.SetUnboundFieldSource ("{ado.Tarif}")
        .usKelas.SetUnboundFieldSource ("{ado.DeskKelas}")
'        .unJumlah.SetUnboundFieldSource ("{ado.Jumlah}")
        .usKomponenTarif.SetUnboundFieldSource ("{ado.NamaKomponen}")
        
         If sUkuranKertas = "" Then
         sUkuranKertas = "5"
         sOrientasKertas = "2"
         sDuplex = "0"
         End If
        
         settingreport rptPendapatanRuangan, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
        End With
    
    CRViewer1.ReportSource = rptPendapatanRuangan
    With CRViewer1
        .EnableGroupTree = False
        .ViewReport
        .Zoom 1
    End With
    Screen.MousePointer = vbDefault
    
    Set dbcmd = Nothing
End Sub

Private Sub Form_Resize()
    With CRViewer1
        .Top = 0
        .Left = 0
        .Height = ScaleHeight
        .Width = ScaleWidth
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakPendapatanRuangan = Nothing
End Sub

