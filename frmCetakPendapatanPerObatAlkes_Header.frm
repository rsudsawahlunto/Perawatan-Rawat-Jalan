VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakPendapatanPerObatAlkes_Header 
   Caption         =   "Medifirst2000 - Laporan"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakPendapatanPerObatAlkes_Header.frx":0000
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
Attribute VB_Name = "frmCetakPendapatanPerObatAlkes_Header"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rptPendapatanPerObatAlkes As crPendapatanPerSMF_Header

Private Sub Form_Load()
    Me.WindowState = 2
    Screen.MousePointer = vbHourglass
    Set dbcmd = New ADODB.Command
    Set dbcmd.ActiveConnection = dbConn
    
    Me.Caption = "Medifirst2000 - Cetak Laporan Pendapatan Obat Alkes-->Unit"
    Set rptPendapatanPerObatAlkes = New crPendapatanPerSMF_Header
    
    'sql diambil dari frmdaftar
    dbcmd.CommandText = strSQL
    dbcmd.CommandType = adCmdText
    With rptPendapatanPerObatAlkes
        .Database.AddADOCommand dbConn, dbcmd
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strWebsite & ", " & strEmail
        .txtPeriode.SetText Format(mdTglAwal, "dd MMMM yyyy HH:mm") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy HH:mm")
        .txtPeriodeFooter.SetText "Periode : " & Format(mdTglAwal, "dd MMMM yyyy HH:mm") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy HH:mm")
       .txtJenisPasien.Suppress = True

        .usJenisKasir.SetUnboundFieldSource ("{ado.JenisKelas}")
        .txtNamaKasir.SetText "Status Laporan : " & mstrKriteria & ""
        .usKomponenUnit.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
                
        Select Case mstrKriteria
            Case "TotalBayar"
                .ucBayar.SetUnboundFieldSource ("{ado.TotalBayar}")
            Case "TotalHutangPenjamin"
                .ucBayar.SetUnboundFieldSource ("{ado.TotalHutangPenjamin}")
            Case "TotalPembebasan"
                .ucBayar.SetUnboundFieldSource ("{ado.TotalPembebasan}")
            Case "TotalTanggunganRS"
                .ucBayar.SetUnboundFieldSource ("{ado.TotalTanggunganRS}")
            Case "TotalSisaTagihan"
                .ucBayar.SetUnboundFieldSource ("{ado.TotalSisaTagihan}")
        End Select
        .txtJudul.SetText "LAPORAN PENDAPATAN OBAT & ALKES --> RUANGAN"
        .UsPenjamin.SetUnboundFieldSource ("{ado.Penjamin}")
        .txtkriteria.SetText "Jenis Pasien"
        .txtRuangan.SetText "Komponen Unit"
        .txtkomponenUnit.SetText "Ruangan"
        .usJenisPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
'        .usRuangan.SetUnboundFieldSource ("{ado.RuanganAsal}")
        .usRuangan.SetUnboundFieldSource ("{ado.KomponenUnit}")
        .usKomponenTarif.SetUnboundFieldSource ("{ado.KomponenTarif}")
    End With
    
    CRViewer1.ReportSource = rptPendapatanPerObatAlkes
    
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
    Set frmCetakPendapatanPerObatAlkes_Header = Nothing
End Sub
