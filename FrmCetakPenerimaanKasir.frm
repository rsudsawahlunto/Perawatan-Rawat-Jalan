VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmCetakPenerimaanKasir 
   Caption         =   "Medifirst2000 - Laporan Penerimaan Kasir"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FrmCetakPenerimaanKasir.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   -1  'True
   End
End
Attribute VB_Name = "FrmCetakPenerimaanKasir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New crPenerimaanKasir

Private Sub Form_Load()
    Dim adocomd As New ADODB.Command
    adocomd.ActiveConnection = dbConn
    
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Dim tanggal As String

    Select Case LCase(mLapPerParameter)
        Case "shift"
            adocomd.CommandText = "select * from " & _
                " V_LaporanPenerimaanKasKasir " & _
                " where KdRuangan = '" & mstrKdRuangan & "' AND TglBKM between '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "' "
        Case Else
            adocomd.CommandText = "select * from " & _
                " V_LaporanPenerimaanKasKasir " & _
                " where KdRuangan = '" & mstrKdRuangan & "' AND IdUser = '" & strIDPegawaiAktif & "' AND TglBKM between '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "' "
    End Select
    
    adocomd.CommandType = adCmdText
    Report.Database.AddADOCommand dbConn, adocomd

    If Format(mdTglAwal, "dd MMMM yyyy") = Format(mdTglAkhir, "dd MMMM yyyy") Then
       tanggal = "Tanggal Kunjungan  : " & " " & Format(mdTglAwal, "dd MMMM yyyy")
    Else
       tanggal = "Periode Kunjungan  : " & " " & Format(mdTglAwal, "dd MMMM yyyy") & " S/d " & Format(mdTglAkhir, "dd MMMM yyyy")
    End If

    With Report
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail
        
        .TxtTanggal.SetText tanggal
        
        .usJenisPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
        
        .usNoBKM.SetUnboundFieldSource ("{ado.NoBKM}")
        .udtTglBKM.SetUnboundFieldSource ("{ado.TglBKM}")
        .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
        .usNamaPasien.SetUnboundFieldSource ("{ado.NamaPasien}")
        .usJK.SetUnboundFieldSource ("{ado.JK}")
                
        .ucTotalBiaya.SetUnboundFieldSource ("{ado.TotalBiaya}")
        .ucHutangPenjamin.SetUnboundFieldSource ("{ado.JmlHutangPenjamin}")
        .ucTanggunganRS.SetUnboundFieldSource ("{ado.JmlTanggunganRS}")
        .ucHarusDibayar.SetUnboundFieldSource ("{ado.JmlHarusDibayar}")
        .ucPembebasan.SetUnboundFieldSource ("{ado.JmlPembebasan}")
        .ucAdministrasi.SetUnboundFieldSource ("{ado.Administrasi}")
        .ucJmlBayar.SetUnboundFieldSource ("{ado.JmlBayar}")
        .ucSisaTagihan.SetUnboundFieldSource ("{ado.SisaTagihan}")
        
        .txtNamaKasir.SetText strNmPegawai
                
'        .SelectPrinter sDriver, sPrinter, vbNull
'        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Dim adojenis As New ADODB.Command
    Set adojenis = New ADODB.Command
    adojenis.ActiveConnection = dbConn
    
    Select Case LCase(mLapPerParameter)
        Case "shift"
            adojenis.CommandText = "select * from " & _
                " V_LaporanPenerimaanKasKasir " & _
                " where KdRuangan = '" & mstrKdRuangan & "' AND TglBKM between '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "' "
        Case Else
            adojenis.CommandText = "select * from " & _
                " V_LaporanPenerimaanKasKasir " & _
                " where KdRuangan = '" & mstrKdRuangan & "' AND IdUser = '" & strIDPegawaiAktif & "' AND TglBKM between '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "' "
    End Select
    
    adojenis.CommandType = adCmdText
    Report.Subreport1.OpenSubreport.Database.AddADOCommand dbConn, adojenis
    With Report
        .Subreport1_usRekapJenisPasien.SetUnboundFieldSource ("{ado.jenispasien}")
        .Subreport1_ucTotalBiaya.SetUnboundFieldSource ("{ado.TotalBiaya}")
        .Subreport1_ucHutangPenjamin.SetUnboundFieldSource ("{ado.JmlHutangPenjamin}")
        .Subreport1_ucTanggunganRS.SetUnboundFieldSource ("{ado.JmlTanggunganRS}")
        .Subreport1_ucHarusDibayar.SetUnboundFieldSource ("{ado.JmlHarusDibayar}")
        .Subreport1_ucPembebasan.SetUnboundFieldSource ("{ado.JmlPembebasan}")
        .Subreport1_ucAdministrasi.SetUnboundFieldSource ("{ado.Administrasi}")
        .Subreport1_ucJmlBayar.SetUnboundFieldSource ("{ado.JmlBayar}")
        .Subreport1_ucSisaTagihan.SetUnboundFieldSource ("{ado.SisaTagihan}")
    End With
    CRViewer1.ReportSource = Report
    
    With CRViewer1
        .ViewReport
        .Zoom 1
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmCetakPenerimaanKasir = Nothing
    mLapPerParameter = ""
End Sub
