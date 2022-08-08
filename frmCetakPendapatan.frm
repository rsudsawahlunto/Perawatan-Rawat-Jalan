VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakPendapatan 
   Caption         =   "Medifirst2000 - Laporan"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakPendapatan.frx":0000
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
Attribute VB_Name = "frmCetakPendapatan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rptLaporanJumlahPendapatan As crLaporanjumlahpendapatan

Private Sub Form_Load()

Dim bln As String

    bln = Format(frmJumlahPasiendanJumlahPendapatan.dtpAwal.Month)
     Select Case bln
        Case "1"
            bln = "Januari"
        Case "2"
            bln = "Februari"
        Case "3"
            bln = "Maret"
        Case "4"
            bln = "April"
        Case "5"
            bln = "Mei"
        Case "6"
            bln = "Juni"
        Case "7"
            bln = "Juli"
        Case "8"
            bln = "Agustus"
        Case "9"
            bln = "September"
        Case "10"
            bln = "Oktober"
        Case "11"
            bln = "November"
        Case "12"
            bln = "Desember"
     End Select



    Me.WindowState = 2
    Screen.MousePointer = vbHourglass
    Set dbcmd = New ADODB.Command
    Set dbcmd.ActiveConnection = dbConn
    
    Me.Caption = "Medifirst2000 - Cetak Laporan Jumlah Pasien dan Jumlah Pendapatan"
    Set rptLaporanJumlahPendapatan = New crLaporanjumlahpendapatan
    
    'sqlnya dari frmdaftar
    
    dbcmd.CommandText = strSQL
    dbcmd.CommandType = adCmdText
    
   
    With rptLaporanJumlahPendapatan
       .Database.AddADOCommand dbConn, dbcmd
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strWebsite & ", " & strEmail
        '.txtPeriode.SetText bln
        '.txtPeriode2.SetText Format(frmjumlahpasiendanpemeriksaan.dtpAwal.Year)
'        .txtData.SetText Format(mdTglAwal, "dd MMMM yyyy HH:mm") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy HH:mm")
        '.txtPeriodeFooter.SetText "Periode : " & Format(mdTglAwal, "dd MMMM yyyy HH:mm") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy HH:mm")
        .Text2.SetText strNNamaRuangan
        .txtPeriode3.SetText Format(frmJumlahPasiendanJumlahPendapatan.dtpAwal.Value, "dd MMMM yyyy HH:mm") & " s/d " & Format(frmJumlahPasiendanJumlahPendapatan.dtpAkhir.Value, "dd MMMM yyyy HH:mm")
        '.txtPeriode3.SetText Format(mdTglAwal, "dd MMMM yyyy HH:mm") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy HH:mm")
        .usJenisPasien.SetUnboundFieldSource ("{ado.IdPenjamin}")
        .unJmlOS.SetUnboundFieldSource ("{ado.JmlOS}")
        .ucPendapatan.SetUnboundFieldSource ("{ado.Pendapatan}")
        
         settingreport rptLaporanJumlahPendapatan, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With
    
    CRViewer1.ReportSource = rptLaporanJumlahPendapatan
    
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
    Call frmJumlahPasiendanJumlahPendapatan.DeleteTable
    Set frmCetakPendapatan = Nothing
End Sub

