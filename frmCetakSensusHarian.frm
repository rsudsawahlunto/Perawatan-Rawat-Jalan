VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakSensusHarian 
   Caption         =   "LAPORAN SENSUS HARIAN"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7110
   ScaleWidth      =   5910
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
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCetakSensusHarian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crCetakSensusHarian

Private Sub Form_Load()

    Dim adocomd As New ADODB.Command
    adocomd.ActiveConnection = dbConn

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
        
    adocomd.CommandText = strSQL
    adocomd.CommandType = adCmdText

    With Report
        .Database.AddADOCommand dbConn, adocomd
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strWebsite & ", " & strEmail
        
        .txtTanggalPilih.SetText Format(frmLapSensusHarian.dtpAwal.Value, "yyyy/MM/dd")
        .txtTanggalPilih2.SetText Format(frmLapSensusHarian.dtpAkhir.Value, "yyyy/MM/dd")
        .txtPelapor.SetText strNmPegawai
                
        .txtjudul.SetText "LAPORAN SENSUS PASIEN"
        
        .usJudul.SetUnboundFieldSource ("{ado.judul}")
        .usSubJudul.SetUnboundFieldSource ("{ado.subJudul}")
        .udtTglSensus.SetUnboundFieldSource ("{ado.TglSensus}")
        .usNoPendaftaran.SetUnboundFieldSource ("{ado.NoPendaftaran}")
        .udtTglSensus.SetUnboundFieldSource ("{ado.TglSensus}")
        .usRecord.SetUnboundFieldSource ("{ado.Record}")
        .usRuanganPelayanan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        
        strSQL = " select SUM(cast(record as int)) as LD from V_SensusHarianFormulirRp2PasienKeluarHidupLamaDiRwt" & _
             " where TglSensus between " & _
             " '" & Format(frmLapSensusHarian.dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' and " & _
             " '" & Format(frmLapSensusHarian.dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' "
        Set rsb = Nothing
        rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        
        .txtLD.SetText rsb("LD").Value
    
    End With
    
    If sUkuranKertas = "" Then
    sUkuranKertas = "5"
    sOrientasKertas = "2"
    sDuplex = "0"
    End If
    
    settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    
    CRViewer1.ReportSource = Report
    
    With CRViewer1
        .EnableGroupTree = False
        .ViewReport
        .Zoom 1
    End With
    
    Screen.MousePointer = vbDefault
    
    Set adocomd = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakSensusHarian = Nothing
End Sub
