VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakRekapKunjunganPI 
   Caption         =   "Medifirst2000 - Laporan"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
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
Attribute VB_Name = "frmCetakRekapKunjunganPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rptRekapKunjunganPI As crRekapKunjunganPI

Private Sub Form_Load()
    Me.WindowState = 2
    Screen.MousePointer = vbHourglass
    Set dbcmd = New ADODB.Command
    Set dbcmd.ActiveConnection = dbConn
    
    Me.Caption = "Medifirst2000 - Cetak Laporan Kas Masuk"
    Set rptRekapKunjunganPI = New crRekapKunjunganPI
    strSQL = " SELECT * " & _
             " FROM V_RekapitulasiPasienBRujukanInternal " & _
             " WHERE TglPendaftaran BETWEEN '" & Format(frmRekapKunjunganPI.dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(frmRekapKunjunganPI.dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "'" & _
             " AND KdRuangan = '" & mstrKdRuangan & "'"
    dbcmd.CommandText = strSQL
    dbcmd.CommandType = adCmdText
    With rptRekapKunjunganPI
        .Database.AddADOCommand dbConn, dbcmd
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS
        .txtAlamat2.SetText strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtPeriode.SetText Format(frmRekapKunjunganPI.dtpAwal.Value, "dd MMMM yyyy") & " s/d " & Format(frmRekapKunjunganPI.dtpAkhir.Value, "dd MMMM yyyy")
        .txtPetugas.SetText strNmPegawai
        .txtRuangan.SetText strNNamaRuangan
        .usRuangPerujuk.SetUnboundFieldSource ("{ado.Ruang Perujuk}")
        .usJenisPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
        .unJumlahPria.SetUnboundFieldSource ("{ado.JmlPasienPria}")
        .unJumlahWanita.SetUnboundFieldSource ("{ado.JmlPasienWanita}")
         settingreport rptRekapKunjunganPI, sPrinter, sDriver, sUkuranKertas, sDuplex, crPortrait
    End With
    CRViewer1.ReportSource = rptRekapKunjunganPI
    
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

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakRekapKunjunganPI = Nothing
End Sub







