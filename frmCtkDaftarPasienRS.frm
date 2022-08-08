VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCtkDaftarPasienRS 
   Caption         =   "Cetak Dokumen Rekam Medis Pasien"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10770
   Icon            =   "frmCtkDaftarPasienRS.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6930
   ScaleWidth      =   10770
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      DisplayGroupTree=   0   'False
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
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCtkDaftarPasienRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crDaftarPasien

Private Sub Form_Load()
    Dim adocomd As New ADODB.Command
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Call openConnection
    Me.WindowState = 2
    Report.txtNamaRS.SetText strNNamaRS
    Report.txtAlamatRS.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
    Report.txtWebsiteRS.SetText strWebsite & ", " & strEmail

    Report.TxtTanggal.SetText ("Periode  : " & Format(frmDaftarPasienRJRIIGD.dtpAwal.Value, "dd MMMM yyyy HH:mm") & " s/d " & Format(frmDaftarPasienRJRIIGD.dtpAkhir, "dd MMMM yyyy HH:mm"))
    Report.txtDaftarKamar.SetText "DAFTAR PASIEN RSUD " & UCase(strNKotaRS)
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    adocomd.CommandType = adCmdUnknown
    Report.Database.AddADOCommand dbConn, adocomd
    With Report
        .usNoRegistrasi.SetUnboundFieldSource ("{ado.NoPendaftaran}")
        .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
        .usNamaPasien.SetUnboundFieldSource ("{ado.NamaPasien}")
        .usJK.SetUnboundFieldSource ("{ado.JK}")
        .usUmur.SetUnboundFieldSource ("{ado.Umur}")
        .usJenisPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
        .usKelas.SetUnboundFieldSource ("{ado.Kelas}")
        .udTglMasuk.SetUnboundFieldSource ("{ado.TglMasuk}")
        .usDokterPemeriksa.SetUnboundFieldSource ("{ado.DokterPemeriksa}")
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
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
    Set frmCtkDaftarPasienRS = Nothing
End Sub
