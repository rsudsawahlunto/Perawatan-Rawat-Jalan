VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakDaftarDokumenRekamMedisPasien 
   Caption         =   "Cetak Daftar Dokumen Rekam Medis Pasien"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   Icon            =   "frmCetakDaftarDokumenRekamMedisPasien.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   5865
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
Attribute VB_Name = "frmCetakDaftarDokumenRekamMedisPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crDaftarDokumenRekamMedisPasien

Private Sub Form_Load()
    On Error GoTo errLoad
    Dim adocomd As New ADODB.Command

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    With Report
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strWebsite & ", " & strEmail

        Set adocomd.ActiveConnection = dbConn
        adocomd.CommandText = strSQL
        adocomd.CommandType = adCmdUnknown
        .Database.AddADOCommand dbConn, adocomd

        If frmDaftarDokumenRekamMedisPasien.optTglMasuk.Value = True Then
            .txtTanggal.SetText "Periode Tanggal Masuk : " & Format(mdTglAwal, "dd MMMM yyyy hh:MM:ss") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy hh:MM:ss")
        ElseIf frmDaftarDokumenRekamMedisPasien.optTglPulang.Value = True Then
            .txtTanggal.SetText "Periode Tanggal Pulang : " & Format(mdTglAwal, "dd MMMM yyyy hh:MM:ss") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy hh:MM:ss")
        ElseIf frmDaftarDokumenRekamMedisPasien.optTglKirim.Value = True Then
            .txtTanggal.SetText "Periode Tanggal Kirim : " & Format(mdTglAwal, "dd MMMM yyyy hh:MM:ss") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy hh:MM:ss")
        ElseIf frmDaftarDokumenRekamMedisPasien.optTglTerima.Value = True Then
            .txtTanggal.SetText "Periode Tanggal Terima : " & Format(mdTglAwal, "dd MMMM yyyy hh:MM:ss") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy hh:MM:ss")
        End If
        .usNoPendaftaran.SetUnboundFieldSource ("{ado.NoPendaftaran}")
        .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
        .usNamaPasien.SetUnboundFieldSource ("{ado.NamaPasien}")
        .usRuangPelayanan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .udtTglMasuk.SetUnboundFieldSource ("{ado.TglMasuk}")
        .udtTglPulang.SetUnboundFieldSource ("{ado.TglPulang}")
        .udtTglKirim.SetUnboundFieldSource ("{ado.TglKirim}")
        .udtTglTerima.SetUnboundFieldSource ("{ado.TglTerima}")
        .usRuangPengirim.SetUnboundFieldSource ("{ado.RuanganPengirim}")

        CRViewer1.ReportSource = Report
    End With

    With CRViewer1
        .EnableGroupTree = False
        .Zoom 1  ' Set the zoom level to fit the page width to the viewer window
        .ViewReport ' Set the viewer to view the report
    End With
    Screen.MousePointer = vbDefault

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakDaftarDokumenRekamMedisPasien = Nothing
End Sub
