VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCtkDaftarPasien 
   Caption         =   "Cetak Dokumen Rekam Medis Pasien"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   Icon            =   "frmCtkDaftarPasien.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   10920
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
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
Attribute VB_Name = "frmCtkDaftarPasien"
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

    Report.TxtTanggal.SetText ("Periode  : " & Format(frmDaftarPasienRJ.dtpAwal.Value, "dd MMMM yyyy HH:mm") & " s/d " & Format(frmDaftarPasienRJ.dtpAkhir, "dd MMMM yyyy HH:mm"))
    If frmDaftarPasienRJ.optPasienPoliklinik.Value = True Then
        adocomd.ActiveConnection = dbConn
        adocomd.CommandText = "select * from V_DaftarPasienLamaRJ where ([Nama Pasien] like '%" & frmDaftarPasienRJ.txtParameter.Text & "%' OR NoCM like '%" & frmDaftarPasienRJ.txtParameter.Text & "%') and Ruangan='" & strNNamaRuangan & "' and TglMasuk between '" & Format(frmDaftarPasienRJ.dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(frmDaftarPasienRJ.dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' AND JenisPasien LIKE '%" & frmDaftarPasienRJ.dcJenisPasien.Text & "%' AND Kelas LIKE '%" & frmDaftarPasienRJ.dcKelas.Text & "%'"
        adocomd.CommandType = adCmdUnknown
        Report.Database.AddADOCommand dbConn, adocomd
        With Report
            .usNoRegistrasi.SetUnboundFieldSource ("{ado.NoPendaftaran}")
            .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
            .usNamaPasien.SetUnboundFieldSource ("{ado.Nama Pasien}")
            .usJK.SetUnboundFieldSource ("{ado.JK}")
            .usUmur.SetUnboundFieldSource ("{ado.Umur}")
            .usJenisPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
            .usKelas.SetUnboundFieldSource ("{ado.Kelas}")
            .udTglMasuk.SetUnboundFieldSource ("{ado.TglMasuk}")
            .usDokterPemeriksa.SetUnboundFieldSource ("{ado.Dokter Pemeriksa}")
        End With
        Report.txtDaftarKamar.SetText (frmDaftarPasienRJ.optPasienPoliklinik.Caption & " " & strNNamaRS)
    ElseIf frmDaftarPasienRJ.optRujukan.Value = True Then
        adocomd.ActiveConnection = dbConn
        adocomd.CommandText = "select * from V_DaftarPasienKonsul where ([Nama Pasien] like '%" & frmDaftarPasienRJ.txtParameter.Text & "%' OR NoCM like '%" & frmDaftarPasienRJ.txtParameter.Text & "%') and KdRuanganTujuan='" & strNKdRuangan & "' and TglDirujuk between '" & Format(frmDaftarPasienRJ.dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(frmDaftarPasienRJ.dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' AND StatusPeriksa = '" & frmDaftarPasienRJ.dcStatusPeriksa.Text & "' AND JenisPasien LIKE '%" & frmDaftarPasienRJ.dcJenisPasien.Text & "%' AND Kelas LIKE '%" & frmDaftarPasienRJ.dcKelas.Text & "%'"
        adocomd.CommandType = adCmdUnknown
        Report.Database.AddADOCommand dbConn, adocomd
        With Report
            .usNoRegistrasi.SetUnboundFieldSource ("{ado.NoPendaftaran}")
            .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
            .usNamaPasien.SetUnboundFieldSource ("{ado.Nama Pasien}")
            .usJK.SetUnboundFieldSource ("{ado.JK}")
            .usUmur.SetUnboundFieldSource ("{ado.Umur}")
            .usJenisPasien.SetUnboundFieldSource ("{ado.StatusPeriksa}")
            .usKelas.SetUnboundFieldSource ("{ado.Ruangan Perujuk}")
            .udTglMasuk.SetUnboundFieldSource ("{ado.TglDirujuk}")
            .usDokterPemeriksa.SetUnboundFieldSource ("{ado.Dokter Perujuk}")
            .Text1.SetText "Dokter Perujuk"
            .Text9.SetText "Ruang Rujukan"
            .Text7.SetText "Status"
            .Text7.HorAlignment = crHorCenterAlign
            .usJenisPasien.HorAlignment = crHorCenterAlign
        End With
        Report.txtDaftarKamar.SetText (frmDaftarPasienRJ.optRujukan.Caption & " " & strNNamaRS)
    End If
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
    Set frmCtkDaftarPasien = Nothing
End Sub
