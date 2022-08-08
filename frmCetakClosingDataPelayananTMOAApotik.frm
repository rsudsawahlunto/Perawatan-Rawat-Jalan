VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakClosingDataPelayananTMOAApotik 
   Caption         =   "Cetak Closing Data Pelayanan TMOA Apotik"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   Icon            =   "frmCetakClosingDataPelayananTMOAApotik.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   5880
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
Attribute VB_Name = "frmCetakClosingDataPelayananTMOAApotik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ReportTindakanMedis As crCetakClosingDataPelayananTMOAApotik2
Dim ReportObatAlkes As crCetakClosingDataPelayananTMOAApotik
Dim ReportApotik As crCetakClosingDataPelayananTMOAApotik3

Private Sub Form_Load()
    Me.WindowState = 2
    Screen.MousePointer = vbHourglass
    Set dbcmd = New ADODB.Command
    Set dbcmd.ActiveConnection = dbConn

    Me.Caption = "Medifirst2000 - Cetak Closing Data Pelayanan TM,OA,Apotik"
    Set ReportTindakanMedis = New crCetakClosingDataPelayananTMOAApotik2
    Set ReportObatAlkes = New crCetakClosingDataPelayananTMOAApotik
    Set ReportApotik = New crCetakClosingDataPelayananTMOAApotik3

    'sqlnya dari frmdaftar
    dbcmd.CommandText = strSQL
    dbcmd.CommandType = adCmdText

    If frmClosingDataPelayananTM_OA_Apotik.optTindakanMedis.Value = True Then
        With ReportTindakanMedis
            .Database.AddADOCommand dbConn, dbcmd
            .txtNamaRS.SetText strNNamaRS
            .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
            .txtAlamat2.SetText strWebsite & ", " & strEmail
            .txtPeriode.SetText Format(mdTglAwal, "dd MMMM yyyy HH:mm") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy HH:mm")
            .txtRuangan.SetText mstrNamaRuangan

            .udtTglPelayanan.SetUnboundFieldSource ("{ado.TglPelayanan}")
            .usJenisPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
            .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
            .usNoPendaftaran.SetUnboundFieldSource ("{ado.NoPendaftaran}")
            .usNamaPasien.SetUnboundFieldSource ("{ado.NamaPasien}")
            .usKelas.SetUnboundFieldSource ("{ado.Kelas}")
            .usJenisItem.SetUnboundFieldSource ("{ado.JenisItem}")
            .usNamaItem.SetUnboundFieldSource ("{ado.NamaItem}")
            .unJumlahItem.SetUnboundFieldSource ("{ado.JmlItem}")
            .ucHargaSatuan.SetUnboundFieldSource ("{ado.HargaSatuan}")
            .ucHargaCito.SetUnboundFieldSource ("{ado.HargaCito}")
            .ucTotalBiaya.SetUnboundFieldSource ("{ado.TotalBiaya}")
            .usNoClosing.SetUnboundFieldSource ("{ado.NoClosing}")
            .UsDokter.SetUnboundFieldSource ("{ado.DokterOperator}")

            With CRViewer1
                .ReportSource = ReportTindakanMedis
                .EnableGroupTree = False
                .ViewReport
                .Zoom 1
            End With

        End With
    ElseIf frmClosingDataPelayananTM_OA_Apotik.optObatAlKes.Value = True Then
        With ReportObatAlkes
            .Database.AddADOCommand dbConn, dbcmd
            .txtNamaRS.SetText strNNamaRS
            .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
            .txtAlamat2.SetText strWebsite & ", " & strEmail
            .txtPeriode.SetText Format(mdTglAwal, "dd MMMM yyyy HH:mm") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy HH:mm")
            .txtRuangan.SetText mstrNamaRuangan

            .udtTglPelayanan.SetUnboundFieldSource ("{ado.TglPelayanan}")
            .usJenisPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
            .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
            .usNoPendaftaran.SetUnboundFieldSource ("{ado.NoPendaftaran}")
            .usNamaPasien.SetUnboundFieldSource ("{ado.NamaPasien}")
            .usKelas.SetUnboundFieldSource ("{ado.Kelas}")
            .usJenisItem.SetUnboundFieldSource ("{ado.JenisItem}")
            .usNamaItem.SetUnboundFieldSource ("{ado.NamaItem}")
            .unJumlahItem.SetUnboundFieldSource ("{ado.JmlItem}")
            .ucHargaSatuan.SetUnboundFieldSource ("{ado.HargaSatuan}")
            .ucHargaService.SetUnboundFieldSource ("{ado.HargaService}")
            .ucAdministrasi.SetUnboundFieldSource ("{ado.Administrasi}")
            .ucTotalBiaya.SetUnboundFieldSource ("{ado.TotalBiaya}")
            .usNoClosing.SetUnboundFieldSource ("{ado.NoClosing}")

            With CRViewer1
                .ReportSource = ReportObatAlkes
                .EnableGroupTree = False
                .ViewReport
                .Zoom 1
            End With

        End With
    ElseIf frmClosingDataPelayananTM_OA_Apotik.optApotik.Value = True Then
        With ReportApotik
            .Database.AddADOCommand dbConn, dbcmd
            .txtNamaRS.SetText strNNamaRS
            .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
            .txtAlamat2.SetText strWebsite & ", " & strEmail
            .txtPeriode.SetText Format(mdTglAwal, "dd MMMM yyyy HH:mm") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy HH:mm")
            .txtRuangan.SetText mstrNamaRuangan
            .udtTglStruk.SetUnboundFieldSource ("{ado.TglStruk}")
            .usJenisPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
            .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
            .usNoPendaftaran.SetUnboundFieldSource ("{ado.NoPendaftaran}")
            .usNamaPasien.SetUnboundFieldSource ("{ado.NamaPasien}")
            .usNoStruk.SetUnboundFieldSource ("{ado.NoStruk}")
            .usJenisItem.SetUnboundFieldSource ("{ado.JenisItem}")
            .usNamaItem.SetUnboundFieldSource ("{ado.NamaItem}")
            .unJumlahItem.SetUnboundFieldSource ("{ado.JmlItem}")
            .ucHargaSatuan.SetUnboundFieldSource ("{ado.HargaSatuan}")
            .ucHargaService.SetUnboundFieldSource ("{ado.HargaService}")
            .ucAdministrasi.SetUnboundFieldSource ("{ado.Administrasi}")
            .ucTotalBiaya.SetUnboundFieldSource ("{ado.TotalBiaya}")
            .usNoClosing.SetUnboundFieldSource ("{ado.NoClosing}")

            With CRViewer1
                .ReportSource = ReportApotik
                .EnableGroupTree = False
                .ViewReport
                .Zoom 1
            End With

        End With

    End If

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
    Set frmCetakClosingDataPelayananTMOAApotik = Nothing
End Sub
