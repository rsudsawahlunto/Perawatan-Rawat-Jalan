VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakDataPelayananTMOAApotik 
   Caption         =   "Cetak Data Pelayanan TMOA Apotik"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   Icon            =   "frmCetakDataPelayananTMOAApotik.frx":0000
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
Attribute VB_Name = "frmCetakDataPelayananTMOAApotik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ReportTindakanMedis As crCetakDataPelayananTMOAApotik2
Dim ReportObatAlkes As crCetakDataPelayananTMOAApotik
Dim ReportApotik As crCetakClosingDataPelayananTMOAApotik3

Private Sub Form_Load()
    Me.WindowState = 2
    Screen.MousePointer = vbHourglass
    Set dbcmd = New ADODB.Command
    Set dbcmd.ActiveConnection = dbConn

    Me.Caption = "Medifirst2000 - Cetak Data Pelayanan TM,OA,Apotik"
    Set ReportTindakanMedis = New crCetakDataPelayananTMOAApotik2
    Set ReportObatAlkes = New crCetakDataPelayananTMOAApotik
    Set ReportApotik = New crCetakClosingDataPelayananTMOAApotik3

    If frmClosingDataPelayananTM_OA_Apotik.optTindakanMedis.Value = True Then
        strSQL = "SELECT TglPelayanan, JenisPasien, NoCM, NamaPasien, NamaItem, TotalBiaya, DokterOperator, JmlHutangPenjamin, JmlTanggunganRS," & _
        " TotalBayar From V_DaftarDataPelayananTMForClosing_Singkat" & _
        " WHERE DokterOperator like '%" & frmClosingDataPelayananTM_OA_Apotik.dcNamaDokter.Text & "%' AND TglPelayanan BETWEEN '" & Format(frmClosingDataPelayananTM_OA_Apotik.dtpAwal.Value, "yyyy/MM/dd hh:mm:00") & "' AND '" & Format(frmClosingDataPelayananTM_OA_Apotik.dtpAkhir.Value, "yyyy/MM/dd hh:mm:59") & "' " & _
        " AND JenisPasien like '%" & frmClosingDataPelayananTM_OA_Apotik.dcJenisPasien.Text & "%' AND Kelas like '%" & frmClosingDataPelayananTM_OA_Apotik.dcKelas.Text & "%' AND JenisItem like '" & frmClosingDataPelayananTM_OA_Apotik.dcJenisItem.Text & "%' AND NamaItem like '%" & frmClosingDataPelayananTM_OA_Apotik.dcNamaItem.Text & "%'  AND AsalPasien LIKE '%" & frmClosingDataPelayananTM_OA_Apotik.dcAsalPasien.Text & "%' AND KdRuangan = '" & mstrKdRuangan & "'  "
    ElseIf frmClosingDataPelayananTM_OA_Apotik.optObatAlKes.Value = True Then
        strSQL = "SELECT TglPelayanan, JenisPasien, NoCM,NoPendaftaran, NamaPasien, JenisItem, NamaItem, TotalBiaya, JmlHutangPenjamin, JmlTanggunganRS," & _
        " TotalBayar From V_DaftarDataPelayananOAForClosing_Singkat" & _
        " WHERE TglPelayanan BETWEEN '" & Format(frmClosingDataPelayananTM_OA_Apotik.dtpAwal.Value, "yyyy/MM/dd hh:mm:00") & "' AND '" & Format(frmClosingDataPelayananTM_OA_Apotik.dtpAkhir.Value, "yyyy/MM/dd hh:mm:59") & "' " & _
        " AND JenisPasien like '%" & frmClosingDataPelayananTM_OA_Apotik.dcJenisPasien.Text & "%' AND Kelas like '%" & frmClosingDataPelayananTM_OA_Apotik.dcKelas.Text & "%' AND JenisItem like '" & frmClosingDataPelayananTM_OA_Apotik.dcJenisItem.Text & "%' AND NamaItem like '%" & frmClosingDataPelayananTM_OA_Apotik.dcNamaItem.Text & "%'  AND AsalPasien LIKE '%" & frmClosingDataPelayananTM_OA_Apotik.dcAsalPasien.Text & "%' AND KdRuangan = '" & mstrKdRuangan & "'  "

    End If

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
            .usNamaPasien.SetUnboundFieldSource ("{ado.NamaPasien}")
            .usNamaItem.SetUnboundFieldSource ("{ado.NamaItem}")
            .ucTotalBiaya.SetUnboundFieldSource ("{ado.TotalBiaya}")
            .UsDokter.SetUnboundFieldSource ("{ado.DokterOperator}")
            .ucJmlTanggunganPenjamin.SetUnboundFieldSource ("{ado.JmlHutangPenjamin}")
            .ucTanggunganRS.SetUnboundFieldSource ("{ado.JmlTanggunganRS}")
            .ucTotalBayar.SetUnboundFieldSource ("{ado.TotalBayar}")

            With CRViewer1
                .ReportSource = ReportTindakanMedis
                .EnableGroupTree = False
                .ViewReport
                .Zoom 1
            End With

        End With
    End If

    If frmClosingDataPelayananTM_OA_Apotik.optObatAlKes.Value = True Then
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
            .usJenisItem.SetUnboundFieldSource ("{ado.JenisItem}")
            .usNamaItem.SetUnboundFieldSource ("{ado.NamaItem}")
            .ucTotalBiaya.SetUnboundFieldSource ("{ado.TotalBiaya}")
            .ucJmlTanggunganPenjamin.SetUnboundFieldSource ("{ado.JmlHutangPenjamin}")
            .ucTanggunganRS.SetUnboundFieldSource ("{ado.JmlTanggunganRS}")
            .ucTotalBayar.SetUnboundFieldSource ("{ado.TotalBayar}")

            With CRViewer1
                .ReportSource = ReportObatAlkes
                .EnableGroupTree = False
                .ViewReport
                .Zoom 1
            End With

        End With
    End If

    If frmClosingDataPelayananTM_OA_Apotik.optApotik.Value = True Then
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
    Set frmCetakDataPelayananTMOAApotik = Nothing
End Sub

