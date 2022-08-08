VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_cetak_RincianBiaya 
   Caption         =   "Cetak Laporan Cetak Rincian Biaya Sementara"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   Icon            =   "frm_cetak_RincianBiaya.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   5805
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
Attribute VB_Name = "frm_cetak_RincianBiaya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New cr_RincianBiaya

Private Sub Form_Load()
    On Error GoTo errLoad
    Dim adocomd As New ADODB.Command

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    strSQL = "select * " & _
    " from V_JudulRincianBiayaSementara where " & _
    " nopendaftaran ='" & mstrNoPen & "'"
    Call msubRecFO(rs, strSQL)

    With Report
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail

        .txtNoPendaftaran.SetText rs("nopendaftaran")
        .txttglpendaftaran.SetText rs("TglPendaftaran")
        .txtNoCM.SetText rs("nocm")
        .txtnmpasien.SetText rs("nama pasien") & " / " & IIf(rs("JK").Value = "P", "Wanita", "Pria")
        .txtUmur.SetText rs("umur")
        .txtAlamat.SetText IIf(IsNull(rs("alamat")), "-", rs("alamat"))
        .txtklpkpasien.SetText rs("jenispasien")
        .txtPenjamin.SetText IIf(IsNull(rs("NamaPenjamin")), "Sendiri", rs("NamaPenjamin"))
        .txtNamaRuangan.SetText mstrNamaRuangan

        .txtPrintTglBKM.SetText strNKotaRS & ", " & Format(Now, "dd MMMM yyyy")
    End With

    Set dbcmd = New ADODB.Command
    dbcmd.CommandText = "sELECT * FROM V_RincianTotalDetailBiayaPelayanan " _
    & "WHERE (NoPendaftaran = '" & mstrNoPen & "')"
    dbcmd.CommandType = adCmdText
    Report.Database.AddADOCommand dbConn, dbcmd
    With Report
        .udTanggal.SetUnboundFieldSource ("{Ado.tglpelayanan}")
        .usruang.SetUnboundFieldSource ("{Ado.ruangan}")
        .usJenisPelayanan.SetUnboundFieldSource ("{Ado.jenis_item}")
        .usKelas.SetUnboundFieldSource ("{Ado.kelas}")
        .unQty.SetUnboundFieldSource ("{Ado.jml_item}")
        .ucBiayaSatuan.SetUnboundFieldSource ("{Ado.Tarif}") '("{Ado.harga_item}")
        .ucTarifCito.SetUnboundFieldSource ("{Ado.TarifCITO}")
        .ucTarifTotal.SetUnboundFieldSource ("{Ado.BiayaTotal}")
        .ustindakan.SetUnboundFieldSource ("{Ado.nama_item}")

        strSQL = "SELECT SUM(TotalHutangPenjamin) as TotJmlTanggungan, SUM(TotalTanggunganRS) as TotTanggunganRS, SUM(TotalHarusDibayar) as TotHrsDibyrPsn, SUM(BiayaTotal) as TotBiayaTotal " & _
        " FROM V_RincianTotalDetailBiayaPelayanan " & _
        " WHERE (NoPendaftaran = '" & mstrNoPen & "')"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        If rs.EOF = True Then
            .txtPembebasan.SetText 0
            .txtTanggunganRS.SetText 0
            .txtTotalBiaya.SetText 0
            .txtTanggungan.SetText 0
            .txtBayar.SetText 0
        Else
            .txtPembebasan.SetText 0
            If mblnAdmin = True Then
                .txtTanggunganRS.SetText IIf(rs("TotTanggunganRS").Value = 0, 0, Format(rs("TotTanggunganRS").Value, "#,###"))
                .txtTanggungan.SetText IIf(rs("TotJmlTanggungan").Value = 0, 0, Format(rs("TotJmlTanggungan").Value, "#,###"))
            Else
                .txtTanggunganRS.SetText "-"
                .txtTanggungan.SetText "-"
            End If

            .txtTotalBiaya.SetText IIf(rs("TotBiayaTotal").Value = 0, 0, Format(rs("TotBiayaTotal").Value, "#,###"))
            .txtBayar.SetText IIf(rs("TotHrsDibyrPsn").Value = 0, 0, Format(rs("TotHrsDibyrPsn").Value, "#,###"))
            If IsNull(rs("TotBiayaTotal").Value) Then
                .txtTerbilang.SetText NumToText(0)
            Else
                .txtTerbilang.SetText NumToText(IIf(rs("TotBiayaTotal").Value = 0, 0, CCur(rs("TotBiayaTotal").Value)))
            End If
        End If
        .txtPetugasKasir.SetText strNmPegawai
        .txtIdPetugas.SetText noidpegawai
    End With
    Dim adojenis As New ADODB.Command
    Set adojenis = New ADODB.Command
    adojenis.CommandText = "select * from " & _
    " V_RincianBiayaPelayananTMPerKomponenNullBS " & _
    " where NoPendaftaran = '" & mstrNoPen & "'"

    adojenis.CommandType = adCmdText
    Report.Subreport1.OpenSubreport.Database.AddADOCommand dbConn, adojenis
    With Report
        .Subreport1_usNamaKomponen.SetUnboundFieldSource ("{ado.KomponenTarif}")
        .Subreport1_usNamaPelayanan.SetUnboundFieldSource ("{ado.Nama_Item}")
        .Subreport1_ucJasaKomponen.SetUnboundFieldSource ("{ado.Harga_Item}")
        .Subreport1_usRuangan.SetUnboundFieldSource ("{ado.Ruangan}")
    End With
    CRViewer1.ReportSource = Report

    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 1
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmDaftarPasienRJ.PostingHutangPenjaminPasien_AU("U")
    Set frm_cetak_RincianBiaya = Nothing
    Set rs = Nothing
    mbolCetakJasaDokter = False
End Sub

