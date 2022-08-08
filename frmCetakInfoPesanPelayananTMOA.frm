VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakInfoPesanPelayananTMOA 
   Caption         =   "Medifirst2000 - Cetak Laporan"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmCetakInfoPesanPelayananTMOA.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
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
Attribute VB_Name = "frmCetakInfoPesanPelayananTMOA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReportTM As New crCetakiInfoPesanPelayananTM
Dim ReportOA As New crCetakiInfoPesanPelayananOA

Private Sub Form_Load()
    On Error GoTo hell
    Dim adocomd As New ADODB.Command

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    If strCetak = "TM" Then
        Call PesanPelayananTM
    Else
        Call PesanPelayananOA
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
hell:
    Screen.MousePointer = vbDefault
    Call msubPesanError
End Sub

Private Sub PesanPelayananTM()
    With ReportTM
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail

        .txtPeriode.SetText Format(mdTglAwal, "dd MMMM yyyy") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy")
        .txtUserPrint.SetText strNmPegawai

        Set adocomd = New ADODB.Command
        adocomd.ActiveConnection = dbConn
        adocomd.CommandText = strSQL
        adocomd.CommandType = adCmdUnknown
        .Database.AddADOCommand dbConn, adocomd

        .usNamaPelayanan.SetUnboundFieldSource ("{ado.NamaPelayanan}")
        .unJmlPelayanan.SetUnboundFieldSource ("{ado.JmlPelayanan}")
        .ucBiayaSatuan.SetUnboundFieldSource ("{ado.BiayaSatuan}")
        .usStatusCito.SetUnboundFieldSource ("{ado.StatusCito}")
        .usNoPendaftaran.SetUnboundFieldSource ("{ado.NoPendaftaran}")
        .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
        .usNamaPasien.SetUnboundFieldSource ("{ado.NamaPasien}")
        .usRuanganTujuan.SetUnboundFieldSource ("{ado.RuanganTujuan}")
        .usDokterOrder.SetUnboundFieldSource ("{ado.DokterOrder}")
        .usUserOrder.SetUnboundFieldSource ("{ado.UserOrder}")
    End With

    If vLaporan = "view" Then
        CRViewer1.ReportSource = ReportTM
        With CRViewer1
            .Zoom 1
            .ViewReport
        End With
    Else
        ReportTM.PrintOut False
        Unload Me
    End If
End Sub

Private Sub PesanPelayananOA()
    With ReportOA
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail

        .txtPeriode.SetText Format(mdTglAwal, "dd MMMM yyyy") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy")
        .txtUserPrint.SetText strNmPegawai

        Set adocomd = New ADODB.Command
        adocomd.ActiveConnection = dbConn
        adocomd.CommandText = strSQL
        adocomd.CommandType = adCmdUnknown
        .Database.AddADOCommand dbConn, adocomd

        .usJenisBarang.SetUnboundFieldSource ("{ado.JenisBarang}")
        .usNamaBarang.SetUnboundFieldSource ("{ado.NamaBarang}")
        .usAsalBarang.SetUnboundFieldSource ("{ado.NamaAsal}")
        .unJmlBarang.SetUnboundFieldSource ("{ado.JmlBarang}")
        .ucHargaFIFO.SetUnboundFieldSource ("{ado.HargaFIFO}")
        .usSatuan.SetUnboundFieldSource ("{ado.Satuan}")
        .usNoPendaftaran.SetUnboundFieldSource ("{ado.NoPendaftaran}")
        .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
        .usNamaPasien.SetUnboundFieldSource ("{ado.NamaPasien}")
        .usRuanganTujuan.SetUnboundFieldSource ("{ado.RuanganTujuan}")
        .usDokterOrder.SetUnboundFieldSource ("{ado.DokterOrder}")
        .usUserOrder.SetUnboundFieldSource ("{ado.UserOrder}")
    End With

    If vLaporan = "view" Then
        CRViewer1.ReportSource = ReportOA
        With CRViewer1
            .Zoom 1
            .ViewReport
        End With
    Else
        ReportOA.PrintOut False
        Unload Me
    End If
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strCetak = ""
    Set frmCetakInfoPesanPelayananTMOA = Nothing
End Sub
