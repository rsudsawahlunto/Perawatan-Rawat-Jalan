VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmCetakInformasiDiagnosaPasien 
   Caption         =   "Viewer Laporan"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10035
   Icon            =   "frmCetakInformasiDiagnosaPasien.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   10035
   WindowState     =   2  'Maximized
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
Attribute VB_Name = "FrmCetakInformasiDiagnosaPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DiagnosaPasien As New crInformasiDiagnosaPasien

Private Sub Form_Load()
    Dim adocomd As New ADODB.Command
    adocomd.ActiveConnection = dbConn

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Dim tanggal As String

    Select Case cetak

        Case "PasienPoliBelum"
            adocomd.CommandText = strSQL
            adocomd.CommandType = adCmdText
            With DiagnosaPasien
                .Text16.SetText strNNamaRS
                .Text18.SetText strNAlamatRS
                .Text19.SetText strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
                .Database.AddADOCommand dbConn, adocomd
                .txtRuang.SetText strNNamaRuangan
                .txtTgl.SetText Format(FrmInformasiDiagnosa.DTPickerAwal, "dd/MM/yyyy") & "  s/d  " & Format(FrmInformasiDiagnosa.DTPickerAkhir, "dd/MM/yyyy")

                .usStatus.SetUnboundFieldSource "{ado.TglMasuk}"
                .usNoDaf.SetUnboundFieldSource "{ado.NoPendaftaran}"
                .usCM.SetUnboundFieldSource "{ado.NoCM}"
                .usPasien.SetUnboundFieldSource "{ado.Nama Pasien}"
                .usUmur.SetUnboundFieldSource "{ado.Umur}"
                .usJK.SetUnboundFieldSource "{ado.JK}"
                .usDiagnosa.SetUnboundFieldSource "{ado.Diagnosa}"
                .udtTglMasuk.SetUnboundFieldSource "{ado.Dokter Pemeriksa}"
                .usAlamat.SetUnboundFieldSource "{ado.Ruangan}"

                .SelectPrinter sDriver, sPrinter, vbNull
                settingreport DiagnosaPasien, sPrinter, sDriver, crPaperLegal, sDuplex, crLandscape
            End With
            CRViewer1.ReportSource = DiagnosaPasien

        Case "PasienPoliSudah"
            adocomd.CommandText = strSQL
            adocomd.CommandType = adCmdText
            With DiagnosaPasien
                .Text16.SetText strNNamaRS
                .Text18.SetText strNAlamatRS
                .Text19.SetText strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
                .Database.AddADOCommand dbConn, adocomd
                .txtRuang.SetText strNNamaRuangan
                .txtTgl.SetText Format(FrmInformasiDiagnosa.DTPickerAwal, "dd/MM/yyyy") & "  s/d  " & Format(FrmInformasiDiagnosa.DTPickerAkhir, "dd/MM/yyyy")

                .usStatus.SetUnboundFieldSource "{ado.TglMasuk}"
                .usNoDaf.SetUnboundFieldSource "{ado.NoPendaftaran}"
                .usCM.SetUnboundFieldSource "{ado.NoCM}"
                .usPasien.SetUnboundFieldSource "{ado.Nama Pasien}"
                .usUmur.SetUnboundFieldSource "{ado.Umur}"
                .usJK.SetUnboundFieldSource "{ado.JK}"
                .usDiagnosa.SetUnboundFieldSource "{ado.Diagnosa}"
                .udtTglMasuk.SetUnboundFieldSource "{ado.Dokter Pemeriksa}"
                .usAlamat.SetUnboundFieldSource "{ado.Ruangan}"

                .SelectPrinter sDriver, sPrinter, vbNull
                settingreport DiagnosaPasien, sPrinter, sDriver, crPaperLegal, sDuplex, crLandscape
            End With
            CRViewer1.ReportSource = DiagnosaPasien

        Case "PasienPoli1"
            adocomd.CommandText = strSQL
            adocomd.CommandType = adCmdText
            With DiagnosaPasien
                .Text16.SetText strNNamaRS
                .Text18.SetText strNAlamatRS
                .Text19.SetText strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
                .Database.AddADOCommand dbConn, adocomd
                .txtRuang.SetText strNNamaRuangan
                .txtTgl.SetText Format(FrmInformasiDiagnosa.DTPickerAwal, "dd/MM/yyyy") & "  s/d  " & Format(FrmInformasiDiagnosa.DTPickerAkhir, "dd/MM/yyyy")

                .usStatus.SetUnboundFieldSource "{ado.TglMasuk}"
                .usNoDaf.SetUnboundFieldSource "{ado.NoPendaftaran}"
                .usCM.SetUnboundFieldSource "{ado.NoCM}"
                .usPasien.SetUnboundFieldSource "{ado.Nama Pasien}"
                .usUmur.SetUnboundFieldSource "{ado.Umur}"
                .usJK.SetUnboundFieldSource "{ado.JK}"
                .usDiagnosa.SetUnboundFieldSource "{ado.Diagnosa}"
                .udtTglMasuk.SetUnboundFieldSource "{ado.Dokter Pemeriksa}"
                .usAlamat.SetUnboundFieldSource "{ado.Ruangan}"

                .SelectPrinter sDriver, sPrinter, vbNull
                settingreport DiagnosaPasien, sPrinter, sDriver, crPaperLegal, sDuplex, crLandscape
            End With
            CRViewer1.ReportSource = DiagnosaPasien

        Case "PasienPoli2"
            adocomd.CommandText = strSQL
            adocomd.CommandType = adCmdText
            With DiagnosaPasien
                .Text16.SetText strNNamaRS
                .Text18.SetText strNAlamatRS
                .Text19.SetText strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
                .Database.AddADOCommand dbConn, adocomd
                .txtRuang.SetText strNNamaRuangan
                .txtTgl.SetText Format(FrmInformasiDiagnosa.DTPickerAwal, "dd/MM/yyyy") & "  s/d  " & Format(FrmInformasiDiagnosa.DTPickerAkhir, "dd/MM/yyyy")

                .usStatus.SetUnboundFieldSource "{ado.TglMasuk}"
                .usNoDaf.SetUnboundFieldSource "{ado.NoPendaftaran}"
                .usCM.SetUnboundFieldSource "{ado.NoCM}"
                .usPasien.SetUnboundFieldSource "{ado.Nama Pasien}"
                .usUmur.SetUnboundFieldSource "{ado.Umur}"
                .usJK.SetUnboundFieldSource "{ado.JK}"
                .usDiagnosa.SetUnboundFieldSource "{ado.Diagnosa}"
                .udtTglMasuk.SetUnboundFieldSource "{ado.Dokter Pemeriksa}"
                .usAlamat.SetUnboundFieldSource "{ado.Ruangan}"

                .SelectPrinter sDriver, sPrinter, vbNull
                settingreport DiagnosaPasien, sPrinter, sDriver, crPaperLegal, sDuplex, crLandscape
            End With
            CRViewer1.ReportSource = DiagnosaPasien

        Case "PasienKonsulBelum"
            adocomd.CommandText = strSQL
            adocomd.CommandType = adCmdText
            With DiagnosaPasien
                .Text16.SetText strNNamaRS
                .Text18.SetText strNAlamatRS
                .Text19.SetText strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
                .Database.AddADOCommand dbConn, adocomd
                .txtRuang.SetText strNNamaRuangan
                .txtTgl.SetText Format(FrmInformasiDiagnosa.DTPickerAwal, "dd/MM/yyyy") & "  s/d  " & Format(FrmInformasiDiagnosa.DTPickerAkhir, "dd/MM/yyyy")

                .Text13.SetText "Tanggal Dirujuk"
                .Text1.SetText "Dokter Perujuk"
                .Text4.SetText "Ruangan Tujuan"

                .usStatus.SetUnboundFieldSource "{ado.TglDirujuk}"
                .usNoDaf.SetUnboundFieldSource "{ado.NoPendaftaran}"
                .usCM.SetUnboundFieldSource "{ado.NoCM}"
                .usPasien.SetUnboundFieldSource "{ado.Nama Pasien}"
                .usUmur.SetUnboundFieldSource "{ado.Umur}"
                .usJK.SetUnboundFieldSource "{ado.JK}"
                .usDiagnosa.SetUnboundFieldSource "{ado.Diagnosa}"
                .udtTglMasuk.SetUnboundFieldSource "{ado.Dokter Perujuk}"
                .usAlamat.SetUnboundFieldSource "{ado.Ruangan Tujuan}"

                .SelectPrinter sDriver, sPrinter, vbNull
                settingreport DiagnosaPasien, sPrinter, sDriver, crPaperLegal, sDuplex, crLandscape
            End With
            CRViewer1.ReportSource = DiagnosaPasien
        Case "PasienKonsulSudah"
            adocomd.CommandText = strSQL
            adocomd.CommandType = adCmdText
            With DiagnosaPasien
                .Text16.SetText strNNamaRS
                .Text18.SetText strNAlamatRS
                .Text19.SetText strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
                .Database.AddADOCommand dbConn, adocomd
                .txtRuang.SetText strNNamaRuangan
                .txtTgl.SetText Format(FrmInformasiDiagnosa.DTPickerAwal, "dd/MM/yyyy") & "  s/d  " & Format(FrmInformasiDiagnosa.DTPickerAkhir, "dd/MM/yyyy")

                .Text13.SetText "Tanggal Dirujuk"
                .Text1.SetText "Dokter Perujuk"
                .Text4.SetText "Ruangan Tujuan"

                .usStatus.SetUnboundFieldSource "{ado.TglDirujuk}"
                .usNoDaf.SetUnboundFieldSource "{ado.NoPendaftaran}"
                .usCM.SetUnboundFieldSource "{ado.NoCM}"
                .usPasien.SetUnboundFieldSource "{ado.Nama Pasien}"
                .usUmur.SetUnboundFieldSource "{ado.Umur}"
                .usJK.SetUnboundFieldSource "{ado.JK}"
                .usDiagnosa.SetUnboundFieldSource "{ado.Diagnosa}"
                .udtTglMasuk.SetUnboundFieldSource "{ado.Dokter Perujuk}"
                .usAlamat.SetUnboundFieldSource "{ado.Ruangan Tujuan}"

                .SelectPrinter sDriver, sPrinter, vbNull
                settingreport DiagnosaPasien, sPrinter, sDriver, crPaperLegal, sDuplex, crLandscape
            End With
            CRViewer1.ReportSource = DiagnosaPasien

        Case "PasienKonsul1"
        Case "PasienKonsul2"
            adocomd.CommandText = strSQL
            adocomd.CommandType = adCmdText
            With DiagnosaPasien
                .Text16.SetText strNNamaRS
                .Text18.SetText strNAlamatRS
                .Text19.SetText strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
                .Database.AddADOCommand dbConn, adocomd
                .txtRuang.SetText strNNamaRuangan
                .txtTgl.SetText Format(FrmInformasiDiagnosa.DTPickerAwal, "dd/MM/yyyy") & "  s/d  " & Format(FrmInformasiDiagnosa.DTPickerAkhir, "dd/MM/yyyy")

                .Text13.SetText "Tanggal Dirujuk"
                .Text1.SetText "Dokter Perujuk"
                .Text4.SetText "Ruangan Tujuan"

                .usStatus.SetUnboundFieldSource "{ado.TglDirujuk}"
                .usNoDaf.SetUnboundFieldSource "{ado.NoPendaftaran}"
                .usCM.SetUnboundFieldSource "{ado.NoCM}"
                .usPasien.SetUnboundFieldSource "{ado.Nama Pasien}"
                .usUmur.SetUnboundFieldSource "{ado.Umur}"
                .usJK.SetUnboundFieldSource "{ado.JK}"
                .usDiagnosa.SetUnboundFieldSource "{ado.Diagnosa}"
                .udtTglMasuk.SetUnboundFieldSource "{ado.Dokter Perujuk}"
                .usAlamat.SetUnboundFieldSource "{ado.Ruangan Tujuan}"

                .SelectPrinter sDriver, sPrinter, vbNull
                settingreport DiagnosaPasien, sPrinter, sDriver, crPaperLegal, sDuplex, crLandscape
            End With
            CRViewer1.ReportSource = DiagnosaPasien

    End Select
    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = DiagnosaPasien
            .ViewReport
            .Zoom 1
        End With
    Else
        If cetak = "PasienPoliBelum" Then
            DiagnosaPasien.PrintOut False
            Unload Me
        ElseIf cetak = "PasienKonsulSudah" Then
            DiagnosaPasien.PrintOut False
            Unload Me
        ElseIf cetak = "PasienPoli1" Then
            DiagnosaPasien.PrintOut False
            Unload Me
        ElseIf cetak = "PasienPoli2" Then
            DiagnosaPasien.PrintOut False
            Unload Me
        ElseIf cetak = "PasienKonsulBelum" Then
            DiagnosaPasien.PrintOut False
            Unload Me
        ElseIf cetak = "PasienKonsulSudah" Then
            DiagnosaPasien.PrintOut False
            Unload Me
        ElseIf cetak = "PasienKonsul1" Then
            DiagnosaPasien.PrintOut False
            Unload Me
        ElseIf cetak = "PasienKonsul2" Then
            DiagnosaPasien.PrintOut False
            Unload Me
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmCetakInformasiDiagnosaPasien = Nothing
End Sub
