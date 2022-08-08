VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakLapDataBayiBulan 
   Caption         =   "Cetak Lap Data Bayi Bulan"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   Icon            =   "frmCetakLapDataBayiBulan.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7050
   ScaleWidth      =   5850
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
Attribute VB_Name = "frmCetakLapDataBayiBulan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crLapDataBayiBulan

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    Dim adocomd As New ADODB.Command
    Call openConnection

    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = "select * from V_DataBayi  WHERE TglMasuk BETWEEN '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "'"
    adocomd.CommandType = adCmdText

    Report.Database.AddADOCommand dbConn, adocomd

    With Report
        .txtnmrs.SetText strNNamaRS
        .txtalmtrs.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
        .txtweb.SetText strWebsite & ", " & strEmail

        .txtPeriode.SetText CStr(Format(mdTglAwal, "dd MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "dd MMMM yyyy"))
        .usNoCM.SetUnboundFieldSource ("{ado.nocm}")
        .udTanggal.SetUnboundFieldSource ("{ado.tglmasuk}")
        .usbarulama.SetUnboundFieldSource ("{ado.status}")
        .usNamaPasien.SetUnboundFieldSource ("{ado.namapasien}")
        .usAlamat.SetUnboundFieldSource ("{ado.alamat}")
        .udtgllahir.SetUnboundFieldSource ("{ado.tgllahir}")
        .usJK.SetUnboundFieldSource ("{ado.jk}")
        .usbb.SetUnboundFieldSource ("{ado.berattinggibadan}")

        .txtRuanganLogin.SetText mstrNamaRuangan
        .txtUser.SetText strNmPegawai
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, crLandscape
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .Zoom (100)
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault

    Exit Sub
errLoad:
    Screen.MousePointer = vbDefault
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakLapDataBayiBulan = Nothing
End Sub
