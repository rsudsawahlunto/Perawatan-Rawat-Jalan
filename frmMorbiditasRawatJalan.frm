VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmMorbiditasRJ 
   Caption         =   "Morbiditas"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5940
   Icon            =   "frmMorbiditasRawatJalan.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   5940
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
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
Attribute VB_Name = "frmMorbiditasRJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crMorbiditasRawatInap
Dim adoCommand As New ADODB.Command
Dim tanggal As String

Private Sub Form_Load()
    On Error GoTo hell
    Set frmMorbiditasRJ = Nothing

    Set adoCommand.ActiveConnection = dbConn
    Call openConnection
    adoCommand.CommandText = strSQL
    adoCommand.CommandType = adCmdText
    If Format(mdTglAwal, "dd MMMM yyyy") = Format(mdTglAkhir, "dd MMMM yyyy") Then
        tanggal = "Tanggal Kunjungan  : " & " " & Format(mdTglAwal, "dd MMMM yyyy") '& " S/d " & Format(mdTglAkhir, "dd MMMM yyyy")
    Else
        tanggal = "Periode Kunjungan  : " & " " & Format(mdTglAwal, "dd MMMM yyyy") & " S/d " & Format(mdTglAkhir, "dd MMMM yyyy")
    End If

    With Report
        .Database.AddADOCommand dbConn, adoCommand
        .txtJudul.SetText "DATA KEADAAN MORBIDITAS RAWAT JALAN SURVEILANS TERPADU RUMAH SAKIT"
        .txtJudul2.SetText "FORMULIR RL 2b1"
        .txtPeriode.SetText tanggal
        .usNoDTD.SetUnboundFieldSource ("{ado.NoDTD}")
        .usNoDT.SetUnboundFieldSource ("{ado.NoDTerperinci}")
        .usNamaDTD.SetUnboundFieldSource ("{ado.NamaDTD}")
        .unKel1.SetUnboundFieldSource ("{ado.Kel_Umur1}")
        .unKel2.SetUnboundFieldSource ("{ado.Kel_Umur2}")
        .unKel3.SetUnboundFieldSource ("{ado.Kel_Umur3}")
        .unKel4.SetUnboundFieldSource ("{ado.Kel_Umur4}")
        .unKel5.SetUnboundFieldSource ("{ado.Kel_Umur5}")
        .unKel6.SetUnboundFieldSource ("{ado.Kel_Umur6}")
        .unKel7.SetUnboundFieldSource ("{ado.Kel_Umur7}")
        .unKel8.SetUnboundFieldSource ("{ado.Kel_Umur8}")
        .unKelL.SetUnboundFieldSource ("{ado.Kel_L}")
        .unKelP.SetUnboundFieldSource ("{ado.Kel_P}")
        .unKelH.SetUnboundFieldSource ("{ado.Kel_Kunj}")
        .txtJmlPasien.SetText "Jumlah Kunjungan Pasien"
        .txtJmlPasienH.SetText "Jumlah Kunjungan Pasien Total"
        .txtJmlPasienM.Suppress = True
        .unKelM.Suppress = True
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS
        .Text3.SetText strNKotaRS & " " & strNKodepos & " Telp. " & strNTeleponRS
        .SelectPrinter sDriver, sPrinter, vbNull
        settingreport Report, sPrinter, sDriver, crPaperLegal, sDuplex, crLandscape
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .Zoom 1
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMorbiditasRJ = Nothing
End Sub

