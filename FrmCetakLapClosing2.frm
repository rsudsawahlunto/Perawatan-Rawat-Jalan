VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmCetakLapClosing2 
   Caption         =   "Cetak Lap Closing 2"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   Icon            =   "FrmCetakLapClosing2.frx":0000
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
      EnableGroupTree =   0   'False
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
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "FrmCetakLapClosing2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim report3 As New CrDaftarJumlahPasien2

Private Sub Form_Load()
    Me.WindowState = 2
    Screen.MousePointer = vbHourglass
    Set dbcmd = New ADODB.Command
    Set dbcmd.ActiveConnection = dbConn

    Me.Caption = "Medifirst2000 - Cetak Closing Data Pelayanan TM,OA,Apotik"
    Set report3 = New CrDaftarJumlahPasien2

    'sqlnya dari frmdaftar
    dbcmd.CommandText = strSQL
    dbcmd.CommandType = adCmdText

    With report3
        .Database.AddADOCommand dbConn, dbcmd
        .txtinstalasi.SetText mstrNamaRuangan
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strWebsite & ", " & strEmail
        .txtPeriode.SetText CStr(Format(mdTglAwal, "dd MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "dd MMMM yyyy"))
        .usNamaDokter.SetUnboundFieldSource ("{ado.DokterPemeriksa}")
        .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
        .usJenisPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
        .usStatus.SetUnboundFieldSource ("{ado.Status}")

        With CRViewer1
            .ReportSource = report3
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    End With

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
    Set FrmCetakLapClosing = Nothing
End Sub
