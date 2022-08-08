VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakPRMRJ 
   Caption         =   "Cetak PRMRJ"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8400
   Icon            =   "frmCetakPRMRJ.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   8400
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
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
Attribute VB_Name = "frmCetakPRMRJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crPRMRJ

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    Dim adocomd As New ADODB.Command
    Call openConnection

    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    adocomd.CommandType = adCmdText

    Report.Database.AddADOCommand dbConn, adocomd

    With Report
        .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
        .usNama.SetUnboundFieldSource ("{Ado.NamaLengkap}")
        .udTglLahir.SetUnboundFieldSource ("{ado.TglLahir}")
        .usJekel.SetUnboundFieldSource ("{Ado.JenisKelamin}")
        .udtTglMasuk.SetUnboundFieldSource ("{Ado.TglMasuk}")
        .usDiagnosa.SetUnboundFieldSource ("{Ado.KdDiagnosa}")
        .usPen.SetUnboundFieldSource ("{Ado.Penunjang}")
        .usObat.SetUnboundFieldSource ("{Ado.Obat}")
        .usRuangan.SetUnboundFieldSource ("{ado.Ruangan}")
    End With

    Screen.MousePointer = vbHourglass
'    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .Zoom 1
        End With
'    Else
'        Report.PrintOut False
'        Unload Me
'    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakPRMRJ = Nothing
End Sub

