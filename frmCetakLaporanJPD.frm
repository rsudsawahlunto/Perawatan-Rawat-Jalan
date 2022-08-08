VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakLaporanJPD 
   Caption         =   "Cetak Laporan JPD"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   Icon            =   "frmCetakLaporanJPD.frx":0000
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
Attribute VB_Name = "frmCetakLaporanJPD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rptJPDokter As New crInfoJasaPelDokter
Dim strLNamaDokter As String

Private Sub Form_Load()
    Me.WindowState = 2
    Screen.MousePointer = vbHourglass
    Set dbcmd = New ADODB.Command
    Set dbcmd.ActiveConnection = dbConn
    Me.Caption = "Medifirst2000 - Informasi Jasa Pelayanan Dokter"
    strSQL = "SELECT * FROM V_LaporanJasaPelayananDokter " _
    & "WHERE [Dokter Pemeriksa]='" & strLNamaDokter & "' AND " _
    & "TglStruk BETWEEN '" _
    & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
    & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "'"

    dbcmd.CommandText = strSQL
    dbcmd.CommandType = adCmdText
    With rptJPDokter
        .Database.AddADOCommand dbConn, dbcmd
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS
        .txtAlamat2.SetText strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtPeriode.SetText Format(mdTglAwal, "dd MMMM yyyy") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy")
        .txtJudul.SetText "INFORMASI JASA PELAYANAN DOKTER"

        .UsDokter.SetUnboundFieldSource ("{ado.Dokter Pemeriksa}")
        .usJenisPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
        .UsPenjamin.SetUnboundFieldSource ("{ado.Penjamin}")
        .usRuangan.SetUnboundFieldSource ("{ado.Ruang Pelayanan}")
        .UsJnsPelayanan.SetUnboundFieldSource ("{ado.Jenis Pelayanan}")
        .usNamaPelayanan.SetUnboundFieldSource ("{ado.Nama Pelayanan}")
        .usKomponenJasa.SetUnboundFieldSource ("{ado.Komponen Jasa}")
        .ucTotalBiaya.SetUnboundFieldSource ("{ado.TotalBiaya}")
        .ucJmlByr.SetUnboundFieldSource ("{ado.JmlBayar}")
        .ucPiutang.SetUnboundFieldSource ("{ado.Piutang}")
        .ucPembebasan.SetUnboundFieldSource ("{ado.Pembebasan}")
        .ucSisaTagihan.SetUnboundFieldSource ("{ado.SisaTagihan}")

        settingreport rptJPDokter, sPrinter, sDriver, sUkuranKertas, sDuplex, crLandscape
    End With
    CRViewer1.ReportSource = rptJPDokter
    With CRViewer1
        .EnableGroupTree = False
        .ViewReport
        .Zoom 1
    End With
    Screen.MousePointer = vbDefault
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
    Set frmCetakLaporanJPD = Nothing
End Sub

'untuk loading data Nama Dokter
Public Sub subLoadNmDokter(strInput As String)
    strLNamaDokter = strInput
End Sub

