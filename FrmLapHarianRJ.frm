VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmLapHarianRJ 
   Caption         =   "MediFirst2000 - Laporan Harian Registrasi "
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   8475
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7695
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
Attribute VB_Name = "FrmLapHarianRJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New LapHarianRJ
Private Sub Form_Load()

Dim laporan As New ADODB.Command
   
    Set rs = New Recordset
   
    Screen.MousePointer = vbHourglass
     Me.WindowState = 2
     openConnection
     With laporan
        .ActiveConnection = dbConn
         .CommandText = "SELECT * " _
                       & "FROM v_daftarpasienRJ where kdruangan='" & strNKdRuangan & "' and dbo.ubahtanggal(tglmasuk) BETWEEN '" & Format(frmtgllap.DTPickerAwal, "yyyy/mm/dd 00:00:00") & "' AND '" & Format(frmtgllap.DTPickerAkhir, "yyyy/mm/dd 23:59:59") & "'"
              

        .CommandType = adCmdText
     End With
      Set rs = laporan.Execute
  Dim tgl As String
  
  If Format(frmtgllap.DTPickerAwal, "dd MMMM yyyy") = Format(frmtgllap.DTPickerAkhir, "dd MMMM yyyy") Then
    tgl = "Tanggal " & Format(frmtgllap.DTPickerAwal, "dd MMMM yyyy") '& " s/d " & Format(frmtgllap.DTPickerAkhir, "dd MMMM yyyy")
  Else
    tgl = "Periode " & Format(frmtgllap.DTPickerAwal, "dd MMMM yyyy") & " s/d " & Format(frmtgllap.DTPickerAkhir, "dd MMMM yyyy")
  End If
    
     With Report
        .Database.AddADOCommand dbConn, laporan
        .txttgl.SetText tgl
        .UnboundString1.SetUnboundFieldSource ("{ado.tglmasuk}")
        .UnboundString2.SetUnboundFieldSource ("{ado.Nopendaftaran}")
        .UnboundString3.SetUnboundFieldSource ("{ado.NOCM}")
        .UnboundString4.SetUnboundFieldSource ("{ado.Namapasien}")
        .UnboundString5.SetUnboundFieldSource ("{ado.Alamat}")
        .UnboundString6.SetUnboundFieldSource ("{ado.Pekerjaan}")
        .UnboundString7.SetUnboundFieldSource ("{ado.Agama}")
        .UnboundString8.SetUnboundFieldSource ("{ado.Jeniskelamin}")
        .UnboundString9.SetUnboundFieldSource ("{ado.Umur}")
        .UnboundString10.SetUnboundFieldSource ("{ado.Status}")
        .UnboundString11.SetUnboundFieldSource ("{ado.JenisPasien}")
'        .UnboundString12.SetUnboundFieldSource ("{ado.Penjamin}")
        .UnboundString13.SetUnboundFieldSource ("{ado.namaSubInstalasi}")
                 
        .txtNamaRS.SetText strNNamaRS
        .txtAlamatRS.SetText strNAlamatRS
        .txtKotaRS.SetText strNKotaRS & " " & strNKodepos & " Telp. " & strNTeleponRS
        .SelectPrinter sDriver, sPrinter, vbNull
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
      End With
     Set rs = Nothing
     
    
    CRViewer1.ReportSource = Report
    CRViewer1.ViewReport
    CRViewer1.DisplayTabs = False
    CRViewer1.DisplayGroupTree = False
    CRViewer1.Zoom 1
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmLapHarianRJ = Nothing
End Sub



