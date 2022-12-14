VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakDaftarPemakaianObatAlkesKaryawan 
   Caption         =   "Medifirst2000 - Cetak Daftar Pemakaian Obat Alkes Karyawan"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakDaftarPemakaianObatAlkesKaryawan.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
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
Attribute VB_Name = "frmCetakDaftarPemakaianObatAlkesKaryawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rptCetakDaftarPemakaianObatAlkesKaryawan As crCetakDaftarPemakaianObatAlkesKaryawan

Private Sub Form_Load()
    Me.WindowState = 2
    Screen.MousePointer = vbHourglass
    Set dbcmd = New ADODB.Command
    Set dbcmd.ActiveConnection = dbConn
    
    Set rptCetakDaftarPemakaianObatAlkesKaryawan = New crCetakDaftarPemakaianObatAlkesKaryawan
    strSQL = "select TglPemakaian,[Nama Barang],[Asal Barang],Satuan,JmlBarang,HargaSatuan,Total,Keperluan,[Penanggung Jawab],Ruangan,KdRuangan " & _
        " from V_DaftarPemakaianObatAlkesKaryawan " & _
        " where Ruangan='" & mstrNamaRuangan & "' AND TglPemakaian BETWEEN '" & Format(mdTglAwal, "yyyy-MM-dd 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy-MM-dd 23:59:59") & "'"
    dbcmd.CommandText = strSQL
    dbcmd.CommandType = adCmdText
    With rptCetakDaftarPemakaianObatAlkesKaryawan
        .Database.AddADOCommand dbConn, dbcmd
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strWebsite & ", " & strEmail
        .txtPeriode.SetText Format(mdTglAwal, "dd MMMM yyyy") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy")
        .txtPetugas.SetText strNmPegawai
        .txtRuangan.SetText strNNamaRuangan
        
        .usNamaBarang.SetUnboundFieldSource ("{ado.Nama Barang}")
        .usAsalBarang.SetUnboundFieldSource ("{ado.Asal Barang}")
        .unJumlahBarang.SetUnboundFieldSource ("{ado.JmlBarang}")
        .usSatuanBarang.SetUnboundFieldSource ("{ado.Satuan}")
        .ucHargaBarang.SetUnboundFieldSource ("{ado.HargaSatuan}")
        .ucTotal.SetUnboundFieldSource ("{ado.Total}")
        .usKeperluan.SetUnboundFieldSource ("{ado.Keperluan}")
        .usPenanggungJawab.SetUnboundFieldSource ("{ado.Penanggung Jawab}")
         
         settingreport rptCetakDaftarPemakaianObatAlkesKaryawan, sPrinter, sDriver, sUkuranKertas, sDuplex, crLandscape
    End With
    CRViewer1.ReportSource = rptCetakDaftarPemakaianObatAlkesKaryawan
    
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
    Set frmCetakDaftarPemakaianObatAlkesKaryawan = Nothing
End Sub











