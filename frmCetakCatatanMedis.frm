VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakCatatanMedis 
   Caption         =   "Cetak Catatan Medis"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   Icon            =   "frmCetakCatatanMedis.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   5880
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   0   'False
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
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCetakCatatanMedis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Dim Report As New crCetakCatatanMedis
'
'Private Sub Form_Load()
'    Dim adocomd As New ADODB.Command
'    Call openConnection
'    Set frmCetakCatatanMedis = Nothing
'    Me.WindowState = 2
'    adocomd.ActiveConnection = dbConn
'    adocomd.CommandText = "sELECT * FROM V_CetakRiwayatMedikPasienRJ " _
'    & "WHERE NoPendaftaran='" & mstrNoPen & "'"
'    adocomd.CommandType = adCmdText
'
'    strSQL = "SELECT * FROM V_CetakCatatanMedikPasien WHERE NoCM='" & mstrNoCM & "'"
'    Set rs = Nothing
'    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
'
'    With Report
'        .Text1.SetText strNNamaRS
'        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
'        .Text3.SetText strWebsite & " " & strEmail
'        .txtNoPen.SetText mstrNoPen
'        .txtNoCM.SetText mstrNoCM
'        .txtRuang.SetText strNNamaRuangan
'        If IsNull(rs("Nama Pasien").Value) = False Then .txtNama.SetText rs("Nama Pasien").Value
'        If IsNull(rs("TglLahir").Value) = False Then .txtTglLahir.SetText rs("TglLahir").Value
'        If IsNull(rs("JK").Value) = False Then
'            If rs("JK").Value = "P" Then
'                .txtL.Font.Strikethrough = True
'            ElseIf rs("JK").Value = "L" Then
'                .txtP.Font.Strikethrough = True
'            End If
'        End If
'        If IsNull(rs("Nama Keluarga").Value) = False Then .txtKeluarga.SetText rs("Nama Keluarga").Value
'        If IsNull(rs("Bin").Value) = False Then .txtBin.SetText rs("Bin").Value
'        If IsNull(rs("Pekerjaan").Value) = False Then .txtPekerjaan.SetText rs("Pekerjaan").Value
'        If IsNull(rs("Alamat").Value) = False Then .txtalamat.SetText rs("Alamat").Value
'        If IsNull(rs("RTRW").Value) = False Then .txtRTRW.SetText rs("RTRW").Value
'        If IsNull(rs("Kelurahan").Value) = False Then .txtKelurahan.SetText rs("Kelurahan").Value
'        If IsNull(rs("Kecamatan").Value) = False Then .txtKecamatan.SetText rs("Kecamatan").Value
'        If IsNull(rs("Kota").Value) = False Then .txtKota.SetText rs("Kota").Value
'        .Database.AddADOCommand dbConn, adocomd
'        .udtgl.SetUnboundFieldSource ("{Ado.TglPeriksa}")
'        .usRuanganPemeriksaan.SetUnboundFieldSource ("{Ado.RuangPemeriksaan}")
'        .usKeluhanUtama.SetUnboundFieldSource ("{Ado.KeluhanUtama}")
'        .usDiagnosa.SetUnboundFieldSource ("{Ado.Diagnosa}")
'        .usPengobatan.SetUnboundFieldSource ("{Ado.Pengobatan}")
'        .usKeterangan.SetUnboundFieldSource ("{Ado.Keterangan}")
'        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
'    End With
'    Screen.MousePointer = vbHourglass
'    If vLaporan = "view" Then
'        With CRViewer1
'            .ReportSource = Report
'            .ViewReport
'            .Zoom 1
'        End With
'    Else
'        Report.PrintOut False
'        Unload Me
'    End If
'    Screen.MousePointer = vbDefault
'End Sub
'
'Private Sub Form_Resize()
'    CRViewer1.Top = 0
'    CRViewer1.Left = 0
'    CRViewer1.Height = ScaleHeight
'    CRViewer1.Width = ScaleWidth
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    Set frmCetakCatatanMedis = Nothing
'End Sub
'
