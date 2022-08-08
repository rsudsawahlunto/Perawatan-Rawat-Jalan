VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakSuratKeteranganKeluhan 
   Caption         =   "Cetak Surat Keterangan Keluhan"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   Icon            =   "frmCetakSuratKeteranganKeluhan.frx":0000
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
Attribute VB_Name = "frmCetakSuratKeteranganKeluhan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rptSuratKeterangan As New crSuratKeteranganKeluhan

Private Sub Form_Load()
    Dim adocomd As New ADODB.Command
    Dim bln As String
    On Error GoTo errLoad

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    With rptSuratKeterangan
        .Text1.SetText "PEMERINTAH " & UCase(strNKotaRS)
        .txtNamaRS.SetText "RSUD KELAS " & UCase(strkelasRS) & " " & UCase(strketkelasRS)
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strWebsite & ", " & strEmail
        .Text2.SetText UCase(strNKotaRS)

        .TxtTanggal.SetText Format(frmSuratKeteranganKeluhan.dtpAwal, "dd MMMM yyyy")
        .txtNama.SetText (frmSuratKeteranganKeluhan.txtNama.Text)
        .txtUmur.SetText (frmSuratKeteranganKeluhan.txtUmur.Text)
        .txtNoCM.SetText (frmSuratKeteranganKeluhan.txtNoCM.Text)
        .txtGolDarah.SetText (frmSuratKeteranganKeluhan.dcGolDarah.Text)
        .TxtAlmt.SetText (frmSuratKeteranganKeluhan.txtAlamat.Text)
        .TxtKeluhanSkrg.SetText (frmSuratKeteranganKeluhan.txtKeluhan.Text)
        .TxtRiwayatPenyakitDahulu.SetText (frmSuratKeteranganKeluhan.txtRiwayat.Text)
        .TxtTinggiBadan.SetText (frmSuratKeteranganKeluhan.txtTinggi.Text) & " Cm "
        .TxtBeratBadan.SetText (frmSuratKeteranganKeluhan.txtBerat.Text) & " Kg "
        .txtBlood.SetText (frmSuratKeteranganKeluhan.txtTekanan.Text)
        .txtBlood2.SetText (frmSuratKeteranganKeluhan.txtTekanan2.Text)
        .txtPernafasan.SetText (frmSuratKeteranganKeluhan.txtPernapasan.Text) & " X/mrt "
        .txtNadi.SetText (frmSuratKeteranganKeluhan.txtNadi.Text) & " X/mrt "
        .txtMataKanan1.SetText (frmSuratKeteranganKeluhan.txtMataKanan1.Text)
        .txtMataKanan2.SetText (frmSuratKeteranganKeluhan.txtMataKanan2.Text)
        .txtMataKiri1.SetText (frmSuratKeteranganKeluhan.txtMataKiri1.Text)
        .txtMataKiri2.SetText (frmSuratKeteranganKeluhan.txtMataKiri2.Text)

        If frmSuratKeteranganKeluhan.chkYa.Value = 1 Then
            .TxtButawarna.SetText ("Ya")
        End If
        If frmSuratKeteranganKeluhan.chkTidak.Value = 1 Then
            .TxtButawarna.SetText ("Tidak")
        End If

        'Buat THT
        .TxtKesan1.SetText (frmSuratKeteranganKeluhan.txtTHT.Text)
        .TxtGigi.SetText (frmSuratKeteranganKeluhan.TxtGigi.Text)
        'Buat Leher
        .TxtKesan2.SetText (frmSuratKeteranganKeluhan.TxtLeher.Text)
        .TxtJantung.SetText (frmSuratKeteranganKeluhan.TxtJantung.Text)
        .TxtParu.SetText (frmSuratKeteranganKeluhan.TxtParu.Text)
        .TxtAbdomen.SetText (frmSuratKeteranganKeluhan.txtPerut.Text)
        .TxtEstrimitas.SetText (frmSuratKeteranganKeluhan.txtEsminitas.Text)
        .txtRadiologi.SetText (frmSuratKeteranganKeluhan.txtRadiologi.Text)
        .TxtPenyakitDalam.SetText (frmSuratKeteranganKeluhan.TxtPenyakitDalam.Text)
        .TxtPemeriksaanBedah.SetText (frmSuratKeteranganKeluhan.txtBedah.Text)
        .txtElektro.SetText (frmSuratKeteranganKeluhan.txtElekto.Text)
        .txtlaboratorium.SetText (frmSuratKeteranganKeluhan.txtlaboratorium.Text)
        .TxtUSG.SetText (frmSuratKeteranganKeluhan.TxtUSG.Text)
        .TxtTreadmill.SetText (frmSuratKeteranganKeluhan.TxtTreadmill.Text)
        .TxtAudiometri.SetText (frmSuratKeteranganKeluhan.TxtAudiometri.Text)
        .txtSimpul.SetText (frmSuratKeteranganKeluhan.txtKesimpulan2.Text)
        .TxtAnjuran.SetText (frmSuratKeteranganKeluhan.txtAnjuran2.Text)
    End With
    Screen.MousePointer = vbDefault

    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = rptSuratKeterangan
            .ViewReport
            .Zoom 1
        End With
    Else
        rptSuratKeterangan.PrintOut False
        Unload Me
    End If
    Exit Sub

errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub
