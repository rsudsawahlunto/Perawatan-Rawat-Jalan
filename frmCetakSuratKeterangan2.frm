VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakSuratKeterangan2 
   Caption         =   "Cetak Surat Keterangan"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   Icon            =   "frmCetakSuratKeterangan2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
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
Attribute VB_Name = "frmCetakSuratKeterangan2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rptSuratKeterangan2 As New crSuratKeterangan2

Private Sub Form_Load()
    Dim adocomd As New ADODB.Command
    Dim bln As String
    On Error GoTo errLoad

    bln = Format(Now, "MM")
    Select Case bln
        Case "01"
            bln = "I"
        Case "02"
            bln = "II"
        Case "03"
            bln = "III"
        Case "04"
            bln = "IV"
        Case "05"
            bln = "V"
        Case "06"
            bln = "VI"
        Case "07"
            bln = "VII"
        Case "08"
            bln = "VIII"
        Case "09"
            bln = "IX"
        Case "10"
            bln = "X"
        Case "11"
            bln = "XI"
        Case "12"
            bln = "XII"
    End Select

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    With rptSuratKeterangan2

        .Text1.SetText "PEMERINTAH " & UCase(strNKotaRS)
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strWebsite & ", " & strEmail
        .Text2.SetText UCase(strNKotaRS)
        .Text6.SetText strNKotaRS & " , dengan ini menerangkan dengan sebenarnya bahwa :"
        .txtBulan.SetText bln
        .txtTahun.SetText Format(Now, "yyyy")
        .txtTanggal2.SetText Format(frmSuratKeterangan.dtpAwal, "dd MMMM yyyy")
        .txtNama.SetText (frmSuratKeterangan.txtNama.Text)
        .txtNIP.SetText (frmSuratKeterangan.txtNIP)
        .txtJenisKelamin.SetText (frmSuratKeterangan.txtJenisKelamin.Text)
        .txtTtl.SetText (frmSuratKeterangan.txtTempat.Text) & ", " & (frmSuratKeterangan.txtTglLahir.Text)
        .txtPekerjaan.SetText (frmSuratKeterangan.dcPekerjaan.Text)
        .txtAdress.SetText (frmSuratKeterangan.txtKeterangan.Text)
        .txtNamaDokter3.SetText (frmSuratKeterangan.dcDokterPenguji.Text)
        .txtNoCM.SetText frmDaftarPasienRJ.dgDaftarPasienRJ.Columns("NoCM").Value
    End With

    CRViewer1.ReportSource = rptSuratKeterangan2
    CRViewer1.Zoom 1
    CRViewer1.ViewReport
    Screen.MousePointer = vbDefault
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
