VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakDaftarTindakan 
   Caption         =   "Cetak Daftar Tindakan Pasien"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   Icon            =   "frmCetakDaftarTindakan.frx":0000
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
Attribute VB_Name = "frmCetakDaftarTindakan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim repDafTindakan As New crDaftarTindakanPasien

Private Sub Form_Load()
    Call openConnection
    Set frmCetakDaftarTindakan = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With repDafTindakan
        .Database.AddADOCommand dbConn, adocomd
        If strCetak2 = "LapDafTindakanPasienHari" Then
            If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
                .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
            Else
                .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
            End If
        Else
            If CStr(Format(mdTglAwal, "dd MMMM yyyy")) = CStr(Format(mdTglAkhir, "dd MMMM yyyy")) Then
                .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
            Else
                .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "dd MMMM yyyy")))
            End If
        End If
        .txtRuangan.SetText mstrNamaRuangan
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail

        .udTglPelayanan.SetUnboundFieldSource ("{ado.TglPelayanan}")
        .UsDokter.SetUnboundFieldSource ("{ado.DokterPemeriksa}")
        .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
        .usNmPasien.SetUnboundFieldSource ("{ado.NamaPasien}")
        .usUmur.SetUnboundFieldSource ("{ado.Umur}")
        .usJK.SetUnboundFieldSource ("{ado.JenisKelamin}")
        .usStatus.SetUnboundFieldSource ("{ado.StatusKasus}")
        .usNamaDiagnosa.SetUnboundFieldSource ("{ado.NamaDiagnosa}")
        .usJnsPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
        .usDiagnosa.SetUnboundFieldSource ("{ado.JenisPelayanan}")
        .ustindakan.SetUnboundFieldSource ("{ado.TindakanPelayanan}")
        .ucTarif.SetUnboundFieldSource ("{ado.Tarif}")
        .ucTarifCito.SetUnboundFieldSource ("{ado.TarifCito}")
        .unJmlTindakan.SetUnboundFieldSource ("{ado.JmlPelayanan}")
        .ucTotBiaya.SetUnboundFieldSource ("{ado.TotalBiaya}")

        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
    End With

    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = repDafTindakan
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        repDafTindakan.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmCetakLapKunjunganPasien = Nothing
    sUkuranKertas = ""
End Sub
