VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCtkLapRekap_Viewer 
   Caption         =   "Cetak  Lap Rekap"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3240
   Icon            =   "frmCtkLapRekap_Viewer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3330
   ScaleWidth      =   3240
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
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
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCtkLapRekap_Viewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim a As String, b As String, c As String
Dim Report As New CrRekapitulasiPsnKunjunganBerdskanPenyakitdanjenis

Private Sub Form_Load()
    Call openConnection
    Dim waktu As String
    waktu = "{Ado.TglPendaftaran}"
    Dim adocomd As New ADODB.Command
    Set frmCtkLapRekap_Viewer = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    '*************************************
    'SET TANGGAL DAN PENAMAAN RSU
    '*************************************
    With Report
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtWebsiteRS.SetText strWebsite & ", " & strEmail
    End With

    '*************************************
    'SET JUDUL
    '*************************************
    Dim Judul1 As String
    Dim Judul2 As String
    Dim Judul3 As String
    Dim judul4 As String
    Dim Judul5 As String
    Dim Judul6 As String
    Dim Judul7 As String
    Dim judul8 As String
    Dim judul9 As String
    Dim judul10 As String
    Dim judul11 As String
    Dim judul12 As String
    Dim Judul13 As String
    Dim judul14 As String

    Judul1 = "LAPORAN REKAPITULASI PASIEN BERDASARKAN STATUS DAN JENIS PASIEN (PERHARI)"
    Judul2 = "LAPORAN REKAPITULASI PASIEN BERDASARKAN STATUS DAN JENIS PASIEN (PERBULAN)"
    Judul3 = "LAPORAN REKAPITULASI PASIEN BERDASARKAN STATUS DAN RUJUKAN (PERHARI)"
    judul4 = "LAPORAN REKAPITULASI PASIEN BERDASARKAN STATUS DAN RUJUKAN (PERBULAN)"
    Judul5 = "LAPORAN REKAPITULASI PASIEN BERDASARKAN STATUS DAN KASUS PENYAKIT PASIEN(PERHARI)"
    Judul6 = "LAPORAN REKAPITULASI PASIEN BERDASARKAN STATUS DAN KASUS PENYAKIT PASIEN(PERBULAN)"
    Judul7 = "LAPORAN REKAPITULASI PASIEN BERDASARKAN STATUS DAN JENIS OPERASI (PERHARI)"
    judul8 = "LAPORAN REKAPITULASI PASIEN BERDASARKAN STATUS DAN JENIS OPERASI(PERBULAN)"
    judul9 = "LAPORAN REKAPITULASI PASIEN BERDASARKAN STATUS DAN KELAS PELAYANAN(PERHARI)"
    judul10 = "LAPORAN REKAPITULASI PASIEN BERDASARKAN STATUS DAN KELAS PELAYANAN(PERBULAN)"
    judul11 = "LAPORAN REKAPITULASI PASIEN BERDASARKAN STATUS DAN KONDISI PULANG (PERHARI)"
    judul12 = "LAPORAN REKAPITULASI PASIEN BERDASARKAN STATUS DAN KONDISI PULANG (PERBULAN)"
    Judul13 = "LAPORAN REKAPITULASI PASIEN BERDASARKAN JENIS PENYAKIT DAN KONDISI PULANG (PERHARI)"
    judul14 = "LAPORAN REKAPITULASI PASIEN BERDASARKAN JENIS PENYAKIT DAN KONDISI PULANG (PERBULAN)"

    Screen.MousePointer = vbHourglass
    '<<MEMILIH FILTER CETAK>>
    Select Case strCetak
            '1 - STATUS DAN JENIS
        Case "LapRekapKPSJ"
            Select Case strCetak2
                    '==================================================
                    '    LAPORAN PER HARI b PSN Status dan Jenis
                    '==================================================
                Case "LapRekapKPSJhr"
                    adocomd.CommandText = strSQL '? STRSQL berdasarkan filter
                    Report.Database.AddADOCommand dbConn, adocomd
                    With Report
                        .usJK.SetUnboundFieldSource ("{Ado.JK}")
                        .usRuangan.SetUnboundFieldSource ("{Ado.RuanganPelayanan}")
                        .UsKasusPenyakit.SetUnboundFieldSource ("{Ado.Detail}")
                        .UnboundDateTime1.SetUnboundFieldSource Format(waktu, "dd mmm yyyy")
                        .UsJenis.SetUnboundFieldSource ("{Ado.Judul}")
                        .unJumlah.SetUnboundFieldSource ("{Ado.JmlPasien}")
                        .txtJudul.SetText Judul1
                        If Format(mdTglAwal, "dd/MMM/yyyy") = Format(mdTglAkhir, "dd/MMM/yyyy") Then
                            .txtBulan.SetText "Tanggal    : " & Format(mdTglAwal, "dd MMMM YYYY")
                        Else
                            .txtBulan.SetText "Periode    : " & Format(mdTglAwal, "dd/MM/yyyy") & "     s/d     " & Format(mdTglAkhir, "dd/MM/yyyy")
                        End If
                    End With
                    '==================================================
                    '    LAPORAN PER BULAN b PSN Status dan Jenis
                    '==================================================
                Case "LapRekapKPSJbln"
                    adocomd.CommandText = strSQL '? STRSQL berdasarkan filter
                    Report.Database.AddADOCommand dbConn, adocomd
                    With Report
                        .usJK.SetUnboundFieldSource ("{Ado.JK}")
                        .usRuangan.SetUnboundFieldSource ("{Ado.RuanganPelayanan}")
                        .UsKasusPenyakit.SetUnboundFieldSource ("{Ado.Detail}")
                        .UnboundDateTime1.SetUnboundFieldSource Format(waktu, "MMM YYYY")
                        .UsJenis.SetUnboundFieldSource ("{Ado.Judul}")
                        .unJumlah.SetUnboundFieldSource ("{Ado.JmlPasien}")
                        .txtJudul.SetText Judul2
                        If Format(mdTglAwal, "MMM/yyyy") = Format(mdTglAkhir, "MMM/yyyy") Then
                            .txtBulan.SetText "Bulan    : " & Format(mdTglAwal, "MMMM YYYY")
                        Else
                            .txtBulan.SetText "Periode    : " & Format(mdTglAwal, "MMMM/yyyy") & "     s/d     " & Format(mdTglAkhir, "MMMM/yyyy")
                        End If

                    End With
            End Select
            '2 - STATUS DAN RUJUKAN
        Case "LapRekapKPSR"
            '==================================================
            '    LAPORAN PER HARI b PSN Status dan Rujukan
            '==================================================
            Select Case strCetak2
                Case "LapRekapKPSRhr"
                    adocomd.CommandText = strSQL '? STRSQL berdasarkan filter
                    Report.Database.AddADOCommand dbConn, adocomd
                    With Report
                        .usJK.SetUnboundFieldSource ("{Ado.JK}")
                        .usRuangan.SetUnboundFieldSource ("{Ado.RuanganPelayanan}")
                        .UsKasusPenyakit.SetUnboundFieldSource ("{Ado.Detail}")
                        .UnboundDateTime1.SetUnboundFieldSource Format(waktu, "dd mmm yyyy")
                        .UsJenis.SetUnboundFieldSource ("{Ado.Judul}")
                        .unJumlah.SetUnboundFieldSource ("{Ado.JmlPasien}")
                        .txtJudul.SetText Judul3
                        If Format(mdTglAwal, "dd/MMM/yyyy") = Format(mdTglAkhir, "dd/MMM/yyyy") Then
                            .txtBulan.SetText "Tanggal    : " & Format(mdTglAwal, "dd MMMM YYYY")
                        Else
                            .txtBulan.SetText "Periode    : " & Format(mdTglAwal, "dd/MM/yyyy") & "     s/d     " & Format(mdTglAkhir, "dd/MM/yyyy")
                        End If
                    End With
                    '==================================================
                    '    LAPORAN PER BULAN b PSN Status dan Rujukan
                    '==================================================
                Case "LapRekapKPSRbln"
                    adocomd.CommandText = strSQL '? STRSQL berdasarkan filter
                    Report.Database.AddADOCommand dbConn, adocomd
                    With Report
                        .usJK.SetUnboundFieldSource ("{Ado.JK}")
                        .usRuangan.SetUnboundFieldSource ("{Ado.RuanganPelayanan}")
                        .UsKasusPenyakit.SetUnboundFieldSource ("{Ado.Detail}")
                        .UnboundDateTime1.SetUnboundFieldSource Format(waktu, " mmm yyyy")
                        .UsJenis.SetUnboundFieldSource ("{Ado.Judul}")
                        .unJumlah.SetUnboundFieldSource ("{Ado.JmlPasien}")
                        .txtJudul.SetText judul4
                        If Format(mdTglAwal, "MMM/yyyy") = Format(mdTglAkhir, "MMM/yyyy") Then
                            .txtBulan.SetText "Bulan    : " & Format(mdTglAwal, "MMMM YYYY")
                        Else
                            .txtBulan.SetText "Periode    : " & Format(mdTglAwal, "MMMM/yyyy") & "     s/d     " & Format(mdTglAkhir, "MMMM/yyyy")
                        End If
                    End With
            End Select
            '3 - STATUS DAN KASUS PENYAKIT
        Case "LapRekapKssPnyktSK"
            '====================================================
            '    LAPORAN PER HARI b PSN status dan Kasus penyakit
            '====================================================
            Select Case strCetak2
                Case "LapRekapKssPnyktHr"
                    adocomd.CommandText = strSQL
                    Report.Database.AddADOCommand dbConn, adocomd
                    With Report
                        .usJK.SetUnboundFieldSource ("{Ado.JK}")
                        .usRuangan.SetUnboundFieldSource ("{Ado.RuanganPelayanan}")
                        .UsKasusPenyakit.SetUnboundFieldSource ("{Ado.Detail}")
                        .UnboundDateTime1.SetUnboundFieldSource Format(waktu, "dd mmm yyyy")
                        .UsJenis.SetUnboundFieldSource ("{Ado.Judul}")
                        .unJumlah.SetUnboundFieldSource ("{Ado.JmlPasien}")
                        .txtJudul.SetText Judul5
                        If Format(mdTglAwal, "dd/MMM/yyyy") = Format(mdTglAkhir, "dd/MMM/yyyy") Then
                            .txtBulan.SetText "Tanggal    : " & Format(mdTglAwal, "dd MMMM YYYY")
                        Else
                            .txtBulan.SetText "Periode    : " & Format(mdTglAwal, "dd/MM/yyyy") & "     s/d     " & Format(mdTglAkhir, "dd/MM/yyyy")
                        End If
                    End With
                    '=====================================================
                    '    LAPORAN PER BULAN b PSN status dan Kasus penyakit
                    '=====================================================
                Case "LapRekapKssPnyktBln"
                    adocomd.CommandText = strSQL
                    Report.Database.AddADOCommand dbConn, adocomd
                    With Report
                        .usJK.SetUnboundFieldSource ("{Ado.JK}")
                        .usRuangan.SetUnboundFieldSource ("{Ado.RuanganPelayanan}")
                        .UsKasusPenyakit.SetUnboundFieldSource ("{Ado.Detail}")
                        .UnboundDateTime1.SetUnboundFieldSource Format(waktu, "mmm yyyy")
                        .UsJenis.SetUnboundFieldSource ("{Ado.Judul}")
                        .unJumlah.SetUnboundFieldSource ("{Ado.JmlPasien}")
                        .txtJudul.SetText Judul6
                        If Format(mdTglAwal, "MMM/yyyy") = Format(mdTglAkhir, "MMM/yyyy") Then
                            .txtBulan.SetText "Bulan    : " & Format(mdTglAwal, "MMMM YYYY")
                        Else
                            .txtBulan.SetText "Periode    : " & Format(mdTglAwal, "MMMM/yyyy") & "     s/d     " & Format(mdTglAkhir, "MMMM/yyyy")
                        End If
                    End With
            End Select
            '4 - STATUS DAN JENIS OPERASI
        Case "LapRekapKPSO"
            '====================================================
            '    LAPORAN PER HARI b PSN status dan Jenis Operasi
            '====================================================
            Select Case LCase(strCetak2)
                Case "laprekapkpsohr"
                    adocomd.CommandText = strSQL
                    Report.Database.AddADOCommand dbConn, adocomd
                    With Report
                        .usJK.SetUnboundFieldSource ("{Ado.JK}")
                        .usRuangan.SetUnboundFieldSource ("{Ado.RuanganPelayanan}")
                        .UsKasusPenyakit.SetUnboundFieldSource ("{Ado.Detail}")
                        .UnboundDateTime1.SetUnboundFieldSource Format(waktu, "dd mmm yyyy")
                        .UsJenis.SetUnboundFieldSource ("{Ado.Judul}")
                        .unJumlah.SetUnboundFieldSource ("{Ado.JmlPasien}")
                        .txtJudul.SetText Judul7
                        If Format(mdTglAwal, "dd/MMM/yyyy") = Format(mdTglAkhir, "dd/MMM/yyyy") Then
                            .txtBulan.SetText "Tanggal    : " & Format(mdTglAwal, "dd MMMM YYYY")
                        Else
                            .txtBulan.SetText "Periode    : " & Format(mdTglAwal, "dd/MM/yyyy") & "     s/d     " & Format(mdTglAkhir, "dd/MM/yyyy")
                        End If

                    End With
                    '=====================================================
                    '    LAPORAN PER BULAN b PSN status dan Jenis Operasi
                    '=====================================================
                Case "laprekapkpsobln"
                    adocomd.CommandText = strSQL
                    Report.Database.AddADOCommand dbConn, adocomd
                    With Report
                        .usJK.SetUnboundFieldSource ("{Ado.JK}")
                        .usRuangan.SetUnboundFieldSource ("{Ado.RuanganPelayanan}")
                        .UsKasusPenyakit.SetUnboundFieldSource ("{Ado.Detail}")
                        .UnboundDateTime1.SetUnboundFieldSource Format(waktu, "mmm yyyy")
                        .UsJenis.SetUnboundFieldSource ("{Ado.Judul}")
                        .unJumlah.SetUnboundFieldSource ("{Ado.JmlPasien}")
                        .txtJudul.SetText judul8
                        If Format(mdTglAwal, "MMM/yyyy") = Format(mdTglAkhir, "MMM/yyyy") Then
                            .txtBulan.SetText "Bulan    : " & Format(mdTglAwal, "MMMM YYYY")
                        Else
                            .txtBulan.SetText "Periode    : " & Format(mdTglAwal, "MMMM/yyyy") & "     s/d     " & Format(mdTglAkhir, "MMMM/yyyy")
                        End If
                    End With

            End Select
            '5 - STATUS DAN KELAS PELAYANAN
        Case "LapRekapKPSkp"
            '====================================================
            '    LAPORAN PER HARI b PSN status dan Kelas Pelayanan
            '====================================================
            Select Case strCetak2
                Case "LapRekapKPSkpHr"
                    adocomd.CommandText = strSQL
                    Report.Database.AddADOCommand dbConn, adocomd
                    With Report
                        .usJK.SetUnboundFieldSource ("{Ado.JK}")
                        .usRuangan.SetUnboundFieldSource ("{Ado.RuanganPelayanan}")
                        .UsKasusPenyakit.SetUnboundFieldSource ("{Ado.Detail}")
                        .UnboundDateTime1.SetUnboundFieldSource Format(waktu, "dd mmm yyyy")
                        .UsJenis.SetUnboundFieldSource ("{Ado.Judul}")
                        .unJumlah.SetUnboundFieldSource ("{Ado.JmlPasien}")
                        .txtJudul.SetText judul9
                        If Format(mdTglAwal, "dd/MMM/yyyy") = Format(mdTglAkhir, "dd/MMM/yyyy") Then
                            .txtBulan.SetText "Tanggal    : " & Format(mdTglAwal, "dd MMMM YYYY")
                        Else
                            .txtBulan.SetText "Periode    : " & Format(mdTglAwal, "dd/MM/yyyy") & "     s/d     " & Format(mdTglAkhir, "dd/MM/yyyy")
                        End If

                    End With
                    '=====================================================
                    '    LAPORAN PER BULAN b PSN status dan Kelas Pelayanan
                    '=====================================================
                Case "LapRekapKPSkpBln"
                    adocomd.CommandText = strSQL
                    Report.Database.AddADOCommand dbConn, adocomd
                    With Report
                        .usJK.SetUnboundFieldSource ("{Ado.JK}")
                        .usRuangan.SetUnboundFieldSource ("{Ado.RuanganPelayanan}")
                        .UsKasusPenyakit.SetUnboundFieldSource ("{Ado.Detail}")
                        .UnboundDateTime1.SetUnboundFieldSource Format(waktu, "mmm yyyy")
                        .UsJenis.SetUnboundFieldSource ("{Ado.Judul}")
                        .unJumlah.SetUnboundFieldSource ("{Ado.JmlPasien}")
                        .txtJudul.SetText judul10
                        If Format(mdTglAwal, "MMM/yyyy") = Format(mdTglAkhir, "MMM/yyyy") Then
                            .txtBulan.SetText "Bulan    : " & Format(mdTglAwal, "MMMM YYYY")
                        Else
                            .txtBulan.SetText "Periode    : " & Format(mdTglAwal, "MMMM/yyyy") & "     s/d     " & Format(mdTglAkhir, "MMMM/yyyy")
                        End If
                    End With
            End Select
            '5 - STATUS PASIEN DAN KONDISI PULANG
        Case "laprekapkpps"
            '====================================================
            '    LAPORAN PER HARI b Status dan kondisi pulang
            '====================================================
            Select Case strCetak2
                Case "laprekapkppshr"
                    adocomd.CommandText = strSQL
                    Report.Database.AddADOCommand dbConn, adocomd
                    With Report
                        .usJK.SetUnboundFieldSource ("{Ado.JK}")
                        .usRuangan.SetUnboundFieldSource ("{Ado.RuanganPelayanan}")
                        .UsKasusPenyakit.SetUnboundFieldSource ("{Ado.Detail}")
                        .UnboundDateTime1.SetUnboundFieldSource Format("{Ado.tglKeluar}")
                        .UsJenis.SetUnboundFieldSource ("{Ado.Judul}")
                        .unJumlah.SetUnboundFieldSource ("{Ado.JmlPasien}")
                        .txtJudul.SetText judul11
                        If Format(mdTglAwal, "dd/MMM/yyyy") = Format(mdTglAkhir, "dd/MMM/yyyy") Then
                            .txtBulan.SetText "Tanggal    : " & Format(mdTglAwal, "dd MMMM YYYY")
                        Else
                            .txtBulan.SetText "Periode    : " & Format(mdTglAwal, "dd/MM/yyyy") & "     s/d     " & Format(mdTglAkhir, "dd/MM/yyyy")
                        End If
                    End With
                    '=====================================================
                    '    LAPORAN PER BULAN b PSN status dan Kondisi Pulang
                    '=====================================================
                Case "laprekapkppsbln"
                    adocomd.CommandText = strSQL
                    Report.Database.AddADOCommand dbConn, adocomd
                    With Report
                        .usJK.SetUnboundFieldSource ("{Ado.JK}")
                        .usRuangan.SetUnboundFieldSource ("{Ado.RuanganPelayanan}")
                        .UsKasusPenyakit.SetUnboundFieldSource ("{Ado.Detail}")
                        .UnboundDateTime1.SetUnboundFieldSource Format("{Ado.tglkeluar}")
                        .UsJenis.SetUnboundFieldSource ("{Ado.Judul}")
                        .unJumlah.SetUnboundFieldSource ("{Ado.JmlPasien}")
                        .txtJudul.SetText judul12
                        If Format(mdTglAwal, "MMM/yyyy") = Format(mdTglAkhir, "MMM/yyyy") Then
                            .txtBulan.SetText "Bulan    : " & Format(mdTglAwal, "MMMM YYYY")
                        Else
                            .txtBulan.SetText "Periode    : " & Format(mdTglAwal, "MMMM/yyyy") & "     s/d     " & Format(mdTglAkhir, "MMMM/yyyy")
                        End If
                    End With
            End Select

            '6 - JENIS PENYAKIT DAN KONDISI PULANG
        Case "laprekapkondisiplgdanKP"
            '====================================================
            '    LAPORAN PER HARI b Jenis Penyakit dan kondisi pulang
            '====================================================
            Select Case strCetak2
                Case "laprekapkondisiplgdanKPhr"
                    adocomd.CommandText = strSQL
                    Report.Database.AddADOCommand dbConn, adocomd
                    With Report
                        .usJK.SetUnboundFieldSource ("{Ado.JK}")
                        .usRuangan.SetUnboundFieldSource ("{Ado.RuanganPelayanan}")
                        .UsKasusPenyakit.SetUnboundFieldSource ("{Ado.Detail}")
                        .UnboundDateTime1.SetUnboundFieldSource Format("{Ado.tglKeluar}")
                        .UsJenis.SetUnboundFieldSource ("{Ado.Judul}")
                        .unJumlah.SetUnboundFieldSource ("{Ado.JmlPasien}")
                        .txtJudul.SetText Judul13
                        If Format(mdTglAwal, "dd/MMM/yyyy") = Format(mdTglAkhir, "dd/MMM/yyyy") Then
                            .txtBulan.SetText "Tanggal    : " & Format(mdTglAwal, "dd MMMM YYYY")
                        Else
                            .txtBulan.SetText "Periode    : " & Format(mdTglAwal, "dd/MM/yyyy") & "     s/d     " & Format(mdTglAkhir, "dd/MM/yyyy")
                        End If
                    End With
                    '=====================================================
                    '    LAPORAN PER BULAN b PSN Jenis Penyakit dan Kondisi Pulang
                    '=====================================================
                Case "laprekapkondisiplgdanKPbln"
                    adocomd.CommandText = strSQL
                    Report.Database.AddADOCommand dbConn, adocomd
                    With Report
                        .usJK.SetUnboundFieldSource ("{Ado.JK}")
                        .usRuangan.SetUnboundFieldSource ("{Ado.RuanganPelayanan}")
                        .UsKasusPenyakit.SetUnboundFieldSource ("{Ado.Detail}")
                        .UnboundDateTime1.SetUnboundFieldSource Format("{Ado.tglkeluar}")
                        .UsJenis.SetUnboundFieldSource ("{Ado.Judul}")
                        .unJumlah.SetUnboundFieldSource ("{Ado.JmlPasien}")
                        .txtJudul.SetText judul14
                        If Format(mdTglAwal, "MMM/yyyy") = Format(mdTglAkhir, "MMM/yyyy") Then
                            .txtBulan.SetText "Bulan    : " & Format(mdTglAwal, "MMMM YYYY")
                        Else
                            .txtBulan.SetText "Periode    : " & Format(mdTglAwal, "MMMM/yyyy") & "     s/d     " & Format(mdTglAkhir, "MMMM/yyyy")
                        End If

                    End With
            End Select
    End Select
    CRViewer1.ReportSource = Report
    CRViewer1.ViewReport
    CRViewer1.Zoom (75)
    Screen.MousePointer = 0
    Exit Sub
errPrint:
    MsgBox "Error cetak!" & vbNewLine & Err.Number & " - " & Err.Description, vbCritical, "Validasi"
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Form_Resize()
    CRViewer1.Width = Me.ScaleWidth
    CRViewer1.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCtkLapRekap_Viewer = Nothing
End Sub

Private Sub subDataRSU()
    On Error Resume Next
End Sub
