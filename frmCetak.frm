VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetak 
   Caption         =   "Cetak"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   Icon            =   "frmCetak.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   5835
   WindowState     =   2  'Maximized
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
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   -1  'True
   End
End
Attribute VB_Name = "frmCetak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bolSuppresDetailSection10 As Boolean

Private Sub Form_Load()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetak = Nothing
End Sub

Public Sub CetakUlang()
    bolSuppresDetailSection10 = True
    Call CetakUlangJenisKuitansi
End Sub

Public Sub CetakUlangJenisKuitansi()
    On Error GoTo errLoad
    Dim rptJnsKuitansi As New crJenisKuitansi
    Set rptJnsKuitansi = New crJenisKuitansi

    strSQL = "SELECT NoPendaftaran FROM V_RincianTotalDetailBiayaPelayananNotNull " _
    & "WHERE (NoStruk = '" & mstrNoStruk & "')"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "No Struk yang dimasukkan tidak terdaftar", vbCritical + vbOKOnly, "Validasi"
        Exit Sub
    End If
    mstrNoPen = rs(0).Value
    strSQL = "SELECT * FROM V_JudulStrukPembayaranPasien " _
    & "WHERE (NoStruk = '" & mstrNoStruk & "')  AND NoBKM = '" & mstrNoBKM & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    With rptJnsKuitansi
        .txtNoCM.SetText rs("NoCM").Value
        .txtNamaPasien.SetText rs("NamaPasien").Value & " / " & IIf(rs("JK").Value = "P", "Wanita", "Pria")
        .txtUmur.SetText rs("Umur").Value
        If Not IsNull(rs("AlamatLengkap")) Then
            .txtAlamatPasien.SetText rs("AlamatLengkap")
        Else
            .txtAlamatPasien.SetText ""
        End If
        .txtTglDaftar.SetText rs("TglMasuk").Value
        .txtTglStruk.SetText rs("TglStruk").Value
        .txtLamaDirawat.SetText "( " & DateDiff("d", rs("TglMasuk").Value, rs("TglStruk").Value) & " hr )"
        .txtNoBKM.SetText rs("NoBKM").Value
        .txtNoBKMRuangan.SetText IIf(IsNull(rs("NoUrut")), "", rs("NoUrut"))
        .txtTglBKM.SetText rs("TglBKM").Value
        .txtPrintTglBKM.SetText strNKotaRS & ", " & Format(rs("TglBKM"), "dd MMMM yyyy")
        .txtPenjamin.SetText rs("NamaPenjamin").Value
        .txtNoKartu.SetText IIf(IsNull(rs("NoKartu")), "", rs("NoKartu"))
        .txtJenisPasien.SetText rs("JenisPasien").Value
        .txtPembayaran.SetText rs("CaraBayar").Value

        .txtBiayaAdministrasi.SetText Format(rs("Administrasi").Value, "#,###.00")

        Set dbcmd = New ADODB.Command
        Set dbcmd.ActiveConnection = dbConn

        dbcmd.CommandText = "sELECT * FROM V_RincianTotalDetailBiayaPelayananNotNull " _
        & "WHERE (NoPendaftaran = '" & mstrNoPen & "') AND (NoStruk='" & mstrNoStruk & "')"

        dbcmd.CommandType = adCmdUnknown

        .Database.AddADOCommand dbConn, dbcmd
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strWebsite & ", " & strEmail
        .txtNoPendaftaran.SetText mstrNoPen
        .txtNoStruk.SetText mstrNoStruk
        .txtIdPetugas.SetText noidpegawai
        .txtPetugasKasir.SetText strNmPegawai
        .txtNamaRuangan.SetText IIf(IsNull(rs("RuanganKasir")), "", rs("RuanganKasir"))
        .txtPembayaranKe.SetText rs("PembayaranKe").Value

        .txtBayar.SetText Format(rs("JmlHarusDibayar").Value, "#,###.00")
        .txtPembebasan.SetText Format(rs("JmlPembebasan").Value, "#,###.00")
        .txtTanggungan.SetText Format(rs("JmlHutangPenjamin").Value, "#,###.00")
        .txtTanggunganRS.SetText Format(rs("JmlTanggunganRS").Value, "#,###.00")
        .txtSisaTagihan.SetText Format(rs("SisaTagihan").Value, "#,###.00")
        .txtJmlDibayar.SetText Format(rs("JmlBayar").Value, "#,###.00")
        .txtPembayaranSebelumnya.SetText Format(rs("JmlBayarSebelumnya").Value, "#,###.00")
        .txtTotalSudahDibayar.SetText Format(rs("JmlSudahDibayar").Value, "#,###.00")
        .txtTerbilang.SetText NumToText(CCur(rs("JmlBayar").Value))

        .usJenisPelayanan.SetUnboundFieldSource ("{ado.Jenis_Item}")
        .ucBiaya.SetUnboundFieldSource ("{ado.Tarif}") '("{ado.Harga_Item}")
        .ucTarifCito.SetUnboundFieldSource ("{ado.TarifCito}")
        .ucTarifTotal.SetUnboundFieldSource ("{ado.BiayaTotal}")
        .udTanggal.SetUnboundFieldSource ("{ado.TglPelayanan}")
        .usLayanan.SetUnboundFieldSource ("{ado.Nama_Item}")
        .unQty.SetUnboundFieldSource ("{ado.Jml_Item}")
        .usKelas.SetUnboundFieldSource ("{ado.Kelas}")
        .usRuangan.SetUnboundFieldSource ("{ado.Ruangan}")

        If bolSuppresDetailSection10 = True Then
            .Text50.Suppress = False
            .Text51.Suppress = False
            .Text52.Suppress = False
            .Text43.Suppress = False
            .Section10.Suppress = False
        Else
            .Text50.Suppress = True
            .Text51.Suppress = True
            .Text52.Suppress = True
            .Text43.Suppress = True
            .Section10.Suppress = True
        End If
        bolSuppresDetailSection10 = False

        Dim adojenis As New ADODB.Command
        Set adojenis = New ADODB.Command
        adojenis.CommandText = "select * from " & _
        " V_RincianBiayaPelayananTMPerKomponenNPBS " & _
        " WHERE (NoPendaftaran = '" & mstrNoPen & "') AND (NoStruk='" & mstrNoStruk & "')"

        adojenis.CommandType = adCmdText
        rptJnsKuitansi.Subreport1.OpenSubreport.Database.AddADOCommand dbConn, adojenis
        With rptJnsKuitansi
            .Subreport1_usNamaKomponen.SetUnboundFieldSource ("{ado.KomponenTarif}")
            .Subreport1_usNamaPelayanan.SetUnboundFieldSource ("{ado.Nama_Item}")
            .Subreport1_ucJasaKomponen.SetUnboundFieldSource ("{ado.Harga_Item}")
            .Subreport1_usRuangan.SetUnboundFieldSource ("{ado.Ruangan}")
        End With

        If vLaporan = "Print" Then
            .PrintOut False
            Unload Me
            Screen.MousePointer = vbDefault
        Else
            With CRViewer1
                .ReportSource = rptJnsKuitansi
                .ViewReport
                .Zoom 1
            End With
            Screen.MousePointer = vbDefault
        End If

    End With
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

