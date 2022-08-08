VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLapRekPerDokter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Kunjungan Pasien "
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9420
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLapRekPerDokter.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   9420
   Begin VB.Frame fraPeriode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   0
      TabIndex        =   12
      Top             =   960
      Width           =   9405
      Begin VB.Frame Frame4 
         Caption         =   "Periode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   9135
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            Caption         =   "Group By"
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   3735
            Begin VB.OptionButton optGroupBy 
               Caption         =   "Total"
               Height          =   210
               Index           =   3
               Left            =   2640
               TabIndex        =   3
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optGroupBy 
               Caption         =   "Tahun"
               Height          =   210
               Index           =   2
               Left            =   1680
               TabIndex        =   2
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optGroupBy 
               Caption         =   "Hari"
               Height          =   210
               Index           =   0
               Left            =   120
               TabIndex        =   0
               Top             =   230
               Value           =   -1  'True
               Width           =   615
            End
            Begin VB.OptionButton optGroupBy 
               Caption         =   "Bulan"
               Height          =   210
               Index           =   1
               Left            =   840
               TabIndex        =   1
               Top             =   230
               Width           =   735
            End
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   3960
            TabIndex        =   4
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            OLEDropMode     =   1
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   60686339
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   6600
            TabIndex        =   5
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   60686339
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6240
            TabIndex        =   14
            Top             =   360
            Width           =   255
         End
      End
   End
   Begin VB.Frame fraButton 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   11
      Top             =   2160
      Width           =   9405
      Begin VB.OptionButton optkonsul 
         Caption         =   "Pasien Konsul"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   3735
      End
      Begin VB.OptionButton optpoli 
         Caption         =   "Pasien Poli"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   3735
      End
      Begin VB.CommandButton CmdGrafik 
         Caption         =   "&Grafik"
         Height          =   615
         Left            =   4200
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   615
         Left            =   5880
         TabIndex        =   9
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   615
         Left            =   7560
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   1800
      _cx             =   3175
      _cy             =   1720
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7560
      Picture         =   "frmLapRekPerDokter.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmLapRekPerDokter.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmLapRekPerDokter.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmLapRekPerDokter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub filter_kriteria()
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value
    Dim mdtBulan As Integer
    Dim MdtTahun As Integer
    If optpoli.Value = True Then
        strCetak = "LapKunjunganPerDokter"
    Else
        strCetak = "LapKunjunganKonsulPerDokter"
    End If

    If optGroupBy(0).Value = True Then
        Select Case strCetak
            Case "LapKunjunganPerDokter"
                strCetak2 = "LapKunjunganPerDokterHari"
                strSQL = "Select * from V_RekapPasienRJPerDokter " & _
                "WHERE (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and  KdRuangan='" & mstrKdRuangan & "'"

            Case "LapKunjunganKonsulPerDokter"
                strCetak2 = "LapKunjunganKonsulPerDokterHari"
                strSQL = "Select KdRuangan,KdRuanganAsal,KdRuanganTujuan,TglDirujuk,Judul,DokterPemeriksa,Detail,JK,JmlPasien from V_RekapPasienKonsulRJPerDokter " & _
                "WHERE (TglDirujuk BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and  KdRuangan='" & mstrKdRuangan & "' and KdRuanganAsal<>'" & mstrKdRuangan & "' AND KdRuanganTujuan='" & mstrKdRuangan & "'"

        End Select

    ElseIf optGroupBy(1).Value = True Then
        mdTglAwal = CDate(Format(dtpAwal.Value, "yyyy-mm ") & "-01 00:00:00") 'TglAwal
        mdtBulan = CStr(Format(dtpAkhir.Value, "mm"))
        MdtTahun = CStr(Format(dtpAkhir.Value, "yyyy"))
        mdTglAkhir = CDate(Format(dtpAkhir.Value, "yyyy-mm") & "-" & funcHitungHari(mdtBulan, MdtTahun) & " 23:59:59")
        Select Case strCetak

            Case "LapKunjunganPerDokter"
                strCetak2 = "LapKunjunganPerDokterBulan"
                strSQL = "Select dbo.FB_TakeBlnThn(TglPendaftaran) as tglPendaftaran,Dokter,Judul,Detail,JK,JmlPasien,KdRuangan from V_RekapPasienRJPerDokter " & _
                "WHERE (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and  KdRuangan='" & mstrKdRuangan & "'"
            Case "LapKunjunganKonsulPerDokter"
                strCetak2 = "LapKunjunganKonsulPerDokterBulan"
                strSQL = "Select KdRuangan,KdRuanganAsal,KdRuanganTujuan,TglDirujuk,Judul,DokterPemeriksa,Detail,JK,JmlPasien from V_RekapPasienKonsulRJPerDokter " & _
                "WHERE (TglDirujuk BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and  KdRuangan='" & mstrKdRuangan & "' and KdRuanganAsal<>'" & mstrKdRuangan & "'AND KdRuanganTujuan='" & mstrKdRuangan & "'"
        End Select
    ElseIf optGroupBy(2).Value = True Then
        mdTglAwal = CDate("01-01-" & Format(dtpAwal.Value, "yyyy 00:00:00")) 'TglAwal
        mdTglAkhir = CDate("31-12-" & Format(dtpAkhir.Value, "yyyy 23:59:59")) 'TglAkhir
        Select Case strCetak

            Case "LapKunjunganPerDokter"
                strCetak2 = "LapKunjunganPerDokterTahun"
                strSQL = "Select * from V_RekapPasienRJPerDokter " & _
                "WHERE (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and  KdRuangan='" & mstrKdRuangan & "'"
            Case "LapKunjunganKonsulPerDokter"
                strCetak2 = "LapKunjunganKonsulPerDokterTahun"
                strSQL = "Select KdRuangan,KdRuanganAsal,KdRuanganTujuan,TglDirujuk,Judul,DokterPemeriksa,Detail,JK,JmlPasien from V_RekapPasienKonsulRJPerDokter " & _
                "WHERE (TglDirujuk BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and  KdRuangan='" & mstrKdRuangan & "' and KdRuanganAsal<>'" & mstrKdRuangan & "'AND KdRuanganTujuan='" & mstrKdRuangan & "'"

        End Select
    ElseIf optGroupBy(3).Value = True Then
        Select Case strCetak

            Case "LapKunjunganPerDokter"
                strCetak2 = "LapKunjunganPerDokterTotal"
                strSQL = "Select * from V_RekapPasienRJPerDokter " & _
                "WHERE (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and  KdRuangan='" & mstrKdRuangan & "'"
            Case "LapKunjunganKonsulPerDokter"
                strCetak2 = "LapKunjunganKonsulPerDokterTotal"
                strSQL = "Select KdRuangan,KdRuanganAsal,KdRuanganTujuan,TglDirujuk,Judul,DokterPemeriksa,Detail,JK,JmlPasien from V_RekapPasienKonsulRJPerDokter " & _
                "WHERE (TglDirujuk BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and  KdRuangan='" & mstrKdRuangan & "' and KdRuanganAsal<>'" & mstrKdRuangan & "'AND KdRuanganTujuan='" & mstrKdRuangan & "'"
        End Select
    End If
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell
    Call filter_kriteria
    If ValidasiTanggal(dtpAwal, dtpAkhir) = False Then Exit Sub
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then MsgBox "Data Tidak Ada", vbExclamation, "Validasi": Exit Sub
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    FrmCetakLapKunjunganPasien.Show
hell:
End Sub

Private Sub cmdgrafik_Click()
    Call filter_kriteria
    If ValidasiTanggal(dtpAwal, dtpAkhir) = False Then Exit Sub
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then MsgBox "Data Tidak Ada", vbExclamation, "Validasi": Exit Sub
    FrmCetakLaporandalamBentukGrafik.Show
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcInstalasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then optGroupBy(0).SetFocus
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then optpoli.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    With Me
        .dtpAwal.Value = Now
        .dtpAkhir.Value = Now
    End With

    Call cekOpt
    Call PlayFlashMovie(Me)
End Sub

Private Sub cekOpt()
    If optGroupBy(0).Value = True Then
        Call optGroupBy_Click(0)
    ElseIf optGroupBy(1).Value = True Then
        Call optGroupBy_Click(1)
    ElseIf optGroupBy(2).Value = True Then
        Call optGroupBy_Click(2)
    End If
End Sub

Private Sub optGroupBy_Click(Index As Integer)
    Select Case Index
        Case 0
            dtpAwal.CustomFormat = "dd MMMM yyyyy"
            dtpAkhir.CustomFormat = "dd MMMM yyyyy"
            optGroupBy(1).Value = False
            optGroupBy(2).Value = False
            optGroupBy(3).Value = False

        Case 1
            dtpAkhir.CustomFormat = "MMMM yyyyy"
            dtpAwal.CustomFormat = "MMMM yyyyy"
            optGroupBy(0).Value = False
            optGroupBy(2).Value = False
            optGroupBy(3).Value = False

        Case 2
            dtpAkhir.CustomFormat = "yyyyy"
            dtpAwal.CustomFormat = "yyyyy"
            optGroupBy(0).Value = False
            optGroupBy(1).Value = False
            optGroupBy(3).Value = False
        Case 3
            dtpAwal.CustomFormat = "dd MMMM yyyyy"
            dtpAkhir.CustomFormat = "dd MMMM yyyyy"
            optGroupBy(0).Value = False
            optGroupBy(1).Value = False
            optGroupBy(2).Value = False
    End Select
End Sub

Private Sub optGroupBy_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAwal.SetFocus
End Sub

Private Sub optkonsul_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetak.SetFocus
End Sub

Private Sub optpoli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetak.SetFocus
End Sub

