VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRKP_RJ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Kunjungan Pasien Berdasarkan Status & Kasus Penyakit Pasien"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRKP_RJ.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   9405
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
      TabIndex        =   7
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
         Height          =   735
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   8895
         Begin VB.OptionButton optGroupBy 
            Caption         =   "Tahun"
            Height          =   210
            Index           =   2
            Left            =   2595
            TabIndex        =   11
            Top             =   360
            Width           =   855
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Group By"
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   840
            TabIndex        =   10
            Top             =   120
            Width           =   2655
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
            Left            =   3600
            TabIndex        =   2
            Top             =   240
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
            Format          =   59179011
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   6240
            TabIndex        =   3
            Top             =   240
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
            Format          =   59179011
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
            Left            =   5880
            TabIndex        =   9
            Top             =   315
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
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   2160
      Width           =   9405
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Spreadsheet"
         Height          =   375
         Left            =   5640
         TabIndex        =   4
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   7440
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   930
      Left            =   -790
      Picture         =   "frmRKP_RJ.frx":08CA
      Top             =   0
      Width           =   10200
   End
End
Attribute VB_Name = "frmRKP_RJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCetak_Click()
    cmdCetak.Enabled = False
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value
    
    mblnGrafik = False
    '******************************************************
    'PILIHAN
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If optGroupBy(0).Value = True Then
        If strCetak = "LapRekapKPSJ" Then
            strSQL = "SELECT *  FROM   V_DatakunjunganPasienMasukBjenisBstausPasien " _
            & "WHERE (TglPendaftaran BETWEEN '" _
            & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
            & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') and  KdRuanganPelayanan = '" & mstrKdRuangan & "' "
            strcatak2 = "LapRekapKPSJhr"
        End If
    ElseIf optGroupBy(1).Value = True Then
        If strCetak = "LapRekapKPSJ" Then
            strSQL = "SELECT { fn MONTHNAME (TglPendaftaran) } AS TglPendaftaran, RuanganPelayanan, Judul, Detail, JK, JmlPasien, KdInstalasi  FROM   V_DatakunjunganPasienMasukBjenisBstausPasien " _
            & "WHERE (TglPendaftaran BETWEEN '" _
            & Format(mdTglAwal, "yyyy/MM/01 00:00:00") & "' AND '" _
            & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') and  KdRuanganPelayanan = '" & mstrKdRuangan & "' "
            strcatak2 = "LapRekapKPSJbln"
        End If
    End If
    
    msubRecFO rs, strSQL
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbExclamation, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
'**********************************************************
'SET JUDUL
'**********************************************************
    Set frmCtkLapRekap_Viewer = Nothing
'    frmCtkLapRekap_Viewer.Show
    If strcatak2 = "LapRekapKPSJhr" Then
        frmCtkLapRekap_Viewer.Caption = "Medifirst2000 - Laporan Rekapitulasi Pasien Per Status & Jenis (hari)"
    ElseIf strcatak2 = "LapRekapKPSJbln" Then
        frmCtkLapRekap_Viewer.Caption = "Medifirst2000 - Laporan Rekapitulasi Pasien Per Status & Jenis (bulan)"
    Else
        
    End If
    cmdCetak.Enabled = True
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
    If KeyCode = 13 Then cmdCetak.SetFocus
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
   
End Sub

Private Sub optGroupBy_Click(Index As Integer)
    If optGroupBy(0).Value = True Then
       dtpAwal.CustomFormat = "dd MMMM yyyy 00:00:00"
       dtpAkhir.CustomFormat = "dd MMMM yyyy 23:59:59"
    ElseIf optGroupBy(1).Value = True Then
      dtpAwal.CustomFormat = "MMMM yyyy 00:00:00"
      dtpAkhir.CustomFormat = "MMMM yyyy 23:59:59"
    End If
End Sub

Private Sub optGroupBy_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAwal.SetFocus
End Sub

'user defined procedure(s) & function(s)



