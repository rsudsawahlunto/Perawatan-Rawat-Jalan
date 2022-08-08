VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRekapKunjunganPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Rekapitulasi Kunjungan Pasien"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRekapKunjunganPasien.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   7575
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   0
      TabIndex        =   13
      Top             =   7320
      Width           =   7575
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   4680
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   6120
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Height          =   4335
      Left            =   0
      TabIndex        =   11
      Top             =   3000
      Width           =   7575
      Begin MSDataGridLib.DataGrid dgRekapKunjungan 
         Height          =   3975
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   7011
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Periode Laporan"
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
      Top             =   2280
      Width           =   7575
      Begin VB.CommandButton cmdCari 
         Caption         =   "&Cari"
         Height          =   375
         Left            =   6600
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   360
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   22675459
         CurrentDate     =   38210
      End
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   360
         Left            =   4200
         TabIndex        =   8
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   22675459
         CurrentDate     =   38210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "s/d"
         Height          =   210
         Left            =   3840
         TabIndex        =   9
         Top             =   315
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Kriteria Kunjungan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   7575
      Begin VB.OptionButton optPasienMasuk 
         Caption         =   "Pasien Masuk Rawat Inap"
         Height          =   255
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   2415
      End
      Begin VB.OptionButton optPasienKeluar 
         Caption         =   "Pasien Keluar Rawat Inap"
         Height          =   255
         Left            =   4920
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Kriteria Laporan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   7575
      Begin VB.OptionButton optStatusPasien 
         Caption         =   "Berdasarkan Status Pasien"
         Height          =   255
         Left            =   4920
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
      Begin VB.OptionButton optJenisPasien 
         Caption         =   "Berdasarkan Jenis Pasien"
         Height          =   255
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   -1920
      Picture         =   "frmRekapKunjunganPasien.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmRekapKunjunganPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCari_Click()
    If (optJenisPasien.Value = True) And (optPasienMasuk.Value = True) Then
        strSQL = "select JenisPasien,sum(JmlPasienPria) as JmlPria,sum(JmlPasienWanita)as JmlWanita,sum(Total)as Total from V_RekapitulasiPasienBJenis where KdRuangan='" & mstrKdRuangan & "' and TglPendaftaran between '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' group by JenisPasien"
    Else
    If (optJenisPasien.Value = True) And (optPasienKeluar.Value = True) Then
        strSQL = "select JenisPasien,sum(JmlPasienPria) as JmlPria,sum(JmlPasienWanita)as JmlWanita,sum(Total)as Total from V_RekapitulasiPasienKeluarRS where KdRuangan='" & mstrKdRuangan & "' and TglKeluar between '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' group by JenisPasien"
    Else
    If (optStatusPasien.Value = True) And (optPasienMasuk.Value = True) Then
        strSQL = "select StatusPasien,sum(JmlPasienPria) as JmlPria,sum(JmlPasienWanita)as JmlWanita,sum(Total)as Total from V_RekapitulasiPasienBStatus where KdRuangan='" & mstrKdRuangan & "' and TglPendaftaran between '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' group by StatusPasien"
    End If
    End If
    End If
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenStatic, adLockReadOnly
    Set dgRekapKunjungan.DataSource = rs
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    optJenisPasien.Value = True
    optPasienMasuk.Value = True
End Sub

