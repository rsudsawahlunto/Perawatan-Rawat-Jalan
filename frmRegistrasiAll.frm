VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmRegistrasiAll 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Registrasi Pasien"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegistrasiAll.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   12270
   Begin VB.TextBox txtNoBKM 
      Height          =   375
      Left            =   2640
      TabIndex        =   78
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame5 
      Caption         =   "Data Penanggungjawab Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   65
      Top             =   5640
      Width           =   12255
      Begin VB.CheckBox chkDiriSendiri 
         Caption         =   "&Diri Sendiri"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   10560
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtTlpRI 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6960
         MaxLength       =   50
         TabIndex        =   33
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox txtAlamatRI 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8160
         MaxLength       =   50
         TabIndex        =   26
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox txtNamaRI 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         MaxLength       =   20
         TabIndex        =   22
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtKodePos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5280
         MaxLength       =   5
         TabIndex        =   32
         Top             =   2040
         Width           =   1455
      End
      Begin MSMask.MaskEdBox meRTRWPJ 
         Height          =   390
         Left            =   4200
         TabIndex        =   31
         Top             =   2040
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo dcKotaPJ 
         Height          =   390
         Left            =   3960
         TabIndex        =   28
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcKecamatanPJ 
         Height          =   390
         Left            =   8040
         TabIndex        =   29
         Top             =   1320
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcKelurahanPJ 
         Height          =   390
         Left            =   120
         TabIndex        =   30
         Top             =   2040
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcPropinsiPJ 
         Height          =   390
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcHubungan 
         Height          =   390
         Left            =   3000
         TabIndex        =   23
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcPekerjaanPJ 
         Height          =   390
         Left            =   5520
         TabIndex        =   24
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Pekerjaan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5520
         TabIndex        =   76
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hubungan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   3000
         TabIndex        =   75
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "Telephone"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   6960
         TabIndex        =   74
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "Alamat Lengkap"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   8160
         TabIndex        =   73
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "Nama Lengkap"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   120
         TabIndex        =   72
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Kode Pos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5280
         TabIndex        =   71
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "RT/RW"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4200
         TabIndex        =   70
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Kelurahan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   69
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Kecamatan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8040
         TabIndex        =   68
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Propinsi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   67
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Kota/Kabupaten"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3960
         TabIndex        =   66
         Top             =   1080
         Width           =   1350
      End
   End
   Begin VB.TextBox txtNoPakai 
      Height          =   495
      Left            =   480
      TabIndex        =   64
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraRawatGabung 
      Caption         =   "Rawat Gabung ?"
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
      Left            =   10440
      TabIndex        =   62
      Top             =   4800
      Width           =   1695
      Begin VB.OptionButton optYa 
         Caption         =   "Ya"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optTidak 
         Caption         =   "Tidak"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin MSComctlLib.StatusBar stbInformasi 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   61
      Top             =   4335
      Visible         =   0   'False
      Width           =   12270
      _ExtentX        =   21643
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   2205
            MinWidth        =   1411
            Text            =   "Cetak Label (F1)"
            TextSave        =   "Cetak Label (F1)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   2558
            MinWidth        =   1764
            Text            =   "Pasien Baru Ctrl+B"
            TextSave        =   "Pasien Baru Ctrl+B"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   2205
            MinWidth        =   1411
            Text            =   "Cari Pasien (F3)"
            TextSave        =   "Cari Pasien (F3)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   2558
            MinWidth        =   1764
            Text            =   "Pasien Lama Ctrl+L"
            TextSave        =   "Pasien Lama Ctrl+L"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   3334
            Text            =   "Lembar Masuk Ctrl+R"
            TextSave        =   "Lembar Masuk Ctrl+R"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   3334
            Text            =   "Surat Keterangan Ctrl+Z"
            TextSave        =   "Surat Keterangan Ctrl+Z"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Enabled         =   0   'False
            Text            =   "Cetak SJP (F9)"
            TextSave        =   "Cetak SJP (F9)"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Text            =   "C. Medis Ctrl+M"
            TextSave        =   "C. Medis Ctrl+M"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraRegistrasiRI 
      Caption         =   "Data Masuk Rawat Inap"
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
      TabIndex        =   55
      Top             =   4680
      Width           =   12255
      Begin MSDataListLib.DataCombo dcCaraMasukRI 
         Height          =   390
         Left            =   2160
         TabIndex        =   15
         Top             =   480
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcKelasKamarRI 
         Height          =   390
         Left            =   5040
         TabIndex        =   16
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcNoKamarRI 
         Height          =   390
         Left            =   7560
         TabIndex        =   17
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcNoBedRI 
         Height          =   390
         Left            =   9600
         TabIndex        =   18
         Top             =   480
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "No. Bed"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   9600
         TabIndex        =   60
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "No. Kamar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   7560
         TabIndex        =   59
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "Kelas Kamar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   5040
         TabIndex        =   58
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "Cara Masuk"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   2160
         TabIndex        =   56
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   44
      Top             =   3840
      Width           =   12255
      Begin VB.CommandButton cmdRujukan 
         Caption         =   "&Data Rujukan"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6975
         TabIndex        =   35
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8670
         TabIndex        =   34
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10365
         TabIndex        =   36
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdAsuransiP 
         Caption         =   "&Asuransi Pasien"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5280
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Data Registrasi Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   37
      Top             =   2040
      Width           =   12255
      Begin MSDataListLib.DataCombo dcInstalasi 
         Height          =   390
         Left            =   2280
         TabIndex        =   8
         Top             =   600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcRuangan 
         Height          =   390
         Left            =   270
         TabIndex        =   11
         Top             =   1320
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   688
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcKelas 
         Height          =   390
         Left            =   8880
         TabIndex        =   10
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpTglPendaftaran 
         Height          =   360
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   51380227
         UpDown          =   -1  'True
         CurrentDate     =   38061
      End
      Begin MSDataListLib.DataCombo dcKelompokPasien 
         Height          =   390
         Left            =   9840
         TabIndex        =   14
         Top             =   1320
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcJenisKelas 
         Height          =   390
         Left            =   5760
         TabIndex        =   9
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcSubInstalasi 
         Height          =   390
         Left            =   3570
         TabIndex        =   12
         Top             =   1320
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcRujukanRI 
         Height          =   390
         Left            =   7200
         TabIndex        =   13
         Top             =   1320
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "Asal Rujukan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   7260
         TabIndex        =   63
         Top             =   1080
         Width           =   1110
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "SMF (Kasus Penyakit)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   3570
         TabIndex        =   57
         Top             =   1080
         Width           =   1845
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelas Pelayanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5760
         TabIndex        =   52
         Top             =   360
         Width           =   1860
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Penjamin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9840
         TabIndex        =   46
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Pendaftaran"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   45
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Kelas Pelayanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8880
         TabIndex        =   40
         Top             =   360
         Width           =   1380
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Ruangan Pemeriksaan / Perawatan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   270
         TabIndex        =   39
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Instalasi Pemeriksaan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2320
         TabIndex        =   38
         Top             =   360
         Width           =   1860
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   41
      Top             =   960
      Width           =   12255
      Begin VB.CheckBox chkDetailPasien 
         Caption         =   "Detail Pasien"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10680
         TabIndex        =   6
         Top             =   550
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         Caption         =   "Umur"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   850
         Left            =   7920
         TabIndex        =   47
         Top             =   150
         Width           =   2535
         Begin VB.TextBox txtHr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1800
            MaxLength       =   6
            TabIndex        =   5
            Top             =   330
            Width           =   375
         End
         Begin VB.TextBox txtBln 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   960
            MaxLength       =   6
            TabIndex        =   4
            Top             =   330
            Width           =   375
         End
         Begin VB.TextBox txtThn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            MaxLength       =   6
            TabIndex        =   3
            Top             =   330
            Width           =   375
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2280
            TabIndex        =   50
            Top             =   360
            Width           =   195
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1440
            TabIndex        =   49
            Top             =   360
            Width           =   270
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   600
            TabIndex        =   48
            Top             =   360
            Width           =   315
         End
      End
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6240
         MaxLength       =   9
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         MaxLength       =   6
         TabIndex        =   0
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label lblJnsKlm 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6240
         TabIndex        =   51
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   43
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblNamaPasien 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   42
         Top             =   360
         Width           =   1350
      End
   End
   Begin VB.TextBox txtNoPendaftaran 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      MaxLength       =   10
      TabIndex        =   53
      Top             =   1200
      Width           =   1695
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   77
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
      Left            =   10440
      Picture         =   "frmRegistrasiAll.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRegistrasiAll.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRegistrasiAll.frx":3816
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "No. Pendaftaran"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   0
      TabIndex        =   54
      Top             =   960
      Width           =   1605
   End
End
Attribute VB_Name = "frmRegistrasiAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFilter As String
Dim intRowNow As Integer
Dim strSubInstalasi As String
Dim strNoAntrian As String
Dim dTglberlaku As Date
Dim curTarif As Currency
Dim curTP As Currency
Dim curTRS As Currency
Dim curPemb As Currency
Dim Qstrsql As String

Private Sub subLoadData()
    sRuangPeriksa = dcRuangan.Text
    sNamaPasien = txtNamaPasien.Text
    sJK = txtJK.Text
    sUmur = txtThn.Text & " th " & txtBln.Text & " bl " & txtHr.Text & " hr"
    sAlamat = ""
    sPenjamin = dcKelompokPasien.Text
    sKelas = dcJenisKelas.Text
    sNoBed = dcNoBedRI.Text
    iNoAntrian = strNoAntrian
End Sub

'Store procedure untuk mengisi struk billing pasien
Private Function sp_AddStrukBuktiKasMasuk() As Boolean
On Error GoTo errload
    
    sp_AddStrukBuktiKasMasuk = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("TglBKM", adDate, adParamInput, , Format(dtpTglPendaftaran.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdCaraBayar", adChar, adParamInput, 2, "01")
        .Parameters.Append .CreateParameter("KdJenisKartu", adChar, adParamInput, 2, Null)
        .Parameters.Append .CreateParameter("NamaBank", adVarChar, adParamInput, 100, Null)
        .Parameters.Append .CreateParameter("NoKartu", adVarChar, adParamInput, 50, Null)
        .Parameters.Append .CreateParameter("AtasNama", adVarChar, adParamInput, 50, Null)
       ' If CCur(txtJmlUang.Text) > CCur(lblTotalTagihan.Caption) Then
        '    .Parameters.Append .CreateParameter("JmlBayar", adCurrency, adParamInput, , CCur(lblTotalTagihan.Caption))
        'Else
        .Parameters.Append .CreateParameter("JmlBayar", adCurrency, adParamInput, , mcurAll_HrsDibyr)
        'End If
'        .Parameters.Append .CreateParameter("JmlBayar", adCurrency, adParamInput, , CCur(txtJmlUang.Text))
        .Parameters.Append .CreateParameter("Administrasi", adCurrency, adParamInput, , 0)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 100, Null)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, "176")
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, noidpegawai)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("OutputNoBKM", adChar, adParamOutput, 10, Null)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_StrukBuktiKasMasukPelayananPasien"
        .CommandType = adCmdStoredProc
        .Execute
    
        If .Parameters("RETURN_VALUE").Value <> 0 Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Struk Billing Pasien", vbCritical, "Validasi"
            sp_AddStrukBuktiKasMasuk = False
        Else
            If Not IsNull(.Parameters("OutputNoBKM").Value) Then txtNoBKM.Text = .Parameters("OutputNoBKM").Value
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    
Exit Function
errload:
    sp_AddStrukBuktiKasMasuk = False
    Call msubPesanError("-Add_StrukBuktiKasMasukPelayananPasien")
End Function
'Store procedure untuk mengisi struk billing pasien
Private Function sp_AddStruk(ByVal adoCommand As ADODB.Command, strStsByr As String) As Boolean
    On Error GoTo errload
    sp_AddStruk = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoBKM", adChar, adParamInput, 10, mstrNoBKM)
        .Parameters.Append .CreateParameter("OutputNoStruk", adChar, adParamOutput, 10, Null)
        .Parameters.Append .CreateParameter("TglStruk", adDate, adParamInput, , Format(dtpTglPendaftaran.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("NoCM", adChar, adParamInput, 6, mstrNoCM)
        .Parameters.Append .CreateParameter("KdKelompokPasien", adChar, adParamInput, 2, dcKelompokPasien.BoundText)
        If dcKelompokPasien.BoundText = "01" Then
            .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, "2222222222")
        Else
            .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, typAsuransi.strIdPenjamin)
        End If
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, "176")
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, noidpegawai)
        .Parameters.Append .CreateParameter("TotalBiaya", adCurrency, adParamInput, , CCur(mcurBayar))
        .Parameters.Append .CreateParameter("JmlHutangPenjamin", adCurrency, adParamInput, , CCur(mcurAll_TP))
        .Parameters.Append .CreateParameter("JmlTanggunganRS", adCurrency, adParamInput, , CCur(mcurAll_TRS))
        .Parameters.Append .CreateParameter("JmlPembebasan", adCurrency, adParamInput, , CCur(mcurAll_Pemb))
        .Parameters.Append .CreateParameter("JmlHrsDibayar", adCurrency, adParamInput, , CCur(mcurAll_HrsDibyr))
        .Parameters.Append .CreateParameter("JmlDiscount", adCurrency, adParamInput, , "0")
            
       
'       '------------------Begin Ditutup
'        'Pembayaran Tagihan Pasien
'        If CCur(txtJmlUang.Text) > CCur(lblTotalTagihan.Caption) Then
'            .Parameters.Append .CreateParameter("JmlBayar", adCurrency, adParamInput, , CCur(lblTotalTagihan.Caption))
'        Else
'            .Parameters.Append .CreateParameter("JmlBayar", adCurrency, adParamInput, , CCur(txtJmlUang.Text))
'        End If
'        .Parameters.Append .CreateParameter("JmlDiscount", adCurrency, adParamInput, , CCur(txtDiscount.Text))
'        .Parameters.Append .CreateParameter("SisaTagihan", adCurrency, adParamInput, , CCur(txtSisaTagihan.Text))
'        .Parameters.Append .CreateParameter("StatusBayar", adChar, adParamInput, 1, strStsByr)
'
'        'Pembayaran Tagihan utk Obat & Alkes
'        .Parameters.Append .CreateParameter("TotalBiayaOA", adCurrency, adParamInput, , mcurOA_TBP)
'        .Parameters.Append .CreateParameter("JmlBayarOA", adCurrency, adParamInput, , mcurOA_JmlByr)
'        .Parameters.Append .CreateParameter("JmlHutangPenjaminOA", adCurrency, adParamInput, , mcurOA_TP)
'        .Parameters.Append .CreateParameter("JmlTanggunganRSOA", adCurrency, adParamInput, , mcurOA_TRS)
'        .Parameters.Append .CreateParameter("JmlPembebasanOA", adCurrency, adParamInput, , mcurOA_Pemb)
'        .Parameters.Append .CreateParameter("JmlHrsDibayarOA", adCurrency, adParamInput, , mcurOA_HrsDibyr)
'        .Parameters.Append .CreateParameter("JmlDiscountOA", adCurrency, adParamInput, , mcurOA_Discount)
'        .Parameters.Append .CreateParameter("SisaTagihanOA", adCurrency, adParamInput, , mcurOA_ST)
'
'        'Pembayaran Tagihan utk Tindakan Medis
'        .Parameters.Append .CreateParameter("TotalBiayaTM", adCurrency, adParamInput, , mcurTM_TBP)
'        .Parameters.Append .CreateParameter("JmlBayarTM", adCurrency, adParamInput, , mcurTM_JmlByr)
'        .Parameters.Append .CreateParameter("JmlHutangPenjaminTM", adCurrency, adParamInput, , mcurTM_TP)
'        .Parameters.Append .CreateParameter("JmlTanggunganRSTM", adCurrency, adParamInput, , mcurTM_TRS)
'        .Parameters.Append .CreateParameter("JmlPembebasanTM", adCurrency, adParamInput, , mcurTM_Pemb)
'        .Parameters.Append .CreateParameter("JmlHrsDibayarTM", adCurrency, adParamInput, , mcurTM_HrsDibyr)
'        .Parameters.Append .CreateParameter("JmlDiscountTM", adCurrency, adParamInput, , mcurTM_Discount)
'        .Parameters.Append .CreateParameter("SisaTagihanTM", adCurrency, adParamInput, , mcurTM_ST)
'        If optJnsBayar(0).Value = True Then
'            .Parameters.Append .CreateParameter("StatusPiutang", adChar, adParamInput, 2, "TM")
'        ElseIf optJnsBayar(1).Value = True Then
'            .Parameters.Append .CreateParameter("StatusPiutang", adChar, adParamInput, 2, "OA")
'        ElseIf optJnsBayar(2).Value = True Then
'            .Parameters.Append .CreateParameter("StatusPiutang", adChar, adParamInput, 2, "SA")
'        End If
'        If optJnsBayar(2).Enabled = True Then
'            If optJnsBayar(2).Value = True Then
'                .Parameters.Append .CreateParameter("StatusBayarSemua", adChar, adParamInput, 1, "Y")
'            Else
'                .Parameters.Append .CreateParameter("StatusBayarSemua", adChar, adParamInput, 1, "T")
'            End If
'        Else
'            .Parameters.Append .CreateParameter("StatusBayarSemua", adChar, adParamInput, 1, "Y")
'        End If
       '------------------ End yg ditutup by Onede
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_NoStrukPelayananPasien"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Struk Billing Pasien", vbCritical, "Validasi"
            sp_AddStruk = False
        Else
            If Not IsNull(.Parameters("OutputNoStruk").Value) Then mstrNoStruk = .Parameters("OutputNoStruk").Value
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Function
errload:
    msubPesanError ("-Add_NoStrukPelayananPasien")
End Function

Private Sub chkDetailPasien_Click()
    If chkDetailPasien.Value = 1 Then
        strPasien = "View"
        strRegistrasi = "PasienLama"
        Load frmPasienBaru
        frmPasienBaru.Show
    Else
        Unload frmPasienBaru
        Unload frmDetailPasien
    End If
End Sub

Private Sub chkDetailPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglPendaftaran.SetFocus
End Sub

Private Sub chkDiriSendiri_Click()
    On Error GoTo errload
    If chkDiriSendiri.Value = vbChecked Then
        strSQL = "SELECT NamaLengkap, Alamat, Telepon,Propinsi,Kota,Kecamatan,Kelurahan,RTRW,Kodepos FROM Pasien WHERE NocM='" & txtNoCM.Text & "'"
        Call msubRecFO(rs, strSQL)
        If rs.RecordCount <> 0 Then
            txtNamaRI.Text = rs("NamaLengkap").Value
            txtAlamatRI.Text = IIf(IsNull(rs("Alamat").Value), "-", rs("Alamat").Value)
            txtTlpRI.Text = IIf(IsNull(rs("Telepon")), "-", rs("Telepon").Value)
            dcPropinsiPJ.Text = IIf(IsNull(rs("Propinsi")), "-", rs("Propinsi"))
            dcKotaPJ.Text = IIf(IsNull(rs("Kota")), "-", rs("Kota"))
            dcKecamatanPJ.Text = IIf(IsNull(rs("Kecamatan")), "-", rs("Kecamatan"))
            dcKelurahanPJ.Text = IIf(IsNull(rs("Kelurahan")), "-", rs("Kelurahan"))
            
            'load Pekerjaan Pasien
            strSQL = "SELECT Pekerjaan FROM detailPasien WHERE NocM='" & txtNoCM.Text & "'"
            Call msubRecFO(rs, strSQL)
            dcPekerjaanPJ.Text = IIf(rs.RecordCount = 0, "-", rs("Pekerjaan"))
            
        Else
        txtNamaRI.Text = ""
        txtAlamatRI.Text = ""
        txtTlpRI.Text = ""
        dcPropinsiPJ.Text = ""
        dcKotaPJ.Text = ""
        dcKecamatanPJ.Text = ""
        dcKelurahanPJ.Text = ""
        End If
    Else
        txtNamaRI.Text = ""
        txtAlamatRI.Text = ""
        txtTlpRI.Text = ""
        dcPropinsiPJ.Text = ""
        dcKotaPJ.Text = ""
        dcKecamatanPJ.Text = ""
        dcKelurahanPJ.Text = ""
    End If
    dcHubungan.BoundText = ""
    Exit Sub
errload:
    msubPesanError
End Sub
Private Sub chkDiriSendiri_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If chkDiriSendiri.Value = vbChecked Then
            cmdSimpan.SetFocus
        Else
            txtNamaRI.SetFocus
        End If
    End If
End Sub

Private Sub cmdAsuransiP_Click()
    mstrNoPen = ""
    mstrNoCM = txtNoCM.Text
    mstrKdJenisPasien = dcKelompokPasien.BoundText
'    mstrKdPenjaminPasien = "2222222222"
    With frmUbahJenisPasien
        .Show
        .txtNamaFormPengirim.Text = "tampung"
        .txtNoCM.Text = mstrNoCM
        .txtNamaPasien.Text = txtNamaPasien.Text
        .txtJK.Text = txtJK.Text
        .txtThn.Text = txtThn.Text
        .txtBln.Text = txtBln.Text
        .txtHr.Text = txtHr.Text
        .txtTglPendaftaran.Text = dtpTglPendaftaran.Value
        .lblNoPendaftaran.Visible = False
        .txtNoPendaftaran.Visible = False
        .dcJenisPasien.BoundText = mstrKdJenisPasien
'        .dcPenjamin.BoundText = mstrKdPenjaminPasien
        .dcAsalRujukan.BoundText = dcRujukanRI.BoundText
    End With
End Sub
Private Sub cmdRujukan_Click()
If dcRujukanRI.BoundText = "01" Then cmdTutup.SetFocus: Exit Sub   ' datang sendiri"
    With frmRujukan
        .Show
        .txtNoCM.Text = txtNoCM
        .txtNamaPasien.Text = txtNamaPasien.Text
        .txtJK.Text = txtJK.Text
        .txtThn.Text = txtThn.Text
        .txtBln.Text = txtBln.Text
        .txtHr.Text = txtHr.Text
        .txtNoPendaftaran.Text = txtNoPendaftaran.Text
        .dcRujukanAsal.Text = dcRujukanRI.Text
        mstrKdInstalasiPerujuk = dcRujukanRI.BoundText
    End With
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo errload
cmdRujukan.Enabled = False
    If funcCekValidasi = False Then Exit Sub
    
    Call msubRecFO(dbRst, "SELECT IdPenjamin FROM PenjaminKelompokPasien WHERE KdKelompokPasien = '" & strKdKelompokPasien & "'")
    If dbRst.EOF = True Then
        MsgBox "Lengkapi dulu data Penjamin Kelompok Pasien " & vbNewLine & "" & dcKelompokPasien.Text & "", vbExclamation, "Validasi"
        dcKelompokPasien.SetFocus
        Exit Sub
    End If
    If dbRst(0).Value <> "2222222222" And typAsuransi.blnSuksesAsuransi = False Then
        cmdSimpan.SetFocus
        'simpan data penjamin
        Call cmdAsuransiP_Click 'mstrKdPenjaminPasien selalu 2222222222
        Exit Sub
    End If
        
    If dcInstalasi.BoundText = "03" Then
        'validasi data registrasi ri
        If Periksa("datacombo", dcCaraMasukRI, "Data cara masuk kosong!") = False Then Exit Sub
        If Periksa("datacombo", dcKelasKamarRI, "Data kelas kamar kosong!") = False Then Exit Sub
        If Periksa("datacombo", dcNoKamarRI, "Data nomor kamar kosong!") = False Then Exit Sub
        If Periksa("datacombo", dcNoBedRI, "Data nomor bed kosong!") = False Then Exit Sub
        If Periksa("text", txtNamaRI, "Data nama penanggung jawab kosong!") = False Then Exit Sub
        If Periksa("text", txtAlamatRI, "Data alamat penanggung jawab kosong!") = False Then Exit Sub
        If Len(Trim(dcHubungan.Text)) > 0 Then
            If Periksa("datacombo", dcHubungan, "Hubungan pasien kosong!") = False Then Exit Sub
        End If
    
        strSQL = "SELECT StatusBed FROM StatusBed WHERE (KdKamar = '" & dcNoKamarRI.BoundText & "') AND (NoBed = '" & dcNoBedRI.BoundText & "')"
        Call msubRecFO(rs, strSQL)
        If UCase(rs(0).Value) = "I" Then
            MsgBox "No. Bed sudah terpakai", vbExclamation, "Validasi"
            strSQL = "SELECT distinct dbo.StatusBed.NoBed, dbo.StatusBed.NoBed AS Alias" & _
                " FROM dbo.NoKamar INNER JOIN dbo.StatusBed ON dbo.NoKamar.KdKamar = dbo.StatusBed.KdKamar" & _
                " WHERE (KdRuangan = '" & dcRuangan.BoundText & "') AND (KdKelas = '" & dcKelasKamarRI.BoundText & "') AND (dbo.NoKamar.KdKamar = '" & dcNoKamarRI.BoundText & "') AND (dbo.StatusBed.StatusBed = 'K')"
            Call msubDcSource(dcNoBedRI, rs, strSQL)
            Exit Sub
        End If
    End If
    
    cmdSimpan.Enabled = False
    
    'simpan data registrasi
    Call sp_RegistrasiAll(dbcmd)
    
    If txtNoPendaftaran = "" Then
        MsgBox "No Pendaftaran kosong !!", vbExclamation, "Validasi"
        Exit Sub
    End If
    
    If dcInstalasi.BoundText = "03" Then
        'simpan registrasi pasien RI
        Call sp_RegistrasiPasienRI(dbcmd)
        'simpan pasien masuk kamar
        Call sp_PasienMasukKamar(dbcmd)
    End If
    
    Call msubRecFO(dbRst, "SELECT IdPenjamin FROM PenjaminKelompokPasien WHERE KdKelompokPasien = '" & strKdKelompokPasien & "'")
'    If dbRst(0).Value <> "2222222222" Then
'        Call sp_AsuransiPasien(dbcmd)
'    End If
    
'    If dcInstalasi.BoundText <> "01" And dcInstalasi.BoundText <> "09" And dcInstalasi.BoundText <> "10" And dcInstalasi.BoundText <> "16" Then
'        If sp_PelayananOtomatis() = False Then Exit Sub
'    End If
    
    cmdSimpan.Enabled = True
    
    Call subEnableButtonReg(True)
cmdRujukan.Enabled = True
Exit Sub
errload:
    Call msubPesanError
    cmdSimpan.Enabled = True
End Sub

Private Sub cmdTutup_Click()
    If cmdSimpan.Enabled = True And txtNamaPasien.Text <> "" Then
        If MsgBox("Simpan Data Registrasi Pasien ?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub dcCaraMasukRI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcKelasKamarRI.SetFocus
End Sub

Private Sub dcHubungan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then dcPekerjaanPJ.SetFocus
End Sub

Private Sub dcHubungan_KeyPress(KeyAscii As Integer)
On Error GoTo errload

    If KeyAscii = 13 Then
        If Len(Trim(dcHubungan.Text)) = 0 Then cmdSimpan.SetFocus
        If dcHubungan.MatchedWithList = True Then dcPekerjaanPJ.SetFocus
        strSQL = "SELECT Hubungan, NamaHubungan FROM HubunganKeluarga WHERE (NamaHubungan LIKE '%" & dcHubungan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcHubungan.BoundText = rs(0).Value
        dcHubungan.Text = rs(1).Value
       
    End If

Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub dcInstalasi_Change()
On Error GoTo errload

    dcJenisKelas.BoundText = ""
    If dcInstalasi.BoundText = "02" Then 'RJ
        dcJenisKelas.BoundText = "01" 'UMUM
    Else
        dcJenisKelas.BoundText = ""
    End If
    dcSubInstalasi.BoundText = ""
    'Call subTampilRegistrasiRI
    Call subDcSource
Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub dcInstalasi_GotFocus()
On Error GoTo errload
Dim tempKode As String

    tempKode = dcInstalasi.BoundText

'    strSQL = "SELECT DISTINCT KdInstalasi,NamaInstalasi FROM V_KelasPelayanan WHERE KdDetailJenisJasaPelayanan='" & dcJenisKelas.BoundText & "'"
    strSQL = "SELECT DISTINCT KdInstalasi,NamaInstalasi FROM V_KelasPelayanan where KdInstalasi = '02'"
    Call msubDcSource(dcInstalasi, rs, strSQL)
    
    dcInstalasi.BoundText = tempKode
    
Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub dcInstalasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcJenisKelas.SetFocus
End Sub

Private Sub dcKecamatanPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcKelurahanPJ.SetFocus
End Sub

Private Sub dcKelurahanPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then meRTRWPJ.SetFocus
End Sub

Private Sub dcKotaPJ_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then dcKecamatanPJ.SetFocus
End Sub

Private Sub dcPekerjaanPJ_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtAlamatRI.SetFocus
End Sub

Private Sub dcPropinsiPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcKotaPJ.SetFocus
End Sub

Private Sub dcSubInstalasi_GotFocus()
On Error GoTo errload
Dim tempKode As String
    
    tempKode = dcSubInstalasi.BoundText
    Call msubDcSource(dcSubInstalasi, rs, "SELECT KdSubInstalasi, NamaSubInstalasi FROM V_SubInstalasiRuangan WHERE (KdRuangan = '" & dcRuangan.BoundText & "') ORDER BY NamaSubInstalasi")
    dcSubInstalasi.BoundText = tempKode

Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub dcJenisKelas_Change()
    dcKelas.Text = ""
End Sub

Private Sub dcJenisKelas_GotFocus()
On Error GoTo errload
Dim tempKode As String
    
    tempKode = dcJenisKelas.BoundText
    strSQL = "SELECT distinct KdDetailJenisJasaPelayanan,DetailJenisJasaPelayanan FROM V_KelasPelayanan"
    Call msubDcSource(dcJenisKelas, rs, strSQL)
    dcJenisKelas.BoundText = tempKode
    
Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub dcJenisKelas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcKelas.SetFocus
End Sub

Private Sub dcKelas_Change()
    dcRuangan.Text = ""
End Sub

Private Sub dcKelas_GotFocus()
On Error GoTo errload
Dim tempKode As String
    
    tempKode = dcKelas.BoundText
    
    strSQL = "SELECT distinct KdKelas, Kelas FROM V_KelasPelayanan WHERE KdInstalasi = '" & dcInstalasi.BoundText & "' and KdDetailJenisJasaPelayanan ='" & dcJenisKelas.BoundText & "' AND KdKelas<>04"
    Call msubDcSource(dcKelas, rs, strSQL)
    
    dcKelas.BoundText = tempKode

Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub dcKelas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcRuangan.SetFocus
End Sub

Private Sub dcKelasKamarRI_GotFocus()
On Error GoTo errload
Dim tempKdKelas As String
Dim tempKdRuangan As String

    tempKdKelas = dcKelasKamarRI.BoundText
    
    'cek kelas intensif
    strSQL = "SELECT Distinct KdKelas, KdRuangan From V_KamarRegRawatInap WHERE KdRuangan = '" & dcRuangan.BoundText & "' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF Then Exit Sub
    tempKdRuangan = rs("KdRuangan").Value
    
    If rs(0).Value = "04" Then
        strSQL = "SELECT DISTINCT KdKelas, Kelas " & _
            " FROM V_KamarRegRawatInap " & _
            " WHERE (KdRuangan = '" & dcRuangan.BoundText & "') AND (KdKelas IN ('" & dcKelas.BoundText & "','04'))"
    Else
        strSQL = "SELECT DISTINCT KdKelas, Kelas " & _
            " FROM V_KamarRegRawatInap " & _
            " WHERE (KdRuangan = '" & dcRuangan.BoundText & "') AND KdKelas in ('" & dcKelas.BoundText & "','04')"
    End If
    
    Call msubDcSource(dcKelasKamarRI, rs, strSQL)
    dcKelasKamarRI.BoundText = tempKdKelas

Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub dcKelasKamarRI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcNoKamarRI.SetFocus
End Sub

Private Sub dcKelompokPasien_Change()
   strKdKelompokPasien = dcKelompokPasien.BoundText
   typAsuransi.blnSuksesAsuransi = False
End Sub

Private Sub dcKelompokPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcKelompokPasien.Text = "" Then
            dcKelompokPasien.SetFocus
            Exit Sub
        End If
        strKdKelompokPasien = dcKelompokPasien.BoundText
        Call msubRecFO(dbRst, "SELECT IdPenjamin FROM PenjaminKelompokPasien WHERE KdKelompokPasien = '" & strKdKelompokPasien & "'")
        If dbRst.EOF = True Then
            MsgBox "Lengkapi dulu data Penjamin Kelompok Pasien " & vbNewLine & "" & dcKelompokPasien.Text & "", vbExclamation, "Validasi"
            dcKelompokPasien.SetFocus
            Exit Sub
        End If
        If dbRst(0).Value <> "2222222222" And typAsuransi.blnSuksesAsuransi = False Then
            Call cmdAsuransiP_Click
        Else
            'Call subTampilRegistrasiRI
            If dcInstalasi.BoundText = "03" Then
                dcCaraMasukRI.SetFocus
            Else
                cmdSimpan.SetFocus
            End If
        End If
    End If
End Sub

Private Sub dcNoBedRI_GotFocus()
On Error GoTo errload
Dim tempKode As String

    tempKode = dcNoBedRI.BoundText
    strSQL = "SELECT distinct dbo.StatusBed.NoBed, dbo.StatusBed.NoBed AS Alias" & _
        " FROM dbo.NoKamar INNER JOIN dbo.StatusBed ON dbo.NoKamar.KdKamar = dbo.StatusBed.KdKamar" & _
        " WHERE (KdRuangan = '" & dcRuangan.BoundText & "') AND (KdKelas = '" & dcKelasKamarRI.BoundText & "') AND (dbo.NoKamar.KdKamar = '" & dcNoKamarRI.BoundText & "') AND (dbo.StatusBed.StatusBed = 'K')"
    Call msubDcSource(dcNoBedRI, rs, strSQL)
    dcNoBedRI.BoundText = tempKode
    
Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub dcNoBedRI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then optTidak.SetFocus
End Sub

Private Sub dcNoKamarRI_GotFocus()
On Error GoTo errload
Dim tempKode As String
    
    tempKode = dcNoKamarRI.BoundText
'    strSQL = "SELECT dbo.NoKamar.NoKamar, dbo.NoKamar.NamaKamar AS Alias" & _
'        " FROM dbo.NoKamar INNER JOIN dbo.StatusBed ON dbo.NoKamar.NoKamar = dbo.StatusBed.NoKamar" & _
'        " WHERE (KdRuangan = '" & dcRuangan.BoundText & "') AND (KdKelas = '" & dcKelasKamarRI.BoundText & "') AND (dbo.StatusBed.StatusBed = 'K')"
    strSQL = "SELECT distinct dbo.NoKamar.KdKamar,dbo.NoKamar.NamaKamar AS Alias" & _
        " FROM dbo.NoKamar INNER JOIN dbo.StatusBed ON dbo.NoKamar.KdKamar = dbo.StatusBed.KdKamar" & _
        " WHERE (KdRuangan = '" & dcRuangan.BoundText & "') AND (KdKelas = '" & dcKelasKamarRI.BoundText & "') AND (dbo.StatusBed.StatusBed = 'K')"
    
    Call msubDcSource(dcNoKamarRI, rs, strSQL)
    dcNoKamarRI.BoundText = tempKode
    
Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub dcNoKamarRI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcNoBedRI.SetFocus
End Sub

Private Sub dcRuangan_Change()
On Error GoTo errload
    
    If dcInstalasi.BoundText = "03" Then
        Call msubDcSource(dcKelasKamarRI, rsB, "SELECT DISTINCT KdKelas, Kelas FROM V_KamarRegRawatInap WHERE (KdRuangan = '" & dcRuangan.BoundText & "')")
        'If rsb.EOF = False Then dcKelasKamarRI.BoundText = rsb(0).Value
'    Else
        dcKelasKamarRI.BoundText = ""
    End If

Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub dcRuangan_GotFocus()
On Error GoTo errload
Dim tempKode As String

    tempKode = dcRuangan.BoundText
    If dcInstalasi.BoundText = "03" Then
        strSQL = "SELECT distinct KdRuangan, NamaRuangan FROM V_KelasPelayanan WHERE (KdInstalasi = '" & dcInstalasi.BoundText & "') AND (KdDetailJenisJasaPelayanan  = '" & dcJenisKelas.BoundText & "') AND (KdKelas = '" & dcKelas.BoundText & "') OR KdKelas='04' ORDER BY NamaRuangan"
    Else
        strSQL = "SELECT distinct KdRuangan, NamaRuangan FROM V_KelasPelayanan WHERE (KdInstalasi = '" & dcInstalasi.BoundText & "') AND (KdDetailJenisJasaPelayanan  = '" & dcJenisKelas.BoundText & "') AND (KdKelas = '" & dcKelas.BoundText & "') ORDER BY NamaRuangan"
    End If
    Call msubDcSource(dcRuangan, rs, strSQL)
    dcRuangan.BoundText = tempKode

Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub dcRuangan_KeyPress(KeyAscii As Integer)
On Error GoTo errload
    If KeyAscii = 13 Then
        strSQL = "SELECT KdRuangan, NamaRuangan FROM V_KelasPelayanan WHERE (KdInstalasi IN ('" & dcInstalasi.BoundText & "','08')) AND (KdDetailJenisJasaPelayanan  = '" & dcJenisKelas.BoundText & "') AND (KdKelas IN ('" & dcKelas.BoundText & "','04')) AND (NamaRuangan LIKE '%" & dcRuangan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcRuangan.BoundText = rs(0).Value
        dcRuangan.Text = rs(1).Value
        
        strSQL = "SELECT KdSubInstalasi, NamaSubInstalasi FROM  V_SubInstalasiRuangan WHERE (KdRuangan = '" & dcRuangan.BoundText & "')"
        Call msubDcSource(dcSubInstalasi, rs, strSQL)
        If rs.EOF = False Then dcSubInstalasi.BoundText = rs(0).Value
        
        dcSubInstalasi.SetFocus
        Exit Sub
    End If
Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub dcRujukanRI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcKelompokPasien.SetFocus
End Sub

Private Sub dcSubInstalasi_KeyPress(KeyAscii As Integer)
On Error GoTo errload
    If KeyAscii = 13 Then
        strSQL = "SELECT KdSubInstalasi, NamaSubInstalasi FROM V_SubInstalasiRuangan WHERE (KdRuangan = '" & dcRuangan.BoundText & "') AND (NamaSubInstalasi LIKE '%" & dcSubInstalasi.Text & "%')"
        Call msubRecFO(dbRst, strSQL)
        If dbRst.EOF = True Then Exit Sub
        dcSubInstalasi.BoundText = dbRst(0).Value
        dcRujukanRI.SetFocus
    End If
    
Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub dtpTglPendaftaran_Change()
    dtpTglPendaftaran.MaxDate = Now
End Sub

Private Sub dtpTglPendaftaran_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcInstalasi.SetFocus
End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo hell
'Dim strCtrlKey As String
'
'    'deklarasi tombol control ditekan
'    strCtrlKey = (Shift + vbCtrlMask)
'
'Select Case KeyCode
'Case vbKeyF1
'mstrNoPen = frmRegistrasiAll.txtNoPendaftaran.Text
'mstrKdInstalasi = frmRegistrasiAll.dcInstalasi.BoundText
'If dcInstalasi.BoundText = "02" Or dcInstalasi.BoundText = "06" Or dcInstalasi.BoundText = "11" Then
'If cmdSimpan.Enabled = True Then Exit Sub
'
''exec ke pasien sudah bayar
''aktifkan jika tidak perlu otomatis bayar
'            strSQL = "SELECT SUM(Tarif) AS Tarif, SUM(JmlHutangPenjamin) AS JmlHutangPenjamin, SUM(JmlTanggunganRS) AS JmlTanggunganRS, SUM(JmlPembebasan) " & _
'                " AS JmlPembebasan From DetailBiayaPelayanan WHERE     (NoPendaftaran = '" & mstrNoPen & "')"
'            Call msubRecFO(rs, strSQL)
'            If dcInstalasi.BoundText = "06" Or dcRuangan.BoundText = "222" Then
'                curTarif = 0
'                curTP = 0
'                curTRS = 0
'                curPemb = 0
'                mcurAll_HrsDibyr = 0
'                mcurBayar = 0
'                mcurAll_TP = 0
'                mcurAll_TRS = 0
'                mcurAll_Pemb = 0
'            Else
'                curTarif = rs.Fields("Tarif")
'                curTP = rs.Fields("JmlHutangPenjamin")
'                curTRS = rs.Fields("JmlTanggunganRS")
'                curPemb = rs.Fields("JmlPembebasan")
'                mcurAll_HrsDibyr = curTarif - (curTP + curTRS + curPemb)
'                mcurBayar = curTarif
'                mcurAll_TP = curTP
'                mcurAll_TRS = curTRS
'                mcurAll_Pemb = curPemb
'            End If
'    If txtNoBKM.Text <> "print" Then
'            If sp_AddStrukBuktiKasMasuk() = False Then Exit Sub
'            mstrNoBKM = txtNoBKM.Text
'            txtNoBKM.Text = "print"
'            If sp_AddStruk(dbcmd, 1) = False Then Exit Sub
'            fStatusPiutang = "TM"
'            fStatusBayarSemua = "Y"
'            Call f_AddStrukPelayananPasienDetail(mstrNoBKM, mstrNoStruk, mstrNoPen, mstrNoCM, CCur(mcurAll_HrsDibyr), 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, mcurBayar, mcurAll_HrsDibyr, mcurAll_TP, mcurAll_TRS, mcurAll_Pemb, mcurAll_HrsDibyr, 0, 0)
'    End If
''end code bayar otomatis
'End If
'frm_cetak_label_viewer.Show
'
''            Call subCetakLabelRegistrasi
'        Case vbKeyB
'            If strCtrlKey = 4 Then
'                Unload Me
'                strPasien = "Baru"
'                frmPasienBaru.Show
'            End If
'        Case vbKeyF2
'            Unload Me
'            frmCariPasien.Show
'        Case vbKeyL
'            If strCtrlKey = 4 Then
'                Unload Me
'                strPasien = "Lama"
'                frmRegistrasiAll.Show
'            End If
'        Case vbKeyR
'            If txtNoPendaftaran.Text = "" Then Exit Sub
'            If dcInstalasi.BoundText <> "03" Then
'               If dcInstalasi.BoundText <> "02" Then Exit Sub
'            End If
'            frmCetakLembarMasukDanKeluarV2.Show
'        Case vbKeyZ
'            If txtNoPendaftaran.Text = "" Then Exit Sub
'            mstrNoPen = txtNoPendaftaran.Text
'            'If dcInstalasi.BoundText <> "03" Then Exit Sub
'            frmCetakSuratKeterangan.Show
'        Case vbKeyF9
'            If cmdSimpan.Enabled = True Then Exit Sub
'            If mstrNoSJP = "" Then
'                MsgBox "No SJP kosong", vbExclamation, "Validasi"
'                Exit Sub
'            End If
'            frmViewerSJP.Show
'        Case vbKeyM
'            If cmdSimpan.Enabled = True Then Exit Sub
'            mstrNoCM = Trim(txtNoCM)
'            frmCetakCatatanMedis0.Show
'    End Select
'Exit Sub
'hell:
'
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpTglPendaftaran.Value = Now
    strRegistrasi = "RJ"
    If mblnCariPasien = True Then frmCariPasien.Enabled = False
    Call subDcSource
   ' Call subTampilRegistrasiRI
'    Call subClearData
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnCariPasien = True Then frmCariPasien.Enabled = True
End Sub

Private Sub hgPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub meRTRWPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtKodePos.SetFocus
    If KeyCode = 39 Then KeyAscii = 0
End Sub
    
Private Sub meRTRWPJ_KeyPress(KeyAscii As Integer)
If KeyCode = 13 Then txtKodePos.SetFocus
End Sub

Private Sub optTidak_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkDiriSendiri.SetFocus
End Sub

Private Sub optYa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkDiriSendiri.SetFocus
End Sub

Private Sub txtAlamatRI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcPropinsiPJ.SetFocus
End Sub


Private Sub txtKodePos_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then txtTlpRI.SetFocus
End Sub

Private Sub txtKodePos_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then txtTlpRI.SetFocus
End Sub

Private Sub txtNamaRI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcHubungan.SetFocus
End Sub

Private Sub txtNoCM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNoBKM = ""
        Call CariData
        If chkDetailPasien.Enabled = True Then chkDetailPasien.SetFocus
    End If
End Sub

Private Sub txtNoPendaftaran_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'untuk enable/disable button reg
Private Sub subEnableButtonReg(blnStatus As Boolean)
    cmdRujukan.Enabled = blnStatus
    cmdAsuransiP.Enabled = blnStatus
    cmdSimpan.Enabled = Not blnStatus
    dtpTglPendaftaran.Enabled = Not blnStatus
    dcInstalasi.Enabled = Not blnStatus
    dcRuangan.Enabled = Not blnStatus
    dcSubInstalasi.Enabled = Not blnStatus
    dcKelompokPasien.Enabled = Not blnStatus
    dcKelas.Enabled = Not blnStatus
    dcJenisKelas.Enabled = Not blnStatus
End Sub

'Store procedure untuk mengisi registrasi pasien RI
Private Sub sp_RegistrasiPasienRI(ByVal adoCommand As ADODB.Command)
 Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adChar, adParamInput, 6, txtNoCM.Text)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, dcSubInstalasi.BoundText)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, dcKelas.BoundText) ' dcKelasKamarRI.BoundText)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(dtpTglPendaftaran.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 2, Null)
        .Parameters.Append .CreateParameter("KdCaraMasuk", adChar, adParamInput, 2, dcCaraMasukRI.BoundText)
        .Parameters.Append .CreateParameter("KdRujukanAsal", adChar, adParamInput, 2, dcRujukanRI.BoundText)
        .Parameters.Append .CreateParameter("NamaPJ", adVarChar, adParamInput, 20, txtNamaRI.Text)
        .Parameters.Append .CreateParameter("PekerjaanPJ", adVarChar, adParamInput, 30, dcPekerjaanPJ.Text)
        .Parameters.Append .CreateParameter("Hubungan", adChar, adParamInput, 2, IIf(dcHubungan.BoundText = "", Null, dcHubungan.BoundText))
        .Parameters.Append .CreateParameter("AlamatPJ", adVarChar, adParamInput, 50, IIf(txtAlamatRI.Text = "", Null, txtAlamatRI.Text))
        .Parameters.Append .CreateParameter("PropinsiPJ", adVarChar, adParamInput, 25, IIf(dcPropinsiPJ.Text = "", Null, dcPropinsiPJ.Text))
        .Parameters.Append .CreateParameter("KotaPJ", adVarChar, adParamInput, 25, IIf(dcKotaPJ.Text = "", Null, dcKotaPJ.Text))
        .Parameters.Append .CreateParameter("KecamatanPJ", adVarChar, adParamInput, 25, IIf(dcKecamatanPJ.Text = "", Null, dcKecamatanPJ.Text))
        .Parameters.Append .CreateParameter("KelurahanPJ", adVarChar, adParamInput, 25, IIf(dcKelurahanPJ.Text = "", Null, dcKelurahanPJ.Text))
        .Parameters.Append .CreateParameter("RTRWPJ", adVarChar, adParamInput, 25, IIf(meRTRWPJ.Text = "", Null, meRTRWPJ.Text))
        .Parameters.Append .CreateParameter("KodePosPJ", adVarChar, adParamInput, 25, IIf(meRTRWPJ.Text = "", Null, txtKodePos.Text))
        .Parameters.Append .CreateParameter("TeleponPJ", adVarChar, adParamInput, 20, IIf(Len(Trim(txtTlpRI.Text)) = 0, Null, Trim(txtTlpRI.Text)))
        
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_RegistrasiPasienRI"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan registrasi RI", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

'Store procedure untuk mengisi pasien masuk RI
Private Sub sp_PasienMasukKamar(ByVal adoCommand As ADODB.Command)
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adChar, adParamInput, 6, txtNoCM.Text)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, dcSubInstalasi.BoundText)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, dcRuangan.BoundText)
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 2, Null)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, dcKelasKamarRI.BoundText)
        .Parameters.Append .CreateParameter("KdKamar", adChar, adParamInput, 4, dcNoKamarRI.BoundText)
        .Parameters.Append .CreateParameter("NoBed", adChar, adParamInput, 2, dcNoBedRI.BoundText)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(dtpTglPendaftaran.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, dcKelas.BoundText)
        
        .Parameters.Append .CreateParameter("OutputNoPakai", adChar, adParamOutput, 10, Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("KdCaraMasuk", adChar, adParamInput, 2, dcCaraMasukRI.BoundText)
        .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, Null)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 2, IIf(optTidak.Value = True, "MA", "RG"))
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PasienMasukKamar"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan pasien masuk kamar", vbCritical, "Validasi"
        Else
            txtNoPakai.Text = .Parameters("OutputNoPakai").Value
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

'Store procedure untuk mengisi registrasi pasien
Private Sub sp_RegistrasiAll(ByVal adoCommand As ADODB.Command)
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoCM", adChar, adParamInput, 6, txtNoCM.Text)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, dcSubInstalasi.BoundText)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, dcRuangan.BoundText)
        .Parameters.Append .CreateParameter("TglPendaftaran", adDate, adParamInput, , Format(dtpTglPendaftaran.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(dtpTglPendaftaran.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, dcKelas.BoundText)
        .Parameters.Append .CreateParameter("KdKelompokPasien", adChar, adParamInput, 2, strKdKelompokPasien)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, noidpegawai)
        .Parameters.Append .CreateParameter("OutputNoPendaftaran", adChar, adParamOutput, 10, Null)
        .Parameters.Append .CreateParameter("OutputNoAntrian", adChar, adParamOutput, 3, Null)
        .Parameters.Append .CreateParameter("KdDetailJenisJasaPelayanan", adChar, adParamInput, 2, dcJenisKelas.BoundText)
        .Parameters.Append .CreateParameter("KdRujukanAsal", adChar, adParamInput, 2, dcRujukanRI.BoundText)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_RegistrasiPasienMRS"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Pendaftaran Pasien ke Instalasi Rawat Jalan", vbCritical, "Validasi"
        Else
            If Not IsNull(.Parameters("OutputNoPendaftaran").Value) Then mstrNoPen = .Parameters("OutputNoPendaftaran").Value
            If Not IsNull(.Parameters("OutputNoAntrian").Value) Then strNoAntrian = .Parameters("OutputNoAntrian").Value
            txtNoPendaftaran.Text = mstrNoPen
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

'Store procedure untuk mengisi pelayanan otomatis
Private Function sp_PelayananOtomatis() As Boolean
On Error GoTo errload
    sp_PelayananOtomatis = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, dcSubInstalasi.BoundText)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, dcRuangan.BoundText)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(dtpTglPendaftaran.Value, "yyyy/MM/dd HH:mm:ss"))
        If dcInstalasi.BoundText <> "03" And dcInstalasi.BoundText <> "08" Then
            .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, Null)
        Else
            .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, dcKelasKamarRI.BoundText)
        End If
        .Parameters.Append .CreateParameter("KdKelasPel", adChar, adParamInput, 2, dcKelas.BoundText)
        If dcInstalasi.BoundText <> "03" And dcInstalasi.BoundText <> "08" Then
            .Parameters.Append .CreateParameter("NoLab_Rad", adChar, adParamInput, 10, Null)
        Else
            .Parameters.Append .CreateParameter("NoLab_Rad", adChar, adParamInput, 10, txtNoPakai.Text)
        End If
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, strIDPegawaiAktif)
        If dcInstalasi.BoundText <> "03" And dcInstalasi.BoundText <> "08" Then
            .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 2, "AL")
        Else
            .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 2, IIf(optTidak.Value = True, "MA", "RG"))
        End If
                
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_BiayaPelayananOtomatis"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            sp_PelayananOtomatis = False
            MsgBox "Ada kesalahan penyimpanan data", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
Exit Function
errload:
    Call msubPesanError("sp_PelayananOtomatis")
End Function

'Store procedure untuk mengisi asuransi pasien
'Private Sub sp_AsuransiPasien(ByVal adoCommand As ADODB.Command)
'    Set dbcmd = New ADODB.Command
'    With dbcmd
'        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
'        .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, typAsuransi.strIdPenjamin)
'        .Parameters.Append .CreateParameter("IdAsuransi", adVarChar, adParamInput, 25, typAsuransi.strIdAsuransi)
'        .Parameters.Append .CreateParameter("NoCM", adChar, adParamInput, 6, mstrNoCM)
'        .Parameters.Append .CreateParameter("NamaPeserta", adVarChar, adParamInput, 50, typAsuransi.strNamaPeserta)
'        If typAsuransi.strIdPeserta <> "" Then
'            .Parameters.Append .CreateParameter("IDPeserta", adVarChar, adParamInput, 16, typAsuransi.strIdPeserta)
'        Else
'            .Parameters.Append .CreateParameter("IDPeserta", adVarChar, adParamInput, 16, Null)
'        End If
'        .Parameters.Append .CreateParameter("KdGolongan", adChar, adParamInput, 2, typAsuransi.strKdGolongan)
'        .Parameters.Append .CreateParameter("TglLahir", adDate, adParamInput, , Format(typAsuransi.dTglLahir, "yyyy/MM/dd"))
'        If typAsuransi.strAlamat <> "" Then
'            .Parameters.Append .CreateParameter("Alamat", adVarChar, adParamInput, 100, typAsuransi.strAlamat)
'        Else
'            .Parameters.Append .CreateParameter("Alamat", adVarChar, adParamInput, 100, Null)
'        End If
'        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
'        .Parameters.Append .CreateParameter("Hubungan", adChar, adParamInput, 2, typAsuransi.strHubungan)
'        If typAsuransi.strNoSJP <> "" Then
'            .Parameters.Append .CreateParameter("NoSJP", adVarChar, adParamInput, 30, typAsuransi.strNoSJP)
'        Else
'            .Parameters.Append .CreateParameter("NoSJP", adVarChar, adParamInput, 30, Null)
'        End If
'        .Parameters.Append .CreateParameter("TglSJP", adDate, adParamInput, , Format(typAsuransi.dTglSJP, "yyyy/MM/dd HH:mm:ss"))
'        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
'        .Parameters.Append .CreateParameter("NoBP", adChar, adParamInput, 3, IIf(Len(Trim(typAsuransi.strNoBp)) = 0, Null, Trim(typAsuransi.strNoBp)))
'        .Parameters.Append .CreateParameter("UnitBagian", adVarChar, adParamInput, 50, IIf(Len(Trim(typAsuransi.strUnitBagian)) = 0, Null, Trim(typAsuransi.strUnitBagian)))
'        .Parameters.Append .CreateParameter("KunjunganKe", adInteger, adParamInput, , typAsuransi.intNoKunjungan)
'        .Parameters.Append .CreateParameter("IdPerusahaan", adChar, adParamInput, 10, IIf(Len(Trim(typAsuransi.strPerusahaanPenjamin)) = 0, Null, Trim(typAsuransi.strPerusahaanPenjamin)))
'
'        .ActiveConnection = dbConn
'        .CommandText = "dbo.AU_AsuransiPasien"
'        .CommandType = adCmdStoredProc
'        .Execute
'
'        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
'            MsgBox "Ada kesalahan dalam pemasukan Asuransi Pasien", vbCritical, "Validasi"
'            mstrNoSJP = ""
'        Else
'            mstrNoSJP = "" ' IIf(IsNull(.Parameters("OutputNoSJP")), "", .Parameters("OutputNoSJP"))
'        End If
'        Set dbcmd = Nothing
'    End With
'    Exit Sub
'End Sub

'untuk cek validasi
Private Function funcCekValidasi() As Boolean
    If txtNamaPasien.Text = "" Then
        MsgBox "No. CM Harus Diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        txtNoCM.SetFocus
        Exit Function
    End If
    If Periksa("datacombo", dcInstalasi, "Nama instalasi kosong") = False Then funcCekValidasi = False: Exit Function
    If Periksa("datacombo", dcJenisKelas, "Jenis kelas pelayanan kosong") = False Then funcCekValidasi = False: Exit Function
    If Periksa("datacombo", dcKelas, "Kelas pelayanan kosong") = False Then funcCekValidasi = False: Exit Function
    If Periksa("datacombo", dcRuangan, "Nama ruangan kosong") = False Then funcCekValidasi = False: Exit Function
    If Periksa("datacombo", dcSubInstalasi, "Nama sub instalasi kosong!") = False Then funcCekValidasi = False: Exit Function
    If Periksa("datacombo", dcRujukanRI, "Data rujukan kosong!") = False Then funcCekValidasi = False: Exit Function
    If Periksa("datacombo", dcKelompokPasien, "Jenis pasien kosong!") = False Then funcCekValidasi = False: Exit Function
    
    funcCekValidasi = True
End Function

'untuk membersihkan data pasien registrasi
Private Sub subClearData()
    txtNoPakai.Text = ""
    txtNoPendaftaran.Text = ""
    txtNamaPasien.Text = ""
    txtJK.Text = ""
    txtThn.Text = ""
    txtBln.Text = ""
    txtHr.Text = ""
    dcHubungan.BoundText = ""
    dtpTglPendaftaran.MaxDate = #9/9/2999#
    dtpTglPendaftaran.Value = Now
    dcInstalasi.Text = ""
    dcRuangan.Text = ""
    dcJenisKelas.Text = ""
    dcKelompokPasien.Text = ""
    dcKelas.Text = ""
End Sub

Private Sub subDcSource()
On Error GoTo errload
    Call msubDcSource(dcKelompokPasien, rs, "SELECT KdKelompokPasien,JenisPasien FROM KelompokPasien WHERE NOT (KdKelompokPasien = '05')") 'askes gakin di tutup by splakuk
    Call msubDcSource(dcRujukanRI, rs, "SELECT KdRujukanAsal, RujukanAsal FROM RujukanAsal")
    dcRujukanRI.BoundText = rs(0).Value
    Call msubDcSource(dcCaraMasukRI, rs, "SELECT KdCaraMasuk, CaraMasuk FROM CaraMasuk")
    Call msubDcSource(dcSubInstalasi, rs, "SELECT KdInstalasi, NamaInstalasi FROM  V_RegistrasiALLRJ WHERE (KdRuangan = '" & mstrKdRuangan & "')")
    Call msubDcSource(dcHubungan, rs, "SELECT Hubungan, NamaHubungan FROM HubunganKeluarga")
    
    strSQL = "SELECT DISTINCT NamaPropinsi, NamaPropinsi AS alias FROM V_Wilayah"
    Call msubDcSource(dcPropinsiPJ, rs, strSQL)
    
    strSQL = "SELECT DISTINCT NamaKotaKabupaten, NamaKotaKabupaten AS alias FROM V_Wilayah"
    Call msubDcSource(dcKotaPJ, rs, strSQL)
    
    strSQL = "SELECT DISTINCT NamaKecamatan, NamaKecamatan AS alias FROM V_Wilayah"
    Call msubDcSource(dcKecamatanPJ, rs, strSQL)
    
    strSQL = "SELECT DISTINCT NamaKelurahan, NamaKelurahan AS alias FROM V_Wilayah"
    Call msubDcSource(dcKelurahanPJ, rs, strSQL)
    
    strSQL = "SELECT DISTINCT Pekerjaan,Pekerjaan AS alias FROM Pekerjaan"
    Call msubDcSource(dcPekerjaanPJ, rs, strSQL)

Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub txtTlpRI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSimpan.SetFocus
End Sub

'Private Sub subTampilRegistrasiRI()
'    If dcInstalasi.BoundText = "03" Then
'        frmRegistrasiAll.Height = 9090
'        Call centerForm(Me, MDIUtama)
'        fraRawatGabung.Visible = True
''        Call Animate(frmRegistrasiAll, 8280, True)
'    Else
'        frmRegistrasiAll.Height = 5100 + stbInformasi.Height
'        Call centerForm(Me, MDIUtama)
'        fraRawatGabung.Visible = False
''        Call Animate(frmRegistrasiAll, 5385, False)
'    End If
'End Sub

'untuk mengganti nocm on change
Public Sub CariData()
On Error GoTo errload
    Call subClearData
    Call subEnableButtonReg(False)
    
    'cek pasien igd
    strSQL = "SELECT NoCM FROM V_DaftarPasienIGDAktif WHERE (NoCM = '" & txtNoCM.Text & "')"
    Call msubRecFO(rs, strSQL)
    If Not rs.EOF Then
        MsgBox "Pasien tersebut belum keluar dari IGD", vbInformation, "Informasi"
        mstrNoCM = ""
        chkDetailPasien.Enabled = False
        cmdSimpan.Enabled = False
        Exit Sub
    End If
    
    'cek pasien ri
    strSQL = "SELECT dbo.RegistrasiRI.NoCM, dbo.Ruangan.NamaRuangan FROM dbo.RegistrasiRI INNER JOIN dbo.Ruangan ON dbo.RegistrasiRI.KdRuangan = dbo.Ruangan.KdRuangan WHERE (NoCM = '" & txtNoCM.Text & "') AND StatusPulang = 'T'"
    Call msubRecFO(rs, strSQL)
    If Not rs.EOF Then
        MsgBox "Pasien tersebut belum keluar dari Rawat Inap," & vbNewLine & "Ruangan " & rs("NamaRuangan") & " ", vbInformation, "Informasi"
        mstrNoCM = ""
        chkDetailPasien.Enabled = False
        cmdSimpan.Enabled = False
        Exit Sub
    End If
    
    strSQL = "Select * from v_CariPasien WHERE [No. CM]='" & txtNoCM.Text & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        mstrNoCM = ""
        chkDetailPasien.Enabled = False
        cmdSimpan.Enabled = False
        Exit Sub
    End If
    
    mstrNoCM = txtNoCM.Text
    txtNamaPasien.Text = rs.Fields("Nama Lengkap").Value
    If rs.Fields("JK").Value = "P" Then
        txtJK.Text = "Perempuan"
    ElseIf rs.Fields("JK").Value = "L" Then
        txtJK.Text = "Laki-laki"
    End If
    txtThn.Text = rs.Fields("UmurTahun").Value
    txtBln.Text = rs.Fields("UmurBulan").Value
    txtHr.Text = rs.Fields("UmurHari").Value
    Set rs = Nothing
    chkDetailPasien.Enabled = True
Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub subCetakLabelRegistrasi()
On Error GoTo errload
    Printer.Print strNNamaRS
    Printer.Print strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
    Printer.Print strWebsite & ", " & strEmail
    
    If (mstrKdInstalasi = "02") Or (mstrKdInstalasi = "11") Or (mstrKdInstalasi = "06") Then
        strSQL = "SELECT * from V_CetakLabelRegistrasiPasienMRS WHERE (NoPendaftaran) =('" & mstrNoPen & "')"
    Else
        strSQL = "SELECT * from V_CetakLabelRegistrasiPasienMRS WHERE (NoPendaftaran) =('" & mstrNoPen & "')"
    End If
    Call msubRecFO(rs, strSQL)
    
    Printer.Print "No. Pendaftaran"
    Printer.Print "No. CM"
    Printer.Print "Nama Pasien"
    Printer.Print "Jenis Kelamin"
    Printer.Print "Kelompok Pasien"
    Printer.Print "Jenis Kelas"
    Printer.Print "Ruangan Tujuan"
    Printer.Print "Lokasi Ruangan"
    Printer.Print "No. Ruangan"

    Printer.Print "No. Antrian"
    Printer.Print "------------------------------"

    strSQL = "SELECT MessageToDay FROM MasterDataPendukung"
    Call msubRecFO(rs, strSQL)
    Printer.Print IIf(IsNull(rs(0)), "", rs(0))
    Printer.Print "------------------------------"
    Printer.Print "User :"


    Printer.EndDoc
Exit Sub
errload:
    Call msubPesanError
End Sub
