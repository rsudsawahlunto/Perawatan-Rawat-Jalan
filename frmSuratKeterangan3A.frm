VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSuratKeterangan3A 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Surat Keterangan"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSuratKeterangan3A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   8985
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   6120
      TabIndex        =   9
      Top             =   6840
      Width           =   1215
   End
   Begin VB.TextBox txtKesimpulan2 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2355
      Left            =   240
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   4440
      Width           =   8415
   End
   Begin VB.TextBox txtUmur 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   4200
      TabIndex        =   2
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtJenisKelamin 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdOut 
      Caption         =   "Tutup"
      Height          =   495
      Left            =   7440
      TabIndex        =   10
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "PERMINTAAN PEMERIKSAAN RADIOLOGY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   0
      TabIndex        =   11
      Top             =   1080
      Width           =   8895
      Begin VB.TextBox txtKodePermintaan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   8760
         TabIndex        =   29
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtPermintaan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   8415
      End
      Begin VB.CheckBox chkSemua 
         Caption         =   "Cek Semua"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8760
         TabIndex        =   28
         Top             =   4920
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkRontgent 
         Caption         =   "Rontgent"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8640
         TabIndex        =   27
         Top             =   2520
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtKiriman 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   6720
         TabIndex        =   6
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtNoCM 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   6720
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtAlamat 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1560
         TabIndex        =   4
         Top             =   1650
         Width           =   3135
      End
      Begin VB.TextBox txtPekerjaan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1560
         TabIndex        =   3
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox txtNama 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   375
         Left            =   6720
         TabIndex        =   23
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   114950147
         UpDown          =   -1  'True
         CurrentDate     =   38209
      End
      Begin MSComCtl2.DTPicker dtpAwal2 
         Height          =   375
         Left            =   6720
         TabIndex        =   24
         Top             =   1680
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   114950147
         UpDown          =   -1  'True
         CurrentDate     =   38209
      End
      Begin MSComctlLib.ListView lvRontgen 
         Height          =   1935
         Left            =   8760
         TabIndex        =   26
         Top             =   2880
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   3413
         View            =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSDataGridLib.DataGrid dgPemeriksaan 
         Height          =   2295
         Left            =   240
         TabIndex        =   31
         Top             =   2880
         Visible         =   0   'False
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   2
         RowHeight       =   16
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
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
               LCID            =   1057
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
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan"
         Height          =   210
         Left            =   240
         TabIndex        =   30
         Top             =   2640
         Width           =   945
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Ront Lama"
         Height          =   210
         Left            =   5640
         TabIndex        =   22
         Top             =   1680
         Width           =   870
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Kiriman dari"
         Height          =   210
         Left            =   5640
         TabIndex        =   21
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Nomor CM"
         Height          =   210
         Left            =   5640
         TabIndex        =   20
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal"
         Height          =   210
         Left            =   5640
         TabIndex        =   19
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Umur"
         Height          =   210
         Left            =   3720
         TabIndex        =   18
         Top             =   720
         Width           =   435
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pemeriksaan Rontgen yang diminta"
         Height          =   210
         Left            =   240
         TabIndex        =   17
         Top             =   2040
         Width           =   3315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Alamat"
         Height          =   210
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Pekerjaan"
         Height          =   210
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1020
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   13
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Keterangan Klinik Penderita"
      Height          =   210
      Left            =   240
      TabIndex        =   25
      Top             =   3840
      Width           =   2235
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   6960
      Picture         =   "frmSuratKeterangan3A.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmSuratKeterangan3A.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmSuratKeterangan3A.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmSuratKeterangan3A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
