VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDaftarPasienLama 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Pasien Lama"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPasienLama.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   12150
   Begin VB.Frame Frame1 
      Caption         =   "Periode Dirawat"
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
      Left            =   0
      TabIndex        =   0
      Top             =   6600
      Width           =   5055
      Begin VB.CommandButton cmdCari 
         Caption         =   "&Cari"
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   -120
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   350
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   49807363
         CurrentDate     =   38209
      End
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   350
         Left            =   2760
         TabIndex        =   3
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   49807363
         CurrentDate     =   38209
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "s/d"
         Height          =   210
         Left            =   2400
         TabIndex        =   4
         Top             =   435
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cari Data Pasien"
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
      Left            =   5160
      TabIndex        =   6
      Top             =   6600
      Width           =   6975
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton cmdPeriksaDiagnosa 
         Caption         =   "Periksa &Diagnosa "
         Height          =   375
         Left            =   3000
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "&Tutup"
         Height          =   375
         Left            =   5040
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan Nama Pasien / No. CM"
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2640
      End
   End
   Begin MSDataGridLib.DataGrid dgPasienMeninggal 
      Height          =   5415
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   9551
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   16
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin VB.Image Image1 
      Height          =   975
      Left            =   2640
      Picture         =   "frmDaftarPasienLama.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarPasienLama.frx":431A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmDaftarPasienLama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
End Sub
