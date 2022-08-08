VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSuratKeteranganKeluhan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Surat Keterangan"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSuratKeteranganKeluhan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   15270
   Begin VB.Frame Frame6 
      Height          =   1575
      Left            =   0
      TabIndex        =   85
      Top             =   6600
      Width           =   15135
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
         Height          =   915
         Left            =   120
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   480
         Width           =   7935
      End
      Begin VB.TextBox txtAnjuran2 
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
         Height          =   915
         Left            =   8160
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   480
         Width           =   6855
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Anjuran"
         Height          =   210
         Left            =   8160
         TabIndex        =   87
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Kesimpulan"
         Height          =   210
         Left            =   120
         TabIndex        =   86
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.TextBox txtAudiometri 
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
      Height          =   315
      Left            =   11640
      MaxLength       =   200
      TabIndex        =   34
      Top             =   6000
      Width           =   3375
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "Simpan"
      Height          =   495
      Left            =   2400
      TabIndex        =   37
      Top             =   8280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtBedah 
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
      Height          =   315
      Left            =   11640
      MaxLength       =   200
      TabIndex        =   30
      Top             =   4560
      Width           =   3375
   End
   Begin VB.TextBox txtTHT 
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
      Height          =   315
      Left            =   7560
      MaxLength       =   200
      TabIndex        =   20
      Top             =   2640
      Width           =   5535
   End
   Begin VB.TextBox txtMataKiri1 
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
      Height          =   315
      Left            =   12720
      MaxLength       =   200
      TabIndex        =   16
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txtMataKanan1 
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
      Height          =   315
      Left            =   10320
      MaxLength       =   200
      TabIndex        =   14
      Top             =   1680
      Width           =   375
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   5280
      TabIndex        =   59
      Top             =   960
      Width           =   9975
      Begin VB.Frame Frame8 
         Caption         =   "KEPALA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   1440
         TabIndex        =   71
         Top             =   240
         Width           =   8055
         Begin VB.TextBox TxtLeher 
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
            Height          =   315
            Left            =   840
            MaxLength       =   200
            TabIndex        =   22
            Top             =   2160
            Width           =   5535
         End
         Begin VB.TextBox txtGigi 
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
            Height          =   315
            Left            =   840
            MaxLength       =   200
            TabIndex        =   21
            Top             =   1800
            Width           =   5535
         End
         Begin VB.Frame Frame5 
            Caption         =   "Mata"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   360
            TabIndex        =   72
            Top             =   240
            Width           =   7455
            Begin VB.TextBox txtMataKiri2 
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
               Height          =   315
               Left            =   6240
               MaxLength       =   200
               TabIndex        =   17
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txtMataKanan2 
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
               Height          =   315
               Left            =   3840
               MaxLength       =   200
               TabIndex        =   15
               Top             =   240
               Width           =   375
            End
            Begin VB.CheckBox chkYa 
               Caption         =   "Ya"
               Height          =   255
               Left            =   2280
               TabIndex        =   18
               Top             =   600
               Width           =   975
            End
            Begin VB.CheckBox chkTidak 
               Caption         =   "Tidak"
               Height          =   255
               Left            =   3960
               TabIndex        =   19
               Top             =   600
               Width           =   975
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               Caption         =   "/"
               Height          =   210
               Left            =   6120
               TabIndex        =   89
               Top             =   250
               Width           =   75
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               Caption         =   "/"
               Height          =   210
               Left            =   3720
               TabIndex        =   88
               Top             =   250
               Width           =   75
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Ketajaman Penglihatan"
               Height          =   210
               Left            =   240
               TabIndex        =   76
               Top             =   240
               Width           =   1860
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Mata Kanan"
               Height          =   330
               Left            =   2280
               TabIndex        =   75
               Top             =   240
               Width           =   945
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Mata Kiri"
               Height          =   210
               Left            =   4920
               TabIndex        =   74
               Top             =   240
               Width           =   675
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Buta Warna"
               Height          =   210
               Left            =   240
               TabIndex        =   73
               Top             =   600
               Width           =   960
            End
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Leher"
            Height          =   210
            Left            =   240
            TabIndex        =   79
            Top             =   2160
            Width           =   465
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Gigi"
            Height          =   210
            Left            =   240
            TabIndex        =   78
            Top             =   1800
            Width           =   285
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "THT"
            Height          =   210
            Left            =   240
            TabIndex        =   77
            Top             =   1440
            Width           =   360
         End
      End
      Begin VB.TextBox TxtParu 
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
         Height          =   315
         Left            =   1200
         MaxLength       =   200
         TabIndex        =   24
         Top             =   3600
         Width           =   3615
      End
      Begin VB.TextBox txtJantung 
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
         Height          =   315
         Left            =   1200
         MaxLength       =   200
         TabIndex        =   23
         Top             =   3240
         Width           =   3615
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   120
         TabIndex        =   60
         Top             =   3000
         Width           =   9735
         Begin VB.TextBox txtTreadmill 
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
            Height          =   315
            Left            =   6240
            MaxLength       =   200
            TabIndex        =   33
            Top             =   1680
            Width           =   3375
         End
         Begin VB.TextBox txtUSG 
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
            Height          =   315
            Left            =   1080
            MaxLength       =   200
            TabIndex        =   28
            Top             =   2040
            Width           =   3615
         End
         Begin VB.TextBox txtElekto 
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
            Height          =   315
            Left            =   6240
            MaxLength       =   200
            TabIndex        =   32
            Top             =   1320
            Width           =   3375
         End
         Begin VB.TextBox txtRadiologi 
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
            Height          =   315
            Left            =   6240
            MaxLength       =   200
            TabIndex        =   31
            Top             =   960
            Width           =   3375
         End
         Begin VB.TextBox txtEsminitas 
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
            Height          =   315
            Left            =   1080
            MaxLength       =   200
            TabIndex        =   26
            Top             =   1320
            Width           =   3615
         End
         Begin VB.TextBox txtPapsmear 
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
            Height          =   315
            Left            =   6960
            MaxLength       =   200
            TabIndex        =   40
            Top             =   3720
            Width           =   1815
         End
         Begin VB.TextBox txtLaboratorium 
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
            Height          =   315
            Left            =   1080
            MaxLength       =   200
            TabIndex        =   27
            Top             =   1680
            Width           =   3615
         End
         Begin VB.TextBox txtKesimpulan 
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
            Height          =   315
            Left            =   1320
            MaxLength       =   200
            TabIndex        =   41
            Top             =   4320
            Width           =   6855
         End
         Begin VB.TextBox txtAnjuran 
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
            Height          =   315
            Left            =   4000
            MaxLength       =   200
            TabIndex        =   42
            Top             =   4680
            Visible         =   0   'False
            Width           =   6855
         End
         Begin VB.TextBox txtPenyakitDalam 
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
            Height          =   315
            Left            =   6240
            MaxLength       =   200
            TabIndex        =   29
            Top             =   240
            Width           =   3375
         End
         Begin VB.TextBox txtPerut 
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
            Height          =   315
            Left            =   1080
            MaxLength       =   200
            TabIndex        =   25
            Top             =   960
            Width           =   3615
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Audiometri"
            Height          =   210
            Left            =   4920
            TabIndex        =   84
            Top             =   2040
            Width           =   885
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Treadmill"
            Height          =   210
            Left            =   4920
            TabIndex        =   83
            Top             =   1680
            Width           =   720
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "USG "
            Height          =   210
            Left            =   240
            TabIndex        =   82
            Top             =   2040
            Width           =   405
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "ElekroRadiografi"
            Height          =   210
            Left            =   4920
            TabIndex        =   81
            Top             =   1320
            Width           =   1275
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Radiologi"
            Height          =   210
            Left            =   4920
            TabIndex        =   80
            Top             =   960
            Width           =   705
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Anjuran"
            Height          =   210
            Left            =   240
            TabIndex        =   70
            Top             =   4800
            Width           =   630
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Kesimpulan"
            Height          =   210
            Left            =   240
            TabIndex        =   69
            Top             =   4320
            Width           =   900
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Papsmear"
            Height          =   210
            Left            =   6960
            TabIndex        =   68
            Top             =   3480
            Width           =   780
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Lab"
            Height          =   210
            Left            =   240
            TabIndex        =   67
            Top             =   1680
            Width           =   285
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Penyakit Dalam"
            Height          =   210
            Left            =   4920
            TabIndex        =   66
            Top             =   240
            Width           =   1230
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Bedah"
            Height          =   210
            Left            =   4920
            TabIndex        =   65
            Top             =   600
            Width           =   510
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Estrimitas"
            Height          =   210
            Left            =   240
            TabIndex        =   64
            Top             =   1320
            Width           =   765
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Perut"
            Height          =   210
            Left            =   240
            TabIndex        =   63
            Top             =   960
            Width           =   450
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Jantung"
            Height          =   210
            Left            =   240
            TabIndex        =   62
            Top             =   240
            Width           =   660
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Paru"
            Height          =   210
            Left            =   240
            TabIndex        =   61
            Top             =   600
            Width           =   360
         End
      End
   End
   Begin VB.TextBox txtUmur 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1560
      TabIndex        =   1
      Top             =   1680
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pengujian"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      TabIndex        =   49
      Top             =   3360
      Width           =   5175
      Begin VB.TextBox txtRiwayat 
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
         Height          =   315
         Left            =   1560
         MaxLength       =   200
         TabIndex        =   6
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txtKeluhan 
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
         Height          =   315
         Left            =   1560
         MaxLength       =   200
         TabIndex        =   5
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox txtNadi 
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
         Height          =   315
         Left            =   1560
         MaxLength       =   200
         TabIndex        =   12
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtPernapasan 
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
         Height          =   315
         Left            =   3840
         MaxLength       =   200
         TabIndex        =   11
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtTekanan2 
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
         Height          =   315
         Left            =   2160
         MaxLength       =   200
         TabIndex        =   10
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtTekanan 
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
         Height          =   315
         Left            =   1560
         MaxLength       =   200
         TabIndex        =   9
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtBerat 
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
         Height          =   315
         Left            =   3840
         MaxLength       =   200
         TabIndex        =   8
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtTinggi 
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
         Height          =   315
         Left            =   1560
         MaxLength       =   200
         TabIndex        =   7
         Top             =   1155
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   2400
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
         Format          =   115998723
         UpDown          =   -1  'True
         CurrentDate     =   38209
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal"
         Height          =   210
         Left            =   120
         TabIndex        =   90
         Top             =   2520
         Width           =   645
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Riwayat Penyakit"
         Height          =   210
         Left            =   120
         TabIndex        =   58
         Top             =   840
         Width           =   1380
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Keluhan"
         Height          =   210
         Left            =   120
         TabIndex        =   57
         Top             =   480
         Width           =   645
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nadi"
         Height          =   210
         Left            =   120
         TabIndex        =   56
         Top             =   1920
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Pernafasan"
         Height          =   210
         Left            =   2760
         TabIndex        =   55
         Top             =   1560
         Width           =   885
      End
      Begin VB.Label Label13 
         Caption         =   "/"
         Height          =   255
         Left            =   2000
         TabIndex        =   53
         Top             =   1560
         Width           =   135
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Tekanan Darah"
         Height          =   210
         Left            =   120
         TabIndex        =   52
         Top             =   1560
         Width           =   1230
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Berat Badan"
         Height          =   210
         Left            =   2760
         TabIndex        =   51
         Top             =   1200
         Width           =   990
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tinggi Badan"
         Height          =   210
         Left            =   120
         TabIndex        =   50
         Top             =   1200
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdOut 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   7920
      TabIndex        =   39
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Cetak"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   38
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      TabIndex        =   43
      Top             =   1080
      Width           =   5175
      Begin VB.TextBox txtAlamat 
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
         Height          =   315
         Left            =   1560
         MaxLength       =   200
         TabIndex        =   4
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox txtNoCM 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1560
         TabIndex        =   2
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtNama 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   2535
      End
      Begin MSDataListLib.DataCombo dcGolDarah 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Umur"
         Height          =   210
         Left            =   240
         TabIndex        =   54
         Top             =   600
         Width           =   435
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Alamat"
         Height          =   210
         Left            =   240
         TabIndex        =   48
         Top             =   1680
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Golongan Darah"
         Height          =   210
         Left            =   240
         TabIndex        =   47
         Top             =   1320
         Width           =   1290
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "No CM"
         Height          =   210
         Left            =   240
         TabIndex        =   46
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   240
         TabIndex        =   44
         Top             =   240
         Width           =   1020
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   45
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
      Left            =   13320
      Picture         =   "frmSuratKeteranganKeluhan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmSuratKeteranganKeluhan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13575
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmSuratKeteranganKeluhan.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmSuratKeteranganKeluhan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkTidak_Click()
    If chkTidak.Value = 1 Then
        chkYa.Enabled = False
    End If

    If chkTidak.Value = 0 Then
        chkYa.Enabled = True
    End If

End Sub

Private Sub chkTidak_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTHT.SetFocus
End Sub

Private Sub chkYa_Click()
    If chkYa.Value = 1 Then
        chkTidak.Enabled = False
    End If

    If chkYa.Value = 0 Then
        chkTidak.Enabled = True
    End If

End Sub

Private Sub chkYa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTHT.SetFocus
End Sub

Private Sub cmdOut_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error Resume Next
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmCetakSuratKeteranganKeluhan.Show
End Sub

Private Sub dcGolDarah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcGolDarah.MatchedWithList = True Then txtAlamat.SetFocus
        strSQL = "Select kdgolongandarah, golongandarah from GolonganDarah Where StatusEnabled='1' and (GolonganDarah LIKE '%" & dcGolDarah.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcGolDarah.Text = ""
            txtAlamat.SetFocus
            Exit Sub
        End If
        dcGolDarah.BoundText = rs(0).Value
        dcGolDarah.Text = rs(1).Value
    End If
End Sub

Private Sub dtpAwal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtMataKanan1.SetFocus
End Sub

Private Sub Form_Load()
    On Error GoTo hell
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call SetComboGolonganDarah
    dtpAwal.Value = Format(Now, "dd MMM yyyy 00:00:00")

    txtNama.Text = frmDaftarPasienRJ.dgDaftarPasienRJ.Columns("Nama Pasien").Value
    txtUmur.Text = frmDaftarPasienRJ.dgDaftarPasienRJ.Columns("UmurTahun").Value
    txtNoCM.Text = frmDaftarPasienRJ.dgDaftarPasienRJ.Columns("NoCM").Value
    txtAlamat.Text = frmDaftarPasienRJ.dgDaftarPasienRJ.Columns("ALamat").Value

    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub SetComboGolonganDarah()
    Set rs = Nothing
    rs.Open "Select * from GolonganDarah Where StatusEnabled='1'", dbConn, , adLockOptimistic
    Set dcGolDarah.RowSource = rs
    dcGolDarah.ListField = rs.Fields(1).Name
    dcGolDarah.BoundColumn = rs.Fields(0).Name
End Sub

Private Sub txtAlamat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeluhan.SetFocus

End Sub

Private Sub txtAnjuran2_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then cmdPrint.SetFocus
End Sub

Private Sub txtAudiometri_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtKesimpulan2.SetFocus
End Sub

Private Sub txtBedah_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtRadiologi.SetFocus
End Sub

Private Sub txtElekto_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtTreadmill.SetFocus
End Sub

Private Sub txtEsminitas_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtLaboratorium.SetFocus
End Sub

Private Sub txtGigi_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then TxtLeher.SetFocus
End Sub

Private Sub txtJantung_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then TxtParu.SetFocus
End Sub

Private Sub txtKeluhan_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtRiwayat.SetFocus
End Sub

Private Sub txtKesimpulan2_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtAnjuran2.SetFocus
End Sub

Private Sub txtLaboratorium_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtUSG.SetFocus
End Sub

Private Sub TxtLeher_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtJantung.SetFocus
End Sub

Private Sub txtNama_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtUmur.SetFocus
End Sub

Private Sub txtNoCM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcGolDarah.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub TxtParu_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtPerut.SetFocus
End Sub

Private Sub txtPenyakitDalam_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtBedah.SetFocus
End Sub

Private Sub txtPerut_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtEsminitas.SetFocus
End Sub

Private Sub txtRadiologi_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtElekto.SetFocus
End Sub

Private Sub txtRiwayat_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtTinggi.SetFocus
End Sub

Private Sub txtTHT_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtGigi.SetFocus
End Sub

Private Sub txtTinggi_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtBerat.SetFocus
End Sub

Private Sub txtBerat_kEYpRESS(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtTekanan.SetFocus
End Sub

Private Sub txtPernapasan_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtNadi.SetFocus
End Sub

Private Sub txtTekanan_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtTekanan2.SetFocus
End Sub

Private Sub txtTekanan2_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtPernapasan.SetFocus
End Sub

Private Sub txtNadi_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then dtpAwal.SetFocus
End Sub

Private Sub txtMataKanan1_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtMataKanan2.SetFocus
End Sub

Private Sub txtMataKanan2_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtMataKiri1.SetFocus
End Sub

Private Sub txtMataKiri1_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtMataKiri2.SetFocus
End Sub

Private Sub txtMataKiri2_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then chkYa.SetFocus
End Sub

Private Sub txtTreadmill_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtAudiometri.SetFocus
End Sub

Private Sub txtUmur_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNoCM.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtUSG_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtPenyakitDalam.SetFocus
End Sub
