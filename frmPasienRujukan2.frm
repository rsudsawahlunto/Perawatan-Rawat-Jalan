VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPasienRujukan2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pasien Konsul ke Unit Lain"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPasienRujukan2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   11295
   Begin VB.Frame Frame4 
      Caption         =   "Pemeriksaan Sederhana"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   11295
      Begin VB.Frame Frame23 
         Height          =   1215
         Left            =   7440
         TabIndex        =   176
         Top             =   6480
         Width           =   3735
         Begin VB.CommandButton cmdSimpan 
            Caption         =   "&Simpan"
            Height          =   465
            Left            =   480
            TabIndex        =   178
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton cmdTutup 
            Caption         =   "Tutu&p"
            Height          =   465
            Left            =   2160
            TabIndex        =   177
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   7440
         TabIndex        =   172
         Top             =   5640
         Width           =   3735
         Begin VB.Frame Frame20 
            Caption         =   "Status CITO"
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
            Left            =   1080
            TabIndex        =   173
            Top             =   120
            Width           =   1695
            Begin VB.OptionButton optCito 
               Caption         =   "Ya"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   175
               Top             =   240
               Width           =   615
            End
            Begin VB.OptionButton optCito 
               Caption         =   "Tidak"
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   174
               Top             =   240
               Value           =   -1  'True
               Width           =   735
            End
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "Khusus"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   3840
         TabIndex        =   159
         Top             =   5640
         Width           =   3495
         Begin VB.CheckBox Check60 
            Caption         =   "Fistulografi"
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
            Left            =   2040
            TabIndex        =   171
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CheckBox Check59 
            Caption         =   "RPG"
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
            Left            =   2040
            TabIndex        =   170
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CheckBox Check58 
            Caption         =   "Urethrogram"
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
            Left            =   2040
            TabIndex        =   169
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CheckBox Check57 
            Caption         =   "HSG"
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
            Left            =   2040
            TabIndex        =   168
            Top             =   840
            Width           =   735
         End
         Begin VB.CheckBox Check56 
            Caption         =   "Myelografi"
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
            Left            =   2040
            TabIndex        =   167
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox Check55 
            Caption         =   "Cystogram"
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
            Left            =   2040
            TabIndex        =   166
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox Check54 
            Caption         =   "OMD"
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
            Left            =   120
            TabIndex        =   165
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CheckBox Check53 
            Caption         =   "Maag Duodenum"
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
            Left            =   120
            TabIndex        =   164
            Top             =   1320
            Width           =   1575
         End
         Begin VB.CheckBox Check52 
            Caption         =   "Oesophagogram"
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
            Left            =   120
            TabIndex        =   163
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CheckBox Check51 
            Caption         =   "Appendicogram"
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
            Left            =   120
            TabIndex        =   162
            Top             =   840
            Width           =   1455
         End
         Begin VB.CheckBox Check50 
            Caption         =   "Colon Inloop"
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
            Left            =   120
            TabIndex        =   161
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox Check49 
            Caption         =   "BNO - IVP"
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
            Left            =   120
            TabIndex        =   160
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "Pemeriksaan Sedang"
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
         Left            =   7440
         TabIndex        =   153
         Top             =   4440
         Width           =   3735
         Begin VB.CheckBox Check48 
            Caption         =   "USG"
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
            Left            =   1680
            TabIndex        =   158
            Top             =   480
            Width           =   1215
         End
         Begin VB.CheckBox Check47 
            Caption         =   "Scoliosis Program"
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
            Left            =   1680
            TabIndex        =   157
            Top             =   240
            Width           =   1695
         End
         Begin VB.CheckBox Check46 
            Caption         =   "Bone - Survey"
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
            Left            =   240
            TabIndex        =   156
            Top             =   720
            Width           =   1455
         End
         Begin VB.CheckBox Check45 
            Caption         =   "Lopografi"
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
            Left            =   240
            TabIndex        =   155
            Top             =   480
            Width           =   1215
         End
         Begin VB.CheckBox Check44 
            Caption         =   "Cor Analisa"
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
            Left            =   240
            TabIndex        =   154
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CheckBox Check35 
         Caption         =   "PA"
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
         Left            =   1920
         TabIndex        =   139
         Top             =   480
         Width           =   615
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
         Height          =   1875
         Left            =   120
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   137
         Top             =   5760
         Width           =   3615
      End
      Begin VB.TextBox txtJmlPelayanan 
         Height          =   315
         Left            =   2640
         TabIndex        =   136
         Top             =   6600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   2520
         TabIndex        =   135
         Top             =   6480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame Frame10 
         Caption         =   "Extremitas Superior"
         Height          =   2295
         Left            =   7440
         TabIndex        =   29
         Top             =   2160
         Width           =   3735
         Begin VB.CheckBox Check28 
            Caption         =   "Sinistra"
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
            Left            =   2760
            TabIndex        =   101
            Top             =   1920
            Width           =   855
         End
         Begin VB.CheckBox Check27 
            Caption         =   "Sinistra"
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
            Left            =   2760
            TabIndex        =   100
            Top             =   1680
            Width           =   855
         End
         Begin VB.CheckBox Check26 
            Caption         =   "Sinistra"
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
            Left            =   2760
            TabIndex        =   99
            Top             =   1440
            Width           =   855
         End
         Begin VB.CheckBox Check25 
            Caption         =   "Sinistra"
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
            Left            =   2760
            TabIndex        =   98
            Top             =   1200
            Width           =   855
         End
         Begin VB.CheckBox Check24 
            Caption         =   "Sinistra"
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
            Left            =   2760
            TabIndex        =   97
            Top             =   960
            Width           =   855
         End
         Begin VB.CheckBox Check23 
            Caption         =   "Sinistra"
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
            Left            =   2760
            TabIndex        =   96
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox Check22 
            Caption         =   "Sinistra"
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
            Left            =   2760
            TabIndex        =   95
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox Check21 
            Caption         =   "Sinistra"
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
            Left            =   2760
            TabIndex        =   94
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox Check20 
            Caption         =   "Dextra"
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
            Left            =   1920
            TabIndex        =   93
            Top             =   1920
            Width           =   855
         End
         Begin VB.CheckBox Check19 
            Caption         =   "Dextra"
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
            Left            =   1920
            TabIndex        =   92
            Top             =   1680
            Width           =   855
         End
         Begin VB.CheckBox Check18 
            Caption         =   "Dextra"
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
            Left            =   1920
            TabIndex        =   91
            Top             =   1440
            Width           =   855
         End
         Begin VB.CheckBox Check17 
            Caption         =   "Dextra"
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
            Left            =   1920
            TabIndex        =   90
            Top             =   1200
            Width           =   855
         End
         Begin VB.CheckBox Check16 
            Caption         =   "Dextra"
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
            Left            =   1920
            TabIndex        =   89
            Top             =   960
            Width           =   855
         End
         Begin VB.CheckBox Check15 
            Caption         =   "Dextra"
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
            Left            =   1920
            TabIndex        =   88
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox Check14 
            Caption         =   "Dextra"
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
            Left            =   1920
            TabIndex        =   87
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Dextra"
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
            Left            =   1920
            TabIndex        =   86
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox chk111032 
            Caption         =   "Manus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   71
            Top             =   1920
            Width           =   855
         End
         Begin VB.CheckBox chk111031 
            Caption         =   "Wrist Joint  "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   70
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CheckBox chk111030 
            Caption         =   "Antebrachii"
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
            Left            =   240
            TabIndex        =   69
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CheckBox chk111029 
            Caption         =   "Humerus / Brachii "
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
            Left            =   240
            TabIndex        =   68
            Top             =   960
            Width           =   1815
         End
         Begin VB.CheckBox chk111026 
            Caption         =   "Scapula"
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
            Left            =   240
            TabIndex        =   67
            Top             =   480
            Width           =   1095
         End
         Begin VB.CheckBox chk111027 
            Caption         =   "Clavicula"
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
            Left            =   240
            TabIndex        =   66
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chkCoxal 
            Caption         =   "Shoulder"
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
            Index           =   13
            Left            =   240
            TabIndex        =   31
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox chkCoxal 
            Caption         =   "Art.Cubiti/Elbow  "
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
            Index           =   11
            Left            =   240
            TabIndex        =   30
            Top             =   1200
            Width           =   1575
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "PELVIS"
         Height          =   1095
         Left            =   6000
         TabIndex        =   26
         Top             =   4440
         Width           =   1335
         Begin VB.CheckBox chkPelvisss 
            Caption         =   "Pelvis"
            Height          =   255
            Left            =   120
            TabIndex        =   128
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox chk111025 
            Caption         =   "Coxcae"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   102
            Top             =   720
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "ABDOMEN"
         Height          =   1815
         Left            =   3360
         TabIndex        =   13
         Top             =   240
         Width           =   3975
         Begin VB.CheckBox chk111008 
            Caption         =   "Abdomen anak 2 posisi"
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
            Left            =   240
            TabIndex        =   133
            Top             =   720
            Width           =   1935
         End
         Begin VB.CheckBox chk111007 
            Caption         =   "Abdomen 3 posisi"
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
            Left            =   240
            TabIndex        =   132
            Top             =   480
            Width           =   1575
         End
         Begin VB.CheckBox chkBNOLateral 
            Caption         =   "Lateral"
            Height          =   255
            Left            =   2880
            TabIndex        =   131
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox chkBNOAP 
            Caption         =   "AP"
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
            Left            =   2280
            TabIndex        =   130
            Top             =   240
            Width           =   495
         End
         Begin VB.CheckBox chkBNOaplateral 
            Caption         =   "BNO"
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
            Left            =   240
            TabIndex        =   129
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox chk111051 
            Caption         =   "Babygram"
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
            Left            =   240
            TabIndex        =   15
            Top             =   1200
            Width           =   1095
         End
         Begin VB.CheckBox chk111042 
            Caption         =   "Invertogram"
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
            Left            =   240
            TabIndex        =   14
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "AP + Lat"
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
            Left            =   2880
            TabIndex        =   18
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "AP + Lat"
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
            Left            =   2880
            TabIndex        =   17
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "AP"
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
            Left            =   2880
            TabIndex        =   16
            Top             =   1200
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "CRANIUM"
         Height          =   3375
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   3615
         Begin VB.CheckBox Check43 
            Caption         =   "For.Opticum"
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
            Left            =   2280
            TabIndex        =   148
            Top             =   2880
            Width           =   1215
         End
         Begin VB.CheckBox Check42 
            Caption         =   "Cadwel"
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
            Left            =   1440
            TabIndex        =   147
            Top             =   2880
            Width           =   855
         End
         Begin VB.CheckBox Check41 
            Caption         =   "Sinistra"
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
            Left            =   2400
            TabIndex        =   146
            Top             =   2400
            Width           =   855
         End
         Begin VB.CheckBox Check40 
            Caption         =   "Dextra"
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
            Left            =   1440
            TabIndex        =   145
            Top             =   2400
            Width           =   855
         End
         Begin VB.CheckBox chk111011 
            Caption         =   "Basis Cranii"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   127
            Top             =   3120
            Width           =   1215
         End
         Begin VB.CheckBox chk111012 
            Caption         =   "Orbita"
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
            Left            =   240
            TabIndex        =   126
            Top             =   2880
            Width           =   855
         End
         Begin VB.CheckBox chk111016 
            Caption         =   "Nasal Bone"
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
            Left            =   240
            TabIndex        =   125
            Top             =   2640
            Width           =   1215
         End
         Begin VB.CheckBox chkEisler 
            Caption         =   "Eisler"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   124
            Top             =   2400
            Width           =   735
         End
         Begin VB.CheckBox chk111014a 
            Caption         =   "TMJ"
            Height          =   255
            Left            =   240
            TabIndex        =   123
            Top             =   2160
            Width           =   735
         End
         Begin VB.CheckBox chk111014 
            Caption         =   "TMJ"
            Height          =   255
            Left            =   240
            TabIndex        =   122
            Top             =   1920
            Width           =   735
         End
         Begin VB.CheckBox chk111013 
            Caption         =   "Mandibula"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   121
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CheckBox chkMaxilla 
            Caption         =   "Maxilla"
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
            Left            =   240
            TabIndex        =   120
            Top             =   1440
            Width           =   855
         End
         Begin VB.CheckBox chk111010 
            Caption         =   "Waters"
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
            Left            =   240
            TabIndex        =   119
            Top             =   1200
            Width           =   975
         End
         Begin VB.CheckBox chk111052 
            Caption         =   "SPN"
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
            Left            =   240
            TabIndex        =   118
            Top             =   960
            Width           =   735
         End
         Begin VB.CheckBox chkStenvers 
            Caption         =   "Stenvers"
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
            Left            =   240
            TabIndex        =   117
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox chk111015 
            Caption         =   "Mastoid"
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
            Left            =   240
            TabIndex        =   116
            Top             =   480
            Width           =   975
         End
         Begin VB.CheckBox chk111009 
            Caption         =   "Scheidel"
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
            Left            =   240
            TabIndex        =   114
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chkTMJSinistra 
            Caption         =   "Sinistra"
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
            Left            =   2400
            TabIndex        =   10
            Top             =   1920
            Width           =   855
         End
         Begin VB.CheckBox chkTMJDextra 
            Caption         =   "Dextra"
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
            Left            =   1440
            TabIndex        =   9
            Top             =   1920
            Width           =   855
         End
         Begin VB.CheckBox chkMastoidSinistra 
            Caption         =   "Sinistra"
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
            Left            =   2520
            TabIndex        =   6
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox chkMastoidDextra 
            Caption         =   "Dextra"
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
            Left            =   1560
            TabIndex        =   5
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "AP + Lat"
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
            Left            =   1560
            TabIndex        =   115
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label19 
            Caption         =   "Lateral"
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
            Left            =   1560
            TabIndex        =   24
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "AP + Lat"
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
            Left            =   1560
            TabIndex        =   12
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Dextra et Sinistra"
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
            Left            =   1560
            TabIndex        =   11
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Dextra et Sinistra"
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
            Left            =   1560
            TabIndex        =   8
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Dextra et Sinistra"
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
            Left            =   1560
            TabIndex        =   7
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "RESPIRASI"
         Height          =   1815
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3135
         Begin VB.CheckBox Check39 
            Caption         =   "Sternum Lat"
            Height          =   255
            Left            =   240
            TabIndex        =   143
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CheckBox Check38 
            Caption         =   "Costae AP"
            Height          =   255
            Left            =   240
            TabIndex        =   142
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CheckBox Check37 
            Caption         =   "Top Lordotic"
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
            Left            =   1800
            TabIndex        =   141
            Top             =   480
            Width           =   1215
         End
         Begin VB.CheckBox Check36 
            Caption         =   "RLD"
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
            Left            =   1200
            TabIndex        =   140
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox Check34 
            Caption         =   "AP"
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
            Left            =   1200
            TabIndex        =   138
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chk111001 
            Caption         =   "Thorax"
            Height          =   255
            Left            =   240
            TabIndex        =   65
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chkThoraxAP 
            Caption         =   "Thorax AP + Lateral"
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
            Index           =   1
            Left            =   240
            TabIndex        =   3
            Top             =   840
            Width           =   1815
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "OTHERS"
         Height          =   1095
         Left            =   3840
         TabIndex        =   19
         Top             =   4440
         Width           =   2055
         Begin VB.CheckBox chkDental 
            Caption         =   "Dental"
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
            Index           =   22
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox chkSoft 
            Caption         =   "Soft Tissue "
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
            Index           =   21
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Leher"
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
            Left            =   360
            TabIndex        =   144
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label12 
            Caption         =   "AP + Lat"
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
            Left            =   1200
            TabIndex        =   22
            Top             =   480
            Width           =   735
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgPelayanan 
         Height          =   615
         Left            =   2040
         TabIndex        =   134
         Top             =   6480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         _Version        =   393216
         Rows            =   50
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   8577768
         BackColorBkg    =   16777215
         FocusRect       =   0
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Frame Frame7 
         Caption         =   "VERTEBRA"
         Height          =   2295
         Left            =   3840
         TabIndex        =   23
         Top             =   2160
         Width           =   3495
         Begin VB.CheckBox Check33 
            Caption         =   "AP + Lat + Obli"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1920
            TabIndex        =   113
            Top             =   2040
            Width           =   1455
         End
         Begin VB.CheckBox Check32 
            Caption         =   "AP + Lat"
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
            Left            =   1920
            TabIndex        =   112
            Top             =   1800
            Width           =   975
         End
         Begin VB.CheckBox chkLumbosacral 
            Caption         =   "Vert.Lumbosacral"
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
            Left            =   240
            TabIndex        =   111
            Top             =   1800
            Width           =   1695
         End
         Begin VB.CheckBox chkVertThoracolumbal 
            Caption         =   "Vert Thoracolumbal"
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
            Left            =   240
            TabIndex        =   110
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CheckBox Check30 
            Caption         =   "AP + Lat"
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
            Left            =   1920
            TabIndex        =   108
            Top             =   1320
            Width           =   975
         End
         Begin VB.CheckBox Check29 
            Caption         =   "AP + Lat + Obli"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1920
            TabIndex        =   107
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CheckBox chkThoracalis 
            Caption         =   "Vert.Thoracalis"
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
            Left            =   240
            TabIndex        =   106
            Top             =   840
            Width           =   1455
         End
         Begin VB.CheckBox chk111018 
            Caption         =   "AP + Lat + Obli"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1920
            TabIndex        =   105
            Top             =   600
            Width           =   1455
         End
         Begin VB.CheckBox chk111017 
            Caption         =   "AP + Lat"
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
            Left            =   1920
            TabIndex        =   104
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox chkCervicali 
            Caption         =   "Vert.Cervicalis"
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
            Left            =   240
            TabIndex        =   103
            Top             =   360
            Width           =   1335
         End
         Begin VB.CheckBox chkVertThoAP 
            Caption         =   "AP + Lat"
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
            Left            =   1920
            TabIndex        =   25
            Top             =   840
            Width           =   975
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Extremitas Inferior"
         Height          =   1815
         Left            =   7440
         TabIndex        =   27
         Top             =   240
         Width           =   3735
         Begin VB.CheckBox Check12 
            Caption         =   "Sinistra"
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
            Left            =   2520
            TabIndex        =   152
            Top             =   1440
            Width           =   855
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Sinitra"
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
            Left            =   2520
            TabIndex        =   151
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Dextra"
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
            Left            =   1680
            TabIndex        =   150
            Top             =   1440
            Width           =   855
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Dextra"
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
            Left            =   1680
            TabIndex        =   85
            Top             =   1200
            Width           =   855
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Dextra"
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
            Left            =   1680
            TabIndex        =   84
            Top             =   960
            Width           =   855
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Sinistra"
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
            Left            =   2520
            TabIndex        =   83
            Top             =   1200
            Width           =   855
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Sinistra"
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
            Left            =   2520
            TabIndex        =   82
            Top             =   960
            Width           =   855
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Sinistra"
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
            Left            =   2520
            TabIndex        =   81
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Dextra"
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
            Left            =   1680
            TabIndex        =   80
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Sinistra"
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
            Left            =   2520
            TabIndex        =   79
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Dextra"
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
            Left            =   1680
            TabIndex        =   78
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Dextra"
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
            Left            =   1680
            TabIndex        =   77
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox chk111050 
            Caption         =   "Calcaneus"
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
            Left            =   240
            TabIndex        =   76
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CheckBox chk111037 
            Caption         =   "Pedis"
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
            Left            =   240
            TabIndex        =   75
            Top             =   1200
            Width           =   735
         End
         Begin VB.CheckBox chk111036 
            Caption         =   "Angkle Joint  "
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
            Left            =   240
            TabIndex        =   74
            Top             =   960
            Width           =   1215
         End
         Begin VB.CheckBox chk111035 
            Caption         =   "Cruris"
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
            Left            =   240
            TabIndex        =   73
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox chk111033 
            Caption         =   "Femur"
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
            Left            =   240
            TabIndex        =   72
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chkCoxal 
            Caption         =   "Art.Genu"
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
            Index           =   1
            Left            =   240
            TabIndex        =   28
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.Label Label6 
         Caption         =   "Keterangan Klinis :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   149
         Top             =   5520
         Width           =   1935
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
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
   Begin VB.Frame Frame11 
      Height          =   375
      Left            =   5400
      TabIndex        =   32
      Top             =   7680
      Width           =   2535
      Begin VB.Frame Frame19 
         Height          =   2415
         Left            =   4080
         TabIndex        =   61
         Top             =   5280
         Width           =   3855
         Begin VB.CheckBox chkSemuaInperium 
            Caption         =   "Cek Semua"
            Height          =   255
            Left            =   240
            TabIndex        =   63
            Top             =   2020
            Width           =   1455
         End
         Begin VB.CheckBox chkExtreIn 
            Caption         =   "Extre Inperium"
            Height          =   210
            Left            =   240
            TabIndex        =   62
            Top             =   240
            Width           =   2175
         End
         Begin MSComctlLib.ListView lvExtremInper 
            Height          =   1455
            Left            =   120
            TabIndex        =   64
            Top             =   480
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   2566
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
      End
      Begin VB.Frame Frame18 
         Height          =   2415
         Left            =   120
         TabIndex        =   57
         Top             =   5280
         Width           =   3855
         Begin VB.CheckBox chkExtre 
            Caption         =   "Extre"
            Height          =   210
            Left            =   240
            TabIndex        =   60
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox chkSemuaExtre 
            Caption         =   "Cek Semua"
            Height          =   255
            Left            =   240
            TabIndex        =   58
            Top             =   2020
            Width           =   1455
         End
         Begin MSComctlLib.ListView lvExtre 
            Height          =   1455
            Left            =   120
            TabIndex        =   59
            Top             =   480
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   2566
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
      End
      Begin VB.Frame Frame17 
         Height          =   2535
         Left            =   8040
         TabIndex        =   53
         Top             =   2760
         Width           =   2415
         Begin VB.CheckBox chkPel 
            Caption         =   "Pelvis"
            Height          =   255
            Left            =   240
            TabIndex        =   56
            Top             =   220
            Width           =   1335
         End
         Begin VB.CheckBox chkSemuaPelvis 
            Caption         =   "Cek Semua"
            Height          =   255
            Left            =   240
            TabIndex        =   54
            Top             =   2180
            Width           =   1455
         End
         Begin MSComctlLib.ListView lvPelvis 
            Height          =   1695
            Left            =   120
            TabIndex        =   55
            Top             =   480
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   2990
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
      End
      Begin VB.Frame Frame16 
         Height          =   2535
         Left            =   4080
         TabIndex        =   49
         Top             =   2760
         Width           =   3855
         Begin VB.CheckBox chkVertebra 
            Caption         =   "Vertebra"
            Height          =   255
            Left            =   360
            TabIndex        =   52
            Top             =   220
            Width           =   1095
         End
         Begin VB.CheckBox chkSemuaVertebra 
            Caption         =   "Cek Semua"
            Height          =   255
            Left            =   240
            TabIndex        =   50
            Top             =   2180
            Width           =   1455
         End
         Begin MSComctlLib.ListView lvVertebra 
            Height          =   1695
            Left            =   120
            TabIndex        =   51
            Top             =   480
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   2990
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
      End
      Begin VB.Frame Frame15 
         Height          =   2535
         Left            =   8040
         TabIndex        =   45
         Top             =   240
         Width           =   2415
         Begin VB.CheckBox chkSemuaOther 
            Caption         =   "Cek Semua"
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   2160
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chkOther 
            Caption         =   "Other"
            Height          =   210
            Left            =   360
            TabIndex        =   46
            Top             =   240
            Width           =   1215
         End
         Begin MSComctlLib.ListView lvOther 
            Height          =   1575
            Left            =   120
            TabIndex        =   48
            Top             =   480
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   2778
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
      End
      Begin VB.CheckBox chkCranium 
         Caption         =   "Cranium"
         Height          =   210
         Left            =   360
         TabIndex        =   43
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Frame Frame14 
         Height          =   2535
         Left            =   4080
         TabIndex        =   40
         Top             =   240
         Width           =   3855
         Begin VB.CheckBox chkAbdo 
            Caption         =   "Abdomen"
            Height          =   210
            Left            =   360
            TabIndex        =   44
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox chkSemuaAbdomen 
            Caption         =   "Cek Semua"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   2160
            Width           =   1455
         End
         Begin MSComctlLib.ListView lvAbdomen 
            Height          =   1575
            Left            =   120
            TabIndex        =   42
            Top             =   480
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   2778
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
      End
      Begin VB.Frame Frame13 
         Height          =   2535
         Left            =   120
         TabIndex        =   37
         Top             =   2760
         Width           =   3855
         Begin VB.CheckBox chkSemuaCranium 
            Caption         =   "Cek Semua"
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   2180
            Width           =   1455
         End
         Begin MSComctlLib.ListView lvCranium 
            Height          =   1695
            Left            =   120
            TabIndex        =   39
            Top             =   480
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   2990
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
      End
      Begin VB.CheckBox chkRespirasi 
         Caption         =   "Respirasi"
         Height          =   210
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   1095
      End
      Begin VB.Frame Frame12 
         Height          =   1935
         Left            =   0
         TabIndex        =   34
         Top             =   720
         Width           =   3855
         Begin VB.CheckBox chkSemuaRespirasi 
            Caption         =   "Cek Semua"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   2160
            Width           =   1455
         End
         Begin MSComctlLib.ListView lvRespirasi 
            Height          =   1575
            Left            =   120
            TabIndex        =   35
            Top             =   480
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   2778
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
      End
   End
   Begin VB.CheckBox Check31 
      Caption         =   "AP + Lat + Obli"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6240
      TabIndex        =   109
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPasienRujukan2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   9480
      Picture         =   "frmPasienRujukan2.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPasienRujukan2.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "frmPasienRujukan2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFilterPelayanan As String
Dim strCito As String
Dim strKodePelayananRS As String
Dim strNamaPelayananRS As String
Dim curBiaya As Currency
Dim curJP As Currency
Dim intJmlPelayanan As Integer
Dim strKdKelas As String
Dim strKelas As String
Dim strKdJenisTarif As String
Dim strJenisTarif As String
Dim subcurTarifBiayaSatuan As Currency
Dim subcurTarifHargaSatuan As Currency
Dim strStatusAPBD As String
Dim MyIndex As Integer
Dim xRow As Integer

'Store procedure untuk mengisi biaya pelayanan pasien
Private Function sp_BiayaPelayanan(ByVal adoCommand As ADODB.Command, strKdPelayananRS As String, curTarif As Currency, intJmlPel As Integer, dtTanggalPelayanan As Date, strkodedokter As String, strStatusCITO As String, f_TarifCito As Currency, f_KdLabLuar As String) As Boolean
    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fQuery As String
    On Error GoTo errLoad

    sp_BiayaPelayanan = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, frmPasienRujukan.dcRuangan.BoundText)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, strKdPelayananRS)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, mstrKdKelas)
        .Parameters.Append .CreateParameter("StatusCITO", adChar, adParamInput, 1, strStatusCITO)
        .Parameters.Append .CreateParameter("Tarif", adInteger, adParamInput, , curTarif)
        .Parameters.Append .CreateParameter("JmlPelayanan", adInteger, adParamInput, , intJmlPel)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(frmPasienRujukan.dtpTglDirujuk.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("NoLab_Rad", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("StatusAPBD", adChar, adParamInput, 2, strStatusAPBD)
        .Parameters.Append .CreateParameter("KdJenisTarif", adChar, adParamInput, 2, strKdJenisTarif)
        .Parameters.Append .CreateParameter("TarifCito", adInteger, adParamInput, , f_TarifCito)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("IdPegawai2", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("KdLaboratory", adChar, adParamInput, 3, IIf(f_KdLabLuar = "", Null, f_KdLabLuar))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_BiayaPelayananPenunjangForKonRa"  ' Adm Radiologi hanya satu
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Biaya Pelayanan Pasien", vbCritical, "Validasi"
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
            sp_BiayaPelayanan = False
            GoTo errLoad
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing

        Set fRS = Nothing
        fQuery = "update BiayaPelayanan set KdRuanganAsal= '" & mstrKdRuangan & "' where NoPendaftaran='" & mstrNoPen & "'"
        Call msubRecFO(fRS, fQuery)

        Set fRS2 = Nothing
        fQuery = "update DetailBiayaPelayanan set KdRuanganAsal = '" & mstrKdRuangan & "' where NoPendaftaran = '" & mstrNoPen & "'"
        Call msubRecFO(fRS2, fQuery)

    End With
    Exit Function
errLoad:
    sp_BiayaPelayanan = False
End Function

Private Sub Check1_Click()
    On Error GoTo a
    If Check1.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000001' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check1.Value = 0
End Sub

Private Sub Check10_Click()
    On Error GoTo a
    If Check10.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000010' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check10.Value = 0
End Sub

Private Sub Check11_Click()
    On Error GoTo a
    If Check11.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000011' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check11.Value = 0

End Sub

Private Sub Check12_Click()
    On Error GoTo a
    If Check12.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000012' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check12.Value = 0

End Sub

Private Sub Check13_Click()
    On Error GoTo a
    If Check13.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000013' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check13.Value = 0

End Sub

Private Sub Check14_Click()
    On Error GoTo a
    If Check14.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000014' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check14.Value = 0

End Sub

Private Sub Check15_Click()
    On Error GoTo a
    If Check15.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000015' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check15.Value = 0

End Sub

Private Sub Check16_Click()
    On Error GoTo a
    If Check16.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000016' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check16.Value = 0
End Sub

Private Sub Check17_Click()
    On Error GoTo a
    If Check17.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000017' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check17.Value = 0
End Sub

Private Sub Check18_Click()
    On Error GoTo a
    If Check18.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000018' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check18.Value = 0
End Sub

Private Sub Check19_Click()
    On Error GoTo a
    If Check19.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000019' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check19.Value = 0
End Sub

Private Sub Check2_Click()
    On Error GoTo a
    If Check2.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000002' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check2.Value = 0
End Sub

Private Sub Check20_Click()
    On Error GoTo a
    If Check20.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000020' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check20.Value = 0
End Sub

Private Sub Check21_Click()
    On Error GoTo a
    If Check21.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000021' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check21.Value = 0
End Sub

Private Sub Check22_Click()
    On Error GoTo a
    If Check22.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000022' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check22.Value = 0
End Sub

Private Sub Check23_Click()
    On Error GoTo a
    If Check23.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000023' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check23.Value = 0
End Sub

Private Sub Check24_Click()
    On Error GoTo a
    If Check24.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000024' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check24.Value = 0
End Sub

Private Sub Check25_Click()
    On Error GoTo a
    If Check25.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000025' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check25.Value = 0
End Sub

Private Sub Check26_Click()
    On Error GoTo a
    If Check26.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000026' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check26.Value = 0
End Sub

Private Sub Check27_Click()
    On Error GoTo a
    If Check27.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000027' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check27.Value = 0
End Sub

Private Sub Check28_Click()
    On Error GoTo a
    If Check28.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000028' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check28.Value = 0
End Sub

Private Sub Check29_Click()
    On Error GoTo a
    If Check29.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000029' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check29.Value = 0
End Sub

Private Sub Check3_Click()
    On Error GoTo a
    If Check3.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000003' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check3.Value = 0
End Sub

Private Sub Check30_Click()
    On Error GoTo a
    If Check30.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000030' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check30.Value = 0
End Sub

Private Sub Check31_Click()
    On Error GoTo a
    If Check31.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000031' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check31.Value = 0
End Sub

Private Sub Check32_Click()
    On Error GoTo a
    If Check32.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000032' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check32.Value = 0
End Sub

Private Sub Check33_Click()
    On Error GoTo a
    If Check33.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000033' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check33.Value = 0
End Sub

Private Sub Check34_Click()
    On Error GoTo a
    If Check34.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000034' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check34.Value = 0
End Sub

Private Sub Check35_Click()
    On Error GoTo a
    If Check35.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000035' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check35.Value = 0
End Sub

Private Sub Check36_Click()
    On Error GoTo a
    If Check36.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000036' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check36.Value = 0
End Sub

Private Sub Check37_Click()
    On Error GoTo a
    If Check37.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000037' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check37.Value = 0
End Sub

Private Sub Check38_Click()
    On Error GoTo a
    If Check38.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000038' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check38.Value = 0
End Sub

Private Sub Check39_Click()
    On Error GoTo a
    If Check39.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000039' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check39.Value = 0
End Sub

Private Sub Check4_Click()
    On Error GoTo a
    If Check4.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000004' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check4.Value = 0
End Sub

Private Sub Check40_Click()
    On Error GoTo a
    If Check40.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000040' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check40.Value = 0
End Sub

Private Sub Check41_Click()
    On Error GoTo a
    If Check41.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000041' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check41.Value = 0
End Sub

Private Sub Check42_Click()
    On Error GoTo a
    If Check42.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000042' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check42.Value = 0
End Sub

Private Sub Check43_Click()
    On Error GoTo a
    If Check43.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000043' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check43.Value = 0
End Sub

Private Sub Check44_Click()
    On Error GoTo a
    If Check44.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000044' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check44.Value = 0
End Sub

Private Sub Check45_Click()
    On Error GoTo a
    If Check45.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000045' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check45.Value = 0
End Sub

Private Sub Check46_Click()
    On Error GoTo a
    If Check46.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000046' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check46.Value = 0
End Sub

Private Sub Check47_Click()
    On Error GoTo a
    If Check47.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000047' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check47.Value = 0
End Sub

Private Sub Check48_Click()
    On Error GoTo a
    If Check48.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000048' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check48.Value = 0
End Sub

Private Sub Check49_Click()
    On Error GoTo a
    If Check49.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000049' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check49.Value = 0
End Sub

Private Sub Check5_Click()
    On Error GoTo a
    If Check5.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000005' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check5.Value = 0
End Sub

Private Sub Check50_Click()
    On Error GoTo a
    If Check50.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000050' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check50.Value = 0
End Sub

Private Sub Check51_Click()
    On Error GoTo a
    If Check51.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000051' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check51.Value = 0
End Sub

Private Sub Check52_Click()
    On Error GoTo a
    If Check52.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000052' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check52.Value = 0
End Sub

Private Sub Check53_Click()
    On Error GoTo a
    If Check53.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000053' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check53.Value = 0
End Sub

Private Sub Check54_Click()
    On Error GoTo a
    If Check54.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000054' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check54.Value = 0
End Sub

Private Sub Check55_Click()
    On Error GoTo a
    If Check55.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000055' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check55.Value = 0
End Sub

Private Sub Check56_Click()
    On Error GoTo a
    If Check56.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000056' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check56.Value = 0
End Sub

Private Sub Check57_Click()
    On Error GoTo a
    If Check57.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000057' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check57.Value = 0
End Sub

Private Sub Check58_Click()
    On Error GoTo a
    If Check58.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000058' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check58.Value = 0
End Sub

Private Sub Check59_Click()
    On Error GoTo a
    If Check59.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000059' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check59.Value = 0
End Sub

Private Sub Check6_Click()
    On Error GoTo a
    If Check6.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000006' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check6.Value = 0
End Sub

Private Sub Check60_Click()
    On Error GoTo a
    If Check60.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000060' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check60.Value = 0
End Sub

Private Sub Check7_Click()
    On Error GoTo a
    If Check7.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000007' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check7.Value = 0
End Sub

Private Sub Check8_Click()
    On Error GoTo a
    If Check8.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000008' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check8.Value = 0
End Sub

Private Sub Check9_Click()
    On Error GoTo a
    If Check9.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '000009' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    Check9.Value = 0
End Sub

Private Sub chk111001_Click()
    On Error GoTo a
    If chk111001.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111001' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111001.Value = 0
End Sub

Private Sub chk111007_Click()

    MyIndex = 45
    On Error GoTo a

    If chk111007.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111007' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click

    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            MyIndex = fgPelayanan.Rows
            fgPelayanan.TextMatrix(xRow, 1) = a
        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111007.Value = 0
End Sub

Private Sub chk111008_Click()

    MyIndex = 46
    On Error GoTo a
    If chk111008.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111008' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click

    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            MyIndex = fgPelayanan.Rows
            fgPelayanan.TextMatrix(xRow, 1) = a
        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111008.Value = 0
End Sub

Private Sub chk111009_Click()
    Dim i As Integer
    Dim h As Integer

    MyIndex = 5
    On Error GoTo a
    If chk111009.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111009' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            MyIndex = fgPelayanan.Rows
            fgPelayanan.TextMatrix(xRow, 1) = a
        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111009.Value = 0
End Sub

Private Sub chk111010_Click()

    MyIndex = 9
    On Error GoTo a
    If chk111010.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111010' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click

    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            MyIndex = fgPelayanan.Rows
            fgPelayanan.TextMatrix(xRow, 1) = a
        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111010.Value = 0
End Sub

Private Sub chk111011_Click()

    MyIndex = 17
    On Error GoTo a
    If chk111011.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111011' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            MyIndex = fgPelayanan.Rows
            fgPelayanan.TextMatrix(xRow, 1) = a
        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111011.Value = 0
End Sub

Private Sub chk111012_Click()

    MyIndex = 16
    On Error GoTo a
    If chk111012.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111012' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            MyIndex = fgPelayanan.Rows
            fgPelayanan.TextMatrix(xRow, 1) = a
        End With
    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111012.Value = 0
End Sub

Private Sub chk111013_Click()

    MyIndex = 11
    On Error GoTo a
    If chk111013.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111013' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            MyIndex = fgPelayanan.Rows
            fgPelayanan.TextMatrix(xRow, 1) = a
        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111013.Value = 0
End Sub

Private Sub chk111014_Click()

    MyIndex = 12
    On Error GoTo a
    If chk111014.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111014' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click

    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            MyIndex = fgPelayanan.Rows
            fgPelayanan.TextMatrix(xRow, 1) = a
        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111014.Value = 0
End Sub

Private Sub chk111014a_Click()
    On Error GoTo a
    If chk111014a.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111014a' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111014a.Value = 0
End Sub

Private Sub chk111015_Click()
    Dim i As Integer
    Dim h As Integer

    MyIndex = 6
    On Error GoTo a
    If chk111015.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111015' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click

    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111015.Value = 0
End Sub

Private Sub msubRemoveItem2(hgrid As Object, MyIndex As Integer)
    Dim i As Integer
    Dim j As Integer
    With hgrid
        Dim intRowNow As Integer
        intRowNow = MyIndex
        For i = 1 To .Rows - 2
            If i = MyIndex Then
                For j = 0 To .Cols - 1
                    .TextMatrix(MyIndex, j) = .TextMatrix(MyIndex + 1, j)
                Next j
                MyIndex = MyIndex + 1
            End If
        Next i
        .Rows = .Rows - 1
        .Row = intRowNow
    End With
End Sub

Private Sub chk111016_Click()

    MyIndex = 15
    On Error GoTo a
    If chk111016.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111016' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click

    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111016.Value = 0
End Sub

Private Sub chk111017_Click()
    On Error GoTo a
    If chk111017.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111017' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111017.Value = 0
End Sub

Private Sub chk111018_Click()
    On Error GoTo a
    If chk111018.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111018' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111018.Value = 0
End Sub

Private Sub chk111025_Click()

    MyIndex = 4
    On Error GoTo a
    If chk111025.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111025' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click

    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111025.Value = 0
End Sub

Private Sub chk111026_Click()
    MyIndex = 19
    On Error GoTo a
    If chk111026.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111026' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111026.Value = 0
End Sub

Private Sub chk111027_Click()
    MyIndex = 18
    On Error GoTo a
    If chk111027.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111027' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111027.Value = 0

End Sub

Private Sub chk111029_Click()
    MyIndex = 21
    On Error GoTo a
    If chk111029.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111029' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111029.Value = 0
End Sub

Private Sub chk111030_Click()
    MyIndex = 23
    On Error GoTo a
    If chk111030.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111030' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111030.Value = 0
End Sub

Private Sub chk111031_Click()
    MyIndex = 24
    On Error GoTo a
    If chk111031.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111031' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With
    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111031.Value = 0
End Sub

Private Sub chk111032_Click()
    MyIndex = 25
    On Error GoTo a
    If chk111032.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111032' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a
        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111032.Value = 0
End Sub

Private Sub chk111033_Click()

    MyIndex = 62
    On Error GoTo a
    If chk111033.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111033' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else
        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111033.Value = 0
End Sub

Private Sub chk111035_Click()
    MyIndex = 64
    On Error GoTo a
    If chk111035.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111035' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else
        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111035.Value = 0
End Sub

Private Sub chk111036_Click()
    MyIndex = 65
    On Error GoTo a
    If chk111036.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111036' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else
        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111036.Value = 0
End Sub

Private Sub chk111037_Click()
    MyIndex = 66
    On Error GoTo a
    If chk111037.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111037' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a
        End With
    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111037.Value = 0
End Sub

Private Sub chk111042_Click()
    MyIndex = 47
    On Error GoTo a
    If chk111042.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111042' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else
        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111042.Value = 0
End Sub

Private Sub chk111050_Click()
    MyIndex = 67
    On Error GoTo a
    If chk111050.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111050' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else
        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a
        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111050.Value = 0
End Sub

Private Sub chk111051_Click()

    MyIndex = 48
    On Error GoTo a
    If chk111051.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111051' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else
        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a
        End With
    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111051.Value = 0
End Sub

Private Sub chk111052_Click()
    Dim i As Integer
    Dim h As Integer

    MyIndex = 8
    On Error GoTo a
    If chk111052.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = '111052' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click

    Else
        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a
        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chk111052.Value = 0
End Sub

Private Sub chkAbdo_Click()
    On Error GoTo a
    If chkAbdo.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'Abdo' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkAbdo.Value = 0
End Sub

Private Sub chkBNOAP_Click()
    On Error GoTo a
    If chkBNOAP.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'BNOAP' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkBNOAP.Value = 0
End Sub

Private Sub chkBNOaplateral_Click()
    On Error GoTo a
    If chkBNOaplateral.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'BNOaplateral' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkBNOaplateral.Value = 0
End Sub

Private Sub chkBNOLateral_Click()
    On Error GoTo a
    If chkBNOLateral.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'BNOLateral' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkBNOLateral.Value = 0
End Sub

Private Sub chkCervicali_Click()
    On Error GoTo a
    If chkCervicali.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'Cervicali' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkCervicali.Value = 0
End Sub

Private Sub chkCranium_Click()
    On Error GoTo a
    If chkCranium.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'Cranium' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkCranium.Value = 0
End Sub

Private Sub chkEisler_Click()
    On Error GoTo a
    If chkEisler.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'Eisler' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkEisler.Value = 0
End Sub

Private Sub chkExtre_Click()
    On Error GoTo a
    If chkExtre.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'Extre' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkExtre.Value = 0
End Sub

Private Sub chkExtreIn_Click()
    On Error GoTo a
    If chkExtreIn.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'ExtreIn' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkExtreIn.Value = 0
End Sub

Private Sub chkLumbosacral_Click()
    On Error GoTo a
    If chkLumbosacral.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'Lumbosacral' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkLumbosacral.Value = 0
End Sub

Private Sub chkMastoidDextra_Click()
    On Error GoTo a
    If chkMastoidDextra.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'MastoidDextra' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkMastoidDextra.Value = 0
End Sub

Private Sub chkMastoidSinistra_Click()
    On Error GoTo a
    If chkMastoidSinistra.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'MastoidSinistra' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkMastoidSinistra.Value = 0
End Sub

Private Sub chkMaxilla_Click()
    On Error GoTo a
    If chkMaxilla.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'Maxilla' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkMaxilla.Value = 0
End Sub

Private Sub chkOther_Click()
    On Error GoTo a
    If chkOther.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'Other' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkOther.Value = 0
End Sub

Private Sub chkPel_Click()
    On Error GoTo a
    If chkPel.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'Pel' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkPel.Value = 0
End Sub

Private Sub chkPelvisss_Click()
    On Error GoTo a
    If chkPelvisss.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'Pelvisss' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkPelvisss.Value = 0
End Sub

Private Sub chkRespirasi_Click()
    On Error GoTo a
    If chkRespirasi.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'Respirasi' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkRespirasi.Value = 0
End Sub

Private Sub chkSemuaAbdomen_Click()
    On Error GoTo a
    If chkSemuaAbdomen.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'SemuaAbdomen' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkSemuaAbdomen.Value = 0
End Sub

Private Sub chkSemuaCranium_Click()
    On Error GoTo a
    If chkSemuaCranium.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'SemuaCranium' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkSemuaCranium.Value = 0
End Sub

Private Sub chkSemuaExtre_Click()
    On Error GoTo a
    If chkSemuaExtre.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'SemuaExtre' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkSemuaExtre.Value = 0
End Sub

Private Sub chkSemuaInperium_Click()
    On Error GoTo a
    If chkSemuaInperium.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'SemuaInperium' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkSemuaInperium.Value = 0
End Sub

Private Sub chkSemuaOther_Click()
    On Error GoTo a
    If chkSemuaOther.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'SemuaOther' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkSemuaOther.Value = 0
End Sub

Private Sub chkSemuaPelvis_Click()
    On Error GoTo a
    If chkSemuaPelvis.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'SemuaPelvis' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkSemuaPelvis.Value = 0
End Sub

Private Sub chkSemuaRespirasi_Click()
    On Error GoTo a
    If chkSemuaRespirasi.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'SemuaRespirasi' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkSemuaRespirasi.Value = 0
End Sub

Private Sub chkSemuaVertebra_Click()
    On Error GoTo a
    If chkSemuaVertebra.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'SemuaVertebra' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkSemuaVertebra.Value = 0
End Sub

Private Sub chkStenvers_Click()
    On Error GoTo a
    If chkStenvers.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'Stenvers' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkStenvers.Value = 0
End Sub

Private Sub chkThoracalis_Click()
    On Error GoTo a
    If chkThoracalis.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'Thoracalis' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkThoracalis.Value = 0
End Sub

Private Sub chkTMJDextra_Click()
    On Error GoTo a
    If chkTMJDextra.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'TMJDextra' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkTMJDextra.Value = 0
End Sub

Private Sub chkTMJSinistra_Click()
    On Error GoTo a
    If chkTMJSinistra.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'TMJSinistra' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkTMJSinistra.Value = 0
End Sub

Private Sub chkVertebra_Click()
    On Error GoTo a
    If chkVertebra.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'Vertebra' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkVertebra.Value = 0
End Sub

Private Sub chkVertThoAP_Click()
    On Error GoTo a
    If chkVertThoAP.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'VertThoAP' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkVertThoAP.Value = 0
End Sub

Private Sub chkVertThoracolumbal_Click()
    On Error GoTo a
    If chkVertThoracolumbal.Value = 1 Then
        strSQL = "select KdPelayananRS,[Nama Pelayanan],KdKelas,Tarif from V_TarifPelayananTindakan where KdPelayananRS = 'VertThoracolumbal' and KdKelas = '" & mstrKdKelas & "' AND  KdRuangan = '" & frmPasienRujukan.dcRuangan.BoundText & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        strKodePelayananRS = rs.Fields(0).Value
        strNamaPelayananRS = rs.Fields(1).Value
        curBiaya = rs.Fields(3).Value

        Call Command1_Click
    Else

        Call Sem

        With fgPelayanan
            a = "Bangkit"
            fgPelayanan.TextMatrix(xRow, 1) = a

        End With

    End If
a:
    MsgBox "Maaf data tarif pelayanan tindakan tidak ada", vbCritical, "Validasi"
    chkVertThoracolumbal.Value = 0
End Sub

Private Sub cmdSimpan_Click()
    Dim adoCommand As New ADODB.Command

    'Input ke Tabel Sementara

    For i = 1 To fgPelayanan.Rows - 2
        If fgPelayanan.TextMatrix(i, 1) <> "Bangkit" Then

            With adoCommand
                .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
                .Parameters.Append .CreateParameter("KdRuanganReq", adChar, adParamInput, 3, mstrKdRuangan)
                .Parameters.Append .CreateParameter("KdPelRSReq", adChar, adParamInput, 6, fgPelayanan.TextMatrix(i, 0))
                .Parameters.Append .CreateParameter("StatusCito", adChar, adParamInput, 1, IIf(optCito(0).Value = True, "Y", "T"))
                .Parameters.Append .CreateParameter("JmlPelReq", adInteger, adParamInput, , txtJmlPelayanan.Text)
                .Parameters.Append .CreateParameter("TglRequest", adDate, adParamInput, , Format(frmPasienRujukan.dtpTglDirujuk.Value, "yyyy/MM/dd HH:mm:ss"))
                .Parameters.Append .CreateParameter("NoLab_Rad", adChar, adParamInput, 10, Null)
                .Parameters.Append .CreateParameter("IdPegawaiReq", adChar, adParamInput, 10, mstrKdDokter)
                .Parameters.Append .CreateParameter("IdUserReq", adChar, adParamInput, 10, strIDPegawaiAktif)
                .Parameters.Append .CreateParameter("KeteranganReq", adVarChar, adParamInput, 50, txtKesimpulan2.Text)
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, Null)

                .ActiveConnection = dbConn
                .CommandText = "dbo.Add_DetailRequestRadiologi"
                .CommandType = adCmdStoredProc
                .Execute
                If Not (.Parameters("RETURN_VALUE").Value = 0) Then
                    MsgBox "Ada Kesalahan dalam Penyimpanan Data", vbCritical, "Validasi"
                Else
                    Call Add_HistoryLoginActivity("Add_DetailRequestRadiologi")
                End If
                Call deleteADOCommandParameters(adoCommand)
                Set adoCommand = Nothing
            End With
            fgPelayanan.Row = fgPelayanan.Row + 1
        End If
    Next i
    On Error GoTo a
    Call SimpanKonsul

    For i = 1 To fgPelayanan.Rows - 2
        If fgPelayanan.TextMatrix(i, 1) <> "Bangkit" Then
            If sp_BiayaPelayanan(dbcmd, fgPelayanan.TextMatrix(i, 0), CCur(fgPelayanan.TextMatrix(i, 3)), fgPelayanan.TextMatrix(i, 2), fgPelayanan.TextMatrix(i, 9), fgPelayanan.TextMatrix(i, 6), fgPelayanan.TextMatrix(i, 7), CCur(fgPelayanan.TextMatrix(i, 8)), fgPelayanan.TextMatrix(i, 10)) = False Then Exit Sub
        End If
    Next i
    MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
    cmdSimpan.Enabled = False
a:
    MsgBox "Maaf data sudah ada", vbCritical, "Validasi"
End Sub

Private Sub SimpanKonsul()

    Dim adoCommand As New ADODB.Command

    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, frmPasienRujukan.txtnocm.Text)
        .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, frmPasienRujukan.dcRuangan.BoundText)
        .Parameters.Append .CreateParameter("IdDokterPerujuk", adChar, adParamInput, 10, mstrKdDokter)
        .Parameters.Append .CreateParameter("TglDirujuk", adDate, adParamInput, , Format(frmPasienRujukan.dtpTglDirujuk.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PasienRujukan"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Data", vbCritical, "Validasi"
        Else

        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With

End Sub

Private Sub Command1_Click()

    With fgPelayanan

        intRowNow = .Rows - 1
        .TextMatrix(intRowNow, 0) = strKodePelayananRS
        .TextMatrix(intRowNow, 1) = strNamaPelayananRS
        .TextMatrix(intRowNow, 2) = CInt(txtJmlPelayanan.Text)

        subcurTarifCito = sp_Take_TarifBPT
        .TextMatrix(intRowNow, 3) = IIf(subcurTarifBiayaSatuan = 0, 0, Format(subcurTarifBiayaSatuan, "#,###")) 'curBiaya
        .TextMatrix(intRowNow, 4) = IIf(Format(funcRoundUp(CStr(subcurTarifBiayaSatuan + subcurTarifCito)) * CInt(txtJmlPelayanan.Text), "#,###") = 0, 0, Format(funcRoundUp(CStr(subcurTarifBiayaSatuan + subcurTarifCito)) * CInt(txtJmlPelayanan.Text), "#,###"))
        .TextMatrix(intRowNow, 8) = subcurTarifCito

        .TextMatrix(intRowNow, 5) = mdTglBerlaku

        .TextMatrix(intRowNow, 7) = strCito
        .TextMatrix(intRowNow, 9) = Now
        .TextMatrix(intRowNow, 12) = MyIndex

        .Rows = .Rows + 1
        .SetFocus
    End With

End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subSetGidPelayanan
    txtJmlPelayanan.Text = 1
    strCito = "0"
    strStatusAPBD = "01"

    Set rs = Nothing

    strSQL = "SELECT KdJenisTarif,JenisTarif " _
    & "FROM v_JenisTarifPasien " _
    & "WHERE NoPendaftaran='" & mstrNoPen & "'"
    Set rs = Nothing

    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockOptimistic
    strKdJenisTarif = rs.Fields(0).Value
    strJenisTarif = rs.Fields(1).Value
    Set rs = Nothing

End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub subSetGidPelayanan()

    With fgPelayanan
        .Clear
        .Rows = 2
        .Cols = 13
        .TextMatrix(0, 0) = "Kode Pelayanan"
        .TextMatrix(0, 1) = "Nama Pelayanan"
        .TextMatrix(0, 2) = "Jumlah"
        .TextMatrix(0, 3) = "Biaya Satuan"
        .TextMatrix(0, 4) = "Biaya Total"
        .TextMatrix(0, 5) = "Tgl Berlaku"
        .TextMatrix(0, 6) = "Kode Dokter"
        .TextMatrix(0, 7) = "Status CITO"
        .TextMatrix(0, 8) = "Biaya CITO"
        .TextMatrix(0, 9) = "Tanggal Pelayanan"
        .TextMatrix(0, 10) = "KodeLabLuar"
        .TextMatrix(0, 11) = "DokterDelegasi"
        .TextMatrix(0, 12) = "Index"

        .ColWidth(0) = 0
        .ColWidth(1) = 4700
        .ColWidth(2) = 1500
        .ColWidth(3) = 1500
        .ColWidth(4) = 1500
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 1500
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 1000
    End With
End Sub

Private Function sp_Take_TarifBPT() As Currency
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, strKodePelayananRS)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, mstrKdKelas)
        .Parameters.Append .CreateParameter("KdJenisTarif", adChar, adParamInput, 2, strKdJenisTarif)
        .Parameters.Append .CreateParameter("TarifCito", adCurrency, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("TarifTotal", adCurrency, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("StatusCito", adChar, adParamInput, 1, IIf(optCito(0).Value = True, "Y", "T"))
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, Null) 'IdDokter sengaja dibuat kosong
        .Parameters.Append .CreateParameter("IdDokter2", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("IdDokter3", adChar, adParamInput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Take_TarifBPT"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam Pengambilan biaya tarif", vbExclamation, "Validasi"
            sp_Take_TarifBPT = 0
            subcurTarifBiayaSatuan = 0
        Else
            sp_Take_TarifBPT = .Parameters("TarifCito").Value
            subcurTarifBiayaSatuan = .Parameters("TarifTotal").Value
            Call Add_HistoryLoginActivity("Take_TarifBPT")
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Sub optCito_Click(Index As Integer)
    If Index = 0 Then
        strCito = "1"
    Else
        strCito = "0"
    End If
End Sub

Private Sub Sem()
    Dim i As Integer

    For i = 1 To fgPelayanan.Rows - 2
        If fgPelayanan.TextMatrix(i, 12) = MyIndex Then
            xRow = i
            Exit Sub
        End If
    Next i
End Sub

