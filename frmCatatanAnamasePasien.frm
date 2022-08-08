VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCatatanAnamasePasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Catatan Anamnesa Mata Pasien"
   ClientHeight    =   10365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10950
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCatatanAnamasePasien.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   10950
   Begin VB.Frame fraDokter 
      Caption         =   "Data Pemeriksa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   960
      TabIndex        =   43
      Top             =   10320
      Visible         =   0   'False
      Width           =   9135
      Begin MSDataGridLib.DataGrid dgDokter 
         Height          =   2295
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   8535
         _ExtentX        =   15055
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
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   40
      Top             =   9480
      Width           =   10935
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   465
         Left            =   6840
         TabIndex        =   27
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   465
         Left            =   8880
         TabIndex        =   28
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Anamnesa Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   0
      TabIndex        =   30
      Top             =   1920
      Width           =   10935
      Begin VB.Frame Frame4 
         Caption         =   "Oculus Dexter (OD)"
         Height          =   2055
         Left            =   120
         TabIndex        =   66
         Top             =   960
         Width           =   10695
         Begin VB.TextBox txtDerajat 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   6480
            MaxLength       =   150
            TabIndex        =   72
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtCyl 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   5160
            MaxLength       =   150
            TabIndex        =   71
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtSPH 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3720
            MaxLength       =   150
            TabIndex        =   70
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtPlacido 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   69
            Top             =   840
            Width           =   7575
         End
         Begin VB.TextBox txtJaval 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   68
            Top             =   1200
            Width           =   7575
         End
         Begin VB.TextBox txtScias 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   67
            Top             =   1560
            Width           =   7575
         End
         Begin MSDataListLib.DataCombo dcVisusMasuk 
            Height          =   330
            Left            =   1320
            TabIndex        =   73
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Appearance      =   0
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo dcVisusKeluar 
            Height          =   330
            Left            =   9360
            TabIndex        =   74
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Appearance      =   0
            Text            =   "DataCombo1"
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "derajat"
            Height          =   210
            Index           =   18
            Left            =   7440
            TabIndex        =   83
            Top             =   360
            Width           =   570
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Visus Masuk"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   15
            Left            =   210
            TabIndex        =   82
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "as"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   13
            Left            =   6210
            TabIndex        =   81
            Top             =   360
            Width           =   195
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Cyl"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   12
            Left            =   4785
            TabIndex        =   80
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Placido"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   210
            TabIndex        =   79
            Top             =   840
            Width           =   600
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "dengan sph"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   2640
            TabIndex        =   78
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Visus Keluar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   16
            Left            =   8280
            TabIndex        =   77
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Javal"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   210
            TabIndex        =   76
            Top             =   1200
            Width           =   450
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Sciascopia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   210
            TabIndex        =   75
            Top             =   1560
            Width           =   885
         End
      End
      Begin VB.TextBox txtGeneralis 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   2520
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   6840
         Width           =   8175
      End
      Begin VB.TextBox txtOcularis 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   2520
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   6120
         Width           =   8175
      End
      Begin VB.TextBox txtOS 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4560
         MaxLength       =   150
         TabIndex        =   24
         Top             =   5640
         Width           =   855
      End
      Begin VB.TextBox txtOD 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3240
         MaxLength       =   150
         TabIndex        =   23
         Top             =   5640
         Width           =   855
      End
      Begin VB.TextBox txtPup 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7860
         MaxLength       =   150
         TabIndex        =   22
         Top             =   5280
         Width           =   615
      End
      Begin VB.TextBox txtDist 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7080
         MaxLength       =   150
         TabIndex        =   21
         Top             =   5280
         Width           =   615
      End
      Begin VB.TextBox txtSamadengan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4560
         MaxLength       =   150
         TabIndex        =   20
         Top             =   5280
         Width           =   855
      End
      Begin VB.TextBox txtDengan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3240
         MaxLength       =   150
         TabIndex        =   19
         Top             =   5280
         Width           =   855
      End
      Begin VB.TextBox txtOOvisus 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1440
         MaxLength       =   150
         TabIndex        =   18
         Top             =   5280
         Width           =   855
      End
      Begin VB.Frame Frame6 
         Caption         =   "Oculus Sinister (OS)"
         Height          =   2055
         Left            =   120
         TabIndex        =   46
         Top             =   3120
         Width           =   10695
         Begin VB.TextBox txtAs 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   6480
            MaxLength       =   150
            TabIndex        =   13
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtCylOs 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   5160
            MaxLength       =   150
            TabIndex        =   12
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtSPHOS 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3720
            MaxLength       =   150
            TabIndex        =   11
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtPlacidoOs 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   15
            Top             =   840
            Width           =   7575
         End
         Begin VB.TextBox txtJavalOs 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   16
            Top             =   1200
            Width           =   7575
         End
         Begin VB.TextBox txtSciasOs 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   17
            Top             =   1560
            Width           =   7575
         End
         Begin MSDataListLib.DataCombo dcVisusMasukOS 
            Height          =   330
            Left            =   1320
            TabIndex        =   10
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Appearance      =   0
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo dcVisusKeluarOs 
            Height          =   330
            Left            =   9360
            TabIndex        =   14
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Appearance      =   0
            Text            =   "DataCombo1"
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "derajat"
            Height          =   210
            Index           =   27
            Left            =   7440
            TabIndex        =   55
            Top             =   360
            Width           =   570
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Visus Masuk"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   26
            Left            =   210
            TabIndex        =   54
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "as"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   25
            Left            =   6210
            TabIndex        =   53
            Top             =   360
            Width           =   195
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Cyl"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   24
            Left            =   4785
            TabIndex        =   52
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Placido"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   23
            Left            =   210
            TabIndex        =   51
            Top             =   840
            Width           =   600
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "dengan sph"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   21
            Left            =   2640
            TabIndex        =   50
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Visus Keluar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   20
            Left            =   8280
            TabIndex        =   49
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Javal"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   19
            Left            =   210
            TabIndex        =   48
            Top             =   1200
            Width           =   450
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Sciascopia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   17
            Left            =   210
            TabIndex        =   47
            Top             =   1560
            Width           =   885
         End
      End
      Begin VB.TextBox txtPemeriksa 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2280
         MaxLength       =   150
         TabIndex        =   8
         Top             =   600
         Width           =   3735
      End
      Begin MSComCtl2.DTPicker dtpTglPeriksa 
         Height          =   330
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   125829123
         UpDown          =   -1  'True
         CurrentDate     =   38076
      End
      Begin MSDataListLib.DataCombo dcPerawat 
         Height          =   330
         Left            =   6120
         TabIndex        =   9
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Historia morbi generalis"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   37
         Left            =   360
         TabIndex        =   65
         Top             =   6840
         Width           =   2025
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Historia morbi ocularis"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   36
         Left            =   360
         TabIndex        =   64
         Top             =   6120
         Width           =   1905
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Tambah buat dekat    :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   35
         Left            =   360
         TabIndex        =   63
         Top             =   5640
         Width           =   1890
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "OS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   34
         Left            =   4200
         TabIndex        =   62
         Top             =   5640
         Width           =   225
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "OD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   33
         Left            =   2880
         TabIndex        =   61
         Top             =   5640
         Width           =   240
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   32
         Left            =   7730
         TabIndex        =   60
         Top             =   5280
         Width           =   90
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Dist.pup"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   31
         Left            =   6240
         TabIndex        =   59
         Top             =   5280
         Width           =   690
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   30
         Left            =   4200
         TabIndex        =   58
         Top             =   5280
         Width           =   135
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "dengan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   29
         Left            =   2520
         TabIndex        =   57
         Top             =   5280
         Width           =   630
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "OO. Visus ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   28
         Left            =   360
         TabIndex        =   56
         Top             =   5280
         Width           =   945
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Paramedis Pemeriksa"
         Height          =   210
         Index           =   22
         Left            =   6120
         TabIndex        =   44
         Top             =   360
         Width           =   1680
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Dokter/Perawat Pemeriksa"
         Height          =   210
         Index           =   11
         Left            =   2280
         TabIndex        =   42
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Periksa"
         Height          =   210
         Index           =   10
         Left            =   240
         TabIndex        =   41
         Top             =   360
         Width           =   1260
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   975
      Left            =   0
      TabIndex        =   31
      Top             =   960
      Width           =   10935
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   6840
         MaxLength       =   9
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3480
         TabIndex        =   2
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         MaxLength       =   10
         TabIndex        =   0
         Top             =   480
         Width           =   1455
      End
      Begin VB.Frame Frame5 
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
         Height          =   580
         Left            =   8280
         TabIndex        =   32
         Top             =   240
         Width           =   2415
         Begin VB.TextBox txtThn 
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   4
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtBln 
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   900
            MaxLength       =   6
            TabIndex        =   5
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtHari 
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   1680
            MaxLength       =   6
            TabIndex        =   6
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            Height          =   210
            Index           =   4
            Left            =   550
            TabIndex        =   35
            Top             =   277
            Width           =   285
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            Height          =   210
            Index           =   5
            Left            =   1350
            TabIndex        =   34
            Top             =   277
            Width           =   240
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            Height          =   210
            Index           =   6
            Left            =   2130
            TabIndex        =   33
            Top             =   270
            Width           =   165
         End
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Index           =   3
         Left            =   6840
         TabIndex        =   39
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Index           =   2
         Left            =   3480
         TabIndex        =   38
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Index           =   1
         Left            =   1800
         TabIndex        =   37
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   1335
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
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmCatatanAnamasePasien.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   9120
      Picture         =   "frmCatatanAnamasePasien.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmCatatanAnamasePasien.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmCatatanAnamasePasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFilterDokter As String

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad

    If Periksa("text", txtPemeriksa, "Nama pemeriksa kosong") = False Then Exit Sub
    If Len(Trim(dcPerawat.Text)) > 0 Then If Periksa("datacombo", dcPerawat, "Nama perawat kosong") = False Then Exit Sub
    If Periksa("datacombo", dcVisusMasuk, "OD Visus Masuk kosong") = False Then Exit Sub
    If Periksa("datacombo", dcVisusKeluar, "OD Visus Keluar kosong") = False Then Exit Sub
    If Periksa("datacombo", dcVisusMasukOS, "OS Visus Masuk kosong") = False Then Exit Sub
    If Periksa("datacombo", dcVisusKeluarOs, "OS Visus Keluar kosong") = False Then Exit Sub

    If mstrKdDokter = "" Then
        MsgBox "Pilih dulu Pemeriksa yang akan menangani Pasien", vbExclamation, "Validasi"
        txtPemeriksa.SetFocus
        Exit Sub
    End If

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtnocm.Text)
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dtpTglPeriksa.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("IdPegawai", adVarChar, adParamInput, 10, mstrKdDokter)

        .Parameters.Append .CreateParameter("OOVisus", adDouble, adParamInput, , IIf(Len(Trim(txtOOvisus.Text)) = 0, Null, Trim(txtOOvisus.Text)))
        .Parameters.Append .CreateParameter("OOsph", adDouble, adParamInput, , IIf(Len(Trim(txtDengan.Text)) = 0, Null, Trim(txtDengan.Text)))
        .Parameters.Append .CreateParameter("OOCyl", adDouble, adParamInput, , IIf(Len(Trim(txtSamadengan.Text)) = 0, Null, Trim(txtSamadengan.Text)))
        .Parameters.Append .CreateParameter("OODist", adDouble, adParamInput, , IIf(Len(Trim(txtDist.Text)) = 0, Null, Trim(txtDist.Text)))
        .Parameters.Append .CreateParameter("OOPup", adDouble, adParamInput, , IIf(Len(Trim(txtPup.Text)) = 0, Null, Trim(txtPup.Text)))
        .Parameters.Append .CreateParameter("OOOD", adDouble, adParamInput, , IIf(Len(Trim(txtOD.Text)) = 0, Null, Trim(txtOD.Text)))
        .Parameters.Append .CreateParameter("OOOS", adDouble, adParamInput, , IIf(Len(Trim(txtOS.Text)) = 0, Null, Trim(txtOS.Text)))
        .Parameters.Append .CreateParameter("HistoryMO", adVarChar, adParamInput, 200, IIf(Len(Trim(txtOcularis.Text)) = 0, Null, Trim(txtOcularis.Text)))
        .Parameters.Append .CreateParameter("HistoryMG", adVarChar, adParamInput, 200, IIf(Len(Trim(txtGeneralis.Text)) = 0, Null, Trim(txtGeneralis.Text)))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("IdParamedis", adChar, adParamInput, 10, IIf(dcPerawat.BoundText = "", Null, dcPerawat.BoundText))

        .Parameters.Append .CreateParameter("ODKdVisusMasuk", adChar, adParamInput, 2, IIf(dcVisusMasuk.BoundText = "", Null, dcVisusMasuk.BoundText))
        .Parameters.Append .CreateParameter("ODsph", adDouble, adParamInput, , IIf(Len(Trim(txtSPH.Text)) = 0, Null, Trim(txtSPH.Text)))
        .Parameters.Append .CreateParameter("ODCyl", adDouble, adParamInput, , IIf(Len(Trim(txtCyl.Text)) = 0, Null, Trim(txtCyl.Text)))
        .Parameters.Append .CreateParameter("ODAs", adDouble, adParamInput, , IIf(Len(Trim(txtDerajat.Text)) = 0, Null, Trim(txtDerajat.Text)))
        .Parameters.Append .CreateParameter("ODKdVisusKeluar", adChar, adParamInput, 2, IIf(dcVisusKeluar.BoundText = "", Null, dcVisusKeluar.BoundText))
        .Parameters.Append .CreateParameter("ODPlacido", adVarChar, adParamInput, 10, IIf(txtPlacido.Text = "", Null, txtPlacido.Text))
        .Parameters.Append .CreateParameter("ODJaval", adVarChar, adParamInput, 10, IIf(txtJaval.Text = "", Null, txtJaval.Text))
        .Parameters.Append .CreateParameter("ODSciascopia", adVarChar, adParamInput, 10, IIf(txtScias.Text = "", Null, txtScias.Text))
        .Parameters.Append .CreateParameter("OSKdVisusMasuk", adChar, adParamInput, 2, IIf(dcVisusMasukOS.BoundText = "", Null, dcVisusMasukOS.BoundText))
        .Parameters.Append .CreateParameter("OSsph", adDouble, adParamInput, , IIf(Len(Trim(txtSPHOS.Text)) = 0, Null, Trim(txtSPHOS.Text)))
        .Parameters.Append .CreateParameter("OSCyl", adDouble, adParamInput, , IIf(Len(Trim(txtCylOs.Text)) = 0, Null, Trim(txtCylOs.Text)))
        .Parameters.Append .CreateParameter("OSAs", adDouble, adParamInput, , IIf(Len(Trim(txtAs.Text)) = 0, Null, Trim(txtAs.Text)))
        .Parameters.Append .CreateParameter("OSKdVisusKeluar", adChar, adParamInput, 2, IIf(dcVisusKeluarOs.BoundText = "", Null, dcVisusKeluarOs.BoundText))
        .Parameters.Append .CreateParameter("OSPlacido", adVarChar, adParamInput, 10, IIf(txtPlacidoOs.Text = "", Null, txtPlacidoOs.Text))
        .Parameters.Append .CreateParameter("OSJaval", adVarChar, adParamInput, 10, IIf(txtJavalOs.Text = "", Null, txtJavalOs.Text))
        .Parameters.Append .CreateParameter("OSSciascopia", adVarChar, adParamInput, 10, IIf(txtSciasOs.Text = "", Null, txtSciasOs.Text))

        .ActiveConnection = dbConn
        .CommandText = "dbo.AU_CatatanAnamnesaMata"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
        Else
            MsgBox "Data Anamnesa Mata berhasil disimpan", vbInformation, "Informasi"
            Call Add_HistoryLoginActivity("AU_CatatanAnamnesaMata")
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    cmdSimpan.Enabled = False
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    If cmdSimpan.Enabled = True Then
        If MsgBox("Simpan catatan Anamnesa pasien ", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub dcPerawat_LostFocus()
    If dcPerawat.MatchedWithList = False Then dcPerawat.BoundText = "": dcPerawat.Text = ""
End Sub

Private Sub dcVisusKeluar_KeyPress(KeyAscii As Integer)
 If KeyAscii = 39 Then KeyAscii = 0
 If KeyAscii = 13 Then
        If dcVisusKeluar.MatchedWithList = True Then txtPlacido.SetFocus
        strSQL = " SELECT KdVisus, NamaVisus" & _
                 " From VisusMata" & _
                 " Where (NamaVisus LIKE '%" & dcVisusKeluar.Text & "%')ORDER BY NamaVisus"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
        dcVisusKeluar.Text = ""
        txtPlacido.SetFocus
        Exit Sub
        End If
        dcVisusKeluar.BoundText = rs(0).Value
        dcVisusKeluar.Text = rs(1).Value
    End If
End Sub

Private Sub dcVisusKeluar_LostFocus()
    If dcVisusKeluar.MatchedWithList = False Then dcVisusKeluar.Text = ""

End Sub

Private Sub dcVisusKeluarOs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcVisusKeluarOs.MatchedWithList = True Then txtPlacidoOs.SetFocus
        strSQL = " SELECT KdVisus, NamaVisus" & _
                 " From VisusMata" & _
                 " Where (NamaVisus LIKE '%" & dcVisusKeluarOs.Text & "%')ORDER BY NamaVisus"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
        dcVisusKeluarOs.Text = ""
        txtPlacidoOs.SetFocus
        Exit Sub
        End If
        dcVisusKeluarOs.BoundText = rs(0).Value
        dcVisusKeluarOs.Text = rs(1).Value
    End If
End Sub

Private Sub dcVisusKeluarOs_LostFocus()
    If dcVisusKeluarOs.MatchedWithList = False Then dcVisusKeluarOs.Text = ""
End Sub

Private Sub dcVisusMasuk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcVisusMasuk.MatchedWithList = True Then txtSPH.SetFocus
        strSQL = " SELECT KdVisus, NamaVisus" & _
                 " From VisusMata" & _
                 " Where (NamaVisus LIKE '%" & dcVisusMasuk.Text & "%')ORDER BY NamaVisus"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
        dcVisusMasuk.Text = ""
        txtSPH.SetFocus
        Exit Sub
        End If
        dcVisusMasuk.BoundText = rs(0).Value
        dcVisusMasuk.Text = rs(1).Value
    End If
End Sub

Private Sub dcPerawat_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcPerawat.Text)) > 0 Then
            strSQL = "SELECT  IdPegawai, [Nama Pemeriksa]" & _
                " From V_DaftarPemeriksaPasien" & _
                " WHERE ([Nama Pemeriksa] LIKE '%" & dcPerawat.Text & "%')and StatusEnabled = '1' "
            Call msubRecFO(rs, strSQL)
            dcPerawat.Text = ""
            If rs.EOF = False Then dcPerawat.BoundText = rs(0).Value: txtSPH.SetFocus
        Else
            txtSPH.SetFocus
        End If
    End If
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcVisusMasuk_LostFocus()
    If dcVisusMasuk.MatchedWithList = False Then dcVisusMasuk.Text = ""

End Sub

Private Sub dcVisusMasukOS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcVisusMasukOS.MatchedWithList = True Then txtSPHOS.SetFocus
        strSQL = " SELECT KdVisus, NamaVisus" & _
                 " From VisusMata" & _
                 " Where (NamaVisus LIKE '%" & dcVisusMasukOS.Text & "%')ORDER BY NamaVisus"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
        dcVisusMasukOS.Text = ""
        txtSPHOS.SetFocus
        Exit Sub
        End If
        dcVisusMasukOS.BoundText = rs(0).Value
        dcVisusMasukOS.Text = rs(1).Value
    End If
End Sub

Private Sub dcVisusMasukOS_LostFocus()
    If dcVisusMasukOS.MatchedWithList = False Then dcVisusMasukOS.Text = ""

End Sub

Private Sub dgDokter_DblClick()
    Call dgDokter_KeyPress(13)
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dgDokter.ApproxCount = 0 Then Exit Sub
        txtPemeriksa.Text = dgDokter.Columns(1).Value
        mstrKdDokter = dgDokter.Columns(0).Value
        If mstrKdDokter = "" Then
            MsgBox "Pilih dulu Pemeriksa yang akan menangani Pasien", vbCritical, "Validasi"
            txtPemeriksa.Text = ""
            dgDokter.SetFocus
            Exit Sub
        End If
        fraDokter.Visible = False
        dcPerawat.SetFocus
    End If
    If KeyAscii = 27 Then
        fraDokter.Visible = False
    End If
End Sub

Private Sub dgDokter_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If dgDokter.Visible = False Then Exit Sub
        txtPemeriksa.SetFocus
    End If
End Sub

Private Sub dtpTglPeriksa_Change()
    dtpTglPeriksa.MaxDate = Now
End Sub

Private Sub txtAs_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then dcVisusKeluarOs.SetFocus
End Sub

Private Sub txtCylOs_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 44 Then
        Call SetKeyPressToNumber(KeyAscii)
    Else
        If cekKoma(txtCylOs, KeyAscii) = False Then
            KeyAscii = 44
         Else
            KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 Then txtAs.SetFocus
End Sub

Private Sub txtDengan_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 44 Then
        Call SetKeyPressToNumber(KeyAscii)
    Else
        If cekKoma(txtDengan, KeyAscii) = False Then
            KeyAscii = 44
        Else
            KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 Then txtSamadengan.SetFocus
End Sub

Private Sub txtDist_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtPup.SetFocus
End Sub

Private Sub txtGeneralis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtJaval_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtScias.SetFocus
End Sub

Private Sub txtJavalOs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtSciasOs.SetFocus
End Sub

Private Sub txtOcularis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtGeneralis.SetFocus
End Sub

Private Sub txtOD_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtOS.SetFocus
End Sub

Private Sub txtOOvisus_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtDengan.SetFocus
End Sub

Private Sub txtOS_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtOcularis.SetFocus
End Sub

Private Sub txtPlacido_KeyPress(KeyAscii As Integer)
'Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtJaval.SetFocus
End Sub


Private Sub txtNadi_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then txtDerajat.SetFocus
 'Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtPemeriksa_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errLoad

    Select Case KeyCode
        Case 13
            If fraDokter.Visible = True Then
                dgDokter.SetFocus
            Else
                dcPerawat.SetFocus
            End If
        Case vbKeyEscape
            fraDokter.Visible = False
    End Select
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dtpTglPeriksa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtPemeriksa.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error GoTo errLoad

    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    
    
    strSQL = "Select * from V_RiwayatCatatanAnamnesaMata Where NoPendaftaran='" & frmDaftarPasienRJ.dgDaftarPasienRJ.Columns("No. Registrasi").Value & "' AND NoCM ='" & frmDaftarPasienRJ.dgDaftarPasienRJ.Columns("NoCM").Value & "' "
    Call msubRecFO(rsB, strSQL)
    
    If rsB.EOF = True Then
        dtpTglPeriksa.Value = Now
        txtPemeriksa.Text = ""
        dcPerawat.BoundText = ""
        dcVisusMasuk.BoundText = ""
        dcVisusKeluar.BoundText = ""
        txtSPH.Text = ""
        txtCyl.Text = ""
        txtDerajat.Text = ""
        txtJaval.Text = ""
        txtPlacido.Text = ""
        txtScias.Text = ""
        
        dcVisusMasukOS.BoundText = ""
        dcVisusKeluarOs.BoundText = ""
        txtSPHOS.Text = ""
        txtCylOs.Text = ""
        txtAs.Text = ""
        txtJavalOs.Text = ""
        txtPlacidoOs.Text = ""
        txtSciasOs.Text = ""
        txtOOvisus.Text = ""
        txtDengan.Text = ""
        txtSamadengan.Text = ""
        txtDist.Text = ""
        txtPup.Text = ""
        txtOD.Text = ""
        txtOS.Text = ""
        txtOcularis.Text = ""
        txtGeneralis.Text = ""
        Call subLoadDcSource
        
'        strSQL = "SELECT  IdPegawai, [Nama Pemeriksa]" & _
'                 " FROM V_DaftarPemeriksaPasien " & _
'                 " WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
'        Call msubRecFO(dbRst, strSQL)
'        If rs.EOF = False Then
'            txtPemeriksa.Text = dbRst(1).Value
'            mstrKdDokter = dbRst(0).Value
'        Else
'            mstrKdDokter = ""
'            txtPemeriksa.Text = ""
'        End If
        
    
    Else
        Call subLoadDcSource
        dtpTglPeriksa.Value = rsB.Fields("TglPeriksa").Value
        txtPemeriksa.Text = rsB.Fields("Dokter Pemeriksa").Value
        mstrKdDokter = rsB.Fields("iddokter").Value
        If rsB.Fields("IdParamedis").Value = Null Then dcPerawat.BoundText = "" Else dcPerawat.BoundText = rsB.Fields("IdParamedis").Value
        dcVisusMasuk.BoundText = rsB.Fields("ODKdVisusMasuk").Value
        dcVisusKeluar.BoundText = rsB.Fields("ODKdVisusKeluar").Value
        If IsNull(rsB.Fields("ODsph")) Then txtSPH.Text = "" Else txtSPH.Text = rsB.Fields("ODsph").Value
        If IsNull(rsB.Fields("ODcyl")) Then txtCyl.Text = "" Else txtCyl.Text = rsB.Fields("ODcyl").Value
        If IsNull(rsB.Fields("ODas")) Then txtDerajat.Text = "" Else txtDerajat.Text = rsB.Fields("ODas").Value
        If IsNull(rsB.Fields("ODJaval")) Then txtJaval.Text = "" Else txtJaval.Text = rsB.Fields("ODJaval").Value
        If IsNull(rsB.Fields("ODPlacido")) Then txtPlacido.Text = "" Else txtPlacido.Text = rsB.Fields("ODPlacido").Value
        If IsNull(rsB.Fields("ODSciascopia")) Then txtScias.Text = "" Else txtScias.Text = rsB.Fields("ODSciascopia").Value
        
        dcVisusMasukOS.BoundText = rsB.Fields("OSKdVisusMasuk").Value
        dcVisusKeluarOs.BoundText = rsB.Fields("OSKdVisusKeluar").Value
        If IsNull(rsB.Fields("OSsph")) Then txtSPHOS.Text = "" Else txtSPHOS.Text = rsB.Fields("OSsph").Value
        If IsNull(rsB.Fields("OScyl")) Then txtCylOs.Text = "" Else txtCylOs.Text = rsB.Fields("OScyl").Value
        If IsNull(rsB.Fields("OSas")) Then txtAs.Text = "" Else txtAs.Text = rsB.Fields("OSas").Value
        If IsNull(rsB.Fields("OSJaval")) Then txtJavalOs.Text = "" Else txtJavalOs.Text = rsB.Fields("OSJaval").Value
        If IsNull(rsB.Fields("OSPlacido")) Then txtPlacidoOs.Text = "" Else txtPlacidoOs.Text = rsB.Fields("OSPlacido").Value
        If IsNull(rsB.Fields("OSSciascopia")) Then txtSciasOs.Text = "" Else txtSciasOs.Text = rsB.Fields("OSSciascopia").Value
        
        If IsNull(rsB.Fields("OOVisus")) Then txtOOvisus.Text = "" Else txtOOvisus.Text = rsB.Fields("OOVisus").Value
        If IsNull(rsB.Fields("OOsph")) Then txtDengan.Text = "" Else txtDengan.Text = rsB.Fields("OOsph").Value
        If IsNull(rsB.Fields("OOcyl")) Then txtSamadengan.Text = "" Else txtSamadengan.Text = rsB.Fields("OOcyl").Value
        If IsNull(rsB.Fields("OODist")) Then txtDist.Text = "" Else txtDist.Text = rsB.Fields("OODist").Value
        If IsNull(rsB.Fields("OOpup")) Then txtPup.Text = "" Else txtPup.Text = rsB.Fields("OOpup").Value
        If IsNull(rsB.Fields("OOOD")) Then txtOD.Text = "" Else txtOD.Text = rsB.Fields("OOOD").Value
        If IsNull(rsB.Fields("HistoryMO")) Then txtOcularis.Text = "" Else txtOcularis.Text = rsB.Fields("HistoryMO").Value
        If IsNull(rsB.Fields("HistoryMG")) Then txtGeneralis.Text = "" Else txtGeneralis.Text = rsB.Fields("HistoryMG").Value
    End If
    
    fraDokter.Visible = False
'    With frmTransaksiPasien
    With frmDaftarPasienRJ.dgDaftarPasienRJ
        txtnopendaftaran = .Columns("No. Registrasi").Value
        txtnocm = .Columns("NoCM").Value
        txtNamaPasien = .Columns("Nama Pasien").Value
        If .Columns("JK").Value = "P" Then txtSex.Text = "Perempuan" Else txtSex.Text = "Laki-Laki"
            
        txtThn = .Columns("UmurTahun").Value
        txtBln = .Columns("UmurBulan").Value
        txtHari = .Columns("UmurHari").Value
           mstrKdKelas = frmDaftarPasienRJ.dgDaftarPasienRJ.Columns("KdKelas").Value
          mstrKdSubInstalasi = frmDaftarPasienRJ.dgDaftarPasienRJ.Columns("KdSubInstalasi").Value
    End With
    

Exit Sub
errLoad:
    Call msubPesanError
    frmDaftarPasienRJ.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmDaftarPasienRJ.Enabled = True
End Sub

Private Sub subLoadDcSource()
On Error GoTo errLoad
    
    strSQL = "SELECT KdVisus, NamaVisus" & _
        " From VisusMata " & _
        " ORDER BY NamaVisus"
    Call msubDcSource(dcVisusMasuk, rs, strSQL)
    If rs.EOF = False Then dcVisusMasuk.BoundText = rs(0).Value
    
    strSQL = "SELECT KdVisus, NamaVisus" & _
        " From VisusMata " & _
        " ORDER BY NamaVisus"
    Call msubDcSource(dcVisusKeluar, rs, strSQL)
    If rs.EOF = False Then dcVisusKeluar.BoundText = rs(0).Value
    
    strSQL = "SELECT KdVisus, NamaVisus" & _
        " From VisusMata " & _
        " ORDER BY NamaVisus"
    Call msubDcSource(dcVisusMasukOS, rs, strSQL)
    If rs.EOF = False Then dcVisusMasukOS.BoundText = rs(0).Value
    
    strSQL = "SELECT KdVisus, NamaVisus" & _
        " From VisusMata " & _
        " ORDER BY NamaVisus"
    Call msubDcSource(dcVisusKeluarOs, rs, strSQL)
    If rs.EOF = False Then dcVisusKeluarOs.BoundText = rs(0).Value

    strSQL = "SELECT  IdPegawai, [Nama Pemeriksa]" & _
        " From V_DaftarPemeriksaPasien " & _
        " ORDER BY  [Nama Pemeriksa]"
    Call msubDcSource(dcPerawat, rs, strSQL)
    If rs.EOF = False Then dcPerawat.BoundText = strIDPegawaiAktif

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadDokter()
On Error GoTo errLoad
    strSQL = "SELECT KodeDokter AS [Kode Dokter],NamaDokter AS [Nama Dokter],JK,Jabatan " & _
             "FROM V_DaftarDokter " & strFilterDokter

'    strSQL = "SELECT IdPegawai AS [Kode Pemeriksa], [Nama Pemeriksa],JK,[Jenis Pemeriksa] " & _
'        " FROM V_DaftarDokterdanPemeriksaPasien " & strFilterDokter
    Call msubRecFO(rs, strSQL)
    Set dgDokter.DataSource = rs
    With dgDokter
        .Columns(0).Width = 1500
        .Columns(1).Width = 4000
        .Columns(2).Width = 400
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Width = 2000
    End With
    fraDokter.Left = 240
    fraDokter.Top = 3000
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub meBeratTingi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcVisusMasuk.SetFocus
    
End Sub

Private Sub txtPemeriksa_Change()
    strFilterDokter = "WHERE NamaDokter like '%" & txtPemeriksa.Text & "%'"
'    mstrKdDokter = ""
    fraDokter.Visible = True
    Call subLoadDokter
End Sub

Private Sub txtPemeriksa_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
End Sub



Private Sub txtCyl_KeyPress(KeyAscii As Integer)
'Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii <> 44 Then
        Call SetKeyPressToNumber(KeyAscii)
    Else
        If cekKoma(txtCyl, KeyAscii) = False Then
            KeyAscii = 44
         Else
            KeyAscii = 0
        End If
    End If

    If KeyAscii = 13 Then txtDerajat.SetFocus
End Sub

Private Sub txtDerajat_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then dcVisusKeluar.SetFocus
End Sub

Private Sub txtPlacidoOs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJavalOs.SetFocus
End Sub

Private Sub txtPup_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtOD.SetFocus
End Sub

Private Sub txtSamadengan_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 44 Then
        Call SetKeyPressToNumber(KeyAscii)
    Else
        If cekKoma(txtSamadengan, KeyAscii) = False Then
            KeyAscii = 44
         Else
            KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 Then txtDist.SetFocus
End Sub

Private Sub txtScias_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcVisusMasukOS.SetFocus
End Sub

Private Sub txtSciasOs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtOOvisus.SetFocus
End Sub

Private Sub txtSPH_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 44 Then
        Call SetKeyPressToNumber(KeyAscii)
    Else
        If cekKoma(txtSPH, KeyAscii) = False Then
            KeyAscii = 44
         Else
            KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 Then txtCyl.SetFocus
End Sub

Private Sub txtSPHOS_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 44 Then
        Call SetKeyPressToNumber(KeyAscii)
    Else
        If cekKoma(txtSPHOS, KeyAscii) = False Then
            KeyAscii = 44
         Else
            KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 Then txtCylOs.SetFocus
End Sub

Private Function cekKoma(objec As TextBox, KeyAscii As Integer) As Boolean
   cekKoma = True
   Dim s() As String, i As Long, a As String
   If Len(objec.Text) > 0 Then
        ReDim s(1 To Len(objec.Text))
        a = Chr(KeyAscii)
         For i = 1 To UBound(s)
             s(i) = Mid$(objec.Text, i, 1)
             If a = s(i) Then Exit Function
         Next i
   Else
       Exit Function
   End If
   
   cekKoma = False
End Function

