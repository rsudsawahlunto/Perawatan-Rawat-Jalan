VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOrderPelayananOA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Order Pelayanan Obat & Alkes"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18735
   Icon            =   "frmOrderPelayananOA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   18735
   Begin VB.TextBox txtNamaForm 
      Height          =   435
      Left            =   0
      TabIndex        =   83
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
   End
   Begin MSDataGridLib.DataGrid dgObatAlkes 
      Height          =   2535
      Left            =   2280
      TabIndex        =   82
      Top             =   -1440
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   19
      AllowAddNew     =   -1  'True
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
   Begin VB.Frame Frame4 
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
      TabIndex        =   64
      Top             =   960
      Width           =   18735
      Begin VB.TextBox txtRP 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10920
         TabIndex        =   76
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtNoPendaftaranOA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   75
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtNoCMOA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   480
         MaxLength       =   15
         TabIndex        =   74
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtNamaPasien 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3840
         MaxLength       =   50
         TabIndex        =   73
         Top             =   480
         Width           =   3135
      End
      Begin VB.Frame Frame5 
         Caption         =   "Umur"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   580
         Left            =   8400
         TabIndex        =   66
         Top             =   240
         Width           =   2415
         Begin VB.TextBox txtHr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            MaxLength       =   6
            TabIndex        =   69
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtBln 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   900
            MaxLength       =   6
            TabIndex        =   68
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtThn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   67
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2130
            TabIndex        =   72
            Top             =   270
            Width           =   165
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1350
            TabIndex        =   71
            Top             =   277
            Width           =   240
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   550
            TabIndex        =   70
            Top             =   277
            Width           =   285
         End
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7080
         TabIndex        =   65
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruangan Perawatan"
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
         Index           =   10
         Left            =   10920
         TabIndex        =   81
         Top             =   240
         Width           =   1725
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Pendaftaran"
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
         Index           =   6
         Left            =   2160
         TabIndex        =   80
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   7
         Left            =   480
         TabIndex        =   79
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   2
         Left            =   3840
         TabIndex        =   78
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label9 
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
         Left            =   7080
         TabIndex        =   77
         Top             =   240
         Width           =   1155
      End
   End
   Begin MSDataGridLib.DataGrid dgDokter 
      Height          =   2295
      Left            =   6960
      TabIndex        =   38
      Top             =   2880
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   19
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
   Begin VB.Frame FraRacikan 
      Caption         =   "Data Obat Racikan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   0
      TabIndex        =   39
      Top             =   3000
      Visible         =   0   'False
      Width           =   18735
      Begin MSDataGridLib.DataGrid dgObatAlkesRacikan 
         Height          =   2505
         Left            =   1440
         TabIndex        =   51
         Top             =   1320
         Visible         =   0   'False
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   4419
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   2
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox TxtIsiRacikan 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7320
         TabIndex        =   46
         Top             =   1440
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame fraHitungObat 
         Caption         =   "Berat Obat"
         Height          =   735
         Left            =   7200
         TabIndex        =   44
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
         Begin VB.TextBox txtBeratObat 
            Height          =   375
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.TextBox txtNoRacikan 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4560
         TabIndex        =   43
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdSelesai 
         BackColor       =   &H0000C000&
         Caption         =   "Selesa&i"
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
         Left            =   15240
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   4200
         Width           =   1455
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   16920
         TabIndex        =   41
         Top             =   4200
         Width           =   1455
      End
      Begin VB.TextBox txtJumlahObatRacik 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   40
         Top             =   240
         Width           =   1935
      End
      Begin MSFlexGridLib.MSFlexGrid fgRacikan 
         Height          =   3375
         Left            =   120
         TabIndex        =   47
         Top             =   840
         Width           =   18495
         _ExtentX        =   32623
         _ExtentY        =   5953
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblJumlahObatRacik 
         Caption         =   "Jumlah Obat Racik"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame frmDataResep 
      Caption         =   "Data Resep"
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
      TabIndex        =   52
      Top             =   2040
      Width           =   18735
      Begin VB.TextBox txtDokter 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6960
         TabIndex        =   55
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox txtNoResep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   3240
         MaxLength       =   15
         TabIndex        =   54
         Top             =   480
         Width           =   1935
      End
      Begin VB.CheckBox chkDokterPemeriksa 
         Caption         =   "Dokter Penulis Resep/ Dokter Order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   18720
         TabIndex        =   53
         Top             =   720
         Visible         =   0   'False
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker dtpTglOrder 
         Height          =   330
         Left            =   960
         TabIndex        =   56
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
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
         CustomFormat    =   "dd/MM/yyyy HH:mm:ss"
         Format          =   178585603
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSComCtl2.DTPicker dtpTglResep 
         Height          =   330
         Left            =   5280
         TabIndex        =   57
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   178585603
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSDataListLib.DataCombo dcRuanganTujuan 
         Height          =   330
         Left            =   10920
         TabIndex        =   58
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ListField       =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dokter Penulis Resep/ Dokter Order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   8
         Left            =   6960
         TabIndex        =   63
         Top             =   240
         Width           =   2940
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Resep"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   3240
         TabIndex        =   62
         Top             =   240
         Width           =   780
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruangan Tujuan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   10920
         TabIndex        =   61
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Resep"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   5280
         TabIndex        =   60
         Top             =   240
         Width           =   870
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   960
         TabIndex        =   59
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.Frame Frame8 
      Height          =   3735
      Left            =   0
      TabIndex        =   24
      Top             =   3000
      Width           =   18735
      Begin VB.TextBox txtKdBarang 
         Height          =   315
         Left            =   3720
         TabIndex        =   37
         Top             =   2280
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtSatuan 
         Height          =   315
         Left            =   4920
         TabIndex        =   36
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtKdAsal 
         Height          =   315
         Left            =   2760
         TabIndex        =   35
         Top             =   2280
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAsalBarang 
         Height          =   315
         Left            =   6000
         TabIndex        =   34
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtKdDokter 
         Height          =   315
         Left            =   1560
         TabIndex        =   33
         Top             =   2280
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtJenisBarang 
         Height          =   315
         Left            =   3000
         TabIndex        =   32
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtHargaBeli 
         Height          =   315
         Left            =   4320
         TabIndex        =   31
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtNoTemporary 
         Height          =   315
         Left            =   7080
         TabIndex        =   30
         Top             =   2280
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtIsi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   120
         TabIndex        =   29
         Top             =   1920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo dcKeteranganPakai2 
         Height          =   330
         Left            =   5160
         TabIndex        =   25
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcKeteranganPakai 
         Height          =   330
         Left            =   3360
         TabIndex        =   26
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcAturanPakai 
         Height          =   330
         Left            =   1560
         TabIndex        =   27
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcNamaPelayananRS 
         Height          =   330
         Left            =   1320
         TabIndex        =   28
         Top             =   1080
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcJenisObat 
         Height          =   330
         Left            =   120
         TabIndex        =   49
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   3375
         Left            =   0
         TabIndex        =   50
         Top             =   120
         Width           =   18495
         _ExtentX        =   32623
         _ExtentY        =   5953
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   -2147483643
         FocusRect       =   2
         HighLight       =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1920
      Top             =   600
   End
   Begin VB.TextBox txtNoOrder 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtNoPakai 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   0
      TabIndex        =   21
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   6840
      Width           =   18735
      Begin VB.TextBox txtTotalBiaya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5640
         MaxLength       =   12
         TabIndex        =   12
         Text            =   "0"
         Top             =   480
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtPembebasan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   12
         TabIndex        =   11
         Text            =   "0"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txtHarusDibayar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   16200
         MaxLength       =   12
         TabIndex        =   10
         Text            =   "0"
         Top             =   480
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtJumlahBayar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   12
         TabIndex        =   9
         Text            =   "0"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txtHutangPenjamin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10920
         MaxLength       =   12
         TabIndex        =   8
         Text            =   "0"
         Top             =   480
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtTanggunganRS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   13560
         MaxLength       =   12
         TabIndex        =   7
         Text            =   "0"
         Top             =   480
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtTotalDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8280
         MaxLength       =   12
         TabIndex        =   6
         Text            =   "0"
         Top             =   480
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtTotPasienBayar 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3240
         TabIndex        =   5
         Text            =   "0"
         Top             =   480
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Biaya"
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
         Left            =   5640
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pembebasan"
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
         Index           =   20
         Left            =   3240
         TabIndex        =   19
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Harus Dibayar"
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
         Index           =   21
         Left            =   16200
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Bayar"
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
         Index           =   22
         Left            =   3000
         TabIndex        =   17
         Top             =   1440
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Hutang Penjamin"
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
         Index           =   23
         Left            =   10920
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Tanggungan RS"
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
         Index           =   24
         Left            =   13560
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Discount"
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
         Index           =   25
         Left            =   8280
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total pasien Harus Bayar"
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
         Left            =   3240
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   2145
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   7800
      Width           =   18735
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
         Left            =   16920
         TabIndex        =   3
         Top             =   240
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
         Left            =   15120
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
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
         Left            =   13320
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   23
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
      Left            =   16920
      Picture         =   "frmOrderPelayananOA.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmOrderPelayananOA.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17055
   End
End
Attribute VB_Name = "frmOrderPelayananOA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim subintJmlArray As Integer
Dim subcurHargaSatuan As Currency
Dim subcurTarifService As Currency
Dim subcurHarusDibayar As Currency
'Dim subarrJumlahJual() As Double
Dim curTanggunganRS As Currency
Dim curHutangPenjamin As Currency
Dim subintJmlService As Integer
Dim tempStatusTampil As Boolean
Dim subJenisHargaNetto As Integer
Dim StrKdRP As String
Dim jmlBayarPerItem As Double
Dim jmlPembebasanTotal As Double
Dim jmlpembebasanPerItem As Double
Dim HargabarangBaru As Double
Dim JmlHargabarangbarupluskekuranganharga As Double
Dim KI As String
Dim Cancel As Boolean
Dim kolom As Integer
Dim riilnya As Double
Dim blt As Integer
Dim TempMstrKdIsntalasi As String
'---------------------------------------------------
Dim subcurTarifServiceRacikan As Currency
Dim subintJmlServiceRacikan As Currency
Dim strNoRacikan As String
Dim subJmlTotal As Integer
Dim cHargabeli As Currency
Dim iJmlService As Integer
Dim cTrfService As Currency
Dim BoolReview As Boolean
Dim boolUpdate As Boolean

Dim curHargaBrg As Currency
Dim curHarusDibayar As Currency
Dim tempKdRuanganTujuan As String
Dim posisiRowDataComboJenisObat As String
'Public UseRacikan As Boolean

'-----------------------------------------------------

'
'Public Sub chkDokterPemeriksa_Click()
'On Error GoTo errLoad
'
'    If chkDokterPemeriksa.Value = vbUnchecked Then
'        txtDokter.Enabled = False
'        txtDokter.Text = ""
'    ElseIf chkDokterPemeriksa.Value = vbChecked Then
'        txtDokter.Enabled = True
'        txtDokter.Text = mstrNamaDokter
'        txtKdDokter.Text = mstrKdDokter
'    End If
'    dgDokter.Visible = False
'
'Exit Sub
'errLoad:
'    Call msubPesanError
'End Sub

Private Sub chkDokterPemeriksa_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        If chkDokterPemeriksa.Value = vbChecked Then txtDokter.SetFocus Else dcRuanganTujuan.SetFocus
    End If
End Sub

Sub subAmbilTarifServiceRacikan()
On Error GoTo errLoad

    strSQL = "SELECT TarifService FROM V_AmbilTarifJenisObat WHERE (KdJenisObat = '" & dcJenisObat.BoundText & "') AND (KdKelompokPasien = '" & mstrKdJenisPasien & "') AND (IdPenjamin = '" & mstrKdPenjaminPasien & "')"
    Call msubRecFO(rsB, strSQL)
    If rsB.EOF = True Then
        strSQL = "SELECT TarifService FROM JenisObat WHERE (KdJenisObat = '" & dcJenisObat.BoundText & "')"
        Set rsB = Nothing
        Call msubRecFO(rsB, strSQL)
        If rsB.EOF = True Then
            subcurTarifServiceRacikan = 0
        Else
            subcurTarifServiceRacikan = IIf(IsNull(rsB(0).Value), 0, rsB(0).Value)
            subintJmlServiceRacikan = 1 'subintJmlService + 1
        End If
    Else
        subcurTarifServiceRacikan = IIf(IsNull(rsB(0).Value), 0, rsB(0).Value)
        subintJmlServiceRacikan = 1 'subintJmlService + 1
    End If
    
'    fgRacikan.TextMatrix(fgData.Row, 1) = dcJenisObat.Text
'    fgData.TextMatrix(fgData.Row, 25) = dcJenisObat.BoundText
    'fgRacikan.TextMatrix(fgData.Row, 15) = subcurTarifServiceRacikan
    'fgRacikan.TextMatrix(fgData.Row, 15) = subintJmlService
    If fgRacikan.Row - 1 = 0 Then Exit Sub
    fgRacikan.TextMatrix(fgRacikan.Row, 0) = fgRacikan.TextMatrix(fgRacikan.Row - 1, 0) + 1
 'Call dcJenisObat_LostFocus
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdBatal_Click()
    Call subAmbilTarifServiceRacikan
    Call subSetGridRacikan
    FraRacikan.Visible = False
    cmdSimpan.Enabled = True
    cmdTutup.Enabled = True
    txtNoRacikan.Text = ""
'    If fgData.TextMatrix(fgData.Row, 2) = "" Then
'    dcJenisObat.Text = ""
    With fgData
'        .TextMatrix(.Row, 0) = ""
'        .TextMatrix(.Row, 1) = ""
        .TextMatrix(.Row, 6) = ""
        .TextMatrix(.Row, 7) = ""
        .TextMatrix(.Row, 10) = ""
        .TextMatrix(.Row, 9) = ""
        .TextMatrix(.Row, 2) = ""
        .TextMatrix(.Row, 3) = ""
    End With
    
'    End If
    fgData.SetFocus
    fgData.Col = 1
    txtJumlahObatRacik.Enabled = True
    frmDataResep.Enabled = True
End Sub

Private Function sp_DetailOrderPelayananOARacikan(f_noOrder As String, f_Noracikan As String, _
     f_KdJenisObat As String, f_ResepKe As Integer) As Boolean
On Error GoTo errLoad

sp_DetailOrderPelayananOARacikan = True
    
    Set dbcmd = New ADODB.Command
    
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, f_noOrder)
        .Parameters.Append .CreateParameter("NoRacikan", adChar, adParamInput, 10, f_Noracikan)
        .Parameters.Append .CreateParameter("NoStruk", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("KdJenisObat", adChar, adParamInput, 2, f_KdJenisObat)
        .Parameters.Append .CreateParameter("ResepKe", adTinyInt, adParamInput, , f_ResepKe)
                        
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_DetailOrderPelayananOARacikan"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_DetailOrderPelayananOARacikan = False

        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    
Exit Function
errLoad:
    Call msubPesanError
    sp_DetailOrderPelayananOARacikan = False

End Function

Private Function sp_DetailOrderPelayananOARacikanTemp(f_KdBarang As String, _
        f_KdAsal As String, f_SatuanJml As String, f_NoTerima As String, f_jmlBrg As Double, _
        f_JmlPembulatan As Integer, f_ResepKe As Integer, f_jmlService, _
        f_TarifService As Integer, f_kebutuhanML As Double, f_kebutuhanTB As Double) As Boolean
On Error GoTo errLoad
    
    sp_DetailOrderPelayananOARacikanTemp = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        
        .Parameters.Append .CreateParameter("NoRacikan", adChar, adParamInput, 10, IIf(Trim(txtNoRacikan.Text) = "", Null, Trim(txtNoRacikan.Text)))
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("satuanJml", adChar, adParamInput, 1, f_SatuanJml)
        .Parameters.Append .CreateParameter("NoTerima", adChar, adParamInput, 10, IIf((f_NoTerima) = "", "0000000000", f_NoTerima))
        .Parameters.Append .CreateParameter("JmlBarang", adDouble, adParamInput, , CDec(f_jmlBrg))
        .Parameters.Append .CreateParameter("JmlPembulatan", adInteger, adParamInput, , f_JmlPembulatan)
        .Parameters.Append .CreateParameter("QtyRacikan", adInteger, adParamInput, , txtJumlahObatRacik.Text)
        .Parameters.Append .CreateParameter("KdJenisObat", adChar, adParamInput, 2, dcJenisObat.BoundText)
        .Parameters.Append .CreateParameter("ResepKe", adTinyInt, adParamInput, , f_ResepKe)
        .Parameters.Append .CreateParameter("JmlService", adInteger, adParamInput, , f_jmlService)
        
        .Parameters.Append .CreateParameter("TarifService", adCurrency, adParamInput, , f_TarifService)
        .Parameters.Append .CreateParameter("KebutuhanML", adDouble, adParamInput, , f_kebutuhanML)
        .Parameters.Append .CreateParameter("KebutuhanTB", adDouble, adParamInput, , f_kebutuhanTB)
       
       .Parameters.Append .CreateParameter("OutputKode", adChar, adParamOutput, 10, Null)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_DetailOrderPelayananOARacikanTemp"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_DetailOrderPelayananOARacikanTemp = False
        Else
             If Not IsNull(.Parameters("OutputKode").Value) Then
                txtNoRacikan.Text = .Parameters("OutputKode").Value
                strNoRacikan = txtNoRacikan.Text
            End If

        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    
Exit Function
errLoad:
    Call msubPesanError
    sp_DetailOrderPelayananOARacikanTemp = False


End Function

Private Sub cmdSelesai_Click()
On Error GoTo errLoad
Dim i As Integer
Dim TotalRacikan As Currency
Dim strAsalBarang As String
Dim strKdAsal As String
Dim strSatuan As String
Dim strNoTerima As String
Dim TempStrjumlah As Integer

If txtJumlahObatRacik.Text = "" Then MsgBox "Jumlah racikan kosong", vbCritical, "Validasi": Exit Sub

    For i = 1 To fgRacikan.Rows - 1
        If fgRacikan.TextMatrix(i, 4) = "" Then
            MsgBox "Lengkapi data pesanan obat", vbInformation, "Validasi"
            fgRacikan.SetFocus
            fgRacikan.Col = 4
            Exit Sub
        End If
         If fgRacikan.TextMatrix(i, 6) = "" Then
            MsgBox "Jumlah data pesanan obat masih kosong", vbInformation, "Validasi"
            fgRacikan.SetFocus
            fgRacikan.Col = 6
            Exit Sub
        End If
        
        If fgRacikan.TextMatrix(i, 6) = "" Then TempStrjumlah = 0 Else TempStrjumlah = fgRacikan.TextMatrix(i, 6)
        If TempStrjumlah > CDbl(fgRacikan.TextMatrix(i, 17)) Then
            MsgBox "Jumlah lebih besar dari stok, " & vbNewLine & " dan stok barang " & fgRacikan.TextMatrix(i, 3) & " yang tersedia " & fgRacikan.TextMatrix(i, 17) & " ", vbExclamation, "Validasi"
            fgRacikan.SetFocus
            Exit Sub
        End If

       
    Next i

    
    If fgRacikan.TextMatrix(1, 0) = "" Then MsgBox "Transaksi racikan belum diisi", vbExclamation, "Validasi": Exit Sub
    With fgRacikan
        For i = 1 To .Rows - 1
        If .TextMatrix(i, 0) = "" Then GoTo lanjutkan_
        .TextMatrix(i, 15) = IIf(.TextMatrix(i, 16) = "", "0", .TextMatrix(i, 16))
        .TextMatrix(i, 16) = IIf(.TextMatrix(i, 16) = "", "0", .TextMatrix(i, 16))
        If sp_DetailOrderPelayananOARacikanTemp(.TextMatrix(i, 0), _
         .TextMatrix(i, 12), _
         .TextMatrix(i, 13), _
         .TextMatrix(i, 14), _
         Val(.TextMatrix(i, 6)), _
         .TextMatrix(i, 7), _
         .TextMatrix(i, 2), _
         .TextMatrix(i, 15), _
         .TextMatrix(i, 16), _
         .TextMatrix(i, 4), _
         .TextMatrix(i, 5)) = False Then Exit Sub
        
        
         TotalRacikan = TotalRacikan + .TextMatrix(i, 9)
         strAsalBarang = .TextMatrix(i, 11)
         strKdAsal = .TextMatrix(i, 12)
         strSatuan = .TextMatrix(i, 13)
         strNoTerima = .TextMatrix(i, 14)
        
lanjutkan_:
        Next i
    End With
    
    cmdTutup.Enabled = True
    cmdSimpan.Enabled = True
    txtJumlahObatRacik.Enabled = True
    FraRacikan.Visible = False
    frmDataResep.Enabled = True
          
    With fgData
        .TextMatrix(.Row, 1) = dcJenisObat.Text
        .TextMatrix(.Row, 5) = strAsalBarang
        .TextMatrix(.Row, 6) = strSatuan
        .TextMatrix(.Row, 8) = 0
        .TextMatrix(.Row, 9) = "-"
        .TextMatrix(.Row, 3) = dcJenisObat.Text  ' nama barang
        .TextMatrix(.Row, 25) = dcJenisObat.BoundText
        .TextMatrix(.Row, 10) = txtJumlahObatRacik.Text
        .TextMatrix(.Row, 11) = TotalRacikan
        .TextMatrix(.Row, 12) = strKdAsal
        .TextMatrix(.Row, 14) = 0
        .TextMatrix(.Row, 15) = 0
        .TextMatrix(.Row, 7) = (TotalRacikan) / Val(txtJumlahObatRacik.Text)
        .TextMatrix(.Row, 2) = "Racikan"
        .TextMatrix(.Row, 33) = IIf(strNoTerima = "", "0000000000", strNoTerima)
        .TextMatrix(.Row, 34) = strNoRacikan
            
        .Rows = .Rows + 1
        .SetFocus
        .Col = 5
    End With
   
   Call Hitung
    Call subHitungTotal
    Call hitungRacikan
    Call subAmbilTarifServiceRacikan
    Call subSetGridRacikan
    
    txtJumlahObatRacik.Text = ""
    
    'Add Arief For Racikan Otomatis
    strNoRacikan = ""
    txtNoRacikan.Text = ""
    'end arief
        
'    MsgBox "Data Racikan Telah Disimpan", vbInformation, "Informasi"
    fgData.SetFocus
    fgData.Col = 27
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub hitungRacikan()
On Error Resume Next
Dim i As Integer
Dim curHutangPenjamin As Currency
Dim curTanggunganRS As Currency
Dim curHarusDibayar As Currency
Dim curTempTotal As Currency
With fgData
    'ambil no temporary
'    If sp_TempDetailApotikJual(CDbl(.TextMatrix(.Row, 7)), .TextMatrix(.Row, 2), .TextMatrix(.Row, 12)) = False Then Exit Sub
    'ambil hutang penjamin dan tanggungan rs
    strSQL = "SELECT HargaSatuan, JmlHutangPenjamin, JmlTanggunganRS" & _
        " FROM TempDetailApotikJual" & _
        " WHERE (NoTemporary = '" & Trim(txtNoTemporary.Text) & "') AND (KdBarang = '" & .TextMatrix(.Row, 2) & "') AND (KdAsal = '" & .TextMatrix(.Row, 12) & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
        curHutangPenjamin = 0
        curTanggunganRS = 0
    Else
        curHutangPenjamin = rs("JmlHutangPenjamin").Value
        curTanggunganRS = rs("JmlTanggunganRS").Value
    End If
    
    .TextMatrix(.Row, 16) = CDbl(.TextMatrix(.Row, 7))
    
    'total harga = ((tarifservice * jmlservice) + _
        (hargasatuan(sebelum ditambah tarifservixe) * jumlah))
    
    
    
'    .TextMatrix(.Row, 11) = ((CDbl(.TextMatrix(.Row, 14)) * CDbl(.TextMatrix(.Row, 15))) + _
'       (CDbl(.TextMatrix(.Row, 16)) * val(.TextMatrix(.Row, 10)))) + val(.TextMatrix(.Row, 26))
'
'    .TextMatrix(.Row, 11) = Format(.TextMatrix(.Row, 11), "#,###.00")
'                    .Col = 11: .CellForeColor = vbBlue: .CellFontBold = True: .Col = 10

    .TextMatrix(.Row, 17) = curHutangPenjamin
    .TextMatrix(.Row, 18) = curTanggunganRS
    
    If curHutangPenjamin > 0 Then
        'akung
        .TextMatrix(.Row, 19) = .TextMatrix(.Row, 11)
        '--
        '.TextMatrix(.Row, 19) = (.TextMatrix(.Row, 14) * .TextMatrix(.Row, 15)) + (val(.TextMatrix(.Row, 10)) * CDbl(curHutangPenjamin)) + val(.TextMatrix(.Row, 26))
    Else
        .TextMatrix(.Row, 19) = 0
    End If
    
    If curTanggunganRS > 0 Then
        .TextMatrix(.Row, 20) = .TextMatrix(.Row, 11)
        '.TextMatrix(.Row, 20) = (.TextMatrix(.Row, 14) * .TextMatrix(.Row, 15)) + (val(.TextMatrix(.Row, 10)) * CDbl(curTanggunganRS)) + val(.TextMatrix(.Row, 26))
    Else
        .TextMatrix(.Row, 20) = 0
    End If
    .TextMatrix(.Row, 21) = CDbl(.TextMatrix(.Row, 10)) * CDbl(.TextMatrix(.Row, 8))
    
    'total harus dibayar = total harga - total discount - _
        total hutang penjamin - totaltanggunganrs
    curHarusDibayar = CDbl(.TextMatrix(.Row, 11)) - (CDbl(.TextMatrix(.Row, 21)) + _
        CDbl(.TextMatrix(.Row, 19)) + CDbl(.TextMatrix(.Row, 20)))
    .TextMatrix(.Row, 22) = FormatPembulatan(IIf(curHarusDibayar < 0, 0, curHarusDibayar), KI)
End With
    Call Hitung
    Call subHitungTotal
    
End Sub

Private Sub subHitungTotal()

On Error GoTo errLoad
Dim i As Integer

    txtTotalBiaya.Text = 0
    txtHutangPenjamin.Text = 0
    txtTanggunganRS.Text = 0
    txtHarusDibayar.Text = 0
    txtTotalDiscount.Text = 0
            
    With fgData
        For i = 1 To IIf(fgData.TextMatrix(fgData.Rows - 1, 2) = "", fgData.Rows - 2, fgData.Rows - 1)
            If .TextMatrix(i, 11) = "" Then .TextMatrix(i, 11) = 0
            If .TextMatrix(i, 22) = "" Then .TextMatrix(i, 22) = 0
            If .TextMatrix(i, 19) = "" Then .TextMatrix(i, 19) = 0
            If .TextMatrix(i, 20) = "" Then .TextMatrix(i, 20) = 0
            If .TextMatrix(i, 21) = "" Then .TextMatrix(i, 21) = 0
            'agar user mengetahui biaya pasien ketika entry peritem obat apakan dijamin atau tdk
        'txtTotalBiaya.Text = txtTotalBiaya.Text + CDbl(.TextMatrix(i, 11))
            txtTotalBiaya.Text = txtTotalBiaya.Text + CDbl(.TextMatrix(i, 22))
            'txtHutangPenjamin.Text = txtTotalBiaya.Text + CDbl(.TextMatrix(i, 19))
            'txtTotalBiaya.Text = txtTotalBiaya.Text + CDbl(.TextMatrix(i, 22))
            
'            txtHutangPenjamin.Text = txtHutangPenjamin.Text + CDbl(IIf(.TextMatrix(i, 19) = "", 0, .TextMatrix(i, 19)))
'            txtTanggunganRS.Text = txtTanggunganRS.Text + CDbl(IIf(.TextMatrix(i, 20) = "", 0, .TextMatrix(i, 20)))
'            txtTotalDiscount.Text = txtTotalDiscount.Text + CDbl(IIf(.TextMatrix(i, 21) = "", 0, .TextMatrix(i, 21)))
             txtHutangPenjamin.Text = txtHutangPenjamin.Text + CDbl(.TextMatrix(i, 19))
             txtTanggunganRS.Text = txtTanggunganRS.Text + CDbl(.TextMatrix(i, 20))
             txtTotalDiscount.Text = txtTotalDiscount.Text + CDbl(.TextMatrix(i, 21))
'            txtHarusDibayar.Text = txtHarusDibayar.Text + CDbl(.TextMatrix(i, 22) - CDbl(.TextMatrix(i, 21) - CDbl(.TextMatrix(i, 19) - CDbl(.TextMatrix(i, 20)))))
            txtHarusDibayar.Text = txtHarusDibayar.Text + CDbl(.TextMatrix(i, 11))
        Next i
    End With
    
     'txtTotalBiaya.Text = funcRound(CStr(CCur(txtTotalBiaya.Text)), CDbl(typSettingDataPendukung.intJmlPembulatanHarga))
    txtTotalBiaya.Text = CStr(CCur(txtTotalBiaya.Text))
    txtTotalBiaya.Text = IIf(Val(txtTotalBiaya.Text) = 0, 0, Format(txtTotalBiaya.Text, "#,###.00"))
    
    'txtHutangPenjamin.Text = funcRound(CStr(CCur(txtHutangPenjamin.Text)), CDbl(typSettingDataPendukung.intJmlPembulatanHarga))
    txtHutangPenjamin.Text = CStr(CCur(txtHutangPenjamin.Text))
    txtHutangPenjamin.Text = IIf(Val(txtHutangPenjamin.Text) = 0, 0, Format(txtHutangPenjamin.Text, "#,###.00"))
    
    'txtTanggunganRS.Text = funcRound(CStr(CCur(txtTanggunganRS.Text)), CDbl(typSettingDataPendukung.intJmlPembulatanHarga))
    txtTanggunganRS.Text = CStr(CCur(txtTanggunganRS.Text))
    txtTanggunganRS.Text = IIf(Val(txtTanggunganRS.Text) = 0, 0, Format(txtTanggunganRS.Text, "#,###.00"))
    
    'txtHarusDibayar.Text = funcRound(CStr(CCur(txtHarusDibayar.Text)), CDbl(typSettingDataPendukung.intJmlPembulatanHarga))
    txtHarusDibayar.Text = CStr(CCur(txtHarusDibayar.Text))
    txtHarusDibayar.Text = IIf(Val(txtHarusDibayar.Text) = 0, 0, Format(txtHarusDibayar.Text, "#,###.00"))
    
    txtTotalDiscount.Text = IIf(Val(txtTotalDiscount.Text) = 0, 0, Format(txtTotalDiscount.Text, "#,###.00"))
    
    subcurHarusDibayar = txtHarusDibayar.Text
    
    

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub Hitung()
Dim i As Integer
    txtHutangPenjamin.Text = 0
    txtHarusDibayar.Text = 0
    txtTanggunganRS.Text = 0
    If mstrKdJenisPasien <> "07" And mstrKdJenisPasien <> "14" Then
        txtHutangPenjamin.Text = (txtTotalBiaya.Text - txtTotPasienBayar.Text)
        txtHarusDibayar.Text = txtTotPasienBayar.Text
    Else
        If txtTotalDiscount.Text > 0 Then
            txtTanggunganRS.Text = (txtTotalBiaya.Text - txtTotalDiscount.Text)
            txtHarusDibayar.Text = txtTotPasienBayar.Text
        Else
            txtTanggunganRS.Text = (txtTotalBiaya.Text - txtTotPasienBayar.Text)
            txtHarusDibayar.Text = txtTotPasienBayar.Text
        End If
        
    End If
    txtTanggunganRS.Text = IIf(Val(txtTanggunganRS.Text) = 0, 0, Format(txtTanggunganRS.Text, "#,###.00"))
    txtHutangPenjamin.Text = IIf(Val(txtHutangPenjamin.Text) = 0, 0, Format(txtHutangPenjamin.Text, "#,###.00"))
    txtHarusDibayar.Text = IIf(Val(txtHarusDibayar.Text) = 0, 0, Format(txtHarusDibayar.Text, "#,###.00"))
    txtTotalDiscount.Text = IIf(Val(txtTotalDiscount.Text) = 0, 0, Format(txtTotalDiscount.Text, "#,###.00"))
End Sub

Private Sub cmdSelesai_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then fgData.SetFocus: fgData.Col = 27

End Sub

Private Sub cmdSelesai_LostFocus()
    fgData.SetFocus
    fgData.Col = 27
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo errLoad

Dim i, j, z, y, X As Integer 'Var x Add By Indra

    If txtDokter.Text = "" Then
    MsgBox "Silahkan isi dokter", vbExclamation, "Validasi"
    txtDokter.SetFocus
    Exit Sub
    End If

    If dcRuanganTujuan.Text = "" Then
    MsgBox "Silahkan isi ruangan tujuan", vbExclamation, "Validasi"
    dcRuanganTujuan.SetFocus
    dcRuanganTujuan.Text = ""
    Exit Sub
    End If
    
    If fgData.TextMatrix(1, 2) = "" Then MsgBox "Data barang harus diisi lengkap", vbExclamation, "Validasi": fgData.SetFocus: Exit Sub
    With fgData
        For z = 1 To .Rows - 1
            If .TextMatrix(z, 2) = "" Then GoTo lanjut0_
            If .TextMatrix(z, 10) = "" Then MsgBox "jumlah barang harus diisi", vbExclamation, "Validasi": Exit Sub
            If .TextMatrix(z, 1) = "" Then MsgBox "Jenis Obat Masih Kosong", vbExclamation, "Validasi": Exit Sub
        Next z
lanjut0_:
    End With

'    With fgData
'    For z = 1 To .Rows - 1
'    If .TextMatrix(z, 1) = "" Then
'        MsgBox "Jenis Obat Harus Diisi", vbExclamation, "Validasi"
'    Exit Sub
'    End If
'
'    If .TextMatrix(z, 2) = "" Then
'        MsgBox "Data Barang Harus Diisi", vbExclamation, "Validasi"
'    Exit Sub
'    End If
'    Next z

'    End With
    
    
    With fgData
        'Add By Indra
        For X = 1 To .Rows - 1
            If Val(.TextMatrix(X, 1)) = 0 Then
                If .TextMatrix(X, 2) = "" Then Exit For
                If .TextMatrix(X, 1) = "" Then
                    MsgBox "Jenis Obat Harus Diisi", vbExclamation, "Validasi"
                    Exit Sub
                End If
            End If
        Next X
        
        For j = 1 To .Rows - 1
            If Val(.TextMatrix(j, 10)) = 0 Then
            If .TextMatrix(j, 2) <> "" Then
                MsgBox "Barang tidak boleh diisi 0", vbExclamation, "Validasi"
                Exit Sub
            End If
            End If
        Next j
        
        
        '-----------
    End With
    
    If sp_Order() = False Then Exit Sub
    strNoOrder = txtNoOrder.Text ' inisialisasi
        
        
    With fgData
        For i = 1 To .Rows - 1
        If .TextMatrix(i, 1) <> "" Then
            If .TextMatrix(i, 1) = "Racikan" Then
'Ditutup by Azein
'         If fgData.TextMatrix(i, 25) = "" Then
'            If fgData.TextMatrix(i, 2) <> "" Or fgData.TextMatrix(i, 12) <> "" Then
'                MsgBox "Jenis obat kosong", vbExclamation, "Validasi"
'                fgData.Col = 1: fgData.SetFocus
'                Exit Sub
'            End If
'         End If
'end tutup
                strSQLx = "select NoRacikan,KdBarang,KdAsal,SatuanJml,NoTerima,JmlBarang,JmlPembulatan,QTYRacikan,KdJenisObat,ResepKe,JmlService,TarifService,KebutuhanML,KebutuhanTB from DetailOrderPelayananOARacikantemp Where NoRacikan='" & .TextMatrix(i, 34) & "' And KdJenisObat='" & .TextMatrix(i, 25) & "' AND ResepKe='" & Val(.TextMatrix(i, 0)) & "'"
                     
                Set rsB = Nothing
                Call msubRecFO(rsB, strSQLx)
                If rsB.EOF = False Then
                   rsB.MoveFirst
                    For j = 1 To rsB.RecordCount
                        If sp_DetailOrderPelayananOARacikan(strNoOrder, .TextMatrix(i, 34), _
                            .TextMatrix(i, 25), .TextMatrix(i, 0)) = False Then Exit Sub
                                
                        If sp_DetailOrderPelayananOA(rsB("kdBarang").Value, rsB("JmlBarang").Value, rsB("ResepKe").Value, _
                          .TextMatrix(i, 35), .TextMatrix(i, 36), .TextMatrix(i, 37), .TextMatrix(i, 32), .TextMatrix(i, 30), CInt(IIf(.TextMatrix(i, 15) = "", 0, .TextMatrix(i, 15))), _
                          CCur(IIf(.TextMatrix(i, 14) = "", 0, .TextMatrix(i, 14))), "No", rsB("KdJenisObat").Value, rsB("KdAsal").Value, .TextMatrix(i, 6), .TextMatrix(i, 33), _
                         .TextMatrix(i, 34)) = False Then Exit Sub

                    rsB.MoveNext
                    Next j
'                Else
'                        GoTo lanjutkan
                End If
'             rsB.MoveFirst
            Else
                If .TextMatrix(i, 2) <> "" Then
                
                  If sp_DetailOrderPelayananOA(.TextMatrix(i, 2), CDbl(.TextMatrix(i, 10)), .TextMatrix(i, 0), _
                    .TextMatrix(i, 35), .TextMatrix(i, 36), .TextMatrix(i, 37), .TextMatrix(i, 32), .TextMatrix(i, 30), CInt(IIf(.TextMatrix(i, 15) = "", 0, .TextMatrix(i, 15))), _
                    CCur(IIf(.TextMatrix(i, 14) = "", 0, .TextMatrix(i, 14))), "No", .TextMatrix(i, 25), .TextMatrix(i, 12), .TextMatrix(i, 6), .TextMatrix(i, 33), _
                    .TextMatrix(i, 34)) = False Then Exit Sub
                    
                End If
            End If
'lanjutkan:
        End If
        Next i
    End With
    
    MsgBox "Data berhasil disimpan", vbInformation, "Informasi"
    cmdSimpan.Enabled = False
    Dim psncetak1 As String
    psncetak1 = MsgBox("Data order pelayanan OA dengan Nama Pasien " & Trim(txtNamaPasien.Text) & " ingin dicetak ?", vbInformation + vbYesNo)
    If psncetak1 = vbYes Then
    cmdCetak_Click
    End If
    
    Call Add_HistoryLoginActivity("Add_DetailOrderPelayananOA")
    
    txtNoResep.Text = ""
    txtTotalBiaya.Text = 0
    txtTotalDiscount.Text = 0
    txtHutangPenjamin.Text = 0
    txtTanggunganRS.Text = 0
    txtHarusDibayar.Text = 0
    txtNoOrder.Text = ""
    
    
    Call subSetGrid

Exit Sub
errLoad:
    Call msubPesanError
    Resume 0
End Sub


Private Sub cmdtutup_Click()
Dim i As Integer
If fgData.TextMatrix(1, 1) <> "" Then
    If cmdSimpan.Enabled = True Then
        If MsgBox("Simpan Data Order Pelayanan Obat dan Alat", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
            Exit Sub
        End If
    
'    For i = 1 To fgData.Rows - 1
'    With fgData
'         If .TextMatrix(i, 2) <> "" Then If sp_StokRuangan(.TextMatrix(i, 2), .TextMatrix(i, 12), CDbl(.TextMatrix(i, 10)), "A") = False Then Exit Sub
'    End With
'    Next i
    End If
    
        For i = 1 To fgData.Rows - 1
            If fgData.TextMatrix(i, 34) <> "0000000000" Then
                dbConn.Execute "DELETE FROM DetailorderPelayananOARacikanTemp WHERE (NoRacikan = '" & fgData.TextMatrix(i, 34) & "')"
            End If
        Next i
    Call frmTransaksiPasien.subLoadRiwayatKonsul
    Unload Me
'    frmDaftarPasienRI.Enabled = True
Else
    Call frmTransaksiPasien.subLoadRiwayatKonsul
    Unload Me
End If
End Sub

Private Sub cmdCetak_Click()
On Error GoTo errLoad
    If cmdSimpan.Enabled = True Then Exit Sub
    mdTglAwal = dtpTglOrder.Value
    mdTglAkhir = dtpTglOrder.Value
    mstrNoPen = txtNoPendaftaranOA.Text
    mstrKdRuanganORS = dcRuanganTujuan.BoundText
    strNamaRuangan = dcRuanganTujuan.Text
    mstrNama = txtDokter.Text
    mstrNoPen = txtNoPendaftaranOA.Text
    mstrNoOrder = strNoOrder
    
    strCetak2 = "OA"
    
    frm_cetak_RincianBiayaKonsul.Show
Exit Sub
errLoad:
    Call msubPesanError
End Sub
'Private Sub cmdBatalRacikan_Click()
'
''fgRacikan.Clear
''Call subSetGridRacikan
''Cancel = False
''FraRacikan.Visible = False
'
'    Call subAmbilTarifServiceRacikan
'    Call subSetGridRacikan
'    FraRacikan.Visible = False
'    cmdSimpan.Enabled = True
'    cmdTutup.Enabled = True
'    txtNoRacikan.Text = ""
'    If fgData.TextMatrix(fgData.Row, 2) = "" Then dcJenisObat.Text = ""
'
'End Sub

Private Sub dcAturanPakai_Change()
On Error GoTo errLoad
    fgData.TextMatrix(fgData.Row, 27) = dcAturanPakai.Text
    fgData.TextMatrix(fgData.Row, 35) = dcAturanPakai.BoundText
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcAturanPakai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        dcAturanPakai.Visible = False
        fgData.Col = 27
        fgData.SetFocus
    End If
End Sub

Private Sub dcAturanPakai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call dcAturanPakai_Change
        dcAturanPakai.Visible = False
        fgData.Col = 27
        fgData.SetFocus
        Call fgData_KeyDown(13, 0)
    End If
End Sub

Private Sub dcAturanPakai_LostFocus()
    dcAturanPakai.Visible = False
End Sub

Private Sub dcJenisObat_Change()
On Error GoTo errLoad
    If dcRuanganTujuan.Text = "" Then
            MsgBox "Silahkan pilih ruangan tujuan", vbInformation + vbOKOnly, "Info"
            dcJenisObat.Text = ""
            dcRuanganTujuan.Text = ""
            dcRuanganTujuan.SetFocus
    Exit Sub
    End If
    strSQL = "SELECT TarifService FROM V_AmbilTarifJenisObat WHERE (KdJenisObat = '" & dcJenisObat.BoundText & "') AND (KdKelompokPasien = '" & mstrKdJenisPasien & "') AND (IdPenjamin = '" & mstrKdPenjaminPasien & "')"
    Call msubRecFO(dbRst, strSQL)
    If dbRst.EOF = True Then
        subcurTarifService = 0
        subintJmlService = 0
    Else
        subcurTarifService = dbRst(0).Value
        subintJmlService = 1
    End If
    fgData.TextMatrix(fgData.Row, 1) = dcJenisObat.Text
    fgData.TextMatrix(fgData.Row, 25) = dcJenisObat.BoundText

    If fgData.TextMatrix(fgData.Row - 1, fgData.Col - 1) = fgData.TextMatrix(fgData.Row, fgData.Col - 1) Then
        subcurTarifService = 0
        subintJmlService = 0
        fgData.TextMatrix(fgData.Row, 14) = 0
        fgData.TextMatrix(fgData.Row, 15) = 0
    Else
        fgData.TextMatrix(fgData.Row, 14) = subcurTarifService
        fgData.TextMatrix(fgData.Row, 15) = subintJmlService
    End If

    If fgData.Row = 1 Then
        If fgData.TextMatrix(fgData.Row, 0) = "" Then fgData.TextMatrix(fgData.Row, 0) = 1
    Else
        If fgData.Row - 1 = 0 Then Exit Sub
        fgData.TextMatrix(fgData.Row, 0) = fgData.TextMatrix(fgData.Row - 1, 0) ' + 1
    End If

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcJenisObat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call dcJenisObat_Change
        dcJenisObat.Visible = False
        fgData.Col = 3
        fgData.SetFocus
    End If
End Sub

Private Sub dcJenisObat_LostFocus()

    dcJenisObat.Visible = False
    If Cancel = False Then
    Dim no As Integer
    If fgData.TextMatrix(fgData.Row, 0) = "" Then
        no = Val(fgData.TextMatrix(fgData.Row - 1, 0)) ' + 1
        fgData.TextMatrix(fgData.Row, 0) = no
    End If
    
    If posisiRowDataComboJenisObat <> fgData.Row Then Exit Sub
    If dcJenisObat.BoundText = "01" And dcJenisObat.BoundText <> "" Then
        Call subSetGridRacikan
        FraRacikan.Visible = True
        dgObatAlkesRacikan.Visible = False
        txtJumlahObatRacik.SetFocus
        txtJumlahObatRacik.Text = ""
        cmdSimpan.Enabled = False
        cmdTutup.Enabled = False
        fgRacikan.TextMatrix(fgRacikan.Row, 2) = fgData.TextMatrix(fgData.Row, 0) ' edit
        
        frmDataResep.Enabled = False
    Else
        With fgData
            .TextMatrix(.Row, 1) = dcJenisObat.Text
            .TextMatrix(.Row, 25) = dcJenisObat.BoundText
        End With
    End If
    
End If

Cancel = False
End Sub



Private Sub dcKeteranganPakai_Change()
On Error GoTo errLoad
    fgData.TextMatrix(fgData.Row, 28) = dcKeteranganPakai.Text
    fgData.TextMatrix(fgData.Row, 36) = dcKeteranganPakai.BoundText
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcKeteranganPakai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        dcKeteranganPakai.Visible = False
        fgData.Col = 28
        fgData.SetFocus
    End If
End Sub

Private Sub dcKeteranganPakai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call dcKeteranganPakai_Change
        dcKeteranganPakai.Visible = False
        fgData.Col = 28
        fgData.SetFocus
        Call fgData_KeyDown(13, 0)
    End If
End Sub

Private Sub dcKeteranganPakai_LostFocus()
    dcKeteranganPakai.Visible = False
End Sub

Private Sub dcKeteranganPakai2_Change()
On Error GoTo errLoad
    fgData.TextMatrix(fgData.Row, 29) = dcKeteranganPakai2.Text
    fgData.TextMatrix(fgData.Row, 37) = dcKeteranganPakai2.BoundText
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcKeteranganPakai2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        dcKeteranganPakai2.Visible = False
        fgData.Col = 29
        fgData.SetFocus
    End If
End Sub

Private Sub dcKeteranganPakai2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call dcKeteranganPakai2_Change
        dcKeteranganPakai2.Visible = False
        fgData.Col = 29
        fgData.SetFocus
        Call fgData_KeyDown(13, 0)
    End If
End Sub

Private Sub dcKeteranganPakai2_LostFocus()
    dcKeteranganPakai2.Visible = False
End Sub

Private Sub dcNamaPelayananRS_Change()
On Error GoTo errLoad
    fgData.TextMatrix(fgData.Row, 31) = dcNamaPelayananRS.Text
    fgData.TextMatrix(fgData.Row, 32) = dcNamaPelayananRS.BoundText
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcNamaPelayananRS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        dcNamaPelayananRS.Visible = False
        fgData.SetFocus
        fgData.Col = 35
    End If
End Sub

Private Sub dcNamaPelayananRS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call dcNamaPelayananRS_Change
        dcNamaPelayananRS.Visible = False
        fgData.SetFocus
        fgData.Col = 35
        Call fgData_KeyDown(13, 0)
    End If
End Sub

Private Sub dcNamaPelayananRS_LostFocus()
    dcNamaPelayananRS.Visible = False
End Sub

Private Sub TampilKdRuangan()
fgData.Clear
Call subSetGrid
Call subLoadDcSource
dcJenisObat.BoundText = ""

strSQL = "SELECT JenisHargaNetto" & _
        " From PersentaseUpTarifOA" & _
        " Where(IdPenjamin = '" & mstrKdPenjaminPasien & "') And (KdKelompokPasien = '" & mstrKdJenisPasien & "')"
    Call msubRecFO(rs, strSQL)
    subJenisHargaNetto = IIf(rs.EOF = True, 1, rs(0))
   
    
    If dcRuanganTujuan.MatchedWithList = True Then fgData.SetFocus
End Sub

Private Sub dcRuanganTujuan_Change()
    Dim pesangantiRuanganTujuan As String
    Dim i As Integer
    If fgData.TextMatrix(1, 2) = "" Then Exit Sub
    If dcRuanganTujuan.BoundText <> tempKdRuanganTujuan Then

    pesangantiRuanganTujuan = MsgBox("Pilih Yes Jika akan menghapus detail Pemesanan Obat?", vbInformation + vbYesNo)
    
        If pesangantiRuanganTujuan = vbYes Then
            
            If FraRacikan.Visible = True Then
                Call cmdBatal_Click
            End If
                
            With fgData
                For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) <> "" Then
                    If .TextMatrix(i, 1) = "Racikan" Then
                        StrSQL12 = "delete DetailOrderPelayananOARacikanTemp where NoRacikan='" & .TextMatrix(i, 34) & "'"
                    End If
                End If
                Next i
            End With
        
            txtNoResep.Text = ""
            txtTotalBiaya.Text = 0
            txtTotalDiscount.Text = 0
            txtHutangPenjamin.Text = 0
            txtTanggunganRS.Text = 0
            txtHarusDibayar.Text = 0
            txtNoOrder.Text = ""
            
            
            Call subSetGrid
            tempKdRuanganTujuan = dcRuanganTujuan.BoundText
        Else
        If dcRuanganTujuan.BoundText <> tempKdRuanganTujuan Then dcRuanganTujuan.BoundText = tempKdRuanganTujuan
            
        '    Exit Sub
        End If
        
       End If

End Sub

Private Sub dcRuanganTujuan_Click(Area As Integer)
''    Call clearRincianDetail
End Sub

Private Sub dcRuanganTujuan_DblClick(Area As Integer)
'    Call clearRincianDetail
End Sub

Private Sub dcRuanganTujuan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then fgData.SetFocus: fgData.Col = 1
End Sub

'end by gantri
Private Sub dcRuanganTujuan_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    If KeyAscii = 13 Then
          
        If dcRuanganTujuan.MatchedWithList = True Then fgData.SetFocus
        
        strSQL = "Select KdRuangan, NamaRuangan, KdInstalasi FROM Ruangan WHERE (NamaRuangan LIKE '%" & dcRuanganTujuan.Text & "%') And kdruangan not in ('701') AND StatusEnabled=1"
        Call msubRecFO(rs, strSQL)
        
        If rs.EOF = True Then dcRuanganTujuan.Text = "": dcRuanganTujuan.SetFocus: Exit Sub
        dcRuanganTujuan.BoundText = rs(0).Value
        dcRuanganTujuan.Text = rs(1).Value
        KI = rs(2).Value
        tempKdRuanganTujuan = rs(0).Value
        
'        dcJenisObat.BoundText = ""


    strSQL = "SELECT JenisHargaNetto" & _
        " From PersentaseUpTarifOA" & _
        " Where(IdPenjamin = '" & mstrKdPenjaminPasien & "') And (KdKelompokPasien = '" & mstrKdJenisPasien & "')"
    Call msubRecFO(rs, strSQL)
    subJenisHargaNetto = IIf(rs.EOF = True, 1, rs(0))
   
    
    If dcRuanganTujuan.MatchedWithList = True Then fgData.SetFocus
    
'    Call clearRincianDetail
'fgData.clear
'Call subSetGrid
'Call subLoadDcSource
                    
    fgData.SetFocus
    fgData.Col = 2
    
   End If
Exit Sub
hell:
    Call msubPesanError
End Sub

'add by gantri

Private Sub dcRuanganTujuan_LostFocus()
On Error GoTo hell
        If dcRuanganTujuan.MatchedWithList = True Then fgData.SetFocus
        'Call clearRincianDetail:
        strSQL = "Select KdRuangan, NamaRuangan, KdInstalasi FROM Ruangan WHERE (NamaRuangan LIKE '%" & dcRuanganTujuan.Text & "%') And kdruangan not in ('701') AND StatusEnabled=1"
        Call msubRecFO(rs, strSQL)

        If rs.EOF = True Then dcRuanganTujuan.Text = "": Exit Sub
        dcRuanganTujuan.BoundText = rs(0).Value
        dcRuanganTujuan.Text = rs(1).Value
        KI = rs(2).Value
        tempKdRuanganTujuan = rs(0).Value
        Exit Sub
hell:
    Call msubPesanError
End Sub

'end by gantri

Public Sub dgDokter_DblClick()
On Error GoTo errLoad
    If dgDokter.ApproxCount = 0 Then Exit Sub
    txtDokter.Text = dgDokter.Columns("Nama Dokter")
    dgDokter.Visible = False
    txtKdDokter.Text = dgDokter.Columns("KodeDokter")
    'fgData.SetFocus
    dcRuanganTujuan.SetFocus
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call dgDokter_DblClick
End Sub

Private Sub dgObatAlkes_DblClick()
On Error GoTo errLoad
Dim i As Integer
Dim tempSettingDataPendukung As Integer
Dim curHargaBrg As Currency
Dim strNoTerima As String

    If fgData.TextMatrix(fgData.Row, 2) <> dgObatAlkes.Columns("KdBarang") And fgData.TextMatrix(fgData.Row, 10) <> "" Then
        MsgBox "Data tidak bisa diubah dengan data yang lain", vbExclamation, "Validasi"
        fgData.SetFocus
        dgObatAlkes.Visible = False
        Exit Sub
    End If

    curHutangPenjamin = 0
    curTanggunganRS = 0
    strNoTerima = ""
    Set rsB = Nothing
    Call msubRecFO(rsB, "select dbo.TakeNoFIFO_F('" & dgObatAlkes.Columns("KdBarang") & "','" & dgObatAlkes.Columns("KdAsal") & "','" & dcRuanganTujuan.BoundText & "') as NoFIFO")
    strNoTerima = IIf(IsNull(rsB("NoFIFO")), "0000000000", rsB("NoFIFO"))
    
    For i = 0 To fgData.Rows - 1
        If dgObatAlkes.Columns("KdBarang") = fgData.TextMatrix(i, 2) And dgObatAlkes.Columns("KdAsal") = fgData.TextMatrix(i, 12) Then
            MsgBox "Data tersebut sudah diinput", vbExclamation, "Validasi"
            dgObatAlkes.Visible = False
            fgData.SetFocus: fgData.Row = i
            Exit Sub
        End If
    Next i
    With fgData
        .TextMatrix(.Row, 2) = dgObatAlkes.Columns("KdBarang")
        .TextMatrix(.Row, 3) = dgObatAlkes.Columns("Nama Barang")
        .TextMatrix(.Row, 4) = dgObatAlkes.Columns("Kekuatan")
        .TextMatrix(.Row, 5) = dgObatAlkes.Columns("AsalBarang")
        .TextMatrix(.Row, 6) = dgObatAlkes.Columns("SatuanJml")
        '.TextMatrix(.Row, 7) = Format(dgObatAlkes.Columns("HargaBarang").Value, "#,###")
        .TextMatrix(.Row, 33) = strNoTerima
        curHargaBrg = 0
        
        strSQL = ""
        Set rsB = Nothing
'        strSQL = "SELECT dbo.FB_TakeHargaNettoOA('" & mstrKdPenjaminPasien & "','" & mstrKdJenisPasien & "','" & dgObatAlkes.Columns("KdBarang") & "','" & dgObatAlkes.Columns("KdAsal") & "','" & dgObatAlkes.Columns("Satuan") & "', '" & mstrKdRuangan & "','" & .TextMatrix(.Row, 33) & "') AS HargaBarang"
'        strSQL = "SELECT dbo.FB_TakeHargaNettoOA('" & mstrKdPenjaminPasien & "','" & mstrKdJenisPasien & "','" & dgObatAlkes.Columns("KdBarang") & "','" & dgObatAlkes.Columns("KdAsal") & "','" & dgObatAlkes.Columns("Satuan") & "', '" & dcRuanganTujuan.BoundText & "','" & .TextMatrix(.Row, 33) & "') AS HargaBarang"
        strSQL = "SELECT dbo.FB_TakeHargaNettoObatAlkesFifo('" & mstrKdPenjaminPasien & "','" & mstrKdJenisPasien & "','" & dgObatAlkes.Columns("KdBarang") & "','" & dgObatAlkes.Columns("KdAsal") & "','" & dgObatAlkes.Columns("Satuan") & "', '" & dcRuanganTujuan.BoundText & "','" & .TextMatrix(.Row, 33) & "') AS HargaBarang"
        
        Call msubRecFO(rsB, strSQL)
        If rsB.EOF = True Then curHargaBrg = 0 Else curHargaBrg = rsB(0).Value
        strSQL = ""
        Set rs = Nothing
        'subcurHargaSatuan = Format(dgObatAlkes.Columns("HargaBarang").Value, "#,###")
        strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "','" & dgObatAlkes.Columns("KdAsal") & "', " & msubKonversiKomaTitik(CStr(curHargaBrg)) & ")  as HargaSatuan"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then subcurHargaSatuan = 0 Else subcurHargaSatuan = rs(0).Value
        .TextMatrix(.Row, 7) = subcurHargaSatuan
        If subcurHargaSatuan = 0 Then
            .TextMatrix(.Row, 7) = 0
        Else
            .TextMatrix(.Row, 7) = IIf(Format(subcurHargaSatuan, "#,###") = "", 0, Format(subcurHargaSatuan, "#,###"))
        End If
        .TextMatrix(.Row, 8) = 0 '(dgObatAlkes.Columns("Discount").Value / 100) * CDbl(.TextMatrix(.Row, 7))
        
                    'khusus OA harga tidak dikalikan lg Ppn krn OA termasuk pelayanan yg include ke tindakan (TM)
    
        Call msubRecFO(rs, "select dbo.FB_TakeStokBrgMedis('" & dcRuanganTujuan.BoundText & "', '" & dgObatAlkes.Columns("KdBarang") & "','" & dgObatAlkes.Columns("KdAsal") & "') as stok")
        .TextMatrix(.Row, 9) = IIf(IsNull(rs("Stok")), 0, rs("Stok"))

'        Call msubRecFO(rs, "select dbo.FB_TakeStokBrgMedis('" & dcRuanganTujuan.BoundText & "', '" & dgObatAlkes.Columns("KdBarang") & "','" & dgObatAlkes.Columns("KdAsal") & "') as stok")
'        .TextMatrix(.Row, 9) = IIf(IsNull(rs("Stok")), 0, rs("Stok"))
        
        .TextMatrix(.Row, 12) = dgObatAlkes.Columns("KdAsal")
        .TextMatrix(.Row, 13) = dgObatAlkes.Columns("Jenis Barang")
'        .TextMatrix(.Row, 14) = subcurTarifService
'        .TextMatrix(.Row, 15) = subintJmlService
        .TextMatrix(.Row, 16) = CDbl(.TextMatrix(.Row, 7))
        .TextMatrix(.Row, 17) = curHutangPenjamin
        .TextMatrix(.Row, 18) = curTanggunganRS
        .TextMatrix(.Row, 19) = 0
        .TextMatrix(.Row, 20) = 0
        .TextMatrix(.Row, 21) = 0
        
        .TextMatrix(.Row, 23) = txtNoTemporary.Text
        txtHargaBeli.Text = curHargaBrg 'dgObatAlkes.Columns("HargaBarang")
        .TextMatrix(.Row, 24) = CDbl(txtHargaBeli.Text)
        .TextMatrix(.Row, 26) = 0
        
    End With
           
    dgObatAlkes.Visible = False
    txtJenisBarang.Text = "": txtKdBarang.Text = "": txtKdAsal.Text = "": txtSatuan.Text = "": txtAsalBarang.Text = "": 'txtKekuatan.Text = ""
    
    With fgData
        .SetFocus
        If .Col = 2 Then
            .Col = 3
        ElseIf .Col = 3 Then
            .Col = 10
        End If
    End With
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dgObatAlkes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call dgObatAlkes_DblClick
End Sub

Private Sub dgObatAlkesRacikan_DblClick()

On Error GoTo errLoad
Dim i As Integer
Dim tempSettingDataPendukung As Integer
Dim curHargaBrg As Currency
Dim Resep As String
    If dgObatAlkesRacikan.Columns("JmlStok") = "" Or dgObatAlkesRacikan.Columns("JmlStok") = "0" Then
        MsgBox "Jumlah Stok" & dgObatAlkesRacikan.Columns("NamaBarang") & " Kosong atau Nol ", vbExclamation, "Validasi"
        Exit Sub
    End If
   
    For i = 0 To fgRacikan.Rows - 1
        If dgObatAlkesRacikan.Columns("KdBarang") = fgRacikan.TextMatrix(i, 0) And dgObatAlkesRacikan.Columns("KdAsal") = fgRacikan.TextMatrix(i, 12) Then
            MsgBox "Data tersebut sudah diinput", vbExclamation, "Validasi"
            dgObatAlkes.Visible = False
            fgRacikan.SetFocus: fgRacikan.Row = i
            Exit Sub
        End If
    Next i

    strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdJenisPasien = rs("KdKelompokPasien").Value
        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
    End If
    
    curHutangPenjamin = 0
    curTanggunganRS = 0

    With fgRacikan
        .TextMatrix(.Row, 0) = dgObatAlkesRacikan.Columns("KdBarang")
        .TextMatrix(.Row, 3) = dgObatAlkesRacikan.Columns("NamaBarang")
        .TextMatrix(.Row, 11) = dgObatAlkesRacikan.Columns("AsalBarang")
        .TextMatrix(.Row, 12) = dgObatAlkesRacikan.Columns("kdAsal")
        .TextMatrix(.Row, 13) = dgObatAlkesRacikan.Columns("Satuan")
        
        Set rsB = Nothing
'        Call msubRecFO(rsB, "select dbo.TakeNoFIFO_F('" & dgObatAlkesRacikan.Columns("KdBarang") & "','" & dgObatAlkesRacikan.Columns("KdAsal") & "','" & mstrKdRuangan & "') as NoFIFO")
        Call msubRecFO(rsB, "select dbo.TakeNoFIFO_F('" & dgObatAlkesRacikan.Columns("KdBarang") & "','" & dgObatAlkesRacikan.Columns("KdAsal") & "','" & dcRuanganTujuan.BoundText & "') as NoFIFO")
        .TextMatrix(.Row, 14) = IIf(IsNull(rsB("NoFIFO")), "0000000000", rsB("NoFIFO"))
        
        .TextMatrix(.Row, 15) = subintJmlService
        .TextMatrix(.Row, 16) = subcurTarifService
        
        .ColAlignment(3) = flexAlignCenterCenter
        .TextMatrix(.Row, 4) = ""
        .TextMatrix(.Row, 5) = ""
        .TextMatrix(.Row, 6) = ""
        .TextMatrix(.Row, 2) = .TextMatrix(1, 2)
        curHargaBrg = 0
        strSQL = ""
        Set rsB = Nothing
'        strSQL = "SELECT dbo.FB_TakeHargaNettoOA('" & mstrKdPenjaminPasien & "','" & mstrKdJenisPasien & "','" & dgObatAlkesRacikan.Columns("KdBarang") & "','" & dgObatAlkesRacikan.Columns("KdAsal") & "','" & dgObatAlkesRacikan.Columns("Satuan") & "', '" & mstrKdRuangan & "','" & .TextMatrix(.Row, 14) & "') AS HargaBarang"
        strSQL = "SELECT dbo.FB_TakeHargaNettoOA('" & mstrKdPenjaminPasien & "','" & mstrKdJenisPasien & "','" & dgObatAlkesRacikan.Columns("KdBarang") & "','" & dgObatAlkesRacikan.Columns("KdAsal") & "','" & dgObatAlkesRacikan.Columns("Satuan") & "', '" & dcRuanganTujuan.BoundText & "','" & .TextMatrix(.Row, 14) & "') AS HargaBarang"
        Call msubRecFO(rsB, strSQL)
        If rsB.EOF = True Then curHargaBrg = 0 Else curHargaBrg = rsB(0).Value
        strSQL = ""
        Set rs = Nothing
        subcurHargaSatuan = 0
        strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "', '" & mstrKdPenjaminPasien & "', '" & dgObatAlkesRacikan.Columns("KdAsal") & "', " & msubKonversiKomaTitik(CStr(curHargaBrg)) & " )  as HargaSatuan "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then subcurHargaSatuan = 0 Else subcurHargaSatuan = rs(0).Value
        .TextMatrix(.Row, 8) = subcurHargaSatuan
        '.TextMatrix(.Row, 8) = Format(subcurHargaSatuan, "#,###")
'        Call msubRecFO(rs, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & dgObatAlkesRacikan.Columns("KdBarang") & "','" & dgObatAlkesRacikan.Columns("KdAsal") & "') as stok")
'        .TextMatrix(.Row, 9) = IIf(IsNull(rs("Stok")), 0, rs("Stok"))
        .TextMatrix(.Row, 10) = dgObatAlkesRacikan.Columns("Kekuatan")
        If dgObatAlkesRacikan.Columns("JmlStok") <> "" Then
            .TextMatrix(.Row, 17) = dgObatAlkesRacikan.Columns("JmlStok")
        Else
            .TextMatrix(.Row, 17) = "0"
        End If
    End With

    For i = 0 To fgRacikan.Rows - 1
        If dgObatAlkesRacikan.Columns("KdBarang") = fgRacikan.TextMatrix(i, 2) Then
            MsgBox "Data tersebut sudah diinput", vbExclamation, "Validasi"
            dgObatAlkesRacikan.Visible = False
            fgRacikan.SetFocus: fgData.Row = i
            Exit Sub
        End If
    Next i
    
    dgObatAlkesRacikan.Visible = False
    With fgRacikan
                '.Rows = .Rows + 1
         .SetFocus
         .Col = 4
    End With
    
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dgObatAlkesRacikan_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then Call dgObatAlkesRacikan_DblClick
    With fgRacikan
        .SetFocus
        .Col = 4
    End With
End Sub

Private Sub dtpTglOrder_Change()
    dtpTglOrder.MaxDate = Now
End Sub

Private Sub dtpTglOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then fgData.SetFocus
End Sub

Private Sub dtpTglResep_Change()
    dtpTglResep.MaxDate = Now
End Sub

Private Sub dtpTglResep_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then chkDokterPemeriksa.SetFocus
End Sub

Private Sub dtpTglResep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkDokterPemeriksa.SetFocus
End Sub

Private Sub fgData_DblClick()
    Call fgData_KeyDown(13, 0)
End Sub

Private Sub fgData_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
    Select Case KeyCode
        Case 13
            If fgData.Col = fgData.Cols - 1 Then
                If fgData.TextMatrix(fgData.Row, 2) <> "" And fgData.TextMatrix(fgData.Row, 10) <> "" Then
                    If fgData.TextMatrix(fgData.Rows - 1, 2) <> "" And fgData.TextMatrix(fgData.Row, 10) <> "" Then
                        fgData.Rows = fgData.Rows + 1
                        If fgData.TextMatrix(fgData.Rows - 2, 25) = "" Then
                            fgData.TextMatrix(fgData.Rows - 1, 0) = "1"
'                        ElseIf fgData.TextMatrix(fgData.Rows - 2, 25) = "01" Then
'                            fgData.TextMatrix(fgData.Rows - 1, 0) = "0"
'                        Else
'                            fgData.TextMatrix(fgData.Rows - 1, 0) = Val(fgData.TextMatrix(fgData.Rows - 2, 0))
                        End If
                    End If
                    fgData.Row = fgData.Rows - 1
                    fgData.Col = 1
                Else
                    If fgData.TextMatrix(fgData.Row, 2) = "" Then
                        fgData.Col = 1
                     ElseIf fgData.TextMatrix(fgData.Row, 10) = "" Then
                        fgData.Col = 10
                    End If
                End If
            Else
            
                For i = 0 To fgData.Cols - 2
                    If fgData.Col = fgData.Cols - 1 Then Exit For
                    fgData.Col = fgData.Col + 1
                    If fgData.ColWidth(fgData.Col) > 0 Then Exit For
                Next i
            End If
            fgData.SetFocus
            If fgData.Col = 1 Then Call subLoadDataCombo(dcJenisObat)
            
        Case 27
            dcJenisObat.Visible = False
            dgObatAlkes.Visible = False
            
        Case vbKeyDelete
            
                Call subHapusDataGrid
                
            
     End Select
End Sub

Private Sub fgData_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    txtIsi.Text = ""
    If Not (KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ Or KeyAscii = 32 Or KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace Or KeyAscii = Asc(",")) Then
        KeyAscii = 0
        Exit Sub
    End If
            
    Select Case fgData.Col
        Case 0 'R/Ke
            txtIsi.MaxLength = 2
            Call subLoadText
            txtIsi.Text = Chr(KeyAscii)
            txtIsi.SelStart = Len(txtIsi.Text)
        
        Case 1 'Jenis Obat
                If fgData.TextMatrix(fgData.Row, 3) = "Racikan" Then Exit Sub

                fgData.Col = 1
                Call subLoadDataCombo(dcJenisObat)
                If KI <> "07" Then
                    dcJenisObat.Visible = True
                End If
            
        Case 2 'Kode Barang
            txtIsi.MaxLength = 9
            Call subLoadText
            txtIsi.Text = Chr(KeyAscii)
            txtIsi.SelStart = Len(txtIsi.Text)
        
        Case 3 'Nama Barang
           If fgData.TextMatrix(fgData.Row, 1) <> "Racikan" And fgData.TextMatrix(fgData.Row, 3) = "" Then
                txtIsi.MaxLength = 20
                Call subLoadText
                txtIsi.Text = Chr(KeyAscii)
                txtIsi.SelStart = Len(txtIsi.Text)
           ElseIf fgData.TextMatrix(fgData.Row, 1) <> "Racikan" And fgData.TextMatrix(fgData.Row, 3) <> "Racikan" Then
                 txtIsi.MaxLength = 20
                Call subLoadText
                txtIsi.Text = Chr(KeyAscii)
                txtIsi.SelStart = Len(txtIsi.Text)
               
           ElseIf fgData.TextMatrix(fgData.Row, 1) = "Racikan" And fgData.TextMatrix(fgData.Row, 3) = "Racikan" Then
            
            Else
                Call subSetGridRacikan
                FraRacikan.Visible = True
                dgObatAlkesRacikan.Visible = False
                txtJumlahObatRacik.SetFocus
                txtJumlahObatRacik.Text = ""
                cmdSimpan.Enabled = False
                cmdTutup.Enabled = False
                fgRacikan.TextMatrix(fgRacikan.Row, 2) = fgData.TextMatrix(fgData.Row, 0) ' edit
                
                frmDataResep.Enabled = False

            End If
 
        
        Case 10 'Jumlah
            
            If fgData.TextMatrix(fgData.Row, 3) = "Racikan" Then Exit Sub
            Call SetKeyPressToNumber(KeyAscii)
            
            txtIsi.MaxLength = 7
            If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyEscape Or KeyAscii = Asc(",")) Then Exit Sub
            Call subLoadText
            txtIsi.Text = Chr(KeyAscii)
            txtIsi.SelStart = Len(txtIsi.Text)
    
        Case 27 'Aturan Pakai
            fgData.Col = 27
            Call subLoadDataCombo(dcAturanPakai)
            
        Case 28 'Keterangan Pakai
            fgData.Col = 28
            Call subLoadDataCombo(dcKeteranganPakai)
            
        Case 29 'Keterangan Pakai2
            fgData.Col = 29
            Call subLoadDataCombo(dcKeteranganPakai2)
            
        Case 30 'Keterangan Lainnya
            txtIsi.MaxLength = 200
            Call subLoadText
            txtIsi.Text = Chr(KeyAscii)
            txtIsi.SelStart = Len(txtIsi.Text)
            
        Case 31
            fgData.Col = 31
            Call subLoadDataCombo(dcNamaPelayananRS)
            
    End Select
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub fgRacikan_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
Dim X As Integer
    Select Case KeyCode
        Case 13
            If fgRacikan.Col = fgRacikan.Cols - 1 Then
                If fgRacikan.TextMatrix(fgRacikan.Row, 4) <> "" Then
                    If fgRacikan.TextMatrix(fgRacikan.Rows - 1, 4) <> "" Then
                        fgRacikan.Rows = fgRacikan.Rows + 1
                    End If
                    fgRacikan.Row = fgRacikan.Rows - 1
                    fgRacikan.Col = 1
                Else
                    fgRacikan.Col = 1
                End If
            Else
                For i = 0 To fgRacikan.Cols - 1
                    If fgRacikan.Col = fgRacikan.Cols - 1 Then Exit For
                        fgRacikan.Col = fgRacikan.Col + 1
                    If fgRacikan.ColWidth(fgRacikan.Col) > 0 Then Exit For
                Next i

            End If
            
            fgRacikan.TextMatrix(fgRacikan.Row, 2) = fgRacikan.TextMatrix(1, 2)
            fgRacikan.SetFocus
            
        Case 27
            dgObatAlkesRacikan.Visible = False
        
        Case vbKeyDelete
            If fgRacikan.Rows - 1 = 1 Then
                For i = 0 To fgRacikan.Cols - 1
                    fgRacikan.TextMatrix(1, i) = ""
                Next i
                If fgRacikan.TextMatrix(fgRacikan.Row, 0) = "" Then FraRacikan.Visible = False
                fgData.SetFocus
                fgData.Col = 1
                Call subLoadDataCombo(dcJenisObat)
            Else
                fgRacikan.RemoveItem fgRacikan.Row
'                If fgRacikan.TextMatrix(fgRacikan.Row, 0) = "" Then FraRacikan.Visible = False
            End If
     End Select
End Sub

Private Sub fgRacikan_KeyPress(KeyAscii As Integer)
 On Error GoTo errLoad
 
    If txtJumlahObatRacik.Text = "" Or txtJumlahObatRacik.Text = "0" Then MsgBox "Jumlah Racikan Masih Kosong atau Nol.", vbCritical, "Medifirst2000 - Validasi": txtJumlahObatRacik.SetFocus: Exit Sub

    Select Case fgRacikan.Col
        Case 4 'Kebutuhan /ML
            Call SetKeyPressToNumber(KeyAscii)
        Case 5 'Kebutuhan /Tablet
            Call SetKeyPressToNumber(KeyAscii)
        Case 6 'Jumlah
            Call SetKeyPressToNumber(KeyAscii)
    End Select
    
    TxtIsiRacikan.Text = ""
    If Not (KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ Or KeyAscii = 32 Or KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace Or KeyAscii = Asc(".")) Then
        KeyAscii = 0
        Exit Sub
    End If
            
    Select Case fgRacikan.Col
            
        Case 3 'Nama Barang
            TxtIsiRacikan.MaxLength = 20
            Call subLoadTextRacikan
            TxtIsiRacikan.Text = Chr(KeyAscii)
            TxtIsiRacikan.SelStart = Len(TxtIsiRacikan.Text)
'            fgRacikan.Rows = fgRacikan.Rows + 1
        
        Case 4 'Kebutuhan /ML

            TxtIsiRacikan.MaxLength = 4
            If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyEscape Or KeyAscii = Asc(".")) Then Exit Sub
            Call subLoadTextRacikan
            TxtIsiRacikan.Text = Chr(KeyAscii)
            TxtIsiRacikan.SelStart = Len(TxtIsiRacikan.Text)
            
        Case 5 'Kebutuhan /Tablet
            TxtIsiRacikan.MaxLength = 4
            If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyEscape Or KeyAscii = Asc(".")) Then Exit Sub
            Call subLoadTextRacikan
            TxtIsiRacikan.Text = Chr(KeyAscii)
            TxtIsiRacikan.SelStart = Len(TxtIsiRacikan.Text)
        Case 6 'Jumlah
            TxtIsiRacikan.MaxLength = 4
            If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyEscape Or KeyAscii = Asc(".")) Then Exit Sub
            Call subLoadTextRacikan
            TxtIsiRacikan.Text = Chr(KeyAscii)
            TxtIsiRacikan.SelStart = Len(TxtIsiRacikan.Text)
            
         Case 7 'Jumlah Pembulatan
            TxtIsiRacikan.MaxLength = 4
            If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyEscape Or KeyAscii = Asc(".")) Then Exit Sub
            Call subLoadTextRacikan
            TxtIsiRacikan.Text = Chr(KeyAscii)
            TxtIsiRacikan.SelStart = Len(TxtIsiRacikan.Text)
    End Select
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Activate()
    txtDokter.Enabled = True
    txtDokter.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF5
            If cmdSimpan.Enabled = False Then frmDaftarBarangGratisRuangan.Show
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
Dim curHargaBrg As Currency
Dim i, j As Integer
Dim tempSettingDataPendukung As Integer
Dim curHarusDibayar As Currency
'dcJenisObat.Visible = True
    
 
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpTglOrder.Value = Now
    dtpTglResep.Value = Now
    
    strSQL11 = "SELECT Value FROM dbo.SettingGlobal WHERE Prefix='KdInstalasiApotek'"
    Call msubRecFO(rsC, strSQL11)
    TempMstrKdIsntalasi = IIf(rsC.EOF = True, "07", rsC(0))
    
    
    Call subSetGrid
    Call subLoadDcSource
    dgDokter.Visible = False
    dcJenisObat.BoundText = ""
    
    dgObatAlkes.Top = 2880
    dgObatAlkes.Left = 2040
    dgObatAlkes.Visible = True
    dgObatAlkes.Visible = False
    
     
     Call subSetGridRacikan
     Cancel = False
    'semua pasien yang masuk UDD bayar tunai. kd penjamin dan kd jenispasien dikembalikan saat form unload
'    mstrKdJenisPasien = "01"
'    mstrKdPenjaminPasien = "2222222222"
    
    strSQL = "SELECT JenisHargaNetto" & _
        " From PersentaseUpTarifOA" & _
        " Where(IdPenjamin = '" & mstrKdPenjaminPasien & "') And (KdKelompokPasien = '" & mstrKdJenisPasien & "')"
    Call msubRecFO(rs, strSQL)
    subJenisHargaNetto = IIf(rs.EOF = True, 1, rs(0))
    
    
'    dcRuanganTujuan.Text = ""
'    dcRuanganTujuan.BoundText = ""
'   add by gantri : untuk membaca kode instalasi pada pilihan ruangan tujuan default pada awal load form
'   If dcRuanganTujuan.MatchedWithList = True Then fgData.SetFocus
        
'        strSQL = "Select KdRuangan, NamaRuangan, KdInstalasi FROM Ruangan WHERE (NamaRuangan LIKE '%" & dcRuanganTujuan.Text & "%') AND StatusEnabled=1"
'        Call msubRecFO(rs, strSQL)
'
'        If Not rs.EOF = True Then Exit Sub
'        dcRuanganTujuan.BoundText = rs(0).Value
'        dcRuanganTujuan.Text = rs(1).Value
'        KI = rs(2).Value
'   txtDokter.SetFocus
'   Call subLoadResep
'    txtDokter.Enabled = True
'   txtDokter.Enabled = False
   
'   dcRuanganTujuan.BoundText = "702"
'   Call dcRuanganTujuan_KeyPress(13)
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errLoad
    strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdJenisPasien = rs("KdKelompokPasien").Value
        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
    End If
    dbConn.Execute "DELETE FROM TempDetailApotikJual WHERE (NoTemporary = '" & txtNoTemporary & "')"
    dbConn.Execute "DELETE FROM DetailOrderPelayananOARacikanTemp"
'    frmDaftarPasienRI.Enabled = True
If txtNamaForm.Text = "frmTransaksiPasien" Then frmTransaksiPasien.Enabled = True
    
errLoad:
End Sub

Private Sub subHapusDataGrid()
On Error GoTo errLoad
Dim i As Integer
Dim strResepKe As String
Dim intBarisYangDihapus As Integer
Dim curHarusDibayar As Currency

    With fgData
        If .Row = 0 Then Exit Sub
        If Val(.TextMatrix(.Row, 11)) = 0 Then GoTo stepHapusData
        intBarisYangDihapus = fgData.Row
'        If .TextMatrix(.Row, 11) <> "01" Then 'jika obat racikan, pastikan jumlah service 1 untuk resep yang sama
        If .TextMatrix(.Row, 25) <> "01" Then 'jika obat racikan, pastikan jumlah service 1 untuk resep yang sama
            strResepKe = .TextMatrix(.Row, 0)
            If Val(.TextMatrix(.Row, 15)) = 0 Then GoTo stepHapusData
            For i = 1 To .Rows - 2
                If .TextMatrix(i, 0) = strResepKe And i <> intBarisYangDihapus Then
                    .TextMatrix(i, 13) = 1
                    Exit For
                End If
            Next i
        End If
        
stepHapusData:
        'add by onede
       ' If sp_StokRuangan(.TextMatrix(.Row, 2), .TextMatrix(.Row, 12), CDbl(.TextMatrix(.Row, 10)), "A") = False Then Exit Sub
        
        dbConn.Execute "DELETE FROM TempDetailApotikJual " & _
            " WHERE (NoTemporary = '" & Trim(.TextMatrix(.Row, 23)) & "') AND (KdBarang = '" & .TextMatrix(.Row, 2) & "') AND (KdAsal = '" & .TextMatrix(.Row, 12) & "')"
        
        If fgData.TextMatrix(fgData.Row, 25) = "01" Then
            dbConn.Execute "DELETE FROM DetailOrderPelayananOARacikanTemp where NoRacikan = '" & fgData.TextMatrix(fgData.Row, 34) & "'"
            
        End If
        
        If .Rows = 2 Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
            Next i
        Else
            .RemoveItem .Row
        End If
        
        If .TextMatrix(.Row, 2) <> "" Then
            'total harga = ((tarifservice * jmlservice) + _
                (hargasatuan(sebelum ditambah tarifservixe) * jumlah))
            .TextMatrix(.Row, 11) = ((CDbl(.TextMatrix(.Row, 14)) * CDbl(.TextMatrix(.Row, 15))) + _
                (CDbl(.TextMatrix(.Row, 16)) * Val(.TextMatrix(.Row, 10))))
            
            'total harus dibayar = total harga - total discount - _
                total hutang penjamin - totaltanggunganrs
            curHarusDibayar = CDbl(.TextMatrix(.Row, 11)) - (CDbl(.TextMatrix(.Row, 21)) + _
                CDbl(.TextMatrix(.Row, 19)) + CDbl(.TextMatrix(.Row, 20)))
            .TextMatrix(.Row, 20) = IIf(curHarusDibayar < 0, 0, curHarusDibayar)
        End If
    End With
    Call subHitungTotal
'    Call rebuildResekKe
Exit Sub
errLoad:
    Call msubPesanError
'    Resume 0
End Sub



Private Sub txtBeratObat_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyEscape Then
        fraHitungObat.Visible = False
        fgRacikan.SetFocus
    End If
End Sub

Private Sub txtBeratObat_LostFocus()
    fraHitungObat.Visible = False
End Sub

'Private Sub Timer1_Timer()
'' dtpTglOrder.Value = Now
'End Sub

Private Sub txtDokter_Change()
On Error GoTo errLoad
    mstrFilterDokter = "WHERE NamaDokter like '%" & txtDokter.Text & "%'"
    txtKdDokter.Text = ""
    dgDokter.Visible = True
    Call subLoadDokter
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtDokter_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13, vbKeyDown
            If dgDokter.Visible = True Then dgDokter.SetFocus Else dcRuanganTujuan.SetFocus
        Case vbKeyEscape
            dgDokter.Visible = False
    End Select
End Sub

Private Sub txtDokter_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtIsi_Change()
Dim i As Integer
    Select Case fgData.Col
        Case 2 'kode barang
      
            If tempStatusTampil = True Then Exit Sub
             'CariBarangNStokMedis_V
'            strSQL = "execute CariBarang_V '" & txtIsi.Text & "%','" & dcRuanganTujuan.BoundText & "'"
            strSQL = "execute CariBarangToAsalBarang_V'" & txtIsi.Text & "%','" & dcRuanganTujuan.BoundText & "','" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "'"
            Call msubRecFO(dbRst, strSQL)
            
            If dcRuanganTujuan.Text = "" Then
            MsgBox "Silahkan pilih ruangan tujuan", vbInformation + vbOKOnly, "Info"
            dcRuanganTujuan.SetFocus
            dcRuanganTujuan.Text = ""
            Exit Sub
            End If
    
            
            
            Set dgObatAlkes.DataSource = dbRst
            With dgObatAlkes
                For i = 0 To .Columns.Count - 1
                    .Columns(i).Width = 0
                Next i
                
                .Columns("KdBarang").Width = 0
                .Columns("Nama Barang").Width = 3000
                .Columns("Jenis Barang").Width = 1500
                .Columns("Kekuatan").Width = 1000
                .Columns("AsalBarang").Width = 1000
                .Columns("Satuan").Width = 675
                
                .Top = txtIsi.Top + txtIsi.Height + Frame8.Top
                .Left = 1820
                .Visible = True
                
            End With
                    
        Case 3 ' nama barang
        
        
            If tempStatusTampil = True Then Exit Sub
'            strSQL = "execute CariBarang_V '" & txtIsi.Text & "%','" & dcRuanganTujuan.BoundText & "'"
            strSQL = "execute CariBarangToAsalBarang_V'" & txtIsi.Text & "%','" & dcRuanganTujuan.BoundText & "','" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "'"
            Call msubRecFO(dbRst, strSQL)
            
           If dcRuanganTujuan.Text = "" Then
            MsgBox "Silahkan pilih ruangan tujuan", vbInformation + vbOKOnly, "Info"
            dcRuanganTujuan.SetFocus
            dcRuanganTujuan.Text = ""
            Exit Sub
            End If

            
            Set dgObatAlkes.DataSource = dbRst
            With dgObatAlkes
                For i = 0 To .Columns.Count - 1
                    .Columns(i).Width = 0
                Next i

                .Columns("KdBarang").Width = 0
                .Columns("Nama Barang").Width = 3000
                .Columns("Jenis Barang").Width = 1500
                .Columns("Kekuatan").Width = 1000
                .Columns("AsalBarang").Width = 1000
                .Columns("Satuan").Width = 675
                
                .Top = txtIsi.Top + txtIsi.Height + Frame8.Top
                .Left = 3000
                .Visible = True
                
            End With
        Case Else
            dgObatAlkes.Visible = False
    End Select
End Sub

Private Sub txtIsi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then If dgObatAlkes.Visible = True Then If dgObatAlkes.ApproxCount > 0 Then dgObatAlkes.SetFocus
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
Dim i As Integer
Dim curHutangPenjamin As Currency
Dim curTanggunganRS As Currency
Dim curHarusDibayar As Currency
Dim KdJnsObat As String

''Add by Azein
'Dim j As Integer
'Dim curDiskon As Currency
'Dim bolNilaiDec As Boolean
'Dim dblSelisih As Double
'Dim intRowTemp As Integer
'Dim strNoTerima As String
'Dim curHargaBrg As Currency
'Dim dblSelisihNow As Double
'Dim dblJmlStokMax As Double
'Dim strKdBrg As String
'Dim strKdAsal As String
'Dim dblJmlTerkecil As Double
'Dim dblTotalStokK As Double
''end azein

    If KeyAscii = 13 Then
        With fgData
            Select Case .Col
                Case 0
                    dgObatAlkes.Visible = False
                    If Val(txtIsi.Text) = 0 Then txtIsi.Text = 1
                    .TextMatrix(.Row, .Col) = CDbl(txtIsi.Text)
                    txtIsi.Visible = False
                    'dcJenisObat.Visible = True
                    dcJenisObat.Left = 120
                    .Col = 1
                    For i = 0 To .Col - 1
                        dcJenisObat.Left = dcJenisObat.Left + .ColWidth(i)
                    Next i
                    dcJenisObat.Visible = False
                    dcJenisObat.Top = .Top - 7
                    
                    For i = 0 To .Row - 1
                        dcJenisObat.Top = dcJenisObat.Top + .RowHeight(i)
                    Next i
                    
                    If .TopRow > 1 Then
                        dcJenisObat.Top = dcJenisObat.Top - ((.TopRow - 1) * .RowHeight(1))
                    End If
                    
                    dcJenisObat.Width = .ColWidth(.Col)
                    dcJenisObat.Height = .RowHeight(.Row)
                    
                    dcJenisObat.Visible = False
'                    dcJenisObat.SetFocus
                    fgData.SetFocus
                    fgData.Col = 1

                Case 1
                    dgObatAlkes.Visible = False
                    If Val(txtIsi.Text) = 0 Then txtIsi.Text = 1
                    .TextMatrix(.Row, .Col) = CDbl(txtIsi.Text)
                    txtIsi.Visible = False
                    'dcJenisObat.Visible = True
                    dcJenisObat.Left = 120
                    .Col = 1
                    For i = 0 To .Col - 1
                        dcJenisObat.Left = dcJenisObat.Left + .ColWidth(i)
                    Next i
                    dcJenisObat.Visible = False
                    dcJenisObat.Top = .Top - 7
                    
                    For i = 0 To .Row - 1
                        dcJenisObat.Top = dcJenisObat.Top + .RowHeight(i)
                    Next i
                    
                    If .TopRow > 1 Then
                        dcJenisObat.Top = dcJenisObat.Top - ((.TopRow - 1) * .RowHeight(1))
                    End If
                    
                    dcJenisObat.Width = .ColWidth(.Col)
                    dcJenisObat.Height = .RowHeight(.Row)
                    
                    dcJenisObat.Visible = False
                    dcJenisObat.SetFocus

                Case 2
                    If dgObatAlkes.Visible = True Then
                        dgObatAlkes.SetFocus
                        Exit Sub
                    Else
                        fgData.SetFocus
                        fgData.Col = 8
                        dcJenisObat.Visible = False
                    End If
     
                Case 3
                    If dgObatAlkes.Visible = True Then
                        dgObatAlkes.SetFocus
                        Exit Sub
                    Else
                        fgData.SetFocus
                        fgData.Col = 8
                        dcJenisObat.Visible = False
                    End If
                
                Case 8
                    dgObatAlkes.Visible = False
                    txtIsi.Visible = False
                    
                    If mblnOperator = False Then
                        If Val(txtIsi.Text) = 0 Then txtIsi.Text = 0
                        
                        'konvert koma col discount
                        .TextMatrix(.Row, .Col) = Val(txtIsi.Text)
                        .TextMatrix(.Row, .Col) = .TextMatrix(.Row, .Col)
                        
                        'konvert koma col jumlah
                        .TextMatrix(.Row, 10) = Val(.TextMatrix(.Row, 10))
                        
                        If .TextMatrix(.Row, 10) <> "0" Then
                            .TextMatrix(.Row, 21) = IIf(Val(Val(.TextMatrix(.Row, 10))) = 0, 0, Val(.TextMatrix(.Row, 10))) * CDbl(.TextMatrix(.Row, 8))
                        
                            curHarusDibayar = CDbl(.TextMatrix(.Row, 11)) - (CDbl(.TextMatrix(.Row, 21)) + _
                                (CDbl(.TextMatrix(.Row, 19)) + CDbl(.TextMatrix(.Row, 20))))
                            .TextMatrix(.Row, 22) = IIf(curHarusDibayar < 0, 0, curHarusDibayar)
                            Call subHitungTotal
                        End If
                        
                    End If
                    
                    fgData.SetFocus
                    fgData.Col = 10

                Case 10
              If fgData.TextMatrix(fgData.Row, 2) = "" Then
                        MsgBox "Nama barang kosong", vbExclamation, "Validasi"
                        txtIsi.Visible = False
                        fgData.Col = 3
                        fgData.SetFocus
                        Exit Sub
                End If
                
                'End azein
                If Trim(txtIsi.Text) = "," Then txtIsi.Text = 0
                If Trim(txtIsi.Text) = "" Then txtIsi.Text = 0
                'Menambahan Azein
                If txtIsi.Text = 0 Then
                       MsgBox "Jumlah barang tidak boleh nol (0)", vbCritical
                       Exit Sub
                End If
                'End azein
'                If CDbl(txtIsi.Text) > CDbl(.TextMatrix(.Row, 9)) Then
'                            MsgBox "Jumlah lebih besar dari stock (" & .TextMatrix(.Row, 9) & ")", vbExclamation, "Validasi"
'                            txtIsi.SelStart = 0: txtIsi.SelLength = Len(txtIsi.Text)
'                            Exit Sub
'                End If
   
                    If CDbl(txtIsi.Text) <= 0 Then txtIsi.Text = 0
  
                    If (fgData.TextMatrix(.Row, 6) = "S") Then
'                     If bolStatusFIFO = False Then
                        If CDbl(txtIsi.Text) > CDbl(.TextMatrix(.Row, 9)) Then
                            MsgBox "Jumlah lebih besar dari stock (" & .TextMatrix(.Row, 9) & ")", vbExclamation, "Validasi"
                            txtIsi.SelStart = 0: txtIsi.SelLength = Len(txtIsi.Text)
                            Exit Sub
                        End If
'                     End If
                    ElseIf (fgData.TextMatrix(.Row, 6) = "K") Then
                        Dim dblJmlTerkecil As Double
                        Dim dblTotalStokK As Double

                        Set rs = Nothing
                        strSQL = "Select JmlTerkecil From MasterBarang Where KdBarang = '" & fgData.TextMatrix(.Row, 2) & "'"
                        Call msubRecFO(rs, strSQL)
                        dblJmlTerkecil = IIf(rs.EOF, 1, rs(0).Value)

                        dblTotalStokK = dblJmlTerkecil * fgData.TextMatrix(.Row, 9)
                        If Val(txtIsi.Text) > Val(dblTotalStokK) Then
                            MsgBox "Jumlah lebih besar dari stock (" & .TextMatrix(.Row, 9) & ")", vbExclamation, "Validasi"
                            txtIsi.SelStart = 0: txtIsi.SelLength = Len(txtIsi.Text)
                            Exit Sub
                        End If
                    End If
                    
                    'konvert koma col jumlah
                    .TextMatrix(.Row, .Col) = txtIsi.Text
                    
                    'konvert koma col discount
                    .TextMatrix(.Row, 8) = .TextMatrix(.Row, 8)

                    txtIsi.Visible = False
                    
'                    If .TextMatrix(.Row, 6) = "S" Then
'                        If sp_StokRuangan(.TextMatrix(.Row, 2), .TextMatrix(.Row, 12), .TextMatrix(.Row, 33), CDbl(.TextMatrix(.Row, 10)), "M") = False Then Exit Sub
'                    End If
'                    If dcJenisObat.Text = "01" Then
'                        subintJmlService = 1
'                    Else
                        'subintJmlService = 1
'                    End If
                    'rubah jumlah service
                    '.TextMatrix(.Row, 15) = subintJmlService
                    
                    'add by onede
         '           If sp_StokRuangan(.TextMatrix(.Row, 2), .TextMatrix(.Row, 12), CDbl(.TextMatrix(.Row, 10)), "M") = False Then Exit Sub
                    
                    'ambil no temporary
                    'If sp_TempDetailApotikJual(CDbl(.TextMatrix(.Row, 7)) + CDbl(.TextMatrix(.Row, 14))) = False Then Exit Sub
                    'ambil hutang penjamin dan tanggungan rs
                    strSQL = "SELECT HargaSatuan, JmlHutangPenjamin, JmlTanggunganRS" & _
                        " FROM TempDetailApotikJual" & _
                        " WHERE (NoTemporary = '" & Trim(txtNoTemporary.Text) & "') AND (KdBarang = '" & .TextMatrix(.Row, 2) & "') AND (KdAsal = '" & .TextMatrix(.Row, 12) & "')"
                    Call msubRecFO(rs, strSQL)
                    If rs.EOF = True Then
                        curHutangPenjamin = 0
                        curTanggunganRS = 0
                    Else
                        curHutangPenjamin = rs("JmlHutangPenjamin").Value
                        curTanggunganRS = rs("JmlTanggunganRS").Value
                    End If
                    
'                     If .TextMatrix(.Row, 6) = "S" Then
'                        If Trim(.TextMatrix(.Row, 10)) = "" Then .TextMatrix(.Row, 10) = 0
'                        If (.TextMatrix(.Row, 10)) <> 0 Then
'                            If sp_StokRuangan(.TextMatrix(.Row, 2), .TextMatrix(.Row, 12), .TextMatrix(.Row, 33), CDbl(.TextMatrix(.Row, 10)), "A") = False Then Exit Sub
'                        End If
'                    End If
                    
                    '.TextMatrix(.Row, 14) = subcurTarifService
                    .TextMatrix(.Row, 16) = CDbl(.TextMatrix(.Row, 7))
                    
                    'total harga = ((tarifservice * jmlservice) + _
                        (hargasatuan(sebelum ditambah tarifservixe) * jumlah))
                    .TextMatrix(.Row, 10) = IIf(.TextMatrix(.Row, 10) = "", 0, .TextMatrix(.Row, 10))
                    .TextMatrix(.Row, 14) = IIf(.TextMatrix(.Row, 14) = "", 0, .TextMatrix(.Row, 14))
                    .TextMatrix(.Row, 15) = IIf(.TextMatrix(.Row, 15) = "", 0, .TextMatrix(.Row, 15))
                    .TextMatrix(.Row, 16) = IIf(.TextMatrix(.Row, 16) = "", 0, .TextMatrix(.Row, 16))
                    .TextMatrix(.Row, 26) = IIf(.TextMatrix(.Row, 26) = "", 0, .TextMatrix(.Row, 26))
                    
'                    .TextMatrix(.Row, 11) = ((CDbl(.TextMatrix(.Row, 14)) * CDbl(.TextMatrix(.Row, 15))) + _
'                        (CDbl(.TextMatrix(.Row, 16)) * Val(.TextMatrix(.Row, 10)))) + Val(.TextMatrix(.Row, 26))
'                    .Col = 11: .CellForeColor = vbBlue: .CellFontBold = True: .Col = 10
                    curHarusDibayar = ((CDbl(.TextMatrix(.Row, 14) * .TextMatrix(.Row, 15)) + CDbl(.TextMatrix(.Row, 10) * CDbl(.TextMatrix(.Row, 16)))) + CDbl(.TextMatrix(.Row, 26))) - _
                                         ((CDbl(.TextMatrix(.Row, 8) / 100)) * (CDbl(.TextMatrix(.Row, 10)) * CDbl(.TextMatrix(.Row, 16))) + _
                                         (CDbl(.TextMatrix(.Row, 19)) + CDbl(.TextMatrix(.Row, 20))))
                        
                    .TextMatrix(.Row, 11) = IIf(curHarusDibayar < 0, 0, curHarusDibayar)
'                    .TextMatrix(.Row, 11) = IIf(.TextMatrix(.Row, 11) = "", 0, .TextMatrix(.Row, 11))
                    .TextMatrix(.Row, 17) = curHutangPenjamin
                    .TextMatrix(.Row, 18) = curTanggunganRS
                    
                    If curHutangPenjamin > 0 Then
                        .TextMatrix(.Row, 19) = (.TextMatrix(.Row, 14) * .TextMatrix(.Row, 15)) + (Val(.TextMatrix(.Row, 10)) * CDbl(.TextMatrix(.Row, 17))) + Val(.TextMatrix(.Row, 26))
                    Else
                        .TextMatrix(.Row, 19) = 0
                    End If
                    
                    If curTanggunganRS > 0 Then
                        .TextMatrix(.Row, 20) = (.TextMatrix(.Row, 14) * .TextMatrix(.Row, 15)) + (Val(.TextMatrix(.Row, 10)) * CDbl(.TextMatrix(.Row, 18))) + Val(.TextMatrix(.Row, 26))
                    Else
                        .TextMatrix(.Row, 20) = 0
                    End If
                    .TextMatrix(.Row, 21) = CDbl(.TextMatrix(.Row, 10)) * CDbl(.TextMatrix(.Row, 8))
                    
                    'total harus dibayar = total harga - total discount - _
                        total hutang penjamin - totaltanggunganrs
'                    curHarusDibayar = CDbl(.TextMatrix(.Row, 11)) - (CDbl(.TextMatrix(.Row, 21)) + _
'                        CDbl(.TextMatrix(.Row, 19)) + CDbl(.TextMatrix(.Row, 20)))
'                    .TextMatrix(.Row, 22) = IIf(curHarusDibayar < 0, 0, curHarusDibayar)
                    .TextMatrix(.Row, 22) = ((.TextMatrix(.Row, 14) * .TextMatrix(.Row, 15)) + _
                        CDbl((.TextMatrix(.Row, 10)) * CDbl(.TextMatrix(.Row, 16)))) + CDbl(.TextMatrix(.Row, 26))
                    Call subHitungTotal
                    fgData.SetFocus
                    fgData.Col = 27
                    'Call subLoadCheck
                    'end fifo
                    
'                    fgData.SetFocus
'                    If txtNamaFormPengirim.Text <> "frmDaftarPenjualan" And txtNamaFormPengirim.Text <> "frmDaftarPenjualanTanpaBayar" Then
'                        fgData.Row = fgData.Rows - 1
'                        fgData.Col = 0
'                    End If
           
                Case 30
                    .TextMatrix(.Row, .Col) = Trim(txtIsi.Text)
                    txtIsi.Visible = False
                    fgData.SetFocus
                    fgData.Col = 31
                
            End Select
        End With

    ElseIf KeyAscii = 27 Then
        txtIsi.Visible = False
        fgData.SetFocus
    ElseIf (KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = Asc(".") Or KeyAscii = vbKeySpace) Then
        If fgData.Col <> 3 Then
            dgObatAlkes.Visible = False
        Else
            dgObatAlkes.Visible = True
            txtIsi.Visible = True
        End If
    Else
        KeyAscii = 0
    End If
    If fgData.Col = 0 Then Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtIsi_LostFocus()
    txtIsi.Visible = False
    
End Sub

Private Sub subSetGrid()
    On Error GoTo errLoad
    With fgData
        .Clear
        .Rows = 2
        .Cols = 38
        
        .RowHeight(0) = 400
        
        .TextMatrix(0, 0) = "R/Ke"
        .TextMatrix(1, 0) = "1"
        .TextMatrix(0, 1) = "Jenis Obat"
        .TextMatrix(0, 2) = "KodeBarang"
        .TextMatrix(0, 3) = "Nama Barang"
        .TextMatrix(0, 4) = "Kekuatan"
        .TextMatrix(0, 5) = "Asal Barang"
        .TextMatrix(0, 6) = "Satuan"
        .TextMatrix(0, 7) = "Harga Satuan" 'udah ditambah tarif service
        .TextMatrix(0, 8) = "Discount" 'per 1 barang
        .TextMatrix(0, 9) = "Stock" 'untuk perbandingan jika jumlah dirubah cek stok
        .TextMatrix(0, 10) = "Jumlah"
        .TextMatrix(0, 11) = "Total Harga" 'total = ((tarifservice * jmlservice) + (hargasatuan(sebelum ditambah tarifservixe) * jumlah))
        
        .TextMatrix(0, 12) = "KdAsal"
        .TextMatrix(0, 13) = "Jenis Barang"
        
        .TextMatrix(0, 14) = "TarifServise" 'per jenis obat
        .TextMatrix(0, 15) = "JmlService" 'jika obat jadi = Jumlah, 1 atau 0
        .TextMatrix(0, 16) = "HargaSebelumTarifService" 'harga satuan sesudah take tarif dan sebelum ditambah tarif service
        
        .TextMatrix(0, 17) = "HutangPenjamin" 'PER BARANG, diambil dari TempDetailApotikJual
        .TextMatrix(0, 18) = "TanggunganRS" 'PER BARANG, diambil dari TempDetailApotikJual
        
        .TextMatrix(0, 19) = "TotalHutangPenjamin" 'jumlah * hutang penjamin
        .TextMatrix(0, 20) = "TotalTanggunganRS" 'jumlah * TotalTanggunganRS
        .TextMatrix(0, 21) = "TotalDiscount" 'jumlah * discount 1 barang
        .TextMatrix(0, 22) = "TotalHarusBayar" 'curHarusDibayar = Total Harga - TotalDiscount - TotalHutangPenjamin - TotalTanggunganRS _
                                iif curHarusDibayar < 0, 0, curHarusDibayar
        .TextMatrix(0, 23) = "NoTemp" 'jika barang dihapus digrid, hapus ke tabel TempDetailApotikJual
        .TextMatrix(0, 24) = "HargaBeli" 'harga satuan sebelum take tarif dan sebelum ditambah tarif service
        .TextMatrix(0, 25) = "KdJenisObat"
        .TextMatrix(0, 26) = "BiayaAdministrasi"
        
        .TextMatrix(0, 27) = "Aturan Pakai"
        .TextMatrix(0, 28) = "Keterangan Pakai"
'        .TextMatrix(0, 29) = "Keterangan Pakai 2"
        .TextMatrix(0, 29) = "Keterangan Waktu"
        .TextMatrix(0, 30) = "Keterangan Lainnya"
        .TextMatrix(0, 31) = "Pemakaian Pemeriksaan"
        .TextMatrix(0, 32) = "KodePelayananRS"
        .TextMatrix(0, 33) = "NoTerima" 'add No Terima
        .TextMatrix(0, 34) = "NoRacikan" ' add No Racikan
        .TextMatrix(0, 35) = "KdSatuanEtiket"
        .TextMatrix(0, 36) = "KdWaktuEtiket"
        .TextMatrix(0, 37) = "KdWaktuEtiket2"
        
        
        .ColWidth(0) = 500
        .ColWidth(1) = 1200
        .ColWidth(2) = 0
        .ColWidth(3) = 3500
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0 '1500
        .ColWidth(8) = 0
        .ColWidth(9) = 700
        .ColWidth(10) = 700
        .ColWidth(11) = 0 '1200
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        .ColWidth(14) = 0   '0
        .ColWidth(15) = 0 '0
        .ColWidth(16) = 0
        .ColWidth(17) = 0
        .ColWidth(18) = 0
        .ColWidth(19) = 0
        .ColWidth(20) = 0
        .ColWidth(21) = 0
        .ColWidth(22) = 0
        .ColWidth(23) = 0
        .ColWidth(24) = 0
        .ColWidth(25) = 0
        .ColWidth(26) = 0
        .ColWidth(27) = 3000
        .ColWidth(28) = 1700
        .ColWidth(29) = 2000
        .ColWidth(30) = 2000
        .ColWidth(31) = 0 '2800
        .ColWidth(32) = 0
        .ColWidth(33) = 0 ' add noRacikan
        .ColWidth(34) = 0 ' add noRacikan
        .ColWidth(35) = 0
        .ColWidth(36) = 0
        .ColWidth(37) = 0
    
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignRightCenter
        .ColAlignment(10) = flexAlignRightCenter
        .ColAlignment(11) = flexAlignRightCenter
    End With
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadDcSource()
On Error GoTo errLoad
'    If UseRacikan = True Then
        Call msubDcSource(dcJenisObat, rs, "SELECT KdJenisObat, JenisObat FROM JenisObat where StatusEnabled=1 ORDER BY JenisObat")
'      Else
'        Call msubDcSource(dcJenisObat, rs, "SELECT     JenisObat.KdJenisObat, JenisObat.JenisObat FROM JenisObat INNER JOIN  SettingGlobal ON JenisObat.KdJenisObat <> SettingGlobal.Value WHERE     (JenisObat.StatusEnabled = 1) and SettingGlobal.Prefix='KdJenisObatRacikan' ORDER BY JenisObat.JenisObat")
'    End If
    'If rs.EOF = False Then dcJenisObat.BoundText = rs(0).Value
    
    strSQL = "SELECT  TOP (200) KdPelayananRS, NamaPelayanan, NoPendaftaran" & _
            " FROM  V_NamaPelayananPerPasien where NoPendaftaran ='" & mstrNoPen & "'"
    Call msubDcSource(dcNamaPelayananRS, rs, strSQL)
    
'    Call msubDcSource(dcAturanPakai, rs, "select KdSatuanEtiket, NamaExternal,SatuanEtiket from SatuanEtiketResep where StatusEnabled=1 Order By SatuanEtiket")
    Call msubDcSource(dcAturanPakai, rs, "select KdSatuanEtiket, SatuanEtiket from SatuanEtiketResep where StatusEnabled=1 Order By SatuanEtiket")
    'If rs.EOF = False Then dcAturanPakai.BoundText = rs(0).Value
    
    Call msubDcSource(dcKeteranganPakai, rs, "select KdWaktuEtiket,WaktuEtiket from WaktuEtiketResep where StatusEnabled=1 order by KdWaktuEtiket")
    'If rs.EOF = False Then dcKeteranganPakai.BoundText = rs(0).Value
    
    Call msubDcSource(dcKeteranganPakai2, rs, "select KdWaktuEtiket2,WaktuEtiket2 from WaktuEtiketResep2 where StatusEnabled=1 order by WaktuEtiket2")
    'If rs.EOF = False Then dcKeteranganPakai2.BoundText = rs(0).Value

    Call msubDcSource(dcRuanganTujuan, rs, "Select KdRuangan,NamaRuangan From Ruangan Where StatusEnabled=1 and KdInstalasi='" & TempMstrKdIsntalasi & "' and kdruangan not in ('701') order by NamaRuangan asc")
    dcRuanganTujuan.BoundText = "702"
    tempKdRuanganTujuan = "702"
Exit Sub
errLoad:
    Call msubPesanError
End Sub

''add by Denki
Private Sub subSetGridRacikan()
On Error GoTo errLoad
    With fgRacikan
        .Visible = True
        .Clear
        .Rows = 2
        .Cols = 18
        
        .RowHeight(0) = 400
        .TextMatrix(0, 0) = "" 'KdBarang
        .TextMatrix(0, 1) = "" 'Jenis obat
        .TextMatrix(0, 2) = "R/Ke"
        .TextMatrix(0, 3) = "Nama Barang"
        .TextMatrix(0, 4) = "/Mg /Ml"
        .TextMatrix(0, 5) = "/Tablet"
        .TextMatrix(0, 6) = "Jumlah"
        .TextMatrix(0, 7) = "Jumlah Pembulatan(untuk harga)"
        .TextMatrix(0, 8) = "Harga Satuan"
        .TextMatrix(0, 9) = "Total Harga"
        .TextMatrix(0, 10) = "Kekuatan"
        .TextMatrix(0, 11) = "AsalBarang"
        .TextMatrix(0, 12) = "kdAsal"
        .TextMatrix(0, 13) = "satuan"
        .TextMatrix(0, 14) = "NoFIFO"
        .TextMatrix(0, 15) = "jmlService" 'add Column Jumlah Service
        .TextMatrix(0, 16) = "TarifService" ' add Column TarifService
        .TextMatrix(0, 17) = "JumlahStok" ' add Column TarifService
        
'        .TextMatrix(0, 18) = "TanggunganRS"
        
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 1000
        .ColWidth(3) = 4800
        .ColWidth(4) = 1800
        .ColWidth(5) = 1200
        .ColWidth(6) = 1200 '0
        .ColWidth(7) = 0 '1200
        .ColWidth(8) = 0 '1800
        .ColWidth(9) = 0 '1800
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        .ColWidth(14) = 0
        .ColWidth(15) = 0
        .ColWidth(16) = 0
        .ColWidth(17) = 0
'        .ColWidth(16) = 0
        
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignCenterCenter
        .ColAlignment(7) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignCenterCenter
        .ColAlignment(9) = flexAlignCenterCenter
        .ColAlignment(10) = flexAlignCenterCenter

    End With
        
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub TxtIsiRacikan_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyDown Then If dgObatAlkesRacikan.Visible = True Then If dgObatAlkesRacikan.ApproxCount > 0 Then dgObatAlkesRacikan.SetFocus
    If KeyCode = vbKeyDown Then If dgObatAlkesRacikan.Visible = True Then If dgObatAlkesRacikan.ApproxCount > 0 Then dgObatAlkesRacikan.SetFocus
    If KeyCode = 13 Then If dgObatAlkesRacikan.Visible = True Then If dgObatAlkesRacikan.ApproxCount > 0 Then dgObatAlkesRacikan.SetFocus

End Sub

Private Sub TxtIsiRacikan_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim i As Integer
Dim curHutangPenjamin As Currency
Dim curTanggunganRS As Currency
Dim curHarusDibayar As Currency
Dim KdJnsObat As String
    Select Case fgRacikan.Col
        Case 4 'Kebutuhan /ML
            Call SetKeyPressToNumber(KeyAscii)
        Case 5 'Kebutuhan /Tablet
            Call SetKeyPressToNumber(KeyAscii)
        Case 6 'Jumlah
            Call SetKeyPressToNumber(KeyAscii)
    End Select

    If KeyAscii = 13 Then
        With fgRacikan
            Select Case .Col
                Case 0
                    If Val(fgRacikan.Text) = 0 Then TxtIsiRacikan.Text = 0
                    .TextMatrix(.Row, .Col) = Val(TxtIsiRacikan.Text)
                    TxtIsiRacikan.Visible = False
                    fgRacikan.SetFocus
                    fgRacikan.Col = 1
                    
                    
                    
                Case 2
                    If Val(fgRacikan.Text) = 0 Then TxtIsiRacikan.Text = 0
                    .TextMatrix(.Row, .Col) = Val(TxtIsiRacikan.Text)
                    TxtIsiRacikan.Visible = False
                    fgRacikan.SetFocus
                    fgRacikan.Col = 3
                    
                Case 4 'Kebutuhan /ML
                    'konvert koma col jumlah
                    
                    .TextMatrix(.Row, .Col) = msubKonversiKomaTitik(TxtIsiRacikan.Text)
                    
                    If Val(TxtIsiRacikan.Text) = 0 Then TxtIsiRacikan.Text = 0
                    TxtIsiRacikan.Visible = False
'                    fraHitungObat.Visible = True
                    .TextMatrix(.Row, .Col) = msubKonversiKomaTitik(TxtIsiRacikan.Text)
                    
                    strSQL = "Select Kekuatan From MasterBarang Where KdBarang = '" & fgRacikan.TextMatrix(fgRacikan.Row, 0) & "'"
                    Call msubRecFO(rs, strSQL)
                    If rs.EOF = False Then
                        strBeratObat = rs.Fields(0)
                    End If
                    
                    Call HitungObat
'                    txtBeratObat.SetFocus
                    
                Case 5 'Kebutuhan Tablet
                    'konvert koma col jumlah
                    .TextMatrix(.Row, .Col) = msubKonversiKomaTitik(TxtIsiRacikan.Text)
                    TxtIsiRacikan.Visible = False
                    
                    If Val(TxtIsiRacikan.Text) = 0 Then TxtIsiRacikan.Text = 0
                                   
                    If .TextMatrix(.Row, .Col) <> 0 Then
                        riilnya = (Val(txtJumlahObatRacik.Text) * Val(.TextMatrix(.Row, 5)))
                    Else
                        MsgBox "GUNAKAN TITIK!", vbCritical, "Peringatan"
                        Exit Sub
                    End If
                    Set dbrs = Nothing
                    Call msubRecFO(dbrs, "select dbo.FB_TakeStokBrgMedis('" & dcRuanganTujuan.BoundText & "', '" & fgRacikan.TextMatrix(fgRacikan.Row, 0) & "','" & fgRacikan.TextMatrix(fgRacikan.Row, 12) & "') as stok")
                    

                    If riilnya > CDbl(dbrs(0).Value) Then
                        MsgBox "Jumlah lebih besar dari stok ", vbExclamation, "Validasi"
'                        TxtIsiRacikan.SelStart = 0: TxtIsiRacikan.SelLength = Len(TxtIsiRacikan.Text)
                        Exit Sub
                    End If
                    .TextMatrix(.Row, 6) = riilnya
                    .TextMatrix(.Row, 6) = msubKonversiKomaTitik(.TextMatrix(.Row, 6))
                    blt = 0
                    blt = Round(riilnya, 1)
'                    If blt < riilnya Then
'                        blt = blt + 1
'                    End If
                    .TextMatrix(.Row, 4) = 0
                    
                    
                    .TextMatrix(.Row, 7) = CStr(blt)
                    '.TextMatrix(.Row, 7) = .TextMatrix(.Row, 6)
                    '.TextMatrix(.Row, 9) = val(.TextMatrix(.Row, 7)) * msubKonversiKomaTitik(.TextMatrix(.Row, 8)) 'Val(.TextMatrix(.Row, 8))
                    .TextMatrix(.Row, 9) = (CDbl(subcurTarifServiceRacikan) * CDbl(subintJmlServiceRacikan)) + (Val(.TextMatrix(.Row, 6)) * CDbl(.TextMatrix(.Row, 8)))
                    
'                    If sp_TempDetailApotikJual(CDbl(.TextMatrix(.Row, 8)), .TextMatrix(.Row, 0), .TextMatrix(.Row, 12)) = False Then Exit Sub
'                    'ambil hutang penjamin dan tanggungan rs
'                    strSQL = "SELECT HargaSatuan, JmlHutangPenjamin, JmlTanggunganRS" & _
'                        " FROM TempDetailApotikJual" & _
'                        " WHERE (NoTemporary = '" & Trim(txtNoTemporary.Text) & "') AND (KdBarang = '" & .TextMatrix(.Row, 0) & "') AND (KdAsal = '" & .TextMatrix(.Row, 12) & "')"
'                    Call msubRecFO(rs, strSQL)
'                    If rs.EOF = True Then
'                        curHutangPenjamin = 0
'                        curTanggunganRS = 0
'                    Else
'                        curHutangPenjamin = rs("JmlHutangPenjamin").Value
'                        curTanggunganRS = rs("JmlTanggunganRS").Value
'                    End If
'
'                    .TextMatrix(.Row, 15) = curHutangPenjamin
'                    .TextMatrix(.Row, 16) = curTanggunganRS
'
'                    .TextMatrix(.Row, .Col) = msubKonversiKomaTitik(TxtIsiRacikan.Text)
                    
                    .TextMatrix(.Row, 9) = Format(.TextMatrix(.Row, 9), "#,###.00")
                    .SetFocus
                    .Col = 9
                    
                 
                    
                Case 6 'Jumlah Real
                    'konvert koma col jumlah
                    .TextMatrix(.Row, .Col) = msubKonversiKomaTitik(TxtIsiRacikan.Text)
                    
                    Set dbrs = Nothing
                    Call msubRecFO(dbrs, "select dbo.FB_TakeStokBrgMedis('" & dcRuanganTujuan.BoundText & "', '" & .TextMatrix(.Row, 0) & "','" & .TextMatrix(.Row, 12) & "') as stok")
                    

                    If CDbl(TxtIsiRacikan.Text) > CDbl(dbrs(0).Value) Then
                        MsgBox "Jumlah lebih besar dari stok ", vbExclamation, "Validasi"
                        TxtIsiRacikan.SelStart = 0: TxtIsiRacikan.SelLength = Len(TxtIsiRacikan.Text)
                        Exit Sub
                    End If
                    

                    
                    
                    TxtIsiRacikan.Visible = False
                    
                    If Val(TxtIsiRacikan.Text) = 0 Then TxtIsiRacikan.Text = 0
                                   
                    If .TextMatrix(.Row, .Col) <> 0 Then
                        riilnya = Val(.TextMatrix(.Row, 6))
                    Else
                        MsgBox "GUNAKAN TITIK!", vbCritical, "Peringatan"
                        Exit Sub
                    End If
                    .TextMatrix(.Row, 6) = msubKonversiKomaTitik(.TextMatrix(.Row, 6))
                  
                  '  ditutup by azein
                    blt = 0
                    blt = Round(riilnya, 1)
'                    If blt > riilnya Then
'                        blt = blt + 1
'                    ElseIf blt < riilnya Then
'                        blt = blt
'                    End If
                  '  end tutup
'                    If .TextMatrix(.Row, 6) >= "1" Then
'                       blt = blt
'                    ElseIf .TextMatrix(.Row, 6) < "1" Then
'                       blt = blt + 1
'                    End If

                    .TextMatrix(.Row, 4) = 0
                    .TextMatrix(.Row, 5) = 0
                    
                    .TextMatrix(.Row, 7) = CStr(blt)
'                    .TextMatrix(.Row, 7) = .TextMatrix(.Row, 6)
                    '.TextMatrix(.Row, 9) = val(.TextMatrix(.Row, 7)) * msubKonversiKomaTitik(.TextMatrix(.Row, 8)) 'Val(.TextMatrix(.Row, 8))
                    'akung karyawan tidak dikenakan uang r dan selain karyawan dikenakan uang r peritem
'                    If mstrKdJenisPasien = "07" Or mstrKdJenisPasien = "14" Then
'                        .TextMatrix(.Row, 9) = (val(.TextMatrix(.Row, 7)) * val(.TextMatrix(.Row, 8)))
'                    Else
                        .TextMatrix(.Row, 9) = (CDbl(subcurTarifServiceRacikan) * CDbl(subintJmlServiceRacikan)) + (Val(.TextMatrix(.Row, 6)) * CDbl(.TextMatrix(.Row, 8)))
'                    End If
                    '---
                    
'                    If sp_TempDetailApotikJual(CDbl(.TextMatrix(.Row, 8)), .TextMatrix(.Row, 0), .TextMatrix(.Row, 12)) = False Then Exit Sub
'                    'ambil hutang penjamin dan tanggungan rs
'                    strSQL = "SELECT HargaSatuan, JmlHutangPenjamin, JmlTanggunganRS" & _
'                        " FROM TempDetailApotikJual" & _
'                        " WHERE (NoTemporary = '" & Trim(txtNoTemporary.Text) & "') AND (KdBarang = '" & .TextMatrix(.Row, 0) & "') AND (KdAsal = '" & .TextMatrix(.Row, 12) & "')"
'                    Call msubRecFO(rs, strSQL)
'                    If rs.EOF = True Then
'                        curHutangPenjamin = 0
'                        curTanggunganRS = 0
'                    Else
'                        curHutangPenjamin = rs("JmlHutangPenjamin").Value
'                        curTanggunganRS = rs("JmlTanggunganRS").Value
'                    End If
'
'                    .TextMatrix(.Row, 15) = curHutangPenjamin
'                    .TextMatrix(.Row, 16) = curTanggunganRS

                    .TextMatrix(.Row, 15) = subintJmlServiceRacikan
                    .TextMatrix(.Row, 16) = subcurTarifServiceRacikan
                    
                    
                    .TextMatrix(.Row, .Col) = msubKonversiKomaTitik(TxtIsiRacikan.Text)
                    
                    .TextMatrix(.Row, 9) = Format(.TextMatrix(.Row, 9), "#,###.00")
                    .SetFocus
                    .Col = 9
                    If .TextMatrix(.Row, 6) = 0 Then
                        MsgBox "GUNAKAN TITIK!", vbCritical, "Peringatan"
'
                        .Col = 6
                        Exit Sub
                    End If
                    
                Case 7 'Jumlah
                    .TextMatrix(.Row, .Col) = msubKonversiKomaTitik(TxtIsiRacikan.Text)
                    .TextMatrix(.Row, 6) = msubKonversiKomaTitik(.TextMatrix(.Row, .Col))
                    
                    .TextMatrix(.Row, 5) = 0
                    .TextMatrix(.Row, 4) = 0
                    
                    .TextMatrix(.Row, 15) = subintJmlServiceRacikan
                    .TextMatrix(.Row, 16) = subcurTarifServiceRacikan
                    
                    .TextMatrix(.Row, 9) = (CDbl(subcurTarifServiceRacikan) * CDbl(subintJmlServiceRacikan)) + (Val(.TextMatrix(.Row, 7)) * Val(.TextMatrix(.Row, 8)))

                    .TextMatrix(.Row, 9) = Format(.TextMatrix(.Row, 9), "#,###.00")
                     TxtIsiRacikan.Visible = False

                    fgRacikan.SetFocus
                    fgRacikan.Col = 8
                    
                Case 8
                  
                    .TextMatrix(.Row, .Col) = TxtIsiRacikan.Text
                    fgRacikan.Visible = False
                    
                    fgRacikan.SetFocus
                    fgRacikan.Col = 8
               
            End Select
        End With

    ElseIf KeyAscii = 27 Then
        TxtIsiRacikan.Visible = False
        dgObatAlkesRacikan.Visible = False
        fgRacikan.SetFocus
    ElseIf (KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = Asc(".") Or KeyAscii = vbKeySpace) Then
        If fgRacikan.Col <> 3 Then
            dgObatAlkesRacikan.Visible = False
        Else
            dgObatAlkesRacikan.Visible = True
            TxtIsiRacikan.Visible = True
        End If
    
    End If
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Private Sub HitungObat()
On Error GoTo errLoad
Dim riilnya As Double
Dim temptriilnya As String
Dim blt As Integer
'If KeyAscii = 13 Then
    If fgRacikan.TextMatrix(fgRacikan.Row, 4) <> 0 Then
        riilnya = (Val(txtJumlahObatRacik.Text) * Val(fgRacikan.TextMatrix(fgRacikan.Row, 4))) / Val(strBeratObat)

    Else
        MsgBox "Kebutuhan /ML Tidak Boleh Nol atau Kosong", vbCritical, "Peringatan"
        Exit Sub
    End If
    
    Set dbrs = Nothing
    Call msubRecFO(dbrs, "select dbo.FB_TakeStokBrgMedis('" & dcRuanganTujuan.BoundText & "', '" & fgRacikan.TextMatrix(fgRacikan.Row, 0) & "','" & fgRacikan.TextMatrix(fgRacikan.Row, 12) & "') as stok")
                   

    If riilnya > CDbl(dbrs(0).Value) Then
        MsgBox "Jumlah lebih besar dari stok ", vbExclamation, "Validasi"
        TxtIsiRacikan.SelStart = 0: TxtIsiRacikan.SelLength = Len(TxtIsiRacikan.Text)
        Exit Sub
    End If
    temptriilnya = riilnya
    fgRacikan.TextMatrix(fgRacikan.Row, 6) = CStr(riilnya)

    With fgRacikan
        .TextMatrix(.Row, 5) = 0
        .TextMatrix(fgRacikan.Row, 6) = funcRoundUp(temptriilnya)
        
'        .TextMatrix(.Row, 6) = msubKonversiKomaTitik(.TextMatrix(.Row, 6))
'         .TextMatrix(.Row, 6) = riilnya
                  
        blt = 0
        blt = Round(riilnya, 1)
        
        .TextMatrix(fgRacikan.Row, 7) = CStr(blt) 'msubKonversiKomaTitik(.TextMatrix(fgRacikan.Row, 7))
'        .TextMatrix(fgRacikan.Row, 7) = CDbl(riilnya)
        .TextMatrix(.Row, 9) = (CDbl(subcurTarifServiceRacikan) * CDbl(subintJmlServiceRacikan)) + (Val(.TextMatrix(.Row, 6)) * CDbl(.TextMatrix(.Row, 8)))
'        .TextMatrix(.Row, 9) = Format(.TextMatrix(.Row, 9), "#,###.00")
        .TextMatrix(.Row, 9) = FormatPembulatan(.TextMatrix(.Row, 9), KI)
        
        .SetFocus
        .Col = 9
    End With
'    fraHitungObat.Visible = False
'    txtBeratObat.Text = ""
'End If

'Call SetKeyPressToNumber(KeyAscii)
Exit Sub
errLoad:
    Call msubPesanError

End Sub
Private Sub TxtIsiRacikan_Change()
On Error GoTo hell
Dim i As Integer
Dim iFifo As Integer
kolom = 0
    Select Case fgRacikan.Col
        
        Case 3 ' nama barang
        
'            strSQLx = "Select MetodeStokBarang from SuratKeputusanRuleRS where statusenabled=1"
'            Call msubRecFO(rsx, strSQLx)
'            If rsx.EOF = False Then iFifo = rsx(0).Value Else iFifo = 0
            'If tempStatusTampil = True Then Exit Sub
'            If subJenisHargaNetto = 2 Then
'                strSQL = "select  TOP 100 JenisBarang, RuanganPelayanan, NamaBarang, Kekuatan, AsalBarang, Satuan, HargaBarang, JmlStok, Discount, KdBarang, KdAsal from V_HargaBarangNStok2 " & _
'                    " where NamaBarang like '" & TxtIsiRacikan.Text & "%' AND KdRuangan like'" & mstrKdRuangan & "%' ORDER BY NamaBarang"
'            Else
                ' CariBarangNStokMedis_V
'             strSQL = "execute CariBarangNStokMedis_V '" & TxtIsiRacikan.Text & "%','" & dcRuanganTujuan.BoundText & "'"
            strSQL = "execute CariBarangNStokMedis_VToAsalBarang'" & txtIsi.Text & "%','" & dcRuanganTujuan.BoundText & "','" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "'"
            
            Call msubRecFO(dbRst, strSQL)
'            Set dgObatAlkesRacikan.DataSource = dbRst
'            With dgObatAlkesRacikan
'                For i = 0 To .Columns.Count - 1
'                    .Columns(i).Width = 0
'                Next i
'
'                .Columns("KdBarang").Width = 0
'                .Columns("NamaBarang").Width = 3000
'                .Columns("JenisBarang").Width = 1500
'                .Columns("Kekuatan").Width = 1000
'                .Columns("AsalBarang").Width = 1000
'                .Columns("Satuan").Width = 675
'
'                .Top = txtIsi.Top + txtIsi.Height + Frame8.Top
'                .Left = 3000
'                .Visible = True
'
'            End With
'
'            If bolStatusFIFO = False Then
'                strSQL = "select  TOP 100 JenisBarang, RuanganPelayanan, NamaBarang, Kekuatan, AsalBarang, Satuan, HargaBarang, JmlStok, Discount, KdBarang, KdAsal from V_HargaBarangNStok1 " & _
'                    " where NamaBarang like '" & TxtIsiRacikan.Text & "%' AND KdRuangan like '" & dcRuanganTujuan.BoundText & "%' ORDER BY NamaBarang"
'            Else
'                strSQL = "select  TOP 100 DetailJenisBarang AS JenisBarang, NamaRuangan AS RuanganPelayanan, NamaBarang, Kekuatan, NamaAsal AS AsalBarang, Satuan, HargaNetto1 AS HargaBarang, JmlStok, Discount, KdBarang, KdAsal from V_StokNHargaGlobalFIFO " & _
'                    " where NamaBarang like '" & TxtIsiRacikan.Text & "%' AND KdRuangan like '" & dcRuanganTujuan.BoundText & "%' ORDER BY NamaBarang"
'            End If
'            Call msubRecFO(dbRst, strSQL)
            
            Set dgObatAlkesRacikan.DataSource = dbRst
            With dgObatAlkesRacikan
                .Columns("RuanganPelayanan").Width = 0
                .Columns("JenisBarang").Width = 1500
                .Columns("NamaBarang").Width = 4000
                .Columns("Kekuatan").Width = 0
                .Columns("AsalBarang").Width = 2000
                .Columns("Satuan").Width = 0
'                .Columns("HargaBarang").Width = 1500
                .Columns("JmlStok").Width = 1000
                
'                .Columns("HargaBarang").NumberFormat = "#,###"
'                .Columns("HargaBarang").Alignment = dbgRight
        
                .Columns("JmlStok").NumberFormat = "#,###"
                .Columns("JmlStok").Alignment = dbgRight
                
                .Columns("Discount").Width = 0
                .Columns("KdBarang").Width = 0
                .Columns("KdAsal").Width = 0
                .Columns("NamaGenerik").Width = 0
                .Columns("KdRuangan").Width = 0
                .Columns("KdGenerikBarang").Width = 0
                .Columns("JenisHarga").Width = 0
                .Columns("SatuanJmlB").Width = 0
                
                .Top = TxtIsiRacikan.Top + TxtIsiRacikan.Height
                .Left = TxtIsiRacikan.Left + TxtIsiRacikan.Height - 400
                
                
                .Visible = True
                
'                .Left = fgData.ColPos(4)  '720
'                .Top = fgRacikan.Top + fgRacikan.RowHeight(0) '1440
                
                
'                For i = 1 To fgRacikan.Row - 1
'                    .Top = .Top + fgRacikan.RowHeight(i)
'                Next i
'                If fgRacikan.TopRow > 1 Then
'                    .Top = .Top - ((fgRacikan.TopRow - 1) * fgRacikan.RowHeight(1))
'                End If
            End With
        Case 4
            kolom = 4
        Case 5
            kolom = 5
        Case 7
            kolom = 7
        Case Else
            dgObatAlkesRacikan.Visible = False
            kolom = 0
    End Select
Exit Sub
hell:
    Call msubPesanError

End Sub

Private Sub txtBeratObat_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
Dim riilnya As Double
Dim blt As Integer
If KeyAscii = 13 Then
    If fgRacikan.TextMatrix(fgRacikan.Row, 4) <> 0 Then
        riilnya = (Val(txtJumlahObatRacik.Text) * Val(fgRacikan.TextMatrix(fgRacikan.Row, 4))) / Val(txtBeratObat.Text)
    Else
        MsgBox "Kebutuhan /ML Tidak Boleh Nol atau Kosong", vbCritical, "Peringatan"
        Exit Sub
    End If
    fgRacikan.TextMatrix(fgRacikan.Row, 6) = CStr(riilnya)

    With fgRacikan
        .TextMatrix(.Row, 5) = 0
        .TextMatrix(fgRacikan.Row, 6) = riilnya
        
        .TextMatrix(.Row, 6) = msubKonversiKomaTitik(.TextMatrix(.Row, 6))
'         .TextMatrix(.Row, 6) = riilnya
                  
        blt = 0
        blt = Round(riilnya, 1)
        .TextMatrix(fgRacikan.Row, 7) = CStr(blt) 'msubKonversiKomaTitik(.TextMatrix(fgRacikan.Row, 7))
        .TextMatrix(.Row, 9) = (CDbl(subcurTarifServiceRacikan) * CDbl(subintJmlServiceRacikan)) + (Val(.TextMatrix(.Row, 6)) * Val(.TextMatrix(.Row, 8)))
        .TextMatrix(.Row, 9) = Format(.TextMatrix(.Row, 9), "#,###.00")
        
        .SetFocus
        .Col = 9
    End With
    fraHitungObat.Visible = False
    txtBeratObat.Text = ""
End If

Call SetKeyPressToNumber(KeyAscii)
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtIsiRacikan_LostFocus()
    With fgRacikan
        Select Case .Col
            Case 5 'Kebutuhan /ML
                If kolom = 4 Then
                    'konvert koma col jumlah
                    .TextMatrix(.Row, kolom) = msubKonversiKomaTitik(TxtIsiRacikan.Text)
                    If Val(TxtIsiRacikan.Text) = 0 Then TxtIsiRacikan.Text = 0
                    TxtIsiRacikan.Visible = False
'                    fraHitungObat.Visible = True
                    .TextMatrix(.Row, kolom) = msubKonversiKomaTitik(TxtIsiRacikan.Text)
'                    txtBeratObat.Text = ""
'                    txtBeratObat.SetFocus
                End If
            Case Else
                TxtIsiRacikan.Visible = False
                
          End Select
    End With
End Sub

Private Sub txtJumlahObatRacik_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        With fgRacikan
            .SetFocus
            .Col = 3
            txtJumlahObatRacik.Text = Val(txtJumlahObatRacik.Text)
        End With
    End If
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtJumlahObatRacik_LostFocus()
'If txtJumlahObatRacik = "" Or txtJumlahObatRacik = "0" Then
'    txtJumlahObatRacik.Enabled = True
'    MsgBox "Jumlah Obat Racikan tidak boleh Nol atau Kosong", vbExclamation, "Medifirst2000 - Validation"
'    txtJumlahObatRacik.SetFocus
'Else
'    txtJumlahObatRacik.Enabled = False
'End If
End Sub

Private Sub txtNoResep_Change()
'On Error GoTo errLoad
'    If Len(Trim(txtNoResep.Text)) = 0 Then Exit Sub
'    Call subLoadDataResep(txtNoResep.Text)
'Exit Sub
'errLoad:
'    Call msubPesanError
End Sub

Private Sub txtNoResep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglResep.SetFocus
End Sub

'untuk meload data dokter di grid
Private Sub subLoadDokter()
On Error GoTo errLoad

    strSQL = "SELECT NamaDokter AS [Nama Dokter],JK,Jabatan,KodeDokter  FROM V_DaftarDokter " & mstrFilterDokter
    Call msubRecFO(rs, strSQL)
    With dgDokter
        Set .DataSource = rs
        .Columns(0).Width = 3500
        .Columns(1).Width = 400
        .Columns(2).Width = 1600
        .Columns(3).Width = 0
    End With
    dgDokter.Left = 6960
    dgDokter.Top = 2880

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Function f_HitungTotal() As Currency
On Error GoTo errLoad
Dim i As Integer

    f_HitungTotal = 0
    For i = 1 To fgData.Rows - 2
        f_HitungTotal = f_HitungTotal + fgData.TextMatrix(i, 11)
    Next i
    
Exit Function
errLoad:
    Call msubPesanError
End Function

Private Function sp_Order() As Boolean
On Error GoTo errLoad
    
    sp_Order = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("TglOrder", adDate, adParamInput, , Format(dtpTglOrder.Value, "yyyy/MM/dd hh:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, IIf(mstrKdRuangan = "", Null, mstrKdRuangan))
        .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, IIf(dcRuanganTujuan.BoundText = "", Null, dcRuanganTujuan.BoundText))
        .Parameters.Append .CreateParameter("KdSupplier", adChar, adParamInput, 4, Null)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("OutKode", adChar, adParamOutput, 10, Null)
        
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_StrukOrder"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_Order = False
        Else
            txtNoOrder.Text = .Parameters("OutKode")
        
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    
Exit Function
errLoad:
    Call msubPesanError
    sp_Order = False
End Function
'Public Function sp_StokRuangan(f_KdBarang As String, f_KdAsal As String, f_NoTerima As String, f_JmlBarang As Double, f_Status As String) As Boolean
'On Error GoTo errLoad
'Dim i As Integer
'
'    sp_StokRuangan = True
'    Set dbcmd = New ADODB.Command
'    With dbcmd
'        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
'        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
'        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
'        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
'        .Parameters.Append .CreateParameter("NoTerima", adChar, adParamInput, 10, f_NoTerima)
'        .Parameters.Append .CreateParameter("JmlBrg", adDouble, adParamInput, , CDbl(f_JmlBarang))
'        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
'
'        .ActiveConnection = dbConn
'        .CommandText = "Update_StokRuanganDynamic" '"dbo.Update_StokRuangan"
'        .CommandType = adCmdStoredProc
'        .Execute
'
'        If .Parameters("return_value").Value <> 0 Then
'            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "validasi"
'            sp_StokRuangan = False
''        Else
''            Call Add_HistoryLoginActivity("Update_StokRuangan")
'        End If
'    End With
'    Set dbcmd = Nothing
'    Call deleteADOCommandParameters(dbcmd)
'
'Exit Function
'errLoad:
'    sp_StokRuangan = False
'    Call msubPesanError("sp_StokRuangan")
'End Function

Private Function sp_DetailOrderPelayananOA(f_KdBarang As String, f_JmlBarang As Double, f_ResepKe As Integer, f_KdSatuanEtiket As String, f_KdWaktuEtiket As String, f_KdWaktuEtiket2 As String, f_KdPelayananRSUsed As String, f_KeteranganLainnya As String, _
        f_jmlService As Integer, f_TarifService As Currency, f_Cito As String, f_KdJenisObat As String, f_KdAsal As String, f_SatuanJml As String, f_NoTerima As String, f_Noracikan As String) As Boolean
On Error GoTo errLoad
    
    sp_DetailOrderPelayananOA = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, txtNoOrder.Text)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaranOA.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCMOA.Text)
        .Parameters.Append .CreateParameter("NoPakai", adVarChar, adParamInput, 10, IIf(Trim(txtNoPakai.Text) = "", Null, Trim(txtNoPakai.Text)))
        .Parameters.Append .CreateParameter("idDokterOrder", adChar, adParamInput, 10, IIf(Trim(txtKdDokter.Text) = "", Null, Trim(txtKdDokter.Text)))
        .Parameters.Append .CreateParameter("StatusCito", adChar, adParamInput, 1, IIf(f_Cito = "Yes", 1, 0))
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("JmlBarang", adDouble, adParamInput, , CDbl(f_JmlBarang))
        .Parameters.Append .CreateParameter("NoResep", adChar, adParamInput, 15, IIf(Trim(txtNoResep.Text) = "", Null, Trim(txtNoResep.Text)))
        .Parameters.Append .CreateParameter("ResepKe", adTinyInt, adParamInput, , f_ResepKe)
        .Parameters.Append .CreateParameter("KdSatuanEtiket", adChar, adParamInput, 3, Trim(f_KdSatuanEtiket))
        .Parameters.Append .CreateParameter("KdWaktuEtiket", adChar, adParamInput, 2, f_KdWaktuEtiket)
        .Parameters.Append .CreateParameter("KdWaktuEtiket2", adChar, adParamInput, 2, f_KdWaktuEtiket2)
        .Parameters.Append .CreateParameter("TglResep", adDate, adParamInput, , Format(dtpTglResep.Value, "yyyy/MM/dd hh:mm:ss"))
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("NoRetur", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("JmlRetur", adInteger, adParamInput, , Null)
        .Parameters.Append .CreateParameter("KdPelayananRSUSed", adChar, adParamInput, 6, IIf(f_KdPelayananRSUsed = "", Null, f_KdPelayananRSUsed))
        .Parameters.Append .CreateParameter("KeteranganLainnya", adVarChar, adParamInput, 200, IIf(f_KeteranganLainnya = "", Null, f_KeteranganLainnya))
        .Parameters.Append .CreateParameter("JmlService", adInteger, adParamInput, , f_jmlService)
        .Parameters.Append .CreateParameter("TarifService", adCurrency, adParamInput, , f_TarifService)
        .Parameters.Append .CreateParameter("KdJenisObat", adChar, adParamInput, 2, f_KdJenisObat)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("SatuanJml", adChar, adParamInput, 1, f_SatuanJml)
        .Parameters.Append .CreateParameter("NoTerima", adChar, adParamInput, 10, IIf((f_NoTerima) = "", Null, f_NoTerima))
        .Parameters.Append .CreateParameter("NoRacikan", adChar, adParamInput, 10, IIf((f_Noracikan) = "", Null, f_Noracikan))
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_DetailOrderPelayananOA"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam proses penyimpanan data", vbCritical, "Validasi"
            sp_DetailOrderPelayananOA = False

        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    
Exit Function
errLoad:
    Call msubPesanError
    sp_DetailOrderPelayananOA = False
End Function

Private Function sp_TempDetailApotikJual(f_HargaSatuan As Currency, f_KdBarang As String, f_KdAsal As String) As Boolean
    sp_TempDetailApotikJual = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoTemporary", adChar, adParamInput, 3, IIf(Len(Trim(txtNoTemporary.Text)) = 0, Null, Trim(txtNoTemporary.Text)))
        .Parameters.Append .CreateParameter("KdKelompokPasien", adChar, adParamInput, 2, mstrKdJenisPasien)
        .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, mstrKdPenjaminPasien)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang) 'fgData.TextMatrix(fgData.Row, 2))
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal) 'fgData.TextMatrix(fgData.Row, 12))
        .Parameters.Append .CreateParameter("HargaSatuan", adCurrency, adParamInput, , f_HargaSatuan)
        .Parameters.Append .CreateParameter("NoTemporaryOutput", adChar, adParamOutput, 3, Null)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_TemporaryDetailApotikJual"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam pengambilan no temporary", vbCritical, "Validasi"
            sp_TempDetailApotikJual = False
        Else
            txtNoTemporary.Text = Trim(.Parameters("NoTemporaryOutput").Value)
            'Call Add_HistoryLoginActivity("Add_TemporaryDetailApotikJual")
        End If
    End With
End Function

Private Sub txtNoResep_LostFocus()
On Error GoTo errLoad
    If Len(Trim((txtNoResep.Text))) = 0 Then Exit Sub
    strSQL = "SELECT NoResep FROM PemakaianAlkes WHERE (NoResep = '" & txtNoResep.Text & "') AND Year(TglPelayanan) = '" & Year(dtpTglOrder.Value) & "'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        MsgBox "No Resep sudah terpakai, Ganti No Resep", vbExclamation, "Validasi"
        txtNoResep.Text = ""
        txtNoResep.SetFocus
        Call subLoadDataResep(txtNoResep.Text)
        Call subHitungTotal


'        If MsgBox("No Resep " & txtNoResep.Text & " sudah terpakai, ganti No Resep?", vbQuestion + vbYesNo, "Validasi") = vbYes Then
'            txtNoResep.SetFocus
'        Else
'            dtpTglResep.SetFocus
'        End If
    End If
    txtNoResep.Text = StrConv(txtNoResep.Text, vbUpperCase)
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadDataResep(f_NoResep As String)
On Error GoTo errLoad
Dim i As Integer
Dim curHutangPenjamin As Currency
Dim curHarusDibayar As Currency
Dim curTanggunganRS As Currency

    strSQL = "SELECT * FROM V_AmbilPemakaianAlkesResep WHERE NoResep = '" & f_NoResep & "'"
    Call msubRecFO(dbRst, strSQL)
    If dbRst.EOF = True Then
        Call subSetGrid
        dtpTglResep.Value = Now
        chkDokterPemeriksa.Value = vbUnchecked
        txtRP.Text = ""
        Exit Sub
    End If
    
    dtpTglResep.Value = dbRst("TglResep")
    If IsNull(dbRst("IdDokter")) Then
        chkDokterPemeriksa.Value = vbUnchecked
        txtKdDokter.Text = ""
        txtDokter.Text = ""
    Else
        chkDokterPemeriksa.Value = vbChecked
        txtKdDokter.Text = dbRst("IdDokter")
        txtDokter.Text = dbRst("Dokter")
    End If
    dgDokter.Visible = False
    txtRP.Text = dbRst("RuanganResep")
   
    For i = 0 To dbRst.RecordCount - 1
        'ambil no temporary
        txtKdBarang.Text = dbRst("KdBarang")
        txtKdAsal.Text = dbRst("KdAsal")
        If sp_TempDetailApotikJual(CDbl(dbRst("HargaSatuan")) + CDbl(dbRst("TarifService")), dbRst("KdBarang"), dbRst("KdAsal")) = False Then Exit Sub 'discount
        'ambil hutang penjamin dan tanggungan rs
        strSQL = "SELECT HargaSatuan, JmlHutangPenjamin, JmlTanggunganRS" & _
            " FROM TempDetailApotikJual" & _
            " WHERE (NoTemporary = '" & Trim(txtNoTemporary.Text) & "') AND (KdBarang = '" & dbRst("KdBarang") & "') AND (KdAsal = '" & dbRst("KdAsal") & "')"
        Call msubRecFO(rsB, strSQL)
        If rsB.EOF = True Then
            curHutangPenjamin = 0
            curHarusDibayar = 0
        Else
            curHutangPenjamin = rsB("JmlHutangPenjamin").Value
            curHarusDibayar = rsB("JmlTanggunganRS").Value
        End If
    
        With fgData
            .TextMatrix(.Rows - 1, 0) = dbRst("ResepKe")
            .TextMatrix(.Rows - 1, 1) = dbRst("JenisObat")
            .TextMatrix(.Rows - 1, 2) = dbRst("KdBarang")
            .TextMatrix(.Rows - 1, 3) = dbRst("NamaBarang")
            .TextMatrix(.Rows - 1, 4) = dbRst("KeKuatan")
            .TextMatrix(.Rows - 1, 5) = dbRst("NamaAsal")
            .TextMatrix(.Rows - 1, 6) = dbRst("SatuanJml")
            .TextMatrix(.Rows - 1, 7) = CDbl(dbRst("HargaSatuan")) + IIf(dbRst("JmlService") = 0, 0, dbRst("TarifService"))
            .TextMatrix(.Rows - 1, 7) = IIf(Val(.TextMatrix(.Rows - 1, 7)) = 0, 0, Format(.TextMatrix(.Rows - 1, 5), "#,###"))
            .TextMatrix(.Rows - 1, 8) = CDbl(0) 'discount
            .TextMatrix(.Rows - 1, 8) = IIf(Val(.TextMatrix(.Rows - 1, 6)) = 0, 0, Format(.TextMatrix(.Rows - 1, 6), "#,###"))
            .TextMatrix(.Rows - 1, 9) = CDbl(dbRst("JmlStok") + dbRst("JmlBarang"))
            .TextMatrix(.Rows - 1, 10) = CDbl(dbRst("JmlBarang"))
            
            'total harga = ((tarifservice * jmlservice) + _
                (hargasatuan(sebelum ditambah tarifservixe) * jumlah))
            .TextMatrix(.Rows - 1, 11) = ((dbRst("TarifService") * dbRst("JmlService")) + _
                (CDbl(dbRst("HargaSatuan")) * CDbl(.TextMatrix(.Rows - 1, 10))))
            .TextMatrix(.Rows - 1, 11) = IIf(Val(.TextMatrix(.Rows - 1, 11)) = 0, 0, Format(.TextMatrix(.Rows - 1, 11), "#,###"))
            
            .TextMatrix(.Rows - 1, 12) = dbRst("KdAsal")
            .TextMatrix(.Rows - 1, 13) = dbRst("JenisBarang")
            .TextMatrix(.Rows - 1, 14) = dbRst("TarifService")
            .TextMatrix(.Rows - 1, 15) = dbRst("JmlService")
            .TextMatrix(.Rows - 1, 16) = CDbl(dbRst("HargaSatuan"))
            .TextMatrix(.Rows - 1, 17) = curHutangPenjamin
            .TextMatrix(.Rows - 1, 18) = curTanggunganRS
            
            .TextMatrix(.Rows - 1, 19) = CDbl(dbRst("JmlBarang")) * curHutangPenjamin
            .TextMatrix(.Rows - 1, 20) = CDbl(dbRst("JmlBarang")) * curTanggunganRS
            .TextMatrix(.Rows - 1, 21) = CDbl(dbRst("JmlBarang")) * CDbl(0) 'discount
            
            'total harus dibayar = total harga - total discount - _
                total hutang penjamin - totaltanggunganrs
            curHarusDibayar = CDbl(.TextMatrix(.Rows - 1, 11)) - (CDbl(.TextMatrix(.Rows - 1, 21)) + _
                CDbl(.TextMatrix(.Rows - 1, 19)) + CDbl(.TextMatrix(.Rows - 1, 120)))
            .TextMatrix(.Rows - 1, 22) = IIf(curHarusDibayar < 0, 0, curHarusDibayar)
            
            .TextMatrix(.Rows - 1, 23) = txtNoTemporary.Text
            
            .TextMatrix(.Rows - 1, 24) = CDbl(dbRst("HargaBeli"))
            .TextMatrix(.Rows - 1, 25) = IIf(IsNull(dbRst("KdJenisObat")), "", dbRst("KdJenisObat"))
            .TextMatrix(.Rows - 1, 26) = dbRst("BiayaAdministrasi")

            .Rows = .Rows + 1
            dbRst.MoveNext
            dbConn.Execute "DELETE FROM TempDetailApotikJual WHERE (NoTemporary = '" & Trim(txtNoTemporary.Text) & "')"
        End With
    Next i
    
    Call subHitungTotal
    
    'txtNamaBarang.Text = "": txtHargaBeli.Text = 0: txtHargaSatuan.Text = 0: txtDiscount.Text = 0: txtStock.Text = 0: txtJumlah.Text = 0
    dgObatAlkes.Visible = False
    txtJenisBarang.Text = "": txtKdBarang.Text = "": txtKdAsal.Text = "": txtSatuan.Text = "": txtAsalBarang.Text = ""
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadText()
Dim i As Integer
    txtIsi.Left = fgData.Left
    
    Select Case fgData.Col
        Case 0
            txtIsi.MaxLength = 2
        
        Case 3
            txtIsi.MaxLength = 20
        
        Case 10
            txtIsi.MaxLength = 7
    End Select
    
    For i = 0 To fgData.Col - 1
        txtIsi.Left = txtIsi.Left + fgData.ColWidth(i)
    Next i
    txtIsi.Visible = True
    txtIsi.Top = fgData.Top - 7
    
    For i = 0 To fgData.Row - 1
        txtIsi.Top = txtIsi.Top + fgData.RowHeight(i)
    Next i
    
    If fgData.TopRow > 1 Then
        txtIsi.Top = txtIsi.Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
    End If
    
    txtIsi.Width = fgData.ColWidth(fgData.Col)
'    txtIsi.Height = fgData.RowHeight(fgData.Row)
    
    txtIsi.Visible = True
    txtIsi.SelStart = Len(txtIsi.Text)
    txtIsi.SetFocus
End Sub
Private Sub subLoadTextRacikan()
Dim i As Integer
    TxtIsiRacikan.Left = fgRacikan.Left
    
    Select Case fgRacikan.Col
        Case 0
            TxtIsiRacikan.MaxLength = 2
        
        Case 3
            TxtIsiRacikan.MaxLength = 20
        
        Case 10
            TxtIsiRacikan.MaxLength = 4
    End Select
    
    For i = 0 To fgRacikan.Col - 1
        TxtIsiRacikan.Left = TxtIsiRacikan.Left + fgRacikan.ColWidth(i)
    Next i
    TxtIsiRacikan.Visible = True
    TxtIsiRacikan.Top = fgRacikan.Top - 7
    
    For i = 0 To fgRacikan.Row - 1
        TxtIsiRacikan.Top = TxtIsiRacikan.Top + fgRacikan.RowHeight(i)
    Next i
    
    If fgData.TopRow > 1 Then
        TxtIsiRacikan.Top = TxtIsiRacikan.Top - ((fgRacikan.TopRow - 1) * fgRacikan.RowHeight(1))
    End If
    
    TxtIsiRacikan.Width = fgRacikan.ColWidth(fgRacikan.Col)
'    txtIsi.Height = fgData.RowHeight(fgData.Row)
    
    TxtIsiRacikan.Visible = True
    TxtIsiRacikan.SelStart = Len(TxtIsiRacikan.Text)
    TxtIsiRacikan.SetFocus
End Sub
Private Sub subLoadDataCombo(s_DcName As Object)
Dim i As Integer
    s_DcName.Left = fgData.Left
    For i = 0 To fgData.Col - 1
        s_DcName.Left = s_DcName.Left + fgData.ColWidth(i)
    Next i
    s_DcName.Visible = True
    s_DcName.Top = fgData.Top - 7
    
    For i = 0 To fgData.Row - 1
        s_DcName.Top = s_DcName.Top + fgData.RowHeight(i)
    Next i
    
    If fgData.TopRow > 1 Then
        s_DcName.Top = s_DcName.Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
    End If
    
    s_DcName.Width = fgData.ColWidth(fgData.Col)
    s_DcName.Height = fgData.RowHeight(fgData.Row)
    
    'untuk mengetahui posisi dcjenisObat pada row
    posisiRowDataComboJenisObat = fgData.Row
    
    s_DcName.Visible = True
    s_DcName.SetFocus
End Sub

Private Sub txtRP_Change()
  'Update 15-05-06 JSPRJ
    If Len(Trim(txtRP.Text)) = 0 Then StrKdRP = "": Exit Sub
    strSQL = "Select KdRuangan from Ruangan  where NamaRuangan='" & txtRP.Text & "'"
    Call msubRecFO(rs, strSQL)
    StrKdRP = IIf(IsNull(rs.Fields(0).Value), "", rs.Fields(0).Value)
End Sub
''========================add by Asep Nur Iman 2013=====================
Private Sub subLoadResep()
On Error GoTo aneh
    Dim rKe As String
    Set rs = Nothing
    strSQL = "Select max(ResepKe) from DetailOrderPelayananOA where NoPendaftaran = '" & txtNoPendaftaranOA.Text & "'"
    Call msubRecFO(rs, strSQL)
    If Not rs.EOF Then
        rKe = rs(0).Value + 1
        fgData.TextMatrix(1, 0) = rKe
    Else
        fgData.TextMatrix(1, 0) = "1"
    End If
    
    Exit Sub
aneh:
    fgData.TextMatrix(1, 0) = "1"
End Sub

Private Sub rebuildResekKe()
    Dim i As Integer
    Dim j As Integer
    
    Dim resepSebelumnya As Integer
    resepSebelumnya = 0
    For i = 1 To fgData.Rows - 1
        Dim resepBerikutnya As Integer
  '      fgData.TextMatrix(i, 0) = 2
        resepBerikutnya = IIf(fgData.TextMatrix(i, 0) = "", "0", fgData.TextMatrix(i, 0))
        
        
        fgData.Row = i
        
        If (resepSebelumnya + 1 <> resepBerikutnya And fgData.CellBackColor <> vbRed) Then
            resepSebelumnya = resepSebelumnya + 1
            If (resepSebelumnya = 0) Then
                MsgBox "asd"
            End If
            fgData.TextMatrix(i, 0) = resepSebelumnya
            
        Else
            If (i = 1 And resepSebelumnya = 0) Then
                resepSebelumnya = 1
                 For j = 1 To fgData.Rows - 1
                    fgData.TextMatrix(j, 0) = 0
                 Next
            End If
            fgData.TextMatrix(i, 0) = resepSebelumnya
        End If
        
        If fgData.TextMatrix(i, 25) = "01" Then
            
'            dbConn.Execute "update DetailOrderPelayananOARacikanTemp set resepke='" & Val(fgData.TextMatrix(i, 0)) - 1 & "' where noracikan='" & fgData.TextMatrix(i, 34) & "' " ' and resepke='" & fgData.TextMatrix(i, 0) & "'"
            dbConn.Execute "update DetailOrderPelayananOARacikanTemp set resepke='" & resepSebelumnya & "' where noracikan='" & fgData.TextMatrix(i, 34) & "' " ' and resepke='" & fgData.TextMatrix(i, 0) & "'"
        
        End If
  
    Next i
End Sub

Private Sub clearRincianDetail()
    Dim pesangantiRuanganTujuan As String
    Dim i As Integer
    If fgData.TextMatrix(1, 1) = "" Then Exit Sub
    pesangantiRuanganTujuan = MsgBox("Pilih Yes Jika akan menghapus detail Pemesanan Obat?", vbInformation + vbYesNo)
    
If pesangantiRuanganTujuan = vbYes Then
    
    If FraRacikan.Visible = True Then
        Call cmdBatal_Click
    End If
        
    With fgData
        For i = 1 To .Rows - 1
        If .TextMatrix(i, 1) <> "" Then
            If .TextMatrix(i, 1) = "Racikan" Then
                StrSQL12 = "delete DetailOrderPelayananOARacikanTemp where NoRacikan='" & .TextMatrix(i, 34) & "'"
            End If
        End If
        Next i
    End With

    txtNoResep.Text = ""
    txtTotalBiaya.Text = 0
    txtTotalDiscount.Text = 0
    txtHutangPenjamin.Text = 0
    txtTanggunganRS.Text = 0
    txtHarusDibayar.Text = 0
    txtNoOrder.Text = ""
    
    
    Call subSetGrid
Else
'    dcRuanganTujuan.BoundText = "702"
'    Exit Sub
End If
    
    


End Sub

