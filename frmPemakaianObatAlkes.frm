VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPemakaianObatAlkes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pemakaian Obat & Alkes"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "frmPemakaianObatAlkes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   17895
   Begin VB.Frame FraRacikan 
      Caption         =   "Obat Racikan"
      Height          =   5655
      Left            =   0
      TabIndex        =   52
      Top             =   2040
      Visible         =   0   'False
      Width           =   17895
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
         TabIndex        =   59
         Top             =   1440
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame fraHitungObat 
         Caption         =   "Berat Obat"
         Height          =   735
         Left            =   7200
         TabIndex        =   57
         Top             =   2400
         Visible         =   0   'False
         Width           =   1815
         Begin VB.TextBox txtBeratObat 
            Height          =   375
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   1575
         End
      End
      Begin MSDataGridLib.DataGrid dgObatAlkesRacikan 
         Height          =   2145
         Left            =   1320
         TabIndex        =   67
         Top             =   2160
         Visible         =   0   'False
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   3784
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
         TabIndex        =   56
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
         Left            =   14280
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   4200
         Width           =   1455
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   495
         Left            =   15840
         TabIndex        =   54
         Top             =   4200
         Width           =   1455
      End
      Begin VB.TextBox txtJumlahObatRacik 
         Height          =   375
         Left            =   1800
         TabIndex        =   53
         Top             =   240
         Width           =   1935
      End
      Begin MSFlexGridLib.MSFlexGrid fgRacikan 
         Height          =   3375
         Left            =   120
         TabIndex        =   66
         Top             =   720
         Width           =   17175
         _ExtentX        =   30295
         _ExtentY        =   5953
         _Version        =   393216
      End
      Begin VB.Label lblJumlahObatRacik 
         Caption         =   "Jumlah Obat Racik"
         Height          =   375
         Left            =   120
         TabIndex        =   60
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   10920
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1920
      Top             =   600
   End
   Begin VB.PictureBox picPenerimaanSementara 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   2400
      ScaleHeight     =   2025
      ScaleWidth      =   7425
      TabIndex        =   42
      Top             =   -1680
      Visible         =   0   'False
      Width           =   7455
      Begin VB.Frame Frame6 
         Caption         =   "Penerimaan Sementara"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   43
         Top             =   120
         Width           =   7215
         Begin VB.CommandButton cmdSimpanTerimaBarang 
            Caption         =   "&Simpan"
            Height          =   495
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   1200
            Width           =   6735
         End
         Begin VB.TextBox txtNamaBarangPenerimaan 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   2160
            TabIndex        =   45
            Top             =   360
            Width           =   4815
         End
         Begin VB.TextBox txtJmlTerima 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            TabIndex        =   44
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   720
            Width           =   1095
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   6960
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Terima Barang"
            Height          =   210
            Index           =   31
            Left            =   240
            TabIndex        =   48
            Top             =   720
            Width           =   1785
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Barang"
            Height          =   210
            Index           =   32
            Left            =   240
            TabIndex        =   47
            Top             =   360
            Width           =   1065
         End
      End
   End
   Begin MSDataGridLib.DataGrid dgObatAlkes 
      Height          =   2535
      Left            =   1920
      TabIndex        =   10
      Top             =   -1920
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   4471
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
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   0
      TabIndex        =   36
      Top             =   6720
      Width           =   17895
      Begin VB.TextBox txtFormPengirimText 
         Height          =   375
         Left            =   1080
         TabIndex        =   69
         Top             =   360
         Visible         =   0   'False
         Width           =   2175
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
         Left            =   14280
         TabIndex        =   38
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
         Left            =   15960
         TabIndex        =   37
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   23
      Top             =   5760
      Width           =   17895
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
         Left            =   15240
         MaxLength       =   12
         TabIndex        =   64
         Text            =   "0"
         Top             =   480
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
         Left            =   7320
         MaxLength       =   12
         TabIndex        =   29
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
         Left            =   12600
         MaxLength       =   12
         TabIndex        =   28
         Text            =   "0"
         Top             =   480
         Width           =   2415
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
         Left            =   9960
         MaxLength       =   12
         TabIndex        =   27
         Text            =   "0"
         Top             =   480
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
         TabIndex        =   26
         Text            =   "0"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2295
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
         TabIndex        =   25
         Text            =   "0"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2295
      End
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
         Left            =   7440
         MaxLength       =   12
         TabIndex        =   24
         Text            =   "0"
         Top             =   480
         Width           =   2415
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
         Left            =   15240
         TabIndex        =   65
         Top             =   240
         Width           =   1695
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
         Left            =   7320
         TabIndex        =   35
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
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
         Left            =   12840
         TabIndex        =   34
         Top             =   240
         Width           =   1860
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
         Left            =   9960
         TabIndex        =   33
         Top             =   240
         Width           =   1950
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
         TabIndex        =   32
         Top             =   1440
         Visible         =   0   'False
         Width           =   1140
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
         TabIndex        =   31
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
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
         Left            =   8880
         TabIndex        =   30
         Top             =   240
         Width           =   945
      End
   End
   Begin MSDataGridLib.DataGrid dgDokter 
      Height          =   2295
      Left            =   11760
      TabIndex        =   8
      Top             =   -1920
      Width           =   6135
      _ExtentX        =   10821
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
   Begin VB.Frame Frame8 
      DragMode        =   1  'Automatic
      Height          =   3855
      Left            =   0
      TabIndex        =   11
      Top             =   1920
      Width           =   17895
      Begin MSDataListLib.DataCombo dcNamaPelayananRS 
         Height          =   330
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcKeteranganPakai2 
         Height          =   330
         Left            =   6720
         TabIndex        =   63
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcKeteranganPakai 
         Height          =   330
         Left            =   5160
         TabIndex        =   62
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcAturanPakai 
         Height          =   330
         Left            =   3600
         TabIndex        =   61
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.CheckBox chkStatusStok 
         Caption         =   "Ya"
         Height          =   495
         Left            =   6120
         TabIndex        =   49
         Top             =   1440
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtIsi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   1320
         TabIndex        =   40
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtNoTemporary 
         Height          =   315
         Left            =   7080
         TabIndex        =   22
         Top             =   2280
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtHargaBeli 
         Height          =   315
         Left            =   4320
         TabIndex        =   21
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtJenisBarang 
         Height          =   315
         Left            =   3000
         TabIndex        =   20
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtKdDokter 
         Height          =   315
         Left            =   1560
         TabIndex        =   19
         Top             =   2280
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtAsalBarang 
         Height          =   315
         Left            =   6000
         TabIndex        =   18
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtKdAsal 
         Height          =   315
         Left            =   2760
         TabIndex        =   14
         Top             =   2280
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtSatuan 
         Height          =   315
         Left            =   4920
         TabIndex        =   13
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtKdBarang 
         Height          =   315
         Left            =   3720
         TabIndex        =   12
         Top             =   2280
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo dcJenisObat 
         Height          =   330
         Left            =   2280
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   3375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   17535
         _ExtentX        =   30930
         _ExtentY        =   5953
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   -2147483643
         FocusRect       =   2
         HighLight       =   2
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame2 
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
      TabIndex        =   15
      Top             =   960
      Width           =   17895
      Begin VB.CommandButton cmdGenerik 
         BackColor       =   &H80000004&
         Caption         =   "Daftar Subtitusi Barang"
         Height          =   495
         Left            =   14520
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   360
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.TextBox txtRP 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   11520
         TabIndex        =   6
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtDokter 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7560
         TabIndex        =   5
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
         TabIndex        =   2
         Top             =   480
         Width           =   2415
      End
      Begin VB.CheckBox chkDokterPemeriksa 
         Caption         =   "Dokter Penulis Resep"
         Height          =   255
         Left            =   7560
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
      Begin VB.CheckBox chkNoResep 
         Caption         =   "No. Resep"
         Height          =   255
         Left            =   3240
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpTglPelayanan 
         Height          =   330
         Left            =   960
         TabIndex        =   0
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
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
         CustomFormat    =   "dd/MM/yyyy HH:mm:ss"
         Format          =   117768195
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSComCtl2.DTPicker dtpTglResep 
         Height          =   330
         Left            =   5880
         TabIndex        =   3
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   117768195
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruangan Perawatan"
         Height          =   210
         Index           =   10
         Left            =   11520
         TabIndex        =   39
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Resep"
         Height          =   210
         Index           =   1
         Left            =   5880
         TabIndex        =   17
         Top             =   240
         Width           =   870
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Pelayanan"
         Height          =   210
         Index           =   0
         Left            =   960
         TabIndex        =   16
         Top             =   240
         Width           =   1185
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   50
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
      Left            =   16200
      Picture         =   "frmPemakaianObatAlkes.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1755
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPemakaianObatAlkes.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPemakaianObatAlkes.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16095
   End
End
Attribute VB_Name = "frmPemakaianObatAlkes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim noRacikan As Integer
Public NotUseRacikan As Boolean
Dim useGeneric As Boolean
Dim tempNoRacikan As Integer
Dim Cancel                    As Boolean

Dim boolStatusSimpanRacikan   As Boolean

Dim boolUpdate                As Boolean

Dim BoolReview                As Boolean

Dim subintJmlArray            As Integer

Dim subcurHargaSatuan         As Currency

Dim subcurTarifService        As Currency

Dim subcurHarusDibayar        As Currency

Dim uniqeId                   As String

Dim curTanggunganRS           As Currency

Dim curHutangPenjamin         As Currency

Dim subintJmlService          As Integer

Dim tempStatusTampil          As Boolean

Dim subJenisHargaNetto        As Integer

Dim subcurTarifServiceRacikan As Currency

Dim subintJmlServiceRacikan   As Currency

Dim cHargabeli                As Currency

Dim iJmlService               As Integer

Dim cTrfService               As Currency

Dim StrKdRP                   As String

Dim strNoRacikan              As String

Dim blt                       As Integer

Dim statTampil                As Boolean

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
        
    Next i
End Sub
Private Sub CountingBiayaAdministrasi()
With fgData
Dim i As Integer
Dim barisValidForCountingResep As Integer
Dim tempSettingDataPendukung As Integer
For i = 1 To .Rows - 1
                        .Row = i
                        If (.CellBackColor <> vbRed) Then
                            barisValidForCountingResep = barisValidForCountingResep + 1
                            .TextMatrix(i, 26) = 0
                        End If
                        

                        If barisValidForCountingResep = (typSettingDataPendukung.intJumlahBAdminOAPerBaris * tempSettingDataPendukung) + 1 Then
                                tempSettingDataPendukung = tempSettingDataPendukung + 1
                                .TextMatrix(i, 26) = typSettingDataPendukung.curBiayaAdministrasi
                        End If
                HitungHutangPenjaminDanTanggunganRs
                Next i
End With
Call subHitungTotal
End Sub

Private Sub HitungHutangPenjaminDanTanggunganRs()
With fgData
If Val(.TextMatrix(.Row, 17)) > 0 Then
      
                        .TextMatrix(.Row, 19) = .TextMatrix(.Row, 11)
        
                        '.TextMatrix(.Row, 19) = (.TextMatrix(.Row, 14) * .TextMatrix(.Row, 15)) + (val(.TextMatrix(.Row, 10)) * CDbl(.TextMatrix(.Row, 17))) + val(.TextMatrix(.Row, 26))
                Else
                        .TextMatrix(.Row, 19) = 0
                End If
    
                If Val(.TextMatrix(.Row, 18)) > 0 Then
                        .TextMatrix(.Row, 20) = .TextMatrix(.Row, 11)
                        '.TextMatrix(.Row, 20) = (.TextMatrix(.Row, 14) * .TextMatrix(.Row, 15)) + (val(.TextMatrix(.Row, 10)) * CDbl(.TextMatrix(.Row, 18))) + val(.TextMatrix(.Row, 26))
                Else
                        .TextMatrix(.Row, 20) = 0
                End If

                .TextMatrix(.Row, 21) = CDbl(.TextMatrix(.Row, 10)) * CDbl(.TextMatrix(.Row, 8))
                .TextMatrix(.Row, 11) = ((CDbl(.TextMatrix(.Row, 14)) * CDbl(IIf(.TextMatrix(.Row, 15) = "", "0", .TextMatrix(.Row, 15)))) + (CDbl(.TextMatrix(.Row, 16)) * Val(.TextMatrix(.Row, 10)))) ' + val(.TextMatrix(.Row, 26))
                'total harus dibayar = total harga - total discount - _
                 total hutang penjamin - totaltanggunganrs
                 Dim curHarusDibayar As Double
                curHarusDibayar = CDbl(.TextMatrix(.Row, 11)) - (CDbl(.TextMatrix(.Row, 21)) + CDbl(.TextMatrix(.Row, 19)) + CDbl(.TextMatrix(.Row, 20)))
                .TextMatrix(.Row, 22) = FormatPembulatan(IIf(curHarusDibayar < 0, 0, curHarusDibayar), mstrKdInstalasiLogin)
End With
End Sub
Public Function sp_StokRuangan(f_KdBarang As String, _
                               f_KdAsal As String, _
                               f_JmlBarang As Double, _
                               f_status As String) As Boolean

        On Error GoTo Errload

        Dim i As Integer

        sp_StokRuangan = True
        Set dbcmd = New ADODB.Command

        With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
                .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
                .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
                .Parameters.Append .CreateParameter("JmlBrg", adDouble, adParamInput, , f_JmlBarang)
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

                .ActiveConnection = dbConn
                .CommandText = "dbo.Update_StokRuangan"
                .CommandType = adCmdStoredProc
                .Execute

                If .Parameters("return_value").Value <> 0 Then
                        MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "validasi"
                        sp_StokRuangan = False
                Else
                        Call Add_HistoryLoginActivity("Update_StokRuangan")
                End If

        End With

        Set dbcmd = Nothing

        Exit Function

Errload:
        sp_StokRuangan = False
        Call msubPesanError
End Function

Public Function sp_StokRealRuangan(f_KdBarang As String, _
                                   f_KdAsal As String, _
                                   f_NoTerima As String, _
                                   f_JmlBarang As Double, _
                                   f_status As String) As Boolean

        On Error GoTo Errload

        Dim i As Integer

        sp_StokRealRuangan = True
        Set dbcmd = New ADODB.Command

        With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
                If (FraRacikan.Visible = True) Then
                    Dim racikanId As String
                    racikanId = uniqeId
                    If (noRacikan < 10) Then
                        Mid(racikanId, 1, 1) = noRacikan
                        Mid(racikanId, 20, 1) = noRacikan
                    Else
                        Mid(racikanId, 1, 2) = noRacikan
                        Mid(racikanId, 20, 2) = noRacikan
                    End If
                    
                    .Parameters.Append .CreateParameter("UniqueId", adVarChar, adParamInput, 32, racikanId)
                 Else
                    .Parameters.Append .CreateParameter("UniqueId", adVarChar, adParamInput, 32, uniqeId)
                 End If
                .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
                .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
                .Parameters.Append .CreateParameter("NoTerima", adChar, adParamInput, 10, f_NoTerima)
                .Parameters.Append .CreateParameter("JmlBrg", adDouble, adParamInput, , CDbl(f_JmlBarang))
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

                .ActiveConnection = dbConn
                .CommandText = "Update_StokRuanganDynamic"
                .CommandType = adCmdStoredProc
                .Execute

                If .Parameters("return_value").Value <> 0 Then
                        MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "validasi"
                        sp_StokRealRuangan = False
                End If

        End With

        Set dbcmd = Nothing
        Call deleteADOCommandParameters(dbcmd)

        Exit Function

Errload:
        sp_StokRealRuangan = False
        Call msubPesanError("sp_StokRuangan")
End Function

Private Function sp_GenerateNoResep() As Boolean

        On Error GoTo Errload

        sp_GenerateNoResep = True
        Set dbcmd = New ADODB.Command

        With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("NoResep", adVarChar, adParamInput, 15, Trim(txtNoResep.Text))
                .Parameters.Append .CreateParameter("TglResep", adDate, adParamInput, , Format(dtpTglResep.Value, "yyyy/MM/dd"))
                .Parameters.Append .CreateParameter("OutputNoResep", adVarChar, adParamOutput, 15, Null)

                .ActiveConnection = dbConn
                .CommandText = "dbo.AU_GenerateNoResep"
                .CommandType = adCmdStoredProc
                .Execute

                If .Parameters("return_value") <> 0 Then
                        MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
                        sp_GenerateNoResep = False
                Else
                        txtNoResep.Text = Trim(.Parameters("OutputNoResep"))

                End If

        End With

        Exit Function

Errload:
        sp_GenerateNoResep = False
        Call msubPesanError("sp_GenerateNoResep")
End Function

Private Sub chkDokterPemeriksa_Click()

        On Error GoTo Errload

        If chkDokterPemeriksa.Value = vbUnchecked Then
                txtDokter.Enabled = False
                txtDokter.Text = ""
        Else
                txtDokter.Enabled = True
                txtDokter.Text = mstrNamaDokter
                txtKdDokter.Text = mstrKdDokter
        End If

        dgDokter.Visible = False

        Exit Sub

Errload:
        Call msubPesanError
End Sub

Private Sub chkDokterPemeriksa_KeyPress(KeyAscii As Integer)

        On Error Resume Next

        If KeyAscii = 13 Then
                If chkDokterPemeriksa.Value = vbChecked Then txtDokter.SetFocus Else fgData.SetFocus
        End If

End Sub

Private Sub chkNoResep_Click()

        If chkNoResep.Value = vbChecked Then
                If (NotUseRacikan = False) Then
                    txtNoResep.Enabled = True
                End If
                dtpTglResep.Enabled = True
                chkDokterPemeriksa.Enabled = True
                txtDokter.Enabled = True
        Else
                txtNoResep.Enabled = False
                dtpTglResep.Enabled = False
                chkDokterPemeriksa.Enabled = False
                txtDokter.Enabled = False
                chkDokterPemeriksa.Value = vbUnchecked
        End If

        dgDokter.Visible = False
End Sub

Private Sub chkNoResep_KeyPress(KeyAscii As Integer)

        If KeyAscii = 13 Then
                If txtNoResep.Enabled = True Then txtNoResep.SetFocus Else fgData.SetFocus
        End If

End Sub

Private Sub chkStatusStok_KeyPress(KeyAscii As Integer)

        If KeyAscii = 13 Then
                chkStatusStok.Visible = False
                fgData.TextMatrix(fgData.Row, fgData.Col) = IIf(chkStatusStok.Value = vbChecked, "Ya", "Tdk")

                With fgData

                        If .RowPos(.Row) >= .Height - 360 Then
                                .SetFocus
                                SendKeys "{DOWN}"

                                Exit Sub

                        End If

                        .SetFocus

                        If fgData.TextMatrix(fgData.Rows - 1, 2) <> "" Then
                                fgData.Rows = fgData.Rows + 1

                                If .TextMatrix(.Rows - 2, 25) = "" Then
                                        .TextMatrix(.Rows - 1, 0) = "1"
                                ElseIf .TextMatrix(.Rows - 2, 25) = "01" Then
                                        .TextMatrix(.Rows - 1, 0) = "1"
                                Else
                                        .TextMatrix(.Rows - 1, 0) = "0"
                                End If
                        End If

                        fgData.SetFocus
                        fgData.Row = fgData.Rows - 1
                        fgData.Col = 0
                End With

        End If

End Sub

Private Sub chkStatusStok_LostFocus()
        chkStatusStok.Visible = False
End Sub

Private Sub cmdBatal_Click()
Dim i As Integer
                i = noRacikan
        If boolStatusSimpanRacikan <> False Then

                With fgData

                        If .TextMatrix(.Row, 3) = "" Then
                                FraRacikan.Visible = False
                                Call ClearFgRacikan
                                .TextMatrix(.Row, 0) = ""
                                .TextMatrix(.Row, 1) = ""
                                .TextMatrix(.Row, 2) = ""
                                .TextMatrix(.Row, 25) = ""
                        Else
                                FraRacikan.Visible = False
                                Call ClearFgRacikan
                        End If

                End With

        Else

                '    dbConn.Execute "DELETE FROM RacikanObatPasienTemp " & _
                '            " WHERE (NoRacikan = '" & fgData.TextMatrix(fgData.Row, 34) & "') AND (NoPendaftaran = '" & mstrNoPen & "')"
                If BoolReview = False Then
                        dbConn.Execute "DELETE FROM RacikanObatPasienTemp " & " WHERE (NoRacikan = '" & fgData.TextMatrix(fgData.Row, 34) & "') AND (NoPendaftaran = '" & mstrNoPen & "')"
                End If
                
                 'For i = 1 To fgRacikan.Rows - 1

                'If fgRacikan.TextMatrix(i, 0) <> "" Then
                       'If sp_StokRealRuangan(fgRacikan.TextMatrix(i, 0), fgRacikan.TextMatrix(i, 12), fgRacikan.TextMatrix(i, 14), fgRacikan.TextMatrix(i, 6), "C") = False Then Exit Sub
                'End If

                'Next i
                
                FraRacikan.Visible = False
                BoolReview = False
                Call ClearFgRacikan
        End If

        txtJumlahObatRacik.Enabled = True
        Dim racikanId As String
                    racikanId = uniqeId
                    If (i < 10) Then
                        Mid(racikanId, 1, 1) = i
                        Mid(racikanId, 20, 1) = i
                    Else
                        Mid(racikanId, 1, 2) = i
                        Mid(racikanId, 20, 2) = i
                    End If
            Call msubRecFO(rs, "execute RestoreStokBarangOtomatis '" & racikanId & "'")
End Sub

Private Sub ClearFgRacikan()

        Call subSetGridRacikan
        txtJumlahObatRacik.Text = ""
        txtBeratObat.Text = ""
End Sub

Private Sub cmdGenerik_Click()

        On Error GoTo Errload
        useGeneric = True
        Dim i As Integer

        strSQL = "execute CariBarangNStokMedis_V_Generik '" & tmpKdBar & "','" & mstrKdRuangan & "'"
        Call msubRecFO(dbRst, strSQL)
    
        Set dgObatAlkes.DataSource = dbRst

        With dgObatAlkes

                For i = 0 To .Columns.Count - 1
                        .Columns(i).Width = 0
                Next i

                .Columns("KdBarang").Width = 1500
                .Columns("NamaBarang").Width = 3000
                .Columns("JenisBarang").Width = 1500
                .Columns("Kekuatan").Width = 1000
                .Columns("AsalBarang").Width = 1000
                .Columns("Satuan").Width = 675

                .Top = 2830
                .Left = 3000
                .Visible = True

                For i = 1 To fgData.Row - 1
                        .Top = .Top + fgData.RowHeight(i)
                Next i

                If fgData.TopRow > 1 Then
                        .Top = .Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
                End If

        End With

        Exit Sub

Errload:
End Sub

Private Sub cmdSelesai_Click()

        On Error GoTo Errload

        Dim i             As Integer

        Dim TotalRacikan  As Currency

        Dim strAsalBarang As String

        Dim strKdAsal     As String

        Dim strSatuan     As String
        
        txtJumlahObatRacik.Enabled = True

        If fgRacikan.TextMatrix(1, 0) = "" Then MsgBox "Transaksi racikan belum diisi", vbExclamation, "Validasi": Exit Sub
        If BoolReview = True Then
                Call msubRecFO(rsB, "Delete from racikanObatPasienTemp Where NoRacikan = '" & fgData.TextMatrix(fgData.Row, 34) & "' and ResepKe = '" & fgData.TextMatrix(fgData.Row, 0) & "'")
                BoolReview = False
        End If

        If txtJumlahObatRacik.Text = "" Or Val(txtJumlahObatRacik.Text) = 0 Then
                MsgBox "Jumlah Obat Racik Harus Di isi", vbCritical, "Informasi"
    
                txtJumlahObatRacik.SetFocus

                Exit Sub

        Else

                With fgRacikan

                        For i = 1 To .Rows - 2

                                If .TextMatrix(i, 4) = "" Or .TextMatrix(i, 5) = "" Or .TextMatrix(i, 6) = "" Then
                                        MsgBox "Data stok/Jumlah belum lengkap", vbCritical, "Validasi"
                                        fgRacikan.SetFocus
                                        fgRacikan.Row = i
                                        fgRacikan.Col = 4

                                        Exit Sub

                                End If

                        Next i

                End With

                With fgRacikan

                        For i = 1 To .Rows - 1

                                If .TextMatrix(i, 0) = "" Then GoTo lanjutkan_
                                If sp_RacikanObatPasienTemp(.TextMatrix(i, 0), .TextMatrix(i, 12), .TextMatrix(i, 13), .TextMatrix(i, 2), .TextMatrix(i, 14), .TextMatrix(i, 7), .TextMatrix(i, 6), .TextMatrix(i, 8), .TextMatrix(i, 4), .TextMatrix(i, 5), .TextMatrix(i, 15), .TextMatrix(i, 16), txtNoRacikan.Text, dcJenisObat.BoundText, txtJumlahObatRacik.Text, mstrKdRuangan, mstrNoPen, "OA") = False Then Exit Sub
                
                                TotalRacikan = TotalRacikan + .TextMatrix(i, 9)
                                strAsalBarang = .TextMatrix(1, 11)
                                strKdAsal = .TextMatrix(1, 12)
                                strSatuan = .TextMatrix(1, 13)
lanjutkan_:
                        Next i
    
                End With

        End If

        cmdTutup.Enabled = True
         If (FraRacikan.Visible = True) Then
                    Dim racikanId As String
                    racikanId = uniqeId
                    If (noRacikan < 10) Then
                        Mid(racikanId, 1, 1) = noRacikan
                        Mid(racikanId, 20, 1) = noRacikan
                    Else
                        Mid(racikanId, 1, 2) = noRacikan
                        Mid(racikanId, 20, 2) = noRacikan
                    End If
                 fgData.TextMatrix(fgData.Row, 38) = racikanId
                 Else
                 End If
        FraRacikan.Visible = False

        With fgData
                .TextMatrix(.Row, 1) = dcJenisObat.Text
                .TextMatrix(.Row, 5) = strAsalBarang
                .TextMatrix(.Row, 6) = strSatuan
                .TextMatrix(.Row, 8) = 0
                .TextMatrix(.Row, 9) = "-"
                .TextMatrix(.Row, 3) = dcJenisObat.Text
                .TextMatrix(.Row, 25) = dcJenisObat.BoundText
    
                If txtJumlahObatRacik.Text = "" Or Val(txtJumlahObatRacik.Text) = 0 Then
                        MsgBox "Jumlah Obat Racik Harus Di isi", vbCritical, "Validasi"
        
                        Exit Sub

                Else
                        .TextMatrix(.Row, 10) = txtJumlahObatRacik.Text
                End If
    
                .TextMatrix(.Row, 11) = TotalRacikan
                .TextMatrix(.Row, 11) = FormatPembulatan(CDbl(.TextMatrix(.Row, 11)), mstrKdInstalasiLogin)
                .TextMatrix(.Row, 12) = strKdAsal
                .TextMatrix(.Row, 7) = (TotalRacikan) / txtJumlahObatRacik.Text
                .TextMatrix(.Row, 7) = FormatPembulatan(CDbl(.TextMatrix(.Row, 7)), mstrKdInstalasiLogin)

                If .TextMatrix(.Row, 7) = "" Then .TextMatrix(.Row, 7) = 0
                .TextMatrix(.Row, 2) = "Racikan"
                .TextMatrix(.Row, 29) = "0000000000"
                .TextMatrix(.Row, 34) = strNoRacikan

                If boolUpdate = False Then
                        '.Rows = .Rows + 1
                Else
                        boolUpdate = False
                End If

        End With

        Call subHitungTotal
        Call hitungRacikan
        MsgBox "Data Racikan Telah Disimpan", vbInformation, "Informasi"
        boolStatusSimpanRacikan = True
        Call subSetGridRacikan
        'Add Arief For Racikan Otomatis
        strNoRacikan = ""
        txtNoRacikan.Text = ""
        'end arief
        txtJumlahObatRacik.Text = ""
        fgData.SetFocus
        fgData.Col = 30
        CountingBiayaAdministrasi
        HitungHutangPenjaminDanTanggunganRs
        Exit Sub

Errload:
        msubPesanError
        boolStatusSimpanRacikan = False
End Sub

Private Sub hitungRacikan()

        On Error Resume Next

        Dim i                 As Integer

        Dim curHutangPenjamin As Currency

        Dim curTanggunganRS   As Currency

        Dim curHarusDibayar   As Currency

        Dim curTempTotal      As Currency

        With fgData

                'ambil no temporary
                If sp_TempDetailApotikJual(CDbl(.TextMatrix(.Row, 7)), .TextMatrix(.Row, 2), .TextMatrix(.Row, 12)) = False Then Exit Sub
                'ambil hutang penjamin dan tanggungan rs
                strSQL = "SELECT HargaSatuan, JmlHutangPenjamin, JmlTanggunganRS" & " FROM TempDetailApotikJual" & " WHERE (NoTemporary = '" & Trim(txtNoTemporary.Text) & "') AND (KdBarang = '" & .TextMatrix(.Row, 2) & "') AND (KdAsal = '" & .TextMatrix(.Row, 12) & "')"
                Call msubRecFO(rs, strSQL)

                If rs.EOF = True Then
                        curHutangPenjamin = 0
                        curTanggunganRS = 0
                Else
                        curHutangPenjamin = rs("JmlHutangPenjamin").Value
                        curTanggunganRS = rs("JmlTanggunganRS").Value
                End If
    
                .TextMatrix(.Row, 16) = CDbl(.TextMatrix(.Row, 7))

                .TextMatrix(.Row, 17) = curHutangPenjamin
                .TextMatrix(.Row, 18) = curTanggunganRS
    
                HitungHutangPenjaminDanTanggunganRs
        End With

        '    Call Hitung
        Call subHitungTotal
    
End Sub


Private Function sp_RacikanObatPasienTemp(f_KdBarang As String, _
                                          f_KdAsal As String, _
                                          f_Satuan As String, _
                                          f_ResepKe As String, _
                                          f_NoTerima As String, _
                                          f_JmlBarang As Single, _
                                          f_JmlPembulatan As Integer, _
                                          f_HargaSatuan As Currency, _
                                          f_kebutuhanML As Double, _
                                          f_kebutuhanTB As Double, _
                                          f_jmlService As Integer, _
                                          f_TarifService As Currency, _
                                          f_Noracikan As String, _
                                          f_KdJenisObat As String, _
                                          f_qtyRacikan As Integer, _
                                          f_KdRuangan As String, _
                                          f_NoPendaftaran As String, _
                                          f_StatusPelyn As String) As Boolean
    
        On Error GoTo Errload
    
        sp_RacikanObatPasienTemp = True
        Set dbcmd = New ADODB.Command

        With dbcmd
                'f_Noracikan = ""
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("NoRacikan", adChar, adParamInput, 10, IIf(Trim(f_Noracikan) = "", Null, Trim(f_Noracikan)))
                .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, IIf(Trim(f_KdBarang) = "", Null, Trim(f_KdBarang)))
                .Parameters.Append .CreateParameter("kdAsal", adChar, adParamInput, 2, IIf(Trim(f_KdAsal) = "", Null, Trim(f_KdAsal)))
                .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, IIf(Trim(f_KdRuangan) = "", Null, Trim(f_KdRuangan)))
                .Parameters.Append .CreateParameter("SatuanJml", adChar, adParamInput, 1, IIf(Trim(f_Satuan) = "", Null, Trim(f_Satuan)))
                .Parameters.Append .CreateParameter("KdJenisObat", adChar, adParamInput, 2, IIf(Trim(f_KdJenisObat) = "", Null, Trim(f_KdJenisObat)))
                .Parameters.Append .CreateParameter("ResepKe", adTinyInt, adParamInput, , IIf(Trim(f_ResepKe) = "", Null, Trim(f_ResepKe)))
                .Parameters.Append .CreateParameter("NoTerima", adChar, adParamInput, 10, IIf(Trim(f_NoTerima) = "", Null, Trim(f_NoTerima)))  ' add no terima
                .Parameters.Append .CreateParameter("JmlBarang", adDouble, adParamInput, , IIf(Trim(f_JmlBarang) = "", Null, CDbl(Trim(f_JmlBarang))))
                .Parameters.Append .CreateParameter("JmlPembulatan", adInteger, adParamInput, , IIf(Trim(f_JmlPembulatan) = "", Null, Trim(f_JmlPembulatan)))
                .Parameters.Append .CreateParameter("qtyRacikan", adInteger, adParamInput, , IIf(IsNull(f_qtyRacikan), 0, f_qtyRacikan))
                Dim Hasil As String
                Hasil = msubKonversiKomaTitik(CStr(f_HargaSatuan))
                .Parameters.Append .CreateParameter("HargaSatuan", adCurrency, adParamInput, , Hasil)
                '.Parameters.Append .CreateParameter("Kekuatan", adDouble, adParamInput, , CDbl(f_kekuatan))
                .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, IIf(mstrNoCM = "", Null, mstrNoCM))
                .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)
                .Parameters.Append .CreateParameter("KebutuhanML", adDouble, adParamInput, , CDbl(f_kebutuhanML))
                
                .Parameters.Append .CreateParameter("KebutuhanTB", adDouble, adParamInput, , CDbl(f_kebutuhanTB))
                .Parameters.Append .CreateParameter("JmlService", adInteger, adParamInput, , f_jmlService) ' add jumlah service
                .Parameters.Append .CreateParameter("TarifService", adCurrency, adParamInput, , f_TarifService) ' add tarif service
                .Parameters.Append .CreateParameter("UniqeId", adVarChar, adParamInput, 32, uniqeId)
                .Parameters.Append .CreateParameter("StatusPelayanan", adChar, adParamInput, 2, f_StatusPelyn) 'add Status Pelayanan
                
        
                .Parameters.Append .CreateParameter("OutputKode", adChar, adParamOutput, 10, Null)
             
                .ActiveConnection = dbConn
                .CommandText = "Add_RacikanObatPasienTemp"
                .CommandType = adCmdStoredProc
                .Execute
        
                If .Parameters("return_value") <> 0 Then
                        MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
                        sp_RacikanObatPasienTemp = False
        
                Else

                        If Not IsNull(.Parameters("OutputKode").Value) Then
                                txtNoRacikan.Text = .Parameters("OutputKode").Value
                                strNoRacikan = .Parameters("OutputKode").Value
                        End If
                End If

                Call deleteADOCommandParameters(dbcmd)
                Set dbcmd = Nothing
        End With
    
        Exit Function

Errload:
        Call msubPesanError
        Call deleteADOCommandParameters(dbcmd)
        ''Resume 0
        sp_RacikanObatPasienTemp = False
End Function

Private Sub subSetGridRacikan()

        On Error GoTo Errload

        With fgRacikan
                .Visible = True
                .Clear
                .Rows = 2
                .Cols = 17
        
                .RowHeight(0) = 400
                .TextMatrix(0, 0) = "" 'KdBarang
                .TextMatrix(0, 1) = "" 'Jenis obat
                .TextMatrix(0, 2) = "R/Ke"
                .TextMatrix(0, 3) = "Nama Barang"
                .TextMatrix(0, 4) = "/Mg /Ml"
                .TextMatrix(0, 5) = "/Tablet"
                .TextMatrix(0, 6) = "Jumlah"
                .TextMatrix(0, 7) = "JmlPembulatan" ' (untuk harga)
                .TextMatrix(0, 8) = "Harga Satuan"
                .TextMatrix(0, 9) = "Total Harga"
                .TextMatrix(0, 10) = "Kekuatan"
                .TextMatrix(0, 11) = "AsalBarang"
                .TextMatrix(0, 12) = "kdAsal"
                .TextMatrix(0, 13) = "satuan"
                .TextMatrix(0, 14) = "NoFIFO"
                .TextMatrix(0, 15) = "jmlService" 'add Column Jumlah Service
                .TextMatrix(0, 16) = "TarifService" ' add Column TarifService

                .ColWidth(0) = 0
                .ColWidth(1) = 0
                .ColWidth(2) = 1000
                .ColWidth(3) = 4800
                .ColWidth(4) = 1800
                .ColWidth(5) = 1200
                .ColWidth(6) = 1800 '0
                .ColWidth(7) = 0 '1200
                .ColWidth(8) = 1800
                .ColWidth(9) = 2000
                .ColWidth(10) = 0
                .ColWidth(11) = 0
                .ColWidth(12) = 0
                .ColWidth(14) = 0
                .ColWidth(15) = 0
                .ColWidth(16) = 0
        
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

Errload:
        Call msubPesanError
End Sub

Private Function sp_RacikanObatPasien(f_Noracikan As String, _
                                      f_KdJenisObat As String, _
                                      f_ResepKe As Integer, _
                                      f_KdSatuanEtiket As String, _
                                      f_KdWaktuEtiket As String, _
                                      f_KdWaktuEtiket2 As String, _
                                      f_Keterangan As String) As Boolean
    
        On Error GoTo Errload
    
        sp_RacikanObatPasien = True
        Set dbcmd = New ADODB.Command

        With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("NoRacikan", adChar, adParamInput, 10, f_Noracikan)
                .Parameters.Append .CreateParameter("NoStruk", adChar, adParamInput, 10, mstrNoStruk)
                If (NotUseRacikan) Then
                    .Parameters.Append .CreateParameter("NoResep", adVarChar, adParamInput, 15, Null)
                Else
                    .Parameters.Append .CreateParameter("NoResep", adVarChar, adParamInput, 15, txtNoResep.Text)
                End If
                
                .Parameters.Append .CreateParameter("tglPelayanan", adDate, adParamInput, , Format(dtpTglPelayanan.Value, "yyyy/mm/dd"))
                .Parameters.Append .CreateParameter("KdJenisObat", adChar, adParamInput, 2, f_KdJenisObat)
                .Parameters.Append .CreateParameter("ResepKe", adInteger, adParamInput, , f_ResepKe)
                .Parameters.Append .CreateParameter("KdSatuanEtiket", adChar, adParamInput, 2, IIf(Len(Trim(f_KdSatuanEtiket)) = 0, Null, f_KdSatuanEtiket)) 'allow null
                .Parameters.Append .CreateParameter("KdWaktuEtiket", adChar, adParamInput, 2, IIf(Len(Trim(f_KdWaktuEtiket)) = 0, Null, f_KdWaktuEtiket))
                .Parameters.Append .CreateParameter("KdWaktuEtiket2", adChar, adParamInput, 2, IIf(Len(Trim(f_KdWaktuEtiket)) = 0, Null, f_KdWaktuEtiket2))
                .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 75, IIf(f_Keterangan = "", Null, f_Keterangan))
        
                .ActiveConnection = dbConn
                .CommandText = "dbo.Add_RacikanObatPasien"
                .CommandType = adCmdStoredProc
                .Execute
        
                If .Parameters("return_value") <> 0 Then
                        MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
                        sp_RacikanObatPasien = False

                End If

        End With

        Set dbcmd = Nothing
        Call deleteADOCommandParameters(dbcmd)

        Exit Function

Errload:
        sp_RacikanObatPasien = False
        Call msubPesanError("sp_RacikanObatPasien")
End Function

Private Sub cmdSimpan_Click()

        On Error GoTo Errload

        Dim i, j, a As Integer
        If (chkDokterPemeriksa.Value = 1) Then
              If txtKdDokter.Text = "" Then
                MsgBox "Dokter Belum di pilih"
                txtDokter.SetFocus
                Exit Sub
              End If
        End If
    If fgData.TextMatrix(1, 2) = "" Then MsgBox "Data Barang Masih Kosong", vbExclamation, "Validasi": Exit Sub
    With fgData
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 2) = "" Then GoTo lanjut0_
            If .TextMatrix(i, 10) = "" Or .TextMatrix(i, 10) = "0" Then MsgBox "Jumlah Tidak Boleh Nol Atau Kosong", vbExclamation, "Validasi": .Col = 10: .SetFocus: Exit Sub
            If .TextMatrix(i, 1) = "" Then MsgBox "Jenis Obat Masih Kosong", vbExclamation, "Validasi": Exit Sub
        Next i
    End With
    
lanjut0_:
        
        Dim arrayNoOrder(20) As String

        'arrayNoOrder = ArrayString(20)

        With fgData
                a = 0

                For i = 1 To fgData.Rows - 1

                        Dim cekData As Boolean

                        Dim Temp    As String
                        cekData = False
                        Temp = fgData.TextMatrix(i, 39)

                        For j = 0 To UBound(arrayNoOrder)

                                If (Temp = arrayNoOrder(j)) Then
                                        cekData = True
                                End If

                        Next j

                        If (cekData = False) Then '
                                arrayNoOrder(a) = Temp
                                a = a + 1
                        End If

                Next i
                
                Dim X As Integer

                For X = 0 To UBound(arrayNoOrder)
                
                        Dim curHargaBrg As Currency

                        If (fgData.TextMatrix(fgData.Row, 2) <> "") Then
                                If fgData.TextMatrix(fgData.Row, 10) = "0" Or fgData.TextMatrix(fgData.Row, 10) = "" Then MsgBox "Jumlah Tidak Boleh Nol Atau Kosong", vbExclamation, "validasi": Exit Sub
                                If fgData.TextMatrix(1, 0) = "" Then MsgBox "Data barang harus diisi", vbExclamation, "Validasi": Exit Sub
                        End If
                        txtNoResep.Text = ""
                        If (NotUseRacikan = False) Then
                                If sp_GenerateNoResep() = False Then Exit Sub
                                If sp_ResepObat() = False Then Exit Sub
                        End If

                        For i = 1 To .Rows - 1
                                dtpTglPelayanan.Value = Now
            
                                'If .TextMatrix(i, 2) = "" Then GoTo lanjutkan
                                Set dbRst = Nothing
                                strSQL = "SELECT * FROM PemakaianAlkes where NoPendaftaran ='" & mstrNoPen & "' and KdBarang like '%" & .TextMatrix(i, 2) & "%' and Kdasal like '%" & .TextMatrix(i, 12) & "%' and TglPelayanan ='" & Format(dtpTglPelayanan.Value, "yyyy/mm/dd hh:mm:ss") & "'"
                                Call msubRecFO(dbRst, strSQL)

                                If dbRst.EOF = False Then dtpTglPelayanan.Value = DateAdd("s", 1, dtpTglPelayanan.Value)
              
                                If .TextMatrix(i, 2) <> "" Then
                                        If .TextMatrix(i, 2) = "Racikan" Then
                                                strSQLx = "Select * from RacikanObatPasienTemp Where NoRacikan='" & .TextMatrix(i, 34) & "' And KdJenisObat='" & .TextMatrix(i, 25) & "' AND ResepKe='" & Val(.TextMatrix(i, 0)) & "'"
                                                Set rsB = Nothing
                                                Call msubRecFO(rsB, strSQLx)

                                                If rsB.EOF = False Then
                                                        If sp_RacikanObatPasien(.TextMatrix(i, 34), .TextMatrix(i, 25), Val(.TextMatrix(i, 0)), .TextMatrix(i, 35), .TextMatrix(i, 36), .TextMatrix(i, 37), .TextMatrix(i, 32)) = False Then Exit Sub
                                                Else
                                                        GoTo lanjutkan
                                                End If

                                                rsB.MoveFirst

                                                For a = 1 To rsB.RecordCount
                                                        curHargaBrg = 0
                                                        strSQLx = "select dbo.FB_TakeHargaNettoOA('" & mstrKdPenjaminPasien & "','" & rsB("KdJenisObat").Value & "','" & rsB("KdBarang") & "', '" & rsB("KdAsal") & "', '" & rsB("SatuanJml") & "','" & mstrKdRuangan & "', '" & rsB("NoTerima").Value & "')"
                                                        Set rsSplakuk = Nothing
                                                        Call msubRecFO(rsSplakuk, strSQLx)

                                                        If rsSplakuk.EOF Then
                                                                curHargaBrg = 0
                                                        Else
                                                                curHargaBrg = IIf(IsNull(rsSplakuk(0)), 0, rsSplakuk(0))
                                                        End If
                        
                                                        subcurHargaSatuan = 0
                                                        strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "','" & rsB("KdAsal") & "', " & msubKonversiKomaTitik(CStr(curHargaBrg)) & ")  as HargaSatuan"
                                                        Call msubRecFO(rsSplakuk, strSQL)

                                                        If rsSplakuk.EOF = True Then subcurHargaSatuan = 0 Else subcurHargaSatuan = rsSplakuk(0).Value
                                                        subcurHargaSatuan = FormatPembulatan(CDbl(subcurHargaSatuan), mstrKdInstalasiLogin)
                                                        iJmlService = rsB("JmlService").Value
                                                        cTrfService = rsB("TarifService").Value
                       
                                                        '                        If sp_PemakaianObatAlkesResep(rsb("KdBarang"), rsb("KdAsal"), rsb("SatuanJml"), rsb("JmlBarang"), _
                                                                                 subcurHargaSatuan, rsb("KdJenisObat"), iJmlService, cTrfService, _
                                                                                 rsb("ResepKe"), IIf(LCase(.TextMatrix(i, 27)) = "ya", "1", "0"), "", "", "", "", "", "", rsb("NoTerima").Value, dtpTglPelayanan.Value) = False Then Exit Sub
                                                        If (chkDokterPemeriksa.Value = 1) Then
                                                            If sp_PemakaianObatAlkesResep(rsB("KdBarang"), rsB("KdAsal"), rsB("SatuanJml"), rsB("JmlPembulatan"), subcurHargaSatuan, rsB("KdJenisObat"), iJmlService, cTrfService, rsB("ResepKe"), IIf(LCase(.TextMatrix(i, 27)) = "ya", "1", "0"), .TextMatrix(i, 40), "", "", "", dgDokter.Columns("KodeDokter"), "", rsB("NoTerima").Value, dtpTglPelayanan.Value) = False Then Exit Sub
                                                        Else
                                                            If sp_PemakaianObatAlkesResep(rsB("KdBarang"), rsB("KdAsal"), rsB("SatuanJml"), rsB("JmlPembulatan"), subcurHargaSatuan, rsB("KdJenisObat"), iJmlService, cTrfService, rsB("ResepKe"), IIf(LCase(.TextMatrix(i, 27)) = "ya", "1", "0"), .TextMatrix(i, 40), "", "", "", "", "", rsB("NoTerima").Value, dtpTglPelayanan.Value) = False Then Exit Sub
                                                        End If
                                                        
                        
                                                        If .TextMatrix(i, 28) = "1" Then If update_DetailOrderTMOA(dbcmd, rsB("KdBarang").Value, "OA", .TextMatrix(i, 39)) = False Then Exit Sub
                                                        rsB.MoveNext
                                                Next a

                                        Else
                                                
                                                If (arrayNoOrder(X) = .TextMatrix(i, 39)) Then
                                                        If (chkDokterPemeriksa.Value = 1) Then
                                                            If sp_PemakaianObatAlkesResep(.TextMatrix(i, 2), .TextMatrix(i, 12), .TextMatrix(i, 6), CDbl(.TextMatrix(i, 10)), .TextMatrix(i, 16), .TextMatrix(i, 25), .TextMatrix(i, 15), .TextMatrix(i, 14), .TextMatrix(i, 0), IIf(LCase(.TextMatrix(i, 27)) = "ya", "1", "0"), .TextMatrix(i, 40), "", "", "", txtKdDokter.Text, "", .TextMatrix(i, 29), dtpTglPelayanan.Value) = False Then Exit Sub
                                                        Else
                                                            If sp_PemakaianObatAlkesResep(.TextMatrix(i, 2), .TextMatrix(i, 12), .TextMatrix(i, 6), CDbl(.TextMatrix(i, 10)), .TextMatrix(i, 16), .TextMatrix(i, 25), .TextMatrix(i, 15), .TextMatrix(i, 14), .TextMatrix(i, 0), IIf(LCase(.TextMatrix(i, 27)) = "ya", "1", "0"), .TextMatrix(i, 40), "", "", "", "", "", .TextMatrix(i, 29), dtpTglPelayanan.Value) = False Then Exit Sub
                                                        End If
                                                        
           
                                                        If .TextMatrix(i, 28) = "1" Then If update_DetailOrderTMOA(dbcmd, fgData.TextMatrix(i, 2), "OA", .TextMatrix(i, 39)) = False Then Exit Sub
                                                End If
                                        End If
                
                                End If
             
                        Next i
                
                Next X

        End With

        dbConn.Execute "DELETE FROM TempDetailApotikJual WHERE (NoTemporary = '" & txtNoTemporary & "')"
    
lanjutkan:
        MsgBox "Penyimpanan data berhasil", vbInformation, "Informasi"
        Call Add_HistoryLoginActivity("AU_GenerateNoResep+Add_ResepObat+Add_PemakaianObatAlkesResep")
        txtNoResep.Text = ""
        txtTotalBiaya.Text = 0
        txtTotalDiscount.Text = 0
        txtHutangPenjamin.Text = 0
        txtTanggunganRS.Text = 0
        txtHarusDibayar.Text = 0
        statTampil = True
        cmdSimpan.Enabled = False
        Call subSetGrid
        dbConn.Execute "DELETE FROM MasatenggangStok WHERE (IdUser = '" & uniqeId & "')"

        For i = 1 To noRacikan

                Dim racikanId As String

                racikanId = uniqeId

                If (i < 10) Then
                        Mid(racikanId, 1, 1) = i
                        Mid(racikanId, 20, 1) = i
                Else
                        Mid(racikanId, 1, 2) = i
                        Mid(racikanId, 20, 2) = i
                End If

                dbConn.Execute "DELETE FROM MasatenggangStok WHERE (IdUser = '" & racikanId & "')"
        Next i

        Exit Sub

Errload:
        Call msubPesanError
        '    Resume 0
End Sub

Private Sub cmdSimpanTerimaBarang_Click()

        On Error GoTo Errload

        If Val(txtJmlTerima.Text) = 0 Then Exit Sub
        If sp_PenerimaanSementara(Now, fgData.TextMatrix(fgData.Row, 2), fgData.TextMatrix(fgData.Row, 12), Val(txtJmlTerima.Text), "A") = False Then Exit Sub
        Call msubRecFO(rs, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & fgData.TextMatrix(fgData.Row, 2) & "','" & fgData.TextMatrix(fgData.Row, 12) & "') as stok")
        fgData.TextMatrix(fgData.Row, 9) = FormatPembulatan(rs(0), mstrKdInstalasi)
        picPenerimaanSementara.Visible = False
        fgData.SetFocus: fgData.Col = 10

        Exit Sub

Errload:
        Call msubPesanError
End Sub

Private Sub cmdTutup_Click()

        Dim X As Integer

        If cmdSimpan.Enabled = True Then
                If fgData.TextMatrix(1, 3) <> "" Then
                        If MsgBox("Simpan data Pemakaian Obat dan Alat Kesehatan", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
                                Call cmdSimpan_Click

                                Exit Sub

                        End If
                End If
        End If
    
        ' hapus RacikanObatPasienTemp add by denki
        For X = 1 To fgData.Rows - 1

                If fgData.TextMatrix(X, 34) <> "0000000000" Then
                        dbConn.Execute "DELETE FROM RacikanObatPasienTemp WHERE (NoRacikan = '" & fgData.TextMatrix(X, 34) & "')"
                End If

        Next X
    
        '10 03 2014
        ' chandra untuk melakukan perubahan stok ke semula
        Dim i As Integer

        For i = 1 To fgData.Rows - 1

                If fgData.TextMatrix(i, 2) <> "" And Val(fgData.TextMatrix(i, 10)) <> 0 Then
                        If fgData.TextMatrix(i, 6) = "S" Then
                                If sp_StokRealRuangan(fgData.TextMatrix(i, 2), fgData.TextMatrix(i, 12), IIf(fgData.TextMatrix(i, 29) = "", fgData.TextMatrix(i, 27), fgData.TextMatrix(i, 29)), CDbl(fgData.TextMatrix(i, 10)), "C") = False Then Exit Sub
                        End If
                End If

        Next i
        For i = 1 To noRacikan
                Dim racikanId As String
                    racikanId = uniqeId
                    If (i < 10) Then
                        Mid(racikanId, 1, 1) = i
                        Mid(racikanId, 20, 1) = i
                    Else
                        Mid(racikanId, 1, 2) = i
                        Mid(racikanId, 20, 2) = i
                    End If
            Call msubRecFO(rs, "execute RestoreStokBarangOtomatis '" & racikanId & "'")
        Next i
        dbConn.Execute "DELETE FROM TempDetailApotikJual WHERE (NoTemporary = '" & txtNoTemporary & "')"
        Unload Me
        Call frmTransaksiPasien.subPemakaianObatAlkes
End Sub

Private Sub dcAturanPakai_KeyDown(KeyCode As Integer, Shift As Integer)

        If KeyCode = vbKeyEscape Then
                dcAturanPakai.Visible = False
                fgData.SetFocus
        End If

End Sub

Private Sub dcAturanPakai_KeyPress(KeyAscii As Integer)

        If KeyAscii = 13 Then
                dcAturanPakai.Visible = False
                fgData.Col = 31
                fgData.SetFocus
        ElseIf Not (KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ Or KeyAscii = 32 Or KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace Or KeyAscii = Asc(",")) Then
                KeyAscii = 0

                Exit Sub

        End If

End Sub

Private Sub dcAturanPakai_LostFocus()

        If Cancel = False Then
                fgData.TextMatrix(fgData.Row, 30) = dcAturanPakai.Text
                fgData.TextMatrix(fgData.Row, 35) = dcAturanPakai.BoundText
        End If

        Cancel = False
        dcAturanPakai.Visible = False
End Sub

Private Sub dcJenisObat_Change()

        On Error GoTo Errload

        subcurTarifService = 0
        fgData.TextMatrix(fgData.Row, 1) = dcJenisObat.Text
        fgData.TextMatrix(fgData.Row, 25) = dcJenisObat.BoundText
        fgData.TextMatrix(fgData.Row, 14) = subcurTarifService

        Exit Sub

Errload:
        Call msubPesanError
End Sub

Private Sub dcJenisObat_KeyDown(KeyCode As Integer, Shift As Integer)

        If KeyCode = 27 Then dcJenisObat.Visible = False: fgData.SetFocus
End Sub

Private Sub dcJenisObat_KeyPress(KeyAscii As Integer)

        If KeyAscii = 13 Then
                Call dcJenisObat_Change

                If fgData.Row = 1 Then
                        If fgData.TextMatrix(fgData.Row, 0) = "" Then fgData.TextMatrix(fgData.Row, 0) = 1
                Else

                        If fgData.Row - 1 = 0 Then Exit Sub
'                        fgData.TextMatrix(fgData.Row, 0) = fgData.TextMatrix(fgData.Row - 1, 0) + 1
                        If fgData.TextMatrix(fgData.Row, 0) <> fgData.TextMatrix(fgData.Row - 1, 0) Then
                            fgData.TextMatrix(fgData.Row, 0) = fgData.TextMatrix(fgData.Row, 0)
                        Else
                            fgData.TextMatrix(fgData.Row, 0) = fgData.TextMatrix(fgData.Row - 1, 0)
                        End If
                End If

                dcJenisObat.Visible = False
                fgData.Col = 3
                fgData.SetFocus
        ElseIf KeyAscii = 27 Then
                dcJenisObat.Visible = False
        End If

End Sub

Private Sub dcJenisObat_LostFocus()
        dcJenisObat.Visible = False

        Dim i As Integer

        dcJenisObat.Visible = False

        If fgData.TextMatrix(fgData.Row, 0) = "" Or fgData.TextMatrix(fgData.Row, 25) <> fgData.TextMatrix(fgData.Row - 1, 25) Then
            If fgData.Row <> 1 Then fgData.TextMatrix(fgData.Row, 0) = Val(fgData.TextMatrix(fgData.Row, 0)) ' + 1
        End If
        
        If dcJenisObat.BoundText <> "02" And dcJenisObat.BoundText <> "03" Then 'And fgData.TextMatrix(fgData.Row, 2) <> "" Then
                If (NotUseRacikan = True) Then
                    Exit Sub
                End If
                Call subSetGridRacikan
                FraRacikan.Visible = True
                tempNoRacikan = -1
                noRacikan = noRacikan + 1
                txtJumlahObatRacik.SetFocus
                txtJumlahObatRacik.Text = ""
                fgRacikan.TextMatrix(fgRacikan.Row, 2) = fgData.TextMatrix(fgData.Row, 0) ' edit'
                BoolReview = False
        ElseIf dcJenisObat.BoundText <> "02" And dcJenisObat.BoundText <> "03" And fgData.TextMatrix(fgData.Row, 2) <> "" Then
                BoolReview = True
                Call subSetGridRacikan
                FraRacikan.Visible = True

                Set rsB = Nothing
                strSQL = "SELECT RacikanObatPasienTemp.NoRacikan, RacikanObatPasienTemp.KdBarang, RacikanObatPasienTemp.KdAsal, RacikanObatPasienTemp.KdRuangan, " & _
                   "RacikanObatPasienTemp.SatuanJml, RacikanObatPasienTemp.KdJenisObat, RacikanObatPasienTemp.ResepKe, RacikanObatPasienTemp.NoTerima, " & _
                   "RacikanObatPasienTemp.JmlBarang, RacikanObatPasienTemp.JmlPembulatan, RacikanObatPasienTemp.QtyRacikan, RacikanObatPasienTemp.HargaSatuan, " & _
                   "RacikanObatPasienTemp.NoCM, RacikanObatPasienTemp.NoPendaftaran, RacikanObatPasienTemp.KebutuhanML, RacikanObatPasienTemp.KebutuhanTB, " & _
                   "RacikanObatPasienTemp.JmlService, RacikanObatPasienTemp.TarifService, RacikanObatPasienTemp.StatusPelayanan, JenisObat.JenisObat, " & _
                   "MasterBarang.NamaBarang , MasterBarang.KeKuatan, AsalBarang.NamaAsal " & _
                   "FROM RacikanObatPasienTemp INNER JOIN " & _
                   "MasterBarang ON RacikanObatPasienTemp.KdBarang = MasterBarang.KdBarang INNER JOIN " & _
                   "JenisObat ON RacikanObatPasienTemp.KdJenisObat = JenisObat.KdJenisObat INNER JOIN " & _
                   "AsalBarang ON RacikanObatPasienTemp.KdAsal = AsalBarang.KdAsal " & _
                   "where nopendaftaran = '" & mstrNoPen & "' and NoRacikan = '" & fgData.TextMatrix(fgData.Row, 34) & "' and ResepKe = '" & fgData.TextMatrix(fgData.Row, 0) & "'"

                Call msubRecFO(rsB, strSQL)

                If rsB.EOF = False Then
                        txtJumlahObatRacik.Text = IIf(IsNull(rsB("qtyRacikan").Value), "", rsB("qtyRacikan").Value)
                        rsB.MoveFirst

                        With fgRacikan

                                For i = 1 To rsB.RecordCount
                                        .TextMatrix(i, 0) = IIf(IsNull(rsB("kdBarang").Value), "", rsB("kdBarang").Value) 'KdBarang
                                        .TextMatrix(i, 1) = IIf(IsNull(rsB("qtyRacikan").Value), "", rsB("qtyRacikan").Value) 'Jenis obat
                                        .TextMatrix(i, 2) = IIf(IsNull(rsB("ResepKe").Value), "", rsB("ResepKe").Value) '"R/Ke"
                                        .TextMatrix(i, 3) = IIf(IsNull(rsB("NamaBarang").Value), "", rsB("NamaBarang").Value) '"Nama Barang"
                                        .TextMatrix(i, 4) = msubKonversiKomaTitik(IIf(IsNull(rsB("KebutuhanML").Value), 0, rsB("KebutuhanML").Value)) '"/Mg /Ml"
                                        .TextMatrix(i, 5) = msubKonversiKomaTitik(IIf(IsNull(rsB("KebutuhanTB").Value), 0, rsB("KebutuhanTB").Value)) '"/Tablet"
                                        .TextMatrix(i, 6) = msubKonversiKomaTitik(IIf(IsNull(rsB("JmlBarang").Value), 0, rsB("JmlBarang").Value)) '"Jumlah"
                                        .TextMatrix(i, 7) = msubKonversiKomaTitik(IIf(IsNull(rsB("JmlPembulatan").Value), 0, rsB("JmlPembulatan").Value)) '"Jumlah Pembulatan(untuk harga)"
                                        .TextMatrix(i, 8) = msubKonversiKomaTitik(IIf(IsNull(rsB("HargaSatuan").Value), 0, rsB("HargaSatuan").Value)) '"Harga Satuan"
                                        .TextMatrix(i, 9) = FormatPembulatan(CDbl(msubKonversiKomaTitik((IIf(IsNull(rsB("HargaSatuan").Value), 0, rsB("HargaSatuan").Value) * IIf(IsNull(rsB("JmlPembulatan").Value), 0, rsB("JmlPembulatan").Value)) + (IIf(IsNull(rsB("JmlService").Value), 0, rsB("JmlService").Value) * IIf(IsNull(rsB("TarifService").Value), 0, rsB("TarifService").Value)))), mstrKdInstalasiLogin) '"Total Harga"
                                        .TextMatrix(i, 10) = IIf(IsNull(rsB("Kekuatan").Value), "", rsB("Kekuatan").Value) '"Kekuatan"
                                        .TextMatrix(i, 11) = IIf(IsNull(rsB("NamaAsal").Value), "", rsB("NamaAsal").Value) '"AsalBarang"
                                        .TextMatrix(i, 12) = IIf(IsNull(rsB("kdAsal").Value), "", rsB("KdAsal").Value) '"kdAsal"
                                        .TextMatrix(i, 13) = IIf(IsNull(rsB("SatuanJml").Value), "", rsB("SatuanJml").Value) '"satuan"
                                        .TextMatrix(i, 14) = IIf(IsNull(rsB("NoTERIMA").Value), "0000000000", rsB("NoTERIMA").Value) '"NoFIFO"
                                        .TextMatrix(i, 15) = IIf(IsNull(rsB("JmlService").Value), 0, rsB("JmlService").Value) '"jmlService" 'add Column Jumlah Service
                                        .TextMatrix(i, 16) = IIf(IsNull(rsB("TarifService").Value), 0, rsB("TarifService").Value) '"TarifService"
                                        .Rows = .Rows + 1
                                        rsB.MoveNext
                                Next i

                        End With

                End If

        Else
                BoolReview = False
                Call dcJenisObat_Change
        End If

End Sub

Private Sub dcKeteranganPakai_KeyDown(KeyCode As Integer, Shift As Integer)

        If KeyCode = vbKeyEscape Then
                dcKeteranganPakai.Visible = False
                fgData.SetFocus
        End If

End Sub

Private Sub dcKeteranganPakai_KeyPress(KeyAscii As Integer)

        If KeyAscii = 13 Then
                dcKeteranganPakai.Visible = False
                fgData.Col = 32
                fgData.SetFocus
        ElseIf Not (KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ Or KeyAscii = 32 Or KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace Or KeyAscii = Asc(",")) Then
                KeyAscii = 0

                Exit Sub

        End If

End Sub

Private Sub dcKeteranganPakai_LostFocus()

        If Cancel = False Then
                fgData.TextMatrix(fgData.Row, 31) = dcKeteranganPakai.Text
                fgData.TextMatrix(fgData.Row, 36) = dcKeteranganPakai.BoundText
        End If

        Cancel = False
End Sub

Private Sub dcNamaPelayananRS_Click(Area As Integer)
    On Error GoTo Errload
    Dim i As Integer

    If bolStatusFIFO = False Then
        fgData.TextMatrix(fgData.Row, 41) = dcNamaPelayananRS.Text
        fgData.TextMatrix(fgData.Row, 40) = dcNamaPelayananRS.BoundText
    Else
        With fgData
            For i = 1 To .Rows - 1
                If .TextMatrix(.Row, 2) = .TextMatrix(i, 2) And .TextMatrix(.Row, 12) = .TextMatrix(i, 12) And .TextMatrix(.Row, 6) = .TextMatrix(i, 6) Then
                If (dcNamaPelayananRS.Text <> "") Then
                    .TextMatrix(i, 41) = dcNamaPelayananRS.Text
                    .TextMatrix(i, 40) = dcNamaPelayananRS.BoundText
                    dcNamaPelayananRS.Visible = False
                    .Col = 41
                End If
                    
                End If
            Next i
        End With
    End If

    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub dgDokter_DblClick()

        On Error GoTo Errload

        If dgDokter.ApproxCount = 0 Then Exit Sub
        txtDokter.Text = dgDokter.Columns("Nama Dokter")
        dgDokter.Visible = False
        txtKdDokter.Text = dgDokter.Columns("KodeDokter")
        fgData.SetFocus

        Exit Sub

Errload:
        Call msubPesanError
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)

        If KeyAscii = 13 Then Call dgDokter_DblClick
End Sub

Private Sub dgObatAlkes_DblClick()

         On Error GoTo Errload
    
        Dim i                        As Integer

        Dim tempSettingDataPendukung As Integer

        Dim curHargaBrg              As Currency

        Dim strNoTerima              As String

        'If fgData.TextMatrix(fgData.Row, 2) <> dgObatAlkes.Columns("KdBarang") And fgData.TextMatrix(fgData.Row, 10) <> "" And useGeneric = False Then
        '        MsgBox "Data tidak bisa diubah dengan data yang lain", vbExclamation, "Validasi"
        '        fgData.SetFocus
        '        dgObatAlkes.Visible = False

        '        Exit Sub

        'End If
         Call msubRecFO(rs, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & dgObatAlkes.Columns("KdBarang") & "','" & dgObatAlkes.Columns("KdAsal") & "') as stok")
         If (rs.EOF = False) Then
            If (rs(0).Value = 0) Then
                MsgBox "Stok " & dgObatAlkes.Columns(1) & " kosong"
                dgObatAlkes.SetFocus
                Exit Sub
            End If
         End If
        curHutangPenjamin = 0
        curTanggunganRS = 0

        strNoTerima = ""

        Set rsB = Nothing
        Call msubRecFO(rsB, "select dbo.TakeNoFIFO_F('" & dgObatAlkes.Columns("KdBarang") & "','" & dgObatAlkes.Columns("KdAsal") & "','" & mstrKdRuangan & "') as NoFIFO")
        strNoTerima = IIf(IsNull(rsB("NoFIFO")), "0000000000", rsB("NoFIFO"))

        For i = 0 To fgData.Rows - 1

                If dgObatAlkes.Columns("KdBarang") = fgData.TextMatrix(i, 2) And dgObatAlkes.Columns("KdAsal") = fgData.TextMatrix(i, 12) And useGeneric = False Then
            
                        MsgBox "Data tersebut sudah diinput", vbExclamation, "Validasi"
                        dgObatAlkes.Visible = False
                        fgData.SetFocus: fgData.Row = i

                        Exit Sub

                End If

        Next i

        With fgData
                .TextMatrix(.Row, 2) = dgObatAlkes.Columns("KdBarang")
        'caun
                .TextMatrix(.Row, 3) = dgObatAlkes.Columns("Nama Barang")
                tmpKdBar = dgObatAlkes.Columns("KdGenerikBarang")
                .TextMatrix(.Row, 4) = dgObatAlkes.Columns("Kekuatan")
                .TextMatrix(.Row, 5) = dgObatAlkes.Columns("AsalBarang")
                .TextMatrix(.Row, 6) = dgObatAlkes.Columns("SatuanJml")
                .TextMatrix(.Row, 10) = 0
                '.TextMatrix(.Row, 10) = dgObatAlkes.Columns("Discount")
                .TextMatrix(.Row, 29) = strNoTerima
                .TextMatrix(.Row, 39) = "0"
                curHargaBrg = 0
                dgObatAlkes.Visible = False
                strSQL = ""
                Set rsB = Nothing
                ' chandra 10 03 2014
                ' dirubah untuk mengambil barang berdasarkan fifo di detail terima bukan berdasarkan no terima
                'strSQL = "SELECT dbo.FB_TakeHargaNettoOA('" & mstrKdPenjaminPasien & "','" & mstrKdJenisPasien & "','" & dgObatAlkes.Columns("KdBarang") & "','" & dgObatAlkes.Columns("KdAsal") & "','" & dgObatAlkes.Columns("SatuanJml") & "', '" & mstrKdRuangan & "','" & .TextMatrix(.Row, 29) & "') AS HargaBarang"
                strSQL = "SELECT dbo.FB_TakeHargaNettoObatAlkesFifo('" & mstrKdPenjaminPasien & "','" & mstrKdJenisPasien & "','" & dgObatAlkes.Columns("KdBarang") & "','" & dgObatAlkes.Columns("KdAsal") & "','" & dgObatAlkes.Columns("SatuanJml") & "', '" & mstrKdRuangan & "','" & .TextMatrix(.Row, 29) & "') AS HargaBarang"
                Call msubRecFO(rsB, strSQL)

                If rsB.EOF = True Then curHargaBrg = 0 Else curHargaBrg = rsB(0).Value
                subcurHargaSatuan = curHargaBrg
                strSQL = ""
                Set rs = Nothing
                'subcurHargaSatuan = 0
                'chandra
                ' selain penjualan bebas tidak ada penaikan harga
                strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "','" & dgObatAlkes.Columns("KdAsal") & "', " & msubKonversiKomaTitik(CStr(curHargaBrg)) & ")  as HargaSatuan"
                Call msubRecFO(rs, strSQL)

                'khusus OA harga tidak dikalikan lg Ppn krn OA termasuk pelayanan yg include ke tindakan (TM)
                If rs.EOF = True Then subcurHargaSatuan = 0 Else subcurHargaSatuan = rs(0).Value
                strSQL = "SELECT Discount FROM dbo.HargaNettoBarangFIFO WHERE NoTerima='" & strNoTerima & "' and KdBarang='" & dgObatAlkes.Columns("KdBarang") & "'"
                Call msubRecFO(rsB, strSQL)
                .TextMatrix(.Row, 8) = 0 'FormatPembulatan((IIf(rsb.EOF = True, 0, rsb(0).Value) / 100) * CDbl(subcurHargaSatuan), mstrKdInstalasiLogin)
                .TextMatrix(.Row, 7) = subcurHargaSatuan
                .TextMatrix(.Row, 7) = FormatPembulatan(CDbl(subcurHargaSatuan), mstrKdInstalasiLogin)   'FormatPembulatan(CDbl(subcurHargaSatuan), mstrKdInstalasiLogin)

                If .TextMatrix(.Row, 7) = "" Then .TextMatrix(.Row, 7) = "0"
                '.TextMatrix(.Row, 8) = 0 '(dgObatAlkes.Columns("Discount").Value / 100) * subcurHargaSatuan

                '        strSQL = ""
                '        Set rs = Nothing
                '        strSQL = "Select JmlStok as Stok From StokRuangan Where KdBarang='" & dgObatAlkes.Columns("KdBarang") & "' and KdAsal='" & dgObatAlkes.Columns("KdAsal") & "' and KdRuangan='" & mstrKdRuangan & "'"
                '        Call msubRecFO(rs, strSQL)
                '        If rs.EOF Then
                '            .TextMatrix(.Row, 9) = 0
                '        Else
                '            .TextMatrix(.Row, 9) = IIf(IsNull(rs("Stok")), 0, rs("Stok"))
                '        End If
                Call msubRecFO(rs, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & dgObatAlkes.Columns("KdBarang") & "','" & dgObatAlkes.Columns("KdAsal") & "') as stok")
                .TextMatrix(.Row, 9) = IIf(IsNull(rs("Stok")), 0, rs("Stok"))

                .TextMatrix(.Row, 12) = dgObatAlkes.Columns("KdAsal")
                .TextMatrix(.Row, 13) = dgObatAlkes.Columns("Jenis Barang")
                .TextMatrix(.Row, 14) = subcurTarifService
                .TextMatrix(.Row, 15) = subintJmlService
                .TextMatrix(.Row, 16) = subcurHargaSatuan
                .TextMatrix(.Row, 17) = curHutangPenjamin
                .TextMatrix(.Row, 18) = curTanggunganRS
                .TextMatrix(.Row, 19) = 0
                .TextMatrix(.Row, 20) = 0
                .TextMatrix(.Row, 21) = 0

                .TextMatrix(.Row, 23) = txtNoTemporary.Text
                txtHargaBeli.Text = curHargaBrg
                .TextMatrix(.Row, 24) = CDbl(txtHargaBeli.Text)

                tempSettingDataPendukung = 0
                Dim barisValidForCountingResep As Integer
                CountingBiayaAdministrasi

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
        useGeneric = False
        Exit Sub

Errload:
End Sub

Private Sub dgObatAlkes_KeyPress(KeyAscii As Integer)

        If KeyAscii = 13 Then Call dgObatAlkes_DblClick
End Sub

Private Sub dgObatAlkesRacikan_DblClick()

        On Error GoTo Errload

        If dgObatAlkesRacikan.ApproxCount = 0 Then Exit Sub

        Dim i                        As Integer

        '----
        'Dim i As Integer
        Dim tempSettingDataPendukung As Integer

        Dim curHargaBrg              As Currency

        Dim Resep                    As String

        Call msubRecFO(rs, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & dgObatAlkesRacikan.Columns("KdBarang") & "','" & dgObatAlkesRacikan.Columns("KdAsal") & "') as stok")

        If (rs.EOF = False) Then
                If (rs(0).Value = 0) Then
                        MsgBox "Stok " & dgObatAlkesRacikan.Columns(2) & " kosong"
                        dgObatAlkesRacikan.SetFocus

                        Exit Sub

                End If
        End If

        strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
        Call msubRecFO(rs, strSQL)

        If rs.EOF = False Then
                mstrKdJenisPasien = rs("KdKelompokPasien").Value
                mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
        End If
    
        curHutangPenjamin = 0
        curTanggunganRS = 0
    
        For i = 0 To fgRacikan.Rows - 1

                If dgObatAlkesRacikan.Columns("KdBarang") = fgRacikan.TextMatrix(i, 0) And dgObatAlkesRacikan.Columns("KdAsal") = fgRacikan.TextMatrix(i, 12) Then
                        MsgBox "Data tersebut sudah diinput", vbExclamation, "Validasi"
dgObatAlkesRacikan.SetFocus
                        'dgObatAlkesRacikan.Visible = False
                        'fgRacikan.SetFocus: fgRacikan.Row = i
                        Exit Sub

                End If

        Next i

        With fgRacikan
                .TextMatrix(.Row, 0) = dgObatAlkesRacikan.Columns("KdBarang")
                .TextMatrix(.Row, 3) = dgObatAlkesRacikan.Columns("NamaBarang")
                .TextMatrix(.Row, 11) = dgObatAlkesRacikan.Columns("AsalBarang")
                .TextMatrix(.Row, 12) = dgObatAlkesRacikan.Columns("kdAsal")
                .TextMatrix(.Row, 13) = dgObatAlkesRacikan.Columns("satuan")
        
                Set rsB = Nothing
                Call msubRecFO(rsB, "select dbo.TakeNoFIFO_F('" & dgObatAlkesRacikan.Columns("KdBarang") & "','" & dgObatAlkesRacikan.Columns("KdAsal") & "','" & mstrKdRuangan & "') as NoFIFO")
                .TextMatrix(.Row, 14) = IIf(IsNull(rsB("NoFIFO")), "0000000000", rsB("NoFIFO"))
        
                .ColAlignment(3) = flexAlignCenterCenter
                .TextMatrix(.Row, 4) = ""
                .TextMatrix(.Row, 5) = ""
                .TextMatrix(.Row, 6) = ""
                .TextMatrix(.Row, 2) = .TextMatrix(1, 2)
                curHargaBrg = 0
                strSQL = ""
                Set rsB = Nothing
                strSQL = "SELECT dbo.FB_TakeHargaNettoOA('" & mstrKdPenjaminPasien & "','" & mstrKdJenisPasien & "','" & dgObatAlkesRacikan.Columns("KdBarang") & "','" & dgObatAlkesRacikan.Columns("KdAsal") & "','" & dgObatAlkesRacikan.Columns("Satuan") & "', '" & mstrKdRuangan & "','" & .TextMatrix(.Row, 14) & "') AS HargaBarang"
                Call msubRecFO(rsB, strSQL)

                If rsB.EOF = True Then curHargaBrg = 0 Else curHargaBrg = rsB(0).Value
                strSQL = ""
                Set rs = Nothing
                subcurHargaSatuan = 0
                strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "', '" & mstrKdPenjaminPasien & "', '" & dgObatAlkesRacikan.Columns("KdAsal") & "', " & msubKonversiKomaTitik(CStr(curHargaBrg)) & " )  as HargaSatuan "
                Call msubRecFO(rs, strSQL)

                If rs.EOF = True Then subcurHargaSatuan = 0 Else subcurHargaSatuan = rs(0).Value
                .TextMatrix(.Row, 8) = FormatPembulatan(CDbl(subcurHargaSatuan), mstrKdInstalasiLogin)
                '.TextMatrix(.Row, 8) = Format(subcurHargaSatuan, "#,###.00")
                '        Call msubRecFO(rs, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & dgObatAlkesRacikan.Columns("KdBarang") & "','" & dgObatAlkesRacikan.Columns("KdAsal") & "') as stok")
                '        .TextMatrix(.Row, 9) = IIf(IsNull(rs("Stok")), 0, rs("Stok"))
                .TextMatrix(.Row, 10) = dgObatAlkesRacikan.Columns("Kekuatan")
        End With
    
        dgObatAlkesRacikan.Visible = False
        TxtIsiRacikan.Visible = False

        With fgRacikan
                .Rows = .Rows + 1
                .SetFocus
                .Col = 4
        End With
    
        Exit Sub

Errload:
        Call msubPesanError
        '    Resume 0
End Sub

Private Sub dgObatAlkesRacikan_KeyPress(KeyAscii As Integer)

        If KeyAscii = 13 Then Call dgObatAlkesRacikan_DblClick
        '    With fgRacikan
        '        .SetFocus
        '        .Col = 4
        '    End With
End Sub

Private Sub dtpTglPelayanan_Change()
        dtpTglPelayanan.MaxDate = Now
End Sub

Private Sub dtpTglPelayanan_KeyDown(KeyCode As Integer, Shift As Integer)

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

        Dim strKdBrg  As String

        Dim strKdAsal As String

        Dim i         As Integer

        Select Case KeyCode

                Case 13

                        If fgData.Col = fgData.Cols - 1 Then
                                If fgData.TextMatrix(fgData.Row, 2) <> "" Then
                                        If fgData.TextMatrix(fgData.Rows - 1, 2) <> "" Then
                                                fgData.Rows = fgData.Rows + 1

                                                If fgData.TextMatrix(fgData.Rows - 2, 25) = "" Then
                                                        fgData.TextMatrix(fgData.Rows - 1, 0) = "1"
                                                ElseIf fgData.TextMatrix(fgData.Rows - 2, 25) = "01" Then
                                                        fgData.TextMatrix(fgData.Rows - 1, 0) = "0"
                                                Else
                                                        fgData.TextMatrix(fgData.Rows - 1, 0) = Val(fgData.TextMatrix(fgData.Rows - 2, 0))
                                                End If
                                        End If

                                        fgData.Row = fgData.Rows - 1
                                        fgData.Col = 1
                                        dcJenisObat.Visible = True
                                        
                                        dcJenisObat.SetFocus
                                Else
                                        fgData.Col = 1
                                End If

                        Else

                                For i = 0 To fgData.Cols - 2

                                        If fgData.Col = fgData.Cols - 1 Then Exit For
                                        fgData.Col = fgData.Col + 1

                                        If fgData.ColWidth(fgData.Col) > 0 Then Exit For
                                Next i

                        End If

                        'fgData.SetFocus

                        If fgData.Col = 1 Then Call subLoadDataCombo(dcJenisObat)

                Case 27
                        dgObatAlkes.Visible = False

                Case vbKeyDelete

                        'validasi FIFO
                        If bolStatusFIFO = True Then
                                If fgData.CellBackColor = vbRed Then
                                        MsgBox "Data yang barisnya berwarna merah tidak bisa di edit", vbExclamation, "validasi"
                                        fgData.SetFocus

                                        Exit Sub

                                End If

                                '                If fgData.TextMatrix(fgData.Row, 28) <> "1" Then
                
                                With fgData
                                        i = .Rows - 1
                                        strKdBrg = .TextMatrix(.Row, 2)
                                        strKdAsal = .TextMatrix(.Row, 12)

                                        Do While i <> 0 'khusus utk delete dr keyboard diset 0 agar ke cek keseluruhannya

                                             '   If .TextMatrix(i, 2) <> "" Then
                                                        If (strKdBrg = .TextMatrix(i, 2)) And (strKdAsal = .TextMatrix(i, 12)) Then
                                                                .Row = i
                                                                Call subHapusDataGrid
                                                                .Row = i - 1
                                                        End If
                                              '  End If

                                                i = i - 1
                                        Loop

                                End With

                                '                Else
                                '                Call subHapusDataGrid
                                '                End If
                        Else
                                Call subHapusDataGrid
                        End If

        End Select

End Sub

Private Sub fgData_KeyPress(KeyAscii As Integer)

        On Error GoTo Errload

        'Validasi jika FIFO
        If bolStatusFIFO = True Then
                If fgData.CellBackColor = vbRed Then
                        MsgBox "Data yang barisnya berwarna merah tidak bisa di edit", vbExclamation, "validasi"
                        fgData.SetFocus

                        Exit Sub

                End If
        End If

        'end fifo

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
                        fgData.Col = 1
                        Call subLoadDataCombo(dcJenisObat)
                        dcJenisObat.SetFocus
                Case 2 'Kode Barang
                        txtIsi.MaxLength = 9
                        Call subLoadText
                        txtIsi.Text = Chr(KeyAscii)
                        txtIsi.SelStart = Len(txtIsi.Text)

                Case 3 'Nama Barang
                        txtIsi.MaxLength = 20
                        Call subLoadText
                        txtIsi.Text = Chr(KeyAscii)
                        txtIsi.SelStart = Len(txtIsi.Text)

                Case 10 'Jumlah
                        txtIsi.MaxLength = 7

                        If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyEscape Or KeyAscii = Asc(",")) Then Exit Sub
                        Call subLoadText
                        txtIsi.Text = Chr(KeyAscii)
                        txtIsi.SelStart = Len(txtIsi.Text)

                        '        Case 27 'Status Stok
                        '            Call subLoadCheck
            
                Case 30 'aturan Pakai
                        fgData.Col = 30
                        Call subLoadDataCombo(dcAturanPakai)
            
                Case 31 'Keterangan Pakai
                        fgData.Col = 31
                        Call subLoadDataCombo(dcKeteranganPakai)
            
                Case 32 'Keterangan Lainnya
                        txtIsi.MaxLength = 100
                        Call subLoadText
                        txtIsi.Text = Chr(KeyAscii)
                        txtIsi.SelStart = Len(txtIsi.Text)
                 Case 41 'nama pelayanan rs yang di gunakan ' ganti pemakain bahan
            fgData.Col = 41
                Call subLoadDataCombo(dcNamaPelayananRS)
                dcNamaPelayananRS.Visible = True
                dcNamaPelayananRS.SetFocus
        End Select

        Exit Sub

Errload:
        Call msubPesanError
End Sub

Private Sub fgRacikan_KeyDown(KeyCode As Integer, Shift As Integer)

        Dim X, i As Integer

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
 If fgRacikan.CellBackColor = vbRed Then
                                        MsgBox "Data yang barisnya berwarna merah tidak bisa di edit", vbExclamation, "validasi"
                                        fgData.SetFocus

                                        Exit Sub

                                End If

                        '            If fgRacikan.Row = 1 Then
                        '                For i = 0 To fgRacikan.Cols - 1
                        '                    fgRacikan.TextMatrix(1, i) = ""
                        '                Next i
                        '            Else
                        ''                With fgRacikan
                        ''                dbConn.Execute "DELETE FROM TempDetailApotikJual " & _
                        ''                    " WHERE (NoTemporary = '" & Trim(txtNoTemporary.Text) & "') AND (KdBarang = '" & .TextMatrix(.Row, 2) & "') AND (KdAsal = '" & .TextMatrix(.Row, 12) & "')"
                        ''                End With
                        '                fgRacikan.RemoveItem fgRacikan.Row
                        '            End If
                        If fgRacikan.Rows - 1 = 1 Then
                                
                                For i = 3 To fgRacikan.Cols - 1
                                        fgRacikan.TextMatrix(1, i) = ""
                                Next i

                                If fgRacikan.TextMatrix(fgRacikan.Row, 0) = "" Then FraRacikan.Visible = False
                                fgData.SetFocus
                                fgData.Col = 1
                                Call subLoadDataCombo(dcJenisObat)
                        Else
                              With fgRacikan
                                        i = .Rows - 1
                                        Dim strKdBrg As String
                                        Dim strKdAsal As String
                                        strKdBrg = .TextMatrix(.Row, 0)
                                        strKdAsal = .TextMatrix(.Row, 12)

                                        Do While i <> 0 'khusus utk delete dr keyboard diset 0 agar ke cek keseluruhannya

                                                If .TextMatrix(i, 0) <> "" Then
                                                        If (strKdBrg = .TextMatrix(i, 0)) And (strKdAsal = .TextMatrix(i, 12)) Then
                                                                .Row = i
                                                                Call subHapusDataGridRacikan
                                                                .Row = i - 1
                                                        End If
                                                End If

                                                i = i - 1
                                        Loop

                                End With
                                '                If fgRacikan.TextMatrix(fgRacikan.Row, 0) = "" Then FraRacikan.Visible = False
                        End If

        End Select

End Sub

Private Sub fgRacikan_KeyPress(KeyAscii As Integer)

        On Error GoTo Errload

        TxtIsiRacikan.Text = ""

        If Not (KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ Or KeyAscii = 32 Or KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace Or KeyAscii = Asc(".")) Then
                KeyAscii = 0

                Exit Sub

        End If
        If fgRacikan.CellBackColor = vbRed Then
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
                        Exit Sub
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

Errload:
        Call msubPesanError
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

'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case vbKeyF5
'            frmDaftarBarangGratisRuangan.Show
'    End Select
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

        If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()

        On Error GoTo Errload

        '10 03 2014 tambah chandra
        uniqeId = CreateGUID
        noRacikan = 0

        Dim curHargaBrg As Currency

        Dim i, j As Integer

        Dim dblJmlTerkecil           As Double

        Dim dblTotalStokK            As Double

        Dim tempSettingDataPendukung As Integer

        Dim curHarusDibayar          As Currency

        Dim strNoTerima              As String

        Dim strKdBrg, strKdAsal As String

        Dim dblSelisih As Double

        Dim rsC        As ADODB.recordset

        Dim bolCekFIFO As Boolean

        Dim k, intRowTemp, iTemp As Integer

        Call PlayFlashMovie(Me)
        Call centerForm(Me, MDIUtama)
        dtpTglPelayanan.Value = Now
        dtpTglResep.Value = Now

        Call subSetGrid
        Call subLoadDcSource
        dgDokter.Visible = False
        dcJenisObat.BoundText = ""
        txtRP.Enabled = False

        dgObatAlkes.Top = 2880
        dgObatAlkes.Left = 2040
        dgObatAlkes.Visible = True
        dgObatAlkes.Visible = False
        chkNoResep.Enabled = Not NotUseRacikan
        txtNoResep.Enabled = Not NotUseRacikan

        If (NotUseRacikan) Then
            
        End If

        txtDokter.Enabled = False
        Call msubRecFO(rs, "execute RestoreStokBarangOtomatis NULL")
        'semua pasien yang masuk UDD bayar tunai. kd penjamin dan kd jenispasien dikembalikan saat form unload

        strSQL = "SELECT JenisHargaNetto" & " From PersentaseUpTarifOA" & " Where(IdPenjamin = '" & mstrKdPenjaminPasien & "') And (KdKelompokPasien = '" & mstrKdJenisPasien & "')"
        Call msubRecFO(rs, strSQL)
        
        If rs.EOF = True Then    'Jika di Program Sys Adm PersentaseUpTarifOA nya belum disetting buat tanggungan penjaminnya maka di set 1
                subJenisHargaNetto = 1
        Else
                subJenisHargaNetto = IIf(rs.EOF = False, 1, rs(0))
        End If
    
        'Untuk load jika ada order obat yang dari ruangan lain----------------------------------------------------------------------------
        '    strSQL = "SELECT * FROM V_DetailOrderOAx where NoPendaftaran ='" & mstrNoPen & "' and KdRuanganTujuan ='" & mstrKdRuangan & "'"
    
        '    If txtFormPengirimText.Text = "frmTransaksiPasien" And statTampil = False Then
        
        '        statTampil = True
        If (NotUseRacikan = False) Then
        
                strSQL = "SELECT * FROM V_DetailOrderOA_Baru where NoPendaftaran ='" & mstrNoPen & "' and KdRuanganTujuan ='" & mstrKdRuangan & "' and NoRiwayat is null order by NoOrder, ResepKe asc"
        
                Set dbRst = Nothing
                Call msubRecFO(rsK, strSQL)
    
                If rsK.EOF = False Then
                        If IsNull(rsK("IdDokterOrder")) Then
                                chkDokterPemeriksa.Value = Unchecked
                        Else
                                chkDokterPemeriksa.Value = Checked
                                txtDokter.Text = rsK("DokterOrder")
                                txtKdDokter.Text = rsK("IdDokterOrder")
                                dgDokter.Visible = False
                        End If

                        Dim jumlahNoTerimaLainnya As Integer

                        jumlahNoTerimaLainnya = 0

                        With fgData

                                For k = 1 To rsK.RecordCount
                                        curHutangPenjamin = 0
                                        curTanggunganRS = 0
                                        .Row = k + jumlahNoTerimaLainnya

                                        If bolStatusFIFO = True Then
                                                iTemp = k

                                                If bolCekFIFO = True Then k = k + 1
                                        End If

                                        .TextMatrix(.Row, 0) = rsK("ResepKe")
                                        .TextMatrix(.Row, 1) = ""
                                        .TextMatrix(.Row, 25) = ""
                                        .TextMatrix(.Row, 2) = rsK("KdBarang")
                                        .TextMatrix(.Row, 3) = rsK("NamaBarang")
                                        .TextMatrix(.Row, 4) = IIf(IsNull(rsK("Kekuatan")), "", rsK("Kekuatan"))
                                        .TextMatrix(.Row, 5) = rsK("NamaAsal")
                                        .TextMatrix(.Row, 6) = rsK("SatuanJml")
                                        .TextMatrix(.Row, 12) = rsK("KdAsal")
                                        .TextMatrix(.Row, 30) = IIf(IsNull(rsK("SatuanEtiket")), "", rsK("SatuanEtiket"))
                                        .TextMatrix(.Row, 31) = IIf(IsNull(rsK("WaktuEtiket")), "", rsK("waktuEtiket"))
                                        .TextMatrix(.Row, 23) = i
                                        .TextMatrix(.Row, 25) = rsK("KdJenisObat")
                                        .TextMatrix(.Row, 39) = rsK("NoOrder")
        
                                        Set rsB = Nothing
                                        Call msubRecFO(rsB, "select dbo.TakeNoFIFO_F('" & .TextMatrix(.Row, 2) & "','" & .TextMatrix(.Row, 12) & "','" & mstrKdRuangan & "') as NoFIFO")
                                        strNoTerima = IIf(IsNull(rsB("NoFIFO")), "0000000000", rsB("NoFIFO"))
                                        .TextMatrix(.Row, 29) = strNoTerima
                                        .TextMatrix(.Row, 28) = 1
                                        .TextMatrix(.Row, 10) = rsK("JmlBarang")
                                        .TextMatrix(.Row, 1) = rsK("JenisObat")
                                        txtIsi.Text = rsK("JmlBarang")

                                        If (fgData.TextMatrix(.Row, 6) = "S") Then
                    
                                                If bolStatusFIFO = False Then
                                                        If CDbl(txtIsi.Text) > CDbl(.TextMatrix(.Row, 9)) Then
                                                                MsgBox "Jumlah lebih besar dari stock (" & .TextMatrix(.Row, 9) & ")", vbExclamation, "Validasi"
                                                                txtIsi.SelStart = 0: txtIsi.SelLength = Len(txtIsi.Text)

                                                                Exit Sub

                                                        End If
                                                End If

                                        ElseIf (fgData.TextMatrix(.Row, 6) = "K") Then
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
                    
                                        If fgData.TextMatrix(.Row, 1) = "" Then
                                                MsgBox "Lengkapi jenis obat", vbExclamation, "Validasi": Exit Sub
                                        End If

                                        'add for FIFO validasi jika terjadi edit jml sto.Row, hapus otomatis
                                        If bolStatusFIFO = True Then
                                                If Trim(.TextMatrix(.Row, 10)) <> "" Then
                                                        i = .Rows - 1
                                                        strKdBrg = .TextMatrix(.Row, 2)
                                                        strKdAsal = .TextMatrix(.Row, 12)

                                                        'For i = 1 To .Rows - 1

                                                        '       If (strKdBrg = .TextMatrix(i, 2)) And (strKdAsal = .TextMatrix(i, 12)) Then
                                                        '              .Row = i

                                                        '             Exit For

                                                        '    End If

                                                        'Next i

                                                End If
                                        
                                                intRowTemp = 0
                                        End If

                                        'add for FIFO jika jml yg diinput melebihi stok penerimaan, mk otomatis muncul di row selanjutnya
                                        If bolStatusFIFO = True Then
                                                Set dbRst = Nothing
                                                Call msubRecFO(dbRst, "select JmlStok as stok from stokruanganfifo where KdRuangan='" & mstrKdRuangan & "' and KdBarang= '" & .TextMatrix(.Row, 2) & "' and KdAsal='" & .TextMatrix(.Row, 12) & "' and NoTerima='" & .TextMatrix(.Row, 29) & "'")

                                                '                        Call msubRecFO(dbRst, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & .TextMatrix(.Row, 2) & "','" & .TextMatrix(.Row, 12) & "') as stok")
                                                If .TextMatrix(.Row, 6) = "S" Then
                                                        dblSelisih = dbRst(0) - CDbl(txtIsi.Text)
                                                Else
                                                        dblSelisih = (dbRst(0) * dblJmlTerkecil) - CDbl(txtIsi.Text)
                                                End If

                                                If dblSelisih < 0 Then
                                                        If .TextMatrix(.Row, 6) = "S" Then
                                                                txtIsi.Text = dbRst(0)
                                                        Else
                                                                txtIsi.Text = dbRst(0) * dblJmlTerkecil
                                                        End If

                                                        .TextMatrix(.Row, 9) = FormatPembulatan(dbRst(0), mstrKdInstalasiLogin)
                                                Else
                                                        Set dbRst = Nothing
                                                        '                            Call msubRecFO(dbRst, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & .TextMatrix(.Row, 2) & "','" & .TextMatrix(.Row, 29) & "') ")
                                                        'Call msubRecFO(dbRst, "select JmlStok as stok from stokruanganfifo where KdRuangan='" & mstrKdRuangan & "' and KdBarang= '" & .TextMatrix(.Row, 2) & "' and KdAsal='" & .TextMatrix(.Row, 12) & "' and NoTerima='" & .TextMatrix(.Row, 29) & "'")
                                                        Call msubRecFO(dbRst, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & rsK("KdBarang") & "','" & rsK("KdAsal") & "') as stok")
                                                        .TextMatrix(.Row, 9) = FormatPembulatan(IIf(IsNull(dbRst("Stok")), 0, dbRst("Stok")), mstrKdInstalasiLogin)
                                                End If
                                        End If

                                        'end FIFO

                                        'konvert koma col jumlah
                                        .TextMatrix(.Row, 10) = txtIsi.Text

                                        'konvert koma col discount
                                        .TextMatrix(.Row, 8) = 0 ' .TextMatrix(.Row, 8)

                                        txtIsi.Visible = False

                                        subintJmlService = 1
                                        'rubah jumlah service
                                        .TextMatrix(.Row, 15) = subintJmlService

                                        'CHANDRA untuk mengurangi stok real 10 03 2014
                                        If .TextMatrix(.Row, 6) = "S" Then
                                                If sp_StokRealRuangan(.TextMatrix(.Row, 2), .TextMatrix(.Row, 12), .TextMatrix(.Row, 29), CDbl(.TextMatrix(.Row, 10)), "M") = False Then Exit Sub
                                        End If

                                        'ambil no temporary
                                        ' chandra
                                        'If sp_TempDetailApotikJual(CDbl(.TextMatrix(.Row, 7)) + CDbl(.TextMatrix(.Row, 14)), .TextMatrix(.Row, 2), .TextMatrix(.Row, 12)) = False Then Exit Sub
                                        'ambil hutang penjamin dan tanggungan rs
                                        strSQL = "SELECT HargaSatuan, JmlHutangPenjamin, JmlTanggunganRS" & " FROM TempDetailApotikJual" & " WHERE (NoTemporary = '" & Trim(txtNoTemporary.Text) & "') AND (KdBarang = '" & .TextMatrix(.Row, 2) & "') AND (KdAsal = '" & .TextMatrix(.Row, 12) & "')"
                                        Call msubRecFO(rs, strSQL)

                                        If rs.EOF = True Then
                                                curHutangPenjamin = 0
                                                curTanggunganRS = 0
                                        Else
                                                curHutangPenjamin = rs("JmlHutangPenjamin").Value
                                                curTanggunganRS = rs("JmlTanggunganRS").Value
                                        End If

                                        .TextMatrix(.Row, 14) = subcurTarifService
                                        strSQL = ""
                                        Set rsB = Nothing
                                        strSQL = "SELECT dbo.FB_TakeHargaNettoOA('" & mstrKdPenjaminPasien & "','" & mstrKdJenisPasien & "','" & .TextMatrix(.Row, 2) & "','" & .TextMatrix(.Row, 12) & "','" & .TextMatrix(.Row, 6) & "', '" & mstrKdRuangan & "','" & .TextMatrix(.Row, 29) & "') AS HargaBarang"
                                        Call msubRecFO(rsB, strSQL)

                                        If rsB.EOF = True Then curHargaBrg = 0 Else curHargaBrg = rsB(0).Value

                                        strSQL = ""
                                        Set rs = Nothing
                                        subcurHargaSatuan = 0

                                        strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "','" & .TextMatrix(.Row, 12) & "', " & msubKonversiKomaTitik(CStr(curHargaBrg)) & ")  as HargaSatuan"
                                        Call msubRecFO(rs, strSQL)

                                        If rs.EOF = True Then subcurHargaSatuan = 0 Else subcurHargaSatuan = rs(0).Value
                                        'khusus OA harga tidak dikalikan lg Ppn krn OA termasuk pelayanan yg include ke tindakan (TM)
                                        'subcurHargaSatuan = (subcurHargaSatuan * typSettingDataPendukung.realPPn / 100) + subcurHargaSatuan
                                        .TextMatrix(.Row, 7) = subcurHargaSatuan
                                        .TextMatrix(.Row, 7) = FormatPembulatan(CDbl(subcurHargaSatuan), mstrKdInstalasiLogin)
                                        .TextMatrix(.Row, 16) = CDbl(.TextMatrix(.Row, 7))

                                        .TextMatrix(.Row, 11) = ((CDbl(.TextMatrix(.Row, 14)) * CDbl(.TextMatrix(.Row, 15))) + (CDbl(.TextMatrix(.Row, 16)) * Val(.TextMatrix(.Row, 10)))) '+ val(.TextMatrix(.Row, 26))

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

                                        curHarusDibayar = CDbl(.TextMatrix(.Row, 11)) - (CDbl(.TextMatrix(.Row, 21)) + CDbl(.TextMatrix(.Row, 19)) + CDbl(.TextMatrix(.Row, 20)))
                                        .TextMatrix(.Row, 22) = IIf(curHarusDibayar < 0, 0, curHarusDibayar)

                                        'add for FIFO jika jml yg diinput melebihi stok penerimaan, mk otomatis muncul di row selanjutnya
                                        
                                        If bolStatusFIFO = True Then
                                                If dblSelisih < 0 Then
                                                        jumlahNoTerimaLainnya = jumlahNoTerimaLainnya + 1

                                                        With fgData
                                                                strSQL = "select NoTerima As NoFIFO,JmlStok from V_StokRuanganFIFO where KdBarang='" & .TextMatrix(.Row, 2) & "' and KdAsal='" & .TextMatrix(.Row, 12) & "' and NoTerima<>'" & .TextMatrix(.Row, 29) & "' and KdRuangan='" & mstrKdRuangan & "' and JmlStok<>0 order by TglTerima asc"
                                                                Set dbRst = Nothing
                                                                Call msubRecFO(dbRst, strSQL)

                                                                If dbRst.EOF = False Then
                                                                        dbRst.MoveFirst

                                                                        For i = 1 To dbRst.RecordCount

                                                                                .Rows = .Rows + 1

                                                                                intRowTemp = .Row

                                                                                If .TextMatrix(.Rows - 2, 2) = "" Then
                                                                                        .Row = .Rows - 2
                                                                                Else
                                                                                        .Row = .Rows - 1
                                                                                End If

                                                                                For j = 0 To .Cols - 1
                                                                                        .Col = j
                                                                                        .CellBackColor = vbRed
                                                                                        .CellForeColor = vbWhite
                                                                                Next j

                                                                                .Row = intRowTemp
                                                                                intRowTemp = 0

                                                                                If .TextMatrix(.Rows - 2, 2) = "" Then
                                                                                        intRowTemp = .Rows - 2
                                                                                Else
                                                                                        intRowTemp = .Rows - 1
                                                                                End If

                                                                                curHutangPenjamin = 0
                                                                                curTanggunganRS = 0

                                                                                .TextMatrix(intRowTemp, 0) = .TextMatrix(.Row, 0)
                                                                                .TextMatrix(intRowTemp, 1) = .TextMatrix(.Row, 1)
                                                                                .TextMatrix(intRowTemp, 2) = .TextMatrix(.Row, 2)
                                                                                .TextMatrix(intRowTemp, 3) = .TextMatrix(.Row, 3)
                                                                                .TextMatrix(intRowTemp, 4) = .TextMatrix(.Row, 4)
                                                                                .TextMatrix(intRowTemp, 5) = .TextMatrix(.Row, 5)
                                                                                .TextMatrix(intRowTemp, 6) = .TextMatrix(.Row, 6)
                                                                                .TextMatrix(intRowTemp, 12) = .TextMatrix(.Row, 12)
                                                                                .TextMatrix(intRowTemp, 9) = dbRst("JmlStok")
                                                                                .TextMatrix(intRowTemp, 39) = .TextMatrix(.Row, 39)

                                                                                strNoTerima = dbRst("NoFIFO")
                                                                                .TextMatrix(intRowTemp, 29) = strNoTerima

                                                                                strSQL = ""
                                                                                Set rsB = Nothing
                                                                                strSQL = "SELECT dbo.FB_TakeHargaNettoOA('" & mstrKdPenjaminPasien & "','" & mstrKdJenisPasien & "','" & .TextMatrix(intRowTemp, 2) & "','" & .TextMatrix(intRowTemp, 12) & "','" & .TextMatrix(intRowTemp, 6) & "', '" & mstrKdRuangan & "','" & .TextMatrix(intRowTemp, 29) & "') AS HargaBarang"
                                                                                Call msubRecFO(rsB, strSQL)

                                                                                If rsB.EOF = True Then curHargaBrg = 0 Else curHargaBrg = rsB(0).Value

                                                                                strSQL = ""
                                                                                Set rs = Nothing
                                                                                subcurHargaSatuan = 0

                                                                                strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "','" & .TextMatrix(intRowTemp, 12) & "', " & msubKonversiKomaTitik(CStr(curHargaBrg)) & ")  as HargaSatuan"
                                                                                Call msubRecFO(rs, strSQL)

                                                                                If rs.EOF = True Then subcurHargaSatuan = 0 Else subcurHargaSatuan = rs(0).Value
                                                                                'khusus OA harga tidak dikalikan lg Ppn krn OA termasuk pelayanan yg include ke tindakan (TM)
                                                                                'subcurHargaSatuan = (subcurHargaSatuan * typSettingDataPendukung.realPPn / 100) + subcurHargaSatuan
                                                                                .TextMatrix(intRowTemp, 7) = subcurHargaSatuan
                                                                                .TextMatrix(intRowTemp, 7) = FormatPembulatan(CDbl(subcurHargaSatuan), mstrKdInstalasiLogin)
                                                                                
                                                                                strSQL = "SELECT Discount FROM dbo.HargaNettoBarangFIFO WHERE NoTerima='" & strNoTerima & "' and KdBarang='" & rsK("KdBarang") & "'"
                                                                                Call msubRecFO(rsB, strSQL)
                                                                                .TextMatrix(intRowTemp, 8) = 0 ' FormatPembulatan((IIf(rsb.EOF = True, 0, rsb(0).Value) / 100) * CDbl(subcurHargaSatuan), mstrKdInstalasiLogin)
                                                                                '.TextMatrix(intRowTemp, 8) = (.TextMatrix(.Row, 8) / 100) * subcurHargaSatuan
                                                                                .TextMatrix(intRowTemp, 10) = Abs(dblSelisih)

                                                                                If .TextMatrix(intRowTemp, 6) = "S" Then
                                                                                        dblSelisih = Abs(dblSelisih) - CDbl(dbRst("JmlStok"))
                                                                                Else
                                                                                        dblSelisih = Abs(dblSelisih) - CDbl(dbRst("JmlStok") * dblJmlTerkecil)
                                                                                End If

                                                                                If dblSelisih >= 0 Then
                                                                                        If .TextMatrix(intRowTemp, 6) = "S" Then
                                                                                                .TextMatrix(intRowTemp, 10) = dbRst("JmlStok")
                                                                                        Else
                                                                                                .TextMatrix(intRowTemp, 10) = dbRst("JmlStok") * dblJmlTerkecil
                                                                                        End If
                                                                                End If

                                                                                .TextMatrix(intRowTemp, 13) = .TextMatrix(.Row, 13)

                                                                                .TextMatrix(intRowTemp, 23) = ""
                                                                                .TextMatrix(intRowTemp, 24) = curHargaBrg
                                                                                .TextMatrix(intRowTemp, 25) = .TextMatrix(.Row, 25)
                                                                                .TextMatrix(intRowTemp, 26) = 0
                                                                                .TextMatrix(intRowTemp, 27) = .TextMatrix(.Row, 27)
                                                                                .TextMatrix(intRowTemp, 28) = .TextMatrix(.Row, 28)

                                                                                .TextMatrix(intRowTemp, 14) = .TextMatrix(.Row, 14)
                                                                                .TextMatrix(intRowTemp, 15) = .TextMatrix(.Row, 15)
                                                                                '.TextMatrix(intRowTemp, 26) = .TextMatrix(.Row, 26)
                                                                                .TextMatrix(intRowTemp, 16) = .TextMatrix(intRowTemp, 7)

                                                                                If .TextMatrix(intRowTemp, 6) = "S" Then
                                                                                        If sp_StokRealRuangan(.TextMatrix(intRowTemp, 2), .TextMatrix(intRowTemp, 12), .TextMatrix(intRowTemp, 29), CDbl(.TextMatrix(intRowTemp, 10)), "M") = False Then Exit Sub
                                                                                End If

                                                                                'ambil no temporary
                                                                                ' chandra
                                                                                'If sp_TempDetailApotikJual(CDbl(.TextMatrix(intRowTemp, 7)) + CDbl(.TextMatrix(intRowTemp, 14)), .TextMatrix(intRowTemp, 2), .TextMatrix(intRowTemp, 12)) = False Then Exit Sub
                                                                                'ambil hutang penjamin dan tanggungan rs
                                                                                strSQL = "SELECT HargaSatuan, JmlHutangPenjamin, JmlTanggunganRS" & " FROM TempDetailApotikJual" & " WHERE (NoTemporary = '" & Trim(txtNoTemporary.Text) & "') AND (KdBarang = '" & .TextMatrix(intRowTemp, 2) & "') AND (KdAsal = '" & .TextMatrix(intRowTemp, 12) & "')"
                                                                                Call msubRecFO(rs, strSQL)

                                                                                If rs.EOF = True Then
                                                                                        curHutangPenjamin = 0
                                                                                        curTanggunganRS = 0
                                                                                Else
                                                                                        curHutangPenjamin = rs("JmlHutangPenjamin").Value
                                                                                        curTanggunganRS = rs("JmlTanggunganRS").Value
                                                                                End If

                                                                                .TextMatrix(intRowTemp, 11) = ((CDbl(.TextMatrix(intRowTemp, 14)) * CDbl(.TextMatrix(intRowTemp, 15))) + (CDbl(.TextMatrix(intRowTemp, 16)) * Val(.TextMatrix(intRowTemp, 10)))) ' + val(.TextMatrix(intRowTemp, 26))

                                                                                .TextMatrix(intRowTemp, 17) = curHutangPenjamin
                                                                                .TextMatrix(intRowTemp, 18) = curTanggunganRS

                                                                                If .TextMatrix(intRowTemp, 17) > 0 Then
                                                                                        .TextMatrix(intRowTemp, 19) = (.TextMatrix(intRowTemp, 14) * .TextMatrix(intRowTemp, 15)) + (Val(.TextMatrix(intRowTemp, 10)) * CDbl(.TextMatrix(intRowTemp, 17))) + Val(.TextMatrix(intRowTemp, 26))
                                                                                Else
                                                                                        .TextMatrix(intRowTemp, 19) = 0
                                                                                End If

                                                                                If .TextMatrix(intRowTemp, 18) > 0 Then
                                                                                        .TextMatrix(intRowTemp, 20) = (.TextMatrix(intRowTemp, 14) * .TextMatrix(intRowTemp, 15)) + (Val(.TextMatrix(intRowTemp, 10)) * CDbl(.TextMatrix(intRowTemp, 18))) + Val(.TextMatrix(intRowTemp, 26))
                                                                                Else
                                                                                        .TextMatrix(intRowTemp, 20) = 0
                                                                                End If

                                                                                '.TextMatrix(intRowTemp, 21) = ((CDbl(.TextMatrix(intRowTemp, 8) / 100)) * (CDbl(.TextMatrix(intRowTemp, 10)))) '* CDbl(.TextMatrix(intRowTemp, 16))))

                                                                                'curHarusDibayar = CDbl(.TextMatrix(intRowTemp, 11)) - (CDbl(.TextMatrix(intRowTemp, 21)) + CDbl(.TextMatrix(intRowTemp, 19)) + CDbl(.TextMatrix(intRowTemp, 20)))
                                                                                '.TextMatrix(intRowTemp, 22) = IIf(curHarusDibayar < 0, 0, curHarusDibayar)
                                                                                
                                                                                .TextMatrix(intRowTemp, 21) = CDbl(.TextMatrix(intRowTemp, 10)) * CDbl(.TextMatrix(intRowTemp, 8))

                                                                                curHarusDibayar = CDbl(.TextMatrix(intRowTemp, 11)) - (CDbl(.TextMatrix(intRowTemp, 21)) + CDbl(.TextMatrix(intRowTemp, 19)) + CDbl(.TextMatrix(intRowTemp, 20)))
                                                                                .TextMatrix(intRowTemp, 22) = IIf(curHarusDibayar < 0, 0, curHarusDibayar)

                                                                                If dblSelisih <= 0 Then Exit For

                                                                                dbRst.MoveNext
                                                                        Next i

                                                                End If

                                                        End With

                                                End If
                                        End If

                                        'If bolStatusFIFO = True Then k = iTemp
                                        .Rows = .Rows + 1
                                        rsK.MoveNext

                                        If .TextMatrix(.Row, 9) = "0" Then .TextMatrix(.Row, 9) = 0
                                Next k

                                Call subHitungTotal
                        End With
       
                End If
        End If

        '    End If
   
        chkNoResep.Value = vbChecked

        ' rebuildResekKe
        If chkDokterPemeriksa.Value = vbUnchecked Then txtDokter.Enabled = False

        Exit Sub
    
Errload:
        Call msubPesanError
        '    Resume 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

        Dim q As String

        If mblnForm = True Then mblnForm = False: Exit Sub
        If valEnd = True Then

                q = MsgBox("Yakin akan membatalkan Penjualan Resep Pasien?", vbOKCancel + vbQuestion, "Validasi")

                If q = 2 Then
                        Cancel = 1
                Else

                        'batalain stok yg keluar
                        Dim i As Integer

                        For i = 1 To fgData.Rows - 1

                                If fgData.TextMatrix(i, 2) <> "" And Val(fgData.TextMatrix(i, 10)) <> 0 Then
                                        If fgData.TextMatrix(i, 6) = "S" Then
                                                If sp_StokRealRuangan(fgData.TextMatrix(i, 2), fgData.TextMatrix(i, 12), fgData.TextMatrix(i, 29), CDbl(fgData.TextMatrix(i, 10)), "C") = False Then Exit Sub
                                        End If
                                End If

                        Next i

                        Cancel = 0
                        Unload Me
                End If
        End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

        On Error GoTo Errload

        strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
        Call msubRecFO(rs, strSQL)

        If rs.EOF = False Then
                mstrKdJenisPasien = rs("KdKelompokPasien").Value
                mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
        End If

        If tutup = False Then
                frmTransaksiPasien.Enabled = True
        End If

Errload:
End Sub

Private Sub subHapusDataGridRacikan()

        On Error GoTo Errload

        Dim i, X As Integer

        Dim strResepKe          As String

        Dim intBarisYangDihapus As Integer

        Dim curHarusDibayar     As Currency

        With fgRacikan

                If .Rows = 2 Then
                        If sp_StokRealRuangan(.TextMatrix(.Row, 0), .TextMatrix(.Row, 12), .TextMatrix(.Row, 14), CDbl(.TextMatrix(.Row, 6)), "C") = False Then Exit Sub
                        For i = 0 To .Cols - 1
                                .TextMatrix(1, i) = ""
                        Next i
                      .TextMatrix(1, 2) = noRacikan
                Else

                        'ddddd
                        'chandra 11 03 2014
                        If sp_StokRealRuangan(.TextMatrix(.Row, 0), .TextMatrix(.Row, 12), .TextMatrix(.Row, 14), CDbl(.TextMatrix(.Row, 6)), "C") = False Then Exit Sub
                        'if(.Row < .Rows)
                        
                        .RemoveItem .Row
                        .TextMatrix(.Row, 2) = noRacikan
                End If
         For i = 1 To fgRacikan.Rows - 1
                    fgRacikan.TextMatrix(i, 2) = fgData.TextMatrix(fgData.Row, 0)
                Next i
        End With

        Call subHitungTotal
Errload:
       ' Call msubPesanError
End Sub

Private Sub subHapusDataGrid()

        On Error GoTo Errload

        Dim i, X As Integer

        Dim strResepKe          As String

        Dim intBarisYangDihapus As Integer

        Dim curHarusDibayar     As Currency

        With fgData

                If .Row = 0 Then Exit Sub
                If Val(.TextMatrix(.Row, 11)) = 0 Then GoTo stepHapusData
                intBarisYangDihapus = fgData.Row

                If .TextMatrix(.Row, 1) = "Racikan" Then 'jika obat racikan, pastikan jumlah service 1 untuk resep yang sama
                        strResepKe = .TextMatrix(.Row, 0)
                        strSQL = "select * from RacikanObatPasienTemp where (NoRacikan = '" & fgData.TextMatrix(.Row, 34) & "')"
                        Call msubRecFO(rs, strSQL)
                        'For i = 1 To rs.RecordCount
                         '   If sp_StokRealRuangan(rs("KdBarang").Value, rs("KdAsal").Value, rs("NoTerima").Value, rs("JmlBarang").Value, "C") = False Then Exit Sub
                         '   rs.MoveNext
                        'Next i
                        
                        Call msubRecFO(rs, "execute RestoreStokBarangOtomatis '" & .TextMatrix(.Row, 38) & "'")
                        If Val(.TextMatrix(.Row, 15)) = 0 Then GoTo stepHapusData

                        For i = 1 To .Rows - 2

                                If .TextMatrix(i, 0) = strResepKe And i <> intBarisYangDihapus Then
                                        .TextMatrix(i, 13) = 1

                                        Exit For

                                End If

                        Next i

                End If

stepHapusData:

                ' hapus RacikanObatPasienTemp add by denki
                '        For x = 1 To fgData.Rows - 1
                '            If fgData.TextMatrix(x, 34) <> "0000000000" And fgData.TextMatrix(.Row, 1) = "Racikan" Then
                
                dbConn.Execute "DELETE FROM RacikanObatPasienTemp WHERE (NoRacikan = '" & fgData.TextMatrix(.Row, 34) & "')"
                '            End If
                '        Next x

                dbConn.Execute "DELETE FROM TempDetailApotikJual " & " WHERE (NoTemporary = '" & Trim(.TextMatrix(.Row, 23)) & "') AND (KdBarang = '" & .TextMatrix(.Row, 2) & "') AND (KdAsal = '" & .TextMatrix(.Row, 12) & "')"

                If .Rows = 2 Then
                        If sp_StokRealRuangan(fgData.TextMatrix(.Row, 2), fgData.TextMatrix(.Row, 12), IIf(fgData.TextMatrix(.Row, 29) = "", fgData.TextMatrix(.Row, 27), fgData.TextMatrix(.Row, 29)), IIf(fgData.TextMatrix(.Row, 10) = "", CDbl(0), CDbl(fgData.TextMatrix(.Row, 10))), "C") = False Then Exit Sub
                        For i = 0 To .Cols - 1
                                .TextMatrix(1, i) = ""
                        Next i

                Else
                        If (fgData.TextMatrix(fgData.Row, 10) = "") Then
                            fgData.TextMatrix(fgData.Row, 10) = "0"
                        End If
                        'chandra 11 03 2014
                        If sp_StokRealRuangan(fgData.TextMatrix(.Row, 2), fgData.TextMatrix(.Row, 12), IIf(fgData.TextMatrix(.Row, 29) = "", fgData.TextMatrix(.Row, 27), fgData.TextMatrix(.Row, 29)), IIf(fgData.TextMatrix(.Row, 10) = "", 0, CDbl(fgData.TextMatrix(.Row, 10))), "C") = False Then Exit Sub
                        .RemoveItem .Row
                End If
        
                .TextMatrix(.Row, 15) = IIf(.TextMatrix(.Row, 15) = "", 0, .TextMatrix(.Row, 15))

                If .TextMatrix(.Row, 2) <> "" Then
                        .TextMatrix(.Row, 11) = ((CDbl(.TextMatrix(.Row, 14)) * CDbl(.TextMatrix(.Row, 15))) + (CDbl(.TextMatrix(.Row, 16)) * Val(.TextMatrix(.Row, 10))))

                        curHarusDibayar = CDbl(.TextMatrix(.Row, 11)) - (CDbl(.TextMatrix(.Row, 21)) + CDbl(.TextMatrix(.Row, 19)) + CDbl(.TextMatrix(.Row, 20)))
                        .TextMatrix(.Row, 20) = IIf(curHarusDibayar < 0, 0, curHarusDibayar)
                End If

        End With

        Call subHitungTotal
'        Call rebuildResekKe

        Exit Sub

Errload:
        Call msubPesanError
End Sub

'Private Sub Timer1_Timer()
'dtpTglPelayanan.Value = Now
'End Sub

Private Sub txtBeratObat_KeyDown(KeyCode As Integer, Shift As Integer)

        If KeyCode = vbKeyEscape Then
                fraHitungObat.Visible = False
                fgRacikan.SetFocus
        End If

End Sub

Private Sub txtBeratObat_KeyPress(KeyAscii As Integer)

        On Error GoTo Gelo

        Dim riilnya As Double

        Dim blt     As Integer

        'SetKeyPressToNumber KeyAscii
   
        If KeyAscii = 13 Then
             If (txtBeratObat.Text = "" Or txtBeratObat.Text = "0") Then
            txtBeratObat.Text = "1"
        End If
                If fgRacikan.TextMatrix(fgRacikan.Row, 4) <> 0 Then
                        riilnya = (Val(txtJumlahObatRacik.Text) * Val(fgRacikan.TextMatrix(fgRacikan.Row, 4))) / Val(txtBeratObat.Text)
                Else
                        MsgBox "Kebutuhan /ML Tidak Boleh Nol atau Kosong", vbCritical, "Peringatan"

                        Exit Sub

                End If

                

                'chandra 11 03 2014
                ' proses fifo untuk racikan
                

                '    fgRacikan.TextMatrix(fgRacikan.Row, 5) = Format(fgRacikan.TextMatrix(fgRacikan.Row, 5), "##.##")
                blt = 0
                blt = Round(riilnya, 1)

                If blt < riilnya Then blt = blt + 1

                Dim i              As Integer

                Dim strKdBrg       As String

                Dim strKdAsal      As String

                Dim intRowTemp     As Integer

                Dim dblJmlTerkecil As Double

                Dim dblSelisih     As Double

                With fgRacikan
                        '.TextMatrix(.Row, 5) = 0
                        
                        'add for FIFO validasi jika terjadi edit jml stok, hapus otomatis
                        If bolStatusFIFO = True Then
                                If Trim(.TextMatrix(.Row, 10)) <> "" Then
                                        i = .Rows - 1

                                        strKdBrg = .TextMatrix(.Row, 0)
                                        strKdAsal = .TextMatrix(.Row, 12)

                                        If (.TextMatrix(.Row, 6) = "") Then
                                                .TextMatrix(.Row, 6) = "0"
                                        End If

                                        If sp_StokRealRuangan(.TextMatrix(.Row, 0), .TextMatrix(.Row, 12), .TextMatrix(.Row, 14), .TextMatrix(.Row, 6), "C") = False Then Exit Sub

                                        Do While i <> 1

                                            '.TextMatrix(i, 0) = .TextMatrix(.Row, 0)
                                            '.TextMatrix(i, 12) = .TextMatrix(.Row, 12)
                                                If .TextMatrix(i, 0) <> "" Then
                                                        If (strKdBrg = .TextMatrix(i, 0)) And (strKdAsal = .TextMatrix(i, 12)) Then
                                                                .Row = i

                                                                If .CellBackColor = vbRed Then
                                                                        Call subHapusDataGridRacikan
                                                                        .Row = i - 1
                                                                End If
                                                        End If
                                                End If

                                                i = i - 1
                                        Loop

                                        For i = 1 To .Rows - 1

                                                If (strKdBrg = .TextMatrix(i, 0)) And (strKdAsal = .TextMatrix(i, 12)) Then
                                                        .Row = i

                                                        Exit For

                                                End If

                                        Next i

                                End If

                                .SetFocus
                                intRowTemp = 0
                        End If
                        fgRacikan.TextMatrix(fgRacikan.Row, 6) = CStr(blt)
                        'fgRacikan.TextMatrix(.Row, 7) = FormatPembulatan(riilnya, mstrKdInstalasi)
                        fgRacikan.TextMatrix(fgRacikan.Row, 7) = CStr(riilnya)
                        fgRacikan.TextMatrix(fgRacikan.Row, 5) = "0"
                        'pengambilan jumlah terkecil
                        strSQL = "Select JmlTerkecil From MasterBarang Where KdBarang = '" & fgRacikan.TextMatrix(.Row, 0) & "'"
                        Call msubRecFO(rs, strSQL)
                        dblJmlTerkecil = IIf(rs.EOF, 1, rs(0).Value)

                        'add for FIFO jika jml yg diinput melebihi stok penerimaan, mk otomatis muncul di row selanjutnya
                        If bolStatusFIFO = True Then
                                Set dbRst = Nothing
                                Call msubRecFO(dbRst, "select JmlStok as stok from stokruanganfifo where KdRuangan='" & mstrKdRuangan & "' and KdBarang= '" & .TextMatrix(.Row, 0) & "' and KdAsal='" & .TextMatrix(.Row, 12) & "' and NoTerima='" & .TextMatrix(.Row, 14) & "'")

                                '                        Call msubRecFO(dbRst, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & .TextMatrix(.Row, 2) & "','" & .TextMatrix(.Row, 12) & "') as stok")
                                If .TextMatrix(.Row, 13) = "S" Then
                                        dblSelisih = IIf(dbRst.EOF, 0, dbRst(0)) - CDbl(blt)
                                Else
                                        dblSelisih = (IIf(dbRst.EOF, 0, dbRst(0)) * dblJmlTerkecil) - CDbl(blt)
                                End If

                                If dblSelisih < 0 Then
                                        If .TextMatrix(.Row, 13) = "S" Then
                                                blt = IIf(dbRst.EOF, 0, dbRst(0))
                                        Else
                                                blt = IIf(dbRst.EOF, 0, dbRst(0)) * dblJmlTerkecil
                                        End If
                                        fgRacikan.TextMatrix(fgRacikan.Row, 6) = CStr(IIf(dbRst.EOF, 0, dbRst(0)))
                                        fgRacikan.TextMatrix(fgRacikan.Row, 7) = CStr(IIf(dbRst.EOF, 0, dbRst(0)))
                                        fgRacikan.TextMatrix(fgRacikan.Row, 5) = "0"
                                        riilnya = riilnya - IIf(dbRst.EOF, 0, dbRst(0))
                                        .TextMatrix(.Row, 6) = msubKonversiKomaTitik(IIf(dbRst.EOF, 0, dbRst(0)))
                                Else
                                        Set dbRst = Nothing
                                        '                            Call msubRecFO(dbRst, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & .TextMatrix(.Row, 2) & "','" & .TextMatrix(.Row, 29) & "') ")
                                        'Call msubRecFO(dbRst, "select JmlStok as stok from stokruanganfifo where KdRuangan='" & mstrKdRuangan & "' and KdBarang= '" & .TextMatrix(.Row, 2) & "' and KdAsal='" & .TextMatrix(.Row, 12) & "' and NoTerima='" & .TextMatrix(.Row, 29) & "'")
                                        Call msubRecFO(dbRst, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & .TextMatrix(.Row, 0) & "','" & .TextMatrix(.Row, 12) & "') as stok")
                                        .TextMatrix(.Row, 6) = (blt)
                                End If
                                If .TextMatrix(.Row, 13) = "S" Then
                                   .TextMatrix(.Row, 10) = IIf(dbRst.EOF, 0, dbRst("stok"))
                                Else
                                   .TextMatrix(.Row, 10) = IIf(dbRst.EOF, 0, dbRst("stok")) * dblJmlTerkecil
                                End If
                                strSQL = "SELECT dbo.FB_TakeHargaNettoOA('" & mstrKdPenjaminPasien & "','" & mstrKdJenisPasien & "','" & .TextMatrix(.Row, 0) & "','" & .TextMatrix(.Row, 12) & "','" & .TextMatrix(.Row, 13) & "', '" & mstrKdRuangan & "','" & .TextMatrix(.Row, 14) & "') AS HargaBarang"
                                                                Call msubRecFO(rsB, strSQL)

                                                                If rsB.EOF = True Then .TextMatrix(.Row, 8) = 0 Else .TextMatrix(.Row, 8) = FormatPembulatan(rsB(0).Value, mstrKdInstalasiLogin)

                                                                strSQL = ""
                                                                Set rs = Nothing
                                                                subcurHargaSatuan = 0

                                                                strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "','" & .TextMatrix(.Row, 12) & "', " & msubKonversiKomaTitik(CStr(.TextMatrix(.Row, 8))) & ")  as HargaSatuan"
                                                                Call msubRecFO(rs, strSQL)

                                                                If rs.EOF = True Then .TextMatrix(.Row, 8) = 0 Else .TextMatrix(.Row, 8) = FormatPembulatan(rs(0).Value, mstrKdInstalasiLogin)

                                .TextMatrix(.Row, 9) = FormatPembulatan(CDbl(.TextMatrix(.Row, 6)) * CDbl(.TextMatrix(.Row, 8)), mstrKdInstalasiLogin)
                        End If

                        'If .TextMatrix(.Row, 13) = "S" Then
                                If sp_StokRealRuangan(.TextMatrix(.Row, 0), .TextMatrix(.Row, 12), .TextMatrix(.Row, 14), .TextMatrix(.Row, 6), "M") = False Then Exit Sub
                        'end If

                        ' pengambilan no terima berikutnya
                        '
                        If bolStatusFIFO = True Then
                                If dblSelisih < 0 Then

                                        With fgRacikan
                                                strSQL = "select NoTerima As NoFIFO,JmlStok from V_StokRuanganFIFO where KdBarang='" & .TextMatrix(.Row, 0) & "' and KdAsal='" & .TextMatrix(.Row, 12) & "' and NoTerima<>'" & .TextMatrix(.Row, 14) & "' and KdRuangan='" & mstrKdRuangan & "' and JmlStok<>0 order by TglTerima asc"
                                                Set dbRst = Nothing
                                                Call msubRecFO(dbRst, strSQL)

                                                If dbRst.EOF = False Then
                                                        dbRst.MoveFirst

                                                        For i = 1 To dbRst.RecordCount

                                                                .Rows = .Rows + 1

                                                                intRowTemp = .Row

                                                                If .TextMatrix(.Rows - 2, 2) = "" Then
                                                                        .Row = .Rows - 2
                                                                Else
                                                                        .Row = .Rows - 1
                                                                End If
                                                                Dim j As Integer
                                                                For j = 0 To .Cols - 1
                                                                        .Col = j
                                                                        .CellBackColor = vbRed
                                                                        .CellForeColor = vbWhite
                                                                Next j

                                                                .Row = intRowTemp
                                                                intRowTemp = 0

                                                                If .TextMatrix(.Rows - 2, 2) = "" Then
                                                                        intRowTemp = .Rows - 2
                                                                Else
                                                                        intRowTemp = .Rows - 1
                                                                End If

                                                                curHutangPenjamin = 0
                                                                curTanggunganRS = 0

                                                                

                                                                For j = 0 To 16
                                                                        .TextMatrix(intRowTemp, j) = .TextMatrix(.Row, j)
                                                                Next j
                                                             
                                                                If (dbRst("JmlStok") < Abs(dblSelisih)) Then
                                                                    fgRacikan.TextMatrix(intRowTemp, 6) = dbRst("JmlStok")
                                                                    fgRacikan.TextMatrix(intRowTemp, 7) = dbRst("JmlStok")
                                                                    fgRacikan.TextMatrix(intRowTemp, 5) = "0"
                                                                    riilnya = riilnya - dbRst("JmlStok")
                                                                    .TextMatrix(intRowTemp, 6) = dbRst("JmlStok")
                                                                Else
                                                                   ' riilnya = riilnya - Abs(dblSelisih)
                                                                    fgRacikan.TextMatrix(intRowTemp, 6) = Abs(dblSelisih)
                                                                    fgRacikan.TextMatrix(intRowTemp, 7) = riilnya
                                                                    fgRacikan.TextMatrix(intRowTemp, 5) = "0"
                                                                    
                                                                    .TextMatrix(intRowTemp, 6) = Abs(dblSelisih)
                                                                End If

                                                                strNoTerima = dbRst("NoFIFO")
                                                                .TextMatrix(intRowTemp, 14) = strNoTerima
                                                                'cccc
                                                                strSQL = ""
                                                                Set rsB = Nothing
                                                                '-
                                                                strSQL = "SELECT dbo.FB_TakeHargaNettoOA('" & mstrKdPenjaminPasien & "','" & mstrKdJenisPasien & "','" & .TextMatrix(intRowTemp, 0) & "','" & .TextMatrix(intRowTemp, 12) & "','" & .TextMatrix(intRowTemp, 13) & "', '" & mstrKdRuangan & "','" & .TextMatrix(intRowTemp, 14) & "') AS HargaBarang"
                                                                Call msubRecFO(rsB, strSQL)

                                                                If rsB.EOF = True Then .TextMatrix(intRowTemp, 8) = 0 Else .TextMatrix(intRowTemp, 8) = FormatPembulatan(rsB(0).Value, mstrKdInstalasiLogin)

                                                                strSQL = ""
                                                                Set rs = Nothing
                                                                subcurHargaSatuan = 0

                                                                strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "','" & .TextMatrix(intRowTemp, 12) & "', " & msubKonversiKomaTitik(CStr(.TextMatrix(intRowTemp, 8))) & ")  as HargaSatuan"
                                                                Call msubRecFO(rs, strSQL)
                                                                'salah
                                                                If rs.EOF = True Then subcurHargaSatuan = 0 Else .TextMatrix(intRowTemp, 8) = FormatPembulatan(rs(0).Value, mstrKdInstalasiLogin)

                                                                If .TextMatrix(intRowTemp, 13) = "S" Then
                                                                        dblSelisih = Abs(dblSelisih) - CDbl(dbRst("JmlStok"))
                                                                Else
                                                                        dblSelisih = Abs(dblSelisih) - CDbl(dbRst("JmlStok") * dblJmlTerkecil)
                                                                End If

                                                             '   If dblSelisih >= 0 Then
                                                                        If .TextMatrix(intRowTemp, 13) = "S" Then
                                                                                .TextMatrix(intRowTemp, 10) = dbRst("JmlStok")
                                                                        Else
                                                                                .TextMatrix(intRowTemp, 10) = dbRst("JmlStok") * dblJmlTerkecil
                                                                        End If
                                                              '  End If

                                                                'If .TextMatrix(intRowTemp, 13) = "S" Then
                                                                        If sp_StokRealRuangan(.TextMatrix(intRowTemp, 0), .TextMatrix(intRowTemp, 12), .TextMatrix(intRowTemp, 14), .TextMatrix(intRowTemp, 6), "M") = False Then Exit Sub
                                                                'End If
                                                                .TextMatrix(intRowTemp, 7) = msubKonversiKomaTitik(.TextMatrix(intRowTemp, 7))
                                                                .TextMatrix(intRowTemp, 9) = FormatPembulatan((CDbl(subcurTarifServiceRacikan) * CDbl(subintJmlServiceRacikan)) + (CDbl(.TextMatrix(intRowTemp, 6)) * CDbl(.TextMatrix(intRowTemp, 8))), mstrKdInstalasiLogin)
                                                                .TextMatrix(intRowTemp, 15) = subintJmlServiceRacikan
                                                                .TextMatrix(intRowTemp, 16) = subcurTarifServiceRacikan
                                                                .TextMatrix(intRowTemp, 9) = FormatPembulatan(CDbl(.TextMatrix(intRowTemp, 9)), mstrKdInstalasiLogin)
                                                                If dblSelisih <= 0 Then Exit For

                                                                dbRst.MoveNext
                                                        Next i

                                                End If

                                        End With

                                End If
                        End If

                        .TextMatrix(.Row, 7) = msubKonversiKomaTitik(.TextMatrix(fgRacikan.Row, 7))
                        .TextMatrix(.Row, 9) = FormatPembulatan((CDbl(subcurTarifServiceRacikan) * CDbl(subintJmlServiceRacikan)) + (CDbl(.TextMatrix(.Row, 6)) * CDbl(.TextMatrix(.Row, 8))), mstrKdInstalasiLogin)
                        .TextMatrix(.Row, 15) = subintJmlServiceRacikan
                        .TextMatrix(.Row, 16) = subcurTarifServiceRacikan
                        .TextMatrix(.Row, 9) = FormatPembulatan(CDbl(.TextMatrix(.Row, 9)), mstrKdInstalasiLogin)
                        .SetFocus
                        .Col = 9
                End With

                fraHitungObat.Visible = False
                txtBeratObat.Text = ""
        End If

        Exit Sub

Gelo:
        Call msubPesanError
        'txtBeratObat.SetFocus
End Sub

Private Sub txtDokter_Change()

        On Error GoTo Errload

        mstrFilterDokter = "WHERE NamaDokter like '%" & txtDokter.Text & "%'"
        txtKdDokter.Text = ""
        dgDokter.Visible = True
        Call subLoadDokter

        Exit Sub

Errload:
        Call msubPesanError
End Sub

Private Sub txtDokter_KeyDown(KeyCode As Integer, Shift As Integer)

        Select Case KeyCode

                Case 13, vbKeyDown

                        If dgDokter.Visible = True Then dgDokter.SetFocus Else fgData.SetFocus

                Case vbKeyEscape
                        dgDokter.Visible = False
        End Select

End Sub

Private Sub txtIsi_Change()

        Dim i As Integer

        Select Case fgData.Col

                Case 2 'kode barang

                        If tempStatusTampil = True Then Exit Sub
                        '            strSQL = "execute CariBarangNStokMedis_V '" & txtIsi.Text & "%','" & mstrKdRuangan & "'"
'                        strSQL = "execute CariBarang_V '" & txtIsi.Text & "%','" & mstrKdRuangan & "'"
                        strSQL = "execute CariBarangToAsalBarang_V '" & txtIsi.Text & "%','" & mstrKdRuangan & "','" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "'"
                        Call msubRecFO(dbRst, strSQL)

                        Set dgObatAlkes.DataSource = dbRst

                        Dim data As String

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

                                .Top = 2830
                                .Left = 1820
                                .Visible = True

                                For i = 1 To fgData.Row - 1
                                        .Top = .Top + fgData.RowHeight(i)
                                Next i

                                If fgData.TopRow > 1 Then
                                        .Top = .Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
                                End If

                        End With

                Case 3 ' nama barang

                        If tempStatusTampil = True Then Exit Sub
            
                        '            strSQL = "execute CariBarangNStokMedis_V '" & txtIsi.Text & "%','" & mstrKdRuangan & "'"
'                        strSQL = "execute CariBarang_V '" & txtIsi.Text & "%','" & mstrKdRuangan & "'"
                        strSQL = "execute CariBarangToAsalBarang_V '" & txtIsi.Text & "%','" & mstrKdRuangan & "','" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "'"
                        Call msubRecFO(dbRst, strSQL)

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

                                .Top = 2830
                                .Left = 3000
                                .Visible = True

                                For i = 1 To fgData.Row - 1
                                        .Top = .Top + fgData.RowHeight(i)
                                Next i

                                If fgData.TopRow > 1 Then
                                        .Top = .Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
                                End If

                        End With

                Case Else
                        dgObatAlkes.Visible = False
        End Select

End Sub

Private Sub txtIsi_KeyDown(KeyCode As Integer, Shift As Integer)

        If KeyCode = vbKeyDown Then If dgObatAlkes.Visible = True Then If dgObatAlkes.ApproxCount > 0 Then dgObatAlkes.SetFocus
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)

        On Error Resume Next

        Dim i, j As Integer

        dcJenisObat.Visible = False
        dcNamaPelayananRS.Visible = False

        Dim curHutangPenjamin As Currency

        Dim curTanggunganRS   As Currency

        Dim curHarusDibayar   As Currency

        Dim KdJnsObat         As String

        Dim dblSelisih        As Double

        Dim intRowTemp        As Integer

        Dim strNoTerima       As String

        Dim curHargaBrg       As Currency

        Dim dblSelisihNow     As Double

        Dim dblJmlStokMax     As Double

        Dim strKdBrg          As String

        Dim strKdAsal         As String

        Dim dblJmlTerkecil    As Double

        Dim dblTotalStokK     As Double
        Select Case fgData.Col
            Case 0
               Call SetKeyPressToNumber(KeyAscii)
        End Select
        '    Call SetKeyPressToNumber(KeyAscii)
    
        If KeyAscii = 13 Then

                With fgData

                        Select Case .Col

                                Case 0
                                        dgObatAlkes.Visible = False

                                        If Val(txtIsi.Text) = 0 Then txtIsi.Text = 1
                                        .TextMatrix(.Row, .Col) = CDbl(txtIsi.Text)
                                        txtIsi.Visible = False

                                        dcJenisObat.Left = 120
                                        .Col = 1

                                        For i = 0 To .Col - 1
                                                dcJenisObat.Left = dcJenisObat.Left + .ColWidth(i)
                                        Next i

                                        dcJenisObat.Visible = True
                                        dcJenisObat.Top = .Top - 7

                                        For i = 0 To .Row - 1
                                                dcJenisObat.Top = dcJenisObat.Top + .RowHeight(i)
                                        Next i

                                        If .TopRow > 1 Then
                                                dcJenisObat.Top = dcJenisObat.Top - ((.TopRow - 1) * .RowHeight(1))
                                        End If

                                        dcJenisObat.Width = .ColWidth(.Col)
                                        dcJenisObat.Height = .RowHeight(.Row)

                                        dcJenisObat.Visible = True
                                        dcJenisObat.SetFocus

                                Case 1

                                Case 2

                                        If dgObatAlkes.Visible = True Then
                                                dgObatAlkes.SetFocus

                                                Exit Sub

                                        Else
                                                fgData.SetFocus
                                                fgData.Col = 8
                                        End If

                                Case 3

                                        If dgObatAlkes.Visible = True Then
                                                dgObatAlkes.SetFocus

                                                Exit Sub

                                        Else
                                                fgData.SetFocus
                                                fgData.Col = 8
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

                                                        curHarusDibayar = CDbl(.TextMatrix(.Row, 11)) - (CDbl(.TextMatrix(.Row, 21)) + (CDbl(.TextMatrix(.Row, 19)) + CDbl(.TextMatrix(.Row, 20))))
                                                        .TextMatrix(.Row, 22) = IIf(curHarusDibayar < 0, 0, curHarusDibayar)
                                                        Call subHitungTotal
                                                End If

                                        End If

                                        fgData.SetFocus
                                        fgData.Col = 10

                                Case 10
                    
                                        If (fgData.TextMatrix(.Row, 6) = "S") Then
                    
                                                If bolStatusFIFO = False Then
                                                        If CDbl(txtIsi.Text) > CDbl(.TextMatrix(.Row, 9)) Then
                                                                MsgBox "Jumlah lebih besar dari stock (" & .TextMatrix(.Row, 9) & ")", vbExclamation, "Validasi"
                                                                txtIsi.SelStart = 0: txtIsi.SelLength = Len(txtIsi.Text)

                                                                Exit Sub

                                                        End If
                                                End If

                                        ElseIf (fgData.TextMatrix(.Row, 6) = "K") Then
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
                    
                                        If fgData.TextMatrix(.Row, 1) = "" Then
                                                MsgBox "Lengkapi jenis obat", vbExclamation, "Validasi": Exit Sub
                                        End If

                                        'add for FIFO validasi jika terjadi edit jml stok, hapus otomatis
                                        If bolStatusFIFO = True Then
                                                If Trim(.TextMatrix(.Row, 10)) <> "" Then
                                                        i = .Rows - 1
                                                        strKdBrg = .TextMatrix(.Row, 2)
                                                        strKdAsal = .TextMatrix(.Row, 12)

                                                        If sp_StokRealRuangan(.TextMatrix(.Row, 2), .TextMatrix(.Row, 12), .TextMatrix(.Row, 29), CDbl(.TextMatrix(.Row, 10)), "C") = False Then Exit Sub

                                                        Do While i <> 1

                                                                If .TextMatrix(i, 2) <> "" Then
                                                                        If (strKdBrg = .TextMatrix(i, 2)) And (strKdAsal = .TextMatrix(i, 12)) Then
                                                                                .Row = i

                                                                                If .CellBackColor = vbRed Then
                                                                                        Call subHapusDataGrid
                                                                                        .Row = i - 1
                                                                                End If
                                                                        End If
                                                                End If

                                                                i = i - 1
                                                        Loop

                                                        For i = 1 To .Rows - 1

                                                                If (strKdBrg = .TextMatrix(i, 2)) And (strKdAsal = .TextMatrix(i, 12)) Then
                                                                        .Row = i

                                                                        Exit For

                                                                End If

                                                        Next i

                                                End If

                                                .SetFocus
                                                intRowTemp = 0
                                        End If

                                        'add for FIFO jika jml yg diinput melebihi stok penerimaan, mk otomatis muncul di row selanjutnya
                                        If bolStatusFIFO = True Then
                                                Set dbRst = Nothing
                                                Call msubRecFO(dbRst, "select JmlStok as stok from stokruanganfifo where KdRuangan='" & mstrKdRuangan & "' and KdBarang= '" & .TextMatrix(.Row, 2) & "' and KdAsal='" & .TextMatrix(.Row, 12) & "' and NoTerima='" & .TextMatrix(.Row, 29) & "'")

                                                '                        Call msubRecFO(dbRst, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & .TextMatrix(.Row, 2) & "','" & .TextMatrix(.Row, 12) & "') as stok")
                                                If .TextMatrix(.Row, 6) = "S" Then
                                                        dblSelisih = dbRst(0) - CDbl(txtIsi.Text)
                                                Else
                                                        dblSelisih = (dbRst(0) * dblJmlTerkecil) - CDbl(txtIsi.Text)
                                                End If

                                                If dblSelisih < 0 Then
                                                        If .TextMatrix(.Row, 6) = "S" Then
                                                                txtIsi.Text = dbRst(0)
                                                        Else
                                                                txtIsi.Text = dbRst(0) * dblJmlTerkecil
                                                        End If

                                                        .TextMatrix(.Row, 9) = FormatPembulatan(dbRst(0), mstrKdInstalasiLogin)
                                                Else
                                                        Set dbRst = Nothing
                                                        '                            Call msubRecFO(dbRst, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & .TextMatrix(.Row, 2) & "','" & .TextMatrix(.Row, 29) & "') ")
                                                        'Call msubRecFO(dbRst, "select JmlStok as stok from stokruanganfifo where KdRuangan='" & mstrKdRuangan & "' and KdBarang= '" & .TextMatrix(.Row, 2) & "' and KdAsal='" & .TextMatrix(.Row, 12) & "' and NoTerima='" & .TextMatrix(.Row, 29) & "'")
                                                        Call msubRecFO(dbRst, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & dgObatAlkes.Columns("KdBarang") & "','" & dgObatAlkes.Columns("KdAsal") & "') as stok")
                                                        .TextMatrix(.Row, 9) = FormatPembulatan(IIf(IsNull(dbRst("Stok")), 0, dbRst("Stok")), mstrKdInstalasiLogin)
                                                End If
                                        End If

                                        'end FIFO

                                        'konvert koma col jumlah
                                        .TextMatrix(.Row, .Col) = txtIsi.Text
                                        .TextMatrix(.Row, 39) = "0"
                                        'konvert koma col discount
                                        .TextMatrix(.Row, 8) = 0 ' .TextMatrix(.Row, 8)

                                        txtIsi.Visible = False

                                        subintJmlService = 1
                                        'rubah jumlah service
                                        .TextMatrix(.Row, 15) = subintJmlService

                                        'CHANDRA untuk mengurangi stok real 10 03 2014
                                        If .TextMatrix(.Row, 6) = "S" Then
                                                If sp_StokRealRuangan(.TextMatrix(.Row, 2), .TextMatrix(.Row, 12), .TextMatrix(.Row, 29), CDbl(.TextMatrix(.Row, 10)), "M") = False Then Exit Sub
                                        End If

                                        'ambil no temporary
                                        ' chandra
                                        If sp_TempDetailApotikJual(CDbl(.TextMatrix(.Row, 7)) + CDbl(.TextMatrix(.Row, 14)), .TextMatrix(.Row, 2), .TextMatrix(.Row, 12)) = False Then Exit Sub
                                        'ambil hutang penjamin dan tanggungan rs
                                        strSQL = "SELECT HargaSatuan, JmlHutangPenjamin, JmlTanggunganRS" & " FROM TempDetailApotikJual" & " WHERE (NoTemporary = '" & Trim(txtNoTemporary.Text) & "') AND (KdBarang = '" & .TextMatrix(.Row, 2) & "') AND (KdAsal = '" & .TextMatrix(.Row, 12) & "')"
                                        Call msubRecFO(rs, strSQL)

                                        If rs.EOF = True Then
                                                curHutangPenjamin = 0
                                                curTanggunganRS = 0
                                        Else
                                                curHutangPenjamin = rs("JmlHutangPenjamin").Value
                                                curTanggunganRS = rs("JmlTanggunganRS").Value
                                        End If

                                        .TextMatrix(.Row, 14) = subcurTarifService
                                        .TextMatrix(.Row, 16) = CDbl(.TextMatrix(.Row, 7))

                                        .TextMatrix(.Row, 11) = ((CDbl(.TextMatrix(.Row, 14)) * CDbl(.TextMatrix(.Row, 15))) + (CDbl(.TextMatrix(.Row, 16)) * Val(.TextMatrix(.Row, 10)))) '+ val(.TextMatrix(.Row, 26))

                                        .TextMatrix(.Row, 17) = curHutangPenjamin
                                        .TextMatrix(.Row, 18) = curTanggunganRS

                                        If curHutangPenjamin > 0 Then
                                                .TextMatrix(.Row, 19) = (.TextMatrix(.Row, 14) * .TextMatrix(.Row, 15)) + (Val(.TextMatrix(.Row, 10)) * CDbl(.TextMatrix(.Row, 17))) ' + val(.TextMatrix(.Row, 26))
                                        Else
                                                .TextMatrix(.Row, 19) = 0
                                        End If

                                        If curTanggunganRS > 0 Then
                                                .TextMatrix(.Row, 20) = (.TextMatrix(.Row, 14) * .TextMatrix(.Row, 15)) + (Val(.TextMatrix(.Row, 10)) * CDbl(.TextMatrix(.Row, 18))) ' + val(.TextMatrix(.Row, 26))
                                        Else
                                                .TextMatrix(.Row, 20) = 0
                                        End If

                                        .TextMatrix(.Row, 21) = CDbl(.TextMatrix(.Row, 10)) * CDbl(.TextMatrix(.Row, 8))

                                        curHarusDibayar = CDbl(.TextMatrix(.Row, 11)) - (CDbl(.TextMatrix(.Row, 21)) + CDbl(.TextMatrix(.Row, 19)) + CDbl(.TextMatrix(.Row, 20)))
                                        .TextMatrix(.Row, 22) = IIf(curHarusDibayar < 0, 0, curHarusDibayar)

                                        'add for FIFO jika jml yg diinput melebihi stok penerimaan, mk otomatis muncul di row selanjutnya
                                        
                                        If bolStatusFIFO = True Then
                                                If dblSelisih < 0 Then

                                                        With fgData
                                                                strSQL = "select NoTerima As NoFIFO,JmlStok from V_StokRuanganFIFO where KdBarang='" & .TextMatrix(.Row, 2) & "' and KdAsal='" & .TextMatrix(.Row, 12) & "' and NoTerima<>'" & .TextMatrix(.Row, 29) & "' and KdRuangan='" & mstrKdRuangan & "' and JmlStok<>0 order by TglTerima asc"
                                                                Set dbRst = Nothing
                                                                Call msubRecFO(dbRst, strSQL)

                                                                If dbRst.EOF = False Then
                                                                        dbRst.MoveFirst

                                                                        For i = 1 To dbRst.RecordCount

                                                                                .Rows = .Rows + 1

                                                                                intRowTemp = .Row

                                                                                If .TextMatrix(.Rows - 2, 2) = "" Then
                                                                                        .Row = .Rows - 2
                                                                                Else
                                                                                        .Row = .Rows - 1
                                                                                End If

                                                                                For j = 0 To .Cols - 1
                                                                                        .Col = j
                                                                                        .CellBackColor = vbRed
                                                                                        .CellForeColor = vbWhite
                                                                                Next j

                                                                                .Row = intRowTemp
                                                                                intRowTemp = 0

                                                                                If .TextMatrix(.Rows - 2, 2) = "" Then
                                                                                        intRowTemp = .Rows - 2
                                                                                Else
                                                                                        intRowTemp = .Rows - 1
                                                                                End If

                                                                                curHutangPenjamin = 0
                                                                                curTanggunganRS = 0

                                                                                .TextMatrix(intRowTemp, 0) = .TextMatrix(.Row, 0)
                                                                                .TextMatrix(intRowTemp, 1) = .TextMatrix(.Row, 1)
                                                                                .TextMatrix(intRowTemp, 2) = .TextMatrix(.Row, 2)
                                                                                .TextMatrix(intRowTemp, 3) = .TextMatrix(.Row, 3)
                                                                                .TextMatrix(intRowTemp, 4) = .TextMatrix(.Row, 4)
                                                                                .TextMatrix(intRowTemp, 5) = .TextMatrix(.Row, 5)
                                                                                .TextMatrix(intRowTemp, 6) = .TextMatrix(.Row, 6)
                                                                                .TextMatrix(intRowTemp, 12) = .TextMatrix(.Row, 12)
                                                                                .TextMatrix(intRowTemp, 9) = dbRst("JmlStok")
                                                                                .TextMatrix(intRowTemp, 23) = .TextMatrix(.Row, 23)
                                                                                .TextMatrix(intRowTemp, 25) = .TextMatrix(.Row, 25)
                                                                                .TextMatrix(intRowTemp, 39) = .TextMatrix(.Row, 39)
                                                                                '.TextMatrix(intRowTemp, 39) = "0"
                                                                                strNoTerima = dbRst("NoFIFO")
                                                                                .TextMatrix(intRowTemp, 29) = strNoTerima

                                                                                strSQL = ""
                                                                                Set rsB = Nothing
                                                                                strSQL = "SELECT dbo.FB_TakeHargaNettoOA('" & mstrKdPenjaminPasien & "','" & mstrKdJenisPasien & "','" & .TextMatrix(intRowTemp, 2) & "','" & .TextMatrix(intRowTemp, 12) & "','" & .TextMatrix(intRowTemp, 6) & "', '" & mstrKdRuangan & "','" & .TextMatrix(intRowTemp, 29) & "') AS HargaBarang"
                                                                                Call msubRecFO(rsB, strSQL)

                                                                                If rsB.EOF = True Then curHargaBrg = 0 Else curHargaBrg = rsB(0).Value

                                                                                strSQL = ""
                                                                                Set rs = Nothing
                                                                                subcurHargaSatuan = 0

                                                                                strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "','" & .TextMatrix(intRowTemp, 12) & "', " & msubKonversiKomaTitik(CStr(curHargaBrg)) & ")  as HargaSatuan"
                                                                                Call msubRecFO(rs, strSQL)

                                                                                If rs.EOF = True Then subcurHargaSatuan = 0 Else subcurHargaSatuan = rs(0).Value
                                                                                'khusus OA harga tidak dikalikan lg Ppn krn OA termasuk pelayanan yg include ke tindakan (TM)
                                                                                'subcurHargaSatuan = (subcurHargaSatuan * typSettingDataPendukung.realPPn / 100) + subcurHargaSatuan
                                                                                .TextMatrix(intRowTemp, 7) = subcurHargaSatuan
                                                                                .TextMatrix(intRowTemp, 7) = FormatPembulatan(CDbl(subcurHargaSatuan), mstrKdInstalasiLogin)
                                                                                
                                                                                strSQL = "SELECT Discount FROM dbo.HargaNettoBarangFIFO WHERE NoTerima='" & strNoTerima & "' and KdBarang='" & dgObatAlkes.Columns("KdBarang") & "'"
                                                                                Call msubRecFO(rsB, strSQL)
                                                                                .TextMatrix(intRowTemp, 8) = 0 ' FormatPembulatan((IIf(rsb.EOF = True, 0, rsb(0).Value) / 100) * CDbl(subcurHargaSatuan), mstrKdInstalasiLogin)
                                                                                '.TextMatrix(intRowTemp, 8) = (.TextMatrix(.Row, 8) / 100) * subcurHargaSatuan
                                                                                .TextMatrix(intRowTemp, 10) = Abs(dblSelisih)

                                                                                If .TextMatrix(intRowTemp, 6) = "S" Then
                                                                                        dblSelisih = Abs(dblSelisih) - CDbl(dbRst("JmlStok"))
                                                                                Else
                                                                                        dblSelisih = Abs(dblSelisih) - CDbl(dbRst("JmlStok") * dblJmlTerkecil)
                                                                                End If

                                                                                If dblSelisih >= 0 Then
                                                                                        If .TextMatrix(intRowTemp, 6) = "S" Then
                                                                                                .TextMatrix(intRowTemp, 10) = dbRst("JmlStok")
                                                                                        Else
                                                                                                .TextMatrix(intRowTemp, 10) = dbRst("JmlStok") * dblJmlTerkecil
                                                                                        End If
                                                                                End If

                                                                                .TextMatrix(intRowTemp, 13) = .TextMatrix(.Row, 13)

                                                                                .TextMatrix(intRowTemp, 23) = ""
                                                                                .TextMatrix(intRowTemp, 24) = curHargaBrg
                                                                                .TextMatrix(intRowTemp, 25) = .TextMatrix(.Row, 25)
                                                                                .TextMatrix(intRowTemp, 26) = 0
                                                                                .TextMatrix(intRowTemp, 27) = .TextMatrix(.Row, 27)
                                                                                .TextMatrix(intRowTemp, 28) = .TextMatrix(.Row, 28)

                                                                                .TextMatrix(intRowTemp, 14) = .TextMatrix(.Row, 14)
                                                                                .TextMatrix(intRowTemp, 15) = .TextMatrix(.Row, 15)
                                                                                '.TextMatrix(intRowTemp, 26) = .TextMatrix(.Row, 26)
                                                                                .TextMatrix(intRowTemp, 16) = .TextMatrix(intRowTemp, 7)

                                                                                If .TextMatrix(intRowTemp, 6) = "S" Then
                                                                                        If sp_StokRealRuangan(.TextMatrix(intRowTemp, 2), .TextMatrix(intRowTemp, 12), .TextMatrix(intRowTemp, 29), CDbl(.TextMatrix(intRowTemp, 10)), "M") = False Then Exit Sub
                                                                                End If

                                                                                'ambil no temporary
                                                                                ' chandra
                                                                                If sp_TempDetailApotikJual(CDbl(.TextMatrix(intRowTemp, 7)) + CDbl(.TextMatrix(intRowTemp, 14)), .TextMatrix(intRowTemp, 2), .TextMatrix(intRowTemp, 12)) = False Then Exit Sub
                                                                                'ambil hutang penjamin dan tanggungan rs
                                                                                
                                                                                strSQL = "SELECT HargaSatuan, JmlHutangPenjamin, JmlTanggunganRS" & " FROM TempDetailApotikJual" & " WHERE (NoTemporary = '" & Trim(txtNoTemporary.Text) & "') AND (KdBarang = '" & .TextMatrix(intRowTemp, 2) & "') AND (KdAsal = '" & .TextMatrix(intRowTemp, 12) & "')"
                                                                                Call msubRecFO(rs, strSQL)

                                                                                If rs.EOF = True Then
                                                                                        curHutangPenjamin = 0
                                                                                        curTanggunganRS = 0
                                                                                Else
                                                                                        curHutangPenjamin = rs("JmlHutangPenjamin").Value
                                                                                        curTanggunganRS = rs("JmlTanggunganRS").Value
                                                                                End If

                                                                                .TextMatrix(intRowTemp, 11) = ((CDbl(.TextMatrix(intRowTemp, 14)) * CDbl(.TextMatrix(intRowTemp, 15))) + (CDbl(.TextMatrix(intRowTemp, 16)) * Val(.TextMatrix(intRowTemp, 10)))) ' + val(.TextMatrix(intRowTemp, 26))

                                                                                .TextMatrix(intRowTemp, 17) = curHutangPenjamin
                                                                                .TextMatrix(intRowTemp, 18) = curTanggunganRS

                                                                                If .TextMatrix(intRowTemp, 17) > 0 Then
                                                                                        .TextMatrix(intRowTemp, 19) = (.TextMatrix(intRowTemp, 14) * .TextMatrix(intRowTemp, 15)) + (Val(.TextMatrix(intRowTemp, 10)) * CDbl(.TextMatrix(intRowTemp, 17))) + Val(.TextMatrix(intRowTemp, 26))
                                                                                Else
                                                                                        .TextMatrix(intRowTemp, 19) = 0
                                                                                End If

                                                                                If .TextMatrix(intRowTemp, 18) > 0 Then
                                                                                        .TextMatrix(intRowTemp, 20) = (.TextMatrix(intRowTemp, 14) * .TextMatrix(intRowTemp, 15)) + (Val(.TextMatrix(intRowTemp, 10)) * CDbl(.TextMatrix(intRowTemp, 18))) + Val(.TextMatrix(intRowTemp, 26))
                                                                                Else
                                                                                        .TextMatrix(intRowTemp, 20) = 0
                                                                                End If

                                                                                '.TextMatrix(intRowTemp, 21) = ((CDbl(.TextMatrix(intRowTemp, 8) / 100)) * (CDbl(.TextMatrix(intRowTemp, 10)))) '* CDbl(.TextMatrix(intRowTemp, 16))))

                                                                                'curHarusDibayar = CDbl(.TextMatrix(intRowTemp, 11)) - (CDbl(.TextMatrix(intRowTemp, 21)) + CDbl(.TextMatrix(intRowTemp, 19)) + CDbl(.TextMatrix(intRowTemp, 20)))
                                                                                '.TextMatrix(intRowTemp, 22) = IIf(curHarusDibayar < 0, 0, curHarusDibayar)
                                                                                
                                                                                .TextMatrix(intRowTemp, 21) = CDbl(.TextMatrix(intRowTemp, 10)) * CDbl(.TextMatrix(intRowTemp, 8))

                                                                                curHarusDibayar = CDbl(.TextMatrix(intRowTemp, 11)) - (CDbl(.TextMatrix(intRowTemp, 21)) + CDbl(.TextMatrix(intRowTemp, 19)) + CDbl(.TextMatrix(intRowTemp, 20)))
                                                                                .TextMatrix(intRowTemp, 22) = IIf(curHarusDibayar < 0, 0, curHarusDibayar)

                                                                                If dblSelisih <= 0 Then Exit For

                                                                                dbRst.MoveNext
                                                                        Next i

                                                                End If

                                                        End With

                                                End If
                                        End If

                                        Call subHitungTotal
                                        'end fifo

                                        fgData.SetFocus
                                        fgData.Col = 27
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

End Sub

Private Sub txtIsi_LostFocus()
        txtIsi.Visible = False
        If (NotUseRacikan = False) Then
            cmdGenerik.Visible = True
        End If
        
End Sub

Private Sub subSetGrid()

        On Error GoTo Errload

        With fgData
                .Clear
                .Rows = 2
                .Cols = 42
        
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
                .TextMatrix(0, 22) = "TotalHarusBayar" 'curHarusDibayar = Total Harga - TotalDiscount - TotalHutangPenjamin - TotalTanggunganRS IIf curHarusDibayar < 0, 0, curHarusDibayar
                .TextMatrix(0, 23) = "NoTemp" 'jika barang dihapus digrid, hapus ke tabel TempDetailApotikJual
                .TextMatrix(0, 24) = "HargaBeli" 'harga satuan sebelum take tarif dan sebelum ditambah tarif service
                .TextMatrix(0, 25) = "KdJenisObat"
                .TextMatrix(0, 26) = "BiayaAdministrasi"
                        
                .TextMatrix(0, 27) = "Kirim"
                .TextMatrix(0, 28) = "StatusOrder" 'for pesan pelayanan
                .TextMatrix(0, 29) = "NoTerima"
                .TextMatrix(0, 30) = "Aturan Pakai"
                .TextMatrix(0, 31) = "Keterangan Pakai"
                .TextMatrix(0, 32) = "Keterangan Waktu"
                .TextMatrix(0, 33) = "Keterangan Lainnya"
                .TextMatrix(0, 34) = "NoRacikan"
                .TextMatrix(0, 35) = "KdSatuanEtiket"
                .TextMatrix(0, 36) = "KdWaktuEtiket"
                .TextMatrix(0, 37) = "kdWaktuEtiket2"
                .TextMatrix(0, 38) = "UniqeId"
                .TextMatrix(0, 39) = "NoOrder"
                .TextMatrix(0, 40) = "KodePelayananRS"
                .TextMatrix(0, 41) = "Pemakaian Pemeriksaan"
        
                .ColWidth(0) = 500
                .ColWidth(1) = 1200
                .ColWidth(2) = 1200
                .ColWidth(3) = 2800
                .ColWidth(4) = 0
                .ColWidth(5) = 1100
                .ColWidth(6) = 0
                .ColWidth(7) = 1200
                .ColWidth(8) = 0
                .ColWidth(9) = 800
                .ColWidth(10) = 800
                .ColWidth(11) = 1200
                .ColWidth(12) = 0
                .ColWidth(13) = 0
                .ColWidth(14) = 0
                .ColWidth(15) = 0
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
                .ColWidth(27) = 0 '800
                .ColWidth(28) = 0
                .ColWidth(29) = 0
                 If (NotUseRacikan = True) Then
                    .ColWidth(30) = 0
                    .ColWidth(31) = 0
                Else
                    .ColWidth(30) = 1300
                    .ColWidth(31) = 2500
                End If
                
                .ColWidth(32) = 0 '1850
                .ColWidth(33) = 0 '1800
                .ColWidth(34) = 0
                .ColWidth(35) = 0
                .ColWidth(36) = 0
                .ColWidth(37) = 0
                .ColWidth(38) = 0
                .ColWidth(39) = 0
                If (NotUseRacikan = True) Then
                    .ColWidth(40) = 0
                    .ColWidth(41) = 2000
                Else
                    .ColWidth(40) = 0
                    .ColWidth(41) = 0
                End If
        
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

Errload:
        Call msubPesanError
End Sub

Private Sub subLoadDcSource()

        On Error GoTo Errload

      If (NotUseRacikan = True) Then
            Call msubDcSource(dcJenisObat, rs, "SELECT     JenisObat.KdJenisObat, JenisObat.JenisObat FROM         JenisObat INNER JOIN  SettingGlobal ON JenisObat.KdJenisObat <> SettingGlobal.Value WHERE     (JenisObat.StatusEnabled = 1) and SettingGlobal.Prefix='KdJenisObatRacikan' ORDER BY JenisObat.JenisObat")

            If rs.EOF = False Then dcJenisObat.BoundText = rs(0).Value
        Else
            Call msubDcSource(dcJenisObat, rs, "SELECT KdJenisObat, JenisObat FROM JenisObat where StatusEnabled=1 ORDER BY JenisObat")

            If rs.EOF = False Then dcJenisObat.BoundText = rs(0).Value
        End If
        Call msubDcSource(dcAturanPakai, rs, "select KdSatuanEtiket,SatuanEtiket from SatuanEtiketResep Order By SatuanEtiket")

        If rs.EOF = False Then dcAturanPakai.BoundText = rs(0).Value
    
        Call msubDcSource(dcKeteranganPakai, rs, "select KdWaktuEtiket,WaktuEtiket from WaktuEtiketResep order by KdWaktuEtiket")

        If rs.EOF = False Then dcKeteranganPakai.BoundText = rs(0).Value
        dcKeteranganPakai.Text = "Sesudah Makan"
        
   

        strSQL = "SELECT  TOP (200) KdPelayananRS, NamaPelayanan, NoPendaftaran" & " FROM  V_NamaPelayananPerPasien where NoPendaftaran ='" & mstrNoPen & "'"
        Call msubDcSource(dcNamaPelayananRS, rs, strSQL)
    
        Exit Sub

Errload:
        Call msubPesanError
End Sub

Private Sub TxtIsiRacikan_Change()

        On Error GoTo hell

        Dim i, kolom As Integer

        kolom = 0

        Select Case fgRacikan.Col
        
                Case 3 ' nama barang
                        'If tempStatusTampil = True Then Exit Sub
                        '            If subJenisHargaNetto = 2 Then
                        '                strSQL = "select  TOP 100 JenisBarang, RuanganPelayanan, NamaBarang, Kekuatan, AsalBarang, Satuan, HargaBarang, JmlStok, Discount, KdBarang, KdAsal from V_HargaBarangNStok2 " & _
                        '                    " where NamaBarang like '" & TxtIsiRacikan.Text & "%' AND KdRuangan like'" & mstrKdRuangan & "%' ORDER BY NamaBarang"
                        '            Else
                        '            If bolStatusFIFO = False Then
                        '                strSQL = "select  TOP 100 JenisBarang, RuanganPelayanan, NamaBarang, Kekuatan, AsalBarang, Satuan, HargaBarang, JmlStok, Discount, KdBarang, KdAsal from V_HargaBarangNStok1 " & _
                        '                    " where NamaBarang like '" & TxtIsiRacikan.Text & "%' AND KdRuangan like '" & mstrKdRuangan & "%' ORDER BY NamaBarang"
                        '            Else
                        '
                        '
                        '                strSQL = "select  TOP 100 DetailJenisBarang AS JenisBarang, NamaRuangan AS RuanganPelayanan, NamaBarang, Kekuatan, NamaAsal AS AsalBarang, Satuan, HargaNetto1 AS HargaBarang, JmlStok, Discount, KdBarang, KdAsal from V_StokNHargaGlobalFIFO " & _
                        '                    " where NamaBarang like '" & TxtIsiRacikan.Text & "%' AND KdRuangan like '" & mstrKdRuangan & "%' ORDER BY NamaBarang"
                        '
                        '            End If
                        '            End If

                        strSQL = "execute CariBarangNStokMedis_V '" & TxtIsiRacikan.Text & "%','" & mstrKdRuangan & "'"

                        Call msubRecFO(dbRst, strSQL)
            
                        Set dgObatAlkesRacikan.DataSource = dbRst

                        With dgObatAlkesRacikan
            
                                '                 For i = 0 To .Columns.Count - 1
                                '                    .Columns(i).Width = 0
                                '                Next i
                
                                .Columns("RuanganPelayanan").Width = 0
                                .Columns("JenisBarang").Width = 1500
                                .Columns("NamaBarang").Width = 3000
                                .Columns("Kekuatan").Width = 0
                                .Columns("AsalBarang").Width = 1500
                                .Columns("Satuan").Width = 0
                                '                .Columns("HargaBarang").Width = 1500
                                .Columns("JmlStok").Width = 0
                                '
                                '                .Columns("HargaBarang").NumberFormat = "#,###.00"
                                '                .Columns("HargaBarang").Alignment = dbgRight
                                '
                                .Columns("JmlStok").NumberFormat = "#,###.00"
                                .Columns("JmlStok").Alignment = dbgRight
                                '                .Columns("Discount").Width = 0
                                .Columns("KdBarang").Width = 0
                                .Columns("KdAsal").Width = 0
                                .Columns("KdRuangan").Width = 0
                                .Columns("JenisHarga").Width = 0
                                .Columns("KdGenerikbarang").Width = 0
                                .Columns("Discount").Width = 0
                                .Columns("NamaGenerik").Width = 0
                
                                .Top = TxtIsiRacikan.Top + TxtIsiRacikan.Height
                                .Left = TxtIsiRacikan.Left '+ TxtIsiRacikan.Height
                                .Visible = True
               
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

Private Sub TxtIsiRacikan_KeyDown(KeyCode As Integer, Shift As Integer)

        If KeyCode = vbKeyDown Then If dgObatAlkesRacikan.Visible = True Then If dgObatAlkesRacikan.ApproxCount > 0 Then dgObatAlkesRacikan.SetFocus
        ' If KeyCode = 13 Then If dgObatAlkesRacikan.Visible = True Then If dgObatAlkesRacikan.ApproxCount > 0 Then dgObatAlkesRacikan.SetFocus

End Sub

Private Sub TxtIsiRacikan_KeyPress(KeyAscii As Integer)

        Dim i, riilnya, subintJmlServiceRacikan As Integer

        Dim curHutangPenjamin         As Currency

        Dim curTanggunganRS           As Currency

        Dim curHarusDibayar           As Currency

        Dim subcurTarifServiceRacikan As Currency

        Dim KdJnsObat                 As String

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
                    
                                        If txtJumlahObatRacik.Text = "" Then
                                                MsgBox "Jumlah Obat Racik Harus Di isi", vbCritical, "Informasi"
    
                                                txtJumlahObatRacik.SetFocus

                                                Exit Sub

                                        End If
                    
                                        .TextMatrix(.Row, .Col) = msubKonversiKomaTitik(TxtIsiRacikan.Text)

                                        If Val(TxtIsiRacikan.Text) = 0 Then TxtIsiRacikan.Text = 0
                                        TxtIsiRacikan.Visible = False
                                        fraHitungObat.Visible = True
                                        .TextMatrix(.Row, .Col) = msubKonversiKomaTitik(TxtIsiRacikan.Text)
                                        txtBeratObat.SetFocus
                    
                                Case 5 'Kebutuhan Tablet
                                        'konvert koma col jumlah
                    
                                        If txtJumlahObatRacik.Text = "" Then
                                                MsgBox "Jumlah Obat Racik Harus Di isi", vbCritical, "Informasi"
    
                                                txtJumlahObatRacik.SetFocus

                                                Exit Sub

                                        End If
                    
                                        
                                        TxtIsiRacikan.Visible = False
                    
                                        If Val(TxtIsiRacikan.Text) = 0 Then TxtIsiRacikan.Text = 0
                                   
                                       ' If .TextMatrix(.Row, .Col) <> 0 Then
                                                riilnya = (Val(txtJumlahObatRacik.Text) * Val(TxtIsiRacikan.Text))
                                        'Else
                                               ' MsgBox "GUNAKAN TITIK!", vbCritical, "Peringatan"

                                                'Exit Sub

                                       ' End If
                                        
                                        '                    .TextMatrix(.Row, 7) = .TextMatrix(.Row, 6)
                                        '.TextMatrix(.Row, 9) = val(.TextMatrix(.Row, 7)) * msubKonversiKomaTitik(.TextMatrix(.Row, 8)) 'Val(.TextMatrix(.Row, 8))
                                        'akung karyawan tidak dikenakan uang r dan selain karyawan dikenakan uang r peritem
                                        '                    If mstrKdJenisPasien = "07" Or mstrKdJenisPasien = "14" Then
                                        '                        .TextMatrix(.Row, 9) = (val(.TextMatrix(.Row, 7)) * val(.TextMatrix(.Row, 8)))
                                        '                    Else
                                        
                                        'End If
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
                                        '
                                        '                    .TextMatrix(.Row, .Col) = msubKonversiKomaTitik(TxtIsiRacikan.Text)
                                       
                                        
                                        
                                        'chandra 11 03 2014
                                        ' proses fifo untuk racikan
                                        

                                        '    fgRacikan.TextMatrix(fgRacikan.Row, 5) = Format(fgRacikan.TextMatrix(fgRacikan.Row, 5), "##.##")
                                        blt = 0
                                        blt = Round(riilnya, 1)

                                        If blt < riilnya Then blt = blt + 1
                                        

                                        Dim strKdBrg       As String

                                        Dim strKdAsal      As String

                                        Dim intRowTemp     As Integer

                                        Dim dblJmlTerkecil As Double

                                        Dim dblSelisih     As Double

                                        With fgRacikan
                                                '.TextMatrix(.Row, 5) = 0
                        
                                                'add for FIFO validasi jika terjadi edit jml stok, hapus otomatis
                                                If bolStatusFIFO = True Then
                                                        If Trim(.TextMatrix(.Row, 10)) <> "" Then
                                                                i = .Rows - 1

                                                                strKdBrg = .TextMatrix(.Row, 0)
                                                                strKdAsal = .TextMatrix(.Row, 12)

                                                                If (.TextMatrix(.Row, 6) = "") Then
                                                                        .TextMatrix(.Row, 6) = "0"
                                                                End If

                                                                If sp_StokRealRuangan(.TextMatrix(.Row, 0), .TextMatrix(.Row, 12), .TextMatrix(.Row, 14), .TextMatrix(.Row, 6), "C") = False Then Exit Sub

                                                                Do While i <> 1

                                                                        '.TextMatrix(i, 0) = .TextMatrix(.Row, 0)
                                                                        '.TextMatrix(i, 12) = .TextMatrix(.Row, 12)
                                                                        If .TextMatrix(i, 0) <> "" Then
                                                                                If (strKdBrg = .TextMatrix(i, 0)) And (strKdAsal = .TextMatrix(i, 12)) Then
                                                                                        .Row = i

                                                                                        If .CellBackColor = vbRed Then
                                                                                                Call subHapusDataGridRacikan
                                                                                                .Row = i - 1
                                                                                        End If
                                                                                End If
                                                                        End If

                                                                        i = i - 1
                                                                Loop

                                                                For i = 1 To .Rows - 1

                                                                        If (strKdBrg = .TextMatrix(i, 0)) And (strKdAsal = .TextMatrix(i, 12)) Then
                                                                                .Row = i

                                                                                Exit For

                                                                        End If

                                                                Next i

                                                        End If

                                                        .SetFocus
                                                        intRowTemp = 0
                                                End If
                                        .SetFocus
                                        .Col = 9
                                        .TextMatrix(.Row, 5) = msubKonversiKomaTitik(TxtIsiRacikan.Text)
                                        .TextMatrix(.Row, 6) = riilnya
                                        .TextMatrix(.Row, 6) = msubKonversiKomaTitik(.TextMatrix(.Row, 6))
                                        blt = 0
                                        '                    blt = Round(riilnya, 1)
                                        blt = riilnya
                                        '                    If blt < riilnya Then
                                        '                        blt = blt + 1
                                        '                    End If
                                        .TextMatrix(.Row, 4) = 0
                    
                                        .TextMatrix(.Row, 7) = CStr(blt)
                                        .TextMatrix(.Row, 9) = (CDbl(subcurTarifServiceRacikan) * CDbl(subintJmlServiceRacikan)) + (CDbl(.TextMatrix(.Row, 7)) * CDbl(.TextMatrix(.Row, 8)))
                        
                                        .TextMatrix(.Row, 15) = subintJmlServiceRacikan
                                        .TextMatrix(.Row, 16) = subcurTarifServiceRacikan
                    
                                        .TextMatrix(.Row, 9) = FormatPembulatan(.TextMatrix(.Row, 9), mstrKdInstalasiLogin)
                                        
                                                'pengambilan jumlah terkecil
                                                strSQL = "Select JmlTerkecil From MasterBarang Where KdBarang = '" & fgRacikan.TextMatrix(.Row, 0) & "'"
                                                Call msubRecFO(rs, strSQL)
                                                dblJmlTerkecil = IIf(rs.EOF, 1, rs(0).Value)

                                                'add for FIFO jika jml yg diinput melebihi stok penerimaan, mk otomatis muncul di row selanjutnya
                                                If bolStatusFIFO = True Then
                                                        Set dbRst = Nothing
                                                        Call msubRecFO(dbRst, "select JmlStok as stok from stokruanganfifo where KdRuangan='" & mstrKdRuangan & "' and KdBarang= '" & .TextMatrix(.Row, 0) & "' and KdAsal='" & .TextMatrix(.Row, 12) & "' and NoTerima='" & .TextMatrix(.Row, 14) & "'")

                                                        '                        Call msubRecFO(dbRst, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & .TextMatrix(.Row, 2) & "','" & .TextMatrix(.Row, 12) & "') as stok")
                                                        If .TextMatrix(.Row, 13) = "S" Then
                                                                dblSelisih = dbRst(0) - CDbl(blt)
                                                        Else
                                                                dblSelisih = (dbRst(0) * dblJmlTerkecil) - CDbl(blt)
                                                        End If

                                                        If dblSelisih < 0 Then
                                                                If .TextMatrix(.Row, 13) = "S" Then
                                                                        blt = dbRst(0)
                                                                Else
                                                                        blt = dbRst(0) * dblJmlTerkecil
                                                                End If

                                                                fgRacikan.TextMatrix(fgRacikan.Row, 6) = CStr(dbRst(0))
                                                                fgRacikan.TextMatrix(fgRacikan.Row, 7) = CStr(dbRst(0))
                                                                'fgRacikan.TextMatrix(fgRacikan.Row, 5) = "0"
                                                                riilnya = riilnya - dbRst(0)
                                                                .TextMatrix(.Row, 6) = msubKonversiKomaTitik(dbRst(0))
                                                        Else
                                                                Set dbRst = Nothing
                                                                '                            Call msubRecFO(dbRst, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & .TextMatrix(.Row, 2) & "','" & .TextMatrix(.Row, 29) & "') ")
                                                                'Call msubRecFO(dbRst, "select JmlStok as stok from stokruanganfifo where KdRuangan='" & mstrKdRuangan & "' and KdBarang= '" & .TextMatrix(.Row, 2) & "' and KdAsal='" & .TextMatrix(.Row, 12) & "' and NoTerima='" & .TextMatrix(.Row, 29) & "'")
                                                                Call msubRecFO(dbRst, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & .TextMatrix(.Row, 0) & "','" & .TextMatrix(.Row, 12) & "') as stok")
                                                                .TextMatrix(.Row, 6) = (blt)
                                                        End If

                                                        If .TextMatrix(.Row, 13) = "S" Then
                                                                .TextMatrix(.Row, 10) = dbRst("stok")
                                                        Else
                                                                .TextMatrix(.Row, 10) = dbRst("stok") * dblJmlTerkecil
                                                        End If

                                                        strSQL = "SELECT dbo.FB_TakeHargaNettoOA('" & mstrKdPenjaminPasien & "','" & mstrKdJenisPasien & "','" & .TextMatrix(.Row, 0) & "','" & .TextMatrix(.Row, 12) & "','" & .TextMatrix(.Row, 13) & "', '" & mstrKdRuangan & "','" & .TextMatrix(.Row, 14) & "') AS HargaBarang"
                                                        Call msubRecFO(rsB, strSQL)

                                                        If rsB.EOF = True Then .TextMatrix(.Row, 8) = 0 Else .TextMatrix(.Row, 8) = rsB(0).Value

                                                        strSQL = ""
                                                        Set rs = Nothing
                                                        subcurHargaSatuan = 0

                                                        strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "','" & .TextMatrix(.Row, 12) & "', " & msubKonversiKomaTitik(CStr(.TextMatrix(.Row, 8))) & ")  as HargaSatuan"
                                                        Call msubRecFO(rs, strSQL)

                                                        If rs.EOF = True Then .TextMatrix(.Row, 8) = 0 Else .TextMatrix(.Row, 8) = rs(0).Value

                                                        .TextMatrix(.Row, 9) = FormatPembulatan(CDbl(.TextMatrix(.Row, 6)) * CDbl(.TextMatrix(.Row, 8)), mstrKdInstalasiLogin)
                                                End If

                                                'If .TextMatrix(.Row, 13) = "S" Then
                                                If sp_StokRealRuangan(.TextMatrix(.Row, 0), .TextMatrix(.Row, 12), .TextMatrix(.Row, 14), .TextMatrix(.Row, 6), "M") = False Then Exit Sub
                                                'end If

                                                ' pengambilan no terima berikutnya
                                                '
                                                If bolStatusFIFO = True Then
                                                        If dblSelisih < 0 Then

                                                                With fgRacikan
                                                                        strSQL = "select NoTerima As NoFIFO,JmlStok from V_StokRuanganFIFO where KdBarang='" & .TextMatrix(.Row, 0) & "' and KdAsal='" & .TextMatrix(.Row, 12) & "' and NoTerima<>'" & .TextMatrix(.Row, 14) & "' and KdRuangan='" & mstrKdRuangan & "' and JmlStok<>0 order by TglTerima asc"
                                                                        Set dbRst = Nothing
                                                                        Call msubRecFO(dbRst, strSQL)

                                                                        If dbRst.EOF = False Then
                                                                                dbRst.MoveFirst

                                                                                For i = 1 To dbRst.RecordCount

                                                                                        .Rows = .Rows + 1

                                                                                        intRowTemp = .Row

                                                                                        If .TextMatrix(.Rows - 2, 2) = "" Then
                                                                                                .Row = .Rows - 2
                                                                                        Else
                                                                                                .Row = .Rows - 1
                                                                                        End If

                                                                                        Dim j As Integer

                                                                                        For j = 0 To .Cols - 1
                                                                                                .Col = j
                                                                                                .CellBackColor = vbRed
                                                                                                .CellForeColor = vbWhite
                                                                                        Next j

                                                                                        .Row = intRowTemp
                                                                                        intRowTemp = 0

                                                                                        If .TextMatrix(.Rows - 2, 2) = "" Then
                                                                                                intRowTemp = .Rows - 2
                                                                                        Else
                                                                                                intRowTemp = .Rows - 1
                                                                                        End If

                                                                                        curHutangPenjamin = 0
                                                                                        curTanggunganRS = 0

                                                                                        For j = 0 To 16
                                                                                                .TextMatrix(intRowTemp, j) = .TextMatrix(.Row, j)
                                                                                        Next j
                                                             
                                                                                        If (dbRst("JmlStok") < Abs(dblSelisih)) Then
                                                                                                fgRacikan.TextMatrix(intRowTemp, 6) = dbRst("JmlStok")
                                                                                                fgRacikan.TextMatrix(intRowTemp, 7) = dbRst("JmlStok")
                                                                                                'fgRacikan.TextMatrix(intRowTemp, 5) = "0"
                                                                                                riilnya = riilnya - dbRst("JmlStok")
                                                                                                .TextMatrix(intRowTemp, 6) = dbRst("JmlStok")
                                                                                        Else
                                                                                                ' riilnya = riilnya - Abs(dblSelisih)
                                                                                                fgRacikan.TextMatrix(intRowTemp, 6) = Abs(dblSelisih)
                                                                                                fgRacikan.TextMatrix(intRowTemp, 7) = riilnya
                                                                                                'fgRacikan.TextMatrix(intRowTemp, 5) = "0"
                                                                    
                                                                                                .TextMatrix(intRowTemp, 6) = Abs(dblSelisih)
                                                                                        End If

                                                                                        strNoTerima = dbRst("NoFIFO")
                                                                                        .TextMatrix(intRowTemp, 14) = strNoTerima
                                                                                        'cccc
                                                                                        strSQL = ""
                                                                                        Set rsB = Nothing
                                                                                        '-
                                                                                        strSQL = "SELECT dbo.FB_TakeHargaNettoOA('" & mstrKdPenjaminPasien & "','" & mstrKdJenisPasien & "','" & .TextMatrix(intRowTemp, 0) & "','" & .TextMatrix(intRowTemp, 12) & "','" & .TextMatrix(intRowTemp, 13) & "', '" & mstrKdRuangan & "','" & .TextMatrix(intRowTemp, 14) & "') AS HargaBarang"
                                                                                        Call msubRecFO(rsB, strSQL)

                                                                                        If rsB.EOF = True Then .TextMatrix(intRowTemp, 8) = 0 Else .TextMatrix(intRowTemp, 8) = rsB(0).Value

                                                                                        strSQL = ""
                                                                                        Set rs = Nothing
                                                                                        subcurHargaSatuan = 0

                                                                                        strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "','" & .TextMatrix(intRowTemp, 12) & "', " & msubKonversiKomaTitik(CStr(.TextMatrix(intRowTemp, 8))) & ")  as HargaSatuan"
                                                                                        Call msubRecFO(rs, strSQL)

                                                                                        'salah
                                                                                        If rs.EOF = True Then subcurHargaSatuan = 0 Else .TextMatrix(intRowTemp, 8) = rs(0).Value

                                                                                        If .TextMatrix(intRowTemp, 13) = "S" Then
                                                                                                dblSelisih = Abs(dblSelisih) - CDbl(dbRst("JmlStok"))
                                                                                        Else
                                                                                                dblSelisih = Abs(dblSelisih) - CDbl(dbRst("JmlStok") * dblJmlTerkecil)
                                                                                        End If

                                                                                        '   If dblSelisih >= 0 Then
                                                                                        If .TextMatrix(intRowTemp, 13) = "S" Then
                                                                                                .TextMatrix(intRowTemp, 10) = dbRst("JmlStok")
                                                                                        Else
                                                                                                .TextMatrix(intRowTemp, 10) = dbRst("JmlStok") * dblJmlTerkecil
                                                                                        End If

                                                                                        '  End If

                                                                                        'If .TextMatrix(intRowTemp, 13) = "S" Then
                                                                                        If sp_StokRealRuangan(.TextMatrix(intRowTemp, 0), .TextMatrix(intRowTemp, 12), .TextMatrix(intRowTemp, 14), .TextMatrix(intRowTemp, 6), "M") = False Then Exit Sub
                                                                                        'End If
                                                                                        .TextMatrix(intRowTemp, 7) = msubKonversiKomaTitik(.TextMatrix(intRowTemp, 7))
                                                                                        .TextMatrix(intRowTemp, 9) = (CDbl(subcurTarifServiceRacikan) * CDbl(subintJmlServiceRacikan)) + (CDbl(.TextMatrix(intRowTemp, 6)) * CDbl(.TextMatrix(intRowTemp, 8)))
                                                                                        .TextMatrix(intRowTemp, 15) = subintJmlServiceRacikan
                                                                                        .TextMatrix(intRowTemp, 16) = subcurTarifServiceRacikan
                                                                                        .TextMatrix(intRowTemp, 9) = FormatPembulatan(.TextMatrix(intRowTemp, 9), mstrKdInstalasiLogin)

                                                                                        If dblSelisih <= 0 Then Exit For

                                                                                        dbRst.MoveNext
                                                                                Next i

                                                                        End If

                                                                End With

                                                        End If
                                                End If

                                                .TextMatrix(.Row, 7) = msubKonversiKomaTitik(.TextMatrix(fgRacikan.Row, 7))
                                                .TextMatrix(.Row, 9) = (CDbl(subcurTarifServiceRacikan) * CDbl(subintJmlServiceRacikan)) + (CDbl(.TextMatrix(.Row, 6)) * CDbl(.TextMatrix(.Row, 8)))
                                                .TextMatrix(.Row, 15) = subintJmlServiceRacikan
                                                .TextMatrix(.Row, 16) = subcurTarifServiceRacikan
                                                .TextMatrix(.Row, 9) = FormatPembulatan(.TextMatrix(.Row, 9), mstrKdInstalasiLogin)
                                                .SetFocus
                                                .Col = 9
                                        End With
                    
                                Case 6 'Jumlah Real

                                        'konvert koma col jumlah
                                        '                    .TextMatrix(.Row, .Col) = msubKonversiKomaTitik(TxtIsiRacikan.Text)
                                        If txtJumlahObatRacik.Text = "" Then
                                                MsgBox "Jumlah Obat Racik Harus Di isi", vbCritical, "Informasi"
    
                                                txtJumlahObatRacik.SetFocus

                                                Exit Sub

                                        End If
                    
                                        .TextMatrix(.Row, .Col) = TxtIsiRacikan.Text
                                        TxtIsiRacikan.Visible = False
                    
                                        If Val(TxtIsiRacikan.Text) = 0 Then TxtIsiRacikan.Text = 0
                                   
                                        If .TextMatrix(.Row, .Col) <> 0 Then
                                                riilnya = Val(.TextMatrix(.Row, 6))
                                        Else
                                                MsgBox "GUNAKAN TITIK!", vbCritical, "Peringatan"

                                                Exit Sub

                                        End If

                                        '                    .TextMatrix(.Row, 6) = msubKonversiKomaTitik(.TextMatrix(.Row, 6))
                                        .TextMatrix(.Row, 6) = .TextMatrix(.Row, 6)
                                        blt = 0
                                        '                    blt = Round(riilnya, 1)
                                        blt = riilnya
                                        '                    If blt <= riilnya Then
                                        '                        blt = blt + 1
                                        '                    End If
                                        .TextMatrix(.Row, 4) = 0
                                        .TextMatrix(.Row, 5) = 0
                 
                                        .TextMatrix(.Row, 7) = msubKonversiKomaTitik(CStr(blt))
                                        '                    .TextMatrix(.Row, 7) = .TextMatrix(.Row, 6)
                                        .TextMatrix(.Row, 9) = (CDbl(subcurTarifServiceRacikan) * CDbl(subintJmlServiceRacikan)) + (CDbl(.TextMatrix(.Row, 7)) * CDbl(.TextMatrix(.Row, 8)))
                                        '
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
                    
                                        ''                    .TextMatrix(.Row, .Col) = msubKonversiKomaTitik(TxtIsiRacikan.Text)
                    
                                        .TextMatrix(.Row, 9) = FormatPembulatan(.TextMatrix(.Row, 9), mstrKdInstalasiLogin)
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

                                        .TextMatrix(.Row, 9) = FormatPembulatan(.TextMatrix(.Row, 9), mstrKdInstalasiLogin)
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

End Sub

Private Sub txtJmlTerima_KeyPress(KeyAscii As Integer)

        If KeyAscii = 13 Then cmdSimpanTerimaBarang.SetFocus
        If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtJumlahObatRacik_KeyDown(KeyCode As Integer, Shift As Integer)

        If KeyCode = 13 Then

                With fgRacikan
                        .Col = 3
                        .SetFocus
                End With

        End If

End Sub

Private Sub txtJumlahObatRacik_KeyPress(KeyAscii As Integer)
        Call SetKeyPressToNumber(KeyAscii)

        If (KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = Asc(".") Or KeyAscii = vbKeySpace) Then Exit Sub
  
        If KeyAscii = 13 Then
        
                If txtJumlahObatRacik.Text = "0" Or txtJumlahObatRacik.Text = "" Or Val(txtJumlahObatRacik.Text) = 0 Then
                        txtJumlahObatRacik.Text = "1"
                End If
          
        End If
  
End Sub

Private Sub txtJumlahObatRacik_LostFocus()

        If txtJumlahObatRacik.Text <> "" Then txtJumlahObatRacik.Enabled = False
End Sub

Private Sub txtNoResep_KeyPress(KeyAscii As Integer)

        If KeyAscii = 13 Then dtpTglResep.SetFocus
End Sub

'untuk meload data dokter di grid
Public Sub subLoadDokter()

        On Error GoTo Errload

        strSQL = "SELECT NamaDokter AS [Nama Dokter],JK,Jabatan,KodeDokter  FROM V_DaftarDokter " & mstrFilterDokter
        Call msubRecFO(rs, strSQL)

        With dgDokter
                Set .DataSource = rs
                .Columns(0).Width = 3500
                .Columns(1).Width = 400
                .Columns(2).Width = 1600
                .Columns(3).Width = 0
        End With

        dgDokter.Left = 5760
        dgDokter.Top = 1920

        Exit Sub

Errload:
        Call msubPesanError
End Sub

Private Function f_HitungTotal() As Currency

        On Error GoTo Errload

        Dim i As Integer

        f_HitungTotal = 0

        For i = 1 To fgData.Rows - 2
                f_HitungTotal = f_HitungTotal + fgData.TextMatrix(i, 11)
        Next i

        Exit Function

Errload:
        Call msubPesanError
End Function

Private Function sp_ResepObat() As Boolean

        On Error GoTo Errload

        sp_ResepObat = True
        Set dbcmd = New ADODB.Command

        With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("NoResep", adVarChar, adParamInput, 15, txtNoResep.Text)
                .Parameters.Append .CreateParameter("TglResep", adDate, adParamInput, , Format(dtpTglResep.Value, "yyyy/MM/dd"))
                .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, IIf(txtKdDokter.Text = "", Null, txtKdDokter.Text))
                .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, IIf(mstrKdRuanganPasien = "", Null, mstrKdRuanganPasien))
                .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
                .Parameters.Append .CreateParameter("ResepBebas", adChar, adParamInput, 1, "T")

                .ActiveConnection = dbConn
                .CommandText = "dbo.Add_ResepObat"
                .CommandType = adCmdStoredProc
                .Execute

                If .Parameters("return_value") <> 0 Then
                        MsgBox "Ada kesalahan dalam penyimpanan data" & .Parameters("return_value"), vbCritical, "Validasi"
                        sp_ResepObat = False

                End If

        End With

        Exit Function

Errload:
        Call msubPesanError
        sp_ResepObat = False
End Function

Public Function sp_PenerimaanSementara(f_Tanggal As Date, _
                                       f_KdBarang As String, _
                                       f_KdAsal As String, _
                                       f_JmlBarang As Double, _
                                       f_status As String) As Boolean

        On Error GoTo Errload

        Dim i As Integer

        sp_PenerimaanSementara = True
        Set dbcmd = New ADODB.Command

        With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("TglTerima", adDate, adParamInput, , Format(f_Tanggal, "yyyy/MM/dd HH:mm:ss"))
                .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
                .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
                .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
                .Parameters.Append .CreateParameter("JmlTerima", adInteger, adParamInput, , f_JmlBarang)
                .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

                .ActiveConnection = dbConn
                .CommandText = "dbo.add_PenerimaanBarangApotikTemp"
                .CommandType = adCmdStoredProc
                .Execute

                If .Parameters("return_value").Value <> 0 Then
                        MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "validasi"
                        sp_PenerimaanSementara = False
                Else
                        Call Add_HistoryLoginActivity("add_PenerimaanBarangApotikTemp")
                End If

        End With

        Set dbcmd = Nothing

        Exit Function

Errload:
        sp_PenerimaanSementara = False
        Call msubPesanError("sp_PenerimaanSementara")
End Function

Private Function sp_PemakaianObatAlkesResep(f_KdBarang As String, _
                                            f_KdAsal As String, _
                                            f_Satuan As String, _
                                            f_Jumlah As Double, _
                                            f_HargaSebelumTarifService As Currency, _
                                            f_KdJenisObat As String, _
                                            f_JumlahServise As Integer, _
                                            f_TarifService As Currency, _
                                            f_Rke As Integer, _
                                            f_StatusStok As String, _
                                            f_KdPelayananUsed As String, _
                                            f_KdStatusHasil As String, _
                                            f_JmlExpose As String, _
                                            f_KdStatusKontras As String, _
                                            f_idPenanggungjawab As String, _
                                            f_Keterangan As String, _
                                            f_NoTerima As String, _
                                            f_TglPelayanan As Date) As Boolean

        On Error GoTo Errload

        Dim i As Integer

        sp_PemakaianObatAlkesResep = True
        Set dbcmd = New ADODB.Command

        With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
                .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
                .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
                .Parameters.Append .CreateParameter("Satuan", adChar, adParamInput, 1, f_Satuan)
                .Parameters.Append .CreateParameter("JmlBrg", adDouble, adParamInput, , CDbl(f_Jumlah))
                .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
                .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
                .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, mstrKdKelas)
                .Parameters.Append .CreateParameter("HargaSatuan", adCurrency, adParamInput, , f_HargaSebelumTarifService)
                '        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(dtpTglPelayanan.Value, "yyyy/MM/dd HH:mm:ss"))
                .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(f_TglPelayanan, "yyyy/MM/dd HH:mm:ss"))
                .Parameters.Append .CreateParameter("NoLabRad", adChar, adParamInput, 10, IIf(mstrNoRad = "", mstrNoLab, mstrNoRad))
                
                
                .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, IIf(txtKdDokter.Text = "", strIDPegawaiAktif, txtKdDokter.Text))
                .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
                .Parameters.Append .CreateParameter("IdPegawai2", adChar, adParamInput, 10, Null)
                .Parameters.Append .CreateParameter("KdJenisObat", adChar, adParamInput, 2, IIf(f_KdJenisObat = "", Null, f_KdJenisObat))
                .Parameters.Append .CreateParameter("JmlService", adInteger, adParamInput, , f_JumlahServise)
                .Parameters.Append .CreateParameter("TarifService", adCurrency, adParamInput, , f_TarifService)
                Dim f_resep As String
                If (txtNoResep.Text = "") Then
                    f_resep = ""
                Else
                    f_resep = txtNoResep.Text
                End If
                
                
                .Parameters.Append .CreateParameter("NoResep", adVarChar, adParamInput, 15, IIf(chkNoResep.Value = vbChecked, IIf(f_resep = "", Null, f_resep), Null))
                .Parameters.Append .CreateParameter("Rke", adTinyInt, adParamInput, , IIf(Len(Trim(f_Rke)) = 0, Null, f_Rke))
                .Parameters.Append .CreateParameter("StatusStok", adChar, adParamInput, 1, f_StatusStok)
                .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, StrKdRP)
                .Parameters.Append .CreateParameter("KdPelayananRSUsed", adChar, adParamInput, 6, IIf(f_KdPelayananUsed = "", Null, f_KdPelayananUsed))

                '.Parameters.Append .CreateParameter("KdPelayananRSUsed", adChar, adParamInput, 6, IIf(f_KdPelayananUsed = "", Null, f_KdPelayananUsed))
                
                
                .Parameters.Append .CreateParameter("KdStatusHasil", adChar, adParamInput, 2, IIf(f_KdStatusHasil = "", Null, f_KdStatusHasil))
                .Parameters.Append .CreateParameter("JmlExpose", adInteger, adParamInput, , IIf(f_JmlExpose = "", Null, f_JmlExpose))
                .Parameters.Append .CreateParameter("KdStatusKontras", adChar, adParamInput, 2, IIf(f_KdStatusKontras = "", Null, f_KdStatusKontras))
                .Parameters.Append .CreateParameter("IdPenanggungjawab", adChar, adParamInput, 10, IIf(f_idPenanggungjawab = "", Null, f_idPenanggungjawab))
                .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 500, IIf(f_Keterangan = "", Null, f_Keterangan))
                .Parameters.Append .CreateParameter("NoTerima", adChar, adParamInput, 10, f_NoTerima)

                .ActiveConnection = dbConn
                .CommandText = "dbo.Add_PemakaianObatAlkesResepNew"
                .CommandType = adCmdStoredProc
                .Execute

                If .Parameters("return_value") <> 0 Then
                        MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
                        sp_PemakaianObatAlkesResep = False

                End If

        End With

        Call deleteADOCommandParameters(dbcmd)

        Exit Function

Errload:
        ' add by onede
        'untuk simpan ulang jika error(Time expired)
        Call deleteADOCommandParameters(dbcmd)

 '       For i = 1 To fgData.Rows - 2
'
 '               If fgData.TextMatrix(i, 2) <> "" Then
  '                      strSQL = "SELECT  NoPendaftaran, KdRuangan, KdBarang, KdAsal, TglPelayanan FROM  PemakaianAlkes" & " WHERE NoPendaftaran = '" & mstrNoPen & "'  AND KdRuangan ='" & mstrKdRuangan & "' AND KdBarang ='" & fgData.TextMatrix(i, 2) & "'  AND KdAsal ='" & fgData.TextMatrix(i, 12) & "'" & "and day(TglPelayanan)=day('" & Format(dtpTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(dtpTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(dtpTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "')"
   '                     Call msubRecFO(rs, strSQL)
'
 '                       If rs.EOF = False Then fgData.RemoveItem i
  '              End If

   '     Next i

    '    MsgBox "Waktu Penyimpanan Habis..Tekan kembali tombol simpan untuk menyimpan barang yg belum tersimpan!!!", vbExclamation, "Validasi"
        Call msubPesanError
        sp_PemakaianObatAlkesResep = False
End Function

Private Function sp_EtiketResep(f_KdBarang As String, _
                                f_KdAsal As String, _
                                f_KdJenisObat As String, _
                                f_Signa As String, _
                                f_KdSatuanEtiket As String, _
                                f_KdWaktuEtiket As String, _
                                f_ResepKe As Integer) As Boolean

        On Error GoTo Errload

        sp_EtiketResep = True
        Set dbcmd = New ADODB.Command

        With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("NoResep", adVarChar, adParamInput, 15, txtNoResep.Text)
                .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
                .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
                .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
                .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(dtpTglPelayanan.Value, "yyyy/MM/dd HH:mm:ss"))
                .Parameters.Append .CreateParameter("KdJenisObat", adChar, adParamInput, 2, IIf(f_KdJenisObat = "", Null, f_KdJenisObat))
                .Parameters.Append .CreateParameter("Signa", adVarChar, adParamInput, 7, f_Signa) 'allow null
                .Parameters.Append .CreateParameter("KdSatuanEtiket", adChar, adParamInput, 2, IIf(Len(Trim(f_KdSatuanEtiket)) = 0, Null, f_KdSatuanEtiket)) 'allow null
                .Parameters.Append .CreateParameter("KdWaktuEtiket", adChar, adParamInput, 2, IIf(Len(Trim(f_KdWaktuEtiket)) = 0, Null, f_KdWaktuEtiket)) 'allow null
                .Parameters.Append .CreateParameter("ResepKe", adTinyInt, adParamInput, , f_ResepKe)

                .ActiveConnection = dbConn
                .CommandText = "dbo.Add_EtiketResep"
                .CommandType = adCmdStoredProc
                .Execute

                If .Parameters("return_value").Value <> 0 Then
                        MsgBox "Ada kesalahan dalam penyimpanan data etiket resep", vbCritical, "Validasi"
                        sp_EtiketResep = False
                Else
                        Call Add_HistoryLoginActivity("Add_EtiketResep")
                End If

        End With

        Exit Function

Errload:
        sp_EtiketResep = False
        Call msubPesanError
End Function

Private Sub txtNoResep_LostFocus()

        On Error GoTo Errload

        If Len(Trim((txtNoResep.Text))) = 0 Then Exit Sub
        strSQL = "SELECT NoResep FROM PemakaianAlkes WHERE (NoResep = '" & txtNoResep.Text & "') AND Year(TglPelayanan) = '" & Year(dtpTglPelayanan.Value) & "'"
        Call msubRecFO(rs, strSQL)

        If rs.EOF = False Then
                MsgBox "No Resep sudah terpakai, Ganti No Resep", vbExclamation, "Validasi"
                txtNoResep.Text = ""
                txtNoResep.SetFocus
                Call subLoadDataResep(txtNoResep.Text)
                Call subHitungTotal

        End If

        txtNoResep.Text = StrConv(txtNoResep.Text, vbUpperCase)

        Exit Sub

Errload:
        Call msubPesanError
End Sub

Private Function sp_TempDetailApotikJual(f_HargaSatuan As Currency, _
                                         f_KdBarang As String, _
                                         f_KdAsal As String) As Boolean
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
                End If

        End With

End Function

Private Sub subHitungTotal()

        On Error GoTo Errload

        Dim i As Integer

        If fgData.TextMatrix(fgData.Row - 1, 11) = "" Then Exit Sub
        txtTotalBiaya.Text = 0
        txtHutangPenjamin.Text = 0
        txtTanggunganRS.Text = 0
        txtHarusDibayar.Text = 0
        txtTotalDiscount.Text = 0

        With fgData

                For i = 1 To IIf(fgData.TextMatrix(fgData.Rows - 1, 2) = "", fgData.Rows - 2, fgData.Rows - 1)

                        If .TextMatrix(i, 22) = "" Then .TextMatrix(i, 22) = 0
                        If .TextMatrix(i, 19) = "" Then .TextMatrix(i, 19) = 0
                        If .TextMatrix(i, 20) = "" Then .TextMatrix(i, 20) = 0
                        If .TextMatrix(i, 21) = "" Then .TextMatrix(i, 21) = 0
            
                        txtTotalBiaya.Text = txtTotalBiaya.Text + CDbl(.TextMatrix(i, 11))
                        txtHutangPenjamin.Text = txtHutangPenjamin.Text + CDbl(.TextMatrix(i, 19))
                        txtTanggunganRS.Text = txtTanggunganRS.Text + CDbl(.TextMatrix(i, 20))
                        txtTotalDiscount.Text = txtTotalDiscount.Text + CDbl(.TextMatrix(i, 21))
                        txtHarusDibayar.Text = txtHarusDibayar.Text + CDbl(.TextMatrix(i, 22))
                Next i

        End With

        txtTotalBiaya.Text = IIf(Val(txtTotalBiaya.Text) = 0, 0, FormatPembulatan(CDbl(txtTotalBiaya.Text), mstrKdInstalasiLogin))
        txtHutangPenjamin.Text = IIf(Val(txtHutangPenjamin.Text) = 0, 0, FormatPembulatan(CDbl(txtHutangPenjamin.Text), mstrKdInstalasiLogin))
        txtTanggunganRS.Text = IIf(Val(txtTanggunganRS.Text) = 0, 0, FormatPembulatan(CDbl(txtTanggunganRS.Text), mstrKdInstalasiLogin))
        txtHarusDibayar.Text = IIf(Val(txtHarusDibayar.Text) = 0, 0, FormatPembulatan(CDbl(txtHarusDibayar.Text), mstrKdInstalasiLogin))
        txtTotalDiscount.Text = IIf(Val(txtTotalDiscount.Text) = 0, 0, FormatPembulatan(CDbl(txtTotalDiscount.Text), mstrKdInstalasiLogin))

        subcurHarusDibayar = txtHarusDibayar.Text

        Exit Sub

Errload:
        Call msubPesanError
        '    Resume 0
End Sub

Private Sub subLoadDataResep(f_NoResep As String)

        On Error GoTo Errload

        Dim i                 As Integer

        Dim curHutangPenjamin As Currency

        Dim curHarusDibayar   As Currency

        Dim curTanggunganRS   As Currency

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
                strSQL = "SELECT HargaSatuan, JmlHutangPenjamin, JmlTanggunganRS" & " FROM TempDetailApotikJual" & " WHERE (NoTemporary = '" & Trim(txtNoTemporary.Text) & "') AND (KdBarang = '" & dbRst("KdBarang") & "') AND (KdAsal = '" & dbRst("KdAsal") & "')"
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
                        .TextMatrix(.Rows - 1, 7) = IIf(Val(.TextMatrix(.Rows - 1, 7)) = 0, 0, FormatPembulatan(CDbl(.TextMatrix(.Rows - 1, 5)), mstrKdInstalasiLogin))
                        .TextMatrix(.Rows - 1, 8) = CDbl(0) 'discount
                        '.TextMatrix(.Rows - 1, 8) = IIf(val(.TextMatrix(.Rows - 1, 6)) = 0, 0, FormatPembulatan(CDbl(.TextMatrix(.Rows - 1, 6)), mstrKdInstalasiLogin))
                        .TextMatrix(.Rows - 1, 9) = CDbl(dbRst("JmlStok") + dbRst("JmlBarang"))
                        .TextMatrix(.Rows - 1, 10) = CDbl(dbRst("JmlBarang"))

                        .TextMatrix(.Rows - 1, 11) = ((dbRst("TarifService") * dbRst("JmlService")) + (CDbl(dbRst("HargaSatuan")) * CDbl(.TextMatrix(.Rows - 1, 10))))
                        .TextMatrix(.Rows - 1, 11) = IIf(Val(.TextMatrix(.Rows - 1, 11)) = 0, 0, FormatPembulatan(CDbl(.TextMatrix(.Rows - 1, 11)), mstrKdInstalasiLogin))

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

                        curHarusDibayar = CDbl(.TextMatrix(.Rows - 1, 11)) - (CDbl(.TextMatrix(.Rows - 1, 21)) + CDbl(.TextMatrix(.Rows - 1, 19)) + CDbl(.TextMatrix(.Rows - 1, 120)))
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

        dgObatAlkes.Visible = False
        txtJenisBarang.Text = "": txtKdBarang.Text = "": txtKdAsal.Text = "": txtSatuan.Text = "": txtAsalBarang.Text = ""

        Exit Sub

Errload:
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

        txtIsi.Visible = True
        txtIsi.SelStart = Len(txtIsi.Text)
        txtIsi.SetFocus
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

        s_DcName.Visible = True
        's_DcName.SetFocus
End Sub

Private Sub subLoadCheck()

        Dim i As Integer

        chkStatusStok.Left = fgData.Left

        For i = 0 To fgData.Col - 1
                chkStatusStok.Left = chkStatusStok.Left + fgData.ColWidth(i)
        Next i

        chkStatusStok.Visible = True
        chkStatusStok.Top = fgData.Top - 7

        For i = 0 To fgData.Row - 1
                chkStatusStok.Top = chkStatusStok.Top + fgData.RowHeight(i)
        Next i

        If fgData.TopRow > 1 Then
                chkStatusStok.Top = chkStatusStok.Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
        End If

        chkStatusStok.Width = fgData.ColWidth(fgData.Col)
        chkStatusStok.Height = fgData.RowHeight(fgData.Row)
        chkStatusStok.BackColor = fgData.BackColor

        chkStatusStok.Visible = True
        chkStatusStok.SetFocus
End Sub

Private Sub txtRP_Change()

        If Len(Trim(txtRP.Text)) = 0 Then StrKdRP = "": Exit Sub
        strSQL = "Select KdRuangan from Ruangan  where NamaRuangan='" & txtRP.Text & "'"
        Call msubRecFO(rs, strSQL)
        StrKdRP = IIf(IsNull(rs.Fields(0).Value), "", rs.Fields(0).Value)
End Sub

Private Function update_DetailOrderTMOA(ByVal adoCommand As ADODB.Command, _
                                        sItem As String, _
                                        sStatus As String, noOrder As String) As Boolean

        On Error GoTo Errload

        update_DetailOrderTMOA = True

        With adoCommand
                .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
                .Parameters.Append .CreateParameter("KdItem", adVarChar, adParamInput, 9, sItem)
                .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
                .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
                .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, noOrder)
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 2, sStatus)

                .ActiveConnection = dbConn
                .CommandText = "dbo.Update_DetailOrderTMOAKhususUntukNoOrderLebihDariSatu"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("RETURN_VALUE").Value = 0) Then
                        MsgBox "Ada Kesalahan dalam Penyimpanan data", vbCritical, "Validasi"
                        Call deleteADOCommandParameters(adoCommand)
                        Set adoCommand = Nothing
                        update_DetailOrderTMOA = False

                End If

                Call deleteADOCommandParameters(adoCommand)
                Set adoCommand = Nothing
        End With

        Exit Function

Errload:
        update_DetailOrderTMOA = False
        Call msubPesanError
End Function


