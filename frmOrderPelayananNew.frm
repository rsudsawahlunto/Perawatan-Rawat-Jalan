VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOrderPelayananNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order Pelayanan Tindakan Medis dan Obat Alkes"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15945
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOrderPelayananNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   15945
   Begin VB.TextBox txtNoPakai 
      Height          =   315
      Left            =   2760
      TabIndex        =   65
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid dgDokterTM 
      Height          =   2295
      Left            =   6120
      TabIndex        =   5
      Top             =   2640
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
   Begin MSDataGridLib.DataGrid dgDokter 
      Height          =   2295
      Left            =   8520
      TabIndex        =   15
      Top             =   3360
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
   Begin MSDataGridLib.DataGrid dgObatAlkes 
      Height          =   2535
      Left            =   2400
      TabIndex        =   17
      Top             =   3600
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
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
   Begin MSDataGridLib.DataGrid dgPelayananRS 
      Height          =   2535
      Left            =   360
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   706
      TabCaption(0)   =   "Tindakan Medis"
      TabPicture(0)   =   "frmOrderPelayananNew.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtTotalBiaya"
      Tab(0).Control(1)=   "txtHutangPenjamin"
      Tab(0).Control(2)=   "txtTanggunganRS"
      Tab(0).Control(3)=   "txtHarusDibayar"
      Tab(0).Control(4)=   "txtTotalDiscount"
      Tab(0).Control(5)=   "chkDokterPelmeriksa"
      Tab(0).Control(6)=   "txtRP"
      Tab(0).Control(7)=   "chkDokterPemeriksa"
      Tab(0).Control(8)=   "txtkdAsal"
      Tab(0).Control(9)=   "txtSatuan"
      Tab(0).Control(10)=   "txtasalbarang"
      Tab(0).Control(11)=   "chkStatusStok"
      Tab(0).Control(12)=   "fgDataTM"
      Tab(0).Control(13)=   "Frame1"
      Tab(0).Control(14)=   "txtIsiTM"
      Tab(0).Control(15)=   "cbCito"
      Tab(0).Control(16)=   "txtKdDokterTM"
      Tab(0).Control(17)=   "txtIsiTMJml"
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Obat Alkes"
      TabPicture(1)   =   "frmOrderPelayananNew.frx":0CE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame8"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame8 
         Height          =   4095
         Left            =   120
         TabIndex        =   50
         Top             =   1560
         Width           =   15375
         Begin VB.TextBox txtIsi 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   330
            Left            =   3600
            TabIndex        =   62
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtNoTemporary 
            Height          =   315
            Left            =   7080
            TabIndex        =   61
            Top             =   2280
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtJenisBarang 
            Height          =   315
            Left            =   5040
            TabIndex        =   60
            Top             =   1800
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtKdDokter 
            Height          =   315
            Left            =   1560
            TabIndex        =   59
            Top             =   2280
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtKdBarang 
            Height          =   315
            Left            =   3720
            TabIndex        =   58
            Top             =   2280
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.ComboBox cbCitoOA 
            Appearance      =   0  'Flat
            Height          =   330
            ItemData        =   "frmOrderPelayananNew.frx":0D02
            Left            =   4320
            List            =   "frmOrderPelayananNew.frx":0D0C
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   0
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txtIsiJml 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   330
            Left            =   5040
            TabIndex        =   56
            Top             =   2520
            Visible         =   0   'False
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo dcKeteranganPakai2 
            Height          =   330
            Left            =   9240
            TabIndex        =   51
            Top             =   840
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcNamaPelayananRS 
            Height          =   330
            Left            =   6120
            TabIndex        =   52
            Top             =   840
            Visible         =   0   'False
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcJenisObat 
            Height          =   330
            Left            =   840
            TabIndex        =   53
            Top             =   960
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcKeteranganPakai 
            Height          =   330
            Left            =   2520
            TabIndex        =   54
            Top             =   960
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcAturanPakai 
            Height          =   330
            Left            =   4440
            TabIndex        =   55
            Top             =   840
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcNamaPelayanan 
            Height          =   330
            Left            =   2160
            TabIndex        =   63
            Top             =   1680
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin MSFlexGridLib.MSFlexGrid fgData 
            Height          =   3735
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   6588
            _Version        =   393216
            FixedCols       =   0
            BackColorSel    =   -2147483643
            FocusRect       =   2
            HighLight       =   2
            Appearance      =   0
         End
      End
      Begin VB.TextBox txtIsiTMJml 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   -70680
         TabIndex        =   49
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtKdDokterTM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   -74280
         MaxLength       =   15
         TabIndex        =   36
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ComboBox cbCito 
         Appearance      =   0  'Flat
         Height          =   330
         ItemData        =   "frmOrderPelayananNew.frx":0D19
         Left            =   -70440
         List            =   "frmOrderPelayananNew.frx":0D23
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   2760
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtIsiTM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   -72000
         TabIndex        =   34
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Order TM"
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
         Left            =   -74760
         TabIndex        =   28
         Top             =   600
         Width           =   15255
         Begin VB.TextBox txtNoPendaftaranTM 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   3960
            MaxLength       =   15
            TabIndex        =   3
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txtDokterTM 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   6000
            TabIndex        =   4
            Top             =   480
            Width           =   3615
         End
         Begin VB.TextBox txtNoCMTM 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2520
            MaxLength       =   15
            TabIndex        =   2
            Top             =   480
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker dtpTglOrderTM 
            Height          =   330
            Left            =   240
            TabIndex        =   1
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
            Format          =   119930883
            UpDown          =   -1  'True
            CurrentDate     =   37760
         End
         Begin MSDataListLib.DataCombo dcRuanganTujuanTM 
            Height          =   330
            Left            =   9960
            TabIndex        =   6
            Top             =   480
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Pendaftaran"
            Height          =   210
            Index           =   6
            Left            =   3960
            TabIndex        =   33
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ruangan Tujuan"
            Height          =   210
            Index           =   9
            Left            =   9960
            TabIndex        =   32
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Dokter Order"
            Height          =   210
            Index           =   8
            Left            =   6000
            TabIndex        =   31
            Top             =   240
            Width           =   1590
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. CM"
            Height          =   210
            Index           =   7
            Left            =   2520
            TabIndex        =   30
            Top             =   240
            Width           =   585
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl. Order"
            Height          =   210
            Index           =   5
            Left            =   240
            TabIndex        =   29
            Top             =   240
            Width           =   840
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Data Order OA"
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
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   15375
         Begin VB.TextBox txtNoResep 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2640
            MaxLength       =   15
            TabIndex        =   12
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtDokter 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   5880
            TabIndex        =   14
            Top             =   480
            Width           =   3615
         End
         Begin MSComCtl2.DTPicker dtpTglOrder 
            Height          =   330
            Left            =   360
            TabIndex        =   11
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
            Format          =   119930883
            UpDown          =   -1  'True
            CurrentDate     =   37760
         End
         Begin MSComCtl2.DTPicker dtpTglResep 
            Height          =   330
            Left            =   4320
            TabIndex        =   13
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
            Format          =   119930883
            UpDown          =   -1  'True
            CurrentDate     =   37760
         End
         Begin MSDataListLib.DataCombo dcRuanganTujuan 
            Height          =   330
            Left            =   9720
            TabIndex        =   16
            Top             =   480
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl. Order"
            Height          =   210
            Index           =   0
            Left            =   360
            TabIndex        =   27
            Top             =   240
            Width           =   840
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl. Resep"
            Height          =   210
            Index           =   1
            Left            =   4560
            TabIndex        =   26
            Top             =   240
            Width           =   870
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Resep"
            Height          =   210
            Index           =   2
            Left            =   2640
            TabIndex        =   25
            Top             =   240
            Width           =   840
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Dokter Order"
            Height          =   210
            Index           =   3
            Left            =   5880
            TabIndex        =   24
            Top             =   240
            Width           =   1590
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ruangan Tujuan"
            Height          =   210
            Index           =   4
            Left            =   9720
            TabIndex        =   23
            Top             =   240
            Width           =   1335
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgDataTM 
         Height          =   3855
         Left            =   -74760
         TabIndex        =   7
         Top             =   1680
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   6800
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   -2147483643
         FocusRect       =   2
         HighLight       =   2
         Appearance      =   0
      End
      Begin VB.CheckBox chkStatusStok 
         Height          =   495
         Left            =   -69360
         TabIndex        =   37
         Top             =   3960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtasalbarang 
         Height          =   495
         Left            =   -70680
         TabIndex        =   38
         Top             =   3960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtSatuan 
         Height          =   495
         Left            =   -72120
         TabIndex        =   39
         Top             =   3960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtkdAsal 
         Height          =   495
         Left            =   -73440
         TabIndex        =   40
         Top             =   3960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chkDokterPemeriksa 
         Height          =   495
         Left            =   -63720
         TabIndex        =   41
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtRP 
         Height          =   495
         Left            =   -65160
         TabIndex        =   42
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chkDokterPelmeriksa 
         Height          =   495
         Left            =   -66600
         TabIndex        =   43
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtTotalDiscount 
         Height          =   495
         Left            =   -68040
         TabIndex        =   44
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtHarusDibayar 
         Height          =   495
         Left            =   -69360
         TabIndex        =   45
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtTanggunganRS 
         Height          =   495
         Left            =   -70680
         TabIndex        =   46
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtHutangPenjamin 
         Height          =   495
         Left            =   -72120
         TabIndex        =   47
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtTotalBiaya 
         Height          =   495
         Left            =   -73440
         TabIndex        =   48
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   18
      Top             =   7080
      Width           =   15735
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   2640
         MaxLength       =   15
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtNoOrder 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   240
         MaxLength       =   15
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
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
         Left            =   11760
         TabIndex        =   9
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
         Left            =   13560
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   66
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
      Left            =   14040
      Picture         =   "frmOrderPelayananNew.frx":0D30
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmOrderPelayananNew.frx":1AB8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmOrderPelayananNew.frx":4479
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "frmOrderPelayananNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim subintJmlArray As Integer
Dim subcurHargaSatuan As Currency
Dim subcurTarifService As Currency
Dim subcurHarusDibayar As Currency
Dim curTanggunganRS As Currency
Dim curHutangPenjamin As Currency
Dim subintJmlService As Integer
Dim tempStatusTampil As Boolean
Dim subJenisHargaNetto As Integer
Dim KI As String

Dim subcurBiayaAdministrasi As Currency

Private Sub cbCito_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        With fgDataTM
            .TextMatrix(.Row, 2) = cbCito.Text
            .SetFocus
            .Col = 3
        End With
    End If
    If KeyCode = vbKeyEscape Then
        cbCito.Visible = False
        fgDataTM.SetFocus
        fgDataTM.Col = 2
    End If
End Sub

Private Sub cbCito_LostFocus()
    cbCito.Visible = False
End Sub

Private Sub cbCitoOA_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        With fgData
            .TextMatrix(.Row, 5) = cbCitoOA.Text
            .SetFocus
            .Col = 6
        End With
    End If
    If KeyCode = vbKeyEscape Then
        cbCito.Visible = False
        fgData.SetFocus
        fgData.Col = 5
    End If
End Sub

Private Sub cbCitoOA_LostFocus()
    cbCitoOA.Visible = False
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad
    Dim i As Integer

    Select Case SSTab1.Tab

        Case 0

            If Periksa("text", txtDokterTM, "Nama Dokter yang memesan kosong") = False Then Exit Sub

            If Periksa("datacombo", dcRuanganTujuanTM, "Ruangan tujuan harus diisi!!") = False Then Exit Sub

            If fgDataTM.TextMatrix(1, 0) = "" Then MsgBox "pelayanan yang akan dipesan harus diisi", vbExclamation, "Validasi": Exit Sub

            If sp_Order() = False Then Exit Sub

            With fgDataTM
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, 4) = "" Then GoTo lanjut_
                    If sp_DetailOrderPelayananTM(.TextMatrix(i, 4), .TextMatrix(i, 1), .TextMatrix(i, 2), _
                        .TextMatrix(i, 3)) = False Then Exit Sub
lanjut_:
                    Next i
                End With
                MsgBox "pemesanan pelayanan berhasil", vbInformation, "Informasi"

                Call subKosong
                Call subSetGrid

            Case 1

                If Periksa("text", txtDokter, "Nama Dokter yang memesan kosong") = False Then Exit Sub
                If Periksa("datacombo", dcRuanganTujuan, "Ruangan tujuan harus diisi!!") = False Then Exit Sub

                If fgData.TextMatrix(1, 2) = "" Then MsgBox "Data barang yang akan dipesan harus diisi", vbExclamation, "Validasi": Exit Sub
                For i = 1 To fgData.Rows - 1
                    If fgData.TextMatrix(i, 2) = "" Then GoTo lanjut1_
                    If fgData.TextMatrix(i, 1) = "" Then MsgBox "Cek Jenis Obat ", vbExclamation, "Validasi": fgData.SetFocus: fgData.Col = 1: Exit Sub
                    If fgData.TextMatrix(i, 10) = 0 Then MsgBox "Cek Jumlah Barang ", vbExclamation, "Validasi": fgData.SetFocus: fgData.Col = 10: Exit Sub
lanjut1_:
                Next i

                If sp_Order() = False Then Exit Sub

                With fgData
                    For i = 1 To .Rows - 1
                        If .TextMatrix(i, 2) = "" Then GoTo lanjutkan_
'                        If sp_DetailOrderPelayananOA(.TextMatrix(i, 2), .TextMatrix(i, 11), .TextMatrix(i, 4), .TextMatrix(i, 0), _
                            .TextMatrix(i, 6), .TextMatrix(i, 7), .TextMatrix(i, 9), .TextMatrix(i, 10), .TextMatrix(i, 5)) = False Then Exit Sub

                        If sp_DetailOrderPelayananOANew(.TextMatrix(i, 2), CDbl(.TextMatrix(i, 10)), .TextMatrix(i, 0), _
                            .TextMatrix(i, 35), .TextMatrix(i, 36), .TextMatrix(i, 37), .TextMatrix(i, 32), .TextMatrix(i, 30), CInt(IIf(.TextMatrix(i, 15) = "", 0, .TextMatrix(i, 15))), _
                            CCur(IIf(.TextMatrix(i, 14) = "", 0, .TextMatrix(i, 14))), "No", .TextMatrix(i, 25), .TextMatrix(i, 12), .TextMatrix(i, 6), .TextMatrix(i, 33), _
                            .TextMatrix(i, 34)) = False Then Exit Sub

lanjutkan_:
                    Next i
                End With
                    MsgBox "pemesanan obat alkes  berhasil", vbInformation, "Informasi"

                    Call subKosong
                    Call subSetGrid

            End Select

            Exit Sub
errLoad:
            msubPesanError

End Sub

Private Sub cmdTutup_Click()

    Unload Me

End Sub

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
'    If dcRuanganTujuan.Text = "" Then
'            MsgBox "Ma'af Pilih Terlebih Dahulu Ruang Tujuan!", vbInformation + vbOKOnly, "Info"
'            dcJenisObat.Text = ""
'            dcRuanganTujuan.Text = ""
'            dcRuanganTujuan.SetFocus
'    Exit Sub
'    End If
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
        fgData.TextMatrix(fgData.Row, 0) = fgData.TextMatrix(fgData.Row - 1, 0) + 1
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
        Call dcJenisObat_Change
        dcJenisObat.Visible = False
        fgData.Col = 3
        fgData.SetFocus

'     dcJenisObat.Visible = False
'    If Cancel = False Then
'    Dim no As Integer
'    If fgData.TextMatrix(fgData.Row, 0) = "" Then
'        no = Val(fgData.TextMatrix(fgData.Row - 1, 0)) + 1
'        fgData.TextMatrix(fgData.Row, 0) = no
'    End If
'
'    If dcJenisObat.BoundText <> "01" And dcJenisObat.BoundText <> "" Then
'        Call subSetGridRacikan
'        FraRacikan.Visible = True
'        dgObatAlkesRacikan.Visible = False
'        txtJumlahObatRacik.SetFocus
'        txtJumlahObatRacik.Text = ""
'        cmdSimpan.Enabled = False
'        cmdTutup.Enabled = False
'        fgRacikan.TextMatrix(fgRacikan.Row, 2) = fgData.TextMatrix(fgData.Row, 0) ' edit
'    Else
'        With fgData
'            .TextMatrix(.Row, 1) = dcJenisObat.Text
'            .TextMatrix(.Row, 25) = dcJenisObat.BoundText
'        End With
'    End If
'
'End If
'
'Cancel = False
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

Private Sub dcNamaPelayanan_Change()
    On Error GoTo errLoad
    fgData.TextMatrix(fgData.Row, 8) = dcNamaPelayanan.Text
    fgData.TextMatrix(fgData.Row, 9) = dcNamaPelayanan.BoundText

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcNamaPelayanan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then dcNamaPelayanan.Visible = False
End Sub

Private Sub dcNamaPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call dcNamaPelayanan_Change
        dcNamaPelayanan.Visible = False
        fgData.Col = 8
        fgData.SetFocus
    End If
End Sub

Private Sub dcNamaPelayanan_LostFocus()
    dcNamaPelayanan.Visible = False
End Sub

Private Sub dcRuanganTujuan_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 13 Then
        If dcRuanganTujuan.MatchedWithList = True Then fgData.SetFocus
        strSQL = "select KdRuangan, NamaRuangan from Ruangan WHERE (NamaRuangan LIKE '%" & dcRuanganTujuan.Text & "%') and StatusEnabled=1"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcRuanganTujuan.BoundText = rs(0).Value
        dcRuanganTujuan.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcRuanganTujuanTM_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 13 Then
        If dcRuanganTujuanTM.MatchedWithList = True Then fgDataTM.SetFocus
        strSQL = "select KdRuangan, NamaRuangan from Ruangan WHERE (NamaRuangan LIKE '%" & dcRuanganTujuanTM.Text & "%') and StatusEnabled=1"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcRuanganTujuanTM.BoundText = rs(0).Value
        dcRuanganTujuanTM.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dgDokter_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDokter
    WheelHook.WheelHook dgDokter
End Sub

Private Sub dgDokter_DblClick()
    On Error GoTo errLoad
    If dgDokter.ApproxCount = 0 Then Exit Sub
    txtDokter.Text = dgDokter.Columns("Nama Dokter")
    dgDokter.Visible = False
    txtKdDokter.Text = dgDokter.Columns("KodeDokter")

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call dgDokter_DblClick
End Sub

Private Sub dgDokter_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If dgDokter.Visible = False Then Exit Sub
        txtDokter.SetFocus
    End If
End Sub

Private Sub dgDokterTM_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDokterTM
    WheelHook.WheelHook dgDokterTM
End Sub

Private Sub dgDokterTM_DblClick()
    On Error GoTo errLoad
    If dgDokterTM.ApproxCount = 0 Then Exit Sub
    txtDokterTM.Text = dgDokterTM.Columns("Nama Dokter")
    dgDokterTM.Visible = False
    txtKdDokterTM.Text = dgDokterTM.Columns("KodeDokter")

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dgDokterTM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call dgDokterTM_DblClick
End Sub

Private Sub dgDokterTM_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If dgDokterTM.Visible = False Then Exit Sub
        txtDokterTM.SetFocus
    End If
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
    Call msubRecFO(rsB, "select dbo.TakeNoFIFO_F('" & dgObatAlkes.Columns("KdBarang") & "','" & dgObatAlkes.Columns("KdAsal") & "','" & mstrKdRuangan & "') as NoFIFO")
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
        .TextMatrix(.Row, 3) = dgObatAlkes.Columns("NamaBarang")
        .TextMatrix(.Row, 4) = dgObatAlkes.Columns("Kekuatan")
        .TextMatrix(.Row, 5) = dgObatAlkes.Columns("AsalBarang")
        .TextMatrix(.Row, 6) = dgObatAlkes.Columns("Satuan")
        '.TextMatrix(.Row, 7) = Format(dgObatAlkes.Columns("HargaBarang").Value, "#,###")
        .TextMatrix(.Row, 33) = strNoTerima
        curHargaBrg = 0
        
        strSQL = ""
        Set rsB = Nothing
        strSQL = "SELECT dbo.FB_TakeHargaNettoOA('" & mstrKdPenjaminPasien & "','" & mstrKdJenisPasien & "','" & dgObatAlkes.Columns("KdBarang") & "','" & dgObatAlkes.Columns("KdAsal") & "','" & dgObatAlkes.Columns("Satuan") & "', '" & mstrKdRuangan & "','" & .TextMatrix(.Row, 34) & "') AS HargaBarang"
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
            .TextMatrix(.Row, 7) = Format(subcurHargaSatuan, "#,###")
        End If
        .TextMatrix(.Row, 8) = (dgObatAlkes.Columns("Discount").Value / 100) * CDbl(.TextMatrix(.Row, 7))

        Call msubRecFO(rs, "select dbo.FB_TakeStokBrgMedis('" & dcRuanganTujuan.BoundText & "', '" & dgObatAlkes.Columns("KdBarang") & "','" & dgObatAlkes.Columns("KdAsal") & "','" & .TextMatrix(.Row, 34) & "') as stok")
        .TextMatrix(.Row, 9) = IIf(IsNull(rs("Stok")), 0, rs("Stok"))
        
        .TextMatrix(.Row, 12) = dgObatAlkes.Columns("KdAsal")
        .TextMatrix(.Row, 13) = dgObatAlkes.Columns("JenisBarang")
        .TextMatrix(.Row, 16) = CDbl(.TextMatrix(.Row, 7))
        .TextMatrix(.Row, 17) = curHutangPenjamin
        .TextMatrix(.Row, 18) = curTanggunganRS
        .TextMatrix(.Row, 19) = 0
        .TextMatrix(.Row, 20) = 0
        .TextMatrix(.Row, 21) = 0
        
        .TextMatrix(.Row, 23) = txtNoTemporary.Text
'        txtHargaBeli.Text = curHargaBrg 'dgObatAlkes.Columns("HargaBarang")
        .TextMatrix(.Row, 24) = curHargaBrg 'CDbl(txtHargaBeli.Text)
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
End Sub

Private Sub dgObatAlkes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call dgObatAlkes_DblClick
End Sub

Private Sub dgPelayananRS_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgPelayananRS
    WheelHook.WheelHook dgPelayananRS
End Sub

Private Sub dgPelayananRS_DblClick()
    On Error GoTo errLoad
    Dim i As Integer

    For i = 0 To fgDataTM.Rows - 1
        If dgPelayananRS.Columns("KdPelayananRS") = fgDataTM.TextMatrix(i, 4) Then
            MsgBox "Data tersebut sudah diinput", vbExclamation, "Validasi"
            dgPelayananRS.Visible = False
            fgDataTM.SetFocus: fgDataTM.Row = i
            Exit Sub
        End If
    Next i

    With fgDataTM
        .TextMatrix(.Row, 4) = dgPelayananRS.Columns("KdPelayananRS")
        .TextMatrix(.Row, 0) = dgPelayananRS.Columns("NamaPelayanan")
        .TextMatrix(.Row, 1) = "1"
        .TextMatrix(.Row, 2) = "No"

    End With

    dgPelayananRS.Visible = False

    With fgDataTM
        .SetFocus
        .Col = 1
    End With

    Exit Sub
errLoad:

End Sub

Private Sub dgPelayananRS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call dgPelayananRS_DblClick
End Sub

Private Sub dtpTglResep_Change()
    dtpTglResep.MaxDate = Now
End Sub

Private Sub dtpTglResep_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtDokter.SetFocus
End Sub

Private Sub dtpTglResep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDokter.SetFocus
End Sub

Private Sub fgData_DblClick()
    Call fgData_KeyDown(13, 0)
End Sub

Private Sub fgData_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Select Case KeyCode
        Case 13
            If fgData.Col = fgData.Cols - 1 Then
                If fgData.TextMatrix(fgData.Row, 2) <> "" Then
                    If fgData.TextMatrix(fgData.Rows - 1, 2) <> "" Then
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
                    fgData.Col = 1
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
        If .TextMatrix(.Row, 11) <> "01" Then 'jika obat racikan, pastikan jumlah service 1 untuk resep yang sama
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
        
        dbConn.Execute "DELETE FROM DetailOrderPelayananOARacikanTemp where NoRacikan = '" & fgData.TextMatrix(fgData.Row, 34) & "'"
        
        If .Rows = 2 Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
                .TextMatrix(1, 0) = 1
            Next i
        Else
            .RemoveItem .Row
        End If
        
        If .TextMatrix(.Row, 2) <> "" Then
         
            .TextMatrix(.Row, 11) = ((CDbl(.TextMatrix(.Row, 14)) * CDbl(.TextMatrix(.Row, 15))) + _
                (CDbl(.TextMatrix(.Row, 16)) * Val(.TextMatrix(.Row, 10))))
          
            curHarusDibayar = CDbl(.TextMatrix(.Row, 11)) - (CDbl(.TextMatrix(.Row, 21)) + _
                CDbl(.TextMatrix(.Row, 19)) + CDbl(.TextMatrix(.Row, 20)))
            .TextMatrix(.Row, 20) = IIf(curHarusDibayar < 0, 0, curHarusDibayar)
        End If
    End With
    Call subHitungTotal

Exit Sub
errLoad:
    Call msubPesanError
End Sub



Private Sub fgData_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    
    If Periksa("datacombo", dcRuanganTujuan, "Pilih ruangan tujuan") = False Then Exit Sub
    
    txtIsi.Text = ""
    txtIsiJml.Text = ""
    If Not (KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ Or KeyAscii = 32 Or KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace Or KeyAscii = Asc(".")) Then
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
            If KI <> "07" Then
            dcJenisObat.Visible = True
            End If
            
        Case 2 'Kode Barang
'            txtIsi.MaxLength = 9
'            Call subLoadText
'            txtIsi.Text = Chr(KeyAscii)
'            txtIsi.SelStart = Len(txtIsi.Text)

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
    
        Case 27 'Aturan Pakai
            fgData.Col = 27
            Call subLoadDataCombo(dcAturanPakai)
            
        Case 28 'Keterangan Pakai
            fgData.Col = 28
            Call subLoadDataCombo(dcKeteranganPakai)
            
'        Case 29 'Keterangan Pakai2
'            fgData.Col = 29
'            Call subLoadDataCombo(dcKeteranganPakai2)
'
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

Private Sub fgDataTM_DblClick()
    Call fgDataTM_KeyDown(13, 0)
End Sub

Private Sub fgDataTM_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Select Case KeyCode
        Case 13
            If fgDataTM.Col = fgDataTM.Cols - 1 Then
                If fgDataTM.TextMatrix(fgDataTM.Row, 4) <> "" Then
                    If fgDataTM.TextMatrix(fgDataTM.Rows - 1, 4) <> "" Then
                        fgDataTM.Rows = fgDataTM.Rows + 1
                    End If
                    fgDataTM.Row = fgDataTM.Rows - 1
                    fgDataTM.Col = 0
                Else
                    fgDataTM.Col = 0
                End If
            Else
                For i = 0 To fgDataTM.Cols - 1
                    If fgDataTM.Col = fgDataTM.Cols - 1 Then Exit For
                    fgDataTM.Col = fgDataTM.Col + 1
                    If fgDataTM.ColWidth(fgDataTM.Col) > 0 Then Exit For
                Next i
            End If

            fgDataTM.SetFocus

        Case 27
            dgPelayananRS.Visible = False

        Case vbKeyDelete
            If fgDataTM.Row = 1 Then
                For i = 0 To fgDataTM.Cols - 1
                    fgDataTM.TextMatrix(1, i) = ""
                Next i
            Else
                fgDataTM.RemoveItem fgDataTM.Row
            End If

    End Select
End Sub

Private Sub fgDataTM_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    
    If Periksa("datacombo", dcRuanganTujuanTM, "Pilih ruangan tujuan") = False Then Exit Sub

    txtIsiTM.Text = ""
    txtIsiTMJml.Text = ""
    If Not (KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ Or KeyAscii = 32 Or KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace Or KeyAscii = Asc(".")) Then
        KeyAscii = 0
        Exit Sub
    End If

    Select Case fgDataTM.Col
        Case 0 'Nama Pelayanan
            txtIsiTM.MaxLength = 0
            Call subLoadTextTM
            txtIsiTM.Text = Chr(KeyAscii)
            txtIsiTM.SelStart = Len(txtIsiTM.Text)

        Case 1 'jumlah
            txtIsiTMJml.MaxLength = 4
            If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyEscape Or KeyAscii = Asc(".")) Then Exit Sub
            Call subLoadTextTMJml
            txtIsiTMJml.Text = Chr(KeyAscii)
            txtIsiTMJml.SelStart = Len(txtIsiTMJml.Text)

        Case 2 'keterangan
            Call subLoadDataComboTM(cbCito)

        Case 3 'cito
            fgData.Col = 1
            txtIsiTM.MaxLength = 200
            Call subLoadTextTM
            txtIsiTM.Text = Chr(KeyAscii)
            txtIsiTM.SelStart = Len(txtIsiTM.Text)

    End Select
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    Call centerForm(Me, MDIUtama)
    dtpTglOrder.Value = Now
    dtpTglResep.Value = Now

    Call subSetGrid
    Call subLoadDcSource
    dgDokter.Visible = False
    dgDokterTM.Visible = False
    dcNamaPelayanan.BoundText = ""

    dgObatAlkes.Top = 2880
    dgObatAlkes.Left = 2040
    dgObatAlkes.Visible = False

    SSTab1.Tab = 0

    Call PlayFlashMovie(Me)
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo hell
    dbConn.Execute "DELETE FROM TempDetailApotikJual WHERE (NoTemporary = '" & txtNoTemporary & "')"
    strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdJenisPasien = rs("KdKelompokPasien").Value
        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Call subSetGrid
    Call cmdBatal_Click
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)

    Select Case SSTab1.Tab
        Case 0
            If KeyAscii = 13 Then txtDokterTM.SetFocus
        Case 1
            If KeyAscii = 13 Then txtNoResep.SetFocus
    End Select

End Sub

Private Sub cmdBatal_Click()
    Select Case SSTab1.Tab
        Case 0 'kelompok pegawai
            Call subKosong

        Case 1 'Jenis Pegawai
            Call subKosong
    End Select
End Sub

Private Sub subKosong()
    Select Case SSTab1.Tab
        Case 0
            txtDokterTM.Text = ""
            dcRuanganTujuanTM.BoundText = ""
            dgDokterTM.Visible = False
            dgPelayananRS.Visible = False

        Case 1
            txtDokter.Text = ""
            dcRuanganTujuan.BoundText = ""
            dgDokter.Visible = False
            dgObatAlkes.Visible = False
    End Select
End Sub

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
    If KeyCode = vbKeyDown Then
        If dgObatAlkes.Visible = False Then Exit Sub
        dgObatAlkes.SetFocus
    End If
End Sub

Private Sub txtDokter_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then
        If dgDokter.Visible = False Then Exit Sub
        dgDokter.SetFocus
    End If
End Sub

Private Sub txtDokterTM_Change()
    On Error GoTo errLoad
    mstrFilterDokter = "WHERE NamaDokter like '%" & txtDokterTM.Text & "%'"
    txtKdDokterTM.Text = ""
    dgDokterTM.Visible = True
    Call subLoadDokterTM
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtDokterTM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If dgDokterTM.Visible = False Then Exit Sub
        dgDokterTM.SetFocus
    End If
End Sub

Private Sub txtDokterTM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dgDokterTM.Visible = False Then Exit Sub
        dgDokterTM.SetFocus
    End If
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtIsi_Change()
    On Error GoTo hell
    Dim i As Integer
    Select Case fgData.Col

'        Case 2 'kode barang
'
'            If tempStatusTampil = True Then Exit Sub
'            strSQL = "execute CariBarangNStokMedis_V '" & txtIsi.Text & "%','" & dcRuanganTujuan.BoundText & "'"
'            Call msubRecFO(dbRst, strSQL)
'
'            If dcRuanganTujuan.Text = "" Then
'            MsgBox "Ma'af Pilih Terlebih Dahulu Ruang Tujuan!", vbInformation + vbOKOnly, "Info"
'            dcRuanganTujuan.SetFocus
'            dcRuanganTujuan.Text = ""
'            Exit Sub
'            End If
'
'            Set dgObatAlkes.DataSource = dbRst
'            With dgObatAlkes
'                For i = 0 To .Columns.Count - 1
'                    .Columns(i).Width = 0
'                Next i
'
'                .Columns("KdBarang").Width = 1500
'                .Columns("NamaBarang").Width = 3000
'                .Columns("JenisBarang").Width = 1500
'                .Columns("Kekuatan").Width = 1000
'                .Columns("AsalBarang").Width = 1000
'                .Columns("Satuan").Width = 675
'
'                .Top = txtIsi.Top + txtIsi.Height + Frame8.Top
'                .Left = 1820
'                .Visible = True
'
'            End With
                    
        Case 3 ' nama barang
        
        
            If tempStatusTampil = True Then Exit Sub
            strSQL = "execute CariBarangNStokMedis_V '" & txtIsi.Text & "%','" & dcRuanganTujuan.BoundText & "'"
            Call msubRecFO(dbRst, strSQL)
            
'           If dcRuanganTujuan.Text = "" Then
'            MsgBox "Pilih Terlebih Dahulu Ruang Tujuan!", vbInformation + vbOKOnly, "Info"
'            dcRuanganTujuan.SetFocus
'            dcRuanganTujuan.Text = ""
'            Exit Sub
'            End If
'
'
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
                
                .Top = 3600 '1680 'txtIsi.Top + txtIsi.Height '+ Frame8.Top
                .Left = 2400 '3360
                .Visible = True
                
            End With
        Case Else
            dgObatAlkes.Visible = False

           
    End Select
    Exit Sub
hell:
    Call msubPesanError
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

        With fgData
            Select Case .Col
                Case 0
                    Call SetKeyPressToNumber(KeyAscii)
                Case 10
                    Call SetKeyPressToNumber(KeyAscii)
            End Select
        End With


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
                    dcJenisObat.SetFocus

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
               
                If Trim(txtIsi.Text) = "," Then txtIsi.Text = 0
                If Trim(txtIsi.Text) = "" Then txtIsi.Text = 0
              
                If txtIsi.Text = 0 Then
                       MsgBox "Jumlah barang tidak boleh nol (0)", vbCritical
                       Exit Sub
                End If
              
                    If CDbl(txtIsi.Text) <= 0 Then txtIsi.Text = 0
                    If (fgData.TextMatrix(.Row, 6) = "S") Then
                        If CDbl(txtIsi.Text) > CDbl(.TextMatrix(.Row, 9)) Then
                            MsgBox "Jumlah lebih besar dari stock (" & .TextMatrix(.Row, 9) & ")", vbExclamation, "Validasi"
                            txtIsi.SelStart = 0: txtIsi.SelLength = Len(txtIsi.Text)
                            Exit Sub
                        End If
                     
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
                    
               
                    .TextMatrix(.Row, 22) = ((.TextMatrix(.Row, 14) * .TextMatrix(.Row, 15)) + _
                        CDbl((.TextMatrix(.Row, 10)) * CDbl(.TextMatrix(.Row, 16)))) + CDbl(.TextMatrix(.Row, 26))
                    Call subHitungTotal
                    fgData.SetFocus
                    fgData.Col = 27
                    'Call subLoadCheck
                    'end fifo
                    
'
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
End Sub

Private Sub txtIsi_LostFocus()
    txtIsi.Visible = False
End Sub

Private Sub txtIsiJml_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim i As Integer
    Dim curHutangPenjamin As Currency
    Dim curTanggunganRS As Currency
    Dim curHarusDibayar As Currency
    Dim KdJnsObat As String

    If KeyAscii = 13 Then
        With fgData
            Select Case .Col
                Case 0
    
                Case 4
                    If Val(txtIsiJml.Text) = 0 Then txtIsiJml.Text = 0

                    'konvert koma col jumlah
                    .TextMatrix(.Row, .Col) = msubKonversiKomaTitik(txtIsiJml.Text)

                    txtIsiJml.Visible = False

                    fgData.SetFocus
                    fgData.Col = 5

                Case 6

                Case 7

            End Select
        End With

    ElseIf KeyAscii = 27 Then
        txtIsiJml.Visible = False
        dgObatAlkes.Visible = False
        fgData.SetFocus
    ElseIf (KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = Asc(".") Or KeyAscii = vbKeySpace) Then
        If fgData.Col <> 3 Then
            dgObatAlkes.Visible = False
        Else
            dgObatAlkes.Visible = True
            txtIsiJml.Visible = True
        End If
    End If
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = 44) Then KeyAscii = 0
End Sub

Private Sub txtIsiJml_LostFocus()
    txtIsiJml.Visible = False
End Sub

Private Sub subSetGrid()
    On Error GoTo errLoad

    Select Case SSTab1.Tab

        Case 0
            With fgDataTM
                .Clear
                .Rows = 2
                .Cols = 5

                .RowHeight(0) = 400

                .TextMatrix(0, 0) = "Nama Pelayanan"
                .TextMatrix(0, 1) = "Jumlah"
                .TextMatrix(0, 2) = "Cito"
                .TextMatrix(0, 3) = "Keterangan Lainnya"
                .TextMatrix(0, 4) = "KdPelayananRS"

                .ColWidth(0) = 5500
                .ColWidth(1) = 1000
                .ColWidth(2) = 1000
                .ColWidth(3) = 4500
                .ColWidth(4) = 0

            End With

        Case 1
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
        .TextMatrix(0, 29) = "Keterangan Pakai 2"
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
        .ColWidth(2) = 1200
        .ColWidth(3) = 2450
        .ColWidth(4) = 0
        .ColWidth(5) = 1100
        .ColWidth(6) = 0
        .ColWidth(7) = 1200
        .ColWidth(8) = 0
        .ColWidth(9) = 700
        .ColWidth(10) = 700
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
        .ColWidth(27) = 1300
        .ColWidth(28) = 1700
        .ColWidth(29) = 0 '1700
        .ColWidth(30) = 0 '1700
        .ColWidth(31) = 0 '1800
        .ColWidth(32) = 0
        .ColWidth(33) = 0 ' add NoTerima
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
           
    End Select
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadDcSource()
    On Error GoTo errLoad

     Call msubDcSource(dcJenisObat, rs, "SELECT KdJenisObat, JenisObat FROM JenisObat where StatusEnabled=1 ORDER BY JenisObat")

    strSQL = "SELECT     KdRuangan, NamaRuangan From Ruangan WHERE     (StatusEnabled = 1)"
    
    
    Call msubDcSource(dcAturanPakai, rs, "select KdSatuanEtiket, NamaExternal,SatuanEtiket from SatuanEtiketResep where StatusEnabled=1 Order By SatuanEtiket")
    Call msubDcSource(dcKeteranganPakai, rs, "select KdWaktuEtiket,WaktuEtiket from WaktuEtiketResep where StatusEnabled=1 order by KdWaktuEtiket")
  
    Call msubDcSource(dcKeteranganPakai2, rs, "select KdWaktuEtiket2,WaktuEtiket2 from WaktuEtiketResep2 where StatusEnabled=1 order by WaktuEtiket2")
    
    Call msubDcSource(dcRuanganTujuanTM, rs, strSQL)
    Call msubDcSource(dcRuanganTujuan, rs, strSQL)
    strSQL = "SELECT TOP (200) KdPelayananRS, NamaPelayanan FROM V_ListPelayanan"
    Call msubDcSource(dcNamaPelayanan, rs, strSQL)

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtIsiTM_Change()
    On Error GoTo hell
    Dim i As Integer
    Select Case fgDataTM.Col

        Case 0 ' nama barang
'            strSQL = "SELECT TOP (200) [Jenis Pelayanan], KdPelayananRS, NamaPelayanan" & _
'            " FROM V_ListPelayanan where NamaPelayanan like'" & txtIsiTM.Text & "%' "
'           If Periksa("datacombo", dcRuanganTujuanTM, "Pilih ruangan tujuan") = False Then Exit Sub
            strSQL = "SELECT TOP (200) JenisPelayanan as [Jenis Pelayanan], KdPelayananRS, NamaPelayanan" & _
            " FROM V_DetailPelayananMedisNew where NamaPelayanan like'" & txtIsiTM.Text & "%' and KdRuangan like'" & dcRuanganTujuanTM.BoundText & "%' and KdKelas='" & mstrKdKelas & "' "

            Call msubRecFO(dbRst, strSQL)

            Set dgPelayananRS.DataSource = dbRst
            With dgPelayananRS
                .Columns("Jenis Pelayanan").Width = 3000
                .Columns("KdPelayananRS").Width = 0
                .Columns("NamaPelayanan").Width = 5000

                .Top = 3400
                .Left = 120
                .Visible = True
                For i = 1 To fgDataTM.Row - 1
                    .Top = .Top + fgDataTM.RowHeight(i)
                Next i
                If fgDataTM.TopRow > 1 Then
                    .Top = .Top - ((fgDataTM.TopRow - 1) * fgDataTM.RowHeight(1))
                End If
            End With
        Case Else
            dgPelayananRS.Visible = False
    End Select
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub txtIsiTM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then If dgPelayananRS.Visible = True Then If dgPelayananRS.ApproxCount > 0 Then dgPelayananRS.SetFocus
End Sub

Private Sub txtIsiTM_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Dim i As Integer

    If KeyAscii = 13 Then
        With fgDataTM
            Select Case .Col

                Case 0 ' nama pel
                    If dgPelayananRS.Visible = True Then
                        dgPelayananRS.SetFocus
                        Exit Sub
                    Else
                        fgDataTM.SetFocus
                        fgDataTM.Col = 1
                    End If

                Case 1 ' jumlah

'                    fgDataTM.TextMatrix(fgDataTM.Row, 1) = txtIsiTM.Text
'                    fgDataTM.SetFocus
'                    fgDataTM.Col = 2

                Case 2 ' cito
                    fgDataTM.SetFocus
                    fgDataTM.Col = 3

                Case 3 ' ket
                    fgDataTM.TextMatrix(fgDataTM.Row, 3) = txtIsiTM.Text
                    fgDataTM.SetFocus
                    fgDataTM.Col = 3

            End Select
        End With

    ElseIf KeyAscii = 27 Then
        txtIsiTM.Visible = False
        dgPelayananRS.Visible = False
        fgDataTM.SetFocus

    End If
End Sub

Private Sub txtIsiTM_LostFocus()
    txtIsiTM.Visible = False
End Sub

Private Sub txtIsiTMJml_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim i As Integer

    If KeyAscii = 13 Then
        With fgDataTM
            Select Case .Col
                Case 0 ' nama pel

                Case 1 ' jumlah
                        fgDataTM.TextMatrix(fgDataTM.Row, 1) = txtIsiTMJml.Text
                        fgDataTM.SetFocus
                        fgDataTM.Col = 2

                Case 2 ' cito
                
                Case 3 ' ket
                   
            End Select
        End With

    ElseIf KeyAscii = 27 Then
        txtIsiTMJml.Visible = False
        dgPelayananRS.Visible = False
        fgDataTM.SetFocus
    End If
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = 44) Then KeyAscii = 0
End Sub

Private Sub txtIsiTMJml_LostFocus()
    txtIsiTMJml.Visible = False
End Sub



Private Sub txtNoResep_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtDokter.SetFocus
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
    dgDokter.Left = 5760
    dgDokter.Top = 2520

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

'untuk meload data dokter di grid
Private Sub subLoadDokterTM()
    On Error GoTo errLoad

    strSQL = "SELECT NamaDokter AS [Nama Dokter],JK,Jabatan,KodeDokter  FROM V_DaftarDokter " & mstrFilterDokter
    Call msubRecFO(rs, strSQL)
    With dgDokterTM
        Set .DataSource = rs
        .Columns(0).Width = 3500
        .Columns(1).Width = 400
        .Columns(2).Width = 1600
        .Columns(3).Width = 0
    End With
    dgDokter.Left = 5760
    dgDokter.Top = 2520

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Function sp_Order() As Boolean
    On Error GoTo errLoad

    sp_Order = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("TglOrder", adDate, adParamInput, , Format(dtpTglOrder.Value, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, IIf(mstrKdRuangan = "", Null, mstrKdRuangan))
        If SSTab1.Tab = 0 Then
            .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, IIf(dcRuanganTujuanTM.Text = "", Null, dcRuanganTujuanTM.BoundText))
        Else
            .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, IIf(dcRuanganTujuan.Text = "", Null, dcRuanganTujuan.BoundText))
        End If
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

Private Function sp_DetailOrderPelayananOANew(f_KdBarang As String, f_JmlBarang As Double, f_ResepKe As Integer, f_KdSatuanEtiket As String, f_KdWaktuEtiket As String, f_KdWaktuEtiket2 As String, f_KdPelayananRSUsed As String, f_KeteranganLainnya As String, _
        f_JmlService As Integer, f_TarifService As Currency, f_Cito As String, f_KdJenisObat As String, f_KdAsal As String, f_SatuanJml As String, f_NoTerima As String, f_Noracikan As String) As Boolean
On Error GoTo errLoad
    
    sp_DetailOrderPelayananOANew = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, txtNoOrder.Text)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtnocm.Text)
        .Parameters.Append .CreateParameter("NoPakai", adVarChar, adParamInput, 10, IIf(Trim(txtNoPakai.Text) = "", Null, Trim(txtNoPakai.Text)))
        .Parameters.Append .CreateParameter("idDokterOrder", adChar, adParamInput, 10, IIf(Trim(txtKdDokter.Text) = "", Null, Trim(txtKdDokter.Text)))
        .Parameters.Append .CreateParameter("StatusCito", adChar, adParamInput, 1, IIf(f_Cito = "Yes", 1, 0))
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("JmlBarang", adDouble, adParamInput, , f_JmlBarang)
        .Parameters.Append .CreateParameter("NoResep", adChar, adParamInput, 15, IIf(Trim(txtNoResep.Text) = "", Null, Trim(txtNoResep.Text)))
        .Parameters.Append .CreateParameter("ResepKe", adTinyInt, adParamInput, , f_ResepKe)
        .Parameters.Append .CreateParameter("KdSatuanEtiket", adChar, adParamInput, 2, f_KdSatuanEtiket)
        .Parameters.Append .CreateParameter("KdWaktuEtiket", adChar, adParamInput, 2, f_KdWaktuEtiket)
        .Parameters.Append .CreateParameter("KdWaktuEtiket2", adChar, adParamInput, 2, f_KdWaktuEtiket2)
        .Parameters.Append .CreateParameter("TglResep", adDate, adParamInput, , Format(dtpTglResep.Value, "yyyy/MM/dd hh:mm:ss"))
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("NoRetur", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("JmlRetur", adInteger, adParamInput, , Null)
        .Parameters.Append .CreateParameter("KdPelayananRSUSed", adChar, adParamInput, 6, IIf(f_KdPelayananRSUsed = "", Null, f_KdPelayananRSUsed))
        .Parameters.Append .CreateParameter("KeteranganLainnya", adVarChar, adParamInput, 200, IIf(f_KeteranganLainnya = "", Null, f_KeteranganLainnya))
        .Parameters.Append .CreateParameter("JmlService", adInteger, adParamInput, , f_JmlService)
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
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_DetailOrderPelayananOANew = False

        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    
Exit Function
errLoad:
    Call msubPesanError
    sp_DetailOrderPelayananOANew = False
End Function



Private Function sp_DetailOrderPelayananOA(f_KdBarang As String, f_KdAsal As String, f_JmlBarang As Integer, f_ResepKe As Integer, f_AturanPakai As String, f_KeteranganPakai As String, f_KdPelayananRSUsed As String, f_KeteranganLainnya As String, f_Cito As String) As Boolean
    On Error GoTo errLoad

    sp_DetailOrderPelayananOA = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, txtNoOrder.Text)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaranTM.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCMTM.Text)
        .Parameters.Append .CreateParameter("NoPakai", adVarChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("idDokterOrder", adChar, adParamInput, 10, txtKdDokter.Text)

        .Parameters.Append .CreateParameter("StatusCito", adChar, adParamInput, 1, IIf(f_Cito = "Yes", 1, 0))

        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("KdAsal", adVarChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("JmlBarang", adInteger, adParamInput, , f_JmlBarang)
        .Parameters.Append .CreateParameter("NoResep", adChar, adParamInput, 15, txtNoResep.Text)
        .Parameters.Append .CreateParameter("ResepKe", adTinyInt, adParamInput, , f_ResepKe)
        .Parameters.Append .CreateParameter("AturanPakai", adVarChar, adParamInput, 50, f_AturanPakai)
        .Parameters.Append .CreateParameter("KeteranganPakai", adVarChar, adParamInput, 50, f_KeteranganPakai)
        .Parameters.Append .CreateParameter("TglResep", adDate, adParamInput, , Format(dtpTglResep.Value, "yyyy/MM/dd hh:mm:ss"))
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("NoRetur", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("JmlRetur", adInteger, adParamInput, , 0)
        .Parameters.Append .CreateParameter("KdPelayananRSUSed", adChar, adParamInput, 6, IIf(f_KdPelayananRSUsed = "", Null, f_KdPelayananRSUsed))
        .Parameters.Append .CreateParameter("KeteranganLainnya", adVarChar, adParamInput, 200, IIf(f_KeteranganLainnya = "", Null, f_KeteranganLainnya))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_DetailOrderPelayananOA"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
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

Private Function sp_DetailOrderPelayananTM(f_KdPelayananRS As String, f_JmlBarang As Integer, f_Cito As String, f_KeteranganLain As String) As Boolean
    On Error GoTo errLoad

    sp_DetailOrderPelayananTM = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, txtNoOrder.Text)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, f_KdPelayananRS)
        .Parameters.Append .CreateParameter("JmlPelayanan", adInteger, adParamInput, , f_JmlBarang)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaranTM.Text)
        .Parameters.Append .CreateParameter("idDokterOrder", adChar, adParamInput, 10, txtKdDokterTM.Text)

        .Parameters.Append .CreateParameter("StatusCito", adChar, adParamInput, 1, IIf(f_Cito = "Yes", 1, 0))

        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCMTM.Text)
        .Parameters.Append .CreateParameter("NoPakai", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("NoRetur", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("JmlRetur", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("KdPelayananRSUsed", adChar, adParamInput, 6, Null)
        .Parameters.Append .CreateParameter("KeteranganLainnya", adVarChar, adParamInput, 200, IIf(f_KeteranganLain = "", Null, f_KeteranganLain))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_DetailOrderPelayananTM"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_DetailOrderPelayananTM = False

        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function
errLoad:
    Call msubPesanError
    sp_DetailOrderPelayananTM = False
End Function

Private Sub txtNoResep_LostFocus()

    txtDokter.SetFocus
End Sub

Private Function sp_TempDetailApotikJual(f_HargaSatuan As Currency) As Boolean
    sp_TempDetailApotikJual = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoTemporary", adChar, adParamInput, 3, IIf(Len(Trim(txtNoTemporary.Text)) = 0, Null, Trim(txtNoTemporary.Text)))
        .Parameters.Append .CreateParameter("KdKelompokPasien", adChar, adParamInput, 2, mstrKdJenisPasien)
        .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, mstrKdPenjaminPasien)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, fgData.TextMatrix(fgData.Row, 2))
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, fgData.TextMatrix(fgData.Row, 12))
        .Parameters.Append .CreateParameter("HargaSatuan", adCurrency, adParamInput, , f_HargaSatuan)
        .Parameters.Append .CreateParameter("NoTemporaryOutput", adChar, adParamOutput, 10, Null)
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
            Call Add_HistoryLoginActivity("Add_TemporaryDetailApotikJual")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Sub subHitungTotal()
    On Error GoTo errLoad
    Dim i As Integer

    If fgData.TextMatrix(fgData.Row - 1, 11) = "" Then Exit Sub
    txtTotalBiaya.Text = 0
    txtHutangPenjamin.Text = 0
    txtTanggunganRS.Text = 0
    txtHarusDibayar.Text = 0
    txtTotalDiscount.Text = 0

    With fgData
        For i = 1 To IIf(fgData.TextMatrix(fgData.Rows - 1, 2) = "", fgData.Rows - 2, fgData.Rows - 1)
            txtTotalBiaya.Text = txtTotalBiaya.Text + CDbl(.TextMatrix(i, 11))
            txtHutangPenjamin.Text = txtHutangPenjamin.Text + CDbl(.TextMatrix(i, 19))
            txtTanggunganRS.Text = txtTanggunganRS.Text + CDbl(.TextMatrix(i, 20))
            txtTotalDiscount.Text = txtTotalDiscount.Text + CDbl(.TextMatrix(i, 21))
            txtHarusDibayar.Text = txtHarusDibayar.Text + CDbl(.TextMatrix(i, 22))
        Next i
    End With

    txtTotalBiaya.Text = IIf(Val(txtTotalBiaya.Text) = 0, 0, Format(txtTotalBiaya.Text, "#,###"))
    txtHutangPenjamin.Text = IIf(Val(txtHutangPenjamin.Text) = 0, 0, Format(txtHutangPenjamin.Text, "#,###"))
    txtTanggunganRS.Text = IIf(Val(txtTanggunganRS.Text) = 0, 0, Format(txtTanggunganRS.Text, "#,###"))
    txtHarusDibayar.Text = IIf(Val(txtHarusDibayar.Text) = 0, 0, Format(txtHarusDibayar.Text, "#,###"))
    txtTotalDiscount.Text = IIf(Val(txtTotalDiscount.Text) = 0, 0, Format(txtTotalDiscount.Text, "#,###"))

    subcurHarusDibayar = txtHarusDibayar.Text

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
        chkDokterPelmeriksa.Value = vbUnchecked
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
        If sp_TempDetailApotikJual(CDbl(dbRst("HargaSatuan"))) = False Then Exit Sub  'discount
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
            curHarusDibayar = CDbl(.TextMatrix(.Rows - 1, 11)) - CDbl(.TextMatrix(.Rows - 1, 21)) - _
            (CDbl(.TextMatrix(.Rows - 1, 19)) + CDbl(.TextMatrix(.Rows - 1, 120)))
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
            txtIsi.MaxLength = 4
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

Private Sub subLoadTextJml()
Dim i As Integer
    txtIsiJml.Left = fgData.Left

    Select Case fgData.Col
        Case 0
            txtIsiJml.MaxLength = 2

        Case 3
            txtIsiJml.MaxLength = 20

        Case 10
            txtIsiJml.MaxLength = 4
    End Select

    For i = 0 To fgData.Col - 1
        txtIsiJml.Left = txtIsiJml.Left + fgData.ColWidth(i)
    Next i
    txtIsiJml.Visible = True
    txtIsiJml.Top = fgData.Top - 7

    For i = 0 To fgData.Row - 1
        txtIsiJml.Top = txtIsiJml.Top + fgData.RowHeight(i)
    Next i

    If fgData.TopRow > 1 Then
        txtIsiJml.Top = txtIsiJml.Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
    End If

    txtIsi.Width = fgData.ColWidth(fgData.Col)

    txtIsiJml.Visible = True
    txtIsiJml.SelStart = Len(txtIsi.Text)
    txtIsiJml.SetFocus
End Sub

Private Sub subLoadTextTM()
    Dim i As Integer
    txtIsiTM.Left = fgDataTM.Left

    For i = 0 To fgDataTM.Col - 1
        txtIsiTM.Left = txtIsiTM.Left + fgDataTM.ColWidth(i)
    Next i
    txtIsiTM.Visible = True
    txtIsiTM.Top = fgDataTM.Top - 7

    For i = 0 To fgDataTM.Row - 1
        txtIsiTM.Top = txtIsiTM.Top + fgDataTM.RowHeight(i)
    Next i

    If fgDataTM.TopRow > 1 Then
        txtIsiTM.Top = txtIsiTM.Top - ((fgDataTM.TopRow - 1) * fgDataTM.RowHeight(1))
    End If

    txtIsiTM.Width = fgDataTM.ColWidth(fgDataTM.Col)

    txtIsiTM.Visible = True
    txtIsiTM.SelStart = Len(txtIsiTM.Text)
    txtIsiTM.SetFocus
End Sub

Private Sub subLoadTextTMJml()
    Dim i As Integer
    txtIsiTMJml.Left = fgDataTM.Left

    For i = 0 To fgDataTM.Col - 1
        txtIsiTMJml.Left = txtIsiTM.Left + fgDataTM.ColWidth(i)
    Next i
    txtIsiTMJml.Visible = True
    txtIsiTMJml.Top = fgDataTM.Top - 7

    For i = 0 To fgDataTM.Row - 1
        txtIsiTMJml.Top = txtIsiTMJml.Top + fgDataTM.RowHeight(i)
    Next i

    If fgDataTM.TopRow > 1 Then
        txtIsiTMJml.Top = txtIsiTMJml.Top - ((fgDataTM.TopRow - 1) * fgDataTM.RowHeight(1))
    End If

    txtIsiTMJml.Width = fgDataTM.ColWidth(fgDataTM.Col)
    txtIsiTMJml.Visible = True
    txtIsiTMJml.SelStart = Len(txtIsiTM.Text)
    txtIsiTMJml.SetFocus
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

    s_DcName.Visible = True
    s_DcName.SetFocus
End Sub

Private Sub subLoadDataComboTM(s_DcName As Object)
    Dim i As Integer
    s_DcName.Left = fgDataTM.Left
    For i = 0 To fgDataTM.Col - 1
        s_DcName.Left = s_DcName.Left + fgDataTM.ColWidth(i)
    Next i
    s_DcName.Visible = True
    s_DcName.Top = fgDataTM.Top - 7

    For i = 0 To fgDataTM.Row - 1
        s_DcName.Top = s_DcName.Top + fgDataTM.RowHeight(i)
    Next i

    If fgDataTM.TopRow > 1 Then
        s_DcName.Top = s_DcName.Top - ((fgDataTM.TopRow - 1) * fgDataTM.RowHeight(1))
    End If

    s_DcName.Width = fgDataTM.ColWidth(fgDataTM.Col)

    s_DcName.Visible = True
    s_DcName.SetFocus
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

