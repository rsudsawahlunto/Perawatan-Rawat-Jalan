VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTransaksiPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Transaksi Pelayanan Pasien"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   15375
   Begin VB.Frame Frame1 
      Caption         =   "Pelayanan Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   0
      TabIndex        =   58
      Top             =   2040
      Width           =   15375
      Begin VB.CommandButton cmdCetakCatatanMedis 
         Caption         =   "Cetak"
         Height          =   375
         Left            =   -9360
         TabIndex        =   109
         Top             =   6240
         Width           =   1575
      End
      Begin VB.CommandButton cmdCetakResume 
         Caption         =   "Cetak &Resume"
         Height          =   375
         Left            =   -9360
         TabIndex        =   108
         Top             =   6240
         Width           =   1575
      End
      Begin VB.CommandButton cmdTotalTanggunganPenjamin 
         Caption         =   "Hitung Claim Ina Cbg's"
         Height          =   375
         Left            =   11040
         TabIndex        =   104
         Top             =   6240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtTotalPenjamin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   6720
         TabIndex        =   103
         Top             =   6240
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   13200
         TabIndex        =   46
         Top             =   6240
         Width           =   2055
      End
      Begin VB.TextBox txtGrandTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         TabIndex        =   45
         Top             =   6240
         Width           =   2415
      End
      Begin TabDlg.SSTab sstTP 
         Height          =   5775
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   10186
         _Version        =   393216
         Tabs            =   10
         Tab             =   6
         TabsPerRow      =   10
         TabHeight       =   1323
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Pelayanan Tindakan"
         TabPicture(0)   =   "frmTransaksiPasien.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(1)=   "cmdUbahPT"
         Tab(0).Control(2)=   "cmdTambahPT"
         Tab(0).Control(3)=   "cmdHapusDataPT"
         Tab(0).Control(4)=   "txtTindakanTotal"
         Tab(0).Control(5)=   "dgTindakan"
         Tab(0).Control(6)=   "Label1"
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "Pemakaian Obat && Alkes"
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label2"
         Tab(1).Control(1)=   "dgObatAlkes"
         Tab(1).Control(2)=   "txtAlkesTotal"
         Tab(1).Control(3)=   "cmdHapusDataPOA"
         Tab(1).Control(4)=   "cmdTambahPOA"
         Tab(1).Control(5)=   "cmdEditData"
         Tab(1).Control(6)=   "cmdUbahOA"
         Tab(1).Control(7)=   "chkAlergi"
         Tab(1).Control(8)=   "dgAlergi"
         Tab(1).Control(9)=   "fraRiwayatResep"
         Tab(1).Control(10)=   "cmdRiwayatResep"
         Tab(1).Control(11)=   "picEditQuanttyBarang"
         Tab(1).ControlCount=   12
         TabCaption(2)   =   "Riwayat Catatan Klinis"
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "dgCatatanKlinis"
         Tab(2).Control(1)=   "cmdTambahCatatanKlinis"
         Tab(2).Control(2)=   "cmdHapusCatataKlinis"
         Tab(2).Control(3)=   "cmdKehamilandanKB"
         Tab(2).ControlCount=   4
         TabCaption(3)   =   "Riwayat Catatan Medis"
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "dgCatatanMedis"
         Tab(3).Control(1)=   "cmdTambahCatatanMedis"
         Tab(3).Control(2)=   "cmdHapusCatatanMedis"
         Tab(3).Control(3)=   "cmdUbahCatatanMedis"
         Tab(3).ControlCount=   4
         TabCaption(4)   =   "Riwayat Diagnosa"
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "dgRiwayatDiagnosa"
         Tab(4).Control(1)=   "cmdDelDiagnosa"
         Tab(4).Control(2)=   "cmdTambahDiagnosa"
         Tab(4).Control(3)=   "cmdCetakDiagnosa"
         Tab(4).Control(4)=   "cmdICD9"
         Tab(4).Control(5)=   "chkTampil"
         Tab(4).ControlCount=   6
         TabCaption(5)   =   "Riwayat Tindakan Medis"
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "dgRiwayatOperasi"
         Tab(5).Control(1)=   "cmdTambahTM"
         Tab(5).Control(2)=   "cmdHapusTM"
         Tab(5).ControlCount=   3
         TabCaption(6)   =   "Riwayat Konsul"
         TabPicture(6)   =   "frmTransaksiPasien.frx":001C
         Tab(6).ControlEnabled=   -1  'True
         Tab(6).Control(0)=   "dgKonsul"
         Tab(6).Control(0).Enabled=   0   'False
         Tab(6).Control(1)=   "cmdTambahKonsul"
         Tab(6).Control(1).Enabled=   0   'False
         Tab(6).Control(2)=   "cmdHapusKonsul"
         Tab(6).Control(2).Enabled=   0   'False
         Tab(6).Control(3)=   "fraOrder"
         Tab(6).Control(3).Enabled=   0   'False
         Tab(6).ControlCount=   4
         TabCaption(7)   =   "Riwayat Kecelakaan"
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "dgKecelakaan"
         Tab(7).Control(1)=   "cmdTambahKecelakaan"
         Tab(7).Control(2)=   "cmdHapusKecelakaan"
         Tab(7).ControlCount=   3
         TabCaption(8)   =   "Riwayat Pemeriksaan"
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "dgRiwayatPemeriksaan"
         Tab(8).Control(1)=   "cmdCetakRP"
         Tab(8).Control(2)=   "chkRP"
         Tab(8).ControlCount=   3
         TabCaption(9)   =   "Riwayat Hasil Pemeriksaan"
         Tab(9).ControlEnabled=   0   'False
         Tab(9).Control(0)=   "cmdCetakHasilPemeriksaan"
         Tab(9).Control(1)=   "dgHasilPemeriksaan"
         Tab(9).ControlCount=   2
         Begin VB.Frame Frame2 
            Caption         =   "Pemeriksaan Dokter"
            Height          =   735
            Left            =   -68880
            TabIndex        =   105
            Top             =   4920
            Width           =   3735
            Begin VB.CommandButton cmdUmum 
               Caption         =   "Dokter &Umum"
               Height          =   375
               Left            =   240
               TabIndex        =   107
               Top             =   240
               Width           =   1575
            End
            Begin VB.CommandButton cmdSpesialis 
               Caption         =   "Dokter &Speisalis"
               Height          =   375
               Left            =   1920
               TabIndex        =   106
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.CommandButton cmdUbahCatatanMedis 
            Caption         =   "U&bah Data"
            Height          =   375
            Left            =   -65040
            TabIndex        =   101
            Top             =   5040
            Width           =   1575
         End
         Begin VB.PictureBox picEditQuanttyBarang 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   2295
            Left            =   -74160
            ScaleHeight     =   2265
            ScaleWidth      =   7065
            TabIndex        =   88
            Top             =   1320
            Visible         =   0   'False
            Width           =   7095
            Begin VB.Frame fraUbahQuantityBarang 
               Caption         =   "Ubah Quantity Barang"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1935
               Left            =   240
               TabIndex        =   89
               Top             =   120
               Width           =   6615
               Begin VB.TextBox txtNoterima 
                  Appearance      =   0  'Flat
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1440
                  TabIndex        =   100
                  Text            =   "txtNoterima"
                  Top             =   1560
                  Visible         =   0   'False
                  Width           =   1095
               End
               Begin VB.CommandButton cmdSimpanEditBarang 
                  Caption         =   "&Simpan"
                  Height          =   375
                  Left            =   3120
                  Style           =   1  'Graphical
                  TabIndex        =   96
                  Top             =   1320
                  Width           =   1575
               End
               Begin VB.TextBox txtNamaBarangEdit 
                  Appearance      =   0  'Flat
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   95
                  Top             =   360
                  Width           =   4815
               End
               Begin VB.TextBox txtJmlBarangEditAwal 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   94
                  TabStop         =   0   'False
                  Text            =   "0"
                  Top             =   720
                  Width           =   1095
               End
               Begin VB.TextBox txtKdBarangEdit 
                  Appearance      =   0  'Flat
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   240
                  TabIndex        =   93
                  Text            =   "txtKdBarangEdit"
                  Top             =   1320
                  Visible         =   0   'False
                  Width           =   1095
               End
               Begin VB.TextBox txtJmlBarangEditBaru 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   4680
                  TabIndex        =   92
                  Text            =   "0"
                  Top             =   720
                  Width           =   1095
               End
               Begin VB.TextBox txtKdAsalEdit 
                  Appearance      =   0  'Flat
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1440
                  TabIndex        =   91
                  Text            =   "txtKdAsalEdit"
                  Top             =   1320
                  Visible         =   0   'False
                  Width           =   1095
               End
               Begin VB.CommandButton cmdBatalEdit 
                  Caption         =   "Tutu&p"
                  Height          =   375
                  Left            =   4800
                  TabIndex        =   90
                  Top             =   1320
                  Width           =   1575
               End
               Begin VB.Line Line1 
                  X1              =   240
                  X2              =   6360
                  Y1              =   1200
                  Y2              =   1200
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Jumlah Awal"
                  Height          =   210
                  Index           =   31
                  Left            =   240
                  TabIndex        =   99
                  Top             =   720
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nama Barang"
                  Height          =   210
                  Index           =   32
                  Left            =   240
                  TabIndex        =   98
                  Top             =   360
                  Width           =   1065
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Jumlah Tambahan"
                  Height          =   210
                  Index           =   14
                  Left            =   3000
                  TabIndex        =   97
                  Top             =   720
                  Width           =   1470
               End
            End
         End
         Begin VB.Frame fraOrder 
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
            Left            =   11880
            TabIndex        =   85
            Top             =   4320
            Visible         =   0   'False
            Width           =   2775
            Begin VB.OptionButton optObat 
               Caption         =   "Obat  Alkes"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   1440
               TabIndex        =   87
               Top             =   120
               Width           =   1215
            End
            Begin VB.OptionButton optTindakan 
               Caption         =   "Tindakan"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   120
               TabIndex        =   86
               Top             =   120
               Width           =   975
            End
         End
         Begin VB.CommandButton cmdRiwayatResep 
            Caption         =   "&Riwayat Resep"
            Height          =   375
            Left            =   -65040
            TabIndex        =   83
            Top             =   5040
            Width           =   1575
         End
         Begin VB.Frame fraRiwayatResep 
            Caption         =   "Daftar Riwayat Resep "
            Height          =   4095
            Left            =   -74880
            TabIndex        =   80
            Top             =   840
            Width           =   14895
            Begin VB.CommandButton cmdTutup2 
               Caption         =   "Tutu&p"
               Height          =   375
               Left            =   12720
               TabIndex        =   81
               Top             =   3600
               Width           =   1935
            End
            Begin MSDataGridLib.DataGrid dgRiwayatResepPasien 
               Height          =   3255
               Left            =   120
               TabIndex        =   82
               Top             =   240
               Width           =   14535
               _ExtentX        =   25638
               _ExtentY        =   5741
               _Version        =   393216
               AllowUpdate     =   -1  'True
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
         Begin MSDataGridLib.DataGrid dgAlergi 
            Height          =   4095
            Left            =   -74520
            TabIndex        =   79
            Top             =   840
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   7223
            _Version        =   393216
            AllowUpdate     =   -1  'True
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
         Begin VB.CheckBox chkAlergi 
            Caption         =   "Data Alergi Pasien"
            Height          =   255
            Left            =   -74760
            TabIndex        =   78
            Top             =   5040
            Width           =   2535
         End
         Begin VB.CommandButton cmdCetakHasilPemeriksaan 
            Caption         =   "Cetak"
            Height          =   375
            Left            =   -61680
            TabIndex        =   77
            Top             =   4920
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusTM 
            Caption         =   "&Hapus Data"
            Height          =   375
            Left            =   -63360
            TabIndex        =   76
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahTM 
            Caption         =   "&Tambah Data"
            Height          =   375
            Left            =   -61680
            TabIndex        =   75
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CheckBox chkTampil 
            Caption         =   "Semua Diagnosa"
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
            Left            =   -74760
            TabIndex        =   30
            Top             =   5040
            Width           =   3015
         End
         Begin VB.CommandButton cmdICD9 
            Caption         =   "&Edit Diagnosa Tindakan [ICD 9]"
            Height          =   375
            Left            =   -66600
            TabIndex        =   32
            Top             =   5040
            Width           =   3135
         End
         Begin VB.CommandButton cmdKehamilandanKB 
            Caption         =   "Data &Kehamilan dan KB"
            Height          =   375
            Left            =   -65760
            TabIndex        =   23
            Top             =   5040
            Width           =   2295
         End
         Begin VB.CommandButton cmdUbahOA 
            Caption         =   "&Ubah Data"
            Height          =   375
            Left            =   -65040
            TabIndex        =   19
            Top             =   5040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdEditData 
            Caption         =   "&Edit Data"
            Height          =   375
            Left            =   -66720
            TabIndex        =   18
            Top             =   5040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdUbahPT 
            Caption         =   "U&bah Data"
            Height          =   375
            Left            =   -65040
            TabIndex        =   13
            Top             =   5160
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusKecelakaan 
            Caption         =   "&Hapus Data"
            Height          =   375
            Left            =   -63360
            TabIndex        =   40
            Top             =   5040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahKecelakaan 
            Caption         =   "&Tambah Data"
            Height          =   375
            Left            =   -61680
            TabIndex        =   41
            Top             =   5040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusKonsul 
            Caption         =   "&Hapus Data"
            Height          =   375
            Left            =   11640
            TabIndex        =   37
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahKonsul 
            Caption         =   "&Tambah Data"
            Height          =   375
            Left            =   13320
            TabIndex        =   38
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusCatatanMedis 
            Caption         =   "&Hapus Data"
            Height          =   375
            Left            =   -63360
            TabIndex        =   27
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahCatatanMedis 
            Caption         =   "&Tambah Data"
            Height          =   375
            Left            =   -61680
            TabIndex        =   28
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusCatataKlinis 
            Caption         =   "&Hapus Data"
            Height          =   375
            Left            =   -63360
            TabIndex        =   24
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahCatatanKlinis 
            Caption         =   "&Tambah Data"
            Height          =   375
            Left            =   -61680
            TabIndex        =   25
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CheckBox chkRP 
            Caption         =   "Tampilkan Semua"
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
            Left            =   -74760
            TabIndex        =   43
            Top             =   5040
            Width           =   2175
         End
         Begin VB.CommandButton cmdCetakRP 
            Caption         =   "&Cetak"
            Height          =   375
            Left            =   -61680
            TabIndex        =   44
            Top             =   5040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdCetakDiagnosa 
            Caption         =   "&Cetak"
            Height          =   375
            Left            =   -68280
            TabIndex        =   31
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahDiagnosa 
            Caption         =   "&Tambah Data"
            Height          =   375
            Left            =   -61680
            TabIndex        =   34
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdDelDiagnosa 
            Caption         =   "&Hapus Data"
            Height          =   375
            Left            =   -63360
            TabIndex        =   33
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahPOA 
            Caption         =   "Ta&mbah Data"
            Height          =   375
            Left            =   -61680
            TabIndex        =   21
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusDataPOA 
            Caption         =   "Hapu&s Data"
            Height          =   375
            Left            =   -63360
            TabIndex        =   20
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahPT 
            Caption         =   "T&ambah Data"
            Height          =   375
            Left            =   -61680
            TabIndex        =   15
            Top             =   5160
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusDataPT 
            Caption         =   "Hapus &Data"
            Height          =   375
            Left            =   -63360
            TabIndex        =   14
            Top             =   5160
            Width           =   1575
         End
         Begin VB.TextBox txtAlkesTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   -71640
            TabIndex        =   17
            Top             =   5040
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox txtTindakanTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   -72000
            TabIndex        =   12
            Top             =   5040
            Visible         =   0   'False
            Width           =   2415
         End
         Begin MSDataGridLib.DataGrid dgTindakan 
            Height          =   3975
            Left            =   -74760
            TabIndex        =   11
            Top             =   840
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   7011
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
            Height          =   4095
            Left            =   -74760
            TabIndex        =   16
            Top             =   840
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   7223
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
         Begin MSDataGridLib.DataGrid dgRiwayatDiagnosa 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   29
            Top             =   840
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   7223
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
         Begin MSDataGridLib.DataGrid dgRiwayatPemeriksaan 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   42
            Top             =   840
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   7223
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
         Begin MSDataGridLib.DataGrid dgCatatanKlinis 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   22
            Top             =   840
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   7223
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
         Begin MSDataGridLib.DataGrid dgCatatanMedis 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   26
            Top             =   840
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   7223
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
         Begin MSDataGridLib.DataGrid dgRiwayatOperasi 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   35
            Top             =   840
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   7223
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
         Begin MSDataGridLib.DataGrid dgKonsul 
            Height          =   4095
            Left            =   240
            TabIndex        =   36
            Top             =   840
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   7223
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
         Begin MSDataGridLib.DataGrid dgKecelakaan 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   39
            Top             =   840
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   7223
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
         Begin MSDataGridLib.DataGrid dgHasilPemeriksaan 
            Height          =   3975
            Left            =   -74640
            TabIndex        =   84
            Top             =   840
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   7011
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Total Biaya Pemakaian Obat && Alkes"
            Height          =   210
            Left            =   -74760
            TabIndex        =   60
            Top             =   5100
            Visible         =   0   'False
            Width           =   2925
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total Biaya Pelayanan Tindakan"
            Height          =   210
            Left            =   -74760
            TabIndex        =   59
            Top             =   5100
            Visible         =   0   'False
            Width           =   2550
         End
      End
      Begin VB.Label LbTanggunganPenjamin 
         AutoSize        =   -1  'True
         Caption         =   "Total Claim Ina Cbg's"
         Height          =   210
         Left            =   4800
         TabIndex        =   102
         Top             =   6315
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Biaya Pelayanan"
         Height          =   210
         Left            =   360
         TabIndex        =   61
         Top             =   6315
         Width           =   1755
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
      Height          =   1095
      Left            =   0
      TabIndex        =   47
      Top             =   960
      Width           =   15375
      Begin VB.TextBox txtKls 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   9840
         TabIndex        =   7
         Top             =   600
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
         Left            =   7320
         TabIndex        =   48
         Top             =   360
         Width           =   2415
         Begin VB.TextBox txtHr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            MaxLength       =   6
            TabIndex        =   6
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtBln 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   900
            MaxLength       =   6
            TabIndex        =   5
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtThn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   4
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            Height          =   210
            Left            =   2130
            TabIndex        =   51
            Top             =   270
            Width           =   165
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            Height          =   210
            Left            =   1350
            TabIndex        =   50
            Top             =   277
            Width           =   240
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            Height          =   210
            Left            =   550
            TabIndex        =   49
            Top             =   277
            Width           =   285
         End
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   0
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3240
         TabIndex        =   2
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   5880
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtJenisPasien 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   11400
         TabIndex        =   8
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtTglDaftar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   13680
         TabIndex        =   9
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Kelas Pelayanan"
         Height          =   210
         Left            =   9840
         TabIndex        =   62
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Left            =   120
         TabIndex        =   57
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   1560
         TabIndex        =   56
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   3240
         TabIndex        =   55
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   5880
         TabIndex        =   54
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pasien"
         Height          =   210
         Left            =   11400
         TabIndex        =   53
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Pendaftaran"
         Height          =   210
         Left            =   13680
         TabIndex        =   52
         Top             =   360
         Width           =   1365
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   74
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
   Begin VB.Frame Frame4 
      Caption         =   "Rekapitulasi Tagihan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   8400
      TabIndex        =   63
      Top             =   3960
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox txtPembebasan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2760
         TabIndex        =   67
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtTRS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2760
         TabIndex        =   66
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtTAsuransi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2760
         TabIndex        =   65
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox TxtTotalBiaya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2760
         TabIndex        =   64
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   240
         TabIndex        =   73
         Top             =   2160
         Width           =   1410
      End
      Begin VB.Label lblTotalTagihan 
         AutoSize        =   -1  'True
         Caption         =   "Rp. 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1680
         TabIndex        =   72
         Top             =   2160
         Width           =   1125
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Pembebasan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   71
         Top             =   1755
         Width           =   1230
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Tanggungan Rumah Sakit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   70
         Top             =   1275
         Width           =   2445
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Tanggungan Penjamin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   69
         Top             =   795
         Width           =   2115
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total Biaya Pelayanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   68
         Top             =   315
         Width           =   2130
      End
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmTransaksiPasien.frx":0038
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13575
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   13560
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmTransaksiPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnCostSharing As Boolean
Dim strKdKelPsn As String
Dim mstrNoCMku As String
Dim vbMsgboxRslt As String

Private Sub chkAlergi_Click()
 If chkAlergi.Value = 1 Then
    cmdEditData.Enabled = False
    cmdUbahOA.Enabled = False
    cmdHapusDataPOA.Enabled = False
    cmdTambahPOA.Enabled = False
    cmdRiwayatResep.Enabled = False
    dgAlergi.Visible = True
    
    dgAlergi.Top = 840
    dgAlergi.Left = 240
    
    Call subLoadDataRiwayatAlergiPasien
  
    
   'frmDataAlergiPasien.Show
  Else
    cmdEditData.Enabled = True
    cmdUbahOA.Enabled = True
    cmdHapusDataPOA.Enabled = True
    cmdTambahPOA.Enabled = True
    cmdRiwayatResep.Enabled = True
   ' Unload frmDataAlergiPasien
    dgAlergi.Visible = False
    
  End If
End Sub



Private Sub chkTampil_Click()
    If chkTampil.Value = 1 Then
        Call subLoadRiwayatDiagnosa(True)
    ElseIf chkTampil.Value = 0 Then
        Call subLoadRiwayatDiagnosa(False)
    End If
End Sub

Private Sub chkRP_Click()
    If chkRP.Value = 0 Then
        subLoadRiwayatPemeriksaan False
    Else
        subLoadRiwayatPemeriksaan True
    End If
End Sub

Private Sub cmdBatalEdit_Click()
    txtKdBarangEdit.Text = "": txtKdAsalEdit.Text = "": txtNamaBarangEdit.Text = "": txtJmlBarangEditAwal.Text = "": txtJmlBarangEditBaru.Text = ""
    picEditQuanttyBarang.Visible = False
End Sub

Private Sub cmdCetakDiagnosa_Click()
    On Error GoTo Jasmed
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frm_cetak_info_diag_viewer.Show
Jasmed:
End Sub

Private Sub cmdCetakHasilPemeriksaan_Click()
    On Error GoTo errLoad
    Dim pesan As VbMsgBoxResult

    If dgHasilPemeriksaan.ApproxCount = 0 Then Exit Sub
    
    pesan = MsgBox("Apakah anda ingin langsung mencetak laporan? " & vbNewLine & "Pilih No jika ingin ditampilkan terlebih dahulu ", vbQuestion + vbYesNo, "Konfirmasi")
    vLaporan = ""
    If pesan = vbYes Then vLaporan = "Print"

    cmdCetakHasilPemeriksaan.Enabled = False
    strSQL = "SELECT * FROM V_RiwayatHasilPemeriksaan" & _
    " WHERE NoCM = '" & mstrNoCM & "'"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount = 0 Then
        cmdCetakHasilPemeriksaan.Enabled = True
        Exit Sub
    End If

    mstrNoLabRad = dgHasilPemeriksaan.Columns("NoLab_Rad").Value

    Select Case dgHasilPemeriksaan.Columns("KdInstalasi").Value
        Case "09" 'lab pk
            strSQL = "select NoVerifikasi from HasilPemeriksaan where NoLab_Rad='" & dgHasilPemeriksaan.Columns("NoLab_Rad") & "' and NoPendaftaran='" & txtnopendaftaran.Text & "'"
            Call msubRecFO(rs, strSQL)
            If IsNull(rs(0)) Then
                MsgBox "Data hasil belum di verifikasi..", vbCritical, "Validasi": GoTo lanjut
            End If
            
            Set frmcetakhasillab = Nothing
            strQuery = "sELECT * from V_CetakHasilLaboratoriumPK WHERE NoLaboratorium = '" & mstrNoLabRad & "'"
            Call msubRecFO(dbRst, strQuery)
            If dbRst.EOF = False Then
                frmcetakhasillab.Show
            Else
               MsgBox "Tidak ada data yang di tampilkan.", vbInformation, "Informasi"
            End If

        Case "16" 'lab pa
            strSQL = "select NoVerifikasi from HasilPemeriksaan where NoLab_Rad='" & dgHasilPemeriksaan.Columns("NoLab_Rad") & "' and NoPendaftaran='" & txtnopendaftaran.Text & "'"
            Call msubRecFO(rs, strSQL)
            If IsNull(rs(0)) Then
                MsgBox "Data hasil belum di verifikasi..", vbCritical, "Validasi": GoTo lanjut
            End If
           
           Set frmCetakHasilLabPA = Nothing
            strQuery = "SELECT * " & _
                     " from V_CetakHasilPeriksaLaboratoryPA " & _
                     " WHERE NoLaboratorium = '" & mstrNoLabRad & "'"
            Call msubRecFO(dbRst, strQuery)
            If dbRst.EOF = False Then
                frmCetakHasilLabPA.Show
            Else
               MsgBox "Tidak ada data yang di tampilkan.", vbInformation, "Informasi"
            End If
            
'            frmCetakHasilLabPA.Show

        Case "10" 'radiologi
             strSQL = "select NoVerifikasi from HasilPemeriksaan where NoLab_Rad='" & dgHasilPemeriksaan.Columns("NoLab_Rad") & "' and NoPendaftaran='" & txtnopendaftaran.Text & "'"
            Call msubRecFO(rs, strSQL)
            If IsNull(rs(0)) Then
                MsgBox "Data hasil belum di verifikasi..", vbCritical, "Validasi": GoTo lanjut
            End If
       
            Set frmCetakHasilRadiologi = Nothing
            strSQL = "SELECT distinct NoRadiology,NoPendaftaran,NoCM,NamaPasien,TglHasil,Umur,AlamatLengkap,RuanganPerujuk,AsalPasien, " & _
                        " JenisKelamin,DokterPerujuk,NamaDetailPeriksa,NamaPelayanan,MemoHasilPeriksa,Catatan  " & _
                        " from V_CetakHasilPemeriksaanRadiology WHERE NoRadiology = '" & mstrNoLabRad & "'"
            Call msubRecFO(dbRst, strSQL)
            If dbRst.EOF = False Then
                frmCetakHasilRadiologi.Show
            Else
               MsgBox "Tidak ada data yang di tampilkan.", vbInformation, "Informasi"
            End If
            
        Case Else
            Call subLoadDiagramOdonto
    End Select
lanjut:
    cmdCetakHasilPemeriksaan.Enabled = True

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdCetakResume_Click()
'If dgRiwayatPemeriksaan.ApproxCount = 0 Then Exit Sub
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmCetakDataRiwayatPemeriksaanPasien.cetakLangsung    '.show
End Sub

Private Sub cmdDelDiagnosa_Click()
    On Error GoTo errHapusData

    Dim vbMsgboxRslt As VbMsgBoxResult
    Dim dbcmd As New ADODB.Command

    If dgRiwayatDiagnosa.ApproxCount = 0 Then Exit Sub

    vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus diagnosa '" _
    & dgRiwayatDiagnosa.Columns("Diagnosa ICD 10").Value & "'" & vbNewLine _
    & "Dengan tanggal pelayanan '" & dgRiwayatDiagnosa.Columns("TglPeriksa").Value _
    & "'", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbNo Then Exit Sub

    'diagnosa utama hanya bisa di-replace
    If dgRiwayatDiagnosa.Columns(14).Value <> "05" Then
        sp_DelDiagnosa dbcmd
        subLoadRiwayatDiagnosa (False)
        MsgBox "Data dihapus...", vbInformation, "Informasi"
    Else
        MsgBox "Diagnosa Utama hanya bisa diganti", vbInformation, "Informasi"
        Exit Sub
    End If

    Exit Sub
errHapusData:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
End Sub

Private Sub cmdEditData_Click()
    On Error GoTo errLoad
    If dgObatAlkes.ApproxCount = 0 Then Exit Sub

    picEditQuanttyBarang.Left = (Me.Width - picEditQuanttyBarang.Width) / 2
    picEditQuanttyBarang.Top = (Me.Height - picEditQuanttyBarang.Height) / 4

    picEditQuanttyBarang.Visible = True
    txtKdBarangEdit.Text = ""
    txtKdAsalEdit.Text = ""

    txtKdBarangEdit.Text = dgObatAlkes.Columns("KdBarang")
    txtKdAsalEdit.Text = dgObatAlkes.Columns("KdAsal")
    txtNamaBarangEdit.Text = dgObatAlkes.Columns("NamaBarang")
    txtJmlBarangEditAwal.Text = dgObatAlkes.Columns("Jml")
    txtNoterima.Text = dgObatAlkes.Columns("NoTerima")

    With txtJmlBarangEditBaru
        .Text = 1
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdHapusCatataKlinis_Click()
    On Error GoTo errHapusData

    Dim vbMsgboxRslt As VbMsgBoxResult
    Dim dbcmd As New ADODB.Command

    If dgCatatanKlinis.ApproxCount = 0 Then Exit Sub

    If dgCatatanKlinis.Columns("Ruang Pemeriksaan").Value <> mstrNamaRuangan Then
        MsgBox "Catatan klinis yang didapat di ruangan lain tidak dapat dihapus di ruangan ini", vbCritical, "Validasi"
        Exit Sub
    End If
    vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus catatan klinis pasien '" _
    & dgCatatanKlinis.Columns("NoPendaftaran").Value & "'" & vbNewLine _
    & "Dengan tanggal periksa '" & dgCatatanKlinis.Columns("TglPeriksa").Value _
    & "'", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbNo Then Exit Sub

    sp_DelBiayaCatatanKlinis dbcmd
    Call subLoadRiwayatCatatanKlinis
    MsgBox "Data dihapus...", vbInformation, "Informasi"

    Exit Sub
errHapusData:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
End Sub

Private Sub cmdHapusCatatanMedis_Click()
    On Error GoTo errHapusData

    Dim vbMsgboxRslt As VbMsgBoxResult
    Dim dbcmd As New ADODB.Command

    If dgCatatanMedis.ApproxCount = 0 Then Exit Sub

    If dgCatatanMedis.Columns("RuangPemeriksaan").Value <> mstrNamaRuangan Then
        MsgBox "Catatan medis yang didapat di ruangan lain tidak dapat dihapus di ruangan ini", vbCritical, "Validasi"
        Exit Sub
    End If
    vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus catatan medis pasien '" _
    & dgCatatanMedis.Columns("NoPendaftaran").Value & "'" & vbNewLine _
    & "Dengan tanggal periksa '" & dgCatatanMedis.Columns("TglPeriksa").Value _
    & "'", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbNo Then Exit Sub

    sp_DelBiayaCatatanMedis dbcmd
    sp_Alergi dbcmd
    Call subLoadRiwayatCatatanMedis
    MsgBox "Data dihapus...", vbInformation, "Informasi"

    Exit Sub
errHapusData:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
End Sub

Private Sub cmdHapusDataPOA_Click()
    On Error GoTo errHapusData

    Dim adoCommand As New ADODB.Command
    Dim vbMsgboxRslt As VbMsgBoxResult
    Dim i As Integer

    If dgObatAlkes.ApproxCount = 0 Then Exit Sub

    If dgObatAlkes.Columns("Status Bayar").Value = "Sudah DiBayar" Then
        MsgBox "Pemakaian Obat dan Alkes yang sudah dibayar tidak dapat dihapus", vbCritical, "Validasi"
        Exit Sub
    ElseIf dgObatAlkes.Columns("Ruang Pelayanan").Value <> mstrNamaRuangan Then
        MsgBox "Pemakaian Obat dan Alkes yang didapat di ruangan lain tidak dapat dihapus di ruangan ini", vbCritical
        Exit Sub
    End If
    vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus pemakaian obat dan alkes '" _
    & dgObatAlkes.Columns("NamaBarang").Value & "'" & vbNewLine _
    & "Dengan tanggal pelayanan '" & dgObatAlkes.Columns("TglPelayanan").Value _
    & "'", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbNo Then Exit Sub

    If bolStatusFIFO = True Then
        strSQL = "SELECT * FROM V_BiayaPemakaianObatAlkes WHERE NoPendaftaran='" & mstrNoPen & "' and KdBarang='" & dgObatAlkes.Columns("KdBarang") & "' " _
        & "and KdRuangan='" & mstrKdRuangan & "' and KdAsal='" & dgObatAlkes.Columns("KdAsal") & "' and SatuanJml='" & dgObatAlkes.Columns("Sat") & "' " _
        & "and tglPelayanan='" & Format(dgObatAlkes.Columns("TglPelayanan").Value, "yyyy/MM/dd HH:mm:ss") & "' Order by NoTerima Desc"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        If rs.EOF = False Then
            rs.MoveFirst
            For i = 1 To rs.RecordCount
                Set dbcmd = New ADODB.Command
                With dbcmd
                    .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
                    .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, rs("KdBarang"))
                    .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, rs("KdAsal"))
                    .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
                    .Parameters.Append .CreateParameter("Satuan", adChar, adParamInput, 1, rs("SatuanJml"))
                    .Parameters.Append .CreateParameter("JmlBrg", adInteger, adParamInput, , rs("JmlBarang"))
                    .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
                    .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(rs("TglPelayanan"), "yyyy/MM/dd HH:mm:ss"))
                    .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
                    .Parameters.Append .CreateParameter("NoTerima", adChar, adParamInput, 10, rs("NoTerima"))
                    .Parameters.Append .CreateParameter("ResepKe", adTinyInt, adParamInput, , rs("ResepKe"))

                    .ActiveConnection = dbConn
                    .CommandText = "dbo.Delete_PemakaianObatAlkes"
                    .CommandType = adCmdStoredProc
                    .Execute

                    If Not (.Parameters("RETURN_VALUE").Value = 0) Then
                        MsgBox "Ada Kesalahan dalam penghapusan data pemakaian obat dan alkes", vbCritical, "Validasi"
                        Exit Sub
                    End If
                    Call deleteADOCommandParameters(dbcmd)
                    Set dbcmd = Nothing
                End With

                rs.MoveNext
            Next i
            Call Add_HistoryLoginActivity("Delete_PemakaianObatAlkes")
        End If
    Else
        Set dbcmd = New ADODB.Command
        With dbcmd
            .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
            .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, dgObatAlkes.Columns("KdBarang").Value)
            .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, dgObatAlkes.Columns("KdAsal").Value)
            .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
            .Parameters.Append .CreateParameter("Satuan", adChar, adParamInput, 1, dgObatAlkes.Columns("Sat").Value)
            .Parameters.Append .CreateParameter("JmlBrg", adInteger, adParamInput, , dgObatAlkes.Columns("Jml").Value)
            .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
            .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(dgObatAlkes.Columns("TglPelayanan").Value, "yyyy/MM/dd HH:mm:ss"))
            .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
            .Parameters.Append .CreateParameter("NoTerima", adChar, adParamInput, 10, dgObatAlkes.Columns("NoTerima").Value)
            .Parameters.Append .CreateParameter("ResepKe", adTinyInt, adParamInput, , dgObatAlkes.Columns("ResepKe").Value)


            .ActiveConnection = dbConn
            .CommandText = "dbo.Delete_PemakaianObatAlkes"
            .CommandType = adCmdStoredProc
            .Execute

            If Not (.Parameters("RETURN_VALUE").Value = 0) Then
                MsgBox "Ada Kesalahan dalam penghapusan data pemakaian obat dan alkes", vbCritical, "Validasi"
                Exit Sub
            Else
                Call Add_HistoryLoginActivity("Delete_PemakaianObatAlkes")
            End If
            Call deleteADOCommandParameters(dbcmd)
            Set dbcmd = Nothing
        End With
    End If
    Call subPemakaianObatAlkes
    MsgBox "Data dihapus...", vbInformation, "Informasi"

    Exit Sub
errHapusData:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
''    Resume 0
End Sub


Private Sub cmdHapusDataPT_Click()
    On Error GoTo errHapusData

    Dim vbMsgboxRslt As VbMsgBoxResult
    Dim dbcmd As New ADODB.Command

    If dgTindakan.ApproxCount = 0 Then Exit Sub

    If dgTindakan.Columns("Status Bayar").Value = "Sudah DiBayar" Then
        MsgBox "Pelayanan yang sudah dibayar tidak dapat dihapus", vbCritical, "Validasi"
        Exit Sub
    ElseIf dgTindakan.Columns("Ruang Pelayanan").Value <> mstrNamaRuangan Then
        MsgBox "Pelayanan yang didapat di ruangan lain tidak dapat dihapus di ruangan ini", vbCritical
        Exit Sub
    End If
    vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus pelayanan '" _
    & dgTindakan.Columns("NamaPelayanan").Value & "'" & vbNewLine _
    & "Dengan tanggal pelayanan '" & dgTindakan.Columns("TglPelayanan").Value _
    & "'", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbNo Then Exit Sub

    sp_DelBiayaPelayanan dbcmd
    Call subLoadPelayananDidapat
    MsgBox "Data dihapus...", vbInformation, "Informasi"

    Exit Sub
errHapusData:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
End Sub

Private Sub cmdHapusKecelakaan_Click()
    On Error GoTo errHapusData

    Dim vbMsgboxRslt As VbMsgBoxResult
    Dim dbcmd As New ADODB.Command

    If dgKecelakaan.ApproxCount = 0 Then Exit Sub

    If dgKecelakaan.Columns("Ruangan Pemeriksa").Value <> mstrNamaRuangan Then
        MsgBox "Catatan klinis yang didapat di ruangan lain tidak dapat dihapus di ruangan ini", vbCritical, "Validasi"
        Exit Sub
    End If
    vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus data kecelakaan '" _
    & dgKecelakaan.Columns("NoPendaftaran").Value & "'" & vbNewLine _
    & "Dengan tanggal periksa '" & dgKecelakaan.Columns("TglPeriksa").Value _
    & "'", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbNo Then Exit Sub

    sp_DelKecelakaan dbcmd
    Call subLoadRiwayatKecelakaan
    MsgBox "Penghapusan data berhasil", vbInformation, "Informasi"

    Exit Sub
errHapusData:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
End Sub

Private Sub cmdHapusKonsul_Click()
    On Error GoTo errHapusData

    Dim vbMsgboxRslt As VbMsgBoxResult
    Dim dbcmd As New ADODB.Command

    If dgKonsul.ApproxCount = 0 Then Exit Sub
    If dgKonsul.Columns("StatusPeriksa").Value = "Sudah" Then
        MsgBox "Data rujukan yang sudah diperiksa tidak dapat dihapus", vbCritical, "Validasi"
        Exit Sub
    End If

    If dgKonsul.Columns("Ruangan Perujuk").Value <> mstrNamaRuangan Then
        MsgBox "Data rujukan yang didapat di ruangan lain tidak dapat dihapus di ruangan ini", vbCritical, "Validasi"
        Exit Sub
    End If
    vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus data konsul pasien '" _
    & dgKonsul.Columns("NoPendaftaran").Value & "'" & vbNewLine _
    & "Dengan tanggal periksa '" & dgKonsul.Columns("TglDirujuk").Value _
    & "'", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbNo Then Exit Sub
    
    StrSQL12 = "SELECT StrukOrder.NoOrder, StrukOrder.TglOrder, StrukOrder.KdRuangan, StrukOrder.KdRuanganTujuan, StrukOrder.KdSupplier, StrukOrder.IdUser," _
            & " DetailOrderPelayananTM.NoPendaftaran FROM StrukOrder INNER JOIN DetailOrderPelayananTM ON StrukOrder.NoOrder = DetailOrderPelayananTM.NoOrder" _
            & " WHERE StrukOrder.KdRuangan = '" & dgKonsul.Columns("KdRuanganAsal").Value & "' AND StrukOrder.KdRuanganTujuan = '" & dgKonsul.Columns("KdRuanganTujuan").Value & "' AND StrukOrder.TglOrder = '" & Format(dgKonsul.Columns("TglDirujuk").Value, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(rsL, StrSQL12)


    sp_DelKonsul dbcmd
    
    strSQL = "Delete DetailOrderPelayananTM where NoPendaftaran='" & txtnopendaftaran.Text & "' AND NoOrder ='" & rsL(0).Value & "' "
    Call msubRecFO(rs, strSQL)

    Call subLoadRiwayatKonsul
    MsgBox "Data dihapus...", vbInformation, "Informasi"

    Exit Sub
errHapusData:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
End Sub

Private Sub cmdHapusTM_Click()
    On Error GoTo errHapusData
    If dgRiwayatOperasi.ApproxCount = 0 Then Exit Sub
    vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus tindakan medis pasien '" _
    & dgRiwayatOperasi.Columns("NoHasilPeriksa").Value & "'" & vbNewLine _
    & "Dengan tanggal periksa '" & dgRiwayatOperasi.Columns("TglMulaiPeriksa").Value _
    & "'", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbNo Then Exit Sub

    dbConn.Execute "DELETE FROM DetailHasilTindakanMedisPasien WHERE NoHasilPeriksa = '" & dgRiwayatOperasi.Columns("NoHasilPeriksa") & "' and NoPendaftaran = '" & dgRiwayatOperasi.Columns("NoPendaftaran") & "'"
    dbConn.Execute "DELETE FROM HasilTindakanMedis WHERE NoHasilPeriksa = '" & dgRiwayatOperasi.Columns("NoHasilPeriksa") & "' and NoPendaftaran = '" & dgRiwayatOperasi.Columns("NoPendaftaran") & "'"

    MsgBox "Data dihapus...", vbInformation, "Informasi"
    Call subLoadRiwayatOperasi
    Exit Sub
errHapusData:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"

End Sub

Private Sub cmdICD9_Click()
    On Error GoTo hell
        
    If dgRiwayatDiagnosa.ApproxCount = 0 Then Exit Sub
    If dgRiwayatDiagnosa.Columns(0) <> txtnopendaftaran.Text Then
        MsgBox "No Pendaftaran tidak sama, mohon isi diagnosanya [ICD 10] dahulu", vbExclamation, "Validasi"
        cmdTambahDiagnosa.SetFocus
        Exit Sub
    End If
    If dgRiwayatDiagnosa.Columns("Ruang Periksa").Value <> mstrNamaRuangan Then
        MsgBox "Diagnosa ICD 9 yang didapat di ruangan lain tidak dapat mengedit Diagnosa di ruangan ini", vbCritical
        Exit Sub
    End If
    Set rs = Nothing
    rs.Open "Select KdJenisDiagnosa From JenisDiagnosa Where KdJenisDiagnosa = '" & dgRiwayatDiagnosa.Columns(14) & "'", dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.EOF = True Then
        MsgBox dgRiwayatDiagnosa.Columns(3) & " tidak terdapat di ruangan " & mstrNamaRuangan, vbExclamation, "Validasi"
        cmdTambahDiagnosa.SetFocus
        Exit Sub
    End If
    Me.Enabled = False
    mstrKdDiagnosa = ""
    mstrKdDiagnosa = dgRiwayatDiagnosa.Columns(4)
    mstrKdJenisDiagnosaTindakan = ""
    mstrKdJenisDiagnosaTindakan = dgRiwayatDiagnosa.Columns(15)
    bolEditDiagnosa = True
    With frmPeriksaDiagnosa
        .Show
        .txtnopendaftaran = txtnopendaftaran.Text
        .txtnocm = txtnocm.Text
        .txtNamaPasien = txtNamaPasien.Text
        .txtSex = txtSex.Text
'        If Left(.txtSex, 1) = "P" Then
'            .txtSex.Text = "Perempuan"
'        Else
'            .txtSex.Text = "Laki-laki"
'        End If
        .txtThn = txtThn.Text
        .txtBln = txtBln.Text
        .txtHari = txtHr.Text

'        strSQL = "SELECT dbo.RegistrasiIGD.IdDokter, dbo.DataPegawai.NamaLengkap " & _
'        " FROM dbo.RegistrasiIGD INNER JOIN dbo.DataPegawai ON dbo.RegistrasiIGD.IdDokter = dbo.DataPegawai.IdPegawai " & _
'        " WHERE (dbo.RegistrasiIGD.NoPendaftaran = '" & txtnopendaftaran.Text & "')"
'        Call msubRecFO(rs, strSQL)
'
'        If Not rs.EOF Then
'            .txtDokter.Text = rs(1).Value
'            mstrKdDokter = rs(0).Value
'            intJmlDokter = rs.RecordCount
'            .fraDokter.Visible = False
'        End If
        .txtDokter.Text = mstrNamaDokter
        .fraDokter.Visible = False

        .dtpTglPeriksa.Value = dgRiwayatDiagnosa.Columns(2)
        .dcJenisDiagnosa.BoundText = dgRiwayatDiagnosa.Columns(14)
        .dcJenisDiagnosaTindakan.BoundText = dgRiwayatDiagnosa.Columns(15)

        .dcJenisDiagnosa.Enabled = False
        .lvwDiagnosa.Enabled = False
        .txtNamaDiagnosa.Enabled = False
        .txtDokter.Enabled = False
        .dtpTglPeriksa.Enabled = False
        .chkICD9.Value = Checked

'        .Show
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdKehamilandanKB_Click()
    frmDataKehamilandanKB.Show
End Sub

Private Sub cmdRadiologi_Click()
    frmSuratKeterangan3C.Show
End Sub

Private Sub cmdRiwayatResep_Click()
    cmdEditData.Enabled = False
    cmdUbahOA.Enabled = False
    cmdHapusDataPOA.Enabled = False
    cmdTambahPOA.Enabled = False
    cmdRiwayatResep.Enabled = False
    cmdTutup.Enabled = False
    chkAlergi.Enabled = False
    'dgRiwayatResepPasien.Visible = True
    fraRiwayatResep.Visible = True
    
    fraRiwayatResep.Top = 840
    fraRiwayatResep.Left = 120
   Call subLoadDataRiwayatResepPasien
End Sub


Public Sub subLoadDataRiwayatResepPasien()
    On Error GoTo errLoad

    strSQL = "SELECT *" & _
        " FROM V_RiwayatReseppasien " & _
        " WHERE (nocm = '" & mstrNoCM & "')"
    Call msubRecFO(rs, strSQL)

    Set dgRiwayatResepPasien.DataSource = rs
    
    With dgRiwayatResepPasien
        .Columns(0).Width = 0 'NoCM
        .Columns(1).Width = 0 'NoPendaftaran
        .Columns(2).Width = 0 'NamaLengkap
        .Columns(3).Width = 1000 'NoResep
        .Columns(4).Width = 2500 'TglResep
        .Columns(5).Width = 2500 'Ruangan
        .Columns(6).Width = 3500 'DokterPenulisResep
        .Columns(7).Width = 3000 'NamaObat
    End With

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdSimpanEditBarang_Click()
    On Error GoTo errLoad

    If Val(txtJmlBarangEditBaru.Text) = 0 Then
        txtJmlBarangEditBaru.SetFocus: txtJmlBarangEditBaru.SelStart = 0: txtJmlBarangEditBaru.SelLength = Len(txtJmlBarangEditBaru.Text)
        Exit Sub
    End If
    If bolStatusFIFO = False Then
         strSQL = "SELECT JmlStok FROM StokRuangan WHERE (KdRuangan = '" & mstrKdRuangan & "') AND (KdBarang = '" & txtKdBarangEdit.Text & "') AND (KdAsal = '" & txtKdAsalEdit.Text & "')"
       Else
         strSQL = "SELECT JmlStok FROM StokRuanganfIFO WHERE (KdRuangan = '" & mstrKdRuangan & "') AND (KdBarang = '" & txtKdBarangEdit.Text & "') AND (KdAsal = '" & txtKdAsalEdit.Text & "')AND (Noterima = '" & txtNoterima.Text & "')"
    End If
    Call msubRecFO(rs, strSQL)

    If Not rs.EOF Then
        If rs(0).Value < Val(txtJmlBarangEditBaru.Text) Then
            MsgBox "Jumlah Barang melebihi stok, stok barang (" & rs(0).Value & ")", vbExclamation, "Validasi"
            txtJmlBarangEditBaru.Text = rs(0).Value: txtJmlBarangEditBaru.SetFocus: txtJmlBarangEditBaru.SelStart = 0: txtJmlBarangEditBaru.SelLength = Len(txtJmlBarangEditBaru.Text)
            Exit Sub
        End If

        strSQL = "SELECT KdJenisObat, ResepKe " & _
        " From PemakaianAlkes " & _
        " Where (NoPendaftaran = '" & txtnopendaftaran.Text & "' ) And (KdRuangan = '" & mstrKdRuangan & "') And (KdBarang = '" & txtKdBarangEdit.Text & "' ) And (KdAsal = '" & txtKdAsalEdit.Text & "') And (TglPelayanan = '" & Format(dgObatAlkes.Columns("TglPelayanan"), "yyyy/MM/dd HH:mm:ss") & "') And (SatuanJml = '" & dgObatAlkes.Columns("Sat") & "')"
        Call msubRecFO(rs, strSQL)
        
        If rs(0).Value = Null Then Exit Sub
        If rs.EOF = True Then Exit Sub
        
        

        If sp_EditQtyBarang(txtKdBarangEdit.Text, txtKdAsalEdit.Text, dgObatAlkes.Columns("Sat"), _
            Val(txtJmlBarangEditAwal.Text), Val(txtJmlBarangEditBaru.Text), _
            dgObatAlkes.Columns("TglPelayanan"), rs(0), rs(1), txtNoterima.Text) = False Then Exit Sub

            picEditQuanttyBarang.Visible = False
            Call subPemakaianObatAlkes
            dgObatAlkes.SetFocus
    End If

        Exit Sub
errLoad:
        Call msubPesanError
'        Resume 0
End Sub

Private Sub cmdSpesialis_Click()
   Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan) 'mstrKdRuanganPasien)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, "155002")
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, "02")
        .Parameters.Append .CreateParameter("StatusCITO", adChar, adParamInput, 1, "0")
        .Parameters.Append .CreateParameter("Tarif", adInteger, adParamInput, , 12000)
        .Parameters.Append .CreateParameter("JmlPelayanan", adInteger, adParamInput, , 1)
        '.Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(dtTanggalPelayanan, "yyyy/MM/dd HH:mm:ss"))
         .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(Now, "yyyy/MM/dd HH:mm:ss"))
         'untuk nambah 1 detik
'         penamwaktu = DateAdd("s", 1, penamwaktu)
         
        .Parameters.Append .CreateParameter("NoLab_Rad", adChar, adParamInput, 10, Null)
        
        Set rs = Nothing
        strSQL = "SELECT IdDokter FROM RegistrasiRJ WHERE NoPendaftaran='" & txtnopendaftaran.Text & "'"
        Call msubRecFO(rs, strSQL)
        
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, rs(0).Value)
        .Parameters.Append .CreateParameter("StatusAPBD", adChar, adParamInput, 2, "01")
        .Parameters.Append .CreateParameter("KdJenisTarif", adChar, adParamInput, 2, "01")
        .Parameters.Append .CreateParameter("TarifCito", adInteger, adParamInput, , 0)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("IdPegawai2", adChar, adParamInput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_BiayaPelayanan"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Biaya Pelayanan Pasien", vbCritical, "Validasi"
            Call deleteADOCommandParameters(dbcmd)
            Set dbcmd = Nothing
'            sp_BiayaPelayanan = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
'        sp_BiayaPelayanan = True
    End With
    
    Call subLoadPelayananDidapat
End Sub

Private Sub cmdTambahCatatanKlinis_Click()
    If txtnocm.Text = "" Then Exit Sub
    frmCatatanKlinisPasien.Show
    frmTransaksiPasien.Enabled = False
End Sub

Private Sub cmdTambahCatatanMedis_Click()
    If txtnocm.Text = "" Then Exit Sub
    frmTransaksiPasien.Enabled = False
    With frmCatatanMedikPasien
        strSQL = "SELECT dbo.RegistrasiIGD.IdDokter, dbo.DataPegawai.NamaLengkap " & _
        " FROM dbo.RegistrasiIGD INNER JOIN dbo.DataPegawai ON dbo.RegistrasiIGD.IdDokter = dbo.DataPegawai.IdPegawai " & _
        " WHERE (dbo.RegistrasiIGD.NoPendaftaran = '" & txtnopendaftaran.Text & "')"
        Call msubRecFO(rs, strSQL)

        If Not rs.EOF Then
            .txtDokter.Text = rs(1).Value
            mstrKdDokter = rs(0).Value
            intJmlDokter = rs.RecordCount
            .fraDokter.Visible = False
        End If
        .Show
    End With
End Sub

Private Sub cmdTambahDiagnosa_Click()
    On Error GoTo errLoad

    Me.Enabled = False
    With frmPeriksaDiagnosa
        .Show
        .txtnopendaftaran = txtnopendaftaran.Text
        .txtnocm = txtnocm.Text
        .txtNamaPasien = txtNamaPasien.Text
        .txtSex = txtSex.Text
        .txtThn = txtThn.Text
        .txtBln = txtBln.Text
        .txtHari = txtHr.Text

'        strSQL = "SELECT dbo.RegistrasiIGD.IdDokter, dbo.DataPegawai.NamaLengkap " & _
'        " FROM dbo.RegistrasiIGD INNER JOIN dbo.DataPegawai ON dbo.RegistrasiIGD.IdDokter = dbo.DataPegawai.IdPegawai " & _
'        " WHERE (dbo.RegistrasiIGD.NoPendaftaran = '" & txtnopendaftaran.Text & "')"
'        Call msubRecFO(rs, strSQL)
'
'        If Not rs.EOF Then
'            .txtDokter.Text = rs(1).Value
'            mstrKdDokter = rs(0).Value
'            intJmlDokter = rs.RecordCount
'            .fraDokter.Visible = False
'        End If
        .txtDokter.Text = mstrNamaDokter
        .fraDokter.Visible = False
'        .Show
    End With

    Exit Sub
errLoad:
    Call msubPesanError
    frmPeriksaDiagnosa.Show
End Sub

Private Sub cmdTambahKecelakaan_Click()
    If txtnocm.Text = "" Then Exit Sub
End Sub
'
'Private Sub cmdTambahKonsul_Click()
'    On Error GoTo errLoad
'
'    If txtNoCM.Text = "" Then Exit Sub
'
''    frmPasienRujukan.Show
''    With frmPasienRujukan
''        .txtNoPendaftaran.Text = txtNoPendaftaran.Text
''        .txtNoCM.Text = txtNoCM.Text
''        .txtNamaPasien.Text = txtNamaPasien.Text
''        .txtSex.Text = txtSex.Text
''        .txtThn.Text = txtThn.Text
''        .txtBln.Text = txtBln.Text
''        .txtHari.Text = txtHr.Text
''        .dtpTglDirujuk.Value = Now
''        strSQL = "SELECT KdSubInstalasi, IdDokter, Dokter FROM V_DokterPasien WHERE NoPendaftaran = '" & txtNoPendaftaran.Text & "'"
''        Call msubRecFO(rs, strSQL)
''        If rs.EOF = False Then
''            mstrKdSubInstalasi = rs("KdSubInstalasi").Value
''            mstrKdDokter = rs("IdDokter").Value
''            intJmlDokter = rs.RecordCount
''            frmPasienRujukan.txtDokter.Text = rs("Dokter").Value
''            frmPasienRujukan.fraDokter.Visible = False
''        End If
''    End With
''
''    Me.Enabled = False
''    frmPasienRujukan.Show
'    frmOrderPelayananKonsul.Show
'    With frmOrderPelayananKonsul
'        .txtNoPendaftaranTM.Text = txtNoPendaftaran.Text
'        .txtNoCMTM.Text = txtNoCM.Text
'        .dtpTglOrderTM.Value = Now
'        strSQL = "SELECT KdSubInstalasi, IdDokter, Dokter FROM V_DokterPasien WHERE NoPendaftaran = '" & txtNoPendaftaran.Text & "'"
'        Call msubRecFO(rs, strSQL)
'        If rs.EOF = False Then
'            mstrKdSubInstalasi = rs("KdSubInstalasi").Value
'            mstrKdDokter = rs("IdDokter").Value
'            intJmlDokter = rs.RecordCount
'        End If
'    End With
'
'    Me.Enabled = False
'    frmOrderPelayananKonsul.Show
'    Exit Sub
'errLoad:
'    Call msubPesanError
'    frmOrderPelayananKonsul.Show
'End Sub

Private Sub cmdTambahPemakaianBahan_Click()
    With frmPemakaianAlkesNonCharge
        .Show
        .txtnopendaftaran.Text = txtnopendaftaran.Text
        .txtnocm.Text = txtnocm.Text
        .txtNamaPasien.Text = txtNamaPasien.Text
        .txtJK.Text = txtSex.Text
        .txtThn.Text = txtThn.Text
        .txtBln.Text = txtBln.Text
        .txtHr.Text = txtHr.Text
    End With
End Sub

Private Sub cmdTambahKonsul_Click()
On Error GoTo Gabril
    fraOrder.Visible = True
    optTindakan.Value = False
    optObat.Value = False
Exit Sub
Gabril:
    Call msubPesanError
End Sub

Private Sub cmdTambahPOA_Click()
    On Error GoTo errLoad
    Me.Enabled = False
'    frmPemakaianObatAlkes2.Show
'    frmPemakaianObatAlkes2.txtRP.Text = mstrNamaRuangan
    frmPemakaianObatAlkes.NotUseRacikan = True
    frmPemakaianObatAlkes.Show
    frmPemakaianObatAlkes.txtRP.Text = mstrNamaRuangan

    Exit Sub
errLoad:
    Call msubPesanError
    frmPemakaianObatAlkes.Show
End Sub

Private Sub cmdTambahPT_Click()
    On Error GoTo errLoad

    strSQL = "SELECT dbo.RegistrasiIGD.IdDokter, dbo.DataPegawai.NamaLengkap " & _
    " FROM dbo.RegistrasiIGD INNER JOIN dbo.DataPegawai ON dbo.RegistrasiIGD.IdDokter = dbo.DataPegawai.IdPegawai " & _
    " WHERE (dbo.RegistrasiIGD.NoPendaftaran = '" & txtnopendaftaran.Text & "')"
    Call msubRecFO(rs, strSQL)

    If Not rs.EOF Then
        frmTindakan.txtDokter.Text = rs(1).Value
        mstrKdDokter = rs(0).Value
        intJmlDokter = rs.RecordCount
        frmTindakan.fraDokter.Visible = False
    Else
        mstrKdDokter = ""
        intJmlDokter = 0
    End If

    frmTindakan.Show

    Exit Sub
errLoad:
    Call msubPesanError
    frmTindakan.Show
End Sub

Private Sub cmdTambahTM_Click()

    On Error GoTo errTO
    Me.Enabled = False
    mstrNoPen = txtnopendaftaran.Text
    mstrNoCM = txtnocm.Text

    With frmTindakanMedisPasien
        .Show
        .txtnopendaftaran.Text = mstrNoPen
        .txtnocm.Text = mstrNoCM
        .txtNamaPasien.Text = txtNamaPasien.Text
        If Left(txtSex.Text, 1) = "P" Then
            .txtSex.Text = "Perempuan"
        Else
            .txtSex.Text = "Laki-Laki"
        End If
        .txtThn.Text = txtThn.Text
        .txtBln.Text = txtBln.Text
        .txtHr.Text = txtHr.Text
        strSQL = "Select NamaSubInstalasi From SubInstalasi Where KdSubInstalasi='" & mstrKdSubInstalasi & "'"
        Set rs = Nothing
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then .txtSubInstalasi.Text = rs(0)
    End With
    Exit Sub
errTO:
    Call msubPesanError
    Me.Enabled = True
End Sub

Private Sub cmdTotalTanggunganPenjamin_Click()
On Error GoTo hell
' If (Dir("C:\Lib\Ina Cbg\result.tlb") <> "") Then
'        Dim context As BridgingInaCbg.context
'        strSQL = "select value from SettingGlobal where Prefix='UrlInaCbg'"
'        Call msubRecFO(rs, strSQL)
'        If (rs.EOF = False) Then
'            Set context = New BridgingInaCbg.context
'            context.SetEndpoint (rs(0).Value)
'            strSQL = "select * from pasien where noCm='" & mstrNoCM & "'"
'            Call msubRecFO(rsC, strSQL)
'            strSQL = "SELECT     PeriksaDiagnosa.KdDiagnosa FROM PeriksaDiagnosa INNER JOIN SettingGlobal ON PeriksaDiagnosa.KdJenisDiagnosa = SettingGlobal.Value WHERE     (PeriksaDiagnosa.NoPendaftaran = '" & mstrNoPen & "') AND (SettingGlobal.Prefix = 'KdDiagnosaUtama' )"
'            Call msubRecFO(rsD, strSQL)
'            Dim diagnosa As String
'            Dim i As Integer
'            For i = 1 To rsD.RecordCount
'                If (diagnosa = "") Then
'                    diagnosa = rsD(0).Value
'                Else
'                    diagnosa = rsD(0).Value & ";"
'                End If
'                rsD.MoveNext
'            Next i
'
'            strSQL = "SELECT     PeriksaDiagnosa.KdDiagnosa FROM PeriksaDiagnosa INNER JOIN SettingGlobal ON PeriksaDiagnosa.KdJenisDiagnosa = SettingGlobal.Value WHERE     (PeriksaDiagnosa.NoPendaftaran = '" & mstrNoPen & "') AND (SettingGlobal.Prefix = 'KdDiagnosaTambahan' )"
'            Call msubRecFO(rsD, strSQL)
'
'            For i = 1 To rsD.RecordCount
'                diagnosa = diagnosa & ";" & rsD(0).Value
'                rsD.MoveNext
'            Next i
'
'             strSQL = "SELECT     DetailPeriksaDiagnosa.KdDiagnosaTindakan FROM DetailPeriksaDiagnosa INNER JOIN SettingGlobal ON DetailPeriksaDiagnosa.KdJenisDiagnosa = SettingGlobal.Value WHERE     (DetailPeriksaDiagnosa.NoPendaftaran = '" & mstrNoPen & "') AND (SettingGlobal.Prefix = 'KdDiagnosaUtama' )"
'            Call msubRecFO(rsD, strSQL)
'            Dim diagnosaTindakan  As String
'            For i = 1 To rsD.RecordCount
'                diagnosaTindakan = diagnosaTindakan & rsD(0).Value + ";"
'                rsD.MoveNext
'            Next i
'
'            strSQL = "SELECT     DetailPeriksaDiagnosa.KdDiagnosaTindakan FROM DetailPeriksaDiagnosa INNER JOIN SettingGlobal ON DetailPeriksaDiagnosa.KdJenisDiagnosa = SettingGlobal.Value WHERE     (DetailPeriksaDiagnosa.NoPendaftaran = '" & mstrNoPen & "') AND (SettingGlobal.Prefix = 'KdDiagnosaTambahan' )"
'            Call msubRecFO(rsD, strSQL)
'
'            For i = 1 To rsD.RecordCount
'                diagnosaTindakan = diagnosaTindakan & rsD(0).Value + ";"
'                rsD.MoveNext
'            Next i
'            'context.SimulasiTarif(
'            Dim hasilTarif As String
'            hasilTarif = context.SimulasiTarif(IIf(rsC("JenisKelamin").Value = "L", "m", "f"), "5", "2222", "444", "rawat jalan", "", "2014-01-04", "2014-01-04", "home", "coba", "20000", diagnosa, diagnosaTindakan, "", "", "joko", Format(rsC("TglLahir").Value, "yyyy-MM-dd"))
'            txtTotalPenjamin.Text = FormatPembulatan(CDbl(hasilTarif), mstrKdInstalasiLogin)
'
'        End If
'
'    End If
'    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdtutup_Click()
    Unload Me
End Sub

Private Sub cmdCetakCatatanMedis_Click()
On Error GoTo hell
    If dgCatatanMedis.ApproxCount = 0 Then Exit Sub
'    cmdCetakCatatanMedis.Enabled = False
        vLaporan = ""
        If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmCetakCatatanMedis.Show
'    cmdCetakCatatanMedis.Enabled = True
    Exit Sub
hell:
    cmdCetakCatatanMedis.Enabled = True

End Sub

Private Sub cmdTutup2_Click()
  fraRiwayatResep.Visible = False
  cmdEditData.Enabled = True
  cmdUbahOA.Enabled = True
  cmdHapusDataPOA.Enabled = True
  cmdTambahPOA.Enabled = True
  cmdRiwayatResep.Enabled = True
  cmdTutup.Enabled = True
  chkAlergi.Enabled = True
End Sub

Private Sub cmdUbahCatatanMedis_Click()
If dgCatatanMedis.ApproxCount = 0 Then Exit Sub
If dgCatatanMedis.Columns("RuangPemeriksaan").Value <> mstrNamaRuangan Then
        MsgBox "Catatan medis yang didapat di ruangan lain tidak dapat diubah di ruangan ini", vbCritical, "Validasi"
        Exit Sub
    End If
Me.Enabled = False
With frmCatatanMedikPasien
    .BoolUbahAlergiCatatanMedikPasien = True
    .Show
    
    .dtpTglPeriksa.Value = dgCatatanMedis.Columns("TglPeriksa").Value
    
    If dgCatatanMedis.Columns("KeluhanUtama") = "" Or IsNull(dgCatatanMedis.Columns("KeluhanUtama")) Then
        .txtKeluhanUtama.Text = ""
    Else
        .txtKeluhanUtama.Text = dgCatatanMedis.Columns("KeluhanUtama").Value
    End If
    If dgCatatanMedis.Columns("Pengobatan") = "" Or IsNull(dgCatatanMedis.Columns("Pengobatan")) Then
        .txtPengobatan.Text = ""
    Else
        .txtPengobatan.Text = dgCatatanMedis.Columns("Pengobatan").Value
    End If
    If dgCatatanMedis.Columns("Keterangan") = "" Or IsNull(dgCatatanMedis.Columns("Keterangan")) Then
        .txtKeterangan.Text = ""
    Else
        .txtKeterangan.Text = dgCatatanMedis.Columns("Keterangan").Value
    End If
    
    If dgCatatanMedis.Columns("KdTriase") = "" Or IsNull(dgCatatanMedis.Columns("KdTriase")) Then
        .dcTriase.BoundText = ""
    Else
       .dcTriase.BoundText = dgCatatanMedis.Columns("KdTriase").Value
        
    End If
    If dgCatatanMedis.Columns("KdImunisasi") = "" Or IsNull(dgCatatanMedis.Columns("KdImunisasi")) Then
        .dcImunisasi.BoundText = ""
    Else
       .dcImunisasi.BoundText = dgCatatanMedis.Columns("KdImunisasi").Value
        
    End If
    If dgCatatanMedis.Columns("KdAlergi") = "" Or IsNull(dgCatatanMedis.Columns("KdAlergi")) Then
        .dcAlergi.BoundText = ""
    Else
       .dcAlergi.BoundText = dgCatatanMedis.Columns("KdAlergi").Value
        
    End If
    If dgCatatanMedis.Columns("Catatan Medis Keluarga") = "" Or IsNull(dgCatatanMedis.Columns("Catatan Medis Keluarga")) Then
        .txtCatatanMedisKeluarga.Text = ""
    Else
        .txtCatatanMedisKeluarga.Text = dgCatatanMedis.Columns("Catatan Medis Keluarga").Value
    End If

    
'    If dgCatatanMedis.Columns("IdParamedis") = "" Or IsNull(dgCatatanMedis.Columns("IdParamedis")) Then
'        .dcPerawat.BoundText = ""
'    Else
'        .dcPerawat.BoundText = dgCatatanMedis.Columns("IdParamedis").Value
'    End If
    
    .txtnopendaftaran = dgCatatanMedis.Columns("NoPendaftaran").Value
        .txtnocm = dgCatatanMedis.Columns("NoCM").Value
        .txtNamaPasien = txtNamaPasien.Text
        .txtSex.Text = txtSex.Text
        .txtThn = txtThn.Text
        .txtBln = txtBln.Text
        .txtHari = txtHr.Text
    

    strSQL = "SELECT  IdPegawai, [Nama Pemeriksa]" & _
    " FROM V_DaftarPemeriksaPasien " & _
    " WHERE (IdPegawai = '" & dgCatatanMedis.Columns("IdDokter").Value & "')"
    Call msubRecFO(dbRst, strSQL)
    If rs.EOF = False Then
        .txtDokter.Text = dbRst(1).Value
        mstrKdDokter = dbRst(0).Value
    Else
        mstrKdDokter = ""
        .txtDokter.Text = ""
    End If
    .fraDokter.Visible = False

End With
End Sub

Private Sub cmdUbahOA_Click()
    On Error GoTo errLoad

    If dgObatAlkes.ApproxCount = 0 Then Exit Sub
    If dgObatAlkes.Columns("Status Bayar").Value = "Sudah DiBayar" Then
        MsgBox "Pelayanan yang sudah dibayar tidak dapat diubah", vbCritical, "Validasi"
        Exit Sub
    ElseIf dgObatAlkes.Columns("Ruang Pelayanan").Value <> mstrNamaRuangan Then
        MsgBox "Pelayanan yang didapat di ruangan lain tidak dapat diubah di ruangan ini", vbCritical
        Exit Sub
    End If
'    frmUpdateBiayaPelayananOA.Show
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdUbahPT_Click()
    On Error GoTo errLoad

    If dgTindakan.ApproxCount = 0 Then Exit Sub
    If dgTindakan.Columns("Status Bayar").Value = "Sudah DiBayar" Then
        MsgBox "Pelayanan yang sudah dibayar tidak dapat diubah", vbCritical, "Validasi"
        Exit Sub
    ElseIf dgTindakan.Columns("Ruang Pelayanan").Value <> mstrNamaRuangan Then
        MsgBox "Pelayanan yang didapat di ruangan lain tidak dapat diubah di ruangan ini", vbCritical
        Exit Sub
    End If
    With frmUpdateBiayaPelayanan
        .txtnopendaftaran = txtnopendaftaran.Text
        strKodePelayananRS = dgTindakan.Columns(12).Value
        .txtnocm.Text = txtnocm.Text
        .txtNamaPasien.Text = txtNamaPasien.Text
        .txtJK.Text = txtSex.Text
        .txtThn.Text = txtThn.Text
        .txtBln.Text = txtBln.Text
        .txtHr.Text = txtHr.Text

        Call .txtNoPendaftaran_KeyPress(13)
    End With

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdUmum_Click()
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan) 'mstrKdRuanganPasien)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, "153012")
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, "02")
        .Parameters.Append .CreateParameter("StatusCITO", adChar, adParamInput, 1, "0")
        .Parameters.Append .CreateParameter("Tarif", adInteger, adParamInput, , 6000)
        .Parameters.Append .CreateParameter("JmlPelayanan", adInteger, adParamInput, , 1)
        '.Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(dtTanggalPelayanan, "yyyy/MM/dd HH:mm:ss"))
         .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(Now, "yyyy/MM/dd HH:mm:ss"))
         'untuk nambah 1 detik
'         penamwaktu = DateAdd("s", 1, penamwaktu)
         
        .Parameters.Append .CreateParameter("NoLab_Rad", adChar, adParamInput, 10, Null)
        
        Set rs = Nothing
        strSQL = "SELECT IdDokter FROM RegistrasiRJ WHERE NoPendaftaran='" & txtnopendaftaran.Text & "'"
        Call msubRecFO(rs, strSQL)
        
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, rs(0).Value)
        .Parameters.Append .CreateParameter("StatusAPBD", adChar, adParamInput, 2, "01")
        .Parameters.Append .CreateParameter("KdJenisTarif", adChar, adParamInput, 2, "01")
        .Parameters.Append .CreateParameter("TarifCito", adInteger, adParamInput, , 0)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("IdPegawai2", adChar, adParamInput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_BiayaPelayanan"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Biaya Pelayanan Pasien", vbCritical, "Validasi"
            Call deleteADOCommandParameters(dbcmd)
            Set dbcmd = Nothing
'            sp_BiayaPelayanan = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
'        sp_BiayaPelayanan = True
    End With
    
    Call subLoadPelayananDidapat
End Sub

Private Sub dgCatatanKlinis_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgCatatanKlinis
    WheelHook.WheelHook dgCatatanKlinis
End Sub

Private Sub dgCatatanKlinis_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then cmdTambahCatatanKlinis.SetFocus
End Sub

Private Sub dgCatatanMedis_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgCatatanMedis
    WheelHook.WheelHook dgCatatanMedis
End Sub

Private Sub dgCatatanMedis_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then cmdTambahCatatanMedis.SetFocus
End Sub

Private Sub dgHasilPemeriksaan_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgHasilPemeriksaan
    WheelHook.WheelHook dgHasilPemeriksaan
End Sub

Private Sub dgKecelakaan_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgKecelakaan
    WheelHook.WheelHook dgKecelakaan
End Sub

Private Sub dgKecelakaan_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then cmdTambahKecelakaan.SetFocus
End Sub

Private Sub dgKonsul_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgKonsul
    WheelHook.WheelHook dgKonsul
End Sub

Private Sub dgKonsul_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then cmdTambahKonsul.SetFocus
End Sub

Private Sub dgObatAlkes_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgObatAlkes
    WheelHook.WheelHook dgObatAlkes

    cmdUbahOA.Enabled = False
End Sub

Private Sub dgObatAlkes_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then cmdTambahPOA.SetFocus
End Sub

Private Sub dgRiwayatDiagnosa_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgRiwayatDiagnosa
    WheelHook.WheelHook dgRiwayatDiagnosa
End Sub

Private Sub dgRiwayatDiagnosa_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then cmdTambahDiagnosa.SetFocus
End Sub

Private Sub dgRiwayatOperasi_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgRiwayatOperasi
    WheelHook.WheelHook dgRiwayatOperasi
End Sub

Private Sub dgRiwayatOperasi_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then dgRiwayatOperasi.SetFocus
End Sub

Private Sub dgRiwayatPemeriksaan_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgRiwayatPemeriksaan
    WheelHook.WheelHook dgRiwayatPemeriksaan
End Sub

Private Sub dgRiwayatPemeriksaan_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then cmdCetakRP.SetFocus
End Sub

Private Sub dgTindakan_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgTindakan
    WheelHook.WheelHook dgTindakan
End Sub

Private Sub dgTindakan_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then cmdTambahPT.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errLoad

    Dim strCtrlKey As String
    Dim strCtrlKeyAlt As String
    
    strCtrlKey = (Shift + vbCtrlMask)
    
    
    
    Select Case KeyCode
        Case vbKey1
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 0
        Case vbKey2
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 1
        Case vbKey3
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 2
        Case vbKey4
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 3
        Case vbKey5
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 4
        Case vbKey6
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 5
        Case vbKey7
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 6
        Case vbKey8
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 7
        Case vbKey9
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 8
        Case vbKey0
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 9
        Case vbKeyF5
            Call subLoadPelayananDidapat
            Call subPemakaianObatAlkes
            Call subLoadRiwayatCatatanKlinis
            Call subLoadRiwayatCatatanMedis
            Call subLoadRiwayatDiagnosa(False)
            Call subLoadRiwayatKecelakaan
            Call subLoadRiwayatOperasi
            Call subLoadRiwayatKonsul
            Call subLoadRiwayatPemeriksaan(False)
            Call subLoadRiwayatHasilPemeriksaan
    End Select

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call subLoadPelayananDidapat
    Call subPemakaianObatAlkes
    Call subLoadRiwayatCatatanKlinis
    Call subLoadRiwayatCatatanMedis
    Call subLoadRiwayatDiagnosa(False)
    Call subLoadRiwayatKecelakaan
    Call subLoadRiwayatOperasi
    Call subLoadRiwayatKonsul
    Call subLoadRiwayatPemeriksaan(False)
    Call subLoadRiwayatHasilPemeriksaan
    dgAlergi.Visible = False
    fraRiwayatResep.Visible = False
     
    If (Dir("C:\Lib\Ina Cbg\result.tlb") = "") Then
         txtTotalPenjamin.Visible = False
         cmdTotalTanggunganPenjamin = False
         LbTanggunganPenjamin.Visible = False
    End If
    sstTP.Tab = 0
    If mblnAdmin = True Then
        cmdHapusDataPT.Enabled = True
        cmdHapusDataPOA.Enabled = True
        cmdEditData.Enabled = True
        cmdHapusCatataKlinis.Enabled = True
        cmdHapusCatatanMedis.Enabled = True
        cmdHapusKonsul.Enabled = True
        cmdHapusKecelakaan.Enabled = True
        cmdUbahPT.Enabled = True
        cmdUbahOA.Enabled = True
    Else
        cmdHapusDataPT.Enabled = False
        cmdHapusDataPOA.Enabled = False
        cmdEditData.Enabled = False
        cmdHapusCatataKlinis.Enabled = False
        cmdHapusCatatanMedis.Enabled = False
        cmdHapusKonsul.Enabled = False
        cmdHapusKecelakaan.Enabled = False
        cmdUbahPT.Enabled = False
        cmdUbahOA.Enabled = False
    End If
    
    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs = Nothing
    If mblnFormDaftarPasienRJ = True Then frmDaftarPasienRJ.Enabled = True: Call frmDaftarPasienRJ.cmdCari_Click
    If mblnFormDaftarAntrian = True Then Call frmDaftarAntrianPasien.cmdCari_Click
    
End Sub



Private Sub optObat_Click()
On Error GoTo errLoad
    mstrNoPen = txtnopendaftaran.Text
    
    With frmOrderPelayananOA
'With frmOrder
'        .UseRacikan = False
        .Show
        .txtNoCMOA.Text = txtnocm.Text
        .txtNoPendaftaranOA.Text = txtnopendaftaran.Text
'        .txtNamaForm.Text = Me.Name
        .txtNamaPasien.Text = txtNamaPasien.Text
        .txtSex.Text = txtSex.Text
 
'        If dgDaftarPasien   RI.Columns(3).Value = "L" Then
'            .txtSex.Text = "Laki-Laki"
'        Else
'            .txtSex.Text = "Perempuan"
'        End If
        .txtThn.Text = txtThn.Text
        .txtBln.Text = txtBln.Text
        .txtHr.Text = txtHr.Text
'        mstrKdKelas = dgDaftarPasienRI.Columns("KdKelas").Value
'        mstrKdSubInstalasi = dgDaftarPasienRI.Columns("KdSubInstalasi").Value
'        mstrValid = dgDaftarPasienRI.Columns("NoPakai")
        .txtRP.Text = mstrNamaRuangan
        .txtDokter.Text = mstrNamaDokter
        .txtKdDokter.Text = mstrKdDokter
        .dgDokter.Visible = False
        .dcRuanganTujuan.BoundText = "702"

    End With
    
    strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdJenisPasien = rs("KdKelompokPasien").Value
        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
    End If
    
'    Me.Enabled = False
    fraOrder.Visible = False
    
Exit Sub
errLoad:
    Call msubPesanError

End Sub

Private Sub optTindakan_Click()
On Error GoTo errLoad
'    If dgDaftarPasienRI.ApproxCount = 0 Then Exit Sub
    mstrNoPen = txtnopendaftaran.Text
    With frmKonsul_OrderPelayanan
        .Show
        .txtNoCMTM.Text = txtnocm.Text
        .txtNoPendaftaranTM.Text = txtnopendaftaran.Text
        
        .txtNamaPasien.Text = txtNamaPasien.Text
        .txtSex.Text = txtSex.Text
'        If dgDaftarPasienRI.Columns(3).Value = "L" Then
'            .txtSex.Text = "Laki-Laki"
'        Else
'            .txtSex.Text = "Perempuan"
'        End If
        .txtThn.Text = txtThn.Text
        .txtBln.Text = txtBln.Text
        .txtHr.Text = txtHr.Text
        .dcDokterPerujuk.BoundText = mstrKdDokter
'        mstrKdKelas = dgDaftarPasienRI.Columns("KdKelas").Value
'        mstrKdSubInstalasi = dgDaftarPasienRI.Columns("KdSubInstalasi").Value
'        mstrValid = dgDaftarPasienRI.Columns("NoPakai")
        .txtNamaForm.Text = Me.Name
        .dcInstalasi.SetFocus
    End With
    Me.Enabled = False
    fraOrder.Visible = False
Exit Sub
errLoad:
    Call msubPesanError

End Sub

Private Sub sstTP_GotFocus()
    If txtSex.Text = "Laki-Laki" Then
        cmdKehamilandanKB.Enabled = False
    Else
        cmdKehamilandanKB.Enabled = True
    End If

End Sub

Private Sub sstTP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case sstTP.Tab
            Case 0 'pelayanan tindakan
                dgTindakan.SetFocus
            Case 1 'pemakaian obat alkes
                dgObatAlkes.SetFocus
                cmdUbahOA.Enabled = False
            Case 2 'riwayat catatan klinis
                dgCatatanKlinis.SetFocus

            Case 3 'riwayat catatan medis
                dgCatatanMedis.SetFocus
            Case 4 'riwayat diagnosa
                dgRiwayatDiagnosa.SetFocus
            Case 5 'riwayat catatan operasi
                dgRiwayatOperasi.SetFocus
            Case 6 'riwayat konsul
                dgKonsul.SetFocus
            Case 7 'riwayat kecelakaan
                dgKecelakaan.SetFocus
            Case 8 'riwayat pemeriksaan
                dgRiwayatPemeriksaan.SetFocus
            Case 9 ' riwayat hasil pemeriksaan
                dgHasilPemeriksaan.SetFocus
        End Select
    End If
End Sub

Private Sub txtAlkesTotal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtJmlBarangEditBaru_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpanEditBarang.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtJmlBarangEditBaru_LostFocus()
    On Error GoTo errLoad
    txtJmlBarangEditBaru.Text = Val(txtJmlBarangEditBaru.Text)
    If Val(txtJmlBarangEditBaru.Text) <> 0 Then
        txtJmlBarangEditBaru.Text = Format(txtJmlBarangEditBaru.Text, "#,###")
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtNoPendaftaran_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then sstTP.SetFocus: sstTP.Tab = 0
End Sub

Private Sub txtTindakanTotal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'Untuk meload riwayat diagnosa yang sudah pernah didapat
Public Sub subLoadRiwayatDiagnosa(blnAll As Boolean)

    On Error GoTo hell
    If blnAll = False Then
        strSQL = "Select * from V_DaftarDiagnosaPasien where nocm = '" & mstrNoCM & "' AND NoPendaftaran = '" & mstrNoPen & "'"
    Else
        strSQL = "Select * from V_DaftarDiagnosaPasien where nocm = '" & mstrNoCM & "'"
    End If
    Call msubRecFO(rs, strSQL)
    Set dgRiwayatDiagnosa.DataSource = rs
    With dgRiwayatDiagnosa
        .Columns(0).Width = 1500 'NoPendaftaran
        .Columns(1).Width = 0 'NoCM
        .Columns(2).Width = 1590 'TglPeriksa
        .Columns(3).Width = 2000 'JenisDiagnosa
        .Columns(4).Width = 1100 'KdDiagnosa ICD 10
        .Columns(4).Caption = "Kode ICD 10"
        .Columns(5).Width = 2700 'Diagnosa ICD 10
        .Columns(5).Caption = "Diagnosa ICD 10"
        .Columns(6).Width = 2500 'JenisDiagnosa
        .Columns(6).Caption = "JenisDiagnosaTindakan"
        .Columns(7).Width = 1000 'KdDiagnosaTindakan ICD 9
        .Columns(7).Caption = "Kode ICD 9"
        .Columns(8).Width = 2700 'DiagnosaTindakan ICD 9
        .Columns(8).Caption = "Diagnosa Tindakan ICD 9"
        .Columns(9).Width = 2000 '[Ruang Periksa]
        .Columns(10).Width = 2700 '[Dokter Pemeriksa]
        .Columns(11).Width = 0 '[Nama Pasien]
        .Columns(12).Width = 0 'JK
        .Columns(13).Width = 0 'Umur
        .Columns(14).Width = 0 'KdJnsDiagnosa
        .Columns(15).Width = 0 'KdJnsDiagnosaTindakan
        .Columns(16).Width = 0 'IdDokterPemeriksa
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

'Untuk meload riwayat pemeriksaan yang sudah pernah didapat
Public Sub subLoadRiwayatPemeriksaan(blnAll As Boolean)
    If blnAll = False Then
        strSQL = "Select * from V_RiwayatPemeriksaanPasien where nocm = '" & mstrNoCM & "' AND KdRuangan='" & mstrKdRuangan & "'"
    Else
        strSQL = "Select * from V_RiwayatPemeriksaanPasien where nocm = '" & mstrNoCM & "'"
    End If
    msubRecFO rs, strSQL
    Set dgRiwayatPemeriksaan.DataSource = rs
    With dgRiwayatPemeriksaan
        .Columns(0).Width = 0 'nocm
        .Columns(1).Width = 0 ' nopendaftaran
        .Columns(2).Width = 2400
        .Columns(3).Width = 1590
        .Columns(4).Width = 4000
        .Columns(5).Width = 3000
        .Columns(6).Width = 2500
        .Columns("KdRuangan").Width = 0
        .Columns("NoLab_Rad").Width = 0
        .Columns("KdJnsPelayanan").Width = 0
        .Columns("KdPelayananRS").Width = 0
    End With
End Sub

'Untuk meload riwayat hasil pemeriksaan yang sudah pernah didapat
Public Sub subLoadRiwayatHasilPemeriksaan()
    strSQL = "Select NoLab_Rad, [Ruang Pemeriksa], [Dokter Pemeriksa], TglPendaftaran, TglHasil, [Asal Rujukan], [Ruangan Perujuk], [Dokter Perujuk], KdInstalasi from V_RiwayatHasilPemeriksaan where nocm = '" & mstrNoCM & "'"
    msubRecFO rs, strSQL
    Set dgHasilPemeriksaan.DataSource = rs
    dgHasilPemeriksaan.Columns("KdInstalasi").Width = 0
End Sub

'Untuk meload pelayanan yang sudah pernah didapat
Public Sub subLoadPelayananDidapat()
    strSQL = "SELECT TglPelayanan,JenisPelayanan,NamaPelayanan,NamaRuangan AS [Ruang Pelayanan]," _
    & "Kelas,JenisTarif,CITO,JmlPelayanan as Jml,Total as Tarif,BiayaTotal," _
    & "DokterPemeriksa,[Status Bayar],KdPelayananRS,KdRuangan,Operator FROM V_BiayaPelayananTindakan WHERE " _
    & "NoPendaftaran='" & mstrNoPen & "' ORDER BY TglPelayanan"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgTindakan.DataSource = rs
    With dgTindakan
        .Columns(0).Width = 1600
        .Columns(1).Width = 2000
        .Columns(2).Width = 2000
        .Columns(3).Width = 1600
        .Columns(4).Width = 900
        .Columns(5).Width = 1000
        .Columns(6).Width = 500
        .Columns(7).Width = 400
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 1200
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 1200
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 2400
        .Columns(11).Width = 1200
        .Columns(12).Width = 0 'KdPelayananRS
        .Columns(13).Width = 0 'KdRuangan
        .Columns(14).Width = 2000

        .Columns(8).NumberFormat = "#,###"
        .Columns(9).NumberFormat = "#,###"
    End With

    strSQL = "SELECT sum(BiayaTotal) as TotalBayar FROM V_BiayaPelayananTindakan " _
    & "WHERE NoPendaftaran='" _
    & mstrNoPen & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount <> 0 Then
        txtTindakanTotal.Text = FormatCurrency(rs.Fields(0).Value, 2)
    Else
        txtTindakanTotal.Text = FormatCurrency(0, 2)
    End If
    If txtAlkesTotal.Text = "" Then
        txtAlkesTotal.Text = 0
        txtAlkesTotal.Text = FormatCurrency(txtAlkesTotal.Text, 2)
    End If
    If txtTindakanTotal.Text = "" Then txtTindakanTotal.Text = 0
    If txtAlkesTotal.Text = "" Then txtAlkesTotal.Text = 0
    txtGrandTotal.Text = CCur(txtTindakanTotal.Text) + CCur(txtAlkesTotal)
    txtGrandTotal.Text = FormatCurrency(txtGrandTotal.Text, 2)
End Sub

'Untuk meload pemakaian alkes yang sudah pernah didapat
Public Sub subPemakaianObatAlkes()
    strSQL = "SELECT TglPelayanan,[Detail Jenis Brg],NamaBarang," _
    & "NamaRuangan AS [Ruang Pelayanan],Kelas,JenisTarif,SatuanJml as Sat," _
    & "JmlBarang as Jml,HargaSatuan as Tarif,BiayaTotal,DokterPemeriksa," _
    & "[Status Bayar],KdBarang,KdAsal,Operator, KdRuangan, NoTerima,ResepKe, idpegawai " _
    & "FROM V_BiayaPemakaianObatAlkes WHERE NoPendaftaran='" _
    & mstrNoPen & "' ORDER BY TglPelayanan"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgObatAlkes.DataSource = rs
    With dgObatAlkes
        .Columns(0).Width = 1600
        .Columns(1).Width = 2000
        .Columns(2).Width = 2000
        .Columns(3).Width = 1600
        .Columns(4).Width = 900
        .Columns(5).Width = 1000
        .Columns(6).Width = 400
        .Columns(7).Width = 400
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 1200
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 1200
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 2400
        .Columns(11).Width = 1200
        .Columns(12).Width = 0
        .Columns(13).Width = 0
        .Columns(14).Width = 2000
        .Columns(15).Width = 0
        .Columns(16).Width = 0
        .Columns("ResepKe").Width = 0
        .Columns(17).Width = 0
        

    End With

    strSQL = "SELECT sum(BiayaTotal) as TotalBayar FROM V_BiayaPemakaianObatAlkes " _
    & "WHERE NoPendaftaran='" _
    & mstrNoPen & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount <> 0 Then
        txtAlkesTotal.Text = FormatCurrency(rs.Fields(0).Value, 2)
        If IsNull(rs.Fields(0)) = True Then txtAlkesTotal.Text = FormatCurrency(0, 2)
    Else
        txtAlkesTotal.Text = FormatCurrency(0, 2)
    End If
    If txtTindakanTotal.Text = "" Then
        txtTindakanTotal.Text = 0
        txtTindakanTotal.Text = FormatCurrency(txtTindakanTotal.Text, 2)
    End If
    If txtTindakanTotal.Text = "" Then txtTindakanTotal.Text = 0
    If txtAlkesTotal.Text = "" Then txtAlkesTotal.Text = 0
    txtGrandTotal.Text = CCur(txtTindakanTotal.Text) + CCur(txtAlkesTotal)
    txtGrandTotal.Text = FormatCurrency(txtGrandTotal.Text, 2)
End Sub

'Store procedure untuk menghapus biaya pelayanan pasien
Private Sub sp_DelBiayaPelayanan(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, dgTindakan.Columns("KdPelayananRS").Value)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(dgTindakan.Columns("TglPelayanan").Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Delete_BiayaPelayananNew"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penghapusan Biaya Pelayanan Pasien", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Delete_BiayaPelayananNew")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

'Store procedure untuk menghapus diagnosa
Private Sub sp_DelDiagnosa(ByVal adoCommand As ADODB.Command)
    Dim rsNew As New ADODB.recordset
    With adoCommand
        strSQL = "SELECT * FROM PeriksaDiagnosa WHERE NoPendaftaran='" & txtnopendaftaran.Text & "' AND KdRuangan='" & mstrKdRuangan & "' AND KdDiagnosa='" & dgRiwayatDiagnosa.Columns("Kode ICD 10").Value & "' AND TglPeriksa='" & Format(dgRiwayatDiagnosa.Columns("TglPeriksa").Value, "yyyy/MM/dd HH:mm:ss") & "'"
        Set rsNew = Nothing
        rsNew.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdDiagnosa", adVarChar, adParamInput, 7, dgRiwayatDiagnosa.Columns("Kode ICD 10").Value)
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dgRiwayatDiagnosa.Columns("TglPeriksa").Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, rsNew("KdSubInstalasi").Value)
        .Parameters.Append .CreateParameter("StatusKasus", adChar, adParamInput, 4, rsNew("StatusKasus").Value)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtnocm.Text)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Delete_Diagnosa"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penghapusan Diagnosa Pasien", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Delete_Diagnosa")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

'Untuk meload riwayat catatan klinis yang sudah pernah didapat
Public Sub subLoadRiwayatCatatanKlinis()
    strSQL = "SELECT * " & _
    " FROM V_RiwayatCatatanKlinisPasien" & _
    " WHERE (nocm = '" & mstrNoCM & "')"
    Call msubRecFO(rs, strSQL)

    Set dgCatatanKlinis.DataSource = rs
    With dgCatatanKlinis
        .Columns(0).Width = 0 'NoCM
        .Columns(1).Width = 0 'NoPendaftaran
        .Columns(2).Width = 1590 'TglPeriksa
        .Columns(3).Width = 1500 '[Ruang Pemeriksaan]
        .Columns(4).Width = 1300 'TekananDarah
        .Columns(5).Width = 1000 'Nadi
        .Columns(6).Width = 1000 'Pernapasan
        .Columns(7).Width = 1000 'Suhu
        .Columns(8).Width = 1500 'BeratTinggiBadan
        .Columns(9).Width = 1500 'Kesadaran
        .Columns(10).Width = 1500 'Keterangan
        .Columns(11).Width = 1500 'Pemeriksa
        .Columns(12).Width = 0 'KdRuangan
    End With
End Sub

'Untuk meload riwayat catatan medis yang sudah pernah didapat
Public Sub subLoadRiwayatCatatanMedis()
    strSQL = "SELECT *" & _
    " FROM V_RiwayatCatatanMedisPasien_N " & _
    " WHERE (nocm = '" & mstrNoCM & "')" 'V_RiwayatCatatanMedisPasien
    Call msubRecFO(rs, strSQL)

    Set dgCatatanMedis.DataSource = rs
    With dgCatatanMedis
        .Columns(0).Width = 0 'NoCM
        .Columns(1).Width = 0 'NoPendaftaran
        .Columns(2).Width = 1600 'TglPeriksa
        .Columns(3).Width = 1590 'RuangPemeriksaan
        .Columns(4).Width = 2000 'KeluhanUtama
'        .Columns(5).Width = 0 'Diagnosa
        .Columns(5).Width = 2500 'Pengobatan
        .Columns(6).Width = 2000 'Keterangan
        .Columns(7).Width = 2000 'RiwayatKeluarga
        .Columns(7).Caption = "Riwayat Medis Keluarga"
        .Columns(8).Width = 2500 '[Dokter Pemeriksa]
        .Columns(9).Width = 0 'KdRuangan
    End With
End Sub

'Untuk meload riwayat Kecelakaan yang sudah pernah didapat
Public Sub subLoadRiwayatKecelakaan()
    strSQL = "SELECT *" & _
    " FROM V_RiwayatKecelakanPasien " & _
    " WHERE (nocm = '" & mstrNoCM & "')"
    Call msubRecFO(rs, strSQL)

    Set dgKecelakaan.DataSource = rs
    With dgKecelakaan
        .Columns(0).Width = 0 'NoCM
        .Columns(1).Width = 0 'NoPendaftaran
        .Columns(2).Width = 1590 'TglPeriksa
        .Columns(3).Width = 1500 '[Ruangan Pemeriksa]
        .Columns(4).Width = 2500 '[Kasus Kecelakaan]
        .Columns(5).Width = 1590  'TglKecelakaan
        .Columns(6).Width = 2500 'TempatKecelakaan
        .Columns(7).Width = 1500 'Pemeriksa
        .Columns(8).Width = 0 'KdRuangan
    End With
End Sub

'Untuk meload riwayat konsul pasien
Public Sub subLoadRiwayatKonsul()
'    strSQL = "SELECT * " & _
'    " FROM V_RiwayatRujukanPasien " & _
'    " where (nocm = '" & mstrNoCM & "')"

' TAMBAH RIWAYAT RESEP APOTIK
    strSQL = "SELECT * " & _
    " FROM V_RiwayatRujukanPasienALL " & _
    " where (nocm = '" & mstrNoCM & "')"
    Call msubRecFO(rs, strSQL)

    Set dgKonsul.DataSource = rs
    With dgKonsul
        .Columns(0).Width = 0 'NoCM
        .Columns(1).Width = 0 'NoPendaftaran
        .Columns(2).Width = 1590 'TglDirujuk
        .Columns(3).Width = 2500 '[Ruangan Perujuk]
        .Columns(4).Width = 2500 '[Ruangan Tujuan]
        .Columns(5).Width = 2500 '[Dokter Perujuk]
        .Columns(6).Width = 1700 'StatusPeriksa
        .Columns(7).Width = 0 'KdRuanganAsal
        .Columns("KdRuanganTujuan").Width = 0 'KdRuanganTujuan
    End With
End Sub

'Untuk meload riwayat operasi yang sudah pernah didapat
Public Sub subLoadRiwayatOperasi()
    strSQL = " SELECT     TOP (200) NoHasilPeriksa, NoCM, NoPendaftaran, KasusPenyakit, JenisTindakanMedis, TglMulaiPeriksa, TglAkhirPeriksa, TglHasilPeriksa" & _
    " FROM   V_HasilTindakanMedis " & _
    " WHERE (nocm = '" & mstrNoCM & "')"
    Call msubRecFO(rs, strSQL)
    Set dgRiwayatOperasi.DataSource = rs
    With dgRiwayatOperasi
        .Columns(0).Width = 1200 'NoHasilPeriksa
        .Columns(1).Width = 1500 'NoCM
'        .Columns("NoCM").Caption = "No. Rekam Medis"
        .Columns(2).Width = 1250 'NoPendaftaran
        .Columns(3).Width = 2000 'kasus penyakt
        .Columns(4).Width = 2000 'jenis tindakan medis
        .Columns(5).Width = 2000 'tgl mulai
        .Columns(6).Width = 2000 'tgl akhir
        .Columns(7).Width = 2000 'tgl hasil

    End With
End Sub

'Store procedure untuk menghapus catatan klinis
Private Sub sp_DelBiayaCatatanKlinis(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dgCatatanKlinis.Columns("TglPeriksa").Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Delete_CatatanKlinis"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penghapusan catatan klinis", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Delete_CatatanKlinis")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

'Store procedure untuk menghapus catatan medis
Private Sub sp_DelBiayaCatatanMedis(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dgCatatanMedis.Columns("TglPeriksa").Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Delete_CatatanMedis"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penghapusan catatan medis", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Delete_CatatanMedis")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub
'Store procedure untuk menghapus catatan medis Alergi
Private Sub sp_Alergi(ByVal adoCommand As ADODB.Command)
    On Error GoTo errLoad

    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtnocm.Text)
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dgCatatanMedis.Columns("TglPeriksa").Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, dgCatatanMedis.Columns("IdDokter").Value)
        .Parameters.Append .CreateParameter("KdAlergi", adChar, adParamInput, 2, IIf(dgCatatanMedis.Columns("kdalergi").Value = "", Null, dgCatatanMedis.Columns("kdalergi").Value))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 1000, Null)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "D")

        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_CatatanAlergiPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Catatan Medis Pasien", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("AUD_CatatanAlergiPasien")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
errLoad:
    Call msubPesanError
End Sub


'Store procedure untuk menghapus data kecelakaan
Private Sub sp_DelKecelakaan(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Delete_KasusKecelakaan"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penghapusan data kecelakaan", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Delete_KasusKecelakaan")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

'Store procedure untuk menghapus data konsul
Private Sub sp_DelKonsul(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("TglDirujuk", adDate, adParamInput, , Format(dgKonsul.Columns("TglDirujuk").Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Delete_PasienRujukan"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penghapusan data kecelakaan", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Delete_PasienRujukan")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

Private Function sp_EditQtyBarang(f_KdBarang As String, f_KdAsal As String, f_Satuan As String, _
    f_JmlBarangLama As Integer, f_JmlBarangTambahan As Integer, _
    f_TglPelayanan As Date, f_KdJenisObat As String, f_ResepKe As Integer, f_NoTerima As String) As Boolean
    On Error GoTo errLoad
    sp_EditQtyBarang = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("Satuan", adChar, adParamInput, 1, f_Satuan)
        .Parameters.Append .CreateParameter("JmlBrgLama", adInteger, adParamInput, , f_JmlBarangLama)
        .Parameters.Append .CreateParameter("JmlBrgTambahan", adInteger, adParamInput, , f_JmlBarangTambahan)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(f_TglPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdJenisObat", adChar, adParamInput, 2, f_KdJenisObat)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("ResepKe", adTinyInt, adParamInput, , f_ResepKe)
        .Parameters.Append .CreateParameter("Noterima", adVarChar, adParamInput, 10, f_NoTerima)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_PemakaianObatAlkes"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_EditQtyBarang = False
        Else
            Call Add_HistoryLoginActivity("Update_PemakaianObatAlkes")
        End If
    End With
    Exit Function
errLoad:
    sp_EditQtyBarang = False
    Call msubPesanError
End Function

Private Sub subLoadDiagramOdonto()
    On Error GoTo hell
    Dim blnSudahAda As Boolean
    Dim strTglPeriksa As String
    Dim i As Integer

    If dgHasilPemeriksaan.ApproxCount = 0 Then Exit Sub

    strSQL = "select NoPendaftaran,TglPeriksa from DetailCatatanOdonto where NoPendaftaran='" & mstrNoPen & "'"
    Call msubRecFO(rs, strSQL)

    With frmDiagramOdonto
        .Show

        For i = 0 To 14
            .optAksi(i).Visible = False
        Next i
        .txtKeterangan.Visible = False
        .Label3.Visible = False
        .lblBelumErupsi.Visible = False
        .lblErupsiSebagian.Visible = False
        .lblAnomaliBentuk.Visible = False
        .lblCalculus.Visible = False
        .picKaries.Visible = False
        .picNonVital.Visible = False
        .picTLogam.Visible = False
        .picTNonLogam.Visible = False
        .picMLogam.Visible = False
        .picMNonLogam.Visible = False
        .picSisaAkar.Visible = False
        .picGigiHilang.Visible = False
        .picJembatan.Visible = False
        .picGigiTiruanLepas.Visible = False
        .cmdSimpan.Visible = False
        .dtpTglPeriksa.Enabled = False

        .Frame2.Height = 800
        .Frame4.Top = .Frame2.Top + .Frame2.Height
        .Height = 8300

        .txtnopendaftaran.Text = mstrNoPen
        .txtnocm.Text = mstrNoCM
        .txtNamaPasien.Text = txtNamaPasien.Text
        If txtSex.Text = "L" Then
            .txtSex.Text = "Laki-Laki"
        Else
            .txtSex.Text = "Perempuan"
        End If
        .txtThn.Text = txtThn.Text
        .txtBln.Text = txtBln.Text
        .txtHr.Text = txtHr.Text
        .txtKls.Text = txtKls.Text
        .txtJenisPasien.Text = txtJenisPasien.Text
        .txtTglDaftar.Text = dgHasilPemeriksaan.Columns("TglHasil")

        strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
            mstrKdJenisPasien = rs("KdKelompokPasien").Value
            mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
        End If
        .subLoadDetailCatatanOdonto
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Public Sub subLoadDataRiwayatAlergiPasien()
    On Error GoTo errLoad

    strSQL = "SELECT *" & _
        " FROM V_RiwayatCatatanAlergipasien " & _
        " WHERE (nocm = '" & mstrNoCM & "')"
    Call msubRecFO(rs, strSQL)

    Set dgAlergi.DataSource = rs
    With dgAlergi
        .Columns(0).Width = 2000 'NoCM
'        .Columns(0).Caption = "No. Rekam Medis" 'NoCM
        .Columns(1).Width = 2000 'NoPendaftaran
        .Columns(2).Width = 0 'NamaPasien
        .Columns(3).Width = 1500 'TglPemeriksaan
        .Columns(4).Width = 2800 'RuanganPemeriksaan
        .Columns(5).Width = 0 'KdRuangan
        .Columns(6).Width = 3350 'NamaAlergi
        .Columns(7).Width = 2400 'Keterangan
    End With

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

