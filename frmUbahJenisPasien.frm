VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmUbahJenisPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Asuransi Pasien"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUbahJenisPasien.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   9750
   Begin VB.TextBox txtNoBKM 
      Height          =   375
      Left            =   4800
      TabIndex        =   69
      Text            =   "Text1"
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtKdInstalasi 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1920
      TabIndex        =   66
      Text            =   "txtKdInstalasi"
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraDataRujukan 
      Caption         =   "Data Rujukan"
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
      TabIndex        =   55
      Top             =   6360
      Width           =   9735
      Begin VB.TextBox txtNoRujukan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2280
         MaxLength       =   30
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   480
         Width           =   3375
      End
      Begin MSDataListLib.DataCombo dcAsalRujukan 
         Height          =   330
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSComCtl2.DTPicker dtpTglDirujuk 
         Height          =   315
         Left            =   240
         TabIndex        =   25
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
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
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   24182787
         UpDown          =   -1  'True
         CurrentDate     =   37694
      End
      Begin MSDataListLib.DataCombo dcNamaPerujuk 
         Height          =   330
         Left            =   2280
         TabIndex        =   26
         Top             =   1080
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcNamaAsalRujukan 
         Height          =   330
         Left            =   5880
         TabIndex        =   24
         Top             =   480
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcDiagnosa 
         Height          =   330
         Left            =   5880
         TabIndex        =   27
         Top             =   1080
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama tempat Perujuk = Nama Puskesmas/ Nama Klinik/ Tempat Dokter Praktek/ Nama Rumah Sakit"
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
         Left            =   240
         TabIndex        =   62
         Top             =   1440
         Width           =   8565
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Perujuk (Dokter, Bidan, Mantri, dll)"
         Height          =   210
         Index           =   21
         Left            =   2280
         TabIndex        =   61
         Top             =   840
         Width           =   3345
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Rujukan"
         Height          =   210
         Index           =   24
         Left            =   2280
         TabIndex        =   60
         Top             =   240
         Width           =   930
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asal Rujukan"
         Height          =   210
         Index           =   25
         Left            =   240
         TabIndex        =   59
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Asal Rujukan (Nama Tempat Rujukan)"
         Height          =   210
         Index           =   27
         Left            =   5880
         TabIndex        =   58
         Top             =   240
         Width           =   3600
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Dirujuk"
         Height          =   210
         Index           =   26
         Left            =   240
         TabIndex        =   57
         Top             =   840
         Width           =   930
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diagnosa (Penyakit) Rujukan"
         Height          =   210
         Index           =   22
         Left            =   5880
         TabIndex        =   56
         Top             =   840
         Width           =   2325
      End
   End
   Begin VB.Frame fraPemakaianAsuransi 
      Caption         =   "Pemakaian Asuransi  (SJP = Surat Jaminan Pelayanan)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   48
      Top             =   4560
      Width           =   9735
      Begin VB.TextBox txtNoSJP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3000
         MaxLength       =   30
         TabIndex        =   17
         Top             =   600
         Width           =   4455
      End
      Begin VB.TextBox txtNoKunjungan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         MaxLength       =   1
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtNoBP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         MaxLength       =   3
         TabIndex        =   19
         Text            =   "a24"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtAnakKe 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2040
         MaxLength       =   1
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox chkNoSJP 
         Caption         =   "No. SJP Otomatis"
         Enabled         =   0   'False
         Height          =   210
         Left            =   3000
         TabIndex        =   16
         Top             =   360
         Width           =   3375
      End
      Begin MSDataListLib.DataCombo dcHubungan 
         Height          =   330
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSComCtl2.DTPicker dtpTglSJP 
         Height          =   315
         Left            =   7560
         TabIndex        =   18
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
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
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   24182787
         UpDown          =   -1  'True
         CurrentDate     =   37694
      End
      Begin MSDataListLib.DataCombo dcUnitKerja 
         Height          =   330
         Left            =   1920
         TabIndex        =   21
         Top             =   1200
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcKelasDitanggung 
         Height          =   330
         Left            =   7560
         TabIndex        =   70
         Top             =   1200
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
         Caption         =   "Kelas Ditanggung"
         Height          =   210
         Index           =   13
         Left            =   7560
         TabIndex        =   71
         Top             =   960
         Width           =   1410
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. SJP"
         Height          =   210
         Index           =   17
         Left            =   7560
         TabIndex        =   54
         Top             =   360
         Width           =   660
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hubungan Pasien"
         Height          =   210
         Index           =   15
         Left            =   240
         TabIndex        =   53
         Top             =   360
         Width           =   1410
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Anak Ke -"
         Height          =   210
         Index           =   16
         Left            =   2040
         TabIndex        =   52
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. BP"
         Height          =   210
         Index           =   18
         Left            =   240
         TabIndex        =   51
         Top             =   960
         Width           =   555
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kunj. Ke -"
         Height          =   210
         Index           =   20
         Left            =   960
         TabIndex        =   50
         Top             =   960
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bertugas di Unit / Bagian"
         Height          =   210
         Index           =   19
         Left            =   1920
         TabIndex        =   49
         Top             =   960
         Width           =   2025
      End
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   7920
      TabIndex        =   29
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   6240
      TabIndex        =   28
      Top             =   8280
      Width           =   1575
   End
   Begin VB.TextBox txtNamaFormPengirim 
      Height          =   495
      Left            =   0
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtTglPendaftaran 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Frame fraDataKartuPeserta 
      Caption         =   "Data Kartu Peserta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   38
      Top             =   2160
      Width           =   9735
      Begin VB.CheckBox chkDiriSendiri 
         Caption         =   "Diri Sendiri"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   8400
         TabIndex        =   7
         Top             =   160
         Width           =   1215
      End
      Begin VB.TextBox txtAlamatPA 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2280
         MaxLength       =   100
         TabIndex        =   13
         Top             =   1800
         Width           =   6615
      End
      Begin VB.TextBox txtNipPA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6960
         MaxLength       =   16
         ScrollBars      =   1  'Horizontal
         TabIndex        =   12
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtNamaPA 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox txtNoKartuPA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         MaxLength       =   15
         TabIndex        =   9
         Top             =   1200
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo dcPenjamin 
         Height          =   330
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSComCtl2.DTPicker dtpTglLahirPA 
         Height          =   315
         Left            =   5400
         TabIndex        =   11
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
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
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   24182785
         UpDown          =   -1  'True
         CurrentDate     =   37694
      End
      Begin MSDataListLib.DataCombo dcPerusahaan 
         Height          =   330
         Left            =   4560
         TabIndex        =   67
         Top             =   600
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
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
      Begin MSDataListLib.DataCombo dcGolonganasuransi 
         Height          =   330
         Left            =   240
         TabIndex        =   72
         Top             =   1800
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
         Caption         =   "Golongan Asuransi"
         Height          =   210
         Index           =   7
         Left            =   240
         TabIndex        =   73
         Top             =   1560
         Width           =   1485
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Perusahaan Penjamin"
         Height          =   210
         Index           =   2
         Left            =   4560
         TabIndex        =   68
         Top             =   360
         Width           =   2235
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat Peserta"
         Height          =   210
         Index           =   14
         Left            =   2280
         TabIndex        =   44
         Top             =   1560
         Width           =   1230
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Lahir"
         Height          =   210
         Index           =   11
         Left            =   5400
         TabIndex        =   43
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. KTP / SIM Peserta"
         Height          =   210
         Index           =   12
         Left            =   6960
         TabIndex        =   42
         Top             =   960
         Width           =   1845
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Peserta"
         Height          =   210
         Index           =   10
         Left            =   2160
         TabIndex        =   41
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Kartu Peserta"
         Height          =   210
         Index           =   9
         Left            =   240
         TabIndex        =   40
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Penjamin"
         Height          =   210
         Index           =   8
         Left            =   240
         TabIndex        =   39
         Top             =   360
         Width           =   1245
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
      TabIndex        =   30
      Top             =   960
      Width           =   9735
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   63
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo dcJenisPasien 
         Height          =   330
         Left            =   7680
         TabIndex        =   6
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         MaxLength       =   6
         TabIndex        =   0
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3840
         MaxLength       =   9
         TabIndex        =   2
         Top             =   600
         Width           =   1095
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
         Height          =   615
         Left            =   5040
         TabIndex        =   31
         Top             =   360
         Width           =   2535
         Begin VB.TextBox txtThn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            MaxLength       =   6
            TabIndex        =   3
            Top             =   250
            Width           =   375
         End
         Begin VB.TextBox txtBln 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   960
            MaxLength       =   6
            TabIndex        =   4
            Top             =   250
            Width           =   375
         End
         Begin VB.TextBox txtHr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            MaxLength       =   6
            TabIndex        =   5
            Top             =   250
            Width           =   375
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "thn"
            Height          =   210
            Index           =   4
            Left            =   600
            TabIndex        =   34
            Top             =   302
            Width           =   285
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "bln"
            Height          =   210
            Index           =   5
            Left            =   1440
            TabIndex        =   33
            Top             =   302
            Width           =   240
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "hr"
            Height          =   210
            Index           =   6
            Left            =   2280
            TabIndex        =   32
            Top             =   302
            Width           =   165
         End
      End
      Begin VB.Label lblNoPendaftaran 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Left            =   7680
         TabIndex        =   64
         Top             =   -120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblJenisPasien 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Pasien"
         Height          =   210
         Left            =   7680
         TabIndex        =   45
         Top             =   360
         Width           =   960
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. CM"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   37
         Top             =   360
         Width           =   585
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pasien"
         Height          =   210
         Index           =   1
         Left            =   1320
         TabIndex        =   36
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Index           =   3
         Left            =   3840
         TabIndex        =   35
         Top             =   360
         Width           =   1065
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   65
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
      Left            =   7920
      Picture         =   "frmUbahJenisPasien.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmUbahJenisPasien.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmUbahJenisPasien.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmUbahJenisPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fTglLahir As Date
Dim fNoPendaftaran As String
Dim fNoSJP As String
Dim fNoBP As String
Dim fNoKunjungan As Integer
Dim fChkNoSJP As String
Dim fDcUnitKerja As String
Dim fNamaAsalRujukan As String
Dim fNamaPerujuk As String
Dim fDiagnosa As String
Dim fAlamatPA As String
Dim fIDPeserta As String
Dim fKdPerusahaan As String

Private Sub subBayarOtomatis()
On Error GoTo hell:
   'exec ke pasien sudah bayar
    'aktifkan jika tidak perlu otomatis bayar
            strSQL = "SELECT SUM(Tarif) AS Tarif, SUM(JmlHutangPenjamin) AS JmlHutangPenjamin, SUM(JmlTanggunganRS) AS JmlTanggunganRS, SUM(JmlPembebasan) " & _
                " AS JmlPembebasan From DetailBiayaPelayanan WHERE     (NoPendaftaran = '" & mstrNoPen & "')"
            Call msubRecFO(rs, strSQL)
            If mstrKdInstalasi = "06" Then
                curTarif = 0
                curTP = 0
                curTRS = 0
                curPemb = 0
                mcurAll_HrsDibyr = 0
                mcurBayar = 0
                mcurAll_TP = 0
                mcurAll_TRS = 0
                mcurAll_Pemb = 0
            Else
                curTarif = rs.Fields("Tarif")
                curTP = rs.Fields("JmlHutangPenjamin")
                curTRS = rs.Fields("JmlTanggunganRS")
                curPemb = rs.Fields("JmlPembebasan")
                mcurAll_HrsDibyr = curTarif - (curTP + curTRS + curPemb)
                mcurBayar = curTarif
                mcurAll_TP = curTP
                mcurAll_TRS = curTRS
                mcurAll_Pemb = curPemb
            End If
            Set rs = Nothing
            Call msubRecFO(rs, "SELECT KdKelompokPasien,IdPenjamin FROM V_JenisPasienNPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')")
            If rs.EOF = True Then mstrKdJenisPasien = "01" Else mstrKdJenisPasien = rs("KdKelompokPasien").Value
            If rs.EOF = True Then
                mstrKdPenjamin = "2222222222"
            Else
                If mstrKdJenisPasien = "01" Then
                    mstrKdPenjamin = "2222222222"
                Else
                    mstrKdPenjamin = rs("IdPenjamin").Value
                End If
            End If
    
            If sp_AddStrukBuktiKasMasuk() = False Then Exit Sub
            mstrNoBKM = txtNoBKM.Text
            If sp_AddStruk(dbcmd, 1) = False Then Exit Sub
            fStatusPiutang = "TM"
            fStatusBayarSemua = "Y"
            Call f_AddStrukPelayananPasienDetail(mstrNoBKM, mstrNoStruk, mstrNoPen, mstrNoCM, CCur(mcurAll_HrsDibyr), 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, mcurBayar, mcurAll_HrsDibyr, mcurAll_TP, mcurAll_TRS, mcurAll_Pemb, mcurAll_HrsDibyr, 0, 0)
    
'end code bayar otomatis
Exit Sub
hell:
    msubPesanError ("Sub Bayar Ototmatis")
End Sub
'Store procedure untuk mengisi struk billing pasien
Public Function sp_AddStrukBuktiKasMasuk() As Boolean
On Error GoTo errload
    
    sp_AddStrukBuktiKasMasuk = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("TglBKM", adDate, adParamInput, , Format(dTglDaftar, "yyyy/MM/dd HH:mm:ss"))
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
Public Function sp_AddStruk(ByVal adoCommand As ADODB.Command, strStsByr As String) As Boolean
    On Error GoTo errload
    sp_AddStruk = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoBKM", adChar, adParamInput, 10, mstrNoBKM)
        .Parameters.Append .CreateParameter("OutputNoStruk", adChar, adParamOutput, 10, Null)
        .Parameters.Append .CreateParameter("TglStruk", adDate, adParamInput, , Format(dTglDaftar, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("NoCM", adChar, adParamInput, 6, mstrNoCM)
        .Parameters.Append .CreateParameter("KdKelompokPasien", adChar, adParamInput, 2, mstrKdJenisPasien)
        
        .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, IIf(mstrKdPenjamin = "", "2222222222", mstrKdPenjamin))
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

Private Sub subLoadPemakaianAsuransi(s_NoPendaftaran As String, s_IdPenjamin As String)
On Error GoTo errload

    strSQL = "SELECT * FROM v_PemakaianAsuransi WHERE NoPendaftaran = '" & s_NoPendaftaran & "' AND IdPenjamin='" & s_IdPenjamin & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.EOF = False Then
        dcHubungan.BoundText = IIf(IsNull(rs("KdHubungan")), "", rs("KdHubungan"))
        txtAnakKe.Text = IIf(IsNull(rs("AnakKe")), "", rs("AnakKe"))
        txtNoSJP.Text = IIf(IsNull(rs("NoSJP")), "", rs("NoSJP"))
        dtpTglSJP.Value = IIf(IsNull(rs("TglSJP")), Now, rs("TglSJP"))
        txtNoBP.Text = IIf(IsNull(rs("NoBP")), "", rs("NoBP"))
        txtNoKunjungan.Text = IIf(IsNull(rs("KunjunganKe")), "", rs("KunjunganKe"))
        dcUnitKerja.Text = IIf(IsNull(rs("UnitBagian")), "", rs("UnitBagian"))
        dcKelasDitanggung.BoundText = IIf(IsNull(rs("KdKelasDiTanggung")), "", rs("KdKelasDiTanggung"))
    Else
        dcHubungan.BoundText = ""
        txtAnakKe.Text = ""
        txtNoSJP.Text = ""
        dtpTglSJP.Value = Now
        txtNoBP.Text = ""
        txtNoKunjungan.Text = ""
        dcUnitKerja.Text = ""
        dcKelasDitanggung.BoundText = ""
    End If
    
Exit Sub
errload:
    Call msubPesanError("subLoadPemakaianAsuransi")
End Sub

Private Sub subTampungDataPenjamin()
    typAsuransi.strIdPenjamin = dcPenjamin.BoundText
    typAsuransi.strIdAsuransi = txtNoKartuPA.Text
    typAsuransi.strNoCm = txtNoCM.Text
    typAsuransi.strNamaPeserta = txtNamaPA.Text
    typAsuransi.strIdPeserta = txtNipPA.Text
    
    typAsuransi.strPerusahaanPenjamin = dcPerusahaan.BoundText
    
    typAsuransi.strKdKelasDitanggung = dcKelasDitanggung.BoundText
    typAsuransi.strKdGolongan = dcGolonganasuransi.BoundText
    typAsuransi.dTglLahir = dtpTglLahirPA.Value
    typAsuransi.strAlamat = txtAlamatPA.Text
    typAsuransi.strNoPendaftaran = txtNoPendaftaran.Text
    typAsuransi.strHubungan = dcHubungan.BoundText
    
    typAsuransi.strNoSJP = txtNoSJP.Text
    typAsuransi.dTglSJP = dtpTglSJP.Value
    typAsuransi.strNoBp = txtNoBP.Text
    typAsuransi.intNoKunjungan = IIf(Val(txtNoKunjungan.Text) = 0, 1, Val(txtNoKunjungan.Text))
    
    typAsuransi.strUnitBagian = dcUnitKerja.Text
    
    typAsuransi.strNoRujukan = txtNoRujukan.Text
    typAsuransi.strKdRujukanAsal = dcAsalRujukan.BoundText
    typAsuransi.strDetailRujukanAsal = dcNamaAsalRujukan.Text
    typAsuransi.strKdDetailRujukanAsal = dcNamaAsalRujukan.BoundText
    typAsuransi.strNamaPerujuk = dcNamaPerujuk.Text
    
    typAsuransi.dTglDirujuk = dtpTglDirujuk.Value
    typAsuransi.strDiagnosaRujukan = dcDiagnosa.Text
    typAsuransi.strKdDiagnosa = dcDiagnosa.BoundText
    
    typAsuransi.blnSuksesAsuransi = True
        
    cmdSimpan.Enabled = False
End Sub

Private Function sp_JenisPasienJoinProgramAskes() As Boolean
On Error GoTo errload

    MousePointer = vbHourglass
    sp_JenisPasienJoinProgramAskes = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, dcPenjamin.BoundText)
        .Parameters.Append .CreateParameter("IdAsuransi", adVarChar, adParamInput, 25, txtNoKartuPA)
        .Parameters.Append .CreateParameter("NoCM", adChar, adParamInput, 6, txtNoCM)
        .Parameters.Append .CreateParameter("NamaPeserta", adVarChar, adParamInput, 50, txtNamaPA.Text)
'5
        .Parameters.Append .CreateParameter("IDPeserta", adVarChar, adParamInput, 16, txtNipPA)
        .Parameters.Append .CreateParameter("KdGolongan", adChar, adParamInput, 2, IIf(Len(Trim(dcGolonganasuransi.Text)) = 0, Null, Trim(dcGolonganasuransi.BoundText)))
        .Parameters.Append .CreateParameter("TglLahir", adDate, adParamInput, , Format(dtpTglLahirPA, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("Alamat", adVarChar, adParamInput, 100, txtAlamatPA)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, IIf(Len(Trim(txtNoPendaftaran.Text)) = 0, Null, txtNoPendaftaran.Text))
'10
        .Parameters.Append .CreateParameter("KdHubungan", adChar, adParamInput, 2, dcHubungan.BoundText)
        .Parameters.Append .CreateParameter("NoSJP", adVarChar, adParamInput, 30, IIf(Len(Trim(txtNoSJP.Text)) = 0, Null, Trim(txtNoSJP.Text)))
        .Parameters.Append .CreateParameter("TglSJP", adDate, adParamInput, , Format(dtpTglSJP, "yyyy/MM/dd hh:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("NoBP", adChar, adParamInput, 3, IIf(Len(Trim(txtNoBP.Text)) = 0, Null, Trim(txtNoBP.Text)))
'15
        .Parameters.Append .CreateParameter("KunjunganKe", adInteger, adParamInput, , IIf(Val(txtNoKunjungan.Text) = 0, "1", txtNoKunjungan.Text))
        .Parameters.Append .CreateParameter("OutputNoSJP", adVarChar, adParamOutput, 30, Null)
        .Parameters.Append .CreateParameter("StatusNoSJP", adChar, adParamInput, 1, IIf(chkNoSJP.Value = vbChecked, "O", "M"))
        .Parameters.Append .CreateParameter("AnakKe", adInteger, adParamInput, , Val(txtAnakKe.Text))
        .Parameters.Append .CreateParameter("UnitBagian", adVarChar, adParamInput, 50, IIf(Len(Trim(dcUnitKerja.Text)) = 0, Null, Trim(dcUnitKerja.Text)))
'20
        .Parameters.Append .CreateParameter("KdPaket", adVarChar, adParamInput, 3, Null)
        .Parameters.Append .CreateParameter("NoRujukan", adVarChar, adParamInput, 30, txtNoRujukan.Text)
        .Parameters.Append .CreateParameter("KdRujukanAsal", adChar, adParamInput, 2, dcAsalRujukan.BoundText)
        .Parameters.Append .CreateParameter("DetailRujukanAsal", adVarChar, adParamInput, 100, IIf(Len(Trim(dcNamaAsalRujukan.Text)) = 0, Null, dcNamaAsalRujukan.Text))
        .Parameters.Append .CreateParameter("KdDetailRujukanAsal", adChar, adParamInput, 8, IIf(chkNoSJP.Value = vbChecked, "12345678", dcNamaAsalRujukan.BoundText))
'25
        .Parameters.Append .CreateParameter("NamaPerujuk", adVarChar, adParamInput, 50, IIf(Len(Trim(dcNamaPerujuk.Text)) = 0, Null, Trim(dcNamaPerujuk.Text)))
        .Parameters.Append .CreateParameter("TglDirujuk", adDate, adParamInput, , Format(dtpTglDirujuk.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("DiagnosaRujukan", adVarChar, adParamInput, 100, IIf(Len(Trim(dcDiagnosa.Text)) = 0, Null, Trim(dcDiagnosa.Text)))
        .Parameters.Append .CreateParameter("KdDiagnosa", adVarChar, adParamInput, 7, dcDiagnosa.BoundText)
        .Parameters.Append .CreateParameter("KdKelompokPasien", adChar, adParamInput, 2, dcJenisPasien.BoundText)
  '      .Parameters.Append .CreateParameter("IdPerusahaan", adVarChar, adParamInput, 10, dcPerusahaan.BoundText)
        .Parameters.Append .CreateParameter("IdPerusahaan", adChar, adParamInput, 10, IIf(dcPerusahaan.Text = "", Null, dcPerusahaan.BoundText))
        .Parameters.Append .CreateParameter("KdKelasDiTanggung", adChar, adParamInput, 2, dcKelasDitanggung.BoundText)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_JenisPasienJoinProgramAskes"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validfasi"
            sp_JenisPasienJoinProgramAskes = False
        Else
            MsgBox "Ubah Jenis Pasien Berhasil", vbInformation, "Informasi"
            txtNoSJP.Text = IIf(IsNull(.Parameters("OutputNoSJP")), "", .Parameters("OutputNoSJP"))
            cmdSimpan.Enabled = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    MousePointer = vbDefault

Exit Function
errload:
    MousePointer = vbDefault
    sp_JenisPasienJoinProgramAskes = False
    Call msubPesanError("sp_JenisPasienJoinProgramAskes")
End Function

Private Sub subKosong()
    txtNoCM.Text = ""
    txtNamaPasien.Text = ""
    txtJK.Text = ""
    txtThn.Text = ""
    txtBln.Text = ""
    txtHr.Text = ""
    txtNoPendaftaran.Text = ""
    
    chkDiriSendiri.Value = vbUnchecked
    
    dcPenjamin.BoundText = ""
    txtNoKartuPA.Text = ""
    txtNamaPA.Text = ""
    dtpTglLahirPA.Value = Now
    txtNipPA.Text = ""
    dcKelasDitanggung.BoundText = ""
    txtAlamatPA.Text = ""
    
    dcHubungan.BoundText = ""
    txtAnakKe.Text = ""
    chkNoSJP.Value = vbUnchecked
    dtpTglSJP.Value = Now
    txtNoBP.Text = ""
    txtNoKunjungan.Text = ""
    dcUnitKerja.BoundText = ""
    
    dcAsalRujukan.BoundText = ""
    txtNoRujukan.Text = ""
    dcNamaAsalRujukan.BoundText = ""
    dtpTglDirujuk.Value = Now
    dcNamaPerujuk.BoundText = ""
    dcDiagnosa.BoundText = ""
    dcGolonganasuransi.BoundText = ""
End Sub

Private Sub subLoadDCSource()
On Error GoTo errload
    Call msubDcSource(dcJenisPasien, rs, "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien Where KdKelompokPasien<>'05'")
    Call msubDcSource(dcHubungan, rs, "SELECT KdHubungan, NamaHubungan FROM HubunganPesertaAsuransi")
    Call msubDcSource(dcKelasDitanggung, rs, "SELECT DISTINCT KdKelas,DeskKelas FROM KelasPelayanan where KdKelas<>'04'")
    Call msubDcSource(dcGolonganasuransi, rs, "SELECT     KdGolongan, NamaGolongan FROM GolonganAsuransi")
    Call msubDcSource(dcUnitKerja, rs, "SELECT KdRuangan, NamaRuangan FROM Ruangan ORDER BY NamaRuangan")
    Call msubDcSource(dcAsalRujukan, rs, "SELECT KdRujukanAsal, RujukanAsal FROM RujukanAsal")
    strSQL = "SELECT KdDetailRujukanAsal, DetailRujukanAsal" & _
        " FROM DetailRujukanAsal " & _
        " WHERE (KdRujukanAsal = '" & dcAsalRujukan.BoundText & "')"
    Call msubDcSource(dcNamaAsalRujukan, rs, strSQL)
    Call msubDcSource(dcNamaPerujuk, rs, "SELECT KodeDokter, NamaDokter FROM V_DaftarDokter")
    Call msubDcSource(dcDiagnosa, rs, "SELECT KdDiagnosa, NamaDiagnosa FROM Diagnosa ORDER BY NamaDiagnosa")
    
    strSQL = "SELECT  IdPenjamin, NamaPenjamin FROM dbo.Penjamin order by NamaPenjamin"
    Call msubDcSource(dcPerusahaan, rs, strSQL)
Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub chkDiriSendiri_Click()
On Error GoTo errload
    If chkDiriSendiri.Value = 1 Then
        strSQL = "SELECT NamaLengkap, NoIdentitas, Alamat FROM v_S_RegistrasiDataPasien WHERE NocM='" & txtNoCM.Text & "'"
        Call msubRecFO(rs, strSQL)
        If rs.RecordCount <> 0 Then
            txtNamaPA.Text = rs("NamaLengkap")
            txtNipPA.Text = rs("NoIdentitas") & ""
            txtAlamatPA.Text = rs("Alamat") & ""
            dcHubungan.Text = "Peserta"
        Else
            txtNamaPA.Text = ""
            txtNipPA.Text = ""
            txtAlamatPA.Text = ""
            dcHubungan.Text = ""
        End If
    Else
        txtNamaPA.Text = ""
        txtNipPA.Text = ""
        txtAlamatPA.Text = ""
        dcHubungan.Text = ""
    End If
Exit Sub
errload:
    msubPesanError
End Sub

Private Sub chkDiriSendiri_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcPenjamin.SetFocus
End Sub

Private Sub chkNoSJP_Click()
    If chkNoSJP.Value = vbChecked Then txtNoSJP.Enabled = False Else txtNoSJP.Enabled = True
End Sub

Private Sub chkNoSJP_KeyPress(KeyAscii As Integer)
    If chkDiriSendiri.Value = vbChecked Then dtpTglSJP.SetFocus Else txtNoSJP.SetFocus
End Sub

Private Sub cmdSimpan_Click()
'On Error GoTo errload
    Call msubRecFO(dbRst, "SELECT IdPenjamin FROM PenjaminKelompokPasien WHERE KdKelompokPasien = '" & dcJenisPasien.BoundText & "'")
    If dbRst(0).Value = "2222222222" Then
        'konversi sp ke fungsi
        If sp_UpdateJenisPasienUmum(dcJenisPasien.BoundText, txtNoPendaftaran.Text) = False Then Exit Sub
        MousePointer = vbHourglass
        'add bayar ototmatis
        If mstrKdInstalasi = "02" Or mstrKdInstalasi = "06" Or mstrKdInstalasi = "11" Then
            Call subBayarOtomatis
        End If
        '--
        'Call f_UpdateJenisPasienUmum(dcJenisPasien.BoundText, txtNoPendaftaran.Text)
        MsgBox "Ubah Data Pasien Berhasil", vbInformation, "Informasi"
        cmdSimpan.Enabled = False
        MousePointer = vbDefault
        Exit Sub
    End If
    
    If Periksa("datacombo", dcPenjamin, "Nama pejamin harus diisi") = False Then Exit Sub
    'If Periksa("text", txtNoSJP, "NoSJP harus diisi") = False Then Exit Sub
    If Periksa("text", txtNoKartuPA, "NoKartu harus diisi") = False Then Exit Sub
    If Periksa("text", txtNamaPA, "Nama Peserta harus diisi") = False Then Exit Sub
    If Periksa("datacombo", dcKelasDitanggung, "Kelas Ditangung harus diisi") = False Then Exit Sub
    If Periksa("datacombo", dcHubungan, "Hubungan peserta harus diisi") = False Then Exit Sub
    'If chkNoSJP.Value = vbUnchecked Then If Periksa("text", txtNoSJP, "No SJP harus diisi") = False Then Exit Sub

    If Periksa("datacombo", dcAsalRujukan, "Asal rujukan harus diisi") = False Then Exit Sub
    If Periksa("text", txtNoRujukan, "No rujukan harus diisi") = False Then Exit Sub
    
'    If sp_AmbulNoKunjungan = False Then Exit Sub
    If txtNamaFormPengirim.Text = "tampung" Then
        Call subTampungDataPenjamin
    Else
        If sp_JenisPasienJoinProgramAskes = False Then Exit Sub
        'add bayar ototmatis
        If mstrKdInstalasi = "02" Or mstrKdInstalasi = "06" Or mstrKdInstalasi = "11" Then
            Call subBayarOtomatis
        End If
'        'Call  f_UpdateJenisPasienJoinProgramAskes(dcPenjamin.BoundText,txtNoKartuPA.Text ,txtNoCM.Text ,txtNamaPA.Text ,txtNipPA,dcKelasDitanggung.BoundText ,Format(dtpTglLahirPA, "yyyy/MM/dd HH:mm:ss"),txtAlamatPA.Text ,IIf(Len(Trim(txtnopendaftaran.Text)) = 0, Null, txtnopendaftaran.Text),dcHubungan.BoundText,IIf(Len(Trim(txtNoSJP.Text)) = 0, Null, Trim(txtNoSJP.Text)),Format(dtpTglSJP, "yyyy/MM/dd hh:mm:ss"),strIDPegawaiAktif,IIf(Len(Trim(txtNoBP.Text)) = 0, Null, Trim(txtNoBP.Text)),IIf(Val(txtNoKunjungan.Text) = 0, "1", txtNoKunjungan.Text),IIf(chkNoSJP.Value = vbChecked, "O", "M"),Val(txtAnakKe.Text),IIf(Len(Trim(dcUnitKerja.Text)) = 0, Null, Trim(dcUnitKerja.Text)),null,txtNoRujukan.Text,
'        fTglLahir = Format(dtpTglLahirPA, "yyyy/MM/dd HH:mm:ss")
'        fNoPendaftaran = IIf(Len(Trim(txtNoPendaftaran.Text)) = 0, Null, txtNoPendaftaran.Text)
'        fNoSJP = IIf(Len(Trim(txtNoSJP.Text)) = "", Null, Trim(txtNoSJP.Text))
'        fNoBP = IIf(Len(Trim(txtNoBP.Text)) = 0, "Null", "'" & Trim(txtNoBP.Text) & "'")
'        fNoKunjungan = IIf(Val(txtNoKunjungan.Text) = 0, "1", txtNoKunjungan.Text)
'        fChkNoSJP = IIf(chkNoSJP.Value = vbChecked, "O", "M")
'        fDcUnitKerja = IIf(Len(Trim(dcUnitKerja.Text)) = 0, "Null", "'" & Trim(dcUnitKerja.Text) & "'")
'        fNamaAsalRujukan = IIf(Len(Trim(dcNamaAsalRujukan.Text)) = 0, "Null", "'" & dcNamaAsalRujukan.Text & "'")
'        fNamaPerujuk = IIf(Len(Trim(dcNamaPerujuk.Text)) = 0, "Null", "'" & Trim(dcNamaPerujuk.Text) & "'")
'        fDiagnosa = IIf(Len(Trim(dcDiagnosa.Text)) = 0, "Null", "'" & Trim(dcDiagnosa.Text) & "'")
'        fKdPerusahaan = IIf(Len(Trim(dcPerusahaan.Text)) = 0, "Null", "'" & Trim(dcPerusahaan.Text) & "'")
'       ' fAlamatPA = IIf(Len(Trim(txtAlamatPA.Text)) = 0, "Null", "'" & txtAlamatPA.Text & "'")
'        'fIDPeserta = IIf(Len(Trim(txtNipPA.Text)) = 0, "Null", "'" & txtNipPA.Text & "'")
'        'fKdGolongan = IIf(Len(Trim(dcKelasDitanggung.BoundText)) = 0, "Null", "'" & dcKelasDitanggung.BoundText & "'")
        
'        Call f_UpdateJenisPasienJoinProgramAskes(dcPenjamin.BoundText, txtNoKartuPA, txtNoCM.Text, txtNamaPA.Text, txtNipPA, dcKelasDitanggung.BoundText, fTglLahir, txtAlamatPA.Text, fNoPendaftaran, dcHubungan.BoundText, fNoSJP, Format(dtpTglSJP, "yyyy/MM/dd hh:mm:ss"), strIDPegawaiAktif, fNoBP, fNoKunjungan, fChkNoSJP, Val(txtAnakKe.Text), fDcUnitKerja, Null, txtNoRujukan.Text, dcAsalRujukan.BoundText, fNamaAsalRujukan, dcNamaAsalRujukan.BoundText, fNamaPerujuk, Format(dtpTglDirujuk.Value, "yyyy/MM/dd HH:mm:ss"), fDiagnosa, dcDiagnosa.BoundText, dcJenisPasien.BoundText, dcPerusahaan.BoundText)
 '       MsgBox "Ubah Data Pasien Sukses", vbInformation, "Informasi"
  '      cmdSimpan.Enabled = False
        
    End If
    MousePointer = vbDefault

'Exit Sub
'errload:
'    Call msubPesanError("cmdSimpan_Click")
'    MousePointer = vbDefault
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcAsalRujukan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNoRujukan.SetFocus
End Sub

Private Sub dcDiagnosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub dcGolonganasuransi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAlamatPA.SetFocus
End Sub

Private Sub dcHubungan_Change()
    txtAnakKe.Text = ""
    If dcHubungan.BoundText = "04" Then txtAnakKe.Enabled = True Else txtAnakKe.Enabled = False
End Sub

Private Sub dcKelasDitanggung_GotFocus()
On Error GoTo errload
Dim tempKode As String
    
    tempKode = dcKelasDitanggung.BoundText
    strSQL = "SELECT DISTINCT KdKelas, DeskKelas FROM V_KelasDitanggungPenjamin WHERE (IdPenjamin = '" & dcPenjamin.BoundText & "') AND KdKelompokPasien = '" & dcJenisPasien.BoundText & "'"
    Call msubDcSource(dcKelasDitanggung, rs, strSQL)
    dcKelasDitanggung.BoundText = tempKode

Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub dcKelasDitanggung_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcAsalRujukan.SetFocus
End Sub

Private Sub dcHubungan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If txtAnakKe.Enabled = True Then txtAnakKe.SetFocus Else txtNoSJP.SetFocus
End Sub

Private Sub dcJenisPasien_Change()
On Error GoTo errload
    Set rs = Nothing
    rs.Open "select * from v_Penjaminpasien where KdKelompokPasien='" & dcJenisPasien.BoundText & "' ORDER BY NamaPenjamin", dbConn, adOpenForwardOnly, adLockReadOnly
    
    Set dcPenjamin.RowSource = rs
    Set dcPerusahaan.RowSource = rs
    
    dcPenjamin.BoundColumn = rs.Fields("idpenjamin").Name
    dcPenjamin.ListField = rs.Fields("namapenjamin").Name
    dcPerusahaan.BoundColumn = rs.Fields("idpenjamin").Name
    dcPerusahaan.ListField = rs.Fields("namapenjamin").Name
    
    dcPenjamin.BoundText = ""
    dcPerusahaan.BoundText = ""

    Call msubRecFO(dbRst, "SELECT IdPenjamin FROM PenjaminKelompokPasien WHERE KdKelompokPasien = '" & dcJenisPasien.BoundText & "'")
    If dbRst(0).Value = "2222222222" Then
        fraDataKartuPeserta.Enabled = False
        fraPemakaianAsuransi.Enabled = False
        fraDataRujukan.Enabled = False
        dcPerusahaan.Text = ""
    Else
        fraDataKartuPeserta.Enabled = True
        fraPemakaianAsuransi.Enabled = True
        fraDataRujukan.Enabled = True
    End If

Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub dcJenisPasien_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then If Frame2.Enabled = True Then chkDiriSendiri.SetFocus Else cmdSimpan.SetFocus
End Sub

Private Sub dcNamaAsalRujukan_GotFocus()
On Error GoTo errload
Dim tempKode As String

    tempKode = dcNamaAsalRujukan.BoundText
    strSQL = "SELECT DetailRujukanAsal.KdDetailRujukanAsal, DetailRujukanAsal.DetailRujukanAsal" & _
        " FROM DetailRujukanAsal " & _
        " WHERE (KdRujukanAsal = '" & dcAsalRujukan.BoundText & "')"
    Call msubDcSource(dcNamaAsalRujukan, rs, strSQL)
    dcNamaAsalRujukan.BoundText = tempKode
    
Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub dcNamaAsalRujukan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglDirujuk.SetFocus
End Sub

Private Sub dcNamaPerujuk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcDiagnosa.SetFocus
End Sub

Private Sub dcPenjamin_Change()
On Error GoTo errload

    strSQL = "SELECT dbo.AsuransiPasien.IdPenjamin, dbo.AsuransiPasien.IdAsuransi, dbo.AsuransiPasien.NoCM, dbo.AsuransiPasien.NamaPeserta, " & _
            " dbo.AsuransiPasien.IDPeserta, dbo.AsuransiPasien.KdGolongan, dbo.AsuransiPasien.TglLahir, dbo.AsuransiPasien.Alamat," & _
            " dbo.AsuransiPasien.IdPerusahaan, dbo.Penjamin.NamaPenjamin AS NamaPerusahaan" & _
            " FROM dbo.AsuransiPasien INNER JOIN" & _
            " dbo.Penjamin ON dbo.AsuransiPasien.IdPerusahaan = dbo.Penjamin.IdPenjamin" & _
        " WHERE (AsuransiPasien.NoCM = '" & txtNoCM.Text & "') AND (AsuransiPasien.IdPenjamin = '" & dcPenjamin.BoundText & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        txtNoKartuPA.Text = IIf(IsNull(rs("IdAsuransi")), "", rs("IdAsuransi"))
        txtNamaPA.Text = IIf(IsNull(rs("NamaPeserta")), "", rs("NamaPeserta"))
        txtNipPA.Text = IIf(IsNull(rs("IDPeserta")), "-", rs("IDPeserta"))
        dcGolonganasuransi.BoundText = IIf(IsNull(rs("KdGolongan")), "", rs("KdGolongan"))
        dtpTglLahirPA.Value = IIf(IsNull(rs("TglLahir")), Now, rs("TglLahir"))
        txtAlamatPA.Text = IIf(IsNull(rs("Alamat")), "", rs("Alamat"))
        dcPerusahaan.Text = IIf(IsNull(rs("NamaPerusahaan")), "", rs("NamaPerusahaan"))
        Call subLoadPemakaianAsuransi(txtNoPendaftaran.Text, dcPenjamin.BoundText)
        dcHubungan.SetFocus
    Else
        txtNoKartuPA.Text = ""
        txtNamaPA.Text = ""
        txtNipPA.Text = ""
        dcGolonganasuransi.BoundText = ""
        dtpTglLahirPA.Value = Now
        txtAlamatPA.Text = ""
    End If

Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub dcPenjamin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcPerusahaan.SetFocus
End Sub

Private Sub dcPerusahaan_KeyPress(KeyAscii As Integer)
    On Error GoTo errload

    If KeyAscii = 13 Then
        If dcPerusahaan.MatchedWithList = True Then txtNoKartuPA.SetFocus
        strSQL = "SELECT  IdPenjamin, NamaPenjamin FROM dbo.Penjamin WHERE (NamaPenjamin LIKE '" & dcPerusahaan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcPerusahaan.BoundText = rs(0).Value
        dcPerusahaan.Text = rs(1).Value
    End If

Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub dcUnitKerja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcKelasDitanggung.SetFocus
End Sub

Private Sub dtpTglDirujuk_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcNamaPerujuk.SetFocus
End Sub

Private Sub dtpTglLahirPA_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtNipPA.SetFocus
End Sub

Private Sub dtpTglSJP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtNoBP.SetFocus
End Sub

'Private Sub Form_Activate()
'On Error GoTo errLoad
'Dim rsAs As ADODB.recordset
'
'    If mblnFormDaftarAntrian = True Then
'        txtNoCM.Text = mstrNoCM
'    End If
''    strSQL = "SELECT * FROM AsuransiPasien WHERE NoCM='" & txtNoCM.Text & "'"
''    msubRecFO rsAs, strSQL
''    If rsAs.RecordCount <> 0 Then
''        dcPenjamin.BoundText = rsAs("IdPenjamin").Value
''        dcPenjamin_Change
''    End If
'
'Exit Sub
'errLoad:
'    msubPesanError
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errload
    Call PlayFlashMovie(Me)
    dtpTglLahirPA.Value = Now
    dtpTglSJP.Value = Now
    dtpTglDirujuk.Value = Now
    
    txtNoBP.Text = ""
    txtNoKunjungan.Text = ""
    
    If mblnFormDaftarAntrian = True Then txtNoCM.Text = mstrNoCM
    txtNoPendaftaran = mstrNoPen
    Call centerForm(Me, MDIUtama)
    Call subLoadDCSource
        
Exit Sub
errload:
    msubPesanError
End Sub

Private Sub txtAlamatPA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcHubungan.SetFocus
End Sub

Private Sub txtAlamatPA_LostFocus()
    txtAlamatPA = StrConv(txtAlamatPA, vbProperCase)
End Sub

Private Sub txtAnakKe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNoSJP.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtNamaPA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglLahirPA.SetFocus
End Sub

Private Sub txtNamaPA_LostFocus()
    txtNamaPA = StrConv(txtNamaPA, vbProperCase)
End Sub

Private Sub txtNipPA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcGolonganasuransi.SetFocus
End Sub

Private Sub txtNipPA_LostFocus()
    txtNipPA = StrConv(txtNipPA, vbProperCase)
End Sub

Private Sub txtNoBP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcUnitKerja.SetFocus
End Sub

Private Sub txtNoBP_LostFocus()
'    If sp_AmbulNoKunjungan = False Then Exit Sub
End Sub

Private Sub txtNoKartuPA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtNamaPA.SetFocus
End Sub

Private Sub txtNoKartuPA_LostFocus()
On Error GoTo errload
Dim strKdGolongan As String
    strSQL = "SELECT * FROM AsuransiPasien " _
        & "WHERE IdPenjamin='" & dcPenjamin.BoundText & "' AND IdAsuransi='" _
        & txtNoKartuPA.Text & "' AND NoCM='" & txtNoCM.Text & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then Exit Sub
    txtNamaPA.Text = rs.Fields("NamaPeserta").Value
    If Not IsNull(rs.Fields("IDPeserta").Value) Then txtNipPA.Text = rs.Fields("IDPeserta").Value
    dtpTglLahirPA.Value = rs.Fields("TglLahir").Value
    strKdGolongan = rs.Fields("KdGolongan").Value
    If Not IsNull(rs.Fields("Alamat").Value) Then txtAlamatPA.Text = rs.Fields("Alamat").Value
    strSQL = "SELECT DISTINCT NamaGolongan,KdGolongan FROM GolonganAsuransi WHERE KdGolongan='" & strKdGolongan & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    dcKelasDitanggung.Text = rs.Fields(0).Value
    dcKelasDitanggung.BoundText = rs.Fields(1).Value
    Set rs = Nothing
    txtNoKartuPA = StrConv(txtNoKartuPA, vbProperCase)
Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub txtNoRujukan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcNamaAsalRujukan.SetFocus
End Sub

Private Sub txtNoSJP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'        If sp_AmbulNoKunjungan = False Then Exit Sub
        dtpTglSJP.SetFocus
    End If
End Sub

Private Sub txtNoSJP_LostFocus()
    txtNoSJP = StrConv(txtNoSJP, vbProperCase)
End Sub

Private Function sp_AmbulNoKunjungan() As Boolean
On Error GoTo errload
    sp_AmbulNoKunjungan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, dcPenjamin.BoundText)
        .Parameters.Append .CreateParameter("IdAsuransi", adChar, adParamInput, 15, Trim(txtNoKartuPA.Text))
        .Parameters.Append .CreateParameter("NoCM", adChar, adParamInput, 6, txtNoCM.Text)
        .Parameters.Append .CreateParameter("KunjunganKe", adInteger, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("TglRujukanOut", adDate, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("TglPendaftaran", adDate, adParamInput, , Format(txtTglPendaftaran.Text, "yyyy/MM/dd hh:mm:ss"))
        .Parameters.Append .CreateParameter("NoSJPRujukan", adVarChar, adParamInput, 30, Trim(txtNoSJP.Text))
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Check_NoRujukan"
        .CommandType = adCmdStoredProc
        .Execute
    
        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam pengambilan No Kunjungan", vbExclamation, "Validasi"
            sp_AmbulNoKunjungan = False
        Else
            txtNoKunjungan.Text = .Parameters("KunjunganKe").Value
'            dtpTglSJP.Value = .Parameters("TglRujukanOut").Value
            If txtNoKunjungan.Text = "0" Then
                MsgBox "Masa berlaku No. Rujukan (SJP) sudah HABIS", vbExclamation, "Informasi"
                sp_AmbulNoKunjungan = False
            ElseIf Val(txtNoKunjungan.Text) > 3 Then
                MsgBox "Masa kunjungan No. Rujukan (SJP) sudah lebih dari 3 kali", vbExclamation, "Informasi"
                sp_AmbulNoKunjungan = False
            Else
'                dtpTglSJP.Value = Now
            End If
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
Exit Function
errload:
    Call msubPesanError
    sp_AmbulNoKunjungan = False
End Function

Private Function sp_UpdateJenisPasienUmum(f_KdKelompokPasien As String, f_NoPendaftaran As String) As Boolean
On Error GoTo errload
    sp_UpdateJenisPasienUmum = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdKelompokpasien", adChar, adParamInput, 2, f_KdKelompokPasien)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)
                
        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_JenisPasienUmum"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_UpdateJenisPasienUmum = False
        End If
        Set dbcmd = Nothing
    End With
Exit Function
errload:
    sp_UpdateJenisPasienUmum = False
    Call msubPesanError
End Function

