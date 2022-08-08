VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPasienPulangRJ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pasien Pulang Poliklinik"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPasienPulangRJ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   11895
   Begin VB.CommandButton cmdValidasidata 
      Caption         =   "Validasi Data"
      Height          =   495
      Left            =   240
      TabIndex        =   56
      Top             =   8520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   8160
      TabIndex        =   15
      ToolTipText     =   "Tekan Tombol Validasi Data Terlebih Dahulu"
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   9975
      TabIndex        =   16
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Frame fraPasienDirujukKeluar 
      Caption         =   "Data Pasien Dirujuk Keluar"
      Enabled         =   0   'False
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
      Left            =   0
      TabIndex        =   47
      Top             =   5520
      Visible         =   0   'False
      Width           =   11895
      Begin VB.TextBox txtDirujukKe 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6360
         MaxLength       =   100
         TabIndex        =   11
         Top             =   600
         Width           =   4455
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         MaxLength       =   100
         TabIndex        =   14
         Top             =   2400
         Width           =   11535
      End
      Begin VB.TextBox txtAlasanDirujuk 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         MaxLength       =   200
         TabIndex        =   13
         Top             =   1800
         Width           =   11535
      End
      Begin VB.TextBox txtAlamatRujukan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         MaxLength       =   100
         TabIndex        =   12
         Top             =   1200
         Width           =   11535
      End
      Begin MSDataListLib.DataCombo dcDokterPerujuk 
         Height          =   330
         Left            =   2400
         TabIndex        =   10
         Top             =   600
         Width           =   3735
         _ExtentX        =   6588
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
      Begin MSComCtl2.DTPicker dtpTglDirujuk 
         Height          =   330
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   128450563
         UpDown          =   -1  'True
         CurrentDate     =   38085
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Dirujuk Ke"
         Height          =   210
         Index           =   3
         Left            =   6360
         TabIndex        =   53
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan"
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   52
         Top             =   2160
         Width           =   945
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Alasan Dirujuk"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   51
         Top             =   1560
         Width           =   1125
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Alamat Rujukan"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   50
         Top             =   960
         Width           =   1260
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Dirujuk"
         Height          =   210
         Left            =   240
         TabIndex        =   49
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Dokter Perujuk"
         Height          =   210
         Left            =   2400
         TabIndex        =   48
         Top             =   360
         Width           =   1230
      End
   End
   Begin VB.Frame framePasienPulang 
      Caption         =   "Pasien Pulang"
      Enabled         =   0   'False
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
      TabIndex        =   41
      Top             =   4440
      Visible         =   0   'False
      Width           =   11895
      Begin VB.CheckBox chkDirujukKeluar 
         Caption         =   "Dirujuk Keluar"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   9360
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtLamadiRS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         MaxLength       =   10
         TabIndex        =   4
         Top             =   600
         Width           =   880
      End
      Begin VB.TextBox txtPenerimaPasien 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   5
         Text            =   "Keluarganya"
         Top             =   600
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo dcStatusPulang2 
         Height          =   330
         Left            =   3120
         TabIndex        =   6
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
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
      Begin MSDataListLib.DataCombo dcKondisiPulang2 
         Height          =   330
         Left            =   6240
         TabIndex        =   7
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
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
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Cara Pulang"
         Height          =   210
         Left            =   3120
         TabIndex        =   45
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Lama di RS"
         Height          =   210
         Left            =   240
         TabIndex        =   44
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Nama Penerima Pasien"
         Height          =   210
         Left            =   1200
         TabIndex        =   43
         Top             =   360
         Width           =   1830
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Kondisi Pulang"
         Height          =   210
         Left            =   6240
         TabIndex        =   42
         Top             =   360
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pasien Keluar Kamar"
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
      TabIndex        =   34
      Top             =   2040
      Width           =   11895
      Begin VB.TextBox txtTglMasuk 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   40
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtLamaDirawat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo dcStatusKeluar 
         Height          =   330
         Left            =   4320
         TabIndex        =   2
         Top             =   600
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
      Begin MSComCtl2.DTPicker dtpTglKeluar 
         Height          =   330
         Left            =   2280
         TabIndex        =   0
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   128450563
         UpDown          =   -1  'True
         CurrentDate     =   38085
      End
      Begin VB.TextBox txtNoPemakaian 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         MaxLength       =   10
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo dcRuanganTujuan 
         Height          =   330
         Left            =   6240
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
      Begin MSDataListLib.DataCombo dcKondisiKeluar 
         Height          =   330
         Left            =   1560
         TabIndex        =   57
         Top             =   1320
         Visible         =   0   'False
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
      Begin MSDataListLib.DataCombo dcKondisiPulang 
         Height          =   330
         Left            =   9240
         TabIndex        =   60
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
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
      Begin MSDataListLib.DataCombo dcStatusPulang 
         Height          =   330
         Left            =   6240
         TabIndex        =   62
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
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
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Cara Pulang"
         Height          =   210
         Left            =   6240
         TabIndex        =   63
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Kondisi Pulang"
         Height          =   210
         Left            =   9240
         TabIndex        =   61
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Kondisi Keluar"
         Height          =   210
         Left            =   1560
         TabIndex        =   54
         Top             =   1080
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Ruangan Tujuan"
         Height          =   210
         Left            =   6240
         TabIndex        =   46
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Masuk"
         Height          =   210
         Left            =   240
         TabIndex        =   39
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Pulang"
         Height          =   210
         Left            =   2280
         TabIndex        =   38
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Status Pulang"
         Height          =   210
         Left            =   4320
         TabIndex        =   37
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Lama Dirawat"
         Height          =   210
         Left            =   240
         TabIndex        =   36
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "No. Pemakaian"
         Height          =   210
         Left            =   240
         TabIndex        =   35
         Top             =   600
         Width           =   1200
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
      TabIndex        =   18
      Top             =   960
      Width           =   11895
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   7200
         TabIndex        =   29
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3720
         TabIndex        =   28
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         MaxLength       =   12
         TabIndex        =   27
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         MaxLength       =   10
         TabIndex        =   26
         Top             =   600
         Width           =   1335
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
         Left            =   8640
         TabIndex        =   19
         Top             =   360
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   20
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   555
            TabIndex        =   25
            Top             =   270
            Width           =   240
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1350
            TabIndex        =   24
            Top             =   270
            Width           =   210
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2130
            TabIndex        =   23
            Top             =   270
            Width           =   150
         End
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   7200
         TabIndex        =   33
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   3720
         TabIndex        =   32
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   1680
         TabIndex        =   31
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.TextBox txtKdDiagnosa 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   2280
      TabIndex        =   55
      Top             =   8640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   58
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
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   8160
      ScaleHeight     =   435
      ScaleWidth      =   1755
      TabIndex        =   59
      ToolTipText     =   "Tekan Tombol Validasi Data Terlebih Dulu"
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Image Image5 
      Height          =   945
      Left            =   10080
      Picture         =   "frmPasienPulangRJ.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image4 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPasienPulangRJ.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10215
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   10080
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
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "frmPasienPulangRJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function sp_PasienDirujukKeluar(f_status As String) As Boolean
On Error GoTo Errload

    sp_PasienDirujukKeluar = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adChar, adParamInput, 6, txtnocm.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("TglDirujuk", adDate, adParamInput, , Format(dtpTglDirujuk.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("DirujukKe", adVarChar, adParamInput, 100, txtDirujukKe.Text)
        .Parameters.Append .CreateParameter("AlamatRujukan", adVarChar, adParamInput, 100, IIf(Len(Trim(txtAlamatRujukan.Text)) = 0, Null, Trim(txtAlamatRujukan.Text)))
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, dcDokterPerujuk.BoundText)
        .Parameters.Append .CreateParameter("AlasanDirujuk", adVarChar, adParamInput, 200, txtAlasanDirujuk.Text)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 100, IIf(Len(Trim(txtKeterangan.Text)) = 0, Null, Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)
                
        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_PasienDirujukKeluar"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada Kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_PasienDirujukKeluar = False
        Else
            Call Add_HistoryLoginActivity("AUD_PasienDirujukKeluar")
        End If
    End With
    Set dbcmd = Nothing

Exit Function
Errload:
    sp_PasienDirujukKeluar = False
    Call msubPesanError("sp_PasienDirujukKeluar")
End Function

Private Sub chkDirujukKeluar_Click()
On Error GoTo Errload
    
    If chkDirujukKeluar.Value = vbChecked Then
        dtpTglDirujuk.Value = Now
        fraPasienDirujukKeluar.Enabled = True
        Call msubDcSource(dcDokterPerujuk, rs, "SELECT KodeDokter, NamaDokter FROM V_DaftarDokter")
    Else
        fraPasienDirujukKeluar.Enabled = False
    End If
    
Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub chkDirujukKeluar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If chkDirujukKeluar.Value = vbChecked Then dtpTglDirujuk.SetFocus Else cmdSimpan.SetFocus
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo Errload
    If funcCekValidasi = False Then Exit Sub
    
    strSQL = "Select * from V_DaftarDiagnosaPasien where nocm = '" & mstrNoCM & "'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        Call SimpanKeluarKamar
        Exit Sub
    Else
        If MsgBox("Riwayat Diagnosa Kosong! Teruskan Simpan? ", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
        Call SimpanKeluarKamar
        Exit Sub
        End If
    End If
Errload:
End Sub

Private Sub cmdTutup_Click()
    If cmdSimpan.Enabled = True Then
        If MsgBox("Simpan data keluar pasien", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
            Exit Sub
        End If
    End If
    Unload Me
    
End Sub

'Private Sub cmdValidasiData_Click()
'    frmValidasiData.Show
'End Sub

Private Sub dcDokterPerujuk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDirujukKe.SetFocus
End Sub

Private Sub dcKondisiKeluar_GotFocus()
    strSQL = "SELECT KdKondisiKeluar,KondisiKeluar FROM KondisiKeluar where StatusEnabled='1' order by KondisiKeluar"
    Call msubDcSource(dcKondisiKeluar, dbRst, strSQL)
End Sub

Private Sub dcKondisiKeluar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcRuanganTujuan.SetFocus
End Sub

Private Sub dcKondisiPulang_GotFocus()
On Error GoTo Errload
    
'    strSQL = "SELECT KdKondisiPulang,KondisiPulang FROM KondisiPulang where KdKondisiPulang NOT IN('03','04','05') and StatusEnabled='1' order by KondisiPulang"
'    Call msubDcSource(dcKondisiPulang, dbRst, strSQL)

Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub dcKondisiPulang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub dcRuanganTujuan_GotFocus()
On Error GoTo Errload

    strSQL = "SELECT KdRuangan,NamaRuangan FROM Ruangan WHERE KdInstalasi IN ('03','08', '24') AND LokasiRuangan <>'Non Aktif' and StatusEnabled='1' order by NamaRuangan"
    Call msubDcSource(dcRuanganTujuan, dbRst, strSQL)

Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub dcRuanganTujuan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If framePasienPulang.Enabled = True Then
            txtPenerimaPasien.SetFocus
        Else
            cmdValidasidata.SetFocus
        End If
    End If
End Sub

Private Sub dcStatusKeluar_Change()
    txtLamadiRS.Text = ""
    txtPenerimaPasien.Text = ""
    dcStatusPulang.Text = ""
    dcKondisiPulang.Text = ""
    dcKondisiKeluar.Text = ""
    dcRuanganTujuan.Text = ""
    If dcStatusKeluar.BoundText <> "01" Then
       framePasienPulang.Enabled = True
       dcKondisiKeluar.Enabled = False
       dcRuanganTujuan.Enabled = False
       txtPenerimaPasien.Text = "Keluarganya"
    Else
       framePasienPulang.Enabled = False
       dcKondisiKeluar.Enabled = True
       dcRuanganTujuan.Enabled = True
    End If
End Sub

Private Sub dcStatusKeluar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcKondisiKeluar.Enabled = True Then
            dcStatusPulang.SetFocus
        ElseIf framePasienPulang.Enabled = True Then
            dcStatusPulang.SetFocus
        End If
    End If
End Sub

Private Sub dcStatusPulang_Change()
On Error GoTo Errload
    
    If dcStatusPulang.BoundText <> "01" And dcStatusPulang.BoundText <> "08" Then
        chkDirujukKeluar.Enabled = True
        chkDirujukKeluar.Value = vbUnchecked
    Else
        chkDirujukKeluar.Enabled = False
        chkDirujukKeluar.Value = vbUnchecked
    End If

Exit Sub
Errload:
    Call msubPesanError("dcStatusPulang_Change")
End Sub

Private Sub dcStatusPulang_GotFocus()
On Error GoTo Errload
Dim tempKode As String

    tempKode = dcStatusPulang.BoundText
    strSQL = "SELECT KdStatusPulang,StatusPulang FROM StatusPulang where StatusEnabled='1' order by StatusPulang"
    Call msubDcSource(dcStatusPulang, dbRst, strSQL)
    dcStatusPulang.BoundText = tempKode
    
Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub dcStatusPulang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcKondisiPulang.SetFocus
End Sub

Private Sub dtpTglDirujuk_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcDokterPerujuk.SetFocus
End Sub

Private Sub dtpTglKeluar_Change()
    dtpTglKeluar.MaxDate = Now
End Sub

Private Sub dtpTglKeluar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcStatusKeluar.SetFocus
End Sub

Private Sub dtpTglKeluar_LostFocus()
On Error GoTo Errload
    If dtpTglKeluar.Value < CDate(txtTglMasuk.Text) Then
        MsgBox "Tanggal keluar tidak boleh melebihi tanggal masuk"
        dtpTglKeluar.Value = txtTglMasuk.Text
    End If
Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub Form_Load()
On Error GoTo Errload
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    framePasienPulang.Enabled = False
    dcRuanganTujuan.Enabled = False
    dtpTglKeluar.Value = Now
  
    
    strSQL = "SELECT KdStatusKeluar,StatusKeluar FROM StatusKeluarKamar where KdStatusKeluar in('02','06','07','11') and StatusEnabled='1' order by StatusKeluar"
    Call msubDcSource(dcStatusKeluar, dbRst, strSQL)
    dcStatusKeluar.BoundText = "02"
    
    strSQL = "SELECT KdStatusPulang,StatusPulang FROM StatusPulang where StatusEnabled='1' order by StatusPulang"
    Call msubDcSource(dcStatusPulang, dbRst, strSQL)
    dcStatusPulang.BoundText = "02"
    
    strSQL = "SELECT KdKondisiPulang,KondisiPulang FROM KondisiPulang where KdKondisiPulang NOT IN('03','04','05') and StatusEnabled='1' order by KondisiPulang"
    Call msubDcSource(dcKondisiPulang, dbRst, strSQL)
    dcKondisiPulang.BoundText = "01"

    
Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Errload
    
    frmDaftarPasienRJ.Enabled = True
  '  Call frmDaftarPasienRJ.optPasNonAktif_GotFocus
  '  Call frmDaftarPasienRJ.optPasAktif_GotFocus
    
Exit Sub
Errload:
    Call msubPesanError
End Sub

'untuk mencek validasi
Private Function funcCekValidasi() As Boolean
    funcCekValidasi = False
    If dcStatusKeluar.Text = "" Then
        MsgBox "Status pulang pasien harus diisi", vbCritical, "Validasi"
        dcStatusKeluar.SetFocus
        Exit Function
    End If
    If dcStatusKeluar.BoundText <> "01" Then
        If txtPenerimaPasien.Text = "" Then
            MsgBox "Penerima pasien harus diisi", vbCritical, "Validasi"
            txtPenerimaPasien.SetFocus
            Exit Function
        End If
        If dcStatusPulang.Text = "" Then
            MsgBox "Cara pulang pasien harus diisi", vbCritical, "Validasi"
            dcStatusPulang.SetFocus
            Exit Function
        End If
        If dcKondisiPulang.Text = "" Then
            MsgBox "Kondisi pulang pasien harus diisi", vbCritical, "Validasi"
            dcKondisiPulang.SetFocus
            Exit Function
        End If
        If chkDirujukKeluar.Value = vbChecked Then
            If Periksa("datacombo", dcDokterPerujuk, "Dokter perujuk kosong") = False Then Exit Function
            If Periksa("text", txtDirujukKe, "Tempat tujuan rujukan kosong") = False Then Exit Function
            If Periksa("text", txtAlasanDirujuk, "Alasan dirujuk kosong") = False Then Exit Function
        End If
    End If
    funcCekValidasi = True
End Function

'untuk enable/disable control2
Private Sub subDisableControl(blnStatus As Boolean)
    dtpTglKeluar.Enabled = blnStatus
    dcStatusKeluar.Enabled = blnStatus
    cmdSimpan.Enabled = blnStatus
End Sub

'untuk save pasien keluar kamar
' rawat jalan g keluar kamar___Parja
Public Sub subSavePsnKelKam(f_NoPakai As String)

End Sub

'untuk save pasien keluar kamar
Public Sub subSavePsnPulang()
On Error GoTo Errload
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, mstrNoCM)
        .Parameters.Append .CreateParameter("TglPulang", adDate, adParamInput, , Format(dtpTglKeluar.Value, "yyyy-MM-dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("NamaPenerima", adVarChar, adParamInput, 30, txtPenerimaPasien.Text)
        .Parameters.Append .CreateParameter("KdKondisiPulang", adChar, adParamInput, 2, dcKondisiPulang.BoundText)
        .Parameters.Append .CreateParameter("KdStatusPulang", adChar, adParamInput, 2, dcStatusPulang.BoundText)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, noidpegawai)
        .Parameters.Append .CreateParameter("OutputLamaRawat", adInteger, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("KdRuanganLogin", adChar, adParamInput, 3, mstrKdRuangan)
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PasienRJPulang"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam penyimpanan data pasien pulang", vbCritical, "Validasi"
            Call deleteADOCommandParameters(dbcmd)
            Set dbcmd = Nothing
            Exit Sub
        Else
            If Not IsNull(.Parameters("OutputLamaRawat").Value) Then _
                txtLamadiRS.Text = .Parameters("OutputLamaRawat").Value
'            MsgBox "Penyimpanan data pasien pulang sukses", vbinformation, "Informasi"

            End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub txtAlamatRujukan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAlasanDirujuk.SetFocus
End Sub

Private Sub txtAlasanDirujuk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
End Sub

Private Sub txtDirujukKe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAlamatRujukan.SetFocus
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtPenerimaPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcStatusPulang.SetFocus
End Sub

Public Sub SimpanKeluarKamar()
On Error GoTo Errload
'/---------------------------\
'---Yang dirubah 2004-08-30---
    If dcKondisiPulang.BoundText = "04" Or dcKondisiPulang.BoundText = "05" Then
        If mblnPsnMati = False Then
            With frmPasienMeninggal
                .Show
                .txtnopendaftaran.Text = mstrNoPen
                .txtnocm.Text = mstrNoCM
                .txtNamaPasien.Text = txtNamaPasien.Text
                .txtSex.Text = txtSex.Text
                .txtThn.Text = txtThn.Text
                .txtBln.Text = txtBln.Text
                .txtHari.Text = txtHari.Text
            End With
            Me.Enabled = False
            Exit Sub
        End If
    End If
    
    
    Call subSavePsnKelKam(txtNoPemakaian.Text)
    If dcStatusKeluar.BoundText <> "01" Then
        Call subSavePsnPulang
        If dcStatusPulang.BoundText <> "01" And dcStatusPulang.BoundText <> "08" Then
            If chkDirujukKeluar.Value = vbChecked Then If sp_PasienDirujukKeluar("A") = False Then Exit Sub
        End If
    End If
        
    
    strSQL = "SELECT NoPakai FROM PasienKeluarKamar WHERE (NoPakai = '" & txtNoPemakaian.Text & "')"
    'Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
        strSQL = "SELECT NoPakai FROM PemakaianKamar WHERE (NoPendaftaran = '" & txtnopendaftaran.Text & "') AND StatusKeluar='T'"
        Call msubRecFO(rs, strSQL)
        'Call subSavePsnKelKam(rs(0))
    End If
    'add onede
    strSQL = "Update PasienDaftar SET KdRuanganAkhir ='" & mstrKdRuangan & "' WHERE NoPendaftaran ='" & txtnopendaftaran.Text & "' "
    dbConn.Execute strSQL
    
    If dcStatusKeluar.BoundText <> "01" Then
        Call Add_HistoryLoginActivity("Add_PasienKeluarKamar+Add_PasienRIPulang")
    Else
        Call Add_HistoryLoginActivity("Add_PasienKeluarKamar")
    End If
    '---
    Call subDisableControl(False)
    mblnPsnMati = False

Exit Sub
Errload:
    Call msubPesanError
End Sub

