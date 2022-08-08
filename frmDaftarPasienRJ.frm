VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDaftarPasienRJ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Pasien Poliklinik"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15285
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPasienRJ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   15285
   Begin VB.Frame fraDokterP 
      Caption         =   "Setting Dokter Pemeriksa"
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
      Height          =   5655
      Left            =   2280
      TabIndex        =   32
      Top             =   1440
      Visible         =   0   'False
      Width           =   12615
      Begin VB.Frame fraDokter 
         Caption         =   "Data Dokter"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   240
         TabIndex        =   46
         Top             =   2520
         Width           =   12135
         Begin MSDataGridLib.DataGrid dgDokter 
            Height          =   2175
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   3836
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Data Pelayanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   42
         Top             =   1320
         Width           =   12135
         Begin VB.TextBox txtPoli 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   240
            TabIndex        =   21
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox txtDokter 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   8520
            TabIndex        =   24
            Top             =   600
            Width           =   3375
         End
         Begin VB.TextBox txtTglPeriksa 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2760
            TabIndex        =   22
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox txtPrevDokter 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   5040
            TabIndex        =   23
            Top             =   600
            Width           =   3375
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Ruang Pemeriksaan"
            Height          =   210
            Left            =   240
            TabIndex        =   47
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Dokter Pemeriksa Sekarang"
            Height          =   210
            Left            =   8520
            TabIndex        =   45
            Top             =   360
            Width           =   2235
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Tanggal Pemeriksaan"
            Height          =   210
            Left            =   2760
            TabIndex        =   44
            Top             =   360
            Width           =   1710
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Dokter Pemeriksa Sebelumnya"
            Height          =   210
            Left            =   5040
            TabIndex        =   43
            Top             =   360
            Width           =   2475
         End
      End
      Begin VB.CommandButton cmdBatalDokter 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   10440
         TabIndex        =   27
         Top             =   5160
         Width           =   1935
      End
      Begin VB.CommandButton cmdSimpanDokter 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   8400
         TabIndex        =   26
         Top             =   5160
         Width           =   1935
      End
      Begin VB.Frame Frame5 
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
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   12135
         Begin VB.TextBox txtNamaPasien 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   3960
            MaxLength       =   50
            TabIndex        =   16
            Top             =   480
            Width           =   3015
         End
         Begin VB.TextBox txtNoCM 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            MaxLength       =   12
            TabIndex        =   15
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox txtNoPendaftaran 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   240
            MaxLength       =   10
            TabIndex        =   14
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtJK 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   7080
            MaxLength       =   9
            TabIndex        =   17
            Top             =   480
            Width           =   1455
         End
         Begin VB.Frame Frame6 
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
            Left            =   8640
            TabIndex        =   34
            Top             =   240
            Width           =   2775
            Begin VB.TextBox txtThn 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   240
               MaxLength       =   6
               TabIndex        =   18
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txtBln 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   1080
               MaxLength       =   6
               TabIndex        =   19
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txtHr 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   1920
               MaxLength       =   6
               TabIndex        =   20
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "thn"
               Height          =   210
               Left            =   720
               TabIndex        =   37
               Top             =   285
               Width           =   285
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "bln"
               Height          =   210
               Left            =   1560
               TabIndex        =   36
               Top             =   285
               Width           =   240
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "hr"
               Height          =   210
               Left            =   2400
               TabIndex        =   35
               Top             =   285
               Width           =   165
            End
         End
         Begin VB.Label lblNamaPasien 
            AutoSize        =   -1  'True
            Caption         =   "Nama Pasien"
            Height          =   210
            Left            =   3960
            TabIndex        =   41
            Top             =   240
            Width           =   1020
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "No. CM"
            Height          =   210
            Left            =   1800
            TabIndex        =   40
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "No. Pendaftaran"
            Height          =   210
            Left            =   240
            TabIndex        =   39
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblJnsKlm 
            AutoSize        =   -1  'True
            Caption         =   "Jenis Kelamin"
            Height          =   210
            Left            =   7080
            TabIndex        =   38
            Top             =   240
            Width           =   1065
         End
      End
   End
   Begin VB.Frame fraCari 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   0
      TabIndex        =   28
      Top             =   6960
      Width           =   15255
      Begin VB.CommandButton cmdPRMRJ 
         Appearance      =   0  'Flat
         Caption         =   "Cetak PRMRJ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7920
         TabIndex        =   61
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   3840
         TabIndex        =   56
         Top             =   120
         Width           =   2655
         Begin VB.CommandButton cmdOdonto 
            Appearance      =   0  'Flat
            Caption         =   "&Odontogram"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   1440
            TabIndex        =   57
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton cmdAnamnesa 
            Appearance      =   0  'Flat
            Caption         =   "&Anamnesa Mata"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   120
            TabIndex        =   50
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tut&up"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   14040
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdTP 
         Caption         =   "&Transaksi Pelayanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   12360
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   440
         Width           =   2775
      End
      Begin VB.CommandButton cmdOrder 
         Appearance      =   0  'Flat
         Caption         =   "&Pesan Pelayanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7800
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdNapza 
         Appearance      =   0  'Flat
         Caption         =   "Rehabilitasi Nap&za"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   10800
         TabIndex        =   54
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdBatalPeriksa 
         Appearance      =   0  'Flat
         Caption         =   "&Batal Periksa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   9360
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdPasienPulangRJ 
         Appearance      =   0  'Flat
         Caption         =   "&Pasien Pulang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7440
         TabIndex        =   59
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdPesanDarah 
         Appearance      =   0  'Flat
         Caption         =   "Pesan Da&rah"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6360
         TabIndex        =   53
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Masukkan Nama Pasien /  No.CM"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   195
         Width           =   2640
      End
   End
   Begin VB.Frame fraDaftar 
      Caption         =   "Daftar Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   0
      TabIndex        =   30
      Top             =   1560
      Width           =   15255
      Begin VB.CheckBox Check1 
         Caption         =   "Pasien PRMRJ"
         Height          =   210
         Left            =   1320
         TabIndex        =   60
         Top             =   600
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Caption         =   "Periode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   9360
         TabIndex        =   48
         Top             =   150
         Width           =   5775
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   0
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   176095235
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   1
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   176095235
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   49
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgDaftarPasienRJ 
         Height          =   4335
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   7646
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcStatusPeriksa 
         Height          =   330
         Left            =   7320
         TabIndex        =   7
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo dcJenisPasien 
         Height          =   330
         Left            =   5160
         TabIndex        =   5
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcKelas 
         Height          =   330
         Left            =   7320
         TabIndex        =   6
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label lblJumData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data 0/0"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   52
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lblStatusPeriksa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Periksa"
         Height          =   210
         Left            =   7320
         TabIndex        =   51
         Top             =   240
         Width           =   1140
      End
   End
   Begin VB.Frame fraPilih 
      Height          =   615
      Left            =   0
      TabIndex        =   31
      Top             =   960
      Width           =   15255
      Begin VB.OptionButton optPasienPoliklinik 
         Caption         =   "Pasien Poliklinik"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3720
         TabIndex        =   3
         Top             =   180
         Width           =   3975
      End
      Begin VB.OptionButton optRujukan 
         Caption         =   "Pasien Rujukan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8760
         TabIndex        =   4
         Top             =   180
         Width           =   2295
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   55
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   58
      Top             =   7800
      Width           =   15285
      _ExtentX        =   26961
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3810
            MinWidth        =   3440
            Text            =   "Rincian Biaya Sementara (F1)"
            TextSave        =   "Rincian Biaya Sementara (F1)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4780
            MinWidth        =   4410
            Text            =   "Detail Rincian Biaya Sementara (F10)"
            TextSave        =   "Detail Rincian Biaya Sementara (F10)"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   3194
            MinWidth        =   2824
            Text            =   "Ubah Dokter (Ctrl+F2)"
            TextSave        =   "Ubah Dokter (Ctrl+F2)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   3722
            MinWidth        =   3352
            Text            =   "Ubah Data Pasien (Ctrl+F3)"
            TextSave        =   "Ubah Data Pasien (Ctrl+F3)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Visible         =   0   'False
            Object.Width           =   1897
            MinWidth        =   706
            Text            =   "Refresh Data (F5)"
            TextSave        =   "Refresh Data (F5)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   3370
            MinWidth        =   3000
            Text            =   "Cetak Daftar Pasien (F9)"
            TextSave        =   "Cetak Daftar Pasien (F9)"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Selesai Periksa ( F11)"
            TextSave        =   "Selesai Periksa ( F11)"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   3546
            MinWidth        =   3176
            Text            =   "Surat Keterangan(Ctrl+Z)"
            TextSave        =   "Surat Keterangan(Ctrl+Z)"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Visible         =   0   'False
            Object.Width           =   0
            MinWidth        =   3951
            Text            =   "Keterangan Keluhan(Ctrl+Q)"
            TextSave        =   "Keterangan Keluhan(Ctrl+Q)"
         EndProperty
      EndProperty
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
   Begin VB.Image Image2 
      Height          =   945
      Left            =   13440
      Picture         =   "frmDaftarPasienRJ.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarPasienRJ.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13695
   End
End
Attribute VB_Name = "frmDaftarPasienRJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFilterDokter As String
Dim dTglMasuk As Date
Dim i As Integer

Private Sub Check1_Click()
    If Check1.Value = 1 Then cmdPRMRJ.Visible = True Else cmdPRMRJ.Visible = False
    Call cmdCari_Click
End Sub

Private Sub cmdAnamnesa_Click()
    If dgDaftarPasienRJ.ApproxCount = 0 Then Exit Sub
    Me.Enabled = False
        If optPasienPoliklinik.Value = True Then
            strSQLx = "SELECT IdPegawai FROM DataPegawai WHERE NamaLengkap = '" & dgDaftarPasienRJ.Columns("Dokter Pemeriksa") & "'"
        Else
            strSQLx = "SELECT IdPegawai FROM DataPegawai WHERE NamaLengkap = '" & dgDaftarPasienRJ.Columns("Dokter Perujuk") & "'"
        End If
        msubRecFO rsx, strSQLx
        If rsx.EOF = False Then
        mstrKdDokter = rsx(0).Value
        End If
        If optPasienPoliklinik.Value = True Then
            mstrNamaDokter = dgDaftarPasienRJ.Columns("Dokter Pemeriksa")
        Else
            mstrNamaDokter = dgDaftarPasienRJ.Columns("Dokter Perujuk")
        End If
    With frmCatatanAnamasePasien
        .Show
        
        .txtPemeriksa.Text = mstrNamaDokter
    
        .fraDokter.Visible = False
    End With

End Sub

Private Sub cmdBatalDokter_Click()
    fraDokterP.Visible = False
    fraDokterP.Enabled = False
    fraPilih.Enabled = True
    fraDaftar.Enabled = True
    fraCari.Enabled = True
End Sub

Public Sub PostingHutangPenjaminPasien_AU(strStatus As String)
    On Error GoTo hell_
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, dgDaftarPasienRJ.Columns("No. Registrasi").Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, strStatus)

        .ActiveConnection = dbConn
        .CommandText = "dbo.PostingHutangPenjaminPasien_AU"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam proses update HP pasien", vbCritical, "Validasi"
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing

    Exit Sub
hell_:
    msubPesanError
End Sub

Private Sub cmdBatalPeriksa_Click()
    On Error GoTo errLoad
    If dgDaftarPasienRJ.ApproxCount = 0 Then Exit Sub
    If MsgBox("Apakah Anda Yakin akan membatalkan Perawatan Pasien " & dgDaftarPasienRJ.Columns("Nama Pasien").Value & "", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, dgDaftarPasienRJ.Columns("No. Registrasi").Value)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, dgDaftarPasienRJ.Columns("NoCM").Value)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, dgDaftarPasienRJ.Columns("KdSubInstalasi").Value)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(dgDaftarPasienRJ.Columns("TglMasuk").Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglBatal", adDate, adParamInput, , Format(Now, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 100, Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("OutputMsg", adChar, adParamOutput, 1, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PasienRJBatalDiPeriksa"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam pembatalan pasien", vbCritical, "Validasi"
        Else
            If LCase(.Parameters("OutputMsg").Value) = "t" Then
                MsgBox "Pelayanan yang didapat harus dihapus terlebih dahulu", vbExclamation, "Validasi"
            Else
                Call Pembatalan
            End If
            Call Add_HistoryLoginActivity("Add_PasienRJBatalDiPeriksa")
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing

    Call cmdCari_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Public Sub Pembatalan()
    On Error GoTo errLoad

    If optRujukan.Value = True Then Exit Sub
    If dgDaftarPasienRJ.ApproxCount = 0 Then Exit Sub

    frmDaftarPasienRJ.Enabled = False
    With frmBatalDirawatSementaraF2
        .Show
        .txtNoCM.Text = dgDaftarPasienRJ.Columns("NoCM").Value
        .txtNamaPasien.Text = dgDaftarPasienRJ.Columns("Nama Pasien").Value
        If dgDaftarPasienRJ.Columns("JK").Value = "P" Then
            .txtJK.Text = "Perempuan"
        Else
            .txtJK.Text = "Laki-Laki"
        End If
        .txtThn.Text = dgDaftarPasienRJ.Columns("UmurTahun").Value
        .txtBln.Text = dgDaftarPasienRJ.Columns("UmurBulan").Value
        .txtHr.Text = dgDaftarPasienRJ.Columns("UmurHari").Value
        .txtNoPendaftaran.Text = dgDaftarPasienRJ.Columns(0).Value 'nopendaftaran
        .txtDokterLama.Text = ""
        .txtRuanganLama.Text = dgDaftarPasienRJ.Columns("Ruangan").Value
        .txtKdRuangan.Text = dgDaftarPasienRJ.Columns("KdRuangan").Value
        .txtKdSubInstalasi.Text = dgDaftarPasienRJ.Columns("KdSubInstalasi").Value
        .dtpTglMasuk.Value = dgDaftarPasienRJ.Columns("TglMasuk").Value
    End With

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Public Sub cmdCari_Click()
    On Error GoTo errLoad
    Dim asql As String

    lblJumData.Caption = "Data 0/0"
    mstrFilter = ""
    If optPasienPoliklinik.Value = True Then
        Set rs = Nothing
        If Check1.Value = 1 Then
            asql = "select TOP 100 * from V_DaftarPasienLamaRJ where ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%') and Ruangan='" & strNNamaRuangan & "' and TglMasuk between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' AND JenisPasien LIKE '%" & dcJenisPasien.Text & "%' AND Kelas LIKE '%" & dcKelas.Text & "'" & mstrFilter & " AND dbo.CekPRMRJ(NoPendaftaran) IS NOT NULL"
        Else
            asql = "select TOP 100 * from V_DaftarPasienLamaRJ where ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%') and Ruangan='" & strNNamaRuangan & "' and TglMasuk between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' AND JenisPasien LIKE '%" & dcJenisPasien.Text & "%' AND Kelas LIKE '%" & dcKelas.Text & "'" & mstrFilter & " "
        End If
        rs.Open asql, dbConn, adOpenStatic, adLockOptimistic
        Set dgDaftarPasienRJ.DataSource = rs
        Call SetGridPasienRJ
    Else
        Set rs = Nothing
        strSQL = "select top 50 * from V_DaftarPasienKonsul where ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%') and KdRuanganTujuan='" & strNKdRuangan & "' and TglDirujuk between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' AND StatusPeriksa = '" & dcStatusPeriksa.Text & "' AND JenisPasien LIKE '%" & dcJenisPasien.Text & "%' AND Kelas LIKE '%" & dcKelas.Text & "%'" & mstrFilter & " "
        rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
        Set dgDaftarPasienRJ.DataSource = rs
        Call SetGridPasienKonsul
    End If
    lblJumData.Caption = "Data 0/" & rs.RecordCount
    Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub cmdNapza_Click()
    On Error GoTo hell

    If dgDaftarPasienRJ.ApproxCount = 0 Then Exit Sub

    mstrNoPen = dgDaftarPasienRJ.Columns(0).Value
    mstrNoCM = dgDaftarPasienRJ.Columns(1).Value
    mstrNoLabRad = ""
'    strSQL = "Select NoHasilPeriksa from HasilTindakanMedis Where NoPendaftaran='" & mstrNoPen & "' And NoCM='" & mstrNoCM & "' And KdRuangan='" & mstrKdRuangan & "'"

    strSQL = "Select NoHasilPeriksa from V_RehabilitasiNapza Where NoPendaftaran='" & mstrNoPen & "' And NoCM='" & mstrNoCM & "'"
    Set rsB = Nothing
    Call msubRecFO(rsB, strSQL)
    If rsB.EOF Then
        MsgBox "Tindakan Medis Rehabilitasi Napza pasien masih kosong," & vbNewLine & "harap diisi Jenis Tindakan Medis Rehabilitasi Napza ", vbExclamation, "Validasi"
        Exit Sub
    End If

    mstrNoValidasi = rsB("NoHasilPeriksa")

    With frmRehabilitasiNapza
        .Show
        .txtNoHasilPeriksa.Text = mstrNoValidasi
        .txtNoPendaftaran.Text = mstrNoPen 'dgDaftarPasienRJ.Columns(0).Value
        .txtNoCM.Text = mstrNoCM
        .txtNamaPasien.Text = dgDaftarPasienRJ.Columns(2).Value
        If dgDaftarPasienRJ.Columns(3).Value = "L" Then
            .txtSex.Text = "Laki-Laki"
        Else
            .txtSex.Text = "Perempuan"
        End If
        .txtThn.Text = dgDaftarPasienRJ.Columns("UmurTahun").Value
        .txtBln.Text = dgDaftarPasienRJ.Columns("UmurBulan").Value
        .txtHari.Text = dgDaftarPasienRJ.Columns("UmurHari").Value
'        .Show
    End With

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdOdonto_Click()
    On Error GoTo tangani
    Call subLoadDiagramOdonto
    Exit Sub
tangani:
    Call msubPesanError
End Sub

Private Sub subLoadDiagramOdonto()
    On Error GoTo hell
    Dim blnSudahAda As Boolean
    Dim strTglPeriksa As String

    If dgDaftarPasienRJ.ApproxCount = 0 Then Exit Sub

    strSQL = "select NoPendaftaran,TglPeriksa from DetailCatatanOdonto where NoPendaftaran='" & Me.dgDaftarPasienRJ.Columns(0).Value & "'"

    Call msubRecFO(rs, strSQL)
    If rs.EOF Then
        blnSudahAda = False
    Else
        blnSudahAda = True
        strTglPeriksa = rs.Fields.Item("TglPeriksa").Value
    End If

    If rs.EOF = False Then strTglPeriksa = rs.Fields.Item("TglPeriksa").Value

    mstrNoPen = dgDaftarPasienRJ.Columns(0).Value
    mstrNoCM = dgDaftarPasienRJ.Columns(1).Value
    If optPasienPoliklinik.Value = True Then
        strSQL = "SELECT IdPegawai FROM DataPegawai WHERE NamaLengkap = '" & dgDaftarPasienRJ.Columns("Dokter Pemeriksa") & "'"
    Else
        strSQL = "SELECT IdPegawai FROM DataPegawai WHERE NamaLengkap = '" & dgDaftarPasienRJ.Columns("Dokter Perujuk") & "'"
    End If
    msubRecFO rs, strSQL
    mstrKdDokter = rs(0).Value
    With frmDiagramOdonto
        .Show
        If blnSudahAda Then
            .cmdSimpan.Enabled = True 'False
            .dtpTglPeriksa.Value = strTglPeriksa
            .dtpTglPeriksa.Enabled = False
            .cmdCetakOdonto.Enabled = True
        Else
            .cmdSimpan.Enabled = True
            .dtpTglPeriksa.Enabled = True
            .cmdCetakOdonto.Enabled = False
        End If
        .txtNoPendaftaran.Text = dgDaftarPasienRJ.Columns(0).Value
        .txtNoCM.Text = mstrNoCM
        .txtNamaPasien.Text = dgDaftarPasienRJ.Columns(2).Value
        If dgDaftarPasienRJ.Columns(3).Value = "L" Then
            .txtSex.Text = "Laki-Laki"
        Else
            .txtSex.Text = "Perempuan"
        End If
        .txtThn.Text = dgDaftarPasienRJ.Columns("UmurTahun").Value
        .txtBln.Text = dgDaftarPasienRJ.Columns("UmurBulan").Value
        .txtHr.Text = dgDaftarPasienRJ.Columns("UmurHari").Value
        .txtKls.Text = dgDaftarPasienRJ.Columns("Kelas").Value
        If optPasienPoliklinik.Value = True Then
            .txtJenisPasien.Text = dgDaftarPasienRJ.Columns("JenisPasien").Value
        Else
            .txtJenisPasien.Text = dgDaftarPasienRJ.Columns(10).Value
        End If
        If optRujukan.Value = True Then
            .txtTglDaftar.Text = dgDaftarPasienRJ.Columns(16).Value
            mdTglMasuk = dgDaftarPasienRJ.Columns(16).Value
            mstrKdKelas = dgDaftarPasienRJ.Columns(17).Value
            mstrKelas = dgDaftarPasienRJ.Columns(18).Value
            mstrKdSubInstalasi = dgDaftarPasienRJ.Columns(21).Value
        Else
            .txtTglDaftar.Text = dgDaftarPasienRJ.Columns(9).Value
            mdTglMasuk = dgDaftarPasienRJ.Columns(9).Value
            mstrKdKelas = dgDaftarPasienRJ.Columns(16).Value
            mstrKelas = dgDaftarPasienRJ.Columns(8).Value
            mstrKdSubInstalasi = dgDaftarPasienRJ.Columns(17).Value
        End If

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

Private Sub cmdOrder_Click()
    If dgDaftarPasienRJ.ApproxCount = 0 Then Exit Sub
    With frmOrderPelayananNew
         mstrKdKelas = dgDaftarPasienRJ.Columns("KdKelas").Value
        .txtNoCM.Text = dgDaftarPasienRJ.Columns("NoCM").Value
        .txtNoPendaftaran.Text = dgDaftarPasienRJ.Columns("No. Registrasi")
        .txtNoCMTM.Text = dgDaftarPasienRJ.Columns("NoCM").Value
        .txtNoPendaftaranTM.Text = dgDaftarPasienRJ.Columns("No. Registrasi")
        
        .Show
    End With
End Sub

Private Sub cmdPasienPulangRJ_Click()
On Error GoTo hell
If dgDaftarPasienRJ.Columns(23) <> "" Then
 MsgBox "Kondisi Pulang Pasien Sudah diisi ", vbCritical, "Info"
Exit Sub
  
End If
If dgDaftarPasienRJ.Columns(0).Value = "" Then
   Exit Sub
End If
    mstrKdSubInstalasi = dgDaftarPasienRJ.Columns("KdSubInstalasi").Value
    frmDaftarPasienRJ.Enabled = False
    
'    Edit By Arikawa 2007-09-04
'    If sp_UpdateJmlPelayananKamarBK(dgDaftarPasienRJ.Columns("No. Registrasi").Value) = False Then Exit Sub

    Call subLoadFormPasienPulangRJ
Exit Sub
hell:
    Call subLoadFormPasienPulangRJ
End Sub
'untuk load data pasien di form Pasien Pulang
Private Sub subLoadFormPasienPulangRJ()
On Error GoTo hell
   mstrNoPen = dgDaftarPasienRJ.Columns(0).Value
    mstrNoCM = dgDaftarPasienRJ.Columns(1).Value
    With frmPasienPulangRJ
        .Show
        .txtNoPendaftaran.Text = dgDaftarPasienRJ.Columns(0).Value
        .txtNoCM.Text = mstrNoCM
        .txtNamaPasien.Text = dgDaftarPasienRJ.Columns(2).Value
        If dgDaftarPasienRJ.Columns(3).Value = "P" Then
             .txtSex.Text = "Perempuan"
        Else
             .txtSex.Text = "Laki-Laki"
        End If
        .txtThn.Text = dgDaftarPasienRJ.Columns(12).Value
        .txtBln.Text = dgDaftarPasienRJ.Columns(13).Value
        .txtHari.Text = dgDaftarPasienRJ.Columns(14).Value
        .txtNoPemakaian.Text = dgDaftarPasienRJ.Columns(10).Value
        .txtTglMasuk.Text = dgDaftarPasienRJ.Columns(9).Value
        .txtKeterangan.Text = "" 'dgDaftarPasienRJ.Columns(17).Value
    End With
Exit Sub
hell:
End Sub

Private Sub cmdPesanDarah_Click()
    If dgDaftarPasienRJ.ApproxCount = 0 Then Exit Sub
    With frmPemesananDarah
        .txtNoCM = dgDaftarPasienRJ.Columns(1).Value
        .txtNoPendaftaran = dgDaftarPasienRJ.Columns(0).Value
        .txtNamaPasien = dgDaftarPasienRJ.Columns("Nama Pasien").Value

        If dgDaftarPasienRJ.Columns("JK").Value = "P" Then
            txtJK.Text = "Perempuan"
        Else
            txtJK.Text = "Laki-Laki"
        End If
        .txtThn.Text = dgDaftarPasienRJ.Columns("UmurTahun").Value
        .txtBln.Text = dgDaftarPasienRJ.Columns("UmurBulan").Value
        .txtHr.Text = dgDaftarPasienRJ.Columns("UmurHari").Value
        .txtSubInstalasi.Text = dgDaftarPasienRJ.Columns("Poliklinik").Value
        mstrKdSubInstalasi = dgDaftarPasienRJ.Columns("KdSubInstalasi").Value
        .Show
    End With
End Sub

Private Sub cmdPRMRJ_Click()
    If dgDaftarPasienRJ.ApproxCount < 1 Then Exit Sub
    Dim tga As String
    
    Set rs = Nothing
    strSQL = "SELECT dbo.CekPRMRJ('" & dgDaftarPasienRJ.Columns(0).Value & "')"
    Call msubRecFO(rs, strSQL)
    tga = rs(0).Value
    
    strSQL = "SELECT * FROM V_PRMRJ WHERE NoCM='" & dgDaftarPasienRJ.Columns(1).Value & "' AND TglMasuk BETWEEN '" & Format(tga, "yyyy-mm-dd 00:00:00") & "' AND '" & Format(dgDaftarPasienRJ.Columns(9).Value, "yyyy-mm-dd 23:59:59") & "' ORDER BY TglMasuk"
    
    Call frmCetakPRMRJ.Show
End Sub

Private Sub cmdSimpanDokter_Click()
    If mstrKdDokter = "" Then
        MsgBox "Pilih dulu dokternya", vbCritical, "Validasi"
        txtDokter.SetFocus
        Exit Sub
    End If
    Call cmdBatalDokter_Click
    If sp_UbahDokter() = False Then Exit Sub
    Call cmdCari_Click
End Sub

Private Sub cmdTP_Click()
    On Error GoTo hell

    Call subLoadFormTP

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdtutup_Click()
    Set rs = Nothing
    Unload Me
End Sub

Private Sub dcJenisPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcJenisPasien.MatchedWithList = True Then dcKelas.SetFocus
        strSQL = "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien where StatusEnabled='1' and (JenisPasien LIKE '%" & dcJenisPasien.Text & "%')order by JenisPasien"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcJenisPasien.Text = ""
            dcKelas.SetFocus
            Exit Sub
        End If
        dcJenisPasien.BoundText = rs(0).Value
        dcJenisPasien.Text = rs(1).Value
        Call cmdCari_Click
    End If
End Sub

Private Sub dcJenisPasien_LostFocus()
    If dcJenisPasien.MatchedWithList = True Then dcKelas.SetFocus
    strSQL = "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien where StatusEnabled='1' and (JenisPasien LIKE '%" & dcJenisPasien.Text & "%')order by JenisPasien"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
        dcJenisPasien.Text = ""
        dcKelas.SetFocus
        Exit Sub
    End If
    dcJenisPasien.BoundText = rs(0).Value
    dcJenisPasien.Text = rs(1).Value
    Call cmdCari_Click
End Sub

Private Sub dcKelas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcKelas.MatchedWithList = True Then cmdCari.SetFocus
        strSQL = "SELECT KdKelas, DeskKelas FROM KelasPelayanan where StatusEnabled='1' and (DeskKelas LIKE '%" & dcKelas.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcKelas.Text = ""
            cmdCari.SetFocus
            Exit Sub
        End If
        dcKelas.BoundText = rs(0).Value
        dcKelas.Text = rs(1).Value
        Call cmdCari_Click
    End If
End Sub

Private Sub dcKelas_LostFocus()
    If dcKelas.MatchedWithList = True Then cmdCari.SetFocus
    strSQL = "SELECT KdKelas, DeskKelas FROM KelasPelayanan where StatusEnabled='1' and (DeskKelas LIKE '%" & dcKelas.Text & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
        dcKelas.Text = ""
        cmdCari.SetFocus
        Exit Sub
    End If
    dcKelas.BoundText = rs(0).Value
    dcKelas.Text = rs(1).Value
    Call cmdCari_Click
End Sub

Private Sub dcStatusPeriksa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcStatusPeriksa.MatchedWithList = True Then cmdCari.SetFocus
        strSQL = "SELECT KdStatusPeriksa, StatusPeriksa FROM StatusPeriksaPasien WHERE StatusEnabled='1' and (StatusPeriksa LIKE '%" & dcStatusPeriksa.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcStatusPeriksa.Text = ""
            cmdCari.SetFocus
            Exit Sub
        End If
        dcStatusPeriksa.BoundText = rs(0).Value
        dcStatusPeriksa.Text = rs(1).Value
    End If
End Sub

Private Sub dcStatusPeriksa_LostFocus()
    If dcStatusPeriksa.MatchedWithList = True Then cmdCari.SetFocus
    strSQL = "SELECT KdStatusPeriksa, StatusPeriksa FROM StatusPeriksaPasien WHERE StatusEnabled='1' and (StatusPeriksa LIKE '%" & dcStatusPeriksa.Text & "%')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
        dcStatusPeriksa.Text = ""
        cmdCari.SetFocus
        Exit Sub
    End If
    dcStatusPeriksa.BoundText = rs(0).Value
    dcStatusPeriksa.Text = rs(1).Value
End Sub

Private Sub dgDaftarPasienRJ_Click()
    Set MyProperty = dgDaftarPasienRJ
End Sub

Private Sub dgDaftarPasienRJ_HeadClick(ByVal ColIndex As Integer)
    Select Case ColIndex
        Case 0
            mstrFilter = " Order By NoPendaftaran"
        Case 1
            mstrFilter = " Order By NoCM"
        Case 2
            mstrFilter = " Order By [Nama Pasien]"
        Case 3
            mstrFilter = " Order By JK"
        Case 4
            mstrFilter = " Order By Umur"
        Case Else
            mstrFilter = ""
    End Select
    Call cmdCari_Click
End Sub

Private Sub dgDaftarPasienRJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdTP.SetFocus
End Sub

Private Sub dgDaftarPasienRJ_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    lblJumData.Caption = "Data " & dgDaftarPasienRJ.Bookmark & "/" & dgDaftarPasienRJ.ApproxCount
End Sub

Private Sub dgDokter_Click()
    Set MyProperty = dgDokter
End Sub

Private Sub dgDokter_DblClick()
    Call dgDokter_KeyPress(13)
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlDokter = 0 Then Exit Sub
        txtDokter.Text = dgDokter.Columns(1).Value
        mstrKdDokter = dgDokter.Columns(0).Value
        If mstrKdDokter = "" Then
            MsgBox "Pilih dulu Dokter yang akan menangani Pasien", vbCritical, "Validasi"
            txtDokter.Text = ""
            dgDokter.SetFocus
            Exit Sub
        End If
        cmdSimpanDokter.SetFocus
    End If
End Sub

Private Sub dgDokter_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If dgDokter.Visible = False Then Exit Sub
        txtDokter.SetFocus
    End If
End Sub

Private Sub dtpAkhir_Change()
    On Error Resume Next
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpAwal_Change()
    On Error Resume Next
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Activate()
    Call centerForm(Me, MDIUtama)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo hell
    Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)
    Select Case KeyCode
        Case vbKeyF1
            If dgDaftarPasienRJ.ApproxCount = 0 Then Exit Sub
            mstrNoPen = dgDaftarPasienRJ.Columns("No. Registrasi").Value
            If dgDaftarPasienRJ.Columns("JenisPasien") <> "UMUM" Then Call PostingHutangPenjaminPasien_AU("A")
            frm_cetak_RincianBiaya.Show
            
        Case vbKeyF10
            If (dgDaftarPasienRJ.ApproxCount = 0) Or (optPasienPoliklinik.Value = False) Then Exit Sub
'            If dgDaftarPasienRJ.ApproxCount = 0 Then Exit Sub
            mstrNoPen = dgDaftarPasienRJ.Columns("No. Registrasi").Value
            If dgDaftarPasienRJ.Columns("JenisPasien") <> "UMUM" Then Call PostingHutangPenjaminPasien_AU("A")
            frm_cetak_RincianBiayaPenjamin.Show

        Case vbKeyF2
            If strCtrlKey = 4 Then
                If optRujukan.Value = True Then MsgBox "Ubah Data Dokter untuk Pasien Poli", vbInformation, "Medifirts2000 - Informasi": Exit Sub
                If dgDaftarPasienRJ.ApproxCount = 0 Then Exit Sub
                With fraDokterP
                    .Left = (Me.Width - .Width) / 2
                    .Top = fraDaftar.Top
                    .Visible = True
                    .Enabled = True
                End With

                fraPilih.Enabled = False
                fraDaftar.Enabled = False
                fraCari.Enabled = False
                txtNoPendaftaran.Text = dgDaftarPasienRJ.Columns(0).Value
                txtNoCM.Text = dgDaftarPasienRJ.Columns(1).Value
                txtNamaPasien.Text = dgDaftarPasienRJ.Columns("Nama Pasien").Value
                If dgDaftarPasienRJ.Columns("JK").Value = "P" Then
                    txtJK.Text = "Perempuan"
                Else
                    txtJK.Text = "Laki-Laki"
                End If
                txtPoli.Text = dgDaftarPasienRJ.Columns("Ruangan").Value
                txtThn.Text = dgDaftarPasienRJ.Columns("UmurTahun").Value
                txtBln.Text = dgDaftarPasienRJ.Columns("UmurBulan").Value
                txtHr.Text = dgDaftarPasienRJ.Columns("UmurHari").Value
                dTglMasuk = dgDaftarPasienRJ.Columns("TglMasuk").Value
                txtTglPeriksa.Text = Format(dTglMasuk, "dd MMMM yyyy HH:mm:ss")
                mstrKdSubInstalasi = dgDaftarPasienRJ.Columns("KdSubInstalasi").Value
                txtDokter.Text = ""
                txtDokter.SetFocus
                txtPrevDokter.Text = dgDaftarPasienRJ.Columns("Dokter Pemeriksa").Value
                mstrKdSubInstalasi = dgDaftarPasienRJ.Columns("KdSubInstalasi").Value
            End If

        Case vbKeyZ
            If strCtrlKey = 4 Then
                If dgDaftarPasienRJ.ApproxCount = 0 Then Exit Sub

                mstrNoPen = dgDaftarPasienRJ.Columns("No. Registrasi")
                mdTglMasuk = Format(Now, "yyyy/MM/dd")
                With frmSuratKeterangan
                    
                If optPasienPoliklinik.Value = True Then
                    .dcDokterPenguji.Text = dgDaftarPasienRJ.Columns("Dokter Pemeriksa").Value
                Else
                   .dcDokterPenguji.Text = dgDaftarPasienRJ.Columns("Dokter Perujuk").Value
                End If
               strSQL = "Select isnull(NIP,'-') as NIP from V_M_DataPegawaiNew where [Nama Lengkap] like '%" & .dcDokterPenguji.Text & "%'"
                Call msubRecFO(rs, strSQL)
                If rs.EOF = True Then
                    .txtNIP.Text = "-"
                  Else
'                rs.Open "Select NIP from V_M_DataPegawaiNew where [Nama Lengkap] like '%" & .dcDokterPenguji.Text & "%'", dbConn, , adLockOptimistic
                    .txtNIP.Text = rs("NIP")
                End If
                .Show
                End With
            End If

        Case vbKeyQ
'            If strCtrlKey = 4 Then
'                If dgDaftarPasienRJ.ApproxCount = 0 Then Exit Sub
'
'                mstrNoPen = dgDaftarPasienRJ.Columns("No. Registrasi")
'                mdTglMasuk = Format(Now, "yyyy/MM/dd")
'                frmSuratKeteranganKeluhan.Show
'            End If

        Case vbKeyF3
            If strCtrlKey = 4 Then
                If dgDaftarPasienRJ.ApproxCount = 0 Then Exit Sub
                strPasien = "Lama"
                mstrKdSubInstalasi = dgDaftarPasienRJ.Columns("KdSubInstalasi").Value
                mstrNoCM = dgDaftarPasienRJ.Columns(1).Value
                With frmPasienBaru
                    .Show
                    .txtTahun.Text = dgDaftarPasienRJ.Columns("UmurTahun").Value
                    .txtBulan.Text = dgDaftarPasienRJ.Columns("UmurBulan").Value
                    .txtHari.Text = dgDaftarPasienRJ.Columns("UmurHari").Value
                End With
            End If
        Case vbKeyF9
            frmCtkDaftarPasien.Show
        Case vbKeyF5
            Call cmdCari_Click

        Case vbKeyF11
            If (dgDaftarPasienRJ.ApproxCount = 0) Or (optRujukan.Value = False) Then Exit Sub

            Set rs = Nothing
            strSQL = "Select StatusPeriksa from PasienRujukan Where NoPendaftaran='" & dgDaftarPasienRJ.Columns(0).Value & "' AND TglDirujuk='" & Format(dgDaftarPasienRJ.Columns(9).Value, "yyyy/MM/dd HH:mm:ss") & "'"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockOptimistic
            If rs(0).Value = "Y" Then
                MsgBox "Pasien Sudah Diperiksa", vbInformation, "Validasi"
                Exit Sub
            End If
            If MsgBox("Apakah pasien sudah selesai diperiksa", vbQuestion + vbYesNo, "Medifirst2000 - Konfirmasi") = vbNo Then Exit Sub
            Set rs = Nothing
            strQuery = "Update PasienRujukan set StatusPeriksa='Y' where NoPendaftaran='" & dgDaftarPasienRJ.Columns(0).Value & "' AND TglDirujuk='" & Format(dgDaftarPasienRJ.Columns(9).Value, "yyyy/MM/dd HH:mm:ss") & "'"
            rs.Open strQuery, dbConn, adOpenDynamic, adLockOptimistic
            Set rs = Nothing
            Call cmdCari_Click
    End Select
    Exit Sub
hell:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo hell

    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)

    mblnFormDaftarPasienRJ = True
    dtpAwal.Value = Format(Now, "dd MMM yyyy 00:00:00")
    dtpAkhir.Value = Now

    Call subLoadDcSource

    optPasienPoliklinik.Caption = "Pasien " + strNNamaRuangan
    optPasienPoliklinik.Value = True

    mblnForm = True

    strSQL = "Select KdRuangan from Ruangan WHERE NamaRuangan LIKE '%Gigi%'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        For i = 1 To rs.RecordCount
            If mstrKdRuangan = rs(0).Value Then cmdOdonto.Visible = True Else cmdOdonto.Visible = False
            If mstrKdRuangan = rs(0).Value Then GoTo Out_
            rs.MoveNext
        Next i
Out_:
    End If

    strSQL = "Select KdRuangan from Ruangan WHERE NamaRuangan LIKE '%Mata%'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        For i = 1 To rs.RecordCount
            If mstrKdRuangan = rs(0).Value Then cmdAnamnesa.Visible = True Else cmdAnamnesa.Visible = False
            If mstrKdRuangan = rs(0).Value Then GoTo OutJuga_
            rs.MoveNext
        Next i
OutJuga_:
    End If
    
    strSQL = "Select * from Ruangan WHERE NamaRuangan LIKE '%Rehab%' or NamaRuangan LIKE '%Umum%'or NamaRuangan LIKE '%Check%'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        For i = 1 To rs.RecordCount
            If mstrKdRuangan = rs(0).Value Then cmdNapza.Enabled = True Else cmdNapza.Enabled = False
            If mstrKdRuangan = rs(0).Value Then GoTo OutNapsa_
            rs.MoveNext
        Next i
OutNapsa_:
    End If
    Call cmdCari_Click
    Exit Sub
hell:
    Call msubPesanError
    Set rs = Nothing
End Sub

Sub SetGridPasienRJ()
    With dgDaftarPasienRJ
        .Columns(0).Width = 1250
        .Columns(0).Caption = "No. Registrasi"
        .Columns(1).Width = 1350
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 2000
        .Columns(3).Width = 300
        .Columns(4).Width = 1000
        .Columns(5).Width = 0
        .Columns(6).Width = 0
        .Columns(7).Width = 1600
        .Columns(8).Width = 1600
        .Columns(9).Width = 1900
        .Columns(10).Width = 2600
        .Columns(11).Width = 800
        .Columns(12).Width = 0
        .Columns(13).Width = 0
        .Columns(14).Width = 0
        .Columns(15).Width = 0
        .Columns(16).Width = 0
        .Columns(17).Width = 0
        .Columns(18).Width = 0
        .Columns(19).Width = 5500
        .Columns(20).Width = 0
        .Columns(21).Width = 2000
        .Columns(22).Width = 2000
        .Columns(23).Width = 0
        .Columns(24).Width = 1500
    End With
End Sub

Sub SetGridPasienKonsul()
    With dgDaftarPasienRJ
        .Columns(0).Width = 1150
        .Columns(0).Caption = "No. Registrasi"
        .Columns(1).Width = 750
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 1800
        .Columns(3).Width = 300
        .Columns(4).Width = 1400
        .Columns(5).Width = 2400
        .Columns(6).Width = 1700
        .Columns(7).Width = 0
        .Columns(8).Width = 2400
        .Columns(9).Width = 1580
        .Columns(10).Width = 0
        .Columns(11).Width = 0
        .Columns(12).Width = 0
        .Columns(13).Width = 0
        .Columns(14).Width = 0
        .Columns(15).Width = 0
        .Columns(16).Width = 0
        .Columns(17).Width = 0
        .Columns(18).Width = 0
        .Columns(19).Width = 0
        .Columns(20).Width = 5500
        .Columns(21).Width = 0
        .Columns(22).Width = 2000
        .Columns(23).Width = 2000
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnFormDaftarPasienRJ = False
    mblnForm = False
End Sub

Private Sub mnuIRBS_Click()
    On Error GoTo hell
    If dgDaftarPasienRJ.ApproxCount = 0 Then Exit Sub
    mstrNoPen = dgDaftarPasienRJ.Columns(0).Value
    frm_cetak_RincianBiaya.Show
    Exit Sub
hell:
End Sub

Private Sub mnuSDokter_Click()
    On Error GoTo errLoad

    If dgDaftarPasienRJ.ApproxCount = 0 Then Exit Sub
    With fraDokterP
        .Left = (Me.Width - .Width) / 2
        .Top = fraDaftar.Top
        .Visible = True
        .Enabled = True
    End With

    fraPilih.Enabled = False
    fraDaftar.Enabled = False
    fraCari.Enabled = False
    txtNoPendaftaran.Text = dgDaftarPasienRJ.Columns(0).Value
    txtNoCM.Text = dgDaftarPasienRJ.Columns(1).Value
    txtNamaPasien.Text = dgDaftarPasienRJ.Columns("Nama Pasien").Value
    If dgDaftarPasienRJ.Columns("JK").Value = "P" Then
        txtJK.Text = "Perempuan"
    Else
        txtJK.Text = "Laki-Laki"
    End If
    txtPoli.Text = dgDaftarPasienRJ.Columns("Ruangan").Value
    txtThn.Text = dgDaftarPasienRJ.Columns("UmurTahun").Value
    txtBln.Text = dgDaftarPasienRJ.Columns("UmurBulan").Value
    txtHr.Text = dgDaftarPasienRJ.Columns("UmurHari").Value
    dTglMasuk = dgDaftarPasienRJ.Columns("TglMasuk").Value
    txtTglPeriksa.Text = Format(dTglMasuk, "dd MMMM yyyy HH:mm:ss")
    mstrKdSubInstalasi = dgDaftarPasienRJ.Columns("KdSubInstalasi").Value
    txtDokter.Text = ""
    txtDokter.SetFocus
    txtPrevDokter.Text = dgDaftarPasienRJ.Columns("Dokter Pemeriksa").Value
    mstrKdSubInstalasi = dgDaftarPasienRJ.Columns("KdSubInstalasi").Value

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub optPasienPoliklinik_Click()
    On Error GoTo errLoad
    StatusBar1.Panels(5).Visible = False
    StatusBar1.Panels(7).Visible = False
    StatusBar1.Panels(2).Visible = True
    Call AturStatusBar
    lblStatusPeriksa.Visible = False
    dcStatusPeriksa.Visible = False
    dcStatusPeriksa.Text = ""
    dcJenisPasien.Visible = True
    dcKelas.Visible = True
    Call cmdCari_Click
    cmdBatalPeriksa.Visible = True
    optPasienPoliklinik.SetFocus
    cmdPasienPulangRJ.Visible = True
    Exit Sub
errLoad:
End Sub

Private Sub optPasienPoliklinik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcJenisPasien.SetFocus
End Sub
Private Sub AturStatusBar()
    If optPasienPoliklinik.Value = True Then
        StatusBar1.Panels(1).Width = "2600,5357"
        StatusBar1.Panels(2).Width = "3100,5357"
        StatusBar1.Panels(3).Width = "4100,5357"
'        StatusBar1.Panels(5).Width = "3100,5357"
        StatusBar1.Panels(6).Width = "4100,5357"
        StatusBar1.Panels(7).Width = "3100,5357"
    Else
'        StatusBar1.Panels(2).Width = "5500,2208"
        StatusBar1.Panels(1).Width = "2600,5357"
        StatusBar1.Panels(3).Width = "4100,5357"
        StatusBar1.Panels(5).Width = "3100,5357"
        StatusBar1.Panels(6).Width = "4100,5357"
        StatusBar1.Panels(7).Width = "3100,5357"
    
    End If
End Sub

Private Sub optPRMRJ_Click()
    StatusBar1.Panels(7).Visible = True
    StatusBar1.Panels(2).Visible = False
    Call AturStatusBar
    lblStatusPeriksa.Visible = True
    dcStatusPeriksa.Visible = True
    dcJenisPasien.Visible = False
    dcJenisPasien.Text = ""
    dcKelas.Visible = False
    dcKelas.Text = ""
    Call subLoadDcSource
    Call cmdCari_Click
    cmdBatalPeriksa.Visible = False
    optRujukan.SetFocus
    cmdPasienPulangRJ.Visible = False
End Sub

Private Sub optRujukan_Click()
'    StatusBar1.Panels(5).Visible = True
    StatusBar1.Panels(7).Visible = True
    StatusBar1.Panels(2).Visible = False
    Call AturStatusBar
    lblStatusPeriksa.Visible = True
    dcStatusPeriksa.Visible = True
    dcJenisPasien.Visible = False
    dcJenisPasien.Text = ""
    dcKelas.Visible = False
    dcKelas.Text = ""
    Call subLoadDcSource
    Call cmdCari_Click
    cmdBatalPeriksa.Visible = False
    optRujukan.SetFocus
    cmdPasienPulangRJ.Visible = False
End Sub

Private Sub optRujukan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcStatusPeriksa.SetFocus
End Sub

Private Sub txtBln_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtHr.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtDokter_Change()
    strFilterDokter = "WHERE NamaDokter like '%" & txtDokter.Text & "%'"
    mstrKdDokter = ""
    Call subLoadDokter
End Sub

Private Sub txtDokter_GotFocus()
    If txtDokter.Text = "" Then strFilterDokter = ""
    Call subLoadDokter
End Sub

Private Sub txtDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlDokter = 0 Then Exit Sub
        dgDokter.SetFocus
    End If
    If KeyAscii = 39 Then KeyAscii = 0
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub subLoadFormTP()
    On Error GoTo hell

    If dgDaftarPasienRJ.ApproxCount = 0 Then Exit Sub
    

    mstrNoPen = dgDaftarPasienRJ.Columns(0).Value
    mstrNoCM = dgDaftarPasienRJ.Columns(1).Value
    If optPasienPoliklinik.Value = True Then
        strSQL = "SELECT IdPegawai FROM DataPegawai WHERE NamaLengkap = '" & dgDaftarPasienRJ.Columns("Dokter Pemeriksa") & "'"
    Else
        strSQL = "SELECT IdPegawai FROM DataPegawai WHERE NamaLengkap = '" & dgDaftarPasienRJ.Columns("Dokter Perujuk") & "'"
    End If
    msubRecFO rs, strSQL
    mstrKdDokter = rs(0).Value
    If optPasienPoliklinik.Value = True Then
    mstrNamaDokter = dgDaftarPasienRJ.Columns("Dokter Pemeriksa")
    Else
    mstrNamaDokter = dgDaftarPasienRJ.Columns("Dokter PErujuk")
    End If
    With frmTransaksiPasien
        .Show
        Me.Enabled = False

        .txtNoPendaftaran.Text = dgDaftarPasienRJ.Columns(0).Value
        .txtNoCM.Text = mstrNoCM
        .txtNamaPasien.Text = dgDaftarPasienRJ.Columns(2).Value
        If dgDaftarPasienRJ.Columns(3).Value = "L" Then
            .txtSex.Text = "Laki-Laki"
        Else
            .txtSex.Text = "Perempuan"
        End If
        .txtThn.Text = dgDaftarPasienRJ.Columns("UmurTahun").Value
        .txtBln.Text = dgDaftarPasienRJ.Columns("UmurBulan").Value
        .txtHr.Text = dgDaftarPasienRJ.Columns("UmurHari").Value
        .txtKls.Text = dgDaftarPasienRJ.Columns("Kelas").Value
        If optPasienPoliklinik.Value = True Then
            .txtJenisPasien.Text = dgDaftarPasienRJ.Columns("JenisPasien").Value
        Else
            .txtJenisPasien.Text = dgDaftarPasienRJ.Columns(10).Value
        End If
        If optRujukan.Value = True Then
            .txtTglDaftar.Text = dgDaftarPasienRJ.Columns(16).Value
            mdTglMasuk = dgDaftarPasienRJ.Columns(16).Value
            mstrKdKelas = dgDaftarPasienRJ.Columns(17).Value
            mstrKelas = dgDaftarPasienRJ.Columns(18).Value
            mstrKdSubInstalasi = dgDaftarPasienRJ.Columns(21).Value

        Else
            .txtTglDaftar.Text = dgDaftarPasienRJ.Columns(9).Value
            mdTglMasuk = dgDaftarPasienRJ.Columns(9).Value
            mstrKdKelas = dgDaftarPasienRJ.Columns(16).Value
            mstrKelas = dgDaftarPasienRJ.Columns(8).Value
            mstrKdSubInstalasi = dgDaftarPasienRJ.Columns(17).Value
        End If

        strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
            mstrKdJenisPasien = rs("KdKelompokPasien").Value
            mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
        End If

    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

'untuk meload data dokter di grid
Private Sub subLoadDokter()
    On Error GoTo hell
    strSQL = "SELECT KodeDokter AS [Kode Dokter],NamaDokter AS [Nama Dokter],JK,Jabatan FROM V_DaftarDokter " & strFilterDokter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlDokter = rs.RecordCount
    Set dgDokter.DataSource = rs
    With dgDokter
        .Columns(0).Width = 1200
        .Columns(1).Width = 4000
        .Columns(2).Width = 400
        .Columns(3).Width = 3000
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

'Store procedure untuk mengisi registrasi pasien
Private Function sp_UbahDokter() As Boolean
    On Error GoTo hell
    sp_UbahDokter = True
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, mstrKdDokter)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(dTglMasuk, "yyyy/MM/dd HH:mm:ss"))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_DokterPemeriksaRJ"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam proses penyimpanan data", vbCritical, "Validasi"
            sp_UbahDokter = False
        Else
            MsgBox "Ubah dokter pemeriksa selesai", vbInformation, "Informasi"
            Call Add_HistoryLoginActivity("Update_DokterPemeriksaRJ")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
hell:
    Call msubPesanError("sp_UbahDokter")
    sp_UbahDokter = False
    Set dbcmd = Nothing
End Function

Private Sub txtHr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPoli.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtJK_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtThn.SetFocus
End Sub

Private Sub txtNamaPasien_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtJK.SetFocus
End Sub

Private Sub txtNoCM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaPasien.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtNoPendaftaran_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtNoCM.SetFocus
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        Call cmdCari_Click
        txtParameter.SetFocus
    End If
End Sub

Private Sub subLoadDcSource()
    On Error GoTo errLoad

    Call msubDcSource(dcStatusPeriksa, rs, "SELECT KdStatusPeriksa, StatusPeriksa FROM StatusPeriksaPasien WHERE StatusEnabled='1'")
    If rs.EOF = False Then dcStatusPeriksa.BoundText = rs(0).Value
    Call msubDcSource(dcJenisPasien, rs, "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien where StatusEnabled='1' order by JenisPasien")
    Call msubDcSource(dcKelas, rs, "SELECT KdKelas, DeskKelas FROM KelasPelayanan where StatusEnabled='1'")

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtPoli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTglPeriksa.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtPrevDokter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If dgDokter.Visible = False Then Exit Sub
        dgDokter.SetFocus
    End If
End Sub

Private Sub txtPrevDokter_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtTglPeriksa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPrevDokter.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtThn_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtBln.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

