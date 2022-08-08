VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDaftarPasienReturStruk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Pasien Retur Struk"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPasienReturStruk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   13215
   Begin VB.Frame fraCariPasien 
      Caption         =   "Cari Data Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   0
      TabIndex        =   7
      Top             =   7560
      Width           =   13185
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Rincian &Biaya"
         Height          =   450
         Left            =   8760
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   450
         Left            =   11040
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   400
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan Nama Pasien /  No.CM"
         Height          =   210
         Left            =   1560
         TabIndex        =   8
         Top             =   150
         Width           =   2640
      End
   End
   Begin VB.Frame fraDafPasien 
      Caption         =   "Daftar Pasien Retur Struk"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   13215
      Begin VB.Frame Frame3 
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
         Left            =   7320
         TabIndex        =   10
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
            Format          =   55836675
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
            Format          =   55836675
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   11
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgDaftarPasienSudahBayar 
         Height          =   5535
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   9763
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
   End
   Begin VB.Image Image2 
      Height          =   930
      Left            =   3008
      Picture         =   "frmDaftarPasienReturStruk.frx":08CA
      Top             =   0
      Width           =   10200
   End
   Begin VB.Image Image1 
      Height          =   930
      Left            =   -2640
      Picture         =   "frmDaftarPasienReturStruk.frx":6012
      Top             =   0
      Width           =   10200
   End
End
Attribute VB_Name = "frmDaftarPasienReturStruk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dTglMasuk As Date

Private Sub cmdcari_Click()
On Error GoTo errLoad
    Set rs = Nothing
    rs.Open "select * from V_DaftarPasienReturStruk where ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%') and TglStruk between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' AND KdRuangan = '" & mstrKdRuangan & "'", dbConn, adOpenStatic, adLockOptimistic
    Set dgDaftarPasienSudahBayar.DataSource = rs
    Call SetGridPasienReturStruk
    If dgDaftarPasienSudahBayar.ApproxCount > 0 Then
        dgDaftarPasienSudahBayar.SetFocus
    Else
        dtpAwal.SetFocus
    End If
    
errLoad:
End Sub

Private Sub cmdPreview_Click()
On Error GoTo hell
    If dgDaftarPasienSudahBayar.ApproxCount = 0 Then Exit Sub
    cmdPreview.Enabled = False
    If Len(dgDaftarPasienSudahBayar.Columns(0).Value) < 1 Then Exit Sub
    mstrNoStruk = dgDaftarPasienSudahBayar.Columns(0).Value
    vLaporan = "Cetak Ulang"
    frmCetak.Show
hell:
    cmdPreview.Enabled = True
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgDaftarPasienSudahBayar_DblClick()
    Call cmdPreview_Click
End Sub

Private Sub dgDaftarPasienSudahBayar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdPreview.SetFocus
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Activate()
    Call cmdcari_Click
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    dtpAwal.Value = Format(Now, "dd MMM yyyy 00:00:00")
    dtpAkhir.Value = Now
    Call cmdcari_Click
End Sub

Sub SetGridPasienReturStruk()
    With dgDaftarPasienSudahBayar
        .Columns(0).Width = 1150
        .Columns(0).Caption = "No. Struk"
        .Columns(1).Width = 1150
        .Columns(1).Caption = "No. Registrasi"
        .Columns(2).Width = 750 'cm
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Width = 1800
        .Columns(4).Width = 300
        .Columns(5).Width = 1450
        .Columns(6).Width = 1500
        .Columns(7).Width = 1569
        .Columns(8).Width = 2700
        .Columns(9).Width = 0
        .Columns(10).Width = 0
        .Columns(11).Width = 0
        .Columns(12).Width = 0
        .Columns(13).Width = 0
    End With
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        Call cmdcari_Click
        If dgDaftarPasienSudahBayar.ApproxCount = 0 Then txtParameter.SetFocus
    End If
End Sub

Private Sub txtParameter_LostFocus()
    txtParameter.Text = StrConv(txtParameter.Text, vbProperCase)
End Sub
