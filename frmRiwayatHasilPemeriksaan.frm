VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmRiwayatHasilPemeriksaan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Riwayat Hasil Pemeriksaan"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRiwayatHasilPemeriksaan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   15030
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   7320
      Width           =   15015
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   495
         Left            =   11400
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   13200
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6375
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   15015
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
         Left            =   9120
         TabIndex        =   8
         Top             =   240
         Width           =   5775
         Begin VB.CommandButton cmdTampilkanTemp 
            Caption         =   "&Cari"
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
            Format          =   60817411
            UpDown          =   -1  'True
            CurrentDate     =   38373
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
            Format          =   60817411
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3120
            TabIndex        =   9
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgData 
         Height          =   5175
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   9128
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   10
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
      Left            =   13200
      Picture         =   "frmRiwayatHasilPemeriksaan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRiwayatHasilPemeriksaan.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRiwayatHasilPemeriksaan.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13215
   End
End
Attribute VB_Name = "frmRiwayatHasilPemeriksaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCetak_Click()
On Error GoTo errload
        
    If dgData.ApproxCount = 0 Then Exit Sub
    
    cmdCetak.Enabled = False
    strSQL = "SELECT * FROM V_RiwayatHasilPemeriksaan" & _
        " WHERE TglPendaftaran BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "'"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount = 0 Then
        cmdCetak.Enabled = True
        Exit Sub
    End If

    mstrNoLabRad = dgData.Columns(0).Value

    Select Case dgData.Columns("KdInstalasi").Value
        Case "09" 'lab pk
            Set frmcetakhasillab = Nothing
            frmcetakhasillab.Show
            
        Case "16" 'lab pa
            Set frmCetakHasilLabPA = Nothing
            frmCetakHasilLabPA.Show
            
        Case "10" 'radiologi
            Set frmCetakHasilRadiologi = Nothing
            frmCetakHasilRadiologi.Show
    End Select

    cmdCetak.Enabled = True

Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdTampilkanTemp_Click()
On Error GoTo errTampilkan
    
    strSQL = "SELECT * FROM V_RiwayatHasilPemeriksaan" & _
        " WHERE TglPendaftaran BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "'"
    Call msubRecFO(rs, strSQL)
    Set dgData.DataSource = rs
    'If dgData.ApproxCount = 0 Then dtpAwal.SetFocus Else dgData.SetFocus
    
Exit Sub
errTampilkan:
    msubPesanError
End Sub

Private Sub dgData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetak.SetFocus
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdTampilkanTemp.SetFocus
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo errFormLoad
    
    Call centerForm(Me, MDIUtama)
    
    dtpAwal.Value = Now
    dtpAkhir.Value = Now
    Call cmdTampilkanTemp_Click
    Call PlayFlashMovie(Me)
Exit Sub
errFormLoad:
    msubPesanError
End Sub




