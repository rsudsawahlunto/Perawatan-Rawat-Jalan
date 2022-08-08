VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLaporanSaldoBarang 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Laporan Saldo Barang"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLaporanSaldoBarang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
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
      Height          =   1215
      Left            =   0
      TabIndex        =   21
      Top             =   960
      Width           =   9255
      Begin VB.Frame Frame3 
         Caption         =   "Group By"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   4335
         Begin VB.OptionButton optHari 
            Caption         =   "Per Hari"
            Height          =   375
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optBulan 
            Caption         =   "Per Bulan"
            Height          =   375
            Left            =   1200
            TabIndex        =   1
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optTotal 
            Caption         =   "Total"
            Height          =   375
            Left            =   3480
            TabIndex        =   3
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optTahun 
            Caption         =   "Per Tahun"
            Height          =   375
            Left            =   2280
            TabIndex        =   2
            Top             =   240
            Width           =   1215
         End
      End
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   330
         Left            =   4560
         TabIndex        =   4
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
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
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   60358659
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   330
         Left            =   7080
         TabIndex        =   5
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
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
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   60358659
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "s/d"
         Height          =   210
         Left            =   6720
         TabIndex        =   23
         Top             =   525
         Width           =   255
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   20
      Top             =   4200
      Width           =   9255
      Begin VB.CommandButton cmdSpreadSheet 
         Caption         =   "&SpreadSheet"
         Height          =   495
         Left            =   4920
         TabIndex        =   15
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "&Tutup"
         Height          =   495
         Left            =   7080
         TabIndex        =   16
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Group By Laporan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   17
      Top             =   2160
      Width           =   9255
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   8775
         Begin VB.OptionButton optGolonganBarang 
            Caption         =   "Golongan Barang"
            Height          =   375
            Left            =   6720
            TabIndex        =   10
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton optStatusBarang 
            Caption         =   "Status Barang"
            Height          =   375
            Left            =   4920
            TabIndex        =   9
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optPabrik 
            Caption         =   "Pabrik"
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optAsalBarang 
            Caption         =   "Asal Barang"
            Height          =   375
            Left            =   1560
            TabIndex        =   7
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optJenisBarang 
            Caption         =   "Jenis Barang"
            Height          =   375
            Left            =   3240
            TabIndex        =   8
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.CheckBox chkNamaBarang 
         Caption         =   "Nama Barang"
         Height          =   255
         Left            =   4560
         TabIndex        =   13
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CheckBox chkGroup 
         Caption         =   "Group yang dipilih"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   2415
      End
      Begin MSDataListLib.DataCombo dcGroup 
         Height          =   390
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   688
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcNamaBarang 
         Height          =   390
         Left            =   4560
         TabIndex        =   14
         Top             =   1440
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   688
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   24
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
      Left            =   7440
      Picture         =   "frmLaporanSaldoBarang.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmLaporanSaldoBarang.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   210
      Left            =   4200
      TabIndex        =   18
      Top             =   2880
      Width           =   60
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmLaporanSaldoBarang.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmLaporanSaldoBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkGroup_Click()
    If chkGroup.Value = vbChecked Then
        dcGroup.Enabled = True
    Else
        dcGroup.Enabled = False
    End If
End Sub

Private Sub chkGroup_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkGroup.Value = vbChecked Then
            dcGroup.SetFocus
        Else
            chkNamaBarang.SetFocus
        End If
    End If
End Sub

Private Sub chkNamaBarang_Click()
    If chkNamaBarang.Value = vbChecked Then
        dcNamaBarang.Enabled = True
        Call msubDcSource(dcNamaBarang, rs, "SELECT KdBarang, NamaBarang FROM MasterBarang ORDER BY NamaBarang")
    Else
        dcNamaBarang.Enabled = False
    End If
End Sub

Private Sub chkNamaBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkNamaBarang.Value = vbChecked Then
            dcNamaBarang.SetFocus
        Else
            cmdSpreadSheet.SetFocus
        End If
    End If
End Sub

Private Sub cmdSpreadSheet_Click()
    strGroup = "": strNama = "": strNoFaktur = ""

    If optHari.Value = True Then
        strCetak = "Hari"
    ElseIf optBulan.Value = True Then
        strCetak = "Bulan"
    ElseIf optTahun.Value = True Then
        strCetak = "Tahun"
    ElseIf optTotal.Value = True Then
        strCetak = "Total"
    Else
        strCetak = ""
    End If

    mdTglAwal = dtpAwal.Value 'TglAwal
    mdTglAkhir = dtpAkhir.Value 'TglAkhir

    If optPabrik.Value = True Then
        strGroup = "Pabrik"
    ElseIf optAsalBarang.Value = True Then
        strGroup = "AsalBarang"
    ElseIf optJenisBarang.Value = True Then
        strGroup = "DetailJenisBarang"
    ElseIf optStatusBarang.Value = True Then
        strGroup = "StatusBarang"
    ElseIf optGolonganBarang.Value = True Then
        strGroup = "GolonganBarang"
    Else
        strGroup = vbNullString
    End If

    If chkGroup.Value = vbChecked Then
        strIsiGroup = dcGroup.Text
    Else
        strIsiGroup = vbNullString
    End If

    If chkNamaBarang.Value = vbChecked Then
        strNama = dcNamaBarang.Text
    Else
        strNama = vbNullString
    End If

    frmCetakLapSaldoBarang.Show
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcGroup_GotFocus()
    On Error GoTo errLoad
    Dim tempKode As String

    tempKode = dcGroup.BoundText
    If optPabrik.Value = True Then
        Call msubDcSource(dcGroup, rs, "SELECT KdPabrik, NamaPabrik FROM Pabrik where StatusEnabled='1' ORDER BY NamaPabrik")
    ElseIf optAsalBarang.Value = True Then
        Call msubDcSource(dcGroup, rs, "SELECT KdAsal, NamaAsal FROM AsalBarang where StatusEnabled='1' ORDER BY NamaAsal")
    ElseIf optJenisBarang.Value = True Then
        Call msubDcSource(dcGroup, rs, "SELECT KdJenisBarang, JenisBarang FROM JenisBarang where StatusEnabled='1' ORDER BY JenisBarang")
    ElseIf optStatusBarang.Value = True Then
        Call msubDcSource(dcGroup, rs, "SELECT KdStatusBarang, StatusBarang FROM StatusBarang where StatusEnabled='1' ORDER BY StatusBarang")
    ElseIf optGolonganBarang.Value = True Then
        Call msubDcSource(dcGroup, rs, "SELECT KdGolBarang, GolonganBarang FROM GolonganBarang where StatusEnabled='1' ORDER BY GolonganBarang")
    Else
        Exit Sub
    End If
    dcGroup.BoundText = tempKode

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcGroup_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(dcGroup.Text)) = 0 Then dcGroup.SetFocus: Exit Sub
        If dcGroup.MatchedWithList = True Then chkNamaBarang.SetFocus: Exit Sub
        If optPabrik.Value = True Then
            strSQL = "SELECT KdPabrik, NamaPabrik FROM Pabrik WHERE (NamaPabrik LIKE '%" & dcGroup.Text & "%')"
        ElseIf optAsalBarang.Value = True Then
            strSQL = "SELECT KdAsal, NamaAsal FROM AsalBarang WHERE (NamaAsal LIKE '%" & dcGroup.Text & "%')"
        ElseIf optJenisBarang.Value = True Then
            strSQL = "SELECT KdJenisBarang, JenisBarang FROM JenisBarang WHERE (JenisBarang LIKE '%" & dcGroup.Text & "%')"
        ElseIf optStatusBarang.Value = True Then
            strSQL = "SELECT KdStatusBarang, StatusBarang FROM StatusBarang WHERE (StatusBarang LIKE '%" & dcGroup.Text & "%')"
        ElseIf optGolonganBarang.Value = True Then
            strSQL = "SELECT KdGolBarang, GolonganBarang FROM GolonganBarang WHERE (GolonganBarang LIKE '%" & dcGroup.Text & "%')"
        Else
            optPabrik.SetFocus
        End If
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcGroup.BoundText = rs(0).Value
        dcGroup.Text = rs(1).Value
    End If
End Sub

Private Sub dcNamaBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(dcNamaBarang.Text)) = 0 Then dcNamaBarang.SetFocus: Exit Sub
        If dcNamaBarang.MatchedWithList = True Then cmdSpreadSheet.SetFocus: Exit Sub
        Call msubRecFO(rs, "SELECT KdBarang, NamaBarang FROM MasterBarang WHERE NamaBarang LIKE '%" & dcNamaBarang.Text & "%'and StatusEnabled='1'")
        If rs.EOF = True Then
            dcNamaBarang.Text = ""
            Exit Sub
        End If
        dcNamaBarang.BoundText = rs(0).Value
        dcNamaBarang.Text = rs(1).Value
    End If
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then optPabrik.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    dtpAwal.Value = Now
    dtpAkhir.Value = Now
End Sub

Private Sub optAsalBarang_Click()
    dcGroup.BoundText = ""
    chkGroup.Caption = optAsalBarang.Caption
End Sub

Private Sub optAsalBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkGroup.SetFocus
End Sub

Private Sub optBulan_Click()
    dtpAwal.CustomFormat = "MMMM yyyy"
    dtpAkhir.CustomFormat = "MMMM yyyy"
End Sub

Private Sub optBulan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAwal.SetFocus
End Sub

Private Sub optGolonganBarang_Click()
    dcGroup.BoundText = ""
    chkGroup.Caption = optGolonganBarang.Caption
End Sub

Private Sub optHari_Click()
    dtpAwal.CustomFormat = "dd MMMM yyyy"
    dtpAkhir.CustomFormat = "dd MMMM yyyy"
End Sub

Private Sub optHari_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAwal.SetFocus
End Sub

Private Sub optGolonganBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkGroup.SetFocus
End Sub

Private Sub optJenisBarang_Click()
    dcGroup.BoundText = ""
    chkGroup.Caption = optJenisBarang.Caption
End Sub

Private Sub optJenisBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkGroup.SetFocus
End Sub

Private Sub optPabrik_Click()
    dcGroup.BoundText = ""
    chkGroup.Caption = optPabrik.Caption
End Sub

Private Sub optPabrik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkGroup.SetFocus
End Sub

Private Sub optStatusBarang_Click()
    dcGroup.BoundText = ""
    chkGroup.Caption = optStatusBarang.Caption
End Sub

Private Sub optStatusBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkGroup.SetFocus
End Sub

Private Sub optTahun_Click()
    dtpAwal.CustomFormat = "yyyy"
    dtpAkhir.CustomFormat = "yyyy"
End Sub

Private Sub optTahun_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAwal.SetFocus
End Sub

Private Sub optTotal_Click()
    Call optHari_Click
End Sub

Private Sub optTotal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAwal.SetFocus
End Sub
