VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStokOpname 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Stok Opname"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15390
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStokOpname.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   15390
   Begin VB.CheckBox Check1 
      Caption         =   "Set Tanggal Kadar Luarsa"
      Height          =   375
      Left            =   360
      TabIndex        =   27
      Top             =   7440
      Width           =   2535
   End
   Begin MSMask.MaskEdBox mebIsi 
      Height          =   330
      Left            =   4320
      TabIndex        =   22
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   16
      Format          =   "dd-MM-yyyy"
      Mask            =   "##/##/#### ##:##"
      PromptChar      =   "_"
   End
   Begin VB.CheckBox chkSetStokReal 
      Caption         =   "Set Stok Real = 0"
      Height          =   375
      Left            =   11280
      TabIndex        =   21
      Top             =   7440
      Width           =   1815
   End
   Begin VB.TextBox txtCariNama 
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
      Height          =   405
      Left            =   3360
      TabIndex        =   7
      Top             =   7440
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtIsi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Left            =   2760
      MaxLength       =   4
      TabIndex        =   18
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   495
      Left            =   0
      TabIndex        =   16
      Top             =   8040
      Visible         =   0   'False
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtNoClosing 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   960
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtKeterangan 
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
      Height          =   405
      Left            =   6240
      TabIndex        =   8
      Top             =   7440
      Width           =   4695
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   11760
      TabIndex        =   9
      Top             =   8040
      Width           =   1815
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   13560
      TabIndex        =   10
      Top             =   8040
      Width           =   1815
   End
   Begin VB.Frame Frame2 
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
      TabIndex        =   11
      Top             =   960
      Width           =   15375
      Begin VB.TextBox txtNoUrutan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3480
         MaxLength       =   10
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Frame Frame1 
         Caption         =   "Group by"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6720
         TabIndex        =   12
         Top             =   120
         Width           =   8535
         Begin VB.TextBox txtKriteriaJenis 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   5760
            TabIndex        =   4
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton optAsal 
            Caption         =   "Asal Barang"
            Height          =   495
            Left            =   2280
            TabIndex        =   1
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optJenis 
            Caption         =   "Jenis Barang"
            Height          =   495
            Left            =   720
            TabIndex        =   0
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            Height          =   390
            Left            =   7560
            TabIndex        =   5
            Top             =   290
            Width           =   855
         End
         Begin MSDataListLib.DataCombo dcCariData 
            Height          =   390
            Left            =   3840
            TabIndex        =   3
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   688
            _Version        =   393216
            MatchEntry      =   -1  'True
            Appearance      =   0
            Style           =   2
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
         Begin VB.OptionButton optLokasi 
            Caption         =   "Lokasi Barang"
            Enabled         =   0   'False
            Height          =   495
            Left            =   3120
            TabIndex        =   2
            Top             =   240
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cari Barang"
            Height          =   210
            Index           =   3
            Left            =   5760
            TabIndex        =   26
            Top             =   120
            Width           =   900
         End
      End
      Begin MSComCtl2.DTPicker dtpTglPenutupan 
         Height          =   450
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   794
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
         CustomFormat    =   "dd MMMM yyyy HH:mm"
         Format          =   122224643
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Penutupan"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1275
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fgData 
      Height          =   4935
      Left            =   0
      TabIndex        =   6
      Top             =   2160
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   8705
      _Version        =   393216
      Rows            =   10
      Cols            =   15
      FixedCols       =   0
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   8535
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Visible         =   0   'False
            Object.Width           =   13229
            Text            =   "F1 : Cetak"
            TextSave        =   "F1 : Cetak"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   27093
            Text            =   "Ctrl + C : Copy Stok System To Stok Real"
            TextSave        =   "Ctrl + C : Copy Stok System To Stok Real"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   17
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
      Left            =   13560
      Picture         =   "frmStokOpname.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmStokOpname.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmStokOpname.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13815
   End
   Begin VB.Label lblJmlData 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Data 0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   14160
      TabIndex        =   20
      Top             =   7440
      Width           =   1035
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cari Barang"
      Height          =   210
      Index           =   2
      Left            =   3360
      TabIndex        =   19
      Top             =   7200
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
      Height          =   210
      Index           =   0
      Left            =   6240
      TabIndex        =   14
      Top             =   7200
      Width           =   945
   End
End
Attribute VB_Name = "frmStokOpname"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrJmlStokReal() As Long
Dim subCopy As Boolean
Dim i As Integer
Dim mstrValue As Integer
Dim tglkadarluarsa2 As Date

Private Sub subLoadText()
    txtIsi.Left = fgData.Left
    mebIsi.Left = fgData.Left
    Select Case fgData.Col
        Case 5, 15, 16, 17
            txtIsi.MaxLength = 12
        Case Else
            Exit Sub
    End Select

    If fgData.Col = 15 Then
        With mebIsi
            For i = 0 To fgData.Col - 1
                .Left = .Left + fgData.ColWidth(i)
            Next i
            .Top = fgData.Top - 7

            If fgData.TopRow > 1 Then
                .Top = .Top + fgData.RowHeight(0)
                For i = fgData.TopRow To fgData.Row - 1
                    .Top = .Top + fgData.RowHeight(i)
                Next i
            Else
                For i = 0 To fgData.Row - 1
                    .Top = .Top + fgData.RowHeight(i)
                Next i
            End If

            .Width = fgData.ColWidth(fgData.Col)
            .Height = fgData.RowHeight(fgData.Row)

            .Visible = True
            .SelStart = Len(mebIsi.Text)
            .SetFocus
            .Text = IIf(fgData.TextMatrix(fgData.Row, fgData.Col) = "0", "__/__/____ __:__", fgData.TextMatrix(fgData.Row, fgData.Col))
            .SelStart = 0
            .SelLength = Len(mebIsi.Text)
        End With
    Else
        With txtIsi
            For i = 0 To fgData.Col - 1
                .Left = .Left + fgData.ColWidth(i)
            Next i
            .Visible = True
            .Top = fgData.Top - 7

            If fgData.TopRow > 1 Then
                .Top = .Top + fgData.RowHeight(0)
                For i = fgData.TopRow To fgData.Row - 1
                    .Top = .Top + fgData.RowHeight(i)
                Next i
            Else
                For i = 0 To fgData.Row - 1
                    .Top = .Top + fgData.RowHeight(i)
                Next i
            End If

            .Width = fgData.ColWidth(fgData.Col)
            .Height = fgData.RowHeight(fgData.Row)

            .Visible = True
            .SelStart = Len(txtIsi.Text)
            .SetFocus
            .Text = Trim(fgData.TextMatrix(fgData.Row, fgData.Col))
            .SelStart = 0
            .SelLength = Len(txtIsi.Text)
        End With
    End If
End Sub

Private Sub Check1_Click()
On Error GoTo errLoad
strSQL = "Select * from settingglobal where prefix='JumlahTambahanBulan'"
 Call msubRecFO(rsB, strSQL)
 
tglkadarluarsa2 = DateAdd("m", rsB("value"), dtpTglPenutupan.Value)



    If Check1.Value = vbChecked Then
        MousePointer = vbHourglass
       

        For i = 1 To fgData.Rows - 1
            With fgData
                If .TextMatrix(i, 6) = "__/__/____" Then
                .TextMatrix(i, 6) = Format(tglkadarluarsa2, "dd/MM/yyyy")
                End If
            End With
        Next i
        MousePointer = vbDefault

    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub chkSetStokReal_Click()
    On Error GoTo errLoad

    If chkSetStokReal.Value = vbChecked Then
        MousePointer = vbHourglass
        ReDim Preserve arrJmlStokReal(fgData.Rows - 1)

        For i = 1 To fgData.Rows - 1
            With fgData
                .TextMatrix(i, 5) = 0
            End With
        Next i
        MousePointer = vbDefault
    Else
        MousePointer = vbHourglass
        For i = 1 To fgData.Rows - 1
            With fgData
                .TextMatrix(i, 5) = .TextMatrix(i, 4)
            End With
        Next i
        MousePointer = vbDefault
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdCari_Click()

    Dim str, strjenis, strasal As String
    On Error GoTo errTampilkan

    MousePointer = vbHourglass
    fgData.Visible = False

    If optJenis.Value = True Then
        mstrFilter = " AND (KdDetailJenisBarang = '" & dcCariData.BoundText & "')"
    ElseIf optAsal.Value = True Then
        mstrFilter = " AND (KdAsal = '" & dcCariData.BoundText & "')"
    End If

    If dcCariData.BoundText = "" Then mstrFilter = ""
    
    
    If bolStatusFIFO = False Then
        strSQL = "SELECT distinct *, '0000000000' as Noterima FROM V_DataStokBarangMedisNonRekap WHERE (NamaBarang like '%" & txtKriteriaJenis.Text & "%') AND (KdRuangan = '" & mstrKdRuangan & "') " & mstrFilter & " ORDER BY JenisBarang, NamaBarang"
    Else
        strSQL = "SELECT distinct * FROM V_DataStokBarangMedisNonRekapFIFO WHERE (NamaBarang like '%" & txtKriteriaJenis.Text & "%') AND (KdRuangan = '" & mstrKdRuangan & "') " & mstrFilter & " ORDER BY JenisBarang, NamaBarang"

    End If
'    strSQL = "Execute DataStokBarangMedisNonRekap_V  '" & strjenis & "','" & strasal & "','%" & txtKriteriaJenis.Text & "%','" & mstrKdRuangan & "'"
    
    Call msubRecFO(rs, strSQL)
    Call subSetGrid
    If IsNull(rs(0)) Then Exit Sub
    For i = 1 To rs.RecordCount
        fgData.Rows = fgData.Rows + 1
        fgData.TextMatrix(i, 0) = IIf(IsNull(rs("JenisBarang").Value), "", rs("JenisBarang").Value)
        fgData.TextMatrix(i, 1) = IIf(IsNull(rs("NamaBarang").Value), "", rs("NamaBarang").Value)
        fgData.TextMatrix(i, 2) = IIf(IsNull(rs("KeKuatan")), "", rs("KeKuatan").Value)
        fgData.TextMatrix(i, 3) = IIf(IsNull(rs("AsalBarang").Value), "", rs("AsalBarang").Value)
        fgData.TextMatrix(i, 4) = IIf(IsNull(rs("StokSystem").Value), "", Format(rs("StokSystem").Value, "#,##0"))
        fgData.TextMatrix(i, 5) = IIf(IsNull(rs("StokSystem").Value), "", Format(rs("StokSystem").Value, "#,##0"))
        
        fgData.TextMatrix(i, 11) = IIf(IsNull(rs("KdBarang").Value), "", rs("KdBarang").Value)
        fgData.TextMatrix(i, 12) = IIf(IsNull(rs("KdAsal").Value), "", rs("KdAsal").Value)
        fgData.TextMatrix(i, 13) = IIf(IsNull(rs("KdRuangan").Value), "", rs("KdRuangan").Value)
        fgData.TextMatrix(i, 18) = rs("NoTerima").Value

        ' ambil dari rs
        Dim intStatusHarga As Integer
        strsqlx = "select MetodeHargaBarang from SuratKeputusanRuleRS where statusenabled=1"
        Call msubRecFO(rsx, strsqlx)
        If rsx.EOF = False Then
            intStatusHarga = rsx(0)
        Else
            intStatusHarga = 0
        End If
        
        Set rsB = Nothing
        If intStatusHarga = 0 Then
            strSQL = "SELECT HargaNetto1,HargaNetto2,Discount,TglKadaluarsa FROM HargaNettoBarang WHERE KdBarang='" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "'"
        Else
            If bolStatusFIFO = True Then
                strSQL = "SELECT HargaNetto1,HargaNetto2,Discount,TglKadaluarsa FROM HargaNettoBarangFIFO WHERE KdBarang='" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "' AND NoTerima ='" & rs("NoTerima") & "'"
            Else
                strSQL = "SELECT distinct HargaNetto1,HargaNetto2,Discount,TglKadaluarsa FROM V_HargaRuanganNonFIFO WHERE KdBarang='" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "' AND TglTerima=(Select min(TglTerima) from V_HargaRuanganFIFO WHERE KdBarang='" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "' AND JmlStok<>0) "
            End If

        End If
        Call msubRecFO(rsB, strSQL)
        If rsB.EOF = False Then
'        Call msubRecFO(rsb, "select dbo.FB_TakeHargaNetto1Medis('" & fgData.TextMatrix(i, 11) & "','" & fgData.TextMatrix(i, 12) & "','S','" & fgData.TextMatrix(i, 18) & "') as NoFIFO")

'        fgData.TextMatrix(i, 6) = IIf(IsNull(dbRst("TglKadaluarsa")), "__/__/____", Format(dbRst("TglKadaluarsa"), "dd/MM/yyyy"))
        fgData.TextMatrix(i, 6) = IIf(IsNull(rsB("TglKadaluarsa")), "__/__/____", Format(rsB("TglKadaluarsa"), "dd/MM/yyyy"))
        fgData.TextMatrix(i, 7) = IIf(IsNull(rsB("HargaNetto1").Value), "", Format(rsB("HargaNetto1").Value, "##,###,##0")) '+ rs("Discount").Value, "##,###,##0"))
        fgData.TextMatrix(i, 8) = IIf(IsNull(rsB("HargaNetto2").Value), "", Format(rsB("HargaNetto2").Value, "##,###,##0"))
'        fgData.TextMatrix(i, 9) = IIf(IsNull(dbRst("Discount").Value), "", Format(dbRst("Discount").Value, "##,###,##0"))
        fgData.TextMatrix(i, 9) = IIf(IsNull(rsB("Discount")), "", Format(rsB("Discount").Value, "##,###,##0"))
        
        Else
        
        fgData.TextMatrix(i, 6) = "__/__/____"
        fgData.TextMatrix(i, 7) = 0 '+ rs("Discount").Value, "##,###,##0"))
        fgData.TextMatrix(i, 8) = 0
        fgData.TextMatrix(i, 9) = 0
        
        
        End If
        
        
        fgData.TextMatrix(i, 10) = IIf(IsNull(rs("Ruangan").Value), "", rs("Ruangan").Value)
        
        
        Dim strNoTerima As String
'        Call msubRecFO(rsB, "select dbo.TakeNoFIFO_F('" & fgData.TextMatrix(i, 11) & "','" & fgData.TextMatrix(i, 12) & "','" & fgData.TextMatrix(i, 13) & "') as NoFIFO")
'        strNoTerima = IIf(IsNull(rsB("NoFIFO")), "0000000000", rsB("NoFIFO"))
'
'        Set dbRst = Nothing
''        strSQL = "SELECT Lokasi FROM StokRuangan WHERE KdBarang='" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "' AND KdRuangan='" & mstrKdRuangan & "'"
'        Call msubRecFO(dbRst, "select dbo.FB_TakeStokBrgMedis('" & fgData.TextMatrix(i, 13) & "', '" & fgData.TextMatrix(i, 11) & "','" & fgData.TextMatrix(i, 12) & "','" & strNoTerima & "') as stok")


'        fgData.TextMatrix(i, 14) = IIf(IsNull(dbRst("Lokasi")), "", dbRst("Lokasi"))
        fgData.TextMatrix(i, 14) = IIf(IsNull(rs("Lokasi")), "", rs("Lokasi"))
'        fgData.TextMatrix(i, 15) = "__/__/____ __:__"
        fgData.TextMatrix(i, 16) = IIf(IsNull(rs("JmlMax")), "", rs("JmlMax").Value)
        fgData.TextMatrix(i, 17) = IIf(IsNull(rs("JmlStokOnHand")), "", rs("JmlStokOnHand").Value)
        
        rs.MoveNext
    Next i
    MousePointer = vbDefault
    fgData.Visible = True
    If fgData.Rows < 1 Then dcCariData.SetFocus Else fgData.SetFocus: fgData.Col = 5
    lblJmlData.Caption = "Data 0 / " & fgData.Rows - 2
    Exit Sub
errTampilkan:
    MousePointer = vbDefault
    fgData.Visible = True
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad
    Dim i As Integer

    If fgData.TextMatrix(1, 11) = "" Then Exit Sub
    If MsgBox("Apakah Anda yakin akan mengupdate Stok Obat dan Alkes ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    cmdSimpan.Enabled = False
    cmdTutup.Enabled = False
    ProgressBar.Visible = True
    ProgressBar.Min = 0
    ProgressBar.Max = fgData.Rows - 2
    ProgressBar.Value = 0

    If sp_Closing = False Then Exit Sub
    For i = 1 To fgData.Rows - 2
        mstrValue = i
        ProgressBar.Value = i
        With fgData
'            If .TextMatrix(i, 15) = "__/__/____ __:__" Then
    If .TextMatrix(i, 6) = "__/__/____" Then
    MsgBox "Sebagian Barang Belum Memiliki Tanggal Kadar luarsa, Setting Terlebih Dahulu", vbInformation
    Exit Sub
    End If
                If sp_DataStokBarangMedis(.TextMatrix(i, 11), .TextMatrix(i, 12), .TextMatrix(i, 4), IIf(Len(.TextMatrix(i, 5)) = 0, 0, .TextMatrix(i, 5)), .TextMatrix(i, 7), .TextMatrix(i, 8), .TextMatrix(i, 9), .TextMatrix(i, 14), .TextMatrix(i, 6), Format(Now, "yyyy/MM/dd HH:mm"), .TextMatrix(i, 16), .TextMatrix(i, 17), .TextMatrix(i, 18)) = False Then Exit Sub
                
                strSQL = "select * from DataStokBarangMedisStokOpname Where KdBarang = '" & fgData.TextMatrix(i, 11) & "' and KdAsal = '" & fgData.TextMatrix(i, 12) & "' and KdRuangan = '" & fgData.TextMatrix(i, 13) & "'"
                Call msubRecFO(rs, strSQL)
                
                If rs.RecordCount = 0 Then
                    If sp_DataStokBarangMedisStokOpname(.TextMatrix(i, 11), .TextMatrix(i, 12), IIf(Len(.TextMatrix(i, 5)) = 0, 0, .TextMatrix(i, 5)), 0, Format(Now, "yyyy/MM/dd HH:mm:ss")) = False Then Exit Sub
                Else
                    If sp_DataStokBarangMedisStokOpname(.TextMatrix(i, 11), .TextMatrix(i, 12), IIf(Len(.TextMatrix(i, 5)) = 0, 0, .TextMatrix(i, 5)), 0, Format(Now, "yyyy/MM/dd HH:mm:ss")) = False Then Exit Sub
                End If
              
        End With
    Next i

    Call Add_HistoryLoginActivity("Add_Closing+AU_DataStokBarangNonMedis")
    ProgressBar.Visible = False
    cmdSimpan.Enabled = False
    cmdTutup.Enabled = True

    MsgBox "Stok opname selesai..", vbInformation, "Informasi"

    Exit Sub
errLoad:
    ProgressBar.Visible = False
    cmdSimpan.Enabled = True
    cmdTutup.Enabled = True
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcCariData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdCari_Click
        If fgData.Rows < 1 Then dcCariData.SetFocus Else fgData.SetFocus
    End If
End Sub

Private Sub dtpTglPenutupan_Change()
    txtKeterangan.Text = "STOK OPNAME BULAN " & UCase(MonthName(dtpTglPenutupan.Month))
End Sub

Private Sub dtpTglPenutupan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then optJenis.SetFocus
End Sub

Private Sub fgData_DblClick()
    If fgData.Row = 0 Or fgData.Row = fgData.Rows - 1 Then Exit Sub
    If fgData.TextMatrix(fgData.Row, 1) = "" Then Exit Sub
    Call fgData_KeyDown(13, 0)
End Sub

Private Sub fgData_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)

    Select Case KeyCode
        Case 13
            If fgData.Row = 0 Or fgData.Row = fgData.Rows - 1 Then Exit Sub
            If fgData.TextMatrix(fgData.Row, 1) = "" Then Exit Sub
            Call subLoadText
            txtIsi.Text = Trim(fgData.TextMatrix(fgData.Row, fgData.Col))
            txtIsi.SelStart = 0
            txtIsi.SelLength = Len(txtIsi.Text)

        Case vbKeyC
            If strCtrlKey = 4 Then
                If fgData.Row = 0 Then Exit Sub
                For i = 1 To fgData.Rows - 2
                    fgData.TextMatrix(i, 5) = fgData.TextMatrix(i, 4)
                Next i
            End If
    End Select
End Sub

Private Sub fgData_RowColChange()
    On Error Resume Next
    lblJmlData.Caption = "Data " & fgData.Row & " / " & fgData.Rows - 2
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpTglPenutupan.Value = Now
    txtKeterangan.Text = "STOK OPNAME BULAN " & UCase(MonthName(dtpTglPenutupan.Month))
    optJenis.Value = True
    Call subSetGrid
    subCopy = False
End Sub

Private Sub mebIsi_LostFocus()
    mebIsi.Visible = False
End Sub

Private Sub optAsal_Click()
    dcCariData.BoundText = ""
    Call msubDcSource(dcCariData, rs, "SELECT KdAsal, NamaAsal FROM  AsalBarang WHERE KdInstalasi='07' and StatusEnabled=1 ORDER BY NamaAsal") 'KdInstalasi='" & mstrKdInstalasiLogin & "' and
End Sub

Private Sub optAsal_GotFocus()
    Call optAsal_Click
End Sub

Private Sub optAsal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcCariData.SetFocus
End Sub

Private Sub optJenis_Click()
    dcCariData.BoundText = ""
    Call msubDcSource(dcCariData, rs, "SELECT KdDetailJenisBarang, DetailJenisBarang FROM v_S_DetailJenisBarangMedis Order By DetailJenisBarang")
End Sub

Private Sub optJenis_GotFocus()
    Call optJenis_Click
End Sub

Private Sub optJenis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcCariData.SetFocus
End Sub

Private Sub optLokasi_Click()
    dcCariData.BoundText = ""
    Call msubDcSource(dcCariData, rs, "SELECT DISTINCT Lokasi, Lokasi FROM StokRuangan WHERE KdRuangan = '" & mstrKdRuangan & "' ORDER BY Lokasi")
End Sub

Private Sub optLokasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcCariData.SetFocus
End Sub

Private Sub txtCariNama_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    Dim bolTemu As Boolean

    If KeyAscii = 13 Then
        With fgData
            .Row = 1
            .Col = 0

            bolTemu = False
            For i = 1 To .Rows - 2
                If UCase(Left(txtCariNama.Text, Len(txtCariNama.Text))) = UCase(Left(fgData.TextMatrix(i, 1), Len(txtCariNama.Text))) Then
                    bolTemu = True
                    Exit For
                End If
            Next i
            If bolTemu = True Then
                .TopRow = i: .Row = i: .Col = 5: .SetFocus
            Else
                MsgBox "Nama barang tidak ada", vbExclamation, "Validasi"
                Exit Sub
            End If
        End With
    End If

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    If KeyAscii = 13 Then
        If Trim(txtIsi.Text) = "" Then txtIsi.Text = 0
        If txtIsi.Text = 0 Then txtIsi.Text = 0

        fgData.TextMatrix(fgData.Row, fgData.Col) = txtIsi.Text
        txtIsi.Visible = False

        If fgData.RowPos(fgData.Row) >= fgData.Height - 360 Then
            fgData.SetFocus
            SendKeys "{DOWN}"
            Exit Sub
        End If
        fgData.SetFocus

    ElseIf KeyAscii = 27 Then
        txtIsi.Visible = False
        fgData.SetFocus
    End If
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = 44) Then KeyAscii = 0
End Sub

Private Function sp_Closing() As Boolean
    On Error GoTo errLoad

    sp_Closing = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoClosing", adChar, adParamInput, 10, IIf(Len(Trim(txtNoClosing.Text)) = 0, Null, Trim(txtNoClosing.Text)))
        .Parameters.Append .CreateParameter("TglClosing", adDate, adParamInput, , Format(dtpTglPenutupan.Value, "yyyy/MM/dd HH:mm"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 200, IIf(Len(Trim(txtKeterangan.Text)) = 0, Null, Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("OutputNoClosing", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "Add_Closing"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data closing", vbCritical, "Validasi"
            sp_Closing = False
        Else
            txtNoClosing.Text = .Parameters("OutputNoClosing").Value
        End If
    End With

    Exit Function
errLoad:
    sp_Closing = False
    Call msubPesanError
    cmdSimpan.Enabled = True
    cmdTutup.Enabled = True
End Function

Private Function sp_DataStokBarangMedisStokOpname(f_KdBarang As String, f_KdAsal As String, f_JmlStokReal As Double, f_JmlKirim As Double, f_tglInputStok As String) As Boolean
    On Error GoTo errLoad

    sp_DataStokBarangMedisStokOpname = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoClosing", adChar, adParamInput, 10, Trim(txtNoClosing.Text))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("JmlStokReal", adDouble, adParamInput, , CDbl(f_JmlStokReal))
        .Parameters.Append .CreateParameter("JmlKirim", adDouble, adParamInput, , CDbl(f_JmlKirim))
        .Parameters.Append .CreateParameter("TglInputStok", adDate, adParamInput, , IIf(f_tglInputStok = "__/__/____ __:__:__", Null, Format(f_tglInputStok, "yyyy/MM/dd hh:mm:ss")))
        .Parameters.Append .CreateParameter("NoUrutanStok", adChar, adParamInput, 10, IIf(Trim(txtNoUrutan.Text) = "", Null, Trim(txtNoUrutan.Text)))
        .Parameters.Append .CreateParameter("OutputNoUrutanStok", adChar, adParamOutput, 10, Null)
         
        .ActiveConnection = dbConn
        .CommandText = "Add_DataStokBarangMedisStokOpname"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data Stok Opname Obat dan Alat Kesehatan", vbCritical, "Validasi"
            sp_DataStokBarangMedisStokOpname = False
            cmdSimpan.Enabled = True
            cmdTutup.Enabled = True
        End If

         If Not IsNull(.Parameters("OutputNoUrutanStok").Value) Then txtNoUrutan.Text = .Parameters("OutputNoUrutanStok").Value
    End With

    Exit Function
errLoad:
    Call msubPesanError
    cmdSimpan.Enabled = True
    cmdTutup.Enabled = True
End Function

Private Function sp_DataStokBarangMedis(f_KdBarang As String, f_KdAsal As String, f_JmlStokSystem As Double, f_JmlStokReal As Double, f_HargaNetto1 As Double, f_HargaNetto2 As Double, f_Discount As Double, f_Lokasi As String, f_tglKadaluarsa As String, f_tglInputStok As String, f_JmlMax As String, f_JmlStokOnHand As String, f_NoTerima As String) As Boolean
    On Error GoTo errLoad

    sp_DataStokBarangMedis = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoClosing", adChar, adParamInput, 10, Trim(txtNoClosing.Text))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("JmlStokSystem", adDouble, adParamInput, , CDbl(f_JmlStokSystem))
        .Parameters.Append .CreateParameter("JmlStokReal", adDouble, adParamInput, , CDbl(f_JmlStokReal))
        .Parameters.Append .CreateParameter("HargaNetto1", adCurrency, adParamInput, , f_HargaNetto1)
        .Parameters.Append .CreateParameter("HargaNetto2", adCurrency, adParamInput, , f_HargaNetto2)
        .Parameters.Append .CreateParameter("Discount", adCurrency, adParamInput, , f_Discount)
        .Parameters.Append .CreateParameter("Lokasi", adVarChar, adParamInput, 12, IIf(f_Lokasi = "", Null, f_Lokasi))
        .Parameters.Append .CreateParameter("TglKadaluarsa", adDate, adParamInput, , IIf(f_tglKadaluarsa = "__/__/____", Null, Format(f_tglKadaluarsa, "yyyy/MM/dd HH:mm")))
        .Parameters.Append .CreateParameter("TglInputStok", adDate, adParamInput, , IIf(f_tglInputStok = "__/__/____ __:__", Null, Format(f_tglInputStok, "yyyy/MM/dd HH:mm")))
        .Parameters.Append .CreateParameter("JmlMax", adInteger, adParamInput, , IIf(f_JmlMax = "", Null, f_JmlMax))
        .Parameters.Append .CreateParameter("JmlStokOnHand", adInteger, adParamInput, , IIf(f_JmlStokOnHand = "", Null, f_JmlStokOnHand))
        .Parameters.Append .CreateParameter("NoTerima", adChar, adParamInput, 10, f_NoTerima)
        .ActiveConnection = dbConn
        .CommandText = "AU_DataStokBarangMedis"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data Stok Opname Obat dan Alat Kesehatan", vbCritical, "Validasi"
            sp_DataStokBarangMedis = False
            cmdSimpan.Enabled = True
            cmdTutup.Enabled = True
        End If
    End With

    Exit Function
errLoad:
    Call msubPesanError
    cmdSimpan.Enabled = True
    cmdTutup.Enabled = True
End Function

Private Sub subSetGrid()
    With fgData
        .Clear
        .Cols = 19
        .Rows = 2

        .TextMatrix(0, 0) = "JenisBarang"
        .TextMatrix(0, 1) = "NamaBarang"
        .TextMatrix(0, 2) = "Kekuatan"
        .TextMatrix(0, 3) = "AsalBarang"
        .TextMatrix(0, 4) = "StokSystem"
        .TextMatrix(0, 5) = "StokReal"
        .TextMatrix(0, 6) = "TglKadaluarsa"
        .TextMatrix(0, 7) = "HargaNetto1"
        .TextMatrix(0, 8) = "HargaNetto2"
        .TextMatrix(0, 9) = "Discount"
        .TextMatrix(0, 10) = "Ruangan"
        .TextMatrix(0, 11) = "KdBarang"
        .TextMatrix(0, 12) = "KdAsal"
        .TextMatrix(0, 13) = "KdRuangan"
        .TextMatrix(0, 14) = "Lokasi"
        .TextMatrix(0, 16) = "JmlMax"
        .TextMatrix(0, 17) = "StockOnHand"
        .TextMatrix(0, 18) = "NoTerima"

        .RowHeight(0) = 400

        .MergeCells = flexMergeFree
        .MergeCol(0) = True
        .MergeCol(3) = True

        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignRightCenter

        .ColWidth(0) = 1300
        .ColWidth(1) = 2500
        .ColWidth(2) = 0
        .ColWidth(3) = 1050
        .ColWidth(4) = 1100
        .ColWidth(5) = 1100
        .ColWidth(6) = 1100
        .ColWidth(7) = 1200
        .ColWidth(8) = 1200
        .ColWidth(9) = 1000
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        .ColWidth(14) = 0
         .ColWidth(15) = 0
        .ColWidth(16) = 1200
        .ColWidth(17) = 1200
        .ColWidth(18) = 1200
        
    End With
End Sub

Private Sub txtIsi_LostFocus()
    txtIsi.Visible = False
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

'
Private Sub mebIsi_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    If KeyAscii = 13 Then
'        If mebIsi.Text = "__/__/____ __:__" Then
'                Exit Function
'            End If
  
'      If funcCekTglValidasi("Tanggal", mebIsi) = "NoErr" Or mebIsi.Text = "__/__/____ __:__" Then
  
        fgData.TextMatrix(fgData.Row, fgData.Col) = mebIsi.Text
        mebIsi.Visible = False

        If fgData.RowPos(fgData.Row) >= fgData.Height - 360 Then
            fgData.SetFocus
            SendKeys "{DOWN}"
            Exit Sub
        End If
        fgData.SetFocus
'      End If
    ElseIf KeyAscii = 27 Then
        mebIsi.Visible = False
        fgData.SetFocus
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtKriteriaJenis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdCari_Click
End Sub
