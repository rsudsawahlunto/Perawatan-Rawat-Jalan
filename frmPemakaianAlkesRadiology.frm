VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmPemakaianBahanAlat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pemakaian Bahan dan Alat"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPemakaianAlkesRadiology.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmPemakaianAlkesRadiology.frx":0CCA
   ScaleHeight     =   6750
   ScaleWidth      =   13185
   Begin VB.TextBox txtKdRuanganPerujuk 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   3600
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   5880
      Width           =   13095
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   9120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   10920
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame8 
      Height          =   3855
      Left            =   0
      TabIndex        =   3
      Top             =   2040
      Width           =   13095
      Begin MSDataListLib.DataCombo dcPelayanan 
         Height          =   330
         Left            =   2640
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataGridLib.DataGrid dgObatAlkes 
         Height          =   2535
         Left            =   2640
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   9735
         _ExtentX        =   17171
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
      Begin MSDataListLib.DataCombo dcStatusKontras 
         Height          =   330
         Left            =   8760
         TabIndex        =   14
         Top             =   2760
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcStatusHasil 
         Height          =   330
         Left            =   8160
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcAsalBarang 
         Height          =   330
         Left            =   7320
         TabIndex        =   12
         Top             =   2400
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtIsi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   3720
         TabIndex        =   11
         Top             =   2400
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   420
         Left            =   9600
         TabIndex        =   0
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid fgAlkes 
         Height          =   3615
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   6376
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
         HighLight       =   2
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Pemakaian"
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
      TabIndex        =   5
      Top             =   1080
      Width           =   13095
      Begin VB.TextBox txtKeperluan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4320
         TabIndex        =   6
         Top             =   480
         Width           =   5175
      End
      Begin MSComCtl2.DTPicker dtpTglPeriksa 
         Height          =   330
         Left            =   9720
         TabIndex        =   18
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
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
         CustomFormat    =   "dd MMMM yyyy HH:mm"
         Format          =   57344003
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSDataListLib.DataCombo dcNamaPelayanan 
         Height          =   330
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Pemakaian"
         Height          =   210
         Left            =   9720
         TabIndex        =   19
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label lblNamaPasien 
         AutoSize        =   -1  'True
         Caption         =   "Pemakaian Untuk "
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label lblJnsKlm 
         AutoSize        =   -1  'True
         Caption         =   "Keperluan"
         Height          =   210
         Left            =   4320
         TabIndex        =   7
         Top             =   240
         Width           =   810
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
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPemakaianAlkesRadiology.frx":1994
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   11400
      Picture         =   "frmPemakaianAlkesRadiology.frx":4355
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPemakaianAlkesRadiology.frx":50DD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmPemakaianBahanAlat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tempStatusTampil As Boolean
Dim subJenisHargaNetto  As Integer

Private Sub dcNamaPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcNamaPelayanan.MatchedWithList = True Then txtKeperluan.SetFocus
        strSQL = "select  kdpelayananrs,namapelayanan from V_ListPemakaianBahan  WHERE (namapelayanan LIKE '%" & dcNamaPelayanan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcNamaPelayanan.BoundText = rs(0).Value
        dcNamaPelayanan.Text = rs(1).Value
    End If
End Sub

Private Sub dcPelayanan_Change()
On Error GoTo errLoad
    fgAlkes.TextMatrix(fgAlkes.Row, 0) = dcPelayanan.Text
    fgAlkes.TextMatrix(fgAlkes.Row, 12) = dcPelayanan.BoundText
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call dcPelayanan_Change
        dcPelayanan.Visible = False
        fgAlkes.Col = 1
        fgAlkes.SetFocus
    End If
End Sub

Private Sub dcPelayanan_LostFocus()
    dcPelayanan.Visible = False
End Sub

Private Sub dcStatusHasil_Change()
On Error GoTo errLoad
    fgAlkes.TextMatrix(fgAlkes.Row, 5) = dcStatusHasil.Text
    fgAlkes.TextMatrix(fgAlkes.Row, 14) = dcStatusHasil.BoundText

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcStatusHasil_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call dcStatusHasil_Change
        dcPelayanan.Visible = False
        fgAlkes.Col = 8
        fgAlkes.SetFocus
    End If
End Sub

Private Sub cmdSimpan_Click()
Dim i As Integer
On Error GoTo aa
If fgAlkes.TextMatrix(1, 10) = "" Then Exit Sub

Set dbcmd = New ADODB.Command
Set dbcmd.ActiveConnection = dbConn
With fgAlkes
For i = 1 To .Rows - 1
If .TextMatrix(i, 10) = "" Then GoTo lanjut_
    If sp_PemakaianBahanAlat(.TextMatrix(i, 10), .TextMatrix(i, 11), .TextMatrix(i, 3), .TextMatrix(i, 4), .TextMatrix(i, 6), .TextMatrix(i, 8)) = False Then Exit Sub
lanjut_:
Next i
End With


MsgBox "Penyimpanan Data Sukses !", vbInformation, "Informasi"
cmdSimpan.Enabled = False
Exit Sub
aa:
    msubPesanError
End Sub

Private Function sp_PemakaianBahanAlat(f_KdBarang As String, f_KdAsal As String, _
    f_Satuan As String, f_Jumlah As Double, f_HargaSatuan As Currency, f_HargaBeli As String) As Boolean
On Error GoTo errLoad
    Dim i As Integer
    sp_PemakaianBahanAlat = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdPelayananRS", adVarChar, adParamInput, 6, dcNamaPelayanan.BoundText)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("TglPemakaian", adDate, adParamInput, , Format(dtpTglPeriksa.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("SatuanJml", adChar, adParamInput, 1, f_Satuan)
        '2010/11/22 ubah arief Cdec = Cdbl agar penyimpanan pd tabel bisa koma
        '.Parameters.Append .CreateParameter("JmlBarang", adDouble, adParamInput, , CDec(f_Jumlah))
        .Parameters.Append .CreateParameter("JmlBarang", adDouble, adParamInput, , CDbl(f_Jumlah))
        
        .Parameters.Append .CreateParameter("HargaSatuan", adCurrency, adParamInput, , f_HargaSatuan)
        .Parameters.Append .CreateParameter("Keperluan", adVarChar, adParamInput, 200, IIf(txtKeperluan.Text = "", Null, txtKeperluan.Text))
        .Parameters.Append .CreateParameter("HargaBeli", adCurrency, adParamInput, , f_HargaBeli)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 10, mstrKdRuangan)
        .Parameters.Append .CreateParameter("idUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PemakaianBahanAlat"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_PemakaianBahanAlat = False
        
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
Exit Function
errLoad:
    sp_PemakaianBahanAlat = False
    msubPesanError
End Function
Private Sub cmdTambah_Click()
Dim i As Integer

'    If Periksa("datacombo", dcPelayanan, "Nama pelayanan kosong") = False Then Exit Sub
'    If Periksa("text", txtNamaBrg, "Nama barang kosong") = False Then Exit Sub
'    If Periksa("text", txtjml, "Jumlah barang kosong") = False Then Exit Sub
'    If Periksa("datacombo", dcStatusHasil, "Status hasil kosong") = False Then Exit Sub
'
'    With fgAlkes
'        For i = 1 To .Rows - 1
'            If .TextMatrix(i, 0) = dcPelayanan.Text And _
'                .TextMatrix(i, 1) = txtNamaBrg And _
'                .TextMatrix(i, 2) = txtasalbarang And _
'                .TextMatrix(i, 3) = txtsatuam And _
'                .TextMatrix(i, 5) = dcStatusHasil.Text Then Exit Sub
'        Next i
'    End With

    ' cek stok barang
'    Set dbcmd = New ADODB.Command
'    Set dbcmd.ActiveConnection = dbConn
'    With dbcmd
'        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
'        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, txtkdbarang)
'        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, txtkdAsal)
'        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, strNKdRuangan)
'        .Parameters.Append .CreateParameter("Satuan", adChar, adParamInput, 1, txtsatuam)
'        .Parameters.Append .CreateParameter("JmlBrg", adInteger, adParamInput, , CInt(txtjml))
'        .Parameters.Append .CreateParameter("OutputPesan", adChar, adParamOutput, 1, Null)
'
'        .CommandText = "dbo.Check_StokBarangRuangan"
'        .CommandType = adCmdStoredProc
'        .Execute
'
'        If .Parameters("OutputPesan") = "T" Then
'            deleteADOCommandParameters dbcmd
'            MsgBox "Stok Barang Tidak Ada"
'            txtjml.SetFocus
'            Exit Sub
'        Else
'            Call Add_HistoryLoginActivity("Check_StokBarangRuangan")
'        End If
'        deleteADOCommandParameters dbcmd
'    End With
'
'    strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "','" & txtkdAsal.Text & "', " & CCur(strNHargaSatuan) & ")  as HargaSatuan"
'    Call msubRecFO(rs, strSQL)
'    If rs.EOF = True Then strNHargaSatuan = 0 Else strNHargaSatuan = rs(0).Value
'
'    fgAlkes.AddItem _
'        dcPelayanan.Text & vbTab & _
'        txtNamaBrg & vbTab & _
'        txtasalbarang & vbTab & _
'        txtsatuam & vbTab & _
'        txtjml.Text & vbTab & _
'        dcStatusHasil.Text & vbTab & _
'        strNHargaSatuan & vbTab & _
'        strNTotal & vbTab & _
'        dcStatusKontras.Text & vbTab & _
'        val(txtJmlExpose) & vbTab & _
'        txtkdbarang & vbTab & _
'        txtkdAsal & vbTab & _
'        dcPelayanan.BoundText & vbTab & _
'        strNKdRuangan & vbTab & _
'        dcStatusHasil.BoundText & vbTab & _
'        dcStatusKontras.BoundText & vbTab & _
'        "Baru" _
'        , fgAlkes.Rows - 1
End Sub


Private Sub cmdTutup_Click()
    
    Unload Me
   
End Sub

Private Sub cmdHapus_Click()
    With fgAlkes
        If .Row = .Rows Then Exit Sub
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        If .TextMatrix(.Row, 12) = "SudahAda" Then Exit Sub
'        .RemoveItem .Row
        msubRemoveItem fgAlkes, .Row
    End With
End Sub

Private Sub dcJnsPelayanan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    dcPelayanan.SetFocus
End If
End Sub

Private Sub dcStatusHasil_LostFocus()
    dcStatusHasil.Visible = False
End Sub

Private Sub dcStatusKontras_Change()
On Error GoTo errLoad
    fgAlkes.TextMatrix(fgAlkes.Row, 8) = dcStatusKontras.Text
    fgAlkes.TextMatrix(fgAlkes.Row, 15) = dcStatusKontras.BoundText

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcStatusKontras_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call dcStatusKontras_Change
        dcStatusKontras.Visible = False
        fgAlkes.Col = 9
        fgAlkes.SetFocus
    End If
End Sub

Private Sub dcStatusKontras_LostFocus()
    dcStatusKontras.Visible = False
End Sub

Private Sub dgObatAlkes_DblClick()
On Error GoTo errLoad
    tempStatusTampil = True
    fgAlkes.TextMatrix(fgAlkes.Row, 10) = dgObatAlkes.Columns("KdBarang")
    fgAlkes.TextMatrix(fgAlkes.Row, 3) = dgObatAlkes.Columns("Satuan")
    fgAlkes.TextMatrix(fgAlkes.Row, 11) = dgObatAlkes.Columns("KdAsal")
    fgAlkes.TextMatrix(fgAlkes.Row, 2) = dgObatAlkes.Columns("AsalBarang")
     fgAlkes.TextMatrix(fgAlkes.Row, 6) = dgObatAlkes.Columns("HargaBarang")
     fgAlkes.TextMatrix(fgAlkes.Row, 1) = dgObatAlkes.Columns("NamaBarang")
    tempStatusTampil = False
    dgObatAlkes.Visible = False
    fgAlkes.TextMatrix(fgAlkes.Row, 4) = 1
    fgAlkes.TextMatrix(fgAlkes.Row, 7) = dgObatAlkes.Columns("HargaBarang")
    fgAlkes.TextMatrix(fgAlkes.Row, 8) = dgObatAlkes.Columns("HargaBarang")
    fgAlkes.SetFocus
    fgAlkes.Col = 4
    
Exit Sub
errLoad:
End Sub


Private Sub dgObatAlkes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call dgObatAlkes_DblClick
End If
End Sub

Private Sub dtpTglPeriksa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then fgAlkes.SetFocus: fgAlkes.Col = 1
End Sub

Private Sub fgAlkes_DblClick()
    Call fgAlkes_KeyDown(13, 0)
End Sub

Private Sub fgAlkes_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
    Select Case KeyCode
        Case 13
            If fgAlkes.Col = fgAlkes.Cols - 1 Then
                If fgAlkes.TextMatrix(fgAlkes.Row, 2) <> "" Then
                    If fgAlkes.TextMatrix(fgAlkes.Rows - 1, 2) <> "" Then fgAlkes.Rows = fgAlkes.Rows + 1
                    fgAlkes.Row = fgAlkes.Rows - 1
                    fgAlkes.Col = 1
                Else
                    fgAlkes.Col = 1
                End If
            Else
                For i = 0 To fgAlkes.Cols - 2
                    If fgAlkes.Col = fgAlkes.Cols - 1 Then Exit For
                    fgAlkes.Col = fgAlkes.Col + 1
                    If fgAlkes.ColWidth(fgAlkes.Col) > 0 Then Exit For
                Next i
            End If
            fgAlkes.SetFocus
            
            If fgAlkes.Col > 7 Then
               fgAlkes.Rows = fgAlkes.Rows + 1
               fgAlkes.Row = fgAlkes.Rows - 1
               fgAlkes.Col = 0
               fgAlkes.SetFocus
            End If
            
        Case 27
            dgObatAlkes.Visible = False
            
        Case vbKeyDelete
            With fgAlkes
                If .Row = .Rows Then Exit Sub
                If .Row = 0 Then Exit Sub
                
                If .Rows = 2 Then
                    For i = 0 To .Cols - 1
                        .TextMatrix(1, i) = ""
                    Next i
                    Exit Sub
                Else
                    .RemoveItem .Row
                End If
            End With
            
    End Select
End Sub

Private Sub fgAlkes_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad

    txtIsi.Text = ""
    If Not (KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ Or KeyAscii = 32 Or KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace Or KeyAscii = Asc(".")) Then
        KeyAscii = 0
        Exit Sub
    End If
            
    Select Case fgAlkes.Col
        Case 0 'nama pemeriksaan
             Call subLoadDataCombo(dcPelayanan)
        
        Case 1 'Nama Barang
            txtIsi.MaxLength = 0
            Call subLoadText
            txtIsi.Text = Chr(KeyAscii)
            txtIsi.SelStart = Len(txtIsi.Text)
        
        Case 2 'asal barang
             Call subLoadDataCombo(dcAsalBarang)
        
        Case 3 'satauan hasil
             
    
        Case 4 'Jumlah
            txtIsi.MaxLength = 4
            Call subLoadText
            txtIsi.Text = Chr(KeyAscii)
            txtIsi.SelStart = Len(txtIsi.Text)
            
        Case 5
            Call subLoadDataCombo(dcStatusHasil)
            
        Case 8 'status kontras
             Call subLoadDataCombo(dcStatusKontras)
             
        Case 9 'Jumlah expose
            txtIsi.MaxLength = 4
            Call subLoadText
            txtIsi.Text = Chr(KeyAscii)
            txtIsi.SelStart = Len(txtIsi.Text)
            
             
             
    End Select
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Load()
On Error GoTo errLoad

    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    dtpTglPeriksa.Value = Format(Now, "yyyy/MMMM/dd HH:mm:ss")
    Call subLoadGridSource
    Call SubLoadDCSource
    
    '2010/11/22 add arief pengambilan JenisHargaNetto tidak ada, krn PemakaianBahanAlat tdk dibebankan ke pasien
    subJenisHargaNetto = 1

Exit Sub
errLoad:
    Call msubPesanError
End Sub
Private Sub subLoadText()
Dim i As Integer
    txtIsi.Left = fgAlkes.Left
    For i = 0 To fgAlkes.Col - 1
        txtIsi.Left = txtIsi.Left + fgAlkes.ColWidth(i)
    Next i
    txtIsi.Visible = True
    txtIsi.Top = fgAlkes.Top - 7
    
    For i = 0 To fgAlkes.Row - 1
        txtIsi.Top = txtIsi.Top + fgAlkes.RowHeight(i)
    Next i
    
    If fgAlkes.TopRow > 1 Then
        txtIsi.Top = txtIsi.Top - ((fgAlkes.TopRow - 1) * fgAlkes.RowHeight(1))
    End If
    
    txtIsi.Width = fgAlkes.ColWidth(fgAlkes.Col)
    txtIsi.Height = fgAlkes.RowHeight(fgAlkes.Row)
    
    txtIsi.Visible = True
    txtIsi.SelStart = Len(txtIsi.Text)
    txtIsi.SetFocus
End Sub
Private Sub txtIsi_Change()
On Error GoTo errLoad
Dim i As Integer
    Select Case fgAlkes.Col
        Case 0  ' nama pemeriksaan
                   
        Case 1 ' nama barang
            '2010/11/22 ubah arief filter KdRuangan diaktifkan, krn pengurangan stok berdsrkan ruangan
             If tempStatusTampil = True Then Exit Sub
                If subJenisHargaNetto = 2 Then
'                    strSQL = "select  TOP 100 JenisBarang, RuanganPelayanan, NamaBarang, Kekuatan, AsalBarang, Satuan, HargaBarang, JmlStok, Discount, KdBarang, KdAsal from V_HargaBarangNStok2 " & _
                        " where NamaBarang like '" & txtIsi.Text & "%' ORDER BY NamaBarang"  'AND KdRuangan = '" & mstrKdRuangan & "'
                    strSQL = "select  TOP 100 JenisBarang, RuanganPelayanan, NamaBarang, Kekuatan, AsalBarang, Satuan, HargaBarang, JmlStok, Discount, KdBarang, KdAsal from V_HargaBarangNStok2 " & _
                        " where NamaBarang like '" & txtIsi.Text & "%' AND KdRuangan = '" & mstrKdRuangan & "' ORDER BY NamaBarang"
                Else
'                    strSQL = "select  TOP 100 JenisBarang, RuanganPelayanan, NamaBarang, Kekuatan, AsalBarang, Satuan, HargaBarang, JmlStok, Discount, KdBarang, KdAsal from V_HargaBarangNStok1 " & _
                        " where NamaBarang like '" & txtIsi.Text & "%'   ORDER BY NamaBarang"  'AND KdRuangan = '" & mstrKdRuangan & "'
                    strSQL = "select  TOP 100 JenisBarang, RuanganPelayanan, NamaBarang, Kekuatan, AsalBarang, Satuan, HargaBarang, JmlStok, Discount, KdBarang, KdAsal from V_HargaBarangNStok1 " & _
                        " where NamaBarang like '" & txtIsi.Text & "%' AND KdRuangan = '" & mstrKdRuangan & "' ORDER BY NamaBarang"
                End If
                Call msubRecFO(dbRst, strSQL)
                
                Set dgObatAlkes.DataSource = dbRst
                With dgObatAlkes
                    .Columns("RuanganPelayanan").Width = 0
                    .Columns("JenisBarang").Width = 1250
                    .Columns("NamaBarang").Width = 2200
                    .Columns("AsalBarang").Width = 1000
            '        .Columns("JenisPasien").Width = 1100
                    .Columns("Satuan").Width = 675
                    .Columns("HargaBarang").Width = 1200
                    .Columns("JmlStok").Width = 700
                    
                    .Columns("HargaBarang").NumberFormat = "#,###"
                    .Columns("HargaBarang").Alignment = dbgRight
                    
                    .Columns("JmlStok").NumberFormat = "#,###"
                    .Columns("JmlStok").Alignment = dbgRight
                
            '        .Columns("NamaGenerik").Width = 0
                    .Columns("Discount").Width = 0
                    .Columns("KdBarang").Width = 0
                    .Columns("KdAsal").Width = 0
            '        .Columns("KdRuangan").Width = 0
            '        .Columns("KdKelompokPasien").Width = 0
            '        .Columns("IdPenjamin").Width = 0
            '        .Columns("JenisHargaNetto").Width = 0
                
              
                dgObatAlkes.Visible = True
                
                .Left = 0
'                .Top = 2950
                .Top = 650
                '.Top = (frmPemesanankeSupplier.Height / 2) + 250
                .Height = 3850
                .Visible = True
  '              For i = 1 To fgAlkes.Row - 1
  '                  .Top = .Top + fgAlkes.RowHeight(i)
  '              Next i
                If fgAlkes.TopRow > 1 Then
                    .Top = .Top - ((fgAlkes.TopRow - 1) * fgAlkes.RowHeight(1))
                End If
            End With

      
        Case Else
            dgObatAlkes.Visible = False
            Exit Sub
    End Select
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub
Private Sub subLoadDataCombo(s_DcName As Object)
Dim i As Integer
    s_DcName.Left = fgAlkes.Left
    For i = 0 To fgAlkes.Col - 1
        s_DcName.Left = s_DcName.Left + fgAlkes.ColWidth(i)
    Next i
    s_DcName.Visible = True
    s_DcName.Top = fgAlkes.Top - 7
    
    For i = 0 To fgAlkes.Row - 1
        s_DcName.Top = s_DcName.Top + fgAlkes.RowHeight(i)
    Next i
    
    If fgAlkes.TopRow > 1 Then
        s_DcName.Top = s_DcName.Top - ((fgAlkes.TopRow - 1) * fgAlkes.RowHeight(1))
    End If
    
    s_DcName.Width = fgAlkes.ColWidth(fgAlkes.Col)
    s_DcName.Height = fgAlkes.RowHeight(fgAlkes.Row)
    
    s_DcName.Visible = True
    s_DcName.SetFocus
End Sub

Private Sub txtIsi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then If dgObatAlkes.Visible = True Then If dgObatAlkes.ApproxCount > 0 Then dgObatAlkes.SetFocus
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
Dim i As Integer
On Error GoTo errLoad
    If KeyAscii = 13 Then
        Select Case fgAlkes.Col
            Case 0
                If dgObatAlkes.Visible = True Then
                    dgObatAlkes.SetFocus
                    Exit Sub
                Else
                    fgAlkes.SetFocus
                    fgAlkes.Col = 1
                End If
                
            Case 1
                If dgObatAlkes.Visible = True Then
                    dgObatAlkes.SetFocus
                    Exit Sub
                Else
                    fgAlkes.SetFocus
                    fgAlkes.Col = 4
                End If
            Case 4
                fgAlkes.TextMatrix(fgAlkes.Row, 4) = txtIsi.Text
                '2010/11/22 ubah arief conversi fgAlkes.TextMatrix(fgAlkes.Row, 4) Cint = Cdbl
                'fgAlkes.TextMatrix(fgAlkes.Row, 7) = CCur(fgAlkes.TextMatrix(fgAlkes.Row, 6)) * Cint(fgAlkes.TextMatrix(fgAlkes.Row, 4))
                fgAlkes.TextMatrix(fgAlkes.Row, 7) = CCur(fgAlkes.TextMatrix(fgAlkes.Row, 6)) * CDbl(fgAlkes.TextMatrix(fgAlkes.Row, 4))
                fgAlkes.SetFocus
                fgAlkes.Col = 5
            
                
            Case 9
               
                With fgAlkes
                    .TextMatrix(.Row, .Col) = txtIsi.Text
                    If .TextMatrix(.Rows - 1, 2) = "" Then
                        .Row = .Rows - 1
                        .Col = 0
                    Else
                        .SetFocus
                        .Rows = .Rows + 1
                        .Row = .Row + 1
                        .Col = 0
                    End If
                End With
                'fgAlkes.Col = 6
        End Select
                        
        txtIsi.Visible = False
                        
        If fgAlkes.RowPos(fgAlkes.Row) >= fgAlkes.Height - 360 Then
            fgAlkes.SetFocus
            SendKeys "{DOWN}"
            Exit Sub
        End If
        
    ElseIf KeyAscii = 27 Then
        txtIsi.Visible = False
        dgObatAlkes.Visible = False
        fgAlkes.SetFocus
    End If

    If fgAlkes.Col = 4 Then
        If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",")) Then KeyAscii = 0
    End If
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtIsi_LostFocus()
    txtIsi.Visible = False
End Sub


Private Sub txtjml_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcStatusHasil.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub




Private Sub SubLoadDCSource()
On Error GoTo errLoad

    strSQL = "SELECT * FROM StatusHasil"
    Call msubDcSource(dcStatusHasil, rs, strSQL)
    
    strSQL = "SELECT * FROM StatusKontras"
    Call msubDcSource(dcStatusKontras, rs, strSQL)
    
    strSQL = "select  kdpelayananrs,namapelayanan from V_ListPemakaianBahan  " '"
    Call msubDcSource(dcNamaPelayanan, dbRst, strSQL)

    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadGridSource()
    With fgAlkes
        .Clear
        .Cols = 17
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 0
        .TextMatrix(0, 1) = "Nama Barang"
        .ColWidth(1) = 3700
        .TextMatrix(0, 2) = "Asal Barang"
        .ColWidth(2) = 1800
        .TextMatrix(0, 3) = "Satuan"
        .ColWidth(3) = 1000
        .TextMatrix(0, 4) = "Jumlah"
        .ColWidth(4) = 1000
        .TextMatrix(0, 5) = "Status Hasil"
        .ColWidth(5) = 0
        .TextMatrix(0, 6) = "Harga Satuan"
        .ColWidth(6) = 1800
        .TextMatrix(0, 7) = "Total Harga"
        .ColWidth(7) = 2000
        .TextMatrix(0, 8) = "Harga Beli"
        .ColWidth(8) = 1200
        .TextMatrix(0, 9) = ""
        .ColWidth(9) = 0
    
        .ColWidth(10) = 0 'KdBarang
        .ColWidth(11) = 0 'KdAsal
        .ColWidth(12) = 0 'KdPelayanRS
        .ColWidth(13) = 0 'KdRuangan
        .ColWidth(14) = 0 'KdStatusHasil
        .ColWidth(15) = 0 'KdStatusKontras
        .ColWidth(16) = 0 '
    End With
End Sub

Private Sub txtKeperluan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglPeriksa.SetFocus
End Sub

Private Sub txtPemakaian_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeperluan.SetFocus
End Sub
