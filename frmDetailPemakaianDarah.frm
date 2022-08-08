VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmDetailPemakaianDarah 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pemakaian Darah"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8835
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDetailPemakaianDarah.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   8835
   Begin VB.ComboBox cbSatuanJml 
      Appearance      =   0  'Flat
      Height          =   330
      ItemData        =   "frmDetailPemakaianDarah.frx":0CCA
      Left            =   2520
      List            =   "frmDetailPemakaianDarah.frx":0CD4
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtIsi 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Left            =   4440
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo dcAsalDarah 
      Height          =   330
      Left            =   2520
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcGolDarah 
      Height          =   330
      Left            =   2520
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcBentukDarah 
      Height          =   330
      Left            =   2520
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   465
      Left            =   7080
      TabIndex        =   2
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   465
      Left            =   5280
      TabIndex        =   1
      Top             =   4680
      Width           =   1695
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   3
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
   Begin MSFlexGridLib.MSFlexGrid fgData 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   6165
      _Version        =   393216
      FixedCols       =   0
      BackColorSel    =   -2147483643
      FocusRect       =   2
      HighLight       =   2
      Appearance      =   0
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7440
      Picture         =   "frmDetailPemakaianDarah.frx":0CDE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1395
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDetailPemakaianDarah.frx":1A66
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "frmDetailPemakaianDarah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub subSetGrid()
    On Error GoTo hell
    With fgData
        .Clear
        .Rows = 2
        .Cols = 8

        .RowHeight(0) = 400

        .TextMatrix(0, 0) = "Bentuk Darah"
        .TextMatrix(0, 1) = "Gol. Darah"
        .TextMatrix(0, 2) = "Asal Darah"
        .TextMatrix(0, 3) = "Satuan"
        .TextMatrix(0, 4) = "Qty Darah"
        .TextMatrix(0, 5) = "KdBentukDarah"
        .TextMatrix(0, 6) = "KdGolDarah"
        .TextMatrix(0, 7) = "KdAsalDarah"

        .ColWidth(0) = 3500
        .ColWidth(1) = 1000
        .ColWidth(2) = 1500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1500
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0

        .ColAlignment(4) = flexAlignRightCenter
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub subLoadDcSource()
    On Error GoTo hell

    Call msubDcSource(dcBentukDarah, rs, "Select KdBentukDarah,BentukDarah From BentukDarah Where StatusEnabled=1")
    Call msubDcSource(dcGolDarah, rs, "Select KdGolonganDarah,GolonganDarah From GolonganDarah Where StatusEnabled=1")
    Call msubDcSource(dcAsalDarah, rs, "Select KdAsal,NamaAsal From AsalBarang Where StatusEnabled=1")

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub subLoadText()
    Dim i As Integer
    txtIsi.Left = fgData.Left

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
    s_DcName.Height = fgData.RowHeight(fgData.Row)
    s_DcName.Visible = True
    s_DcName.SetFocus
End Sub

Private Sub subLoadComboBox(s_CbName As Object)
    Dim i As Integer
    s_CbName.Left = fgData.Left
    
    For i = 0 To fgData.Col - 1
        s_CbName.Left = s_CbName.Left + fgData.ColWidth(i)
    Next i
    s_CbName.Visible = True
    s_CbName.Top = fgData.Top - 7

    For i = 0 To fgData.Row - 1
        s_CbName.Top = s_CbName.Top + fgData.RowHeight(i)
    Next i

    If fgData.TopRow > 1 Then
        s_CbName.Top = s_CbName.Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
    End If

    s_CbName.Width = fgData.ColWidth(fgData.Col)
    s_CbName.Visible = True
    s_CbName.SetFocus
End Sub

Private Sub cbSatuanJml_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 27 Then
        If cbSatuanJml.Text = "" Then Exit Sub
        fgData.TextMatrix(fgData.Row, 3) = cbSatuanJml.Text
        cbSatuanJml.Visible = False
        fgData.Col = 4
        fgData.SetFocus
    End If
End Sub

Private Sub cbSatuanJml_LostFocus()
    cbSatuanJml.Visible = False
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell
    Dim i As Integer

    If fgData.TextMatrix(1, 0) = "" Then MsgBox "Data barang harus diisi", vbExclamation, "Validasi": fgData.SetFocus: fgData.Col = 0: Exit Sub

    For i = 1 To fgData.Rows - 2
        With fgData
            If .TextMatrix(i, 5) = "" Then
                MsgBox "Silahkan isi bentuk darah ", vbExclamation, "Validasi"
                .SetFocus: .Row = i: .Col = 0
                Exit Sub
            End If

            If .TextMatrix(i, 6) = "" Then
                MsgBox "Silahkan isi golongan darah ", vbExclamation, "Validasi"
                .SetFocus: .Row = i: .Col = 2
                Exit Sub
            End If
            If .TextMatrix(i, 7) = "" Then
                MsgBox "Silahkan isi asal darah ", vbExclamation, "Validasi"
                .SetFocus: .Row = i: .Col = 2
                Exit Sub
            End If
            If .TextMatrix(i, 3) = "" Then
                MsgBox "Silahkan isi satuan jumlah", vbExclamation, "Validasi"
                .SetFocus: .Row = i: .Col = 3
                Exit Sub
            End If
        End With
    Next i

    For i = 1 To fgData.Rows - 1
        With fgData
            If .TextMatrix(i, 0) = "" Then Exit For
            If SP_AUDDetailPemakaianDarah(.TextMatrix(i, 5), .TextMatrix(i, 6), .TextMatrix(i, 7), .TextMatrix(i, 3), .TextMatrix(i, 4), "A") = False Then Exit Sub
        End With
    Next i

    MsgBox "Data berhasil disimpan", vbInformation, "Sukses"
    Call subSetGrid

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    strKodePelayananRS = ""
    mdTglBerlaku = Now
    frmTindakan.Visible = True
    With frmTindakan
        .txtNamaPelayanan.Text = ""
        .txtKuantitas.Text = 1
        .fraPelayanan.Visible = False
        .txtNamaPelayanan.SetFocus
'       .chkPerawat.SetFocus
    End With
    Unload Me
End Sub

Private Sub dcAsalDarah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 27 Then
        If dcAsalDarah.BoundText = "" Then Exit Sub
        fgData.TextMatrix(fgData.Row, 2) = dcAsalDarah.Text
        fgData.TextMatrix(fgData.Row, 7) = dcAsalDarah.BoundText
        dcAsalDarah.Visible = False
        fgData.Col = 3
        fgData.SetFocus
    End If
End Sub

Private Sub dcAsalDarah_LostFocus()
    dcAsalDarah.Visible = False
End Sub

Private Sub dcBentukDarah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 27 Then
        If dcBentukDarah.BoundText = "" Then Exit Sub
        fgData.TextMatrix(fgData.Row, 0) = dcBentukDarah.Text
        fgData.TextMatrix(fgData.Row, 5) = dcBentukDarah.BoundText
        dcBentukDarah.Visible = False
        fgData.Col = 1
        fgData.SetFocus
    End If
End Sub

Private Sub dcBentukDarah_LostFocus()
    dcBentukDarah.Visible = False
End Sub

Private Sub dcGolDarah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 27 Then
        If dcGolDarah.BoundText = "" Then Exit Sub
        fgData.TextMatrix(fgData.Row, 1) = dcGolDarah.Text
        fgData.TextMatrix(fgData.Row, 6) = dcGolDarah.BoundText
        dcGolDarah.Visible = False
        fgData.Col = 2
        fgData.SetFocus
    End If
End Sub

Private Sub dcGolDarah_LostFocus()
    dcGolDarah.Visible = False
End Sub

Private Sub fgData_DblClick()
    Call fgData_KeyDown(13, 0)
End Sub

Private Sub fgData_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo hell
    Select Case KeyCode
        Case 13
            If fgData.Col = fgData.Cols - 1 Then
                If fgData.TextMatrix(fgData.Row, 2) <> "" Then
                    If fgData.TextMatrix(fgData.Rows - 1, 2) <> "" Then fgData.Rows = fgData.Rows + 1
                    fgData.Row = fgData.Rows - 1
                    fgData.Col = 0
                Else
                    fgData.Col = 0
                End If
            Else
                For i = 0 To fgData.Cols - 2
                    If fgData.Col = fgData.Cols - 1 Then Exit For
                    fgData.Col = fgData.Col + 1
                    If fgData.ColWidth(fgData.Col) > 0 Then Exit For
                Next i
            End If
            fgData.SetFocus

        Case 27
            txtIsi.Visible = False
            dcBentukDarah.Visible = False
            dcGolDarah.Visible = False
            dcAsalDarah.Visible = False

        Case vbKeyDelete
            With fgData
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

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub fgData_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    txtIsi.Text = ""
    If Not (KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ Or KeyAscii = 32 Or KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace Or KeyAscii = Asc(".")) Then
        KeyAscii = 0
        Exit Sub
    End If

    Select Case fgData.Col
        Case 0 'Bentuk Darah
            fgData.Col = 0
            Call subLoadDataCombo(dcBentukDarah)

        Case 1 'Gol. Darah
            fgData.Col = 1
            Call subLoadDataCombo(dcGolDarah)

        Case 2 'Asal Darah
            fgData.Col = 2
            Call subLoadDataCombo(dcAsalDarah)

        Case 3 'Satuan
            fgData.Col = 3
            Call subLoadComboBox(cbSatuanJml)

        Case 4 'Qty Darah
            Call SetKeyPressToNumber(KeyAscii)
            Call subLoadText
            txtIsi.Text = Chr(KeyAscii)
            txtIsi.SelStart = Len(txtIsi.Text)

    End Select

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call subSetGrid
    Call subLoadDcSource
End Sub

Private Function SP_AUDDetailPemakaianDarah(f_KdBentukDarah As Integer, f_KdGolonganDarah As String, f_KdAsalDarah As String, f_SatuanJml As String, f_JmlDarah As Double, f_status As String) As Boolean
    On Error GoTo hell
    SP_AUDDetailPemakaianDarah = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, "117053")
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(mdTglBerlaku, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdBentukDarah", adTinyInt, adParamInput, , f_KdBentukDarah)
        .Parameters.Append .CreateParameter("KdGolonganDarah", adChar, adParamInput, 2, f_KdGolonganDarah)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsalDarah)
        .Parameters.Append .CreateParameter("SatuanJml", adChar, adParamInput, 1, f_SatuanJml)
        .Parameters.Append .CreateParameter("JmlDarah", adDouble, adParamInput, 1, f_JmlDarah)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_DetailPemakaianDarah"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data ", vbCritical, "Validasi"
            SP_AUDDetailPemakaianDarah = False
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Exit Function
hell:
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    SP_AUDDetailPemakaianDarah = False
    Call msubPesanError
End Function

Private Sub Form_Unload(Cancel As Integer)
    Call cmdTutup_Click
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    Dim i As Integer
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
        Select Case fgData.Col
            Case 4
                fgData.TextMatrix(fgData.Row, fgData.Col) = txtIsi.Text
                fgData.SetFocus: fgData.Col = 4
        End Select
    End If

    If fgData.Col = 4 Then
        If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace Or KeyAscii = Asc(",")) Then KeyAscii = 0
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtIsi_LostFocus()
    txtIsi.Visible = False
End Sub
