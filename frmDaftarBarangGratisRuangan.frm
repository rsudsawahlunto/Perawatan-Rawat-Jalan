VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmDaftarBarangGratisRuangan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Daftar Barang Gratis Yang Keluar"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarBarangGratisRuangan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   8925
   Begin VB.TextBox txtNoStruk 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtNoRetur 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   120
      MaxLength       =   15
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtIsi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2160
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   5880
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid fgData 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   7646
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   0
      Appearance      =   0
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   6
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
      Left            =   7080
      Picture         =   "frmDaftarBarangGratisRuangan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarBarangGratisRuangan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarBarangGratisRuangan.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10680
      Y1              =   5760
      Y2              =   5760
   End
End
Attribute VB_Name = "frmDaftarBarangGratisRuangan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim substrNomorRetur As String
Dim i As Integer

Private Sub subLoadText()
    txtIsi.Left = fgData.Left
    Select Case fgData.Col
        Case 4
            txtIsi.MaxLength = 10
        Case Else
            Exit Sub
    End Select

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
    txtIsi.Height = fgData.RowHeight(fgData.Row)

    txtIsi.Visible = True
    txtIsi.SelStart = Len(txtIsi.Text)
    txtIsi.SetFocus
End Sub

Private Sub subLoadData()
    strSQL = "SELECT  NamaBarang, AsalBarang, JmlStok, KdRuangan, KdAsal, KdBarang" & _
    " FROM  V_DaftarBarangGratis"
    Call msubRecFO(rs, strSQL)

    For i = 1 To rs.RecordCount
        With fgData
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = rs("NamaBarang")
            .TextMatrix(i, 2) = rs("AsalBarang")
            .TextMatrix(i, 3) = rs("JmlStok")
            .TextMatrix(i, 4) = 0
            .TextMatrix(i, 5) = rs("KdRuangan")
            .TextMatrix(i, 6) = rs("KdBarang")
            .TextMatrix(i, 7) = rs("KdAsal")

            .Rows = .Rows + 1
            rs.MoveNext
        End With
    Next i
    fgData.Row = 1
End Sub

Private Sub subSetGrid()
    With fgData
        .Clear
        .Rows = 2
        .Cols = 8

        .RowHeight(0) = 400
        .TextMatrix(0, 0) = "No"
        .TextMatrix(0, 1) = "Nama Barang"
        .TextMatrix(0, 2) = "Asal Barang"
        .TextMatrix(0, 3) = "Jumlah Stok"
        .TextMatrix(0, 4) = "Jumlah Keluar"
        .TextMatrix(0, 5) = "KdRuangan"
        .TextMatrix(0, 6) = "KdBarang"
        .TextMatrix(0, 7) = "KdAsal"

        .ColWidth(0) = 500
        .ColWidth(1) = 3700
        .ColWidth(2) = 2000
        .ColWidth(3) = 1300
        .ColWidth(4) = 1300
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
    End With
End Sub

Private Sub cmdSimpan_Click()
    Dim i As Integer

    Call txtIsi_KeyPress(13)

    For i = 1 To fgData.Rows - 1
        With fgData
            If Val(.TextMatrix(i, 4)) <> 0 Then If sp_AddBarangGratisApotikRuangan(mstrNoPen, .TextMatrix(i, 5), .TextMatrix(i, 6), .TextMatrix(i, 7), .TextMatrix(i, 4)) = False Then Exit Sub
        End With
    Next i

    cmdSimpan.Enabled = False
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub fgData_DblClick()
    Call fgData_KeyDown(13, 0)
End Sub

Private Sub fgData_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)

    Select Case KeyCode
        Case 13
            If fgData.TextMatrix(fgData.Row, 2) = "" Then Exit Sub
            Call subLoadText
            txtIsi.Text = Trim(fgData.TextMatrix(fgData.Row, fgData.Col))
            txtIsi.SelStart = 0
            txtIsi.SelLength = Len(txtIsi.Text)
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call subSetGrid
    Call subLoadData
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    If KeyAscii = 13 Then
        If Val(txtIsi.Text) = 0 Then txtIsi.Text = 0
        If Val(txtIsi.Text) > Val(fgData.TextMatrix(fgData.Row, 3)) Then
            MsgBox "Jumlah keluar lebih besar dari jumlah stok", vbExclamation, "Validasi"
            txtIsi.SelStart = 0
            txtIsi.SelLength = Len(txtIsi.Text)
            Exit Sub
        End If

        fgData.TextMatrix(fgData.Row, 4) = txtIsi.Text
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
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Function sp_AddBarangGratisApotikRuangan(f_NoPendaftaran As String, f_KdRuangan As String, f_KdBarang As String, f_KdAsal As String, f_JumlahKeluar As String) As Boolean
    On Error GoTo errLoad

    sp_AddBarangGratisApotikRuangan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(Now, "yyyy/mm/dd hh:mm:00"))
        .Parameters.Append .CreateParameter("JmlBarang", adDouble, adParamInput, , f_JumlahKeluar)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_BarangGratisRuangan"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data Retur Pengiriman Barang Dari Ruangan", vbCritical, "Validasi"
            sp_AddBarangGratisApotikRuangan = False
        Else
            Call Add_HistoryLoginActivity("Add_BarangGratisRuangan")
        End If
    End With
    Set dbcmd = Nothing
    Call deleteADOCommandParameters(dbcmd)
    Exit Function
errLoad:
    sp_AddBarangGratisApotikRuangan = False
    Call msubPesanError
End Function

Private Sub txtIsi_LostFocus()
    txtIsi.Visible = False
End Sub

