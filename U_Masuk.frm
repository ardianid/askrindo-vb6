VERSION 5.00
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form U_Masuk 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "U_Masuk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frm_Pwd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Conection"
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton Cmd_Cancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   4440
         TabIndex        =   16
         Top             =   1800
         Width           =   735
      End
      Begin VB.CommandButton Cmd_Browse 
         Caption         =   "Command1"
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   2760
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Cmd_Ok 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   3600
         TabIndex        =   13
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Txt_Nama 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox Txt_Pwd 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   12
         Top             =   480
         Width           =   570
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   630
         TabIndex        =   14
         Top             =   960
         Width           =   900
      End
      Begin VB.Image Image1 
         Height          =   1575
         Left            =   240
         Picture         =   "U_Masuk.frx":08CA
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1215
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Daftar 
      Height          =   5055
      Left            =   1320
      TabIndex        =   0
      Top             =   6480
      Visible         =   0   'False
      Width           =   5190
      _Version        =   65536
      _ExtentX        =   9155
      _ExtentY        =   8916
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "U_Masuk.frx":4137
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "U_Masuk.frx":4153
      Childs          =   "U_Masuk.frx":41FF
      Begin VB.TextBox Txt_Cr_Daftar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   840
         TabIndex        =   2
         Top             =   720
         Width           =   4050
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   2
         Left            =   255
         TabIndex        =   1
         Top             =   480
         Width           =   4605
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Daftar 
         Height          =   3585
         Left            =   240
         OleObjectBlob   =   "U_Masuk.frx":421B
         TabIndex        =   4
         Top             =   1200
         Width           =   4665
      End
      Begin VB.TextBox Txt_Cr_Daftar 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   3
         Top             =   1800
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   14
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   15
         Left            =   480
         TabIndex        =   6
         Top             =   1800
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN USER"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   20
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1605
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D tdb_pwd 
      Height          =   3135
      Left            =   -1200
      TabIndex        =   8
      Top             =   6960
      Width           =   7455
      _Version        =   65536
      _ExtentX        =   13150
      _ExtentY        =   5530
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "U_Masuk.frx":6C9F
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "U_Masuk.frx":6CBB
      Childs          =   "U_Masuk.frx":6D67
   End
End
Attribute VB_Name = "U_Masuk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_daftar As New XArrayDB
Dim jumlah  As Integer
Dim kode_karyawan As String
Dim Moving As Boolean
Dim yold, xold As Long

Private Sub Cmd_Browse_Click()
    
    With TDB_Daftar
                
        .Left = tdb_pwd.Left + Frm_Pwd.Left + Cmd_Browse.Left + Cmd_Browse.Width / 2 - .Width / 2
        .Top = tdb_pwd.Top + Frm_Pwd.Top + Cmd_Browse.Top + Cmd_Browse.Height
                
        If .Visible = False Then
            
            Txt_Cr_Daftar(1).Text = ""
            
            txt_cr_daftar_KeyUp 0, 0, 0
            
            .Visible = True
            
             Txt_Cr_Daftar(1).SetFocus
            
        Else
            .Visible = False
        End If
        
    End With
    
End Sub

Private Sub Cmd_Browse_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Cmd_Cancel_Click
End Sub

Private Sub Cmd_Cancel_Click()
    
    If MsgBox("Yakin akan keluar dari program", vbYesNo + vbQuestion, "Konfirmasi") = vbYes Then
        
        If kon.State = adStateOpen Then
            kon.Close
            Set kon = Nothing
            End
        End If
        
    End If
    
End Sub

Private Sub Cmd_Ok_Click()
On Error GoTo er_handler

    Dim konfirm As Integer
    
    If Txt_Nama.Text = "" Then
        
        konfirm = CInt(MsgBox("Nama user tidak boleh kosong", vbOKOnly + vbInformation, "Informasi"))
        
        Txt_Nama.SetFocus
        On Error GoTo 0
        Exit Sub
    End If
    
    If Txt_Pwd.Text = "" Then
        
        konfirm = CInt(MsgBox("Password tidak boleh kosong", vbOKOnly + vbInformation, "Informasi"))
        
        Txt_Pwd.SetFocus
        On Error GoTo 0
        Exit Sub
        
    End If
        
    Dim sql As String
    Dim rs As Recordset
        
            
            sql = "select Kode_Karyawan,Nama_Karyawan,Status from VIEW_User where Nama_Karyawan='" & Trim(Txt_Nama.Text) & "'"
            
            Set rs = New ADODB.Recordset
                rs.Open sql, kon
                
                With rs
                    
                    If Not .EOF Then
                        
                        If !status = 0 Then
                            
                            konfirm = CInt(MsgBox("Anda sudah tidak diberikan hak untuk menggunakan program ini, hubungi administrator", vbOKOnly + vbInformation, "Informasi"))
                            
                            On Error GoTo 0
                            Exit Sub
                            
                        End If
                        
                        kode_karyawan = IIf(Not IsNull(!kode_karyawan), !kode_karyawan, "")
                    Else
                    
                            konfirm = CInt(MsgBox("Nama anda tidak ditemukan dalam otoritas pemakai program ini", vbOKOnly + vbInformation, "Informasi"))
                            
                        Txt_Nama.SetFocus
                        
                        On Error GoTo 0
                        Exit Sub
                    End If
                    
                End With
    

        sql = "select Id_User,Nama_Karyawan,Pwd from VIEW_User where Kode_Karyawan ='" & kode_karyawan & "'" ' and Pwd ='" & encrypt_pwd(Trim(Txt_Pwd.Text)) & "'"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon
            
            With rs
                
                If Not .EOF Then
                    
                    If IsNull(!pwd) Then
                        konfirm = CInt(MsgBox("Password anda kosong", vbOKOnly + vbInformation, "Informasi"))
                        
                        On Error GoTo 0
                        Exit Sub
                    Else
                            
                        Dim decrypt As String
                            decrypt = decrypt_pwd(!pwd)
                            
                            If UCase(decrypt) <> UCase(Trim(Txt_Pwd.Text)) Then
                            
                                konfirm = CInt(MsgBox("Password yang anda masukkan salah", vbOKOnly + vbInformation, "Informasi"))
                    
                                jumlah = jumlah + 1
                                
                                If jumlah = 3 Then
                                    kon.Close
                                    Set kon = Nothing
                                    End
                                    
                                    On Error GoTo 0
                                    Exit Sub
                                    
                                    
                                End If
                                
                                Txt_Pwd.SetFocus
                                     
                                On Error GoTo 0
                                Exit Sub
                                     
                            End If
                    
                    End If
                    
                    Dim Id_User As String
                    Dim Nama_Kar As String
                        Id_User = !Id_User
                        Id_User = Trim(Id_User)
                        Flag_tempat = Id_User
                        Nama_Kar = IIf(Not IsNull(!Nama_Karyawan), !Nama_Karyawan, "")
                        
                        
                    With Utama
                        .SetAktifMenu "select nama_form from VIEW_Hak_Akses where id_user=" & Id_User
                        .StatusBar1.Panels(1).Text = "User Actived : " & Nama_Kar
                        .logof.Enabled = True
                        .login.Enabled = False
                    End With
                    
                    Unload Me
                    
                    jumlah = 0
                    
                    If kon.State = adStateOpen Then
                        kon.Close
                        Set kon = Nothing
                    End If
                    
'                    If remind = True Then
'
'                        If Not (Frm Is Nothing) Then
'                            Unload Frm
'                            Set Frm = Nothing
'                            Set Frm = Frm_Reminder
'                            Frm.Show
'                        Else
'                            Set Frm = Frm_Reminder
'                            Frm.Show
'                        End If
'
'                    End If
                    
                    
                Else
                    
                    konfirm = CInt(MsgBox("Password yang anda masukkan salah", vbOKOnly + vbInformation, "Informasi"))
                    
                    
                    
                    jumlah = jumlah + 1
                    
                    If jumlah = 3 Then
                        kon.Close
                        Set kon = Nothing
                        End
                    End If
                    
                    Txt_Pwd.SetFocus
                    
                End If
                
            End With
        
        
        
        On Error GoTo 0
        Exit Sub

er_handler:
    
    konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
        Err.Clear
    

End Sub


Private Sub Form_Activate()
    On Error Resume Next
        Txt_Nama.SetFocus
End Sub

Private Sub Form_Load()
    
    Me.Height = 2790
    Me.Width = 5640
    
    
    
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = (Screen.Height / 2 - Me.Height / 2) - 1500

    Dim status As String
    status = Buka_Koneksi
    If status = "-2147467259" Then
        Dim konfirm As Integer
            konfirm = CInt(MsgBox("Koneksi terputus ....", vbOKOnly + vbInformation, "Informasi"))
            
            End
            Exit Sub
    End If

'    SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / 2, _
'                        Me.Top / 2, Me.Width / 2, _
'                      Me.Height / 2, SWP_NOACTIVATE Or SWP_SHOWWINDOW

    
'    Me.PaintPicture Utama.Picture, 0, 0
    jumlah = 0
    
'    Grid_Daftar.Array = arr_daftar
    
        
'    With TDB_Daftar
'        .Left = Frm_Pwd.Left + Cmd_Browse.Left + Cmd_Browse.Width / 2 - .Width + 750
'        .Top = Frm_Pwd.Top + Cmd_Browse.Top + Cmd_Browse.Height + 50
'    End With
        
'    Frm_Pwd.BackColor = RGB(187, 223, 249)
'    Txt_Pwd.BackColor = RGB(187, 223, 249)
'    Txt_Nama.BackColor = RGB(187, 223, 249)
    
End Sub

Private Sub Form_Resize()

    
'    With tdb_pwd
'        .Left = Me.Width / 2 - .Width / 2
'        .Top = 1000
'    End With
'
'    Image1.Move 0, 0, Frm_Pwd.Width - 0, Frm_Pwd.Height - 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If kon.State = adStateOpen Then
        kon.Close
        Set kon = Nothing
    End If
    
End Sub

Private Sub grid_daftar_DblClick()
    
    If arr_daftar.UpperBound(1) = 1 And arr_daftar(1, 1) = Empty Then Exit Sub
    
    Txt_Nama.Text = arr_daftar(Grid_Daftar.Bookmark, 1)
    kode_karyawan = arr_daftar(Grid_Daftar.Bookmark, 0)
    
    TDB_Daftar.Visible = False
    Txt_Pwd.SetFocus
    
End Sub

Private Sub grid_daftar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then grid_daftar_DblClick
    If KeyCode = vbKeyEscape Then TDB_Daftar.Visible = False: Cmd_Browse.SetFocus
End Sub

Private Sub TDB_Daftar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub TDB_Daftar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   TDB_Daftar.Top = TDB_Daftar.Top - (yold - Y)
   TDB_Daftar.Left = TDB_Daftar.Left - (xold - X)
End If

End Sub

Private Sub TDB_Daftar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub

Private Sub tdb_pwd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub tdb_pwd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   tdb_pwd.Top = tdb_pwd.Top - (yold - Y)
   tdb_pwd.Left = tdb_pwd.Left - (xold - X)
End If

End Sub

Private Sub tdb_pwd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub

Private Sub txt_cr_daftar_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Daftar.SetFocus
    If KeyCode = vbKeyEscape Then TDB_Daftar.Visible = False: Cmd_Browse.SetFocus
End Sub

Private Sub txt_cr_daftar_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    Dim sql As String
    Dim rs As Recordset
        
        arr_daftar.ReDim 0, 0, 0, 0
        arr_daftar.ReDim 1, 1, 1, 1
            Grid_Daftar.ReBind
            Grid_Daftar.Refresh
        
        sql = "select Kode_Karyawan,Nama_Karyawan from VIEW_User"
        
        Select Case Index
            Case 0
                sql = sql & " where Kode_Karyawan like '%" & Trim(Txt_Cr_Daftar(0).Text) & "%'"
            Case 1
                sql = sql & " where Nama_Karyawan like '%" & Trim(Txt_Cr_Daftar(1).Text) & "%'"
        End Select
        
        sql = sql & " order by Kode_Karyawan,Nama_Karyawan asc"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon, adOpenKeyset
            
            Dim a As Long
            Dim kode, NAMA As String
            
            a = 1
            
            With rs
            
                Do While Not .EOF
                    arr_daftar.ReDim 1, a, 0, Grid_Daftar.Columns.Count
                        Grid_Daftar.ReBind
                        Grid_Daftar.Refresh
                        
                        kode = IIf(Not IsNull(!kode_karyawan), !kode_karyawan, "")
                        NAMA = IIf(Not IsNull(!Nama_Karyawan), !Nama_Karyawan, "")
                    
                    arr_daftar(a, 0) = kode
                    arr_daftar(a, 1) = NAMA
                    
                a = a + 1
                .MoveNext
                Loop
                
                Grid_Daftar.ReBind
                Grid_Daftar.Refresh
                
                Grid_Daftar.MoveFirst
            
            End With
            
        
End Sub

Private Sub txt_nama_GotFocus()
    Call Focus_(Txt_Nama)
End Sub

Private Sub Txt_Nama_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyEscape Then Cmd_Cancel_Click
'    If KeyCode = vbKeyF3 Then Cmd_Browse_Click
    If KeyCode = 13 Then
        
        Dim sql As String
        Dim rs As Recordset
            sql = "select Kode_Karyawan,Nama_Karyawan from VIEW_User where Nama_Karyawan='" & Trim(Txt_Nama.Text) & "'"
            
            Set rs = New ADODB.Recordset
                rs.Open sql, kon
                
                With rs
                    
                    If Not .EOF Then
                        
                        kode_karyawan = IIf(Not IsNull(!kode_karyawan), !kode_karyawan, "")
                        
                        Txt_Pwd.SetFocus
                        
                    Else
                    
                        Dim konfirm As Integer
                            konfirm = CInt(MsgBox("Nama anda tidak ditemukan dalam otoritas pemakai program ini", vbOKOnly + vbInformation, "Informasi"))
                            
                        
                    End If
                    
                End With
                
        
    End If
    
End Sub

Private Sub Txt_Pwd_GotFocus()
    Call Focus_(Txt_Pwd)
End Sub

Private Sub Txt_Pwd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_OK.SetFocus
    If KeyCode = vbKeyEscape Then Cmd_Cancel_Click
End Sub
