VERSION 5.00
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form Frm_Rubah_Pwd 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RUBAH PASSWORD"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_Rubah_Pwd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Karyawan 
      Height          =   3735
      Left            =   -5640
      TabIndex        =   19
      Top             =   4440
      Visible         =   0   'False
      Width           =   6255
      _Version        =   65536
      _ExtentX        =   11033
      _ExtentY        =   6588
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Frm_Rubah_Pwd.frx":27C92
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Frm_Rubah_Pwd.frx":27CAE
      Childs          =   "Frm_Rubah_Pwd.frx":27D5A
      Begin VB.TextBox Txt_Cr_Kar 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   1
         Left            =   3720
         TabIndex        =   22
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox Txt_Cr_Kar 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   0
         Left            =   1080
         TabIndex        =   21
         Top             =   840
         Width           =   1215
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   5655
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Karyawan 
         Height          =   2295
         Left            =   240
         OleObjectBlob   =   "Frm_Rubah_Pwd.frx":27D76
         TabIndex        =   23
         Top             =   1200
         Width           =   5655
      End
      Begin VB.Label Label2 
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
         Index           =   15
         Left            =   3000
         TabIndex        =   26
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label2 
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
         Index           =   14
         Left            =   480
         TabIndex        =   25
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN DATA KARYAWAN"
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
         TabIndex        =   24
         Top             =   240
         Width           =   2760
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   16
      Top             =   2040
      Width           =   2055
      Begin VB.CommandButton Cmd_Keluar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Keluar"
         Height          =   375
         Left            =   1080
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Cmd_Simpan 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "DATA KARYAWAN"
      Height          =   1575
      Left            =   -4680
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   7695
      Begin VB.TextBox Txt_Kode_Karyawan 
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
         Left            =   2400
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Cmd_Browse_Karyawan 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Txt_Nama_Karyawan 
         Enabled         =   0   'False
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
         Left            =   2400
         TabIndex        =   7
         Top             =   840
         Width           =   5055
      End
      Begin VB.TextBox Txt_Jabatan 
         Enabled         =   0   'False
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
         Left            =   2400
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   5055
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
         Index           =   0
         Left            =   360
         TabIndex        =   15
         Top             =   480
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Index           =   1
         Left            =   2160
         TabIndex        =   14
         Top             =   480
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama "
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
         Index           =   2
         Left            =   360
         TabIndex        =   13
         Top             =   840
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Index           =   3
         Left            =   2160
         TabIndex        =   12
         Top             =   840
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan"
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
         Index           =   4
         Left            =   360
         TabIndex        =   11
         Top             =   1200
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Index           =   5
         Left            =   2160
         TabIndex        =   10
         Top             =   1200
         Visible         =   0   'False
         Width           =   60
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Width           =   4455
      Begin VB.TextBox Txt_Password 
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Txt_Verifikasi 
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Password :"
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   4
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Verification Password :"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   1635
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1800
      TabIndex        =   27
      Top             =   0
      Width           =   4455
      Begin VB.TextBox Txt_Pwd_Lama 
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   28
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password :"
         Height          =   195
         Index           =   1
         Left            =   870
         TabIndex        =   29
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   120
      Picture         =   "Frm_Rubah_Pwd.frx":2B1A8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "Frm_Rubah_Pwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim yold, xold As Long
Dim Moving As Boolean
Dim Arr_Karyawan As New XArrayDB

Private Sub Isi_Grid(ByVal grid As Object, ByVal arr As Object, ByVal rec As Recordset)
    
    Dim a As Long
    Dim kode, NAMA, jab As String
        
        With rec
            
            a = 1
                
                arr.ReDim 0, 0, 0, 0
                arr.ReDim 1, 1, 1, 1
                    grid.ReBind
                    grid.Refresh
                    
                    Do While Not .EOF
                        arr.ReDim 1, a, 0, grid.Columns.Count
                            grid.ReBind
                            grid.Refresh
                            
                            kode = IIf(Not IsNull(!kode_karyawan), !kode_karyawan, "")
                            NAMA = IIf(Not IsNull(!Nama_Karyawan), !Nama_Karyawan, "")
'                            jab = IIf(Not IsNull(!jabatan), !jabatan, "")
                            
                            arr(a, 0) = kode
                            arr(a, 1) = NAMA
                            arr(a, 2) = ""
                        
                    a = a + 1
                    .MoveNext
                    Loop
                    
                    grid.ReBind
                    grid.Refresh
                    
                    grid.MoveFirst
        
        End With
        
End Sub


Private Sub Cmd_Browse_Karyawan_Click()

With TDB_Karyawan
    
    If .Visible = False Then
        
        Txt_Cr_Kar(0).Text = ""
        Txt_Cr_Kar(1).Text = ""
        
        Txt_Cr_Kar_KeyUp 0, 0, 0
        
        .Visible = True
        
        Txt_Cr_Kar(0).SetFocus
    
    Else
        .Visible = False
    End If
    
End With

End Sub

Private Sub Cmd_Keluar_Click()
    Unload Me
End Sub

Private Sub Cmd_Simpan_Click()
On Error GoTo err_handler

    Dim sql As String
    Dim rs As Recordset
    Dim sql1 As String
    Dim rs1 As Recordset
    Dim konfirm As Integer
    
    If Txt_Kode_Karyawan.Text = "" Then
        
        konfirm = CInt(MsgBox("Kode Karyawan tidak boleh kosong", vbOKOnly + vbInformation, "Informasi"))
        
        Txt_Kode_Karyawan.SetFocus
        On Error GoTo 0
        Exit Sub
    End If
    
    If Txt_Pwd_Lama.Text = "" Then
        
        konfirm = CInt(MsgBox("Password lama tidak boleh kosong", vbOKOnly + vbInformation, "Informasi"))
        
        Txt_Pwd_Lama.SetFocus
        On Error GoTo 0
        Exit Sub
    End If
    
    If Txt_Password.Text = "" Then
        
        konfirm = CInt(MsgBox("Password tidak boleh kosong", vbOKOnly + vbInformation, "Informasi"))
        
        Txt_Password.SetFocus
        On Error GoTo 0
        Exit Sub
    End If
    
    If Txt_Verifikasi.Text = "" Then
        
        konfirm = CInt(MsgBox("Verifikasi password tidak boleh kosong", vbOKOnly + vbInformation, "Informasi"))
        
        Txt_Verifikasi.SetFocus
        On Error GoTo 0
        Exit Sub
    End If

        If Trim(Txt_Password.Text) <> Trim(Txt_Verifikasi.Text) Then
            
            
                konfirm = CInt(MsgBox("Password dan Verifikasi Password tidak sama", vbOKOnly + vbInformation, "Informasi"))
                
                Txt_Verifikasi.SetFocus
            
            On Error GoTo 0
            Exit Sub
            
        End If
    
    sql1 = "select Kode_Karyawan from Tb_User where Kode_Karyawan='" & Trim(Txt_Kode_Karyawan.Text) & "'"
            
        Set rs1 = New ADODB.Recordset
            rs1.Open sql1, kon
            
            With rs1
                
                If .EOF Then
                    
                    konfirm = CInt(MsgBox("Data Karyawan tidak ditemukan dalam daftar user", vbOKOnly + vbInformation, "Informasi"))
                    
                    Txt_Kode_Karyawan.SetFocus
                    On Error GoTo 0
                    Exit Sub
                    
                End If
                
            End With
    
    sql1 = "select Pwd from Tb_User where Kode_Karyawan='" & Trim(Txt_Kode_Karyawan.Text) & "'"
        
        Set rs1 = New ADODB.Recordset
            rs1.Open sql1, kon
            
            With rs1
                
                If .EOF Then
                    
                    konfirm = CInt(MsgBox("Password lama tidak ditemukan dalam daftar user", vbOKOnly + vbInformation, "Informasi"))
                    
                    Txt_Pwd_Lama.SetFocus
                    On Error GoTo 0
                    Exit Sub
                    
                Else
                    
                    If Trim(Txt_Pwd_Lama.Text) <> decrypt_pwd(Trim(!pwd)) Then
                    
                        konfirm = CInt(MsgBox("Password lama tidak sama dengan password terdahulu", vbOKOnly + vbInformation, "Informasi"))
                    
                        Txt_Pwd_Lama.SetFocus
                        On Error GoTo 0
                        Exit Sub
                    
                    End If
                    
                End If
                
            End With
    
    kon.BeginTrans
    
    sql = "update Tb_User set Pwd='" & encrypt_pwd(Trim(Txt_Password.Text)) & "' where Kode_Karyawan='" & Trim(Txt_Kode_Karyawan.Text) & "'"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon
    
    kon.CommitTrans
    
        konfirm = CInt(MsgBox("Password user telah dirubah", vbOKOnly + vbInformation, "Informasi"))
        
'        Txt_Kode_Karyawan.Text = ""
'        Txt_Nama_Karyawan.Text = ""
'        Txt_Jabatan.Text = ""
        
'        Txt_Pwd_Lama.Text = ""
'        Txt_Pwd_Lama.Enabled = False
'        Txt_Password.Text = ""
'        Txt_Password.Enabled = False
'        Txt_Verifikasi.Text = ""
'        Txt_Verifikasi.Enabled = False
        
'        Txt_Kode_Karyawan.SetFocus
        
        On Error GoTo 0
        Exit Sub

err_handler:
    
    kon.RollbackTrans
    
    konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Informasi"))
        Err.Clear

End Sub

Private Sub Form_Activate()
    On Error Resume Next
        Txt_Pwd_Lama.SetFocus
End Sub

Private Sub Form_Load()
    
    Me.Left = Utama.Width / 2 - Me.Width / 2
    Me.Top = (Utama.Height / 2 - Me.Height / 2) - 1500

Dim status As String
status = Buka_Koneksi
If status = "-2147467259" Then
    Dim konfirm As Integer
        konfirm = CInt(MsgBox("Koneksi terputus ....", vbOKOnly + vbInformation, "Informasi"))
        
        End
        Exit Sub
End If
        
    With Me
        .Left = Screen.Width / 2 - .Width / 2
        .Top = 750
    End With
     
    If cari_karyawan_aktif = False Then Unload Me
                    
    With TDB_Karyawan
        .Left = 360
        .Top = 840
    End With
                        
    Grid_Karyawan.Array = Arr_Karyawan
                    
End Sub

Private Function cari_karyawan_aktif() As Boolean
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select kode_karyawan from VIEW_User where id_user=" & Flag_tempat
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon, adOpenKeyset
        
        With rs
            If Not .EOF Then
                Txt_Kode_Karyawan.Text = !kode_karyawan
                    cari_karyawan_aktif = True
            Else
                    cari_karyawan_aktif = False
            End If
        End With
    
End Function



Private Sub Form_Unload(Cancel As Integer)

If kon.State = adStateOpen Then
    
    kon.Close
    Set kon = Nothing
End If

End Sub

Private Sub Grid_Karyawan_DblClick()

    If Arr_Karyawan.UpperBound(1) = 1 And Arr_Karyawan(1, 1) = Empty Then Exit Sub
    
    Txt_Kode_Karyawan = Arr_Karyawan(Grid_Karyawan.Bookmark, 0)
    Txt_Nama_Karyawan = Arr_Karyawan(Grid_Karyawan.Bookmark, 1)
'    Txt_Jabatan = Arr_Karyawan(Grid_Karyawan.Bookmark, 2)
    
    Txt_Password.Enabled = True
    Txt_Verifikasi.Enabled = True
    Txt_Pwd_Lama.Enabled = True
    
    TDB_Karyawan.Visible = False
    Txt_Pwd_Lama.SetFocus

End Sub

Private Sub Grid_Karyawan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Karyawan_DblClick
    If KeyCode = vbKeyEscape Then TDB_Karyawan.Visible = False: Cmd_Browse_Karyawan.SetFocus
End Sub

Private Sub TDB_Karyawan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub TDB_Karyawan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   TDB_Karyawan.Top = TDB_Karyawan.Top - (yold - Y)
   TDB_Karyawan.Left = TDB_Karyawan.Left - (xold - X)
End If

End Sub

Private Sub TDB_Karyawan_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub

Private Sub Txt_Cr_Kar_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Karyawan.SetFocus
    If KeyCode = vbKeyEscape Then TDB_Karyawan.Visible = False: Cmd_Browse_Karyawan.SetFocus
End Sub

Private Sub Txt_Cr_Kar_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim sql As String
Dim rs As Recordset
    
    sql = "select Kode_Karyawan,Nama_Karyawan from VIEW_Karyawan"
    
    Select Case Index
        Case 0
            sql = sql & " where Kode_Karyawan like '%" & Trim(Txt_Cr_Kar(0).Text) & "%'"
        Case 1
            sql = sql & " where Nama_Karyawan like '%" & Trim(Txt_Cr_Kar(1).Text) & "%'"
    End Select
    
    sql = sql & " order by Kode_Karyawan asc"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
        
        Isi_Grid Grid_Karyawan, Arr_Karyawan, rs
    
End Sub

Private Sub Txt_Kode_Karyawan_GotFocus()
    Call Focus_(Txt_Kode_Karyawan)
End Sub

Private Sub Txt_Kode_Karyawan_KeyDown(KeyCode As Integer, Shift As Integer)
    
'    If KeyCode = vbKeyF3 Then Cmd_Browse_Karyawan_Click
    If KeyCode = 13 Then
    
        Dim konfirm As Integer
        If Txt_Kode_Karyawan.Text = "" Then
            konfirm = CInt(MsgBox("Kode Karyawan tidak boleh Kosong", vbOKOnly + vbInformation, "Informasi"))
            
            Txt_Kode_Karyawan.SetFocus
            Exit Sub
        End If
        
        Dim sql As String
        Dim rs As Recordset
                
        Txt_Nama_Karyawan.Text = ""
        Txt_Jabatan.Text = ""
        Txt_Password.Text = ""
        Txt_Verifikasi.Text = ""
        Txt_Pwd_Lama.Text = ""
                
            sql = "select Kode_Karyawan,Nama_Karyawan from VIEW_User where Kode_Karyawan ='" & Trim(Txt_Kode_Karyawan.Text) & "'"
            
            Set rs = New ADODB.Recordset
                rs.Open sql, kon
                        
            With rs
                
                If Not .EOF Then
                    
                    Txt_Nama_Karyawan.Text = IIf(Not IsNull(!Nama_Karyawan), !Nama_Karyawan, "")
'                    Txt_Jabatan.Text = IIf(Not IsNull(!jabatan), !jabatan, "")
                               
                    Txt_Password.Enabled = True
                    Txt_Verifikasi.Enabled = True
                    Txt_Pwd_Lama.Enabled = True
                    Txt_Pwd_Lama.SetFocus
                    
               
                    
                Else
                    
                    konfirm = CInt(MsgBox("Kode karyawan tidak ditemukan", vbOKOnly + vbInformation, "Information"))
                    
                    
                    Txt_Password.Enabled = False
                    Txt_Verifikasi.Enabled = False
                    Txt_Pwd_Lama.Enabled = False
                    
                End If
                
            End With
    End If
End Sub

Private Sub Txt_Password_GotFocus()
    Call Focus_(Txt_Password)
End Sub

Private Sub Txt_Password_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Verifikasi.SetFocus
End Sub

Private Sub Txt_Pwd_Lama_GotFocus()
    Call Focus_(Txt_Pwd_Lama)
End Sub

Private Sub Txt_Pwd_Lama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Password.SetFocus
End Sub

Private Sub Txt_Verifikasi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Simpan.SetFocus
End Sub

Private Sub Txt_Verifikasi_LostFocus()
    
    If Txt_Verifikasi.Text <> "" Then
        
        If Trim(Txt_Password.Text) <> Trim(Txt_Verifikasi.Text) Then
            
            Dim konfirm As Integer
                konfirm = CInt(MsgBox("Password dan Verifikasi Password tidak sama", vbOKOnly + vbInformation, "Informasi"))
                
                Txt_Verifikasi.SetFocus
            
        End If
        
    End If
    
End Sub
