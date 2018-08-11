VERSION 5.00
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form User_Baru 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New User"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7620
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "User_Baru.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Daftar 
      Height          =   3495
      Left            =   -2400
      TabIndex        =   35
      Top             =   4440
      Visible         =   0   'False
      Width           =   6615
      _Version        =   65536
      _ExtentX        =   11668
      _ExtentY        =   6165
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "User_Baru.frx":27C92
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "User_Baru.frx":27CAE
      Childs          =   "User_Baru.frx":27D5A
      Begin VB.TextBox Txt_Cr_Daftar 
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   42
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox Txt_Cr_Daftar 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   41
         Top             =   600
         Width           =   1215
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Daftar 
         Height          =   2415
         Left            =   240
         OleObjectBlob   =   "User_Baru.frx":27D76
         TabIndex        =   37
         Top             =   960
         Width           =   6135
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
         Height          =   135
         Index           =   0
         Left            =   240
         TabIndex        =   36
         Top             =   360
         Width           =   6135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   195
         Index           =   2
         Left            =   3240
         TabIndex        =   40
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   39
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN DATA KARYAWAN"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   38
         Top             =   120
         Width           =   2220
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Karyawan 
      Height          =   3495
      Left            =   -6000
      TabIndex        =   27
      Top             =   4200
      Visible         =   0   'False
      Width           =   6615
      _Version        =   65536
      _ExtentX        =   11668
      _ExtentY        =   6165
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "User_Baru.frx":2B7BE
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "User_Baru.frx":2B7DA
      Childs          =   "User_Baru.frx":2B886
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
         Height          =   15
         Index           =   2
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   6135
      End
      Begin VB.TextBox Txt_Cr_Kar 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   28
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Txt_Cr_Kar 
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   29
         Top             =   480
         Width           =   2415
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Karyawan 
         Height          =   2535
         Left            =   240
         OleObjectBlob   =   "User_Baru.frx":2B8A2
         TabIndex        =   31
         Top             =   840
         Width           =   6135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN DATA KARYAWAN"
         Height          =   195
         Index           =   20
         Left            =   240
         TabIndex        =   34
         Top             =   120
         Width           =   2220
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         Height          =   195
         Index           =   14
         Left            =   480
         TabIndex        =   33
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   195
         Index           =   15
         Left            =   3240
         TabIndex        =   32
         Top             =   480
         Width           =   405
      End
   End
   Begin VB.Frame Frame4 
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
      Left            =   2760
      TabIndex        =   19
      Top             =   3960
      Width           =   4455
      Begin VB.CommandButton Cmd_Keluar 
         Caption         =   "&Keluar"
         Height          =   375
         Left            =   3480
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Cmd_Daftar 
         Caption         =   "&Daftar"
         Height          =   375
         Left            =   2640
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Cmd_Hapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   1800
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Cmd_Rubah 
         Caption         =   "&Rubah"
         Height          =   375
         Left            =   960
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Cmd_Tambah 
         Caption         =   "&Tambah"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Cmd_Batal 
         Caption         =   "&Batal"
         Height          =   375
         Left            =   960
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Cmd_Simpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame Frame_Nav 
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
      Left            =   360
      TabIndex        =   14
      Top             =   3960
      Width           =   2175
      Begin VB.CommandButton Cmd_Navigasi 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1560
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Cmd_Navigasi 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Cmd_Navigasi 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Cmd_Navigasi 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Password"
      Height          =   1815
      Left            =   360
      TabIndex        =   9
      Top             =   1920
      Width           =   6855
      Begin VB.TextBox Txt_Verifikasi 
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
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox Txt_Password 
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
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   480
         Width           =   2655
      End
      Begin VB.Frame Frame_Stats 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   3360
         TabIndex        =   44
         Top             =   1080
         Width           =   1335
         Begin VB.CheckBox Check_Aktif 
            Alignment       =   1  'Right Justify
            Caption         =   "&Aktif"
            Height          =   375
            Left            =   360
            TabIndex        =   45
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Verifikasi Password :"
         Height          =   195
         Left            =   255
         TabIndex        =   13
         Top             =   840
         Width           =   1470
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         Height          =   195
         Left            =   960
         TabIndex        =   12
         Top             =   480
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Karyawan"
      Height          =   1455
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6855
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
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   5895
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
         Left            =   1800
         TabIndex        =   7
         Top             =   840
         Width           =   3015
      End
      Begin VB.CommandButton Cmd_Browse_Karyawan 
         Caption         =   "..."
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
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
         Left            =   1800
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Lbl_Info 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lbl_Info"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   6105
         TabIndex        =   43
         Top             =   150
         Width           =   585
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
         TabIndex        =   4
         Top             =   1200
         Visible         =   0   'False
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
         TabIndex        =   3
         Top             =   1200
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama :"
         Height          =   195
         Index           =   2
         Left            =   1200
         TabIndex        =   2
         Top             =   840
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode :"
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   1
         Top             =   480
         Width           =   465
      End
   End
End
Attribute VB_Name = "User_Baru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rubah As Boolean
Dim yold, xold As Long
Dim Moving As Boolean
Dim Arr_Karyawan As New XArrayDB
Dim Id_Pwd As String
Dim Tujuan As Long
Dim arr_daftar As New XArrayDB

Private Sub Check_Aktif_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Simpan.SetFocus
End Sub



Private Sub Cmd_Batal_Click()

    Frame_Nav.Enabled = True
    rubah = False
             
        Cmd_Tambah.Visible = True
        Cmd_Tambah.Enabled = True
        Cmd_Simpan.Visible = False
        Cmd_Rubah.Visible = True
        Cmd_Rubah.Enabled = True
        Cmd_Hapus.Enabled = True
        Cmd_Daftar.Enabled = True
        Cmd_Keluar.Enabled = True
        
        Dim n As Object
            
For Each n In Me

        If TypeOf n Is TextBox Then
            If Left(UCase(n.Name), 6) <> UCase("Txt_Cr") Then
                n.Enabled = False
              '  n.Text = ""
            End If
        End If
        
        If TypeOf n Is TDBContainer3D Then n.Visible = False
        
Next

Set n = Nothing

Frame_Stats.Enabled = False
Cmd_Browse_Karyawan.Enabled = False

 If Cmd_Tambah.Enabled = True Then Cmd_Tambah.SetFocus

    txt_cr_daftar_KeyUp 0, 0, 0
        Cmd_Navigasi_Click 3


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

Private Sub Cmd_Daftar_Click()

rubah = False
Tujuan = 2

With TDB_Daftar

If .Visible = False Then
    
    Frame_Nav.Enabled = False
    
    Cmd_Tambah.Enabled = False
    Cmd_Rubah.Visible = False
    Cmd_Batal.Visible = True
    Cmd_Hapus.Enabled = False
    Cmd_Daftar.Enabled = False
    Cmd_Keluar.Enabled = False
    
    Txt_Cr_Daftar(0).Text = ""
    Txt_Cr_Daftar(1).Text = ""
    
    txt_cr_daftar_KeyUp 0, 0, 0
    
    .Visible = True
    
    Txt_Cr_Daftar(0).SetFocus
    
Else
    
    Frame_Nav.Enabled = True
    
    .Visible = False
    
End If

End With

End Sub

Private Sub Cmd_Hapus_Click()

rubah = False
Tujuan = 1

With TDB_Daftar

If .Visible = False Then
    
    Frame_Nav.Enabled = False
    
    Cmd_Tambah.Enabled = False
    Cmd_Rubah.Visible = False
    Cmd_Batal.Visible = True
    Cmd_Hapus.Enabled = False
    Cmd_Daftar.Enabled = False
    Cmd_Keluar.Enabled = False
    
    Txt_Cr_Daftar(0).Text = ""
    Txt_Cr_Daftar(1).Text = ""
    
    txt_cr_daftar_KeyUp 0, 0, 0
    
    .Visible = True
    
    Txt_Cr_Daftar(0).SetFocus
    
Else
    
    Frame_Nav.Enabled = True
    
    .Visible = False
    
End If

End With

End Sub

Private Sub Cmd_Keluar_Click()
    Unload Me
End Sub

Private Sub Cmd_Navigasi_Click(Index As Integer)

On Error Resume Next

With Rs_Nav
Select Case Index
    Case 0
        .MoveFirst
    Case 1
        
        If .BOF Then .MoveFirst
        
        .MovePrevious
        
        If .BOF Then .MoveFirst
        
    Case 2
        
        If .EOF Then .MoveLast
        
        .MoveNext
        
        If .EOF Then .MoveLast
        
    Case 3
        
        .MoveLast
        
End Select
End With

isi_semua Rs_Nav


End Sub

Private Sub isi_semua(ByVal rec As Recordset)
On Error Resume Next
    
    With rec
        
        If .BOF Then .MoveFirst
        If .EOF Then .MoveLast
        
        
        Id_Pwd = !Id_User
        Txt_Kode_Karyawan.Text = IIf(Not IsNull(!kode_karyawan), !kode_karyawan, "")
        Txt_Nama_Karyawan.Text = IIf(Not IsNull(!Nama_Karyawan), !Nama_Karyawan, "")
'        Txt_Jabatan.Text = IIf(Not IsNull(!jabatan), !jabatan, "")
        
        Txt_Password.Text = IIf(Not IsNull(!pwd), !pwd, "")
        Txt_Verifikasi.Text = IIf(Not IsNull(!pwd), !pwd, "")
        Check_Aktif.Value = !status
        
        If .RecordCount = 0 Then
            Lbl_Info.Caption = "Record Ke " & 0 & " Dari " & .RecordCount & " Record"
        Else
            Lbl_Info.Caption = "Record Ke " & .AbsolutePosition & " Dari " & .RecordCount & " Record"
        End If
        
    End With
    
End Sub


Private Sub Cmd_Rubah_Click()

rubah = False
Tujuan = 0

With TDB_Daftar

If .Visible = False Then
    
    Frame_Nav.Enabled = False
    
    Cmd_Tambah.Visible = False
    Cmd_Simpan.Visible = True
    Cmd_Simpan.Enabled = False
    Cmd_Rubah.Visible = False
    Cmd_Batal.Visible = True
    Cmd_Hapus.Enabled = False
    Cmd_Daftar.Enabled = False
    Cmd_Keluar.Enabled = False
    
    Txt_Cr_Daftar(0).Text = ""
    Txt_Cr_Daftar(1).Text = ""
    
    txt_cr_daftar_KeyUp 0, 0, 0
    
    .Visible = True
    
    Txt_Cr_Daftar(0).SetFocus
    
Else
    
    .Visible = False
    
    Frame_Nav.Enabled = True
    
End If

End With

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
    
    
         kon.BeginTrans
    
    If rubah = False Then
    
    sql1 = "select Kode_Karyawan from Tb_User where Kode_Karyawan='" & Trim(Txt_Kode_Karyawan.Text) & "'"
    
    Set rs1 = New ADODB.Recordset
        rs1.Open sql1, kon
        
        With rs1
            
            If Not .EOF Then
                
                konfirm = CInt(MsgBox("Karyawan sudah ada dalam daftar user", vbOKOnly + vbInformation, "Information"))
                
            Else
                                
                sql = "insert into Tb_User (Kode_Karyawan,Pwd,Status) values('" & Trim(Txt_Kode_Karyawan.Text) & "','" & encrypt_pwd(Trim(Txt_Password.Text)) & "'," & Check_Aktif.Value & ")"
                
                Set rs = New ADODB.Recordset
                    rs.Open sql, kon
                
'                kon.CommitTrans
'
                konfirm = CInt(MsgBox("Data user telah tersimpan", vbOKOnly + vbInformation, "Informasi"))
                
                                
            End If
        End With
    
    
    
    Else
    
        sql = "update Tb_User set Status= " & Check_Aktif.Value & " where Id_User =" & Id_Pwd
            
            Set rs = New ADODB.Recordset
                rs.Open sql, kon
        
'        kon.CommitTrans
        
        konfirm = CInt(MsgBox("Data user telah dirubah", vbOKOnly + vbInformation, "Informasi"))
        
    End If
    
    kon.CommitTrans
    
    Cmd_Batal_Click
    
    On Error GoTo 0
    Exit Sub

err_handler:
    
    kon.RollbackTrans
        
        konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
            Err.Clear

End Sub

Private Sub Cmd_Tambah_Click()

    rubah = False
    
    Frame_Nav.Enabled = False
    Cmd_Tambah.Visible = False
    Cmd_Simpan.Visible = True
    Cmd_Simpan.Enabled = False
     Cmd_Rubah.Visible = False
     Cmd_Batal.Visible = True
     Cmd_Hapus.Enabled = False
     Cmd_Daftar.Enabled = False
     Cmd_Keluar.Enabled = False
        
     Lbl_Info.Caption = "Record Ke " & 0
        
     Cmd_Browse_Karyawan.Enabled = True
     Txt_Kode_Karyawan.Enabled = True
     Txt_Kode_Karyawan.Text = ""
     Txt_Kode_Karyawan.SetFocus
        

End Sub

Private Sub Form_Activate()
    On Error Resume Next
        Cmd_Tambah.SetFocus
End Sub

Private Sub Form_Load()
    
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
    
'' akses command ''

'    hak_akses_percommand CStr(Me.Name)
'
'    Cmd_Tambah.Enabled = c_tambah
'    Cmd_Rubah.Enabled = c_rubah
'    Cmd_Hapus.Enabled = c_hapus

'' stop here ''
    
    rubah = False
    Txt_Kode_Karyawan.Enabled = False
    Cmd_Browse_Karyawan.Enabled = False
    Txt_Password.Enabled = False
    Txt_Verifikasi.Enabled = False
    Frame_Stats.Enabled = False
            
    With TDB_Karyawan
        .Left = 600
        .Top = 840
    End With
        
    With TDB_Daftar
        .Left = Me.Width / 2 - .Width / 2
        .Top = Me.Height / 2 - .Height / 2
    End With
    
    Grid_Karyawan.Array = Arr_Karyawan
    Grid_Daftar.Array = arr_daftar
                    
    txt_cr_daftar_KeyUp 0, 0, 0
        Cmd_Navigasi_Click 3
                    
End Sub

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

Private Sub Isi_Grid_Transaksif(ByVal grid As Object, ByVal arr As Object, ByVal rec As Recordset)
    
    Dim a As Long
    Dim id, kode, NAMA, jab As String
        
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
                            
                            id = !Id_User
                            kode = IIf(Not IsNull(!kode_karyawan), !kode_karyawan, "")
                            NAMA = IIf(Not IsNull(!Nama_Karyawan), !Nama_Karyawan, "")
'                            jab = IIf(Not IsNull(!jabatan), !jabatan, "")
                            
                            arr(a, 0) = id
                            arr(a, 1) = kode
                            arr(a, 2) = NAMA
                            arr(a, 3) = ""
                        
                    a = a + 1
                    .MoveNext
                    Loop
                    
                    grid.ReBind
                    grid.Refresh
                    
                    grid.MoveFirst
        
        End With
        
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Cmd_Keluar.Enabled = False Then
        Cancel = True
    Else
        Cancel = False

    
    If kon.State = adStateOpen Then
        
        kon.Close
        Set kon = Nothing
    End If
        
    End If
End Sub

Private Sub grid_daftar_DblClick()
On Error GoTo err_handler

If arr_daftar.UpperBound(1) = 1 And arr_daftar(1, 1) = Empty Then Exit Sub
    
    With Rs_Nav
        
        .MoveFirst
        
        .Find "Id_User='" & arr_daftar(Grid_Daftar.Bookmark, 0) & "'"
        
    Id_Pwd = arr_daftar(Grid_Daftar.Bookmark, 0)
    Txt_Kode_Karyawan.Text = arr_daftar(Grid_Daftar.Bookmark, 1)
    Txt_Nama_Karyawan.Text = arr_daftar(Grid_Daftar.Bookmark, 2)
'    Txt_Jabatan.Text = arr_daftar(Grid_Daftar.Bookmark, 3)
        
    Txt_Password.Text = IIf(Not IsNull(!pwd), !pwd, "")
    Txt_Verifikasi.Text = IIf(Not IsNull(!pwd), !pwd, "")
    Check_Aktif.Value = !status
        
    End With
    
    
Select Case Tujuan

    Case 0
        
        rubah = True
        
        Frame_Stats.Enabled = True
        TDB_Daftar.Visible = False
        Check_Aktif.Enabled = True
        Check_Aktif.SetFocus
        Cmd_Simpan.Enabled = True
             
    Case 1
        
        Dim sql As String
        Dim rs As Recordset
        
        If MsgBox("Yakin akan dihapus", vbYesNo + vbQuestion, "Hapus") = vbNo Then
            
            On Error GoTo 0
            Exit Sub
        End If
        
        kon.BeginTrans
        
        sql = "delete from Tb_User where Id_User=" & Id_Pwd
        Set rs = New ADODB.Recordset
            rs.Open sql, kon
            
            kon.CommitTrans
            
        Dim konfirm As Integer
            konfirm = CInt(MsgBox("Data user berhasil dihapus", vbOKOnly + vbInformation, "Informasi"))
            
            Cmd_Batal_Click
        
    Case 2
               
         Cmd_Navigasi_Click 3
            
         TDB_Daftar.Visible = False
         Frame_Nav.Enabled = True
         Cmd_Navigasi(1).SetFocus
         
         
End Select

On Error GoTo 0
Exit Sub


err_handler:
    
    If Tujuan = 1 Then kon.RollbackTrans
        
        konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Informasi"))
            Err.Clear

End Sub

Private Sub grid_daftar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then grid_daftar_DblClick
    If KeyCode = vbKeyEscape Then Cmd_Batal_Click
End Sub

Private Sub Grid_Karyawan_DblClick()

    If Arr_Karyawan.UpperBound(1) = 1 And Arr_Karyawan(1, 1) = Empty Then Exit Sub
    
    Txt_Kode_Karyawan = Arr_Karyawan(Grid_Karyawan.Bookmark, 0)
    Txt_Nama_Karyawan = Arr_Karyawan(Grid_Karyawan.Bookmark, 1)
'    Txt_Jabatan = Arr_Karyawan(Grid_Karyawan.Bookmark, 2)
    
    Txt_Password.Enabled = True
    Txt_Verifikasi.Enabled = True
    Cmd_Simpan.Enabled = True
    Frame_Stats.Enabled = True
    Check_Aktif.Enabled = True
    
    Txt_Password.Text = ""
    Txt_Verifikasi.Text = ""
    
    Check_Aktif.Value = 1
    
    TDB_Karyawan.Visible = False
    Txt_Password.SetFocus

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

Private Sub txt_cr_daftar_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Daftar.SetFocus
    If KeyCode = vbKeyEscape Then Cmd_Batal_Click
End Sub

Private Sub txt_cr_daftar_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim sql As String
    
    sql = "select * from VIEW_User"
    
    Select Case Index
        Case 0
            sql = sql & " where Kode_Karyawan like '%" & Trim(Txt_Cr_Daftar(0).Text) & "%'"
        Case 1
            sql = sql & " where Nama_Karyawan like '%" & Trim(Txt_Cr_Daftar(1).Text) & "%'"
    End Select
    
    sql = sql & " order by Kode_Karyawan asc"
    
    Set Rs_Nav = New ADODB.Recordset
        Rs_Nav.Open sql, kon, adOpenKeyset
        
        Isi_Grid_Transaksif Grid_Daftar, arr_daftar, Rs_Nav


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
    
    If KeyCode = vbKeyF3 Then Cmd_Browse_Karyawan_Click
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
'        Txt_Jabatan.Text = ""
        Txt_Password.Text = ""
        Txt_Verifikasi.Text = ""
        Check_Aktif.Value = 0
                
            sql = "select Kode_Karyawan,Nama_Karyawan from VIEW_Karyawan where Kode_Karyawan ='" & Trim(Txt_Kode_Karyawan.Text) & "'"
            
            Set rs = New ADODB.Recordset
                rs.Open sql, kon
            
            With rs
                
                If Not .EOF Then
                    
                    Txt_Nama_Karyawan.Text = IIf(Not IsNull(!Nama_Karyawan), !Nama_Karyawan, "")
'                    Txt_Jabatan.Text = IIf(Not IsNull(!jabatan), !jabatan, "")
                    
                    Cmd_Simpan.Enabled = True
                    Txt_Password.Enabled = True
                    Txt_Verifikasi.Enabled = True
                    Frame_Stats.Enabled = True
                    Check_Aktif.Enabled = True
                    Txt_Password.SetFocus
                    
                    Txt_Password.Text = ""
                    Txt_Verifikasi.Text = ""
                    Check_Aktif.Value = 1
                    
                Else
                    
                    konfirm = CInt(MsgBox("Kode karyawan tidak ditemukan", vbOKOnly + vbInformation, "Information"))
                    
                    Cmd_Simpan.Enabled = False
                    Txt_Password.Enabled = False
                    Txt_Verifikasi.Enabled = False
                    Check_Aktif.Enabled = False
                    
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

Private Sub Txt_Verifikasi_GotFocus()
    Call Focus_(Txt_Verifikasi)
End Sub

Private Sub Txt_Verifikasi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Check_Aktif.SetFocus
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
