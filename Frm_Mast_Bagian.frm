VERSION 5.00
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form Frm_Mast_Bagian 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Master Bagian"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Cari 
      Height          =   2055
      Left            =   -360
      TabIndex        =   0
      Top             =   3360
      Visible         =   0   'False
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   3625
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Frm_Mast_Bagian.frx":0000
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Frm_Mast_Bagian.frx":001C
      Childs          =   "Frm_Mast_Bagian.frx":00C8
      Begin VB.CommandButton Cmd_Cancel 
         Caption         =   "&Keluar"
         Height          =   400
         Left            =   4680
         TabIndex        =   4
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton Cmd_OK 
         Caption         =   "&OK"
         Height          =   400
         Left            =   3840
         TabIndex        =   3
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Txt_Cr_Nama 
         Height          =   320
         Left            =   1920
         TabIndex        =   2
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox Txt_Cr_Kode 
         Height          =   320
         Left            =   1920
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   210
         Index           =   3
         Left            =   1800
         TabIndex        =   9
         Top             =   960
         Width           =   60
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   210
         Index           =   2
         Left            =   1800
         TabIndex        =   8
         Top             =   600
         Width           =   60
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Bagian"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Bagian"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   720
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   5400
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pencarian"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   960
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   120
      ScaleHeight     =   5025
      ScaleWidth      =   8145
      TabIndex        =   10
      Top             =   2040
      Width           =   8175
      Begin TrueOleDBGrid60.TDBGrid Grid_Status 
         Height          =   4815
         Left            =   120
         OleObjectBlob   =   "Frm_Mast_Bagian.frx":00E4
         TabIndex        =   11
         Top             =   120
         Width           =   7935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   8175
      Begin VB.Frame v 
         BackColor       =   &H00E0E0E0&
         Height          =   855
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   7935
         Begin VB.CommandButton cmd_keluar 
            Caption         =   "&Keluar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6840
            TabIndex        =   19
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton Cmd_Cari 
            Caption         =   "&Cari"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5880
            TabIndex        =   20
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmd_hapus 
            Caption         =   "&Hapus"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4920
            TabIndex        =   21
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmd_rubah 
            Caption         =   "&Rubah"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3960
            TabIndex        =   22
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmd_navigasi 
            Caption         =   "8"
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   2040
            TabIndex        =   23
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton cmd_navigasi 
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   1440
            TabIndex        =   24
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton cmd_navigasi 
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   840
            TabIndex        =   25
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton cmd_navigasi 
            Caption         =   "7"
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   240
            TabIndex        =   28
            Top             =   240
            Width           =   615
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Frame2"
            Height          =   855
            Left            =   2880
            TabIndex        =   27
            Top             =   0
            Width           =   15
         End
         Begin VB.CommandButton cmd_tambah 
            Caption         =   "&Tambah"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3000
            TabIndex        =   26
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmd_simpan 
            Caption         =   "&Simpan"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3000
            TabIndex        =   29
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmd_batal 
            Caption         =   "&Batal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3960
            TabIndex        =   30
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         TabIndex        =   13
         Top             =   -120
         Width           =   7935
         Begin VB.TextBox txt_kode 
            Height          =   320
            Left            =   2040
            TabIndex        =   15
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txt_jenis 
            Height          =   320
            Left            =   2040
            TabIndex        =   14
            Top             =   600
            Width           =   3495
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Bagian"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Bagian"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   240
            Width           =   825
         End
      End
   End
End
Attribute VB_Name = "Frm_Mast_Bagian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rubah As Boolean
Dim Moving As Boolean
Dim yold, xold As Long
'Dim idnya As Double

Private Sub IsiSemua()
    
    Dim sql As String
        sql = "select * from Tb_Bagian order by No_Bagian desc"
        
        Set Rs_Nav = New ADODB.Recordset
            Rs_Nav.Open sql, kon, adOpenKeyset
        
        Set Grid_Status.DataSource = Rs_Nav
            Grid_Status.Refresh
    
End Sub

Private Sub Cmd_Batal_Click()

    rubah = False
    
    cmd_simpan.Visible = False
    cmd_tambah.Visible = True
    cmd_tambah.Enabled = True
    cmd_batal.Visible = False
    cmd_rubah.Visible = True
    cmd_rubah.Enabled = True
    cmd_hapus.Enabled = True
    Cmd_Cari.Enabled = True
    cmd_keluar.Enabled = True
    'frame_nav.Enabled = False
    
    cmd_simpan.Enabled = True
        
    Dim n As Object
        For Each n In Me
            If TypeOf n Is TextBox Then
                If Left(n.Name, 6) <> "Txt_Cr" Then
                    n.Enabled = False
                End If
            End If
            
            If TypeOf n Is DTPicker Then n.Enabled = False
            
            If TypeOf n Is TDBContainer3D Then
                n.Visible = False
            End If
        Next
    Set n = Nothing
    
    cmd_tambah.SetFocus
        
    If Rs_Nav.State = adStateOpen Then
        If Rs_Nav.RecordCount > 0 Then Rs_Nav.MoveLast
    End If
    
End Sub

Private Sub Cmd_Cancel_Click()
    
    cmd_tambah.Enabled = True
    cmd_rubah.Visible = True
    cmd_batal.Visible = False
    cmd_hapus.Enabled = True
    Cmd_Cari.Enabled = True
    cmd_keluar.Enabled = True
    
    TDB_Cari.Visible = False
End Sub

Private Sub Cmd_Cari_Click()
        
    If Rs_Nav.RecordCount <= 0 Then Exit Sub
            
    cmd_tambah.Enabled = False
    cmd_rubah.Visible = False
    cmd_batal.Visible = True
    cmd_hapus.Enabled = False
'    Cmd_Cari.Enabled = False
    cmd_keluar.Enabled = False
        
    With TDB_Cari
        
        .Left = Me.Width / 2 - .Width / 2
        .Top = Me.Height / 2 - .Height / 2
        
        If .Visible = False Then
            Txt_Cr_Kode.Text = ""
            Txt_Cr_Nama.Text = ""
            .Visible = True
            Txt_Cr_Kode.SetFocus
        Else
            .Visible = False
        End If
        
    End With
    
End Sub

Private Sub Cmd_Hapus_Click()
On Error GoTo err_handler

'    If idnya = "" Then
'        On Error GoTo 0
'        Exit Sub
'    End If
    
    If Rs_Nav.RecordCount <= 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    
    If MsgBox("Yakin akan menghapus data bagian " & txt_jenis.Text, vbYesNo + vbQuestion, "Konfirmasi") = vbNo Then
        On Error GoTo 0
        Exit Sub
    End If
    
    Dim sql As String
    Dim rs As Recordset
    
    sql = "delete from Tb_Bagian where No_Bagian='" & Trim(txt_kode.Text) & "'"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon
    
    Dim konfirm As Integer
        konfirm = CInt(MsgBox("Data telah dihapus", vbOKOnly + vbInformation, "Informasi"))
    
    IsiSemua
    
    On Error GoTo 0
    Exit Sub

err_handler:
    
    'Dim Konfirm As Integer
        konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Informasi"))
            Err.Clear

End Sub

Private Sub Cmd_Keluar_Click()
    Unload Me
End Sub

Private Sub Cmd_Navigasi_Click(Index As Integer)

On Error Resume Next

With Rs_Nav
    Select Case Index
        Case 0
            .MoveLast
        Case 1
            
            If .EOF Then .MoveLast
                
                .MoveNext
                
            If .EOF Then .MoveLast
            
        Case 2
            
            If .BOF Then .MoveFirst
                
                .MovePrevious
                
            If .BOF Then .MoveFirst
            
        Case 3
            
            .MoveFirst
            
    End Select
End With

Set Grid_Status.DataSource = Rs_Nav
    Grid_Status.Refresh

End Sub

Private Sub Cmd_Ok_Click()
On Error Resume Next

    With Rs_Nav
        .MoveFirst
        
        If Txt_Cr_Kode.Text <> "" Then
'        If Len(Txt_Cr_Kode.Text) = 10 Then
            .Find "No_Bagian like '%" & Trim(Txt_Cr_Kode.Text) & "%'"
'        End If
        ElseIf Txt_Cr_Nama.Text <> "" And Txt_Cr_Kode.Text = "" Then
            .Find "Nama_Bagian like '%" & Trim(Txt_Cr_Nama.Text) & "%'"
'        ElseIf Txt_Cr_Nama.Text <> "" And Txt_Cr_Kode.Text <> "" Then
'            .Find "Tgl='" & Format(Trim(Txt_Cr_Kode.Text), "yyyy/mm/dd") & "' and Pendidikan like '%" & Trim(Txt_Cr_Nama.Text) & "%'"
        End If
        
    End With
    
    Set Grid_Status.DataSource = Rs_Nav
        Grid_Status.Refresh
    
   'TDB_Cari.Visible = False
    
End Sub

Private Sub Cmd_Rubah_Click()
    
    If Rs_Nav.RecordCount <= 0 Then Exit Sub
    
    cmd_tambah.Visible = False
    cmd_simpan.Visible = True
    cmd_rubah.Visible = False
    cmd_batal.Visible = True
    cmd_hapus.Enabled = False
    Cmd_Cari.Enabled = False
    cmd_keluar.Enabled = False

    rubah = True
    txt_jenis.Enabled = True
    
    
    txt_jenis.SetFocus
    
End Sub

Private Sub Cmd_Simpan_Click()
On Error GoTo err_handler

Dim konfirm As Integer
    If txt_kode.Text = "" Then
        konfirm = CInt(MsgBox("No Bagian tidak boleh kosong", vbOKOnly + vbInformation, "Informasi"))

        txt_kode.SetFocus
        On Error GoTo 0
        Exit Sub
    End If
    
    If txt_jenis.Text = "" Then
        konfirm = CInt(MsgBox("Nama Bagian tidak boleh kosong", vbOKOnly + vbInformation, "Informasi"))
        
        txt_jenis.SetFocus
        On Error GoTo 0
        Exit Sub
    End If
    
'    Dim harga As Double
'    If TDB_harga.ValueIsNull Then
'        harga = 0
'    Else
'        harga = Replace(Trim(TDB_harga.Value), ",", "")
'    End If
'
'    If harga = 0 Then
'        konfirm = CInt(MsgBox("Harga perjenis customer tidak boleh 0", vbOKOnly + vbInformation, "Informasi"))
'
'        TDB_harga.SetFocus
'        On Error GoTo 0
'        Exit Sub
'    End If
    
    Dim sql, sql1 As String
    Dim rs As Recordset
    Dim rs1 As Recordset
    
    If rubah = False Then
        
'        sql1 = "select Kode from Tb_Pendidikan where Kode='" & Trim(Txt_Kode.Text) & "'"
'
'        Set rs1 = New ADODB.Recordset
'            rs1.Open sql1, kon
'
'        With rs1
'            If Not .EOF Then
'                konfirm = CInt(MsgBox("Kode pendidikan yang anda masukkan sudah ada", vbOKOnly + vbInformation, "Informasi"))
'
'                Txt_Kode.SetFocus
'                On Error GoTo 0
'                Exit Sub
'            Else
                
                sql = "insert into Tb_Bagian (no_bagian,Nama_bagian) values('" & Trim(txt_kode.Text) & "','" & Trim(txt_jenis.Text) & "')"
                Set rs = New ADODB.Recordset
                    rs.Open sql, kon
                
                konfirm = CInt(MsgBox("Data sudah disimpan", vbOKOnly + vbInformation, "Informasi"))
                
'
'            End If
'        End With
    
    Else
        
        sql = "update Tb_Bagian set Nama_bagian='" & Trim(txt_jenis.Text) & "' where no_bagian='" & Trim(txt_kode.Text) & "'"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon
            
        konfirm = CInt(MsgBox("Data telah dirubah", vbOKOnly + vbInformation, "Informasi"))
        
    End If
    
    IsiSemua
    
    
    Cmd_Batal_Click
    On Error GoTo 0
    Exit Sub
    
err_handler:
    
    konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Informaton"))
        Err.Clear
    
End Sub

Private Sub Cmd_Tambah_Click()

rubah = False
'    Dim n As Object
'        For Each n In Me
'            If TypeOf n Is TextBox Then
'                If Left(n.Name, 6) <> "Txt_Cr" Then
'                    n.Text = ""
'                End If
'            End If
'        Next
'    Set n = Nothing
    
    txt_kode.Enabled = True
    txt_kode.Text = ""
    
    
    cmd_tambah.Visible = False
    cmd_simpan.Visible = True
    cmd_rubah.Visible = False
    cmd_batal.Visible = True
    cmd_hapus.Enabled = False
    Cmd_Cari.Enabled = False
    cmd_keluar.Enabled = False
    
'    cmd_simpan.Enabled = False
    
    txt_kode.SetFocus

End Sub

Private Sub Form_Activate()
    On Error Resume Next
        cmd_tambah.SetFocus
End Sub

Private Sub Form_Load()

'Dim status As String
'status = Buka_Koneksi
'If status = "-2147467259" Then
'    Dim konfirm As Integer
'        konfirm = CInt(MsgBox("Koneksi terputus ....", vbOKOnly + vbInformation, "Informasi"))
'
'        End
'        Exit Sub
'End If

With Me
    .Left = Screen.Width / 2 - .Width / 2
    .Top = 300
End With

IsiSemua

rubah = False
txt_kode.Enabled = False
txt_jenis.Enabled = False


End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    If cmd_keluar.Enabled = False Then
        Cancel = True
    Else
        Cancel = False
        
        If kon.State = adStateOpen Then
            
            kon.Close
            Set kon = Nothing
        End If
        
'        If kon1.State = adStateOpen Then
'
'            kon1.Close
'            Set kon1 = Nothing
'        End If
        
'        With Utama
'            .balikkan_samping
'        End With
        
    End If
    
End Sub

Private Sub Grid_Status_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    If Rs_Nav.RecordCount = 0 Then
        txt_kode.Text = ""
        txt_jenis.Text = ""
        
    Else
        
'        If Rs_Nav.BOF Then Rs_Nav.MoveFirst
'        If Rs_Nav.EOF Then Rs_Nav.MoveLast
        
        txt_kode.Text = Rs_Nav!no_bagian
        txt_jenis.Text = IIf(Not IsNull(Rs_Nav!nama_bagian), Rs_Nav!nama_bagian, "")
'        TDB_harga.Value = IIf(Not IsNull(Rs_Nav!harga), Rs_Nav!harga, Null)
        
    End If
    
End Sub

Private Sub TDB_Cari_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub TDB_Cari_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   TDB_Cari.Top = TDB_Cari.Top - (yold - Y)
   TDB_Cari.Left = TDB_Cari.Left - (xold - X)
End If

End Sub

Private Sub TDB_Cari_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub


Private Sub Txt_Cr_Kode_GotFocus()
    Call Focus_(Txt_Cr_Kode)
End Sub

Private Sub Txt_Cr_Kode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Cr_Nama.SetFocus
End Sub

Private Sub Txt_Cr_Nama_GotFocus()
    Call Focus_(Txt_Cr_Nama)
End Sub

Private Sub Txt_Cr_Nama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_OK.SetFocus
End Sub

'Private Sub txt_kode_GotFocus()
'    Call Focus_(Txt_Kode)
'End Sub
'
'Private Sub txt_kode_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
'
'        Dim konfirm As Integer
'            If Txt_Kode.Text = "" Then
'                konfirm = CInt(MsgBox("Kode pendidikan tidak boleh kosong", vbOKOnly + vbInformation, "Informasi"))
'                Exit Sub
'            End If
'
'            Dim sql As String
'            Dim rs As Recordset
'
'                sql = "select Kode from Tb_Pendidikan where Kode='" & Trim(Txt_Kode.Text) & "'"
'
'                Set rs = New ADODB.Recordset
'                    rs.Open sql, kon
'
'
'                With rs
'                    If Not .EOF Then
'                        konfirm = CInt(MsgBox("Kode pendidikan yang anda masukkan sudah ada", vbOKOnly + vbInformation, "Informasi"))
'                        cmd_simpan.Enabled = False
'                        txt_Pend.Enabled = False
'                    Else
'                        cmd_simpan.Enabled = True
'                        txt_Pend.Enabled = True
'                        txt_Pend.Text = ""
'                        txt_Pend.SetFocus
'                    End If
'                End With
'
'    End If
'End Sub

Private Sub txt_Pend_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmd_simpan.SetFocus
End Sub

Private Sub txt_jenis_GotFocus()
    Call Focus_(txt_jenis)
End Sub

Private Sub txt_jenis_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmd_simpan.SetFocus
End Sub

Private Sub txt_kode_GotFocus()
    Call Focus_(txt_kode)
End Sub

Private Sub txt_kode_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
    
    Dim konfirm As Integer
        If txt_kode.Text = "" Then
            konfirm = CInt(MsgBox("No Bagian tidak boleh kosong", vbOKOnly + vbInformation, "Informasi"))
            Exit Sub
        End If
    
    
    If Rs_Nav.RecordCount = 0 Then
        
        txt_jenis.Text = ""
        txt_jenis.Enabled = True
        
        txt_jenis.SetFocus
        
        Exit Sub
    End If
    
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select no_bagian from Tb_Bagian where no_bagian='" & Trim(txt_kode.Text) & "'"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon
        With rs
        If Not .EOF Then
            konfirm = CInt(MsgBox("No Bagian yang anda masukkan sudah ada", vbOKOnly + vbInformation, "Informasi"))
        Else
'            txt_kode.Text = txt_kode.Text
            txt_jenis.Text = ""
            txt_jenis.Enabled = True
            
            
            txt_jenis.SetFocus
            
        End If
        End With
    
    End If
    
End Sub


