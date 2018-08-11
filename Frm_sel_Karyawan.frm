VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frm_sel_Karyawan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleksi Karyawan"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_sel_Karyawan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   0
      ScaleHeight     =   4185
      ScaleWidth      =   5145
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton Cmd_Keluar 
         Caption         =   "&Keluar"
         Height          =   615
         Left            =   4080
         TabIndex        =   22
         Top             =   3360
         Width           =   855
      End
      Begin VB.CommandButton Cmd_Lihat 
         Caption         =   "&Tampil"
         Height          =   615
         Left            =   3120
         TabIndex        =   21
         Top             =   3360
         Width           =   855
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   4815
         Begin VB.TextBox Txt_Telp 
            Height          =   320
            Left            =   1200
            TabIndex        =   12
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox Txt_Alamat 
            Height          =   320
            Left            =   1200
            TabIndex        =   11
            Top             =   1440
            Width           =   3375
         End
         Begin VB.TextBox Txt_Nama 
            Height          =   320
            Left            =   1200
            TabIndex        =   7
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox Txt_Kode 
            Height          =   320
            Left            =   1200
            TabIndex        =   6
            Top             =   360
            Width           =   1215
         End
         Begin MSMask.MaskEdBox Tgl_Lhr1 
            Height          =   315
            Left            =   1200
            TabIndex        =   8
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Tgl_Lhr2 
            Height          =   315
            Left            =   3120
            TabIndex        =   9
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Tgl_Masuk1 
            Height          =   315
            Left            =   1200
            TabIndex        =   13
            Top             =   2160
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Tgl_Masuk2 
            Height          =   315
            Left            =   3120
            TabIndex        =   14
            Top             =   2160
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl Masuk :"
            Height          =   195
            Index           =   2
            Left            =   255
            TabIndex        =   20
            Top             =   2160
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telp :"
            Height          =   195
            Index           =   5
            Left            =   660
            TabIndex        =   19
            Top             =   1800
            Width           =   405
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat :"
            Height          =   195
            Index           =   1
            Left            =   465
            TabIndex        =   18
            Top             =   1440
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl Lhr :"
            Height          =   195
            Index           =   4
            Left            =   480
            TabIndex        =   17
            Top             =   1080
            Width           =   585
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama :"
            Height          =   195
            Index           =   3
            Left            =   555
            TabIndex        =   16
            Top             =   720
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "S/D"
            Height          =   195
            Index           =   19
            Left            =   2760
            TabIndex        =   15
            Top             =   2160
            Width           =   255
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "S/D"
            Height          =   195
            Index           =   18
            Left            =   2760
            TabIndex        =   10
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode :"
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   5
            Top             =   360
            Width           =   465
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   3855
         Begin VB.OptionButton Opt_Kriteria 
            Caption         =   "&Berdasarkan Kriteria"
            Height          =   255
            Left            =   960
            TabIndex        =   3
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton Opt_Semua 
            Caption         =   "&Semua"
            Height          =   255
            Left            =   0
            TabIndex        =   2
            Top             =   120
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "Frm_sel_Karyawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check_Foto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Lihat.SetFocus
End Sub

Private Sub Cmd_Keluar_Click()

    Unload Me

End Sub

Private Sub Cmd_Lihat_Click()
    
    Dim sql As String
    
    If Opt_Semua.Value = True Then
    
    sql = "select * from VIEW_Karyawan order by Nama_Karyawan asc"
    
    Else
    
    If Txt_Kode.Text <> "" Or Txt_Nama.Text <> "" Or Tgl_Lhr1.Text <> "__/__/____" Or Tgl_Lhr2.Text <> "__/__/____" Or Txt_Alamat.Text <> "" Or _
        Txt_Telp.Text <> "" Or Tgl_Masuk1.Text <> "__/__/____" Or Tgl_Masuk2.Text <> "__/__/____" Then
        
        sql = "select * from VIEW_Karyawan where"
        
        If Txt_Kode.Text <> "" Then
            sql = sql & " Kode_Karyawan like '%" & Trim(Txt_Kode.Text) & "%'"
        End If
        
        If Txt_Nama.Text <> "" And Txt_Kode.Text = "" Then
            sql = sql & " Nama_Karyawan like '%" & Trim(Txt_Nama.Text) & "%'"
        End If
        
        If Txt_Nama.Text <> "" And Txt_Kode.Text <> "" Then
            sql = sql & " and Nama_Karyawan like '%" & Trim(Txt_Nama.Text) & "%'"
        End If
        
        If Tgl_Lhr1.Text <> "__/__/____" And Tgl_Lhr2.Text <> "__/__/____" And Txt_Nama.Text = "" And Txt_Kode.Text = "" Then
            sql = sql & " Tgl_Lhr >='" & Format(Trim(Tgl_Lhr1.Text), "yyyy/mm/dd") & "' and Tgl_Lhr <='" & Format(Trim(Tgl_Lhr2.Text), "yyyy/mm/dd") & "'"
        End If
        
        If Tgl_Lhr1.Text <> "__/__/____" And Tgl_Lhr2.Text <> "__/__/____" And (Txt_Nama.Text <> "" Or Txt_Kode.Text <> "") Then
            sql = sql & " and Tgl_Lhr >='" & Format(Trim(Tgl_Lhr1.Text), "yyyy/mm/dd") & "' and Tgl_Lhr <='" & Format(Trim(Tgl_Lhr2.Text), "yyyy/mm/dd") & "'"
        End If
        
        If Txt_Alamat.Text <> "" And Tgl_Lhr1.Text = "__/__/____" And Tgl_Lhr2.Text = "__/__/____" And Txt_Nama.Text = "" And Txt_Kode.Text = "" Then
            sql = sql & " Alamat_1 like '%" & Trim(Txt_Alamat.Text) & "%' or Alamat_2 like '%" & Trim(Txt_Alamat.Text) & "%' or Alamat_3 like '%" & Trim(Txt_Alamat.Text) & "%'"
        End If
        
        If Txt_Alamat.Text <> "" And ((Tgl_Lhr1.Text <> "__/__/____" And Tgl_Lhr2.Text <> "__/__/____") Or Txt_Nama.Text <> "" Or Txt_Kode.Text <> "") Then
            sql = sql & " and Alamat_1 like '%" & Trim(Txt_Alamat.Text) & "%' or Alamat_2 like '%" & Trim(Txt_Alamat.Text) & "%' or Alamat_3 like '%" & Trim(Txt_Alamat.Text) & "%'"
        End If
        
        If Txt_Telp.Text <> "" And Txt_Alamat.Text = "" And Tgl_Lhr1.Text = "__/__/____" And Tgl_Lhr2.Text = "__/__/____" And Txt_Nama.Text = "" And Txt_Kode.Text = "" Then
            sql = sql & " No_telp like '%" & Trim(Txt_Telp.Text) & "%' or No_telp_Hp like '%" & Trim(Txt_Telp.Text) & "%'"
        End If
        
        If Txt_Telp.Text <> "" And (Txt_Alamat.Text <> "" Or (Tgl_Lhr1.Text <> "__/__/____" And Tgl_Lhr2.Text <> "__/__/____") Or Txt_Nama.Text <> "" Or Txt_Kode.Text <> "") Then
            sql = sql & " and No_telp like '%" & Trim(Txt_Telp.Text) & "%' or No_telp_Hp like '%" & Trim(Txt_Telp.Text) & "%'"
        End If
        
        If Tgl_Masuk1.Text <> "__/__/____" And Tgl_Masuk2.Text <> "__/__/____" And Txt_Telp.Text = "" And Txt_Alamat.Text = "" And Tgl_Lhr1.Text = "__/__/____" And Tgl_Lhr2.Text = "__/__/____" And Txt_Nama.Text = "" And Txt_Kode.Text = "" Then
            sql = sql & " Tgl_masuk >= '" & Format(Trim(Tgl_Masuk1.Text), "yyyy/mm/dd") & "' and Tgl_Masuk <='" & Format(Trim(Tgl_Masuk2.Text), "yyyy/mm/dd") & "'"
        End If
        
        If Tgl_Masuk1.Text <> "__/__/____" And Tgl_Masuk2.Text <> "__/__/____" And (Txt_Telp.Text <> "" Or Txt_Alamat.Text <> "" Or (Tgl_Lhr1.Text <> "__/__/____" And Tgl_Lhr2.Text <> "__/__/____") Or Txt_Nama.Text <> "" Or Txt_Kode.Text <> "") Then
            sql = sql & " and Tgl_masuk >= '" & Format(Trim(Tgl_Masuk1.Text), "yyyy/mm/dd") & "' and Tgl_Masuk <='" & Format(Trim(Tgl_Masuk2.Text), "yyyy/mm/dd") & "'"
        End If
        
        sql = sql & " order by Nama_Karyawan asc"
        
        
    Else
        
        Dim konfirm As Integer
            konfirm = CInt(MsgBox("Kriteria pencarian harus diisi", vbOKOnly + vbInformation, "Informasi"))
        
        Exit Sub
    End If
    
    End If
    
'    khusus_user = Mid(Utama.StatusBar1.Panels(5).Text, 7, Len(Utama.StatusBar1.Panels(5).Text))
    
    Mysq = sql
    
    Load frm_lap_kary
        frm_lap_kary.Show
    
    
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
        .Top = 250
    End With
    
    Opt_Semua.Value = True

'' akses command ''

'    hak_akses_percommand CStr(Me.Name)
'
'    Cmd_Lihat.Enabled = c_laporan

'' stop here ''


End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If kon.State = adStateOpen Then
        
        kon.Close
        Set kon = Nothing
    End If
    
 
End Sub

Private Sub Opt_Kriteria_Click()
    
    If Opt_Kriteria.Value = True Then Frame2.Enabled = True
    
End Sub

Private Sub Opt_Kriteria_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Kode.SetFocus
End Sub

Private Sub Opt_Semua_Click()
    If Opt_Semua.Value = True Then
        Frame2.Enabled = False
    
    Dim a As Object
        For Each a In Me
            If TypeOf a Is TextBox Then
                a.Text = ""
            End If
            
            If TypeOf a Is MaskEdBox Then a.Text = "__/__/____"
        Next
        
        Set a = Nothing
    
    End If
End Sub

Private Sub Opt_Semua_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Cmd_Lihat.Enabled = True Then Cmd_Lihat.SetFocus
    End If
End Sub

Private Sub Tgl_Lhr1_GotFocus()
    Call Focus_(Tgl_Lhr1)
End Sub

Private Sub Tgl_Lhr1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Tgl_Lhr2.SetFocus
End Sub

Private Sub Tgl_Lhr2_GotFocus()
    Call Focus_(Tgl_Lhr2)
End Sub

Private Sub Tgl_Lhr2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Alamat.SetFocus
End Sub

Private Sub Tgl_Masuk1_GotFocus()
    Call Focus_(Tgl_Masuk1)
End Sub

Private Sub Tgl_Masuk1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Tgl_Masuk2.SetFocus
End Sub

Private Sub Tgl_Masuk2_GotFocus()
    Call Focus_(Tgl_Masuk2)
End Sub

Private Sub Tgl_Masuk2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Cmd_Lihat.Enabled = True Then Cmd_Lihat.SetFocus
    End If
        
End Sub

Private Sub Txt_Alamat_GotFocus()
    Call Focus_(Txt_Alamat)
End Sub

Private Sub Txt_Alamat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Telp.SetFocus
End Sub


Private Sub txt_kode_GotFocus()
    Call Focus_(Txt_Kode)
End Sub

Private Sub Txt_Kode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Nama.SetFocus
End Sub

Private Sub txt_nama_GotFocus()
    Call Focus_(Txt_Nama)
End Sub

Private Sub Txt_Nama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Tgl_Lhr1.SetFocus
End Sub



Private Sub Txt_Telp_GotFocus()
    Call Focus_(Txt_Telp)
End Sub

Private Sub Txt_Telp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Tgl_Masuk1.SetFocus
End Sub
