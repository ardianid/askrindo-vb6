VERSION 5.00
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form Karyawan2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Jabatan Detail..."
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Jabatan 
      Height          =   2415
      Left            =   -5520
      TabIndex        =   13
      Top             =   2760
      Visible         =   0   'False
      Width           =   6255
      _Version        =   65536
      _ExtentX        =   11033
      _ExtentY        =   4260
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Karyawan2.frx":0000
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Karyawan2.frx":001C
      Childs          =   "Karyawan2.frx":00C8
      Begin VB.TextBox txt_cr_jabatan 
         Height          =   300
         Index           =   0
         Left            =   1200
         TabIndex        =   15
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txt_cr_jabatan 
         Height          =   300
         Index           =   1
         Left            =   3960
         TabIndex        =   14
         Top             =   120
         Width           =   1815
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Jabatan 
         Height          =   1695
         Left            =   240
         OleObjectBlob   =   "Karyawan2.frx":00E4
         TabIndex        =   16
         Top             =   480
         Width           =   5775
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kd Jabatan :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   25
         Left            =   240
         TabIndex        =   18
         Top             =   120
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Jabatan :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   27
         Left            =   2760
         TabIndex        =   17
         Top             =   120
         Width           =   1140
      End
   End
   Begin VB.TextBox t_no_bagian 
      Height          =   315
      Left            =   4080
      TabIndex        =   19
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Selesai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   12
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Tambah"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   11
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox tket 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   1920
      Width           =   5655
   End
   Begin VB.TextBox t_nama_bagian 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   7
      Top             =   1200
      Width           =   5655
   End
   Begin VB.TextBox t_kd_jabatan 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox t_nama_jabatan 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   5655
   End
   Begin VB.CommandButton cmd_browse_jab 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   3000
      TabIndex        =   2
      Top             =   480
      Width           =   300
   End
   Begin TDBDate6Ctl.TDBDate ttgl 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   556
      Calendar        =   "Karyawan2.frx":37D9
      Caption         =   "Karyawan2.frx":38F1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Karyawan2.frx":395D
      Keys            =   "Karyawan2.frx":397B
      Spin            =   "Karyawan2.frx":39D9
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   1863103
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "05/01/2007"
      ValidateMode    =   0
      ValueVT         =   2010382343
      Value           =   39087
      CenturyMode     =   0
   End
   Begin TDBNumber6Ctl.TDBNumber tdb_gapok 
      Height          =   345
      Left            =   1320
      TabIndex        =   20
      Top             =   1560
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   609
      Calculator      =   "Karyawan2.frx":3A01
      Caption         =   "Karyawan2.frx":3A21
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Karyawan2.frx":3A8D
      Keys            =   "Karyawan2.frx":3AAB
      Spin            =   "Karyawan2.frx":3AF5
      AlignHorizontal =   1
      AlignVertical   =   2
      Appearance      =   1
      BackColor       =   16777215
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###;;0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   12582912
      Format          =   "###,###,###"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999
      MinValue        =   -999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   1
      Value           =   0
      MaxValueVT      =   1028849669
      MinValueVT      =   1598423045
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gaji Pokok :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   360
      TabIndex        =   21
      Top             =   1560
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Bagian :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kd Jabatan :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   345
      TabIndex        =   6
      Top             =   480
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Jabatan :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Mulai :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   225
      TabIndex        =   1
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "Karyawan2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Moving As Boolean
Dim yold, xold As Long

Private Sub Command1_Click()
    
    Dim a As Integer
    If t_kd_jabatan.Text = "" Then
       a = CInt(MsgBox("Jabatan tidak boleh kosong...", vbOKOnly + vbInformation, "Informasi"))
        t_kd_jabatan.SetFocus
        Exit Sub
    End If
    
    If tdb_gapok.ValueIsNull Then
        a = CInt(MsgBox("Gaji pokok tidak boleh kosong...", vbOKOnly + vbInformation, "Informasi"))
        tdb_gapok.SetFocus
        Exit Sub
    End If
    
    If tdb_gapok.Value = 0 Then
        a = CInt(MsgBox("Gaji pokok tidak boleh kosong...", vbOKOnly + vbInformation, "Informasi"))
        tdb_gapok.SetFocus
        Exit Sub
    End If
    
    With Karyawan
        .tambah_jab Trim(t_no_bagian.Text), Trim(t_nama_bagian.Text), Trim(t_kd_jabatan.Text), Trim(t_nama_jabatan.Text), ttgl.Value, Trim(tket.Text), tdb_gapok.Value
    End With
    
    Unload Me
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    ttgl.SetFocus
End Sub

Private Sub cmd_browse_jab_Click()

    
    With TDB_Jabatan
        If .Visible = False Then
            
            .Left = cmd_browse_jab.Left + cmd_browse_jab.Width / 2 - .Width / 2
            .Top = cmd_browse_jab.Top + cmd_browse_jab.Height + 15
            
            txt_cr_jabatan(0).Text = ""
            txt_cr_jabatan(1).Text = ""
            
            txt_cr_jabatan_KeyUp 0, 0, 0
            .Visible = True
            txt_cr_jabatan(0).SetFocus
            
        Else
            .Visible = False
        End If
    End With


End Sub

Private Sub Form_Load()
        
    ttgl.Value = Date
    
    t_kd_jabatan.Text = ""
    t_nama_jabatan.Text = ""
    t_nama_bagian.Text = ""
    tket.Text = ""
    
End Sub

Private Sub ttgl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        t_kd_jabatan.SetFocus
    End If
End Sub

Private Sub txt_cr_jabatan_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Jabatan.SetFocus
    If KeyCode = vbKeyEscape Then cmd_browse_jab_Click
End Sub

Private Sub txt_cr_jabatan_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    Dim sql As String
    Dim rs As Recordset
    
        sql = "select top 100 * from VIEW_Jabatan"
        
    If txt_cr_jabatan(0).Text <> "" Or txt_cr_jabatan(1).Text <> "" Then
        
        sql = sql & " where "
        
        Select Case Index
            Case 0
                sql = sql & "kode_jab like '%" & Trim(txt_cr_jabatan(0).Text) & "%'"
            Case 1
                sql = sql & "nama_jab like '%" & Trim(txt_cr_jabatan(1).Text) & "%'"
        End Select
    End If
    
    sql = sql & " order by kode_jab desc"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
    Set Grid_Jabatan.DataSource = rs
        Grid_Jabatan.Refresh

End Sub

Private Sub t_kd_jabatan_GotFocus()
    Call Focus_(t_kd_jabatan)
End Sub

Private Sub t_kd_jabatan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then t_kd_jabatan_LostFocus
    If KeyCode = vbKeyF3 Then cmd_browse_jab_Click
End Sub

Private Sub t_kd_jabatan_LostFocus()
    
    If t_kd_jabatan.Text = "" Then Exit Sub
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select * from VIEW_Jabatan where kode_jab='" & Trim(t_kd_jabatan.Text) & "'"
      '  sql = sql & " and no_bagian='" & Trim(t_no_bagian.Text) & "'"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon, adOpenKeyset
        
        With rs
            If Not .EOF Then
                
                t_nama_jabatan.Text = IIf(Not IsNull(!nama_jab), !nama_jab, "")
                t_nama_bagian.Text = IIf(Not IsNull(!nama_bagian), !nama_bagian, "")
                t_no_bagian.Text = IIf(Not IsNull(!no_bagian), !no_bagian, "")
                
                tdb_gapok.SetFocus
            
            Else
                
                a = CInt(MsgBox("Kode jabatan yang anda masukkan tidak ditemukan ", vbOKOnly + vbInformation, "Informasi"))
                
                t_kd_jabatan.Text = ""
                
                t_kd_jabatan.SetFocus
                
            End If
        End With
        
    
End Sub

Private Sub Grid_Jabatan_DblClick()

On Error GoTo err_handler

'If Grid_Bagian.Row < 0 Then Exit Sub

    t_kd_jabatan.Text = Grid_Jabatan.Columns(2).Text
    t_nama_jabatan.Text = Grid_Jabatan.Columns(3).Text
    t_nama_bagian.Text = Grid_Jabatan.Columns(1).Text
    t_no_bagian.Text = Grid_Jabatan.Columns(0).Text
    
    TDB_Jabatan.Visible = False
    t_kd_jabatan_LostFocus
    
    Exit Sub
    
err_handler:
        
        Dim konfirm As Integer
        
        konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
        Err.Clear


End Sub

Private Sub Grid_Jabatan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Jabatan_DblClick
    If KeyCode = vbKeyEscape Then cmd_browse_jab_Click
End Sub

Private Sub TDB_Jabatan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub TDB_Jabatan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   TDB_Jabatan.Top = TDB_Jabatan.Top - (yold - Y)
   TDB_Jabatan.Left = TDB_Jabatan.Left - (xold - X)
End If

End Sub

Private Sub TDB_Jabatan_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub
