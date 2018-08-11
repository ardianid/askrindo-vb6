VERSION 5.00
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form Frm_Mast_Jabatan 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Master Jabatan"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8535
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
   ScaleHeight     =   4710
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_bagian 
      Height          =   3255
      Left            =   8040
      TabIndex        =   58
      Top             =   1440
      Visible         =   0   'False
      Width           =   6495
      _Version        =   65536
      _ExtentX        =   11456
      _ExtentY        =   5741
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Frm_Mast_Jabatan.frx":0000
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Frm_Mast_Jabatan.frx":001C
      Childs          =   "Frm_Mast_Jabatan.frx":00C8
      Begin VB.TextBox txt_cr_bag 
         Height          =   285
         Index           =   1
         Left            =   4200
         TabIndex        =   65
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txt_cr_bag 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   64
         Top             =   600
         Width           =   1335
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Bagian 
         Height          =   2055
         Left            =   240
         OleObjectBlob   =   "Frm_Mast_Jabatan.frx":00E4
         TabIndex        =   60
         Top             =   960
         Width           =   5895
      End
      Begin VB.Frame Frame6 
         Height          =   135
         Index           =   0
         Left            =   240
         TabIndex        =   59
         Top             =   360
         Width           =   5895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
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
         Index           =   0
         Left            =   240
         TabIndex        =   62
         Top             =   120
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Bagian :"
         Height          =   195
         Index           =   7
         Left            =   3000
         TabIndex        =   63
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Bagian :"
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   61
         Top             =   600
         Width           =   825
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Hapus 
      Height          =   3855
      Left            =   -3000
      TabIndex        =   24
      Top             =   4440
      Visible         =   0   'False
      Width           =   6255
      _Version        =   65536
      _ExtentX        =   11033
      _ExtentY        =   6800
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Frm_Mast_Jabatan.frx":2DDC
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Frm_Mast_Jabatan.frx":2DF8
      Childs          =   "Frm_Mast_Jabatan.frx":2EA4
      Begin VB.TextBox Txt_Cr_Hapus 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   29
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Txt_Cr_Hapus 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   28
         Top             =   960
         Width           =   1455
      End
      Begin VB.Frame Frame6 
         Height          =   135
         Index           =   3
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   5655
      End
      Begin VB.TextBox Txt_Cr_Hapus 
         Height          =   285
         Index           =   2
         Left            =   4200
         TabIndex        =   26
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox Txt_Cr_Hapus 
         Height          =   285
         Index           =   3
         Left            =   4200
         TabIndex        =   25
         Top             =   960
         Width           =   1575
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Hapus 
         Height          =   2295
         Left            =   240
         OleObjectBlob   =   "Frm_Mast_Jabatan.frx":2EC0
         TabIndex        =   67
         Top             =   1320
         Width           =   5775
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Bagian :"
         Height          =   195
         Index           =   38
         Left            =   360
         TabIndex        =   34
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Bagian :"
         Height          =   195
         Index           =   39
         Left            =   600
         TabIndex        =   33
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
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
         Index           =   3
         Left            =   240
         TabIndex        =   32
         Top             =   120
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kd Jabatan :"
         Height          =   195
         Index           =   14
         Left            =   3240
         TabIndex        =   31
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Jabatan :"
         Height          =   195
         Index           =   17
         Left            =   3000
         TabIndex        =   30
         Top             =   960
         Width           =   1140
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Daftar 
      Height          =   3855
      Left            =   5280
      TabIndex        =   35
      Top             =   3480
      Visible         =   0   'False
      Width           =   6375
      _Version        =   65536
      _ExtentX        =   11245
      _ExtentY        =   6800
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Frm_Mast_Jabatan.frx":65B3
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Frm_Mast_Jabatan.frx":65CF
      Childs          =   "Frm_Mast_Jabatan.frx":667B
      Begin VB.TextBox Txt_Cr_Daftar 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   40
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Txt_Cr_Daftar 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   39
         Top             =   960
         Width           =   1455
      End
      Begin VB.Frame Frame6 
         Height          =   135
         Index           =   2
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   5655
      End
      Begin VB.TextBox Txt_Cr_Daftar 
         Height          =   285
         Index           =   2
         Left            =   4080
         TabIndex        =   37
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Txt_Cr_Daftar 
         Height          =   285
         Index           =   3
         Left            =   4080
         TabIndex        =   36
         Top             =   960
         Width           =   1695
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Daftar 
         Height          =   2295
         Left            =   240
         OleObjectBlob   =   "Frm_Mast_Jabatan.frx":6697
         TabIndex        =   66
         Top             =   1320
         Width           =   5775
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Bagian :"
         Height          =   195
         Index           =   40
         Left            =   360
         TabIndex        =   45
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Bagian :"
         Height          =   195
         Index           =   41
         Left            =   240
         TabIndex        =   44
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
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
         Index           =   2
         Left            =   240
         TabIndex        =   43
         Top             =   120
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kd Jabatan :"
         Height          =   195
         Index           =   13
         Left            =   3120
         TabIndex        =   42
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Jabatan :"
         Height          =   195
         Index           =   16
         Left            =   2880
         TabIndex        =   41
         Top             =   960
         Width           =   1140
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Rubah 
      Height          =   3975
      Left            =   -1680
      TabIndex        =   46
      Top             =   4320
      Visible         =   0   'False
      Width           =   6255
      _Version        =   65536
      _ExtentX        =   11033
      _ExtentY        =   7011
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Frm_Mast_Jabatan.frx":9D8B
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Frm_Mast_Jabatan.frx":9DA7
      Childs          =   "Frm_Mast_Jabatan.frx":9E53
      Begin VB.TextBox Txt_Cr_Rubah 
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   51
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Txt_Cr_Rubah 
         Height          =   315
         Index           =   1
         Left            =   1440
         TabIndex        =   50
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Frame Frame6 
         Height          =   135
         Index           =   1
         Left            =   240
         TabIndex        =   49
         Top             =   360
         Width           =   5655
      End
      Begin VB.TextBox Txt_Cr_Rubah 
         Height          =   315
         Index           =   2
         Left            =   4200
         TabIndex        =   48
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Txt_Cr_Rubah 
         Height          =   315
         Index           =   3
         Left            =   4200
         TabIndex        =   47
         Top             =   1080
         Width           =   1575
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Rubah 
         Height          =   2295
         Left            =   240
         OleObjectBlob   =   "Frm_Mast_Jabatan.frx":9E6F
         TabIndex        =   52
         Top             =   1440
         Width           =   5775
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Bagian :"
         Height          =   195
         Index           =   10
         Left            =   480
         TabIndex        =   57
         Top             =   720
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Bagian :"
         Height          =   195
         Index           =   11
         Left            =   360
         TabIndex        =   56
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
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
         Index           =   1
         Left            =   240
         TabIndex        =   55
         Top             =   120
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kd Jabatan :"
         Height          =   195
         Index           =   12
         Left            =   3240
         TabIndex        =   54
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Jabatan :"
         Height          =   195
         Index           =   18
         Left            =   3000
         TabIndex        =   53
         Top             =   1080
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.CommandButton cmd_browse_bag 
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
         Height          =   230
         Left            =   4080
         TabIndex        =   22
         Top             =   360
         Width           =   255
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
         Left            =   3360
         TabIndex        =   14
         Top             =   3600
         Width           =   4455
         Begin VB.CommandButton Cmd_Keluar 
            Caption         =   "&Keluar"
            Height          =   375
            Left            =   3480
            TabIndex        =   19
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Daftar 
            Caption         =   "&Info"
            Height          =   375
            Left            =   2640
            TabIndex        =   18
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Hapus 
            Caption         =   "&Hapus"
            Height          =   375
            Left            =   1800
            TabIndex        =   17
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Rubah 
            Caption         =   "&Rubah"
            Height          =   375
            Left            =   960
            TabIndex        =   16
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Tambah 
            Caption         =   "&Tambah"
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Simpan 
            Caption         =   "&Simpan"
            Height          =   375
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Batal 
            Caption         =   "&Batal"
            Height          =   375
            Left            =   960
            TabIndex        =   20
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
         Left            =   240
         TabIndex        =   9
         Top             =   3600
         Width           =   2295
         Begin VB.CommandButton Cmd_Navigasi 
            Caption         =   ">>"
            Height          =   375
            Index           =   3
            Left            =   1680
            TabIndex        =   13
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Cmd_Navigasi 
            Caption         =   "<"
            Height          =   375
            Index           =   1
            Left            =   600
            TabIndex        =   11
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Cmd_Navigasi 
            Caption         =   ">"
            Height          =   375
            Index           =   2
            Left            =   1200
            TabIndex        =   12
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Cmd_Navigasi 
            Caption         =   "<<"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.TextBox t_nama_jab 
         Height          =   405
         Left            =   2040
         TabIndex        =   8
         Top             =   1800
         Width           =   4575
      End
      Begin VB.TextBox t_kode_jab 
         Height          =   405
         Left            =   2040
         TabIndex        =   7
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox t_nama_bag 
         Enabled         =   0   'False
         Height          =   405
         Left            =   2040
         TabIndex        =   6
         Top             =   840
         Width           =   4575
      End
      Begin VB.TextBox t_no_bag 
         Height          =   405
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin TDBNumber6Ctl.TDBNumber tdb_lembur 
         Height          =   345
         Left            =   2040
         TabIndex        =   71
         Top             =   2280
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   609
         Calculator      =   "Frm_Mast_Jabatan.frx":D562
         Caption         =   "Frm_Mast_Jabatan.frx":D582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Frm_Mast_Jabatan.frx":D5EE
         Keys            =   "Frm_Mast_Jabatan.frx":D60C
         Spin            =   "Frm_Mast_Jabatan.frx":D656
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
      Begin TDBNumber6Ctl.TDBNumber tdb_makan 
         Height          =   345
         Left            =   2040
         TabIndex        =   72
         Top             =   2640
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   609
         Calculator      =   "Frm_Mast_Jabatan.frx":D67E
         Caption         =   "Frm_Mast_Jabatan.frx":D69E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Frm_Mast_Jabatan.frx":D70A
         Keys            =   "Frm_Mast_Jabatan.frx":D728
         Spin            =   "Frm_Mast_Jabatan.frx":D772
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
      Begin TDBNumber6Ctl.TDBNumber tdb_transport 
         Height          =   345
         Left            =   2040
         TabIndex        =   73
         Top             =   3000
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   609
         Calculator      =   "Frm_Mast_Jabatan.frx":D79A
         Caption         =   "Frm_Mast_Jabatan.frx":D7BA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Frm_Mast_Jabatan.frx":D826
         Keys            =   "Frm_Mast_Jabatan.frx":D844
         Spin            =   "Frm_Mast_Jabatan.frx":D88E
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
         Caption         =   "Uang Transport :"
         Height          =   195
         Index           =   8
         Left            =   720
         TabIndex        =   70
         Top             =   3000
         Width           =   1230
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Uang Makan :"
         Height          =   195
         Index           =   6
         Left            =   960
         TabIndex        =   69
         Top             =   2640
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Uang Lembur :"
         Height          =   195
         Index           =   5
         Left            =   930
         TabIndex        =   68
         Top             =   2280
         Width           =   1050
      End
      Begin VB.Label Lbl_Info 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lbl_Info"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   7320
         TabIndex        =   23
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nama Jabatan :"
         Height          =   195
         Index           =   3
         Left            =   855
         TabIndex        =   4
         Top             =   1800
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Kode Jabatan :"
         Height          =   195
         Index           =   2
         Left            =   900
         TabIndex        =   3
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nama Bagian :"
         Height          =   195
         Index           =   1
         Left            =   975
         TabIndex        =   2
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "No. Bagian :"
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   885
      End
   End
End
Attribute VB_Name = "Frm_Mast_Jabatan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rubah As Boolean
Dim Moving As Boolean
Dim yold, xold As Long

Private Sub isi_semua(ByVal rec As Recordset)
    
    With rec
        If .EOF Then .MoveLast
        If .BOF Then .MoveFirst
        
       t_no_bag.Text = IIf(Not IsNull(!no_bagian), !no_bagian, "")
       t_nama_bag.Text = IIf(Not IsNull(!nama_bagian), !nama_bagian, "")
       t_kode_jab.Text = IIf(Not IsNull(!kode_jab), !kode_jab, "")
       t_nama_jab.Text = IIf(Not IsNull(!nama_jab), !nama_jab, "")
       tdb_lembur.Value = IIf(Not IsNull(!lembur), !lembur, Null)
       tdb_makan.Value = IIf(Not IsNull(!makan), !makan, Null)
       tdb_transport.Value = IIf(Not IsNull(!transport), !transport, Null)
        
        If .RecordCount = 0 Then
            Lbl_Info.Caption = "Record Ke " & 0 & " Dari " & .RecordCount & " Record"
        Else
            Lbl_Info.Caption = "Record Ke " & .AbsolutePosition & " Dari " & .RecordCount & " Record"
        End If
        
    End With
    
End Sub



Private Sub Cmd_Batal_Click()

    Frame_Nav.Enabled = True
    rubah = False
             
        Cmd_Tambah.Visible = True
        Cmd_Tambah.Enabled = True
        Cmd_Simpan.Visible = False
        Cmd_Rubah.Visible = True
        Cmd_Rubah.Enabled = True
        Cmd_Daftar.Enabled = True
        Cmd_Keluar.Enabled = True
        Cmd_Hapus.Enabled = True
        
        Dim n As Object
            
For Each n In Me

        If TypeOf n Is TextBox Then
            If Left(UCase(n.Name), 6) <> UCase("Txt_Cr") Then
                n.Enabled = False
            End If
        End If
        
'        If TypeOf n Is DTPicker Then n.Enabled = False
'        If TypeOf n Is ComboBox Then n.Enabled = False
        
        If TypeOf n Is TDBNumber Then
            n.Enabled = False
        End If
        
        
        If TypeOf n Is TDBContainer3D Then n.Visible = False
'        If TypeOf n Is OptionButton Then n.Enabled = False
'        If TypeOf n Is CheckBox Then n.Enabled = False
        
Next

Set n = Nothing

cmd_browse_bag.Enabled = False

txt_cr_daftar_KeyUp 0, 0, 0
    Cmd_Navigasi_Click 0


 Cmd_Tambah.SetFocus


End Sub

Private Sub cmd_browse_bag_Click()

    With TDB_Bagian
        If .Visible = False Then
            
            .Left = Frame1.Left + cmd_browse_bag.Left + cmd_browse_bag.Width / 2 - .Width / 2
            .Top = Frame1.Top + cmd_browse_bag.Top + cmd_browse_bag.Height + 15
            
            txt_cr_bag(0).Text = ""
            txt_cr_bag(1).Text = ""
            
            txt_cr_bag_KeyUp 0, 0, 0
            .Visible = True
            txt_cr_bag(0).SetFocus
            
        Else
            .Visible = False
        End If
    End With


End Sub

Private Sub Cmd_Daftar_Click()

Frame_Nav.Enabled = False
With TDB_Daftar

If .Visible = False Then

    .Left = Me.Width / 2 - .Width / 2
    .Top = Me.Height / 2 - .Height / 2

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
    .Visible = False
End If

End With


End Sub

Private Sub Cmd_Hapus_Click()

Frame_Nav.Enabled = False
With TDB_Hapus

If .Visible = False Then

    .Left = Me.Width / 2 - .Width / 2
    .Top = Me.Height / 2 - .Height / 2

    Cmd_Tambah.Enabled = False
    Cmd_Rubah.Visible = False
    Cmd_Batal.Visible = True
    Cmd_Hapus.Enabled = False
    Cmd_Daftar.Enabled = False
    Cmd_Keluar.Enabled = False
    
    Txt_Cr_Hapus(0).Text = ""
    Txt_Cr_Hapus(1).Text = ""
    
    Txt_Cr_Hapus_KeyUp 0, 0, 0
    
    .Visible = True
    
    Txt_Cr_Hapus(0).SetFocus
    
Else
    .Visible = False
End If

End With


End Sub

Private Sub Cmd_Keluar_Click()
    Unload Me
End Sub

Private Sub Cmd_Navigasi_Click(Index As Integer)

With Rs_Nav
Select Case Index
    Case 0
        If .RecordCount = 0 Then
            Exit Sub
        Else
            .MoveFirst
        End If
    Case 1
        
        If .BOF Then .MoveFirst
        
        .MovePrevious
        
        If .BOF Then .MoveFirst
        
    Case 2
        
        If .EOF Then .MoveLast
        
        .MoveNext
        
        If .EOF Then .MoveLast
        
    Case 3
        If .RecordCount = 0 Then
            Exit Sub
        Else
            .MoveLast
        End If
End Select
End With

isi_semua Rs_Nav


End Sub

Private Sub Cmd_Rubah_Click()

Frame_Nav.Enabled = False
With TDB_Rubah

If .Visible = False Then
    
    .Left = Me.Width / 2 - .Width / 2
    .Top = Me.Height / 2 - .Height / 2
    
    Cmd_Tambah.Visible = False
    Cmd_Simpan.Visible = True
    Cmd_Simpan.Enabled = False
    Cmd_Rubah.Visible = False
    Cmd_Batal.Visible = True
    Cmd_Hapus.Enabled = False
    Cmd_Daftar.Enabled = False
    Cmd_Keluar.Enabled = False
    
    Txt_Cr_Rubah(0).Text = ""
    Txt_Cr_Rubah(1).Text = ""
    Txt_Cr_Rubah(2).Text = ""
    
    Txt_Cr_Rubah_KeyUp 0, 0, 0
    
    .Visible = True
    
    Txt_Cr_Rubah(0).SetFocus
    
Else
    .Visible = False
End If

End With


End Sub

Private Sub Cmd_Simpan_Click()

On Error GoTo err_handler

kon.BeginTrans

Dim message As Integer
Dim konfirm As Integer

If t_no_bag.Text = "" Then
    message = CInt(MsgBox("No bagian tidak boleh kosong", vbOKOnly + vbInformation, "Informasi"))
    
    t_no_bag.SetFocus
    kon.RollbackTrans
    Exit Sub
End If

If t_kode_jab.Text = "" Then
    message = CInt(MsgBox("Kode Jabatan tidak boleh kosong", vbOKOnly + vbInformation, "Informasi"))
    
    t_kode_jab.SetFocus
    kon.RollbackTrans
    Exit Sub
End If

If t_nama_jab.Text = "" Then
    message = CInt(MsgBox("Nama jabatan tidak boleh kosong", vbOKOnly + vbInformation, "Informasi"))
    
    t_nama_jab.SetFocus
    kon.RollbackTrans
    Exit Sub
End If

Dim sql As String
Dim rs As Recordset

Dim lembur As Double
If tdb_lembur.ValueIsNull Then
    lembur = 0
Else
    lembur = Replace(Trim(tdb_lembur.Value), ",", "")
End If

Dim makan As Double
If tdb_makan.ValueIsNull Then
    makan = 0
Else
    makan = Replace(Trim(tdb_makan.Value), ",", "")
End If

Dim transport As Double
If tdb_transport.ValueIsNull Then
    transport = 0
Else
    transport = Replace(Trim(tdb_transport.Value), ",", "")
End If


If rubah = True Then
    sql = "update tb_jabatan set nama_jab='" & Trim(t_nama_jab.Text) & "',lembur=" & lembur & ",makan=" & makan & ",transport=" & transport & " where kode_jab='" & Trim(t_kode_jab.Text) & "'"
Else
    sql = "insert into tb_jabatan (kode_jab,nama_jab,no_bagian,lembur,makan,transport) values('" & Trim(t_kode_jab.Text) & "','" & Trim(t_nama_jab.Text) & "','" & Trim(t_no_bag.Text) & "'," & lembur & "," & makan & "," & transport & ")"
End If

Set rs = New ADODB.Recordset
    rs.Open sql, kon

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
     Cmd_Daftar.Enabled = False
     Cmd_Keluar.Enabled = False
     Cmd_Hapus.Enabled = False
    
    cmd_browse_bag.Enabled = True
    
    t_no_bag.Enabled = True
    t_no_bag.Text = ""
    t_no_bag.SetFocus

End Sub

Private Sub Form_Activate()
On Error Resume Next
    Cmd_Tambah.SetFocus
End Sub

Private Sub Form_Load()

    With Me
        .Left = Screen.Width / 2 - Me.Width / 2
        .Top = 250
    End With


t_no_bag.Enabled = False
t_nama_bag.Enabled = False
t_kode_jab.Enabled = False
t_nama_jab.Enabled = False
tdb_lembur.Enabled = False
tdb_makan.Enabled = False
tdb_transport.Enabled = False
cmd_browse_bag.Enabled = False

Lbl_Info.Caption = "Record Ke " & 0 & " Dari " & 0 & " Record"

txt_cr_daftar_KeyUp 0, 0, 0
    Cmd_Navigasi_Click 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Cmd_Keluar.Enabled = False Then
        Cancel = True
        Exit Sub
    End If
    
    If kon.State = adStateOpen Then
        
        kon.Close
        Set kon = Nothing
    End If

End Sub

Private Sub Grid_Bagian_DblClick()

On Error GoTo err_handler

If Grid_Bagian.Row < 0 Then Exit Sub

    t_no_bag.Text = Grid_Bagian.Columns(0).Text
    t_nama_bag.Text = Grid_Bagian.Columns(1).Text
    
    TDB_Bagian.Visible = False
    t_no_bag_LostFocus
    
    Exit Sub
    
err_handler:
        
        Dim konfirm As Integer
        
        konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
        Err.Clear

End Sub

Private Sub Grid_Bagian_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Bagian_DblClick
    If KeyCode = vbKeyEscape Then cmd_browse_bag_Click
End Sub

Private Sub grid_daftar_DblClick()
    If Grid_Daftar.Row < 0 Then Exit Sub
    
    Dim nobuk As String
        nobuk = Grid_Daftar.Columns(2).Text
    
    Rs_Nav.MoveFirst
    
    Rs_Nav.Find "kode_jab='" & nobuk & "'"

    isi_semua Rs_Nav
    
    TDB_Daftar.Visible = False
    Frame_Nav.Enabled = True
    Cmd_Navigasi(0).SetFocus

End Sub


Private Sub grid_daftar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then grid_daftar_DblClick
    If KeyCode = vbKeyEscape Then Cmd_Batal_Click
End Sub


Private Sub Grid_Hapus_DblClick()

On Error GoTo err_handler
    
    If Grid_Hapus.Row < 0 Then Exit Sub
    
    kon.BeginTrans
    
    If MsgBox("Yakin akan hapus : " & Grid_Hapus.Columns(2).Text & " ...?", vbYesNo + vbQuestion, "Hapus") = vbNo Then
        kon.RollbackTrans
        On Error GoTo 0
        Exit Sub
    End If


    Dim sql As String
    Dim rs As Recordset
        sql = "delete from Tb_Jabatan where kode_jab='" & Grid_Hapus.Columns(2).Text & "'"
            
            Set rs = New ADODB.Recordset
                rs.Open sql, kon
        
        kon.CommitTrans
        Dim konfirm As Integer
            
            konfirm = CInt(MsgBox(Grid_Hapus.Columns(2).Text & " Berhasil dihapus", vbOKOnly + vbInformation, "Hapus"))
            
            Txt_Cr_Hapus_KeyUp 0, 0, 0
            
'            Cmd_Batal_Click
        
        On Error GoTo 0
        Exit Sub
        
err_handler:
    
    kon.RollbackTrans
    
    konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
        Err.Clear


End Sub

Private Sub Grid_Hapus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Hapus_DblClick
    If KeyCode = vbKeyEscape Then Cmd_Batal_Click
End Sub


Private Sub Grid_Rubah_DblClick()

If Grid_Rubah.Row < 0 Then Exit Sub
    
    Dim nobuk As String
        nobuk = Grid_Rubah.Columns(2).Text
    
    Rs_Nav.MoveFirst
    
    Rs_Nav.Find "kode_jab='" & nobuk & "'"

    isi_semua Rs_Nav
    
    TDB_Rubah.Visible = False
        
    t_nama_jab.Enabled = True
    tdb_lembur.Enabled = True
    tdb_makan.Enabled = True
    tdb_transport.Enabled = True
    
    Cmd_Simpan.Enabled = True
    rubah = True
    
   t_nama_jab.SetFocus


End Sub

Private Sub Grid_Rubah_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Rubah_DblClick
    If KeyCode = vbKeyEscape Then TDB_Rubah.Visible = False: Cmd_Batal_Click
End Sub

Private Sub t_kode_jab_GotFocus()
    Call Focus_(t_kode_jab)
End Sub

Private Sub t_kode_jab_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then t_kode_jab_LostFocus
End Sub

Private Sub t_kode_jab_LostFocus()
    
    If t_kode_jab.Text = "" Then Exit Sub
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select * from Tb_Jabatan where kode_jab='" & Trim(t_kode_jab.Text) & "' and no_bagian='" & Trim(t_no_bag.Text) & "'"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    With rs
        
        If Not .EOF Then
            
            Dim a As Integer
                a = CInt(MsgBox("Kode jabatan yang anda masukkan sudah ada", vbOKOnly + vbInformation, "Informasi"))
                
                t_kode_jab.Text = ""
                t_kode_jab.SetFocus
                
                tdb_lembur.Value = Null
                tdb_makan.Value = Null
                tdb_transport.Value = Null
                
                t_nama_jab.Enabled = False
                tdb_lembur.Enabled = False
                tdb_makan.Enabled = False
                tdb_transport.Enabled = False
                
                Cmd_Simpan.Enabled = False
                
        Else
                
                t_nama_jab.Enabled = True
                tdb_lembur.Enabled = True
                tdb_makan.Enabled = True
                tdb_transport.Enabled = True
                
                Cmd_Simpan.Enabled = True
                t_nama_jab.SetFocus
        
        End If
        
    End With
        
    
    
End Sub

Private Sub t_nama_jab_GotFocus()
    Call Focus_(t_nama_jab)
End Sub

Private Sub t_nama_jab_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        tdb_lembur.SetFocus
        'If Cmd_Simpan.Enabled = True Then Cmd_Simpan.SetFocus
    End If
End Sub

Private Sub t_no_bag_GotFocus()
    Call Focus_(t_no_bag)
End Sub

Private Sub t_no_bag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then t_no_bag_LostFocus
    If KeyCode = vbKeyF3 Then cmd_browse_bag_Click
End Sub

Private Sub t_no_bag_LostFocus()
    
    If t_no_bag.Text = "" Then Exit Sub
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select no_bagian,nama_bagian from Tb_Bagian where no_bagian='" & Trim(t_no_bag.Text) & "'"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon, adOpenKeyset
        
        With rs
            If .EOF Then
                Dim a As Integer
                    a = CInt(MsgBox("No bagian yang anda masukkan tidak ditemukan", vbOKOnly + vbInformation, "Informasi"))
                    
                    t_kode_jab.Enabled = False
'                    t_nama_jab.Enabled = False
                    t_no_bag.Text = ""
                    t_nama_bag.Text = ""
                    tdb_lembur.Value = Null
                    tdb_makan.Value = Null
                    tdb_transport.Value = Null
                    t_no_bag.SetFocus
                    
                    
            Else
                
               t_nama_bag.Text = IIf(Not IsNull(!nama_bagian), !nama_bagian, "")
               
               t_kode_jab.Enabled = True
'               t_nama_jab.Enabled = True
               
               t_kode_jab.Text = ""
               t_nama_jab.Text = ""
               tdb_lembur.Value = Null
                    tdb_makan.Value = Null
                    tdb_transport.Value = Null
               t_kode_jab.SetFocus
                
            End If
        End With
    
End Sub

Private Sub TDB_bagian_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub TDB_bagian_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   TDB_Bagian.Top = TDB_Bagian.Top - (yold - Y)
   TDB_Bagian.Left = TDB_Bagian.Left - (xold - X)
End If

End Sub

Private Sub TDB_bagian_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
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

Private Sub TDB_Hapus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub TDB_Hapus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   TDB_Hapus.Top = TDB_Hapus.Top - (yold - Y)
   TDB_Hapus.Left = TDB_Hapus.Left - (xold - X)
End If

End Sub

Private Sub TDB_Hapus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub

Private Sub tdb_lembur_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        tdb_makan.SetFocus
    End If
End Sub

Private Sub tdb_makan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        tdb_transport.SetFocus
    End If
End Sub

Private Sub TDB_Rubah_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub TDB_Rubah_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   TDB_Rubah.Top = TDB_Rubah.Top - (yold - Y)
   TDB_Rubah.Left = TDB_Rubah.Left - (xold - X)
End If

End Sub

Private Sub TDB_Rubah_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub

Private Sub tdb_transport_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Cmd_Simpan.Enabled = True Then Cmd_Simpan.SetFocus
    End If
End Sub

Private Sub txt_cr_bag_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Bagian.SetFocus
    If KeyCode = vbKeyEscape Then cmd_browse_bag_Click
End Sub

Private Sub txt_cr_bag_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    Dim sql As String
    Dim rs As Recordset
    
        sql = "select top 100 * from Tb_Bagian"
        
    If txt_cr_bag(0).Text <> "" Or txt_cr_bag(1).Text <> "" Then
        
        sql = sql & " where "
        
        Select Case Index
            Case 0
                sql = sql & "no_bagian like '%" & Trim(txt_cr_bag(0).Text) & "%'"
            Case 1
                sql = sql & "nama_bagian like '%" & Trim(txt_cr_bag(1).Text) & "%'"
        End Select
    End If
    
    sql = sql & " order by no_bagian desc"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
    Set Grid_Bagian.DataSource = rs
        Grid_Bagian.Refresh


End Sub

Private Sub txt_cr_daftar_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Daftar.SetFocus
    If KeyCode = vbKeyEscape Then Cmd_Batal_Click
End Sub

Private Sub txt_cr_daftar_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    Dim sql As String
'    Dim rs As Recordset
    
        sql = "select top 100 * from VIEW_Jabatan"
        
    If Txt_Cr_Daftar(0).Text <> "" Or Txt_Cr_Daftar(1).Text <> "" _
        Or Txt_Cr_Daftar(2).Text <> "" Or Txt_Cr_Daftar(3).Text <> "" Then
        
        sql = sql & " where "
        
        Select Case Index
            Case 0
                sql = sql & "no_bagian like '%" & Trim(Txt_Cr_Daftar(0).Text) & "%'"
            Case 1
                sql = sql & "nama_bagian like '%" & Trim(Txt_Cr_Daftar(1).Text) & "%'"
            Case 2
               sql = sql & "kode_jab like '%" & Trim(Txt_Cr_Daftar(2).Text) & "%'"
            Case 3
                sql = sql & "nama_jab like '%" & Trim(Txt_Cr_Daftar(3).Text) & "%'"
        End Select
    End If
    
    sql = sql & " order by kode_jab desc"
    
    Set Rs_Nav = New ADODB.Recordset
        Rs_Nav.Open sql, kon, adOpenKeyset
    
    Set Grid_Daftar.DataSource = Rs_Nav
        Grid_Daftar.Refresh


End Sub


Private Sub Txt_Cr_Hapus_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Hapus.SetFocus
    If KeyCode = vbKeyEscape Then Cmd_Batal_Click
End Sub

Private Sub Txt_Cr_Hapus_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    Dim sql As String
    Dim rs As Recordset
    
        sql = "select top 100 * from VIEW_Jabatan"
        
    If Txt_Cr_Hapus(0).Text <> "" Or Txt_Cr_Hapus(1).Text <> "" _
        Or Txt_Cr_Hapus(2).Text <> "" Or Txt_Cr_Hapus(3).Text <> "" Then
        
        sql = sql & " where "
        
        Select Case Index
            Case 0
                sql = sql & "no_bagian like '%" & Trim(Txt_Cr_Hapus(0).Text) & "%'"
            Case 1
                sql = sql & "nama_bagian like '%" & Trim(Txt_Cr_Hapus(1).Text) & "%'"
            Case 2
               sql = sql & "kode_jab like '%" & Trim(Txt_Cr_Hapus(2).Text) & "%'"
            Case 3
                sql = sql & "nama_jab like '%" & Trim(Txt_Cr_Hapus(3).Text) & "%'"
        End Select
    End If
    
    sql = sql & " order by kode_jab desc"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
    Set Grid_Hapus.DataSource = rs
        Grid_Hapus.Refresh


End Sub

Private Sub Txt_Cr_Rubah_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Rubah.SetFocus
    If KeyCode = vbKeyEscape Then Cmd_Batal_Click
End Sub

Private Sub Txt_Cr_Rubah_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    Dim sql As String
    Dim rs As Recordset
    
        sql = "select top 100 * from VIEW_Jabatan"
        
    If Txt_Cr_Rubah(0).Text <> "" Or Txt_Cr_Rubah(1).Text <> "" _
        Or Txt_Cr_Rubah(2).Text <> "" Or Txt_Cr_Rubah(3).Text <> "" Then
        
        sql = sql & " where "
        
        Select Case Index
            Case 0
                sql = sql & "no_bagian like '%" & Trim(Txt_Cr_Rubah(0).Text) & "%'"
            Case 1
                sql = sql & "nama_bagian like '%" & Trim(Txt_Cr_Rubah(1).Text) & "%'"
            Case 2
               sql = sql & "kode_jab like '%" & Trim(Txt_Cr_Rubah(2).Text) & "%'"
            Case 3
                sql = sql & "nama_jab like '%" & Trim(Txt_Cr_Rubah(3).Text) & "%'"
        End Select
    End If
    
    sql = sql & " order by kode_jab desc"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
    Set Grid_Rubah.DataSource = rs
        Grid_Rubah.Refresh


End Sub
