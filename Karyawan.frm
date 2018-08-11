VERSION 5.00
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form Karyawan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DATA PEGAWAI"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Karyawan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Jabatan 
      Height          =   3615
      Left            =   9120
      TabIndex        =   139
      Top             =   8280
      Visible         =   0   'False
      Width           =   6255
      _Version        =   65536
      _ExtentX        =   11033
      _ExtentY        =   6376
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Karyawan.frx":27C92
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Karyawan.frx":27CAE
      Childs          =   "Karyawan.frx":27D5A
      Begin VB.TextBox txt_cr_jabatan 
         Height          =   300
         Index           =   1
         Left            =   3960
         TabIndex        =   146
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txt_cr_jabatan 
         Height          =   300
         Index           =   0
         Left            =   1320
         TabIndex        =   145
         Top             =   600
         Width           =   1335
      End
      Begin VB.Frame Frame6 
         Height          =   135
         Index           =   5
         Left            =   240
         TabIndex        =   140
         Top             =   360
         Width           =   5655
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Jabatan 
         Height          =   2415
         Left            =   240
         OleObjectBlob   =   "Karyawan.frx":27D76
         TabIndex        =   141
         Top             =   960
         Width           =   5775
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Jabatan :"
         Height          =   195
         Index           =   27
         Left            =   2760
         TabIndex        =   144
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kd Jabatan :"
         Height          =   195
         Index           =   25
         Left            =   360
         TabIndex        =   143
         Top             =   600
         Width           =   915
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
         Index           =   5
         Left            =   240
         TabIndex        =   142
         Top             =   120
         Width           =   1065
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Bagian 
      Height          =   3255
      Left            =   8520
      TabIndex        =   131
      Top             =   8280
      Visible         =   0   'False
      Width           =   6495
      _Version        =   65536
      _ExtentX        =   11456
      _ExtentY        =   5741
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Karyawan.frx":2B46B
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Karyawan.frx":2B487
      Childs          =   "Karyawan.frx":2B533
      Begin VB.Frame Frame6 
         Height          =   135
         Index           =   4
         Left            =   240
         TabIndex        =   135
         Top             =   360
         Width           =   5895
      End
      Begin VB.TextBox txt_cr_bag 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   133
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txt_cr_bag 
         Height          =   285
         Index           =   1
         Left            =   4200
         TabIndex        =   132
         Top             =   600
         Width           =   1695
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Bagian 
         Height          =   2055
         Left            =   240
         OleObjectBlob   =   "Karyawan.frx":2B54F
         TabIndex        =   134
         Top             =   960
         Width           =   5895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Bagian :"
         Height          =   195
         Index           =   15
         Left            =   360
         TabIndex        =   138
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Bagian :"
         Height          =   195
         Index           =   13
         Left            =   3000
         TabIndex        =   137
         Top             =   600
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
         Index           =   4
         Left            =   240
         TabIndex        =   136
         Top             =   120
         Width           =   1065
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Counter 
      Height          =   3975
      Left            =   -840
      TabIndex        =   107
      Top             =   9840
      Visible         =   0   'False
      Width           =   7095
      _Version        =   65536
      _ExtentX        =   12515
      _ExtentY        =   7011
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Karyawan.frx":2E247
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Karyawan.frx":2E263
      Childs          =   "Karyawan.frx":2E30F
      Begin VB.TextBox Txt_Cr_Counter 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3480
         TabIndex        =   114
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox Txt_Cr_Counter 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   113
         Top             =   720
         Width           =   1335
      End
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   3
         Left            =   240
         TabIndex        =   108
         Top             =   480
         Width           =   6495
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Counter 
         Height          =   2655
         Left            =   240
         OleObjectBlob   =   "Karyawan.frx":2E32B
         TabIndex        =   109
         Top             =   1080
         Width           =   6495
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
         Index           =   24
         Left            =   480
         TabIndex        =   112
         Top             =   720
         Width           =   420
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
         Index           =   23
         Left            =   2880
         TabIndex        =   111
         Top             =   720
         Width           =   450
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
         TabIndex        =   110
         Top             =   240
         Width           =   1065
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Daftar 
      Height          =   3855
      Left            =   -960
      TabIndex        =   53
      Top             =   8280
      Visible         =   0   'False
      Width           =   6495
      _Version        =   65536
      _ExtentX        =   11456
      _ExtentY        =   6800
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Karyawan.frx":30DC0
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Karyawan.frx":30DDC
      Childs          =   "Karyawan.frx":30E88
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   2
         Left            =   240
         TabIndex        =   94
         Top             =   360
         Width           =   6015
      End
      Begin VB.TextBox Txt_Cr_Daftar 
         Height          =   285
         Index           =   1
         Left            =   3240
         TabIndex        =   58
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox Txt_Cr_Daftar 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   57
         Top             =   600
         Width           =   1215
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Daftar 
         Height          =   2655
         Left            =   240
         OleObjectBlob   =   "Karyawan.frx":30EA4
         TabIndex        =   54
         Top             =   1080
         Width           =   6015
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   93
         Top             =   120
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   195
         Index           =   41
         Left            =   2640
         TabIndex        =   56
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NoPeg"
         Height          =   195
         Index           =   40
         Left            =   600
         TabIndex        =   55
         Top             =   600
         Width           =   465
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Hapus 
      Height          =   3615
      Left            =   -4680
      TabIndex        =   47
      Top             =   8280
      Visible         =   0   'False
      Width           =   6495
      _Version        =   65536
      _ExtentX        =   11456
      _ExtentY        =   6376
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Karyawan.frx":33E14
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Karyawan.frx":33E30
      Childs          =   "Karyawan.frx":33EDC
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   1
         Left            =   240
         TabIndex        =   92
         Top             =   360
         Width           =   5895
      End
      Begin VB.TextBox Txt_Cr_Hapus 
         Height          =   285
         Index           =   1
         Left            =   3240
         TabIndex        =   52
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox Txt_Cr_Hapus 
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   51
         Top             =   600
         Width           =   1215
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Hapus 
         Height          =   2535
         Left            =   240
         OleObjectBlob   =   "Karyawan.frx":33EF8
         TabIndex        =   48
         Top             =   960
         Width           =   6015
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   91
         Top             =   120
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NoPeg"
         Height          =   195
         Index           =   39
         Left            =   360
         TabIndex        =   50
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   195
         Index           =   38
         Left            =   2640
         TabIndex        =   49
         Top             =   600
         Width           =   405
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Rubah 
      Height          =   3735
      Left            =   -2640
      TabIndex        =   41
      Top             =   8280
      Visible         =   0   'False
      Width           =   6255
      _Version        =   65536
      _ExtentX        =   11033
      _ExtentY        =   6588
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Karyawan.frx":36E67
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Karyawan.frx":36E83
      Childs          =   "Karyawan.frx":36F2F
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   0
         Left            =   240
         TabIndex        =   90
         Top             =   360
         Width           =   5775
      End
      Begin VB.TextBox Txt_Cr_Rubah 
         Height          =   315
         Index           =   1
         Left            =   3240
         TabIndex        =   43
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox Txt_Cr_Rubah 
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   42
         Top             =   600
         Width           =   1215
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Rubah 
         Height          =   2655
         Left            =   240
         OleObjectBlob   =   "Karyawan.frx":36F4B
         TabIndex        =   44
         Top             =   960
         Width           =   5775
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   89
         Top             =   120
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   195
         Index           =   37
         Left            =   2760
         TabIndex        =   46
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NoPeg"
         Height          =   195
         Index           =   36
         Left            =   360
         TabIndex        =   45
         Top             =   600
         Width           =   465
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Pendidikan 
      Height          =   2295
      Left            =   -3600
      TabIndex        =   34
      Top             =   8520
      Visible         =   0   'False
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   4048
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Karyawan.frx":39EBA
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Karyawan.frx":39ED6
      Childs          =   "Karyawan.frx":39F82
      Begin VB.TextBox Txt_Cr_Pendidikan 
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   35
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Txt_Cr_Pendidikan 
         Height          =   285
         Index           =   1
         Left            =   3240
         TabIndex        =   36
         Top             =   240
         Width           =   2175
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Pendidikan 
         Height          =   1455
         Left            =   120
         OleObjectBlob   =   "Karyawan.frx":39F9E
         TabIndex        =   37
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         Height          =   195
         Index           =   33
         Left            =   240
         TabIndex        =   39
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pendidikan"
         Height          =   195
         Index           =   32
         Left            =   2400
         TabIndex        =   38
         Top             =   240
         Width           =   765
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Status_P 
      Height          =   2295
      Left            =   3840
      TabIndex        =   28
      Top             =   8520
      Visible         =   0   'False
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   4048
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Karyawan.frx":3C8FE
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Karyawan.frx":3C91A
      Childs          =   "Karyawan.frx":3C9C6
      Begin VB.TextBox Txt_Cr_Status 
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   29
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Txt_Cr_Status 
         Height          =   285
         Index           =   1
         Left            =   3000
         TabIndex        =   30
         Top             =   240
         Width           =   2415
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Status 
         Height          =   1455
         Left            =   120
         OleObjectBlob   =   "Karyawan.frx":3C9E2
         TabIndex        =   31
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         Height          =   195
         Index           =   31
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   195
         Index           =   30
         Left            =   2400
         TabIndex        =   32
         Top             =   240
         Width           =   465
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Agama 
      Height          =   2295
      Left            =   0
      TabIndex        =   22
      Top             =   8280
      Visible         =   0   'False
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   4048
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Karyawan.frx":3F346
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Karyawan.frx":3F362
      Childs          =   "Karyawan.frx":3F40E
      Begin VB.TextBox Txt_Cr_Agama 
         Height          =   285
         Index           =   1
         Left            =   3000
         TabIndex        =   24
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox Txt_Cr_Agama 
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   23
         Top             =   240
         Width           =   1575
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Agama 
         Height          =   1455
         Left            =   120
         OleObjectBlob   =   "Karyawan.frx":3F42A
         TabIndex        =   25
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agama"
         Height          =   195
         Index           =   29
         Left            =   2400
         TabIndex        =   27
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         Height          =   195
         Index           =   28
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   6720
         TabIndex        =   20
         Top             =   0
         Width           =   2655
         Begin VB.Label Lbl_Umur 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   600
            TabIndex        =   40
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Umur :"
            Height          =   195
            Index           =   16
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame Frame9 
         Height          =   615
         Left            =   7200
         TabIndex        =   156
         Top             =   600
         Width           =   2175
         Begin VB.Label lblmasakerja 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1440
            TabIndex        =   158
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sisa Masa Kerja  :"
            Height          =   195
            Index           =   34
            Left            =   120
            TabIndex        =   157
            Top             =   240
            Width           =   1275
         End
      End
      Begin VB.CommandButton cmd_del2 
         Caption         =   "-"
         Height          =   315
         Left            =   480
         TabIndex        =   152
         Top             =   6720
         Width           =   375
      End
      Begin VB.CommandButton cmd_add2 
         Caption         =   "+"
         Height          =   315
         Left            =   120
         TabIndex        =   153
         Top             =   6720
         Width           =   375
      End
      Begin VB.TextBox tkode_abs 
         Height          =   315
         Left            =   5520
         TabIndex        =   150
         Top             =   3240
         Width           =   1695
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
         Height          =   230
         Left            =   13800
         TabIndex        =   130
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
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
         Left            =   13800
         TabIndex        =   129
         Top             =   960
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox t_peringkat 
         Height          =   315
         Left            =   12120
         TabIndex        =   128
         Top             =   2760
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox t_jenjang 
         Height          =   315
         Left            =   12120
         TabIndex        =   127
         Top             =   2400
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox t_nama_jabatan 
         Height          =   315
         Left            =   12120
         TabIndex        =   124
         Top             =   2040
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.TextBox t_kd_jabatan 
         Height          =   315
         Left            =   12120
         TabIndex        =   123
         Top             =   1680
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox t_nama_bagian 
         Height          =   315
         Left            =   12120
         TabIndex        =   122
         Top             =   1320
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.TextBox t_no_bagian 
         Height          =   315
         Left            =   12120
         TabIndex        =   121
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   101
         Top             =   10920
         Visible         =   0   'False
         Width           =   4695
         Begin VB.CommandButton Cmd_Browse_Counter 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   230
            Left            =   4200
            TabIndex        =   105
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Lbl_Kode_Counter 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   106
            Top             =   120
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Lbl_Nama_Counter 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   1320
            TabIndex        =   104
            Top             =   360
            Width           =   2760
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cabang"
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
            Index           =   22
            Left            =   120
            TabIndex        =   103
            Top             =   360
            Width           =   600
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
            Index           =   21
            Left            =   1200
            TabIndex        =   102
            Top             =   360
            Width           =   60
         End
      End
      Begin VB.TextBox Txt_Kode_Agama 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   88
         Top             =   1800
         Width           =   495
      End
      Begin VB.ComboBox Cbo_Agama 
         Height          =   315
         ItemData        =   "Karyawan.frx":41D81
         Left            =   2040
         List            =   "Karyawan.frx":41D83
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Frame Frame_Nav 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   81
         Top             =   7560
         Width           =   2175
         Begin VB.CommandButton Cmd_Navigasi 
            Caption         =   ">>"
            Height          =   495
            Index           =   3
            Left            =   1560
            TabIndex        =   85
            Top             =   120
            Width           =   495
         End
         Begin VB.CommandButton Cmd_Navigasi 
            Caption         =   ">"
            Height          =   495
            Index           =   2
            Left            =   1080
            TabIndex        =   84
            Top             =   120
            Width           =   495
         End
         Begin VB.CommandButton Cmd_Navigasi 
            Caption         =   "<"
            Height          =   495
            Index           =   1
            Left            =   600
            TabIndex        =   83
            Top             =   120
            Width           =   495
         End
         Begin VB.CommandButton Cmd_Navigasi 
            Caption         =   "<<"
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   82
            Top             =   120
            Width           =   495
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   8520
         TabIndex        =   63
         Top             =   8160
         Visible         =   0   'False
         Width           =   4695
         Begin VB.Frame Frame7 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   120
            TabIndex        =   95
            Top             =   600
            Width           =   4455
            Begin VB.OptionButton Opt_Hari 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Per&Hari"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2520
               TabIndex        =   100
               Top             =   650
               Width           =   975
            End
            Begin VB.OptionButton Opt_Bulan 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Per&Bulan"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   99
               Top             =   650
               Width           =   975
            End
            Begin TDBNumber6Ctl.TDBNumber TDB_Gaji 
               Height          =   320
               Left            =   1320
               TabIndex        =   96
               Top             =   240
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
               _ExtentY        =   564
               Calculator      =   "Karyawan.frx":41D85
               Caption         =   "Karyawan.frx":41DA5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "Karyawan.frx":41E11
               Keys            =   "Karyawan.frx":41E2F
               Spin            =   "Karyawan.frx":41E79
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,###;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
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
               ShowContextMenu =   -1
               ValueVT         =   1
               Value           =   0
               MaxValueVT      =   1028849669
               MinValueVT      =   1598423045
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
               Index           =   20
               Left            =   1200
               TabIndex        =   98
               Top             =   240
               Width           =   60
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Gaji Pokok"
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
               Index           =   26
               Left            =   120
               TabIndex        =   97
               Top             =   240
               Width           =   840
            End
         End
         Begin VB.TextBox Txt_Ket 
            Height          =   765
            Left            =   1440
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   72
            Top             =   2400
            Width           =   3015
         End
         Begin TDBNumber6Ctl.TDBNumber TDB_Tunjangan 
            Height          =   320
            Left            =   1440
            TabIndex        =   66
            Top             =   1680
            Width           =   3015
            _Version        =   65536
            _ExtentX        =   5318
            _ExtentY        =   564
            Calculator      =   "Karyawan.frx":41EA1
            Caption         =   "Karyawan.frx":41EC1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Karyawan.frx":41F2D
            Keys            =   "Karyawan.frx":41F4B
            Spin            =   "Karyawan.frx":41F95
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###;;0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
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
            ShowContextMenu =   -1
            ValueVT         =   1
            Value           =   0
            MaxValueVT      =   1028849669
            MinValueVT      =   1598423045
         End
         Begin TDBNumber6Ctl.TDBNumber TDB_Uang_Makan 
            Height          =   320
            Left            =   1440
            TabIndex        =   69
            Top             =   2040
            Width           =   3015
            _Version        =   65536
            _ExtentX        =   5318
            _ExtentY        =   564
            Calculator      =   "Karyawan.frx":41FBD
            Caption         =   "Karyawan.frx":41FDD
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Karyawan.frx":42049
            Keys            =   "Karyawan.frx":42067
            Spin            =   "Karyawan.frx":420B1
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###;;0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
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
            ShowContextMenu =   -1
            ValueVT         =   2089877505
            Value           =   0
            MaxValueVT      =   1028849669
            MinValueVT      =   1598423045
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
            Index           =   59
            Left            =   1320
            TabIndex        =   71
            Top             =   2400
            Width           =   60
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ket"
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
            Index           =   58
            Left            =   240
            TabIndex        =   70
            Top             =   2400
            Width           =   285
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
            Index           =   49
            Left            =   1320
            TabIndex        =   68
            Top             =   2040
            Width           =   60
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Uang Makan"
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
            Index           =   48
            Left            =   240
            TabIndex        =   67
            Top             =   2040
            Width           =   990
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
            Index           =   47
            Left            =   1320
            TabIndex        =   65
            Top             =   1680
            Width           =   60
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tunjangan"
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
            Index           =   46
            Left            =   240
            TabIndex        =   64
            Top             =   1680
            Width           =   870
         End
      End
      Begin VB.TextBox Txt_Kode_Jenis_Kelamin 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   61
         Top             =   1440
         Width           =   495
      End
      Begin VB.ComboBox Cbo_Jenis_Kelamin 
         Height          =   315
         ItemData        =   "Karyawan.frx":420D9
         Left            =   2040
         List            =   "Karyawan.frx":420DB
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   1440
         Width           =   2415
      End
      Begin TDBDate6Ctl.TDBDate TDB_Tgl_Lhr 
         Height          =   315
         Left            =   5640
         TabIndex        =   19
         Top             =   1080
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   556
         Calendar        =   "Karyawan.frx":420DD
         Caption         =   "Karyawan.frx":421F5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Karyawan.frx":42261
         Keys            =   "Karyawan.frx":4227F
         Spin            =   "Karyawan.frx":422DD
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
      Begin VB.TextBox Txt_Tempat_Lhr 
         Height          =   315
         Left            =   1560
         TabIndex        =   17
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox Txt_Kodepos 
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
         Left            =   1560
         TabIndex        =   10
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox Txt_Alamat_3 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   2880
         Width           =   5655
      End
      Begin VB.TextBox Txt_Alamat_2 
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   2520
         Width           =   5655
      End
      Begin VB.TextBox Txt_Alamat_1 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   2160
         Width           =   5655
      End
      Begin VB.TextBox Txt_Nama 
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   720
         Width           =   5415
      End
      Begin VB.TextBox Txt_Kode 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin TDBDate6Ctl.TDBDate TDB_Tgl_Masuk 
         Height          =   315
         Left            =   1560
         TabIndex        =   115
         Top             =   4320
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   547
         Calendar        =   "Karyawan.frx":42305
         Caption         =   "Karyawan.frx":4241D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Karyawan.frx":42489
         Keys            =   "Karyawan.frx":424A7
         Spin            =   "Karyawan.frx":42505
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
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "No. Telp"
         Height          =   855
         Left            =   360
         TabIndex        =   11
         Top             =   3480
         Width           =   3615
         Begin VB.TextBox Txt_Telp_Hp 
            Height          =   285
            Left            =   1200
            TabIndex        =   15
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox Txt_Telp_Rumah 
            Height          =   285
            Left            =   1200
            TabIndex        =   14
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Telp Hp :"
            Height          =   195
            Index           =   9
            Left            =   255
            TabIndex        =   13
            Top             =   480
            Width           =   885
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Telp Rumah :"
            Height          =   195
            Index           =   8
            Left            =   15
            TabIndex        =   12
            Top             =   120
            Width           =   1185
         End
      End
      Begin TDBNumber6Ctl.TDBNumber tdb_gapok 
         Height          =   345
         Left            =   12120
         TabIndex        =   148
         Top             =   3120
         Visible         =   0   'False
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   609
         Calculator      =   "Karyawan.frx":4252D
         Caption         =   "Karyawan.frx":4254D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Karyawan.frx":425B9
         Keys            =   "Karyawan.frx":425D7
         Spin            =   "Karyawan.frx":42621
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
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4800
         TabIndex        =   73
         Top             =   7560
         Width           =   4455
         Begin VB.CommandButton Cmd_Keluar 
            Caption         =   "&Keluar"
            Height          =   495
            Left            =   3480
            TabIndex        =   78
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Daftar 
            Caption         =   "&Daftar"
            Height          =   495
            Left            =   2640
            TabIndex        =   77
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Hapus 
            Caption         =   "&Hapus"
            Height          =   495
            Left            =   1800
            TabIndex        =   76
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Rubah 
            Caption         =   "&Rubah"
            Height          =   495
            Left            =   960
            TabIndex        =   75
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Tambah 
            Caption         =   "&Tambah"
            Height          =   495
            Left            =   120
            TabIndex        =   74
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdnonaktif 
            Caption         =   "&Non Aktifkan Pegawai"
            Height          =   495
            Left            =   1800
            TabIndex        =   154
            Top             =   120
            Width           =   1695
         End
         Begin VB.CommandButton Cmd_Batal 
            Caption         =   "&Batal"
            Height          =   495
            Left            =   960
            TabIndex        =   79
            Top             =   120
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Simpan 
            Caption         =   "&Simpan"
            Height          =   495
            Left            =   120
            TabIndex        =   80
            Top             =   120
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin TrueOleDBGrid60.TDBGrid grid_jb 
         Height          =   2055
         Left            =   120
         OleObjectBlob   =   "Karyawan.frx":42649
         TabIndex        =   151
         Top             =   4680
         Width           =   9135
      End
      Begin VB.Label lblstatus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STATUS : NON AKTIF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   5520
         TabIndex        =   155
         Top             =   4080
         Width           =   3300
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   8400
         TabIndex        =   62
         Top             =   7320
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Absen :"
         Height          =   195
         Index           =   19
         Left            =   4560
         TabIndex        =   149
         Top             =   3240
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gaji Pokok :"
         Height          =   195
         Index           =   17
         Left            =   11280
         TabIndex        =   147
         Top             =   3120
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peringkat :"
         Height          =   195
         Index           =   11
         Left            =   11760
         TabIndex        =   126
         Top             =   2760
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenjang Jabatan :"
         Height          =   195
         Index           =   10
         Left            =   11640
         TabIndex        =   125
         Top             =   2520
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Jabatan :"
         Height          =   195
         Index           =   7
         Left            =   11160
         TabIndex        =   120
         Top             =   2040
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kd Jabatan :"
         Height          =   195
         Index           =   5
         Left            =   11115
         TabIndex        =   119
         Top             =   1680
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Bagian :"
         Height          =   195
         Index           =   4
         Left            =   11160
         TabIndex        =   118
         Top             =   1320
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kd Bagian :"
         Height          =   195
         Index           =   1
         Left            =   11220
         TabIndex        =   117
         Top             =   960
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Masuk :"
         Height          =   195
         Index           =   50
         Left            =   720
         TabIndex        =   116
         Top             =   4320
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agama :"
         Height          =   195
         Index           =   18
         Left            =   960
         TabIndex        =   86
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelamin :"
         Height          =   195
         Index           =   42
         Left            =   495
         TabIndex        =   59
         Top             =   1440
         Width           =   1050
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Lhr :"
         Height          =   195
         Index           =   14
         Left            =   5040
         TabIndex        =   18
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat Lhr :"
         Height          =   195
         Index           =   12
         Left            =   600
         TabIndex        =   16
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Pos :"
         Height          =   195
         Index           =   6
         Left            =   720
         TabIndex        =   9
         Top             =   3240
         Width           =   765
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat :"
         Height          =   195
         Index           =   3
         Left            =   960
         TabIndex        =   4
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama :"
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   3
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Peg :"
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "Karyawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rubah As Boolean
Dim Moving As Boolean
Dim yold, xold As Long
Dim Arr_Rubah As New XArrayDB
Dim Arr_Hapus As New XArrayDB
Dim arr_daftar As New XArrayDB
Dim arr_jb As New XArrayDB

Public Sub set_setelah_edit_status()
    lblstatus.Caption = "STATUS : NON-AKTIF"
    Cmd_Batal_Click
End Sub

Private Sub Kosong_Rubah()

    Arr_Rubah.ReDim 0, 0, 0, 0
    Arr_Rubah.ReDim 1, 1, 1, 1
    Grid_Rubah.ReBind
    Grid_Rubah.Refresh
    
End Sub

Private Sub Kosong_Hapus()
    Arr_Hapus.ReDim 0, 0, 0, 0
    Arr_Hapus.ReDim 1, 1, 1, 1
    Grid_Hapus.ReBind
    Grid_Hapus.Refresh
End Sub

Private Sub kosong_daftar()
    arr_daftar.ReDim 0, 0, 0, 0
    arr_daftar.ReDim 1, 1, 1, 1
    Grid_Daftar.ReBind
    Grid_Daftar.Refresh
End Sub

Private Sub kosong_jb()
    arr_jb.ReDim 0, 0, 0, 0
    arr_jb.ReDim 1, 1, 1, grid_jb.Columns.Count
    grid_jb.ReBind
    grid_jb.Refresh
End Sub

Private Sub isi_jb()

    Dim sql As String
    Dim rs As Recordset
    
        sql = "select * from VIEW_Karyawan_Jab where kd_karyawan='" & Trim(Txt_Kode.Text) & "'"
        
    sql = sql & " order by no_id asc"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
    Isi_grid_jb rs
    
End Sub

Public Sub tambah_jab(ByVal nobagian As String, ByVal namabagian As String, ByVal kdjab As String, ByVal namajab As String, _
                ByVal ttgl As String, ByVal ket As String, ByVal gapok As Double)
                
     Dim i As Integer
    If arr_jb.UpperBound(1) = 1 And arr_jb(1, 1) = Empty Then
        i = 1
    Else
        i = arr_jb.UpperBound(1) + 1
    End If
    
    arr_jb.ReDim 1, i, 0, grid_jb.Columns.Count
        grid_jb.ReBind
        grid_jb.Refresh
    
    arr_jb(i, 0) = nobagian
    arr_jb(i, 1) = namabagian
    arr_jb(i, 2) = kdjab
    arr_jb(i, 3) = namajab
    arr_jb(i, 4) = ttgl
    arr_jb(i, 5) = gapok
    arr_jb(i, 6) = ket
    arr_jb(i, 7) = 0
    
    grid_jb.ReBind
    grid_jb.Refresh
    
End Sub

Private Sub Cbo_Agama_Change()
    With Cbo_Agama
        Txt_Kode_Agama.Text = Left(.Text, 2)
    End With
End Sub

Private Sub Cbo_Agama_Click()
    Cbo_Agama_Change
End Sub

Private Sub Cbo_Agama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Alamat_1.SetFocus
End Sub

Private Sub Cbo_Jenis_Kelamin_Change()
    With Cbo_Jenis_Kelamin
        Txt_Kode_Jenis_Kelamin.Text = Left(.Text, 2)
    End With
End Sub

Private Sub Cbo_Jenis_Kelamin_Click()
    Cbo_Jenis_Kelamin_Change
End Sub

Private Sub cbo_jenis_kelamin_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
    Cbo_Agama.SetFocus
End If

End Sub

Private Sub cmd_add2_Click()
    Karyawan2.Show 1
End Sub

Private Sub Cmd_Batal_Click()
    
    If rubah <> True Then Lbl_Umur.Caption = "0 Thn"
    
    Frame_Nav.Enabled = True
    rubah = False
             
        Cmd_Tambah.Visible = True
        
        Cmd_Tambah.Enabled = True
    
        Cmd_Simpan.Visible = False
        Cmd_Rubah.Visible = True
        Cmd_Rubah.Enabled = True
        Cmd_Hapus.Visible = True
        Cmd_Daftar.Visible = True
        Cmd_Hapus.Enabled = True
        Cmd_Daftar.Enabled = True
        Cmd_Keluar.Enabled = True
        cmdnonaktif.Visible = False
        
        cmd_add2.Enabled = False
        cmd_del2.Enabled = False
        
        Dim n As Object
            
For Each n In Me

        If TypeOf n Is TextBox Then
            If Left(UCase(n.Name), 6) <> UCase("Txt_Cr") Then
                n.Enabled = False
            End If
        End If
        
        If TypeOf n Is TDBDate Then n.Enabled = False
        If TypeOf n Is ComboBox Then n.Enabled = False
        If TypeOf n Is MaskEdBox Then n.Enabled = False
        
        If TypeOf n Is TDBNumber Then
            n.Enabled = False
        End If
        
        
        If TypeOf n Is TDBContainer3D Then n.Visible = False
        If TypeOf n Is OptionButton Then n.Enabled = False
        
        If TypeOf n Is CommandButton Then
            
            If UCase(Left(n.Name, 10)) = UCase("cmd_browse") Then
                n.Enabled = False
            End If
            
        End If
        
        

Next

Set n = Nothing

 If Cmd_Tambah.Enabled = True Then Cmd_Tambah.SetFocus

 txt_cr_daftar_KeyUp 0, 0, 0
 Cmd_Navigasi_Click 3
    
End Sub

Private Sub cmd_browse_bag_Click()

    With TDB_bagian
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

Private Sub Cmd_Browse_Counter_Click()

With TDB_Counter
    .Left = 2640
    .Top = 1200
    
    If .Visible = False Then
    
    Txt_Cr_Counter(0).Text = ""
    Txt_Cr_Counter(1).Text = ""
    
    Txt_Cr_Counter_KeyUp 0, 0, 0
    
    .Visible = True
    
    Txt_Cr_Counter(0).SetFocus
    
    Else
        .Visible = False
    End If
    
End With

End Sub

Private Sub cmd_browse_jab_Click()
    
    If t_no_bagian.Text = "" Then Exit Sub
    
    With TDB_Jabatan
        If .Visible = False Then
            
            .Left = Frame1.Left + cmd_browse_jab.Left + cmd_browse_jab.Width / 2 - .Width / 2
            .Top = Frame1.Top + cmd_browse_jab.Top + cmd_browse_jab.Height + 15
            
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

Private Sub Cmd_Daftar_Click()

Frame_Nav.Enabled = False
With TDB_Daftar

If .Visible = False Then
    
    Cmd_Tambah.Enabled = False
    Cmd_Rubah.Visible = False
    Cmd_Batal.Visible = True
    Cmd_Hapus.Enabled = False
    Cmd_Daftar.Enabled = False
    Cmd_Keluar.Enabled = False
    
    cmd_add2.Enabled = False
    cmd_del2.Enabled = False
    
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

Private Sub cmd_del2_Click()
On Error GoTo err_handler

    Dim konfirm As Integer

    If arr_jb.UpperBound(1) = 1 And arr_jb(1, 1) = Empty Then Exit Sub
    
    If arr_jb(grid_jb.Bookmark, 6) <> 0 Then
        
        Dim rs As Recordset
        Dim sql As String
            sql = "delete from Tb_Karyawan_Jab where no_id=" & arr_jb(grid_jb.Bookmark, 7)
        
        kon.BeginTrans
        
        Set rs = New ADODB.Recordset
        rs.Open sql, kon
        
        kon.CommitTrans
        
    End If
    
    If arr_jb.UpperBound(1) > 1 Then
        grid_jb.Delete
      Else
        arr_jb.ReDim 0, 0, 0, 0
        arr_jb.ReDim 1, 1, 1, grid_jb.Columns.Count
      End If
        
        grid_jb.ReBind
        grid_jb.Refresh
    
    On Error GoTo 0
    Exit Sub
    
err_handler:
        
            konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
                Err.Clear
End Sub

Private Sub Cmd_Hapus_Click()

Frame_Nav.Enabled = False
With TDB_Hapus

If .Visible = False Then
    
    Cmd_Tambah.Enabled = False
    Cmd_Rubah.Visible = False
    Cmd_Batal.Visible = True
    Cmd_Hapus.Enabled = False
    Cmd_Daftar.Enabled = False
    Cmd_Keluar.Enabled = False
    
    cmd_add2.Enabled = False
    cmd_del2.Enabled = False
    
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

Private Sub Cmd_Rubah_Click()

Frame_Nav.Enabled = False
With TDB_Rubah

If .Visible = False Then
    
    
    Cmd_Tambah.Visible = False
    Cmd_Simpan.Visible = True
    Cmd_Simpan.Enabled = False
    Cmd_Rubah.Visible = False
    Cmd_Batal.Visible = True
    Cmd_Hapus.Visible = False
    Cmd_Daftar.Visible = False
    Cmd_Keluar.Enabled = False
    cmdnonaktif.Visible = True
    
'    cmd_add2.Enabled = True
'    cmd_del2.Enabled = True
    cmdnonaktif.Enabled = False
    
    Txt_Cr_Rubah(0).Text = ""
    Txt_Cr_Rubah(1).Text = ""
    
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

Dim konfirm As Integer
            
            Dim gapok As Double
            If tdb_gapok.ValueIsNull Then
                gapok = 0
            Else
                gapok = Replace(Trim(tdb_gapok.Value), ",", "")
            End If
    
Dim tgl_keluar As String
'    If ttgl_keluar.Text = "__/__/____" Then
        tgl_keluar = ""
'    Else
'        tgl_keluar = Trim(ttgl_keluar.Text)
'    End If
    
'            Dim Tunj As Double
'            If TDB_Tunjangan.ValueIsNull Then
'               Tunj = 0
'            Else
'                Tunj = Replace(Trim(TDB_Tunjangan.Value), ",", "")
'            End If
'
'            Dim Uang_Mkan As Double
'            If TDB_Uang_Makan.ValueIsNull Then
'                Uang_Mkan = 0
'            Else
'                Uang_Mkan = Replace(Trim(TDB_Uang_Makan.Value), ",", "")
'            End If
            
kon.BeginTrans
Dim sql, sql1 As String
Dim rs As Recordset
Dim rs1 As Recordset
            
'Dim fl_gaji As String
'    If Opt_Bulan.Value = True Then
'        fl_gaji = "b"
'    ElseIf Opt_Hari.Value = True Then
'        fl_gaji = "h"
'    End If
            
If rubah = False Then
    
    If Txt_Kode.Text = "" Then
        konfirm = CInt(MsgBox("Kode karyawan tidak boleh kosong", vbOKOnly + vbInformation, "Informasi"))
        Txt_Kode.SetFocus
        
        On Error GoTo 0
        Exit Sub
    Else
    
    sql1 = "select Kode_Karyawan from Tb_Karyawan where Kode_Karyawan='" & Trim(Txt_Kode.Text) & "'"
    
    Set rs1 = New ADODB.Recordset
        rs1.Open sql1, kon
    
    With rs1
        If Not .EOF Then
            konfirm = CInt(MsgBox("Kode karyawan yang anda masukkan sudah ada", vbOKOnly + vbInformation, "Informasi"))
            
            Txt_Kode.SetFocus
            
            kon.RollbackTrans
            On Error GoTo 0
            Exit Sub
        End If
    End With
        
        sql = "insert into Tb_Karyawan (Kode_Karyawan,Nama_Karyawan,Jenis_Kelamin,Agama,Alamat_1,Alamat_2,Alamat_3,Kode_Pos,No_Telp,No_Telp_HP,Tempat_Lhr,Tgl_Lhr,Tgl_Masuk,Jml_Hutang,kd_jab,jenjang_jbt,peringkat,kd_absen,gapok,tgl_keluar)"
        sql = sql & " values('" & Trim(Txt_Kode.Text) & "','" & Trim(Txt_Nama.Text) & "','" & Trim(Txt_Kode_Jenis_Kelamin.Text) & "','" & Trim(Txt_Kode_Agama.Text) & "','" & Trim(Txt_Alamat_1.Text) & "','" & Trim(Txt_Alamat_2.Text) & "','" & Trim(Txt_Alamat_3.Text) & "','" & Trim(Txt_Kodepos.Text) & "'"
        sql = sql & ",'" & Trim(Txt_Telp_Rumah.Text) & "','" & Trim(Txt_Telp_Hp.Text) & "','" & Trim(Txt_Tempat_Lhr.Text) & "','" & Format(Trim(TDB_Tgl_Lhr.Text), "yyyy/mm/dd") & "','" & Format(Trim(TDB_Tgl_Masuk.Text), "yyyy/mm/dd") & "',0,'" & Trim(t_kd_jabatan.Text) & "','" & Trim(t_jenjang.Text) & "','" & Trim(t_peringkat.Text) & "','" & Trim(tkode_abs.Text) & "'," & gapok
        
        If Len(tgl_keluar) = 0 Then
            sql = sql & ",NULL)"
        Else
        sql = sql & ",'" & Format(Trim(tgl_keluar), "yyyy/mm/dd") & "')"
        End If
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon
        
        simpan_detail
        
        kon.CommitTrans
        
        konfirm = CInt(MsgBox("Data karyawan telah disimpan ...", vbOKOnly + vbInformation, "Informasi"))
        
        Cmd_Batal_Click
        
    End If
    
Else

    sql = "update Tb_Karyawan set Nama_Karyawan='" & Trim(Txt_Nama.Text) & "',Jenis_Kelamin='" & Trim(Txt_Kode_Jenis_Kelamin.Text) & "',Alamat_1='" & Trim(Txt_Alamat_1.Text) & "',Alamat_2='" & Trim(Txt_Alamat_2.Text) & "',Alamat_3='" & Trim(Txt_Alamat_3.Text) & "',"
    sql = sql & "Kode_Pos='" & Trim(Txt_Kodepos.Text) & "',No_Telp_Hp='" & Trim(Txt_Telp_Hp.Text) & "',No_Telp='" & Trim(Txt_Telp_Rumah.Text) & "',Agama='" & Trim(Txt_Kode_Agama.Text) & "',Tempat_Lhr='" & Trim(Txt_Tempat_Lhr.Text) & "',Tgl_Lhr='" & Format(Trim(TDB_Tgl_Lhr.Text), "yyyy/mm/dd") & "',"
    sql = sql & "Tgl_Masuk='" & Format(Trim(TDB_Tgl_Masuk.Text), "yyyy/mm/dd") & "',kd_jab='" & Trim(t_kd_jabatan.Text) & "',jenjang_jbt='" & Trim(t_jenjang.Text) & "',Peringkat='" & Trim(t_peringkat.Text) & "',kd_absen='" & Trim(tkode_abs.Text) & "',gapok=" & gapok
    'sql = sql & ",tgl_keluar=" & IIf(Len(tgl_keluar) = 0, Null, "'" & Format(Trim(tgl_keluar), "yyyy/mm/dd") & "'")
    
    If Len(tgl_keluar) = 0 Then
            sql = sql & ",tgl_keluar=NULL"
        Else
        sql = sql & ",tgl_keluar='" & Format(Trim(tgl_keluar), "yyyy/mm/dd") & "'"
        End If
    
    sql = sql & " where Kode_Karyawan='" & Trim(Txt_Kode.Text) & "'"
        
    Set rs = New ADODB.Recordset
        rs.Open sql, kon
        
        simpan_detail
        
        kon.CommitTrans
        
        konfirm = CInt(MsgBox("Data karyawan telah dirubah ...", vbOKOnly + vbInformation, "Informasi"))
        
        Cmd_Batal_Click
    
End If

isi_jb

rubah = False
On Error GoTo 0
Exit Sub

err_handler:
    
    kon.RollbackTrans
        konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
            Err.Clear

End Sub

Private Sub simpan_detail()
    Dim sql As String
    Dim i As Integer
    Dim rs As Recordset
    
    Dim kdjab As String
    Dim tgl1 As String
    Dim ket As String
    Dim gapok As Double
    
        i = 1
        
    If Not (arr_jb.UpperBound(1) = 1 And arr_jb(1, 1) = Empty) Then
    
        For i = 1 To arr_jb.UpperBound(1)
            
            If arr_jb(i, 7) = 0 Then
                
                kdjab = arr_jb(i, 2)
                tgl1 = arr_jb(i, 4)
                gapok = arr_jb(i, 5)
                ket = arr_jb(i, 6)
                
                sql = "insert into Tb_Karyawan_Jab (kd_karyawan,kd_jab,tgl1,keterangan,gapok) values("
                sql = sql & "'" & Trim(Txt_Kode.Text) & "','" & kdjab & "','" & Format(tgl1, "yyyy/mm/dd") & "','" & ket & "'," & gapok & ")"
                
                Set rs = New ADODB.Recordset
                rs.Open sql, kon
                
            End If
            
        Next
    
    End If
        
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
    
     cmd_add2.Enabled = False
     cmd_del2.Enabled = False
     cmdnonaktif.Enabled = False
    
     Txt_Kode.Text = ""
     Txt_Kode.Enabled = True
     Txt_Kode.SetFocus
        
        
End Sub

Private Sub cmdnonaktif_Click()
    karyawan3.Show 1
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

rubah = False

'' akses command ''

'    hak_akses_percommand CStr(Me.Name)
'
'    Cmd_Tambah.Enabled = c_tambah
'    Cmd_Rubah.Enabled = c_rubah
'    Cmd_Hapus.Enabled = c_hapus

'' stop here ''


With Me
    .Left = Screen.Width / 2 - .Width / 2
    .Top = 50
End With

grid_jb.Array = arr_jb
kosong_jb

Dim n As Object
    For Each n In Me
    
        If TypeOf n Is TextBox Then
            If Left(UCase(n.Name), 6) <> UCase("Txt_Cr") Then
                n.Enabled = False
            End If
        End If
        
'        If TypeOf n Is CheckBox Then n.Enabled = False
        If TypeOf n Is MaskEdBox Then n.Enabled = False
        If TypeOf n Is TDBDate Then n.Enabled = False
        If TypeOf n Is TDBNumber Then n.Enabled = False
        If TypeOf n Is ComboBox Then n.Enabled = False
        If TypeOf n Is OptionButton Then n.Enabled = False
        If TypeOf n Is CommandButton Then
            If n.Caption = "..." Then
                n.Enabled = False
            End If
        End If
            
    Next

Set n = Nothing


Grid_Rubah.Array = Arr_Rubah
Grid_Hapus.Array = Arr_Hapus
Grid_Daftar.Array = arr_daftar


cmd_add2.Enabled = False
cmd_del2.Enabled = False
'cmdnonaktif.Enabled = False

With Cbo_Agama
    .Clear
    .AddItem "01. Islam"
    .AddItem "02. Kristen"
    .AddItem "03. Hindu"
    .AddItem "04. Budha"
    .AddItem "05. Konghucu"
End With

atur_grid_transaksi

Isi_Combo

Cmd_Simpan.TabIndex = Txt_Ket.TabIndex + 1

Lbl_Umur.Caption = "0 Thn"

txt_cr_daftar_KeyUp 0, 0, 0

Cmd_Navigasi_Click 3

End Sub

Sub Isi_Combo()
    With Cbo_Jenis_Kelamin
        .Clear
        .AddItem "01. Pria"
        .AddItem "02. Wanita"
    End With
End Sub

Sub atur_grid_transaksi()
    
    With TDB_Rubah
        .Left = Me.Width / 2 - .Width / 2
        .Top = Me.Height / 2 - .Height / 2
    End With
    
    With TDB_Hapus
        .Left = Me.Width / 2 - .Width / 2
        .Top = Me.Height / 2 - .Height / 2
    End With

    With TDB_Daftar
        .Left = Me.Width / 2 - .Width / 2
        .Top = Me.Height / 2 - .Height / 2
    End With

End Sub

Sub Isi_grid_transaksi(ByVal rec As Recordset, ByVal gridnya As Integer)
    
    Dim a As Long
    Dim kode, NAMA, alamat As String
        
        Select Case gridnya
            Case 0
                Kosong_Rubah
            Case 1
                Kosong_Hapus
            Case 2
                kosong_daftar
        End Select
        
        a = 1
        
        With rec
            
           Do While Not .EOF
            
           Select Case gridnya
            Case 0
                Arr_Rubah.ReDim 1, a, 0, Grid_Rubah.Columns.Count
                Grid_Rubah.ReBind
                Grid_Rubah.Refresh
             Case 1
                Arr_Hapus.ReDim 1, a, 0, Grid_Hapus.Columns.Count
                Grid_Hapus.ReBind
                Grid_Hapus.Refresh
             Case 2
                arr_daftar.ReDim 1, a, 0, Grid_Daftar.Columns.Count
                Grid_Daftar.ReBind
                Grid_Daftar.Refresh
           End Select
            
            kode = IIf(Not IsNull(!kode_karyawan), !kode_karyawan, "")
            NAMA = IIf(Not IsNull(!Nama_Karyawan), !Nama_Karyawan, "")
            alamat = IIf(Not IsNull(!alamat_1), !alamat_1, "")
            
            Select Case gridnya
                Case 0
                    Arr_Rubah(a, 0) = kode
                    Arr_Rubah(a, 1) = NAMA
                    Arr_Rubah(a, 2) = alamat
                Case 1
                    Arr_Hapus(a, 0) = kode
                    Arr_Hapus(a, 1) = NAMA
                    Arr_Hapus(a, 2) = alamat
                Case 2
                    arr_daftar(a, 0) = kode
                    arr_daftar(a, 1) = NAMA
                    arr_daftar(a, 2) = alamat
            End Select
            
           a = a + 1
           .MoveNext
           Loop
            
           Select Case gridnya
            Case 0
            
                Grid_Rubah.ReBind
                Grid_Rubah.Refresh
                
                Grid_Rubah.MoveFirst
                
            Case 1
                
                Grid_Hapus.ReBind
                Grid_Hapus.Refresh
                
                Grid_Hapus.MoveLast
                
            Case 2
                
                Grid_Daftar.ReBind
                Grid_Daftar.Refresh
                
                Grid_Daftar.MoveLast
                
           End Select
            
        End With
End Sub

Sub Isi_grid_jb(ByVal rec As Recordset)
    
    Dim noid As Integer
    Dim nobagian As String
    Dim namabagian As String
    Dim kdjab As String
    Dim namajab As String
    Dim tgl1 As String
    Dim ket As String
    Dim gapok As Double
    
    kosong_jb
        
        Dim a As Integer
        a = 1
        
        With rec
            
           Do While Not .EOF
            
                arr_jb.ReDim 1, a, 0, grid_jb.Columns.Count
                grid_jb.ReBind
                grid_jb.Refresh
            
            noid = IIf(Not IsNull(!no_id), !no_id, 0)
            nobagian = IIf(Not IsNull(!no_bagian), !no_bagian, "")
            namabagian = IIf(Not IsNull(!nama_bagian), !nama_bagian, "")
            kdjab = IIf(Not IsNull(!kode_jab), !kode_jab, "")
            namajab = IIf(Not IsNull(!nama_jab), !nama_jab, "")
            tgl1 = IIf(Not IsNull(!tgl1), !tgl1, "")
            ket = IIf(Not IsNull(!keterangan), !keterangan, "")
            gapok = IIf(Not IsNull(!gapok), !gapok, 0)
            
                    arr_jb(a, 0) = nobagian
                    arr_jb(a, 1) = namabagian
                    arr_jb(a, 2) = kdjab
                    arr_jb(a, 3) = namajab
                    arr_jb(a, 4) = Format(tgl1, "dd/mm/yyyy")
                    arr_jb(a, 5) = gapok
                    arr_jb(a, 6) = ket
                    arr_jb(a, 7) = noid
            
           a = a + 1
           .MoveNext
           Loop

                grid_jb.ReBind
                grid_jb.Refresh
                
              If a > 1 Then grid_jb.MoveLast
            
        End With
End Sub

Sub Atur_Tdb(ByVal TDB As TDBContainer3D, ByVal frme As Frame, ByVal comd As CommandButton)
    With TDB
        .Left = frme.Left + comd.Left - .Width / 2
        .Top = frme.Top + comd.Top + comd.Height
    End With
End Sub

Sub Atur_Tdb_Atas(ByVal TDB As TDBContainer3D, ByVal frme As Frame, ByVal comd As CommandButton)
    With TDB
        .Left = frme.Left + comd.Left - .Width / 2
        .Top = frme.Top + comd.Top - .Height
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

Private Sub Grid_Bagian_DblClick()

On Error GoTo err_handler

If Grid_Bagian.Row < 0 Then Exit Sub

    t_no_bagian.Text = Grid_Bagian.Columns(0).Text
    t_nama_bagian.Text = Grid_Bagian.Columns(1).Text
    
    TDB_bagian.Visible = False
    t_no_bagian_LostFocus
    
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

Private Sub Grid_Counter_DblClick()

If Grid_Counter.Row < 0 Then Exit Sub

    Lbl_Kode_Counter.Caption = Grid_Counter.Columns(0).Text
    Lbl_Nama_Counter.Caption = Grid_Counter.Columns(1).Text
    
    TDB_Counter.Visible = False
    Cmd_Simpan.SetFocus

End Sub

Private Sub Grid_Counter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Counter_DblClick
    If KeyCode = vbKeyEscape Then TDB_Counter.Visible = False: Cmd_Browse_Counter.SetFocus
End Sub

Private Sub grid_daftar_DblClick()
    
    If arr_daftar.UpperBound(1) = 1 And arr_daftar(1, 1) = Empty Then Exit Sub

    Rs_Nav.MoveFirst
    
    Rs_Nav.Find "Kode_Karyawan='" & arr_daftar(Grid_Daftar.Bookmark, 0) & "'"

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
    
    If Arr_Hapus.UpperBound(1) = 1 And Arr_Hapus(1, 1) = Empty Then Exit Sub
    
    kon.BeginTrans
    
    If MsgBox("Yakin akan hapus : " & Arr_Hapus(Grid_Hapus.Bookmark, 0) & " ...?", vbYesNo + vbQuestion, "Hapus") = vbNo Then
        kon.RollbackTrans
        On Error GoTo 0
        Exit Sub
    End If
    
    Dim sql As String
    Dim rs As Recordset
        sql = "delete from Tb_Karyawan where Kode_Karyawan='" & Arr_Hapus(Grid_Hapus.Bookmark, 0) & "'"
            
            Set rs = New ADODB.Recordset
                rs.Open sql, kon
        
        kon.CommitTrans
        Dim konfirm As Integer
            
            konfirm = CInt(MsgBox(Arr_Hapus(Grid_Hapus.Bookmark, 0) & " Berhasil dihapus", vbOKOnly + vbInformation, "Hapus"))
            
            Cmd_Batal_Click
        
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

Private Sub Grid_Jabatan_DblClick()

On Error GoTo err_handler

If Grid_Bagian.Row < 0 Then Exit Sub

    t_kd_jabatan.Text = Grid_Jabatan.Columns(2).Text
    t_nama_jabatan.Text = Grid_Jabatan.Columns(3).Text
    
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

Private Sub Grid_Rubah_DblClick()

If Arr_Rubah.UpperBound(1) = 1 And Arr_Rubah(1, 1) = Empty Then Exit Sub

    Txt_Cr_Rubah(0).Text = Arr_Rubah(Grid_Rubah.Bookmark, 0)

    Txt_Cr_Rubah_KeyUp 0, 0, 0

    isi_semua Rs_Nav
    
    TDB_Rubah.Visible = False
        
        
    Dim n As Object
        For Each n In Me
                        If TypeOf n Is TextBox Then
                        
                         If Not (Left(UCase(n.Name), 9) = UCase("Txt_Kode_") Or n.Name = "Txt_Agama" Or n.Name = "Txt_Status" Or n.Name = "Txt_Pendidikan" Or n.Name = "Txt_Jabatan" Or n.Name = "Txt_Kode") Then
                            n.Enabled = True
                         End If
                         
                        End If
            
            If TypeOf n Is TDBDate Then n.Enabled = True
            If TypeOf n Is TDBNumber Then n.Enabled = True
            If TypeOf n Is ComboBox Then n.Enabled = True
            If TypeOf n Is OptionButton Then n.Enabled = True
            If TypeOf n Is MaskEdBox Then n.Enabled = True
'            If TypeOf n Is CheckBox Then n.Enabled = True
            
            If TypeOf n Is CommandButton Then
                If n.Caption = "..." Then
                    n.Enabled = True
                End If
            End If
            
        Next
    
    t_nama_jabatan.Enabled = False
    t_nama_bagian.Enabled = False
    
    cmd_add2.Enabled = True
    cmd_del2.Enabled = True
    cmdnonaktif.Enabled = True
    
    Cmd_Simpan.Enabled = True
    rubah = True
    
    Txt_Nama.SetFocus
    
End Sub

Sub isi_semua(ByVal rec As Recordset)
On Error Resume Next

    With rec
        
        If .EOF Then .MoveLast
        If .BOF Then .MoveFirst
        
        Txt_Kode = IIf(Not IsNull(!kode_karyawan), !kode_karyawan, "")
        Txt_Nama = IIf(Not IsNull(!Nama_Karyawan), !Nama_Karyawan, "")
        Txt_Tempat_Lhr = IIf(Not IsNull(!tempat_lhr), !tempat_lhr, "")
        TDB_Tgl_Lhr.Value = IIf(Not IsNull(!Tgl_Lhr), !Tgl_Lhr, Date)
        TDB_Tgl_Masuk.Value = IIf(Not IsNull(!Tgl_Masuk), !Tgl_Masuk, Date)
        
        TDB_Tgl_Lhr_LostFocus
        
        Txt_Kode_Jenis_Kelamin = IIf(Not IsNull(!jenis_kelamin), !jenis_kelamin, "")
        
        Txt_Alamat_1 = IIf(Not IsNull(!alamat_1), !alamat_1, "")
        Txt_Alamat_2 = IIf(Not IsNull(!alamat_2), !alamat_2, "")
        Txt_Alamat_3 = IIf(Not IsNull(!alamat_3), !alamat_3, "")
        Txt_Kodepos = IIf(Not IsNull(!Kode_Pos), !Kode_Pos, "")
        Txt_Telp_Rumah = IIf(Not IsNull(!No_Telp), !No_Telp, "")
        Txt_Telp_Hp = IIf(Not IsNull(!No_Telp_Hp), !No_Telp_Hp, "")
        Txt_Kode_Agama = IIf(Not IsNull(!Agama), !Agama, "")
        TDB_Gaji.Value = IIf(Not IsNull(!gaji), !gaji, Null)
        
        Txt_Ket.Text = IIf(Not IsNull(!ket), !ket, "")
        TDB_Uang_Makan.Value = IIf(Not IsNull(!Uang_Makan), !Uang_Makan, Null)
        TDB_Tunjangan.Value = IIf(Not IsNull(!Tunjangan), !Tunjangan, Null)
        
        Dim fl_gaji As String
            fl_gaji = IIf(Not IsNull(!flag_gaji), !flag_gaji, "")
            
            If fl_gaji = "b" Then
                Opt_Bulan.Value = True
            ElseIf fl_gaji = "h" Then
                Opt_Hari.Value = True
            End If
        
        
        
        Lbl_Kode_Counter.Caption = IIf(Not IsNull(!kode_counter), !kode_counter, "")
        Lbl_Nama_Counter.Caption = IIf(Not IsNull(!nama_counter), !nama_counter, "")
        
        
        t_no_bagian.Text = IIf(Not IsNull(!no_bagian), !no_bagian, "")
        t_nama_bagian.Text = IIf(Not IsNull(!nama_bagian), !nama_bagian, "")
        t_kd_jabatan.Text = IIf(Not IsNull(!kd_jab), !kd_jab, "")
        t_nama_jabatan.Text = IIf(Not IsNull(!nama_jab), !nama_jab, "")
        t_jenjang.Text = IIf(Not IsNull(!jenjang_jbt), !jenjang_jbt, "")
        t_peringkat.Text = IIf(Not IsNull(!peringkat), !peringkat, "")
        
        tkode_abs.Text = IIf(Not IsNull(!kd_absen), !kd_absen, "")
        tdb_gapok.Value = IIf(Not IsNull(!gapok), !gapok, Null)
        
        If Not IsNull(!tgl_keluar) Then
            lblstatus.Caption = "STATUS : NON-AKTIF"
        Else
            lblstatus.Caption = "STATUS : AKTIF"
        End If
        
        If .RecordCount = 0 Then
            Lbl_Info.Caption = "Record Ke " & 0 & " Dari " & .RecordCount & " Record"
        Else
            Lbl_Info.Caption = "Record Ke " & .AbsolutePosition & " Dari " & .RecordCount & " Record"
        End If
    End With
    
End Sub

Private Sub Grid_Rubah_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Rubah_DblClick
    If KeyCode = vbKeyEscape Then TDB_Rubah.Visible = False: Cmd_Batal_Click
End Sub

Private Sub Opt_Aka_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Simpan.SetFocus
End Sub

Private Sub Opt_Bulan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then TDB_Tunjangan.SetFocus
End Sub

Private Sub Opt_Hari_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then TDB_Tunjangan.SetFocus
End Sub

Private Sub Opt_Istana_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Simpan.SetFocus
End Sub

Private Sub t_jenjang_GotFocus()
    Call Focus_(t_jenjang)
End Sub

Private Sub t_jenjang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then t_peringkat.SetFocus
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
    If t_no_bagian.Text = "" Then
        
        Dim a As Long
            a = CInt(MsgBox("No bagian tidak boleh kosong", vbOKOnly + vbInformation, "Informasi"))
            
            t_no_bagian.SetFocus
            Exit Sub
    End If
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select * from VIEW_Jabatan where kode_jab='" & Trim(t_kd_jabatan.Text) & "'"
        sql = sql & " and no_bagian='" & Trim(t_no_bagian.Text) & "'"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon, adOpenKeyset
        
        With rs
            If Not .EOF Then
                
                t_nama_jabatan.Text = IIf(Not IsNull(!nama_jab), !nama_jab, "")
                
                Txt_Tempat_Lhr.SetFocus
            
            Else
                
                a = CInt(MsgBox("Kode jabatan yang anda masukkan tidak ditemukan dalam bagian " & t_nama_bagian.Text, vbOKOnly + vbInformation, "Informasi"))
                
                t_kd_jabatan.Text = ""
                t_kd_jabatan.SetFocus
                
            End If
        End With
        
    
End Sub

Private Sub t_no_bagian_GotFocus()
    Call Focus_(t_no_bagian)
End Sub

Private Sub t_no_bagian_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then t_no_bagian_LostFocus
    If KeyCode = vbKeyF3 Then cmd_browse_bag_Click
End Sub

Private Sub t_no_bagian_LostFocus()

    If t_no_bagian.Text = "" Then Exit Sub
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select no_bagian,nama_bagian from Tb_Bagian where no_bagian='" & Trim(t_no_bagian.Text) & "'"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon, adOpenKeyset
        
        With rs
            If .EOF Then
                Dim a As Integer
                    a = CInt(MsgBox("No bagian yang anda masukkan tidak ditemukan", vbOKOnly + vbInformation, "Informasi"))
                    
                    t_kd_jabatan.Enabled = False
'                    t_nama_jab.Enabled = False
                    t_no_bagian.Text = ""
                    t_nama_bagian.Text = ""
                    t_no_bagian.SetFocus
                    
                    
            Else
                
               t_nama_bagian.Text = IIf(Not IsNull(!nama_bagian), !nama_bagian, "")
               
               t_kd_jabatan.Enabled = True
'               t_nama_jab.Enabled = True
               
               t_kd_jabatan.Text = ""
               t_nama_jabatan.Text = ""
               
               t_kd_jabatan.SetFocus
                
            End If
        End With


End Sub

Private Sub t_peringkat_GotFocus()
    Call Focus_(t_peringkat)
End Sub

Private Sub t_peringkat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Tempat_Lhr.SetFocus
End Sub

Private Sub TabStrip1_Click()

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
   TDB_bagian.Top = TDB_bagian.Top - (yold - Y)
   TDB_bagian.Left = TDB_bagian.Left - (xold - X)
End If

End Sub

Private Sub TDB_bagian_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub

Private Sub TDB_Counter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub TDB_Counter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   TDB_Counter.Top = TDB_Counter.Top - (yold - Y)
   TDB_Counter.Left = TDB_Counter.Left - (xold - X)
End If

End Sub

Private Sub TDB_Counter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub

Private Sub TDB_Gaji_GotFocus()
    Call Focus_(TDB_Gaji)
End Sub

Private Sub TDB_Gaji_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Opt_Bulan.SetFocus
    End If
End Sub

Private Sub TDB_Gaji_LostFocus()
    
    If TDB_Gaji.ValueIsNull Then
        TDB_Gaji.Value = Null
    End If
    
End Sub

Private Sub TDB_Hutang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Cmd_Simpan.SetFocus
    End If
End Sub

Private Sub tdb_gapok_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then ttgl_keluar.SetFocus
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

Private Sub TDB_Tgl_Lhr_GotFocus()
    Call Focus_(TDB_Tgl_Lhr)
End Sub

Private Sub TDB_Tgl_Lhr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Cbo_Jenis_Kelamin_Click
        Cbo_Jenis_Kelamin.SetFocus
    End If
End Sub

Private Sub TDB_Tgl_Lhr_LostFocus()

On Error GoTo err_handler

    Dim tahun As Long
        
        'TAHUN = Year(Now) - Year(TDB_Tgl_Lhr.Text)
        
        kalkulasi_umur TDB_Tgl_Lhr.Text
        
        Lbl_Umur.Caption = tahun_u & " Tahun"
        
        '  bulan_u & " Bulan " & hari_u & " Hari"
        
        tahun = CDbl(Year(TDB_Tgl_Masuk.Text))
        tahun = tahun - CDbl(Year(TDB_Tgl_Lhr.Text))
        tahun = 54 - tahun
      '  tahun = tahun + (56 - CDbl(tahun_u))
     '   tahun = tahun - (CDbl(Year(Now)))
        
        lblmasakerja.Caption = tahun & " Thn"
        
On Error GoTo 0
Exit Sub

err_handler:
    
    Dim konfirm As Integer
        konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
        Err.Clear
        
End Sub

Private Sub TDB_Tgl_Masuk_GotFocus()
    Call Focus_(TDB_Tgl_Masuk)
End Sub

Private Sub TDB_Tgl_Masuk_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tkode_abs.SetFocus
End Sub

Private Sub TDB_Tunjangan_GotFocus()
    Call Focus_(TDB_Tunjangan)
End Sub

Private Sub TDB_Tunjangan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then TDB_Uang_Makan.SetFocus
End Sub

Private Sub TDB_Tunjangan_LostFocus()
    If TDB_Tunjangan.ValueIsNull Then
        TDB_Tunjangan.Value = Null
    End If
End Sub

Private Sub TDB_Uang_Makan_GotFocus()
    Call Focus_(TDB_Uang_Makan)
End Sub

Private Sub TDB_Uang_Makan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Ket.SetFocus
End Sub

Private Sub TDB_Uang_Makan_LostFocus()
    If TDB_Uang_Makan.ValueIsNull Then
        TDB_Uang_Makan.Value = Null
    End If
End Sub

Private Sub TDBContainer3D1_Click()

End Sub

Private Sub tkode_abs_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Simpan.SetFocus
End Sub

Private Sub ttgl_keluar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Simpan.SetFocus
End Sub

Private Sub Txt_Alamat_1_GotFocus()
    Call Focus_(Txt_Alamat_1)
End Sub

Private Sub Txt_Alamat_1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Txt_Alamat_2.SetFocus
    End If
End Sub

Private Sub Txt_Alamat_2_GotFocus()
    Call Focus_(Txt_Alamat_2)
End Sub

Private Sub Txt_Alamat_2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Txt_Alamat_3.SetFocus
    End If
End Sub

Private Sub Txt_Alamat_3_GotFocus()
    Call Focus_(Txt_Alamat_3)
End Sub

Private Sub Txt_Alamat_3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Txt_Kodepos.SetFocus
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

Private Sub Txt_Cr_Counter_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Counter.SetFocus
    If KeyCode = vbKeyEscape Then TDB_Counter.Visible = False: Cmd_Browse_Counter.SetFocus
End Sub

Private Sub Txt_Cr_Counter_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select top 100 * from Tb_Mast_Counter"
        
        Select Case Index
            Case 0
                sql = sql & " where Kode like '%" & Trim(Txt_Cr_Counter(0).Text) & "%'"
            Case 1
                sql = sql & " where Nama_Counter like '%" & Trim(Txt_Cr_Counter(1).Text) & "%'"
        End Select
        
        sql = sql & " order by Kode,Nama_Counter asc"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon, adOpenKeyset
            
            Set Grid_Counter.DataSource = rs
                Grid_Counter.Refresh
        
End Sub

Private Sub txt_cr_daftar_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Daftar.SetFocus
    If KeyCode = vbKeyEscape Then Cmd_Batal_Click
End Sub

Private Sub txt_cr_daftar_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
           
    Dim sql As String
        sql = "select top 100 * from VIEW_Karyawan"
        
    If Txt_Cr_Daftar(0).Text <> "" Or Txt_Cr_Daftar(1).Text <> "" Then
        Select Case Index
            Case 0
                sql = sql & " where Kode_Karyawan like '%" & Trim(Txt_Cr_Daftar(0).Text) & "%'"
            Case 1
                sql = sql & " where Nama_Karyawan like '%" & Trim(Txt_Cr_Daftar(1).Text) & "%'"
        End Select
    End If
    
    sql = sql & " order by Kode_Karyawan asc"
    
    Set Rs_Nav = New ADODB.Recordset
        Rs_Nav.Open sql, kon, adOpenKeyset
    
    Isi_grid_transaksi Rs_Nav, 2

End Sub

Private Sub Txt_Cr_Hapus_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Hapus.SetFocus
    If KeyCode = vbKeyEscape Then Cmd_Batal_Click
End Sub

Private Sub Txt_Cr_Hapus_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    Dim sql As String
        sql = "select top 100 * from VIEW_Karyawan"
            
    If Txt_Cr_Hapus(0).Text <> "" Or Txt_Cr_Hapus(1).Text <> "" Then
        Select Case Index
            Case 0
                sql = sql & " where Kode_Karyawan like '%" & Trim(Txt_Cr_Hapus(0).Text) & "%'"
            Case 1
                sql = sql & " where Nama_Karyawan like '%" & Trim(Txt_Cr_Hapus(1).Text) & "%'"
        End Select
    End If

    sql = sql & " order by Kode_Karyawan asc"
    
    Set Rs_Nav = New ADODB.Recordset
        Rs_Nav.Open sql, kon, adOpenKeyset
    
    Isi_grid_transaksi Rs_Nav, 1
    
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

Private Sub Txt_Cr_Rubah_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Rubah.SetFocus
    If KeyCode = vbKeyEscape Then Cmd_Batal_Click
End Sub

Private Sub Txt_Cr_Rubah_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)


            
    Dim sql As String
        sql = "select top 100 * from View_Karyawan"
        
    If Txt_Cr_Rubah(0).Text <> "" Or Txt_Cr_Rubah(1).Text <> "" Then
        Select Case Index
            Case 0
                sql = sql & " where Kode_Karyawan like '%" & Trim(Txt_Cr_Rubah(0).Text) & "%'"
            Case 1
                sql = sql & " where Nama_Karyawan like '%" & Trim(Txt_Cr_Rubah(1).Text) & "%'"
        End Select
    End If
    
    sql = sql & " order by Kode_Karyawan asc"
    
    Set Rs_Nav = New ADODB.Recordset
        Rs_Nav.Open sql, kon, adOpenKeyset
    
    Isi_grid_transaksi Rs_Nav, 0
    
End Sub


Private Sub Txt_Ket_GotFocus()
    Call Focus_(Txt_Ket)
End Sub

Private Sub Txt_Ket_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Browse_Counter.SetFocus
End Sub

Private Sub Txt_Kode_Agama_Change()
    With Txt_Kode_Agama
        If .Text = "01" Then
            Cbo_Agama.ListIndex = 0
        ElseIf .Text = "02" Then
            Cbo_Agama.ListIndex = 1
        ElseIf .Text = "03" Then
            Cbo_Agama.ListIndex = 2
        ElseIf .Text = "04" Then
            Cbo_Agama.ListIndex = 3
        ElseIf .Text = "05" Then
            Cbo_Agama.ListIndex = 4
        End If
    End With
End Sub

Private Sub Txt_Kode_Change()

    If Len(Trim(Txt_Kode.Text)) = 0 Then
        kosong_jb
    Else
        isi_jb
    End If
    
End Sub

Private Sub Txt_Kode_Jenis_Kelamin_Change()
    With Txt_Kode_Jenis_Kelamin
        If .Text = "01" Then
            Cbo_Jenis_Kelamin.ListIndex = 0
        Else
            Cbo_Jenis_Kelamin.ListIndex = 1
        End If
    End With
End Sub

Private Sub Txt_Kode_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo err_handler
    
    Dim n As Object
    If KeyCode = 13 And Txt_Kode.Text <> "" Then
    
    Dim sql As String
    Dim rs As Recordset
    
    sql = "select Kode_Karyawan from Tb_Karyawan where Kode_Karyawan='" & Trim(Txt_Kode.Text) & "'"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon
        
     If Not rs.EOF Then
            Dim konfirm As Integer
                konfirm = CInt(MsgBox("Kode Sudah ada ...", vbOKOnly + vbInformation, "Informasi"))

    For Each n In Me
    
        If TypeOf n Is TextBox Then
            If Left(UCase(n.Name), 6) <> UCase("Txt_Cr") And UCase(n.Name) <> UCase("txt_kode") Then
                n.Enabled = False
            End If
        End If
        
        If TypeOf n Is TDBDate Then n.Enabled = False
        If TypeOf n Is TDBNumber Then n.Enabled = False
        If TypeOf n Is MaskEdBox Then n.Enabled = False
        If TypeOf n Is ComboBox Then n.Enabled = False
        If TypeOf n Is OptionButton Then n.Enabled = False
        If TypeOf n Is CommandButton Then
            If n.Caption = "..." Then
                n.Enabled = False
            End If
        End If
            
    Next

    Set n = Nothing
                
                t_nama_bagian.Enabled = False
                t_nama_jabatan.Enabled = False
                
                Txt_Kode.SetFocus
                Cmd_Simpan.Enabled = False
                 
                cmd_add2.Enabled = False
                cmd_del2.Enabled = False
                 
                On Error GoTo 0
                Exit Sub
        Else
                    
                    For Each n In Me
                    
                        If TypeOf n Is TextBox Then
                        
                         If Not (UCase(n.Name) = UCase("Txt_Kode") Or n.Name = "Txt_Agama" Or n.Name = "Txt_Status" Or n.Name = "Txt_Pendidikan" Or n.Name = "Txt_Jabatan") Then
                            n.Enabled = True
                         End If
                         
                         If Not (Left(UCase(n.Name), 9) = UCase("Txt_Kode")) Then
                            n.Text = ""
                         End If
                         
                        End If
                        
                       If TypeOf n Is TDBDate Then n.Enabled = True
                        If TypeOf n Is TDBNumber Then
                            n.Enabled = True
                            n.Text = ""
                        End If
                        If TypeOf n Is ComboBox Then n.Enabled = True
                        If TypeOf n Is OptionButton Then n.Enabled = True
                        If TypeOf n Is MaskEdBox Then n.Enabled = True
                        If TypeOf n Is CommandButton Then
                            If n.Caption = "..." Then
                                n.Enabled = True
                            End If
                        End If
                 
                        
                    Next
                    
                    Set n = Nothing
                
                t_nama_bagian.Enabled = False
                t_nama_jabatan.Enabled = False
                
                cmd_add2.Enabled = True
                cmd_del2.Enabled = True
                
                Txt_Ket.Text = ""
                Txt_Ket.Enabled = True
                Lbl_Umur.Caption = "0 Thn"
                Lbl_Nama_Counter.Caption = ""
                Lbl_Kode_Counter.Caption = ""
                Cmd_Simpan.Enabled = True
                Txt_Nama.SetFocus
                
        End If
            
    End If
    
On Error GoTo 0
Exit Sub

err_handler:
    
    Dim p As Integer
        p = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
        Err.Clear

End Sub

Private Sub Txt_Kodepos_GotFocus()
    Call Focus_(Txt_Kodepos)
End Sub

Private Sub Txt_Kodepos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Telp_Rumah.SetFocus
End Sub

Private Sub txt_nama_GotFocus()
    Call Focus_(Txt_Nama)
End Sub

Private Sub Txt_Nama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Txt_Tempat_Lhr.SetFocus
    End If
End Sub

Private Sub Txt_Telp_Hp_GotFocus()
    Call Focus_(Txt_Telp_Hp)
End Sub

Private Sub Txt_Telp_Hp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then TDB_Tgl_Masuk.SetFocus
End Sub

Private Sub Txt_Telp_Rumah_GotFocus()
    Call Focus_(Txt_Telp_Rumah)
End Sub

Private Sub Txt_Telp_Rumah_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Telp_Hp.SetFocus
End Sub

Private Sub Txt_Tempat_Lhr_GotFocus()
    Call Focus_(Txt_Tempat_Lhr)
End Sub

Private Sub Txt_Tempat_Lhr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        TDB_Tgl_Lhr.SetFocus
    End If
End Sub
