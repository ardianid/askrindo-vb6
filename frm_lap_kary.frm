VERSION 5.00
Object = "{EC76FE26-BAFD-4E89-AA40-E748DA83A570}#1.0#0"; "IsButton_Ard.ocx"
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_lap_kary 
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   9645
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_lap_kary.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7665
   ScaleWidth      =   9645
   WindowState     =   2  'Maximized
   Begin IsButton_Ard.isButton isButton1 
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   741
      Icon            =   "frm_lap_kary.frx":000C
      Style           =   8
      Caption         =   "x"
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7005
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5805
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frm_lap_kary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As Report_Kary
Option Explicit

Private Sub Form_Load()
On Error GoTo err_handler

Screen.MousePointer = vbHourglass

Dim sql As String
Dim rs As Recordset
    sql = Mysq
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon

Set Report = New Report_Kary
    Report.Database.SetDataSource rs, , 1
    Report.Subreport1.OpenSubreport.Database.SetDataSource rs, , 1
    
    CRViewer1.Zoom 1

CRViewer1.ReportSource = Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault

On Error GoTo 0
Exit Sub

err_handler:
    
Screen.MousePointer = vbDefault
    
    Dim konfirm As Integer
        konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
            Err.Clear
End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

isButton1.Left = Me.Width - isButton1.Width - 100
isButton1.Top = 0

End Sub

Private Sub isButton1_Click()
    Unload Me
    
    Set Report = Nothing
    
End Sub



