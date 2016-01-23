VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   3210
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   2640
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "v1.0"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   1
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Protector 2015 by DamDam"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   3165
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Top             =   0
      Width           =   7500
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim efek As Integer

Private Sub Timer1_Timer()
On Error Resume Next
efek = efek + 5
ProgressBar1.Value = ProgressBar1.Value + 400 / 400

If efek > 500 Then
    Timer1.Enabled = False
    Screen.MousePointer = vbNormal
    Me.WindowState = 0
    Do
    Me.Left = Me.Left + 20
    Me.Move Me.Left, Me.Top
    DoEvents
    Loop Until Me.Left > Screen.Width
    Load frmUtama
    frmUtama.Show
    Unload Me
End If
End Sub
