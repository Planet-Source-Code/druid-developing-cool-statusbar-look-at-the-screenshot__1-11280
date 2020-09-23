VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStatusBar 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Status Bar Magic"
   ClientHeight    =   3435
   ClientLeft      =   3780
   ClientTop       =   2760
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmStatusBar.frx":0000
   ScaleHeight     =   3435
   ScaleWidth      =   7275
   Begin VB.Timer tmrDisconnect 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6840
      Top             =   1560
   End
   Begin VB.Timer tmrConnect 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6360
      Top             =   1560
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      Picture         =   "frmStatusBar.frx":00A6
      Style           =   1  'Grafisch
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   615
      Left            =   120
      Picture         =   "frmStatusBar.frx":0430
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.Timer tmrAnimate2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5880
      Top             =   1560
   End
   Begin MSComctlLib.ProgressBar prStatus 
      Height          =   255
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Min             =   1e-4
      Max             =   100
      Scrolling       =   1
   End
   Begin VB.Timer tmrAnimate1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5400
      Top             =   1560
   End
   Begin MSComctlLib.StatusBar staStatus 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   3120
      Width           =   6670
      _ExtentX        =   11774
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4057
            MinWidth        =   4057
            Picture         =   "frmStatusBar.frx":07BA
            Text            =   "Disconnected"
            TextSave        =   "Disconnected"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3881
            MinWidth        =   3881
            Text            =   "Ready"
            TextSave        =   "Ready"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   882
            MinWidth        =   882
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
      EndProperty
   End
   Begin VB.Image imgEmpty 
      Height          =   255
      Left            =   6000
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgAnimate2 
      Height          =   240
      Index           =   3
      Left            =   6480
      Picture         =   "frmStatusBar.frx":0B56
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate2 
      Height          =   240
      Index           =   2
      Left            =   6480
      Picture         =   "frmStatusBar.frx":0EE0
      Top             =   960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate2 
      Height          =   240
      Index           =   1
      Left            =   6480
      Picture         =   "frmStatusBar.frx":126A
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate2 
      Height          =   240
      Index           =   0
      Left            =   6480
      Picture         =   "frmStatusBar.frx":15F4
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imStatus 
      Height          =   240
      Index           =   2
      Left            =   6240
      Picture         =   "frmStatusBar.frx":197E
      Top             =   960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imStatus 
      Height          =   240
      Index           =   1
      Left            =   6240
      Picture         =   "frmStatusBar.frx":1D08
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imStatus 
      Height          =   240
      Index           =   0
      Left            =   6240
      Picture         =   "frmStatusBar.frx":2092
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imWorking 
      Height          =   240
      Left            =   6990
      Picture         =   "frmStatusBar.frx":241C
      Top             =   3170
      Width           =   240
   End
   Begin VB.Image imOK 
      Height          =   240
      Left            =   6700
      Picture         =   "frmStatusBar.frx":27A6
      Top             =   3170
      Width           =   240
   End
   Begin VB.Image imgAnimate1 
      Height          =   240
      Index           =   3
      Left            =   6720
      Picture         =   "frmStatusBar.frx":2B30
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate1 
      Height          =   240
      Index           =   2
      Left            =   6720
      Picture         =   "frmStatusBar.frx":2EBA
      Top             =   960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate1 
      Height          =   240
      Index           =   1
      Left            =   6720
      Picture         =   "frmStatusBar.frx":3244
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAnimate1 
      Height          =   240
      Index           =   0
      Left            =   6720
      Picture         =   "frmStatusBar.frx":35CE
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgConnected 
      Height          =   240
      Left            =   6960
      Picture         =   "frmStatusBar.frx":3958
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare Variables
Dim AnimateStatus1 As Long
Dim AnimateStatus2 As Long

Private Sub cmdConnect_Click()
    tmrConnect.Enabled = True
    cmdConnect.Enabled = False
    'Set the pictures to working
    imOK.Picture = imStatus(2).Picture
    imWorking.Picture = imStatus(1).Picture
    staStatus.Panels(2).Text = "Connecting..."
    staStatus.Panels(2).Picture = imgAnimate2(0).Picture
    tmrAnimate2.Enabled = True
End Sub

Private Sub cmdDisconnect_Click()
    tmrDisconnect.Enabled = True
    cmdDisconnect.Enabled = False
    'Set the pictures to working
    imOK.Picture = imStatus(2).Picture
    imWorking.Picture = imStatus(1).Picture
    staStatus.Panels(2).Text = "Disconnecting..."
    staStatus.Panels(2).Picture = imgAnimate1(0).Picture
    tmrAnimate1.Enabled = True
End Sub

Private Sub Form_Load()
    'Set the AnimateStatus to 0
    AnimateStatus1 = 0
    AnimateStatus2 = 0
    'Place the ProgressBar into the panel
    prStatus.Left = staStatus.Panels(4).Left + 40
    prStatus.Top = staStatus.Top + 60
    prStatus.Width = staStatus.Width - staStatus.Panels(4).Left - 100
    prStatus.Height = staStatus.Height - 90
End Sub

Private Sub tmrAnimate1_Timer()
    'If the Animation loop has finished restart it
    If AnimateStatus1 = imgAnimate1.Count Then
        AnimateStatus1 = 0
    End If
    'Replace the actual StatusBar picture with the next one
    staStatus.Panels(2).Picture = imgAnimate1(AnimateStatus1).Picture
    AnimateStatus1 = AnimateStatus1 + 1
End Sub

Private Sub tmrAnimate2_Timer()
    'If the Animation loop has finished restart it
    If AnimateStatus2 = imgAnimate2.Count Then
        AnimateStatus2 = 0
    End If
    'Replace the actual StatusBar picture with the next one
    staStatus.Panels(2).Picture = imgAnimate2(AnimateStatus2).Picture
    AnimateStatus2 = AnimateStatus2 + 1
End Sub

Private Sub tmrConnect_Timer()
    'Add 1 to the value of the ProgressBar
    prStatus.Value = prStatus.Value + 1
    'Display the percent value of the progressbar in the StatusBar
    staStatus.Panels(3).Text = Round(prStatus.Value) & "%"
    'If the ProgressBar reaches 100 the computer is "connected"
    If prStatus.Value >= 100 Then
    'Set the properties of the controls to "connected"
    prStatus.Value = prStatus.Min
    tmrAnimate2.Enabled = False
    tmrConnect.Enabled = False
    imOK.Picture = imStatus(0).Picture
    imWorking.Picture = imStatus(2).Picture
    staStatus.Panels(2).Text = ""
    staStatus.Panels(2).Picture = imgEmpty.Picture
    staStatus.Panels(1).Text = "Connected to the Internet"
    staStatus.Panels(1).Picture = imgConnected.Picture
    staStatus.Panels(3).Text = ""
    cmdDisconnect.Enabled = True
    End If
End Sub

Private Sub tmrDisconnect_Timer()
    'Add 1 to the value of the ProgressBar
    prStatus.Value = prStatus.Value + 1
    'Display the percent value of the progressbar in the StatusBar
    staStatus.Panels(3).Text = Round(prStatus.Value) & "%"
    'If the ProgressBar reaches 100 the computer is "Disconnected"
    If prStatus.Value >= 100 Then
    'Set the properties of the controls to "Disconnected"
    prStatus.Value = prStatus.Min
    tmrAnimate1.Enabled = False
    tmrDisconnect.Enabled = False
    imOK.Picture = imStatus(0).Picture
    imWorking.Picture = imStatus(2).Picture
    staStatus.Panels(2).Text = ""
    staStatus.Panels(2).Picture = imgEmpty.Picture
    staStatus.Panels(1).Text = "Disconnected"
    staStatus.Panels(1).Picture = imgAnimate1(1).Picture
    staStatus.Panels(3).Text = ""
    cmdConnect.Enabled = True
    End If
End Sub
