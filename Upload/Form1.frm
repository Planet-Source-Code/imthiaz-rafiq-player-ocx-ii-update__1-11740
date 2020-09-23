VERSION 5.00
Object = "{A74DF9B6-95A3-11D4-9012-CF206198051B}#6.0#0"; "playerX.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Xplayer By Iimthiaz Beta Test 2"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3690
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   3690
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2040
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   500
      Left            =   60
      TabIndex        =   1
      Top             =   3780
      Width           =   2055
      Begin VB.CommandButton Command1 
         Caption         =   "Load"
         Height          =   315
         Left            =   30
         TabIndex        =   2
         Top             =   90
         Width           =   915
      End
   End
   Begin PlayerX.Player Player1 
      Height          =   3705
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   6535
      CurrentPosition =   0
      ChildNo         =   0
      VolumePercent   =   8
      BalancePercent  =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo Errhand

cd1.Filter = "All Supported files (*.*)|*.dat;*.avi;*.mp3;*.wav;*.mpg"
cd1.DialogTitle = "Open a media file to play"
cd1.InitDir = App.Path
cd1.ShowOpen
If cd1.FileName <> "" Then
  Player1.File = cd1.FileName
End If
Exit Sub
Errhand:

End Sub

Private Sub Form_Load()
'Player1.File = "E:\Bloom\Junck\sdsj.dat"
End Sub

Private Sub Form_Resize()
Player1.Top = 10
Player1.Left = 10
If Form1.Height > 4000 Then
  Player1.Height = Form1.Height - 950
  Frame1.Top = Player1.Height + 40
Else
  Form1.Height = 4450
End If
If Form1.Width > 4149 Then
  Player1.Width = Form1.Width - 150
Else
  Form1.Width = 4150
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim x As VbMsgBoxResult

x = MsgBox("I think so you liked my ocx if so can you vote me in planetSourcecode.com", vbQuestion + vbYesNo, "Vote Me")

If x = vbYes Then
  Shell "start http://www.planet-source-code.com/xq/ASP/txtCodeId.11740/lngWId.1/qx/vb/scripts/ShowCode.htm", vbHide
Else
  MsgBox "It's Alright Man Chear up Wait for the Source code in Few Days with all bugs fixed" + vbCrLf + vbCrLf + "Mail Me At imthiazrafiq@hotmail.com or Icq # 24294947", vbInformation
End If
End Sub
