VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form client1 
   Caption         =   "client1"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "发送"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   1935
      Left            =   1080
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "连接"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin MSWinsockLib.Winsock Winsockclient 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
      LocalPort       =   100
   End
   Begin VB.Label Label1 
      Caption         =   "IP"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "发送"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   135
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Width           =   15
   End
   Begin VB.Label Label4 
      Caption         =   "接收"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   3120
      Width           =   495
   End
End
Attribute VB_Name = "client1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'主窗口及其配置
'send button

Private Sub Command1_Click()
Dim a As String
a = "客户端信息=>" & Text2.Text
MsgBox a
Winsockclient.SendData a
Text2.Text = ""
End Sub

'cancel

Private Sub Command2_Click()
Winsockclient.Close
End Sub

Private Sub Form_Load()
Text1.Text = "输入服务器的ip"
End Sub


'连接按钮
Private Sub Command3_Click()
Winsockclient.Connect
End Sub


'ip号
Private Sub Text1_Change()
Winsockclient.RemoteHost = Text1.Text
End Sub
'send


Private Sub winsockclient_DataArrival(ByVal requestID As Long)
Dim str As String
Dim a As String

Winsockclient.GetData str
a = str
Text3.Text = a

End Sub

