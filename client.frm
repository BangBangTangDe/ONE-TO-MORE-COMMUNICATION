VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form client 
   Caption         =   "client"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin MSWinsockLib.Winsock Winsockclient0 
      Left            =   120
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
      LocalPort       =   101
   End
   Begin VB.CommandButton Command3 
      Caption         =   "连接"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   1935
      Left            =   1080
      TabIndex        =   7
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "发送"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "接收"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   135
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   15
   End
   Begin VB.Label Label2 
      Caption         =   "发送"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "IP"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "client"
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
Winsockclient0.SendData a
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
Winsockclient0.Connect
End Sub


'ip号
Private Sub Text1_Change()
Winsockclient0.RemoteHost = Text1.Text
End Sub
'send


Private Sub winsockclient0_DataArrival(ByVal requestID As Long)
Dim str As String
Dim a As String

Winsockclient0.GetData str
a = str
Text3.Text = a

End Sub
