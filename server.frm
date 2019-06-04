VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form server 
   Caption         =   "server"
   ClientHeight    =   5310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   5175
   StartUpPosition =   3  '窗口缺省
   Begin MSWinsockLib.Winsock listener 
      Left            =   120
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   100
      LocalPort       =   80
   End
   Begin VB.CommandButton button_cancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton button_send 
      Caption         =   "发送"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   2535
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox send_text 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
   Begin MSWinsockLib.Winsock WinsockServer 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   100
      LocalPort       =   81
   End
   Begin VB.Label server_send 
      Caption         =   "发送"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   495
   End
   Begin VB.Label server_get 
      Caption         =   "接收"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   1440
      Width           =   735
   End
End
Attribute VB_Name = "server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'取消
Dim sockindex As Integer
Private Sub button_cancel_Click()
listener.Close
For i = 1 To WinsockServer.UBound
WinsockServer(i).Close
Next i
End
End Sub

'加载

Private Sub Form_Load()
sockindex = 0
listener.Listen
client.Show
client1.Show

End Sub

'发送

Private Sub button_send_Click()
Dim a As String
For i = 1 To WinsockServer.UBound
If WinsockServer(i).State = 7 Then
a = "服务器" & "客户端" & i & ":" & send_text.Text
WinsockServer(i).SendData a
End If
Next i
send_text.Text = ""
End Sub

'接收

Private Sub winsockserver_DataArrival(Index As Integer, ByVal requestID As Long)
Dim str As String
Dim a As String

WinsockServer(Index).GetData str
a = "客户端" & Index & ":" & str
If Text1.Text = "" Then
Text1.Text = a
Else
Text1.Text = Text1.Text & vbCrLf & a

End If

End Sub
'连接

Private Sub listener_ConnectionRequest(ByVal requestID As Long)
Dim msg As String
For i = 1 To WinsockServer.UBound - 1
  If WinsockServer(i).State = 7 Then sockindex = i
  Next
sockindex = sockindex + 1
Load WinsockServer(sockindex)
WinsockServer(sockindex).Accept requestID
msg = "连接客户端" & sockindex & "成功"
MsgBox (msg)
End Sub


Private Sub winsockserver_close(Index As Integer)
If WinsockServer(Index).State <> sckClosed Then
WinsockServer(Index).Close
End If
End Sub



