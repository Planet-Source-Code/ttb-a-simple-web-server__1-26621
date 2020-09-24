VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Simple Webserver."
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3315
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   3315
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Hoe 
      Index           =   0
      Left            =   2040
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Start"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save Html File"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Html File"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "webs.frx":0000
      Top             =   600
      Width           =   3135
   End
   Begin MSWinsockLib.Winsock Pimp 
      Left            =   2520
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
   Begin VB.Label current 
      Caption         =   "0"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label curr 
      Caption         =   "Connections:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function LoadFile(filename1 As String) As String
On Error GoTo hell
Open filename1 For Binary As #1
LoadFile = Input(FileLen(filename1), #1)
Close #1
hell:
If Err.Number = 76 Then LoadFile = "Cant find file! Oh no!"
End Function

Private Sub Command1_Click()
Dim fn As String
fn = InputBox("What file name?", "Html File", App.Path & "\ main.html")
Open fn For Binary As #1
Text1 = Input(FileLen(fn), #1)
Close #1
End Sub

Private Sub Command2_Click()
Dim fn As String
fn = InputBox("What Filename?", "Enter File Name")
Open fn For Output As #1
Print #1, Text1
Close #1
End Sub

Private Sub Command3_Click()
If Command3.Caption = "Start" Then
Pimp.Close
Pimp.LocalPort = 80
Pimp.Listen
Command3.Caption = "Stop"
Exit Sub
End If
If Command3.Caption = "Stop" Then

Pimp.Close
Exit Sub
End If
End Sub









Private Sub Form_Load()
Dim i As Integer
For i = 1 To 200
Load Hoe(i)
Next i
End Sub

Private Sub Hoe_Close(Index As Integer)

End Sub

Private Sub Hoe_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String
Dim strGet As String
Dim spc2 As Long
Dim page As String
Hoe(Index).GetData strData
If Mid(strData, 1, 3) = "GET" Then
strGet = InStr(strData, "GET ")
spc2 = InStr(strGet + 5, strData, " ")
page = Trim(Mid(strData, strGet + 5, spc2 - (strGet + 4)))
If Right(page, 1) = "/" Then page = Left(page, Len(page) - 1)
If page = "/" Then page = "index.html"
If page = "" Then page = "index.html"
Hoe(Index).SendData LoadFile(App.Path & "\" & page)
End If
End Sub

Private Sub Hoe_SendComplete(Index As Integer)
current.Caption = current.Caption - 1
Hoe(Index).Close
End Sub

Private Sub Pimp_ConnectionRequest(ByVal requestID As Long)
Dim i As Integer
For i = 0 To 200
If Hoe(i).State = sckClosed Then
Hoe(i).Close
Hoe(i).Accept (requestID)
current.Caption = current.Caption + 1
Exit Sub
End If
Next i
End Sub

