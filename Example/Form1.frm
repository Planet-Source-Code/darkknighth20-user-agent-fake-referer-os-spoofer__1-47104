VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spoofer by DarkKnight"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1200
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Caption         =   "&Referer"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   3375
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "http://google.com"
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "&Directories"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   3375
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "Directory/Directory/Directory/Index.php"
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Host"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   3375
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Text            =   "Domain.com"
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Incoming Data"
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3375
      Begin VB.TextBox Text1 
         Height          =   1815
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Disconnect"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Connect"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Winsock1.Close
'Closes the winsock just incase it was connected already.
'Doing so prevents errors.
Form1.Caption = "Connecting"
Winsock1.Connect Text2, "80"
'Connects to the host
End Sub

Private Sub Command2_Click()
Winsock1.Close
Form1.Caption = "Closed"
End Sub

Private Sub Form_Load()
MsgBox "This example was made by DarkKnight. It can be pretty useful if you use what is taught on here in a useful way."
End Sub

Private Sub Winsock1_Connect()
Form1.Caption = "Connected"
Winsock1.SendData "GET /" & Text3 & " HTTP/1.0" & Chr(13) & Chr(10) & _
"Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, application/x-comet, */*" & Chr(13) & Chr(10) & _
"Accept-Language: en" & Chr(13) & Chr(10) & _
"Accept-Encoding: gzip , deflate" & Chr(13) & Chr(10) & _
"Cache-Control: no-cache" & Chr(13) & Chr(10) & _
"Proxy-Connection: Keep-Alive" & Chr(13) & Chr(10) & _
"User-Agent: Mozilla/4.0 (compatible; MSIE 9.99; Linux 2.4.18-4 i686) Opera 6.0  [en]" & Chr(13) & Chr(10) & _
"Referer: " & Text4 & vbCrLf & vbCrLf

'----------------------------------------------
'Well, putting "'"s would mess the above up so heres a copy
'----------------------------------------------
'Form1.Caption = "Connected"

'Winsock1.SendData "GET /" & Text3 & "HTTP/1.0" & Chr(13) & Chr(10) & _
'#The above requests the page from the host

'"Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, application/x-comet, */*" & Chr(13) & Chr(10) & _
'"Accept-Language: en" & Chr(13) & Chr(10) & _
'"Accept-Encoding: gzip , deflate" & Chr(13) & Chr(10) & _
'"Cache-Control: no-cache" & Chr(13) & Chr(10) & _
'"Proxy-Connection: Keep-Alive" & Chr(13) & Chr(10) & _
'#The above crap (except for the GET / HTTP/1.0) is useless.
'#I'm just trying to make this very browser like.
'#It's also not what this project is about :)

'"User-Agent: Mozilla/4.0 (compatible; MSIE 9.99; Linux 2.4.18-4 i686) Opera 6.0  [en]" & Chr(13) & Chr(10) & _
'#The good part. The above says that we are using
'#Internet explorer 9.99 and linux.

'"Referer: " & Text4 & vbCrLf & vbCrLf
'#This tells the website where we came from.
'#Changing this can be useful ;)
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim Wdata As String
Winsock1.GetData Wdata, vbString
'#Lets us get the nice data so we can read the page

Wdata = Replace(Wdata, vbLf, vbCrLf)
'#Replace the gay squares to vbCrLf so that it's not all
'#Gay and on the same line

Text1 = Wdata
Form1.Caption = "Data"
End Sub

