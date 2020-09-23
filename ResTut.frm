VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   8490
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   180
      TabIndex        =   7
      Text            =   "By Megalos@mail.com"
      Top             =   8160
      Width           =   2295
   End
   Begin VB.CommandButton cmtnav 
      Caption         =   "< Back  "
      Height          =   345
      Index           =   1
      Left            =   5040
      TabIndex        =   6
      Top             =   8130
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   150
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   10755
   End
   Begin VB.CommandButton cmtnav 
      Caption         =   "Forward >"
      Height          =   345
      Index           =   0
      Left            =   6030
      TabIndex        =   1
      Top             =   8130
      Width           =   915
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6225
      Left            =   150
      ScaleHeight     =   6195
      ScaleWidth      =   10725
      TabIndex        =   0
      Top             =   1860
      Width           =   10755
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CC3401&
      Height          =   345
      Left            =   1200
      TabIndex        =   5
      Top             =   120
      Width           =   315
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Lesson"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CC3401&
      Height          =   345
      Left            =   210
      TabIndex        =   4
      Top             =   120
      Width           =   915
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   2100
      Picture         =   "ResTut.frx":0000
      Top             =   0
      Width           =   510
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Working With Resource File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   465
      Left            =   2970
      TabIndex        =   3
      Top             =   30
      Width           =   4605
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'the API declaration that play Wav file
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Dim imgNo As Long

Private Sub Form_Load()
'Me.MouseIcon = LoadResPicture(101, vbResCursor)  'this will load the  icon from the res"Icon" folder
Do_Click 1
LoadCustom "MYSOUND", "c:\tmpfile.$$$" 'this will load sound from res "CUSTOM" folder and then
sndPlaySound "c:\tmpfile.$$$", 1 'play the Wav file using the sndPlaySound API function
End Sub

Private Sub cmtnav_Click(Index As Integer)
Dim what As Long
If Index = 0 Then what = 1 Else what = (-1)
Do_Click what
End Sub

'the LoadCustomfunction
 Sub LoadCustom(Name As String, FileName As String)
   Dim myArray() As Byte
   Dim myFile As Long
   If Dir(FileName) = "" Then
       myArray = LoadResData(Name, "CUSTOM")
       myFile = FreeFile
       Open FileName For Binary Access Write As #myFile
       Put #myFile, , myArray
       Close #myFile
   End If
End Sub

Sub Do_Click(Index As Long)
imgNo = imgNo + Index
If imgNo > 7 Or imgNo < 1 Then imgNo = 1

   LoadCustom "MYIMAGE" & imgNo, "c:\tmpfile." & imgNo & "$$" 'this will copy the GIF/JPG file to c:\tmpfile.$$$
   Picture1.Picture = LoadPicture("c:\tmpfile." & imgNo & "$$")
       Text1 = LoadResString(10 & imgNo)
   Label3.Caption = imgNo
End Sub






