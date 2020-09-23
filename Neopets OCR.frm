VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmNeopetsOCR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Neopets OCR"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   2640
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox picCode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   720
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   1
      Top             =   240
      Width           =   1200
   End
   Begin VB.CommandButton cmdReadCode 
      Caption         =   "Read Code"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblCode 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Coded by Machiavellian"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
End
Attribute VB_Name = "frmNeopetsOCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Neopets OCR - Machiavellian
'
' This program reads the new distorted Neopets magic code
' Do whatever you wish with it, but if you incorporate any or
' all of it into your own programs, credit would be nice.

Private Sub cmdReadCode_Click()
    lblCode.Caption = strCode
End Sub

Private Function strCode() As String
    Dim strLetter As String
    Dim strBuffer As String
    Dim strInput As String
    Dim strData As String
    Dim strCharacter As String

    Dim intScore As Integer
    Dim intHighScore As Integer
    Dim intCount As Integer

    Dim ax As Byte
    Dim ay As Byte
    Dim by As Byte

    Dim Image() As Byte

    Image() = Inet.OpenURL("http://www.neopets.com/rscheck.phtml", icByteArray)

    Do
    DoEvents
    Loop While Inet.StillExecuting

    Open App.Path & "\" & "Code.gif" For Binary As #1
    Put #1, , Image()
    Close #1

    picCode.Picture = LoadPicture("Code.gif")

    For z% = 0 To 2
    strLetter = ""
    strBuffer = ""
    intHighScore = 0
    ax = 255
    ay = 255

    For y% = 0 To 30
    For x% = 0 To 22
        If picCode.Point(x% + z% * 24, y%) = 8421504 And picCode.Point(x% + z% * 24 + 1, y%) = 8421504 And picCode.Point(x% + z% * 24 + 2, y%) = 8421504 Then
            If ax > x% Then ax = x%
            If ay > y% Then ay = y%
            If by < y% Then by = y%

            strBuffer = strBuffer & "000"
            x% = x% + 2
        Else
            strBuffer = strBuffer & " "
        End If
    Next x%
    Next y%

    For y% = ay To by
        strLetter = strLetter & Mid(strBuffer, ax + 1 + y% * 23, 18)
    Next y%

    Open App.Path & "\OCR.dat" For Input As #1
        Do While Not EOF(1)
            Line Input #1, strInput

            strData = Split(strInput, ":")(1)
            intScore = 0
            intCount = 0

            For x% = 1 To Len(strLetter)
                If Mid(strLetter, x%, 1) = "0" Then
                    intCount = intCount + 1
                End If
                
                If Mid(strData, x%, 1) = "0" Then
                    intCount = intCount - 1
                End If
            
                If Mid(strLetter, x%, 1) = "0" And Mid(strData, x%, 1) = "0" Then
                    intScore = intScore + 1
                End If
            Next x%

            intScore = intScore - Abs(intCount) / 2
            
            If intScore > intHighScore Then
                intHighScore = intScore
                strCharacter = Split(strInput, ":")(0)
            End If
        Loop
    Close #1

    strCode = strCode & strCharacter
    Next z%
End Function
