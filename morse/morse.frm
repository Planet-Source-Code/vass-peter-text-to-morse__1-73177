VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Morse "
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open Play dialog"
      Height          =   735
      Left            =   2760
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convert to morse code!!"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   2520
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim normaltext As String
Dim morsetext As String
Dim letter As String
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Sub Command1_Click()
normaltext = Text1
If normaltext = "" Then
MsgBox "This field can't be blank!!", vbCritical, "Error"
Exit Sub
End If
For m = 1 To Len(normaltext)
letter = Mid(normaltext, m, 1)
Select Case letter

Case Is = "a"
morsetext = morsetext & " " & " .–"


Case Is = "b"
morsetext = morsetext & " " & " –..."


Case Is = "c"
morsetext = morsetext & " " & " –.–."


Case Is = "d"
morsetext = morsetext & " " & " –.."

Case Is = "e"
morsetext = morsetext & " " & " ."

Case Is = "f"
morsetext = morsetext & " " & " ..–."

Case Is = "g"
morsetext = morsetext & " " & " ––."

Case Is = "h"
morsetext = morsetext & " " & " ...."

Case Is = "i"
morsetext = morsetext & " " & " .."

Case Is = "j"
morsetext = morsetext & " " & " .---"

Case Is = "k"
morsetext = morsetext & " " & " -.-"

Case Is = "l"
morsetext = morsetext & " " & " .-.."

Case Is = "m"
morsetext = morsetext & " " & " --"

Case Is = "n"
morsetext = morsetext & " " & " -."

Case Is = "o"
morsetext = morsetext & " " & " ---"

Case Is = "p"
morsetext = morsetext & " " & " .--."

Case Is = "q"
morsetext = morsetext & " " & " --.-"

Case Is = "r"
morsetext = morsetext & " " & " .-."

Case Is = "s"
morsetext = morsetext & " " & " ..."

Case Is = "t"
morsetext = morsetext & " " & " -"

Case Is = "u"
morsetext = morsetext & " " & " ..-"

Case Is = "v"
morsetext = morsetext & " " & " ...-"

Case Is = "X"
morsetext = morsetext & " " & " -.--"

Case Is = "y"
morsetext = morsetext & " " & " -.--"

Case Is = "z"
morsetext = morsetext & " " & " --.."

End Select

Next m
Open App.Path & "\morse.txt" For Output As #1
Print #1, "********CREATED WHIT MORSE CONVERTER BY VASS PÉTER********"
Print #1, morsetext
Close #1
MsgBox "Succesfuly converted!!Filepath: " & App.Path & "\morse.txt", vbInformation, "SUCCES!!"
morsetext = ""
Text1 = ""
Text1.SetFocus
End Sub

Private Sub Command2_Click()
CommonDialog1.ShowOpen
Open CommonDialog1.FileTitle For Input As #1
While EOF(1) = 0
Line Input #1, morsetext
Wend
For i = 1 To Len(morsetext)
letter = Mid(morsetext, i, 1)
If letter = "." Then
Beep 200, 300
wait 200
End If
If letter = "-" Then
Beep 200, 600
wait 200
End If
Next i
Close #1
End Sub

Private Sub Form_Load()
MsgBox "Created by Vass Péter whit VB 6.0!!", vbInformation, " "
End Sub
