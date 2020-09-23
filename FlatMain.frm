VERSION 5.00
Begin VB.Form FlatMain 
   Caption         =   "Object Flattening"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6855
   Icon            =   "FlatMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   375
      Left            =   4920
      TabIndex        =   18
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   375
      Left            =   4920
      TabIndex        =   17
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   195
      Left            =   2640
      TabIndex        =   15
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   195
      Left            =   2640
      TabIndex        =   14
      Top             =   3360
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   3720
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   4920
      TabIndex        =   9
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4920
      TabIndex        =   8
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   2640
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   1215
      Left            =   1440
      TabIndex        =   1
      Top             =   1200
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "JO :)"
      Height          =   255
      Left            =   6360
      TabIndex        =   19
      Top             =   120
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   6
      X1              =   120
      X2              =   6720
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   5
      X1              =   6720
      X2              =   6720
      Y1              =   720
      Y2              =   4320
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   4
      X1              =   120
      X2              =   120
      Y1              =   720
      Y2              =   4320
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   3
      X1              =   4680
      X2              =   4680
      Y1              =   720
      Y2              =   4320
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   2
      X1              =   2160
      X2              =   2160
      Y1              =   720
      Y2              =   4320
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   1
      X1              =   120
      X2              =   6720
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   120
      X2              =   6720
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scroll Bars"
      Height          =   255
      Index           =   5
      Left            =   4920
      TabIndex        =   16
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Check Boxes"
      Height          =   255
      Index           =   4
      Left            =   2640
      TabIndex        =   13
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Option Boxes"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   10
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Command Buttons"
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   7
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Text Boxes"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scroll Bars"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This project shows one method for flattening objects..."
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "FlatMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''
' Spinflip@graalmail.com  '
'''''''''''''''''''''''''''
'
'
'This code could be implemented in
'a large scale project to add another
'edge in the overall gui


Private Sub Form_Load()
Label1.Caption = Label1.Caption & vbCrLf & "By: Spinflip@graalmail.com"

Flatten (VScroll2.hwnd)
Flatten (Text2.hwnd)
Flatten (Command2.hwnd)
Flatten (Option2.hwnd)
Flatten (Check2.hwnd)
Flatten (Frame2.hwnd)
End Sub
