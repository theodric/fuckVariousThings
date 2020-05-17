VERSION 5.00
Object = "{5A0DDB3C-8039-442B-B424-4E72C21D60F9}#1.0#0"; "AxMrquee.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   Caption         =   "Fuck Various Things v0.01 alpha 2"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12615
   DrawMode        =   16  'Merge Pen
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   12615
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   12000
      TabIndex        =   10
      ToolTipText     =   "Scroll right to left"
      Top             =   2100
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   12330
      TabIndex        =   9
      ToolTipText     =   "Scroll left to right"
      Top             =   2100
      Width           =   255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "<-- Fuck this"
      Height          =   615
      Left            =   9240
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   7
      Text            =   "Enter something to fuck!"
      Top             =   1200
      Width           =   5175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "You"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Stop"
      Height          =   375
      Left            =   11160
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Start"
      Height          =   375
      Left            =   11160
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Horse you rode in on"
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Father"
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mother"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin AXMarqueeCtl.AXMarquee AXMarquee1 
      Height          =   690
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   1217
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scroll      < L   R >"
      Height          =   495
      Left            =   11160
      TabIndex        =   11
      Top             =   1920
      Width           =   1455
      Begin VB.CommandButton Command8 
         Caption         =   "Set"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   195
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
AXMarquee1.Text = "Fuck your mother"
End Sub

Private Sub Command2_Click()
AXMarquee1.Text = "Fuck your father"
End Sub

Private Sub Command3_Click()
AXMarquee1.Text = "Fuck the horse you rode in on (like a Celtic chieftain)"
End Sub

Private Sub Command4_Click()
AXMarquee1.Scrolling = True
End Sub

Private Sub Command5_Click()
AXMarquee1.Scrolling = False
End Sub

Private Sub Command6_Click()
AXMarquee1.Text = "Fuck you!"
End Sub

Private Sub Command7_Click()
AXMarquee1.Text = "Fuck " + Text1.Text
End Sub


Private Sub Command8_Click()
MsgBox "LOL BITCH, I THREW A MODAL - direction-setting not implemented yet"
End Sub

Private Sub Form_Load()
AXMarquee1.Scrolling = True
AXMarquee1.Text = "Please choose something to fuck or input your own!"
Option2.Enabled = True
End Sub

