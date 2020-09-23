VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scroll Bar Example by Nathan Musto"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   5
      Left            =   0
      Max             =   100
      TabIndex        =   2
      Top             =   3600
      Value           =   50
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3495
      LargeChange     =   5
      Left            =   4200
      Max             =   100
      TabIndex        =   0
      Top             =   0
      Value           =   50
      Width           =   255
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This is a sample program showing how to use a
' Vertical Scroll Bar and a Horizontal Scroll Bar.
' It is the programmer's resposibility to write
' code to react with the Scroll Bar.  The Horizontal
' Scroll Bar works the same way, only Left to Right.  I used the
' "Change" subs to update the button when the user
' just presses the arrow buttons.  The "Scroll"
' subs are used so that the button is updated
' while the Bar is being dragged.  If the "Scroll"
' subs are taken out, the button would not
' update until the drag was complete.

Private Sub Command1_Click()
MsgBox "I hope that this example will help you a great deal.  I know it isn't much, but it is sort of a warm up for those people just starting to program.", vbExclamation, "Scroll Bar Example by Nathan Musto"
End Sub

Private Sub Form_Load()
Command1.Caption = "V: " & VScroll1.Value & "; H: " & HScroll1.Value
Command1.Top = VScroll1.Value * 15 * 2
Command1.Left = HScroll1.Value * 15 * 2
End Sub

Private Sub HScroll1_Change()
Command1.Left = HScroll1.Value * 15 * 2
Command1.Caption = "V: " & VScroll1.Value & "; H: " & HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
Command1.Left = HScroll1.Value * 15 * 2
Command1.Caption = "V: " & VScroll1.Value & "; H: " & HScroll1.Value
End Sub

Private Sub VScroll1_Change()
Command1.Top = VScroll1.Value * 15 * 2
Command1.Caption = "V: " & VScroll1.Value & "; H: " & HScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
Command1.Top = VScroll1.Value * 15 * 2
Command1.Caption = "V: " & VScroll1.Value & "; H: " & HScroll1.Value
End Sub
