VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   ScaleHeight     =   485
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   644
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Clear!"
      Height          =   495
      Left            =   7440
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CheckBox chkBackslash 
      Caption         =   "Backslash"
      Height          =   255
      Left            =   7320
      TabIndex        =   4
      Top             =   960
      Width           =   2175
   End
   Begin VB.CheckBox chkSlash 
      Caption         =   "Slash"
      Height          =   255
      Left            =   7320
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.CheckBox chkHorizontal 
      Caption         =   "Horizontal"
      Height          =   255
      Left            =   7320
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.CheckBox chkVertical 
      Caption         =   "Vertical"
      Height          =   255
      Left            =   7320
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   6855
      Left            =   240
      ScaleHeight     =   453
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   453
      TabIndex        =   0
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Mirror Paint
' by Remerico Cruz
' http://walalang-.blogspot.com (Tagalog)
'
' February 27, 2006, 7:03 PM
'
' After seeing the cool "mirror brush" effect in MacPaint 2.0, I decided to emulate the effect in VB just to
' see if I can do it. ;) The program will "mirror" your drawing on the other side of the canvas, so it'll
' have a reversed symmetrical mirror image.

Public mVertical As Boolean, mHorizontal As Boolean, mSlash As Boolean, mBackslash As Boolean
Public oldX As Long, oldY As Long

Private Sub chkBackslash_Click()
mBackslash = Not mBackslash
End Sub

Private Sub chkHorizontal_Click()
mHorizontal = Not mHorizontal
End Sub

Private Sub chkSlash_Click()
mSlash = Not mSlash
End Sub

Private Sub chkVertical_Click()
mVertical = Not mVertical
End Sub

Private Sub Command1_Click()
Picture1.Cls
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Button = 1 Then ' If mousebutton is pressed (equals to 1), then..
  
    Picture1.Line (oldX, oldY)-(X, Y) ' draw your line
  
    ' when the checkboxes are enabled, the other mirrored lines will appear also.
    If mVertical = True Then Picture1.Line (Picture1.ScaleWidth - oldX, oldY)-(Picture1.ScaleWidth - X, Y)
    If mHorizontal = True Then Picture1.Line (oldX, Picture1.ScaleHeight - oldY)-(X, Picture1.ScaleHeight - Y)
        
    If mSlash = True Then Picture1.Line (Picture1.ScaleWidth - oldY, Picture1.ScaleHeight - oldX)-(Picture1.ScaleWidth - Y, Picture1.ScaleHeight - X)
    If mBackslash = True Then Picture1.Line (oldY, oldX)-(Y, X)
    
    If mVertical = True And mSlash = True Then Picture1.Line (Picture1.ScaleWidth - oldY, oldX)-(Picture1.ScaleWidth - Y, X)
    If mHorizontal = True And mBackslash = True Then Picture1.Line (oldY, Picture1.ScaleHeight - oldX)-(Y, Picture1.ScaleHeight - X)
    
    If (mHorizontal = True And mVertical = True) Or _
       (mSlash = True And mBackslash = True) Then Picture1.Line (Picture1.ScaleWidth - oldX, Picture1.ScaleHeight - oldY)-(Picture1.ScaleWidth - X, Picture1.ScaleHeight - Y)
       
  End If
  
  oldX = X
  oldY = Y
End Sub
