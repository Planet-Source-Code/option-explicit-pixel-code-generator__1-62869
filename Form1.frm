VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   322
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   305
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   240
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   5
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy code to Redraw routine, rerun, then click!"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Code"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   240
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************
'* Bitmap Code Generator                                                 *
'* Matthew R. Usner, October 12, 2005.  No copyright, it's too dumb.     *
'*************************************************************************
'* This small, almost totally useless project scans all the pixels in a  *
'* bitmap stored in a picturebox and generates SetPixelV code to repli-  *
'* cate that bitmap.  Automatically generates For loop code if enough    *
'* consecutive pixels are identical to exceed a FORLOOP_THRESHOLD value. *
'* Generated code is written to the clipboard for pasting into other     *
'* projects.  Why did I do something so useless, you ask?  Well, as many *
'* of you know, my hobby at PSC is inflicting usercontrols on you.  I    *
'* prefer things like checkmarks for checkboxes drawn by code as opposed *
'* to being stored in imagelists or the like.  I don't like dependencies *
'* of any kind in controls.  So, I draw them using LineTo or SetPixelV.  *
'* The problem is, I can't even draw a stick figure.  I see these nice   *
'* custom checkmarks in web sites and I couldn't replicate them, so I    *
'* wrote this.  As far as I know checkmarks aren't copyrighted, so hope- *
'* fully this wouldn't be considered stealing!  To use, just use your    *
'* favorite screen grabber to grab the checkmark or whatever, then MS    *
'* Paint or whatever to save just the area you want to disk.  (You may   *
'* want to edit colors, or resize, or some such.) Place this bitmap      *
'* in Picture1 and run the program.  Stop the program, and paste the     *
'* newly generated code into the "Redraw" sub.  Run again and click the  *
'* button under Picture2.  Voila!  You can tweak the code to take out    *
'* parts you don't want fairly easily.  Only try this with SMALL (32x32  *
'* size or less) bitmaps.  I tried with larger bitmaps and while the     *
'* code generates properly, the code for even a smaller photgraph can    *
'* easily exceed 10,000 lines.  A VB procedure can't be larger than 64K. *
'* I included a small calculator icon example with the code already      *
'* loaded into the Redraw routine to get you started.  I wouldn't even   *
'* use this for something the size of the calculator icon but did it     *
'* just to give you the idea.  I'm sure this could be optimized (espec-  *
'* ially the string concatenation) but it runs great on small bitmaps    *
'* and I'm not going to endlessly tweak this.  Let the flames begin :-)  *
'*************************************************************************

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

' fyi only.  Paste into your projects.
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private MyhDC As Long

' a for loop generates if number of consecutive equal pixels is greater than this constant.
Private Const FORLOOP_THRESHOLD As Long = 1

Option Explicit

Private Sub Command1_Click()

   GeneratePixelCodeSnippet

End Sub

Private Sub GeneratePixelCodeSnippet()

'*************************************************************************
'* generates the code to the clipboard.  It gets the value for each      *
'* pixel.  Then it scans for all identical pixels to the end of the row, *
'* and creates a For Loop if the number of consecutive pixels in that    *
'* row exceeds the FORLOOP_THRESHOLD constant.  If there are not enough  *
'* consecutive identical pixels to exceed the threshold, a simple        *
'* no-loop SetPixelV line is generated.                                  *
'*************************************************************************

   Dim X As Long, Y As Long, Z As Long, Pixel1 As Long, Pixel2 As Long
   Dim NumX As Long, NumY As Long
   Dim Snippet As String

   NumX = ScaleX(Picture1.Picture.Width, vbTwips, vbPixels) - 1
   NumY = ScaleY(Picture1.Picture.Height, vbTwips, vbPixels) - 1
   Label1.Caption = "X: " & CStr(NumX) & " Y: " & CStr(NumY)

   For Y = 0 To NumY
      Label2.Caption = "Current Y: " & CStr(Y)
      Me.Refresh
      For X = 0 To NumX
         Pixel1 = GetPixel(Picture1.hdc, X, Y)
         For Z = X + 1 To NumX
            Pixel2 = GetPixel(Picture1.hdc, Z, Y)
            If Pixel1 <> Pixel2 Then
               Z = Z - 1
               Exit For
            End If
         Next Z
         If Z > NumX Then Z = NumX
         If Z - X >= FORLOOP_THRESHOLD Then
            Snippet = Snippet & GenerateForLoopPixelCodeLine(X, Y, Z)
            X = Z
         Else
            Snippet = Snippet & GenerateSinglePixelCodeLine(X, Y)
         End If
      Next X
   Next Y

'  write the generated code to the clipboard.
   Clipboard.Clear
   Clipboard.SetText Snippet

End Sub

Private Function GenerateSinglePixelCodeLine(ByVal X As Long, ByVal Y As Long) As String

   GenerateSinglePixelCodeLine = "SetPixelV MyhDC, " & CStr(X) & ", " & CStr(Y) & ", &H" & Hex(GetPixel(Picture1.hdc, X, Y)) & "&" & vbCrLf

End Function

Private Function GenerateForLoopPixelCodeLine(ByVal X As Long, ByVal Y As Long, ByVal Z As Long) As String

   GenerateForLoopPixelCodeLine = "For i = " & CStr(X) & " To " & CStr(Z) & ": " & _
                                  "SetPixelV MyhDC, i, " & CStr(Y) & ", &H" & _
                                  Hex(GetPixel(Picture1.hdc, X, Y)) & "&" & _
                                  ": Next i" & vbCrLf

End Function

Private Sub Command2_Click()
   Redraw
End Sub

Private Sub Form_Load()
   MyhDC = Picture2.hdc
End Sub

Private Sub Redraw()

'*************************************************************************
'* place generated code in this routine.  I put the calculator icon code *
'* in to give you an example how quickly the code grows.  The best way   *
'* to use this tool is for VERY small areas (like a custom checkmark).   *
'*************************************************************************

Dim i As Long

'------- paste here!
For i = 0 To 55: SetPixelV MyhDC, i, 0, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 1, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 2, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 3, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 4, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 5, &HFFFFFF: Next i
For i = 0 To 1: SetPixelV MyhDC, i, 6, &HFFFFFF: Next i
For i = 2 To 30: SetPixelV MyhDC, i, 6, &H7F7F00: Next i
For i = 31 To 55: SetPixelV MyhDC, i, 6, &HFFFFFF: Next i
SetPixelV MyhDC, 0, 7, &HFFFFFF
SetPixelV MyhDC, 1, 7, &H7F7F00
SetPixelV MyhDC, 2, 7, &HFFFFFF
SetPixelV MyhDC, 3, 7, &HFFFF00
SetPixelV MyhDC, 4, 7, &HFFFFFF
SetPixelV MyhDC, 5, 7, &HFFFF00
SetPixelV MyhDC, 6, 7, &HFFFFFF
SetPixelV MyhDC, 7, 7, &HFFFF00
SetPixelV MyhDC, 8, 7, &HFFFFFF
SetPixelV MyhDC, 9, 7, &HFFFF00
SetPixelV MyhDC, 10, 7, &HFFFFFF
SetPixelV MyhDC, 11, 7, &HFFFF00
SetPixelV MyhDC, 12, 7, &HFFFFFF
SetPixelV MyhDC, 13, 7, &HFFFF00
SetPixelV MyhDC, 14, 7, &HFFFFFF
SetPixelV MyhDC, 15, 7, &HFFFF00
SetPixelV MyhDC, 16, 7, &HFFFFFF
SetPixelV MyhDC, 17, 7, &HFFFF00
SetPixelV MyhDC, 18, 7, &HFFFFFF
SetPixelV MyhDC, 19, 7, &HFFFF00
SetPixelV MyhDC, 20, 7, &HFFFFFF
SetPixelV MyhDC, 21, 7, &HFFFF00
SetPixelV MyhDC, 22, 7, &HFFFFFF
SetPixelV MyhDC, 23, 7, &HFFFF00
SetPixelV MyhDC, 24, 7, &HFFFFFF
SetPixelV MyhDC, 25, 7, &HFFFF00
SetPixelV MyhDC, 26, 7, &HFFFFFF
SetPixelV MyhDC, 27, 7, &HFFFF00
SetPixelV MyhDC, 28, 7, &HFFFFFF
SetPixelV MyhDC, 29, 7, &HFFFF00
SetPixelV MyhDC, 30, 7, &H7F0000
SetPixelV MyhDC, 31, 7, &H0&
For i = 32 To 55: SetPixelV MyhDC, i, 7, &HFFFFFF: Next i
SetPixelV MyhDC, 0, 8, &HFFFFFF
SetPixelV MyhDC, 1, 8, &H7F7F00
SetPixelV MyhDC, 2, 8, &HFFFF00
For i = 3 To 4: SetPixelV MyhDC, i, 8, &H7F7F00: Next i
For i = 5 To 19: SetPixelV MyhDC, i, 8, &H0&: Next i
For i = 20 To 29: SetPixelV MyhDC, i, 8, &H7F7F00: Next i
SetPixelV MyhDC, 30, 8, &H7F0000
SetPixelV MyhDC, 31, 8, &H0&
For i = 32 To 55: SetPixelV MyhDC, i, 8, &HFFFFFF: Next i
SetPixelV MyhDC, 0, 9, &HFFFFFF
SetPixelV MyhDC, 1, 9, &H7F7F00
SetPixelV MyhDC, 2, 9, &HFFFFFF
For i = 3 To 4: SetPixelV MyhDC, i, 9, &H7F7F00: Next i
SetPixelV MyhDC, 5, 9, &H0&
For i = 6 To 19: SetPixelV MyhDC, i, 9, &HBFBFBF: Next i
For i = 20 To 29: SetPixelV MyhDC, i, 9, &H7F7F00: Next i
SetPixelV MyhDC, 30, 9, &H7F0000
SetPixelV MyhDC, 31, 9, &H0&
For i = 32 To 55: SetPixelV MyhDC, i, 9, &HFFFFFF: Next i
SetPixelV MyhDC, 0, 10, &HFFFFFF
SetPixelV MyhDC, 1, 10, &H7F7F00
SetPixelV MyhDC, 2, 10, &HFFFF00
For i = 3 To 4: SetPixelV MyhDC, i, 10, &H7F7F00: Next i
SetPixelV MyhDC, 5, 10, &H0&
For i = 6 To 18: SetPixelV MyhDC, i, 10, &HFFFFFF: Next i
SetPixelV MyhDC, 19, 10, &HBFBFBF
For i = 20 To 29: SetPixelV MyhDC, i, 10, &H7F7F00: Next i
SetPixelV MyhDC, 30, 10, &H7F0000
SetPixelV MyhDC, 31, 10, &H0&
For i = 32 To 55: SetPixelV MyhDC, i, 10, &HFFFFFF: Next i
SetPixelV MyhDC, 0, 11, &HFFFFFF
SetPixelV MyhDC, 1, 11, &H7F7F00
SetPixelV MyhDC, 2, 11, &HFFFFFF
For i = 3 To 4: SetPixelV MyhDC, i, 11, &H7F7F00: Next i
SetPixelV MyhDC, 5, 11, &H0&
For i = 6 To 19: SetPixelV MyhDC, i, 11, &HBFBFBF: Next i
For i = 20 To 29: SetPixelV MyhDC, i, 11, &H7F7F00: Next i
SetPixelV MyhDC, 30, 11, &H7F0000
SetPixelV MyhDC, 31, 11, &H0&
For i = 32 To 55: SetPixelV MyhDC, i, 11, &HFFFFFF: Next i
SetPixelV MyhDC, 0, 12, &HFFFFFF
SetPixelV MyhDC, 1, 12, &H7F7F00
SetPixelV MyhDC, 2, 12, &HFFFF00
For i = 3 To 29: SetPixelV MyhDC, i, 12, &H7F7F00: Next i
SetPixelV MyhDC, 30, 12, &H7F0000
SetPixelV MyhDC, 31, 12, &H0&
For i = 32 To 55: SetPixelV MyhDC, i, 12, &HFFFFFF: Next i
SetPixelV MyhDC, 0, 13, &HFFFFFF
SetPixelV MyhDC, 1, 13, &H7F7F00
SetPixelV MyhDC, 2, 13, &HFFFFFF
For i = 3 To 4: SetPixelV MyhDC, i, 13, &H7F7F00: Next i
For i = 5 To 6: SetPixelV MyhDC, i, 13, &HFFFFFF: Next i
SetPixelV MyhDC, 7, 13, &H0&
SetPixelV MyhDC, 8, 13, &H7F7F00
For i = 9 To 10: SetPixelV MyhDC, i, 13, &HFFFFFF: Next i
SetPixelV MyhDC, 11, 13, &H0&
SetPixelV MyhDC, 12, 13, &H7F7F00
For i = 13 To 14: SetPixelV MyhDC, i, 13, &HFFFFFF: Next i
SetPixelV MyhDC, 15, 13, &H0&
SetPixelV MyhDC, 16, 13, &H7F7F00
For i = 17 To 18: SetPixelV MyhDC, i, 13, &HFFFFFF: Next i
SetPixelV MyhDC, 19, 13, &H0&
SetPixelV MyhDC, 20, 13, &H7F7F00
For i = 21 To 22: SetPixelV MyhDC, i, 13, &HFFFFFF: Next i
SetPixelV MyhDC, 23, 13, &H0&
SetPixelV MyhDC, 24, 13, &H7F7F00
For i = 25 To 26: SetPixelV MyhDC, i, 13, &HFFFFFF: Next i
SetPixelV MyhDC, 27, 13, &H0&
For i = 28 To 29: SetPixelV MyhDC, i, 13, &H7F7F00: Next i
SetPixelV MyhDC, 30, 13, &H7F0000
SetPixelV MyhDC, 31, 13, &H0&
For i = 32 To 55: SetPixelV MyhDC, i, 13, &HFFFFFF: Next i
SetPixelV MyhDC, 0, 14, &HFFFFFF
SetPixelV MyhDC, 1, 14, &H7F7F00
SetPixelV MyhDC, 2, 14, &HFFFF00
For i = 3 To 4: SetPixelV MyhDC, i, 14, &H7F7F00: Next i
SetPixelV MyhDC, 5, 14, &HFFFFFF
SetPixelV MyhDC, 6, 14, &HBFBFBF
SetPixelV MyhDC, 7, 14, &H0&
SetPixelV MyhDC, 8, 14, &H7F7F00
SetPixelV MyhDC, 9, 14, &HFFFFFF
SetPixelV MyhDC, 10, 14, &HBFBFBF
SetPixelV MyhDC, 11, 14, &H0&
SetPixelV MyhDC, 12, 14, &H7F7F00
SetPixelV MyhDC, 13, 14, &HFFFFFF
SetPixelV MyhDC, 14, 14, &HBFBFBF
SetPixelV MyhDC, 15, 14, &H0&
SetPixelV MyhDC, 16, 14, &H7F7F00
SetPixelV MyhDC, 17, 14, &HFFFFFF
SetPixelV MyhDC, 18, 14, &HBFBFBF
SetPixelV MyhDC, 19, 14, &H0&
SetPixelV MyhDC, 20, 14, &H7F7F00
SetPixelV MyhDC, 21, 14, &HFFFFFF
SetPixelV MyhDC, 22, 14, &HBFBFBF
SetPixelV MyhDC, 23, 14, &H0&
SetPixelV MyhDC, 24, 14, &H7F7F00
SetPixelV MyhDC, 25, 14, &HFFFFFF
SetPixelV MyhDC, 26, 14, &HBFBFBF
SetPixelV MyhDC, 27, 14, &H0&
For i = 28 To 29: SetPixelV MyhDC, i, 14, &H7F7F00: Next i
SetPixelV MyhDC, 30, 14, &H7F0000
SetPixelV MyhDC, 31, 14, &H0&
For i = 32 To 55: SetPixelV MyhDC, i, 14, &HFFFFFF: Next i
SetPixelV MyhDC, 0, 15, &HFFFFFF
SetPixelV MyhDC, 1, 15, &H7F7F00
SetPixelV MyhDC, 2, 15, &HFFFFFF
For i = 3 To 4: SetPixelV MyhDC, i, 15, &H7F7F00: Next i
For i = 5 To 7: SetPixelV MyhDC, i, 15, &H0&: Next i
SetPixelV MyhDC, 8, 15, &H7F7F00
For i = 9 To 11: SetPixelV MyhDC, i, 15, &H0&: Next i
SetPixelV MyhDC, 12, 15, &H7F7F00
For i = 13 To 15: SetPixelV MyhDC, i, 15, &H0&: Next i
SetPixelV MyhDC, 16, 15, &H7F7F00
For i = 17 To 19: SetPixelV MyhDC, i, 15, &H0&: Next i
SetPixelV MyhDC, 20, 15, &H7F7F00
For i = 21 To 23: SetPixelV MyhDC, i, 15, &H0&: Next i
SetPixelV MyhDC, 24, 15, &H7F7F00
For i = 25 To 27: SetPixelV MyhDC, i, 15, &H0&: Next i
For i = 28 To 29: SetPixelV MyhDC, i, 15, &H7F7F00: Next i
SetPixelV MyhDC, 30, 15, &H7F0000
SetPixelV MyhDC, 31, 15, &H0&
For i = 32 To 55: SetPixelV MyhDC, i, 15, &HFFFFFF: Next i
SetPixelV MyhDC, 0, 16, &HFFFFFF
SetPixelV MyhDC, 1, 16, &H7F7F00
SetPixelV MyhDC, 2, 16, &HFFFF00
For i = 3 To 29: SetPixelV MyhDC, i, 16, &H7F7F00: Next i
SetPixelV MyhDC, 30, 16, &H7F0000
SetPixelV MyhDC, 31, 16, &H0&
For i = 32 To 55: SetPixelV MyhDC, i, 16, &HFFFFFF: Next i
SetPixelV MyhDC, 0, 17, &HFFFFFF
SetPixelV MyhDC, 1, 17, &H7F7F00
SetPixelV MyhDC, 2, 17, &HFFFFFF
For i = 3 To 4: SetPixelV MyhDC, i, 17, &H7F7F00: Next i
For i = 5 To 6: SetPixelV MyhDC, i, 17, &HFFFFFF: Next i
SetPixelV MyhDC, 7, 17, &H0&
SetPixelV MyhDC, 8, 17, &H7F7F00
For i = 9 To 10: SetPixelV MyhDC, i, 17, &HFFFFFF: Next i
SetPixelV MyhDC, 11, 17, &H0&
SetPixelV MyhDC, 12, 17, &H7F7F00
For i = 13 To 14: SetPixelV MyhDC, i, 17, &HFFFFFF: Next i
SetPixelV MyhDC, 15, 17, &H0&
SetPixelV MyhDC, 16, 17, &H7F7F00
For i = 17 To 18: SetPixelV MyhDC, i, 17, &HFFFFFF: Next i
SetPixelV MyhDC, 19, 17, &H0&
SetPixelV MyhDC, 20, 17, &H7F7F00
For i = 21 To 22: SetPixelV MyhDC, i, 17, &HFFFFFF: Next i
SetPixelV MyhDC, 23, 17, &H0&
SetPixelV MyhDC, 24, 17, &H7F7F00
For i = 25 To 26: SetPixelV MyhDC, i, 17, &HFFFFFF: Next i
SetPixelV MyhDC, 27, 17, &H0&
For i = 28 To 29: SetPixelV MyhDC, i, 17, &H7F7F00: Next i
SetPixelV MyhDC, 30, 17, &H7F0000
SetPixelV MyhDC, 31, 17, &H0&
For i = 32 To 55: SetPixelV MyhDC, i, 17, &HFFFFFF: Next i
SetPixelV MyhDC, 0, 18, &HFFFFFF
SetPixelV MyhDC, 1, 18, &H7F7F00
SetPixelV MyhDC, 2, 18, &HFFFF00
For i = 3 To 4: SetPixelV MyhDC, i, 18, &H7F7F00: Next i
SetPixelV MyhDC, 5, 18, &HFFFFFF
SetPixelV MyhDC, 6, 18, &HBFBFBF
SetPixelV MyhDC, 7, 18, &H0&
SetPixelV MyhDC, 8, 18, &H7F7F00
SetPixelV MyhDC, 9, 18, &HFFFFFF
SetPixelV MyhDC, 10, 18, &HBFBFBF
SetPixelV MyhDC, 11, 18, &H0&
SetPixelV MyhDC, 12, 18, &H7F7F00
SetPixelV MyhDC, 13, 18, &HFFFFFF
SetPixelV MyhDC, 14, 18, &HBFBFBF
SetPixelV MyhDC, 15, 18, &H0&
SetPixelV MyhDC, 16, 18, &H7F7F00
SetPixelV MyhDC, 17, 18, &HFFFFFF
SetPixelV MyhDC, 18, 18, &HBFBFBF
SetPixelV MyhDC, 19, 18, &H0&
SetPixelV MyhDC, 20, 18, &H7F7F00
SetPixelV MyhDC, 21, 18, &HFFFFFF
SetPixelV MyhDC, 22, 18, &HBFBFBF
SetPixelV MyhDC, 23, 18, &H0&
SetPixelV MyhDC, 24, 18, &H7F7F00
SetPixelV MyhDC, 25, 18, &HFFFFFF
SetPixelV MyhDC, 26, 18, &HBFBFBF
SetPixelV MyhDC, 27, 18, &H0&
For i = 28 To 29: SetPixelV MyhDC, i, 18, &H7F7F00: Next i
SetPixelV MyhDC, 30, 18, &H7F0000
SetPixelV MyhDC, 31, 18, &H0&
For i = 32 To 55: SetPixelV MyhDC, i, 18, &HFFFFFF: Next i
SetPixelV MyhDC, 0, 19, &HFFFFFF
SetPixelV MyhDC, 1, 19, &H7F7F00
SetPixelV MyhDC, 2, 19, &HFFFFFF
For i = 3 To 4: SetPixelV MyhDC, i, 19, &H7F7F00: Next i
For i = 5 To 7: SetPixelV MyhDC, i, 19, &H0&: Next i
SetPixelV MyhDC, 8, 19, &H7F7F00
For i = 9 To 11: SetPixelV MyhDC, i, 19, &H0&: Next i
SetPixelV MyhDC, 12, 19, &H7F7F00
For i = 13 To 15: SetPixelV MyhDC, i, 19, &H0&: Next i
SetPixelV MyhDC, 16, 19, &H7F7F00
For i = 17 To 19: SetPixelV MyhDC, i, 19, &H0&: Next i
SetPixelV MyhDC, 20, 19, &H7F7F00
For i = 21 To 23: SetPixelV MyhDC, i, 19, &H0&: Next i
SetPixelV MyhDC, 24, 19, &H7F7F00
For i = 25 To 27: SetPixelV MyhDC, i, 19, &H0&: Next i
For i = 28 To 29: SetPixelV MyhDC, i, 19, &H7F7F00: Next i
SetPixelV MyhDC, 30, 19, &H7F0000
SetPixelV MyhDC, 31, 19, &H0&
For i = 32 To 55: SetPixelV MyhDC, i, 19, &HFFFFFF: Next i
SetPixelV MyhDC, 0, 20, &HFFFFFF
SetPixelV MyhDC, 1, 20, &H7F7F00
SetPixelV MyhDC, 2, 20, &HFFFF00
For i = 3 To 29: SetPixelV MyhDC, i, 20, &H7F7F00: Next i
SetPixelV MyhDC, 30, 20, &H7F0000
SetPixelV MyhDC, 31, 20, &H0&
For i = 32 To 55: SetPixelV MyhDC, i, 20, &HFFFFFF: Next i
SetPixelV MyhDC, 0, 21, &HFFFFFF
SetPixelV MyhDC, 1, 21, &H7F7F00
SetPixelV MyhDC, 2, 21, &HFFFFFF
For i = 3 To 4: SetPixelV MyhDC, i, 21, &H7F7F00: Next i
For i = 5 To 6: SetPixelV MyhDC, i, 21, &HFFFFFF: Next i
SetPixelV MyhDC, 7, 21, &H0&
SetPixelV MyhDC, 8, 21, &H7F7F00
For i = 9 To 10: SetPixelV MyhDC, i, 21, &HFFFFFF: Next i
SetPixelV MyhDC, 11, 21, &H0&
SetPixelV MyhDC, 12, 21, &H7F7F00
For i = 13 To 14: SetPixelV MyhDC, i, 21, &HFFFFFF: Next i
SetPixelV MyhDC, 15, 21, &H0&
SetPixelV MyhDC, 16, 21, &H7F7F00
For i = 17 To 18: SetPixelV MyhDC, i, 21, &HFFFFFF: Next i
SetPixelV MyhDC, 19, 21, &H0&
SetPixelV MyhDC, 20, 21, &H7F7F00
For i = 21 To 26: SetPixelV MyhDC, i, 21, &HFFFFFF: Next i
SetPixelV MyhDC, 27, 21, &H0&
For i = 28 To 29: SetPixelV MyhDC, i, 21, &H7F7F00: Next i
SetPixelV MyhDC, 30, 21, &H7F0000
SetPixelV MyhDC, 31, 21, &H0&
For i = 32 To 55: SetPixelV MyhDC, i, 21, &HFFFFFF: Next i
SetPixelV MyhDC, 0, 22, &HFFFFFF
SetPixelV MyhDC, 1, 22, &H7F7F00
SetPixelV MyhDC, 2, 22, &HFFFF00
For i = 3 To 4: SetPixelV MyhDC, i, 22, &H7F7F00: Next i
SetPixelV MyhDC, 5, 22, &HFFFFFF
SetPixelV MyhDC, 6, 22, &HBFBFBF
SetPixelV MyhDC, 7, 22, &H0&
SetPixelV MyhDC, 8, 22, &H7F7F00
SetPixelV MyhDC, 9, 22, &HFFFFFF
SetPixelV MyhDC, 10, 22, &HBFBFBF
SetPixelV MyhDC, 11, 22, &H0&
SetPixelV MyhDC, 12, 22, &H7F7F00
SetPixelV MyhDC, 13, 22, &HFFFFFF
SetPixelV MyhDC, 14, 22, &HBFBFBF
SetPixelV MyhDC, 15, 22, &H0&
SetPixelV MyhDC, 16, 22, &H7F7F00
SetPixelV MyhDC, 17, 22, &HFFFFFF
SetPixelV MyhDC, 18, 22, &HBFBFBF
SetPixelV MyhDC, 19, 22, &H0&
SetPixelV MyhDC, 20, 22, &H7F7F00
SetPixelV MyhDC, 21, 22, &HFFFFFF
For i = 22 To 26: SetPixelV MyhDC, i, 22, &HBFBFBF: Next i
SetPixelV MyhDC, 27, 22, &H0&
For i = 28 To 29: SetPixelV MyhDC, i, 22, &H7F7F00: Next i
SetPixelV MyhDC, 30, 22, &H7F0000
SetPixelV MyhDC, 31, 22, &H0&
For i = 32 To 55: SetPixelV MyhDC, i, 22, &HFFFFFF: Next i
SetPixelV MyhDC, 0, 23, &HFFFFFF
SetPixelV MyhDC, 1, 23, &H7F7F00
SetPixelV MyhDC, 2, 23, &HFFFFFF
For i = 3 To 4: SetPixelV MyhDC, i, 23, &H7F7F00: Next i
For i = 5 To 7: SetPixelV MyhDC, i, 23, &H0&: Next i
SetPixelV MyhDC, 8, 23, &H7F7F00
For i = 9 To 11: SetPixelV MyhDC, i, 23, &H0&: Next i
SetPixelV MyhDC, 12, 23, &H7F7F00
For i = 13 To 15: SetPixelV MyhDC, i, 23, &H0&: Next i
SetPixelV MyhDC, 16, 23, &H7F7F00
For i = 17 To 19: SetPixelV MyhDC, i, 23, &H0&: Next i
SetPixelV MyhDC, 20, 23, &H7F7F00
For i = 21 To 27: SetPixelV MyhDC, i, 23, &H0&: Next i
For i = 28 To 29: SetPixelV MyhDC, i, 23, &H7F7F00: Next i
SetPixelV MyhDC, 30, 23, &H7F0000
SetPixelV MyhDC, 31, 23, &H0&
For i = 32 To 55: SetPixelV MyhDC, i, 23, &HFFFFFF: Next i
SetPixelV MyhDC, 0, 24, &HFFFFFF
SetPixelV MyhDC, 1, 24, &H7F7F00
SetPixelV MyhDC, 2, 24, &HFFFF00
For i = 3 To 29: SetPixelV MyhDC, i, 24, &H7F7F00: Next i
SetPixelV MyhDC, 30, 24, &H7F0000
SetPixelV MyhDC, 31, 24, &H0&
For i = 32 To 55: SetPixelV MyhDC, i, 24, &HFFFFFF: Next i
SetPixelV MyhDC, 0, 25, &HFFFFFF
For i = 1 To 2: SetPixelV MyhDC, i, 25, &H7F7F00: Next i
For i = 3 To 30: SetPixelV MyhDC, i, 25, &H7F0000: Next i
SetPixelV MyhDC, 31, 25, &H0&
For i = 32 To 55: SetPixelV MyhDC, i, 25, &HFFFFFF: Next i
For i = 0 To 1: SetPixelV MyhDC, i, 26, &HFFFFFF: Next i
For i = 2 To 30: SetPixelV MyhDC, i, 26, &H0&: Next i
For i = 31 To 55: SetPixelV MyhDC, i, 26, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 27, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 28, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 29, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 30, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 31, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 32, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 33, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 34, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 35, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 36, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 37, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 38, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 39, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 40, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 41, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 42, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 43, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 44, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 45, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 46, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 47, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 48, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 49, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 50, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 51, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 52, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 53, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 54, &HFFFFFF: Next i
For i = 0 To 55: SetPixelV MyhDC, i, 55, &HFFFFFF: Next i

End Sub
