VERSION 5.00
Begin VB.UserControl AXMarquee 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000D&
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4605
   PropertyPages   =   "Marquee.ctx":0000
   ScaleHeight     =   182
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   307
   ToolboxBitmap   =   "Marquee.ctx":0011
   Begin VB.PictureBox picBlankCol 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   420
      Picture         =   "Marquee.ctx":010B
      ScaleHeight     =   46
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   2
      Top             =   672
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox picCaps 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   -2148
      Picture         =   "Marquee.ctx":06BD
      ScaleHeight     =   35.752
      ScaleMode       =   0  'User
      ScaleWidth      =   889.6
      TabIndex        =   1
      Top             =   2130
      Width           =   13350
   End
   Begin VB.PictureBox picMsg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   0
      Top             =   1485
      Width           =   1170
   End
   Begin VB.Timer tAni 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   204
      Top             =   156
   End
End
Attribute VB_Name = "AXMarquee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ActiveX Marquee Control"
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Enum ScrollModeValue
  R_to_L = 0
  L_to_R = 1
End Enum

'Vars for tracking BMP size and position
Private lBMPWidth   As Long     'Total width of the Message Bitmap to be drawn on the background
Private bRestart    As Boolean
Private lCtlWidth   As Long     'Corrected bitmap width for drawing - rounded up to multiple of 5

'Y position of where the Message Bitmap will be drawn on the background
Const SRC_Y = 0

'Height of the control - don't allow it to change
Const CTL_HEIGHT = 683    'Twips

'Default Property Values:
Const m_def_ScrollMode = R_to_L
Const m_def_Text = "ActiveX Marquee"
Const m_def_Scrolling = False

'Property Variables:
Dim m_ScrollMode As ScrollModeValue   'Tracks which direction the control scrolls from.
Dim m_Text As String                  'Holds the message text to be displayed.
Dim m_Scrolling As Boolean               'Tracks whether Scrolling is enabled or disabled.

Private Sub picCaps_Click()

End Sub

Private Sub tAni_Timer()
  Static lX           As Long   'Absolute X position to track message bitmap
  Static lX2          As Long   'X position on the control to draw the message
  Static lSrcOffset   As Long   'Offset into the Message bitmap
  Static lSrcWidth    As Long   'Width from the offset in the Message bitmap to draw

  If bRestart Then
    'Determine which side to scroll from
    If m_ScrollMode = R_to_L Then
      'Scroll Right to Left
      lX = lCtlWidth - BULB_WIDTH
      lSrcOffset = 0
      lSrcWidth = BULB_WIDTH
    Else
      'Assume scroll Left to Right
      lX = BULB_WIDTH
      lSrcOffset = BULB_WIDTH
      lSrcWidth = BULB_WIDTH
    End If

    bRestart = False
  End If  'If bRestart
  
  If m_ScrollMode = R_to_L Then
    If lX > 0 Then
      lX2 = lX
      If lCtlWidth - lX <= lBMPWidth Then
        lSrcWidth = lCtlWidth - lX
      Else
        lSrcWidth = lBMPWidth
      End If
    Else ' assume lx <= 0
      lX2 = 0
      lSrcOffset = Abs(lX)
      lSrcWidth = lBMPWidth - lSrcOffset
    End If
  Else  'Assume m_ScrollMode = L_to_R
    If lX < lCtlWidth Then
      If lX <= lBMPWidth Then
        lX2 = 0
        lSrcWidth = lX
        lSrcOffset = lBMPWidth - lX
      Else
        lX2 = lX2 + BULB_WIDTH
        lSrcWidth = lBMPWidth
        lSrcOffset = 0
      End If
    Else  'assume lx >= lctlwidth
      If lX > lBMPWidth Then
        lX2 = lX2 + BULB_WIDTH
        lSrcWidth = lBMPWidth
      Else
        lSrcOffset = lBMPWidth - lX
        lSrcWidth = lCtlWidth
      End If
    End If
  End If
  
  UserControl.PaintPicture picMsg.Picture, lX2, SRC_Y, , , _
                           lSrcOffset, , lSrcWidth, , _
                           vbSrcCopy
  
  If m_ScrollMode = R_to_L Then
    If lSrcOffset + BULB_WIDTH = lBMPWidth Then
      bRestart = True
    Else
      lX = lX - BULB_WIDTH
    End If
  Else  'Assume m_ScrollMode = L_to_R
    If lX2 + BULB_WIDTH = lCtlWidth Then
      bRestart = True
    Else
      lX = lX + BULB_WIDTH
    End If
  End If
  
End Sub

Private Sub UserControl_Initialize()
  InitBMPStruct
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  m_ScrollMode = m_def_ScrollMode
  m_Text = m_def_Text
  m_Scrolling = m_def_Scrolling
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  ScrollMode = PropBag.ReadProperty("ScrollMode", m_def_ScrollMode)
  Text = PropBag.ReadProperty("Text", m_def_Text)
  Scrolling = PropBag.ReadProperty("Scrolling", m_def_Scrolling)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("ScrollMode", m_ScrollMode, m_def_ScrollMode)
  Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
  
  Call PropBag.WriteProperty("Scrolling", m_Scrolling, m_def_Scrolling)
End Sub

Private Sub UserControl_Resize()
  
  'Don't allow the control to change height
  UserControl.Height = CTL_HEIGHT
  
  'Determine the closest LED to begin drawing from
  lCtlWidth = UserControl.ScaleWidth - UserControl.ScaleWidth Mod 5
  
  'Repaint the unlit LED grid
  DrawBackground
  
End Sub

Public Property Get Text() As String
Attribute Text.VB_Description = "Text string to display on the marquee"
Attribute Text.VB_ProcData.VB_Invoke_Property = ";Text"
  Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
  m_Text = New_Text
  PropertyChanged "Text"
  
  'Force a reset of the Timer painting code since the direction changed.
  If m_Scrolling Then
    tAni.Enabled = False
    bRestart = True
    DrawBackground
    BuildTheBmp (m_Text)
    tAni.Enabled = True
  Else
    tAni.Enabled = False
    bRestart = False
  End If

End Property

Public Property Get Scrolling() As Boolean
Attribute Scrolling.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Scrolling.VB_ProcData.VB_Invoke_Property = ";Behavior"
  Scrolling = m_Scrolling
End Property

Public Property Let Scrolling(ByVal bScrolling As Boolean)
  
  m_Scrolling = bScrolling
  
  PropertyChanged "Scrolling"
  
  If m_Scrolling Then
    DrawBackground
    BuildTheBmp (m_Text)
    tAni.Enabled = True
  Else
    tAni.Enabled = False
    bRestart = False
  End If
  
End Property

Public Property Get ScrollMode() As ScrollModeValue
  ScrollMode = m_ScrollMode
End Property

Public Property Let ScrollMode(ByVal New_ScrollMode As ScrollModeValue)
  m_ScrollMode = New_ScrollMode
  PropertyChanged "ScrollMode"
  
  'Force a reset of the Timer painting code since the direction changed.
  If m_Scrolling Then
    tAni.Enabled = False
    bRestart = True
    DrawBackground
    BuildTheBmp (m_Text)
    tAni.Enabled = True
  Else
    tAni.Enabled = False
    bRestart = False
  End If

End Property

Private Sub DrawBackground()
  Dim lColX As Long
  
  With UserControl
        
    'Turn this on so that what is drawn becomes part of the UserControl's picture.
    .AutoRedraw = True
    
    For lColX = 0 To .ScaleWidth Step 5  'Unlit columns are 5 pixels wide
    
      .PaintPicture picBlankCol.Picture, lColX, 0, _
                    aCharSpace.Width, , _
                    aCharSpace.Left, 0, _
                    aCharSpace.Width
    
    Next lColX
    
    'Turn off so painting performance is faster
    .AutoRedraw = False
    
  End With 'UserControl
  
End Sub

Private Function BuildTheBmp(sText As String) As Long
  Dim lChar     As Long     'Character in the string that we are working on.
  Dim lOffset   As Long     'Tracks the offset into the destination bitmap.
  Dim lCharVal  As Long     'Value of the character at the current offset.
  Dim lCounter  As Long     'Temp counter
  Dim lMsgLength As Long    'Length of the message string
  
  'No support for lower case yet...  Convert all msgs to uppercase.
  sText = UCase$(sText)
  lMsgLength = Len(sText)
  
  With picMsg
  
    'Set to true so the drawing will become part of the picture property.
    .AutoRedraw = True
    
    'Calculating the width of the picture first by accessing the array values in memory is
    'much faster than setting the .Width property each time through the loops below.
    For lChar = 1 To lMsgLength
      lCharVal = Asc(Mid$(sText, lChar, 1))
      If lCharVal = 32 Then 'A space
        For lCounter = 1 To 4
          lOffset = lOffset + aCharSpace.Width
        Next lCounter
      
      ElseIf lCharVal >= 65 And lCharVal <= 90 Then
        lOffset = lOffset + aChars(lCharVal).Width  'Make the Picture wide enough to handle the bitmap
      End If
      
    Next lChar
    
    'Set the picture control to the total width of the message to be created.
    .Width = lOffset + aCharSpace.Width
    
    lOffset = 0
    
    For lChar = 1 To lMsgLength
      
      'Get the ASCII value of the character - This is the index into the bmp array.
      lCharVal = Asc(Mid$(sText, lChar, 1))
      
      If lCharVal = 32 Then 'A space
      
        For lCounter = 1 To 4
          .PaintPicture picCaps.Picture, lOffset, 0, _
                        aCharSpace.Width, , _
                        aCharSpace.Left, 0, _
                        aCharSpace.Width
          
          lOffset = lOffset + aCharSpace.Width

        Next lCounter
              
      ElseIf lCharVal >= 65 And lCharVal <= 90 Then
            
        'Paint the region containing the desired character onto the Msg picturebox at
        'at offset lOffset.
        .PaintPicture picCaps.Picture, lOffset, 0, _
                      aChars(lCharVal).Width, , _
                      aChars(lCharVal).Left, 0, _
                      aChars(lCharVal).Width
                      
        'Increment lOffset by the width of the last Bmp painted on the Msg picturebox.
        lOffset = lOffset + aChars(lCharVal).Width
      
      Else
        Debug.Print "Unsupported character entered - " & Mid$(sText, lChar, 1) & "ASCII = " & Asc(Mid$(sText, lChar, 1))
      
      End If
      
    Next lChar
    
    'Add a blank row of LEDs to the end of the message
    .PaintPicture picCaps.Picture, lOffset, 0, _
                  aCharSpace.Width, , _
                  aCharSpace.Left, 0, _
                  aCharSpace.Width
                  
    lOffset = lOffset + aCharSpace.Width
    
    'Now that we're done drawing turn this off for better paint performance.
    .AutoRedraw = False
    
    .Picture = picMsg.Image
    
  End With  'picMsg
  
    lBMPWidth = lOffset
  
  BuildTheBmp = 0
  
  bRestart = True
End Function
