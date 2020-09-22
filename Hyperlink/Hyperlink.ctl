VERSION 5.00
Begin VB.UserControl Hyperlink 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1755
   ForeColor       =   &H00000000&
   MousePointer    =   1  'Arrow
   PropertyPages   =   "Hyperlink.ctx":0000
   ScaleHeight     =   915
   ScaleWidth      =   1755
   ToolboxBitmap   =   "Hyperlink.ctx":002E
   Begin VB.Timer myTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   630
      Top             =   435
   End
   Begin VB.Label lblHyperlink 
      AutoSize        =   -1  'True
      Caption         =   "Hyperlink AX II"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1065
   End
End
Attribute VB_Name = "Hyperlink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'******************************************************************
'Date: 02/13/00
'
'Project:     HyperlinkAX
'Description: Label control that acts like a hyperlink
'             with hover, sound, etc...
'Programmer:  Shannon Harmon (shannohh@vpcusa.com)
'Copyright:   Virtual PC's, Inc. (c)2000
'             All rights reserved.
'Properties:  Quite a few, look them up with your object browser.
'Methods:     No public methods.
'Language:    Written for Microsoft(tm) Visual Basic(tm) version
'             6, with service pack 3 installed but may work in
'             other versions with little or no modification.
'License:     You may modify this code to fit your needs,
'             but you may not redistribute it in any
'             fashion except as compiled in your executable
'             in binary format.  You may not redistribute this
'             source code.  Failure to do so is a violation of
'             copyright law.  This code is provided 'as is' by
'             Virtual PC's, Inc., without any warranties as to
'             performance, fitness, mechantability, and any other
'             warranty (whether expressed or implied).
'Notice:      You must set your hyperlink to a valid email or
'             web url string such as: http://www.yourlink.com or
'             for email: mailto://person@server.com
'             You could also use this control to open a document
'             on your system with it's default program by setting
'             the Hyperlink property to a valid file.  Although,
'             it will not return any errors if the file does not
'             have a default program for it's extension or does
'             not exist.
'******************************************************************

Option Explicit

'Private types
Private Type POINTAPI
    x As Long
    y As Long
End Type

'Public enums
Public Enum Appearance
  vbFlat = 0
  vb3D = 1
End Enum

Public Enum BorderStyle
  vbNone = 0
  vbFixedSingle = 1
End Enum

Public Enum SoundID
  vbMusica = 101
  vbUtopia = 102
End Enum

'Private constants for sndPlaySound
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_MEMORY = &H4
Private SoundArray() As Byte

'Default property constants:
Const m_def_SoundOn = True
Const m_def_MOUnderline = True
Const m_def_MOPointer = True
Const m_def_SoundID = 101
Const m_def_AutoToolTip = True
Const m_def_HyperLink = "http://www.yourlink.com"
Const m_def_HoverColor = vbRed

'Default property variables:
Dim m_SoundOn As Boolean
Dim m_MOUnderline As Boolean
Dim m_MOPointer As Boolean
Dim m_SoundID As Long
Dim m_AutoToolTip As Boolean
Dim m_HyperLink As String
Dim m_HoverColor As OLE_COLOR
Dim m_Hover As Boolean
Dim m_OriginalColor As OLE_COLOR

'Event declarations:
Event Click() 'MappingInfo=lblHyperlink,lblHyperlink,-1,Click
Attribute Click.VB_Description = "Event that runs when the object is clicked."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lblHyperlink,lblHyperlink,-1,MouseDown
Attribute MouseDown.VB_Description = "Event that occurs when the mouse button is down."
Event MouseOver(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lblHyperlink,lblHyperlink,-1,MouseMove
Event MouseOut()
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lblHyperlink,lblHyperlink,-1,MouseUp
Attribute MouseUp.VB_Description = "Event that occurs when the mouse button has been released."
Event Change() 'MappingInfo=lblHyperlink,lblHyperlink,-1,Change
Attribute Change.VB_Description = "Event that runs when the caption changes."

'Private function declarations
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Byte, ByVal uFlags As Long) As Long

'Play a wav sound form a resource file.  Will not return any errors
'so there is really no need to check for a sound card.
Private Sub PlayResWav(vResID As Variant, sResType As String)
On Error Resume Next
  SoundArray = LoadResData(vResID, sResType)
  sndPlaySound SoundArray(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
End Sub

'Occurs when the mouse has left the control
'Sets the labels color back to normal
'Sets underline back to normal
'Sets pointer back to normal
Private Sub MouseExit()
  m_Hover = False
  lblHyperlink.ForeColor = m_OriginalColor
  If m_MOUnderline Then
    If Not Ambient.Font.Underline Then lblHyperlink.Font.Underline = False
  End If
  If m_MOPointer Then
    lblHyperlink.MousePointer = 99
    lblHyperlink.MouseIcon = LoadResPicture(102, vbResCursor)
  End If
  RaiseEvent MouseOut
End Sub

'Check to see if the control is under the cursor
Private Function UnderMouse() As Boolean
Dim ptMouse As POINTAPI
  GetCursorPos ptMouse
  If WindowFromPoint(ptMouse.x, ptMouse.y) = UserControl.hwnd Then
    UnderMouse = True
  Else
    UnderMouse = False
  End If
End Function

'Resize control/label based on labels autosize propety value
Private Sub Resize()
  If lblHyperlink.AutoSize Then
    UserControl.Width = lblHyperlink.Width
    UserControl.Height = lblHyperlink.Height
  Else
    lblHyperlink.Width = UserControl.Width
    lblHyperlink.Height = UserControl.Height
  End If
End Sub

Private Sub lblHyperlink_Click()
On Error Resume Next
Dim lngMouse As Long, lngRet As Long
  'If there is no hyperlink/mail address to go to just run event only
  If m_HyperLink = "" Then GoTo ProcEvent
  'Store mouse's current state
  lngMouse = Screen.MousePointer
  'If sound is enabled then play the sound
  If m_SoundOn Then PlayResWav m_SoundID, "SOUND"
  'Change pointer to hourglass
  Screen.MousePointer = vbHourglass
  'Execute the link's default program
  lngRet = ShellExecute(0&, "Open", m_HyperLink, "", vbNullString, 1)
  'Change mousepointer back to whatever it was
  Screen.MousePointer = lngMouse
ProcEvent:
  RaiseEvent Click
End Sub

Private Sub lblHyperlink_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub lblHyperlink_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  'If not currently hovering over the control
  If Not m_Hover Then
    'Set hovering to true
    m_Hover = True
    'If custom pointer on mouseover enabled then apply browser cursor
    If m_MOPointer Then
        lblHyperlink.MousePointer = 99
        lblHyperlink.MouseIcon = LoadResPicture(101, vbResCursor)
    End If
    'Store original color value of the label
    m_OriginalColor = lblHyperlink.ForeColor
    'Change the color of the label to the hover color property
    lblHyperlink.ForeColor = m_HoverColor
    'If underline on mouseover enabled then underline the caption
    If m_MOUnderline Then
        lblHyperlink.Font.Underline = True
    End If
    'If autotooltip enabled then set the tooltip to current hyperlink property
    If m_AutoToolTip Then
      lblHyperlink.ToolTipText = m_HyperLink
    'Else use the default tooltiptext property
    Else
      lblHyperlink.ToolTipText = UserControl.Extender.ToolTipText
    End If
    'Enable timer to watch for the mouse to exit the control
    myTimer.Enabled = True
    RaiseEvent MouseOver(Button, Shift, x, y)
  End If
End Sub

Private Sub lblHyperlink_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub lblHyperlink_Change()
  Resize
  RaiseEvent Change
End Sub

Private Sub myTimer_Timer()
  'If the mouse is not over our control then run mouse exit code
  If Not UnderMouse Then
    MouseExit
    myTimer.Enabled = False
  End If
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  m_HyperLink = m_def_HyperLink
  m_HoverColor = m_def_HoverColor
  m_AutoToolTip = m_def_AutoToolTip
  m_SoundOn = m_def_SoundOn
  m_SoundID = m_def_SoundID
  m_MOUnderline = m_def_MOUnderline
  m_MOPointer = m_def_MOPointer
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Set lblHyperlink.Font = PropBag.ReadProperty("Font", Ambient.Font)
  lblHyperlink.Enabled = PropBag.ReadProperty("Enabled", True)
  lblHyperlink.Caption = PropBag.ReadProperty("Caption", "Hyperlink AX II")
  lblHyperlink.Alignment = PropBag.ReadProperty("Alignment", 0)
  lblHyperlink.AutoSize = PropBag.ReadProperty("AutoSize", True)
  m_HyperLink = PropBag.ReadProperty("HyperLink", m_def_HyperLink)
  m_HoverColor = PropBag.ReadProperty("HoverColor", m_def_HoverColor)
  lblHyperlink.Appearance = PropBag.ReadProperty("Appearance", 1)
  lblHyperlink.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
  lblHyperlink.ForeColor = PropBag.ReadProperty("ForeColor", Ambient.ForeColor)
  lblHyperlink.BackColor = PropBag.ReadProperty("BackColor", Ambient.BackColor)
  m_AutoToolTip = PropBag.ReadProperty("AutoToolTip", m_def_AutoToolTip)
  m_SoundOn = PropBag.ReadProperty("SoundOn", m_def_SoundOn)
  m_SoundID = PropBag.ReadProperty("SoundID", m_def_SoundID)
  m_MOUnderline = PropBag.ReadProperty("MOUnderline", m_def_MOUnderline)
  m_MOPointer = PropBag.ReadProperty("MOPointer", m_def_MOPointer)
End Sub

Private Sub UserControl_Resize()
  Resize
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Enabled", lblHyperlink.Enabled, True)
  Call PropBag.WriteProperty("Font", lblHyperlink.Font, Ambient.Font)
  Call PropBag.WriteProperty("Caption", lblHyperlink.Caption, "Hyperlink AX")
  Call PropBag.WriteProperty("Alignment", lblHyperlink.Alignment, 0)
  Call PropBag.WriteProperty("AutoSize", lblHyperlink.AutoSize, True)
  Call PropBag.WriteProperty("HoverColor", m_HoverColor, m_def_HoverColor)
  Call PropBag.WriteProperty("HyperLink", m_HyperLink, m_def_HyperLink)
  Call PropBag.WriteProperty("Appearance", lblHyperlink.Appearance, 1)
  Call PropBag.WriteProperty("BorderStyle", lblHyperlink.BorderStyle, 0)
  Call PropBag.WriteProperty("ForeColor", lblHyperlink.ForeColor, &H404040)
  Call PropBag.WriteProperty("BackColor", lblHyperlink.BackColor, &H8000000F)
  Call PropBag.WriteProperty("AutoToolTip", m_AutoToolTip, m_def_AutoToolTip)
  Call PropBag.WriteProperty("SoundOn", m_SoundOn, m_def_SoundOn)
  Call PropBag.WriteProperty("SoundID", m_SoundID, m_def_SoundID)
  Call PropBag.WriteProperty("MOUnderline", m_MOUnderline, m_def_MOUnderline)
  Call PropBag.WriteProperty("MOPointer", m_MOPointer, m_def_MOPointer)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblHyperlink,lblHyperlink,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
  Enabled = lblHyperlink.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  lblHyperlink.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblHyperlink,lblHyperlink,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
  Set Font = lblHyperlink.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set lblHyperlink.Font = New_Font
  PropertyChanged "Font"
  Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,http://www.vsoftusa.com
Public Property Get HyperLink() As String
Attribute HyperLink.VB_Description = "Returns/sets the website or email address to open when clicked."
  HyperLink = m_HyperLink
End Property

Public Property Let HyperLink(ByVal New_HyperLink As String)
  m_HyperLink = Trim(New_HyperLink)
  PropertyChanged "HyperLink"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblHyperlink,lblHyperlink,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets an objects caption."
  Caption = lblHyperlink.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
  lblHyperlink.Caption() = New_Caption
  PropertyChanged "Caption"
  Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblHyperlink,lblHyperlink,-1,Alignment
Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets whether or not an object is aligned left, middle or right."
  Alignment = lblHyperlink.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
  lblHyperlink.Alignment() = New_Alignment
  PropertyChanged "Alignment"
  Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblHyperlink,lblHyperlink,-1,AutoSize
Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Returns/sets whether or not an object is sized to fit it's caption."
Attribute AutoSize.VB_ProcData.VB_Invoke_Property = "Settings"
  AutoSize = lblHyperlink.AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
  lblHyperlink.AutoSize() = New_AutoSize
  PropertyChanged "AutoSize"
  Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbRed
Public Property Get HoverColor() As OLE_COLOR
Attribute HoverColor.VB_Description = "Returns/sets the hover color used to display text and graphics in an object."
  HoverColor = m_HoverColor
End Property

Public Property Let HoverColor(ByVal New_HoverColor As OLE_COLOR)
  m_HoverColor = New_HoverColor
  PropertyChanged "HoverColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblHyperlink,lblHyperlink,-1,Appearance
Public Property Get Appearance() As Appearance
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
  Appearance = lblHyperlink.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Appearance)
  lblHyperlink.Appearance() = New_Appearance
  PropertyChanged "Appearance"
  Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblHyperlink,lblHyperlink,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
  BorderStyle = lblHyperlink.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyle)
  lblHyperlink.BorderStyle() = New_BorderStyle
  PropertyChanged "BorderStyle"
  Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblHyperlink,lblHyperlink,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
  ForeColor = lblHyperlink.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  lblHyperlink.ForeColor() = New_ForeColor
  PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblHyperlink,lblHyperlink,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
  BackColor = lblHyperlink.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  lblHyperlink.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get AutoToolTip() As Boolean
Attribute AutoToolTip.VB_Description = "Returns/sets the auto tooltip property."
  AutoToolTip = m_AutoToolTip
End Property

Public Property Let AutoToolTip(ByVal New_AutoToolTip As Boolean)
  m_AutoToolTip = New_AutoToolTip
  PropertyChanged "AutoToolTip"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get SoundOn() As Boolean
Attribute SoundOn.VB_Description = "Sets/Returns sound when clicked On/Off"
  SoundOn = m_SoundOn
End Property

Public Property Let SoundOn(ByVal New_SoundOn As Boolean)
  m_SoundOn = New_SoundOn
  PropertyChanged "SoundOn"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get MOUnderline() As Boolean
  MOUnderline = m_MOUnderline
End Property

Public Property Let MOUnderline(ByVal New_MOUnderline As Boolean)
  m_MOUnderline = New_MOUnderline
  PropertyChanged "MOUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get MOPointer() As Boolean
  MOPointer = m_MOPointer
End Property

Public Property Let MOPointer(ByVal New_MOPointer As Boolean)
  m_MOPointer = New_MOPointer
  PropertyChanged "MOPointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,101
Public Property Get SoundID() As SoundID
Attribute SoundID.VB_Description = "Sets/Returns current sound to be used when clicked."
  SoundID = m_SoundID
End Property

Public Property Let SoundID(ByVal New_SoundID As SoundID)
  m_SoundID = New_SoundID
  PropertyChanged "SoundID"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub About()
Attribute About.VB_Description = "Show about dialog box."
Attribute About.VB_UserMemId = -552
   Load frmAbout
   frmAbout.Show vbModal
End Sub

