VERSION 5.00
Begin VB.UserControl Button 
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3180
   PropertyPages   =   "Button.ctx":0000
   ScaleHeight     =   2295
   ScaleWidth      =   3180
   ToolboxBitmap   =   "Button.ctx":0025
   Begin VB.Image HOTSPOT 
      Height          =   495
      Left            =   960
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Captions 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      Height          =   195
      Left            =   1350
      TabIndex        =   0
      Top             =   360
      Width           =   555
   End
   Begin VB.Image DOWN_TOP 
      Height          =   90
      Left            =   720
      Picture         =   "Button.ctx":0337
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   4290
   End
   Begin VB.Image DOWN_RIGHT 
      Height          =   3390
      Left            =   1800
      Picture         =   "Button.ctx":17A1
      Stretch         =   -1  'True
      Top             =   2760
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image DOWN_MIDDLE 
      Height          =   3210
      Left            =   3000
      Picture         =   "Button.ctx":298B
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   4290
   End
   Begin VB.Image DOWN_LEFT 
      Height          =   3390
      Left            =   1560
      Picture         =   "Button.ctx":2F8B5
      Stretch         =   -1  'True
      Top             =   3240
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image DOWN_BOTTOM 
      Height          =   90
      Left            =   720
      Picture         =   "Button.ctx":30A9F
      Stretch         =   -1  'True
      Top             =   5040
      Visible         =   0   'False
      Width           =   4290
   End
   Begin VB.Image MIDDLE 
      Height          =   1695
      Left            =   360
      Picture         =   "Button.ctx":31F09
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2775
   End
   Begin VB.Image RIGHT 
      Height          =   3435
      Left            =   3480
      Picture         =   "Button.ctx":63C17
      Stretch         =   -1  'True
      Top             =   -1320
      Width           =   90
   End
   Begin VB.Image TOP 
      Height          =   90
      Left            =   -120
      Picture         =   "Button.ctx":64E3D
      Stretch         =   -1  'True
      Top             =   960
      Width           =   4695
   End
   Begin VB.Image LEFT 
      Height          =   3555
      Left            =   2280
      Picture         =   "Button.ctx":66487
      Stretch         =   -1  'True
      Top             =   -840
      Width           =   90
   End
   Begin VB.Image BOTTOM 
      Height          =   90
      Left            =   -120
      Picture         =   "Button.ctx":676AD
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   4695
   End
End
Attribute VB_Name = "BUTTON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Created soley by Richard Taylor
'For planet-source-code
'
'You can use this control in any program
'that you upload to pscode only or
'program that you distribute
'
'You must leave all the commented lines in
'control or you may suffer legal action
'
''''''Please take the time to vote for this
'control on pscode, If you can take the
'time to look at and use the control then
'it would be worth just leaving a comment
'on the web site and vote, just that
'30 seconds of your life spent writing
'or clicking on a web page. I think
'that its deffinatly worth the time
'As the more votes and feedback I get then
'the more I can upgrade the control so
'that its easier to use, looks better,
'works better
'Thanks for downloading
'
'Richard Taylor

Event Click()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=HOTSPOT,HOTSPOT,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=HOTSPOT,HOTSPOT,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=HOTSPOT,HOTSPOT,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."




Private Sub HOTSPOT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
Press_Down True
End Sub

Private Sub HOTSPOT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
Press_Down False
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
LEFT.LEFT = 0
LEFT.TOP = 0
LEFT.Height = UserControl.Height
RIGHT.LEFT = UserControl.Width - RIGHT.Width
RIGHT.TOP = 0
RIGHT.Height = UserControl.Height
TOP.LEFT = LEFT.Width
TOP.Width = UserControl.Width - LEFT.Width - RIGHT.Width
TOP.TOP = 0
BOTTOM.TOP = UserControl.Height - BOTTOM.Height
BOTTOM.Width = UserControl.Width - LEFT.Width - RIGHT.Width
BOTTOM.LEFT = LEFT.Width
MIDDLE.TOP = TOP.Height
MIDDLE.Height = UserControl.Height - TOP.Height - BOTTOM.Height
MIDDLE.Width = UserControl.Width - LEFT.Width - RIGHT.Width
MIDDLE.LEFT = LEFT.Width
DOWN_LEFT.LEFT = 0
DOWN_LEFT.TOP = 0
DOWN_LEFT.Height = UserControl.Height
DOWN_RIGHT.LEFT = UserControl.Width - DOWN_RIGHT.Width
DOWN_RIGHT.TOP = 0
DOWN_RIGHT.Height = UserControl.Height
DOWN_TOP.LEFT = DOWN_LEFT.Width
DOWN_TOP.Width = UserControl.Width - DOWN_LEFT.Width - DOWN_RIGHT.Width
DOWN_TOP.TOP = 0
DOWN_BOTTOM.TOP = UserControl.Height - DOWN_BOTTOM.Height
DOWN_BOTTOM.Width = UserControl.Width - DOWN_LEFT.Width - DOWN_RIGHT.Width
DOWN_BOTTOM.LEFT = DOWN_LEFT.Width
DOWN_MIDDLE.TOP = DOWN_TOP.Height
DOWN_MIDDLE.Height = UserControl.Height - DOWN_TOP.Height - DOWN_BOTTOM.Height
DOWN_MIDDLE.Width = UserControl.Width - DOWN_LEFT.Width - DOWN_RIGHT.Width
DOWN_MIDDLE.LEFT = DOWN_LEFT.Width
Captions.TOP = UserControl.Height / 2 - (Captions.Height / 2)
Captions.LEFT = UserControl.Width / 2 - (Captions.Width / 2)
HOTSPOT.Height = UserControl.Height
HOTSPOT.Width = UserControl.Width
HOTSPOT.TOP = 0
HOTSPOT.LEFT = 0
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Captions,Captions,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = Captions.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Captions.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Captions,Captions,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Captions.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Captions.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Captions,Captions,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = Captions.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    Captions.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Captions,Captions,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = Captions.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    Captions.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Captions,Captions,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = Captions.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    Captions.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Captions,Captions,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = Captions.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    Captions.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Captions,Captions,-1,WordWrap
Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "Returns/sets a value that determines whether a control expands to fit the text in its Caption."
    WordWrap = Captions.WordWrap
End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
    Captions.WordWrap() = New_WordWrap
    PropertyChanged "WordWrap"
End Property


Private Sub UserControl_InitProperties()
Set m_MouseIcon = LoadPicture("")
m_MousePointer = m_def_MousePointer
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Captions.Caption = PropBag.ReadProperty("Caption", "Caption")
    Set Captions.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Captions.FontBold = PropBag.ReadProperty("FontBold", 0)
    Captions.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    Captions.FontName = PropBag.ReadProperty("FontName", "")
    Captions.FontSize = PropBag.ReadProperty("FontSize", 0)
    Set m_MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    m_MousePointer = PropBag.ReadProperty("MousePointer", m_def_MousePointer)
    Captions.WordWrap = PropBag.ReadProperty("WordWrap", False)
    HOTSPOT.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    HOTSPOT.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
End Sub
Public Function About()
frmAbout.Show
frmAbout.SetFocus
End Function
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", Captions.Caption, "Caption")
    Call PropBag.WriteProperty("Font", Captions.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", Captions.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", Captions.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", Captions.FontName, "")
    Call PropBag.WriteProperty("FontSize", Captions.FontSize, 0)
    Call PropBag.WriteProperty("MouseIcon", m_MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", m_MousePointer, m_def_MousePointer)
    Call PropBag.WriteProperty("WordWrap", Captions.WordWrap, False)
    Call PropBag.WriteProperty("MousePointer", HOTSPOT.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", HOTSPOT.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
End Sub

Private Function Press_Down(Hide As Boolean)
If Hide = False Then
DOWN_LEFT.Visible = False
DOWN_RIGHT.Visible = False
DOWN_TOP.Visible = False
DOWN_BOTTOM.Visible = False
DOWN_MIDDLE.Visible = False
LEFT.Visible = True
RIGHT.Visible = True
TOP.Visible = True
BOTTOM.Visible = True
MIDDLE.Visible = True
Else
DOWN_LEFT.Visible = True
DOWN_RIGHT.Visible = True
DOWN_TOP.Visible = True
DOWN_BOTTOM.Visible = True
DOWN_MIDDLE.Visible = True
LEFT.Visible = False
RIGHT.Visible = False
TOP.Visible = False
BOTTOM.Visible = False
MIDDLE.Visible = False
End If
End Function


Private Sub HOTSPOT_Click()
    RaiseEvent Click
End Sub

Private Sub HOTSPOT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = HOTSPOT.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    HOTSPOT.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=HOTSPOT,HOTSPOT,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = HOTSPOT.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set HOTSPOT.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

