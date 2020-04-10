VERSION 5.00
Begin VB.UserControl SGELabel 
   CanGetFocus     =   0   'False
   ClientHeight    =   1440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3045
   ScaleHeight     =   1440
   ScaleWidth      =   3045
   ToolboxBitmap   =   "SGELabel.ctx":0000
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   285
      TabIndex        =   0
      Top             =   210
      Width           =   1845
   End
End
Attribute VB_Name = "SGELabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Enumeracao

''Enumeracao Alignment
'Public Enum ENUM_ALIGNMENT
'    LeftJustify
'    RightJustify
'    Center
'End Enum

'Enumeracao Back Style
Public Enum ENUM_BACK_STYLE
    Opaque
    Transparent
End Enum

'Enumeracao Border Style
Public Enum ENUM_BORDER_STYLE
    None
    FixedSingle
End Enum

'Enumeracao Border Style
Public Enum ENUM_OLEDROPMODE
    None
    Manual
End Enum

'Enumeracao Appearance
Public Enum ENUM_APPEARANCE
    Flat
    DDD
End Enum

'Event Declarations:
Event OLECompleteDrag(Effect As Long) 'MappingInfo=Label1,Label1,-1,OLECompleteDrag
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a manual or automatic drag/drop has been completed or canceled."
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Label1,Label1,-1,OLEDragDrop
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual."
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer) 'MappingInfo=Label1,Label1,-1,OLEDragOver
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=Label1,Label1,-1,OLEGiveFeedback
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=Label1,Label1,-1,OLESetData
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=Label1,Label1,-1,OLEStartDrag
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."
Event Click() 'MappingInfo=Label1,Label1,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=Label1,Label1,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Label1,Label1,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Label1,Label1,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Label1,Label1,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Change() 'MappingInfo=Label1,Label1,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Label1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Label1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = Label1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Label1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,BackStyle
Public Property Get BackStyle() As ENUM_BACK_STYLE
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = Label1.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As ENUM_BACK_STYLE)
    Label1.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,BorderStyle
Public Property Get BorderStyle() As ENUM_BORDER_STYLE
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = Label1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As ENUM_BORDER_STYLE)
    Label1.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    Label1.Refresh
End Sub

Private Sub Label1_Click()
    RaiseEvent Click
End Sub

Private Sub Label1_DblClick()
    RaiseEvent DblClick
End Sub

'Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
'
'    Call Controle_DragDrop(Me, Source, X, Y)
'    giProxButtonUp = 0
''    RaiseEvent DragDrop(Source, X, Y)
'End Sub

Public Sub Drag(iModo As Integer)
    UserControl.Extender.Drag iModo
End Sub

Public Sub Move(iLeft As Integer, iTop As Integer)
    Call UserControl.Extender.Move(iLeft, iTop)
End Sub

Public Property Let Left(ByVal iValor As Integer)
    UserControl.Extender.Left = iValor
End Property

Public Property Get Left() As Integer
    Left = UserControl.Extender.Left
End Property

Public Property Let Top(ByVal iValor As Integer)
    UserControl.Extender.Top = iValor
End Property

Public Property Get Top() As Integer
    Top = UserControl.Extender.Top
End Property

Public Property Let Width(ByVal iValor As Integer)
    UserControl.Extender.Width = iValor
    Label1.Width = iValor
End Property

Public Property Get Width() As Integer
    Width = UserControl.Extender.Width
End Property

Public Property Let Height(ByVal iValor As Integer)
    UserControl.Extender.Height = iValor
    Label1.Height = iValor
End Property

Public Property Get Height() As Integer
    Height = UserControl.Extender.Height
End Property

Public Property Set Container(ByVal objValor As Object)
    Set UserControl.Extender.Container = objValor
End Property

Public Property Get Container() As Object
    Set Container = UserControl.Extender.Container
End Property

Public Property Let TabIndex(ByVal iValor As Integer)
End Property

Public Property Get TabIndex() As Integer
    TabIndex = 1
End Property

Public Property Get Name() As String
    Name = UserControl.Extender.Name
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Index() As Integer

On Error GoTo Erro_Index

    Index = UserControl.Extender.Index
    
    Exit Property
    
Erro_Index:

    Index = -1
    
    Exit Property
    
End Property

Public Property Let Visible(ByVal bValor As Boolean)
    UserControl.Extender.Visible = bValor
End Property

Public Property Get Visible() As Boolean
    Visible = UserControl.Extender.Visible
End Property

'Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    Set gobjControleDrag = Me
'    giProxButtonUp = 0
'    Call Controle_MouseDown(Me, Button, Shift, X, Y)
''    RaiseEvent MouseDown(Button, Shift, X, Y)
'
'End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Alignment
Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = Label1.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    Label1.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Appearance
Public Property Get Appearance() As ENUM_APPEARANCE
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = Label1.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As ENUM_APPEARANCE)
    Label1.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,AutoSize
Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Determines whether a control is automatically resized to display its entire contents."
    AutoSize = Label1.AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    Label1.AutoSize() = New_AutoSize
    PropertyChanged "AutoSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
Attribute Caption.VB_UserMemId = 0
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption = New_Caption
    PropertyChanged "Caption"
End Property

Private Sub Label1_Change()
    RaiseEvent Change
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Label1,Label1,-1,DataFormat
'Public Property Get DataFormat() As IStdDataFormatDisp
'    Set DataFormat = Label1.DataFormat
'End Property
'
'Public Property Set DataFormat(ByVal New_DataFormat As IStdDataFormatDisp)
'    Set Label1.DataFormat = New_DataFormat
'    PropertyChanged "DataFormat"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Label1,Label1,-1,DataMember
'Public Property Get DataMember() As String
'    DataMember = Label1.DataMember
'End Property
'
'Public Property Let DataMember(ByVal New_DataMember As String)
'    Label1.DataMember() = New_DataMember
'    PropertyChanged "DataMember"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Label1,Label1,-1,DataSource
'Public Property Get DataSource() As DataSource
'    Set DataSource = Label1.DataSource
'End Property
'
'Public Property Set DataSource(ByVal New_DataSource As DataSource)
'    Set Label1.DataSource = New_DataSource
'    PropertyChanged "DataSource"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,LinkItem
Public Property Get LinkItem() As String
Attribute LinkItem.VB_Description = "Returns/sets the data passed to a destination control in a DDE conversation with another application."
    LinkItem = Label1.LinkItem
End Property

Public Property Let LinkItem(ByVal New_LinkItem As String)
    Label1.LinkItem() = New_LinkItem
    PropertyChanged "LinkItem"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,LinkMode
Public Property Get LinkMode() As LinkModeConstants
Attribute LinkMode.VB_Description = "Returns/sets the type of link used for a DDE conversation and activates the connection."
    LinkMode = Label1.LinkMode
End Property

Public Property Let LinkMode(ByVal New_LinkMode As LinkModeConstants)
    Label1.LinkMode() = New_LinkMode
    PropertyChanged "LinkMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,LinkTimeout
Public Property Get LinkTimeout() As Integer
Attribute LinkTimeout.VB_Description = "Returns/sets the amount of time a control waits for a response to a DDE message."
    LinkTimeout = Label1.LinkTimeout
End Property

Public Property Let LinkTimeout(ByVal New_LinkTimeout As Integer)
    Label1.LinkTimeout() = New_LinkTimeout
    PropertyChanged "LinkTimeout"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,LinkTopic
Public Property Get LinkTopic() As String
Attribute LinkTopic.VB_Description = "Returns/sets the source application and topic for a destination control."
    LinkTopic = Label1.LinkTopic
End Property

Public Property Let LinkTopic(ByVal New_LinkTopic As String)
    Label1.LinkTopic() = New_LinkTopic
    PropertyChanged "LinkTopic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = Label1.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set Label1.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = Label1.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    Label1.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,OLEDropMode
Public Property Get OLEDropMode() As ENUM_OLEDROPMODE
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
    OLEDropMode = Label1.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As ENUM_OLEDROPMODE)
    Label1.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,RightToLeft
Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "Determines text display direction and control visual appearance on a bidirectional system."
    RightToLeft = Label1.RightToLeft
End Property

Public Property Let RightToLeft(ByVal New_RightToLeft As Boolean)
    Label1.RightToLeft() = New_RightToLeft
    PropertyChanged "RightToLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = Label1.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    Label1.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,UseMnemonic
Public Property Get UseMnemonic() As Boolean
Attribute UseMnemonic.VB_Description = "Returns/sets a value that specifies whether an & in a Label's Caption property defines an access key."
    UseMnemonic = Label1.UseMnemonic
End Property

Public Property Let UseMnemonic(ByVal New_UseMnemonic As Boolean)
    Label1.UseMnemonic() = New_UseMnemonic
    PropertyChanged "UseMnemonic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,WhatsThisHelpID
Public Property Get WhatsThisHelpID() As Long
Attribute WhatsThisHelpID.VB_Description = "Returns/sets an associated context number for an object."
    WhatsThisHelpID = Label1.WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal New_WhatsThisHelpID As Long)
    Label1.WhatsThisHelpID() = New_WhatsThisHelpID
    PropertyChanged "WhatsThisHelpID"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,WordWrap
Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "Returns/sets a value that determines whether a control expands to fit the text in its Caption."
    WordWrap = Label1.WordWrap
End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
    Label1.WordWrap() = New_WordWrap
    PropertyChanged "WordWrap"
End Property

Private Sub UserControl_Initialize()
    
    'Posiciona o Label no canto superior esquerdo
    Label1.Top = 0
    Label1.Left = 0
 
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Label1.Appearance = PropBag.ReadProperty("Appearance", 1)
    Label1.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    'Label1.ForeColor =PropBag.ReadProperty("ForeColor", &H80000012)
    Me.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Label1.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Label1.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    Label1.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Label1.Alignment = PropBag.ReadProperty("Alignment", 0)
    Label1.AutoSize = PropBag.ReadProperty("AutoSize", False)
    Label1.Caption = PropBag.ReadProperty("Caption", "")
'    Set DataFormat = PropBag.ReadProperty("DataFormat", Nothing)
'    Label1.DataMember = PropBag.ReadProperty("DataMember", "")
'    Set DataSource = PropBag.ReadProperty("DataSource", Nothing)
    Label1.LinkItem = PropBag.ReadProperty("LinkItem", "")
    Label1.LinkMode = PropBag.ReadProperty("LinkMode", 0)
    Label1.LinkTimeout = PropBag.ReadProperty("LinkTimeout", 50)
    Label1.LinkTopic = PropBag.ReadProperty("LinkTopic", "")
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    Label1.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Label1.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    Label1.RightToLeft = PropBag.ReadProperty("RightToLeft", False)
    Label1.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    Label1.UseMnemonic = PropBag.ReadProperty("UseMnemonic", True)
    Label1.WhatsThisHelpID = PropBag.ReadProperty("WhatsThisHelpID", 0)
    Label1.WordWrap = PropBag.ReadProperty("WordWrap", False)

'''    Label1.FontBold = PropBag.ReadProperty("FontBold", 0)
'''    Label1.FontItalic = PropBag.ReadProperty("FontItalic", 0)
'''    Label1.FontName = PropBag.ReadProperty("FontName", "")
'''    Label1.FontSize = PropBag.ReadProperty("FontSize", 0)
'''    Label1.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
'''    Label1.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
End Sub

Private Sub UserControl_Resize()

    'Coloca o mesmo tamanho do controle para o Label
    Label1.Height = UserControl.Height
    Label1.Width = UserControl.Width
    
End Sub

Private Sub UserControl_Show()

    Label1.ForeColor = Me.ForeColor
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", Label1.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", Label1.Enabled, True)
    Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", Label1.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", Label1.BorderStyle, 0)
    Call PropBag.WriteProperty("Alignment", Label1.Alignment, 0)
    Call PropBag.WriteProperty("Appearance", Label1.Appearance, 1)
    Call PropBag.WriteProperty("AutoSize", Label1.AutoSize, False)
    Call PropBag.WriteProperty("Caption", Label1.Caption, "")
'    Call PropBag.WriteProperty("DataFormat", DataFormat, Nothing)
'    Call PropBag.WriteProperty("DataMember", Label1.DataMember, "")
'    Call PropBag.WriteProperty("DataSource", DataSource, Nothing)
    Call PropBag.WriteProperty("LinkItem", Label1.LinkItem, "")
    Call PropBag.WriteProperty("LinkMode", Label1.LinkMode, 0)
    Call PropBag.WriteProperty("LinkTimeout", Label1.LinkTimeout, 50)
    Call PropBag.WriteProperty("LinkTopic", Label1.LinkTopic, "")
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", Label1.MousePointer, 0)
    Call PropBag.WriteProperty("OLEDropMode", Label1.OLEDropMode, 0)
    Call PropBag.WriteProperty("RightToLeft", Label1.RightToLeft, False)
    Call PropBag.WriteProperty("ToolTipText", Label1.ToolTipText, "")
    Call PropBag.WriteProperty("UseMnemonic", Label1.UseMnemonic, True)
    Call PropBag.WriteProperty("WhatsThisHelpID", Label1.WhatsThisHelpID, 0)
    Call PropBag.WriteProperty("WordWrap", Label1.WordWrap, False)
    
'''    Call PropBag.WriteProperty("FontBold", Label1.FontBold, 0)
'''    Call PropBag.WriteProperty("FontItalic", Label1.FontItalic, 0)
'''    Call PropBag.WriteProperty("FontName", Label1.FontName, "")
'''    Call PropBag.WriteProperty("FontSize", Label1.FontSize, 0)
'''    Call PropBag.WriteProperty("FontStrikethru", Label1.FontStrikethru, 0)
'''    Call PropBag.WriteProperty("FontUnderline", Label1.FontUnderline, 0)

End Sub
Private Sub Label1_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub Label1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub Label1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub Label1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub Label1_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub Label1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = Label1.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    Label1.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = Label1.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    Label1.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = Label1.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    Label1.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = Label1.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    Label1.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
    FontStrikethru = Label1.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    Label1.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
    FontUnderline = Label1.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    Label1.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,LinkExecute
Public Sub LinkExecute(ByVal Command As String)
Attribute LinkExecute.VB_Description = "Sends a command string to the source application in a DDE conversation."
    Label1.LinkExecute Command
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,LinkRequest
Public Sub LinkRequest()
Attribute LinkRequest.VB_Description = "Asks the source DDE application to update the contents of a Label, PictureBox, or Textbox control."
    Label1.LinkRequest
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,LinkSend
Public Sub LinkSend()
Attribute LinkSend.VB_Description = "Transfers contents of PictureBox to destination application in DDE conversation."
    Label1.LinkSend
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,LinkPoke
Public Sub LinkPoke()
Attribute LinkPoke.VB_Description = "Transfers contents of Label, PictureBox, or TextBox to source application in DDE conversation."
    Label1.LinkPoke
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,OLEDrag
Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
    Label1.OLEDrag
End Sub

