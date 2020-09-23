VERSION 5.00
Begin VB.UserControl XPCheckBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2040
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   495
   ScaleWidth      =   2040
   ToolboxBitmap   =   "ctlXPCheckBox.ctx":0000
End
Attribute VB_Name = "XPCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type Size
    cx As Long
    cy As Long
End Type

Public Enum CHECKBOX_TYPE
    [CheckBox] = 1
    [RadioButton] = 2
End Enum

Public Enum CHECKBOX_DIRECTION
    [dcLeft] = 1
    [dcRight] = 2
End Enum

Public Enum CHECKBOX_MARKSTYLE
    [msRegular] = 0
    [msCross] = 1
    [msBrokenCross] = 2
End Enum

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function FloodFill Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long

Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNFACE = 15

' Data for the radio button
Private Const RadioPoint As String = "0000000000000000011000011110001111000011000000000"
Private Const RadioOutline As String = "000011110000001100001100010000000010010000000010100000000001100000000001100000000001100000000001010000000010010000000010001100001100000011110000" ' 12x12

Private COLOR_USER_OUTLINE As Long
Private COLOR_USER_MARKER As Long
Private COLOR_USER_MARKER_SEL As Long
Private COLOR_USER_MARKER_DOWN As Long

Private m_Caption As String
Private m_Font As IFontDisp
Private m_Checked As Boolean
Private m_Enabled As Boolean
Private m_AutoSize As Boolean
Private m_Type As CHECKBOX_TYPE
Private m_Direction As CHECKBOX_DIRECTION
Private m_MarkStyle As CHECKBOX_MARKSTYLE

Private blnOutOfRange As Boolean
Private blnHasFocus As Boolean
Private CheckBoxImg(2) As String

Event Click()
Event Checked()

Public Property Let MarkStyle(intStyle As CHECKBOX_MARKSTYLE)
    m_MarkStyle = intStyle

    UpdateCheckBox
    PropertyChanged "MarkStyle"
End Property

Public Property Get MarkStyle() As CHECKBOX_MARKSTYLE
    MarkStyle = m_MarkStyle
End Property

Public Property Let Direction(intDirection As CHECKBOX_DIRECTION)
    m_Direction = intDirection

    UpdateCaption
    UpdateCheckBox

    PropertyChanged "Direction"
End Property

Public Property Get Direction() As CHECKBOX_DIRECTION
    Direction = m_Direction
End Property

Public Property Let Caption(Str As String)
    m_Caption = Str

    If m_AutoSize Then Call ResizeToFitText
    Call UpdateCaption
    PropertyChanged "Caption"
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Checked(State As Boolean)
    m_Checked = State

    If State Then Call CheckOtherRadioCtrls
    Call UpdateCheckBox
    PropertyChanged "Checked"
End Property

Public Property Get Checked() As Boolean
    Checked = m_Checked
End Property

Public Property Let ButtonType(iType As CHECKBOX_TYPE)
    m_Type = iType

    UpdateCheckBox
    PropertyChanged "Type"
End Property

Public Property Get ButtonType() As CHECKBOX_TYPE
    ButtonType = m_Type
End Property

Public Property Let Enabled(State As Boolean)
    m_Enabled = State

    UserControl.Extender.TabStop = m_Enabled
    UserControl.ForeColor = IIf(m_Enabled, 0, GetSysColor(COLOR_BTNSHADOW))

    UpdateCaption
    UpdateCheckBox
    PropertyChanged "Enabled"
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let AutoSize(State As Boolean)
    m_AutoSize = State

    If m_AutoSize Then ResizeToFitText
    UpdateCaption
    PropertyChanged "AutoSize"
End Property

Public Property Get AutoSize() As Boolean
    AutoSize = m_AutoSize
End Property

Public Property Set Font(Font As IFontDisp)
    Set m_Font = Font

    UpdateCaption
    PropertyChanged "Font"
End Property

Public Property Get Font() As IFontDisp
    Set Font = m_Font
End Property

Private Sub UserControl_GotFocus()
    blnHasFocus = True
    UpdateCaption
End Sub

Private Sub UserControl_Initialize()
    COLOR_USER_OUTLINE = RGB(10, 36, 106)
    COLOR_USER_MARKER = RGB(212, 213, 216)
    COLOR_USER_MARKER_SEL = RGB(182, 189, 210)
    COLOR_USER_MARKER_DOWN = RGB(133, 146, 181)

    ' Regular check 'mark'
    CheckBoxImg(0) = "0000001000001110001111101110111110001110000010000"
    ' Alternate check 'mark'  : Cross
    CheckBoxImg(1) = "1100011111011101111100011100011111011101111100011"
    ' Alternate check 'mark' 2: Broken Cross
    CheckBoxImg(2) = "1100011111011101101100000000011011011101111100011"
End Sub

Private Sub UserControl_InitProperties()
    Set m_Font = Ambient.Font
    Set UserControl.Font = m_Font
    m_Caption = Ambient.DisplayName
    m_Enabled = True
    m_Checked = False
    m_AutoSize = False
    m_Type = CheckBox
    m_Direction = dcLeft
    m_MarkStyle = msCross

    UpdateCaption
    UpdateCheckBox
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If m_Type = CheckBox Or (m_Type = RadioButton And Not m_Checked) Then
            m_Checked = m_Checked Xor True
        End If

        CheckOtherRadioCtrls
    End If
End Sub

Private Sub UserControl_LostFocus()
    blnHasFocus = False
    UpdateCaption
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then UpdateCheckBox COLOR_USER_MARKER_DOWN
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnOutOfRange = True
    If m_Enabled Then
        If X >= 0 And X <= UserControl.Width Then
            If Y >= 0 And Y <= UserControl.Height Then
                blnOutOfRange = False
            End If
        End If
    
        If blnOutOfRange And GetCapture() = UserControl.hWnd Then
            Call ReleaseCapture
        ElseIf Not blnOutOfRange And GetCapture() <> UserControl.hWnd Then
            Call SetCapture(UserControl.hWnd)
        End If
        
        UpdateCheckBox IIf(Not blnOutOfRange, COLOR_USER_MARKER_SEL, COLOR_USER_MARKER)
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then
        If Not blnOutOfRange Then
            If m_Type = CheckBox Or (m_Type = RadioButton And Not m_Checked) Then
                m_Checked = m_Checked Xor True
            End If

            CheckOtherRadioCtrls

            RaiseEvent Click
        End If
    
        UpdateCheckBox
    End If
End Sub

Private Sub UserControl_Resize()
    If Not m_AutoSize Then If UserControl.Height < 240 Then UserControl.Height = 240

    UpdateCaption
    UpdateCheckBox
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    m_Checked = PropBag.ReadProperty("Checked", False)
    m_Enabled = PropBag.ReadProperty("Enabled", True)
    m_AutoSize = PropBag.ReadProperty("AutoSize", False)
    m_Type = PropBag.ReadProperty("Type", CheckBox)
    m_Direction = PropBag.ReadProperty("Direction", dcLeft)
    m_MarkStyle = PropBag.ReadProperty("MarkStyle", msCross)
    
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set UserControl.Font = m_Font

    UserControl.Extender.TabStop = m_Enabled
    UserControl.ForeColor = IIf(m_Enabled, 0, GetSysColor(COLOR_BTNSHADOW))

    If m_AutoSize Then ResizeToFitText
    UpdateCaption
    UpdateCheckBox
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", m_Caption, Ambient.DisplayName
    PropBag.WriteProperty "Checked", m_Checked, False
    PropBag.WriteProperty "Enabled", m_Enabled, True
    PropBag.WriteProperty "AutoSize", m_AutoSize, False
    PropBag.WriteProperty "Font", m_Font, Ambient.Font
    PropBag.WriteProperty "Direction", m_Direction, dcLeft
    PropBag.WriteProperty "Type", m_Type, CheckBox
    PropBag.WriteProperty "MarkStyle", m_MarkStyle, msCross
End Sub

Private Sub UpdateCaption()
    Dim sizSize As Size
    Dim lngY As Long, lngX As Long, lngSub  As Long
    Dim rctErase As RECT
    Dim hBrush As Long

    Call GetTextExtentPoint32(UserControl.hDC, m_Caption, Len(m_Caption), sizSize)

    lngY = (ScaleX(UserControl.Height, vbTwips, vbPixels) / 2) - sizSize.cy / 2
    If lngY < 0 Then lngY = 0
    hBrush = CreateSolidBrush(GetSysColor(COLOR_BTNFACE))

    lngX = IIf(m_Direction = dcLeft, 17, ScaleX(UserControl.Width, vbTwips, vbPixels) - (sizSize.cx + 17))
    lngSub = IIf(m_Direction = dcLeft, 0, 13)

    Call SetRect(rctErase, IIf(m_Direction = dcLeft, 16, 0), 0, ScaleX(UserControl.Width, vbTwips, vbPixels) - lngSub, ScaleY(UserControl.Height, vbTwips, vbPixels))
    Call FillRect(UserControl.hDC, rctErase, hBrush)
    Call TextOut(UserControl.hDC, lngX + 1, lngY, m_Caption, Len(m_Caption))
    Call DeleteObject(hBrush)

    If blnHasFocus And m_Enabled Then
        Call SetRect(rctErase, lngX - 1, lngY - 1, lngX + sizSize.cx + 3, lngY + sizSize.cy + 1)
        Call DrawFocusRect(UserControl.hDC, rctErase)
    End If

    UserControl.Refresh
End Sub

Private Sub UpdateCheckBox(Optional FillColor As Long)
    Dim hBrush As Long
    
    Dim intX As Integer
    Dim intY As Integer
    Dim lngClr As Long
    
    Dim lngY As Long, lngX As Long
    Dim intSize As Integer
    Dim rctOutline As RECT

    intSize = 13
    lngY = (ScaleX(UserControl.Height, vbTwips, vbPixels) / 2) - intSize / 2
    lngX = IIf(m_Direction = dcLeft, 0, ScaleX(UserControl.Width, vbTwips, vbPixels) - intSize)

    If lngY < 0 Then lngY = 0

    If IsMissing(FillColor) Or FillColor = 0 Then FillColor = COLOR_USER_MARKER
    If Not m_Enabled Then FillColor = GetSysColor(COLOR_BTNFACE)
    hBrush = CreateSolidBrush(GetSysColor(COLOR_BTNFACE))

    Call SetRect(rctOutline, lngX, 0, lngX + intSize, ScaleY(UserControl.Height, vbTwips, vbPixels))
    Call FillRect(UserControl.hDC, rctOutline, hBrush)
    Call DeleteObject(hBrush)

    If m_Type = CheckBox Then
        hBrush = CreateSolidBrush(FillColor)

        Call SetRect(rctOutline, lngX, lngY, lngX + intSize, lngY + intSize)
        Call FillRect(UserControl.hDC, rctOutline, hBrush)
        Call DeleteObject(hBrush)

        hBrush = CreateSolidBrush(IIf(m_Enabled, COLOR_USER_OUTLINE, GetSysColor(COLOR_BTNSHADOW)))
    
        Call FrameRect(UserControl.hDC, rctOutline, hBrush)
        Call DeleteObject(hBrush)
    Else
        Dim hOldBrush As Long

        lngClr = IIf(m_Enabled, COLOR_USER_OUTLINE, GetSysColor(COLOR_BTNSHADOW))
        hBrush = CreateSolidBrush(FillColor)
        hOldBrush = SelectObject(UserControl.hDC, hBrush)
        
        For intY = 0 To 12
            For intX = 1 To 12
                If Mid(RadioOutline, (intY * 12) + intX, 1) = "1" Then
                    Call SetPixel(UserControl.hDC, lngX + intX, (lngY + intY) + 1, lngClr)
                End If
            Next
        Next
        
        Call FloodFill(UserControl.hDC, lngX + 5, lngY + 2, lngClr)
        Call SelectObject(UserControl.hDC, hBrush)
        Call DeleteObject(hBrush)
    End If

    If m_Checked Then
        Dim strImg As String

        strImg = IIf(m_Type = CheckBox, CheckBoxImg(m_MarkStyle), RadioPoint)
        lngClr = IIf(m_Enabled, 0, GetSysColor(COLOR_BTNSHADOW))
    
        For intY = 0 To 6
            For intX = 1 To 7
                If Mid(strImg, (intY * 7) + intX, 1) = "1" Then
                    Call SetPixel(UserControl.hDC, lngX + intX + 2, lngY + intY + 3, lngClr)
                End If
            Next
        Next
    End If

    UserControl.Refresh
End Sub

Private Sub ResizeToFitText()
    Dim sizSize As Size

    Call GetTextExtentPoint32(UserControl.hDC, m_Caption, Len(m_Caption), sizSize)

    UserControl.Width = ScaleX(sizSize.cx + 20, vbPixels, vbTwips)
    UserControl.Height = ScaleY(sizSize.cy + 4, vbPixels, vbTwips)

    Call UserControl_Resize
End Sub

Private Sub CheckOtherRadioCtrls()
    Dim Ctrl As Object
    Dim strOwnName As String
    
    If m_Type = RadioButton Then
        If InStr(Ambient.DisplayName, "(") > 0 Then
            strOwnName = LCase(Left(Ambient.DisplayName, InStr(Ambient.DisplayName, "(") - 1))
    
            For Each Ctrl In UserControl.ParentControls
                If TypeName(Ctrl) = "XPCheckBox" Then
                    If LCase(Ctrl.Name) = strOwnName Then
                        If Ctrl.Index <> UserControl.Extender.Index And Ctrl.Checked Then Ctrl.Checked = False
                    End If
                End If
            Next
        End If
    End If
End Sub
