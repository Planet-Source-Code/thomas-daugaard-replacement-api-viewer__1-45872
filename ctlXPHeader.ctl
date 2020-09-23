VERSION 5.00
Begin VB.UserControl XPHeader 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   ControlContainer=   -1  'True
   ScaleHeight     =   900
   ScaleWidth      =   2400
   ToolboxBitmap   =   "ctlXPHeader.ctx":0000
   Begin VB.Line lneBorder 
      Index           =   1
      X1              =   210
      X2              =   870
      Y1              =   375
      Y2              =   375
   End
   Begin VB.Line lneBorder 
      Index           =   0
      X1              =   210
      X2              =   870
      Y1              =   225
      Y2              =   225
   End
End
Attribute VB_Name = "XPHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Public Enum enmBorderType
    Inset = 1
    Raised = 2
End Enum

Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_WINDOW = 5

Private m_BorderType As enmBorderType
Private m_Caption As String
Private m_Description As String
Private m_Font As IFontDisp
Private cGradient  As New clsGradient

Private Sub UserControl_Initialize()
    UserControl.BackColor = RGB(255, 255, 255)
End Sub

Private Sub UserControl_InitProperties()
    m_BorderType = Inset
    m_Caption = Ambient.DisplayName
    m_Description = "Add you description text here"
    Set m_Font = Ambient.Font

    ChangeBorder
End Sub

Public Property Let Caption(Str As String)
    m_Caption = Str

    UpdateCaption
    PropertyChanged "Caption"
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Description(Str As String)
    m_Description = Str

    UpdateCaption
    PropertyChanged "Description"
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Set Font(Font As IFontDisp)
    Set m_Font = Font
    Set UserControl.Font = m_Font

    UpdateCaption
    PropertyChanged "Font"
End Property

Public Property Get Font() As IFontDisp
    Set Font = m_Font
End Property

Public Property Let BorderStyle(Style As enmBorderType)
    m_BorderType = Style

    ChangeBorder
    PropertyChanged "BorderStyle"
End Property

Public Property Get BorderStyle() As enmBorderType
    BorderStyle = m_BorderType
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BorderStyle", m_BorderType, Inset
    PropBag.WriteProperty "Font", m_Font, Ambient.Font
    PropBag.WriteProperty "Caption", m_Caption, Ambient.DisplayName
    PropBag.WriteProperty "Description", m_Description, "Add you description text here"
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BorderType = PropBag.ReadProperty("BorderStyle", Inset)
    m_Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    m_Description = PropBag.ReadProperty("Description", "Add you description text here")
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set UserControl.Font = m_Font

    UpdateCaption
    ChangeBorder
End Sub

Private Sub UserControl_Resize()
    Dim lWidth As Long
    Dim lHeight As Long
    
    lWidth = UserControl.Width - 15
    lHeight = UserControl.Height - 15

    With lneBorder(0)
        .X1 = 0
        .X2 = lWidth + 15
        .Y1 = lHeight - 15
        .Y2 = lHeight - 15
    End With
    
    With lneBorder(1)
        .X1 = 0
        .X2 = lWidth + 15
        .Y1 = lHeight
        .Y2 = lHeight
    End With
    
    UpdateCaption
End Sub

Private Sub ChangeBorder()
    Dim lShadow As Long
    Dim lWhite As Long

    lShadow = IIf(m_BorderType = Inset, GetSysColor(COLOR_BTNSHADOW), GetSysColor(COLOR_WINDOW))
    lWhite = IIf(m_BorderType = Inset, GetSysColor(COLOR_WINDOW), GetSysColor(COLOR_BTNSHADOW))

    lneBorder(0).BorderColor = lShadow
    lneBorder(1).BorderColor = lWhite
End Sub

Private Sub UpdateCaption()
    UserControl.Cls

    cGradient.Color1 = RGB(255, 255, 255)  'RGB(212, 213, 216)
    cGradient.Color2 = RGB(119, 145, 173) 'RGB(10, 36, 106)
    
    Call cGradient.Draw(UserControl.hWnd, UserControl.hDC)
    
    UserControl.FontBold = True
    Call TextOut(UserControl.hDC, 10, 10, m_Caption, Len(m_Caption))
    UserControl.FontBold = False

    Call TextOut(UserControl.hDC, 20, 25, m_Description, Len(m_Description))

    UserControl.Refresh
End Sub
