VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Rem colors
Rem color background colors
Private m_color_bg_pushed As Long
Private m_color_bg_default As Long
Private m_color_bg_flagged As Long
Private m_color_bg_mine As Long
Rem color value font colors
Private m_color_value_mine As Long
Private m_color_value_flagged As Long
Private m_color_value_1 As Long
Private m_color_value_2 As Long
Private m_color_value_3 As Long
Private m_color_value_4 As Long
Private m_color_value_5 As Long
Private m_color_value_6 As Long
Private m_color_value_7 As Long
Private m_color_value_8 As Long
Rem table_border line colors
Private m_color_table_bg As Long

Rem border of cells
Private m_cell_border_default_weight As XlBorderWeight
Private m_cell_border_default_style As XlLineStyle
Private m_cell_border_pushed_weight As XlBorderWeight
Private m_cell_border_pushed_style As XlLineStyle

' ----------------------------------------------------------------
' Procedure Name: Class_Initialize
' Purpose: Set up style variables
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Private Sub Class_Initialize()

    m_color_bg_pushed = CLng(RGB(200, 200, 200))
    m_color_bg_default = CLng(RGB(242, 242, 242))
    m_color_bg_flagged = CLng(RGB(255, 165, 0))
    m_color_bg_mine = CLng(RGB(254, 0, 0))
    
    m_color_value_mine = CLng(RGB(0, 0, 0))
    m_color_value_flagged = CLng(RGB(0, 0, 0))
    m_color_value_1 = CLng(RGB(0, 0, 255))
    m_color_value_2 = CLng(RGB(0, 128, 0))
    m_color_value_3 = CLng(RGB(255, 0, 0))
    m_color_value_4 = CLng(RGB(0, 0, 139))
    m_color_value_5 = CLng(RGB(165, 42, 42))
    m_color_value_6 = CLng(RGB(0, 255, 255))
    m_color_value_7 = CLng(RGB(0, 0, 0))
    m_color_value_8 = CLng(RGB(128, 128, 128))
    
    m_color_table_bg = CLng(RGB(0, 128, 0))
    
    m_cell_border_default_weight = xlThick
    m_cell_border_default_style = xlDouble
    m_cell_border_pushed_weight = xlThin
    m_cell_border_pushed_style = xlContinuous

End Sub

' ----------------------------------------------------------------
' Procedure Name: getTextColorByValue (Long)
' Purpose: Return the text color based on the value.
' Procedure Kind: Function
' Procedure Access: Public
' Parameter value (Integer): Value found in selected cell
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Public Function getTextColorByValue(ByVal value As Integer) As Long

    Select Case value
        Case -1: getTextColorByValue = m_color_value_mine
        Case 0: getTextColorByValue = 0
        Case 1: getTextColorByValue = m_color_value_1
        Case 2: getTextColorByValue = m_color_value_2
        Case 3: getTextColorByValue = m_color_value_3
        Case 4: getTextColorByValue = m_color_value_4
        Case 5: getTextColorByValue = m_color_value_5
        Case 6: getTextColorByValue = m_color_value_6
        Case 7: getTextColorByValue = m_color_value_7
        Case 8: getTextColorByValue = m_color_value_8
    End Select

End Function

' ----------------------------------------------------------------
' Procedure Name: getBgColorByValue (Long)
' Purpose: Return the background color based on the value.
' Procedure Kind: Function
' Procedure Access: Public
' Parameter value (Integer): Value found in selected cell
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Public Function getBgColorByValue(ByVal value As Integer) As Long

    getBgColorByValue = IIf(value = -1, m_color_bg_mine, m_color_bg_pushed)

End Function

' ----------------------------------------------------------------
' Procedure Name: Property
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Public Property Get colorValueFlagged() As Long: colorValueFlagged = m_color_value_flagged: End Property
Public Property Get colorBgPushed() As Long: colorBgPushed = m_color_bg_pushed: End Property
Public Property Get colorBgDefault() As Long: colorBgDefault = m_color_bg_default: End Property
Public Property Get colorBgFlagged() As Long: colorBgFlagged = m_color_bg_flagged: End Property
Public Property Get colorBgMine() As Long: colorBgMine = m_color_bg_mine: End Property
Public Property Get colorTableBackground() As Long: colorTableBackground = m_color_table_bg: End Property
Public Property Get borderDefaultWeight() As XlBorderWeight: borderDefaultWeight = m_cell_border_default_weight: End Property
Public Property Get borderDefaultSytle() As XlLineStyle: borderDefaultSytle = m_cell_border_default_style: End Property
Public Property Get borderPushedWeight() As XlBorderWeight: borderPushedWeight = m_cell_border_pushed_weight: End Property
Public Property Get borderPushedStyle() As XlLineStyle: borderPushedStyle = m_cell_border_pushed_style: End Property
