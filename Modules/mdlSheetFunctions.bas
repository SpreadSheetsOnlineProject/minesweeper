Attribute VB_Name = "mdlSheetFunctions"
Option Explicit

' ----------------------------------------------------------------
' Procedure Name: setRangeStyle
' Purpose: This procedure sets the color and borders of the target range.
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter Target (Range): The ragne which will be set.
' Parameter cell_value (Integer): The value of the target cell, -1: mine, 0: empty, 1-8.
' Parameter cell_value_color (Long): The text color of the cell value.
' Parameter bg_color (Long): The background color of the target cell.
' Parameter border_weight (XlBorderWeight): Part of the border style.
' Parameter border_style (XlLineStyle): Part of the border style.
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Public Sub setRangeStyle(ByRef Target As Range, _
                        cell_value As Integer, _
                        cell_value_color As Long, _
                        bg_color As Long, _
                        border_weight As XlBorderWeight, _
                        border_style As XlLineStyle)

    Rem Left border
    With Target.Borders(xlEdgeLeft)
        .LineStyle = border_style
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = border_weight
    End With
    Rem Top border
    With Target.Borders(xlEdgeTop)
        .LineStyle = border_style
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = border_weight
    End With
    Rem Bottom border
    With Target.Borders(xlEdgeBottom)
        .LineStyle = border_style
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = border_weight
    End With
    Rem Right border
    With Target.Borders(xlEdgeRight)
        .LineStyle = border_style
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = border_weight
    End With
    Rem Inside borders
    With Target.Borders(xlInsideHorizontal)
        .LineStyle = border_style
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = border_weight
    End With
    With Target.Borders(xlInsideVertical)
        .LineStyle = border_style
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = border_weight
    End With
    
    Rem Bacground color
    With Target.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = bg_color
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    Rem set cell value
    Select Case cell_value
        Case -1: Target.Value2 = "*"
        Case 0: Target.Value2 = ""
        Case Else: Target.Value2 = cell_value
    End Select
    
    Rem Alignment
    With Target
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Rem Cell value color
    With Target.Font
        .Color = cell_value_color
        .Bold = True
        .TintAndShade = 0
    End With

End Sub

' ----------------------------------------------------------------
' Procedure Name: showMines
' Purpose: This procedure shows the cell of mines and define style of them.
' Procedure Kind: Function
' Procedure Access: Public
' Parameter win (Boolean): Indicates if the user won the game or failed.
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Public Function showMines(ByVal win As Boolean)

    Rem Collection to store the cells with mine
    Dim coll_mines As Collection: Set coll_mines = mdlGlobals.m_field.getMines()

    Dim mine_range As Range: Set mine_range = coll_mines(1)
    
    Rem Union ranges of the coll_mines into one range.
    Dim c_range As Range
    For Each c_range In coll_mines
        Set mine_range = Union(mine_range, c_range)
    Next c_range
    
    Rem Set up style variables
    Dim cell_value_color As Long: cell_value_color = mdlGlobals.m_style.getTextColorByValue(-1)
    Dim border_style As XlLineStyle: border_style = mdlGlobals.m_style.borderPushedStyle
    Dim border_weight As XlBorderWeight: border_weight = mdlGlobals.m_style.borderPushedWeight
    Rem cell_bg_color based on the value.
    Dim cell_bg_color As Long: cell_bg_color = IIf(win, mdlGlobals.m_style.colorBgFlagged, mdlGlobals.m_style.colorBgMine)
    
    Call setRangeStyle(Target:=mine_range, _
                       cell_value:=-1, _
                       cell_value_color:=cell_value_color, _
                       bg_color:=cell_bg_color, _
                       border_weight:=border_weight, _
                       border_style:=border_style)

End Function

' ----------------------------------------------------------------
' Procedure Name: radioSelect
' Purpose: When the user is changing the size of the board, then this procedure modify the number of mines.
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter win (Boolean): Indicates if the user won the game or failed.
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Public Sub radioSelect(ByVal size As String)

    Select Case size
        Case "9x9": shField.lbMineDinamic.Caption = 10
        Case "16x16": shField.lbMineDinamic.Caption = 40
        Case "30x16": shField.lbMineDinamic.Caption = 100
    End Select

End Sub

' ----------------------------------------------------------------
' Procedure Name: formatTableBackground
' Purpose: When the games starts, this procedure display the board of the game.
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter range_ (Range): Represents the table of the game.
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Public Sub formatTableBackground(ByRef range_ As Range)
    
    Rem The background of the edge of the board.
    Dim color_table_bg As Long: color_table_bg = mdlGlobals.m_style.colorTableBackground
    
    Dim Target As Range
    Rem Move the range by 1 cell column left and 1 row up
    Set Target = range_.Offset(-1, -1)
    Rem Resize the range to include the bottom right corner
    Set Target = Target.Resize(Target.Rows.Count + 2, Target.Columns.Count + 2)

    Rem Set width and height of the columns and rows
    Target.ColumnWidth = 2.43
    Target.RowHeight = 18
    
    With Target.Rows(1).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = color_table_bg
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Target.Rows(1).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Target.Rows(Target.Rows.Count).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = color_table_bg
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Target.Rows(Target.Rows.Count).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Target.Columns(1).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = color_table_bg
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Target.Columns(1).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Target.Columns(Target.Columns.Count).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = color_table_bg
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Target.Columns(Target.Columns.Count).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With

End Sub

' ----------------------------------------------------------------
' Procedure Name: setSheetControls
' Purpose: Procedure to change the availability of start and reset buttons.
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter started (Boolean): Indicates if the games started or not.
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Public Sub setSheetControls(ByVal started As Boolean)

    Rem Button availabilities
    shField.btnStart.Enabled = Not started
    shField.btnReset.Enabled = started
    
    Rem Show remaining mines if the game started, or 0
    If started Then
        shField.tbRemMines.Text = mdlGlobals.m_field.remainingMines
    Else
        shField.tbRemMines.Text = 0
        
    End If

End Sub

' ----------------------------------------------------------------
' Procedure Name: formatGameToReset
' Purpose: Format the whole sheet of the game.
' Procedure Kind: Function
' Procedure Access: Public
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Public Function formatGameToReset()
    
    Rem From 1:1 to 50:100 will be enough
    Dim range_ As Range
    Set range_ = shField.Range(shField.Cells(1, 1), shField.Cells(50, 100))
    
    Rem Clear the cells
    range_.ClearContents
    range_.ClearFormats
    
    Rem set up widht and height of each column and row of the rnage
    range_.ColumnWidth = 2.43
    range_.RowHeight = 18

End Function
