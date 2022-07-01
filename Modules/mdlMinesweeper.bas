Attribute VB_Name = "mdlMinesweeper"
Option Explicit

' ----------------------------------------------------------------
' Procedure Name: resetGame
' Purpose: This procedure resets the game
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Public Sub resetGame()

    mdlGlobals.game_is_running = False
    mdlGlobals.game_end_time = Now()

    Rem Reset controls in field sheet
    Call mdlSheetFunctions.setSheetControls(False)
    Rem Reset board in field sheet
    Call mdlSheetFunctions.formatGameToReset
    
    Rem Set global variables to nothing
    Set mdlGlobals.m_field = Nothing
    Set mdlGlobals.m_style = Nothing

End Sub

' ----------------------------------------------------------------
' Procedure Name: startGame
' Purpose: This procedure prepaires and starts the game
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Public Sub startGame()

    Rem Set field and color classes
    Set mdlGlobals.m_field = New clsField
    Set mdlGlobals.m_style = New clsStyle
    
    mdlGlobals.game_is_running = True

    Rem Reset the board
    Call mdlSheetFunctions.formatGameToReset
    
    Rem Check which checkbox was selected and set the field based on it
    If shField.cb9x9.value Then Call m_field.setField(9, 9, 10)
    If shField.cb16x16.value Then Call m_field.setField(16, 16, 40)
    If shField.cb30x16.value Then Call m_field.setField(16, 30, 100)
    
    Call mdlSheetFunctions.setSheetControls(True)
    
    Dim board_start_row As Integer: board_start_row = mdlGlobals.m_field.boardStartRow
    Dim board_start_col As Integer: board_start_col = mdlGlobals.m_field.boardStartCol
    Dim number_of_rows As Integer: number_of_rows = mdlGlobals.m_field.rowNum
    Dim number_of_cols As Integer: number_of_cols = mdlGlobals.m_field.colNum
    
    Rem Set the playground defined by field range
    Dim field As Range
    Set field = shField.Range(Cells(board_start_row, board_start_col), _
                              Cells(board_start_row, board_start_col).Offset(number_of_rows - 1, number_of_cols - 1))

    Rem Set the background of the board
    Call mdlSheetFunctions.formatTableBackground(field)
    
    Rem Set the cells of the field.
    Call mdlSheetFunctions.setRangeStyle(Target:=field, _
                                         cell_value:=0, _
                                         cell_value_color:=0, _
                                         bg_color:=mdlGlobals.m_style.colorBgDefault, _
                                         border_weight:=mdlGlobals.m_style.borderDefaultWeight, _
                                         border_style:=mdlGlobals.m_style.borderDefaultSytle)
    
    mdlGlobals.game_start_time = Now()
    
End Sub

' ----------------------------------------------------------------
' Procedure Name: gameFinished
' Purpose: This procedure finishes the game.
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter win (Boolean): Indicates if the user win the game or not.
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Public Sub gameFinished(ByVal win As Boolean)

    mdlGlobals.game_is_running = False
    mdlGlobals.game_end_time = Now()
    
    Call mdlSheetFunctions.showMines(win)
    
    If win Then
        MsgBox "Gratulálok! Minden aknát megtaláltál." & vbNewLine & _
            "Játék kezdete: " & mdlGlobals.game_start_time & vbNewLine & _
            "Játék vége: " & mdlGlobals.game_end_time, vbOKOnly, "Gratulálok!"
    Else
        MsgBox "Sajnos aknára léptél! Próbáld meg újra." & vbNewLine & _
            "Játék kezdete: " & mdlGlobals.game_start_time & vbNewLine & _
            "Játék vége: " & mdlGlobals.game_end_time, vbOKOnly, "Sajnos aknára léptél!"
    End If

End Sub

' ----------------------------------------------------------------
' Procedure Name: doubleClicked
' Purpose: This procedure runs when the user double clicked on a cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter Target (Range): The cell or range what the user clicked.
' Parameter Cancel (Boolean): Deny or allow to get into the cell.
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Public Sub doubleClicked(ByRef Target As Range, _
                         ByRef Cancel As Boolean)

On Error GoTo eh

    Rem If the game is not running, do nothing
    If Not mdlGlobals.game_is_running Then Exit Sub
    
    Rem If the user double clicked on multiple field then do nothing
    If Target.Cells.Count > 1 Then
        Cancel = True
        Exit Sub
    End If
    
    Rem If the user clicked a cell outside of the table then do nothing
    If Not mdlGlobals.m_field.isValidField(Target) Then Exit Sub

    Rem If the cell has been flagged then do nothing
    If Target.Interior.Color = mdlGlobals.m_style.colorBgFlagged Then
        Cancel = True
        Exit Sub
    End If
    
    Rem If the cell has been clicked then do nothing
    If Target.Interior.Color = mdlGlobals.m_style.colorBgPushed Then
        Cancel = True
        Exit Sub
    End If
    
    Cancel = True
    
    Rem Get the value of the selected cell
    Dim cell_value As Integer: cell_value = m_field.getCellValue(Target)
    
    Rem If the selected cell is mine then end the game
    If cell_value = -1 Then
        Call gameFinished(False)
        Exit Sub
    End If
    
    Rem In case of the selected cell is empty this variable counts how many cells are opened.
    Rem This is 1 by default.
    Dim touched_fields As Integer: touched_fields = 1
    
    Rem If the cell is an empty field
    If cell_value = 0 Then
        
        Rem This collection is going to store the empty cells around the selected one.
        Dim coll_empty As New Collection
        Call mdlGlobals.m_field.getEmptyCells(Target.row, Target.Column, coll_empty)
        
        Rem Due to the selected cell is empty, this value is set by the count of surrounding empty cells.
        touched_fields = coll_empty.Count
        
        Rem Loop through the empty cell collection
        Dim c_range As Range
        For Each c_range In coll_empty

            Rem If the current cell has already pushed and it is not empty.
            If Not IsEmpty(c_range) Then
                touched_fields = touched_fields - 1
            End If
            
            Rem Set the current cell of the field as pushed.
            cell_value = m_field.getCellValue(c_range)
            Call setTouchedCells(c_range, cell_value)
        
        Next c_range
        
    Else
        Call setTouchedCells(Target, cell_value)
    End If
    
    Rem Decrease the number of untouched cells to decide if every non-mine cell has been found
    mdlGlobals.m_field.untouchedFields = mdlGlobals.m_field.untouchedFields - touched_fields
    
    Rem If ervery non-mine cell has been opened
    If mdlGlobals.m_field.untouchedFields <= 0 Then
        Call gameFinished(True)
        Exit Sub
    End If
    
done:
    Exit Sub
eh:
    Rem If error occure, the game must be reset.
    MsgBox "Sajnost váratlan hiba lépett fel a játék közben!" & vbNewLine & _
        "A játékot újra kell indítanom." & vbNewLine & _
        "További jó szórakozást!", vbCritical, "Váratlan hiba!"
    Call resetGame
    
End Sub

' ----------------------------------------------------------------
' Procedure Name: rightClicked
' Purpose: This procedure runs when the user right clicked on a cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter Target (Range): The cell or range what the user clicked.
' Parameter Cancel (Boolean): Deny or allow to get into the cell.
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Public Sub rightClicked(ByRef Target As Range, _
                        ByRef Cancel As Boolean)

On Error GoTo eh

    Rem If the game is not running then do nothing
    If Not mdlGlobals.game_is_running Then Exit Sub

    Rem If multiple cells were clicked, then do nothing
    If Target.Cells.Count > 1 Then Exit Sub
    
    Rem Check if the clicked cell is inside of the gameboard.
    If Not mdlGlobals.m_field.isValidField(Target) Then Exit Sub
    
    Cancel = True
    
    Rem If the color of the clicked cell is flagged then remove flag
    If Target.Interior.Color = mdlGlobals.m_style.colorBgFlagged Then
    
        Call mdlSheetFunctions.setRangeStyle(Target:=Target, _
                                             cell_value:=0, _
                                             cell_value_color:=0, _
                                             bg_color:=mdlGlobals.m_style.colorBgDefault, _
                                             border_weight:=mdlGlobals.m_style.borderDefaultWeight, _
                                             border_style:=mdlGlobals.m_style.borderDefaultSytle)
                                             
        Rem Increase the number of remaining mines.
        mdlGlobals.m_field.remainingMines = mdlGlobals.m_field.remainingMines + 1
        shField.tbRemMines.value = mdlGlobals.m_field.remainingMines
                
        Exit Sub
        
    End If
    
    Rem If the color of the clicked cell is the default background, then mark as flag
    If Target.Interior.Color = mdlGlobals.m_style.colorBgDefault Then
    
        If CInt(shField.tbRemMines.value) = 0 Then Exit Sub

        Call mdlSheetFunctions.setRangeStyle(Target:=Target, _
                                             cell_value:=0, _
                                             cell_value_color:=0, _
                                             bg_color:=mdlGlobals.m_style.colorBgFlagged, _
                                             border_weight:=mdlGlobals.m_style.borderDefaultWeight, _
                                             border_style:=mdlGlobals.m_style.borderDefaultSytle)
        
        Rem Decrease the number of remaining mines.
        mdlGlobals.m_field.remainingMines = mdlGlobals.m_field.remainingMines - 1
        shField.tbRemMines.value = mdlGlobals.m_field.remainingMines
        
        Exit Sub
    End If
    
done:
    Exit Sub
eh:
    Rem In case of error, the game must be reset.
    MsgBox "Sajnost váratlan hiba lépett fel a játék közben!" & vbNewLine & _
        "A játékot újra kell indítanom." & vbNewLine & _
        "További jó szórakozást!", vbCritical, "Váratlan hiba!"
    Call resetGame
    
End Sub

' ----------------------------------------------------------------
' Procedure Name: setTouchedCells
' Purpose: Set the style of the target range
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter Target (Range): The cell or range where the style needs to be set.
' Parameter cell_value (Integer): The value in the cell of the field -1 = mine, 0 = empty
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Public Sub setTouchedCells(ByRef Target As Range, _
                           ByVal cell_value As Integer)

    Rem Style variables
    Dim new_color_value As Long: new_color_value = mdlGlobals.m_style.getTextColorByValue(cell_value)
    Dim new_color_cell As Long: new_color_cell = mdlGlobals.m_style.getBgColorByValue(cell_value)
    Dim new_border_weight As XlBorderWeight: new_border_weight = mdlGlobals.m_style.borderPushedWeight
    Dim new_border_style As XlLineStyle: new_border_style = mdlGlobals.m_style.borderPushedStyle

    Call mdlSheetFunctions.setRangeStyle(Target:=Target, _
                                         cell_value:=cell_value, _
                                         cell_value_color:=new_color_value, _
                                         bg_color:=new_color_cell, _
                                         border_weight:=new_border_weight, _
                                         border_style:=new_border_style)

End Sub
