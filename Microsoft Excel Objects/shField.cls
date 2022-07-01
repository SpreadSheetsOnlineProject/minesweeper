VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub btnReset_Click()
    Call mdlMinesweeper.resetGame
End Sub

Private Sub btnStart_Click()
    Call mdlMinesweeper.startGame
End Sub

Private Sub cb16x16_Click()
    Call mdlSheetFunctions.radioSelect("16x16")
End Sub

Private Sub cb30x16_Click()
    Call mdlSheetFunctions.radioSelect("30x16")
End Sub

Private Sub cb9x9_Click()
    Call mdlSheetFunctions.radioSelect("9x9")
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Call mdlMinesweeper.doubleClicked(Target, Cancel)
End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    Call mdlMinesweeper.rightClicked(Target, Cancel)
End Sub
