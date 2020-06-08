Option Explicit
'==============================================================================
'   > FSW - Working Suit Equipment v1.00<
'
'   Author          : Frixell Script Works
'   Version         : 1.00
'   Release Date    : 09/06/2020
'   Last Update     : 09/06/2020
'   Language        : Visual Basic for Application (VBA)
'   Platform        : Excel (2013) 32-bit
'------------------------------------------------------------------------------
'   Level Script    : Easy, Intermediate
'   Compatibility   : Medium
'   Requires        : Sheet Configuration
'==============================================================================
' > Changelog
'
' 09 June 2020
'   - Script Finished and Released.
'
'==============================================================================
'
' > Introduction
'
' This script provide equipping working suit to character, make him/her better during work.
' Also minimizes risk to be hurt during work.
'
' Command button uses macro to call this script to make a character wear the battle suit
' and switches if there's already equipped equipment.
'
'
'==============================================================================
'
' > Usage
'
' *Place this script in the code editor within the sheet or module.
'
' Design the EQUIPMENT sheet and INVENTORY SHEET like in the sample sheet.
' You can download it here (https://github.com/FrixellScriptWorks/Excel-war)
'
' To call this script, you can make a control button (not ActiveX), and asiign it
' with macro "equip_work_eq" (without quote, you will know by looking it)
'
' * = important
'
'==============================================================================
Sub equip_work_eq() 'Init Procedure
'==============================================================================
' CONFIGURATION SECTION
'
' Below is the configuration section where you can edit variables in the script.
' This VBA variables needs to be declared first, so be careful to not changing or
' deleting the declaration line (i.e. Dim VARIABLES As String: <- this is declaration)
'==============================================================================

' This is where the EQUIPMENT SHEET and INVENTORY SHEET present. Change the words in the brackets into the sheet you want.
Dim EQ_WORKSHEET As Worksheet:                   Set EQ_WORKSHEET = ThisWorkbook.Sheets("Equipment")
Dim INV_WORKSHEET As Worksheet:                 Set INV_WORKSHEET = ThisWorkbook.Sheets("Inventory")

' This is the equipment slot takes place. This variables is stored in the EQUIPMENT SHEET.
Dim START_WORK_ROW As String:                      START_WORK_ROW = "4" 'Default = 4
Dim NAME_COLUMN As String:                            NAME_COLUMN = "C" 'Default = C
Dim PRODUCTIVITY_N_COLUMN As String:        PRODUCTIVITY_N_COLUMN = "D" 'Default = D
Dim PRODUCTIVITY_P_COLUMN As String:        PRODUCTIVITY_P_COLUMN = "E" 'Default = E
Dim HURT_COLUMN As String:                            HURT_COLUMN = "F" 'Default = F
Dim RESOURCES_COLUMN As String:                  RESOURCES_COLUMN = "G" 'Default = G
Dim BALANCE_COLUMN As String:                      BALANCE_COLUMN = "H" 'Default = H
Dim POWER_RATE_COLUMN As String:                POWER_RATE_COLUMN = "I" 'Default = I
Dim POWER_MULT_COLUMN As String:                POWER_MULT_COLUMN = "J" 'Default = J

' This is where the equipment stored in inventory. Note that this is the column of the NAME.
' You can change it according to your equipment table location.
Dim WORK_EQ_COL As Integer:                           WORK_EQ_COL = 2 'Default = 2

' This is message when you encounter some minor error when switching equipment.
Dim MSGBOX_EMPTY As String:                          MSGBOX_EMPTY = "This is empty, it's no use."
Dim MSGBOX_ALREADY_EQUIPPED As String:    MSGBOX_ALREADY_EQUIPPED = "There's already equipped item, switching..."
Dim MSGBOX_NOT_WORK_EQ As String:              MSGBOX_NOT_WORK_EQ = "Please select an appropiate work equipment." & vbNewLine & "Select the name of your work equipment."

' This is the name of TABLE object of the inventory. Change the name in the brackets with your table name.
Dim workEqTable As Object:                        Set workEqTable = INV_WORKSHEET.ListObjects("workEqTable")

'==============================================================================
' Rest of the script
'
' It's different from the configuration section. Main script starts here.
' Please do not edit past this point unless you know what are you doing.
'
'==============================================================================

'Preparation
Dim selectedEqRow As Integer: selectedEqRow = ActiveCell.Row
Dim selectedEqColName As Integer: selectedEqColName = ActiveCell.Column
Dim selectedEqColSlot As Integer: selectedEqColSlot = ActiveCell.Offset(0, 1).Column
Dim selectedEqColPdN As Integer: selectedEqColPdN = ActiveCell.Offset(0, 2).Column
Dim selectedEqColPdP As Integer: selectedEqColPdP = ActiveCell.Offset(0, 3).Column
Dim selectedEqColHurt As Integer: selectedEqColHurt = ActiveCell.Offset(0, 4).Column
Dim selectedEqColReso As Integer: selectedEqColReso = ActiveCell.Offset(0, 5).Column
Dim selectedEqColBal As Integer: selectedEqColBal = ActiveCell.Offset(0, 6).Column
Dim selectedEqColPowRate As Integer: selectedEqColPowRate = ActiveCell.Offset(0, 7).Column
Dim selectedEqColPowMult As Integer: selectedEqColPowMult = ActiveCell.Offset(0, 8).Column


'Error Handler
On Error GoTo ErrHandler

'Error Handler 2
If selectedEqColName <> WORK_EQ_COL Then
    MsgBox (MSGBOX_NOT_WORK_EQ)
    Exit Sub
ElseIf INV_WORKSHEET.Cells(selectedEqRow, selectedEqColSlot).Value Then
    MsgBox (MSGBOX_EMPTY)
    Exit Sub
End If

'Slot Check
Dim selectRow As String
Select Case INV_WORKSHEET.Cells(selectedEqRow, selectedEqColSlot).Value
Case "Head"
    selectRow = START_WORK_ROW
Case "Vision"
    selectRow = START_WORK_ROW + 1
Case "Body"
    selectRow = START_WORK_ROW + 2
Case "Pants"
    selectRow = START_WORK_ROW + 3
Case "Boots"
    selectRow = START_WORK_ROW + 4
Case "Charm"
    selectRow = START_WORK_ROW + 5
Case "Offhand"
    selectRow = START_WORK_ROW + 6
Case Else
    MsgBox ("Not a valid slot")
    Exit Sub
End Select

'Check if there's an equipment
If EQ_WORKSHEET.Range(NAME_COLUMN & selectRow).Value <> "" Then
    MsgBox (MSGBOX_ALREADY_EQUIPPED)
    workEqTable.ListRows.Add (selectedEqRow - 4)
    INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Offset(1, 0).Activate
    selectedEqRow = ActiveCell.Row
    INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Offset(-1, 0).Value = EQ_WORKSHEET.Range(NAME_COLUMN & selectRow).Value
    INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Offset(-1, 1).Value = INV_WORKSHEET.Cells(selectedEqRow, selectedEqColSlot).Value
    INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Offset(-1, 2).Value = EQ_WORKSHEET.Range(PRODUCTIVITY_N_COLUMN & selectRow).Value
    INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Offset(-1, 3).Value = EQ_WORKSHEET.Range(PRODUCTIVITY_P_COLUMN & selectRow).Value
    INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Offset(-1, 4).Value = EQ_WORKSHEET.Range(HURT_COLUMN & selectRow).Value
    INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Offset(-1, 5).Value = EQ_WORKSHEET.Range(RESOURCES_COLUMN & selectRow).Value
    INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Offset(-1, 6).Value = EQ_WORKSHEET.Range(BALANCE_COLUMN & selectRow).Value
    INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Offset(-1, 7).Value = EQ_WORKSHEET.Range(POWER_RATE_COLUMN & selectRow).Value
    INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Offset(-1, 8).Value = EQ_WORKSHEET.Range(POWER_MULT_COLUMN & selectRow).Value
Else:
End If

'Equip this
EQ_WORKSHEET.Range(NAME_COLUMN & selectRow).Value = INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Value
EQ_WORKSHEET.Range(PRODUCTIVITY_N_COLUMN & selectRow).Value = INV_WORKSHEET.Cells(selectedEqRow, selectedEqColPdN).Value
EQ_WORKSHEET.Range(PRODUCTIVITY_P_COLUMN & selectRow).Value = INV_WORKSHEET.Cells(selectedEqRow, selectedEqColPdP).Value
EQ_WORKSHEET.Range(HURT_COLUMN & selectRow).Value = INV_WORKSHEET.Cells(selectedEqRow, selectedEqColHurt).Value
EQ_WORKSHEET.Range(RESOURCES_COLUMN & selectRow).Value = INV_WORKSHEET.Cells(selectedEqRow, selectedEqColReso).Value
EQ_WORKSHEET.Range(BALANCE_COLUMN & selectRow).Value = INV_WORKSHEET.Cells(selectedEqRow, selectedEqColBal).Value
EQ_WORKSHEET.Range(POWER_RATE_COLUMN & selectRow).Value = INV_WORKSHEET.Cells(selectedEqRow, selectedEqColPowRate).Value
EQ_WORKSHEET.Range(POWER_MULT_COLUMN & selectRow).Value = INV_WORKSHEET.Cells(selectedEqRow, selectedEqColPowMult).Value

'Deleting Eq
workEqTable.ListRows(selectedEqRow - 4).Delete

'Moving view
EQ_WORKSHEET.Activate
EQ_WORKSHEET.Range(NAME_COLUMN & selectRow).Select
Exit Sub

'ErrHandler label
ErrHandler:
    MsgBox ("Error detected.")
    
End Sub
'===============================================================
'
' > End of the script
'
'===============================================================

