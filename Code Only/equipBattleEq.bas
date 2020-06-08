Option Explicit
'==============================================================================
'   > FSW - Battle Suit Equipment v1.02<
'
'   Author          : Frixell Script Works
'   Version         : 1.02
'   Release Date    : 08/06/2020
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
'   - Fixed Bug: Error occurs when equipping empty equipment
' 08 June 2020
'   - Script Finished and Released.
'
'==============================================================================
'
' > Introduction
'
' This script provide equipping battle suit to character
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
' with macro "equip_battle_eq" (without quote, you will know by looking it)
'
' * = important
'
'==============================================================================
Sub equip_battle_eq() 'Init Procedure
'==============================================================================
' CONFIGURATION SECTION
'
' Below is the configuration section where you can edit variables in the script.
' This VBA variables needs to be declared first, so be careful to not changing or
' deleting the declaration line (i.e. Dim VARIABLES As String: <- this is declaration)
'==============================================================================

' This is where the EQUIPMENT SHEET and INVENTORY SHEET present. Change the words in the brackets into the sheet you want.
Dim EQ_WORKSHEET As Worksheet:                  Set EQ_WORKSHEET = ThisWorkbook.Sheets("Equipment")
Dim INV_WORKSHEET As Worksheet:                Set INV_WORKSHEET = ThisWorkbook.Sheets("Inventory")

' This is the equipment slot takes place. This variables is stored in the EQUIPMENT SHEET.
Dim START_BATTLE_ROW As String:                 START_BATTLE_ROW = "17"   'Default = 17
Dim NAME_COLUMN As String:                           NAME_COLUMN = "C"    'Default = C
Dim DAMAGE_COLUMN As String:                       DAMAGE_COLUMN = "D"    'Default = D
Dim HP_COLUMN As String:                               HP_COLUMN = "E"    'Default = E
Dim ARMOR_COLUMN As String:                         ARMOR_COLUMN = "F"    'Default = F
Dim PENETRATION_COLUMN As String:             PENETRATION_COLUMN = "G"    'Default = G
Dim HIT_COLUMN As String:                             HIT_COLUMN = "H"    'Default = H
Dim EVASION_COLUMN As String:                     EVASION_COLUMN = "I"    'Default = I
Dim CRIT_RATE_COLUMN As String:                 CRIT_RATE_COLUMN = "J"    'Default = J
Dim CRIT_EVASION_COLUMN As String:           CRIT_EVASION_COLUMN = "K"    'Default = K
Dim CRIT_MULTIPLIER_COLUMN As String:     CRIT_MULTIPLIER_COLUMN = "L"    'Default = L

' This is where the equipment stored in inventory. Note that this is the column of the NAME.
' You can change it according to your equipment table location.
Dim BATTLE_EQ_COL As Integer:                      BATTLE_EQ_COL = 13     'Default = 13

' This is message when you encounter some minor error when switching equipment.
Dim MSGBOX_EMPTY As String:                         MSGBOX_EMPTY = "This is empty, it's no use."
Dim MSGBOX_ALREADY_EQUIPPED As String:   MSGBOX_ALREADY_EQUIPPED = "There's already equipped item, switching..."
Dim MSGBOX_NOT_BATTLE_EQ As String:         MSGBOX_NOT_BATTLE_EQ = "Please select an appropiate battle equipment." & vbNewLine & "Select the name of your battle equipment."

' This is the name of TABLE object of the inventory. Change the name in the brackets with your table name.
Dim battleEqTable As Object:                   Set battleEqTable = INV_WORKSHEET.ListObjects("battleEqTable")

'==============================================================================
' Rest of the script
'
' It's different from the configuration section. Main script starts here.
' Please do not edit past this point unless you know what are you doing.
'
'==============================================================================

'Preparation
Dim selectedEqRow As Integer:                      selectedEqRow = ActiveCell.Row
Dim selectedEqColName As Integer:              selectedEqColName = ActiveCell.Column
Dim selectedEqColSlot As Integer:              selectedEqColSlot = ActiveCell.Offset(0, 1).Column
Dim selectedEqColDmg As Integer:                selectedEqColDmg = ActiveCell.Offset(0, 2).Column
Dim selectedEqColHp As Integer:                  selectedEqColHp = ActiveCell.Offset(0, 3).Column
Dim selectedEqColArmor As Integer:            selectedEqColArmor = ActiveCell.Offset(0, 4).Column
Dim selectedEqColPenet As Integer:            selectedEqColPenet = ActiveCell.Offset(0, 5).Column
Dim selectedEqColHit As Integer:                selectedEqColHit = ActiveCell.Offset(0, 6).Column
Dim selectedEqColEva As Integer:                selectedEqColEva = ActiveCell.Offset(0, 7).Column
Dim selectedEqColCritRate As Integer:      selectedEqColCritRate = ActiveCell.Offset(0, 8).Column
Dim selectedEqColCritEva As Integer:        selectedEqColCritEva = ActiveCell.Offset(0, 9).Column
Dim selectedEqColCritMult As Integer:      selectedEqColCritMult = ActiveCell.Offset(0, 10).Column

'Error Handler
On Error GoTo ErrHandler

'Error Handler 2
If selectedEqColName <> BATTLE_EQ_COL Then
    MsgBox (MSGBOX_NOT_BATTLE_EQ)
    Exit Sub
ElseIf INV_WORKSHEET.Cells(selectedEqRow, selectedEqColSlot).Value = "" Then
    MsgBox (MSGBOX_EMPTY)
    Exit Sub
End If

'Slot Check
Dim selectRow As String
Select Case INV_WORKSHEET.Cells(selectedEqRow, selectedEqColSlot).Value
Case "Primary"
    selectRow = START_BATTLE_ROW
Case "Secondary"
    selectRow = START_BATTLE_ROW + 1
Case "Helmet"
    selectRow = START_BATTLE_ROW + 2
Case "Armor"
    selectRow = START_BATTLE_ROW + 3
Case "Boots"
    selectRow = START_BATTLE_ROW + 4
Case "Sights"
    selectRow = START_BATTLE_ROW + 5
Case Else
    MsgBox ("Not a valid slot")
    Exit Sub
End Select

'Check if there's an equipment
If EQ_WORKSHEET.Range(NAME_COLUMN & selectRow).Value <> "" Then
    MsgBox (MSGBOX_ALREADY_EQUIPPED)
    battleEqTable.ListRows.Add (selectedEqRow - 4)
    INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Offset(1, 0).Activate
    selectedEqRow = ActiveCell.Row
    INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Offset(-1, 0).Value = EQ_WORKSHEET.Range(NAME_COLUMN & selectRow).Value
    INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Offset(-1, 1).Value = INV_WORKSHEET.Cells(selectedEqRow, selectedEqColSlot).Value
    INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Offset(-1, 2).Value = EQ_WORKSHEET.Range(DAMAGE_COLUMN & selectRow).Value
    INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Offset(-1, 3).Value = EQ_WORKSHEET.Range(HP_COLUMN & selectRow).Value
    INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Offset(-1, 4).Value = EQ_WORKSHEET.Range(ARMOR_COLUMN & selectRow).Value
    INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Offset(-1, 5).Value = EQ_WORKSHEET.Range(PENETRATION_COLUMN & selectRow).Value
    INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Offset(-1, 6).Value = EQ_WORKSHEET.Range(HIT_COLUMN & selectRow).Value
    INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Offset(-1, 7).Value = EQ_WORKSHEET.Range(EVASION_COLUMN & selectRow).Value
    INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Offset(-1, 8).Value = EQ_WORKSHEET.Range(CRIT_RATE_COLUMN & selectRow).Value
    INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Offset(-1, 9).Value = EQ_WORKSHEET.Range(CRIT_EVASION_COLUMN & selectRow).Value
    INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Offset(-1, 10).Value = EQ_WORKSHEET.Range(CRIT_MULTIPLIER_COLUMN & selectRow).Value
Else:
End If

'Equip this
EQ_WORKSHEET.Range(NAME_COLUMN & selectRow).Value = INV_WORKSHEET.Cells(selectedEqRow, selectedEqColName).Value
EQ_WORKSHEET.Range(DAMAGE_COLUMN & selectRow).Value = INV_WORKSHEET.Cells(selectedEqRow, selectedEqColDmg).Value
EQ_WORKSHEET.Range(HP_COLUMN & selectRow).Value = INV_WORKSHEET.Cells(selectedEqRow, selectedEqColHp).Value
EQ_WORKSHEET.Range(ARMOR_COLUMN & selectRow).Value = INV_WORKSHEET.Cells(selectedEqRow, selectedEqColArmor).Value
EQ_WORKSHEET.Range(PENETRATION_COLUMN & selectRow).Value = INV_WORKSHEET.Cells(selectedEqRow, selectedEqColPenet).Value
EQ_WORKSHEET.Range(HIT_COLUMN & selectRow).Value = INV_WORKSHEET.Cells(selectedEqRow, selectedEqColHit).Value
EQ_WORKSHEET.Range(EVASION_COLUMN & selectRow).Value = INV_WORKSHEET.Cells(selectedEqRow, selectedEqColEva).Value
EQ_WORKSHEET.Range(CRIT_RATE_COLUMN & selectRow).Value = INV_WORKSHEET.Cells(selectedEqRow, selectedEqColCritRate).Value
EQ_WORKSHEET.Range(CRIT_EVASION_COLUMN & selectRow).Value = INV_WORKSHEET.Cells(selectedEqRow, selectedEqColCritEva).Value
EQ_WORKSHEET.Range(CRIT_MULTIPLIER_COLUMN & selectRow).Value = INV_WORKSHEET.Cells(selectedEqRow, selectedEqColCritMult).Value

'Deleting Eq
battleEqTable.ListRows(selectedEqRow - 4).Delete

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

