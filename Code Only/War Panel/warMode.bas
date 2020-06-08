Option Explicit
'==============================================================================
'   > FSW - Simple Duel Battle with Result Form v1.04<
'
'   Author          : Frixell Script Works
'   Version         : 1.04
'   Release Date    : 03/06/2020
'   Last Update     : 08/06/2020
'   Language        : Visual Basic for Application (VBA)
'   Platform        : Excel (2013) 32-bit
'------------------------------------------------------------------------------
'   Level Script    : Easy
'   Compatibility   : High
'   Requires        : battleResultForm & Sheet Configuration
'==============================================================================
' > Changelog
'
' 08 June 2020
'   - Added Feature    : Damage variance defined by user in config section.
'   - Minor Change     : Formatting numbers.
'   - Minor Change     : Font style in result form.
' 07 June 2020
'   - Added Feature    : Configuration section let you configure a thing.
'   - Minor Change     : Define the sheet name, no more ActiveSheet
' 05 June 2020
'   - Added Feature    : Battle shout when hit or miss.
'   - Fixed bug        : Damage bar overflowed.
' 04 June 2020
'   - Added Feature    : Battle Result UserForm with Damage Level bar.
'   - Fixed Bug        : Score reset to 0 after hit more than one time.
'   - Fixed Bug        : Random Number doesn't work correctly.
' 03 June 2020
'   - Script Finished and Released.
'
'==============================================================================
'
' > Introduction
'
' This script can make a duel simulator in one Worksheet.
'
' One button can make player 1 and 2 attacking each other with some adjustment
' on the excel cells' value. Like other simulator, there are some parameter that
' you 'must' preserve before running this script, such as:
'
' - Damage
' - Health Point
' - Armor
' - Evasion, etc. (you can see this on the excel)
'
'==============================================================================
'
' > Usage
'
' *Place this script in the code editor within the sheet or module.
'
' Design the battler stat parameter like in the example sheet. You can make anything you
' like as long as the total stats is on the ROW 11 (player 1) and ROW 25 (player 2), and
' COLUMN starting from D to L
'
' Score is on COLUMN O
'
' To call this script, you can make a control button (not ActiveX), and asiign it
' with macro "fight_enemy" (without quote, you will know by looking it)
'
' * = important
'
'==============================================================================
Sub fight_enemy() 'Init procedure
'==============================================================================
' CONFIGURATION SECTION
'
' Below is the configuration section where you can edit variables in the script.
' This VBA variables needs to be declared first, so be careful to not changing or
' deleting the declaration line (i.e. Dim VARIABLES As String: <- this is declaration)
'==============================================================================
' This is where the sheet and workbook is present. You can change to any sheet you want.
Dim THIS_WORKSHEET As Worksheet:  Set THIS_WORKSHEET = ThisWorkbook.Sheets("Fight")

' This is type of the army. Change the value if you want the type in different cell.
' Type of units contain: Infantry, Vehicle, Tank, Aircraft, Helicopter, Artillery.
' Each of them has advantages and disadvantages againts the others.
' Right now, you cannot add or remove those types. If you got the wrong value,
' then you didn't get the army type bonuses.
Dim YOUR_TYPE_CELL As String:       YOUR_TYPE_CELL = "C2"   'Default = C2
Dim ENEMY_TYPE_CELL As String:     ENEMY_TYPE_CELL = "C16"  'Default = C16

' This is where the TOTAL stats is stored. You can change the row and column of the
' stat cell as you like. Remember, they are in the same ROW.
Dim YOUR_STAT_ROW As String:         YOUR_STAT_ROW = "11"   'Default = 11
Dim ENEMY_STAT_ROW As String:       ENEMY_STAT_ROW = "25"   'Default = 25
Dim DAMAGE As String:                       DAMAGE = "D"    'Default = D
Dim HEALTH_POINT As String:           HEALTH_POINT = "E"    'Default = E
Dim ARMOR As String:                         ARMOR = "F"    'Default = F
Dim PENETRATION As String:             PENETRATION = "G"    'Default = G
Dim HIT_RATE As String:                   HIT_RATE = "H"    'Default = H
Dim EVASION As String:                     EVASION = "I"    'Default = I
Dim CRIT_RATE As String:                 CRIT_RATE = "J"    'Default = J
Dim CRIT_EVASION As String:           CRIT_EVASION = "K"    'Default = K
Dim CRIT_MULTIPLIER As String:     CRIT_MULTIPLIER = "L"    'Default = L

' This cell is represent where the CURRENT hit points and score is stored. You can
' change it as well.
Dim YOUR_HP_CELL As String:            YOUR_HP_CELL = "O3"  'Default = O3
Dim ENEMY_HP_CELL As String:          ENEMY_HP_CELL = "O4"  'Default = O4
Dim YOUR_SCORE_CELL As String:      YOUR_SCORE_CELL = "O9"  'Default = O9
Dim ENEMY_SCORE_CELL As String:    ENEMY_SCORE_CELL = "O10" 'Default = O10

' This is a shout text that appear in the result box. Variable begins with YOUR means
' the shout performed by your action. If you miss, then you shout miss.
' Change it to a text do you like.
Dim YOUR_SHOUT_CRIT As String: YOUR_SHOUT_CRIT = "Critical hit!"
Dim YOUR_SHOUT_HIT As String: YOUR_SHOUT_HIT = "Direct hit!"
Dim YOUR_SHOUT_MISS As String: YOUR_SHOUT_MISS = "It's a miss!"
Dim YOUR_SHOUT_REFLECT1 As String: YOUR_SHOUT_REFLECT1 = "It's a reflection!"
Dim YOUR_SHOUT_REFLECT2 As String: YOUR_SHOUT_REFLECT2 = "We didn't even scratch them!"
Dim YOUR_SHOUT_REFLECT3 As String: YOUR_SHOUT_REFLECT3 = "Ricochet!"

' This is a shout text made BY YOU after THE ENEMY taken action. You can see a relief
' text when the enemy missed a shot. You can change it too.
Dim ENEMY_SHOUT_CRIT As String: ENEMY_SHOUT_CRIT = "No, a critical hit!"
Dim ENEMY_SHOUT_HIT As String: ENEMY_SHOUT_HIT = "We've been hit!"
Dim ENEMY_SHOUT_MISS As String: ENEMY_SHOUT_MISS = "Woah, that was close!"
Dim ENEMY_SHOUT_REFLECT1 As String: ENEMY_SHOUT_REFLECT1 = "It's a reflection!"
Dim ENEMY_SHOUT_REFLECT2 As String: ENEMY_SHOUT_REFLECT2 = "Goodness, we saved by armor!"
Dim ENEMY_SHOUT_REFLECT3 As String: ENEMY_SHOUT_REFLECT3 = "Ah, I'm still alive!"

' If you wanna configure the damage variance. By default, damage variance is
' set to 30%.
Dim VARIANCE_EDIT As Boolean: VARIANCE_EDIT = False
Dim DMG_VARIANCE As Single:    DMG_VARIANCE = 0.3 '30% (True variance_edit first)

'==============================================================================
' Rest of the script
'
' It's different from the configuration section. Main script starts here.
' Please do not edit past this point unless you know what are you doing.
'
'==============================================================================
'Initiate Form
battleResultForm.yourDamageLabel.Caption = 0
battleResultForm.enemyDamageLabel.Caption = 0
battleResultForm.bonusDamageLabel.Caption = "Bonus Damage Type: "
battleResultForm.yourHitLabel.Caption = 0
battleResultForm.enemyHitLabel.Caption = 0
battleResultForm.yourDamageBar.Width = 60
battleResultForm.yourDamageBar.Left = 0
battleResultForm.enemyDamageBar.Width = 60
battleResultForm.enemyDamageBar.Left = 204

'Initiate Type
Dim armyType(1 To 2) As String
armyType(1) = THIS_WORKSHEET.Range(YOUR_TYPE_CELL).Value
armyType(2) = THIS_WORKSHEET.Range(ENEMY_TYPE_CELL).Value

'Initiate Bonus
Dim bonusType(1 To 2) As Single
If armyType(1) = "Infantry" And armyType(2) = "Artillery" Then
    bonusType(1) = 1.2
    ElseIf armyType(1) = "Vehicle" And armyType(2) = "Infantry" Then
        bonusType(1) = 1.2
        ElseIf armyType(1) = "Tank" And armyType(2) = "Vehicle" Then
            bonusType(1) = 1.2
            ElseIf armyType(1) = "Helicopter" And armyType(2) = "Tank" Then
                bonusType(1) = 1.2
                ElseIf armyType(1) = "Aircraft" And armyType(2) = "Helicopter" Then
                    bonusType(1) = 1.2
                    ElseIf armyType(1) = "Artillery" And armyType(2) = "Aircraft" Then
                        bonusType(1) = 1.2
                        Else: bonusType(1) = 1
End If
'Bonus Musuh
If armyType(2) = "Infantry" And armyType(1) = "Artillery" Then
    bonusType(2) = 1.2
    ElseIf armyType(2) = "Vehicle" And armyType(1) = "Infantry" Then
        bonusType(2) = 1.2
        ElseIf armyType(2) = "Tank" And armyType(1) = "Vehicle" Then
            bonusType(2) = 1.2
            ElseIf armyType(2) = "Helicopter" And armyType(1) = "Tank" Then
                bonusType(2) = 1.2
                ElseIf armyType(2) = "Aircraft" And armyType(1) = "Helicopter" Then
                    bonusType(2) = 1.2
                    ElseIf armyType(2) = "Artillery" And armyType(1) = "Aircraft" Then
                        bonusType(2) = 1.2
                        Else: bonusType(2) = 1
End If

'Accuracy and Critical
Dim acc(1 To 2), crit(1 To 2), resultDamage(1 To 2) As Single, randNum(1 To 2) As Single

'acc
acc(1) = 100 * (THIS_WORKSHEET.Range(HIT_RATE & YOUR_STAT_ROW).Value / (THIS_WORKSHEET.Range(HIT_RATE & YOUR_STAT_ROW).Value + THIS_WORKSHEET.Range(EVASION & ENEMY_STAT_ROW).Value))
acc(2) = 100 * (THIS_WORKSHEET.Range(HIT_RATE & ENEMY_STAT_ROW).Value / (THIS_WORKSHEET.Range(HIT_RATE & ENEMY_STAT_ROW).Value + THIS_WORKSHEET.Range(EVASION & YOUR_STAT_ROW).Value))
'crit
crit(1) = (THIS_WORKSHEET.Range(CRIT_RATE & YOUR_STAT_ROW).Value - THIS_WORKSHEET.Range(CRIT_EVASION & ENEMY_STAT_ROW).Value)
crit(2) = (THIS_WORKSHEET.Range(CRIT_RATE & ENEMY_STAT_ROW).Value - THIS_WORKSHEET.Range(CRIT_EVASION & YOUR_STAT_ROW).Value)

'random number generator
randNum(1) = Int(1 + (Rnd * (100 - 1 + 1)))
randNum(2) = Int(1 + (Rnd * (100 - 1 + 1)))
'result damage teman
If randNum(1) < acc(1) And randNum(1) < crit(1) Then
    resultDamage(1) = bonusType(1) * (THIS_WORKSHEET.Range(DAMAGE & YOUR_STAT_ROW).Value * (THIS_WORKSHEET.Range(CRIT_MULTIPLIER & YOUR_STAT_ROW).Value + 100) / 100)
    ElseIf randNum(1) < acc(1) And Not randNum(1) < crit(1) Then
        resultDamage(1) = bonusType(1) * THIS_WORKSHEET.Range(DAMAGE & YOUR_STAT_ROW).Value
        Else: resultDamage(1) = 0
End If
'result damage musuh
If randNum(2) < acc(2) And randNum(2) < crit(2) Then
    resultDamage(2) = bonusType(2) * (THIS_WORKSHEET.Range(DAMAGE & ENEMY_STAT_ROW).Value * (THIS_WORKSHEET.Range(CRIT_MULTIPLIER & ENEMY_STAT_ROW).Value + 100) / 100)
    ElseIf randNum(2) < acc(2) And Not randNum(2) < crit(2) Then
        resultDamage(2) = bonusType(2) * THIS_WORKSHEET.Range(DAMAGE & ENEMY_STAT_ROW).Value
        Else: resultDamage(2) = 0
End If

'armor
Dim penetHit(1 To 2) As Single
If (THIS_WORKSHEET.Range(PENETRATION & YOUR_STAT_ROW).Value - THIS_WORKSHEET.Range(ARMOR & ENEMY_STAT_ROW).Value) > 0 Then
    penetHit(1) = 1
    ElseIf (THIS_WORKSHEET.Range(PENETRATION & YOUR_STAT_ROW).Value - THIS_WORKSHEET.Range(ARMOR & ENEMY_STAT_ROW).Value) = 0 Then
        penetHit(1) = 0.5
    Else: penetHit(1) = 0
End If
If (THIS_WORKSHEET.Range(PENETRATION & ENEMY_STAT_ROW).Value - THIS_WORKSHEET.Range(ARMOR & YOUR_STAT_ROW).Value) > 0 Then
    penetHit(2) = 1
    ElseIf (THIS_WORKSHEET.Range(PENETRATION & ENEMY_STAT_ROW).Value - THIS_WORKSHEET.Range(ARMOR & YOUR_STAT_ROW).Value) = 0 Then
        penetHit(2) = 0.5
    Else: penetHit(2) = 0
End If

'final result damage with armor
Dim finalDamage(1 To 2) As Single
Dim dmgVar As Single
If VARIANCE_EDIT = True Then
    dmgVar = DMG_VARIANCE
Else
    dmgVar = 0.3
End If
finalDamage(1) = (resultDamage(1) * ((1 - dmgVar) + (Rnd * (dmgVar * 2)))) * penetHit(1)
finalDamage(2) = (resultDamage(2) * ((1 - dmgVar) + (Rnd * (dmgVar * 2)))) * penetHit(2)

'Waktunya pengurangan HP
THIS_WORKSHEET.Range(YOUR_HP_CELL).Value = THIS_WORKSHEET.Range(YOUR_HP_CELL).Value - finalDamage(2)
THIS_WORKSHEET.Range(ENEMY_HP_CELL).Value = THIS_WORKSHEET.Range(ENEMY_HP_CELL).Value - finalDamage(1)

'Tambahkan Point
THIS_WORKSHEET.Range(YOUR_SCORE_CELL).Value = THIS_WORKSHEET.Range(YOUR_SCORE_CELL).Value + finalDamage(1)
THIS_WORKSHEET.Range(ENEMY_SCORE_CELL).Value = THIS_WORKSHEET.Range(ENEMY_SCORE_CELL).Value + finalDamage(2)

'Apply Battle Result Form
battleResultForm.yourDamageLabel.Caption = Format(finalDamage(1), "#,##0.00")
battleResultForm.enemyDamageLabel.Caption = Format(finalDamage(2), "#,##0.00")

If bonusType(1) > 1 Then
    battleResultForm.bonusDamageLabel.Caption = battleResultForm.bonusDamageLabel.Caption & "Active"
Else: battleResultForm.bonusDamageLabel.Caption = battleResultForm.bonusDamageLabel.Caption & "Inactive"
End If

If randNum(1) < acc(1) And randNum(1) < crit(1) Then
    If finalDamage(1) = 0 Then
        Select Case Int((3 * Rnd) + 1)
            Case 1
                battleResultForm.yourHitLabel.Caption = YOUR_SHOUT_REFLECT1
            Case 2
                battleResultForm.yourHitLabel.Caption = YOUR_SHOUT_REFLECT2
            Case 3
                battleResultForm.yourHitLabel.Caption = YOUR_SHOUT_REFLECT3
        End Select
        Else: battleResultForm.yourHitLabel.Caption = YOUR_SHOUT_CRIT
    End If
    ElseIf randNum(1) < acc(1) And Not randNum(1) < crit(1) Then
        If finalDamage(1) = 0 Then
            Select Case Int((3 * Rnd) + 1)
                Case 1
                    battleResultForm.yourHitLabel.Caption = YOUR_SHOUT_REFLECT1
                Case 2
                    battleResultForm.yourHitLabel.Caption = YOUR_SHOUT_REFLECT2
                Case 3
                    battleResultForm.yourHitLabel.Caption = YOUR_SHOUT_REFLECT3
            End Select
            Else: battleResultForm.yourHitLabel.Caption = YOUR_SHOUT_HIT
        End If
    Else:
        battleResultForm.yourHitLabel.Caption = YOUR_SHOUT_MISS
End If
If randNum(2) < acc(2) And randNum(2) < crit(2) Then 'Critical Hit
    If finalDamage(2) = 0 Then
        Select Case Int((3 * Rnd) + 1)
            Case 1
                battleResultForm.enemyHitLabel.Caption = ENEMY_SHOUT_REFLECT1
            Case 2
                battleResultForm.enemyHitLabel.Caption = ENEMY_SHOUT_REFLECT2
            Case 3
                battleResultForm.enemyHitLabel.Caption = ENEMY_SHOUT_REFLECT3
        End Select
        Else: battleResultForm.enemyHitLabel.Caption = ENEMY_SHOUT_CRIT
    End If
    ElseIf randNum(2) < acc(2) And Not randNum(2) < crit(2) Then 'Normal hit
    If finalDamage(2) = 0 Then
        Select Case Int((3 * Rnd) + 1)
            Case 1
                battleResultForm.enemyHitLabel.Caption = ENEMY_SHOUT_REFLECT1
            Case 2
                battleResultForm.enemyHitLabel.Caption = ENEMY_SHOUT_REFLECT2
            Case 3
                battleResultForm.enemyHitLabel.Caption = ENEMY_SHOUT_REFLECT3
        End Select
        Else: battleResultForm.enemyHitLabel.Caption = ENEMY_SHOUT_HIT
    End If
    Else:
    battleResultForm.enemyHitLabel.Caption = ENEMY_SHOUT_MISS 'Miss
End If

battleResultForm.yourDamageBar.Width = 144 * (finalDamage(1) / ((1 + dmgVar) * 1.2 * THIS_WORKSHEET.Range(DAMAGE & YOUR_STAT_ROW).Value * ((THIS_WORKSHEET.Range(CRIT_MULTIPLIER & YOUR_STAT_ROW).Value + 100) / 100)))
battleResultForm.enemyDamageBar.Width = 144 * (finalDamage(2) / ((1 + dmgVar) * 1.2 * THIS_WORKSHEET.Range(DAMAGE & ENEMY_STAT_ROW).Value * ((THIS_WORKSHEET.Range(CRIT_MULTIPLIER & ENEMY_STAT_ROW).Value + 100) / 100)))
battleResultForm.enemyDamageBar.Left = 204 + (144 - battleResultForm.enemyDamageBar.Width)

'Showing Result form
battleResultForm.Show

End Sub
'===============================================================
'
' > End of the script
'
'===============================================================

