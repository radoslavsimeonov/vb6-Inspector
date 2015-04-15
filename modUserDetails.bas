Attribute VB_Name = "modUserDetails"
Option Explicit

Public Type Elements
    ID      As Integer
    Desc    As String
    Abbr    As String
    Serial  As String
End Type

Public RegisterHDDIndex As Integer
Public bWorkstationHasOwner As Boolean

Public Rank()        As Elements
Public UnitType()    As Elements

Private iWorkstationOwnerIndex As Integer

Public Property Get WorkstationOwnerIndex() As Integer
    WorkstationOwnerIndex = iWorkstationOwnerIndex
End Property

Property Let WorkstationOwnerIndex(iValue As Integer)
    iWorkstationOwnerIndex = iValue
End Property

Private Sub AddElement(el() As Elements, sAbbr As String, sDesc As String)
Dim i As Integer
    
    i = UBound(el)
    el(i).ID = i
    el(i).Abbr = sAbbr
    el(i).Desc = sDesc
    i = i + 1
    ReDim Preserve el(i)
End Sub

Public Sub FillElements2(el() As Elements, ctrl As Control)
Dim i As Integer
Dim LI As ListItem

    ctrl.ListItems.Clear
    ctrl.ListItems.Add , , ""
    
    For i = 0 To UBound(el) - 1
        Set LI = ctrl.ListItems.Add(, , el(i).Desc)
    Next i
End Sub

Public Function Fill_Rank(cmb As Control, Optional sCase As Integer = 0) As Elements()
Dim tmpStr() As String
ReDim Rank(0)
    
    AddElement Rank, "OR-1", RankSplit("Private;Матрос 1 клас", sCase)
    AddElement Rank, "OR-4", RankSplit("Lance Corporal;Старши матрос", sCase)
'    AddElement Rank, "OR-4", RankSplit("Ефрейтор 1 клас;Старши матрос 1 клас", sCase)
'    AddElement Rank, "OR-4", RankSplit("Ефрейтор 2 клас;Старши матрос 2 клас", sCase)
    AddElement Rank, "OR-5", RankSplit("Corporal;Старшина II степен", sCase)
    AddElement Rank, "OR-6", RankSplit("Sergeant;Старшина I степен", sCase)
    AddElement Rank, "OR-7", RankSplit("Staff/Colour Sergeant;Главен старшина", sCase)
    AddElement Rank, "OR-8", RankSplit("Warrant Officer Class 2;Мичман", sCase)
    AddElement Rank, "OR-9", RankSplit("Warrant Officer Class 1;Офицерски кандидат", sCase)
'    AddElement Rank, "OR-9", RankSplit("Офицерски кандидат 1 клас;Офицерски кандидат 1 клас", sCase)
'    AddElement Rank, "OR-9", RankSplit("Офицерски кандидат 2 клас;Офицерски кандидат 2 клас", sCase)
    AddElement Rank, "OF-1", RankSplit("Officer Cadet;Лейтенант", sCase)
    AddElement Rank, "OF-1", RankSplit("Lieutenant;Старши лейтенант", sCase)
    AddElement Rank, "OF-2", RankSplit("Captain;Капитан-лейтенант", sCase)
    AddElement Rank, "OF-3", RankSplit("Major;Капитан III ранг", sCase)
    AddElement Rank, "OF-4", RankSplit("Lieutenant Colonel;Капитан II ранг", sCase)
    AddElement Rank, "OF-5", RankSplit("Colonel;Капитан I ранг", sCase)
    AddElement Rank, "OF-6", RankSplit("Brigadier;Комодор", sCase)
    AddElement Rank, "OF-7", RankSplit("Major General;Контраадмирал", sCase)
    AddElement Rank, "OF-8", RankSplit("Lieutenant General;Вицеадмирал", sCase)
    AddElement Rank, "OF-9", RankSplit("General;Адмирал", sCase)
    AddElement Rank, "CIV", RankSplit("Civilian;Цивилен служител", sCase)

    Fill_Rank = Rank
    FillElements2 Rank, cmb
End Function

Private Function RankSplit(str As String, Pos As Integer) As String
Dim tmpStr() As String
    
    tmpStr = Split(str, ";")
    RankSplit = tmpStr(Pos)
End Function

Public Function Fill_UnitType(cmb As Control) As Elements()
ReDim UnitType(0)

    AddElement UnitType, "MOD", "Министерство на отбраната"
    AddElement UnitType, "JFC", "Съвместно Командване на силите"
    AddElement UnitType, "LF", "Сухопътни Войски"
    AddElement UnitType, "AF", "Военновъздушни сили"
    AddElement UnitType, "NAVY", "Военноморски сили"
    AddElement UnitType, "SCI", "Научно-приложна или образователна"
    AddElement UnitType, "DRU", "Друга, пряко подчинена на Министъра"

    Fill_UnitType = UnitType
    FillElements2 UnitType, cmb
End Function

Public Sub SwitchForcesRanks(iList As Integer, cmb As Control)

    Dim tmpIdx As Integer
    
    tmpIdx = cmb.SelectedItem.Index
    
    cmb.Clear
    Rank = Fill_Rank(cmb, iList)
    If Len(cmb.Text) > 0 Then cmb.ListIndex = tmpIdx
    
End Sub

