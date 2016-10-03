VERSION 1.0 CLASS
BEGIN
  MultiUse = -1 
END

Attribute VB_Name = "clsCars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colCars As Collection

Private Sub Class_Initialize()
    Set colCars = New Collection
End Sub

Private Sub Class_Terminate()
    Set colCars = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
    Attribute NewEnum.VB_UserMemId = -4
    Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = colCars.[_NewEnum]
End Property

Public Sub Add(obj As clsCar)
    colCars.Add obj
End Sub

Public Sub Remove(Index As Variant)
    colCars.Remove Index
End Sub

Public Property Get Item(Index As Variant) As clsCar
    Attribute Item.VB_UserMemId = 0
    Set Item = colCars.Item(Index)
End Property

Property Get Count() As Long
    Count = colCars.Count
End Property

Public Sub Clear()
    Set colCars = New Collection
End Sub


Public Sub FillFromSheet(wks As Worksheet)
    Const cFirstRow = 2, cMotorName = 1, cMotorColor = 2, cMotorMileage = 3, cMotorFuelBudget = 4, cWhellName = 5, cWhellRadius = 6

    Dim i As Long, obj As clsCar
    With wks
        For i = cFirstRow To .Cells(Rows.Count, 1).End(xlUp).Row
            Set obj = New clsCar
            obj.Car.Name = .Cells(i, cMotorName)
            Me.Add obj
        Next
    End With
End Sub