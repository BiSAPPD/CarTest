VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private varMotor As clsMotorCars
Private varWhell As clsWhellCars



Function FuelBudget(FuelCost As Double, Distance As Double) As Double
FuelBudget = (Distance / Mileage) * FuelCost
End Function

