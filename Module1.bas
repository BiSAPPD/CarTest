Sub test()
    Dim Cars As clsCars
    Set Cars = New clsCars

    Cars.FillFromSheet ActiveSheet
    
    Dim car As clsCar
    Debug.Print "Test 1: Return all Trades"
    For Each car In cars
        Debug.Print _
            car.motor.Name & vbTab & _
    Next
End Sub