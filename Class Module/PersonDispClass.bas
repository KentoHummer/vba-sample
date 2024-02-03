Private Sub Class_Initialize()

End Sub

Sub DisplayPersonInfo(person As PersonClass)
    ' 引数として渡されたPersonインスタンスの情報を表示
    MsgBox "Name: " & person.name & vbCrLf & "Age: " & person.age
    MsgBox "Gun: " & person.personWeapon.gun & vbCrLf & "Sword: " & person.personWeapon.sword
End Sub
