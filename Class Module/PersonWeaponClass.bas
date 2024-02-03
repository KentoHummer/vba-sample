' PersonWeapon Class モジュール

Private mGun As String
Private mSword As String

Private Sub Class_Initialize()

End Sub

Public Property Get gun() As String
    gun = mGun
End Property

Public Property Let gun(value As String)
    mGun = value
End Property

Public Property Get sword() As String
    sword = mSword
End Property

Public Property Let sword(value As String)
    mSword = value
End Property
