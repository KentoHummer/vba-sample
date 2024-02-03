Private mName As String
Private mAge As Integer
Private mPersonWeapon As PersonWeaponClass

Private Sub Class_Initialize()
    Set mPersonWeapon = New PersonWeaponClass
End Sub

Public Property Get name() As String
    name = mName
End Property

Public Property Let name(value As String)
    mName = value
End Property

Public Property Get age() As Integer
    age = mAge
End Property

Public Property Let age(value As Integer)
    mAge = value
End Property

Public Property Get personWeapon() As PersonWeaponClass
    Set personWeapon = mPersonWeapon
End Property
