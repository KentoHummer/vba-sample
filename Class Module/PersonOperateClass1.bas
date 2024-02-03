Implements PersonOperateInterface

Private Sub PersonOperateInterface_operatePersonInfo(ByRef person As PersonClass)
    person.name = "aaa"
    person.age = 22
    person.personWeapon.gun = "SG234"
    person.personWeapon.sword = "アルテマウェポン"
End Sub
