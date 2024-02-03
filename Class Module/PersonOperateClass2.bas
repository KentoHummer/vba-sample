Implements PersonOperateInterface

Private Sub PersonOperateInterface_operatePersonInfo(ByRef person As PersonClass)
    person.name = "bbb"
    person.age = 25
    person.personWeapon.gun = "SVG"
    person.personWeapon.sword = "エクスカリバー"
End Sub
