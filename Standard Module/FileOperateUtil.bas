Sub execute(ByRef personOperate As PersonOperateInterface)
    Dim person As New PersonClass
    personOperate.operatePersonInfo person

    ' Classモジュールのインスタンスを引数として渡す
    Dim personDisp As New PersonDispClass
    personDisp.DisplayPersonInfo person
End Sub
