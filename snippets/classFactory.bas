Option Explicit
Option Private Module

Public Function CreateUser(Age As Long, Human As Boolean) As clsUser

    Set CreateUser = New clsUser

    CreateUser.InitiateProperties Age:=Age, Human:=Human

End Function

