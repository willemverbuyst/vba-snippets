Option Explicit

Private Type TUser
    UserAge As Long
    UserIsHuman As Boolean
    UserName As String
End Type

Private This As TUser

Public Property Get UserAge() As Long
    UserAge = This.UserAge
End Property

Public Property Get UserIsHuman() As Boolean
    UserIsHuman = This.UserIsHuman
End Property

Public Property Get UserName() As String
    UserName = This.UserName
End Property

Private Sub Class_Initialize()

    This.UserName = CapitalizeString(Application.UserName)
    
End Sub

Public Sub InitiateProperties(Age As Long, Human As Boolean)

    This.UserAge = Age
    This.UserIsHuman = Human

End Sub

