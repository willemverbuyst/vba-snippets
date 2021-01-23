Option Explicit

Public Sub RunWelcome()
    Dim User As New clsUser
    Dim HumanVerified As String
    
    Set User = CreateUser(93, False)
   
    HumanVerified = VerifyHuman(User)
    
    Call SayHi(User, HumanVerified)

End Sub




