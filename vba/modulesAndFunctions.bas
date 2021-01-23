Option Explicit
Option Private Module

Public Function VerifyHuman(UserToCheck As clsUser) As String

    If Not UserToCheck.UserName = vbNullString _
        And CBool(UserToCheck.UserAge) _
        And UserToCheck.UserIsHuman _
    Then
        VerifyHuman = "human."
    Else
        VerifyHuman = "NOT human."
    End If

End Function

Public Sub SayHi(UserToWelcome As clsUser, HumanVerified As String)

    Dim Message As WelcomeAddIn.clsMessage
    Set Message = InstantiateMessageClass
    
    If UserToWelcome.UserAge > 30 Then
        MsgBox Prompt:=Message.WelcomeMessageFormal _
            & UserToWelcome.UserName & "." _
            & vbNewLine _
            & "You are " & UserToWelcome.UserAge & " years old." _
            & vbNewLine _
            & "You are " & HumanVerified _
            & vbNewLine _
            & Message.GoodbyeMessageFormal, _
            Title:="USER PROFILE"
    Else
        MsgBox Prompt:=Message.WelcomeMessageInFormal _
            & UserToWelcome.UserName & "!" _
            & vbNewLine _
            & "You are " & UserToWelcome.UserAge & " years old." _
            & vbNewLine _
            & "You are " & HumanVerified _
            & vbNewLine _
            & Message.GoodbyeMessageInFormal, _
            Title:="USER PROFILE"
    End If
    
End Sub

