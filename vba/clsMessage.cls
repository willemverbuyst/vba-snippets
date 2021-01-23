Option Explicit

Private Type TMessage
    GoodByeMessageFormal As String
    GoodByeMessageInFormal As String
    WelcomeMessageFormal As String
    WelcomeMessageInFormal As String
End Type

Private This As TMessage

Public Property Get GoodByeMessageFormal() As String
    GoodByeMessageFormal = This.GoodByeMessageFormal
End Property

Public Property Get GoodByeMessageInFormal() As String
    GoodByeMessageInFormal = This.GoodByeMessageInFormal
End Property

Public Property Get WelcomeMessageFormal() As String
    WelcomeMessageFormal = This.WelcomeMessageFormal
End Property

Public Property Get WelcomeMessageInFormal() As String
    WelcomeMessageInFormal = This.WelcomeMessageInFormal
End Property

Private Sub Class_Initialize()
    This.GoodByeMessageFormal = "Have a nice day!"
    This.GoodByeMessageInFormal = "Say you later aligator!"
    This.WelcomeMessageFormal = "Hello, how are you "
    This.WelcomeMessageInFormal = "Wazzup "
End Sub

