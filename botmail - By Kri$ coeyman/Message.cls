VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private pMessageCC As String 'local copy
Private pMessageText As String 'local copy
Private pReturnPath As String 'local copy
Private pMessageId As String 'local copy
Private pSubject As String 'local copy
Private pMessageTo As String 'local copy
Private pMessageFrom As String 'local copy
Private pMessageDate As String 'local copy
Private pContentType As String 'local copy
Private pXHeaders As Collection 'local copy
Private ponlytext As String 'local copy
Private pFileAttached(1 To 10) As String
Private pnameFileAttached(1 To 10) As String

'variables locales de stockage des valeurs de propriétés
Private mvarRecepientname As Variant 'copie locale
Public Property Let Recepientname(ByVal vData As Variant)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.Recepientname = 5
    mvarRecepientname = vData
End Property


Public Property Set Recepientname(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.Recepientname = Form1
    Set mvarRecepientname = vData
End Property


Public Property Get Recepientname() As Variant
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.Recepientname
    If IsObject(mvarRecepientname) Then
        Set Recepientname = mvarRecepientname
    Else
        Recepientname = mvarRecepientname
    End If
End Property

Public Property Let FileAttached(ByVal I As Integer, ByVal vData As String)
    pFileAttached(I) = vData
End Property
Public Property Get FileAttached(ByVal I As Integer) As String
    FileAttached = pFileAttached(I)
End Property
Public Property Let nameFileAttached(ByVal I As Integer, ByVal vData As String)
    pnameFileAttached(I) = vData
End Property
Public Property Get nameFileAttached(ByVal I As Integer) As String
    nameFileAttached = pnameFileAttached(I)
End Property


Public Property Set XHeaders(ByVal vData As Collection)
    Set pXHeaders = vData
End Property
Public Property Get XHeaders() As Collection
    Set XHeaders = pXHeaders
End Property

Public Property Let ContentType(ByVal vData As String)
    pContentType = vData
End Property
Public Property Get ContentType() As String
    ContentType = pContentType
End Property

Public Property Let MessageDate(ByVal vData As String)
    pMessageDate = vData
End Property
Public Property Get MessageDate() As String
    MessageDate = pMessageDate
End Property

Public Property Let MessageFrom(ByVal vData As String)
    pMessageFrom = vData
End Property
Public Property Get MessageFrom() As String
    MessageFrom = pMessageFrom
End Property

Public Property Let MessageTo(ByVal vData As String)
    pMessageTo = vData
End Property
Public Property Get MessageTo() As String
    MessageTo = pMessageTo
End Property

Public Property Let Onlytext(ByVal vData As String)
    ponlytext = vData
End Property
Public Property Get Onlytext() As String
    Onlytext = ponlytext
End Property

Public Property Let Subject(ByVal vData As String)
    pSubject = vData
End Property
Public Property Get Subject() As String
    Subject = pSubject
End Property

Public Property Let MessageId(ByVal vData As String)
    pMessageId = vData
End Property
Public Property Get MessageId() As String
    MessageId = pMessageId
End Property

Public Property Let ReturnPath(ByVal vData As String)
    pReturnPath = vData
End Property
Public Property Get ReturnPath() As String
    ReturnPath = pReturnPath
End Property

Public Property Let MessageText(ByVal vData As String)
    pMessageText = vData
End Property
Public Property Get MessageText() As String
    MessageText = pMessageText
End Property

Public Property Let MessageCC(ByVal vData As String)
    pMessageCC = vData
End Property
Public Property Get MessageCC() As String
    MessageCC = pMessageCC
End Property

Private Sub Class_Initialize()
    Set pXHeaders = New Collection
End Sub


