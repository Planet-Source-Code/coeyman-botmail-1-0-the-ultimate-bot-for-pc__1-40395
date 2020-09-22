VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Serversetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales de stockage des valeurs de propriétés
Private mpopserver As Variant 'copie locale
Private msmtpserver As Variant 'copie locale
Private musername As Variant 'copie locale
Private mdomainname As Variant 'copie locale
Private mpassword As Variant 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mPOPAuth As Variant 'copie locale
Private mverify As Variant 'copie locale
Public Property Let verify(ByVal vData As Variant)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.verify = 5
    mverify = vData
End Property

Public Property Set verify(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.verify = Form1
    Set mverify = vData
End Property

Public Property Get verify() As Variant
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.verify
    If IsObject(mverify) Then
        Set verify = mverify
    Else
        verify = mverify
    End If
End Property

Public Property Let POPAuth(ByVal vData As Variant)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.POPAuth = 5
    mPOPAuth = vData
End Property


Public Property Set POPAuth(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.POPAuth = Form1
    Set mPOPAuth = vData
End Property


Public Property Get POPAuth() As Variant
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.POPAuth
    If IsObject(mPOPAuth) Then
        Set POPAuth = mPOPAuth
    Else
        POPAuth = mPOPAuth
    End If
End Property

Public Sub initserv()
Dim tmpstr As String
Dim NC As Integer
tmpstr = String(255, 0)
NC = GetPrivateProfileString("server", "smtpserver", "", tmpstr, 255, App.Path + "\" + App.EXEName + ".ini")
msmtpserver = Left(tmpstr, NC)
tmpstr = String(255, 0)
NC = GetPrivateProfileString("server", "popserver", "", tmpstr, 255, App.Path + "\" + App.EXEName + ".ini")
mpopserver = Left(tmpstr, NC)
tmpstr = String(255, 0)
NC = GetPrivateProfileString("server", "username", "", tmpstr, 255, App.Path + "\" + App.EXEName + ".ini")
musername = Left(tmpstr, NC)
tmpstr = String(255, 0)
NC = GetPrivateProfileString("server", "domainname", "", tmpstr, 255, App.Path + "\" + App.EXEName + ".ini")
mdomainname = Left(tmpstr, NC)
tmpstr = String(255, 0)
NC = GetPrivateProfileString("server", "Password", "", tmpstr, 255, App.Path + "\" + App.EXEName + ".ini")
mpassword = Left(tmpstr, NC)

End Sub

Public Property Let password(ByVal vData As Variant)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.password = 5
    mpassword = vData
End Property

Public Property Set password(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.password = Form1
    Set mpassword = vData
End Property

Public Property Get password() As Variant
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.password
    If IsObject(mpassword) Then
        Set password = mpassword
    Else
        password = mpassword
    End If
End Property



Public Property Let domainname(ByVal vData As Variant)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.domainname = 5
    mdomainname = vData
End Property


Public Property Set domainname(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.domainname = Form1
    Set mdomainname = vData
End Property


Public Property Get domainname() As Variant
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.domainname
    If IsObject(mdomainname) Then
        Set domainname = mdomainname
    Else
        domainname = mdomainname
    End If
End Property

Public Property Let username(ByVal vData As Variant)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.username = 5
    musername = vData
End Property

Public Property Set username(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.username = Form1
    Set musername = vData
End Property

Public Property Get username() As Variant
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.username
    If IsObject(musername) Then
        Set username = musername
    Else
        username = musername
    End If
End Property

Public Property Let smtpserver(ByVal vData As Variant)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.smtpserver = 5
    msmtpserver = vData
End Property

Public Property Set smtpserver(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.smtpserver = Form1
    Set msmtpserver = vData
End Property

Public Property Get smtpserver() As Variant
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.smtpserver
    If IsObject(msmtpserver) Then
        Set smtpserver = msmtpserver
    Else
        smtpserver = msmtpserver
    End If
End Property

Public Property Let popserver(ByVal vData As Variant)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.popserver = 5
    mpopserver = vData
End Property

Public Property Set popserver(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.popserver = Form1
    Set mpopserver = vData
End Property

Public Property Get popserver() As Variant
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.popserver
    If IsObject(mpopserver) Then
        Set popserver = mpopserver
    Else
        popserver = mpopserver
    End If
End Property
