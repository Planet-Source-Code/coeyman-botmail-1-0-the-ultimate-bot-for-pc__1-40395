Attribute VB_Name = "mdlscript"
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public les_mails As String
Public Const maxshort = 100
Public Const maxinfo = 5

Dim tabshort(1 To maxshort, 1 To maxinfo) As String

Public dmdtrans As Boolean
Public transfer As Boolean
Public emailatrans As String

Public onlynew As Boolean
Public onlysel As Boolean
Public emailsel As String
Public istorefresh As Boolean

Public Sendmecom As String
Public vmailcom As String
Public executecom As String
Public aspipage As String
Public botstop As String

Public getemailliste As String
Public getnewemail As String
Public getallemail  As String
Public getselemail As String

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long

Dim site As String
Dim addresse As String
Dim filelist As String
Dim zipfile As String

Const scUserAgent = "API-Guide test program"
Const INTERNET_OPEN_TYPE_DIRECT = 1
Const INTERNET_OPEN_TYPE_PROXY = 3
Const INTERNET_FLAG_RELOAD = &H80000000
Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long

Function aspiweb(ByVal url As String, email As String)
Dim addresseweb As String
Dim tabpiece(1 To 10) As String
Dim tabattach(1 To 10) As String
On Error Resume Next


If (InStr(1, UCase(email), "SMTP:") <> 0) Then
  email = Mid(email, InStr(1, email, "smtp:") + 6)
  emailatrans = email
End If
' Suppose que l'adresse URL est toujours valide.
If (InStr(1, UCase(url), Chr(10)) <> 0) Then
  url = Left(url, InStr(1, url, Chr(10)) - 1)
  
End If
If (InStr(1, UCase(url), Chr(13)) <> 0) Then
  url = Left(url, InStr(1, url, Chr(13)) - 1)
End If

addresseweb = Trim(url)
url = addresseweb
addresse = getaddress(addresseweb)
site = getsite(addresseweb)
filelist = ""
MkDir (App.Path + "\tmpaspi")
Debug.Print "en entrée " & addresseweb
Debug.Print "addresse  " & addresse
Debug.Print "site      " & site

VBMail.Inet1.AccessType = icUseDefault
Dim b As String
'b() = VBMail.Inet1.OpenURL(addresseweb, icByteArray)

    Dim hOpen As Long, hFile As Long, sBuffer As String, Ret As Long
    'Create a buffer for the file we're going to download
     sBuffer = Space(1)
    'Create an internet connection
    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    'Open the url
    hFile = InternetOpenUrl(hOpen, addresseweb, vbNullString, ByVal 0&, INTERNET_FLAG_RELOAD, ByVal 0&)
    'Read the first 50 bytes of the file
    InternetReadFile hFile, sBuffer, 1, Ret
    While Ret > 0
        b = b + sBuffer
        sBuffer = Space(1)
        InternetReadFile hFile, sBuffer, 1, Ret
    Wend

    InternetCloseHandle hFile
    InternetCloseHandle hOpen
    'Show our file
    'b = sBuffer

Open "tmpaspi\" + repbywhite(Mid(addresseweb, Len(getaddress(addresseweb)) + 1)) For Binary Access _
Write As #1
Put #1, , b
Close #1

zipfile = "tmpaspi\" + repbywhite(Left(Mid(addresseweb, Len(getaddress(addresseweb)) + 1), InStr(1, Mid(addresseweb, Len(getaddress(addresseweb)) + 1), ".")) + "zip ")
If zipfile = "tmpaspi\zip " Then
    zipfile = "tmpaspi\aspi.zip "
End If

Open zipfile For Binary Access Write As #1
Close #1
zipfile = "tmpaspi\" + GetShortFile(zipfile)
Kill zipfile

Call dowloadlinx(repbywhite(Mid(addresseweb, Len(getaddress(addresseweb)) + 1)), "<link", "href=")
Call dowloadlinx(repbywhite(Mid(addresseweb, Len(getaddress(addresseweb)) + 1)), "<img", "src=")
Call dowloadlinx(repbywhite(Mid(addresseweb, Len(getaddress(addresseweb)) + 1)), "background=", "background=")

Call Shell("C:\program Files\Winzip\Winzip32.exe -a -ex -hs -r " + App.Path + "\" + zipfile + " tmpaspi\*.*", vbMinimizedFocus)
DoEvents
While FindWindow(vbNullString, "WinZip - " + repbywhite(Left(Mid(addresseweb, Len(getaddress(addresseweb)) + 1), InStr(1, Mid(addresseweb, Len(getaddress(addresseweb)) + 1), "."))) + "zip ") <> 0
DoEvents
Wend

' if you want to use PKZIP remove commentary
'Call Shell("PKZIP.EXE -a " + zipfile + " tmpaspi\*.*", vbMinimizedFocus)
'DoEvents
'While FindWindow(vbNullString, "PKZIP.EXE") <> 0
'DoEvents
'Wend

    Dim frmsend As New frmSMTP
    On Error Resume Next
    tabpiece(1) = App.Path + "\" + zipfile
    tabattach(1) = url + ".zip"
    Call frmsend.sendmail(email, "Aspipage " + url, "Botmail <BoTmAiL@nombidon.com>", "", " En piece jointe le fichier " + zipfile, tabpiece, tabattach)
    While frmsend.varquitprog <> True
    DoEvents
    Wend
    Set frmsend = Nothing
    If Err Then
        Debug.Print "erreur"
    End If
    
    Kill "tmpaspi\*.*"
    RmDir ("tmpaspi")
End Function

Function getaddress(ByVal adr As String) As String
Dim Pos As Integer
Pos = 0
While (InStr(Pos + 1, adr, "/") <> 0)
Pos = InStr(Pos + 1, adr, "/")
Wend
getaddress = Mid(adr, 1, Pos)
End Function

Function getpix(ByVal adr As String) As String
Dim Pos As Integer
Pos = 0
While (InStr(Pos + 1, adr, "/") <> 0)
Pos = InStr(Pos + 1, adr, "/")
Wend

getpix = repbywhite(Mid(adr, Pos + 1))

If (InStr(1, getpix, "?") <> 0) Then
    getpix = Mid(getpix, 1, InStr(1, getpix, "?") - 1)
End If
End Function


Function getsite(ByVal adr As String) As String
Dim Pos As Integer

getsite = Mid(adr, InStr(1, adr, "//") + 2)
getsite = Left$(getsite, InStr(1, getsite, "/") - 1)
getsite = "http://" + getsite
End Function


Function dowloadlinx(ByVal webpage As String, ByVal balisedep, ByVal debfic)
Call VBMail.rtb.LoadFile("tmpaspi\" + webpage, rtfText)

Dim possimg As Long
Dim oldpossimg As Long
Dim startimgadr As Long
Dim imgadrlen As Long
Dim pixadr As String
possimg = 1
oldpossimg = 0
While oldpossimg < possimg
    startimgadr = VBMail.rtb.Find(debfic + Chr(34), VBMail.rtb.Find(balisedep, possimg)) + Len(debfic) + 1
    If (startimgadr > possimg) Then
        imgadrlen = VBMail.rtb.Find(Chr(34), startimgadr) - startimgadr
        VBMail.rtb.SelStart = startimgadr
        VBMail.rtb.SelLength = imgadrlen
        VBMail.rtb.SelBold = True
        oldpossimg = possimg
        possimg = startimgadr + imgadrlen
        If (InStr(1, UCase(VBMail.rtb.SelText), "HTTP")) Then
            pixadr = VBMail.rtb.SelText
        Else
            If (Mid(VBMail.rtb.SelText, 1, 1) = "/") Then
                pixadr = site + VBMail.rtb.SelText
            Else
                pixadr = addresse + VBMail.rtb.SelText
            End If
        End If
            Debug.Print pixadr
        
        If (Dir(getpix(VBMail.rtb.SelText)) = "") Then
            VBMail.Inet1.AccessType = icUseDefault
            Dim b() As Byte
            b() = VBMail.Inet1.OpenURL(pixadr, icByteArray)
            Open "tmpaspi\" + getpix(VBMail.rtb.SelText) For Binary Access _
                Write As #1
                Put #1, , b()
            Close #1
        End If
        filelist = filelist + " " + getpix(VBMail.rtb.SelText)
        DoEvents
        VBMail.rtb.SelText = GetShortFile(App.Path + "\tmpaspi\" + getpix(VBMail.rtb.SelText))
        Debug.Print VBMail.rtb.SelText
        
    Else
        possimg = -1
        VBMail.rtb.SelStart = 0
        VBMail.rtb.SelLength = 0
    End If
Wend
Call VBMail.rtb.SaveFile("tmpaspi\" + webpage, rtfText)
End Function

Function repbywhite(ByVal fic As String) As String
While InStr(1, fic, "%20") <> 0
    fic = Left$(fic, InStr(1, fic, "%20") - 1) + " " + Mid(fic, InStr(1, fic, "%20") + 3)
Wend
repbywhite = fic
End Function


Public Function GetShortPath(strFileName As String) As String
    Dim lngRes As Long, strPath As String
    strPath = String$(165, 0)
    lngRes = GetShortPathName(strFileName, strPath, 164)
    'remove all unnecessary chr$(0)'s
    GetShortPath = Left$(strPath, lngRes)
End Function

Public Function GetShortFile(strFileName As String) As String
Dim Pos As Integer

    Dim lngRes As Long, strPath As String
    strPath = String$(165, 0)
    'retrieve the short pathname
    lngRes = GetShortPathName(strFileName, strPath, 164)
    'remove all unnecessary chr$(0)'s
    GetShortFile = Left$(strPath, lngRes)
    Pos = 0
    While (InStr(Pos + 1, GetShortFile, "\") <> 0)
        Pos = InStr(Pos + 1, GetShortFile, "\")
    Wend
    GetShortFile = Mid(GetShortFile, Pos + 1)
    
End Function


Public Function init_tabshort()

 Sendmecom = ""
 vmailcom = ""
 executecom = ""
 botstop = ""


If Dir(App.Path + "\" + App.EXEName + ".ini") = "" Then
    WritePrivateProfileString "Botcom", "Sendmecom", "sendme", App.Path + "\" + App.EXEName + ".ini"
    WritePrivateProfileString "Botcom", "vmailcom", "vmail", App.Path + "\" + App.EXEName + ".ini"
    WritePrivateProfileString "Botcom", "executecom", "execcmd", App.Path + "\" + App.EXEName + ".ini"
    WritePrivateProfileString "Botcom", "aspipage", "aspipage", App.Path + "\" + App.EXEName + ".ini"
    WritePrivateProfileString "Botcom", "botstop", "stopme", App.Path + "\" + App.EXEName + ".ini"
    
    WritePrivateProfileString "vmailcmd", "getemailliste", "getlis", App.Path + "\" + App.EXEName + ".ini"
    WritePrivateProfileString "vmailcmd", "getnewemail", "getnew", App.Path + "\" + App.EXEName + ".ini"
    WritePrivateProfileString "vmailcmd", "getallemail", "getall", App.Path + "\" + App.EXEName + ".ini"
    WritePrivateProfileString "vmailcmd", "getselemail", "getsel", App.Path + "\" + App.EXEName + ".ini"
    
    WritePrivateProfileString "execcmd1", "KeyName", "Botmail", App.Path + "\" + App.EXEName + ".ini"
    WritePrivateProfileString "execcmd1", "AttachName", "Botmail.exe", App.Path + "\" + App.EXEName + ".ini"
    WritePrivateProfileString "execcmd1", "PathName", App.Path + "\" + App.EXEName + ".exe", App.Path + "\" + App.EXEName + ".ini"
    WritePrivateProfileString "execcmd1", "Subject", "Envoie de Botmail", App.Path + "\" + App.EXEName + ".ini"
    WritePrivateProfileString "execcmd1", "texte", "Voici le fichier executable BOTMAIL", App.Path + "\" + App.EXEName + ".ini"
    
    WritePrivateProfileString "server", "smtpserver", InputBox("adress of smtp server ex: smtpuunet.mysite.com", "Server setup"), App.Path + "\" + App.EXEName + ".ini"
    WritePrivateProfileString "server", "popserver", InputBox("adress of pop server ex: Popuunet.mysite.com", "Server setup"), App.Path + "\" + App.EXEName + ".ini"
    WritePrivateProfileString "server", "username", InputBox("username ex: for myname@mysite.com -> myname", "Server setup"), App.Path + "\" + App.EXEName + ".ini"
    WritePrivateProfileString "server", "domainname", InputBox("domain name  ex: for myname@mysite.com -> mysite.com", "Server setup"), App.Path + "\" + App.EXEName + ".ini"
    WritePrivateProfileString "server", "Password", InputBox("Password : ", "server setup"), App.Path + "\" + App.EXEName + ".ini"
    
End If
    Sendmecom = String(255, 0)
    vmailcom = String(255, 0)
    executecom = String(255, 0)
    aspipage = String(255, 0)
    botstop = String(255, 0)
    getemailliste = String(255, 0)
    getnewemail = String(255, 0)
    getallemail = String(255, 0)
    getselemail = String(255, 0)
    
    
    NC = GetPrivateProfileString("Botcom", "Sendmecom", "", Sendmecom, 255, App.Path + "\" + App.EXEName + ".ini")
    If NC <> 0 Then Sendmecom = UCase(Left$(Sendmecom, NC))
    NC = GetPrivateProfileString("Botcom", "vmailcom", "", vmailcom, 255, App.Path + "\" + App.EXEName + ".ini")
    If NC <> 0 Then vmailcom = UCase(Left$(vmailcom, NC))
    NC = GetPrivateProfileString("Botcom", "executecom", "", executecom, 255, App.Path + "\" + App.EXEName + ".ini")
    If NC <> 0 Then executecom = UCase(Left$(executecom, NC))
    NC = GetPrivateProfileString("Botcom", "aspipage", "", aspipage, 255, App.Path + "\" + App.EXEName + ".ini")
    If NC <> 0 Then aspipage = UCase(Left$(aspipage, NC))
    NC = GetPrivateProfileString("Botcom", "botstop", "", botstop, 255, App.Path + "\" + App.EXEName + ".ini")
    If NC <> 0 Then botstop = UCase(Left$(botstop, NC))
    
    NC = GetPrivateProfileString("vmailcmd", "getemailliste", "", getemailliste, 255, App.Path + "\" + App.EXEName + ".ini")
    If NC <> 0 Then getemailliste = UCase(Left$(getemailliste, NC))
    NC = GetPrivateProfileString("vmailcmd", "getnewemail", "", getnewemail, 255, App.Path + "\" + App.EXEName + ".ini")
    If NC <> 0 Then getnewemail = UCase(Left$(getnewemail, NC))
    NC = GetPrivateProfileString("vmailcmd", "getallemail", "", getallemail, 255, App.Path + "\" + App.EXEName + ".ini")
    If NC <> 0 Then getallemail = UCase(Left$(getallemail, NC))
    NC = GetPrivateProfileString("vmailcmd", "getselemail", "", getselemail, 255, App.Path + "\" + App.EXEName + ".ini")
    If NC <> 0 Then getselemail = UCase(Left$(getselemail, NC))


For I = 1 To maxshort
    For j = 1 To maxinfo
            tabshort(I, j) = String(255, 0)
    Next j
    
    NC = GetPrivateProfileString("execcmd" + CStr(I), "KeyName", "", tabshort(I, 1), 255, App.Path + "\" + App.EXEName + ".ini")
    
    If NC <> 0 Then
        tabshort(I, 1) = Left$(tabshort(I, 1), NC)
    Else
        tabshort(I, 1) = ""
    End If
    
    NC = GetPrivateProfileString("execcmd" + CStr(I), "AttachName", "", tabshort(I, 2), 255, App.Path + "\" + App.EXEName + ".ini")
    If NC <> 0 Then tabshort(I, 2) = Left$(tabshort(I, 2), NC)
    NC = GetPrivateProfileString("execcmd" + CStr(I), "PathName", "", tabshort(I, 3), 255, App.Path + "\" + App.EXEName + ".ini")
    If NC <> 0 Then tabshort(I, 3) = Left$(tabshort(I, 3), NC)
    NC = GetPrivateProfileString("execcmd" + CStr(I), "Subject", "", tabshort(I, 4), 255, App.Path + "\" + App.EXEName + ".ini")
    If NC <> 0 Then tabshort(I, 4) = Left$(tabshort(I, 4), NC)
    NC = GetPrivateProfileString("execcmd" + CStr(I), "texte", "", tabshort(I, 5), 255, App.Path + "\" + App.EXEName + ".ini")
    If NC <> 0 Then tabshort(I, 5) = Left$(tabshort(I, 5), NC)
Next I

End Function

Public Function send_shortcut(ByVal shortc As String, email As String) As Boolean
Dim trouve As Boolean
Dim tabpiece(1 To 10) As String
Dim tabattach(1 To 10) As String
Dim frmsend As New frmSMTP
Dim tmptxt As String
On Error Resume Next
send_shortcut = False
trouve = False
VBMail.Inet1.AccessType = icUseDefault
Dim b() As Byte

If (InStr(1, UCase(email), "SMTP:") <> 0) Then
      email = Mid(email, InStr(1, email, "smtp:") + 6)
End If
    
If (shortc <> "") Then
    For I = 1 To maxshort
        If (InStr(1, UCase(shortc), UCase(tabshort(I, 1))) <> 0 And tabshort(I, 1) <> "") Then
            trouve = True
                    If (Dir(tabshort(I, 3), 30) = "") Then
                        b() = VBMail.Inet1.OpenURL(tabshort(I, 3), icByteArray)
                        Open App.Path + "\download\" + "Download" + Right(tabshort(I, 3), 4) For Binary Access _
                        Write As #1
                        Put #1, , b()
                        Close #1
                    End If
                    
                    On Error Resume Next
                    If (Dir(tabshort(I, 3), 30) = "") Then
                        tabpiece(1) = App.Path + "\download\" + "Download" + Right(tabshort(I, 3), 4)
                        tabattach(1) = tabshort(I, 2)
                        Call frmsend.sendmail(email, tabshort(I, 4), "Botmail <BoTmAiL@nombidon.com>", "", tabshort(I, 5) + vbCrLf, tabpiece, tabattach)
                    Else
                        tabpiece(1) = tabshort(I, 3)
                        tabattach(1) = tabshort(I, 2)
                        Call frmsend.sendmail(email, tabshort(I, 4), "Botmail <BoTmAiL@nombidon.com>", "", tabshort(I, 5) + vbCrLf, tabpiece, tabattach)
                    End If
                        While frmsend.varquitprog <> True
                            DoEvents
                        Wend
                        Set frmsend = Nothing
                
                If Err Then
                    send_shortcut = False
                Else
                    send_shortcut = True
                End If
                
        End If
    Next I

    If trouve <> True Then
         
        On Error Resume Next
        tmptxt = "Not found"
        For I = 1 To 10
            tabpiece(I) = ""
            tabattach(I) = ""
        Next I
        For I = 1 To maxshort
            If (tabshort(I, 1) <> "") Then tmptxt = tmptxt + vbCrLf + tabshort(I, 1) + " -> " + tabshort(I, 4)
        Next I
        Call frmsend.sendmail(email, "EXECCMD Return", "Botmail <BoTmAiL@nombidon.com>", "", tmptxt, tabpiece, tabattach)
            While frmsend.varquitprog <> True
             DoEvents
            Wend
            Set frmsend = Nothing
     End If
    
End If
End Function

Public Function send_mail(ByVal shortc As String, email As String) As Boolean
Dim trouve As Boolean
Dim chaine As String
Dim frmsend As New frmSMTP
Dim tabpiece(1 To 10) As String
Dim tabattach(1 To 10) As String
send_mail = False
trouve = False
emailatrans = ""
onlysel = False
onlynew = False
dmdtrans = False

If (InStr(1, UCase(email), "SMTP:") <> 0) Then
      email = Mid(email, InStr(1, email, "smtp:") + 6)
      emailatrans = email
End If
emailsel = ""

If (shortc <> "") Then
        If (InStr(1, UCase(shortc), getallemail) <> 0 And Trim(getallemail) <> "") Then
            dmdtrans = True
            onlynew = False
            onlysel = False
        End If
        If (InStr(1, UCase(shortc), getnewemail) <> 0 And Trim(getnewemail) <> "") Then
            dmdtrans = True
            onlynew = True
            
        End If
        If (InStr(1, UCase(shortc), getselemail) <> 0 And Trim(getselemail) <> "") Then
            dmdtrans = True
            onlysel = True
            emailsel = UCase(shortc)
        End If
        
        If (InStr(1, UCase(shortc), getemailliste) <> 0 And Trim(getemailliste) <> "") Then
            trouve = True
        
                For I = 1 To 10
                    tabpiece(I) = ""
                    tabattach(I) = ""
                Next I
                Call frmsend.sendmail(email, "Les emails", "Botmail <BoTmAiL@nombidon.com>", "", les_mails + vbCrLf, tabpiece, tabattach)
                
                If Err Then
                    send_mail = False
                Else
                    send_mail = True
                End If
        End If
    
End If

End Function

Public Sub sendme(ByVal url As String, ByVal email As String)

On Error Resume Next
Dim tabpiece(1 To 10) As String
Dim tabattach(1 To 10) As String
    VBMail.Inet1.AccessType = icUseDefault
    Dim b() As Byte
    
    If (InStr(1, UCase(email), "SMTP:") <> 0) Then
      email = Mid(email, InStr(1, email, "smtp:") + 6)
      emailatrans = email
    End If
    ' Suppose que l'adresse URL est toujours valide.
    If (InStr(1, UCase(url), Chr(10)) <> 0) Then
      url = Left(url, InStr(1, url, Chr(10)) - 1)
      
    End If
    If (InStr(1, UCase(url), Chr(13)) <> 0) Then
      url = Left(url, InStr(1, url, Chr(13)) - 1)
    End If
    url = Trim(url)

    If (Dir(url, 30) = "") Then
        ' Récupère le fichier dans un tableau d'octets.
        b() = VBMail.Inet1.OpenURL(url, icByteArray)
    
        Open App.Path + "\download\" + "Download" + Right(url, 4) For Binary Access _
        Write As #1
        Put #1, , b()
        Close #1
    End If
     Dim frmsend As New frmSMTP
        On Error Resume Next
        If (Dir(url, 30) = "") Then
            tabpiece(1) = App.Path + "\download\" + "Download" + Right(url, 4)
            tabattach(1) = url + Right(url, 4)
            Call frmsend.sendmail(email, "send me" + url, "Botmail <BoTmAiL@nombidon.com>", "", " En piece jointe le fichier " + url + " " + vbCrLf, tabpiece, tabattach)
        Else
            tabpiece(1) = url
            tabattach(1) = ""
            Call frmsend.sendmail(email, "send me " + url, "Botmail <BoTmAiL@nombidon.com>", "", " En piece jointe le fichier " + url + " " + vbCrLf, tabpiece, tabattach)
        End If
        While frmsend.varquitprog <> True
            DoEvents
        Wend
        Set frmsend = Nothing
    
End Sub
