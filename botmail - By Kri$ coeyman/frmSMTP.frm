VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmSMTP 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "send email"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   1665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   0
      Top             =   480
   End
   Begin MSWinsockLib.Winsock sckSMTP 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
   End
End
Attribute VB_Name = "frmSMTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim PrevCommand As String
Dim IsSMTPConnected As Boolean
Dim CCAdd As String
Dim BCCAdd As String
Dim Msg As String
Dim intWait As Integer
Dim IsVerified As Boolean
Dim IsAvailable As Boolean
Dim IsCanceled As Boolean
Dim POPAuthorised As Boolean
Dim POPError As Boolean
Dim SignedOut As Boolean
Dim NumOfFilesAttached As Integer
Dim srvsetup As New Serversetup
Dim sendmsg As New Message
Dim K As Variant
Public varquitprog As Boolean
Option Explicit

Public Function sendmail(ByVal destinataire As String, ByVal sujet As String, ByVal nom_envoyeur As String, ByVal nom_receveur As String, ByVal le_texte As String, ByRef piece_jointe() As String, ByRef nom_piece() As String)
Dim indpiece As Integer

 srvsetup.verify = False
 IsCanceled = False
 srvsetup.initserv
 varquitprog = False
 
For indpiece = 1 To 10
     sendmsg.FileAttached(indpiece) = ""
     sendmsg.nameFileAttached(indpiece) = ""
Next indpiece
For indpiece = LBound(piece_jointe) To UBound(piece_jointe)
    If (indpiece < 9) Then
        sendmsg.FileAttached(indpiece + 1) = piece_jointe(indpiece)
        sendmsg.nameFileAttached(indpiece + 1) = nom_piece(indpiece)
    End If
Next indpiece

sendmsg.Subject = sujet
sendmsg.MessageFrom = nom_envoyeur
sendmsg.Recepientname = nom_receveur
sendmsg.Onlytext = le_texte
sendmsg.MessageTo = destinataire


'If need POP authorisation,

 'sckSMTP.Connect srvsetup.popserver, 110
 'PrevCommand = "POPConnect"
'
' 'Wait for authorisation
' Do Until POPAuthorised = True Or POPError = True Or IsCanceled = True
'  DoEvents
' Loop
'
' If POPError Then
'  'If error occured in authorisation.Inform and exit
'  MsgBox "Error in authorisation", vbOKOnly, "Error"
'
'  Exit Function
' End If
'
' 'CLose connection
' sckSMTP.Close
'


'Connect to the SMTP server

sckSMTP.Connect srvsetup.smtpserver, 25

'Start the timer
tmrTimeout.Enabled = True
End Function


Private Sub Form_Load()

 srvsetup.verify = False
 IsCanceled = False
 srvsetup.initserv
 varquitprog = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 'CLose if needed
 If sckSMTP.State <> sckClosed Then sckSMTP.Close
 
 'If we sign in the POP earlier
  If srvsetup.POPAuth = 1 And POPAuthorised = True Then
  
   sckSMTP.Connect srvsetup.popserver, 110
   PrevCommand = "SignOut"
   
   Do Until SignedOut = True Or POPError = True
    DoEvents
   Loop
   'CLose it
   sckSMTP.Close
  
  End If
 
End Sub



Private Sub Label1_Click()

End Sub

Private Sub sckSMTP_Close()
'CLose it
If sckSMTP.State <> sckClosed Then sckSMTP.Close

End Sub


Private Sub sckSMTP_DataArrival(ByVal bytesTotal As Long)

Dim DatRec As String
Dim strBuffer As String
Dim Filenum As Integer
Dim I As Integer
Dim cptcount As Long
Dim Pos As Long
Dim resuline As Long
Dim TempPath As String

sckSMTP.GetData DatRec

'For intercepting SMTP data
Debug.Print DatRec
Select Case Val(Left$(DatRec, 3))

Case 220

 If Not IsSMTPConnected Then
  IsSMTPConnected = True
  tmrTimeout.Enabled = False
  'Notify address
  smtpSendData "HELO " & srvsetup.smtpserver & vbCrLf 'smtp.
  
 
 End If

Case 250

 Select Case PrevCommand
 
 Case "HELO"
 
 
 '--------Start of Verify Codes--------
 
 'Verify the e-mail address.( If srvsetup.verify flag turned on )

  If srvsetup.verify Then
  
   smtpSendData "VRFY " & Left$(sendmsg.MessageTo, InStr(sendmsg.MessageTo, InStr(sendmsg.MessageTo, "@") - 1))
   
   
 
   Do Until intWait >= 10 Or IsVerified = True   'Wait for verification or timeout
     
     If IsCanceled Then  'User Canceled
     
      IsCanceled = False
    
      Exit Sub
     End If
     
     intWait = intWait + 1
     
     DoEvents
   Loop
 
   If IsAvailable Then
   
   
   ElseIf Not IsAvailable Then
    'Inform user
    MsgBox "The recepient not available or server busy.Please try again", vbYesNo, "Recepient Address Verification Failure"
    'Enable the Send button.


    Exit Sub
   End If
  End If
  
  '-----End of Verify Code-------
  
  'Send the sender address..
  smtpSendData "MAIL FROM: <" & srvsetup.username & "@" & srvsetup.domainname & ">" & vbCrLf ' & srvsetup.username & Mid$(srvsetup.smtpserver, 6) & ">" & vbCrLf

  
 Case "MAIL"
  smtpSendData "RCPT TO: <" & Trim$(sendmsg.MessageTo) & ">" & vbCrLf

  
 Case "RCPT"

  smtpSendData "DATA" & vbCrLf
  
 Case "VRFY"
  
  IsVerified = True
  
 Case "DATE"
  'Email sent.
  

  varquitprog = True
  'CLose the connection
  sckSMTP.Close
  IsSMTPConnected = False
  'Sign Out from Server
  If srvsetup.POPAuth = 1 Then
  
   sckSMTP.Connect srvsetup.popserver, 110
   PrevCommand = "SignOut"
   
   Do Until SignedOut = True Or POPError = True
    DoEvents
   Loop
   
   
   'CLose it
   sckSMTP.Close
  
  End If

 
 End Select
 
Case 251
 
 Select Case PrevCommand
 
 Case "RCPT"
  smtpSendData "DATA" & vbCrLf
  
  Debug.Print "Email forwarded to :" & K
 
 End Select
 
Case 354

 Select Case PrevCommand
 
  Case "DATA"
  

   'Server ready for message.Compose the message.
  sendmsg.Onlytext = sendmsg.Onlytext & vbCrLf & vbCrLf
   cptcount = 0
   Pos = 0
    While (InStr(Pos + 1, sendmsg.Onlytext, vbCrLf) <> 0)
        Pos = InStr(Pos + 1, sendmsg.Onlytext, vbCrLf)
        cptcount = cptcount + 1
    Wend
    
   Msg = "DATE: " & Format(Now, "dd mmm yy ttttt") & vbCrLf & "FROM: " & sendmsg.MessageFrom & vbCrLf & "TO: " & sendmsg.Recepientname & vbCrLf & "SUBJECT: " & sendmsg.Subject & vbCrLf
   
   
  
'====================Email Attachments=================

'Attach the UUencoded files...If there are any..

Msg = Msg & "Encoding: " & CStr(cptcount) & " TEXT"

TempPath = App.Path
If Right$(TempPath, 1) <> "\" Then TempPath = TempPath & "\"

'Kill any stupid files out there...
If Dir(TempPath & "Temp.dat") <> "" Then Kill TempPath & "Temp.dat"

Filenum = FreeFile
For I = 1 To 10
    If (sendmsg.FileAttached(I) <> "") Then
        resuline = UUEncode(sendmsg.FileAttached(I), TempPath & "Temp.dat", True, sendmsg.nameFileAttached(I))
  
        Msg = Msg + " , " + CStr(resuline + 2) + " UUENCODE"
    End If
Next

Msg = Msg & vbCrLf & vbCrLf & sendmsg.Onlytext
Open TempPath & "Temp.dat" For Append As #Filenum
 Print #Filenum, "."
Close #Filenum

'Send the header first
 smtpSendData Msg


   strBuffer = Space$(8192)
   sckSMTP.SendData strBuffer
   
 Open TempPath & "Temp.dat" For Binary As Filenum

 Do Until EOF(Filenum)
  strBuffer = Space$(8192)
  Get #Filenum, , strBuffer
  
  sckSMTP.SendData strBuffer
 Loop
 
 
 
 'Mark the end of message
 sckSMTP.SendData vbCrLf & "." & vbCrLf
 
 Close Filenum

 
  
'======================================================


 End Select
'Errors
Case Is >= 400
 
 MsgBox Mid$(DatRec, 4), vbInformation, "Error in Email Transaction"
 varquitprog = True
 'Reenable Send Button


 'Close socket
 sckSMTP.Close
 
 If POPAuthorised Then POPAuthorised = False
 IsSMTPConnected = False
 
End Select

'Intercepting POP messages
Select Case Left$(DatRec, 3)

Case "+OK"
 Select Case PrevCommand
  Case "POPConnect"
   sckSMTP.SendData "USER " & srvsetup.username & vbCrLf
   PrevCommand = "Username"

  Case "Username"
   sckSMTP.SendData "PASS " & srvsetup.password & vbCrLf
   PrevCommand = "Pass"

  Case "Pass"
   POPAuthorised = True

  Case "SignOut"
   sckSMTP.SendData "QUIT"
   POPAuthorised = False
   SignedOut = True
 End Select
Case "-ERR"
 'Error from POP server
 MsgBox "Error " & Mid$(DatRec, 5) & " at POP Server", vbOKOnly, "Error from server"
 varquitprog = True
End Select

If varquitprog = True Then
Unload Me
End If

End Sub

Private Sub smtpSendData(strMessage As String)

PrevCommand = Left$(Trim$(strMessage), 4)

 sckSMTP.SendData strMessage
End Sub

Private Sub sckSMTP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'AN error occured.Display the error
 MsgBox "An error occured [" & Number & "] : " & Description, vbYes, "Error"
 'Close the connection
 sckSMTP.Close
 'Mark unconnected
 IsSMTPConnected = False
 'If User pressed Send button...reenable it

 
End Sub

Private Sub tmrTimeout_Timer()

'Timeourt occured.Display the error


'close the connection
sckSMTP.Close
IsSMTPConnected = False

'disable th etimer
tmrTimeout.Enabled = False
End Sub





