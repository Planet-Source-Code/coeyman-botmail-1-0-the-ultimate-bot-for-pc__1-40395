VERSION 5.00
Begin VB.Form NewMsg 
   Caption         =   "Envoyer un message"
   ClientHeight    =   4755
   ClientLeft      =   1875
   ClientTop       =   2160
   ClientWidth     =   8505
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4755
   ScaleWidth      =   8505
   Begin VB.PictureBox AttachWin 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C0C0&
      Height          =   825
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   8445
      TabIndex        =   14
      Top             =   3930
      Visible         =   0   'False
      Width           =   8505
      Begin VB.ListBox aList 
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   300
         Width           =   5835
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pièces jointes:"
         Height          =   195
         Left            =   60
         TabIndex        =   17
         Top             =   60
         Width           =   930
      End
      Begin VB.Label NumAtt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1200
         TabIndex        =   16
         Top             =   60
         Width           =   45
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   8505
      TabIndex        =   8
      Top             =   0
      Width           =   8505
      Begin VB.CommandButton CompAdd 
         Caption         =   "A&dresse"
         Height          =   330
         Left            =   5940
         TabIndex        =   13
         Top             =   90
         Width           =   1335
      End
      Begin VB.CommandButton CompOpt 
         Caption         =   "&Options"
         Height          =   330
         Left            =   4500
         TabIndex        =   12
         Top             =   90
         Width           =   1335
      End
      Begin VB.CommandButton Attach 
         Caption         =   "&Pièce jointe"
         Height          =   330
         Left            =   3060
         TabIndex        =   11
         Top             =   90
         Width           =   1335
      End
      Begin VB.CommandButton ChkNames 
         Caption         =   "&Vérifier les noms"
         Height          =   330
         Left            =   1620
         TabIndex        =   10
         Top             =   90
         Width           =   1335
      End
      Begin VB.CommandButton Send 
         Caption         =   "&Envoyer"
         Height          =   330
         Left            =   180
         TabIndex        =   9
         Top             =   90
         Width           =   1335
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   540
      End
      Begin VB.Line TopLine2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   7380
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.TextBox txtNoteText 
      Height          =   1275
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1890
      Width           =   7395
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   1395
      Left            =   0
      ScaleHeight     =   1395
      ScaleWidth      =   8505
      TabIndex        =   4
      Top             =   495
      Width           =   8505
      Begin VB.TextBox txtTo 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   900
         TabIndex        =   0
         Top             =   180
         Width           =   4995
      End
      Begin VB.TextBox txtcc 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   900
         TabIndex        =   1
         Top             =   540
         Width           =   4995
      End
      Begin VB.TextBox txtsubject 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   900
         TabIndex        =   2
         Top             =   900
         Width           =   4995
      End
      Begin VB.Line TopLine 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   0
         X2              =   7320
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line LeftLine 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1380
      End
      Begin VB.Line TopLine 
         BorderColor     =   &H00000000&
         Index           =   0
         X1              =   0
         X2              =   7320
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&A:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cc:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Sujet:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   720
      End
   End
End
Attribute VB_Name = "NewMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Attach_Click()
' Identifie une pièce jointe.
On Error Resume Next
   VBMail.CMDialog1.DialogTitle = "Pièce jointe"
   VBMail.CMDialog1.Filter = "Tous les fichiers(*.*)|*.*|Fichiers texte(*.txt)|*.txt"
   VBMail.CMDialog1.ShowOpen
   If Err = 0 Then
        On Error GoTo 0
        VBMail.MapiMess.AttachmentIndex = VBMail.MapiMess.AttachmentCount
        VBMail.MapiMess.AttachmentName = VBMail.CMDialog1.FileTitle
        VBMail.MapiMess.AttachmentPathName = VBMail.CMDialog1.FileName
        VBMail.MapiMess.AttachmentPosition = VBMail.MapiMess.AttachmentIndex
        VBMail.MapiMess.AttachmentType = vbAttachTypeData
   End If
End Sub

Private Sub ChkNames_Click()
    ' Résout les noms.
    Call CopyNamestoMsgBuffer(Me, True)
    Call UpdateRecips(Me)
End Sub

Private Sub CompAdd_Click()
    ' Affiche le Carnet d'adresses et le met à jour lors du retour.
    Call CopyNamestoMsgBuffer(Me, False)
    VBMail.MapiMess.Action = vbMessageShowAdBook
    Call UpdateRecips(Me)
End Sub

Private Sub CompOpt_Click()
    ' Affiche la feuille d'options du message.
    OptionType = conOptionMessage
    MailOptFrm.Show 1
End Sub

Private Sub Form_Activate()
    ' Affecte la valeur -1 à la propriété MsgIndex (Tampon de composition) quand cette fenêtre est activée.
    VBMail.MapiMess.MsgIndex = -1
End Sub

Private Sub Form_Load()
    ' S'assure que les fenêtres sont dimensionnées correctement.
    Call Picture1_Resize
    Call Picture2_Resize
    Call Form_Resize
End Sub

Private Sub Form_Resize()
    ' Ajuste la taille de la fenêtre si la feuille n'est pas réduite.
    If WindowState <> 1 Then
        If ScaleHeight > txtNoteText.Top Then
            txtNoteText.Height = ScaleHeight - txtNoteText.Top
            txtNoteText.Width = ScaleWidth
        End If
    End If
End Sub

Private Sub Picture1_Resize()
    ' Met à jour les largeurs des champs en ajustant la ligne des contrôles.
    TopLine(0).X2 = Picture1.Width
    TopLine(1).X2 = Picture1.Width
    Picture1.Refresh
End Sub

Private Sub Picture2_Resize()
    ' Met à jour les largeurs des champs en ajustant la ligne des contrôles.
    TopLine2.X2 = Picture2.Width
    Picture2.Refresh
End Sub

Private Sub Send_Click()
    ' Place le texte du sujet et du message dans le tampon.
    ' Ajoute des caractères au début du message pour les fichiers des pièces jointes.
    If VBMail.MapiMess.AttachmentCount > 0 Then
        txtNoteText = String$(VBMail.MapiMess.AttachmentCount, "*") + txtNoteText
    End If
    VBMail.MapiMess.MsgSubject = txtsubject
    VBMail.MapiMess.MsgNoteText = txtNoteText
    VBMail.MapiMess.MsgReceiptRequested = ReturnRequest
    Call CopyNamestoMsgBuffer(Me, True)
                  
    On Error Resume Next
    VBMail.MapiMess.Action = vbMessageSend
    If Err Then
        MsgBox "Une erreur est survenue pendant l'envoi d'un message: " + Str$(Err)
    Else
        Unload Me
    End If
End Sub

