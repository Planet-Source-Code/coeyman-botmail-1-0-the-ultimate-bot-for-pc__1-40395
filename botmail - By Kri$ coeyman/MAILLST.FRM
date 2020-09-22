VERSION 5.00
Begin VB.Form MailLst 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Liste de la Messagerie"
   ClientHeight    =   6120
   ClientLeft      =   2070
   ClientTop       =   3270
   ClientWidth     =   7845
   Icon            =   "MAILLST.frx":0000
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6120
   ScaleWidth      =   7845
   Begin VB.PictureBox Tools 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   7845
      TabIndex        =   2
      Top             =   5505
      Width           =   7845
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   0
         X2              =   6660
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   0
         X1              =   0
         X2              =   6660
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.ListBox MList 
      Height          =   2010
      Left            =   90
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Label Headings 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "En-tête de la zone de liste"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   1
      Top             =   60
      Width           =   1680
   End
   Begin VB.Image NewMail 
      Height          =   480
      Left            =   5880
      Picture         =   "MAILLST.frx":030A
      Top             =   2820
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image nonew 
      Height          =   480
      Left            =   5280
      Picture         =   "MAILLST.frx":0614
      Top             =   2880
      Width           =   480
   End
End
Attribute VB_Name = "MailLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variables du module, utilisées pour conserver les informations
' de position lors de l'événement MouseDown.
Dim ListX, ListY

Private Sub Form_Load()
    ' Redimensionne la feuille.
    Height = 3945
    Call Tools_Resize

     ' Définit l'en-tête de la zone de liste.
     a$ = Mid$(Format$("De", "!" + String$(25, "@")), 1, 25)
     b$ = Mid$(Format$("Sujet", "!" + String$(35, "@")), 1, 35)
     c$ = "Date"
     Headings = a$ + b$ + c$
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Si l'utilisateur ferme l'application, laisse décharger la feuille.
    If UnloadMode = 4 Then
        ' Le déchargement est autorisé.
    Else
        ' Si l'utilisateur est toujours connecté, réduit la feuille au lieu de la fermer.
        If VBMail.MapiMess.SessionID <> 0 Then
            Me.WindowState = 1
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_Resize()
    ' Si la feuille n'est pas réduite, redimensionne la zone de
    ' liste pour rentrer dans la feuille.
    If WindowState <> 1 And WindowState <> 2 Then

        If VBMail.DispTools.Checked Then
            xHeight% = Tools.Height
        Else
            xHeight% = 0
        End If

        ' Vérifie la taille minimum de la feuille.
        If Height < 2500 - xHeight% Then
            Height = 2500
            Exit Sub
        End If

        MList.Width = ScaleWidth - MList.Left - 90
        MList.Height = ScaleHeight - 90 - MList.Top - xHeight%
    End If
End Sub

Private Sub MList_Click()
' Définit l'index du message et autorise les boutons

' Précédent et Suivant si nécessaire.
On Error Resume Next
    Select Case MList.ListIndex
        Case 0
            
        Case MList.ListCount - 1
            
        Case Else
            
    End Select
    VBMail.MapiMess.MsgIndex = MList.ListIndex
End Sub

Private Sub MList_DBLClick()
' Vérifie si le message est visible; s'il n'est pas visible,
' le charge dans une nouvelle feuille.
    If Not MailLst.MList.ItemData(MailLst.MList.ListIndex) Then

       MailLst.MList.ItemData(MailLst.MList.ListIndex) = True

     End If
End Sub

Private Sub MList_KeyPress(KeyAscii As Integer)
    ' Si l'utilisateur appuie sur la touche Entrée, traite
    ' l'action comme un événement DblClick.
    If KeyAscii = 13 Then
        Call MList_DBLClick
    End If
End Sub

Private Sub MList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Sauvegarde les coordonnées X et Y afin de déterminer le
    ' départ d'une action Glisser/Déplacer.
    ListX = X
    ListY = Y
End Sub

Private Sub MList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Si le bouton de la souris est cliqué et si les coordonnées X
    ' et Y ont changé, commence à glisser la liste.
    If Button = 1 And ((X <> ListX) Or (Y <> ListY)) Then
        MList.Drag 1
    End If
End Sub



Private Sub Tools_Resize()
    ' Ajuste la largeur des lignes en haut de la barre d'outils.
    Line1(0).X2 = Tools.Width
    Line1(1).X2 = Tools.Width
    Tools.Refresh
End Sub



