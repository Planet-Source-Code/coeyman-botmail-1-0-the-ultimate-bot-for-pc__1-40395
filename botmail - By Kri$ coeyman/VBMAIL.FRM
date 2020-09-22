VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "richtx32.ocx"
Begin VB.MDIForm VBMail 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Botmail"
   ClientHeight    =   3285
   ClientLeft      =   1230
   ClientTop       =   2070
   ClientWidth     =   7680
   Icon            =   "VBMAIL.frx":0000
   ScrollBars      =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   7680
      TabIndex        =   7
      Top             =   855
      Visible         =   0   'False
      Width           =   7680
      Begin VB.Line TopLine 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   0
         X2              =   10800
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   540
      End
      Begin VB.Line TopLine 
         BorderColor     =   &H00000000&
         Index           =   1
         X1              =   15
         X2              =   10800
         Y1              =   505
         Y2              =   505
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   7680
      TabIndex        =   0
      Top             =   2955
      Width           =   7680
      Begin VB.Line MsgBoxSide 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   7260
         X2              =   7260
         Y1              =   60
         Y2              =   300
      End
      Begin VB.Line MsgBoxSide 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   60
         X2              =   60
         Y1              =   60
         Y2              =   300
      End
      Begin VB.Line MsgBoxLine 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   60
         X2              =   7260
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Line MsgBoxLine 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   60
         X2              =   7260
         Y1              =   60
         Y2              =   60
      End
      Begin VB.Line TimeBoxSide 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   8580
         X2              =   8580
         Y1              =   60
         Y2              =   300
      End
      Begin VB.Line TimeBoxLine 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   7320
         X2              =   8580
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Line TimeBoxSide 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   7320
         X2              =   7320
         Y1              =   60
         Y2              =   300
      End
      Begin VB.Line TimeBoxLine 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   7320
         X2              =   8580
         Y1              =   60
         Y2              =   60
      End
      Begin VB.Line TopLine2 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   0
         X2              =   10800
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Label UnreadLbl 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3000
         TabIndex        =   1
         Top             =   60
         Width           =   1575
      End
      Begin VB.Line TopLine2 
         BorderColor     =   &H00000000&
         Index           =   0
         X1              =   0
         X2              =   10800
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label TimeLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Temps"
         Height          =   205
         Left            =   7500
         TabIndex        =   4
         Top             =   75
         Width           =   345
      End
      Begin VB.Label MsgCountLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Information compteur de messages"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   75
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   7620
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   7680
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   6960
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   327681
      End
      Begin VB.Timer Timer2 
         Interval        =   10000
         Left            =   6480
         Top             =   120
      End
      Begin VB.Timer Timer1 
         Interval        =   15000
         Left            =   180
         Top             =   120
      End
      Begin MSMAPI.MAPIMessages MapiMess 
         Left            =   1320
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   327681
         AddressEditFieldCount=   0
         AddressModifiable=   0   'False
         AddressResolveUI=   0   'False
         FetchSorted     =   0   'False
         FetchUnreadOnly =   -1  'True
      End
      Begin MSMAPI.MAPISession MapiSess 
         Left            =   720
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   327681
         DownloadMail    =   -1  'True
         LogonUI         =   -1  'True
         NewSession      =   0   'False
      End
      Begin MSComDlg.CommonDialog CMDialog1 
         Left            =   1920
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         FilterIndex     =   672
      End
      Begin RichTextLib.RichTextBox rtb 
         Height          =   2055
         Left            =   7560
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   3625
         _Version        =   327681
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"VBMAIL.frx":030A
      End
      Begin VB.Label Label1 
         Caption         =   "Ces contrôles ne sont pas visibles durant l'exécution"
         Height          =   315
         Left            =   2700
         TabIndex        =   3
         Top             =   300
         Width           =   3945
      End
   End
   Begin VB.Menu FileMenu 
      Caption         =   "&Fichier"
      Begin VB.Menu PrintMessage 
         Caption         =   "&Imprimer un message"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu PrSetup 
         Caption         =   "&Configuration de l'impression..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu EditMenu 
      Caption         =   "&Edition"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu EditDelete 
         Caption         =   "&Supprimer"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Mail 
      Caption         =   "&Messagerie"
      Begin VB.Menu Logon 
         Caption         =   "&Connexion"
      End
      Begin VB.Menu LogOff 
         Caption         =   "&Déconnexion"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu rMsgList 
         Caption         =   "&Mettre à jour la liste des messages"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Options 
      Caption         =   "&Options"
      Begin VB.Menu MailOpts 
         Caption         =   "&Messagerie..."
      End
      Begin VB.Menu FontS 
         Caption         =   "&Polices"
         Enabled         =   0   'False
         Visible         =   0   'False
         Begin VB.Menu FontScreen 
            Caption         =   "É&cran..."
            Enabled         =   0   'False
         End
         Begin VB.Menu FontPrt 
            Caption         =   "&Imprimante..."
            Enabled         =   0   'False
            Visible         =   0   'False
         End
      End
      Begin VB.Menu DispTools 
         Caption         =   "&Outils d'affichage"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu HelpMenu 
      Caption         =   "&?"
      Begin VB.Menu About 
         Caption         =   "À &propos de..."
      End
   End
End
Attribute VB_Name = "VBMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub About_Click()
    MsgBox "Botmail", 0, "By Kri$"
End Sub

Private Sub DispTools_Click()
    DispTools.Checked = Not DispTools.Checked
    MailLst.Tools.Visible = DispTools.Checked

    
    If MailLst.Tools.Visible Then
        Factor = 1
        ToolsSize% = -MailLst.Tools.Height
    Else
        Factor = -1
        ToolsSize% = 0
    End If

    Select Case MailLst.WindowState
        Case 0    ' Change la taille de la feuille pour refléter un changement dans la barre d'outils.
            MailLst.Height = MailLst.Height + (Factor * MailLst.Tools.Height)
        Case 2    ' Si agrandie, ajuste la taille de la zone de liste.
            MailLst.MList.Height = ScaleHeight - 90 - MailLst.MList.Top + ToolsSize%
    End Select
End Sub



Private Sub Exit_Click()
    ' Ferme l'application et se déconnecte.
    If MapiSess.SessionID <> 0 Then
        Call logoff_Click
    End If
    End
End Sub

Private Sub FontPrt_Click()
    ' Définit les polices d'imprimante.
    On Error Resume Next
    CMDialog1.Flags = 2
    CMDialog1.FontName = Printer.FontName
    CMDialog1.FontSize = Printer.FontSize
    CMDialog1.FontBold = Printer.FontBold
    CMDialog1.FontItalic = Printer.FontItalic
    CMDialog1.ShowFont
    If Err = 0 Then
        Printer.FontName = CMDialog1.FontName
        Printer.FontSize = CMDialog1.FontSize
        Printer.FontBold = CMDialog1.FontBold
        Printer.FontItalic = CMDialog1.FontItalic
    End If

End Sub

Private Sub FontScreen_Click()
    ' Définit les polices écran pour le contrôle actif.
    On Error Resume Next
    CMDialog1.Flags = 1
    CMDialog1.FontName = VBMail.ActiveForm.ActiveControl.FontName
    CMDialog1.FontSize = VBMail.ActiveForm.ActiveControl.FontSize
    CMDialog1.FontBold = VBMail.ActiveForm.ActiveControl.FontBold
    CMDialog1.FontItalic = VBMail.ActiveForm.ActiveControl.FontItalic
    CMDialog1.ShowFont
    If Err = 0 Then
        VBMail.ActiveForm.ActiveControl.FontName = CMDialog1.FontName
        VBMail.ActiveForm.ActiveControl.FontSize = CMDialog1.FontSize
        VBMail.ActiveForm.ActiveControl.FontBold = CMDialog1.FontBold
        VBMail.ActiveForm.ActiveControl.FontItalic = CMDialog1.FontItalic
    End If
End Sub


Private Sub logoff_Click()
    ' Se déconnecte de la messagerie.
    Call LogOffUser
End Sub

Private Sub Logon_Click()
    ' Se connecte à la messagerie.
    On Error Resume Next
    MapiSess.Action = 1
    If Err <> 0 Then
        MsgBox "Échec de connexion: " + Error$
    Else
        Screen.MousePointer = 11
        MapiMess.SessionID = MapiSess.SessionID
        ' Obtient le nombre de messages.
        GetMessageCount
        ' Charge la liste des messages avec l'information enveloppe.
        Screen.MousePointer = 11
        Call LoadList(MapiMess)
        Screen.MousePointer = 0
        ' Ajuste les boutons si nécessaire.
        Logon.Enabled = False
        LogOff.Enabled = True
        
      End If

End Sub

Private Sub MailOpts_Click()
    ' Affiche la feuille des options de la messagerie.
    OptionType = conOptionGeneral
    MailOptFrm.Show 1
End Sub

Private Sub MDIForm_Load()
   On Error Resume Next
    ' S'assure que tous les contrôles sont dimensionnés correctement.
     Call init_tabshort
     istorefresh = False
     TimeLbl = Time$
     SendWithMapi = True
     Call Picture1_Resize
     Call Picture2_Resize
     VBMail.MsgCountLbl = "Hors Ligne"
     Call Logon_Click
     
     MkDir App.Path + "\download"
     
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

Private Sub Picture1_Resize()
Const TimeBoxStartOffset = 1200
Const TimeBoxEndOffset = 60
Const MsgBoxStartOffset = 60
Const MsgBoxEndOffset = TimeBoxStartOffset + 90

    ' Ajuste les dimensions des lignes et la position des étiquettes indiquant le temps.
    TimeLbl.Left = Picture1.Width - TimeLbl.Width - 265
    TopLine2(0).X2 = Picture1.Width
    TopLine2(1).X2 = Picture1.Width

    TimeBoxLine(0).X1 = Picture1.Width - TimeBoxStartOffset
    TimeBoxLine(0).X2 = Picture1.Width - TimeBoxEndOffset

    TimeBoxLine(1).X1 = Picture1.Width - TimeBoxStartOffset
    TimeBoxLine(1).X2 = Picture1.Width - TimeBoxEndOffset

    TimeBoxSide(0).X1 = Picture1.Width - TimeBoxStartOffset
    TimeBoxSide(0).X2 = Picture1.Width - TimeBoxStartOffset

    TimeBoxSide(1).X1 = Picture1.Width - TimeBoxEndOffset
    TimeBoxSide(1).X2 = Picture1.Width - TimeBoxEndOffset

    MsgBoxLine(0).X2 = Picture1.Width - MsgBoxEndOffset
    MsgBoxLine(1).X2 = Picture1.Width - MsgBoxEndOffset

    MsgBoxSide(1).X1 = Picture1.Width - MsgBoxEndOffset
    MsgBoxSide(1).X2 = Picture1.Width - MsgBoxEndOffset

    Picture1.Refresh
End Sub

Private Sub Picture2_Resize()
    ' Ajuste les positions des lignes.
    TopLine(0).X2 = Picture2.Width
    TopLine(1).X2 = Picture2.Width
    Picture2.Refresh
End Sub

Private Sub PrSetup_Click()
' Appelle la procédure d'initialisation de l'imprimante
' dans le contrôle des boîtes de dialogue communes.
On Error Resume Next
    CMDialog1.Flags = &H40  ' Boîte de dialogue Configuration de l'impression seulement.
    CMDialog1.ShowPrinter
End Sub

Private Sub rMsgList_Click()
        Screen.MousePointer = 11
        GetMessageCount
        Call LoadList(MapiMess)
        Screen.MousePointer = 0
End Sub

Private Sub Timer1_Timer()
    ' Met à jour les étiquettes indiquant le temps.
    TimeLbl = Time$
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If istorefresh = True Then Load frmrefresh
tmpind = MailLst.MList.ListIndex
Call rMsgList_Click
MailLst.MList.ListIndex = tmpind
End Sub





