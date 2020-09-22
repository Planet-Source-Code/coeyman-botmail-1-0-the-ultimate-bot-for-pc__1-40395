VERSION 5.00
Begin VB.Form MailOptFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options de la Messagerie"
   ClientHeight    =   3075
   ClientLeft      =   2025
   ClientTop       =   2910
   ClientWidth     =   4740
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3075
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox GeneralOption 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2535
      ScaleWidth      =   4605
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   4605
      Begin VB.CheckBox Chkrefresh 
         Caption         =   "Force le refresh des message à 10 secondes"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2655
      End
      Begin VB.CheckBox retRecip 
         Caption         =   "&Renvoie un accusé de réception"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   3225
      End
      Begin VB.CheckBox SendMapiDLG 
         Caption         =   "&Utiliser la boîte de dialogue Envoyer de MAPI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   60
         Width           =   4395
      End
   End
   Begin VB.CommandButton CancelBt 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   2640
      Width           =   1035
   End
   Begin VB.CommandButton OkBt 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   1035
   End
End
Attribute VB_Name = "MailOptFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelBt_Click()
    ' Décharge la feuille en cours, et n'enregistre pas les changements.
    Unload Me
End Sub









Private Sub Form_Load()
    ' Initialise les valeurs pour les cases à cocher et les champs d'édition.
    username = VBMail.MapiSess.username
    MailPassWord = VBMail.MapiSess.password
    
    LogonUI = Abs(VBMail.MapiSess.LogonUI)
    SendMapiDLG = Abs(SendWithMapi)
    retRecip = Abs(ReturnRequest)
    If (istorefresh = True) Then
        Chkrefresh.Value = 1
    Else
        Chkrefresh.Value = 0
    End If
    Select Case OptionType
        Case conOptionMessage
            Call SetupOptionForm(GeneralOption)
        Case conOptionGeneral
            Call SetupOptionForm(GeneralOption)
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Select Case OptionType
        Case conOptionMessage
            Call SetupOptionForm(GeneralOption)
        Case conOptionGeneral
            Call SetupOptionForm(GeneralOption)
    End Select
End Sub





Private Sub OkBt_Click()
    ' Enregistre les changements de l'utilisateur.
    ' Fonctionnalité supplémentaire: enregistre les valeurs dans un fichier INI.
    VBMail.MapiSess.username = username
    VBMail.MapiSess.password = MailPassWord
    
    VBMail.MapiSess.LogonUI = LogonUI
    SendWithMapi = SendMapiDLG
    ReturnRequest = retRecip
    If Chkrefresh.Value = 1 Then
        istorefresh = True
    Else
        istorefresh = False
    End If
    Unload Me
End Sub

Private Sub SetupOptionForm(BasePic As Control)
    BasePic.Top = 0
    BasePic.Left = 0
    BasePic.Visible = True
    BasePic.Enabled = True
    OkBt.Top = BasePic.Height + 120
    CancelBt.Top = BasePic.Height + 120
    'Me.Width = BasePic.Width + 120
    Me.Height = OkBt.Top + OkBt.Height + 495
End Sub

