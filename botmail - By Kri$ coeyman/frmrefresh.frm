VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Begin VB.Form frmrefresh 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "refresh from"
   ClientHeight    =   570
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   1185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSMAPI.MAPIMessages MapiMess 
      Left            =   600
      Top             =   0
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
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327681
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "frmrefresh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ' Se connecte à la messagerie.
    On Error Resume Next
    MapiSess.Action = 1
    If Err <> 0 Then
        MsgBox "Échec de connexion: " + Error$
    Else
        Screen.MousePointer = 11
        MapiMess.SessionID = MapiSess.SessionID
        
    End If
      
Unload Me
End Sub
