VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmBrowser 
   BorderStyle     =   0  'None
   ClientHeight    =   1275
   ClientLeft      =   3000
   ClientTop       =   2895
   ClientWidth     =   1710
   LinkTopic       =   "Form1"
   ScaleHeight     =   1275
   ScaleWidth      =   1710
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1680
      ExtentX         =   2963
      ExtentY         =   2143
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
