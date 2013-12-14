VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmPlugins 
   BorderStyle     =   0  'None
   Caption         =   "Plugins"
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   645
   ScaleWidth      =   615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSScriptControlCtl.ScriptControl SC 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   0   'False
   End
End
Attribute VB_Name = "frmPlugins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ExecScript(ByRef Statement As String)

On Error Resume Next
SC.ExecuteStatement Statement


End Sub
