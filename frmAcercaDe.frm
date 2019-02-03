VERSION 5.00
Begin VB.Form frmAcercaDe 
   BackColor       =   &H00004000&
   BorderStyle     =   0  'None
   Caption         =   "Acerca de CLINICA"
   ClientHeight    =   5085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraFondo 
      BackColor       =   &H8000000E&
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         Height          =   495
         Left            =   5760
         Picture         =   "frmAcercaDe.frx":0000
         TabIndex        =   3
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Frame fraClinica 
         BackColor       =   &H80000009&
         Height          =   2175
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   7935
         Begin VB.Label lblClinicaAtencionPacientes 
            BackColor       =   &H80000005&
            Caption         =   "Clinica - Atención Pacientes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   27.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   360
            TabIndex        =   2
            Top             =   240
            Width           =   4935
         End
         Begin VB.Image Image2 
            Height          =   1425
            Left            =   5760
            Picture         =   "frmAcercaDe.frx":60C1
            Top             =   240
            Width           =   1740
         End
      End
      Begin VB.Label lblUniversidad 
         BackColor       =   &H80000014&
         Caption         =   "Universidad Nacional de La Matanza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         TabIndex        =   9
         Top             =   2520
         Width           =   3375
      End
      Begin VB.Label lblTaller 
         BackColor       =   &H80000014&
         Caption         =   "Taller Programacion Visual Cliente/servidor"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   360
         TabIndex        =   8
         Top             =   2520
         Width           =   3615
      End
      Begin VB.Label lblIntegrantes 
         BackColor       =   &H80000014&
         Caption         =   "Barrios Alejandro    Coronado Raul   Fabbri Daniela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1320
         TabIndex        =   7
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label lblAlumnos 
         BackColor       =   &H80000014&
         Caption         =   "Alumnos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label lblprof 
         BackColor       =   &H80000005&
         Caption         =   "* Martin Battaglia         * Marcelo Rodriguez"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1440
         TabIndex        =   5
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label lblDocentes 
         BackColor       =   &H80000014&
         Caption         =   "Docentes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   3120
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmAcercaDe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
On Error GoTo HuboError

Unload Me

HuboError:
rtaError = evaluarError(Err)
'de acuerdo a la respuesta, realiza...
Select Case rtaError
    Case Finalizar
        End
    Case Reintentar
        Resume
    Case Ignorar
        Resume Next
    Case Cancelar
        'no hace nada
End Select
End Sub

