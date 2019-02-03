VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8145
   ForeColor       =   &H00004000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFondo 
      BackColor       =   &H8000000E&
      Height          =   4635
      Left            =   50
      TabIndex        =   0
      Top             =   60
      Width           =   8040
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   4440
         Top             =   1320
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   3960
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Min             =   1e-4
         Max             =   50000
      End
      Begin VB.Label lblCargaAplic 
         BackColor       =   &H8000000E&
         Caption         =   "Cargando Aplicacion ..."
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Image Image2 
         Height          =   1425
         Left            =   5760
         Picture         =   "frmSplash.frx":0000
         Top             =   480
         Width           =   1740
      End
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
         Height          =   1455
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   4935
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
         TabIndex        =   6
         Top             =   3000
         Width           =   1215
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
         Top             =   3000
         Width           =   1935
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
         Left            =   4200
         TabIndex        =   4
         Top             =   3000
         Width           =   855
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
         Left            =   5400
         TabIndex        =   3
         Top             =   2880
         Width           =   2055
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
         TabIndex        =   2
         Top             =   2160
         Width           =   3615
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
         Left            =   4440
         TabIndex        =   1
         Top             =   2280
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const p = 0

Private Sub Timer1_Timer()
Dim p As Long

While Me.ProgressBar1.Value < Me.ProgressBar1.Max
p = p + 5
Me.ProgressBar1.Value = p
Wend
Timer1.Enabled = False
Load frmLogin
frmLogin.Show
Unload Me

End Sub
