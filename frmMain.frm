VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "IRAQ-WAR: BodyCount"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3765
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer tmrGetHTML 
      Interval        =   60000
      Left            =   2160
      Top             =   1920
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2280
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Move mouse over header for more Informations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   8
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Count:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   660
   End
   Begin VB.Label lblUS 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Height          =   300
      Left            =   1320
      TabIndex        =   6
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "U.S. Troops Killed in War"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   0
      Left            =   570
      TabIndex        =   5
      Top             =   1560
      Width           =   2640
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3840
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Civilians reported killed in Iraq by US-led military action during 2003"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   585
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3540
   End
   Begin VB.Label lblMax 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lblMin 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Maximum:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Minimum:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   3855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      Top             =   1440
      Width           =   3855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private GetHTML As String
    
    Private WWW_Adress(1) As String
    Private strCutFirst(1) As String
    Private strCutLast(1) As String
    Private strFirstString(1) As String
    Private strLastString(1) As String
    Private Const US = 0
    Private Const IRAQ = 1
    
    Private strMinimum As String
    Private strMaximum As String
    Private strKilled As String

    
Private Sub Form_Load()
  ' Settings for US
    WWW_Adress(US) = "http://www.nynewsday.com/ny-uswarcasualtiesgallery,0,5198262.photogallery"
    strCutFirst(US) = "&nbsp; <font color=""white"">"
    strCutLast(US) = "</font>"
    strFirstString(US) = "> "
    strLastString(US) = " <"
  
  ' Settings for IRAQ
    WWW_Adress(IRAQ) = "http://www.iraqbodycount.net/bodycount.htm"
    strCutFirst(IRAQ) = "The date and time at your location is"
    strCutLast(IRAQ) = "are all online (not print) versions"
    strFirstString(IRAQ) = "<b>"
    strLastString(IRAQ) = "</b>"
    
    
    Me.Show
    tmrGetHTML_Timer
End Sub

Private Sub tmrGetHTML_Timer()
  Dim lngCount As Long
  Dim strTemp As String

    For lngCount = 0 To 1
        lblHeader(lngCount).ToolTipText = "Informations from: " & WWW_Adress(lngCount)
      ' Read HTML-Code in Variable
        GetHTML = Inet1.OpenURL(WWW_Adress(lngCount), icString)
        
      ' Cut HTML-Code in smaller variable
        strTemp = Mid$(GetHTML, InStr(LCase$(GetHTML), LCase$(strCutFirst(lngCount))))
        strTemp = Left$(strTemp, InStr(strTemp, strCutLast(lngCount)))
        
        Select Case lngCount
            Case IRAQ
              ' Read out the minimum
                strMinimum = Mid$(strTemp, InStr(LCase$(strTemp), strFirstString(lngCount)) + Len(strFirstString(lngCount)))
                strMinimum = Left$(strMinimum, InStr(LCase$(strMinimum), strLastString(lngCount)) - 1)
        
              ' Cut HTML-Code after the minimum
                strTemp = Mid$(strTemp, InStr(strTemp, strLastString(lngCount)) + Len(strLastString(lngCount)))
        
              ' Read out the maximum
                strMaximum = Mid$(strTemp, InStr(LCase$(strTemp), strFirstString(lngCount)) + Len(strFirstString(lngCount)))
                strMaximum = Left$(strMaximum, InStr(LCase$(strMaximum), strLastString(lngCount)) - 1)
        
                lblMin.Caption = strMinimum
                lblMax.Caption = strMaximum
            Case US
              ' Read the count
                strKilled = Mid$(strTemp, InStr(LCase$(strTemp), strFirstString(lngCount)) + Len(strFirstString(lngCount)))
                strKilled = Left$(strKilled, InStr(LCase$(strKilled), strLastString(lngCount)) - 1)
                lblUS.Caption = Mid$(strKilled, InStrRev(strKilled, " ") + 1)
        End Select
    Next lngCount
End Sub
