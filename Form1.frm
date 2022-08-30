VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form principal 
   Caption         =   "The Miner"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15825
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   15825
   Begin VB.CommandButton stop 
      Caption         =   "X"
      Height          =   435
      Left            =   15000
      TabIndex        =   12
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton okbtn 
      Caption         =   "Iniciar"
      Height          =   495
      Left            =   13920
      TabIndex        =   5
      Top             =   6720
      Width           =   975
   End
   Begin VB.Timer temporizador 
      Interval        =   1000
      Left            =   11760
      Top             =   6120
   End
   Begin SHDocVwCtl.WebBrowser navegador 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15855
      ExtentX         =   27966
      ExtentY         =   11456
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
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
   Begin VB.Label endereco_lbl 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   7200
      Width           =   6855
   End
   Begin VB.Label proprietario_lbl 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   6600
      Width           =   6975
   End
   Begin VB.Label Label6 
      Caption         =   "Endereço"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Proprietário"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Acertos:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label acertos_lbl 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   4
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Acertos / Velocidade :"
      Height          =   255
      Left            =   9720
      TabIndex        =   3
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Inscrição atual:"
      Height          =   255
      Left            =   10320
      TabIndex        =   2
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label inscr_municipal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   1
      Top             =   6480
      Width           =   1935
   End
End
Attribute VB_Name = "principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Option Explicit
 
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
 
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
 
Private Declare Sub mouse_event Lib "user32.dll" ( _
    ByVal dwFlags As Long, _
    ByVal dx As Long, _
    ByVal dy As Long, _
    ByVal cButtons As Long, _
    ByVal dwExtraInfo As Long)
    
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String _
) As Long
 
Private Declare Function GetWindowRect Lib "user32" ( _
    ByVal hwnd As Long, _
    lpRect As RECT _
) As Long
 
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
 
Private Declare Function SetCursorPos Lib "user32" ( _
    ByVal x As Long, _
    ByVal y As Long _
) As Long
 
Private Declare Function GetCursorPos Lib "user32" ( _
    lpPoint As POINTAPI _
) As Long
 
Private Type POINTAPI
    x As Long
    y As Long
End Type
Dim Pt As POINTAPI
 
Private Const SW_SHOWNORMAL = 1
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10
Private Const MOUSEEVENTF_ABSOLUTE = &H8000
Dim winRECT As RECT, Hwnd2 As Long

'-------------------------------Meu código --------------------------------------------------


Private Sub Cliqueduplo()
Hwnd2 = FindWindow("icoPMsgAIM", vbNullString)
 Hwnd2 = FindWindow("icoPMsgAIM", vbNullString)
 GetWindowRect Hwnd2, winRECT
 SetCursorPos Pt.x, Pt.y
 SetWindowPos Hwnd2, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
 mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
 mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
 mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
 mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub



Public Function Reinicia()
    
    Unload Me
    Load Me
    Me.Show
    controle = 10
    navegador.Navigate "http://cidadaoscarlos.giap.com.br/siamun/f?p=111:57:3779871646466756"

End Function









Private Sub navegador_DocumentComplete(ByVal pDisp As Object, URL As Variant)

Dim oDoc As MSHTML.IHTMLDocument2

If navegador.LocationURL <> "" Then
Set oDoc = navegador.Document
codigo = oDoc.body.innerHTML
End If


Me.Caption = setores + 1 & " | " & quadras + 1 & " | " & lotes + 1 & " | " & unidades + 1 & " | " & matricula
If matricula <> 0 Then SanSaveFile App.Path & "\matricula.txt", CStr(matricula)

inscr_municipal.Caption = CStr(matricula)
acertos_lbl.Caption = CStr(acertos) & " - " & Format((acertos * 3600) / (tempo + 1), "0")
proprietario_lbl.Caption = proprietario
endereco_lbl.Caption = endereco
    
'Verifcação de sucesso na identificação do proprietário

If InStr(1, codigo, "P0_BAIRRO_CAB") <> 0 Then

    'Extração do html os valores de proprietário, bairro e endereço
    
    Texto = SanGetEnclosedText(CStr(codigo), "P0_CONTRIBUINTE", "P0_ESTADO")
    proprietario = SanGetEnclosedText(CStr(Texto), "value=", "P0_ENDERECO")
    proprietario = SanGetEnclosedText(CStr(Texto), Chr(34), Chr(34) & " type=")
    endereco = SanGetEnclosedText(CStr(Texto), "P0_ENDERECO value=" & Chr(34), Chr(34) & " type")
    
    
    'Salvamento incremental
    SanSaveFile App.Path & "\db.csv", SanReadFile(App.Path & "\db.csv") & CStr(matricula) & ";" & proprietario & ";" & endereco ' & Chr(13) & Chr(10)
    
    'Incremento do numero de matriculas
    erro = 0
    errolote = 0
    erroquadra = 0
    acertos = acertos + 1
    matricula = matricula + 1
    ultimoincremento = "unidade"
    
    Reinicia
    
ElseIf InStr(1, codigo, "encontrada") <> 0 Then
'Comentar esse código
    Select Case ultimoincremento
        Case "unidade"
        
            If erro > 0 Then
                ultimoincremento = "lote"
                unidades = 0
                lotes = lotes + 1
                matricula = 1001001001 + 1000 * lotes + 1000000 * quadras + 1000000000 * setores
            Else
                erro = erro + 1
                ultimoincremento = "unidade"
                matricula = matricula + 1
            End If
            
    
            Reinicia
        Case "lote"
            
            If errolote > 0 Then
                ultimoincremento = "quadra"
                unidades = 0
                lotes = 0
                quadras = quadras + 1
                matricula = 1001001001 + 1000 * lotes + 1000000 * quadras + 1000000000 * setores
            Else
                errolote = errolote + 1
                lotes = lotes + 1
                ultimoincremento = "lote"
                matricula = 1001001001 + 1000 * lotes + 1000000 * quadras + 1000000000 * setores
            End If
        
        
            Reinicia
            
        Case "quadra"
            If erroquadra > 0 Then
                ultimoincremento = "setor"
                unidades = 0
                lotes = 0
                quadras = 0
                setores = setores + 1
                matricula = 1001001001 + 1000 * lotes + 1000000 * quadras + 1000000000 * setores
            Else
                erroquadra = erroquadra + 1
                quadras = quadras + 1
                ultimoincremento = "quadra"
                matricula = 1001001001 + 1000 * lotes + 1000000 * quadras + 1000000000 * setores
            End If
            'Prossegue no loop
            Reinicia
            
    End Select
End If


inscricao_municipal = CStr(matricula)

If Len(inscricao_municipal) = 10 Then inscricao_municipal = "0" & inscricao_municipal

Select Case controle
    'Casos para navegação primeira
    Case 1
        Pt.x = 655
        Pt.y = 300
        controle = 2
        Call Cliqueduplo
    Case 2
        Pt.x = 100
        Pt.y = 185
        controle = 3
        Call Cliqueduplo
    'Rotina de envio de código
    Case 3
        Pt.x = 455
        Pt.y = 295
        navegador.SetFocus
        Call Cliqueduplo
        DoEvents
        'Rotina de escrever a inscrição
        
        For j = 1 To 11
            SendKeys Mid(inscricao_municipal, j, 1), 100
        Next j
        Pt.x = 655
        Call Cliqueduplo
        controle = 1
        
    
    Case 10
        controle = 11
    Case 11
        Pt.x = 655
        Pt.y = 300
        controle = 12
        Call Cliqueduplo
    Case 12
        Pt.x = 100
        Pt.y = 185
        controle = 13
        Call Cliqueduplo
    'Rotina de envio de código
    Case 13
        Pt.x = 455
        Pt.y = 295
        navegador.SetFocus
        Call Cliqueduplo
        DoEvents
        'Rotina de escrever a inscrição
        
        For j = 1 To 11
            SendKeys Mid(inscricao_municipal, j, 1), 100
        Next j
        Pt.x = 655
        Call Cliqueduplo
End Select
End Sub

'------------------------------------- Fim do meu código


Private Sub okbtn_Click()
Dim diferenca As Double
ultimoincremento = "unidade"

valoranterior = Val(SanReadFile(App.Path & "\matricula.txt"))
diferenca = valoranterior - 1001001001
matricula = valoranterior

setores = Int(diferenca / 1000000000)
quadras = Int(diferenca / 1000000) - setores * 1000
lotes = Int(diferenca / 1000) - setores * 1000000 - quadras * 1000
unidades = diferenca - setores * 1000000000 - quadras * 1000000 - lotes * 1000


controle = 1
navegador.Navigate "http://cidadaoscarlos.giap.com.br/siamun/f?p=111:57:3779871646466756"

End Sub

Private Sub stop_Click()
navegador.stop
End Sub

Private Sub temporizador_Timer()
tempo = tempo + 1
End Sub
