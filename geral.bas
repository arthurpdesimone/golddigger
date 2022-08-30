Attribute VB_Name = "geral"

Public urldaem() As Byte
Public arrWebSites() As String
Public util As Long
Dim GKW() As contador
Public Relaciona As Byte
Public NovoRG As String
Public NovoNascimento As String
Public Novocep As String
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpSFlags As Long, ByVal dwReserved As Long) As Long
Public Const INTERNET_CONNECTION_LAN As Long = &H2
Public Const INTERNET_CONNECTION_MODEM As Long = &H1
'Variaveis sempre usadas
Public NL As String 'ENTER, onde for usado
Public io As String 'Entrada e Saida de Subrotina, onde forem usadas
Public projeto As String 'Nome do Projeto, onde for usado
Public globalpath As String 'Caminho de abertura do projeto, onde for usado
Public achado As Boolean


Type website
ctrl As Long
URL As String
Title As String
Keywords As String
End Type

Type ramo
    Index As Long
    Level As Long
    Key As String
    Path As String
    URI As String
    Title As String
    Navbar As String
    Content As String
    Keywords As String
    Description As String
End Type

Type contador
    palavra As String
    contagem As Long
    Entropia As Long
    Porcentagem As Single
End Type


Public Function Crypt(Text As String) As String
'funcao que encripta e descripta
Dim strTempChar As String
For I = 1 To Len(Text)
If Asc(Mid$(Text, I, 1)) < 128 Then
strTempChar = Asc(Mid$(Text, I, 1)) + 128
ElseIf Asc(Mid$(Text, I, 1)) > 128 Then
strTempChar = Asc(Mid$(Text, I, 1)) - 128
End If
Mid$(Text, I, 1) = Chr(strTempChar)
Next I
Crypt = Text
End Function




'********MANIPULAÇÂO DE ARQUIVOS
'Listar diretorios para String em formato Array
Function SanListDir(fn_SanListDir_path As String, fn_SanListDir_extension As String) As String
    Static fn_SanListDir_fnum As Integer
    Static fn_SanListDir_myPath As String
    Static fn_SanListDir_myName As String
    Static fn_SanListDir_ThePath As String
    Static fn_SanListDir_TheDir As String
    fn_SanListDir_myPath = fn_SanListDir_path
    fn_SanListDir_myName = Dir(fn_SanListDir_myPath + "\" & fn_SanListDir_extension, vbArchive)
    fn_SanListDir_TheDir = ""
    
    Do While fn_SanListDir_myName <> ""
       If fn_SanListDir_myName <> "." And fn_SanListDir_myName <> ".." Then
          fn_SanListDir_ThePath = fn_SanListDir_myPath & "\" & fn_SanListDir_myName
          If (GetAttr(fn_SanListDir_ThePath) And vbArchive) = vbArchive Then
            fn_SanListDir_TheDir = Replace(fn_SanListDir_TheDir, "\\", "\") & fn_SanListDir_ThePath & "|"
          End If   ' it represents a directory.
       End If
       fn_SanListDir_myName = Dir   ' Get next entry.
       DoEvents
    Loop
    If fn_SanListDir_TheDir <> "" Then SanListDir = fn_SanListDir_TheDir
End Function
Function SanOpenDir(fn_SanOpenDir_Dialog As Object, fn_SanOpenDir_Extension As String) As String
    fn_SanOpenDir_Dialog.Filter = fn_SanOpenDir_Extension
    fn_SanOpenDir_Dialog.FileName = fn_SanOpenDir_Extension
    fn_SanOpenDir_Dialog.ShowOpen
    If fn_SanOpenDir_Dialog.FileName <> "" And fn_SanOpenDir_Dialog.FileName <> fn_SanOpenDir_Extension Then
        SanOpenDir = Left(fn_SanOpenDir_Dialog.FileName, InStrRev(fn_SanOpenDir_Dialog.FileName, "\"))
    Else
        SanOpenDir = ""
    End If
End Function

Function SanOpenFile(fn_SanOpenFile_Dialog As Object, fn_SanOpenFile_Extension As String) As String
    fn_SanOpenFile_Dialog.Filter = fn_SanOpenFile_Extension
    fn_SanOpenFile_Dialog.FileName = fn_SanOpenFile_Extension
    fn_SanOpenFile_Dialog.ShowOpen
    If fn_SanOpenFile_Dialog.FileName <> "" And fn_SanOpenFile_Dialog.FileName <> fn_SanOpenFile_Extension Then
        SanOpenFile = SanReadFile(fn_SanOpenFile_Dialog.FileName)
    Else
        SanOpenFile = ""
    End If
End Function
Sub SanTrunkFile(fn_santrunkfile_file As String)
    Static fn_santrunkfile_fnum As Integer
    fn_santrunkfile_fnum = FreeFile
    Open fn_santrunkfile_file For Output As fn_santrunkfile_fnum
    Close #fn_santrunkfile_fnum
End Sub
Sub SanAppendFile(fn_SanAppendFile_file As String, fn_SanAppendFile_text As String)
    Static fn_SanAppendFile_fnum As Integer
    fn_SanAppendFile_fnum = FreeFile
    Open fn_SanAppendFile_file For Append As fn_SanAppendFile_fnum
    Print #fn_SanAppendFile_fnum, fn_SanAppendFile_text
    Close #fn_SanAppendFile_fnum
End Sub
Sub SanSaveFile(fn_SanSaveFile_file As String, fn_SanSaveFile_text As String)
    Static fn_SanSaveFile_fnum As Integer
    fn_SanSaveFile_fnum = FreeFile
    Open fn_SanSaveFile_file For Output As fn_SanSaveFile_fnum
    Print #fn_SanSaveFile_fnum, fn_SanSaveFile_text
    Close #fn_SanSaveFile_fnum
End Sub
Function SanReadFile(fn_SanReadFile_file As String) As String
    If fn_SanReadFile_file = "" Or Dir(fn_SanReadFile_file, vbArchive) = "" Or Right(fn_SanReadFile_file, 1) = "\" Then
        SanReadFile = ""
        Exit Function
    End If
    'Cria variaveis estáticas para não aumentar HEAP
    Static fn_sanreadfile_fnum As Integer
    Static fn_sanreadfile_readchunksize&
    Static fn_sanreadfile_r$
    'Seta seus estados de inicialização / aloca recursos
    fn_sanreadfile_r$ = ""
    fn_sanreadfile_readchunksize& = 0
    'Excecuta a função
    fn_sanreadfile_fnum = FreeFile
    Open fn_SanReadFile_file For Binary As #fn_sanreadfile_fnum
    fn_sanreadfile_readchunksize& = LOF(fn_sanreadfile_fnum)
    fn_sanreadfile_r$ = String$(fn_sanreadfile_readchunksize&, Chr$(0))
    Get #fn_sanreadfile_fnum, , fn_sanreadfile_r$
    'Desaloca recursos
    Close #fn_sanreadfile_fnum
    'Retorna os dados
    SanReadFile = fn_sanreadfile_r$
End Function
'**************MANIPULAÇÂO DE DIRETORIOS
Function SanCreateDir(fn_SanCreateDir_dir As String) As String
    If Dir(fn_SanCreateDir_dir, vbDirectory) = "" Then
        MkDir (fn_SanCreateDir_dir)
        'Tratar erros, quando existirem e retornar vazio se falhar
        SanCreateDir = fn_SanCreateDir_dir
    Else
        SanCreateDir = fn_SanCreateDir_dir
    End If
End Function
Function SanNormalizeDir(diretorio As String) As String
    Msg = diretorio
    Msg = LCase(Msg)
    'Limpa espaços
    'Limpa letra a
    Msg = Replace(Msg, "á", "a")
    Msg = Replace(Msg, "â", "a")
    Msg = Replace(Msg, "ã", "a")
    Msg = Replace(Msg, "à", "a")
    Msg = Replace(Msg, "ä", "a")
    'Limpa letra e
    Msg = Replace(Msg, "é", "e")
    Msg = Replace(Msg, "ê", "e")
    Msg = Replace(Msg, "è", "e")
    Msg = Replace(Msg, "ë", "e")
    'Limpa letra i
    Msg = Replace(Msg, "í", "i")
    Msg = Replace(Msg, "î", "i")
    Msg = Replace(Msg, "ì", "i")
    Msg = Replace(Msg, "ï", "i")
    'Limpa letra o
    Msg = Replace(Msg, "ó", "o")
    Msg = Replace(Msg, "ô", "o")
    Msg = Replace(Msg, "õ", "o")
    Msg = Replace(Msg, "ò", "o")
    Msg = Replace(Msg, "ö", "o")
    'Limpa letra u
    Msg = Replace(Msg, "ú", "u")
    Msg = Replace(Msg, "û", "u")
    Msg = Replace(Msg, "ù", "u")
    Msg = Replace(Msg, "ü", "u")
    'Cedilha
    Msg = Replace(Msg, "ç", "c")
    'Substitui os demais caracteres invalidos por TRAVESSÃO
    For t = 1 To Len(Msg)
        caractere = Mid(Msg, t, 1)
        If InStr(1, "abcdefghijklmnopqrstuvxwyz0123456789_\/&=?@", caractere) = 0 Then Msg = Replace(Msg, caractere, "_")
    Next t
   SanNormalizeDir = Msg
End Function
'*************FUNÇÕES PARA MANIPULAÇÂO DE TEXTOS
Function SanGetEnclosedText(var_SanGetEnclosedText_Texto As String, var_SanGetEnclosedText_Prefix As String, var_SanGetEnclosedText_Postfix As String) As String
    'Especificar depois PREFIXO VAZIO = INICIO DA STRING SUFIXO VAZIO AO FIM DA STRING
    If var_SanGetEnclosedText_Texto = "" Or var_SanGetEnclosedText_Prefix = "" Or var_SanGetEnclosedText_Postfix = "" Then SanGetEnclosedText = "": Exit Function
    Static var_SanGetEnclosedText_pin As Long
    Static var_SanGetEnclosedText_pout As Long
    Static var_SanGetEnclosedText_saida As String
    var_SanGetEnclosedText_pin = InStr(1, var_SanGetEnclosedText_Texto, var_SanGetEnclosedText_Prefix) + Len(var_SanGetEnclosedText_Prefix)
    var_SanGetEnclosedText_pout = InStr(var_SanGetEnclosedText_pin, var_SanGetEnclosedText_Texto, var_SanGetEnclosedText_Postfix)
    If var_SanGetEnclosedText_pin > 0 And var_SanGetEnclosedText_pout <= Len(var_SanGetEnclosedText_Texto) And var_SanGetEnclosedText_pout > var_SanGetEnclosedText_pin Then
        SanGetEnclosedText = Mid(var_SanGetEnclosedText_Texto, var_SanGetEnclosedText_pin, var_SanGetEnclosedText_pout - var_SanGetEnclosedText_pin)
    Else
        SanGetEnclosedText = ""
    End If
End Function
Function SanSoletras(frase As String) As String
    Static caractere As String
    Static t As Long
    Static Msg As String
    'msg = SanNormalizeText(frase)
    Msg = LCase(frase)
    'Limpa espaços
    'Limpa letra a
    Msg = Replace(Msg, "á", "a")
    Msg = Replace(Msg, "â", "a")
    Msg = Replace(Msg, "ã", "a")
    Msg = Replace(Msg, "à", "a")
    Msg = Replace(Msg, "ä", "a")
    'Limpa letra e
    Msg = Replace(Msg, "é", "e")
    Msg = Replace(Msg, "ê", "e")
    Msg = Replace(Msg, "è", "e")
    Msg = Replace(Msg, "ë", "e")
    'Limpa letra i
    Msg = Replace(Msg, "í", "i")
    Msg = Replace(Msg, "î", "i")
    Msg = Replace(Msg, "ì", "i")
    Msg = Replace(Msg, "ï", "i")
    'Limpa letra o
    Msg = Replace(Msg, "ó", "o")
    Msg = Replace(Msg, "ô", "o")
    Msg = Replace(Msg, "õ", "o")
    Msg = Replace(Msg, "ò", "o")
    Msg = Replace(Msg, "ö", "o")
    'Limpa letra u
    Msg = Replace(Msg, "ú", "u")
    Msg = Replace(Msg, "û", "u")
    Msg = Replace(Msg, "ù", "u")
    Msg = Replace(Msg, "ü", "u")
    'Cedilha
    Msg = Replace(Msg, "ç", "c")
    'Substitui os demais caracteres invalidos por TRAVESSÃO
    For t = 1 To Len(Msg)
        caractere = Mid(Msg, t, 1)
        If InStr(1, "abcdefghijklmnopqrstuvxwyz", caractere) = 0 Then Msg = Replace(Msg, caractere, "_")
    Next t
    Msg = Replace(Msg, "_", "")
    SanSoletras = Msg
End Function
Function SanNormalizeText(frase As String) As String
'***************************************
'FUNÇÃO DE NORMALIZAÇÃO DE TEXTOS
'Substitui caracteres acentuados e separadores
'Usa processo de conversão de STRING para ARRAY e ARRAY para STRING
'***************************************
Static uni_iniciado As Boolean
If uni_iniciado = Empty Then
    GoTo Inicialize_uni
Else
    GoTo execute_uni
End If
Inicialize_uni:
    uni_iniciado = True
    Static lut(256) As Byte
    Static ache As String
    Static subs As String
    Static tamanho As Long
    Static bytArray() As Byte
    'CRIA MODULO PARA BLOQUEAR ACENTUAÇÃO, ALTERNATIVAMENTE
    'ache = "áãâàäåéëêèíìïîóõôòöøüûùúýÿ¥çñ$ðæþÁÃÂÀÄÅÉËÊÈÍÌÏÎÓÕÔÒÖØÜÛÙÚÝŸ¥ÇÑ$ß§ÐÆ£Þ¢¤¶ªº°ƒµ0123456789'.,,+:/)!(&?;»=´´```´“”#®©@_~’\–-²*³%][×<¬>½{}¡·^¿«««|" + Chr(134) + Chr(147) + Chr(148) + Chr(34) + Chr(160) + Chr(133) + Chr(166) + Chr(153) + Chr(150) + Chr(9)
    'subs = "aaaaaaeeeeiiiioooooouuuuyyycnsdepAAAAAAEEEEIIIIOOOOOOUUUUYYYCNSBSDELPCOPAOOFU                                                                                   " + Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(32)
    'Permitir acentuação
    ache = "åøý¥ñðæþÅØÝ¥ÇÑ$ß§ÐÆ£Þ¢¤¶ªº°ƒµ0123456789'.,,+:/)!(&?;»=´´```´“”#®©@_~’\–-²*³%][×<¬>½{}¡·^¿«««|" + Chr(134) + Chr(147) + Chr(148) + Chr(34) + Chr(160) + Chr(133) + Chr(166) + Chr(153) + Chr(150) + Chr(9)
    subs = "aoyyndepAOYYCNSBSDELpcoPaoofu                                                                                   " + Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(32)
    tamanho = Len(ache)
    For s = 0 To 255
        lut(s) = CByte(s)
    Next s
    For s = 1 To tamanho
        lut(Asc(Mid(ache, s, 1))) = CByte(Asc(Mid(subs, s, 1)))
    Next s
execute_uni:
    'Converte codigo de pagina para LATIN1
    If frase = "" Then SanNormalizeText = "": Exit Function
    bytArray = UCase(frase) & Chr(0)
    'Processa o Array
    tamanho = 2 * Len(frase)
    For x = 0 To tamanho Step 2
        bytArray(x) = lut(bytArray(x))
    Next x
    'Converte para texto capitalizado, melhor condição de buscador
    frase = StrConv(bytArray, vbUpperCase)
    iPos = InStr(1, frase, Chr(0))
    If iPos > 0 Then frase = Left(frase, iPos - 1)
    'Excessões estranhas
    If InStr(1, frase, Chr(9)) > 0 Then frase = Replace(frase, Chr(9), " ")
    If InStr(1, frase, Chr(10)) > 0 Then frase = Replace(frase, Chr(10), " ")
    If InStr(1, frase, Chr(13)) > 0 Then frase = Replace(frase, Chr(13), " ")
    If InStr(1, frase, Chr(134)) > 0 Then frase = Replace(frase, Chr(134), " ")
    If InStr(1, frase, Chr(147)) > 0 Then frase = Replace(frase, Chr(147), " ")
    If InStr(1, frase, Chr(148)) > 0 Then frase = Replace(frase, Chr(148), " ")
    If InStr(1, frase, Chr(150)) > 0 Then frase = Replace(frase, Chr(150), " ")
    SanNormalizeText = frase
End Function
Function StrCount(Msg As String, caractere As String) As Long
    Static I As Long
    Static carac As String
    Static sum As Long
    sum = 0
    For I = 1 To Len(Msg) - Len(caractere) + 1
        If Mid(Msg, I, Len(caractere)) = caractere Then sum = sum + 1
    Next I
    StrCount = sum
End Function
Function SanArray2Str(var_Array As Variant, Delimiter As String) As String
Static result As String
result = Join(var_Array, Delimiter)
SanArray2Str = Trim(result)
End Function

Function SanStr2Array(fn_SanStr2Array_String As String, fn_SanStr2Array_Separator) As Variant
    Static fn_SanStr2Array_dummy(0)
    'Verifica se existem condições de particionamento da string
    If InStr(1, fn_SanStr2Array_String, fn_SanStr2Array_Separator) > 0 And Trim(fn_SanStr2Array_String) <> "" Then
        'Retorna Array
        SanStr2Array = Split(fn_SanStr2Array_String, fn_SanStr2Array_Separator)
    Else
        'Retorna Array Vazio
        SanStr2Array = fn_SanStr2Array_dummy
    End If
End Function
'Converte conjunto de caracteres UTF,HTML e JAVA para LATIN1
'Funciona bem PALAVRA POR PALAVRA
Public Function ANSIText(var_ansitext_msg As String) As String
'***************************************
'FUNÇÃO DE CONVERSÃO DE CODIGOS DE CARACTERES
'Substitui tags HTML especiais por caracteres em português ISO8559-1
'***************************************
Static uni_iniciado As Boolean
If uni_iniciado = Empty Then
    GoTo Inicialize_uni
Else
    GoTo execute_uni
End If
Inicialize_uni:
    Static cp(9, 100) As String
    Static fnum As Integer
    Static leitura As String '
    Static var_ansitext_contador As Long '
    Static var_ansitext_contalinha As Long '
    'CARREGA CODIGO DE PAGINA DE ARQUIVO DE DADOS ANEXO latin1.dat
    fnum = FreeFile
    Open App.Path + "\latin1.dat" For Input As #fnum
    var_ansitext_contalinha = 0
    Do
        Line Input #fnum, leitura
        partes = Split(leitura, Chr(9))
        For var_ansitext_contador = 0 To UBound(partes)
            cp(var_ansitext_contador, var_ansitext_contalinha) = partes(var_ansitext_contador)
        Next var_ansitext_contador
        var_ansitext_contalinha = var_ansitext_contalinha + 1
    Loop While Not EOF(fnum)
    Close fnum
execute_uni:
    Static codepage As String
    Static coluna As Long
    Static colunaprocurada As String
    codepage = "Latin1"
    'If InStr(1, var_ansitext_msg, "Indi") > 0 Then MsgBox "AQUI"
    If InStr(1, var_ansitext_msg, "Â") > 0 Or InStr(1, var_ansitext_msg, "Ã") > 0 Then codepage = "UTF-8": GoTo vaidireto
    If InStr(1, var_ansitext_msg, "£") > 0 Or InStr(1, var_ansitext_msg, "¢") > 0 Then codepage = "UTF-7,5": GoTo vaidireto
    If InStr(1, var_ansitext_msg, "+A") > 0 Then codepage = "UTF-7": GoTo vaidireto
    If InStr(1, var_ansitext_msg, "&#") > 0 Then
        codepage = "HTML": GoTo vaidireto
    ElseIf InStr(1, var_ansitext_msg, "&") > 0 Then
        codepage = "EHTML": GoTo vaidireto
    End If
vaidireto:
    If codepage <> "Latin1" Then
        GoSub detecta_coluna
        'Procura e substitui
        For var_ansitext_contador = 1 To 100
            If InStr(1, var_ansitext_msg, cp(colunaprocurada, var_ansitext_contador)) > 0 Then
                var_ansitext_msg = Replace(var_ansitext_msg, cp(colunaprocurada, var_ansitext_contador), cp(0, var_ansitext_contador))
                'Exit For
            End If
        Next var_ansitext_contador
    End If
    ANSIText = var_ansitext_msg
Exit Function
detecta_coluna:
colunaprocurada = -1
For var_ansitext_contador = 0 To 9
    If cp(var_ansitext_contador, 0) = codepage Then colunaprocurada = var_ansitext_contador: Exit For
Next
Return
End Function

'*********************Funções HTML
Function SanHTML2Text(pagina As String) As String
    Static linhas
    Static linha As String
    Static captura As Boolean
    Static tag As String
    Static saida As String
    linhas = Split(pagina, Chr(13) & Chr(10))
    frase = ""
    captura = True
    For aa = 0 To UBound(linhas)
        partes = Split(linhas(aa), "<")
        For ab = 0 To UBound(partes)
            celula = partes(ab)
            posicao = InStr(1, celula, ">")
            If posicao > 0 Then
                tag = UCase(Left(celula, Len(celula) - 1))
                If Left(tag, 6) = "SCRIPT" Or Left(tag, 5) = "STYLE" Or Left(tag, 4) = "<!--" Then
                    captura = False
                End If
                If Left(tag, 7) = "/SCRIPT" Or Left(tag, 6) = "/STYLE" Or Left(tag, 3) = "-->" Then
                    captura = True
                End If
                saida = Trim(Right(celula, Len(celula) - posicao))
                If saida <> "" And captura Then
                    'If InStr(1, saida, "Kapyn") > 0 Then MsgBox "aqui"
                    'saida = Trim(SanNormalizeText(ANSIText(saida)))
                    While InStr(1, saida, "  ") > 0
                        saida = Replace(saida, "  ", " ")
                    Wend
                    'DEBUG
                    If InStr(1, saida, "[") = 0 And saida <> "" Then frase = frase & " " & saida
                End If
            Else
                'algo errado
            End If
        Next ab
    Next aa
    'Consolida erros na pagina
    'Elimina muitas linhas puladas
    While InStr(1, frase, Chr(13) & Chr(10) & Chr(13) & Chr(10)) > 0
        frase = Replace(frase, Chr(13) & Chr(10) & Chr(13) & Chr(10), Chr(13) & Chr(10))
    Wend
    SanHTML2Text = Trim(frase)
End Function
Function SanHTML2Text2(pagina As String) As String
    Static linhas
    Static linha As String
    Static captura As Boolean
    Static tag As String
    Static saida As String
    linhas = Split(pagina, Chr(13) & Chr(10))
    frase = ""
    captura = True
    For aa = 0 To UBound(linhas)
        partes = Split(linhas(aa), "<")
        For ab = 0 To UBound(partes)
            celula = partes(ab)
            posicao = InStr(1, celula, ">")
            If posicao > 0 Then
                tag = UCase(Left(celula, Len(celula) - 1))
                If Left(tag, 6) = "SCRIPT" Or Left(tag, 5) = "STYLE" Or Left(tag, 4) = "<!--" Then
                    captura = False
                End If
                If Left(tag, 7) = "/SCRIPT" Or Left(tag, 6) = "/STYLE" Or Left(tag, 3) = "-->" Then
                    captura = True
                End If
                saida = Trim(Right(celula, Len(celula) - posicao))
                If saida <> "" And captura Then
                    'saida = Trim(SanNormalizeText(ANSIText(saida)))
                    While InStr(1, saida, "  ") > 0
                        saida = Replace(saida, "  ", " ")
                    Wend
                    'DEBUG
                    If InStr(1, saida, "[") = 0 And saida <> "" Then frase = frase & " " & saida
                End If
            Else
                'algo errado
            End If
        Next ab
        DoEvents
    Next aa
    'Consolida erros na pagina
    'Elimina muitas linhas puladas
    While InStr(1, frase, Chr(13) & Chr(10) & Chr(13) & Chr(10)) > 0
        frase = Replace(frase, Chr(13) & Chr(10) & Chr(13) & Chr(10), Chr(13) & Chr(10))
    Wend
    SanHTML2Text2 = Trim(frase)
End Function
'***********ESTATISTICAS DE PALAVRAS
Function SanGetKeywords(ClearText As String, Optional peso As Integer) As String
    Static aa As Long
    Static ab As Long
    Static Keywords As String
    Static buck() As contador
    Static sobenum As Long
    Static rotate1 As String
    Static rotate2 As String
    Static palavra2 As String
    Static palavra3 As String
    Static palavra4 As String
    'Se for 0, reinicia buck, se for >0 faz incremental
    If peso = 0 Then ReDim buck(0)
    'Pensar sobre isso, parece faltar eliminação de ENTERS
    ClearText = Replace(ClearText, Chr(13) & Chr(10), " ")
    palavras = Split(LCase(SanNormalizeText(ClearText)), " ")
    Keywords = ""
    rotate1 = ""
    rotate2 = ""
    rotate3 = ""
    palavra2 = ""
    palavra3 = ""
    palavra4 = ""
    For aa = 0 To UBound(palavras)
        palavra = Trim(palavras(aa))
        '###################################################################################
        'Rotina de criação de EXPRESSÕES
        palavra2 = Trim(rotate1 & " " & palavra)
        palavra3 = Trim(rotate2 & " " & rotate1 & " " & palavra)
        palavra4 = Trim(rotate3 & " " & rotate2 & " " & rotate1 & palavra)
        '###################################################################################
        'Por palavra unica
        If palavra <> "" And Len(palavra) > 1 And Len(palavra2) < 22 And Val(palavra) = 0 And InStr(1, "de|que|os|do|se|para|não|da|em|as|com|dos|ao|por|no|como|ele|seu|ele|seu|um|pois|sua|na|seus|te|eu|mas|me|porque|sobre|lhe|então|todos|disse|meu|também|assim|vos|teu|conosco|menos|fique|alguns|ou|aos|home|proxima|anterior|voltar|ainda|nbsp|das|esta|este|foi|ha|mais|nao|rs|sao|sem|ser|tem|uma", palavra) = 0 Then
            parse = palavra
            If Trim(parse) <> "" Then GoSub buck_add
        End If
        'Por expressão dupla maior que 5 caracteres
        partes = Split(palavra2, " ")
        If UBound(partes) = 1 And Len(palavra2) > 5 And Len(palavra2) < 32 Then
            'Por expressão dupla
            parse = palavra2
            If parse <> "" Then GoSub buck_add
        End If
        'Por expressão tripla maior que 8 caracteres
        partes = Split(palavra3, " ")
        If UBound(partes) = 2 And Len(palavra3) > 8 And Len(palavra3) < 32 Then
            'Por expressão tripla eliminando espaços duplo
            parse = Replace(palavra3, "  ", " ")
            If parse <> "" Then GoSub buck_add
        End If
        'Por expressão tripla maior que 8 caracteres
        partes = Split(palavra4, " ")
        If UBound(partes) = 3 And Len(palavra4) > 12 And Len(palavra4) < 48 Then
            'Por expressão tripla eliminando espaços duplo
            parse = Replace(palavra3, "  ", " ")
            If parse <> "" Then GoSub buck_add
        End If
        'Grava e toda as ultimas posições
        rotate3 = rotate2
        rotate2 = rotate1
        rotate1 = palavra
    Next aa
    'Rotina de ordenação
    Do
        achado = False
        For aa = 1 To UBound(buck)
            If buck(aa - 1).contagem < buck(aa).contagem Then
                sobe = buck(aa - 1).palavra
                buck(aa - 1).palavra = buck(aa).palavra
                buck(aa).palavra = sobe
                sobenum = buck(aa - 1).contagem
                buck(aa - 1).contagem = buck(aa).contagem
                buck(aa).contagem = sobenum
                achado = True
            End If
        Next aa
    Loop While achado
    For aa = 0 To UBound(buck)
        If buck(aa).palavra <> "" Then Keywords = Keywords + Format(buck(aa).contagem, "000000") & ":" & buck(aa).palavra & Chr(13) & Chr(10)
    Next aa
    SanGetKeywords = Keywords
Exit Function
'Incrementar essa função com pesos:
'TITULO -> 3 pontos
'Descrição -> 2 pontos
'Meta Tags -> 2 pontos
'Texto -> 1 ponto
buck_add:
    achado = False
    For ab = 0 To UBound(buck)
        If buck(ab).palavra = parse Then
            If ab > (0 + peso) Then
                buck(ab).contagem = buck(ab).contagem + 1
                sobe = buck(ab - 1 - peso).palavra
                buck(ab - 1 - peso).palavra = buck(ab).palavra
                buck(ab).palavra = sobe
                sobenum = buck(ab - 1 - peso).contagem
                buck(ab - 1 - peso).contagem = buck(ab).contagem
                buck(ab).contagem = sobenum
            End If
            achado = True
            Exit For
        End If
    Next ab
    If Not achado Then
        If UBound(buck) < 2049 Then
            ReDim Preserve buck(UBound(buck) + 1)
            buck(UBound(buck)).palavra = parse
            buck(UBound(buck)).contagem = 1
        Else
            'Se a ultima palavra tiver contagem 1, simplesmente troca
            If buck(UBound(buck)).contagem = 1 Then buck(UBound(buck)).palavra = parse
        End If
    End If
Return
End Function
Function SanResetGKW()
    ReDim GKW(0)
End Function
Sub SanAddGKW(lista As String)
    Static aa As Long
    Static ab As Long
    Static sobenum As Long '
    linhas = Split(lista, Chr(13) & Chr(10))
    For aa = 0 To UBound(linhas)
        partes = Split(linhas(aa), ":")
        If UBound(partes) > 0 Then
            achado = False
            For ab = 0 To UBound(GKW)
                If GKW(ab).palavra = partes(1) Then
                    GKW(ab).contagem = GKW(ab).contagem + Val(partes(0))
                    If ab > 0 Then
                        sobe = GKW(ab - 1).palavra
                        GKW(ab - 1).palavra = GKW(ab).palavra
                        GKW(ab).palavra = sobe
                        sobenum = GKW(ab - 1).contagem
                        GKW(ab - 1).contagem = GKW(ab).contagem
                        GKW(ab).contagem = sobenum
                    End If
                    achado = True
                    Exit For
                End If
            Next ab
            If Not achado Then
                If UBound(GKW) < 2049 And Val(partes(0)) > 1 Then
                    ReDim Preserve GKW(UBound(GKW) + 1)
                    GKW(UBound(GKW)).palavra = partes(1)
                    GKW(UBound(GKW)).contagem = Val(partes(0))
                Else
                    'Se a ultima palavra tiver contagem 2 e o novo dado for maior, troca
                    If GKW(UBound(GKW)).contagem = 2 And partes(0) > 2 Then GKW(UBound(GKW)).palavra = partes(1)
                End If
            End If
            'Adiciona em GKW
        End If
    Next aa
End Sub
Function SanGetGKW() As String
    Static aa As Long
    Static ab As Long
    Static Keywords As String
    Static sobenum As Long '
    Static sum As Long
    Keywords = ""
    SanEntropiaGKW
    'Rotina de ordenação -> COMPARAR ENTROPIAS
    Do
        achado = False
        For aa = 1 To UBound(GKW)
            If GKW(aa - 1).Entropia < GKW(aa).Entropia Then
                'GKW(aa).contagem = GKW(aa).contagem + 1
                sobe = GKW(aa - 1).palavra
                GKW(aa - 1).palavra = GKW(aa).palavra
                GKW(aa).palavra = sobe
                sobenum = GKW(aa - 1).contagem
                GKW(aa - 1).contagem = GKW(aa).contagem
                GKW(aa).contagem = sobenum
                sobenum = GKW(aa - 1).Entropia
                GKW(aa - 1).Entropia = GKW(aa).Entropia
                GKW(aa).Entropia = sobenum
                achado = True
            End If
        Next aa
    Loop While achado
    'Calcula somatória
    sum = 0
    For aa = 0 To UBound(GKW)
        If GKW(aa).palavra <> "" And GKW(aa).contagem > 1 Then
            sum = sum + GKW(aa).Entropia
        End If
    Next aa
    'Calcula porcentagem
    For aa = 0 To UBound(GKW)
        If GKW(aa).palavra <> "" And GKW(aa).contagem > 1 Then
            GKW(aa).Porcentagem = GKW(aa).Entropia / sum
        End If
    Next aa
    'Monta lista para impressão
    For aa = 0 To UBound(GKW)
        If GKW(aa).palavra <> "" And GKW(aa).contagem > 1 Then
            Keywords = Keywords & Format(GKW(aa).Porcentagem, "00.0000%") & Chr(9) & Format(GKW(aa).Entropia, "000000") & Chr(9) & Format(GKW(aa).contagem, "000000") & Chr(9) & GKW(aa).palavra & Chr(13) & Chr(10)
        End If
    Next aa
    SanGetGKW = Keywords
End Function
Sub SanEntropiaGKW()
    'Gera entropias
    For aa = 0 To UBound(GKW)
        GKW(aa).Entropia = Len(GKW(aa).palavra) * GKW(aa).contagem
    Next aa
End Sub
'Retorna true ou falso se a pagina se encaixa no contexto da pesquisa
Function SanCompareGKW(pagina As String) As Boolean
    Static aa As Long
    Static ab As Long
    Static conta_palavras_pagina As Long
    Static conta_achadas_pagina As Long
    'Existem palavras para serem pesquisadas?
    If UBound(GKW) > 0 Then
        pagina = Replace(pagina, Chr(13) & Chr(10), " ")
        If pagina = "" Then SanCompareGKW = False: Exit Function
        palavras = Split(LCase(SanNormalizeText(pagina)), " ")
        conta_palavras_pagina = 0
        conta_achadas_pagina = 0
        For aa = 0 To UBound(palavras)
            palavra = Trim(palavras(aa))
            If palavra <> "" Then
                'Procurar palavras no dicionario
                conta_palavras_pagina = conta_palavras_pagina + 1
                achado = False
                For ab = 1 To UBound(GKW)
                    If palavra = GKW(ab).palavra Then achado = True: Exit For
                Next ab
                If achado Then conta_achadas_pagina = conta_achadas_pagina + 1
            End If
        Next aa
        If conta_palavras_pagina > 0 Then
            'Se o auto-aprendizado tiver menos de 128 palavras (6%), escapa
            If UBound(GKW) < 128 Then SanCompareGKW = True: Exit Function
            ratio = conta_achadas_pagina * conta_achadas_pagina / (conta_palavras_pagina * UBound(GKW))
            'Lei de Pareto
            If ratio > 0.002 Then
                SanCompareGKW = True
            Else
                SanCompareGKW = False
            End If
        Else
            SanCompareGKW = False
        End If
    Else
        'Não existem palavras no dicionario
        SanCompareGKW = True
    End If
End Function
'***********************************************************
'FUNÇÕES DE DAEMONS
'***********************************************************
Function daemon(daem() As Byte, operacao As String, Optional dado As String) As Boolean
    If operacao = "RESET" Then
        Static ender As Integer
        Static modulus As Integer
        Static el
        Static ep
        ReDim daem(0)
        ReDim daem(0 To 41, 0 To 140, 8192)
        Dim t As Long
    End If
    dado = LCase(dado)
    If Len(dado) > 1 Then
        el = InStr(1, "%abcdefghijklmnopqrstuvxwyz0123456789-._@", Left(dado, 1)) - 1
        ep = Len(dado)
        If ep > 140 Then ep = 140
        t = m_CRC.CalculateString(Right(dado, ep - 1))
        ender = CLng(t / 8)
        modulus = CLng(2 ^ (t Mod 8))
        If operacao = "TEST" Then
            daemon = (daem(el, ep, ender) And modulus > 0)
        ElseIf operacao = "ADD" Then
            daem(el, ep, ender) = daem(el, ep, ender) Or modulus
        End If
    End If
End Function
Function DUPLICIDADE(Msg As String)
    If daemon(urldaem, "TEST", Msg) Then
        DUPLICIDADE = ""
    Else
        daemon urldaem, "ADD", Msg
        DUPLICIDADE = Msg
    End If
End Function
'*********************************************************
'CAPTURA E GRAVAÇÃO DE LINKS
'*********************************************************
'VARRE TODAS AS HTML CAPTURADAS NESTE DIRETORIO



'*****************************************************************
'FUNÇÕES AUXILIARES DE CONVERSÃO DE TEXTO
Function LimpaTitulo(Msg)
    titulo = Msg
    'LIMPA TAGS HTML NO TITULO
    If InStr(1, titulo, "<") > 0 Then
        titulo = Replace(titulo, "<b>", "")
        titulo = Replace(titulo, "</b>", "")
        titulo = Replace(titulo, "<", "")
    End If
    If InStr(1, titulo, ">") > 0 Then titulo = Replace(titulo, ">", "")
    If InStr(1, titulo, "&") > 0 Then
        titulo = Replace(titulo, "&amp;", "&")
        titulo = Replace(titulo, "&gt;", ">")
        titulo = Replace(titulo, "&lt;", "<")
        titulo = Replace(titulo, "&quot;", Chr(34))
        titulo = Replace(titulo, "&nbsp;", " ")
    End If
    'CARACTERES ESPECIAIS
    If InStr(1, titulo, "Ã") > 0 Then
        titulo = Replace(titulo, "Ã§", "ç")
        titulo = Replace(titulo, "Ã£", "ã")
        titulo = Replace(titulo, "Ãµ", "õ")
        titulo = Replace(titulo, "Ãª", "ê")
        titulo = Replace(titulo, "Ã³", "ó")
        titulo = Replace(titulo, "Ã´", "ô")
        titulo = Replace(titulo, "Ã¡", "á")
        titulo = Replace(titulo, "Ã©", "é")
        titulo = Replace(titulo, "Ã­", "í")
        titulo = Replace(titulo, "Ãº", "ú")
        titulo = Replace(titulo, "Ã ", "à")
        titulo = Replace(titulo, "Ã¼", "ü")
        titulo = Replace(titulo, "Ã¢", "â")
        titulo = Replace(titulo, "Ã", "Á")
        titulo = Replace(titulo, "Ãš", "Ú")
        titulo = Replace(titulo, "Ã‡", "Ç")
        titulo = Replace(titulo, "Ãƒ", "Ã")
        titulo = Replace(titulo, "Ã", "Í")
        titulo = Replace(titulo, "ÃŠ", "Ê")
        titulo = Replace(titulo, "Ã”", "Ô")
        titulo = Replace(titulo, "Ãº", "º")
        titulo = Replace(titulo, "Ã“", "Ó")
        titulo = Replace(titulo, "Ã®", "®")
        titulo = Replace(titulo, "Â®", "®")
        titulo = Replace(titulo, "Ãª", "ª")
    End If
    If InStr(1, titulo, "â") > 0 Then titulo = Replace(titulo, "â„¢", "™")
    If InStr(1, titulo, "|") > 0 Then titulo = Replace(titulo, "|", " ")
    If InStr(1, titulo, "Â") > 0 Then
        titulo = Replace(titulo, "Âª", "ª")
        titulo = Replace(titulo, "Âº", "º")
        titulo = Replace(titulo, "Â·", "·")
        titulo = Replace(titulo, "Â°", "º")
    End If
    LimpaTitulo = titulo
End Function
Function fnCheckFileName(fname As String)
    Static result
    fnCheckFileName = True
    If InStr(1, result, "\") > 0 Then fnCheckFileName = False: Exit Function
    If InStr(1, result, "/") > 0 Then fnCheckFileName = False: Exit Function
    If InStr(1, result, ":") > 0 Then fnCheckFileName = False: Exit Function
    If InStr(1, result, "*") > 0 Then fnCheckFileName = False: Exit Function
    If InStr(1, result, "?") > 0 Then fnCheckFileName = False: Exit Function
    If InStr(1, result, "<") > 0 Then fnCheckFileName = False: Exit Function
    If InStr(1, result, ">") > 0 Then fnCheckFileName = False: Exit Function
    If InStr(1, result, "|") > 0 Then fnCheckFileName = False: Exit Function
    If InStr(1, result, Chr(34)) > 0 Then fnCheckFileName = False: Exit Function
End Function
Function URL2FILE(URL)
    salvarcomo = URL
    salvarcomo = Replace(salvarcomo, "\", "-")
    salvarcomo = Replace(salvarcomo, "/", "-")
    salvarcomo = Replace(salvarcomo, ":", "-")
    salvarcomo = Replace(salvarcomo, "*", "-")
    salvarcomo = Replace(salvarcomo, "?", "-")
    salvarcomo = Replace(salvarcomo, Chr(34), "-")
    salvarcomo = Replace(salvarcomo, "<", "-")
    salvarcomo = Replace(salvarcomo, ">", "-")
    salvarcomo = Replace(salvarcomo, "|", "-")
    URL2FILE = salvarcomo
End Function
'cad clienbtes ---------------------

Public Sub SelecionaTexto(txtObjeto As TextBox)

  With txtObjeto
  .SelStart = 0
  .SelLength = Len(.Text)
  End With

End Sub

Public Sub SoNumeros(Tecla As Integer)
  If Not IsNumeric(Chr(Tecla)) And Tecla <> vbKeyBack Then
    Tecla = 0
  End If
End Sub

Public Function ValidaRG(ByVal RG As String) As Boolean
  Dim I As Long
  Dim caractere As String
  NovoRG = ""

  For I = 1 To Len(RG)
    caractere = Mid(RG, I, 1)
      
    If IsNumeric(caractere) Then
      NovoRG = NovoRG & caractere
    End If
  
  Next
    
    If Len(NovoRG) = 8 Or Len(NovoRG) = 9 Or Len(NovoRG) = 10 Then
        ValidaRG = True
    End If

End Function

Public Function ValidaNascimento(ByVal Nascimento As TextBox) As Boolean

  Dim I As Long
  Dim caractere As String
  NovoNascimento = ""
    
  For I = 1 To Len(Nascimento)
    caractere = Mid(Nascimento, I, 1)
      
    If IsNumeric(caractere) Then
     NovoNascimento = NovoNascimento & caractere
    End If
  
  Next
    
  If Len(NovoNascimento) = 8 Then
    ValidaNascimento = True
  End If
    
End Function

Public Function ValidaCEP(ByVal CEP As String) As Boolean

  Dim I As Long
  Dim caractere As String
  Dim Novocep As String
   
  For I = 1 To Len(CEP)
    caractere = Mid(CEP, I, 1)
      
    If IsNumeric(caractere) Then
      Novocep = Novocep & caractere
    End If
  
  Next
      
    If Len(Novocep) = 8 Then
      ValidaCEP = True
    End If

End Function
Public Function ValidEMail(sEMail As String) As Boolean
  Dim nCharacter As Integer
  Dim Count As Integer
  Dim sLetra As String
  If Len(sEMail) < 5 Then
    ValidEMail = False
    Exit Function
  End If
  For nCharacter = 1 To Len(sEMail)
    If Mid(sEMail, nCharacter, 1) = "@" Then
      Count = Count + 1
    End If
  Next
  If Count <> 1 Then
    ValidEMail = False
    Exit Function
  Else
    If InStr(sEMail, "@") = 1 Then
      ValidEMail = False
      Exit Function
    ElseIf InStr(sEMail, "@") = Len(sEMail) Then
      ValidEMail = False
      Exit Function
    End If
  End If
  nCharacter = 0
  Count = 0
  For nCharacter = 1 To Len(sEMail)
    If Mid(sEMail, nCharacter, 1) = "." Then
      Count = Count + 1
    End If
  Next
  If Count < 1 Then
    ValidEMail = False
    Exit Function
  Else
    If InStr(sEMail, ".") = 1 Then
      ValidEMail = False
      Exit Function
    ElseIf InStr(sEMail, ".") = Len(sEMail) Then
      ValidEMail = False
      Exit Function
    ElseIf InStr(InStr(sEMail, "@"), sEMail, ".") = 0 Then
      ValidEMail = False
      Exit Function
    End If
  End If
  nCharacter = 0
  Count = 0
  If InStr(sEMail, "..") > InStr(sEMail, "@") Then
    ValidEMail = False
    Exit Function
  End If
  For nCharacter = 1 To Len(sEMail)
    sLetra = Mid$(sEMail, nCharacter, 1)
    If Not (LCase(sLetra) Like "[a-z]" Or sLetra = _
          "@" Or sLetra = "." Or sLetra = "-" Or _
          sLetra = "_" Or IsNumeric(sLetra)) Then
      ValidEMail = False
      Exit Function
    End If
  Next
  nCharacter = 0
  ValidEMail = True
End Function
Public Function FU_ValidaCPF(cpf As String) As Boolean
Dim soma As Integer
Dim Resto As Integer
Dim I As Integer
'Valida argumento
    If Len(cpf) <> 11 Then
    FU_ValidaCPF = False
    Exit Function
    End If
  soma = 0
  For I = 1 To 9
    soma = soma + Val(Mid$(cpf, I, 1)) * (11 - I)
  Next I
  Resto = 11 - (soma - (Int(soma / 11) * 11))
  If Resto = 10 Or Resto = 11 Then Resto = 0
  If Resto <> Val(Mid$(cpf, 10, 1)) Then
    FU_ValidaCPF = False
    Exit Function
  End If
  soma = 0
  For I = 1 To 10
    soma = soma + Val(Mid$(cpf, I, 1)) * (12 - I)
  Next I
  Resto = 11 - (soma - (Int(soma / 11) * 11))
  If Resto = 10 Or Resto = 11 Then Resto = 0
  If Resto <> Val(Mid$(cpf, 11, 1)) Then
    FU_ValidaCPF = False
    Exit Function
  End If
  FU_ValidaCPF = True
End Function
Public Function Online() As Boolean
Online = InternetGetConnectedState(0&, 0&)
End Function


Public Function replacetag(inputstr As String)
                        texto = Replace(texto, "Preço:", " ")
                        texto = Replace(texto, "Bairro:", " ", , , vbTextCompare)
                        texto = Replace(texto, "Cidade:", " ", , , vbTextCompare)
                        texto = Replace(texto, "Estado:", " ", , , vbTextCompare)
                        texto = Replace(texto, "São Paulo", " ", , , vbTextCompare)
                        texto = Replace(texto, "SP", " ")
                        texto = Replace(texto, "Área útil:", " ", , , vbTextCompare)
                        texto = Replace(texto, "Vago: Sim", " ", , , vbTextCompare)
                        texto = Replace(texto, "Vago: Não", " ", , , vbTextCompare)
                        texto = Replace(texto, "CEP:", " ", , , vbTextCompare)
                        texto = Replace(texto, "-", " ", , , vbTextCompare)
                        texto = Replace(texto, "(", " ", , , vbTextCompare)
                        texto = Replace(texto, ")", " ", , , vbTextCompare)
                        'texto = Replace(texto, ",00", " ", , , vbTextCompare)
                        texto = Replace(texto, ",", " ", , , vbTextCompare)
                        'texto = Replace(texto, ".", " ", , , vbTextCompare)
                        'CONTRAI NUMERAIS
                        texto = Replace(texto, "uma ", "1", , , vbTextCompare)
                        texto = Replace(texto, "um ", "1", , , vbTextCompare)
                        texto = Replace(texto, "duas", "2", , , vbTextCompare)
                        texto = Replace(texto, "dois", "2", , , vbTextCompare)
                        texto = Replace(texto, "tres", "3", , , vbTextCompare)
                        texto = Replace(texto, "três", "3", , , vbTextCompare)
                        texto = Replace(texto, "meio", "1/2", , , vbTextCompare)
                        texto = Replace(texto, "meia", "1/2", , , vbTextCompare)
                        'REMOVE ADJETIVOS
                        texto = Replace(texto, "bom", " ", , , vbTextCompare)
                        texto = Replace(texto, "ampla", " ", , , vbTextCompare)
                        texto = Replace(texto, "amplo", " ", , , vbTextCompare)
                        texto = Replace(texto, "bem", " ", , , vbTextCompare)
                        texto = Replace(texto, "grandes", " ", , , vbTextCompare)
                        texto = Replace(texto, "lindos", " ", , , vbTextCompare)
                        texto = Replace(texto, "lindas", " ", , , vbTextCompare)
                        texto = Replace(texto, "lindo", " ", , , vbTextCompare)
                        texto = Replace(texto, "linda", " ", , , vbTextCompare)
                        texto = Replace(texto, "boa", " ", , , vbTextCompare)
                        texto = Replace(texto, "muito", " ", , , vbTextCompare)
                        texto = Replace(texto, "também", " ", , , vbTextCompare)
                        texto = Replace(texto, "grande", " ", , , vbTextCompare)
                        texto = Replace(texto, " e ", " ", , , vbTextCompare)
                        texto = Replace(texto, " é ", " ", , , vbTextCompare)
                        texto = Replace(texto, "excelente", " ", , , vbTextCompare)
                        texto = Replace(texto, "totalmente", " ", , , vbTextCompare)
                        texto = Replace(texto, "otimo", " ", , , vbTextCompare)
                        texto = Replace(texto, "otima", " ", , , vbTextCompare)
                        texto = Replace(texto, "ótimo", " ", , , vbTextCompare)
                        texto = Replace(texto, "ótima", " ", , , vbTextCompare)
                        'RESUME O TEXTO
                        texto = Replace(texto, "acabamento", "Acab", , , vbTextCompare)
                        texto = Replace(texto, "Travessa", "Trav", , , vbTextCompare)
                        texto = Replace(texto, "pavimento", "Pav", , , vbTextCompare)
                        texto = Replace(texto, "terreno", "Terr", , , vbTextCompare)
                        texto = Replace(texto, "coberta", "Cobrt", , , vbTextCompare)
                        texto = Replace(texto, "andares", "And", , , vbTextCompare)
                        texto = Replace(texto, "andar", "And", , , vbTextCompare)
                        texto = Replace(texto, "comercial", "Cml", , , vbTextCompare)
                        texto = Replace(texto, "comércio", "Cml", , , vbTextCompare)
                        texto = Replace(texto, "aceitamos", "Ac", , , vbTextCompare)
                        texto = Replace(texto, "aceito", "Ac", , , vbTextCompare)
                        texto = Replace(texto, "parcelamento", "Parc", , , vbTextCompare)
                        texto = Replace(texto, "parcelo", "Parc", , , vbTextCompare)
                        texto = Replace(texto, "parcela", "Parc", , , vbTextCompare)
                        texto = Replace(texto, "infantil", "Inf", , , vbTextCompare)
                        texto = Replace(texto, "quintal", "Qtl", , , vbTextCompare)
                        texto = Replace(texto, "lavanderia", "Lav", , , vbTextCompare)
                        texto = Replace(texto, "empregada", "Emp", , , vbTextCompare)
                        texto = Replace(texto, "vendo", "Vdo", , , vbTextCompare)
                        texto = Replace(texto, "entrada", "Entr", , , vbTextCompare)
                        texto = Replace(texto, "independente", "Indep", , , vbTextCompare)
                        texto = Replace(texto, "quartos", "Qto", , , vbTextCompare)
                        texto = Replace(texto, "quarto", "Qto", , , vbTextCompare)
                        texto = Replace(texto, "varanda", "Var", , , vbTextCompare)
                        texto = Replace(texto, "esquina", "Esq", , , vbTextCompare)
                        texto = Replace(texto, "festas", "Fest", , , vbTextCompare)
                        texto = Replace(texto, "festa", "Fest", , , vbTextCompare)
                        texto = Replace(texto, "dependência", "Dep", , , vbTextCompare)
                        texto = Replace(texto, "dependencia", "Dep", , , vbTextCompare)
                        texto = Replace(texto, "Apartamento", "Ap ", , , vbTextCompare)
                        texto = Replace(texto, "Venda", "Vd ", , , vbTextCompare)
                        texto = Replace(texto, "Comprar", "Cp ", , , vbTextCompare)
                        texto = Replace(texto, "Alugar", "Alu ", , , vbTextCompare)
                        texto = Replace(texto, "TROCAR", "Tcr ", , , vbTextCompare)
                        texto = Replace(texto, "TROCA", "Tct ", , , vbTextCompare)
                        texto = Replace(texto, "Rua", "R ", , , vbTextCompare)
                        texto = Replace(texto, "Apartamento.", "Ap ", , , vbTextCompare)
                        texto = Replace(texto, "Apto", "Ap ", , , vbTextCompare)
                        texto = Replace(texto, "Apto.", "Ap ", , , vbTextCompare)
                        texto = Replace(texto, "Apt.", "Ap ", , , vbTextCompare)
                        texto = Replace(texto, "Bloco", "Bl ", , , vbTextCompare)
                        texto = Replace(texto, "Vila", "Vl ", , , vbTextCompare)
                        texto = Replace(texto, "Dormitório s ", "Dorm ", , , vbTextCompare)
                        texto = Replace(texto, "Dormitório", "Dorm ", , , vbTextCompare)
                        texto = Replace(texto, "Vaga s  de Garagem", "Vg Grg ", , , vbTextCompare)
                        texto = Replace(texto, "Cozinha", "Coz ", , , vbTextCompare)
                        texto = Replace(texto, "Jardim", "Jd ", , , vbTextCompare)
                        texto = Replace(texto, "Condomínio Fechado", "Cond Fech", , , vbTextCompare)
                        texto = Replace(texto, "Condominio", "Cond ", , , vbTextCompare)
                        texto = Replace(texto, "Sobrado", "Sobr ", , , vbTextCompare)
                        texto = Replace(texto, "dormitorio", "Dorm ", , , vbTextCompare)
                        texto = Replace(texto, "Casas", "Cs ", , , vbTextCompare)
                        texto = Replace(texto, "Casa", "Cs ", , , vbTextCompare)
                        texto = Replace(texto, "Térrea", "Tr ", , , vbTextCompare)
                        texto = Replace(texto, "Terrea", "Tr ", , , vbTextCompare)
                        texto = Replace(texto, "salão", "Sl ", , , vbTextCompare)
                        texto = Replace(texto, "Sala ", "Sl ", , , vbTextCompare)
                        texto = Replace(texto, "Locação ", "Alu ", , , vbTextCompare)
                        texto = Replace(texto, "Aproximadamente ", "Aprox ", , , vbTextCompare)
                        texto = Replace(texto, "Próximo", "Prox ", , , vbTextCompare)
                        texto = Replace(texto, "Proximo", "Prox ", , , vbTextCompare)
                        texto = Replace(texto, "Piscina", "Pisc ", , , vbTextCompare)
                        texto = Replace(texto, "salao", "Sl ", , , vbTextCompare)
                        texto = Replace(texto, "Ponto", "Pt ", , , vbTextCompare)
                        texto = Replace(texto, "Zona Norte", "ZN ", , , vbTextCompare)
                        texto = Replace(texto, "Zona Sul", "ZS ", , , vbTextCompare)
                        texto = Replace(texto, "Zona Leste", "ZL ", , , vbTextCompare)
                        texto = Replace(texto, "Zona Oeste", "ZO ", , , vbTextCompare)
                        texto = Replace(texto, "Avenida", "Av ", , , vbTextCompare)
                        texto = Replace(texto, "Vila", "Vl ", , , vbTextCompare)
                        texto = Replace(texto, "Suites", "Sui ", , , vbTextCompare)
                        texto = Replace(texto, "Suite", "Sui ", , , vbTextCompare)
                        texto = Replace(texto, "Serviço", "Svc ", , , vbTextCompare)
                        texto = Replace(texto, "Banheiro", "Banh ", , , vbTextCompare)
                        texto = Replace(texto, "Quadra", "Qd ", , , vbTextCompare)
                        texto = Replace(texto, "Suíte s ", "Sui ", , , vbTextCompare)
                        texto = Replace(texto, "imóvel", "Imv ", , , vbTextCompare)
                        texto = Replace(texto, "imovel", "Imv ", , , vbTextCompare)
                        texto = Replace(texto, "garagem", "Grg ", , , vbTextCompare)
                        texto = Replace(texto, "garagen", "Grg ", , , vbTextCompare)
                        texto = Replace(texto, "fechada", "Fech ", , , vbTextCompare)
                        texto = Replace(texto, "carro", "Car ", , , vbTextCompare)
                        texto = Replace(texto, "carros", "Car ", , , vbTextCompare)
                        texto = Replace(texto, "area", "Ar ", , , vbTextCompare)
                        texto = Replace(texto, "área", "Ar ", , , vbTextCompare)
                        texto = Replace(texto, "àrea", "Ar ", , , vbTextCompare)
                        texto = Replace(texto, "aréa", "Ar ", , , vbTextCompare)
                        texto = Replace(texto, "areia", "Ar ", , , vbTextCompare)
                        texto = Replace(texto, "lazer", "Laz", , , vbTextCompare)
                        texto = Replace(texto, "churrasqueira", "Churr ", , , vbTextCompare)
                        texto = Replace(texto, "ambientes", "Amb ", , , vbTextCompare)
                        texto = Replace(texto, "ambiente", "Amb ", , , vbTextCompare)
                        texto = Replace(texto, "suítes", "Sui ", , , vbTextCompare)
                        texto = Replace(texto, "fundo", "Fdo ", , , vbTextCompare)
                        texto = Replace(texto, "vende", "Vd ", , , vbTextCompare)
                        texto = Replace(texto, "localizacao", "Loc ", , , vbTextCompare)
                        texto = Replace(texto, "localização", "Loc ", , , vbTextCompare)
                        texto = Replace(texto, "local", "Loc ", , , vbTextCompare)
                        texto = Replace(texto, "visitação", "Vis ", , , vbTextCompare)
                        texto = Replace(texto, "visitas", "Vis ", , , vbTextCompare)
                        texto = Replace(texto, "visita", "Vis ", , , vbTextCompare)
                        texto = Replace(texto, "qualquer", "Qquer ", , , vbTextCompare)
                        texto = Replace(texto, "proprietários", "Prop ", , , vbTextCompare)
                        texto = Replace(texto, "proprietarios", "Prop ", , , vbTextCompare)
                        texto = Replace(texto, "proprietario", "Prop ", , , vbTextCompare)
                        texto = Replace(texto, "proprietário", "Prop ", , , vbTextCompare)
                        texto = Replace(texto, "entre em contato", "Fone ", , , vbTextCompare)
                        texto = Replace(texto, "para", "p/ ", , , vbTextCompare)
                        texto = Replace(texto, "com ", "c/", , , vbTextCompare)
                        'ULTIMAS
                            texto = Replace(texto, " s ", " ", , , vbTextCompare)

                        'ELIMINA ESPAÇOS DESNECESSARIOS
                        While InStr(1, texto, "  ")
                            texto = Replace(texto, "  ", " ", , , vbTextCompare)
                        
                        Wend
                    'Coloca numeros em bold
                    For z = 0 To 9
                        texto = Replace(texto, CStr(z), "<b>" & CStr(z) & "</b>")
                    Next z
End Function

 

