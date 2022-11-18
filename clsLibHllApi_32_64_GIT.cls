VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLibhllapi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------

'Classe para automação do terminal PW3270 usando a libhllapi.dll

'Dúvidas/Sugestões: Erick Lopes de Souza - sodapraia@hotmail.com

'------------------------------------------------------

'Se Office for 64 bits...
#If Win64 Then
    Private Const pwPath = "C:\Program Files\PW3270\libhllapi64.dll"
    Private Declare PtrSafe Function hllapi_init Lib "libhllapi64.dll" (ByVal tp As String) As Long
    Private Declare PtrSafe Function hllapi_deinit Lib "libhllapi64.dll" () As Long
    Private Declare PtrSafe Function hllapi_get_revision Lib "libhllapi64.dll" () As Long
    Private Declare PtrSafe Function hllapi_connect Lib "libhllapi64.dll" (ByVal uri As String, ByVal wait As Integer) As Long
    Private Declare PtrSafe Function hllapi_disconnect Lib "libhllapi64.dll" () As Long
    Private Declare PtrSafe Function hllapi_wait_for_ready Lib "libhllapi64.dll" (ByVal timeOut As Integer) As Long
    Private Declare PtrSafe Function hllapi_get_screen_at Lib "libhllapi64.dll" (ByVal row As Integer, ByVal col As Integer, ByVal text As String) As Long
    Private Declare PtrSafe Function hllapi_enter Lib "libhllapi64.dll" () As Long
    Private Declare PtrSafe Function hllapi_get_message_id Lib "libhllapi64.dll" () As Long
    Private Declare PtrSafe Function hllapi_set_text_at Lib "libhllapi64.dll" (ByVal row As Integer, ByVal col As Integer, ByVal text As String) As Long
    Private Declare PtrSafe Function hllapi_wait Lib "libhllapi64.dll" (ByVal timeOut As Integer) As Long
    Private Declare PtrSafe Function hllapi_pfkey Lib "libhllapi64.dll" (ByVal keycode As Integer) As Long
    Private Declare PtrSafe Function hllapi_pakey Lib "libhllapi64.dll" (ByVal keycode As Integer) As Long
    Private Declare PtrSafe Function hllapi_cmp_text_at Lib "libhllapi64.dll" (ByVal row As Integer, ByVal col As Integer, ByVal text As String) As Long
    Private Declare PtrSafe Function hllapi_is_connected Lib "libhllapi64.dll" () As Long
    Private Declare PtrSafe Function hllapi_set_unlock_delay Lib "libhllapi64.dll" (ByVal delay As Integer) As Long
    Private Declare PtrSafe Function hllapi_get_cursor_address Lib "libhllapi64.dll" () As Long
    Private Declare PtrSafe Function hllapi_set_cursor_address Lib "libhllapi64.dll" (ByVal addr As Integer) As Long
    Private Declare PtrSafe Function hllapi_action Lib "libhllapi64.dll" (ByVal keyname As String) As Long
    Private Declare PtrSafe Function hllapi_find_text Lib "libhllapi64.dll" (ByVal text As String) As Long
    Private Declare PtrSafe Function hllapi_set_session_parameter Lib "libhllapi64.dll" (ByVal text As String, ByVal size As Integer, ByVal value As Integer) As Long
#Else
    'Se office for 32 bits
    Private Const pwPath = "C:\Program Files\PW3270\libhllapi32.dll"
    Private Declare Function hllapi_init Lib "libhllapi32.dll" (ByVal tp As String) As Long
    Private Declare Function hllapi_deinit Lib "libhllapi32.dll" () As Long
    Private Declare Function hllapi_get_revision Lib "libhllapi32.dll" () As Long
    Private Declare Function hllapi_connect Lib "libhllapi32.dll" (ByVal uri As String, ByVal wait As Integer) As Long
    Private Declare Function hllapi_disconnect Lib "libhllapi32.dll" () As Long
    Private Declare Function hllapi_wait_for_ready Lib "libhllapi32.dll" (ByVal timeOut As Integer) As Long
    Private Declare Function hllapi_get_screen_at Lib "libhllapi32.dll" (ByVal row As Integer, ByVal col As Integer, ByVal text As String) As Long
    Private Declare Function hllapi_enter Lib "libhllapi32.dll" () As Long
    Private Declare Function hllapi_get_message_id Lib "libhllapi32.dll" () As Long
    Private Declare Function hllapi_set_text_at Lib "libhllapi32.dll" (ByVal row As Integer, ByVal col As Integer, ByVal text As String) As Long
    Private Declare Function hllapi_wait Lib "libhllapi32.dll" (ByVal timeOut As Integer) As Long
    Private Declare Function hllapi_pfkey Lib "libhllapi32.dll" (ByVal keycode As Integer) As Long
    Private Declare Function hllapi_pakey Lib "libhllapi32.dll" (ByVal keycode As Integer) As Long
    Private Declare Function hllapi_cmp_text_at Lib "libhllapi32.dll" (ByVal row As Integer, ByVal col As Integer, ByVal text As String) As Long
    Private Declare Function hllapi_is_connected Lib "libhllapi32.dll" () As Long
    Private Declare Function hllapi_set_unlock_delay Lib "libhllapi32.dll" (ByVal delay As Integer) As Long
    Private Declare Function hllapi_get_cursor_address Lib "libhllapi32.dll" () As Long
    Private Declare Function hllapi_set_cursor_address Lib "libhllapi32.dll" (ByVal addr As Integer) As Long
    Private Declare Function hllapi_action Lib "libhllapi32.dll" (ByVal keyname As String) As Long
    Private Declare Function hllapi_find_text Lib "libhllapi32.dll" (ByVal text As String) As Long
    Private Declare Function hllapi_set_session_parameter Lib "libhllapi32.dll" (ByVal text As String, ByVal size As Integer, ByVal value As Integer) As Long
#End If


Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare PtrSafe Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function GetWindowText Lib "USER32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function GetWindowTextLength Lib "USER32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function GetWindow Lib "USER32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "USER32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const GW_HWNDNEXT = 2
Dim WinTitle As String
Dim FullTitle As String
Dim pw As New clsLibhllapi


Public Function Connect(Optional host As String) As Boolean

    Dim hostStr() As Variant
    
    'Carregando a DLL no local indicado
    LoadLibrary (pwPath)
    
    'Inicializando conexão com o nome do host já aberto
    If host <> "" Then
        If hllapi_init(host) <> 0 Then
            MsgBox ("Erro ao conectar! Verifique se a sessão " & host & " do Terminal está aberto!")
            Connect = False
        Else
            hllapi_set_unlock_delay 1
            Connect = True
        End If
    Else
        hostStr() = Array("pw3270:A", "pw3270:B", "pw3270:C", "pw3270:D")
        For I = 0 To 3
            If hllapi_init(hostStr(I)) <> 0 Then
                If I >= 3 Then
                    MsgBox ("Tentativas de conexão esgotadas! Verifique se há alguma sessão do Terminal aberta!")
                    Connect = False
                End If
            Else
                hllapi_set_unlock_delay 1
                Connect = True
                Exit Function
            End If
        Next
    End If
    
End Function

Public Sub Disconnect()

    'Desconectando...
    hllapi_deinit

End Sub

Public Sub WaitHost(ByVal tempo As Integer)
    
    'Pausa a execução do código
    hllapi_wait tempo

End Sub

Public Function WaitHostOK(ByVal tempo As Integer)

    'Aguardar host ate que esteja ok
    hllapi_wait_for_ready tempo
    hllapi_wait 0.1
    While Status <> 0: DoEvents: Wend

End Function

Public Function GetString(ByVal LINHA As Integer, ByVal coluna As Integer, tamanho As Long)

    'Definindo o tamanho da string
    Dim txt As String

    txt = Space(tamanho)
    hllapi_get_screen_at LINHA, coluna, txt
    GetString = txt

End Function

Public Sub Enter()

    'Nao precisa explicar... ;)
    hllapi_enter
    While Status <> 0: DoEvents: Wend

End Sub

Public Function PutString(LINHA As Integer, coluna As Integer, texto As String)

    'Inserir texto na posicao desejada

    hllapi_set_text_at CInt(LINHA), CInt(coluna), texto
    While Status <> 0: DoEvents: Wend

End Function

Public Sub SendPFKey(ByVal numero As Integer)

    'Envia tecla de funcao (F1, F8, etc)
    hllapi_pfkey numero
    While Status <> 0: DoEvents: Wend

End Sub

Public Sub SendPAKey(ByVal codigo As Integer)

    'Envia teclas de acordo com a tabela ASCII
    hllapi_pakey codigo
    While Status <> 0: DoEvents: Wend

End Sub

Public Function SendEspKey(keyname As String)

    Dim newKey As String
    'Home = firstfield
    'End = fieldend

    Select Case keyname
        Case "HOME"
            newKey = "firstfield"
        Case "END"
            newKey = "fieldend"
        Case "TAB"
            newKey = "nextfield"
    End Select

    hllapi_action newKey
    While Status <> 0: DoEvents: Wend

End Function

Public Function FindText(txt As String)

    'Localizar um texto na tela. Se encontrar, <> 0
    FindText = hllapi_find_text(txt)

End Function

Public Function SetCursor(LINHA As Integer, col As Integer)

    'Posicao do cursor:  A  /  B
    '                   008 / 003 -> ((A - 1)*80 + B) -> 563
    Dim addr As Integer

    addr = ((LINHA - 1) * 80) + col
    hllapi_set_cursor_address addr
    While Status <> 0: DoEvents: Wend

End Function

Public Function SetCursorADDR(addr As Integer)

    'Move o cursor para um endereço específico, como por exemplo, o retorno do GetCursor (925)
    hllapi_set_cursor_address addr
    While Status <> 0: DoEvents: Wend

End Function

Public Function GetCursor()

    GetCursor = hllapi_get_cursor_address

End Function

Public Function WaitForCursor(LINHA As Integer, coluna As Integer) As Boolean

    'Verifica se o cursor esta em determinada posicao, retornando true
    Dim addr As Integer, cmpAddr As Integer

    addr = hllapi_get_cursor_address

    cmpAddr = ((coluna - 1) * 80) + LINHA

    If addr = cmpAddr Then
        WaitForCursor = True
    Else
        WaitForCursor = False
    End If

End Function

Public Function GetScreen(Optional start_row As Integer = 1, Optional end_row As Integer = 24)

    'Função para pegar a tela do host com parâmetros opcionais

    'Ex.: GetScreen() pega todo o conteúdo da tela
    '     GetScreen(10) pega o conteúdo da tela a partir da linha 10 até a última linha
    '     GetScreen(20,22) pega o conteúdo da tela da linha 20 até a linha 22

    'Obs: O HOST SÓ TEM 24 LINHAS... ;)

    Dim txt As String
    Dim resTxt As String
    Dim I As Integer

    txt = Space(80)

    For I = start_row To end_row
        hllapi_get_screen_at I, 1, txt
        resTxt = resTxt & txt & vbCrLf
    Next

    GetScreen = resTxt
    While Status <> 0: DoEvents: Wend

End Function

Public Sub ChangeParam(ByVal name As String, ByVal size As Integer, ByVal value As Integer)

    hllapi_set_session_parameter name, size, value

End Sub

Public Function Status()

    'Status do Host, "0" se ok
    Status = hllapi_get_message_id()

End Function

Public Function NewNamedSess(sessName As String, Optional hidden As Boolean = True)

'CRIA UMA SESSÃO COM UM NOME ESPECÍFICO E ESCONDE A JANELA

Dim wh As Long
Dim oShell: Set oShell = CreateObject("WScript.Shell")
Dim seCs, timeOut As Integer

'Cria uma nova sessão com o nome desejado
Shell "C:\Arquivos de Programas\pw3270\pw3270.EXE --session " & sessName

'Define tempo para timeout
timeOut = 35 'segundos aprox.

'Loop de tempo
For I = 1 To timeOut
    
    'Verifica se o Rede Caixa foi autorizado e está aberta.
    If GetHndPartCaption(wh, sessName) = True Then
        If hidden Then
            ShowWindow wh, 0
        End If
        hllapi_connect WinTitle, 0
        hllapi_wait 1
        Exit For
    Else
        Sleep 1000
        DoEvents
        seCs = seCs + 1
    End If

Next

'Verifica se o tempo de conexão foi excedido
If seCs >= timeOut Then
    MsgBox "Tempo esgotado para autorização do terminal! Reinicie a macro!"
    Exit Function
End If

End Function

Public Sub DisconnectNamedSess(sessName As String)

Dim wh As Long
Dim oShell: Set oShell = CreateObject("WScript.Shell")
Dim seCs, timeOut As Integer


'Disconectar
hllapi_deinit

'Define tempo para timeout
timeOut = 35 'segundos aprox.

'Loop de tempo
For I = 1 To timeOut
    
    'Verifica se o Rede Caixa foi autorizado e está aberta.
    If GetHndPartCaption(wh, sessName) = True Then
        ShowWindow wh, 1
        pw.WaitHost 1
        Exit For
    Else
        Sleep 1000
        DoEvents
        seCs = seCs + 1
    End If

Next

'Verifica se o tempo de conexão foi excedido
If seCs >= timeOut Then
    MsgBox "Tempo esgotado para autorização do terminal! Reinicie a macro!"
    Exit Sub
End If

'Elimina a janela aberta
oShell.Run "taskkill /F /FI ""WINDOWTITLE eq " & FullTitle & """ /T"

End Sub

'Função para pegar janela por nome parcial
Private Function GetHndPartCaption(ByRef lWnd As Long, ByVal sCaption As String) As Boolean

    Dim lhWndP As Long
    Dim sStr As String
    
    GetHndPartCaption = False
    
    lhWndP = FindWindow(vbNullString, vbNullString)
    Do While lhWndP <> 0
        sStr = String(GetWindowTextLength(lhWndP) + 1, Chr$(0))
        GetWindowText lhWndP, sStr, Len(sStr)
        sStr = Left$(sStr, Len(sStr) - 1)
        If InStr(1, sStr, sCaption) > 0 Then
            GetHndPartCaption = True
            lWnd = lhWndP
            WinTitle = Mid(sStr, InStr(1, sStr, sCaption), Len(sCaption) + 2)
            FullTitle = sStr
            Exit Do
        End If
        lhWndP = GetWindow(lhWndP, GW_HWNDNEXT)
    Loop
    
End Function

Public Sub SendKeys(Str As Variant)

Dim mPos As Integer
Dim line, col, charIniPos, charFinPos, I As Integer
Dim cMds(0 To 255) As Variant
Dim espK, tmp, reSt, strCmd As Variant

'Populando array de Esp Keys e Enter
espK = "<ENTER><HOME><END><TAB>"

'Posição atual do cursor
mPos = GetCursor

'Gerando linha e coluna
line = 1 + Int(mPos / 80)
col = mPos Mod 80

'Corrigindo erro da última coluna
If col = 0 Then
    line = line - 1
    col = 80
End If
    
'Separando os comandos
I = 0
While Len(Str) <> 0

    charIniPos = InStr(1, Str, "<")
    charFinPos = InStr(1, Str, ">")
    
    If charIniPos <> 0 Then
    
        If charFinPos <> 0 Then
            tmp = Mid(Str, charIniPos, (charFinPos - charIniPos) + 1)
            reSt = Mid(Str, 1, charIniPos - 1)
            If reSt <> "" Then
                cMds(I) = reSt
                I = I + 1
                cMds(I) = tmp
                Str = Replace(Str, reSt, "", 1, 1)
                Str = Replace(Str, tmp, "", 1, 1)
                GoTo Jump
            Else
                cMds(I) = UCase(tmp)
                Str = Replace(Str, tmp, "", 1, 1)
                GoTo Jump
            End If
        Else
            'Comando não fechado, erro de digitação
            MsgBox "Comandos inválidos no parâmetro SendKeys. Verifique!", vbCritical, "Erro"
        End If
        
    End If
    
    cMds(I) = UCase(Str)
    Str = ""

Jump:

    I = I + 1
    
Wend
            
'Agora, executar os comandos do Array
For z = 0 To I - 1

    'Vendo o tipo de comando ou texto
    If InStr(1, espK, cMds(z)) Then
    
        If cMds(z) = "<ENTER>" Then
            'Manda Enter
            pw.Enter
        Else
            'Manda EspKey
            cMds(z) = UCase(Replace(Replace(cMds(z), "<", ""), ">", ""))
            pw.SendEspKey CStr(cMds(z))
        End If
        
    Else
    
        If UCase(Mid(cMds(z), 1, 3)) = "<PF" Then
            'Manda PF
            cMds(z) = UCase(Replace(Replace(cMds(z), "<PF", ""), ">", ""))
            pw.SendPFKey CInt(cMds(z))
        Else
            'PutString
            pw.PutString CInt(line), CInt(col), CStr(cMds(z))
        End If
        
    End If
    
Next

End Sub

Public Sub WaitForString(text As String, line As Integer, col As Integer, Optional timeOut As Integer):

'Aguarda string ser encontrada na tela

Dim count As Integer

'Padrão de 5 minutos caso não seja especificado
If timeOut = 0 Or timeOut = Null Then timeOut = 300

count = 1
While count < timeOut
    'Testa o texto na tela
    If text = pw.GetString(line, col, Len(text)) Then
        Exit Sub
    Else
        Sleep 1000
        DoEvents
    End If
    count = count + 1
Wend

'Se exceder o timeout, pausa a macro.
If count >= timeOut Then
    MsgBox ("Tempo de espera esgotado! Verifique a sessão do Rede Caixa.")
    Stop
End If

End Sub

Public Sub WaitExec()
    
    'Verifica se o host está ocupado, semelhante ao WaitHostOK, mas sem delay
    While pw.Status <> 0: DoEvents: Wend
    
End Sub

Public Sub MoveTo(LINHA As Variant, col As Variant, Optional page As Integer)

    'Posiciona o cursor - ignora Page
    SetCursor CInt(LINHA), CInt(col)

End Sub

