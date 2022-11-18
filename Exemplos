## Exemplos de uso da classe

**1. Exemplo básico:**
```vbs
Public Sub Exemplo()

    Dim someTxt As String

    'Cria um novo objeto da classe
    Dim pw As New clsLibHllapi 'SUBSTITUA PELO NOME DA CLASSE QUE VOCÊ CRIOU, Classe1, por exemplo :)

    'Necessario para iniciar
    pw.Connect "pw3270:A"

    'Pega texto na tela
    someTxt = pw.GetString(1, 2, 26)

    'Mostra o texto
    MsgBox someTxt

    'Envia o texto AOF na posicao desejada
    pw.PutString 21, 22, "AOF   "

    'Enter
    pw.Enter

    'Enviando a tecla TAB
    pw.SendEspKey "TAB"

    'Pesquisando texto
    If pw.FindText("5697") <> 0 Then
        MsgBox "Texto Existe!"
    End If

    'Envia pressionamento de F3
    pw.SendPFKey 3

    'Desconecta o terminal
    pw.Disconnect


End Sub
```
