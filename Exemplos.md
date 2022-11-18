## Exemplos de uso da classe VBA - PW3270

**1. Exemplo básico:**
```vb
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

**2. Exemplo WORD**

```vb
Public Sub Pegatela()

    'Pega a tela do host usando GetScreen, copia no word e formata
    
    Dim pw As New clsLibHllapi
    
    pw.Connect "pw3270:A"
    
    Selection.TypeText text:=pw.GetScreen()
    Selection.TypeParagraph
    Selection.WholeStory
    Selection.Font.Name = "Courier New"
    Selection.Font.Size = 9
    
    pw.Disconnect
    
End Sub
```

**3. Exemplo EXCEL**

```vb
Public Sub Exemplo()

    Dim someTxt As String

    'Cria um novo objeto da classe
    Dim pw As New clsLibHllapi 'SUBSTITUA PELO NOME DA CLASSE QUE VOCÊ CRIOU, Classe1, por exemplo :)


    'Necessario para iniciar
    'O parâmetro "pw3270:A" é opcional, poderia ser só pw.Connect
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
