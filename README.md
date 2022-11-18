## Classe de automação VBA para comunicação com o terminal pw3270

Classe em Visual Basic For Aplications - VBA, para criação de macros/automação do pw3270 através do MS Office (Excel, Word, etc):

Para instalar, tem duas opções:

**1. Criação da Classe:**

  - Crie uma classe com o nome clsLibHllapi (Ou o nome que preferir, a classe é sua!), 
    clicando no menu *inserir > Módulo de classe*;
  - Copie o código do arquivo [clsLibHllApi_32_64_GIT.cls](clsLibHllApi_32_64_GIT.cls) **a partir da linha 10** e cole na classe criada;
  - Divirta-se!

**2. Importando a Classe:**

  - Faça o download do arquivo [clsLibHllApi_32_64_GIT.cls](clsLibHllApi_32_64_GIT.cls);
  - No explorador de projetos, no canto superior direito, clique com o botão direito do mouse em 
    '*Módulos de Classe*' e selecione '*Importar Arquivo*';
  - Selecione o arquivo baixado.
  - Divirta-se!

**OBS**: Dependendo da instalação, a sua dll pode ter o nome de libhllapi.dll. Assim, basta mudar o nome nas declarações.

Por Exemplo:
```vba
  Private Declare Function hllapi_init Lib "libhllapi32.dll" (ByVal tp As String) As Long
```
Ficaria:
```vba
  Private Declare Function hllapi_init Lib "libhllapi.dll" (ByVal tp As String) As Long
```
