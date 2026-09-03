# Instalação

## O que você precisa

PowerPoint com VBA habilitado, e um arquivo salvo como `.pptm`. Não existe outra dependência obrigatória: nenhuma DLL, nenhuma referência a marcar em Tools, nenhum instalador.

A lib usa duas funções da API do Windows, declaradas dentro dos próprios módulos: `QueryPerformanceCounter` para o relógio e `GetAsyncKeyState` para o teclado do demo. As duas já existem em qualquer Windows.

## Importando

1. Abra o VBE com `Alt+F11`.
2. Menu **File** e depois **Import File**, ou `Ctrl+M`.
3. Importe os arquivos de `engine/`.

A ordem não importa para o compilador, só para leitura.

```
PCore.bas        tipos, matemática, cor, relógio
PLighting.bas    luz e névoa
PMaterial.cls    uma superfície
PMaterials.bas   registro de materiais
PCanvas.cls      onde projetar
PCamera.cls      de onde olhar
PRenderer.cls    como desenhar
PScene.cls       o que existe
PBody.cls        o que se move e colide      (opcional)
PLevel.bas       ler e escrever JSON         (opcional)
Psu3D.bas        fachada
PSelfTest.bas    autoteste                   (opcional)
PDemo.bas        fase jogável de exemplo     (opcional)
```

## O mínimo

Se você só quer desenhar 3D, oito módulos bastam:

`PCore`, `PLighting`, `PMaterial`, `PMaterials`, `PCanvas`, `PCamera`, `PRenderer`, `PScene`.

`Psu3D.bas` é conveniência em cima desses oito e cabe junto sem custo.

## Módulos de fora

Psu3D não embute código de terceiros. O VBA compila o projeto inteiro, então importar um módulo sem o que ele chama não compila. Por isso o núcleo não chama nada de fora: os oito módulos do núcleo, mais o `Psu3D`, o `PBody` e o `PSelfTest`, compilam sozinhos.

Sobram dois casos, e os dois são opcionais.

### JSON, para `PLevel`

`PLevel.bas` é o único que usa, e usa em toda parte. Sem ele o `PLevel` não compila.

Pegue em **https://github.com/vbacollective/json** e importe o `JSON.cls`.

Se você não vai ler nem escrever fase em arquivo, simplesmente não importe o `PLevel.bas` e o JSON deixa de ser assunto.

### UCursor, para mouse

Não faz parte da lib, não é obrigatório, e **nenhum módulo do núcleo o menciona**. Os doze módulos fora do `PDemo` compilam sem ele.

O único que chama é o `PDemo.bas`, que é um demo em primeira pessoa e portanto precisa de mouse. Se você não vai rodar o demo, não importe o `PDemo` e o assunto some.

A canvas mapeia coordenadas e não lê hardware. Quem tem um ponteiro subtrai `CenterX` e `CenterY` por conta própria, que é a mesma conta:

```vba
dx = MeuCursorX - cv.CenterX
dy = MeuCursorY - cv.CenterY
cam.AddAngles dx * 0.0026, -dy * 0.0026
```

Serve qualquer fonte de posição do ponteiro. O `UCursor` é uma delas, e resolve letterboxing, DPI e modo janela por você.

## Conferindo que deu certo

Rode `PSelfTest.Psu3DSelfTest`. São 174 asserções contra respostas conhecidas, rodando inteiramente em memória: o renderer entra em `DryRun`, então o pipeline inteiro executa sem slide e sem deixar nada para trás. Dá para rodar do VBE com nenhuma apresentação aberta.

Ele existe porque as falhas que importam numa lib importada à mão são silenciosas: um módulo que ficou de fora, uma classe que não veio junto. O relatório nomeia cada verificação que falhou.

Detalhe do `.cls`: ao importar uma classe, o VBE lê o nome do atributo `VB_Name` no topo do arquivo, não o nome do arquivo. Se você renomear o arquivo, a classe continua com o nome antigo. Não renomeie.

## Habilitando macro

O PowerPoint bloqueia macro por padrão em arquivo baixado da internet. Clique com o botão direito no `.pptm`, **Propriedades**, e marque **Desbloquear** na parte de baixo da aba Geral. Sem isso o VBE abre mas nada roda.
