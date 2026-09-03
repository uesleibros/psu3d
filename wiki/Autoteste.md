# Autoteste

```vba
PSelfTest.Psu3DSelfTest        ' abre um MsgBox com o resultado
txt = PSelfTest.RunAll()       ' devolve o relatório como string
```

174 asserções contra respostas conhecidas, rodando inteiramente em memória. O renderer entra em `DryRun`, então o pipeline inteiro executa sem slide e sem deixar nada para trás. Dá para rodar do VBE com nenhuma apresentação aberta.

Ele existe porque as falhas que importam numa lib importada à mão são silenciosas: um módulo que ficou de fora, uma classe que não veio junto, o `JSON.cls` ausente. O relatório nomeia cada verificação que falhou.

## O que cada bloco cobre

| bloco | o que prova |
|---|---|
| `CheckMath` | clamp, lerp, ângulos, overlap, random determinístico |
| `CheckColor` | empacotar, misturar, hex de ida e volta |
| `CheckLighting` | direção, intensidade, faixas de névoa |
| `CheckMaterials` | registro, ids, tabela de sombreamento, invalidação |
| `CheckCanvas` | retângulo, projeção, coordenadas locais |
| `CheckCamera` | yaw, pitch, limites, `LookAt`, transformação |
| `CheckScene` | adicionar, editar, span, rampa, caixa girada, reaproveitamento de slot |
| `CheckIndex` | a grade responde exatamente o que uma varredura responderia |
| `CheckMotion` | oscilação, passo reportado, giro |
| `CheckBody` | gravidade, parede, degrau, escada, água, gatilho, quique, carona |
| `CheckOrder` | a moeda é pintada antes da plataforma vista de baixo, e depois vista de cima |
| `CheckRenderer` | o pipeline inteiro sem slide |
| `CheckLevel` | parse, `repeat`, spawn, orçamento, erros nomeados |
| `CheckSave` | escrever e reler devolve o mesmo mundo |
| `CheckHardening` | typo em arquivo e coordenada absurda não derrubam nada |

## Quatro blocos escritos com cuidado extra

A cor de teste do material é clara o suficiente para que topo e base fiquem distinguíveis depois do sombreamento. Com uma cor escura os dois arredondam para o mesmo valor e o teste passaria por acidente.

No `CheckBody`, tudo que existe para ser pisado fica acima do chão. Duas superfícies na mesma altura fariam o teste depender de qual a consulta alcançou primeiro, e um teste desses passa e falha sem ninguém ter mudado nada.

No `CheckOrder`, além da moeda e da plataforma existe um chamariz distante que não divide um pixel com nenhuma das duas. Ele é o ponto do teste: era ele quem fechava o ciclo que sacrificava a moeda.

O `CheckSave` verifica que o documento salvo tem ponto decimal. É a única falha do salvar que não aparece como número errado: numa máquina configurada com vírgula, o `Format` escreveria um arquivo que nenhum leitor de JSON aceita de volta.

## Se faltar o JSON

`CheckLevel`, `CheckSave` e `CheckHardening` precisam do `JSON.cls`. Sem ele o projeto não compila, e o autoteste nem chega a rodar. É esse o sintoma de "esqueci de importar o JSON".
