# OpenSlideV2

Um fork do [open-slide](https://github.com/1weiho/open-slide), de [Yiwei Ho](https://github.com/1weiho), com as três coisas que faltavam para um deck sair daqui direto para o cliente: edição direta na tela, exportação para PowerPoint editável e resolução automática das fontes e dos assets de que o deck depende.

*Read this in [English](README.md).*

## O que é

O open-slide é um framework de apresentações escrito para agentes de código. Você descreve o deck em linguagem natural, o agente escreve os componentes React, e o framework cuida do canvas fixo de 1920x1080, da escala, da navegação, do hot reload e do modo apresentador. Todo o mérito disso é do autor original.

Este fork não muda nada dessa base. Ele acrescenta três recursos por cima, e todo o resto acompanha o projeto original.

## Começando

**Requisitos**

- Node.js 18 ou mais novo
- pnpm (o repositório é um monorepo pnpm com Turbo)

**Instalação**

```bash
git clone https://github.com/bruno-org/OpenSlideV2.git
cd OpenSlideV2
pnpm install
pnpm setup:fonts
```

O `pnpm setup:fonts` instala as fontes do projeto no perfil do usuário atual, sem precisar de administrador. Ele importa porque a exportação para PPTX editável referencia as fontes pelo nome: sem elas instaladas, quem abrir o arquivo vê a substituição automática do PowerPoint e o desenho da letra muda.

**Publicar este fork em um workspace open-slide existente**

```bash
pnpm install:workspace ../MeuWorkspace
```

O comando compila, empacota com `npm pack` e reinstala. Ele não copia arquivos soltos para dentro de `node_modules`, e isso é proposital: o Vite mantém cache de transformação para arquivos dentro de `node_modules` e continuaria servindo a versão anterior mesmo depois de reiniciar.

## O que este fork acrescenta

### Arrastar e redimensionar na tela

O inspetor do projeto original edita as propriedades do elemento selecionado (texto, tipografia, cor, imagem). Aqui ele também move e redimensiona:

- arrastar o corpo do elemento reposiciona;
- as quatro alças de canto redimensionam, com a borda oposta ancorada.

O gesto termina virando uma operação `set-style` comum (`translate`, `width`, `height`), a mesma que o painel já usava, então salvar, desfazer e refazer funcionam sem nada extra e o resultado é gravado no `.tsx`. Ajuste fino deixa de exigir uma ida ao agente.

### Exportação para PPTX editável

O menu `Download` ganhou **Export as editable PPTX**. O texto sai como texto de verdade do PowerPoint, editável, e o slide continua igual ao que aparece na tela.

A ordem das operações é o que garante isso:

1. cada página é capturada como imagem de fundo, com os textos promovidos tornados transparentes. Gradiente, sombra, SVG, `clip-path` e filtro continuam exatamente onde estavam, porque é a captura da própria página;
2. por cima entram caixas de texto nativas, uma por linha renderizada, no retângulo que o navegador mediu, com quebra automática desligada. Uma fonte substituída pode mudar o desenho da letra, mas nunca refluir o deck;
3. o que não dá para reproduzir com fidelidade como texto do PowerPoint (texto preenchido com gradiente, sombra de texto, texto girado ou recortado) não é promovido: fica pintado no fundo. Perde-se a edição daquele trecho, nunca a aparência.

Três detalhes separam "parecido" de "igual":

- **peso**: o PowerPoint não tem eixo de peso, só negrito ligado ou desligado. Um peso 500 ou 800 é apontado para a face nomeada correspondente ("Geist SemiBold", "Geist ExtraBold"). Esses nomes vêm do que cada arquivo de fonte declara internamente, e não do nome do arquivo: o `Geist-UltraBlack.ttf` se declara como "Geist ExtraBold";
- **kerning**: o PowerPoint lê apenas a tabela `kern` legada, então fontes que guardam o kerning em GPOS (a maioria das modernas) renderizam sem kerning e cada linha sai mais larga. A calibração mede com kerning desligado justamente para reproduzir isso, e transforma a diferença em espaçamento entre letras;
- **rede de segurança**: uma família que não dá para garantir na máquina de destino é mapeada para uma fonte presente em qualquer Windows (Segoe UI, Georgia, Consolas), escolhida aqui e não pelo PowerPoint, com a largura da linha calibrada contra ela.

### Preflight de dependências

Todo `open-slide dev` começa conferindo se a máquina tem o que os decks pedem, e resolvendo o que falta:

- lê os decks em `slides/` e os temas em `themes/` e levanta as famílias de fonte usadas, incluindo a que o runtime aplica por padrão;
- verifica quais já estão instaladas no sistema;
- busca o que falta: primeiro no pacote npm correspondente, depois no repositório [google/fonts](https://github.com/google/fonts), que é a fonte pública estável de TTF instalável (o endpoint de download do site responde com HTML, e a API de CSS serve woff2, que nenhum sistema operacional instala);
- instala para o usuário atual, em Windows, macOS e Linux;
- assets: um deck pode trazer `slides/<id>/assets.manifest.json` mapeando nome de arquivo para URL. O que estiver faltando em disco é baixado.

Nada disso derruba o servidor. Falha de rede vira linha de relatório. Cada deck aberto deixa a máquina um pouco mais equipada do que a encontrou.

Para rodar sozinho: `open-slide preflight`, ou `--no-install` para apenas listar o que falta.

## Verificação

Os dois recursos têm checagem executável. Rode a partir de `packages/core`, com o servidor de desenvolvimento do workspace no ar:

```bash
pnpm dev:demo                               # em outro terminal, na raiz do repositório
node tools/verify-drag-resize.mjs           # arrasta, redimensiona, confere o .tsx
node tools/verify-pptx-editable.mjs         # exporta, abre no PowerPoint, compara
```

Os dois apontam para fixtures que vivem neste repositório (`apps/demo/slides/verify-*`), então rodam em qualquer clone sem preparação. O segundo exporta o mesmo deck nos dois formatos, usa o PPTX de imagem como gabarito, abre o PPTX editável no PowerPoint de verdade (via COM, no Windows), exporta os slides e compara. Ele mede geometria, que é a prova real de fidelidade, e mantém a contagem de pixels divergentes como rede contra regressão grosseira.

Resultado no fixture de estresse, que reúne gradiente, brilho radial, sombra, SVG embutido, títulos de várias linhas, pesos misturados e todos os alinhamentos de texto:

| medida | resultado |
| --- | --- |
| deslocamento horizontal | até 4,9px |
| deslocamento vertical | até 14,7px |
| largura do texto | dentro de 0,3% |
| altura do texto | dentro de 0,5% |
| pixels divergentes | 2,1% a 3,9% |

Os pixels que sobram são antialiasing de borda e a rasterização do próprio PowerPoint. O deslocamento vertical se concentra em texto grande cujo `line-height` é menor que a linha natural da fonte: o PowerPoint não comprime a linha abaixo desse mínimo, então o texto fica alguns pixels mais alto. Num canvas de 1080px, isso é cerca de 1%.

Os dois scripts sobem um navegador próprio, com perfil temporário, e não encostam em nenhum navegador que você já tenha aberto. O passo 3 da checagem de PPTX precisa do PowerPoint instalado, então só roda no Windows; todo o resto funciona em qualquer sistema.

## Limites conhecidos

- Texto com preenchimento em gradiente, sombra, rotação ou recorte não vira texto editável. Fica na imagem de fundo, visualmente idêntico.
- O PPTX editável depende das fontes instaladas na máquina que abre o arquivo. O preflight resolve isso para quem usa o OpenSlideV2; para quem apenas recebe o `.pptx`, vale a rede de segurança com fonte do sistema.
- Formas e imagens não viram objetos nativos: permanecem no fundo. Só o texto é promovido.

## Atribuição

O [open-slide](https://github.com/1weiho/open-slide) foi criado por [Yiwei Ho](https://github.com/1weiho) e é a base inteira deste projeto. Este repositório é um fork independente que acrescenta recursos sobre aquele trabalho.

Se o OpenSlideV2 for útil para você, vá primeiro à fonte: [dê uma estrela no repositório original](https://github.com/1weiho/open-slide) e considere [apoiar o desenvolvimento dele](https://ko-fi.com/D1D11YPUP1).

## Licença

MIT, para os dois. Uso livre para copiar, modificar e distribuir, comercialmente ou não, mantendo o aviso de copyright. Veja [LICENSE](LICENSE).

## Contribuindo

Issues e pull requests são bem-vindos. Mudanças no núcleo do framework costumam render mais enviadas ao projeto original, onde todo mundo se beneficia delas.
