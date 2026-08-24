# Web Part de Carrossel — SharePoint Online (SPFx)

## O que este web part faz
- Carrossel de imagens com **tempo de transição configurável** (2 a 30 segundos) e reprodução automática opcional.
- Para cada slide você pode escolher a imagem **do computador (upload) ou da biblioteca de documentos do SharePoint**, além de definir: **Pré-cabeçalho, Título, Descrição** e **Chamada à ação** (Botão, Ícone, Texto ou Cartão inteiro), com link.
- Setas de navegação e indicadores (bolinhas) que podem ser ligados/desligados.
- **Altura** ajustável em pixels pelo painel de propriedades. A **largura** acompanha automaticamente a coluna da seção onde o web part é inserido (padrão de qualquer web part moderno do SharePoint — não é possível arrastar a largura como se faz com o web part de Imagem).

## Pré-requisitos na sua máquina
- Node.js 16.x ou 18.x (SPFx 1.18 não é compatível com Node 20+)
- npm
- Gulp CLI: `npm install -g gulp-cli`
- Yeoman + gerador SPFx (opcional, só se quiser recriar o scaffold do zero): `npm install -g yo @microsoft/generator-sharepoint`

## Passo a passo para gerar o .sppkg

1. Extraia esta pasta `carousel-webpart` em qualquer diretório.
2. Abra o terminal dentro da pasta e instale as dependências:
   ```
   npm install
   ```
3. (Opcional, mas recomendado) Teste localmente no Workbench:
   ```
   gulp serve
   ```
   Isso abre `https://localhost:5432/workbench` no navegador para testar o carrossel antes de publicar.
4. Gere o pacote de produção:
   ```
   gulp bundle --ship
   gulp package-solution --ship
   ```
5. O arquivo final estará em:
   ```
   sharepoint/solution/carousel-webpart.sppkg
   ```

## Como instalar no SharePoint Online
1. Acesse o **App Catalog** do seu tenant (Central de Administração > Mais recursos > Apps > App Catalog). Se não existir, um administrador precisa criar um.
2. Vá em **Apps para SharePoint** e clique em **Carregar**.
3. Selecione o arquivo `carousel-webpart.sppkg`.
4. Marque a opção **"Confiar nesse aplicativo"** quando solicitado (o pacote já está configurado com `skipFeatureDeployment: true`, então fica disponível automaticamente em todos os sites do tenant).
5. Vá até a página do SharePoint onde quer usar o carrossel, edite a página, adicione uma seção (ex.: seção flexível) e insira o web part **"Carrossel Customizado"** pelo seletor de web parts.
6. Edite o web part (ícone de lápis) para adicionar os slides, imagens, textos e ajustar tempo de transição e altura.

## Observações importantes
- O upload de imagens do computador grava o arquivo na biblioteca **Site Assets** do site atual (comportamento padrão do SharePoint para este tipo de seletor). Se preferir outra biblioteca, me avise que ajusto o caminho no código (`CarouselWebPart.ts`, método `onImageSave`).
- As versões de dependências (`@microsoft/sp-*: 1.18.2`) são compatíveis com SharePoint Online atual. Se o seu tenant/CLI usar outra versão do SPFx, pode ser necessário ajustar as versões no `package.json`.
- Se o `npm install` reclamar de conflitos de versão do `@fluentui/react` ou `@pnp/spfx-property-controls`, rode com `--legacy-peer-deps`.
- Não testei a compilação real deste código (meu ambiente aqui não tem acesso à internet para baixar pacotes npm), então é possível que precise de pequenos ajustes de tipos/versões no primeiro build — me envie o erro que eu corrijo rapidamente.

## Estrutura de arquivos
```
carousel-webpart/
├── config/                  → configs do SPFx (bundle, package, serve)
├── src/webparts/carousel/
│   ├── CarouselWebPart.ts   → classe principal + Property Pane
│   ├── CarouselWebPart.manifest.json
│   ├── ICarouselWebPartProps.ts
│   ├── components/
│   │   ├── Carousel.tsx     → componente React do carrossel
│   │   ├── ICarouselProps.ts
│   │   └── Carousel.module.scss
│   └── loc/                 → strings de localização
├── package.json
├── tsconfig.json
└── gulpfile.js
```
