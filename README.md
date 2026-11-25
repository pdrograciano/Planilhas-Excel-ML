# MIXBONIX ORDER

Aplicação web interna para converter o Excel de vendas do Mercado Livre no modelo de pedidos da MIXBONIX. O navegador lê e gera os arquivos localmente. A aplicação funciona em hospedagem estática com HTTPS.

## Instalação e execução

Requer Node.js 22.13 ou superior.

```sh
npm ci
npm run dev
```

Abra o endereço que o Vite imprimir no terminal. No Windows, se o PowerShell bloquear `npm.ps1`, execute os mesmos comandos com `npm.cmd`.

## Uso

1. No Mercado Livre, baixe o relatório `.xlsx` com a aba **Vendas BR**.
2. Arraste o arquivo para a área de seleção ou clique em **Selecionar arquivo**.
3. Clique em **Gerar planilhas** e escolha **Arquivo Formatado** para manter a sequência do relatório ou **Arquivo Ordenado por Itens** para agrupar pela coluna **ITEM** de A a Z. Em ambas as versões, a coluna **ORDEM** começa em 1 e segue a sequência das linhas exibidas.
4. No **Histórico**, use **Arquivo original ML** para baixar o arquivo enviado ou selecione uma das versões formatadas. Use a lixeira para apagar uma entrada ou **Limpar histórico** para apagar todas as entradas da guia.

As duas versões usam a data e a hora locais da geração. O **Arquivo Formatado** usa o nome `24-11-2026_12-34-56_Formatado_MercadoLivre_mixbonixgerator.xlsx` e o **Arquivo Ordenado por Itens** usa `24-11-2026_12-34-56_PorItens_MercadoLivre_mixbonixgerator.xlsx`.

Quando o Mercado Livre registra um pacote com vários produtos, a coluna **Título do anúncio** mostra **(X pedidos feitos)** na linha do comprador, seguido das linhas dos produtos. A versão por itens ordena o bloco pelo primeiro item e mantém as linhas juntas. A coluna **ORDEM** numera as linhas dos produtos.

Use **PDF nomes e sku** para gerar etiquetas de **40 mm de largura por 25 mm de altura**. O PDF segue a sequência do relatório original e gera uma página por envio, com o nome do comprador, o identificador do envio, o título de cada produto, a quantidade e o SKU. Selecione o papel de **40 × 25 mm** na impressora e imprima em **tamanho real (100%)**.

## Conferência de pedidos do dia

1. Abra **Conferência de pedidos do dia** e selecione o PDF das etiquetas de envio do Mercado Livre.
2. O site lê os QR codes e os destinatários no aparelho. Para ver os produtos e quantidades, use uma planilha correspondente do histórico ou carregue a planilha de vendas original do lote. A associação usa o identificador do envio.
3. Abra a câmera, permita o acesso e aponte para o QR de cada etiqueta. Compare o destinatário e os itens exibidos com o pacote. Leituras repetidas não aumentam o total.
4. Use **Finalizar** para ver os envios identificados, pendentes e o total de leituras que não puderam ser associadas ao PDF.

O ícone de sol ou lua no canto superior direito alterna o tema do site. A preferência é lembrada no navegador.

Os arquivos do **Histórico** ficam na guia e sobrevivem à recarga. Se o navegador atingir o limite de armazenamento, baixe os arquivos da entrada sinalizada antes de recarregar. As regras de abreviação da coluna **ITEM** ficam em `src/formatador/item.ts`.

## Testes e publicação

```sh
npm test
npm run build
npm run test:browser
```

Os testes de navegador usam o Google Chrome instalado e a compilação da pasta `dist/`. Para usar o Edge, configure `MIXBONIX_CANAL_BROWSER` como `msedge`. Os testes usam dados fictícios. Para conferir o relatório real localmente, configure `MIXBONIX_EXEMPLO` com o caminho do Excel de referência e execute `npm test`.

O GitHub Pages publica o conteúdo de `dist/` após cada envio à branch `main`. Em **Settings → Pages**, a origem da publicação deve ser **GitHub Actions**. Use `npm run preview` para conferir a compilação localmente.
