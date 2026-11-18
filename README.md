# Gabina Impressões 3D — Calculadora e Estoque

- Calculadora rápida: https://adseads.github.io/calculadora-gabina/
- Estoque: https://adseads.github.io/calculadora-gabina/estoque.html
- Calculadora completa: https://adseads.github.io/calculadora-gabina/completa.html

## Como ligar o estoque (uma vez só)

Os dados do estoque ficam numa Planilha Google. A página lê e grava nela.

1. Crie uma Planilha Google em branco (sheets.new) com o nome **Gabina 3D — Estoque**.
2. Na planilha: **Extensões → Apps Script**.
3. Apague o que estiver no editor, cole todo o conteúdo de `apps-script/Codigo.gs` e salve (ícone de disquete).
4. (Opcional) Troque a senha na linha `const SENHA = '1234';`.
5. **Implantar → Nova implantação** → engrenagem → **App da Web**.
   - Executar como: **Eu**
   - Quem pode acessar: **Qualquer pessoa**
   - **Implantar** → autorize com sua conta Google (se aparecer "app não verificado": Avançado → Acessar).
6. Copie o **URL do app da Web** (termina em `/exec`) e cole em `config.js`, entre as aspas.

Na primeira vez que a página abrir, a planilha cria as abas Filamentos e Insumos sozinha, já com os filamentos cadastrados.

Se mudar o código do Apps Script depois: **Implantar → Gerenciar implantações → editar → Nova versão** (o endereço continua o mesmo).
