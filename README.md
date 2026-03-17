# DEV_JSON (VBA / Access) — Motor JSON modular

Este repositório contém o motor JSON em VBA (Access/Office), organizado em módulos por responsabilidade (`modJSON_*`), além do integrador e testes.

## Versão estável

- **Release/Tag:** `stable-2026-03-17`
- **Testes (Access/VBA):** `RunAllTests_32: Passed=6 Failed=0` (2026-03-17)

---

## Estrutura (o que tem aqui)

O código VBA fica em: **`src/vba/`**

### Módulos do motor (modJSON_*)
- `modJSON_IO.bas`  
  Rotinas de entrada/saída e tratamento de texto JSON.
- `modJSON_Normalize.bas`  
  Normalização e pré-processamento de JSON/texto.
- `modJSON_Format.bas`  
  Formatação/pretty-print (indentação) e utilidades de apresentação.
- `modJSON_Flatten.bas`  
  Conversão de JSON para estrutura “achatada” (flatten).
- `modJSON_Unflatten.bas`  
  Reconstrução (unflatten) da estrutura original (objetos/arrays aninhados).
- `modJSON_Utils.bas`  
  Utilitários comuns (strings, arrays, validações).
- `modJSON_Stubs.bas`  
  Stubs/compatibilidade e pontos de extensão.

### Integração e testes
- `mod_Globals.bas`  
  Globais/constantes usadas pelo projeto.
- `mod_Integrador.bas` e `Form_frmIntegrador.cls`  
  Camada de integração com o app.
- `mod_Testes.bas`  
  Testes automatizados (executar `RunAllTests_32`).

> Observação: o módulo `mod_Auxiliar.bas` pode existir apenas como “placeholder”/organização (não deve conter duplicatas de funções públicas).

---

## API pública (funções principais)

### `JSON_Normalizar(...)`
Normaliza/ajusta texto JSON para consumo/parse/uso interno.

Uso típico:
- limpar/normalizar strings JSON recebidas de fontes externas
- preparar JSON antes de formatar/flatten/unflatten

### `JSON_Indentar(...)`
Gera uma versão “pretty” do JSON (indentado), para leitura/depuração/log.

### `JSON_Flatten(...)`
Converte um JSON (com objetos/arrays aninhados) para uma forma “achatada”, adequada para:
- persistência tabular
- transporte em estruturas simples
- comparação/diff

### `JSON_Unflatten(...)`
Reconstrói o JSON original (aninhado) a partir da forma “flatten”.

---

## Como usar no Access (importar módulos)

1. No Access, abra o Editor VBA (ALT+F11).
2. Importe os arquivos `.bas` e `.cls` a partir de `src/vba/`:
   - `src/vba/modJSON_*.bas`
   - `src/vba/mod_Globals.bas`, `src/vba/mod_Integrador.bas`, `src/vba/mod_Testes.bas`
   - `src/vba/Form_frmIntegrador.cls` (se aplicável ao seu app)
3. Compile o projeto:
   - VBA Editor → **Debug → Compile**

> Dica: evite manter módulos antigos/legado com as mesmas funções públicas (ex.: versões antigas de `JSON_Unflatten`), para não causar conflito de compilação.

---

## Rodando os testes

No VBA Editor:
1. Abra o módulo `mod_Testes.bas`
2. Execute a macro:
   - `RunAllTests_32`

Resultado esperado (estado estável em 2026-03-17):
- `Passed=6 Failed=0`

---

## Notas / Compatibilidade

- O motor foi refatorado para módulos `modJSON_*` por responsabilidade.
- Recomendado manter este repositório como **fonte da verdade** no Git e importar para o `.accdb` quando necessário.
- Se você encontrar “texto mesclando” ao digitar no Copilot Chat no navegador, um hard refresh (Ctrl+Shift+R) costuma resolver; se persistir, limpe cookies/dados do `github.com`.
