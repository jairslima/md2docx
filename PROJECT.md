# MD to DOCX Converter by Jair Lima

Conversor de arquivos Markdown (.md) para Word (.docx) com formatação completa e fiel ao padrão DOCX. Funciona como executável standalone chamável de qualquer pasta via terminal.

## Stack e Dependências

- **Python 3.14+**
- **python-docx 1.2+** — geração do arquivo DOCX
- **mistune 3.2+** — parser de Markdown (AST-based)
- **lxml 6+** — dependência do python-docx
- **PyInstaller 6+** — geração do executável .exe

## Estrutura de Arquivos

```
ConversorMD2DocX/
├── md2docx.py          # Script principal (conversor + CLI)
├── build.bat           # Script de build para gerar o .exe
├── requirements.txt    # Dependências Python
├── test_sample.md      # Arquivo MD de teste com todos os elementos
├── test_sample.docx    # DOCX gerado pelo teste
├── dist/
│   └── md2docx.exe     # Executável final (standalone)
└── PROJECT.md          # Este arquivo
```

## Comandos Essenciais

```bash
# Instalar dependências
pip install -r requirements.txt

# Executar via Python diretamente
python md2docx.py                     # Converter todos .md da pasta atual
python md2docx.py arquivo.md          # Converter arquivo específico
python md2docx.py arquivo.md --force  # Forçar sobrescrita
python md2docx.py "C:\pasta\"         # Converter todos .md da pasta (batch)

# Build do executável
build.bat                             # Gera dist/md2docx.exe

# Usar o executável (de qualquer pasta)
md2docx                               # Converter todos .md da pasta atual
md2docx README.md                     # Converter arquivo específico
md2docx README.md -f                  # Forçar sobrescrita
md2docx "C:\pasta\"                   # Batch — pasta como argumento posicional
md2docx --folder C:\docs              # Batch — via flag
md2docx --output C:\saida README.md   # Pasta de saída customizada
```

## Instalar o executável globalmente

Copiar `dist/md2docx.exe` para uma pasta no PATH do sistema:
- `C:\Windows\System32` (requer admin)
- `C:\Tools` ou similar (adicionar ao PATH manualmente)

## Elementos Markdown Suportados

| Elemento          | Suporte    | Estilo DOCX               |
|-------------------|-----------|---------------------------|
| H1 – H6           | ✅ completo | Heading 1-6              |
| Parágrafo         | ✅          | Normal                   |
| **Negrito**       | ✅          | Run bold                 |
| *Itálico*         | ✅          | Run italic               |
| ***Bold+Italic*** | ✅          | Run bold+italic          |
| ~~Tachado~~       | ✅          | Run strikethrough        |
| `Código inline`   | ✅          | Courier New + highlight  |
| Bloco de código   | ✅          | Code Block + shading     |
| Lista bullets     | ✅ aninhada | List Bullet 1-3          |
| Lista numerada    | ✅ aninhada | List Number 1-3          |
| Tabela            | ✅ alinhada | Word Table com borda     |
| Blockquote        | ✅          | Block Quote + borda esq  |
| Citação bíblica   | ✅          | Âmbar + ref direita      |
| Link              | ✅          | Hyperlink DOCX nativo    |
| Imagem local      | ✅          | Inline picture           |
| Linha horizontal  | ✅          | Paragraph border         |
| YAML front matter | ✅ ignorado | —                        |

## Funcionalidades Especiais (v2)

- **Página de capa automática** — detecta `# Título\n## Subtítulo\n*Autor*\n---` no início do MD e gera capa centralizada com estilos tipográficos
- **Sumário (TOC)** — inserido após a capa; no Word pressionar **Ctrl+A → F9** para popular
- **Número de página no rodapé** — campo PAGE centralizado, cinza, em todas as páginas
- **Quebra de página antes de H1** — cada H1 (exceto o primeiro) inicia automaticamente em nova página
- **Citações bíblicas estilizadas** — blockquotes com padrão `> texto\n> — Referência` ganham borda âmbar e referência alinhada à direita
- **Batch por pasta** — passar uma pasta como argumento posicional funciona (ex: `md2docx "C:\pasta\"`)

## Decisões Arquiteturais

- **mistune 3 (AST mode)**: Usado pela API `BaseRenderer` que recebe tokens estruturados em vez de HTML string, permitindo renderização direta para DOCX sem intermediário HTML.
- **Estilos nomeados DOCX**: Headings, Code Block, Block Quote usam estilos Word nomeados para compatibilidade máxima com editores DOCX.
- **PyInstaller --onefile**: Executável único, sem dependências externas após build.
- **Encoding UTF-8**: `sys.stdout.reconfigure(encoding='utf-8')` resolve problema com terminal Windows (cp1252).
- **Detecção de capa por regex**: `extract_cover()` usa regex para extrair o bloco de capa sem interferir no parser mistune.

## Estado Atual (2026-03-18) — v2.2

- ✅ Conversor funcional e testado com livro real (6 arquivos .md, 0 erros)
- ✅ Executável `dist/md2docx.exe` compilado e atualizado (instalado em System32)
- ✅ Capa automática, TOC, rodapé, quebra H1, citações bíblicas implementados
- ✅ Batch por pasta via argumento posicional corrigido
- ✅ Banner "MD to DOCX Converter by Jair Lima" exibido em toda execução
- ✅ Spinner animado durante conversão (evita aparência de travamento)
- ✅ Tempo de execução e tamanho do DOCX exibidos no `[OK]`
- ✅ `--version` / `-v` exibe versão e encerra
- ✅ `--recursive` / `-r` varre subpastas recursivamente
- ✅ Imagens por URL remota: download automático para temp file
- ✅ Ajuste automático de largura de colunas proporcional ao conteúdo

## Próximos Passos (sugestões)

- Suporte a footnotes DOCX nativo (atualmente ignorado)
- Task lists com checkboxes (☑ / ☐)
- Ícone customizado no .exe

## Problemas Conhecidos

- TOC requer atualização manual no Word (Ctrl+A → F9) — limitação do formato DOCX
- Imagens por URL dependem de conexão ativa; falha silenciosa exibe texto de alt
