# RDDADOX

RDDADOX é um RDD (Replaceable Database Driver) para [Harbour](https://harbour.github.io/)
que permite acessar bancos relacionais usando ADO/OLE DB e ODBC, mantendo a
interface tradicional de áreas, campos, navegação e gravação do Harbour.

O projeto é especialmente útil para aplicações Harbour que precisam trabalhar
com bancos SQL sem abandonar operações como `DbGoTop()`, `DbSkip()`, `DbSeek()`,
`DbAppend()` e `DbCommit()`.

> **Status:** em desenvolvimento. Teste o comportamento do RDD e do driver
> específico antes de utilizá-lo em produção.

## Recursos

- RDD completo registrado como `RDDADOX`.
- Acesso a tabelas e consultas SQL por meio de recordsets ADO.
- Leitura de metadados e conversão de tipos ADO para tipos de campo Harbour.
- Navegação, inclusão, alteração, exclusão, filtros e transações.
- Criação e remoção de tabelas conforme o dialeto do banco.
- Leitura e gravação opcional de campos BLOB e MEMO.
- Compatibilidade com builds Harbour de 32 e 64 bits, desde que os drivers
  instalados tenham a mesma arquitetura da aplicação.

## Bancos e drivers

As conexões são construídas a partir do engine informado por
`RDDADOX_SetEngine()`. O driver correspondente precisa estar instalado no
Windows e disponível para a arquitetura do executável.

| Engine | Tecnologia esperada |
| --- | --- |
| `MDB` / `ACCESS` | Microsoft Jet OLE DB 4.0 |
| `ACCDB` / `ACCDB64` / `ACEOLEDB` | Microsoft ACE OLE DB 12.0 |
| `SQLITE` | SQLite3 ODBC Driver |
| `DUCKDB` | DuckDB ODBC Driver |
| `MYSQL` / `MYSQL64` | MySQL ODBC 8.0 ou 9.0 |
| `MARIADB` | MariaDB ODBC 3.2 |
| `PGSQL` / `POSTGRESQL` / `PGSQL64` | PostgreSQL ANSI ODBC |
| `MSSQL` / `SQLSERVER` / `SQL` | SQL Server via SQLOLEDB |
| `FIREBIRD` / `FDB` / `GDB` / `IB` | Firebird ODBC Driver |

Também há suporte específico para arquivos Excel (`.xls`) e Paradox (`.db`)
em operações de criação/abertura.

## Requisitos

- Windows.
- Harbour com `hbmk2` disponível no `PATH`.
- Acesso às bibliotecas padrão do Harbour, incluindo `xhb.hbc`.
- Microsoft Data Access Components/ADO disponível no sistema.
- Driver ODBC ou provedor OLE DB adequado ao banco utilizado.
- Aplicação e driver com a mesma arquitetura: 32 bits com 32 bits ou 64 bits
  com 64 bits.

## Compilação

O projeto usa o arquivo `rddadox.hbp`. Com o ambiente Harbour configurado,
execute:

```bat
hbmk2.exe rddadox.hbp
```

Há scripts prontos para os ambientes utilizados no projeto:

```bat
comprddadow32.bat
comprddadow64.bat
```

Esses scripts carregam os ambientes Harbour locais antes de chamar o `hbmk2`.
Se os caminhos da sua instalação forem diferentes, ajuste os scripts ou
compile diretamente com `hbmk2`.

O resultado é colocado em `lib\<plataforma>\<compilador>\`, conforme a
configuração de `rddadox.hbp`.

## Uso básico

O RDD é registrado automaticamente durante a inicialização da biblioteca.
Antes de abrir uma área, configure a tabela e o engine:

```harbour
#include "rddsys.ch"
#include "rddadox.ch"

PROCEDURE AbrirClientes()
   LOCAL cArquivo := "clientes.sqlite"

   RDDADOX_SetEngine( "SQLITE" )
   RDDADOX_SetTable( "clientes" )
   RDDADOX_SetQuery( "SELECT * FROM clientes ORDER BY nome" )

   IF ! DbUseArea( .T., "RDDADOX", cArquivo, "clientes", .T., .F. )
      ? "Não foi possível abrir a tabela"
      RETURN
   ENDIF

   DbGoTop()
   DO WHILE ! Eof()
      ? clientes->nome
      DbSkip()
   ENDDO

   DbCloseArea()
RETURN
```

Para servidores, configure também servidor, usuário e senha antes de abrir a
área:

```harbour
RDDADOX_SetEngine( "POSTGRESQL" )
RDDADOX_SetServer( "localhost" )
RDDADOX_SetUser( "app_user" )
RDDADOX_SetPassword( "use uma senha fora do código-fonte" )
RDDADOX_SetTable( "clientes" )
```

O valor de `RDDADOX_SetQuery()` é opcional. Quando não informado, o RDD usa
`SELECT * FROM <tabela>`.

## BLOBs e MEMOs

Para evitar carregar dados grandes na memória em todas as leituras, BLOBs e
MEMOs não são carregados por padrão. Use as funções auxiliares quando
necessário:

```harbour
// Imagem
RDDADOX_PegarBlobJpg( "foto", "C:\temp\foto.jpg" )
RDDADOX_GravarBlobJpg( "foto", "C:\temp\foto.jpg" )

// Texto longo
cTexto := RDDADOX_PegarMemo( "observacao" )
RDDADOX_GravarMemo( "observacao", cTexto )
```

Também é possível controlar esse comportamento por área com
`RDDADOX_SetLoadBlobs()` e `RDDADOX_SetLoadMemos()`.

## API pública

As principais funções e procedimentos exportados são:

- Configuração: `RDDADOX_SetTable()`, `RDDADOX_SetEngine()`,
  `RDDADOX_SetServer()`, `RDDADOX_SetUser()`, `RDDADOX_SetPassword()`,
  `RDDADOX_SetQuery()` e `RDDADOX_SetLocateFor()`.
- Dados binários e textos longos: `RDDADOX_PegarBlobJpg()`,
  `RDDADOX_GravarBlobJpg()`, `RDDADOX_PegarMemo()` e
  `RDDADOX_GravarMemo()`.
- Opções de carregamento: `RDDADOX_SetLoadBlobs()` e
  `RDDADOX_SetLoadMemos()`.
- Objetos ADO da área atual: `RDDADOX_GetConnection()`,
  `RDDADOX_GetCatalog()` e `RDDADOX_GetRecordSet()`.

O arquivo `rddadox.ch` contém constantes ADO úteis para código Harbour que
precise trabalhar diretamente com os objetos ADO retornados pelo RDD.

## Solução de problemas

- **Driver não encontrado:** confirme o nome e a arquitetura do driver ODBC
  no Administrador de Fontes de Dados ODBC.
- **Falha ao abrir Access/Excel:** instale o provedor Jet/ACE compatível com
  a arquitetura do executável.
- **Conexão recusada:** verifique servidor, porta, banco, usuário, senha e
  permissões no servidor.
- **Consulta sem dados:** confirme o nome configurado em
  `RDDADOX_SetTable()` e, quando usado, o SQL passado para
  `RDDADOX_SetQuery()`.
- **Erro durante a compilação:** confirme que `hbmk2`, `xhb.hbc` e as
  variáveis de ambiente do Harbour estão disponíveis.

## Estrutura do projeto

| Arquivo | Finalidade |
| --- | --- |
| `rddadox.prg` | Implementação do RDD e funções públicas |
| `rddadox.ch` | Constantes ADO para aplicações Harbour |
| `rddadox.hbp` | Configuração de compilação do `hbmk2` |
| `rddadox.hbx` | Lista de símbolos externos gerada pelo Harbour |
| `comprddadow32.bat` | Compilação para ambiente Harbour 32 bits |
| `comprddadow64.bat` | Compilação para ambiente Harbour 64 bits |

## Contribuindo

Antes de enviar uma alteração:

1. Compile a biblioteca com `hbmk2.exe rddadox.hbp`.
2. Teste a alteração com o banco e o driver afetados.
3. Descreva no pull request o ambiente Harbour, a arquitetura e o driver
   utilizados.

## Apoie o projeto

Doações ajudam na manutenção e evolução do RDDADOX. Para contribuir, entre em
contato pelo e-mail `jcassianojr@gmail.com` e solicite os dados atualizados
para PIX ou PayPal.
