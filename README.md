# Notas — organização de NFS-e

Aplicação web local para importar notas de serviço emitidas e recebidas, conferir dados e exportar Excel com colunas personalizadas. Interface em português, sem serviços externos para processar os documentos. **Esta aplicação ainda não emite NFS-e:** a página **Emitir NFS-e** abre o Emissor Público Nacional para emissão com validade fiscal.

## Instalação para cada emitente

O produto está sendo preparado para uma instalação separada por CNPJ emitente. Em **Emitir NFS-e**, cadastre o emitente da instalação (razão social, CNPJ, município, código IBGE, inscrição municipal e regime tributário). Este cadastro não transmite notas nem armazena o certificado A1. A integração fiscal direta ainda depende de DPS assinada, credenciamento no Emissor Nacional e validação em homologação para o regime e serviço do emitente.

Cada instalação deve usar seu próprio banco `data/notas.sqlite3`. Em implantações separadas que usam o mesmo código, defina a variável de ambiente `NFSE_DATABASE` com um caminho absoluto diferente para cada cliente antes de iniciar. Não copie a pasta `data` de um cliente para outro nem inclua certificados ou senhas no projeto distribuído.

## Iniciar no Windows

Requer Python 3.11 ou superior. Dê dois cliques em **iniciar.bat** e abra **[http://127.0.0.1:8000](http://127.0.0.1:8000)**. A primeira execução instala as dependências pela internet. Mantenha a janela aberta; `Ctrl+C` encerra o servidor.

Alternativa pelo PowerShell, na pasta do projeto:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe app.py
```

O servidor usa Waitress e atende apenas neste computador (`127.0.0.1`). Esta versão não tem contas de usuário, autorização multiusuário ou publicação em nuvem. O acesso pela rede não está configurado.

## Usar

1. Cadastre nome, CNPJ e cidade em **Clientes**. O CNPJ permite classificar uma nota como emitida ou recebida.
2. Em **Importar notas**, escolha o cliente e envie XML, ZIP com XML/PDF ou PDF individual. Limites: 20 MB por arquivo, 60 MB por requisição, 150 arquivos e 100 MB descompactados por lote; 500 notas por XML e 30 páginas por PDF. XMLs em listas ABRASF são aceitos.
3. Abra a nota pelo número ou seta. Confira os campos e use os links para baixar os originais. Notas sem vínculo aparecem em **Sem vínculo**. Todas as notas começam pendentes, inclusive XMLs.
4. Para marcar como conferida, informe número, emissão, valor e um documento do prestador/tomador que corresponda ao cliente. Situação fiscal e conferência são campos separados.
5. Filtre por cliente, emissão/competência, período, tipo, situação, conferência e busca textual. Selecione notas específicas pelas caixas, se necessário.
6. Use **Exportar Excel** para selecionar campos, renomear cabeçalhos e ordenar com as setas. Exporte todas as notas filtradas ou somente as selecionadas. Notas pendentes são incluídas se o filtro não as excluir.
7. Salve uma configuração em **Modelos de Excel** para os próximos meses. Salvar com o mesmo nome atualiza o modelo existente.

O Excel mantém valores como números, datas como datas e CNPJs/chaves como texto. Conteúdo começando com `=` é escrito como texto literal, sem executar fórmulas. Dados ausentes ficam vazios; zero explícito continua zero. A segunda aba registra a quantidade de pendências e a origem local dos dados.

## Documentos suportados e limites

- **NFS-e nacional:** XML com `NFSe/infNFSe` e `DPS/infDPS`, com identificação das partes, datas, descrição, valores e tributos disponíveis.
- **ABRASF:** XML com `InfNfse`, nas estruturas 1.x/2.x usuais. Leiautes proprietários que não seguem essas estruturas precisam de adaptação a partir de amostras.
- **PDF:** extrai texto e tenta reconhecer campos explicitamente rotulados, como número, data e valores. O texto fica disponível na conferência. Uma nota por PDF. A extração é assistida, não universal.
- **PDF digitalizado:** o original é preservado, mas **OCR não está incluído**. Preencha os campos manualmente ou importe o XML.
- **XML e PDF da mesma nota:** vinculados automaticamente quando possuem a mesma chave de acesso nacional (50 dígitos extraídos sem espaços do PDF). O XML é a fonte prioritária. Sem chave comum, ficam separados para conferência; não há associação por suposição de nome ou valor.
- **Duplicatas:** arquivos idênticos são ignorados por cliente; XMLs são identificados pela chave ou por prestador, município, número, emissão e código de verificação. Outra versão de XML é anexada à mesma nota e exige nova conferência. Notas iguais importadas para dois clientes são registros separados, pois o vínculo pode ser diferente.
- **Situação:** não há consulta online de cancelamentos, autorização fiscal ou validação criptográfica da assinatura. O arquivo isolado não comprova a situação atual. Um cancelamento incorporado em `CompNfse` pode ser reconhecido; eventos isolados e DPS/RPS não são aceitos como notas.
- **Totais da tela:** soma dos valores informados para as notas filtradas, excluindo aquelas marcadas como canceladas. Não representa apuração tributária.
- **PDF a partir de XML:** não é gerado nesta versão. Downloads são dos originais que foram importados.

## Busca automática — pendente

A página **Integrações** registra o próximo passo para conectar o Ambiente de Dados Nacional (ADN). Não há botão que simule sincronização, coleta de credenciais ou conexão ativa. Faltam confirmar tipo e fornecedor do certificado, implementar a autenticação correspondente e validar a cobertura com clientes de Santa Rita do Sapucaí, Pouso Alegre e Itajubá/MG. Notas recebidas dependem também dos municípios de origem dos prestadores. Histórico municipal antigo pode exigir importação complementar.

Referências técnicas consultadas:

- [Documentação nacional da NFS-e](https://www.gov.br/nfse/pt-br/biblioteca/documentacao-tecnica/documentacao-atual)
- [API de distribuição para contribuintes do ADN](https://www.gov.br/nfse/pt-br/biblioteca/documentacao-tecnica/documentacao-atual/manual-contribuintes-apis-adn-sistema-nacional-nfse.pdf)

## Experimentar com documentos fictícios

Cadastre **Cliente de demonstração**, CNPJ **12.345.678/0001-95**, cidade **Santa Rita do Sapucaí / MG**. Importe os dois arquivos em `examples/`: uma nota emitida nacional e uma recebida ABRASF. São exemplos sintéticos **sem validade fiscal**, sem assinatura; não são enviados a nenhum portal. A aplicação inicia vazia e não insere dados fictícios automaticamente.

## Armazenamento e cópia de segurança

O SQLite em `data/notas.sqlite3` armazena clientes, notas, modelos, histórico e os próprios arquivos originais. Arquivos de ZIP são lidos em memória, sem extrair caminhos no disco. DTDs e entidades externas de XML são bloqueados. Escritas exigem cabeçalho da aplicação e mesma origem no navegador; os nomes de host aceitos são locais.

Para backup, **encerre o aplicativo e copie a pasta `data` inteira**. Para restaurar, mantenha o aplicativo encerrado e restaure essa pasta. O banco não possui criptografia própria. Como este projeto está em uma pasta OneDrive, os dados também podem ser sincronizados pelo OneDrive conforme a configuração do computador. Documentos não são transmitidos pelo código da aplicação.

## Desenvolvimento e verificações

Backend: Flask, SQLite, defusedxml, pypdf, openpyxl e Waitress. Frontend: HTML, CSS e JavaScript, sem Node e sem dependências por CDN.

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe -m pytest -q
```

Os testes usam banco temporário e documentos sintéticos. Cobrem leitura, vínculo, duplicatas, XML/PDF, ZIP parcial, arquivos malformados, limites, persistência, filtros, revisão e o conteúdo real do XLSX.

Para a verificação de interface com Edge instalado:

```powershell
.\.venv\Scripts\python.exe tests\browser_check.py
```

Essa verificação usa uma instância temporária na porta 8001; não adiciona dados à aplicação principal. Capturas são gravadas em `test-results/`.
