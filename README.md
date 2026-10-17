# Separador de Listas e Provas

Aplicação local em Python para identificar, separar e relacionar automaticamente páginas de **listas de treinamento** e **provas** dentro de um PDF, gerando arquivos finais organizados por colaborador, além de relatórios de conferência e uma etapa de **revisão manual assistida** para os casos em que o pareamento automático não é conclusivo.

## Contexto

Este projeto nasceu de uma necessidade real. Uma empresa recebia listas de presença e provas de treinamento misturadas em um único PDF, muitas vezes escaneado, e precisava separar e conferir cada par manualmente, página por página.

Desenvolvi a ferramenta de forma voluntária para automatizar essa rotina, e ela passou a ser usada no dia a dia da equipe responsável.

## Visão geral

O sistema foi desenvolvido para reduzir o trabalho manual envolvido na conferência de documentos de treinamentos, principalmente em cenários em que um único PDF reúne páginas de listas e provas do mesmo colaborador. A aplicação analisa o arquivo, reconhece padrões textuais, identifica códigos, separa blocos de prova, faz o pareamento com as listas correspondentes e exporta os PDFs finais automaticamente sempre que houver confiança suficiente.

Quando o sistema encontra ambiguidades, ele não força uma associação incorreta. Nesses casos, o fluxo segue para uma **revisão manual assistida**, com visualização lado a lado da lista e da prova candidata, permitindo ao usuário confirmar, corrigir ou marcar uma pendência como sem correspondência.

Além disso, a aplicação gera relatórios em **TXT**, **CSV** e **XLSX**, facilitando auditoria, validação e rastreabilidade do processamento.

## Principais recursos

- Interface gráfica local com **Tkinter**, simples e direta.
- Leitura de PDFs com **PyMuPDF**.
- OCR com **Tesseract** e **pytesseract** para PDFs escaneados ou páginas com leitura difícil.
- Identificação automática de:
  - páginas de lista;
  - início e continuação de provas;
  - códigos lógicos de documentos;
  - possíveis nomes de colaboradores.
- Pareamento automático entre lista e prova.
- Geração de PDFs individuais com **1 página de lista + bloco completo da prova correspondente**.
- Tratamento de ambiguidades por **revisão manual assistida**.
- Relatórios em texto e planilha.
- Conversão automática de CSV para **XLSX**.
- Estrutura de saída organizada por colaborador e por execução.
- Prevenção de sobrescrita de arquivos com criação de nomes alternativos automáticos.
- Possibilidade de execução com interface gráfica ou por linha de comando.
- Possibilidade de empacotamento em modo portátil para uso em ambientes sem instalação de Python na máquina final.

## Público-alvo

Este projeto é especialmente útil para equipes administrativas, RH, treinamento, SMS, qualidade ou áreas operacionais que precisam:

- separar documentos misturados em um único PDF;
- conferir listas e provas de treinamentos;
- reduzir erros de associação manual;
- manter histórico e rastreabilidade do processamento;
- gerar saídas padronizadas com mais agilidade.

## Estrutura do projeto

A estrutura real do repositório atual é a seguinte:

```text
projeto/
├─ colaboradores_local.py
├─ Colaboradores.spec
├─ environment.yml
├─ gerar_exe_pyinstaller.bat
├─ iniciar_colaboradores.bat
└─ .gitignore
```

### Estrutura criada ou esperada em execução

Durante o uso do sistema, outras pastas podem existir ou ser necessárias, mesmo que não façam parte do repositório em si. Exemplo:

```text
projeto/
├─ tesseract/
│  ├─ tesseract.exe
│  └─ tessdata/
│     └─ por.traineddata
└─ colaboradores_data/
   ├─ bruto/
   └─ processado/
```

A pasta `colaboradores_data` é utilizada pelo sistema para organizar arquivos de entrada e saída e pode ser criada automaticamente pela aplicação quando necessário.

A pasta `tesseract` pode ficar fora do repositório, mas é importante em muitos cenários de uso portátil, pois o processo de build atual empacota essa pasta quando ela estiver presente ao lado do projeto.

### Descrição dos arquivos principais

#### `colaboradores_local.py`
Arquivo principal da aplicação. Contém:

- lógica de análise e classificação de páginas;
- OCR e pré-processamento de imagem;
- identificação de listas e provas;
- pareamento automático;
- exportação dos PDFs;
- geração de relatórios;
- interface gráfica;
- modo de revisão manual assistida;
- modo de execução por linha de comando.

#### `Colaboradores.spec`
Arquivo de especificação do **PyInstaller**. Ele pode servir como base para ajustes manuais de empacotamento, referência técnica e futura padronização do processo de build.

> Observação: no fluxo atual, o script `gerar_exe_pyinstaller.bat` remove o `.spec` anterior e executa o empacotamento por linha de comando. Ainda assim, manter esse arquivo no repositório continua sendo útil como referência.

#### `environment.yml`
Arquivo de ambiente **Conda** recomendado exclusivamente para este projeto, voltado ao desenvolvimento e à geração do executável em máquina de trabalho. Ele centraliza as dependências Python do projeto e evita misturar este sistema com ambientes de outros projetos.

> Observação: o `environment.yml` prepara o ambiente Python, mas não substitui a necessidade do Tesseract disponível na pasta `tesseract` ou no `PATH`, conforme o modo de uso adotado.

#### `gerar_exe_pyinstaller.bat`
Script de empacotamento da aplicação com **PyInstaller**, voltado ao ambiente Windows. Ele:

- limpa builds anteriores;
- instala dependências mínimas para empacotamento;
- gera o executável `Colaboradores.exe`;
- inclui a pasta `tesseract` no pacote, quando ela existir;
- copia o `iniciar_colaboradores.bat` para a pasta final gerada.

#### `iniciar_colaboradores.bat`
Script simples para iniciar a aplicação. Ele tenta:

1. abrir `Colaboradores.exe`, se existir;
2. caso contrário, abrir `python\python.exe` com o script `colaboradores_local.py`, caso exista uma estrutura portátil adicional montada manualmente com Python embarcado.


## Tecnologias utilizadas

- **Python**
- **Tkinter** — interface gráfica local
- **PyMuPDF (fitz)** — leitura, renderização e extração de conteúdo de PDFs
- **pytesseract** — integração Python com o Tesseract OCR
- **Tesseract OCR** — reconhecimento óptico de caracteres
- **Pillow (PIL)** — tratamento de imagem para OCR e pré-visualização
- **openpyxl** — geração do arquivo XLSX a partir do CSV

## Requisitos para desenvolvimento

Para rodar o projeto a partir do código-fonte, recomenda-se:

- Miniconda ou Anaconda instalado na máquina de desenvolvimento;
- uso de um ambiente Conda separado e exclusivo para este projeto, definido em `environment.yml`;
- Tesseract OCR disponível:
  - na pasta `tesseract` ao lado do projeto, **ou**
  - instalado e acessível pelo `PATH` do sistema;
- Windows com suporte ao Tkinter.

### Criação do ambiente Conda

No diretório do projeto, execute:

```bash
conda env create -f environment.yml
```

Depois, ative o ambiente:

```bash
conda activate colaboradores-python
```

Se o ambiente já existir e você quiser sincronizá-lo novamente com o arquivo:

```bash
conda env update -f environment.yml --prune
```

### O que o `environment.yml` cobre

O arquivo foi preparado para centralizar um ambiente Python separado para este projeto, incluindo:

- Python;
- Tk;
- PyMuPDF;
- pytesseract;
- Pillow;
- openpyxl;
- PyInstaller.

### Observação importante sobre o Tesseract

O `environment.yml` foi pensado para o ambiente Python de desenvolvimento e build. O **Tesseract OCR** continua sendo tratado separadamente, porque o processo atual de empacotamento incorpora a pasta `tesseract` quando ela está presente ao lado do projeto.

Em outras palavras:

- o ambiente Conda prepara o Python e as bibliotecas do projeto;
- a pasta `tesseract` continua sendo a forma mais segura de garantir OCR na versão portátil gerada para outro computador.


## Requisitos para uso portátil

Para distribuição sem necessidade de instalar Python na máquina final, o caminho mais indicado neste projeto é utilizar o **PyInstaller** a partir do ambiente Conda definido em `environment.yml` e entregar a pasta gerada em `dist/Colaboradores`.

Nesse cenário, o ideal é que a estrutura final contenha:

```text
dist/
└─ Colaboradores/
   ├─ Colaboradores.exe
   ├─ iniciar_colaboradores.bat
   └─ tesseract/
      ├─ tesseract.exe
      └─ tessdata/
         └─ por.traineddata
```

### Observação importante

A distribuição portátil padrão deste projeto é baseada no executável `Colaboradores.exe`.

O fallback via `python\python.exe`, previsto no `iniciar_colaboradores.bat`, só funcionará se uma estrutura portátil complementar com Python embarcado for preparada manualmente, o que não faz parte obrigatória da saída padrão do build atual.


## Configuração do Tesseract

A aplicação tenta localizar o Tesseract da seguinte forma:

1. primeiro, procura por `tesseract\tesseract.exe` ao lado do programa;
2. se não encontrar, tenta localizar o executável `tesseract` no `PATH` do sistema;
3. se não encontrar nenhum dos dois, interrompe a execução com erro.

### Estrutura esperada

```text
tesseract/
├─ tesseract.exe
└─ tessdata/
   └─ por.traineddata
```

### Observação importante

Como o idioma padrão de OCR configurado no projeto é **português** (`por`), o arquivo `por.traineddata` precisa estar presente na pasta `tessdata`.

## Como executar em modo desenvolvimento

### Opção 1 — Interface gráfica

No diretório do projeto, execute:

```bash
python colaboradores_local.py
```

A interface será aberta e permitirá:

- selecionar um PDF;
- fazer pré-análise;
- processar o arquivo;
- abrir a revisão manual;
- abrir a pasta base;
- abrir a pasta da última saída gerada.

### Opção 2 — Linha de comando

Também é possível processar um PDF diretamente via terminal:

```bash
python colaboradores_local.py "caminho_do_arquivo.pdf"
```

Nesse modo, a aplicação processa o documento e imprime no console:

- progresso da execução;
- colaborador identificado;
- quantidade de arquivos gerados;
- pasta de saída;
- caminhos dos relatórios principais.

## Como gerar a versão portátil com PyInstaller

### 1. Ativar o ambiente do projeto

No Windows, dentro da pasta do projeto, ative o ambiente:

```powershell
conda activate colaboradores-python
```

### 2. Executar o script de build

Depois, execute:

```powershell
.\gerar_exe_pyinstaller.bat
```

### O que esse script faz

- remove pastas de build antigas;
- remove o arquivo `.spec` anterior;
- atualiza o `pip`;
- garante as dependências mínimas de empacotamento;
- gera o executável com janela gráfica;
- exclui módulos desnecessários do empacotamento para reduzir o pacote;
- adiciona a pasta `tesseract` ao build, caso ela exista;
- copia o `iniciar_colaboradores.bat` para a pasta final.

### Saída esperada

Após a conclusão, a pasta principal gerada será:

```text
dist/Colaboradores
```

Dentro dela deve estar o executável `Colaboradores.exe`.

### Recomendação prática

Para manter o build mais previsível entre diferentes máquinas, recomenda-se gerar o executável sempre com o ambiente `colaboradores-python` ativo e manter a pasta `tesseract` ao lado do projeto quando a intenção for distribuir uma versão realmente portátil.


## Como iniciar a aplicação empacotada

Dentro da pasta final gerada, é possível abrir:

- `Colaboradores.exe`, diretamente; ou
- `iniciar_colaboradores.bat`, que tenta localizar e iniciar automaticamente a melhor opção disponível.

## Fluxo de uso da aplicação

### 1. Selecionar o PDF

Use o botão **Selecionar PDF** para escolher o arquivo a ser processado.

### 2. Copiar para a pasta bruto

Opcionalmente, use **Copiar para pasta bruto** para armazenar uma cópia do PDF original em:

```text
colaboradores_data/bruto/
```

Isso ajuda na organização e preserva um local padrão para os arquivos de entrada.

### 3. Fazer a pré-análise

A opção **Pré-análise** informa:

- nome do arquivo;
- quantidade de páginas;
- se o PDF provavelmente é escaneado.

Essa etapa é útil para uma conferência rápida antes do processamento completo.

### 4. Processar o documento

Ao clicar em **Processar**, o sistema executa as etapas abaixo:

1. análise de todas as páginas;
2. leitura textual e OCR quando necessário;
3. identificação de listas;
4. identificação de blocos de prova;
5. definição do colaborador principal;
6. criação da estrutura de saída;
7. pareamento entre listas e provas;
8. exportação automática dos pares confiáveis;
9. geração de relatórios.

Ao final, o sistema informa na tela:

- colaborador identificado;
- quantidade de listas encontradas;
- quantidade de provas encontradas;
- PDFs gerados automaticamente;
- pendências para revisão manual;
- nome dos relatórios gerados;
- pasta de saída da execução.

### 5. Revisar manualmente os casos pendentes

Se houver casos não resolvidos automaticamente, use o botão **Revisão manual**.

A revisão manual assistida permite:

- navegar entre as pendências;
- visualizar a página da lista e a prova candidata lado a lado;
- ajustar o código final, se necessário;
- confirmar uma prova candidata;
- marcar um caso como sem correspondência;
- gerar todos os PDFs confirmados manualmente.

Esse recurso foi pensado para manter a produtividade sem sacrificar a confiabilidade do resultado final.

## Estrutura de saída gerada

Cada execução cria uma pasta específica dentro de:

```text
colaboradores_data/processado/
```

A estrutura segue o padrão:

```text
colaboradores_data/
└─ processado/
   └─ NomeDoColaborador/
      └─ AAAAMMDD_HHMMSS/
         ├─ pdfs/
         │  ├─ NomeDoColaborador_CODIGO_1.pdf
         │  ├─ NomeDoColaborador_CODIGO_2.pdf
         │  └─ ...
         └─ relatorios/
            ├─ relatorio.txt
            ├─ resumo.csv
            └─ resumo.xlsx
```

### Organização por colaborador

O sistema tenta definir um nome principal de colaborador a partir de:

1. nome do arquivo de entrada;
2. nomes extraídos das provas;
3. nomes extraídos das páginas analisadas.

Com isso, a pasta do processamento e os arquivos gerados ficam mais organizados.

## Nome dos arquivos gerados

Os PDFs exportados seguem o padrão:

```text
NomeDoColaborador_IDENTIFICADOR.pdf
```

Exemplo conceitual:

```text
JoaoExemplo_PE-AB-12345.pdf
```

Quando já existe um arquivo com o mesmo nome, o sistema evita sobrescrever e passa a gerar:

```text
JoaoExemplo_PE-AB-12345_2.pdf
JoaoExemplo_PE-AB-12345_3.pdf
```

## Relatórios gerados

A aplicação gera três formatos principais de relatório.

### `relatorio.txt`
Relatório textual detalhado da execução, contendo, entre outros dados:

- arquivo de entrada;
- colaborador identificado;
- total de páginas;
- listas encontradas;
- provas encontradas;
- quantidade de pares automáticos e manuais;
- pasta de saída;
- agrupamento por situação de processamento;
- observações, avisos e erros.

### `resumo.csv`
Resumo tabular da execução, ideal para conferência rápida e integração com planilhas.

### `resumo.xlsx`
Versão em Excel gerada automaticamente a partir do CSV, facilitando leitura, filtro e compartilhamento.

## Lógica geral de identificação

Sem entrar em todos os detalhes internos do algoritmo, o sistema utiliza uma combinação de:

- leitura textual de páginas;
- OCR em regiões estratégicas;
- identificação de termos característicos de lista;
- identificação de termos característicos de prova;
- extração de códigos lógicos;
- validações por proximidade e compatibilidade;
- heurísticas para início e continuação de blocos de prova.

Essa abordagem híbrida aumenta a robustez do processamento em documentos reais, nos quais podem existir:

- digitalizações imperfeitas;
- páginas com baixa qualidade;
- nomes incompletos;
- códigos parcialmente legíveis;
- anexos ou pequenas variações de layout.

## Situações de processamento

Os pares podem seguir diferentes estados de fluxo, como:

- **OK** — pareamento automático confiável;
- **OK_BASE_CODE** — pareamento automático confiável com base no código-base;
- **MANUAL_CONFIRMADO** — confirmação manual realizada;
- **MANUAL_SEM_CORRESPONDENCIA** — item marcado manualmente como sem correspondência;
- **PENDENTE_REVISAO_MANUAL** — item ainda aguardando análise manual.

Essas situações também aparecem nos relatórios, permitindo controle completo do processamento.

## Benefícios do projeto

- Redução do tempo gasto com conferência manual.
- Diminuição de erros humanos no pareamento de documentos.
- Organização padronizada da saída.
- Rastreabilidade por execução.
- Possibilidade de intervenção humana apenas quando necessário.
- Melhor visualização gerencial por meio dos relatórios em planilha.

## Limitações conhecidas

Como qualquer sistema baseado em leitura de documentos reais e OCR, alguns cenários podem exigir atenção adicional:

- PDFs extremamente degradados podem reduzir a qualidade da leitura;
- páginas muito inclinadas, cortadas ou borradas podem afetar a extração;
- códigos ausentes ou ilegíveis podem exigir revisão manual;
- layouts muito diferentes do padrão esperado podem demandar ajustes futuros na lógica.

Ainda assim, o sistema foi estruturado para não comprometer a confiabilidade: quando a automação não alcança segurança suficiente, o fluxo é encaminhado para validação manual.

## Boas práticas de uso

- Manter a pasta `tesseract` junto ao programa na versão portátil.
- Garantir que `por.traineddata` esteja presente em `tessdata`.
- Trabalhar com cópias dos PDFs originais quando possível.
- Utilizar a pasta `bruto` como repositório de entrada organizado.
- Conferir o `resumo.xlsx` após cada processamento.
- Executar a revisão manual sempre que houver pendências.
- Preservar a pasta de saída de cada execução para fins de auditoria.

## Solução de problemas

### O programa informa que o Tesseract não foi encontrado
Verifique se existe a estrutura:

```text
tesseract/tesseract.exe
tesseract/tessdata/por.traineddata
```

Caso prefira, também é possível deixar o Tesseract instalado e disponível no `PATH` da máquina de desenvolvimento, desde que o ambiente `colaboradores-python` esteja ativo quando o projeto for executado.

### O executável não abre
Verifique se:

- a pasta gerada pelo PyInstaller foi criada corretamente;
- o `Colaboradores.exe` está presente em `dist/Colaboradores`;
- a pasta `tesseract` foi incluída quando necessária;
- o antivírus ou políticas da máquina não bloquearam a execução.

### O PowerShell não reconhece o arquivo `.bat`
No Windows PowerShell, execute informando o caminho relativo do arquivo.

Use:

```powershell
.\gerar_exe_pyinstaller.bat
```

ou:

```powershell
.\iniciar_colaboradores.bat
```

### O sistema processou, mas alguns arquivos ficaram pendentes
Isso significa que a automação identificou necessidade de validação humana. Nessa situação, abra a opção **Revisão manual** e conclua a associação dos casos pendentes.

### Os nomes dos arquivos ficaram repetidos
O sistema já possui proteção contra sobrescrita. Quando necessário, ele cria automaticamente versões com sufixos como `_2`, `_3` e assim por diante.

## Segurança e privacidade

A aplicação foi concebida para uso local, sem necessidade de envio do PDF para serviços externos. Isso favorece o uso em ambientes corporativos com maior restrição de rede, confidencialidade ou instalação.

## Possibilidades de evolução

Como próximos passos, o projeto pode evoluir para incluir, por exemplo:

- melhorias adicionais de OCR;
- suporte ampliado a layouts diferentes;
- exportações mais avançadas;
- logs estruturados por execução;
- pacote ainda mais simplificado para distribuição corporativa;
- futura migração para uma arquitetura mais modular, se desejado.

## Autor

**Gilleade Gonçalves Silva**  
E-mail: gilleadesilva@gmail.com  
LinkedIn: https://br.linkedin.com/in/gilleade-silva  
GitHub: https://github.com/Gilleade

## Licença e uso

Este projeto ainda não possui uma licença formal definida no repositório.

Até que uma licença específica seja adicionada por meio de um arquivo `LICENSE`, recomenda-se considerar o código e os materiais deste projeto como de uso controlado pelo autor, especialmente em contextos de distribuição, modificação, reutilização ou uso corporativo.

Caso o projeto venha a ser publicado ou distribuído de forma mais ampla, é recomendável incluir uma licença apropriada ao objetivo pretendido.

## Observação final

Este README foi elaborado com base nos arquivos operacionais atuais do projeto, especialmente o script principal `colaboradores_local.py`, o arquivo `Colaboradores.spec`, o novo `environment.yml`, os scripts de inicialização e empacotamento e a estrutura de uso esperada para OCR e geração de saídas. Ele pode ser adaptado futuramente para refletir novas funcionalidades, reorganização da arquitetura, formalização adicional do ambiente de desenvolvimento ou novas formas de distribuição.
