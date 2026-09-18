# IMOBILIARIA
Projeto SIAI

## Abertura de Cad_Recebimento e RelRecebimento (18/09/2026)

`EnsureRuntimeFields` deixa de executar a descoberta de campos carregando
registros nas consultas locais desses dois formularios. Para cada consulta MySQL
baseada em `TZAbstractRODataset`, incluindo `TZQuery` e `TZReadOnlyQuery`, uma
consulta privada usa o SQL original com `LIMIT 0`, copia somente as definicoes
de campos e e descartada. O SQL, parametros, eventos, filtros, cursores e estado
de todos os datasets, inclusive os globais de `DM_Tabelas`, permanecem
inalterados. SQLs nao elegiveis e qualquer falha usam imediatamente o caminho
legado `FieldDefs.Update`.

O diagnostico de `RelRecebimento` agora separa preparacao dos lookups, abertura
do filtro temporario, abertura de empreendimentos e termino da preparacao do
formulario. Nenhuma consulta foi antecipada, fechada ou movida; as novas linhas
apenas medem as chamadas existentes. Nenhum form foi retirado do AutoCreate e
nenhum DFM foi alterado.

Validacao em MySQL descartavel: 32 testes de metadados, 16 de cobranca e 7 de
negociacao aprovados. Release e Debug completos compilaram; executaveis em
`build/startup-audit/rel-receb-ro-metadata-release` e
`build/startup-audit/rel-receb-ro-metadata-debug`.

Teste manual com `/diagnostico`: abrir Recebimentos, fechar e abrir novamente;
depois usar Relatorios da propria tela e aguardar o formulario de filtros. O
log deve registrar as etapas `TFrmCad_Recebimento.*: metadados` e
`RelRecebimento:*`, permitindo medir a base real sem expor SQL ou dados.

## Cobranca: totalizador sem lookups irrelevantes (18/09/2026)

Terceiro lote da cobranca. O totalizador preserva ZQRecebimento4 aberto para o
grid, mas percorre uma TZQuery privada que reutiliza o mesmo filtro e traz
somente venda_idvenda, Dt_Vencimento e saldo. Assim, a soma nao aciona os
campos lookup que nao participam de multa, mora ou valor corrigido. O loop usa
a mesma rotina de formulas e arredondamentos dos campos calculados do dataset.

As taxas continuam sendo obtidas por titulo, em outra TZQuery privada do
formulario. ZQVenda_cobr, os datasets globais e a negociacao nao sao alterados
pelo totalizador. O estado e restaurado mesmo se ocorrer erro; a protecao
existente tambem limpa totais parciais e a negociacao antes de propagar a falha.

Validacao em banco MySQL descartavel: 16 cenarios de cobranca e 7 de negociacao
aprovados. Inclui calculo unico por titulo, totalizador sem lookups, contratos
com taxas diferentes, normal/Price, filtros, bancos, relatorio, pesquisa vazia,
Post e Delete. Os executaveis desta revisao serao gerados em
build/startup-audit/cobranca-totais-raw-release e
build/startup-audit/cobranca-totais-raw-debug.

Teste manual apos Build no Delphi com /diagnostico: pesquisar um cliente com
muitos titulos, conferir quantidade, saldo, multa, mora e corrigido; gerar a
negociacao e os dois relatorios. Enviar o log novo para medir este lote. Nenhum
form foi retirado do AutoCreate e nenhum DFM foi alterado neste lote.

## Cobranca: navegacao e lookup conjunto (17/09/2026)

Segundo lote: a validade dos totais usa participante + filtro de loteamento +
modo normal/Price + data, em vez do identificador da linha selecionada. Os
 totais ja usam taxas por titulo desde o lote anterior. Antes de reutilizar
um resultado, Open/Refresh/Post/Delete continuam invalidando a cobranca.
Mudar de titulo do mesmo participante nao refaz totais nem apaga negociacao.
Mudar participante, filtro, modo ou data continua solicitando atualizacao.
Uma alteracao externa no banco requer atualizar/reabrir a consulta, como no
fluxo existente; esta chave nao e um mecanismo de sincronizacao entre usuarios.

Nome e CPF sao obtidos em uma unica chamada Lookup com dois campos de retorno,
sem cache persistente. Foram preservados o nome alternativo do titulo antigo,
documentos nulos e fontes ausentes/fechadas. Nenhum SQL/filtro/evento/cursor dos
datasets globais foi modificado pela otimizacao. O diagnostico passa a mostrar
"lookup participante/nome e CPF juntos" e registra se uma atualizacao foi
motivada por dados invalidados ou mudanca no contexto. Esses motivos nao
incluem identificadores ou dados pessoais.

Validacao: 16 casos de cobranca e 7 de negociacao aprovados em banco descartavel.
Inclui duas vendas do mesmo cliente com taxas distintas: normal 433,00 e Price
435,50, iguais no grid e totalizador. Navegar entre contratos preserva o total
e a negociacao, mas atualiza a venda e o banco da linha. Post e Delete forcam
novos totais. Alterar/reabrir a origem do nome/CPF aparece na proxima leitura;
nulos e cadastros ausentes nao herdam dados da linha anterior. O contador de
Lookup confirma um acesso ao participante por preenchimento, e o cursor global
permanece na mesma posicao. O relatorio real foi preparado com dois titulos e
os respectivos bancos. Nao houve impressao fisica ou teste na base de producao.

Logs: build/startup-audit/metadata-a63d0713ad6f41f7bef59860ca219092/tests.log
 e build/startup-audit/metadata-41f3d32fc2a64b01a8bc6644894286f1/tests.log.
Executaveis de validacao: build/startup-audit/cobranca-lookups-release/SIAI.exe
 e build/startup-audit/cobranca-lookups-debug/SIAI.exe.
Permanece o aviso de encerramento my_thread_global_end no teste com Fortes.
Nenhum form foi retirado do AutoCreate. Nenhum DFM foi alterado neste lote.

Teste manual apos Build no Delphi com /diagnostico: pesquisar cliente com
varios contratos, selecionar parcelas e gerar negociacao. Mudar de linha,
sair e voltar ao grid; conferir preservacao dos totais e parcelas. Repetir
com outro cliente, filtro e modo Price. Conferir nome/CPF e bancos nos dois
relatorios. O ganho de tempo real depende do proximo log do usuario.

## Cobranca: consistencia dos totais e pesquisa unica (17/09/2026)

Primeiro lote apos a medicao detalhada. Apenas Cad_Recebimento.pas foi alterado
no codigo do aplicativo; nenhum DFM, global do DM ou AutoCreate foi modificado
neste lote. Os lookups e a preparacao do relatorio continuam para lotes seguintes.

O teste com dois contratos do mesmo participante reproduziu uma divergencia
anterior: soma do grid 433,00 contra totalizador 448,00 no modo normal;
435,50 contra 446,00 no modo Price. O totalizador utilizava taxas de uma unica
venda e, no Price, a taxa do registro atual do grid. Agora uma TZQuery privada,
pertencente ao formulario, consulta as taxas da venda de cada titulo. A consulta
reutiliza somente a ultima venda durante a operacao e e invalidada no inicio
da atualizacao dos totais. Nao altera ZQVenda_cobr, usado pela negociacao.
As formulas e arredondamentos permanecem; a origem das taxas foi corrigida.
Venda ausente interrompe o calculo e limpa os totais parciais e a negociacao.

A pesquisa nao executa mais o recalc completo anterior ao Refresh.
FinalizarPesquisaCobranca executa Refresh e a atualizacao completa explicitamente;
o OnEnter seguinte reconhece a atualizacao concluida e nao repete a operacao.
O vinculo bancario original do DFM e mantido. Nao foi removido o Refresh.

Validacao em MySQL descartavel, porta 33557, somente dados ficticios:
- 12 casos CobrancaPerformanceSmoke aprovados, incluindo contratos distintos,
  normal/Price, atualizacao de taxa, filtro por loteamento, pesquisa vazia,
  erro/recuperacao, preservacao de negociacao e relacao bancaria original.
- O DFM real de RelCobranca2 gerou paginas com dois titulos; o teste verificou
  os valores e os bancos durante a geracao. Arquivo de teste em
  build/startup-audit/metadata-tests/cobranca-mixed-test.rpf. Nao houve impressao
  fisica nem validacao visual completa do layout.
- 7 casos de NegotiationSmoke aprovados, incluindo selecao e mascaras.
- SIAI completo compilado em Release e Debug, em build/startup-audit/
  cobranca-consistency-release e cobranca-consistency-debug. Release: 1413
  warnings; Debug: 1414, incluindo W1036 em XBanner.pas de componente externo.
- Log principal: build/startup-audit/metadata-2c01ef015c7345e38571c16baf3a4159/tests.log.
- Regressao: build/startup-audit/metadata-4cd8dcb367314aa9b76ce52b93f1be13/tests.log.
- A biblioteca MySQL antiga ainda emite no encerramento do teste com Fortes
  o aviso my_thread_global_end sobre duas threads. Assertions aprovadas, exit 0.

Roteiro manual: encerrar o SIAI, executar Build no Delphi e iniciar com
/diagnostico. Na aba Cobranca, pesquisar por participante e quadra/lote,
com/sem filtro de loteamento; conferir quantidade, saldo, multa, mora e
corrigido contra as linhas. Repetir normal e Price com contratos de taxas
diferentes. Selecionar parcelas, gerar negociacao, voltar ao grid e abrir os
dois relatorios. Conferir banco/agencia/conta nos diferentes loteamentos.
O ganho de tempo em producao depende do novo log; nao foi medido neste lote.

Limites: a regra existente para juros futuros de uma negociacao que mistura
contratos nao foi redefinida. Os testes de valores mistos cobrem os titulos
vencidos, seus totais e o relatorio do grid. A tabela auxiliar compartilhada
cobranca e o filtro legado dos totais tambem permanecem para analise separada.

## Medicoes detalhadas da Cobranca (17/09/2026)

Com `/diagnostico`, o log agora separa os lookups de nome, CPF e loteamento,
as formulas financeiras e as operacoes First, Next, Locate, EnableControls
e recalculo explicito. Cada grupo informa quantidade de chamadas e tempo
acumulado desde a ultima gravacao do log. As medidas de navegacao/recalculo
incluem eventos internos e podem conter os lookups: nao somar pais e filhos.
A resolucao e de milissegundos; operacoes muito curtas podem aparecer como zero.

Tambem foram separados abertura da pesquisa, carga dos titulos selecionados,
dados bancarios, Refresh, consulta de saldo/quantidade, DROP, CREATE/copia,
abertura dos titulos dos totais, soma dos totais e redesenho do grid.
Os dois relatorios de Cobranca registram o intervalo BeforePrint/AfterPrint
do Fortes, correspondente a geracao de paginas, preservando os eventos
anteriores. Esse intervalo nao inclui a espera do usuario na visualizacao
nem mede o tempo fisico da impressora. A preparacao dos dados do relatorio
do grid possui sua propria medicao.

Nao foram alteradas formulas, filtros, comandos SQL, vinculos de datasets ou
globais nesta instrumentacao. Nenhum DFM alterado; nenhum form retirado do
AutoCreate. Sem `/diagnostico`, nao sao instalados os eventos de medicao nos
relatorios e as funcoes de registro retornam sem produzir log.

Validacao: Release e Debug compilados. Seis cenarios de calculo/desempenho e
encadeamento dos eventos de relatorio passaram com diagnostico ligado e
desligado; os sete cenarios anteriores de negociacao/mascaras tambem passaram.
O teste de relatorio verifica os callbacks e a preservacao do cancelamento,
sem imprimir documentos. A DLL MySQL emitiu `my_thread_global_end(): 2 threads
didn't exit` no encerramento dos dois testes que instanciam o relatorio;
ambos terminaram com codigo zero e todas as verificacoes aprovadas. Esse aviso
de encerramento do ambiente de teste nao foi corrigido neste lote.

Reproduzir com `tests/RunMetadataTests.ps1 -Cobranca -Diagnostic` e sem
`-Diagnostic`. Resultados em `build/startup-audit/metadata-8a6515a573d54054b235032931ea6974`
e `build/startup-audit/metadata-51acf05705c9433891427e6dd852b614`.

Para medir a base real: fazer Build no Delphi, executar com `/diagnostico`,
abrir Cobranca, pesquisar/selecionar um cliente, aguardar o grid e abrir a
visualizacao do relatorio. Depois informar que concluiu para leitura do log.

## Desempenho da aba Cobranca: primeiro lote (17/09/2026)

Alteracao restrita a `Cad_Recebimento.pas`. O calculo dos campos reutiliza
`ZQVenda_cobr` somente enquanto a venda carregada tem o mesmo ID. O proprio
dataset local continua disponivel para a negociacao e os demais leitores;
nenhum dataset global foi substituido, liberado ou teve seu SQL alterado.
Antes de abrir/atualizar os titulos e apos Post/Delete, o estado reutilizavel
e invalidado. Os quatro eventos estavam sem manipuladores no DFM, verificado
tambem no teste. As formulas financeiras existentes foram preservadas.

O retorno ao grid reaproveita os totais apenas depois de uma atualizacao
concluida e quando titulo, filtro de loteamento, modo de mora e data continuam
iguais. Isso tambem evita apagar a negociacao apenas ao retornar ao grid.
Reabrir a pesquisa/aba ou atualizar os titulos recarrega os dados; mudancas
feitas por outro usuario no banco nao sao consultadas a cada simples foco.

Com `/diagnostico`, o log registra consultas de venda, tempo de recalculo dos
titulos, tempo de totais/recriacao da tabela e retornos ao grid sem recarga.
O ganho em segundos no ambiente real ainda precisa de medicao. A recriacao
da tabela `cobranca`, os lookups globais e a pesquisa ampla permanecem como
pontos para um proximo lote, apos validar este.

Validacao: Release e Debug compilados; 1.413 avisos W em Release, mesma
quantidade anterior. Cinco cenarios de desempenho passaram em MySQL
descartavel, mais os sete testes anteriores de negociacao e mascaras.
Vinte ciclos de navegacao na mesma venda acrescentaram zero consultas de
venda; dez retornos ao grid acrescentaram zero consultas de venda/totais.
Conferidos multa, mora, valores corrigidos, modo Price, alteracao de taxa
seguida de Refresh, troca de titulo/cliente, preservacao da negociacao e
permanencia dos lookups globais ativos. Nenhum acesso a base real nos testes.

Reproducao: `tests/RunMetadataTests.ps1 -Cobranca`. Resultado:
`build/startup-audit/metadata-79c947b966f74a5cb8963bbfa87c2b03/tests.log`.
Regressao: `tests/RunMetadataTests.ps1 -Negotiation`, resultado em
`build/startup-audit/metadata-37c3588d74414eaa8ab5644a4e014b3f/tests.log`.

Teste manual: fazer Build no Delphi, manter `/diagnostico`, abrir
Recebimentos > Cobranca, navegar entre titulos e alternar foco entre campos e
grid. Conferir multas, mora e totais; negociar uma e varias parcelas; retornar
ao grid e conferir que a negociacao nao desaparece so pelo foco. Pesquisar
outro cliente/loteamento, testar com e sem Price, fechar/reabrir a aba e
conferir os dados. Visualizar os relatorios usados normalmente. Informar
qual acao ainda demora para comparar as novas fases do log.

Nenhum formulario retirado do AutoCreate neste lote. Nenhum DFM modificado.
Executaveis: `build/startup-audit/cobranca-performance-release/SIAI.exe` e
`build/startup-audit/cobranca-performance-debug/SIAI.exe`.

## Correcao: negociacao de parcelas e pesquisa de cobranca (16/09/2026)

Corrigido em `Cad_Recebimento.pas` o caminho com "Negociacao por parcelas"
marcado: sem linhas explicitamente selecionadas, a rotina agora limpa a saida
anterior, informa como selecionar com Ctrl+clique e interrompe a geracao.
Tambem bloqueia uma selecao sem registros ou com vencimento/saldo ausente ou
invalido. Vencimentos nulos sao preservados como nulos na copia para o dataset
auxiliar, permitindo a validacao sem erro de conversao. As formulas existentes
de juros, multa e consolidacao nao foram alteradas.

Em `PesqCobranca.pas`, Valor e Saldo receberam a mascara `###,###,##0.00`.
Nenhum DFM foi alterado nesta correcao; nenhum form foi retirado do AutoCreate.

Validacao: SIAI completo compilado em Release e Debug. Sete testes passaram
em MySQL descartavel com dados ficticios: falta de selecao, origem fechada,
uma parcela, duas parcelas, vencimento nulo, modo sem selecao por parcelas,
mascaras monetarias e sua preservacao ao reabrir. O teste usa o DFM e os eventos
reais, mas substitui as consultas por dados ficticios e desliga os eventos de
abertura/conexao da aplicacao. A base real nao foi acessada.

Reproducao: `tests/RunMetadataTests.ps1 -Negotiation` executa
`tests/NegotiationSmoke.dpr`. Resultado desta rodada:
`build/startup-audit/metadata-fe6dc89388be4284bfbb6c79845cb612/tests.log`.

Teste manual: recompilar no Delphi, abrir Recebimentos > Cobranca e, com
"Negociacao por parcelas" marcado, tentar gerar sem selecao (deve avisar,
sem linha de 1899). Selecionar uma e depois varias linhas com Ctrl+clique,
informar o vencimento e sair do campo; conferir datas e totais. No botao
Pesquisar abaixo dessa opcao, conferir Valor e Saldo com duas casas decimais.

## Desempenho: lote 3 (16/09/2026)

Nenhum form retirado do AutoCreate neste lote. A lista continua com 18 objetos.
O diagnostico anterior mediu 33,7 s ate o login: 24,8 s no modulo de dados,
8,1 s nos forms auxiliares e aproximadamente 0,55 s nas consultas iniciais e
na atualizacao financeira. O ganho deste lote ainda depende de nova medicao
com a base do usuario.

`uSiaiMetadata.pas` prepara uma TZQuery temporaria com o SELECT original seguido
de `LIMIT 0`, exclusivamente para descobrir a estrutura. Copia as definicoes
dos campos e descarta a consulta temporaria. O SQL, parametros, eventos e estado
de preparacao da consulta original sao preservados; sua abertura posterior
continua retornando registros. Nenhum LIMIT e acrescentado ao SQL das telas.

A chamada foi limitada em `uRuntimeFields.pas` aos proprietarios medidos:

- `TDM_Tabelas`;
- `TFrmPesqRecebimento_bx`;
- `TFrmPesqRecebimento2`;
- `TFrmPesqEndereco`;
- `TFrmPesqCobranca`;
- `TFrmPesqRecebimento`.

Somente TZQuery MySQL recebe a otimizacao. Consultas com LIMIT existente, UNION,
comentarios, multiplos comandos, clausulas de bloqueio, INTO, variaveis ou
formas ambiguas continuam no caminho anterior. Consultas ligadas a DataSource,
ou parametrizadas com MasterSource, tambem conservam o caminho anterior.
Se a consulta temporaria falhar, a descoberta original e utilizada. A criacao
dos campos, suas classes adaptadas ao driver, lookups e ordem de abertura foram
mantidas. O UPDATE financeiro de inicializacao nao foi alterado.

Validacao:

- SIAI completo compilado em Release e Debug; 1.413 avisos W na configuracao
  Release, mesma quantidade do lote 2.
- 20 testes anteriores passaram (`build/startup-audit/lot3-startup-tests-run.log`).
- 30 testes de metadados passaram em MySQL 5.5.16 local descartavel com dados
  ficticios (`build/startup-audit/lot3-metadata-tests-run.log`).
- Comparadas definicoes de campos: nome, tipo, tamanho, precisao, obrigatoriedade
  e atributos. Exercitados joins com colunas repetidas, agregacoes, parametros,
  resultados vazios, abertura, reabertura, campos persistentes e fallback.
- O teste integrado comparou a preparacao anterior com a otimizada, incluindo
  classes dos campos, evento de calculo, lookup nativo, comportamento de lookup
  calculado, abertura automatica, reutilizacao dos campos e edicao/Post.
  O lookup calculado da fixture ficou vazio em ambos os caminhos; esse caso
  valida equivalencia, nao confirma todos os lookups reais do sistema.
- Confirmado: o SELECT temporario retorna zero registros, enquanto a consulta
  original retorna os registros esperados antes e depois de reabrir.
- A base real do SIAI nao foi utilizada pelos testes automatizados.

Para reproduzir os testes novos: executar `tests/RunMetadataTests.ps1`. O script
exige o MySQL 5.5 e Delphi nos caminhos instalados nesta maquina, recusa uma porta
33557 ocupada, cria dados ficticios em pasta exclusiva, compila o teste e encerra
o servidor descartavel. Logs permanecem em `build/startup-audit/metadata-*`.
Os testes anteriores continuam em `tests/RunStartupTests.bat`.

Roteiro manual deste lote:

1. No Delphi, executar Build do projeto SIAI e manter `/diagnostico` em
   Run > Parameters. Medir ate aparecer o login e comparar com os 33,7 s.
2. Entrar no sistema e testar pesquisas de recebimentos, baixas, cobranca e
   enderecos: buscar registros conhecidos, selecionar, cancelar e reabrir.
3. Conferir nomes de clientes, descricoes, valores e datas nas pesquisas;
   testar resultados vazios e com varios registros.
4. Abrir Vendas, Participantes, Imoveis e Configuracoes, que usam o modulo
   de dados compartilhado. Em homologacao, validar uma edicao e gravacao.
5. Informar quando o login aparecer para comparar o novo log de inicializacao.

Executaveis gerados: `build/startup-audit/lot3-release/SIAI.exe` e
`build/startup-audit/lot3-debug/SIAI.exe`. Ao testar pelo Delphi, recompilar o
projeto; nao utilizar um executavel antigo apenas porque o nome e igual.

## Desempenho: lote 2 (16/09/2026)

Lote 1 validado pelo usuario sem problemas relatados. O lote 2 reduz o autocreate
de 26 para 18 objetos (11 retirados no total desde a lista original de 29).

| Form retirado neste lote | Unit | Funcao que garante a criacao |
| --- | --- | --- |
| `Frm_AchaParticipante` | `AchaParticipante.pas` | `AchaCliente` |
| `Frm_AchaCidade` | `AchaCidade.pas` | `Acha_Cidade` |
| `Frm_AchaCorretor` | `AchaCorretor.pas` | `AchaCorretor` |
| `Frm_AchaPlanoDeContas` | `AchaPlanoDeContas.pas` | `AchaPlanoDeConta` |
| `Frm_Acha_Contabancaria` | `Acha_Contabancaria.pas` | `Acha_Contabancaria` |
| `FormMensagem` | `Finanmsg.pas` | `Mensagem` |
| `FrmPerguntaSIMNAO` | `PerguntaSIMNAO.pas` | `SIMNAO` |
| `FrmPergunta` | `Pergunta.pas` | `pergunta` |

As oito criacoes ficam antes do primeiro acesso aos controles em `funcoes.pas`.
Cada form continua global, e criado com `Application` como proprietario e e
reutilizado ate o encerramento. Nada e liberado ao retornar do dialogo: os controles
continuam disponiveis para a leitura da resposta. As consultas, retornos e regras
de negocio foram preservados. Nenhum DFM foi modificado neste lote.

Verificacao de dependencias: as referencias ativas do SIAI a essas instancias
estao centralizadas em `funcoes.pas`; nao foram encontradas referencias a seus
componentes a partir de outros DFM. Os oito forms nao possuem datasets proprios.
Uma referencia em `CopiaAutomatica.pas` pertence a outro fluxo/projeto e seu
autocreate nao foi alterado.

Validacao executada:

- SIAI completo compilado em Release e Debug; mesmos avisos na comparacao Release
  com o lote 1.
- 20 testes isolados passaram: os 9 anteriores, 6 casos de primeiro uso/reabertura
  dos dialogos e 5 testes de construcao dos DFM das pesquisas.
- As respostas SIM/NAO e de texto foram verificadas depois do fechamento;
  a reabertura reutiliza a mesma instancia. Enter e Escape foram exercitados.
- Comparacao do codigo confirmou que `funcoes.pas` mudou somente pelas oito
  guardas de criacao, sem alteracoes nas consultas.
- As consultas das pesquisas nao foram executadas contra um banco de dados.

Executavel de teste: `build\startup-audit\lot2-release\SIAI.exe` (usar na
instalacao de homologacao com INI/DLLs/recursos corretos). Resultados dos testes:
`build\startup-audit\lot2-tests-run.log`. O script `tests\RunStartupTests.bat`
recompila e executa a suite atual de 20 casos.

Roteiro manual do lote 2:

1. Pesquisar participante/cliente em Vendas e Recebimentos.
2. Pesquisar cidade no cadastro de participante ou empresa.
3. Pesquisar corretor em Vendas.
4. Pesquisar plano de contas em Recebimentos e Configuracoes.
5. Pesquisar conta bancaria nas opcoes que utilizam essa selecao.
6. Para cada pesquisa, testar nenhum, um e varios resultados; selecionar,
   cancelar e reabrir. Usar termos amplos para realmente abrir o dialogo de selecao.
7. Testar mensagens com OK, confirmacoes com SIM e NAO e perguntas de texto
   com Enter e Escape. Repetir as chamadas para conferir a reutilizacao.

Os testes isolados nao substituem a validacao das selecoes com a base de
homologacao. O ganho de tempo deste lote ainda nao foi medido com o banco real.

## Desempenho: lote 1 (16/09/2026)

Autocreate reduzido de 29 para 26 objetos. Foram retirados somente:

- `Frm_Funcoes` (`funcoes.pas`): as funcoes globais da unit continuam disponiveis.
- `Frm_AchaImovel` (`AchaImovel.pas`, singular).
- `Frm_Achavendas` (`Achavendas.pas`, plural).

`Frm_Achaimoveis` (plural) e `Frm_AchaVenda` (singular), usados nas pesquisas,
permanecem no autocreate. Nenhuma unit ou DFM foi excluido.

A apresentacao agora e exibida apos a criacao do principal e antes do DataModule,
enquanto os componentes sao preparados. Ela e liberada em `finally`, sem esperar
os 470 disparos do timer. A primeira pintura ainda depende da inicializacao das
units e da criacao do principal. A barra visual nao mede progresso real.

As cotacoes sao consultadas em uma thread depois da entrada e da agenda, com
timeouts de conexao/leitura de 3 segundos e ate 3 redirecionamentos. Esses limites
nao representam um prazo total garantido, inclusive para DNS. A thread nao acessa
forms nem datasets; o callback atualiza a interface na thread principal e e
desvinculado ao destruir o principal. Falhas deixam o painel de cotacoes oculto.
O timer de memoria continua mostrando o consumo, sem forcar a reducao do working
set. A entrada negada encerra o evento antes das etapas seguintes.

O DataModule, o monitor SQL, as consultas financeiras e o UPDATE de saldos
permanecem neste lote. A instrumentacao nao muda campos ou consultas.

### Diagnostico opcional

Executar `SIAI.exe /diagnostico` grava `%TEMP%\SIAI-startup-<PID>.log`.
Sem esse parametro nao e gravado log de desempenho. O arquivo registra etapas
da entrada e tempos de preparacao/metadados dos datasets, sem SQL, credenciais
ou valores dos registros. Os tempos de autenticacao e agenda podem incluir a
espera do usuario. O modo de diagnostico possui custo adicional de gravacao;
compare execucoes com o mesmo modo e a mesma base.

### Testes deste lote

Na raiz do projeto, executar `tests\RunStartupTests.bat` com Delphi XE10.1
instalado nos caminhos indicados no script. Saidas em `build\startup-audit\tests`;
logs em `build\startup-audit`. Os testes nao criam o DataModule nem executam o
FormShow principal, portanto nao conectam ao banco e nao fazem alteracoes financeiras.

Os nove cenarios verificam cotacoes validas/textuais/numericas, JSON invalido,
nulo ou incompleto, resposta vazia, falha de requisicao simulada, cancelamento
com destinatario encerrado e construcao do principal/apresentacao sem os tres
forms retirados. O callback tambem e verificado na thread principal.

Para homologacao manual, usar o executavel compilado com as DLLs, INI e demais
recursos da instalacao de teste. Nao copiar cegamente a configuracao de producao:
o sistema ainda executa o UPDATE financeiro existente durante a entrada.

1. Abrir normalmente e verificar apresentacao, login e tela principal.
2. Negar/cancelar o acesso e confirmar que a entrada nao continua.
3. Testar com internet indisponivel e fechar enquanto a cotacao esta pendente.
4. Pesquisar imoveis, selecionar/cancelar e reabrir a pesquisa.
5. Pesquisar vendas; conferir tambem os acessos por cessao e troca de lote.
6. Validar mensagens, confirmacoes e funcoes comuns em cadastros/recebimentos.
7. Alternar Vendas, Participantes e Recebimentos e observar a resposta apos 10 segundos.

Compilacao e testes isolados nao substituem a homologacao com banco nem medem
o ganho real no ambiente do cliente.

### Regra para proximos lotes

Sempre informar ao usuario os forms retirados do autocreate, com nome da unit,
os testes executados, seus resultados, as limitacoes e um roteiro de teste manual.
Continuar em lotes pequenos; mudancas financeiras dependem da proxima autorizacao.
