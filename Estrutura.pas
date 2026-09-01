unit Estrutura;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, JvExControls, JvSpecialProgress, dxButton, StdCtrls,
  XBanner, db, IniFiles, dxCore2;

type
  TFrmEstrutura = class(TForm)
    DXBPesquisar: TdxButton;
    dxButton1: TdxButton;
    JSP1: TJvSpecialProgress;
    ListBox1: TListBox;
    XBanner1: TXBanner;
    Label2: TLabel;
    procedure dxButton1Click(Sender: TObject);
    procedure DXBPesquisarClick(Sender: TObject);
    function Altera(Tipo,vartabela,varnovo,varcomp : string) : boolean;
    procedure FormShow(Sender: TObject);
    function verifica : Boolean;
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmEstrutura: TFrmEstrutura;
  texto: TStringList;
  varschemata : string;
implementation

uses Principal, Tabelas, Funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrmEstrutura.dxButton1Click(Sender: TObject);
begin
  Close;
end;


procedure TFrmEstrutura.DXBPesquisarClick(Sender: TObject);
begin
  if not Verif_senha('Estrutura','Atualizar a Estrutura do Sistema','') then exit;
  verifica;

end;

Function TFrmEstrutura.Altera(tipo,vartabela,varnovo,varcomp : string): Boolean;
VAr
  VarResult : Boolean;
Begin
  if tipo='CREATE' Then
    ListBox1.Items.Add('Criando Tabela '+Vartabela)
  else
    ListBox1.Items.Add('Tabela '+Vartabela+' -> '+varnovo);
  JSP1.Position := JSP1.Position +1;
  DM_tabelas.ZQEmpresa.SQL.Clear;
  if tipo='CREATE' Then
    DM_Tabelas.ZQEmpresa.SQL.Add('CREATE TABLE `'+varschemata+'`.`'+Vartabela+'` ('+Varnovo+' '+varcomp)
  else if tipo='ADD CONSTRAINT' Then
    DM_Tabelas.ZQEmpresa.SQL.Add('ALTER TABLE `'+varschemata+'`.`'+Vartabela+'` '+tipo+'ADD CONSTRAINT `)'+Varnovo)
  else if tipo='DROP INDEX' Then
    DM_Tabelas.ZQEmpresa.SQL.Add('ALTER TABLE `'+varschemata+'`.`'+Vartabela+'` DROP INDEX `'+Varnovo+'` '+varcomp)
  else
    DM_Tabelas.ZQEmpresa.SQL.Add('ALTER TABLE `'+varschemata+'`.`'+Vartabela+'` '+tipo+' COLUMN `'+Varnovo+'` '+varcomp);
  VarResult := False;
  try
  dm_tABELAS.ZQEmpresa.ExecSQL;
  except
    On EConvertError do VarResult := False;
  else
    VarResult := True;
  end;
  If Not VarResult Then Begin
    if tipo='ADD' Then
      ListBox1.Items.Append('                         Incluido')
    else if tipo='CREATE' Then
      ListBox1.Items.Append('                         Criado tabela')
    else if tipo='DROP' Then
      ListBox1.Items.Append('                         Excluido')
    else
      ListBox1.Items.Append('                         Alterado');
  end;
  ListBox1.Refresh;
  Result := VarResult;
end;

procedure TFrmEstrutura.FormShow(Sender: TObject);
Var
  ArqIni : tIniFile;
  VarPath : string;
begin
  VarPath := ExtractFilePath( Application.ExeName );
  ArqIni := tIniFile.Create(varpath+'siai.Ini');
  try
    //  ArqIni := tIniFile.Create('c:\alpha\gead.Ini');
    varschemata := ArqIni.ReadString('BANCO_DE_DADOS','SCHEMATA', varschemata );
    if senhaestrut then Begin
    end;
    if not senhaestrut then
      verifica;
  finally
    ArqIni.Free;
  end;

end;

function TFrmEstrutura.verifica : Boolean;
Begin
  JSP1.Maximum := 181;
  Altera('ADD','loteamento','pasta_mapa','varchar(100)');                              // 12/11/09
  Altera('DROP','participante','tpfon1','');                                           // 15/11/09
  Altera('DROP','participante','fpfon2','');                                           // 15/11/09
  Altera('DROP','participante','fpfon3','');                                           // 15/11/09
  Altera('ADD','loteamento','nomeloteamento','varchar(100)');                          // 15/11/09
  Altera('ADD','participante','naturalidade','integer');                               // 16/11/09
  Altera('ADD','participante','localdetrab','varchar(100)');                           // 16/11/09
  Altera('ADD','participante','profissao','varchar(30)');                              // 16/11/09
  Altera('ADD','incorporador_loteamento','percent_parte','double');                    // 16/11/09
  Altera('ADD','cidade','cepgeral','varchar(10)');                                     // 16/11/09
  Altera('ADD','participante','estadocivil','varchar(20)');                            // 19/11/09
  Altera('ADD','participante','renda','double');                                       // 19/11/09
  Altera('ADD','loteamento','tipo','varchar(15)');                                     // 19/11/09
  Altera('ADD','representa_participante','data','date');                               // 19/11/09
  Altera('ADD','representa_participante','cartorio','varchar(15)');                    // 19/11/09
  Altera('ADD','representa_participante','livro','varchar(4)');                        // 19/11/09
  Altera('ADD','representa_participante','folha','integer');                           // 19/11/09
  Altera('ADD','representa_participante','cidade_cart','integer');                     // 19/11/09
  Altera('ADD','representa_participante','comarca','integer');                         // 19/11/09
  Altera('ADD','representa_participante','regime','varchar(30)');                      // 20/11/09
  Altera('ADD','representa_participante','trabalho','varchar(30)');                    // 20/11/09
  Altera('ADD','representa_participante','admissao','date');                           // 20/11/09
  Altera('ADD','representa_participante','renda','double');                            // 20/11/09
  Altera('ADD','participante','observacao','text');                                    // 20/11/09
  Altera('ADD','imovel','disponivel','varchar(3)');                                    // 25/11/09
  Altera('ADD','corretor','codcidade','integer');                                      // 26/11/09
  Altera('CREATE','tipodocumento','idtipodocumento','integer UNSIGNED NOT NULL AUTO_INCREMENT, PRIMARY KEY(idtipodocumento))');  // 07/01/09
  Altera('ADD','tipodocumento','tipodoc','varchar(2)');                                // 27/11/09
  Altera('ADD','tipodocumento','descricao','varchar(40)');                             // 27/11/09
  Altera('ADD','tipodocumento','vend_receb','varchar(1)');                             // 27/11/09
  Altera('ADD','tipodocumento','vend_caixa','varchar(1)');                             // 27/11/09
  Altera('ADD','tipodocumento','receb_receb','varchar(1)');                            // 27/11/09
  Altera('ADD','tipodocumento','receb_caixa','varchar(1)');                            // 27/11/09
  Altera('ADD','tipodocumento','dados_chequ','varchar(1)');                            // 27/11/09
  Altera('ADD','tipodocumento','so_avista','varchar(1)');                              // 27/11/09
  Altera('ADD','tipodocumento','lancabanco','varchar(1)');                             // 27/11/09
  Altera('ADD','tipodocumento','somapaga','varchar(1)');                               // 27/11/09
  Altera('ADD','tipodocumento','lancinclus','varchar(1)');                             // 27/11/09
  Altera('MODIFY','imovel','medidafrente','VARCHAR(20) DEFAULT NULL');                 // 30/11/09
  Altera('MODIFY','imovel','medidafundo','VARCHAR(20) DEFAULT NULL');                  // 30/11/09
  Altera('MODIFY','imovel','medidaesquerda','VARCHAR(20) DEFAULT NULL');               // 30/11/09
  Altera('MODIFY','imovel','medidadedireita','VARCHAR(20) DEFAULT NULL');              // 30/11/09
  Altera('MODIFY','imovel','esquinalinha','VARCHAR(20) DEFAULT NULL');                 // 30/11/09
  Altera('MODIFY','imovel','esquinacurva','VARCHAR(20) DEFAULT NULL');                 // 30/11/09
  Altera('CREATE','plano_contas','codigo','integer UNSIGNED NOT NULL AUTO_INCREMENT, PRIMARY KEY(codigo))');  // 30/11/09
  Altera('ADD','plano_contas','doccomum','varchar(2)');                                // 30/11/09
  Altera('ADD','plano_contas','mascara','varchar(16)');                                // 30/11/09
  Altera('ADD','plano_contas','descricao','varchar(50)');                              // 30/11/09
  Altera('ADD','plano_contas','classificacao','varchar(1)');                           // 30/11/09
  Altera('ADD','plano_contas','valor','double');                                       // 30/11/09
  Altera('ADD','plano_contas','cod_reduzido','integer');                               // 30/11/09
  Altera('ADD','plano_contas','cred_debi','varchar(1)');                               // 30/11/09
  Altera('ADD','plano_contas','usoativo','varchar(3)');                                // 30/11/09
  Altera('ADD','plano_contas','CDPlano','varchar(1)');                                 // 30/11/09
  Altera('ADD','plano_contas','DespFixVar','varchar(10)');                             // 30/11/09
  Altera('CREATE','cheque_rec','idcheque','integer UNSIGNED NOT NULL AUTO_INCREMENT, PRIMARY KEY(idcheque))'); // 07/05/09
  Altera('ADD','cheque_rec','numero','varchar(20)');                                   // 01/12/09
  Altera('DROP','cheque_rec','ch_conta','');                                           // 15/11/09
  Altera('ADD','cheque_rec','banco','varchar(3)');                                     // 01/12/09
  Altera('ADD','cheque_rec','agencia','varchar(3)');                                   // 01/10/09
  Altera('ADD','cheque_rec','dono','varchar(50)');                                     // 01/12/09
  Altera('ADD','cheque_rec','CPF_CNPJ','varchar(20)');                                 // 01/12/09
  Altera('ADD','cheque_rec','valor','float');                                          // 01/12/09
  Altera('ADD','cheque_rec','emissao','date');                                         // 01/12/09
  Altera('ADD','cheque_rec','deposito','date');                                        // 01/12/09
  Altera('ADD','cheque_rec','alias','varchar(12)');                                    // 01/12/09
  Altera('ADD','cheque_rec','prorrogado','date');                                      // 01/12/09
  Altera('ADD','cheque_rec','conta','varchar(15)');                                    // 01/12/09
  Altera('ADD','cheque_rec','idrecebimento','integer');                                // 01/12/09
  Altera('ADD','cheque_rec','idpagamento','integer');                                  // 01/12/09
  Altera('ADD','cheque_rec','idvinculo','integer');                                    // 01/12/09
  Altera('CREATE','limite_especial','idespecial','integer UNSIGNED NOT NULL AUTO_INCREMENT, PRIMARY KEY(idespecial))'); // 01/12/09
  Altera('ADD','limite_especial','data_valida','date');                                // 01/12/09
  Altera('ADD','limite_especial','Conta_Bancaria_cod_banco','integer');                // 01/12/09
  Altera('ADD','limite_especial','valor','double');                                    // 01/12/09
  Altera('ADD','conta_bancaria','nometitular','varchar(50)');                          // 01/12/09
  Altera('ADD','conta_bancaria','aberta','date');                                      // 01/12/09
  Altera('ADD','conta_bancaria','nometitular','varchar(50)');                          // 01/12/09
  Altera('ADD','conta_bancaria','senha_interna','varchar(50)');                        // 01/12/09
  Altera('ADD','conta_bancaria','senha_conta','varchar(50)');                          // 01/12/09
  Altera('ADD','conta_bancaria','ativa','varchar(1)');                                 // 01/12/09
  Altera('ADD','conta_bancaria','jurosemboleto','double');                             // 01/12/09
  Altera('ADD','conta_bancaria','taxadiaria','double');                                // 01/12/09
  Altera('ADD','conta_bancaria','valoroupercent','varchar(1)');                        // 01/12/09
  Altera('ADD','conta_bancaria','carteira','varchar(3)');                              // 01/12/09
  Altera('ADD','conta_bancaria','boleto','varchar(15)');                               // 01/12/09
  Altera('CREATE','boleto_Mens','idboleto_Mens','integer UNSIGNED NOT NULL AUTO_INCREMENT, PRIMARY KEY(idboleto_Mens))');        // 19/10/09
  Altera('ADD','boleto_Mens','idordem','integer');                                     // 01/12/09
  Altera('ADD','boleto_Mens','linha1','text');                                         // 01/12/09
  Altera('ADD','boleto_Mens','linha2','text');                                         // 01/12/09
  Altera('ADD','boleto_Mens','linha3','text');                                         // 01/12/09
  Altera('ADD','boleto_Mens','linha4','text');                                         // 01/12/09
  Altera('ADD','Conta_bancaria','boleto','varchar(15)');                               // 01/12/09
  Altera('CREATE','quadras','idquadras','integer UNSIGNED NOT NULL AUTO_INCREMENT, PRIMARY KEY(idquadras))');        // 01/12/09
  Altera('ADD','quadras','loteamento_idLoteamento','integer');                         // 01/12/09
  Altera('ADD','quadras','descricao','varchar(25)');                                   // 01/12/09
  Altera('ADD','quadras','apedlido','varchar(10)');                                    // 01/12/09
  Altera('ADD','venda','valorvenda','double');                                         // 03/12/09
  Altera('CREATE','recebimento','idrecebimento','integer UNSIGNED NOT NULL AUTO_INCREMENT, PRIMARY KEY(idrecebimento))');        // 04/12/09
  Altera('ADD','recebimento','documento','varchar(20)');                               // 04/12/09
  Altera('ADD','recebimento','cliente','integer');                                     // 04/12/09
  Altera('ADD','recebimento','usuario','integer');                                     // 04/12/09
  Altera('ADD','recebimento','Dt_Entrada','date');                                     // 04/12/09
  Altera('ADD','recebimento','Dt_Vencimento','date');                                  // 04/12/09
  Altera('ADD','recebimento','Valor','double');                                        // 04/12/09
  Altera('ADD','recebimento','Observ','Text');                                         // 04/12/09
  Altera('ADD','recebimento','VrDoc','Double');                                        // 04/12/09
  Altera('ADD','recebimento','ordem','varchar(14)');                                   // 04/12/09
  Altera('ADD','recebimento','TipDoc','varchar(2)');                                   // 04/12/09
  Altera('ADD','recebimento','saldo','double');                                        // 04/12/09
  Altera('ADD','recebimento','marcar','varchar(1)');                                   // 04/12/09
  Altera('ADD','recebimento','RefBaixa','integer');                                    // 04/12/09
  Altera('ADD','recebimento','refvinda','integer');                                    // 04/12/09
  Altera('ADD','recebimento','contabil','integer');                                    // 04/12/09
  Altera('ADD','recebimento','empresa','integer');                                     // 04/12/09
  Altera('ADD','recebimento','custodaparcela','double');                               // 04/12/09
  Altera('ADD','recebimento','origem','varchar(10)');                                  // 04/12/09
  Altera('ADD','recebimento','adversa','integer');                                     // 04/12/09
  Altera('ADD','recebimento','recpag','varchar(1)');                                   // 04/12/09
  Altera('ADD','recebimento','numordem','integer');                                    // 04/12/09
  Altera('CREATE','receb_baixa','idreceb_baixa','integer UNSIGNED NOT NULL AUTO_INCREMENT, PRIMARY KEY(idreceb_baixa))');        // 04/12/09
  Altera('ADD','receb_baixa','Dt_rec','date');                                         // 04/12/09
  Altera('ADD','receb_baixa','Vr_rec','double');                                       // 04/12/09
  Altera('ADD','receb_baixa','TipDoc','varchar(2)');                                   // 04/12/09
  Altera('ADD','receb_baixa','Docum','varchar(20)');                                   // 04/12/09
  Altera('ADD','receb_baixa','Juros_Vr','Double');                                     // 04/12/09
  Altera('ADD','receb_baixa','Desc_Vr','Double');                                      // 04/12/09
  Altera('ADD','receb_baixa','RefBaixa','integer');                                    // 04/12/09
  Altera('ADD','receb_baixa','obsebx','Text');                                         // 04/12/09
  Altera('ADD','receb_baixa','Codcontabil','integer');                                 // 04/12/09
  Altera('ADD','receb_baixa','dataref','date');                                        // 04/12/09
  Altera('ADD','receb_baixa','vencimento','date');                                     // 04/12/09
  Altera('CREATE','recbxhist','idrecbxhist','integer UNSIGNED NOT NULL AUTO_INCREMENT, PRIMARY KEY(idrecbxhist))');        // 04/12/09
  Altera('ADD','recbxhist','refer','integer');                                         // 04/12/09
  Altera('ADD','recbxhist','idrecib','integer');                                       // 04/12/09
  Altera('ADD','recbxhist','valor','double');                                          // 04/12/09
  Altera('ADD','recbxhist','descon','double');                                         // 04/12/09
  Altera('ADD','recbxhist','juros','double');                                          // 04/12/09
  Altera('ADD','recbxhist','percent_usado','double');                                  // 04/12/09
  Altera('ADD','recbxhist','data','date');                                             // 04/12/09
  Altera('CREATE','numordem','idnumordem','integer UNSIGNED NOT NULL AUTO_INCREMENT, PRIMARY KEY(idnumordem))');    // 04/12/09
  Altera('CREATE','contabxrec','cotagem','integer UNSIGNED NOT NULL AUTO_INCREMENT, PRIMARY KEY(cotagem))');        // 04/12/09
  Altera('ADD','contabxrec','hoje','date');                                            // 04/12/09
  Altera('ADD','contabxrec','entrada','date');                                         // 04/12/09
  Altera('CREATE','configuracoes','idconfiguracoes','integer UNSIGNED NOT NULL AUTO_INCREMENT, PRIMARY KEY(idconfiguracoes))');        // 04/12/09
  Altera('ADD','configuracoes','venda_planodecontas','integer');                       // 04/12/09
  Altera('ADD','participante','codrecinc','integer');                                  // 08/12/09
  Altera('ADD','participante','codrecBx','integer');                                   // 08/12/09
  Altera('ADD','participante','codpaginc','integer');                                  // 08/12/09
  Altera('ADD','participante','codpagBx','integer');                                   // 08/12/09
  Altera('CREATE','caixa','codigo','integer UNSIGNED NOT NULL AUTO_INCREMENT, PRIMARY KEY(codigo))');  // 09/02/09
  Altera('ADD','caixa','plano_contas_codigo','INTEGER');                                         // 09/02/09
  Altera('ADD','caixa','data_lan','date');                                                       // 09/02/09
  Altera('ADD','caixa','vr_lan','float');                                                        // 09/02/09
  Altera('ADD','caixa','obs','text');                                                            // 09/02/09
  Altera('ADD','caixa','documento','VARCHAR(15)');                                                   // 09/02/09
  Altera('ADD','caixa','empresa','INTEGER');                                                     // 09/02/09
  Altera('ADD','caixa','vinculo','VARCHAR(10)');                                                 // 09/02/09
  Altera('ADD','caixa','saldo','DOUBLE');                                                        // 09/02/09
  Altera('ADD','caixa','credeb','VARCHAR(1)');                                                   // 09/02/09
  Altera('ADD','caixa','contraquem','VARCHAR(1)');                                               // 09/02/09
  Altera('ADD','caixa','contracodigo','integer');                                                // 09/02/09
  Altera('CREATE','dividido','iddividido','integer UNSIGNED NOT NULL AUTO_INCREMENT, PRIMARY KEY(iddividido))');     // 19/02/09
  Altera('ADD','dividido','participante','integer');                                             // 09/12/09
  Altera('ADD','dividido','percentual','float');                                                 // 09/12/09
  Altera('ADD','dividido','planodeconta','integer');                                             // 09/12/09
  Altera('ADD','dividido','tipodocumento','varchar(2)');                                         // 09/12/09
  Altera('ADD','dividido','ordem','integer');                                                    // 09/12/09
  Altera('ADD','configuracoes','jurosmensal','double');                                          // 09/12/09
  Altera('ADD','configuracoes','paticipantedefault','integer');                                  // 09/12/09
  Altera('ADD','configuracoes','diasdeatrazo','integer');                                        // 09/12/09
  Altera('CREATE','mov_banco','lancamento','integer UNSIGNED NOT NULL AUTO_INCREMENT, PRIMARY KEY(lancamento))');  // 09/12/09
  Altera('ADD','mov_banco','conta_bancaria_cod_banco','integer');                                // 09/12/09
  Altera('ADD','mov_banco','plano_contas_codigo','integer');                                     // 09/12/09
  Altera('ADD','mov_banco','documento','varchar(15)');                                           // 09/12/09
  Altera('ADD','mov_banco','dt_lanc','date');                                                    // 09/12/09
  Altera('ADD','mov_banco','hora_lanc','time');                                                  // 09/12/09
  Altera('ADD','mov_banco','dt_conciliado','date');                                              // 09/12/09
  Altera('ADD','mov_banco','vr_lanc','float');                                                   // 09/12/09
  Altera('ADD','mov_banco','credeb','varchar(1)');                                               // 09/12/09
  Altera('ADD','mov_banco','saldo_lanc','double');                                               // 09/12/09
  Altera('ADD','mov_banco','saldo_conci','double');                                              // 09/12/09
  Altera('ADD','mov_banco','obs','text');                                                        // 09/12/09
  Altera('ADD','mov_banco','contraquem','varchar(1)');                                           // 09/12/09
  Altera('ADD','mov_banco','contracodigo','integer');                                            // 09/12/09
  Altera('ADD','mov_banco','vinculo','varchar(10)');                                             // 09/12/09
  Altera('ADD','loteamento','codcontabancaria','integer');                                       // 11/12/09
  Altera('ADD','conta_bancaria','idparticipante','integer');                                     // 11/12/09
  Altera('ADD','incorporador_loteamento','codcontabancaria','integer');                          // 15/12/09
  Altera('CREATE','banco_remessa','idbanco_remessa','integer UNSIGNED NOT NULL AUTO_INCREMENT, PRIMARY KEY(idbanco_remessa))');     // 16/12/09
  Altera('ADD','banco_remessa','remessa','integer');                                             // 16/12/09
  Altera('ADD','banco_remessa','idbanco','integer');                                             // 16/12/09
  Altera('ADD','banco_remessa','gerado','date');                                                 // 16/12/09
  Altera('CREATE','remessa_receb','idremessa_receb','integer UNSIGNED NOT NULL AUTO_INCREMENT, PRIMARY KEY(idremessa_receb))');     // 16/12/09
  Altera('ADD','remessa_receb','remessa','integer');                                             // 16/12/09
  Altera('ADD','remessa_receb','idrec','integer');                                               // 16/12/09
  Altera('ADD','banco_remessa','nomearq','varchar(20)');                                         // 24/12/09
  Altera('ADD','participante','fone4','varchar(14)');                                            // 05/01/10
  Altera('DROP','participante','apelido','');                                                    // 05/01/10
  Altera('ADD','participante','complemento','varchar(30)');                                      // 19/01/10
  Altera('ADD','participante','complemento_cob','varchar(30)');                                  // 19/01/10
  Altera('ADD','participante','cadastrado','date');                                              // 26/01/10
  Altera('ADD','loteamento','apelido','varchar(50)');                                            // 28/01/10
  Altera('ADD','configuracoes','recebe_planodecontas','integer');                                // 23/02/10
  Altera('CREATE','conjuge','idconjuge','integer UNSIGNED NOT NULL AUTO_INCREMENT, PRIMARY KEY(idconjuge))');     // 25/02/10
  Altera('ADD','conjuge','nome','varchar(100)');                                                 // 23/02/10
  Altera('ADD','conjuge','doc1','varchar(20)');                                                  // 23/02/10
  Altera('ADD','conjuge','doc2','varchar(20)');                                                  // 23/02/10
  Altera('ADD','conjuge','nasc','date');                                                         // 23/02/10
  Altera('ADD','conjuge','natural','integer');                                                   // 23/02/10
  Altera('ADD','conjuge','idparti','integer');                                                   // 23/02/10
  Altera('ADD','conjuge','email','varchar(100)');                                                // 23/02/10
  Altera('ADD','conjuge','assina','varchar(1)');                                                 // 23/02/10
  Altera('ADD','conjuge','empresa','varchar(100)');                                              // 23/02/10
  Altera('ADD','conjuge','cargo','varchar(20)');                                                 // 23/02/10
  Altera('ADD','conjuge','admissa','date');                                                      // 23/02/10
  Altera('ADD','conjuge','renda','double');                                                      // 23/02/10
  Altera('ADD','conjuge','regime','varchar(30)');                                                // 23/02/10
  Altera('ADD','conjuge','comarca','double');                                                    // 23/02/10
  Altera('ADD','conjuge','cidade_cart','double');                                                // 23/02/10
  Altera('ADD','conjuge','folha','varchar(10)');                                                 // 23/02/10
  Altera('ADD','conjuge','livro','varchar(7)');                                                  // 23/02/10
  Altera('ADD','conjuge','cartorio','varchar(15)');                                              // 23/02/10
  Altera('ADD','conjuge','data','date');                                                         // 23/02/10
  Altera('ADD','recebimento','idloteamento','integer');                                          // 23/02/10
  Altera('ADD','recebimento','venda_idvenda','integer');                                         // 03/03/10
  Altera('ADD','recebimento','quadralote','varchar(20)');                                        // 03/03/10
  Altera('ADD','recebimento','numboleto','varchar(11)');                                         // 03/03/10
  Altera('ADD','representa_participante','nomeconjuge','varchar(100)');                          // 04/03/10
  Altera('ADD','representa_participante','doc1conjuge','varchar(20)');                           // 04/03/10
  Altera('ADD','representa_participante','doc2conjuge','varchar(20)');                           // 04/03/10
  Altera('ADD','representa_participante','nascconjuge','date');                                  // 04/03/10
  Altera('ADD','representa_participante','assconjuge','varchar(1)');                             // 04/03/10
  Altera('ADD','representa_participante','cidconjuge','integer');                                // 04/03/10
  Altera('CREATE','mens_retorno','idmens_retorno','integer UNSIGNED NOT NULL AUTO_INCREMENT, PRIMARY KEY(idmens_retorno))');   // 08/03/10
  Altera('ADD','mens_retorno','banco','integer');                                   // 24/03/10
  Altera('ADD','mens_retorno','base','integer');                                    // 24/03/10
  Altera('ADD','mens_retorno','codigo','varchar(2)');                               // 24/03/10
  Altera('ADD','mens_retorno','mensagem','varchar(100)');                           // 24/03/10
  Altera('ADD','mens_retorno','atitude','varchar(1)');                              // 24/03/10
  Altera('CREATE','mens_motivo','idmens_motivo','integer UNSIGNED NOT NULL AUTO_INCREMENT, PRIMARY KEY(idmens_motivo))');   // 12/03/10
  Altera('ADD','mens_motivo','retorno','varchar(2)');                               // 24/03/10
  Altera('ADD','mens_motivo','codmotivo','varchar(3)');                             // 24/03/10
  Altera('ADD','mens_motivo','mensagem','varchar(100)');                            // 24/03/10
  Altera('ADD','banco_remessa','arquivo','varchar(15)');                            // 24/03/10
  Altera('ADD','remessa_receb','retorno_motivo','varchar(5)');                      // 24/03/10
  Altera('ADD','remessa_receb','acao','varchar(1)');                                // 24/03/10
  Altera('ADD','configuracoes','bxautodoc','varchar(2)');                           // 24/03/10
  Altera('ADD','configuracoes','bxautoplc','integer');                              // 24/03/10
  Altera('MODIFY','conjuge','cartorio','varchar(50) CHARACTER SET latin1 COLLATE latin1_swedish_ci DEFAULT NULL'); // 22/10/08
  Altera('MODIFY','cheque_rec','agencia','varchar(6) CHARACTER SET latin1 COLLATE latin1_swedish_ci DEFAULT NULL'); // 22/10/08
  Altera('ADD','contabxrec','tipobaixa','varchar(15)');                             // 28/04/10

  showmessage('Ao sair desta tela, o sistema ser� fechado para que as atualiza��es sejam concideradas,'+ chr(13) +'execute novamente o sistema para poder trabalhar normalmente...');
End;


procedure TFrmEstrutura.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;


end.


