unit importar_part;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset,
  Mask, ComCtrls, Grids, DBGrids, ZConnection;

type
  Tfrm_baixa = class(TForm)
    Button1: TButton;
    Button2: TButton;
    DS_Imovel: TDataSource;
    ZQImovel: TZQuery;

























    DS_Participante: TDataSource;
    ZqParticipante: TZQuery;
    DS_Responsavel: TDataSource;
    ZQresponsavel: TZQuery;
    DS_Comprador: TDataSource;
    ZQComprador: TZQuery;






    ZQVendedor: TZQuery;



    DS_Vendedor: TDataSource;
    ZQDividido: TZQuery;
    DS_Dividido: TDataSource;
    ZQConjuge: TZQuery;




















    DS_Conjuge: TDataSource;
    DS_Compr_conjuge: TDataSource;
    ZQCompr_conjuge: TZQuery;
    Dataprocuradores: TDataSource;
    ZQProcuradores: TZQuery;





    xlote: TMaskEdit;
    Label1: TLabel;
    ZQRemes_Receb: TZQuery;
    DS_Remes_Receb: TDataSource;
    DataZQhistorico_cessao: TDataSource;
    ZQrepresentantes: TZQuery;
    Datarepresentantes: TDataSource;




    nome: TMaskEdit;




















    ProgressBar1: TProgressBar;












































































    msg: TLabel;
    msg1: TLabel;
    DBGrid6: TDBGrid;
    DBGrid9: TDBGrid;
    DBGrid10: TDBGrid;
    DBGrid17: TDBGrid;
    DBGrid19: TDBGrid;
    DBGrid21: TDBGrid;
    DBGrid22: TDBGrid;
    zconeccao: TZConnection;
    ZParticipante: TZQuery;










































    ZResponsavel: TZQuery;


























    ZConjuge: TZQuery;




















    ZVendedor: TZQuery;



    ZCompr_conjuge: TZQuery;




















    ZProcuradores: TZQuery;





    Zrepresentantes: TZQuery;













    ProgressBar2: TProgressBar;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure xloteExit(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  frm_baixa: Tfrm_baixa;
  nlt:integer;
implementation

uses tabelas, Achaloteamento, uRuntimeFields;

{$R *.dfm}

procedure Tfrm_baixa.Button1Click(Sender: TObject);
Var
  Varrefbaixa, varregistro : integer;
  varsqltexto : string;
  Varpago, percentbaixa, baixatot, varjuros, varx,
  avanco,varsomarepasse, varsomarepasseautomatico : Double;
  contador,varreg : Integer;
  varordem, varnumordem1, varnumordem2 : string;

begin

    ZqParticipante.First;
    if ZqParticipante.RecordCount>0 then
    begin
      msg.Caption:='Participantes                          ';
      application.ProcessMessages;
      ProgressBar2.Position:=0;
      ProgressBar2.Max:=ZqParticipante.RecordCount;
      while not ZqParticipante.Eof do
      begin
        ProgressBar2.Position:=ZqParticipante.RecNo;

//         if not ZParticipante.Locate('nome_parte;doc1',VarArrayOf([ZqParticipantenome_parte.Value,charrem(ZqParticipantedoc1.Value)]),[]) then
         if not ZParticipante.Locate('nome_parte',ZQParticipante.FieldByName('nome_parte').Value,[]) then
         begin
              ZParticipante.insert;
              //ZParticipanteidpaticipante.Value:=ZqParticipanteidpaticipante.Value;
              ZParticipante.FieldByName('nome_parte').Value:=ZQParticipante.FieldByName('nome_parte').Value;
              ZParticipante.FieldByName('doc1').Value:=ZqParticipante.FieldByName('doc1').Value;
              ZParticipante.FieldByName('doc2').Value:=ZqParticipante.FieldByName('doc2').Value;
              ZParticipante.FieldByName('endereco').Value:=ZqParticipante.FieldByName('endereco').Value;
              ZParticipante.FieldByName('bairro').Value:=ZqParticipante.FieldByName('bairro').Value;
              ZParticipante.FieldByName('cidade').Value:=ZqParticipante.FieldByName('cidade').Value;
              ZParticipante.FieldByName('cep').Value:=ZqParticipante.FieldByName('cep').Value;
              ZParticipante.FieldByName('ende_cob').Value:=ZqParticipante.FieldByName('ende_cob').Value;
              ZParticipante.FieldByName('bairro_cob').Value:=ZqParticipante.FieldByName('bairro_cob').Value;
              ZParticipante.FieldByName('cidade_cob').Value:=ZqParticipante.FieldByName('cidade_cob').Value;
              ZParticipante.FieldByName('cep_cob').Value:=ZqParticipante.FieldByName('cep_cob').Value;
              ZParticipante.FieldByName('Fone1').Value:=ZqParticipante.FieldByName('Fone1').Value;
              ZParticipante.FieldByName('fone2').Value:=ZqParticipante.FieldByName('fone2').Value;
              ZParticipante.FieldByName('fone3').Value:=ZqParticipante.FieldByName('fone3').Value;
              ZParticipante.FieldByName('tipopessoa').Value:=ZqParticipante.FieldByName('tipopessoa').Value;
              ZParticipante.FieldByName('aniversario').Value:=ZqParticipante.FieldByName('aniversario').Value;
              ZParticipante.FieldByName('nacionalidade').Value:=ZqParticipante.FieldByName('nacionalidade').Value;
              ZParticipante.FieldByName('email').Value:=ZqParticipante.FieldByName('email').Value;
              ZParticipante.FieldByName('naturalidade').Value:=ZqParticipante.FieldByName('naturalidade').Value;
              ZParticipante.FieldByName('localdetrab').Value:=ZqParticipante.FieldByName('localdetrab').Value;
              ZParticipante.FieldByName('profissao').Value:=ZqParticipante.FieldByName('profissao').Value;
              ZParticipante.FieldByName('estadocivil').Value:=ZqParticipante.FieldByName('ESTADOCIVIL').Value;
              ZParticipante.FieldByName('renda').Value:=ZqParticipante.FieldByName('renda').Value;
              ZParticipante.FieldByName('observacao').Value:=ZqParticipante.FieldByName('observacao').Value;
              ZParticipante.FieldByName('codrecinc').Value:=ZqParticipante.FieldByName('codrecinc').Value;
              ZParticipante.FieldByName('codrecBx').Value:=ZqParticipante.FieldByName('codrecBx').Value;
              ZParticipante.FieldByName('codpaginc').Value:=ZqParticipante.FieldByName('codpaginc').Value;
              ZParticipante.FieldByName('codpagBx').Value:=ZqParticipante.FieldByName('codpagBx').Value;
              ZParticipante.FieldByName('fone4').Value:=ZqParticipante.FieldByName('fone4').Value;
              ZParticipante.FieldByName('complemento').Value:=ZqParticipante.FieldByName('complemento').Value;
              ZParticipante.FieldByName('complemento_cob').Value:=ZqParticipante.FieldByName('complemento_cob').Value;
              ZParticipante.FieldByName('cadastrado').Value:=ZqParticipante.FieldByName('cadastrado').Value;
              ZParticipante.FieldByName('Nome_Firma').Value:=ZqParticipante.FieldByName('Nome_Firma').Value;
              ZParticipante.FieldByName('Inativo').value:=ZqParticipante.FieldByName('Inativo').value;
              ZParticipante.FieldByName('Sexo').value:=ZqParticipante.FieldByName('sexo').value;
              ZParticipante.FieldByName('cc').value:=ZqParticipante.FieldByName('cc').value;
              ZParticipante.FieldByName('cex').value:=ZqParticipante.FieldByName('cex').value;
              ZParticipante.FieldByName('jb').value:=ZqParticipante.FieldByName('jb').value;
              ZParticipante.FieldByName('judicial').value:=ZqParticipante.FieldByName('judicial').value;
              ZParticipante.FieldByName('xquadra').value:=ZqParticipante.FieldByName('xquadra').value;
              ZParticipante.FieldByName('xlote').value:=ZqParticipante.FieldByName('xlote').value;
              ZParticipante.post;
              if (ZqParticipante.FieldByName('ESTADOCIVIL').Value='CASADO(A)') or (ZqParticipante.FieldByName('ESTADOCIVIL').Value='AMASIADO(A)') then
              begin
                ZCompr_conjuge.Insert;
                ZCompr_conjuge.FieldByName('nome').Value:=ZQCompr_conjuge.FieldByName('nome').Value;
                ZCompr_conjuge.FieldByName('doc1').Value:=ZQCompr_conjuge.FieldByName('doc1').Value;
                ZCompr_conjuge.FieldByName('doc2').Value:=ZQCompr_conjuge.FieldByName('doc2').Value;
                ZCompr_conjuge.FieldByName('nasc').Value:=ZQCompr_conjuge.FieldByName('nasc').Value;
                ZCompr_conjuge.FieldByName('natural').Value:=ZQCompr_conjuge.FieldByName('natural').Value;
                ZCompr_conjuge.FieldByName('idparti').Value:=ZParticipante.FieldByName('idpaticipante').Value;
                ZCompr_conjuge.FieldByName('email').Value:=ZQCompr_conjuge.FieldByName('email').Value;
                ZCompr_conjuge.FieldByName('assina').Value:=ZQCompr_conjuge.FieldByName('assina').Value;
                ZCompr_conjuge.FieldByName('empresa').Value:=ZQCompr_conjuge.FieldByName('empresa').Value;
                ZCompr_conjuge.FieldByName('cargo').Value:=ZQCompr_conjuge.FieldByName('cargo').Value;
                ZCompr_conjuge.FieldByName('admissa').Value:=ZQCompr_conjuge.FieldByName('admissa').Value;
                ZCompr_conjuge.FieldByName('renda').Value:=ZQCompr_conjuge.FieldByName('renda').Value;
                ZCompr_conjuge.FieldByName('regime').Value:=ZQCompr_conjuge.FieldByName('regime').Value;
                ZCompr_conjuge.FieldByName('comarca').Value:=ZQCompr_conjuge.FieldByName('comarca').Value;
                ZCompr_conjuge.FieldByName('cidade_cart').Value:=ZQCompr_conjuge.FieldByName('cidade_cart').Value;
                ZCompr_conjuge.FieldByName('folha').Value:=ZQCompr_conjuge.FieldByName('folha').Value;
                ZCompr_conjuge.FieldByName('livro').Value:=ZQCompr_conjuge.FieldByName('livro').Value;
                ZCompr_conjuge.FieldByName('cartorio').Value:=ZQCompr_conjuge.FieldByName('cartorio').Value;
                ZCompr_conjuge.FieldByName('data').Value:=ZQCompr_conjuge.FieldByName('data').Value;
                ZCompr_conjuge.Post;
              end;

         end;
         ZqParticipante.next;
      end;
    end;
{    ZQresponsavel.First;



       if ZQresponsavel.RecordCount>0 then
       begin
         msg.Caption:='Responsaveis                       ';
         application.ProcessMessages;

         while not ZQresponsavel.Eof do
         begin
           if not Zresponsavel.Locate('paticipante_idpaticipante',VarArrayOf([ZQResponsavelpaticipante_idpaticipante.Value]),[]) then
           begin
                  Zresponsavel.Insert;
                  Zresponsavelpaticipante_idpaticipante.Value:=ZQresponsavelpaticipante_idpaticipante.Value;
                  Zresponsaveldocumen1.Value:=ZQresponsaveldocumen1.Value;
                  Zresponsaveldocumen2.Value:=ZQresponsaveldocumen2.Value;
                  Zresponsavelnasicmento.Value:=ZQresponsavelnasicmento.Value;
                  Zresponsavelcid_natural.Value:=ZQresponsavelcid_natural.Value;
                  Zresponsavelnomerepres.Value:=ZQresponsavelnomerepres.Value;
                  Zresponsavelemailrepres.Value:=ZQresponsavelemailrepres.Value;
                  Zresponsavelprofiss.Value:=ZQresponsavelprofiss.Value;
                  Zresponsavelassina.Value:=ZQresponsavelassina.Value;
                  Zresponsaveldata.Value:=ZQresponsaveldata.Value;
                  Zresponsavelcartorio.Value:=ZQresponsavelcartorio.Value;
                  Zresponsavellivro.Value:=ZQresponsavellivro.Value;
                  Zresponsavelfolha.Value:=ZQresponsavelfolha.Value;
                  Zresponsavelcidade_cart.Value:=ZQresponsavelcidade_cart.Value;
                  Zresponsavelcomarca.Value:=ZQresponsavelcomarca.Value;
                  Zresponsavelregime.Value:=ZQresponsavelregime.Value;
                  Zresponsaveltrabalho.Value:=ZQresponsaveltrabalho.Value;
                  Zresponsaveladmissao.Value:=ZQresponsaveladmissao.Value;
                  Zresponsavelrenda.Value:=ZQresponsavelrenda.Value;
                  Zresponsavelnomeconjuge.Value:=ZQresponsavelnomeconjuge.Value;
                  Zresponsaveldoc1conjuge.Value:=ZQresponsaveldoc1conjuge.Value;
                  Zresponsaveldoc2conjuge.Value:=ZQresponsaveldoc2conjuge.Value;
                  Zresponsavelnascconjuge.Value:=ZQresponsavelnascconjuge.Value;
                  Zresponsavelassconjuge.Value:=ZQresponsavelassconjuge.Value;
                  Zresponsavelcidconjuge.Value:=ZQresponsavelcidconjuge.Value;
                  Zresponsavel.post;
           end;
           ZQresponsavel.next;

         end;
       end;
       ZQCompr_conjuge.first;
       if ZQCompr_conjuge.RecordCount>0 then
       begin
         msg.Caption:='Conjug�                                ';
         application.ProcessMessages;
         ProgressBar2.Position:=0;
         ProgressBar2.Max:=ZQCompr_conjuge.RecordCount;

         while not ZQCompr_conjuge.Eof do
         begin
           ProgressBar2.Position:=ZQCompr_conjuge.RecNo;
           if not ZCompr_conjuge.Locate('nome;idparti',VarArrayOf([ZQCompr_conjugenome.Value,ZQCompr_conjugeidparti.Value]),[]) then
           begin
             ZCompr_conjuge.Insert;
             ZCompr_conjugenome.Value:=ZQCompr_conjugenome.Value;
             ZCompr_conjugedoc1.Value:=ZQCompr_conjugedoc1.Value;
             ZCompr_conjugedoc2.Value:=ZQCompr_conjugedoc2.Value;
             ZCompr_conjugenasc.Value:=ZQCompr_conjugenasc.Value;
             ZCompr_conjugenatural.Value:=ZQCompr_conjugenatural.Value;
             ZCompr_conjugeidparti.Value:=ZQCompr_conjugeidparti.Value;
             ZCompr_conjugeemail.Value:=ZQCompr_conjugeemail.Value;
             ZCompr_conjugeassina.Value:=ZQCompr_conjugeassina.Value;
             ZCompr_conjugeempresa.Value:=ZQCompr_conjugeempresa.Value;
             ZCompr_conjugecargo.Value:=ZQCompr_conjugecargo.Value;
             ZCompr_conjugeadmissa.Value:=ZQCompr_conjugeadmissa.Value;
             ZCompr_conjugerenda.Value:=ZQCompr_conjugerenda.Value;
             ZCompr_conjugeregime.Value:=ZQCompr_conjugeregime.Value;
             ZCompr_conjugecomarca.Value:=ZQCompr_conjugecomarca.Value;
             ZCompr_conjugecidade_cart.Value:=ZQCompr_conjugecidade_cart.Value;
             ZCompr_conjugefolha.Value:=ZQCompr_conjugefolha.Value;
             ZCompr_conjugelivro.Value:=ZQCompr_conjugelivro.Value;
             ZCompr_conjugecartorio.Value:=ZQCompr_conjugecartorio.Value;
             ZCompr_conjugedata.Value:=ZQCompr_conjugedata.Value;
             ZCompr_conjuge.Post;
           end;
           ZQCompr_conjuge.next;


         end;
       end;
   { ZQProcuradores.First;
    if ZQProcuradores.RecordCount>0 then
    begin
      msg.Caption:='Procuradores                                                ';
      application.ProcessMessages;

      while not ZQProcuradores.Eof do
      begin
        if not ZProcuradores.Locate('idloteamento;nome',VarArrayOf([ZQProcuradoresidloteamento.Value,ZQProcuradoresnome.Value]),[]) then
        begin
          ZProcuradores.Insert;
          ZProcuradoresidloteamento.Value:=ZQProcuradoresidloteamento.Value;
          ZProcuradoresidparti.Value:=ZQProcuradoresidparti.Value;
          ZProcuradoresnome.Value:=ZQProcuradoresnome.Value;
          ZProcuradoresassina.Value:=ZQProcuradoresassina.Value;
          ZProcuradores.post;
        end;
        ZQProcuradores.Next;
      end;
    end;
    }
  ProgressBar1.Position:=0;
  msg.Caption:='Finalizado                                                   ';
  application.ProcessMessages;
  msg1.Caption:='Processo                ';
  application.ProcessMessages;
  showmessage('Fim de Processo');
//  inc(nlt);
//  xlote.Text:=inttostr(nlt);
//  Button1Click(self);
end;

procedure Tfrm_baixa.Button2Click(Sender: TObject);
begin
  close;
end;

procedure Tfrm_baixa.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  ZQVendedor.Close;
  ZQProcuradores.Close;
  ZQrepresentantes.Close;
  ZqParticipante.Close;
  ZQresponsavel.Close;
  ZQCompr_conjuge.Close;

  ZVendedor.Close;
  ZProcuradores.Close;
  Zrepresentantes.Close;
  ZParticipante.Close;
  Zresponsavel.Close;
  ZCompr_conjuge.Close;

end;

procedure Tfrm_baixa.FormShow(Sender: TObject);
begin
  ZQVendedor.open;
  ZQProcuradores.open;
  ZQrepresentantes.open;
  ZqParticipante.open;
  ZQresponsavel.open;
  ZQCompr_conjuge.open;

  ZVendedor.open;
  ZProcuradores.open;
  Zrepresentantes.open;
  ZParticipante.open;
  Zresponsavel.open;
  ZCompr_conjuge.open;
   xlote.SetFocus;
end;

procedure Tfrm_baixa.xloteExit(Sender: TObject);
begin
  Button1.Enabled:=true;
  Button1.SetFocus;


end;


procedure Tfrm_baixa.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQImovel', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImovelloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImovelquadra', 'quadra', TStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImovellote', 'lote', TStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImovellado', 'lado', TStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImovelarea', 'area', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImovelconfrontofrente', 'confrontofrente', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImovelconfrontofundo', 'confrontofundo', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImovelconfrontodireita', 'confrontodireita', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImovelconfrontoesquerda', 'confrontoesquerda', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImovelesquinanome', 'esquinanome', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImovelvalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '##,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImovelObervacaoloteamento', 'Obervacaoloteamento', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImovelcornomapa', 'cornomapa', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImovelnomeloteamento', 'nomeloteamento', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'loteamento_idloteamento', 'ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImovelidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImoveldisponivel', 'disponivel', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImovelmedidafrente', 'medidafrente', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImovelmedidafundo', 'medidafundo', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImovelmedidadedireita', 'medidadedireita', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImovelesquinalinha', 'esquinalinha', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImovelesquinacurva', 'esquinacurva', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImovelmedidaesquerda', 'medidaesquerda', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImoveldescricao', 'descricao', TStringField, fkCalculated, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImovelmatri', 'matri', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQImovel', 'ZQImovelmatricula', 'matricula', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZqParticipante', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipanteidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantenome_parte', 'nome_parte', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantedoc1', 'doc1', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantedoc2', 'doc2', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipanteendereco', 'endereco', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantebairro', 'bairro', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantecidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantecep', 'cep', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipanteende_cob', 'ende_cob', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantebairro_cob', 'bairro_cob', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantecidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantecep_cob', 'cep_cob', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipanteFone1', 'Fone1', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantefone2', 'fone2', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantefone3', 'fone3', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantetipopessoa', 'tipopessoa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipanteaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantenacionalidade', 'nacionalidade', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipanteemail', 'email', TStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantenaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantelocaldetrab', 'localdetrab', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipanteprofissao', 'profissao', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipanteestadocivil', 'estadocivil', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipanterenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipanteobservacao', 'observacao', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantecodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantecodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantecodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantecodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantefone4', 'fone4', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantecomplemento', 'complemento', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantecomplemento_cob', 'complemento_cob', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantecadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipanteNome_Firma', 'Nome_Firma', TStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipanteInativo', 'Inativo', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipanteSexo', 'Sexo', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantecc', 'cc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantecex', 'cex', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantejb', 'jb', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantejudicial', 'judicial', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantexquadra', 'xquadra', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZqParticipante', 'ZqParticipantexlote', 'xlote', TStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQresponsavel', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsavelidrepresenta_participante', 'idrepresenta_participante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsavelpaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsaveldocumen1', 'documen1', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsaveldocumen2', 'documen2', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsavelnasicmento', 'nasicmento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsavelcid_natural', 'cid_natural', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsavelnomerepres', 'nomerepres', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsavelemailrepres', 'emailrepres', TStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsavelprofiss', 'profiss', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsavelassina', 'assina', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsaveldata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsavelcartorio', 'cartorio', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsavellivro', 'livro', TStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsavelfolha', 'folha', TStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsavelcidade_cart', 'cidade_cart', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsavelcomarca', 'comarca', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsavelregime', 'regime', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsaveltrabalho', 'trabalho', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsaveladmissao', 'admissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsavelrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsavelnomeconjuge', 'nomeconjuge', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsaveldoc1conjuge', 'doc1conjuge', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsaveldoc2conjuge', 'doc2conjuge', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsavelnascconjuge', 'nascconjuge', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsavelassconjuge', 'assconjuge', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQresponsavel', 'ZQresponsavelcidconjuge', 'cidconjuge', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQComprador', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQComprador', 'ZQCompradoridcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQComprador', 'ZQCompradorpaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQComprador', 'ZQCompradorvenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQComprador', 'ZQCompradorpromissario', 'promissario', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQComprador', 'ZQCompradornomecomprador', 'nomecomprador', TStringField, fkLookup, 140, 0, False, '', '', '', '', 0, 'paticipante_idpaticipante', 'ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(Tfrm_baixa, 'ZQComprador', 'ZQCompradorpercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQComprador', 'ZQCompradorMarcar', 'Marcar', TSmallintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQVendedor', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQVendedor', 'ZQVendedoridvendedor', 'idvendedor', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQVendedor', 'ZQVendedorcorretor_idcorretor', 'corretor_idcorretor', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQVendedor', 'ZQVendedorvenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQDividido', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQDividido', 'ZQDivididoiddividido', 'iddividido', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQDividido', 'ZQDivididoparticipante', 'participante', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQDividido', 'ZQDivididopercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQDividido', 'ZQDivididoplanodeconta', 'planodeconta', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQDividido', 'ZQDivididotipodocumento', 'tipodocumento', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQDividido', 'ZQDivididoordem', 'ordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQConjuge', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugeidconjuge', 'idconjuge', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugenome', 'nome', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugedoc1', 'doc1', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugedoc2', 'doc2', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugenasc', 'nasc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugenatural', 'natural', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugeemail', 'email', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugeassina', 'assina', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugeempresa', 'empresa', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugecargo', 'cargo', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugeadmissa', 'admissa', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugerenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugeidparti', 'idparti', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugeregime', 'regime', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugecomarca', 'comarca', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugecidade_cart', 'cidade_cart', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugedata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugefolha', 'folha', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugelivro', 'livro', TStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugecartorio', 'cartorio', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQCompr_conjuge', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCompr_conjuge', 'ZQCompr_conjugeidconjuge', 'idconjuge', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCompr_conjuge', 'ZQCompr_conjugenome', 'nome', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCompr_conjuge', 'ZQCompr_conjugedoc1', 'doc1', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCompr_conjuge', 'ZQCompr_conjugedoc2', 'doc2', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCompr_conjuge', 'ZQCompr_conjugenasc', 'nasc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCompr_conjuge', 'ZQCompr_conjugenatural', 'natural', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCompr_conjuge', 'ZQCompr_conjugeidparti', 'idparti', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCompr_conjuge', 'ZQCompr_conjugeemail', 'email', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCompr_conjuge', 'ZQCompr_conjugeassina', 'assina', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCompr_conjuge', 'ZQCompr_conjugeempresa', 'empresa', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCompr_conjuge', 'ZQCompr_conjugecargo', 'cargo', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCompr_conjuge', 'ZQCompr_conjugeadmissa', 'admissa', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCompr_conjuge', 'ZQCompr_conjugerenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCompr_conjuge', 'ZQCompr_conjugeregime', 'regime', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCompr_conjuge', 'ZQCompr_conjugecomarca', 'comarca', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCompr_conjuge', 'ZQCompr_conjugecidade_cart', 'cidade_cart', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCompr_conjuge', 'ZQCompr_conjugefolha', 'folha', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCompr_conjuge', 'ZQCompr_conjugelivro', 'livro', TStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCompr_conjuge', 'ZQCompr_conjugecartorio', 'cartorio', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCompr_conjuge', 'ZQCompr_conjugedata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQProcuradores', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQProcuradores', 'ZQProcuradoresidProcuradores', 'idProcuradores', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQProcuradores', 'ZQProcuradoresidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQProcuradores', 'ZQProcuradoresidparti', 'idparti', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQProcuradores', 'ZQProcuradoresnome', 'nome', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQProcuradores', 'ZQProcuradoresassina', 'assina', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQRemes_Receb', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRemes_Receb', 'ZQRemes_Recebidremessa_receb', 'idremessa_receb', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRemes_Receb', 'ZQRemes_Recebremessa', 'remessa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRemes_Receb', 'ZQRemes_Recebidrec', 'idrec', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRemes_Receb', 'ZQRemes_Recebretorno_motivo', 'retorno_motivo', TStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRemes_Receb', 'ZQRemes_Recebacao', 'acao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRemes_Receb', 'ZQRemes_RecebNossonumero', 'Nossonumero', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRemes_Receb', 'ZQRemes_Recebocorrencia', 'ocorrencia', TStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRemes_Receb', 'ZQRemes_Recebdesc_motivo', 'desc_motivo', TStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRemes_Receb', 'ZQRemes_Recebcredito', 'credito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRemes_Receb', 'ZQRemes_Recebdt_ocorrencia', 'dt_ocorrencia', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQrepresentantes', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrepresentantes', 'ZQrepresentantesidrepresentantes', 'idrepresentantes', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrepresentantes', 'ZQrepresentantesincorporador_idincorporador', 'incorporador_idincorporador', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrepresentantes', 'ZQrepresentantesnomerepre', 'nomerepre', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrepresentantes', 'ZQrepresentantesdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZParticipante', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipanteidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantenome_parte', 'nome_parte', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantedoc1', 'doc1', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantedoc2', 'doc2', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipanteendereco', 'endereco', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantebairro', 'bairro', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantecidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantecep', 'cep', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipanteende_cob', 'ende_cob', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantebairro_cob', 'bairro_cob', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantecidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantecep_cob', 'cep_cob', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipanteFone1', 'Fone1', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantefone2', 'fone2', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantefone3', 'fone3', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantetipopessoa', 'tipopessoa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipanteaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantenacionalidade', 'nacionalidade', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipanteemail', 'email', TStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantenaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantelocaldetrab', 'localdetrab', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipanteprofissao', 'profissao', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipanteestadocivil', 'estadocivil', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipanterenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipanteobservacao', 'observacao', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantecodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantecodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantecodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantecodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantefone4', 'fone4', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantecomplemento', 'complemento', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantecomplemento_cob', 'complemento_cob', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantecadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipanteNome_Firma', 'Nome_Firma', TStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipanteInativo', 'Inativo', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipanteSexo', 'Sexo', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantecc', 'cc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantecex', 'cex', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantejb', 'jb', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantejudicial', 'judicial', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantexquadra', 'xquadra', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParticipante', 'ZParticipantexlote', 'xlote', TStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZResponsavel', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsavelidrepresenta_participante', 'idrepresenta_participante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsavelpaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsaveldocumen1', 'documen1', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsaveldocumen2', 'documen2', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsavelnasicmento', 'nasicmento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsavelcid_natural', 'cid_natural', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsavelnomerepres', 'nomerepres', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsavelemailrepres', 'emailrepres', TStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsavelprofiss', 'profiss', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsavelassina', 'assina', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsaveldata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsavelcartorio', 'cartorio', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsavellivro', 'livro', TStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsavelfolha', 'folha', TStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsavelcidade_cart', 'cidade_cart', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsavelcomarca', 'comarca', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsavelregime', 'regime', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsaveltrabalho', 'trabalho', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsaveladmissao', 'admissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsavelrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsavelnomeconjuge', 'nomeconjuge', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsaveldoc1conjuge', 'doc1conjuge', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsaveldoc2conjuge', 'doc2conjuge', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsavelnascconjuge', 'nascconjuge', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsavelassconjuge', 'assconjuge', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZResponsavel', 'ZResponsavelcidconjuge', 'cidconjuge', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZConjuge', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConjuge', 'ZConjugeidconjuge', 'idconjuge', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConjuge', 'ZConjugenome', 'nome', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConjuge', 'ZConjugedoc1', 'doc1', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConjuge', 'ZConjugedoc2', 'doc2', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConjuge', 'ZConjugenasc', 'nasc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConjuge', 'ZConjugenatural', 'natural', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConjuge', 'ZConjugeidparti', 'idparti', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConjuge', 'ZConjugeemail', 'email', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConjuge', 'ZConjugeassina', 'assina', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConjuge', 'ZConjugeempresa', 'empresa', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConjuge', 'ZConjugecargo', 'cargo', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConjuge', 'ZConjugeadmissa', 'admissa', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConjuge', 'ZConjugerenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConjuge', 'ZConjugeregime', 'regime', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConjuge', 'ZConjugecomarca', 'comarca', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConjuge', 'ZConjugecidade_cart', 'cidade_cart', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConjuge', 'ZConjugefolha', 'folha', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConjuge', 'ZConjugelivro', 'livro', TStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConjuge', 'ZConjugecartorio', 'cartorio', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConjuge', 'ZConjugedata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZVendedor', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZVendedor', 'ZVendedoridvendedor', 'idvendedor', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZVendedor', 'ZVendedorcorretor_idcorretor', 'corretor_idcorretor', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZVendedor', 'ZVendedorvenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZCompr_conjuge', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZCompr_conjuge', 'ZCompr_conjugeidconjuge', 'idconjuge', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZCompr_conjuge', 'ZCompr_conjugenome', 'nome', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZCompr_conjuge', 'ZCompr_conjugedoc1', 'doc1', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZCompr_conjuge', 'ZCompr_conjugedoc2', 'doc2', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZCompr_conjuge', 'ZCompr_conjugenasc', 'nasc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZCompr_conjuge', 'ZCompr_conjugenatural', 'natural', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZCompr_conjuge', 'ZCompr_conjugeidparti', 'idparti', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZCompr_conjuge', 'ZCompr_conjugeemail', 'email', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZCompr_conjuge', 'ZCompr_conjugeassina', 'assina', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZCompr_conjuge', 'ZCompr_conjugeempresa', 'empresa', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZCompr_conjuge', 'ZCompr_conjugecargo', 'cargo', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZCompr_conjuge', 'ZCompr_conjugeadmissa', 'admissa', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZCompr_conjuge', 'ZCompr_conjugerenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZCompr_conjuge', 'ZCompr_conjugeregime', 'regime', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZCompr_conjuge', 'ZCompr_conjugecomarca', 'comarca', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZCompr_conjuge', 'ZCompr_conjugecidade_cart', 'cidade_cart', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZCompr_conjuge', 'ZCompr_conjugefolha', 'folha', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZCompr_conjuge', 'ZCompr_conjugelivro', 'livro', TStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZCompr_conjuge', 'ZCompr_conjugecartorio', 'cartorio', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZCompr_conjuge', 'ZCompr_conjugedata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZProcuradores', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZProcuradores', 'ZProcuradoresidProcuradores', 'idProcuradores', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZProcuradores', 'ZProcuradoresidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZProcuradores', 'ZProcuradoresidparti', 'idparti', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZProcuradores', 'ZProcuradoresnome', 'nome', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZProcuradores', 'ZProcuradoresassina', 'assina', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'Zrepresentantes', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrepresentantes', 'Zrepresentantesidrepresentantes', 'idrepresentantes', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrepresentantes', 'Zrepresentantesincorporador_idincorporador', 'incorporador_idincorporador', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrepresentantes', 'Zrepresentantesnomerepre', 'nomerepre', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrepresentantes', 'Zrepresentantesdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
