unit importar;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset,
  Mask, ComCtrls, Grids, DBGrids, ZConnection;

type
  Tfrm_baixa = class(TForm)
    Button1: TButton;
    Button2: TButton;
    ZQLoteamento: TZQuery;
    DS_Loteamento: TDataSource;
    DS_Imovel: TDataSource;
    ZQImovel: TZQuery;

























    ZQincorp_loteame: TZQuery;
    DSincorp_loteame: TDataSource;
    DS_Quadras: TDataSource;
    ZQQuadras: TZQuery;




    DS_Participante: TDataSource;
    ZqParticipante: TZQuery;
    DS_Responsavel: TDataSource;
    ZQresponsavel: TZQuery;
    DS_Comprador: TDataSource;
    ZQComprador: TZQuery;






    DS_Receb_baixa: TDataSource;
    ZQReceb_Baixa: TZQuery;















    ZQVendedor: TZQuery;



    DS_Vendedor: TDataSource;
    DS_Recebimento: TDataSource;
    ZQRecebimento: TZQuery;
    DS_Venda: TDataSource;
    ZQVenda: TZQuery;
    DS_ReBxHi: TDataSource;
    ZQReBxHi: TZQuery;









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


































    Datarecebimento_historico_cessao: TDataSource;
    ZQrecebimento_historico_cessao: TZQuery;





































    ZQrecebimento_historico_excluidos: TZQuery;
    Datarecebimento_historico_excluidos: TDataSource;
































    ZQRemes_Receb: TZQuery;
    DS_Remes_Receb: TDataSource;
    ZQCheque: TZQuery;

















    DS_Cheque: TDataSource;
    ZQAditamento: TZQuery;
    DataZQAditamento: TDataSource;











    Dataquitacao: TDataSource;
    ZQquitacao: TZQuery;











    DataZQhistorico_cessao: TDataSource;
    ZQhistorico_cessao: TZQuery;



    ZQComprador_cessao: TZQuery;
    DataZQComprador_cessao: TDataSource;











    ZQParcelas: TZQuery;
    Dataparcelas: TDataSource;








    ZQrepresentantes: TZQuery;
    Datarepresentantes: TDataSource;




    nome: TMaskEdit;




















    ProgressBar1: TProgressBar;







































































































    msg: TLabel;
    msg1: TLabel;
    ZQConta_bancaria: TZQuery;
    DS_ZQConta_bancaria: TDataSource;























    DBGrid1: TDBGrid;
    DBGrid2: TDBGrid;
    DBGrid3: TDBGrid;
    DBGrid4: TDBGrid;
    DBGrid5: TDBGrid;
    DBGrid6: TDBGrid;
    DBGrid7: TDBGrid;
    DBGrid8: TDBGrid;
    DBGrid9: TDBGrid;
    DBGrid10: TDBGrid;
    DBGrid11: TDBGrid;
    DBGrid12: TDBGrid;
    DBGrid13: TDBGrid;
    DBGrid14: TDBGrid;
    DBGrid15: TDBGrid;
    DBGrid16: TDBGrid;
    DBGrid17: TDBGrid;
    DBGrid18: TDBGrid;
    DBGrid19: TDBGrid;
    DBGrid20: TDBGrid;
    DBGrid21: TDBGrid;
    DBGrid22: TDBGrid;
    zconeccao: TZConnection;
    Zloteamento: TZQuery;




















    ZParticipante: TZQuery;










































    ZImovel: TZQuery;























    ZResponsavel: TZQuery;


























    ZConjuge: TZQuery;




















    ZComprador: TZQuery;






    ZRecebimento: TZQuery;





































    ZReceb_baixa: TZQuery;
















    Zcontabancaria: TZQuery;























    ZVenda: TZQuery;







    ZComprador_cessao: TZQuery;











    ZRebxHi: TZQuery;










    ZAditamento: TZQuery;











    ZQuadras: TZQuery;




    ZVendedor: TZQuery;



    ZCompr_conjuge: TZQuery;




















    Zrecebimento_historico_cessao: TZQuery;





































    ZProcuradores: TZQuery;





    ZRemes_Receb: TZQuery;










    Zrecebimento_historico_excluidos: TZQuery;
































    ZQuitacao: TZQuery;











    Zhistorico_cessao: TZQuery;



    ZParcelas: TZQuery;








    Zincorp_loteame: TZQuery;
    Zrepresentantes: TZQuery;




    Zcheque: TZQuery;























    ZConta_bancaria: TZQuery;































    ProgressBar2: TProgressBar;











    procedure Button1Click(Sender: TObject);
    procedure xloteEnter(Sender: TObject);
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

  ZQLoteamento.close;
  ZQLoteamento.SQL.Clear;
  ZQLoteamento.SQL.add('Select * from Loteamento where idloteamento='+quotedstr(xlote.text));
  ZQLoteamento.open;
  ZQLoteamento.first;
  ProgressBar1.Position:=0;
  nlt:= strtoint(xlote.text);
  while not ZQLoteamento.Eof do
  begin
    zqimovel.First;
    if zqimovel.RecordCount>0 then
    begin
      msg1.Caption:='Imoveis                ';
      application.ProcessMessages;
      ProgressBar2.Max:=zqimovel.RecordCount;
      ProgressBar2.Position:=0;
      while not zqimovel.eof do
      begin
        ProgressBar2.Position:=zqimovel.RecNo;
        application.ProcessMessages;
       { zqrecebimento.first;

        if zqrecebimento.RecordCount>0 then
        begin
          msg.Caption:='recebimento              ';
          application.ProcessMessages;
          ProgressBar1.Max:=zqrecebimento.RecordCount;
          avanco:=100/zqrecebimento.RecordCount;
          while not zqrecebimento.Eof do
          begin
            application.ProcessMessages;
            avanco:=avanco+1;
            ProgressBar1.Position:=Trunc(avanco);

            msg.Caption:='Hist.recebimento        ';
            application.ProcessMessages;
            ZQReBxHi.First;
            if ZQReBxHi.RecordCount>0 then
            begin
              while not ZQReBxHi.eof do
              begin
                msg.Caption:='Bx.recebimento         ';
                application.ProcessMessages;

                ZQReceb_Baixa.First;
                if ZQReceb_Baixa.RecordCount>0 then
                begin
                  while not ZQReceb_Baixa.Eof do
                  begin
                    if not ZReceb_Baixa.Locate('RefBaixa;Dt_rec;dataref',VarArrayOf([ZQReceb_baixaRefBaixa.Value,ZQReceb_baixaDt_rec.Value,ZQReceb_baixadataref.Value]),[]) then
                    begin
                      ZReceb_Baixa.insert;
                      ZReceb_baixaDt_rec.Value:=ZQReceb_baixaDt_rec.Value;
                      ZReceb_baixaVr_rec.Value:=ZQReceb_baixaVr_rec.Value;
                      ZReceb_baixaTipDoc.Value:=ZQReceb_baixaTipDoc.Value;
                      ZReceb_baixaDocum.Value:=ZQReceb_baixaDocum.Value;
                      ZReceb_baixaJuros_Vr.Value:=ZQReceb_baixaJuros_Vr.Value;
                      ZReceb_baixaDesc_Vr.Value:=ZQReceb_baixaDesc_Vr.Value;
                      ZReceb_baixaRefBaixa.Value:=ZQReceb_baixaRefBaixa.Value;
                      ZReceb_baixaobsebx.Value:=ZQReceb_baixaobsebx.Value;
                      ZReceb_baixaCodcontabil.Value:=ZQReceb_baixaCodcontabil.Value;
                      ZReceb_baixadataref.Value:=ZQReceb_baixadataref.Value;
                      ZReceb_baixavencimento.Value:=ZQReceb_baixavencimento.Value;
                      ZReceb_baixaSubstituicao.Value:=ZQReceb_baixaSubstituicao.Value;
                      ZReceb_baixasq.Value:=ZQReceb_baixasq.Value;
                      ZReceb_baixaVR_AB.Value:=ZQReceb_baixaVR_AB.Value;
                      ZReceb_baixavalor_parcela.Value:=ZQReceb_baixavalor_parcela.Value;
                      ZReceb_Baixa.post;
                    end;
                    ZQReceb_Baixa.Next;
                  end;
                end;
                if not ZReBxHi.Locate('refer;idrecib;data',VarArrayOf([ZQReBxHirefer.Value,ZQReBxHiidrecib.Value,ZQReBxHidata.Value]),[]) then
                begin
                  ZReBxHi.Insert;
                  ZReBxHirefer.Value:=ZQReBxHirefer.Value;
                  ZReBxHiidrecib.Value:=ZQReBxHiidrecib.Value;
                  ZReBxHivalor.Value:=ZQReBxHivalor.Value;
                  ZReBxHidescon.Value:=ZQReBxHidescon.Value;
                  ZReBxHijuros.Value:=ZQReBxHijuros.Value;
                  ZReBxHipercent_usado.Value:=ZQReBxHipercent_usado.Value;
                  ZReBxHidata.Value:=ZQReBxHidata.Value;
                  ZReBxHisq.Value:=ZQReBxHisq.Value;
                  ZReBxHivalor_parcela.Value:=ZQReBxHivalor_parcela.Value;
                  ZReBxHi.post;
                end;
                ZQReBxHi.next;
              end;
              ZQrecebimento_historico_cessao.First;
              if ZQrecebimento_historico_cessao.RecordCount>0 then
              begin
                msg.Caption:='Hist. cessao recebimento';
                application.ProcessMessages;

                while not ZQrecebimento_historico_cessao.eof do
                begin
                  if not Zrecebimento_historico_cessao.Locate('idrecebimento',ZQrecebimento_historico_cessaoidrecebimento.Value,[]) then
                  begin
                    Zrecebimento_historico_cessao.Insert;
                    Zrecebimento_historico_cessaoidrecebimento.Value:=ZQrecebimento_historico_cessaoidrecebimento.Value;
                    Zrecebimento_historico_cessaodocumento.Value:=ZQrecebimento_historico_cessaodocumento.Value;
                    Zrecebimento_historico_cessaocliente.Value:=ZQrecebimento_historico_cessaocliente.Value;
                    Zrecebimento_historico_cessaousuario.Value:=ZQrecebimento_historico_cessaousuario.Value;
                    Zrecebimento_historico_cessaoDt_Entrada.Value:=ZQrecebimento_historico_cessaoDt_Entrada.Value;
                    Zrecebimento_historico_cessaoDt_Vencimento.Value:=ZQrecebimento_historico_cessaoDt_Vencimento.Value;
                    Zrecebimento_historico_cessaoValor.Value:=ZQrecebimento_historico_cessaoValor.Value;
                    Zrecebimento_historico_cessaoObserv.Value:=ZQrecebimento_historico_cessaoObserv.Value;
                    Zrecebimento_historico_cessaoVrDoc.Value:=ZQrecebimento_historico_cessaoVrDoc.Value;
                    Zrecebimento_historico_cessaoordem.Value:=ZQrecebimento_historico_cessaoordem.Value;
                    Zrecebimento_historico_cessaoTipDoc.Value:=ZQrecebimento_historico_cessaoTipDoc.Value;
                    Zrecebimento_historico_cessaosaldo.Value:=ZQrecebimento_historico_cessaosaldo.Value;
                    Zrecebimento_historico_cessaomarcar.Value:=ZQrecebimento_historico_cessaomarcar.Value;
                    Zrecebimento_historico_cessaoRefBaixa.Value:=ZQrecebimento_historico_cessaoRefBaixa.Value;
                    Zrecebimento_historico_cessaorefvinda.Value:=ZQrecebimento_historico_cessaorefvinda.Value;
                    Zrecebimento_historico_cessaocontabil.Value:=ZQrecebimento_historico_cessaocontabil.Value;
                    Zrecebimento_historico_cessaoempresa.Value:=ZQrecebimento_historico_cessaoempresa.Value;
                    Zrecebimento_historico_cessaocustodaparcela.Value:=ZQrecebimento_historico_cessaocustodaparcela.Value;
                    Zrecebimento_historico_cessaoorigem.Value:=ZQrecebimento_historico_cessaoorigem.Value;
                    Zrecebimento_historico_cessaoadversa.Value:=ZQrecebimento_historico_cessaoadversa.Value;
                    Zrecebimento_historico_cessaorecpag.Value:=ZQrecebimento_historico_cessaorecpag.Value;
                    Zrecebimento_historico_cessaonumordem.Value:=ZQrecebimento_historico_cessaonumordem.Value;
                    Zrecebimento_historico_cessaoidloteamento.Value:=ZQrecebimento_historico_cessaoidloteamento.Value;
                    Zrecebimento_historico_cessaovenda_idvenda.Value:=ZQrecebimento_historico_cessaovenda_idvenda.Value;
                    Zrecebimento_historico_cessaoquadralote.Value:=ZQrecebimento_historico_cessaoquadralote.Value;
                    Zrecebimento_historico_cessaonumboleto.Value:=ZQrecebimento_historico_cessaonumboleto.Value;
                    Zrecebimento_historico_cessaoSubstituicao.Value:=ZQrecebimento_historico_cessaoSubstituicao.Value;
                    Zrecebimento_historico_cessaosq.Value:=ZQrecebimento_historico_cessaosq.Value;
                    Zrecebimento_historico_cessaosomar.Value:=ZQrecebimento_historico_cessaosomar.Value;
                    Zrecebimento_historico_cessaonomeadversa.Value:=ZQrecebimento_historico_cessaonomeadversa.Value;
                    Zrecebimento_historico_cessaoReajustado.Value:=ZQrecebimento_historico_cessaoReajustado.Value;
                    Zrecebimento_historico_cessaoData_reajuste.Value:=ZQrecebimento_historico_cessaoData_reajuste.Value;
                    Zrecebimento_historico_cessaoProximo_Reajuste.Value:=ZQrecebimento_historico_cessaoProximo_Reajuste.Value;
                    Zrecebimento_historico_cessaoParcelas_fixas.Value:=ZQrecebimento_historico_cessaoParcelas_fixas.Value;
                    Zrecebimento_historico_cessaodata_cessao.Value:=ZQrecebimento_historico_cessaodata_cessao.Value;
                    Zrecebimento_historico_cessaohora_cessao.Value:=ZQrecebimento_historico_cessaohora_cessao.Value;
                    Zrecebimento_historico_cessao.post;
                  end;
                  ZQrecebimento_historico_cessao.next;
                end;
              end;

              ZQrecebimento_historico_excluidos.First;
              if ZQrecebimento_historico_excluidos.RecordCount>0 then
              begin
                msg.Caption:='Hist. excl. recebimento';
                application.ProcessMessages;

                while not ZQrecebimento_historico_excluidos.eof do
                begin
                  if not Zrecebimento_historico_excluidos.Locate('idrecebimento',ZQrecebimento_historico_excluidosidrecebimento.Value,[]) then
                  begin
                    Zrecebimento_historico_excluidos.Insert;
                    Zrecebimento_historico_excluidosidrecebimento.Value:=ZQrecebimento_historico_excluidosidrecebimento.Value;
                    Zrecebimento_historico_excluidosdocumento.Value:=ZQrecebimento_historico_excluidosdocumento.Value;
                    Zrecebimento_historico_excluidoscliente.Value:=ZQrecebimento_historico_excluidoscliente.Value;
                    Zrecebimento_historico_excluidosusuario.Value:=ZQrecebimento_historico_excluidosusuario.Value;
                    Zrecebimento_historico_excluidosDt_Entrada.Value:=ZQrecebimento_historico_excluidosDt_Entrada.Value;
                    Zrecebimento_historico_excluidosDt_Vencimento.Value:=ZQrecebimento_historico_excluidosDt_Vencimento.Value;
                    Zrecebimento_historico_excluidosValor.Value:=ZQrecebimento_historico_excluidosValor.Value;
                    Zrecebimento_historico_excluidosObserv.Value:=ZQrecebimento_historico_excluidosObserv.Value;
                    Zrecebimento_historico_excluidosVrDoc.Value:=ZQrecebimento_historico_excluidosVrDoc.Value;
                    Zrecebimento_historico_excluidosordem.Value:=ZQrecebimento_historico_excluidosordem.Value;
                    Zrecebimento_historico_excluidosTipDoc.Value:=ZQrecebimento_historico_excluidosTipDoc.Value;
                    Zrecebimento_historico_excluidossaldo.Value:=ZQrecebimento_historico_excluidossaldo.Value;
                    Zrecebimento_historico_excluidosmarcar.Value:=ZQrecebimento_historico_excluidosmarcar.Value;
                    Zrecebimento_historico_excluidosRefBaixa.Value:=ZQrecebimento_historico_excluidosRefBaixa.Value;
                    Zrecebimento_historico_excluidosrefvinda.Value:=ZQrecebimento_historico_excluidosrefvinda.Value;
                    Zrecebimento_historico_excluidoscontabil.Value:=ZQrecebimento_historico_excluidoscontabil.Value;
                    Zrecebimento_historico_excluidosempresa.Value:=ZQrecebimento_historico_excluidosempresa.Value;
                    Zrecebimento_historico_excluidoscustodaparcela.Value:=ZQrecebimento_historico_excluidoscustodaparcela.Value;
                    Zrecebimento_historico_excluidosorigem.Value:=ZQrecebimento_historico_excluidosorigem.Value;
                    Zrecebimento_historico_excluidosadversa.Value:=ZQrecebimento_historico_excluidosadversa.Value;
                    Zrecebimento_historico_excluidosrecpag.Value:=ZQrecebimento_historico_excluidosrecpag.Value;
                    Zrecebimento_historico_excluidosnumordem.Value:=ZQrecebimento_historico_excluidosnumordem.Value;
                    Zrecebimento_historico_excluidosidloteamento.Value:=ZQrecebimento_historico_excluidosidloteamento.Value;
                    Zrecebimento_historico_excluidosvenda_idvenda.Value:=ZQrecebimento_historico_excluidosvenda_idvenda.Value;
                    Zrecebimento_historico_excluidosquadralote.Value:=ZQrecebimento_historico_excluidosquadralote.Value;
                    Zrecebimento_historico_excluidosnumboleto.Value:=ZQrecebimento_historico_excluidosnumboleto.Value;
                    Zrecebimento_historico_excluidosSubstituicao.Value:=ZQrecebimento_historico_excluidosSubstituicao.Value;
                    Zrecebimento_historico_excluidossq.Value:=ZQrecebimento_historico_excluidossq.Value;
                    Zrecebimento_historico_excluidossomar.Value:=ZQrecebimento_historico_excluidossomar.Value;
                    Zrecebimento_historico_excluidosmotivo.Value:=ZQrecebimento_historico_excluidosmotivo.Value;
                    Zrecebimento_historico_excluidosProximo_Reajuste.Value:=ZQrecebimento_historico_excluidosProximo_Reajuste.Value;
                    Zrecebimento_historico_excluidos.post;
                  end;
                  ZQrecebimento_historico_excluidos.next;
                end;
              end;
              ZQRemes_Receb.First;
              if ZQRemes_Receb.RecordCount>0 then
              begin
                 while not ZQRemes_Receb.Eof do
                 begin
                   if not ZRemes_Receb.Locate('idrec,remessa',VarArrayOf([ZQRemes_Recebidrec.Value,ZQRemes_Recebremessa.Value]),[]) then
                   begin
                     ZRemes_Receb.Insert;
                     ZRemes_Recebremessa.Value:=ZQRemes_Recebremessa.Value;
                     ZRemes_Recebidrec.Value:=ZQRemes_Recebidrec.Value;
                     ZRemes_Recebretorno_motivo.Value:=ZQRemes_Recebretorno_motivo.Value;
                     ZRemes_Recebacao.Value:=ZQRemes_Recebacao.Value;
                     ZRemes_RecebNossonumero.Value:=ZQRemes_RecebNossonumero.Value;
                     ZRemes_Recebocorrencia.Value:=ZQRemes_Recebocorrencia.Value;
                     ZRemes_Recebdesc_motivo.Value:=ZQRemes_Recebdesc_motivo.Value;
                     ZRemes_Recebcredito.Value:=ZQRemes_Recebcredito.Value;
                     ZRemes_Recebdt_ocorrencia.Value:=ZQRemes_Recebdt_ocorrencia.Value;
                     ZRemes_Receb.post;
                   end;
                   ZQRemes_Receb.next;
                 end;
              end;
            end;
            ZQCheque.First;
            if ZQCheque.RecordCount>0 then
            begin
              msg.Caption:='Cheque                   ';
              application.ProcessMessages;

              while not ZQCheque.eof do
              begin
                if not ZCheque.Locate('idrecebimento,numero',VarArrayOf([ZQChequeidrecebimento.Value,ZQChequenumero.Value]),[]) then
                begin
                  ZCheque.Insert;
                  ZChequebanco.value:=ZQChequebanco.value;
                  ZChequeagencia.value:=ZQChequeagencia.value;
                  ZChequedono.Value:=ZQChequedono.Value;
                  ZChequeCPF_CNPJ.Value:=ZQChequeCPF_CNPJ.Value;
                  ZChequevalor.Value:=ZQChequevalor.Value;
                  ZChequeemissao.value:=ZQChequeemissao.value;
                  ZChequedeposito.Value:=ZQChequedeposito.Value;
                  ZChequealias.Value:=ZQChequealias.Value;
                  ZChequeprorrogado.Value:=ZQChequeprorrogado.Value;
                  ZChequeconta.Value:=ZQChequeconta.Value;
                  ZChequeidrecebimento.Value:=ZQChequeidrecebimento.Value;
                  ZChequeidpagamento.Value:=ZQChequeidpagamento.Value;
                  ZChequenumero.Value:=ZQChequenumero.Value;
                  ZChequeidvinculo.Value:=ZQChequeidvinculo.Value;
                  ZChequeSubstituicao.Value:=ZQChequeSubstituicao.Value;
                  ZChequesq.Value:=ZQChequesq.Value;
                  ZCheque.Post;
                end;
                ZQCheque.next;
              end;
            end;
            ZQVenda.First;
            if ZQVenda.RecordCount>0 then
            begin
              msg.Caption:='Venda              ';
              application.ProcessMessages;

              while not ZQVenda.eof do
              begin
                ZQVendedor.First;
                if ZQVendedor.RecordCount>0 then
                begin
                  msg.Caption:='Vendedor               ';
                  application.ProcessMessages;

                  while not ZQVendedor.Eof do
                  begin
                    if not ZVendedor.Locate('idvendedor',ZQVendedoridvendedor.Value,[]) then
                    begin
                      ZVendedor.Insert;
                      ZVendedoridvendedor.Value:=ZQVendedoridvendedor.Value;
                      ZVendedorcorretor_idcorretor.Value:=ZQVendedorcorretor_idcorretor.Value;
                      ZVendedorvenda_idvenda.Value:=ZQVendedorvenda_idvenda.Value;
                      ZVendedor.Post;
                    end;
                    ZQVendedor.next;

                  end;
                end;
                ZQAditamento.first;
                if ZQAditamento.RecordCount>0 then
                begin
                   msg.Caption:='Aditamento                ';
                   application.ProcessMessages;

                  while not ZQAditamento.eof do
                  begin
                    if not ZAditamento.Locate('controle;idvenda',VarArrayOf([ZQAditamentocontrole.Value,ZQAditamentoidvenda.Value]),[]) then
                    begin
                      ZAditamento.Insert;
                      ZAditamentoidvenda.Value:=ZQAditamentoidvenda.Value;
                      ZAditamentodatavenda.Value:=ZQAditamentodatavenda.Value;
                      ZAditamentodatacessao.Value:=ZQAditamentodatacessao.Value;
                      ZAditamentoimovel.Value:=ZQAditamentoimovel.Value;
                      ZAditamentovalorvenda.Value:=ZQAditamentovalorvenda.Value;
                      ZAditamentoidcomprador.Value:=ZQAditamentoidcomprador.Value;
                      ZAditamentoidparticipante.Value:=ZQAditamentoidparticipante.Value;
                      ZAditamentonome_parte.Value:=ZQAditamentonome_parte.Value;
                      ZAditamentohistorico.Value:=ZQAditamentohistorico.Value;
                      ZAditamentohora.Value:=ZQAditamentohora.Value;
                      ZAditamento.Post;
                    end;
                    ZQAditamento.next;
                  end;
                end;
                ZQhistorico_cessao.first;
                if ZQhistorico_cessao.RecordCount>0 then
                begin
                  msg.Caption:='Hist.Cessao recebimento';
                  application.ProcessMessages;

                  while not ZQhistorico_cessao.eof do
                  begin
                    if not Zhistorico_cessao.Locate('idhistorico_cessao;idvenda',VarArrayOf([ZQhistorico_cessaoidhistorico_cessao.Value,ZQhistorico_cessaoidvenda.Value]),[]) then
                    begin
                      Zhistorico_cessao.insert;
                      Zhistorico_cessaoidvenda.Value:= ZQhistorico_cessaoidvenda.Value;
                      Zhistorico_cessaohistorico.Value:= ZQhistorico_cessaohistorico.Value;
                      Zhistorico_cessao.Post;
                    end;
                    ZQhistorico_cessao.Next;
                  end;
                end;
                ZQQuitacao.first;
                if ZQQuitacao.RecordCount>0 then
                begin
                  msg.Caption:='Quita��o                    ';
                  application.ProcessMessages;

                  while not ZQQuitacao.eof do
                  begin
                    if not ZQuitacao.Locate('controle;idvenda',VarArrayOf([ZQQuitacaocontrole.Value,ZQQuitacaoidvenda.Value]),[]) then
                    begin
                      ZQuitacao.Insert;
                      ZQuitacaoidvenda.Value:=ZQQuitacaoidvenda.Value;
                      ZQuitacaodatavenda.Value:=ZQQuitacaodatavenda.Value;
                      ZQuitacaodatacessao.Value:=ZQQuitacaodatacessao.Value;
                      ZQuitacaoimovel.Value:=ZQQuitacaoimovel.Value;
                      ZQuitacaovalorvenda.Value:=ZQQuitacaovalorvenda.Value;
                      ZQuitacaoidcomprador.Value:=ZQQuitacaoidcomprador.Value;
                      ZQuitacaoidparticipante.Value:=ZQQuitacaoidparticipante.Value;
                      ZQuitacaonome_parte.Value:=ZQQuitacaonome_parte.Value;
                      ZQuitacaohistorico.Value:=ZQQuitacaohistorico.Value;
                      ZQuitacaohora.Value:=ZQQuitacaohora.Value;
                      ZQuitacao.Post;
                    end;
                    ZQQuitacao.Next;
                  end;
                end;

                ZQComprador.first;
                if ZQComprador.RecordCount>0 then
                begin
                  msg.Caption:='Comprador                    ';
                  application.ProcessMessages;

                  while not ZQComprador.eof do
                  begin
                    if not ZComprador.Locate('paticipante_idpaticipante;venda_idvenda',VarArrayOf([ZQCompradorpaticipante_idpaticipante.Value,ZQCompradorvenda_idvenda.Value]),[]) then
                    begin
                      ZComprador.Insert;
                      ZCompradorpaticipante_idpaticipante.Value:=ZQCompradorpaticipante_idpaticipante.Value;
                      ZCompradorvenda_idvenda.Value:=ZQCompradorvenda_idvenda.Value;
                      ZCompradorpromissario.Value:=ZQCompradorpromissario.Value;
                      ZCompradorpercentual.Value:=ZQCompradorpercentual.Value;
                      ZCompradorMarcar.Value:=ZQCompradorMarcar.Value;
                      ZComprador.Post;
                    end;
                    ZQComprador.next;
                  end;
                end;

                ZQComprador_cessao.first;
                if ZQComprador_cessao.RecordCount>0 then
                begin
                  msg.Caption:='Comprador Cess�o         ';
                  application.ProcessMessages;

                  while not ZQComprador_cessao.eof do
                  begin
                    if not ZComprador_cessao.Locate('idvenda;datacessao',VarArrayOf([ZQComprador_cessaoidvenda.Value,ZQComprador_cessaodatacessao.Value]),[]) then
                    begin
                      ZComprador_cessao.Insert;
                      ZComprador_cessaoidvenda.Value:=ZQComprador_cessaoidvenda.Value;
                      ZComprador_cessaodatavenda.Value:=ZQComprador_cessaodatavenda.Value;
                      ZComprador_cessaodatacessao.Value:=ZQComprador_cessaodatacessao.Value;
                      ZComprador_cessaoimovel.Value:=ZQComprador_cessaoimovel.Value;
                      ZComprador_cessaovalorvenda.Value:=ZQComprador_cessaovalorvenda.Value;
                      ZComprador_cessaoidcomprador.Value:=ZQComprador_cessaoidcomprador.Value;
                      ZComprador_cessaoidparticipante.Value:=ZQComprador_cessaoidparticipante.Value;
                      ZComprador_cessaonome_parte.Value:=ZQComprador_cessaonome_parte.Value;
                      ZComprador_cessaohistorico.Value:=ZQComprador_cessaohistorico.Value;
                      ZComprador_cessaohora_cessao.Value:=ZQComprador_cessaohora_cessao.Value;
                      ZComprador_cessao.Post;
                    end;
                    ZQComprador_cessao.Next;
                  end;
                end;

                ZQParcelas.first;
                if ZQParcelas.RecordCount>0 then
                begin
                  msg.Caption:='Parcelas                 ';
                  application.ProcessMessages;

                  while not ZQParcelas.eof do
                  begin
                    ZParcelas.Insert;
                    ZParcelasvenda_idvenda.Value:=ZQParcelasvenda_idvenda.Value;
                    ZParcelasordem.Value:=ZQParcelasordem.Value;
                    ZParcelasdatavenci.Value:=ZQParcelasdatavenci.Value;
                    ZParcelasvalor.Value:=ZQParcelasvalor.Value;
                    ZParcelassaldo.Value:=ZQParcelassaldo.Value;
                    ZParcelasvalorvenda.Value:=ZQParcelasvalorvenda.Value;
                    ZParcelasnobanco.Value:=ZQParcelasnobanco.Value;
                    ZParcelas.post;
                    ZQParcelas.next;
                  end;
                end;
                if not ZVenda.Locate('idvenda',ZQVendaidvenda.Value,[]) then
                begin
                  ZVenda.Insert;
                  ZVendaidvenda.Value:=ZQVendaidvenda.Value;
                  ZVendadatavenda.Value:=ZQVendadatavenda.Value;
                  ZVendaimovel.Value:=ZQVendaimovel.Value;
                  ZVendavalorvenda.Value:=ZQVendavalorvenda.Value;
                  ZVendaforma_reajuste.Value:=ZQVendaforma_reajuste.Value;
                  ZVendatabela_Price.Value:=ZQVendatabela_Price.Value;
                  ZVendaEscriturado.Value:=ZQVendaEscriturado.Value;
                  ZVenda.Post;
                end;
                ZQVenda.next;
              end;
            end;
            if not zrecebimento.Locate('idrecebimento',zqrecebimentoidrecebimento.Value,[]) then
            begin
              zrecebimento.Insert;
              zrecebimentoidrecebimento.Value:=zqrecebimentoidrecebimento.Value;
              zrecebimentodocumento.Value:=zqrecebimentodocumento.Value;
              zrecebimentocliente.Value:=zqrecebimentocliente.Value;
              zrecebimentousuario.Value:=zqrecebimentousuario.Value;
              zrecebimentoDt_Entrada.Value:=zqrecebimentoDt_Entrada.Value;
              zrecebimentoDt_Vencimento.Value:=zqrecebimentoDt_Vencimento.Value;
              zrecebimentoValor.Value:=zqrecebimentoValor.Value;
              zrecebimentoObserv.Value:=zqrecebimentoObserv.Value;
              zrecebimentoVrDoc.Value:=zqrecebimentoVrDoc.Value;
              zrecebimentoordem.Value:=zqrecebimentoordem.Value;
              zrecebimentoTipDoc.Value:=zqrecebimentoTipDoc.Value;
              zrecebimentosaldo.Value:=zqrecebimentosaldo.Value;
              zrecebimentomarcar.Value:=zqrecebimentomarcar.Value;
              zrecebimentoRefBaixa.Value:=zqrecebimentoRefBaixa.Value;
              zrecebimentorefvinda.Value:=zqrecebimentorefvinda.Value;
              zrecebimentocontabil.Value:=zqrecebimentocontabil.Value;
              zrecebimentoempresa.Value:=zqrecebimentoempresa.Value;
              zrecebimentocustodaparcela.Value:=zqrecebimentocustodaparcela.Value;
              zrecebimentoorigem.Value:=zqrecebimentoorigem.Value;
              zrecebimentoadversa.Value:=zqrecebimentoadversa.Value;
              zrecebimentorecpag.Value:=zqrecebimentorecpag.Value;
              zrecebimentonumordem.Value:=zqrecebimentonumordem.Value;
              zrecebimentoidloteamento.Value:=zqrecebimentoidloteamento.Value;
              zrecebimentovenda_idvenda.Value:=zqrecebimentovenda_idvenda.Value;
              zrecebimentoquadralote.Value:=zqrecebimentoquadralote.Value;
              zrecebimentonumboleto.Value:=zqrecebimentonumboleto.Value;
              zrecebimentoSubstituicao.Value:=zqrecebimentoSubstituicao.Value;
              zrecebimentosq.Value:=zqrecebimentosq.Value;
              zrecebimentosomar.Value:=zqrecebimentosomar.Value;
              zrecebimentonomeadversa.Value:=zqrecebimentonomeadversa.Value;
              zrecebimentoReajustado.Value:=zqrecebimentoReajustado.Value;
              zrecebimentoData_reajuste.Value:=zqrecebimentoData_reajuste.Value;
              zrecebimentoProximo_Reajuste.Value:=zqrecebimentoProximo_Reajuste.Value;
              zrecebimentoParcelas_fixas.Value:=zqrecebimentoParcelas_fixas.Value;
              zrecebimentojuros.Value:=zqrecebimentojuros.Value;
              zrecebimentodescontos.Value:=zqrecebimentodescontos.Value;
              zrecebimentoData_Quitacao.Value:=zqrecebimentoData_Quitacao.Value;
              zrecebimento.post;
            end;
            zqrecebimento.next;
          end;
        end; }


        ZQQuadras.First;
        if ZQQuadras.RecordCount>0 then
        begin
          msg.Caption:='Quadras                                                   ';
          application.ProcessMessages;

          while not ZQQuadras.Eof do
          begin
           if not ZQuadras.Locate('loteamento_idLoteamento;descricao',VarArrayOf([ZQquadras.FieldByName('loteamento_idLoteamento').Value,ZQquadras.FieldByName('descricao').Value]),[]) then
           begin
             ZQuadras.Insert;
//           ZQuadrasidquadras.Value:=ZQQuadrasidquadras.Value;
             ZQuadras.FieldByName('loteamento_idLoteamento').Value:=ZQquadras.FieldByName('loteamento_idLoteamento').Value;
             ZQuadras.FieldByName('descricao').Value:=ZQquadras.FieldByName('descricao').Value;
             ZQuadras.FieldByName('apedlido').Value:=ZQquadras.FieldByName('apedlido').Value;
             ZQuadras.Post;
           end;
           ZQQuadras.Next;
         end;
       end;
       if not zimovel.Locate('loteamento_idLoteamento;quadra;lote',VarArrayOf([ZQimovel.FieldByName('loteamento_idloteamento').Value,ZQimovel.FieldByName('quadra').Value,ZQimovel.FieldByName('lote').Value]),[]) then
       begin
         zimovel.Insert;
         //zimovelidimovel.Value:=zqimovelidimovel.Value;
         ZImovel.FieldByName('loteamento_idloteamento').Value:=ZQimovel.FieldByName('loteamento_idloteamento').Value;
         ZImovel.FieldByName('quadra').Value:=ZQimovel.FieldByName('quadra').Value;
         ZImovel.FieldByName('lote').Value:=ZQimovel.FieldByName('lote').Value;
         ZImovel.FieldByName('lado').Value:=ZQimovel.FieldByName('lado').Value;
         ZImovel.FieldByName('medidafrente').Value:=ZQimovel.FieldByName('medidafrente').Value;
         ZImovel.FieldByName('medidafundo').Value:=ZQimovel.FieldByName('medidafundo').Value;
         ZImovel.FieldByName('medidaesquerda').Value:=ZQimovel.FieldByName('medidaesquerda').Value;
         ZImovel.FieldByName('medidadedireita').Value:=ZQimovel.FieldByName('medidadedireita').Value;
         ZImovel.FieldByName('area').Value:=ZQimovel.FieldByName('area').Value;
         ZImovel.FieldByName('confrontofrente').Value:=ZQimovel.FieldByName('confrontofrente').Value;
         ZImovel.FieldByName('confrontofundo').Value:=ZQimovel.FieldByName('confrontofundo').Value;
         ZImovel.FieldByName('confrontodireita').Value:=ZQimovel.FieldByName('confrontodireita').Value;
         ZImovel.FieldByName('confrontoesquerda').Value:=ZQimovel.FieldByName('confrontoesquerda').Value;
         ZImovel.FieldByName('esquinanome').Value:=ZQimovel.FieldByName('esquinanome').Value;
         ZImovel.FieldByName('esquinalinha').Value:=ZQimovel.FieldByName('esquinalinha').Value;
         ZImovel.FieldByName('esquinacurva').Value:=ZQimovel.FieldByName('esquinacurva').Value;
         ZImovel.FieldByName('valorvenal').Value:=ZQimovel.FieldByName('valorvenal').Value;
         ZImovel.FieldByName('Obervacaoloteamento').Value:=ZQimovel.FieldByName('Obervacaoloteamento').Value;
         ZImovel.FieldByName('cornomapa').Value:=ZQimovel.FieldByName('cornomapa').Value;
         ZImovel.FieldByName('disponivel').Value:=ZQimovel.FieldByName('disponivel').Value;
         ZImovel.FieldByName('matri').Value:=ZQimovel.FieldByName('matri').Value;
         ZImovel.FieldByName('matricula').Value:=ZQimovel.FieldByName('matricula').Value;
         zimovel.post;
       end;
       zqimovel.next;
      end;
    end;
    ProgressBar2.Position:=0;
    ZQincorp_loteame.First;
    if ZQincorp_loteame.RecordCount>0 then
    begin
      msg.Caption:='Incorporadores/Responsaveis';
      application.ProcessMessages;

      while not ZQincorp_loteame.Eof do
      begin
        ZQrepresentantes.First;
        if ZQrepresentantes.RecordCount>0 then
        begin
          msg.Caption:='Representantes           ';
          application.ProcessMessages;

          while not ZQrepresentantes.eof do
          begin
            if not Zrepresentantes.Locate('incorporador_idincorporador;nomerepre;documento',VarArrayOf([ZQrepresentantes.FieldByName('incorporador_idincorporador').Value,ZQrepresentantes.FieldByName('nomerepre').value,ZQrepresentantes.FieldByName('documento').Value]),[]) then
            begin
              Zrepresentantes.insert;
              Zrepresentantes.FieldByName('incorporador_idincorporador').Value:=ZQrepresentantes.FieldByName('incorporador_idincorporador').Value;
              Zrepresentantes.FieldByName('nomerepre').value:=ZQrepresentantes.FieldByName('nomerepre').value;
              Zrepresentantes.FieldByName('documento').Value:=ZQrepresentantes.FieldByName('documento').Value;
              Zrepresentantes.post;
            end;
            ZQrepresentantes.Next;
          end;
        end;
      {  ZqParticipante.First;
        if ZqParticipante.RecordCount>0 then
        begin
          msg.Caption:='Participantes                          ';
          application.ProcessMessages;

          while not ZqParticipante.Eof do
          begin
            ZQresponsavel.First;
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

              while not ZQCompr_conjuge.Eof do
              begin
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
            end;}
{            ZQDividido.First;
            if ZQDividido.RecordCount>0 then
            begin
              while not ZQDividido.Eof do
              begin
                ZQDividido.delete;
              end;
            end;}

          {  if ZQConta_bancaria.RecordCount>0 then
            begin
              msg.Caption:='Conta Banc�ria                                ';
              application.ProcessMessages;
              ZQConta_bancaria.First;
              while not ZQConta_bancaria.Eof do
              begin
                if not ZConta_bancaria.Locate('idconta_bancaria;n_conta',VarArrayOf([ZQConta_bancariaidconta_bancaria.value,ZQConta_bancarian_conta.Value]),[]) then
                begin
                  ZConta_bancaria.insert;
                  ZConta_bancariaidconta_bancaria.value:=ZQConta_bancariaidconta_bancaria.value;
                  ZConta_bancarian_banco.value:=ZQConta_bancarian_banco.value;
                  ZConta_bancarian_agencia.Value:=ZQConta_bancarian_agencia.Value;
                  ZConta_bancarian_agencia_v.Value:=ZQConta_bancarian_agencia_v.Value;
                  ZConta_bancarian_conta.Value:=ZQConta_bancarian_conta.Value;
                  ZConta_bancarian_conta_v.Value:=ZQConta_bancarian_conta_v.Value;
                  ZConta_bancarianomebanco.Value:=ZQConta_bancarianomebanco.Value;
                  ZConta_bancarian_no_banco.Value:=ZQConta_bancarian_no_banco.Value;
                  ZConta_bancarianomeresposavel.Value:=ZQConta_bancarianomeresposavel.Value;
                  ZConta_bancarianometitular.Value:=ZQConta_bancarianometitular.Value;
                  ZConta_bancariaaberta.Value:=ZQConta_bancariaaberta.Value;
                  ZConta_bancariasenha_interna.Value:=ZQConta_bancariasenha_interna.Value;
                  ZConta_bancariasenha_conta.Value:=ZQConta_bancariasenha_conta.Value;
                  ZConta_bancariaativa.Value:=ZQConta_bancariaativa.Value;
                  ZConta_bancariajurosemboleto.Value:=ZQConta_bancariajurosemboleto.Value;
                  ZConta_bancariataxadiaria.Value:=ZQConta_bancariataxadiaria.Value;
                  ZConta_bancariavaloroupercent.Value:=ZQConta_bancariavaloroupercent.Value;
                  ZConta_bancariacarteira.Value:=ZQConta_bancariacarteira.Value;
                  ZConta_bancariaboleto.Value:=ZQConta_bancariaboleto.Value;
                  ZConta_bancariaidparticipante.Value:=ZQConta_bancariaidparticipante.Value;
                  ZConta_bancarian_dif_empreed.Value:=ZQConta_bancarian_dif_empreed.Value;
                  ZConta_bancariaconvenio.Value:=ZQConta_bancariaconvenio.Value;
                  ZConta_bancariadoc_titular.Value:=ZQConta_bancariadoc_titular.Value;
                  ZConta_bancaria.post;
                end;
                ZQConta_bancaria.next;
              end;
            end;
            if not ZParticipante.Locate('idpaticipante;doc1',VarArrayOf([ZqParticipanteidpaticipante.Value,ZqParticipantedoc1.Value]),[]) then
            begin
              ZParticipante.insert;
              ZParticipanteidpaticipante.Value:=ZqParticipanteidpaticipante.Value;
              ZParticipantenome_parte.Value:=ZqParticipantenome_parte.Value;
              ZParticipantedoc1.Value:=ZqParticipantedoc1.Value;
              ZParticipantedoc2.Value:=ZqParticipantedoc2.Value;
              ZParticipanteendereco.Value:=ZqParticipanteendereco.Value;
              ZParticipantebairro.Value:=ZqParticipantebairro.Value;
              ZParticipantecidade.Value:=ZqParticipantecidade.Value;
              ZParticipantecep.Value:=ZqParticipantecep.Value;
              ZParticipanteende_cob.Value:=ZqParticipanteende_cob.Value;
              ZParticipantebairro_cob.Value:=ZqParticipantebairro_cob.Value;
              ZParticipantecidade_cob.Value:=ZqParticipantecidade_cob.Value;
              ZParticipantecep_cob.Value:=ZqParticipantecep_cob.Value;
              ZParticipanteFone1.Value:=ZqParticipanteFone1.Value;
              ZParticipantefone2.Value:=ZqParticipantefone2.Value;
              ZParticipantefone3.Value:=ZqParticipantefone3.Value;
              ZParticipantetipopessoa.Value:=ZqParticipantetipopessoa.Value;
              ZParticipanteaniversario.Value:=ZqParticipanteaniversario.Value;
              ZParticipantenacionalidade.Value:=ZqParticipantenacionalidade.Value;
              ZParticipanteemail.Value:=ZqParticipanteemail.Value;
              ZParticipantenaturalidade.Value:=ZqParticipantenaturalidade.Value;
              ZParticipantelocaldetrab.Value:=ZqParticipantelocaldetrab.Value;
              ZParticipanteprofissao.Value:=ZqParticipanteprofissao.Value;
              ZParticipanteestadocivil.Value:=ZqParticipanteestadocivil.Value;
              ZParticipanterenda.Value:=ZqParticipanterenda.Value;
              ZParticipanteobservacao.Value:=ZqParticipanteobservacao.Value;
              ZParticipantecodrecinc.Value:=ZqParticipantecodrecinc.Value;
              ZParticipantecodrecBx.Value:=ZqParticipantecodrecBx.Value;
              ZParticipantecodpaginc.Value:=ZqParticipantecodpaginc.Value;
              ZParticipantecodpagBx.Value:=ZqParticipantecodpagBx.Value;
              ZParticipantefone4.Value:=ZqParticipantefone4.Value;
              ZParticipantecomplemento.Value:=ZqParticipantecomplemento.Value;
              ZParticipantecomplemento_cob.Value:=ZqParticipantecomplemento_cob.Value;
              ZParticipantecadastrado.Value:=ZqParticipantecadastrado.Value;
              ZParticipanteNome_Firma.Value:=ZqParticipanteNome_Firma.Value;
              ZParticipanteInativo.value:=ZqParticipanteInativo.value;
              ZParticipantesexo.value:=ZqParticipantesexo.value;
              ZParticipantecc.value:=ZqParticipantecc.value;
              ZParticipantecex.value:=ZqParticipantecex.value;
              ZParticipantejb.value:=ZqParticipantejb.value;
              ZParticipantejudicial.value:=ZqParticipantejudicial.value;
              ZParticipantexquadra.value:=ZqParticipantexquadra.value;
              ZParticipantexlote.value:=ZqParticipantexlote.value;
              ZParticipante.post;
            end;
            ZqParticipante.next;
          end;
        end;}
        if not Zincorp_loteame.Locate('loteamento_idloteamento',VarArrayOf([ZQincorp_loteame.FieldByName('loteamento_idloteamento').Value]),[]) then
        begin
          Zincorp_loteame.insert;
          Zincorp_loteame.FieldByName('incorporador_idincorporador').Value:=ZQincorp_loteame.FieldByName('incorporador_idincorporador').Value;
          Zincorp_loteame.FieldByName('loteamento_idloteamento').Value:=ZQincorp_loteame.FieldByName('loteamento_idloteamento').Value;
          Zincorp_loteame.FieldByName('percent_parte').Value:=ZQincorp_loteame.FieldByName('percent_parte').Value;
          Zincorp_loteame.FieldByName('codcontabancaria').Value:=ZQincorp_loteame.FieldByName('codcontabancaria').Value;
          Zincorp_loteame.FieldByName('digito_dif').Value:=ZQincorp_loteame.FieldByName('digito_dif').Value;
          Zincorp_loteame.post;
        end;
        ZQincorp_loteame.Next;

      end;

    end;

    ZQProcuradores.First;
    if ZQProcuradores.RecordCount>0 then
    begin
      msg.Caption:='Procuradores                                                ';
      application.ProcessMessages;

      while not ZQProcuradores.Eof do
      begin
        if not ZProcuradores.Locate('idloteamento;nome',VarArrayOf([ZQProcuradores.FieldByName('idloteamento').Value,ZQProcuradores.FieldByName('nome').Value]),[]) then
        begin
          ZProcuradores.Insert;
          ZProcuradores.FieldByName('idloteamento').Value:=ZQProcuradores.FieldByName('idloteamento').Value;
          ZProcuradores.FieldByName('idparti').Value:=ZQProcuradores.FieldByName('idparti').Value;
          ZProcuradores.FieldByName('nome').Value:=ZQProcuradores.FieldByName('nome').Value;
          ZProcuradores.FieldByName('assina').Value:=ZQProcuradores.FieldByName('assina').Value;
          ZProcuradores.post;
        end;
        ZQProcuradores.Next;
      end;
    end;

    if not ZLoteamento.Locate('idloteamento',VarArrayOf([ZQLoteamento.FieldByName('idloteamento').Value]),[]) then
    begin
      ZLoteamento.Insert;
      Zloteamento.FieldByName('idloteamento').Value:=ZQLoteamento.FieldByName('idloteamento').Value;
      Zloteamento.FieldByName('cidade_idcidade').Value:=ZQLoteamento.FieldByName('cidade_idcidade').Value;
      Zloteamento.FieldByName('matriculaloteamento').Value:=ZQLoteamento.FieldByName('matriculaloteamento').Value;
      Zloteamento.FieldByName('datacadastro').Value:=ZQLoteamento.FieldByName('datacadastro').Value;
      Zloteamento.FieldByName('datainalguracao').Value:=ZQLoteamento.FieldByName('datainalguracao').Value;
      Zloteamento.FieldByName('dataconclusao').Value:=ZQLoteamento.FieldByName('dataconclusao').Value;
      Zloteamento.FieldByName('logotipo').Value:=ZQLoteamento.FieldByName('logotipo').Value;
      Zloteamento.FieldByName('mapa').Value:=ZQLoteamento.FieldByName('mapa').Value;
      Zloteamento.FieldByName('testemunha1').Value:=ZQLoteamento.FieldByName('testemunha1').Value;
      Zloteamento.FieldByName('testemunha2').Value:=ZQLoteamento.FieldByName('testemunha2').Value;
      Zloteamento.FieldByName('testemunha3').Value:=ZQLoteamento.FieldByName('testemunha3').Value;
      Zloteamento.FieldByName('testemunha4').Value:=ZQLoteamento.FieldByName('testemunha4').Value;
      Zloteamento.FieldByName('Observincorp').Value:=ZQLoteamento.FieldByName('Observincorp').Value;
      Zloteamento.FieldByName('bairro').Value:=ZQLoteamento.FieldByName('bairro').Value;
      Zloteamento.FieldByName('pasta_mapa').Value:=ZQLoteamento.FieldByName('pasta_mapa').Value;
      Zloteamento.FieldByName('nomeloteamento').Value:=ZQLoteamento.FieldByName('nomeloteamento').Value;
      Zloteamento.FieldByName('tipo').Value:=ZQLoteamento.FieldByName('tipo').Value;
      Zloteamento.FieldByName('codcontabancaria').Value:=ZQLoteamento.FieldByName('codcontabancaria').Value;
      Zloteamento.FieldByName('apelido').Value:=ZQLoteamento.FieldByName('apelido').Value;
      Zloteamento.FieldByName('digito_dif').Value:=ZQLoteamento.FieldByName('digito_dif').Value;
      ZLoteamento.post;
    end;
    ZQLoteamento.next;
  end;
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

procedure Tfrm_baixa.xloteEnter(Sender: TObject);
begin
  Frm_Achaloteamento.Top := frm_baixa.Top+105;
  Frm_Achaloteamento.Left := frm_baixa.Left+(frm_baixa.Width-Frm_Achaloteamento.Width)  ;
  Frm_Achaloteamento.showmodal;
  xlote.Text:=DM_Tabelas.ZQAchaLotea.FieldByName('idloteamento').text;
  nome.Text:=DM_Tabelas.ZQAchaLotea.FieldByName('apelido').Value;
  DM_Tabelas.ZQAchaLotea.Close;
end;

procedure Tfrm_baixa.Button2Click(Sender: TObject);
begin
  close;
end;

procedure Tfrm_baixa.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  ZQConta_bancaria.Close;
  zqimovel.Close;
  zqrecebimento.Close;
  ZQReBxHi.Close;
  ZQReceb_Baixa.Close;
  ZQRemes_Receb.Close;
  ZQrecebimento_historico_cessao.Close;
  ZQrecebimento_historico_excluidos.Close;
  ZQVenda.Close;
  ZQVendedor.Close;
  ZQCheque.Close;
  ZQAditamento.Close;
  ZQhistorico_cessao.Close;
  ZQQuitacao.Close;
  ZQComprador.Close;
  ZQComprador_cessao.Close;
  ZQincorp_loteame.Close;
  ZQParcelas.Close;
  ZQQuadras.Close;
  ZQProcuradores.Close;
  ZQrepresentantes.Close;
  ZqParticipante.Close;
  ZQresponsavel.Close;
  ZQCompr_conjuge.Close;
  ZQLoteamento.close;
  ZQDividido.close;
  ZQConjuge.close;
  ZQLoteamento.close;

  ZConta_bancaria.Close;
  zimovel.Close;
  zrecebimento.Close;
  ZReBxHi.Close;
  ZReceb_Baixa.Close;
  ZRemes_Receb.Close;
  Zrecebimento_historico_cessao.Close;
  Zrecebimento_historico_excluidos.Close;
  ZVenda.Close;
  ZVendedor.Close;
  ZCheque.Close;
  ZAditamento.Close;
  Zhistorico_cessao.Close;
  ZQuitacao.Close;
  ZComprador.Close;
  ZComprador_cessao.Close;
  Zincorp_loteame.Close;
  ZParcelas.Close;
  ZQuadras.Close;
  ZProcuradores.Close;
  Zrepresentantes.Close;
  ZParticipante.Close;
  Zresponsavel.Close;
  ZCompr_conjuge.Close;
  ZLoteamento.close;
  ZConjuge.close;

end;

procedure Tfrm_baixa.FormShow(Sender: TObject);
begin
  ZQConta_bancaria.Open;
  zqimovel.Open;
  zqrecebimento.open;
  ZQReBxHi.Open;
  ZQReceb_Baixa.open;
  ZQRemes_Receb.open;
  ZQrecebimento_historico_cessao.open;
  ZQrecebimento_historico_excluidos.open;
  ZQVenda.open;
  ZQVendedor.open;
  ZQCheque.open;
  ZQAditamento.open;
  ZQhistorico_cessao.open;
  ZQQuitacao.open;
  ZQComprador.open;
  ZQComprador_cessao.open;
  ZQincorp_loteame.open;
  ZQParcelas.open;
  ZQQuadras.open;
  ZQProcuradores.open;
  ZQrepresentantes.open;
  ZqParticipante.open;
  ZQresponsavel.open;
  ZQCompr_conjuge.open;
  ZQDividido.open;
  ZQConjuge.open;
  ZQLoteamento.open;

  ZConta_bancaria.Open;
  zimovel.Open;
  zrecebimento.open;
  ZReBxHi.Open;
  ZReceb_Baixa.open;
  ZRemes_Receb.open;
  Zrecebimento_historico_cessao.open;
  Zrecebimento_historico_excluidos.open;
  ZVenda.open;
  ZVendedor.open;
  ZCheque.open;
  ZAditamento.open;
  Zhistorico_cessao.open;
  ZQuitacao.open;
  ZComprador.open;
  ZComprador_cessao.open;
  Zincorp_loteame.open;
  ZParcelas.open;
  ZQuadras.open;
  ZProcuradores.open;
  Zrepresentantes.open;
  ZParticipante.open;
  Zresponsavel.open;
  ZCompr_conjuge.open;
  ZConjuge.open;
  ZLoteamento.open;

  xlote.Text:='4';
  xlote.SetFocus;
end;

procedure Tfrm_baixa.xloteExit(Sender: TObject);
begin
  Button1.Enabled:=true;
  Button1.SetFocus;

  ZQLoteamento.close;
  ZQLoteamento.SQL.Clear;
  ZQLoteamento.SQL.add('Select * from Loteamento where idloteamento='+quotedstr(xlote.text));
  ZQLoteamento.open;
  ZQLoteamento.first;


end;


procedure Tfrm_baixa.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQLoteamento', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQLoteamento', 'ZQLoteamentoidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQLoteamento', 'ZQLoteamentocidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQLoteamento', 'ZQLoteamentomatriculaloteamento', 'matriculaloteamento', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQLoteamento', 'ZQLoteamentodatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQLoteamento', 'ZQLoteamentodatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQLoteamento', 'ZQLoteamentodataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQLoteamento', 'ZQLoteamentologotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQLoteamento', 'ZQLoteamentomapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQLoteamento', 'ZQLoteamentotestemunha1', 'testemunha1', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQLoteamento', 'ZQLoteamentotestemunha2', 'testemunha2', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQLoteamento', 'ZQLoteamentotestemunha3', 'testemunha3', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQLoteamento', 'ZQLoteamentotestemunha4', 'testemunha4', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQLoteamento', 'ZQLoteamentoObservincorp', 'Observincorp', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQLoteamento', 'ZQLoteamentobairro', 'bairro', TStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQLoteamento', 'ZQLoteamentopasta_mapa', 'pasta_mapa', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQLoteamento', 'ZQLoteamentonomeloteamento', 'nomeloteamento', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQLoteamento', 'ZQLoteamentotipo', 'tipo', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQLoteamento', 'ZQLoteamentocodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQLoteamento', 'ZQLoteamentoapelido', 'apelido', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQLoteamento', 'ZQLoteamentodigito_dif', 'digito_dif', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
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
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQincorp_loteame', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQincorp_loteame', 'ZQincorp_loteameincorporador_idincorporador', 'incorporador_idincorporador', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQincorp_loteame', 'ZQincorp_loteameloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQincorp_loteame', 'ZQincorp_loteamepercent_parte', 'percent_parte', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQincorp_loteame', 'ZQincorp_loteamecodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQincorp_loteame', 'ZQincorp_loteamedigito_dif', 'digito_dif', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQincorp_loteame', 'ZQincorp_loteamenconta', 'nconta', TStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'codcontabancaria', 'ZQConta_bancaria', 'idconta_bancaria', 'n_conta', True);
  RegisterRuntimeField(Tfrm_baixa, 'ZQincorp_loteame', 'ZQincorp_loteameconta_v', 'conta_v', TStringField, fkLookup, 3, 0, False, '', '', '', '', 0, 'codcontabancaria', 'ZQConta_bancaria', 'idconta_bancaria', 'n_conta_v', True);
  RegisterRuntimeField(Tfrm_baixa, 'ZQincorp_loteame', 'ZQincorp_loteameagencia', 'agencia', TStringField, fkLookup, 6, 0, False, '', '', '', '', 0, 'codcontabancaria', 'ZQConta_bancaria', 'idconta_bancaria', 'n_agencia', True);
  RegisterRuntimeField(Tfrm_baixa, 'ZQincorp_loteame', 'ZQincorp_loteameagencia_v', 'agencia_v', TStringField, fkLookup, 3, 0, False, '', '', '', '', 0, 'codcontabancaria', 'ZQConta_bancaria', 'idconta_bancaria', 'n_agencia_v', True);
  RegisterRuntimeField(Tfrm_baixa, 'ZQincorp_loteame', 'ZQincorp_loteamen_banco', 'n_banco', TStringField, fkLookup, 3, 0, False, '', '', '', '', 0, 'codcontabancaria', 'ZQConta_bancaria', 'idconta_bancaria', 'n_banco', True);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQQuadras', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQQuadras', 'ZQQuadrasidquadras', 'idquadras', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQQuadras', 'ZQQuadrasloteamento_idLoteamento', 'loteamento_idLoteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQQuadras', 'ZQQuadrasdescricao', 'descricao', TStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQQuadras', 'ZQQuadrasapedlido', 'apedlido', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
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
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQReceb_Baixa', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReceb_Baixa', 'ZQReceb_Baixaidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReceb_Baixa', 'ZQReceb_BaixaDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReceb_Baixa', 'ZQReceb_BaixaVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReceb_Baixa', 'ZQReceb_BaixaTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReceb_Baixa', 'ZQReceb_BaixaDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReceb_Baixa', 'ZQReceb_BaixaJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReceb_Baixa', 'ZQReceb_BaixaDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReceb_Baixa', 'ZQReceb_BaixaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReceb_Baixa', 'ZQReceb_Baixaobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReceb_Baixa', 'ZQReceb_BaixaCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReceb_Baixa', 'ZQReceb_Baixadataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReceb_Baixa', 'ZQReceb_Baixavencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReceb_Baixa', 'ZQReceb_Baixasubstituicao', 'substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReceb_Baixa', 'ZQReceb_Baixavr_ab', 'vr_ab', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReceb_Baixa', 'ZQReceb_Baixasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReceb_Baixa', 'ZQReceb_Baixavalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQVendedor', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQVendedor', 'ZQVendedoridvendedor', 'idvendedor', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQVendedor', 'ZQVendedorcorretor_idcorretor', 'corretor_idcorretor', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQVendedor', 'ZQVendedorvenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQRecebimento', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentodocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentoObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentoordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentoTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentomarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentoorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentorecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentoquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentonumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentoSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentosomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentonomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentoReajustado', 'Reajustado', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentoProximo_Reajuste', 'Proximo_Reajuste', TStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentoParcelas_fixas', 'Parcelas_fixas', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentodescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQRecebimento', 'ZQRecebimentoData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQVenda', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQVenda', 'ZQVendaidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQVenda', 'ZQVendadatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQVenda', 'ZQVendaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQVenda', 'ZQVendavalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQVenda', 'ZQVendaforma_reajuste', 'forma_reajuste', TStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQVenda', 'ZQVendatabela_Price', 'tabela_Price', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQVenda', 'ZQVendaEscriturado', 'Escriturado', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQReBxHi', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReBxHi', 'ZQReBxHiidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReBxHi', 'ZQReBxHirefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReBxHi', 'ZQReBxHiidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReBxHi', 'ZQReBxHivalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReBxHi', 'ZQReBxHidescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReBxHi', 'ZQReBxHijuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReBxHi', 'ZQReBxHipercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReBxHi', 'ZQReBxHidata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReBxHi', 'ZQReBxHisq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQReBxHi', 'ZQReBxHivalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
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
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQrecebimento_historico_cessao', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoSequencial', 'Sequencial', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaodocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaomarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaorecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaonumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaosomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaonomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoReajustado', 'Reajustado', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoProximo_Reajuste', 'Proximo_Reajuste', TStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoParcelas_fixas', 'Parcelas_fixas', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaodata_cessao', 'data_cessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaohora_cessao', 'hora_cessao', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidoscliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidossaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosmarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidoscontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidoscustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosrecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosnumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidossq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidossomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosmotivo', 'motivo', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosProximo_Reajuste', 'Proximo_Reajuste', TStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidossequencial', 'sequencial', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
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
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQCheque', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCheque', 'ZQChequeBanco', 'Banco', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCheque', 'ZQChequeDono', 'Dono', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCheque', 'ZQChequeCPF_CNPJ', 'CPF_CNPJ', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCheque', 'ZQChequevalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCheque', 'ZQChequeemissao', 'emissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCheque', 'ZQChequedeposito', 'deposito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCheque', 'ZQChequealias', 'alias', TStringField, fkData, 12, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCheque', 'ZQChequeprorrogado', 'prorrogado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCheque', 'ZQChequeconta', 'conta', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCheque', 'ZQChequeidrecebimento', 'idrecebimento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCheque', 'ZQChequeidpagamento', 'idpagamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCheque', 'ZQChequeidcheque', 'idcheque', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCheque', 'ZQChequeidvinculo', 'idvinculo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCheque', 'ZQChequenumero', 'numero', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCheque', 'ZQChequeagencia', 'agencia', TStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCheque', 'ZQChequeSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQCheque', 'ZQChequesq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQAditamento', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQAditamento', 'ZQAditamentoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQAditamento', 'ZQAditamentodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQAditamento', 'ZQAditamentodatacessao', 'datacessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQAditamento', 'ZQAditamentoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQAditamento', 'ZQAditamentovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQAditamento', 'ZQAditamentoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQAditamento', 'ZQAditamentoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQAditamento', 'ZQAditamentonome_parte', 'nome_parte', TStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQAditamento', 'ZQAditamentohistorico', 'historico', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQAditamento', 'ZQAditamentohora', 'hora', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQAditamento', 'ZQAditamentocontrole', 'controle', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQquitacao', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQquitacao', 'ZQquitacaoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQquitacao', 'ZQquitacaodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQquitacao', 'ZQquitacaodatacessao', 'datacessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQquitacao', 'ZQquitacaoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQquitacao', 'ZQquitacaovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQquitacao', 'ZQquitacaoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQquitacao', 'ZQquitacaoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQquitacao', 'ZQquitacaonome_parte', 'nome_parte', TStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQquitacao', 'ZQquitacaohistorico', 'historico', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQquitacao', 'ZQquitacaohora', 'hora', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQquitacao', 'ZQquitacaocontrole', 'controle', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQhistorico_cessao', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQhistorico_cessao', 'ZQhistorico_cessaoidhistorico_cessao', 'idhistorico_cessao', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQhistorico_cessao', 'ZQhistorico_cessaoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQhistorico_cessao', 'ZQhistorico_cessaohistorico', 'historico', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQComprador_cessao', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQComprador_cessao', 'ZQComprador_cessaoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQComprador_cessao', 'ZQComprador_cessaodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQComprador_cessao', 'ZQComprador_cessaodatacessao', 'datacessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQComprador_cessao', 'ZQComprador_cessaoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQComprador_cessao', 'ZQComprador_cessaovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQComprador_cessao', 'ZQComprador_cessaoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQComprador_cessao', 'ZQComprador_cessaoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQComprador_cessao', 'ZQComprador_cessaonome_parte', 'nome_parte', TStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQComprador_cessao', 'ZQComprador_cessaohistorico', 'historico', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQComprador_cessao', 'ZQComprador_cessaohora_cessao', 'hora_cessao', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQComprador_cessao', 'ZQComprador_cessaoautocodigo', 'autocodigo', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQParcelas', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQParcelas', 'ZQParcelasidparcelas', 'idparcelas', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQParcelas', 'ZQParcelasvenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQParcelas', 'ZQParcelasordem', 'ordem', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQParcelas', 'ZQParcelasdatavenci', 'datavenci', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQParcelas', 'ZQParcelasvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQParcelas', 'ZQParcelassaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQParcelas', 'ZQParcelasvalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQParcelas', 'ZQParcelasnobanco', 'nobanco', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQrepresentantes', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrepresentantes', 'ZQrepresentantesidrepresentantes', 'idrepresentantes', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrepresentantes', 'ZQrepresentantesincorporador_idincorporador', 'incorporador_idincorporador', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrepresentantes', 'ZQrepresentantesnomerepre', 'nomerepre', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQrepresentantes', 'ZQrepresentantesdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQConta_bancaria', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConta_bancaria', 'ZQConta_bancariaidconta_bancaria', 'idconta_bancaria', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConta_bancaria', 'ZQConta_bancarian_banco', 'n_banco', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConta_bancaria', 'ZQConta_bancarian_agencia', 'n_agencia', TStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConta_bancaria', 'ZQConta_bancarian_agencia_v', 'n_agencia_v', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConta_bancaria', 'ZQConta_bancarian_conta', 'n_conta', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConta_bancaria', 'ZQConta_bancarian_conta_v', 'n_conta_v', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConta_bancaria', 'ZQConta_bancarianomebanco', 'nomebanco', TStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConta_bancaria', 'ZQConta_bancarian_no_banco', 'n_no_banco', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConta_bancaria', 'ZQConta_bancarianomeresposavel', 'nomeresposavel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConta_bancaria', 'ZQConta_bancarianometitular', 'nometitular', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConta_bancaria', 'ZQConta_bancariaaberta', 'aberta', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConta_bancaria', 'ZQConta_bancariasenha_interna', 'senha_interna', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConta_bancaria', 'ZQConta_bancariasenha_conta', 'senha_conta', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConta_bancaria', 'ZQConta_bancariaativa', 'ativa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConta_bancaria', 'ZQConta_bancariajurosemboleto', 'jurosemboleto', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConta_bancaria', 'ZQConta_bancariataxadiaria', 'taxadiaria', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConta_bancaria', 'ZQConta_bancariavaloroupercent', 'valoroupercent', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConta_bancaria', 'ZQConta_bancariacarteira', 'carteira', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConta_bancaria', 'ZQConta_bancariaboleto', 'boleto', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConta_bancaria', 'ZQConta_bancariaidparticipante', 'idparticipante', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConta_bancaria', 'ZQConta_bancarian_dif_empreed', 'n_dif_empreed', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConta_bancaria', 'ZQConta_bancariaconvenio', 'convenio', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConta_bancaria', 'ZQConta_bancariadoc_titular', 'doc_titular', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'Zloteamento', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zloteamento', 'Zloteamentoidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zloteamento', 'Zloteamentocidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zloteamento', 'Zloteamentomatriculaloteamento', 'matriculaloteamento', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zloteamento', 'Zloteamentodatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zloteamento', 'Zloteamentodatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zloteamento', 'Zloteamentodataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zloteamento', 'Zloteamentologotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zloteamento', 'Zloteamentomapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zloteamento', 'Zloteamentotestemunha1', 'testemunha1', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zloteamento', 'Zloteamentotestemunha2', 'testemunha2', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zloteamento', 'Zloteamentotestemunha3', 'testemunha3', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zloteamento', 'Zloteamentotestemunha4', 'testemunha4', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zloteamento', 'ZloteamentoObservincorp', 'Observincorp', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zloteamento', 'Zloteamentobairro', 'bairro', TStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zloteamento', 'Zloteamentopasta_mapa', 'pasta_mapa', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zloteamento', 'Zloteamentonomeloteamento', 'nomeloteamento', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zloteamento', 'Zloteamentotipo', 'tipo', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zloteamento', 'Zloteamentocodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zloteamento', 'Zloteamentoapelido', 'apelido', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zloteamento', 'Zloteamentodigito_dif', 'digito_dif', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
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
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZImovel', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZImovel', 'ZImovelidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZImovel', 'ZImovelloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZImovel', 'ZImovelquadra', 'quadra', TStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZImovel', 'ZImovellote', 'lote', TStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZImovel', 'ZImovellado', 'lado', TStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZImovel', 'ZImovelmedidafrente', 'medidafrente', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZImovel', 'ZImovelmedidafundo', 'medidafundo', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZImovel', 'ZImovelmedidaesquerda', 'medidaesquerda', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZImovel', 'ZImovelmedidadedireita', 'medidadedireita', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZImovel', 'ZImovelarea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZImovel', 'ZImovelconfrontofrente', 'confrontofrente', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZImovel', 'ZImovelconfrontofundo', 'confrontofundo', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZImovel', 'ZImovelconfrontodireita', 'confrontodireita', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZImovel', 'ZImovelconfrontoesquerda', 'confrontoesquerda', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZImovel', 'ZImovelesquinanome', 'esquinanome', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZImovel', 'ZImovelesquinalinha', 'esquinalinha', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZImovel', 'ZImovelesquinacurva', 'esquinacurva', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZImovel', 'ZImovelvalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZImovel', 'ZImovelObervacaoloteamento', 'Obervacaoloteamento', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZImovel', 'ZImovelcornomapa', 'cornomapa', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZImovel', 'ZImoveldisponivel', 'disponivel', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZImovel', 'ZImovelmatri', 'matri', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZImovel', 'ZImovelmatricula', 'matricula', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
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
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZComprador', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZComprador', 'ZCompradoridcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZComprador', 'ZCompradorpaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZComprador', 'ZCompradorvenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZComprador', 'ZCompradorpromissario', 'promissario', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZComprador', 'ZCompradorpercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZComprador', 'ZCompradorMarcar', 'Marcar', TSmallintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZRecebimento', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentodocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentoObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentoordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentoTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentomarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentoorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentorecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentoquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentonumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentoSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentosomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentonomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentoReajustado', 'Reajustado', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentoProximo_Reajuste', 'Proximo_Reajuste', TStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentoParcelas_fixas', 'Parcelas_fixas', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentodescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRecebimento', 'ZRecebimentoData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZReceb_baixa', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZReceb_baixa', 'ZReceb_baixaidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZReceb_baixa', 'ZReceb_baixaDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZReceb_baixa', 'ZReceb_baixaVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZReceb_baixa', 'ZReceb_baixaTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZReceb_baixa', 'ZReceb_baixaDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZReceb_baixa', 'ZReceb_baixaJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZReceb_baixa', 'ZReceb_baixaDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZReceb_baixa', 'ZReceb_baixaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZReceb_baixa', 'ZReceb_baixaobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZReceb_baixa', 'ZReceb_baixaCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZReceb_baixa', 'ZReceb_baixadataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZReceb_baixa', 'ZReceb_baixavencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZReceb_baixa', 'ZReceb_baixaSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZReceb_baixa', 'ZReceb_baixasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZReceb_baixa', 'ZReceb_baixaVR_AB', 'VR_AB', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZReceb_baixa', 'ZReceb_baixavalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZReceb_baixa', 'ZReceb_baixaidreceb', 'idreceb', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'Zcontabancaria', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcontabancaria', 'Zcontabancariaidconta_bancaria', 'idconta_bancaria', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcontabancaria', 'Zcontabancarian_banco', 'n_banco', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcontabancaria', 'Zcontabancarian_agencia', 'n_agencia', TStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcontabancaria', 'Zcontabancarian_agencia_v', 'n_agencia_v', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcontabancaria', 'Zcontabancarian_conta', 'n_conta', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcontabancaria', 'Zcontabancarian_conta_v', 'n_conta_v', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcontabancaria', 'Zcontabancarianomebanco', 'nomebanco', TStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcontabancaria', 'Zcontabancarian_no_banco', 'n_no_banco', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcontabancaria', 'Zcontabancarianomeresposavel', 'nomeresposavel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcontabancaria', 'Zcontabancarianometitular', 'nometitular', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcontabancaria', 'Zcontabancariaaberta', 'aberta', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcontabancaria', 'Zcontabancariasenha_interna', 'senha_interna', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcontabancaria', 'Zcontabancariasenha_conta', 'senha_conta', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcontabancaria', 'Zcontabancariaativa', 'ativa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcontabancaria', 'Zcontabancariajurosemboleto', 'jurosemboleto', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcontabancaria', 'Zcontabancariataxadiaria', 'taxadiaria', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcontabancaria', 'Zcontabancariavaloroupercent', 'valoroupercent', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcontabancaria', 'Zcontabancariacarteira', 'carteira', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcontabancaria', 'Zcontabancariaboleto', 'boleto', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcontabancaria', 'Zcontabancariaidparticipante', 'idparticipante', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcontabancaria', 'Zcontabancarian_dif_empreed', 'n_dif_empreed', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcontabancaria', 'Zcontabancariaconvenio', 'convenio', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcontabancaria', 'Zcontabancariadoc_titular', 'doc_titular', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZVenda', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZVenda', 'ZVendaidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZVenda', 'ZVendadatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZVenda', 'ZVendaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZVenda', 'ZVendavalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZVenda', 'ZVendaforma_reajuste', 'forma_reajuste', TStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZVenda', 'ZVendatabela_Price', 'tabela_Price', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZVenda', 'ZVendaEscriturado', 'Escriturado', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZComprador_cessao', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZComprador_cessao', 'ZComprador_cessaoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZComprador_cessao', 'ZComprador_cessaodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZComprador_cessao', 'ZComprador_cessaodatacessao', 'datacessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZComprador_cessao', 'ZComprador_cessaoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZComprador_cessao', 'ZComprador_cessaovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZComprador_cessao', 'ZComprador_cessaoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZComprador_cessao', 'ZComprador_cessaoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZComprador_cessao', 'ZComprador_cessaonome_parte', 'nome_parte', TStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZComprador_cessao', 'ZComprador_cessaohistorico', 'historico', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZComprador_cessao', 'ZComprador_cessaohora_cessao', 'hora_cessao', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZComprador_cessao', 'ZComprador_cessaoautocodigo', 'autocodigo', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZRebxHi', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRebxHi', 'ZRebxHiidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRebxHi', 'ZRebxHirefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRebxHi', 'ZRebxHiidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRebxHi', 'ZRebxHivalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRebxHi', 'ZRebxHidescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRebxHi', 'ZRebxHijuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRebxHi', 'ZRebxHipercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRebxHi', 'ZRebxHidata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRebxHi', 'ZRebxHisq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRebxHi', 'ZRebxHivalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZAditamento', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZAditamento', 'ZAditamentoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZAditamento', 'ZAditamentodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZAditamento', 'ZAditamentodatacessao', 'datacessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZAditamento', 'ZAditamentoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZAditamento', 'ZAditamentovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZAditamento', 'ZAditamentoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZAditamento', 'ZAditamentoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZAditamento', 'ZAditamentonome_parte', 'nome_parte', TStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZAditamento', 'ZAditamentohistorico', 'historico', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZAditamento', 'ZAditamentohora', 'hora', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZAditamento', 'ZAditamentocontrole', 'controle', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQuadras', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQuadras', 'ZQuadrasidquadras', 'idquadras', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQuadras', 'ZQuadrasloteamento_idLoteamento', 'loteamento_idLoteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQuadras', 'ZQuadrasdescricao', 'descricao', TStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQuadras', 'ZQuadrasapedlido', 'apedlido', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
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
  RegisterRuntimeDataSet(Tfrm_baixa, 'Zrecebimento_historico_cessao', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaoSequencial', 'Sequencial', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaodocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaoObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaoordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaoTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaomarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaoorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaorecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaoquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaonumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaoSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaosomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaonomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaoReajustado', 'Reajustado', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaoProximo_Reajuste', 'Proximo_Reajuste', TStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaoParcelas_fixas', 'Parcelas_fixas', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaodata_cessao', 'data_cessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_cessao', 'Zrecebimento_historico_cessaohora_cessao', 'hora_cessao', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZProcuradores', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZProcuradores', 'ZProcuradoresidProcuradores', 'idProcuradores', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZProcuradores', 'ZProcuradoresidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZProcuradores', 'ZProcuradoresidparti', 'idparti', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZProcuradores', 'ZProcuradoresnome', 'nome', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZProcuradores', 'ZProcuradoresassina', 'assina', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZRemes_Receb', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRemes_Receb', 'ZRemes_Recebidremessa_receb', 'idremessa_receb', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRemes_Receb', 'ZRemes_Recebremessa', 'remessa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRemes_Receb', 'ZRemes_Recebidrec', 'idrec', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRemes_Receb', 'ZRemes_Recebretorno_motivo', 'retorno_motivo', TStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRemes_Receb', 'ZRemes_Recebacao', 'acao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRemes_Receb', 'ZRemes_RecebNossonumero', 'Nossonumero', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRemes_Receb', 'ZRemes_Recebocorrencia', 'ocorrencia', TStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRemes_Receb', 'ZRemes_Recebdesc_motivo', 'desc_motivo', TStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRemes_Receb', 'ZRemes_Recebcredito', 'credito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZRemes_Receb', 'ZRemes_Recebdt_ocorrencia', 'dt_ocorrencia', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'Zrecebimento_historico_excluidos', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidoscliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidossaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosmarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidoscontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidoscustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosrecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosnumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidossq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidossomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosmotivo', 'motivo', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidosProximo_Reajuste', 'Proximo_Reajuste', TStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrecebimento_historico_excluidos', 'Zrecebimento_historico_excluidossequencial', 'sequencial', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZQuitacao', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQuitacao', 'ZQuitacaoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQuitacao', 'ZQuitacaodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQuitacao', 'ZQuitacaodatacessao', 'datacessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQuitacao', 'ZQuitacaoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQuitacao', 'ZQuitacaovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQuitacao', 'ZQuitacaoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQuitacao', 'ZQuitacaoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQuitacao', 'ZQuitacaonome_parte', 'nome_parte', TStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQuitacao', 'ZQuitacaohistorico', 'historico', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQuitacao', 'ZQuitacaohora', 'hora', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQuitacao', 'ZQuitacaocontrole', 'controle', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'Zhistorico_cessao', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zhistorico_cessao', 'Zhistorico_cessaoidhistorico_cessao', 'idhistorico_cessao', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zhistorico_cessao', 'Zhistorico_cessaoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zhistorico_cessao', 'Zhistorico_cessaohistorico', 'historico', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZParcelas', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParcelas', 'ZParcelasidparcelas', 'idparcelas', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParcelas', 'ZParcelasvenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParcelas', 'ZParcelasordem', 'ordem', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParcelas', 'ZParcelasdatavenci', 'datavenci', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParcelas', 'ZParcelasvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParcelas', 'ZParcelassaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParcelas', 'ZParcelasvalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZParcelas', 'ZParcelasnobanco', 'nobanco', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'Zincorp_loteame', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zincorp_loteame', 'Zincorp_loteameincorporador_idincorporador', 'incorporador_idincorporador', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zincorp_loteame', 'Zincorp_loteameloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zincorp_loteame', 'Zincorp_loteamepercent_parte', 'percent_parte', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zincorp_loteame', 'Zincorp_loteamecodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zincorp_loteame', 'Zincorp_loteamedigito_dif', 'digito_dif', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'Zrepresentantes', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrepresentantes', 'Zrepresentantesidrepresentantes', 'idrepresentantes', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrepresentantes', 'Zrepresentantesincorporador_idincorporador', 'incorporador_idincorporador', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrepresentantes', 'Zrepresentantesnomerepre', 'nomerepre', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zrepresentantes', 'Zrepresentantesdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'Zcheque', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcheque', 'Zchequeidcheque', 'idcheque', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcheque', 'Zchequebanco', 'banco', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcheque', 'Zchequeagencia', 'agencia', TStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcheque', 'Zchequedono', 'dono', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcheque', 'ZchequeCPF_CNPJ', 'CPF_CNPJ', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcheque', 'Zchequevalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcheque', 'Zchequeemissao', 'emissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcheque', 'Zchequedeposito', 'deposito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcheque', 'Zchequealias', 'alias', TStringField, fkData, 12, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcheque', 'Zchequeprorrogado', 'prorrogado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcheque', 'Zchequeconta', 'conta', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcheque', 'Zchequeidrecebimento', 'idrecebimento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcheque', 'Zchequeidpagamento', 'idpagamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcheque', 'Zchequenumero', 'numero', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcheque', 'Zchequeidvinculo', 'idvinculo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcheque', 'ZchequeSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'Zcheque', 'Zchequesq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(Tfrm_baixa, 'ZConta_bancaria', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConta_bancaria', 'ZConta_bancariaidconta_bancaria', 'idconta_bancaria', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConta_bancaria', 'ZConta_bancarian_banco', 'n_banco', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConta_bancaria', 'ZConta_bancarian_agencia', 'n_agencia', TStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConta_bancaria', 'ZConta_bancarian_agencia_v', 'n_agencia_v', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConta_bancaria', 'ZConta_bancarian_conta', 'n_conta', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConta_bancaria', 'ZConta_bancarian_conta_v', 'n_conta_v', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConta_bancaria', 'ZConta_bancarianomebanco', 'nomebanco', TStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConta_bancaria', 'ZConta_bancarian_no_banco', 'n_no_banco', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConta_bancaria', 'ZConta_bancarianomeresposavel', 'nomeresposavel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConta_bancaria', 'ZConta_bancarianometitular', 'nometitular', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConta_bancaria', 'ZConta_bancariaaberta', 'aberta', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConta_bancaria', 'ZConta_bancariasenha_interna', 'senha_interna', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConta_bancaria', 'ZConta_bancariasenha_conta', 'senha_conta', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConta_bancaria', 'ZConta_bancariaativa', 'ativa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConta_bancaria', 'ZConta_bancariajurosemboleto', 'jurosemboleto', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConta_bancaria', 'ZConta_bancariataxadiaria', 'taxadiaria', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConta_bancaria', 'ZConta_bancariavaloroupercent', 'valoroupercent', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConta_bancaria', 'ZConta_bancariacarteira', 'carteira', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConta_bancaria', 'ZConta_bancariaboleto', 'boleto', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConta_bancaria', 'ZConta_bancariaidparticipante', 'idparticipante', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConta_bancaria', 'ZConta_bancarian_dif_empreed', 'n_dif_empreed', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConta_bancaria', 'ZConta_bancariaconvenio', 'convenio', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZConta_bancaria', 'ZConta_bancariadoc_titular', 'doc_titular', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
