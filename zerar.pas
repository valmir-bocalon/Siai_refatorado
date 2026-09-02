unit zerar;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset,
  Mask, ComCtrls;

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
  if ZQLoteamento.RecordCount=0 then
  begin
    showmessage('Fim de Processo !');
    exit;
  end;
  ProgressBar1.Position:=0;
  nlt:=strtoint(xlote.text);
  while not ZQLoteamento.Eof do
  begin
    zqimovel.First;
    if zqimovel.RecordCount>0 then
    begin
      msg1.Caption:='Imoveis                ';
      application.ProcessMessages;

      while not zqimovel.eof do
      begin
        application.ProcessMessages;
        zqrecebimento.first;

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
                    ZQReceb_Baixa.Delete;
                  end;
                end;
                ZQReBxHi.Delete;
              end;
              ZQrecebimento_historico_cessao.First;
              if ZQrecebimento_historico_cessao.RecordCount>0 then
              begin
                msg.Caption:='Hist. cessao recebimento';
                application.ProcessMessages;

                while not ZQrecebimento_historico_cessao.eof do
                begin
                  ZQrecebimento_historico_cessao.Delete;
                end;
              end;

              ZQrecebimento_historico_excluidos.First;
              if ZQrecebimento_historico_excluidos.RecordCount>0 then
              begin
                msg.Caption:='Hist. excl. recebimento';
                application.ProcessMessages;

                while not ZQrecebimento_historico_excluidos.eof do
                begin
                  ZQrecebimento_historico_excluidos.Delete;
                end;
              end;
              ZQRemes_Receb.First;
              if ZQRemes_Receb.RecordCount>0 then
              begin
                 while not ZQRemes_Receb.Eof do
                 begin
                   ZQRemes_Receb.Delete;
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
                ZQCheque.Delete;
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
                    ZQVendedor.Delete;
                  end;
                end;
                ZQAditamento.first;
                if ZQAditamento.RecordCount>0 then
                begin
                   msg.Caption:='Aditamento                ';
                   application.ProcessMessages;

                  while not ZQAditamento.eof do
                  begin
                    ZQAditamento.delete;
                  end;  
                end;
                ZQhistorico_cessao.first;
                if ZQhistorico_cessao.RecordCount>0 then
                begin
                  msg.Caption:='Hist.Cessao recebimento';
                  application.ProcessMessages;

                  while not ZQhistorico_cessao.eof do
                  begin
                    ZQhistorico_cessao.delete;
                  end;
                end;
                ZQQuitacao.first;
                if ZQQuitacao.RecordCount>0 then
                begin
                  msg.Caption:='Quitação                    ';
                  application.ProcessMessages;

                  while not ZQQuitacao.eof do
                  begin
                    ZQQuitacao.delete;
                  end;
                end;

                ZQComprador.first;
                if ZQComprador.RecordCount>0 then
                begin
                  msg.Caption:='Comprador                    ';
                  application.ProcessMessages;

                  while not ZQComprador.eof do
                  begin
                    ZQComprador.delete;
                  end;
                end;

                ZQComprador_cessao.first;
                if ZQComprador_cessao.RecordCount>0 then
                begin
                  msg.Caption:='Comprador Cessão         ';
                  application.ProcessMessages;

                  while not ZQComprador_cessao.eof do
                  begin
                    ZQComprador_cessao.delete;
                  end;
                end;

                ZQParcelas.first;
                if ZQParcelas.RecordCount>0 then
                begin
                  msg.Caption:='Parcelas                 ';
                  application.ProcessMessages;

                  while not ZQParcelas.eof do
                  begin
                    ZQParcelas.delete;
                  end;
                end;

                ZQVenda.delete;
              end;
            end;
            zqrecebimento.Delete;
          end;
        end;
        zqimovel.Delete;
      end;
    end;

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
            ZQrepresentantes.Delete;
          end;
        end;
        {ZqParticipante.First;
        if ZqParticipante.RecordCount>0 then
        begin
          msg.Caption:='Participantes                          ';
          application.ProcessMessages;

        {  while not ZqParticipante.Eof do
          begin
            ZQresponsavel.First;
            if ZQresponsavel.RecordCount>0 then
            begin
              msg.Caption:='Responsaveis                       ';
              application.ProcessMessages;

              while not ZQresponsavel.Eof do
              begin
                ZQresponsavel.Delete;
              end;
            end;
            ZQCompr_conjuge.first;
            if ZQCompr_conjuge.RecordCount>0 then
            begin
              msg.Caption:='Cônjuge                                ';
              application.ProcessMessages;

              while not ZQCompr_conjuge.Eof do
              begin
                ZQCompr_conjuge.Delete;
              end;
            end;
            ZQDividido.First;
            if ZQDividido.RecordCount>0 then
            begin
              while not ZQDividido.Eof do
              begin
                ZQDividido.delete;
              end;
            end;

            {if ZQConta_bancaria.RecordCount>0 then
            begin
              msg.Caption:='Conta Bancária                                ';
              application.ProcessMessages;
              ZQConta_bancaria.First;
              while not ZQConta_bancaria.Eof do
              begin
                ZQConta_bancaria.delete;
              end;
            end;}

          {  ZqParticipante.Delete;
          end;
        end;  }

        ZQincorp_loteame.Delete;
      end;

    end;

    ZQProcuradores.First;
    if ZQProcuradores.RecordCount>0 then
    begin
      msg.Caption:='Procuradores                                                ';
      application.ProcessMessages;

      while not ZQProcuradores.Eof do
      begin
        ZQProcuradores.Delete;
      end;
    end;

    ZQQuadras.First;
    if ZQQuadras.RecordCount>0 then
    begin
      msg.Caption:='Quadras                                                   ';
      application.ProcessMessages;

      while not ZQQuadras.Eof do
      begin
        ZQQuadras.Delete;
      end;
    end;
    ZQLoteamento.delete;
  end;
  ProgressBar1.Position:=0;
  msg.Caption:='Finalizado                                                   ';
  application.ProcessMessages;
  msg1.Caption:='Processo                ';
  application.ProcessMessages;
  inc(nlt);
  {if (nlt=1) or (nlt=4) or (nlt=6) or (nlt=7) or (nlt=8) or (nlt=13) or (nlt=15) or (nlt=22) and (not zqloteamento.Eof) then
  begin
    inc(nlt);
    if (nlt<>1) and (nlt<>4) and (nlt<>6) and (nlt<>7) and (nlt<>8) and (nlt<>13) and (nlt<>15) and (nlt<>22) then
    begin
      xlote.Text:=inttostr(nlt);
      Button1Click(self);
    end
    else
    begin
      inc(nlt);
      if (nlt<>1) and (nlt<>4) and (nlt<>6) and (nlt<>7) and (nlt<>8) and (nlt<>13) and (nlt<>15) and (nlt<>22) then
      begin
        xlote.Text:=inttostr(nlt);
        Button1Click(self);
      end
      else
      begin
        inc(nlt);
        if (nlt<>1) and (nlt<>4) and (nlt<>6) and (nlt<>7) and (nlt<>8) and (nlt<>13) and (nlt<>15) and (nlt<>22) then
        begin
          xlote.Text:=inttostr(nlt);
          Button1Click(self);
        end
        else
        begin
           inc(nlt);
           if (nlt<>1) and (nlt<>4) and (nlt<>6) and (nlt<>7) and (nlt<>8) and (nlt<>13) and (nlt<>15) and (nlt<>22) then
           begin
             xlote.Text:=inttostr(nlt);
             Button1Click(self);
           end;
        end;
      end;
    end;
  end}

 { if (nlt=17)then
  begin
    inc(nlt);
    xlote.Text:=inttostr(nlt);
//    Button1Click(self);
  end;
  if (nlt=28)  then
  begin
    inc(nlt);
    xlote.Text:=inttostr(nlt);
  //  Button1Click(self);
  end;

//  if nlt<=29 then
//  begin
    xlote.Text:=inttostr(nlt);
    Button1Click(self);
//  end;}
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
  xlote.Text:='4';
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
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugenomecida', 'nomecida', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'natural', '', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugeestcida', 'estcida', TStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'natural', '', 'idcidade', 'estado', True);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugeidparti', 'idparti', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugeregime', 'regime', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugecomarca', 'comarca', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugecidade_cart', 'cidade_cart', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugedata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugefolha', 'folha', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugelivro', 'livro', TStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugecomar_cidnome', 'comar_cidnome', TStringField, fkLookup, 10, 0, False, '', '', '', '', 0, 'comarca', '', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugecomar_cidest', 'comar_cidest', TStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'comarca', '', 'idcidade', 'estado', True);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugecart_cidnome', 'cart_cidnome', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cidade_cart', '', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(Tfrm_baixa, 'ZQConjuge', 'ZQConjugecart_cidest', 'cart_cidest', TStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade_cart', '', 'idcidade', 'estado', True);
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

end.
