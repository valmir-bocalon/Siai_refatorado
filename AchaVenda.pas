unit AchaVenda;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, XBanner, Grids, DBGrids, Mask, Data.DB;

type
  TFrm_AchaVenda = class(TForm)
    XBanner4: TXBanner;
    Label19: TLabel;
    DBGAchavenda: TDBGrid;
    Epesquisa: TEdit;
    Label3: TLabel;
    edlote: TMaskEdit;
    procedure FormShow(Sender: TObject);
    procedure EpesquisaChange(Sender: TObject);
    procedure EpesquisaKeyPress(Sender: TObject; var Key: Char);
    procedure DBGAchavendaKeyPress(Sender: TObject; var Key: Char);
    procedure edloteChange(Sender: TObject);
    procedure edloteKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_AchaVenda: TFrm_AchaVenda;

implementation

uses tabelas, Funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrm_AchaVenda.FormShow(Sender: TObject);
begin
  Epesquisa.SetFocus;
end;

procedure TFrm_AchaVenda.EpesquisaChange(Sender: TObject);
begin
  DM_TAbelas.ZQAchaVenda.close;
  DM_TAbelas.ZQAchaVenda.SQL.Clear;
  DM_TAbelas.ZQAchaVenda.SQL.Add('select idvenda,datavenda,imovel,valorvenda, ');
  DM_TAbelas.ZQAchaVenda.SQL.Add('       idimovel,loteamento_idloteamento,quadra,lote,lado,medidafrente,medidafundo,medidaesquerda,medidadedireita,area,confrontofrente,confrontofundo,confrontodireita,');
  DM_TAbelas.ZQAchaVenda.SQL.Add('       confrontoesquerda,esquinanome,esquinalinha,esquinacurva,valorvenal,Obervacaoloteamento,cornomapa,disponivel,');
  DM_TAbelas.ZQAchaVenda.SQL.Add('       idcomprador,paticipante_idpaticipante,venda_idvenda,promissario,');
  DM_TAbelas.ZQAchaVenda.SQL.Add('       idpaticipante,nome_parte,doc1,doc2,endereco,bairro,cidade,cep,ende_cob,bairro_cob,cidade_cob,cep_cob,Fone1,fone2,fone3,tipopessoa,aniversario,nacionalidade,');
  DM_TAbelas.ZQAchaVenda.SQL.Add('       email,naturalidade,localdetrab,profissao,estadocivil,renda,observacao,codrecinc,codrecBx,codpaginc,codpagBx,fone4,complemento,complemento_cob,cadastrado,Perc_comissao,vlr_comissao,path_pdf ');
  DM_TAbelas.ZQAchaVenda.SQL.Add('  from venda');
  DM_TAbelas.ZQAchaVenda.SQL.Add(' join imovel on idimovel=imovel');
  DM_TAbelas.ZQAchaVenda.SQL.Add(' join comprador on venda_idvenda=idvenda');
  DM_TAbelas.ZQAchaVenda.SQL.Add(' join participante on paticipante_idpaticipante=idpaticipante');
  DM_TAbelas.ZQAchaVenda.SQL.Add(' where nome_parte like '+quotedstr('%'+Epesquisa.Text+'%'));
  DM_TAbelas.ZQAchaVenda.SQL.Add(' or quadra like '+quotedstr(Epesquisa.Text+'%'));
  DM_TAbelas.ZQAchaVenda.SQL.Add(' or idvenda = '+quotedstr(Epesquisa.Text));
  DM_TAbelas.ZQAchaVenda.Open;
end;

procedure TFrm_AchaVenda.EpesquisaKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #27 Then
    Close;
  if key = #13 Then Begin
    if DM_Tabelas.ZQAchaVenda.RecordCount>1 Then
      DBGAchavenda.SetFocus
    else if DM_Tabelas.ZQAchaVenda.RecordCount=1 Then
    begin
      DM_TAbelas.ZQVenda.Locate('idvenda',DM_Tabelas.ZQAchaVenda.FieldByName('idvenda').Text,[]);
      close;
    end  
    else
      close;
  end;
end;

procedure TFrm_AchaVenda.DBGAchavendaKeyPress(Sender: TObject;
  var Key: Char);
begin
  if (key = #27) or (key = #13) Then
  Begin
    DM_TAbelas.ZQVenda.close;
    DM_TAbelas.ZQVenda.SQL.Clear;
    DM_TAbelas.ZQVenda.sql.Add('select idvenda,datavenda,imovel,valorvenda,forma_reajuste,tabela_Price,Escriturado,marca,codigo_contrato_ref,Multa,Mora,Perc_comissao,vlr_comissao,path_pdf from venda where idvenda='+quotedstr(DM_Tabelas.ZQAchaVenda.FieldByName('idvenda').Text));
    DM_TAbelas.ZQVenda.open;
    DM_TAbelas.ZQVenda.Locate('idvenda',DM_Tabelas.ZQAchaVenda.FieldByName('idvenda').Text,[]);
    Close;
  end;

end;

procedure TFrm_AchaVenda.edloteChange(Sender: TObject);
begin
  DM_TAbelas.ZQAchaVenda.close;
  DM_TAbelas.ZQAchaVenda.SQL.Clear;
  DM_TAbelas.ZQAchaVenda.SQL.Add('select idvenda,datavenda,imovel,valorvenda, ');
  DM_TAbelas.ZQAchaVenda.SQL.Add('       idimovel,loteamento_idloteamento,quadra,lote,lado,medidafrente,medidafundo,medidaesquerda,medidadedireita,area,confrontofrente,confrontofundo,confrontodireita,');
  DM_TAbelas.ZQAchaVenda.SQL.Add('       confrontoesquerda,esquinanome,esquinalinha,esquinacurva,valorvenal,Obervacaoloteamento,cornomapa,disponivel,');
  DM_TAbelas.ZQAchaVenda.SQL.Add('       idcomprador,paticipante_idpaticipante,venda_idvenda,promissario,');
  DM_TAbelas.ZQAchaVenda.SQL.Add('       idpaticipante,nome_parte,doc1,doc2,endereco,bairro,cidade,cep,ende_cob,bairro_cob,cidade_cob,cep_cob,Fone1,fone2,fone3,tipopessoa,aniversario,nacionalidade,');
  DM_TAbelas.ZQAchaVenda.SQL.Add('       email,naturalidade,localdetrab,profissao,estadocivil,renda,observacao,codrecinc,codrecBx,codpaginc,codpagBx,fone4,complemento,complemento_cob,cadastrado,Perc_comissao,vlr_comissao,path_pdf ');
  DM_TAbelas.ZQAchaVenda.SQL.Add(' from venda');
  DM_TAbelas.ZQAchaVenda.SQL.Add(' join imovel on idimovel=imovel');
  DM_TAbelas.ZQAchaVenda.SQL.Add(' join comprador on venda_idvenda=idvenda');
  DM_TAbelas.ZQAchaVenda.SQL.Add(' join participante on paticipante_idpaticipante=idpaticipante');
  DM_TAbelas.ZQAchaVenda.SQL.Add(' where quadra like '+quotedstr(Epesquisa.Text+'%'));
  DM_TAbelas.ZQAchaVenda.SQL.Add(' and lote like '+quotedstr(edlote.Text+'%'));
  DM_TAbelas.ZQAchaVenda.Open;

end;

procedure TFrm_AchaVenda.edloteKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #27 Then
    Close;
  if key = #13 Then Begin
    if DM_Tabelas.ZQAchaVenda.RecordCount>1 Then
    begin
      DBGAchavenda.SetFocus;
    end
    else if DM_Tabelas.ZQAchaVenda.RecordCount=1 Then
    begin
      DM_TAbelas.ZQVenda.close;
      DM_TAbelas.ZQVenda.SQL.Clear;
      DM_TAbelas.ZQVenda.sql.Add('select idvenda,datavenda,imovel,valorvenda,forma_reajuste,tabela_Price,Escriturado,marca,codigo_contrato_ref,Multa,Mora,Perc_comissao,vlr_comissao,path_pdf from venda where idvenda='+quotedstr(DM_Tabelas.ZQAchaVenda.FieldByName('idvenda').Text));
      DM_TAbelas.ZQVenda.open;
      DM_TAbelas.ZQVenda.Locate('idvenda',DM_Tabelas.ZQAchaVenda.FieldByName('idvenda').Text,[])
    end
    else
    begin
      close;
    end;  
  end;

end;


procedure TFrm_AchaVenda.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
