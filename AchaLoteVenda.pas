unit AchaLoteVenda;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, StdCtrls, XBanner, Mask, DBCtrls, Data.DB;

type
  TFrm_AchaLoteVenda = class(TForm)
    XBanner4: TXBanner;
    Label19: TLabel;
    CBLote: TComboBox;
    CBQuadra: TComboBox;
    Empreendimento: TComboBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    DBGImovel: TDBGrid;
    Label4: TLabel;
    DBECod: TDBEdit;
    procedure FormShow(Sender: TObject);
    procedure CBLoteKeyPress(Sender: TObject; var Key: Char);
    procedure DBGImovelKeyPress(Sender: TObject; var Key: Char);
    procedure CBLoteChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DBGImovelDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_AchaLoteVenda: TFrm_AchaLoteVenda;

implementation

uses Tabelas, funcoes, uRuntimeFields;
{$R *.dfm}

procedure TFrm_AchaLoteVenda.FormShow(Sender: TObject);
begin
  CBLote.Clear;
  CBQuadra.Clear;
  Empreendimento.Clear;
  DM_Tabelas.ZQImovelTemp.SQL.Clear;


  DM_Tabelas.ZQImovelTemp.SQL.Add('SELECT idimovel,loteamento_idloteamento,quadra,lote,lado,medidafrente,medidafundo,medidaesquerda,medidadedireita,area,confrontofrente,confrontofundo,confrontodireita,confrontoesquerda,');
  DM_Tabelas.ZQImovelTemp.SQL.Add(' esquinanome,esquinalinha,esquinacurva,valorvenal,Obervacaoloteamento,cornomapa,disponivel,');
  DM_Tabelas.ZQImovelTemp.SQL.Add(' idloteamento,cidade_idcidade,matriculaloteamento,datacadastro,datainalguracao,dataconclusao,logotipo,mapa,testemunha1,testemunha2,testemunha3,');
  DM_Tabelas.ZQImovelTemp.SQL.Add(' testemunha4,Observincorp,bairro,pasta_mapa,nomeloteamento,tipo,codcontabancaria,apelido,Perc_comissao ');
  DM_Tabelas.ZQImovelTemp.SQL.Add(' FROM imovel join loteamento on loteamento_idloteamento=idloteamento ');
//  DM_Tabelas.ZQImovelTemp.SQL.Add('SELECT * FROM imovel join loteamento on loteamento_idloteamento=idloteamento');
  DM_Tabelas.ZQImovelTemp.SQL.Add(' where (disponivel is null or disponivel ='+quotedstr('SIM')+')');
  if DM_Tabelas.ZQVenda.RecordCount>0 Then
    DM_Tabelas.ZQImovelTemp.SQL.Add(' or idimovel = '+quotedstr(DM_Tabelas.ZQVenda.FieldByName('imovel').Text))
  else
    DM_Tabelas.ZQImovelTemp.SQL.Add(' and 1=1');
  DM_Tabelas.ZQImovelTemp.SQL.Add(' group by lote');
  DM_Tabelas.ZQImovelTemp.SQL.Add(' order by lote,quadra');
  DM_Tabelas.ZQImovelTemp.Open;
  while not DM_Tabelas.ZQImovelTemp.Eof do begin
    CBLote.Items.Add(DM_tabelas.ZQImovelTemp.FieldByName('lote').AsString);
    DM_Tabelas.ZQImovelTemp.Next;
  end;
  DM_Tabelas.ZQImovelTemp.Close;
  DM_Tabelas.ZQImovelTemp.SQL[7] := ' group by quadra';
  DM_Tabelas.ZQImovelTemp.Open;
  while not DM_Tabelas.ZQImovelTemp.Eof do begin
    CBQuadra.Items.Add(DM_tabelas.ZQImovelTemp.FieldByName('quadra').AsString);
    DM_Tabelas.ZQImovelTemp.Next;
  end;
  DM_Tabelas.ZQImovelTemp.Close;
  DM_Tabelas.ZQImovelTemp.SQL[7] := ' group by loteamento_idloteamento';
  DM_Tabelas.ZQImovelTemp.Open;
  while not DM_Tabelas.ZQImovelTemp.Eof do begin
    Empreendimento.Items.Add(DM_tabelas.ZQImovelTemp.FieldByName('apelido').AsString);
    DM_Tabelas.ZQImovelTemp.Next;
  end;
  DM_Tabelas.ZQImovelTemp.Close;
  DM_Tabelas.ZQImovelTemp.SQL[7] := ' ';
  DM_Tabelas.ZQImovelTemp.Open;
  DM_Tabelas.ZQImovelTemp.Locate('idimovel',DM_Tabelas.ZQVenda.FieldByName('imovel').AsLargeInt,[]);
  Label4.Caption := 'F';
end;

procedure TFrm_AchaLoteVenda.CBLoteKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #27 Then
    Close;
  if key = #13 then Begin
    Label4.Caption := 'T';
    DM_Tabelas.ZQImovelTemp.Locate('idimovel',DM_Tabelas.ZQVenda.FieldByName('imovel').AsLargeInt,[]);
    if DM_Tabelas.ZQImovelTemp.RecordCount>1 Then
      DBGImovel.SetFocus
    else
      Close;
  end;
end;

procedure TFrm_AchaLoteVenda.DBGImovelKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #27 Then
    Label4.Caption := 'F';
  if key = #13 Then
    Label4.Caption := 'T';
  if (key = #27) or (key = #13) Then
    Close;
end;

procedure TFrm_AchaLoteVenda.CBLoteChange(Sender: TObject);
Var
  varfilt : string;
begin
  varfilt := '';
  DM_Tabelas.ZQImovelTemp.Close;
  varfilt := 'lote like '+quotedstr(CBLote.Text+'%')+' and (disponivel is null or disponivel ='+quotedstr('SIM')+')';
  if not empty(CBQuadra.Text) Then Begin
    if not empty(varfilt) then
      varfilt := varfilt +' and ';
    varfilt := varfilt +'quadra like '+quotedstr(CBQuadra.Text+'%');
  end;
  if not empty(Empreendimento.Text) Then Begin
    if not empty(varfilt) then
      varfilt := varfilt +' and ';
    varfilt := varfilt +' apelido like '+quotedstr('%'+Empreendimento.Text+'%');
  end;
  DM_Tabelas.ZQImovelTemp.SQL[5] := 'where '+varfilt;
  DM_Tabelas.ZQImovelTemp.Open;
  
end;

procedure TFrm_AchaLoteVenda.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  if Label4.Caption = 'T' then
  begin
    DM_TAbelas.ZQVenda.FieldByName('imovel').AsLargeInt := DM_Tabelas.ZQImovelTemp.FieldByName('idimovel').AsLargeInt;
    DM_TAbelas.ZQVenda.FieldByName('Perc_comissao').AsFloat := DM_Tabelas.ZQImovelTemp.FieldByName('Perc_comissao').AsFloat;
  end;
end;

procedure TFrm_AchaLoteVenda.DBGImovelDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBECod.Text = DM_Tabelas.ZQImovelTemp.FieldByName('idimovel').Text Then Begin
    DBGImovel.Canvas.Brush.Color :=$006CFFFF;
    DBGImovel.Canvas.Font.Color := $00A80000;
    DBGImovel.Canvas.Font.Style := [FsBold];
  end;
  DBGImovel.DefaultDrawDataCell(Rect, DBGImovel.columns[datacol].field, State);

end;


procedure TFrm_AchaLoteVenda.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
