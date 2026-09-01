unit Achaimoveis;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, StdCtrls, XBanner, Mask, DBCtrls, Data.DB;

type
  TFrm_Achaimoveis = class(TForm)
    XBanner2: TXBanner;
    Label2: TLabel;
    Eparticipante: TEdit;
    DBGPart: TDBGrid;
    Label1: TLabel;
    DBECod: TDBEdit;
    edlote: TMaskEdit;
    Label3: TLabel;
    Label4: TLabel;
    procedure EparticipanteChange(Sender: TObject);
    procedure EparticipanteKeyPress(Sender: TObject; var Key: Char);
    procedure DBGPartKeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
    procedure DBGPartDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure edloteChange(Sender: TObject);
    procedure edloteKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_Achaimoveis: TFrm_Achaimoveis;

implementation

uses tabelas, uRuntimeFields;

{$R *.dfm}

procedure TFrm_Achaimoveis.EparticipanteChange(Sender: TObject);
begin
    DM_Tabelas.Acha_imovel.SQL.Clear;
    DM_Tabelas.Acha_imovel.SQL.Add('SELECT * FROM imovel join loteamento on loteamento_idloteamento=idloteamento');
    DM_Tabelas.Acha_imovel.SQL.Add(' where quadra like '+quotedstr('%'+Eparticipante.Text+'%'));
    DM_Tabelas.Acha_imovel.SQL.Add(' or lote like '+quotedstr('%'+Eparticipante.Text+'%'));
    DM_Tabelas.Acha_imovel.SQL.Add(' or apelido like '+quotedstr('%'+Eparticipante.Text+'%'));
    DM_Tabelas.Acha_imovel.SQL.Add(' or idimovel = '+quotedstr(Eparticipante.Text));    
    DM_Tabelas.Acha_imovel.SQL.Add(' or (Exists (select * from venda join comprador on venda_idvenda=idvenda where paticipante_idpaticipante ='+quotedstr('%'+Eparticipante.Text+'%')+' and imovel=idimovel))');
    DM_Tabelas.Acha_imovel.SQL.Add(' group by lote');
    DM_Tabelas.Acha_imovel.SQL.Add(' order by lote,quadra');
    DM_Tabelas.Acha_imovel.Open;
end;

procedure TFrm_Achaimoveis.EparticipanteKeyPress(Sender: TObject;
  var Key: Char);
begin
  Label1.Caption := 'V';
  if key = #27 Then  Begin
    Label1.Caption := 'F';
    Close;
  end;
  if Key = #13 Then
  Begin
    Label1.Caption := 'V';
    close;
  end;
end;

procedure TFrm_Achaimoveis.DBGPartKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #27 Then  Begin
    Label1.Caption := 'F';
    Close;
  end
  else if key = #13 Then
  begin
    Label1.Caption := 'V';
    Close;
  end;  
end;

procedure TFrm_Achaimoveis.FormShow(Sender: TObject);
begin
  DM_Tabelas.Acha_imovel.close;
  DM_Tabelas.Acha_imovel.SQL.Clear;
  DM_Tabelas.Acha_imovel.SQL.Add('SELECT * FROM imovel join loteamento on loteamento_idloteamento=idloteamento');
  DM_Tabelas.Acha_imovel.SQL.Add(' where 1=1');
  DM_Tabelas.Acha_imovel.SQL.Add(' group by lote');
  DM_Tabelas.Acha_imovel.SQL.Add(' order by lote,quadra');
  DM_Tabelas.Acha_imovel.Open;
  Label1.Caption := 'F';
  edlote.Clear;
  Eparticipante.Clear;
  Eparticipante.SetFocus;
end;

procedure TFrm_Achaimoveis.DBGPartDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBECod.Text = DM_Tabelas.Acha_imovel.FieldByName('idimovel').Text Then Begin
    DBGPart.Canvas.Brush.Color :=$006CFFFF;
    DBGPart.Canvas.Font.Color := $00A80000;
  end;
  DBGPart.DefaultDrawDataCell(Rect, DBGPart.columns[datacol].field, State);

end;

procedure TFrm_Achaimoveis.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  if  Label1.Caption = 'V' then
      DM_TAbelas.ZQImovel.Locate('idimovel',DM_Tabelas.Acha_imovel.FieldByName('idimovel').AsLargeInt,[]);
  edlote.Clear;
  Eparticipante.Clear;
  DM_Tabelas.Acha_imovel.close;
end;

procedure TFrm_Achaimoveis.edloteChange(Sender: TObject);
begin
    DM_Tabelas.Acha_imovel.SQL.Clear;
    DM_Tabelas.Acha_imovel.SQL.Add('SELECT * FROM imovel join loteamento on loteamento_idloteamento=idloteamento');
    DM_Tabelas.Acha_imovel.SQL.Add(' where quadra like '+quotedstr('%'+Eparticipante.Text+'%'));
    DM_Tabelas.Acha_imovel.SQL.Add(' and lote like '+quotedstr('%'+edlote.Text+'%'));
    DM_Tabelas.Acha_imovel.SQL.Add(' group by lote');
    DM_Tabelas.Acha_imovel.SQL.Add(' order by lote,quadra');
    DM_Tabelas.Acha_imovel.Open;
end;

procedure TFrm_Achaimoveis.edloteKeyPress(Sender: TObject; var Key: Char);
begin
  Label1.Caption := 'V';
  if key = #27 Then  Begin
    Label1.Caption := 'F';
    Close;
  end;
  if Key = #13 Then
  Begin
    Label1.Caption := 'V';
    close;
  end;

end;


procedure TFrm_Achaimoveis.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
