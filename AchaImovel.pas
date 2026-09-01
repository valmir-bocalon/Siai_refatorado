unit AchaImovel;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, StdCtrls, XBanner, Mask, DBCtrls, Data.DB;

type
  TFrm_AchaImovel = class(TForm)
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
    Label5: TLabel;
    Label6: TLabel;
    EComprador: TEdit;
    Codpart: TEdit;
    procedure FormShow(Sender: TObject);
    procedure CBLoteKeyPress(Sender: TObject; var Key: Char);
    procedure DBGImovelKeyPress(Sender: TObject; var Key: Char);
    procedure CBLoteChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DBGImovelDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure ECompradorKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_AchaImovel: TFrm_AchaImovel;

implementation

uses Tabelas, funcoes, uRuntimeFields;
{$R *.dfm}

procedure TFrm_AchaImovel.FormShow(Sender: TObject);
begin
  Codpart.Text :='';
  EComprador.Text := '';
  CBLote.Clear;
  CBQuadra.Clear;
  Empreendimento.Clear;
  DM_Tabelas.ZQImovelTemp.SQL.Clear;
  DM_Tabelas.ZQImovelTemp.SQL.Add('SELECT * FROM imovel join loteamento on loteamento_idloteamento=idloteamento');
  DM_Tabelas.ZQImovelTemp.SQL.Add(' where 1=1');
  DM_Tabelas.ZQImovelTemp.SQL.Add(' group by lote');
  DM_Tabelas.ZQImovelTemp.SQL.Add(' order by lote,quadra');
  DM_Tabelas.ZQImovelTemp.Open;
  while not DM_Tabelas.ZQImovelTemp.Eof do begin
    CBLote.Items.Add(DM_tabelas.ZQImovelTemp.FieldByName('lote').AsString);
    DM_Tabelas.ZQImovelTemp.Next;
  end;
  DM_Tabelas.ZQImovelTemp.Close;
  DM_Tabelas.ZQImovelTemp.SQL[2] := ' group by quadra';
  DM_Tabelas.ZQImovelTemp.Open;
  while not DM_Tabelas.ZQImovelTemp.Eof do begin
    CBQuadra.Items.Add(DM_tabelas.ZQImovelTemp.FieldByName('quadra').AsString);
    DM_Tabelas.ZQImovelTemp.Next;
  end;
  DM_Tabelas.ZQImovelTemp.Close;
  DM_Tabelas.ZQImovelTemp.SQL[2] := ' group by loteamento_idloteamento';
  DM_Tabelas.ZQImovelTemp.Open;
  while not DM_Tabelas.ZQImovelTemp.Eof do begin
    Empreendimento.Items.Add(DM_tabelas.ZQImovelTemp.FieldByName('apelido').AsString);
    DM_Tabelas.ZQImovelTemp.Next;
  end;
  DM_Tabelas.ZQImovelTemp.Close;
  DM_Tabelas.ZQImovelTemp.SQL[2] := ' ';
  DM_Tabelas.ZQImovelTemp.Open;
  Label4.Caption := 'F';
  Label5.Caption := inttostr(DM_Tabelas.ZQImovelTemp.RecordCount);
end;

procedure TFrm_AchaImovel.CBLoteKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #27 Then

    Close;
  if key = #13 then Begin
    Label4.Caption := 'T';
    if DM_Tabelas.ZQImovelTemp.RecordCount>1 Then
      DBGImovel.SetFocus
    else
      Close;
  end;
end;

procedure TFrm_AchaImovel.DBGImovelKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #27 Then
    Label4.Caption := 'F';
  if key = #13 Then
    Label4.Caption := 'T';
  if (key = #27) or (key = #13) Then
    Close;
end;

procedure TFrm_AchaImovel.CBLoteChange(Sender: TObject);
Var
  varfilt : string;
begin
  varfilt := '';
  DM_Tabelas.ZQImovelTemp.Close;
  varfilt := 'lote like '+quotedstr(CBLote.Text+'%');
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
  if not empty(Codpart.text) Then Begin
    varfilt := varfilt + ' and Exists (select * from venda join comprador on venda_idvenda=idvenda where paticipante_idpaticipante ='+Codpart.Text+' and imovel=idimovel)';
  end;
  DM_Tabelas.ZQImovelTemp.SQL[1] := 'where '+varfilt;
  DM_Tabelas.ZQImovelTemp.Open;
  Label5.Caption := inttostr(DM_Tabelas.ZQImovelTemp.RecordCount);
end;

procedure TFrm_AchaImovel.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  if Label4.Caption = 'T' then
    DM_TAbelas.ZQImovel.Locate('idimovel',DM_Tabelas.ZQImovelTemp.FieldByName('idimovel').AsLargeInt,[]);
end;

procedure TFrm_AchaImovel.DBGImovelDrawColumnCell(Sender: TObject;
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

procedure TFrm_AchaImovel.ECompradorKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (key=13) and (not empty(EComprador.Text)) Then
    if AchaCLiente(200, 10,Ecomprador.Text) Then Begin
      Ecomprador.Text := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').Text;
      Codpart.Text := DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').Text;
    end
    else Begin
      Ecomprador.Text := '';
      Codpart.Text := '';
    end;
end;


procedure TFrm_AchaImovel.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
