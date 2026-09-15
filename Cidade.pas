unit Cidade;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, wwdbdatetimepicker, StdCtrls, DBCtrls, Mask, Grids, DBGrids, ExtCtrls, XBanner, DB, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, ComCtrls;

type
  TFrm_Cidade = class(TForm)
    XBanner14: TXBanner;
    Label19: TLabel;
    Label18: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    Panel4: TPanel;
    XBanner13: TXBanner;
    BtExcluir: TdxButtonArround;
    BtEditar: TdxButtonArround;
    BtIncluir: TdxButtonArround;
    BtCancelar: TdxButtonArround;
    BtGravar: TdxButtonArround;
    BtPesquisar: TdxButtonArround;
    BtRelatorio: TdxButtonArround;
    BtUltimo: TdxButtonArround;
    BTProximo: TdxButtonArround;
    BTAnterior: TdxButtonArround;
    BTPrimeiro: TdxButtonArround;
    DBGCid: TDBGrid;
    DBECodCid: TDBEdit;
    DBECidNome: TDBEdit;
    DBECidDDD: TDBEdit;
    DBCBCidEstado: TDBComboBox;
    DBECidEmanc: TwwDBDateTimePicker;
    BtFechar: TdxButtonArround;
    Label1: TLabel;
    DBECep: TDBEdit;
    Registros: TLabel;
    LReg: TLabel;
    dxButton1: TdxButtonArround;
    ZQlog_localidade: TZQuery;
    DS_log_localidade: TDataSource;










    ProgressBar1: TProgressBar;
    procedure BTPrimeiroClick(Sender: TObject);
    procedure BTAnteriorClick(Sender: TObject);
    procedure ativacidade;
    procedure desativacidade;
    procedure botoes_cidade;
    procedure BTProximoClick(Sender: TObject);
    procedure BtUltimoClick(Sender: TObject);
    procedure BtGravarClick(Sender: TObject);
    procedure BtCancelarClick(Sender: TObject);
    procedure BtIncluirClick(Sender: TObject);
    procedure BtEditarClick(Sender: TObject);
    procedure BtExcluirClick(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure BtFecharClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure DBGCidKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGCidMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBGCidEnter(Sender: TObject);
    procedure DBGCidDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGCidTitleClick(Column: TColumn);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure dxButton1Click(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_Cidade: TFrm_Cidade;

implementation

uses tabelas, Funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrm_Cidade.BTPrimeiroClick(Sender: TObject);
begin
  DM_Tabelas.ZQCidade.First;
  botoes_cidade;
end;

procedure TFrm_Cidade.BTAnteriorClick(Sender: TObject);
begin
  DM_Tabelas.ZQCidade.Prior;
  botoes_cidade;
end;

procedure TFrm_Cidade.ativacidade;
Begin
  DBECidNome.ReadOnly := False;
  DBCBCidEstado.ReadOnly := False;
  DBECidDDD.ReadOnly := False;
  DBECep.ReadOnly := False;
  DBECidEmanc.ReadOnly := False;
  BTPrimeiro.Enabled := False;
  BTAnterior.Enabled := False;
  BTProximo.Enabled := False;
  BtUltimo.Enabled := False;
  BtGravar.Enabled := True;
  BtCancelar.Enabled := True;
  BtIncluir.Enabled := False;
  BtEditar.Enabled := False;
  BtExcluir.Enabled := False;
  BtPesquisar.Enabled := False;
  BtRelatorio.Enabled := False;
  DBGCid.Enabled := False;
end;

procedure TFrm_Cidade.desativacidade;
Begin
  DBECidNome.ReadOnly := true;
  DBCBCidEstado.ReadOnly := true;
  DBECep.ReadOnly := True;
  DBECidDDD.ReadOnly := true;
  DBECidEmanc.ReadOnly := true;
  BTPrimeiro.Enabled := true;
  BTAnterior.Enabled := true;
  BTProximo.Enabled := true;
  BtUltimo.Enabled := true;
  BtGravar.Enabled := False;
  BtCancelar.Enabled := False;
  BtIncluir.Enabled := true;
  BtEditar.Enabled := true;
  BtExcluir.Enabled := true;
  BtPesquisar.Enabled := true;
  BtRelatorio.Enabled := true;
  DBGCid.Enabled := True;
end;

procedure TFrm_Cidade.botoes_cidade;
Begin
  BTPrimeiro.Enabled := True;
  BTAnterior.Enabled := True;
  BTProximo.Enabled := True;
  BtUltimo.Enabled := True;
  if DM_Tabelas.ZQCidade.Recno<=1 Then Begin
    BTPrimeiro.Enabled := False;
    BTAnterior.Enabled := False;
  end;
  if DM_Tabelas.ZQCidade.RecNo=DM_Tabelas.ZQCidade.RecordCount Then Begin
    BTProximo.Enabled := False;
    BtUltimo.Enabled := False;
  end;
  LReg.Caption := inttostr(DM_Tabelas.ZQCidade.RecordCount);
End;

procedure TFrm_Cidade.BTProximoClick(Sender: TObject);
begin
  DM_Tabelas.ZQCidade.Next;;
  botoes_cidade;
end;

procedure TFrm_Cidade.BtUltimoClick(Sender: TObject);
begin
  DM_Tabelas.ZQCidade.Last;
  botoes_cidade;
end;

procedure TFrm_Cidade.BtGravarClick(Sender: TObject);
begin
  if not Verif_senha('Cidade','Gravar inclusão ou edição','cidade: '+DM_Tabelas.ZQCidade.FieldByName('nomecid').AsString) Then Exit;
  DM_Tabelas.ZQCidade.Post;
  desativacidade;
  DBGCid.SetFocus;
end;

procedure TFrm_Cidade.BtCancelarClick(Sender: TObject);
begin
  if not Verif_senha('Cidade','Cancelar inclusão ou edição','cidade: '+DM_Tabelas.ZQCidade.FieldByName('nomecid').AsString) Then Exit;
  DM_Tabelas.ZQCidade.Cancel;
  desativacidade;
  DBGCid.SetFocus;
end;

procedure TFrm_Cidade.BtIncluirClick(Sender: TObject);
begin
  if not Verif_senha('Cidade','Inclusão','') Then Exit;
  DM_Tabelas.ZQCidade.Insert;
  ativacidade;
  DBECidNome.SetFocus;

end;

procedure TFrm_Cidade.BtEditarClick(Sender: TObject);
begin
  if not Verif_senha('Cidade','Edição','cidade: '+DM_Tabelas.ZQCidade.FieldByName('idcidade').Text+' - '+DM_Tabelas.ZQCidade.FieldByName('nomecid').AsString) Then Exit;
  DM_Tabelas.ZQCidade.Edit;
  ativacidade;
  DBECidNome.SetFocus;

end;

procedure TFrm_Cidade.BtExcluirClick(Sender: TObject);
begin
  if not Verif_senha('Cidade','Exclusão','cidade: '+DM_Tabelas.ZQCidade.FieldByName('idcidade').Text+' - '+DM_Tabelas.ZQCidade.FieldByName('nomecid').AsString) Then Exit;
  if simnao('Confirma a exclusão da cidade ?','SIM') then
  begin
    if DM_Tabelas.ZQCidade.RecordCount>0 Then
      DM_Tabelas.ZQCidade.Delete;
  end;
end;

procedure TFrm_Cidade.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrm_Cidade.BtFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrm_Cidade.FormShow(Sender: TObject);
begin
  DM_Tabelas.ZQCidade.open;
  desativacidade;
  DBGCid.SetFocus;
end;

procedure TFrm_Cidade.DBGCidKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  botoes_cidade;
end;

procedure TFrm_Cidade.DBGCidMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  botoes_cidade;
end;

procedure TFrm_Cidade.DBGCidEnter(Sender: TObject);
begin
  botoes_cidade;
end;

procedure TFrm_Cidade.DBGCidDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBECodCid.Text = DM_Tabelas.ZQCidade.FieldByName('idcidade').Text Then Begin
    DBGCid.Canvas.Brush.Color :=$006CFFFF;
    DBGCid.Canvas.Font.Color := $00A80000;
    DBGCid.Canvas.Font.Style := [FsBold];
  end;
  DBGCid.DefaultDrawDataCell(Rect, DBGCid.columns[datacol].field, State);

end;

procedure TFrm_Cidade.DBGCidTitleClick(Column: TColumn);
var
  campo : string;
  varcodigo : integer;
begin
  campo:=column.fieldname;
  DBGCid.Columns[0].Color := clWindow;
  DBGCid.Columns[1].Color := clWindow;
  DBGCid.Columns[2].Color := clWindow;
  DBGCid.Columns[3].Color := clWindow;
  DBGCid.Columns[4].Color := clWindow;
  DBGCid.Columns[5].Color := clWindow;
  application.processmessages;
  varcodigo := DM_Tabelas.ZQCidade.FieldByName('idcidade').AsLargeInt;
  DM_Tabelas.ZQCidade.sql.clear;
  DM_Tabelas.ZQCidade.SQL.Add('select idcidade,nomecid,ddd,emancipacao,estado,cepgeral from cidade');
  DM_Tabelas.ZQCidade.SQL.Add(' order by '+campo);
  DM_Tabelas.ZQCidade.Open;
  DM_Tabelas.ZQCidade.Locate('idcidade',varcodigo,[]);
  column.color:=$00EDF0F1;
  botoes_cidade;
end;

procedure TFrm_Cidade.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  DM_Tabelas.ZQCidade.close;
  Frm_Cidade:=nil;
  Action:=CaFree;
end;

procedure TFrm_Cidade.dxButton1Click(Sender: TObject);
begin
  DM_Tabelas.ZQCidade.sql.clear;
  DM_Tabelas.ZQCidade.SQL.Add('select idcidade,nomecid,ddd,emancipacao,estado,cepgeral from cidade');
  DM_Tabelas.ZQCidade.SQL.Add(' order by nomecid');
  DM_Tabelas.ZQCidade.Open;
  ZQlog_localidade.open;
  ZQlog_localidade.First;
  ProgressBar1.Position:=0;
  ProgressBar1.Max:=ZQlog_localidade.RecordCount;
  ProgressBar1.Visible:=true;
  ZQlog_localidade.DisableControls;
  while not ZQlog_localidade.eof do
  begin
    ProgressBar1.Position:=ZQlog_localidade.RecNo;
    if uppercase(removeacentos(ZQlog_localidade.FieldByName('loc_no').AsString))=uppercase(removeacentos(DM_Tabelas.ZQCidade.FieldByName('nomecid').AsString)) then
    begin
       ZQlog_localidade.Next;
    end
    else
    begin
      DM_Tabelas.ZQCidade.Insert;
      DM_Tabelas.ZQCidade.FieldByName('nomecid').AsString:=uppercase(removeacentos(ZQlog_localidade.FieldByName('loc_no').AsString));
      DM_Tabelas.ZQCidade.FieldByName('estado').AsString:=uppercase(ZQlog_localidade.FieldByName('ufe_sg').AsString);
      DM_Tabelas.ZQCidade.FieldByName('cepgeral').AsString:=copy(ZQlog_localidade.FieldByName('cep').AsString,1,2)+'.'+copy(ZQlog_localidade.FieldByName('cep').AsString,3,3)+'-'+copy(ZQlog_localidade.FieldByName('cep').AsString,6,3);
      DM_Tabelas.ZQCidade.post;
      ZQlog_localidade.Next;
    end;
  end;
  ZQlog_localidade.EnableControls;
  ProgressBar1.Position:=0;
  ProgressBar1.Visible:=false;
  ZQlog_localidade.close;
  DM_Tabelas.ZQCidade.Refresh; 
end;


procedure TFrm_Cidade.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_Cidade, 'ZQlog_localidade', False);
  RegisterRuntimeField(TFrm_Cidade, 'ZQlog_localidade', 'ZQlog_localidadeloc_nu_sequencial', 'loc_nu_sequencial', TIntegerField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cidade, 'ZQlog_localidade', 'ZQlog_localidadeloc_nosub', 'loc_nosub', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cidade, 'ZQlog_localidade', 'ZQlog_localidadeloc_no', 'loc_no', TWideStringField, fkData, 60, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cidade, 'ZQlog_localidade', 'ZQlog_localidadecep', 'cep', TWideStringField, fkData, 16, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cidade, 'ZQlog_localidade', 'ZQlog_localidadeufe_sg', 'ufe_sg', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cidade, 'ZQlog_localidade', 'ZQlog_localidadeloc_in_situacao', 'loc_in_situacao', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cidade, 'ZQlog_localidade', 'ZQlog_localidadeloc_in_tipo_localidade', 'loc_in_tipo_localidade', TWideStringField, fkData, 1, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cidade, 'ZQlog_localidade', 'ZQlog_localidadeloc_nu_sequencial_sub', 'loc_nu_sequencial_sub', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cidade, 'ZQlog_localidade', 'ZQlog_localidadeloc_key_dne', 'loc_key_dne', TWideStringField, fkData, 16, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cidade, 'ZQlog_localidade', 'ZQlog_localidadetemp', 'temp', TWideStringField, fkData, 8, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.

