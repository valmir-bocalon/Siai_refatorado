unit ipca;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, wwdbdatetimepicker, StdCtrls, DBCtrls, Mask, Grids, DBGrids,
    dxButton, ExtCtrls, XBanner, DB, dxCore2;

type
  TFrm_ipca = class(TForm)
    XBanner14: TXBanner;
    Label19: TLabel;
    Label18: TLabel;
    Label20: TLabel;
    Panel4: TPanel;
    XBanner13: TXBanner;
    BtExcluir: TdxButton;
    BtEditar: TdxButton;
    BtIncluir: TdxButton;
    BtCancelar: TdxButton;
    BtGravar: TdxButton;
    BtPesquisar: TdxButton;
    BtRelatorio: TdxButton;
    BtUltimo: TdxButton;
    BTProximo: TdxButton;
    BTAnterior: TdxButton;
    BTPrimeiro: TdxButton;
    DBGCid: TDBGrid;
    DBECodCid: TDBEdit;
    DBECidNome: TDBEdit;
    BtFechar: TdxButton;
    Registros: TLabel;
    LReg: TLabel;
    Label2: TLabel;
    DBEdit1: TDBEdit;
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
    procedure BtRelatorioClick(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_ipca: TFrm_ipca;

implementation

uses tabelas, Funcoes, RelIpca, uRuntimeFields;

{$R *.dfm}

procedure TFrm_ipca.BTPrimeiroClick(Sender: TObject);
begin
  DM_Tabelas.ZQipca.First;
  botoes_cidade;
end;

procedure TFrm_ipca.BTAnteriorClick(Sender: TObject);
begin
  DM_Tabelas.ZQipca.Prior;
  botoes_cidade;
end;

procedure TFrm_ipca.ativacidade;
Begin
  DBECidNome.ReadOnly := False;
  DBEdit1.ReadOnly    := False;
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

procedure TFrm_ipca.desativacidade;
Begin
  DBECidNome.ReadOnly := true;
  DBEdit1.ReadOnly    := true;
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

procedure TFrm_ipca.botoes_cidade;
Begin
  BTPrimeiro.Enabled := True;
  BTAnterior.Enabled := True;
  BTProximo.Enabled := True;
  BtUltimo.Enabled := True;
  if DM_Tabelas.ZQipca.Recno<=1 Then Begin
    BTPrimeiro.Enabled := False;
    BTAnterior.Enabled := False;
  end;
  if DM_Tabelas.ZQipca.RecNo=DM_Tabelas.ZQipca.RecordCount Then Begin
    BTProximo.Enabled := False;
    BtUltimo.Enabled := False;
  end;
  LReg.Caption := inttostr(DM_Tabelas.ZQipca.RecordCount);
End;

procedure TFrm_ipca.BTProximoClick(Sender: TObject);
begin
  DM_Tabelas.ZQipca.Next;;
  botoes_cidade;
end;

procedure TFrm_ipca.BtUltimoClick(Sender: TObject);
begin
  DM_Tabelas.ZQipca.Last;
  botoes_cidade;
end;

procedure TFrm_ipca.BtGravarClick(Sender: TObject);
begin
  if not Verif_senha('IPCA','Gravar inclusão ou edição','Taxa Mês: '+DM_Tabelas.ZQIPCA.FieldByName('percentual').Text) Then Exit;
  DM_Tabelas.ZQipca.Post;
  DM_Tabelas.ZQipca.Refresh;
  desativacidade;
  DBGCid.Refresh;
  DBGCid.SetFocus;
end;

procedure TFrm_ipca.BtCancelarClick(Sender: TObject);
begin
  if not Verif_senha('IPCA','Cancelar inclusão ou edição','Taxa Mês: '+DM_Tabelas.ZQIPCA.FieldByName('percentual').Text) Then Exit;
  DM_Tabelas.ZQipca.Cancel;
  desativacidade;
  DBGCid.SetFocus;
end;

procedure TFrm_ipca.BtIncluirClick(Sender: TObject);
begin
  if not Verif_senha('IPCA','Inclusão','') Then Exit;
  DM_Tabelas.ZQipca.Insert;
  ativacidade;
  DBECidNome.SetFocus;

end;

procedure TFrm_ipca.BtEditarClick(Sender: TObject);
begin
  if not Verif_senha('IPCA','Edição','Código: '+DM_Tabelas.ZQIPCA.FieldByName('idipca').Text+' - '+DM_Tabelas.ZQIPCA.FieldByName('percentual').Text) Then Exit;
  DM_Tabelas.ZQipca.Edit;
  ativacidade;
  DBECidNome.SetFocus;

end;

procedure TFrm_ipca.BtExcluirClick(Sender: TObject);
begin
  if not Verif_senha('IPCA','Exclusão','Código: '+DM_Tabelas.ZQIPCA.FieldByName('idipca').Text+' - '+DM_Tabelas.ZQIPCA.FieldByName('percentual').Text) Then Exit;
  if DM_Tabelas.ZQipca.RecordCount>0 Then
    DM_Tabelas.ZQipca.Delete;

end;

procedure TFrm_ipca.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrm_ipca.BtFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrm_ipca.FormShow(Sender: TObject);
begin
  DM_Tabelas.ZQipca.Open;
  desativacidade;
  DBGCid.SetFocus;
end;

procedure TFrm_ipca.DBGCidKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  botoes_cidade;
end;

procedure TFrm_ipca.DBGCidMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  botoes_cidade;
end;

procedure TFrm_ipca.DBGCidEnter(Sender: TObject);
begin
  botoes_cidade;
end;

procedure TFrm_ipca.DBGCidDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBECodCid.Text = DM_Tabelas.ZQIPCA.FieldByName('idipca').Text Then Begin
    DBGCid.Canvas.Brush.Color :=$006CFFFF;
    DBGCid.Canvas.Font.Color := $00A80000;
    DBGCid.Canvas.Font.Style := [FsBold];
  end;
  DBGCid.DefaultDrawDataCell(Rect, DBGCid.columns[datacol].field, State);

end;

procedure TFrm_ipca.DBGCidTitleClick(Column: TColumn);
var
  campo : string;
  varcodigo : integer;
begin
  campo:=column.fieldname;
  DBGCid.Columns[0].Color := clWindow;
  DBGCid.Columns[1].Color := clWindow;
  DBGCid.Columns[2].Color := clWindow;
  application.processmessages;
  varcodigo := DM_Tabelas.ZQIPCA.FieldByName('idipca').AsLargeInt;
  DM_Tabelas.ZQipca.sql.clear;
  DM_Tabelas.ZQipca.SQL.Add('select  idipca,Mes_ano,percentual from ipca');
  DM_Tabelas.ZQipca.SQL.Add(' order by '+campo);
  DM_Tabelas.ZQipca.Open;
  DM_Tabelas.ZQipca.Locate('idipca',varcodigo,[]);
  column.color:=$00EDF0F1;
  botoes_cidade;
end;

procedure TFrm_ipca.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  DM_Tabelas.ZQipca.close;
  Frm_ipca:=nil;
  Action:=CaFree;
end;

procedure TFrm_ipca.BtRelatorioClick(Sender: TObject);
begin
  Frm_ipca.FormStyle:=fsNormal;
  if Frm_Relipca=nil then
    Frm_Relipca:=TFrm_Relipca.Create(Application);
  Frm_Relipca.RLReport1.PreviewModal;
  Frm_Relipca:=nil;
  Frm_ipca.FormStyle:=fsStayOnTop;
end;


procedure TFrm_ipca.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
