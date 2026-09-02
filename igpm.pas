unit igpm;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, wwdbdatetimepicker, StdCtrls, DBCtrls, Mask, Grids, DBGrids,
    dxButton, ExtCtrls, XBanner, DB, dxCore2;

type
  TFrm_igpm = class(TForm)
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
  Frm_igpm: TFrm_igpm;

implementation

uses tabelas, Funcoes, RelIgpm, uRuntimeFields;

{$R *.dfm}

procedure TFrm_igpm.BTPrimeiroClick(Sender: TObject);
begin
  DM_Tabelas.ZQigpm.First;
  botoes_cidade;
end;

procedure TFrm_igpm.BTAnteriorClick(Sender: TObject);
begin
  DM_Tabelas.ZQigpm.Prior;
  botoes_cidade;
end;

procedure TFrm_igpm.ativacidade;
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

procedure TFrm_igpm.desativacidade;
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

procedure TFrm_igpm.botoes_cidade;
Begin
  BTPrimeiro.Enabled := True;
  BTAnterior.Enabled := True;
  BTProximo.Enabled := True;
  BtUltimo.Enabled := True;
  if DM_Tabelas.ZQigpm.Recno<=1 Then Begin
    BTPrimeiro.Enabled := False;
    BTAnterior.Enabled := False;
  end;
  if DM_Tabelas.ZQigpm.RecNo=DM_Tabelas.ZQigpm.RecordCount Then Begin
    BTProximo.Enabled := False;
    BtUltimo.Enabled := False;
  end;
  LReg.Caption := inttostr(DM_Tabelas.ZQigpm.RecordCount);
End;

procedure TFrm_igpm.BTProximoClick(Sender: TObject);
begin
  DM_Tabelas.ZQigpm.Next;;
  botoes_cidade;
end;

procedure TFrm_igpm.BtUltimoClick(Sender: TObject);
begin
  DM_Tabelas.ZQigpm.Last;
  botoes_cidade;
end;

procedure TFrm_igpm.BtGravarClick(Sender: TObject);
begin
  if not Verif_senha('IGPM','Gravar inclusão ou edição','Taxa Mês: '+DM_Tabelas.ZQigpm.FieldByName('percentual').Text) Then Exit;
  DM_Tabelas.ZQigpm.Post;
  DM_Tabelas.ZQigpm.Refresh;
  desativacidade;
  DBGCid.Refresh;
  DBGCid.SetFocus;
end;

procedure TFrm_igpm.BtCancelarClick(Sender: TObject);
begin
  if not Verif_senha('IGPM','Cancelar inclusão ou edição','Taxa Mês: '+DM_Tabelas.ZQigpm.FieldByName('percentual').Text) Then Exit;
  DM_Tabelas.ZQigpm.Cancel;
  desativacidade;
  DBGCid.SetFocus;
end;

procedure TFrm_igpm.BtIncluirClick(Sender: TObject);
begin
  if not Verif_senha('IGPM','Inclusão','') Then Exit;
  DM_Tabelas.ZQigpm.Insert;
  ativacidade;
  DBECidNome.SetFocus;

end;

procedure TFrm_igpm.BtEditarClick(Sender: TObject);
begin
  if not Verif_senha('IGPM','Edição','Código: '+DM_Tabelas.ZQigpm.FieldByName('idIGPM').Text+' - '+DM_Tabelas.ZQigpm.FieldByName('percentual').Text) Then Exit;
  DM_Tabelas.ZQigpm.Edit;
  ativacidade;
  DBECidNome.SetFocus;

end;

procedure TFrm_igpm.BtExcluirClick(Sender: TObject);
begin
  if not Verif_senha('IGPM','Exclusão','Código: '+DM_Tabelas.ZQigpm.FieldByName('idIGPM').Text+' - '+DM_Tabelas.ZQigpm.FieldByName('percentual').Text) Then Exit;
  if DM_Tabelas.ZQigpm.RecordCount>0 Then
    DM_Tabelas.ZQigpm.Delete;

end;

procedure TFrm_igpm.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrm_igpm.BtFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrm_igpm.FormShow(Sender: TObject);
begin
  DM_Tabelas.ZQigpm.Open;
  desativacidade;
  DBGCid.SetFocus;
end;

procedure TFrm_igpm.DBGCidKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  botoes_cidade;
end;

procedure TFrm_igpm.DBGCidMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  botoes_cidade;
end;

procedure TFrm_igpm.DBGCidEnter(Sender: TObject);
begin
  botoes_cidade;
end;

procedure TFrm_igpm.DBGCidDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBECodCid.Text = DM_Tabelas.ZQigpm.FieldByName('idIGPM').Text Then Begin
    DBGCid.Canvas.Brush.Color :=$006CFFFF;
    DBGCid.Canvas.Font.Color := $00A80000;
    DBGCid.Canvas.Font.Style := [FsBold];
  end;
  DBGCid.DefaultDrawDataCell(Rect, DBGCid.columns[datacol].field, State);

end;

procedure TFrm_igpm.DBGCidTitleClick(Column: TColumn);
var
  campo : string;
  varcodigo : integer;
begin
  campo:=column.fieldname;
  DBGCid.Columns[0].Color := clWindow;
  DBGCid.Columns[1].Color := clWindow;
  DBGCid.Columns[2].Color := clWindow;
  application.processmessages;
  varcodigo := DM_Tabelas.ZQigpm.FieldByName('idIGPM').AsLargeInt;
  DM_Tabelas.ZQigpm.sql.clear;
  DM_Tabelas.ZQigpm.SQL.Add('select  idigpm,Mes_ano,percentual from igpm');
  DM_Tabelas.ZQigpm.SQL.Add(' order by '+campo);
  DM_Tabelas.ZQigpm.Open;
  DM_Tabelas.ZQigpm.Locate('idigpm',varcodigo,[]);
  column.color:=$00EDF0F1;
  botoes_cidade;
end;

procedure TFrm_igpm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  DM_Tabelas.ZQigpm.close;
  Frm_igpm:=nil;
  Action:=CaFree;
end;

procedure TFrm_igpm.BtRelatorioClick(Sender: TObject);
begin
  Frm_igpm.FormStyle:=fsNormal;
  if Frm_Religpm=nil then
    Frm_Religpm:=TFrm_Religpm.Create(Application);
  Frm_Religpm.RLReport1.PreviewModal;
  Frm_Religpm:=nil;
  Frm_igpm.FormStyle:=fsStayOnTop;
end;


procedure TFrm_igpm.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
