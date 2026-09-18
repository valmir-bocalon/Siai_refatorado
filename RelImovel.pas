unit RelImovel;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, XBanner, TFlatCheckBoxUnit, CheckLst,
  DB, DBClient, Grids, DBGrids, ZAbstractRODataset, ZAbstractDataset,
  ZDataset, Vcl.ComCtrls;

type
  TFrm_RelImovel = class(TForm)
    XBanner10: TXBanner;
    Label2: TLabel;
    GroupBox1: TGroupBox;
    FCBS: TFlatCheckBox;
    FCBN: TFlatCheckBox;
    XBanner1: TXBanner;
    BTFechar: TdxButtonArround;
    BtRelatorio: TdxButtonArround;
    GroupBox2: TGroupBox;
    XBanner3: TXBanner;
    XBanner4: TXBanner;
    DXBMarcEmpree: TdxButtonArround;
    DS_Empreend: TDataSource;
    CDSEmpreend: TClientDataSet;
    CDSEmpreendnomeempre: TStringField;
    CDSEmpreendcodempre: TIntegerField;
    CDSEmpreendmarca: TIntegerField;
    DBGEmpre: TDBGrid;
    DS_Imov: TDataSource;
    ZQImov: TZQuery;




    CDSEmpreendlogo: TBlobField;
    CDSEmpreendapelido: TStringField;




















    ProgressBar1: TProgressBar;
    CBObs: TCheckBox;
    procedure BTFecharClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure DXBMarcEmpreeClick(Sender: TObject);
    procedure DBGEmpreDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGEmpreKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGEmpreMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure mudamarcaEmpreend;
    procedure BtRelatorioClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_RelImovel: TFrm_RelImovel;

implementation

uses tabelas, RelImovel01, RelImovel02, uRuntimeFields;

{$R *.dfm}



procedure TFrm_RelImovel.BTFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrm_RelImovel.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  CDSEmpreend.Close;
  ZQImov.Close;
  Frm_RelImovel:=nil;
  Action:=CaFree;
end;

procedure TFrm_RelImovel.FormShow(Sender: TObject);
begin
  DM_Tabelas.ZQLoteamento.First;
  CDSEmpreend.Close;
  CDSEmpreend.CreateDataSet;
  DM_Tabelas.ZQLoteamento.DisableControls;
  ProgressBar1.Visible:=true;
  ProgressBar1.Max:=DM_Tabelas.ZQLoteamento.RecordCount;
  ProgressBar1.Position:=0;
  while not DM_Tabelas.ZQLoteamento.Eof do begin
    ProgressBar1.Position:=DM_Tabelas.ZQLoteamento.RecNo;
    CDSEmpreend.Insert;
    CDSEmpreendmarca.Value := 1;
    CDSEmpreendcodempre.Value := DM_Tabelas.ZQLoteamento.FieldByName('idloteamento').AsLargeInt;
//    CDSEmpreendnomeempre.Value := DM_TAbelas.ZQLoteamentonomeloteamento.Value;
    CDSEmpreendnomeempre.Value := DM_TAbelas.ZQLoteamento.FieldByName('apelido').AsString;
    CDSEmpreendlogo.Assign(DM_Tabelas.ZQLoteamento.FieldByName('logotipo'));
    CDSEmpreend.Post;
    DM_Tabelas.ZQLoteamento.Next;
  end;
  DM_Tabelas.ZQLoteamento.EnableControls;
  ProgressBar1.Position:=0;
  ProgressBar1.Visible:=false;
end;

procedure TFrm_RelImovel.DXBMarcEmpreeClick(Sender: TObject);
begin
  if DXBMarcEmpree.Caption = 'Desmarcar' Then Begin
    CDSEmpreend.First;
    CDSEmpreend.DisableControls;
    ProgressBar1.Visible:=true;
    ProgressBar1.Max:=CDSEmpreend.RecordCount;
    ProgressBar1.Position:=0;
    while not CDSEmpreend.Eof do begin
      ProgressBar1.Position:=CDSEmpreend.RecNo;
      CDSEmpreend.Edit;
      CDSEmpreendmarca.Value := 1;
      CDSEmpreend.Post;
      CDSEmpreend.Next;
    end;
    DXBMarcEmpree.Caption := 'Marcar';
    CDSEmpreend.EnableControls;
    ProgressBar1.Position:=0;
    ProgressBar1.Visible:=false;
  end
  Else
  Begin
    CDSEmpreend.First;
    CDSEmpreend.DisableControls;
    ProgressBar1.Visible:=true;
    ProgressBar1.Max:=CDSEmpreend.RecordCount;
    ProgressBar1.Position:=0;

    while not CDSEmpreend.Eof do begin
      ProgressBar1.Position:=CDSEmpreend.RecNo;
      CDSEmpreend.Edit;
      CDSEmpreendmarca.Value := 0;
      CDSEmpreend.Post;
      CDSEmpreend.Next;
    end;
    CDSEmpreend.EnableControls;
    DXBMarcEmpree.Caption := 'Desmarcar';
    ProgressBar1.Position:=0;
    ProgressBar1.Visible:=false;

  End;
end;

procedure TFrm_RelImovel.DBGEmpreDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  DBGEmpre.DefaultDrawDataCell(Rect, DBGEmpre.columns[datacol].field, State);
  if Column.Field = CDSEmpreendmarca then begin
    DBGEmpre.Canvas.FillRect(Rect);
    if CDSEmpreendmarca.Value = 0 then
      DM_Tabelas.ImageList1.Draw(DBGEmpre.Canvas, Rect.Left + 15,Rect.Top + 1, 0)
    else
      DM_Tabelas.ImageList1.Draw(DBGEmpre.Canvas, Rect.Left + 15,Rect.Top + 1, 1);
  end;
end;

procedure TFrm_RelImovel.DBGEmpreKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = 13 Then
    mudamarcaEmpreend;
end;

procedure TFrm_RelImovel.DBGEmpreMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  mudamarcaEmpreend;
end;

procedure TFrm_RelImovel.mudamarcaEmpreend;
Begin
  if DBGEmpre.SelectedField.FieldName = 'marca' Then Begin
    CDSEmpreend.Edit;
    if CDSEmpreendmarca.Value = 0 Then
      CDSEmpreendmarca.Value := 1
    else
      CDSEmpreendmarca.Value := 0;
    CDSEmpreend.Post;
  end;
End;

procedure TFrm_RelImovel.BtRelatorioClick(Sender: TObject);
begin
  Frm_RelImovel.FormStyle:=fsNormal;
  CDSEmpreend.Filtered := false;
  CDSEmpreend.Filter := 'marca=0';
  CDSEmpreend.Filtered := True;
  ZQImov.Open;
  ZQImov.Filtered := False;
  if FCBS.Checked<> FCBN.Checked Then begin
    if FCBS.Checked Then
      ZQImov.Filter := 'disponivel='+quotedstr('SIM')
    else
      ZQImov.Filter := 'disponivel='+quotedstr('NºO');
    ZQImov.Filtered := True;
  end;
  if CBObs.Checked=false then
  begin
    if Frm_RelImovel01=nil then
       Frm_RelImovel01:=TFrm_RelImovel01.Create(Application);

    Frm_RelImovel01.XConta.Value := 0;
    ZQImov.RecordCount;
    Frm_RelImovel01.RLReport1.PreviewModal;
    FreeAndNil(Frm_RelImovel01);
  end
  else
  begin
    if Frm_RelImovel02=nil then
       Frm_RelImovel02:=TFrm_RelImovel02.Create(Application);

    Frm_RelImovel02.XConta.Value := 0;
    ZQImov.RecordCount;
    Frm_RelImovel02.RLReport1.PreviewModal;
    FreeAndNil(Frm_RelImovel02);
  end;
  CDSEmpreend.Filtered := False;
  Frm_RelImovel.FormStyle:=fsStayOnTop;
end;


procedure TFrm_RelImovel.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_RelImovel, 'ZQImov', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovarea', 'area', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovvalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovquadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovlote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovlado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovmedidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovmedidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovmedidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovmedidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovconfrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovconfrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovconfrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovconfrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovesquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovesquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovesquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovObervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovcornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovdisponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovmatri', 'matri', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovmatricula', 'matricula', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelImovel, 'ZQImov', 'ZQImovproposta', 'proposta', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
