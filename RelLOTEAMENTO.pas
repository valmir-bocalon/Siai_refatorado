unit RelLOTEAMENTO;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, XBanner, TFlatCheckBoxUnit, CheckLst,
  DB, DBClient, Grids, DBGrids, ZAbstractRODataset, ZAbstractDataset,
  ZDataset, Vcl.ComCtrls;

type
  TFrm_Relloteamento = class(TForm)
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
    CDSEmpreendmatricula: TStringField;





















    ProgressBar1: TProgressBar;
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
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_Relloteamento: TFrm_Relloteamento;

implementation

uses tabelas, RelImovel01, Relloteamento01, uRuntimeFields;

{$R *.dfm}

procedure TFrm_Relloteamento.BTFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrm_Relloteamento.FormShow(Sender: TObject);
begin

  DM_Tabelas.ZQLoteamento.First;
  CDSEmpreend.Close;
  CDSEmpreend.CreateDataSet;
  DM_Tabelas.ZQLoteamento.DisableControls;
  ProgressBar1.Visible:=true;
  ProgressBar1.Position:=0;
  ProgressBar1.Max:=DM_Tabelas.ZQLoteamento.RecordCount;
  while not DM_Tabelas.ZQLoteamento.Eof do begin
    ProgressBar1.Position:=DM_Tabelas.ZQLoteamento.RecNo;
    CDSEmpreend.Insert;
    CDSEmpreendmarca.Value     := 1;
    CDSEmpreendcodempre.Value  := DM_Tabelas.ZQLoteamento.FieldByName('idloteamento').AsLargeInt;
    CDSEmpreendnomeempre.Value := DM_TAbelas.ZQLoteamento.FieldByName('nomeloteamento').AsString;
    CDSEmpreendlogo.Assign(DM_Tabelas.ZQLoteamento.FieldByName('logotipo'));
    CDSEmpreendapelido.Value   := DM_Tabelas.ZQLoteamento.FieldByName('apelido').AsString;
    CDSEmpreendmatricula.Value := DM_Tabelas.ZQLoteamento.FieldByName('matriculaloteamento').AsString;
    CDSEmpreend.Post;
    DM_Tabelas.ZQLoteamento.Next;
  end;
  DM_Tabelas.ZQLoteamento.EnableControls;
  ProgressBar1.Position:=0;
  ProgressBar1.Visible:=false;


end;

procedure TFrm_Relloteamento.DXBMarcEmpreeClick(Sender: TObject);
begin
  if DXBMarcEmpree.Caption = 'Desmarcar' Then Begin
    CDSEmpreend.First;
    ProgressBar1.Visible:=true;
    ProgressBar1.Position:=0;
    ProgressBar1.Max:=CDSEmpreend.RecordCount;
    CDSEmpreend.DisableControls;
    while not CDSEmpreend.Eof do begin
      ProgressBar1.Position:=CDSEmpreend.RecNo;
      CDSEmpreend.Edit;
      CDSEmpreendmarca.Value := 1;
      CDSEmpreend.Post;
      CDSEmpreend.Next;
    end;
    CDSEmpreend.EnableControls;
    DXBMarcEmpree.Caption := 'Marcar';
    ProgressBar1.Position:=0;
    ProgressBar1.Visible:=false;

  end
  Else
  Begin
    CDSEmpreend.First;
    ProgressBar1.Visible:=true;
    ProgressBar1.Position:=0;
    ProgressBar1.Max:=CDSEmpreend.RecordCount;
    CDSEmpreend.DisableControls;
    while not CDSEmpreend.Eof do begin
      ProgressBar1.Position:=CDSEmpreend.RecNo;
      CDSEmpreend.Edit;
      CDSEmpreendmarca.Value := 0;
      CDSEmpreend.Post;
      CDSEmpreend.Next;
    end;
    DXBMarcEmpree.Caption := 'Desmarcar';
    CDSEmpreend.EnableControls;
    ProgressBar1.Position:=0;
    ProgressBar1.Visible:=false;

  End;
end;

procedure TFrm_Relloteamento.DBGEmpreDrawColumnCell(Sender: TObject;
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

procedure TFrm_Relloteamento.DBGEmpreKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = 13 Then
    mudamarcaEmpreend;
end;

procedure TFrm_Relloteamento.DBGEmpreMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  mudamarcaEmpreend;
end;

procedure TFrm_Relloteamento.mudamarcaEmpreend;
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

procedure TFrm_Relloteamento.BtRelatorioClick(Sender: TObject);
begin
  try

      Frm_Relloteamento.FormStyle:=fsNormal;
      CDSEmpreend.Filtered:=false;
      CDSEmpreend.Filter := 'marca=0';
      CDSEmpreend.Filtered := True;
      CDSEmpreend.recordcount;
      ZQImov.Open;
      ZQImov.Filtered := False;
      if FCBS.Checked<> FCBN.Checked Then begin
        if FCBS.Checked Then
          ZQImov.Filter := 'disponivel='+quotedstr('SIM')
        else
          ZQImov.Filter := 'disponivel='+quotedstr('NºO');
        ZQImov.Filtered := True;
      end;
      ZQImov.RecordCount;
      if Frm_Relloteamento01=nil then
         Frm_Relloteamento01:=TFrm_Relloteamento01.Create(Application);
      Frm_Relloteamento01.XConta.Value := 0;
      CDSEmpreend.First;
      Frm_Relloteamento01.RLReport1.PreviewModal;

  finally
   FreeAndNil(Frm_Relloteamento01);
   CDSEmpreend.Filtered := False;
   Frm_Relloteamento.FormStyle:=fsStayOnTop;
  end;

end;


procedure TFrm_Relloteamento.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_Relloteamento, 'ZQImov', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovvalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovquadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovlote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovlado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovmedidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovmedidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovmedidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovmedidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovarea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovconfrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovconfrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovconfrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovconfrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovesquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovesquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovesquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovObervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovcornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovdisponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovmatri', 'matri', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovmatricula', 'matricula', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Relloteamento, 'ZQImov', 'ZQImovproposta', 'proposta', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
