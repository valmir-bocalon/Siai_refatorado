unit RelREMESSA;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, XBanner, TFlatCheckBoxUnit, CheckLst,
  DB, DBClient, Grids, DBGrids, ZAbstractRODataset, ZAbstractDataset,
  ZDataset, ExtCtrls;

type
  TFrm_RelREMESSA = class(TForm)
    XBanner10: TXBanner;
    Label2: TLabel;
    BTFechar: TdxButtonArround;
    BtRelatorio: TdxButtonArround;
    GroupBox2: TGroupBox;
    XBanner3: TXBanner;
    XBanner4: TXBanner;
    DXBMarcEmpree: TdxButtonArround;
    DS_Remess: TDataSource;
    CDSRemess: TClientDataSet;
    DBGEmpre: TDBGrid;
    CDSRemessmarca: TIntegerField;
    CDSRemessidbanco_remessa: TLargeintField;
    CDSRemessremessa: TIntegerField;
    CDSRemessidbanco: TIntegerField;
    CDSRemessgerado: TDateField;
    CDSRemessnomearq: TStringField;
    CDSRemessarquivo: TStringField;
    dxButton1: TdxButtonArround;
    Timer1: TTimer;
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
    procedure dxButton1Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_RelREMESSA: TFrm_RelREMESSA;

implementation

uses tabelas, RelImovel01, RelImovelREMESSA, UFrmContasBancarias,
  RelImovelREMESSA2, uRuntimeFields;

{$R *.dfm}

procedure TFrm_RelREMESSA.BTFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrm_RelREMESSA.FormShow(Sender: TObject);
begin
  DM_Tabelas.ZQBancRemes.First;
  CDSRemess.Close;
  CDSRemess.CreateDataSet;
  DM_Tabelas.ZQBancRemes.DisableControls;
  while not DM_Tabelas.ZQBancRemes.Eof do
  begin
  if DM_Tabelas.ZQBancRemes.FieldByName('gerado').Text+DM_Tabelas.ZQBancRemes.FieldByName('nomearq').AsString=FrmCad_ContasBancarias.numero.Caption then
    begin
      CDSRemess.Insert;
      CDSRemessmarca.Value := 0;
      CDSRemessremessa.Value := DM_Tabelas.ZQBancRemes.FieldByName('remessa').AsInteger;
      CDSRemessgerado.Value := DM_Tabelas.ZQBancRemes.FieldByName('gerado').AsDateTime;
      CDSRemessidbanco_remessa.Value:=DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
      CDSRemessidbanco.Value:=DM_Tabelas.ZQBancRemes.FieldByName('idbanco').AsInteger;
      CDSRemessnomearq.Value:=DM_Tabelas.ZQBancRemes.FieldByName('nomearq').AsString;
      CDSRemess.Post;
    end;
{    else
       CDSRemessmarca.Value := 1;
    CDSRemessremessa.Value := DM_Tabelas.ZQBancRemesremessa.Value;
    CDSRemessgerado.Value := DM_Tabelas.ZQBancRemesgerado.Value;
    CDSRemessidbanco_remessa.Value:=DM_Tabelas.ZQBancRemesidbanco_remessa.Value;
    CDSRemessidbanco.Value:=DM_Tabelas.ZQBancRemesidbanco.Value;
    CDSRemessnomearq.Value:=DM_Tabelas.ZQBancRemesnomearq.Value;
    CDSRemess.Post;}
    DM_Tabelas.ZQBancRemes.Next;
  end;

  FrmCad_ContasBancarias.numero.Caption:='';
  DM_Tabelas.ZQRemes_Receb2.open;
  Timer1.enabled:=true;
end;

procedure TFrm_RelREMESSA.DXBMarcEmpreeClick(Sender: TObject);
begin
  if DXBMarcEmpree.Caption = 'Desmarcar' Then Begin
    CDSRemess.First;
    while not CDSRemess.Eof do begin
      CDSRemess.Edit;
      CDSRemessmarca.Value := 1;
      CDSRemess.Post;
      CDSRemess.Next;
    end;
    DXBMarcEmpree.Caption := 'Marcar';
  end
  Else
  Begin
    CDSRemess.First;
    while not CDSRemess.Eof do begin
      CDSRemess.Edit;
      CDSRemessmarca.Value := 0;
      CDSRemess.Post;
      CDSRemess.Next;
    end;
    DXBMarcEmpree.Caption := 'Desmarcar';
  End;
end;

procedure TFrm_RelREMESSA.DBGEmpreDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  DBGEmpre.DefaultDrawDataCell(Rect, DBGEmpre.columns[datacol].field, State);
  if Column.Field = CDSRemessmarca then begin
    DBGEmpre.Canvas.FillRect(Rect);
    if CDSRemessmarca.Value = 0 then
      DM_Tabelas.ImageList1.Draw(DBGEmpre.Canvas, Rect.Left + 15,Rect.Top + 1, 0)
    else
      DM_Tabelas.ImageList1.Draw(DBGEmpre.Canvas, Rect.Left + 15,Rect.Top + 1, 1);
  end;
end;

procedure TFrm_RelREMESSA.DBGEmpreKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = 13 Then
    mudamarcaEmpreend;
end;

procedure TFrm_RelREMESSA.DBGEmpreMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  mudamarcaEmpreend;
end;

procedure TFrm_RelREMESSA.mudamarcaEmpreend;
Begin
  if DBGEmpre.SelectedField.FieldName = 'marca' Then Begin
    CDSRemess.Edit;
    if CDSRemessmarca.Value = 0 Then
      CDSRemessmarca.Value := 1
    else
      CDSRemessmarca.Value := 0;
    CDSRemess.Post;
  end;
End;

procedure TFrm_RelREMESSA.BtRelatorioClick(Sender: TObject);
begin
  CDSRemess.Filtered:=false;
  CDSRemess.Filter := 'marca=0';
  CDSRemess.Filtered := True;
  CDSRemess.RecordCount;

  DM_Tabelas.ZQRemes_Receb2.MasterSource:=DS_Remess;
  DM_Tabelas.ZQRemes_Receb2.MasterFields:='idbanco_remessa';
  DM_Tabelas.ZQRemes_Receb2.LinkedFields:='remessa';
 
  if RLReportREMESSA=nil then
     RLReportREMESSA:=TRLReportREMESSA.Create(Application);
  RLReportREMESSA.RLLabel2.Caption:=CDSRemessnomearq.Value;
  RLReportREMESSA.RLReportREMESSA.PreviewModal;
  RLReportREMESSA:=nil;
  CDSRemess.Filtered := False;
end;

procedure TFrm_RelREMESSA.dxButton1Click(Sender: TObject);
begin
  CDSRemess.Filtered:=false;
  CDSRemess.Filter := 'marca=0';
  CDSRemess.Filtered := True;


  DM_Tabelas.ZQRemes_Receb2.MasterSource:=DS_Remess;
  DM_Tabelas.ZQRemes_Receb2.MasterFields:='idbanco_remessa';
  DM_Tabelas.ZQRemes_Receb2.LinkedFields:='remessa';

  if RLReportREMESSA2=nil then
     RLReportREMESSA2:=TRLReportREMESSA2.Create(Application);
  RLReportREMESSA2.RLLabel2.Caption:=CDSRemessnomearq.Value;
  RLReportREMESSA2.RLReportREMESSA2.PreviewModal;
  RLReportREMESSA2:=nil;
  CDSRemess.Filtered := False;

end;

procedure TFrm_RelREMESSA.Timer1Timer(Sender: TObject);
var
posi:integer;
ext:string;
begin
  //simula pressionamento dos botões
  posi:=pos('.',CDSRemessnomearq.Value);
  inc(posi);
  ext:=uppercase(copy(CDSRemessnomearq.Value,posi,3));
  if (ext='REM') or (ext='TST') then
  begin
    if assigned(BtRelatorio.OnClick) then
    begin
      Timer1.enabled:=false;
      BtRelatorio.OnClick(BtRelatorio);
    end;  
  end
  else
  begin
    if assigned(dxButton1.OnClick) then
    begin
      Timer1.enabled:=false;
      dxButton1.OnClick(dxButton1);
    end;
  end;

end;

procedure TFrm_RelREMESSA.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  DM_Tabelas.ZQBancRemes.EnableControls;
  Frm_RelREMESSA:=nil;
  Action:=CaFree;
end;


procedure TFrm_RelREMESSA.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
