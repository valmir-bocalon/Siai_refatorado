unit TrocaCPF_RG;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Gauges;

type
  TFrm_TrocaCPF_RG = class(TForm)
    Gauge1: TGauge;
    Button1: TButton;
    Button2: TButton;
    procedure Button2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_TrocaCPF_RG: TFrm_TrocaCPF_RG;

implementation

uses tabelas, uRuntimeFields;

{$R *.dfm}

procedure TFrm_TrocaCPF_RG.Button2Click(Sender: TObject);
begin
  Close;
end;

procedure TFrm_TrocaCPF_RG.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Frm_TrocaCPF_RG:=nil;
  Action:=CaFree;
end;

procedure TFrm_TrocaCPF_RG.Button1Click(Sender: TObject);
var
  vardoc1, vardoc2 : string;
begin
  DM_Tabelas.ZQresponsavel.First;
  Gauge1.MaxValue := DM_Tabelas.ZQresponsavel.RecordCount;
  while not DM_Tabelas.ZQresponsavel.Eof do Begin
    Gauge1.Progress := DM_Tabelas.ZQresponsavel.RecNo;
    vardoc1 := DM_Tabelas.ZQresponsavel.FieldByName('documen1').AsString;
    vardoc2 := DM_Tabelas.ZQresponsavel.FieldByName('documen2').AsString;
    if length(vardoc1) > length(vardoc2) Then Begin
      DM_Tabelas.ZQresponsavel.Edit;
      DM_Tabelas.ZQresponsavel.FieldByName('documen1').AsString := vardoc2;
      DM_Tabelas.ZQresponsavel.FieldByName('documen2').AsString := vardoc1;
      DM_Tabelas.ZQresponsavel.Post;
    end;
    DM_TAbelas.ZQresponsavel.Next;
  end;
  DM_Tabelas.ZqParticipante.First;
  Gauge1.MaxValue := DM_tabelas.ZqParticipante.RecordCount;
  while not DM_Tabelas.ZqParticipante.Eof do begin
    Gauge1.Progress := DM_TAbelas.ZqParticipante.RecNo;
    DM_Tabelas.ZqParticipante.Edit;
    if copy(DM_TAbelas.ZqParticipante.FieldByName('endereco').AsString,1,4)='Rua.' Then
      DM_Tabelas.ZqParticipante.FieldByName('endereco').AsString := 'RUA'+uppercase(copy(DM_TAbelas.ZqParticipante.FieldByName('endereco').AsString,5,200));
    if copy(DM_TAbelas.ZqParticipante.FieldByName('endereco').AsString,1,4)='Ave.' Then
      DM_Tabelas.ZqParticipante.FieldByName('endereco').AsString := 'AV'+uppercase(copy(DM_TAbelas.ZqParticipante.FieldByName('endereco').AsString,5,200));
    if copy(DM_TAbelas.ZqParticipante.FieldByName('ende_cob').AsString,1,4)='Rua.' Then
      DM_Tabelas.ZqParticipante.FieldByName('ende_cob').AsString := 'RUA'+uppercase(copy(DM_TAbelas.ZqParticipante.FieldByName('ende_cob').AsString,5,200));
    if copy(DM_TAbelas.ZqParticipante.FieldByName('ende_cob').AsString,1,4)='Ave.' Then
      DM_Tabelas.ZqParticipante.FieldByName('ende_cob').AsString := 'AV'+uppercase(copy(DM_TAbelas.ZqParticipante.FieldByName('ende_cob').AsString,5,200));
    DM_Tabelas.ZqParticipante.Post;
    DM_Tabelas.ZqParticipante.Next;
  end;
end;


procedure TFrm_TrocaCPF_RG.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
