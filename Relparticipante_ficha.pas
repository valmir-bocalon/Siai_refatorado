unit Relparticipante_ficha;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, RLReport;

type
  TFrmRelparticipante_ficha = class(TForm)
  published
    RLReport1: TRLReport;
    RLBand1: TRLBand;
    RLBand2: TRLBand;
    RLLabelParticipante: TRLLabel;
    RLLabelConjugeSecao: TRLLabel;
    RLDrawSeparador: TRLDraw;
    RLLabelNome: TRLLabel;
    RLLabelRG: TRLLabel;
    RLLabelCPF: TRLLabel;
    RLLabelNacionalidade: TRLLabel;
    RLLabelNaturalidade: TRLLabel;
    RLLabelOcupacao: TRLLabel;
    RLLabelEndereco: TRLLabel;
    RLLabelBairro: TRLLabel;
    RLLabelCidade: TRLLabel;
    RLLabelCEP: TRLLabel;
    RLLabelTelCel: TRLLabel;
    RLDBNome: TRLDBText;
    RLDBRG: TRLDBText;
    RLDBCPF: TRLDBText;
    RLDBNacionalidade: TRLDBText;
    RLDBNaturalidade: TRLDBText;
    RLDBNatUF: TRLDBText;
    RLLabelNatBarra: TRLLabel;
    RLDBOcupacao: TRLDBText;
    RLDBEndereco: TRLDBText;
    RLDBBairro: TRLDBText;
    RLDBCidade: TRLDBText;
    RLDBCidadeUF: TRLDBText;
    RLLabelCidadeBarra: TRLLabel;
    RLDBCEP: TRLDBText;
    RLDBTelCel: TRLDBText;
    RLLabelEstadoCivil: TRLLabel;
    RLLabelRegime: TRLLabel;
    RLLabelConjCidade: TRLLabel;
    RLLabelConjNome: TRLLabel;
    RLLabelConjNaturalidade: TRLLabel;
    RLLabelConjCPF: TRLLabel;
    RLLabelConjRG: TRLLabel;
    RLDBEstadoCivil: TRLDBText;
    RLDBRegime: TRLDBText;
    RLDBConjCidade: TRLDBText;
    RLDBConjCidadeUF: TRLDBText;
    RLLabelConjCidadeBarra: TRLLabel;
    RLDBConjNome: TRLDBText;
    RLDBConjNaturalidade: TRLDBText;
    RLDBConjNaturalidadeUF: TRLDBText;
    RLLabelConjNaturalidadeBarra: TRLLabel;
    RLDBConjCPF: TRLDBText;
    RLDBConjRG: TRLDBText;
    RLLabel3: TRLLabel;
    RLLabel1: TRLLabel;
    RLLabel4: TRLLabel;
    RLSystemInfo4: TRLSystemInfo;
    RLLabel7: TRLLabel;
    RLSystemInfo3: TRLSystemInfo;
    RLSystemInfo2: TRLSystemInfo;
    RLLabel6: TRLLabel;
    RLLabel5: TRLLabel;
    RLSystemInfo1: TRLSystemInfo;
    RLDraw1: TRLDraw;
    RLLabel8: TRLLabel;
    procedure RLReport1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand2BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
  public
    constructor Create(AOwner: TComponent); override;
    procedure AfterConstruction; override;
  end;

var
  FrmRelparticipante_ficha: TFrmRelparticipante_ficha;

implementation

uses
  Tabelas, Participante, principal, uRuntimeFields, DB;

{$R *.dfm}

function ReportConjugeLookup(const AKeyField, AResultField: string): string;
var
  LField: TField;
  LValue: Variant;
begin
  Result := '';
  if (Frm_Participante = nil) or (Frm_Participante.ZQConjuge = nil) or
     (not Frm_Participante.ZQConjuge.Active) or
     Frm_Participante.ZQConjuge.IsEmpty or (DM_Tabelas = nil) or
     (DM_Tabelas.ZQCidade = nil) or (not DM_Tabelas.ZQCidade.Active) then
    Exit;
  LField := Frm_Participante.ZQConjuge.FindField(AKeyField);
  if (LField = nil) or LField.IsNull then
    Exit;
  LValue := DM_Tabelas.ZQCidade.Lookup('idcidade', LField.Value, AResultField);
  if not VarIsNull(LValue) and not VarIsEmpty(LValue) then
    Result := VarToStr(LValue);
end;

procedure EnsureReportConjugeFields(AForm: TFrm_Participante);
var
  LWasActive: Boolean;

  procedure EnsureTextField(const AComponentName, AFieldName: string;
    ASize: Integer);
  var
    LField: TWideStringField;
  begin
    if AForm.ZQConjuge.FindField(AFieldName) <> nil then
      Exit;
    LField := TWideStringField.Create(AForm);
    try
      LField.Name := AComponentName;
      LField.FieldName := AFieldName;
      LField.FieldKind := fkCalculated;
      LField.Size := ASize;
      LField.DataSet := AForm.ZQConjuge;
    except
      LField.Free;
      raise;
    end;
  end;
begin
  if (AForm = nil) or (AForm.ZQConjuge = nil) then
    Exit;
  if (AForm.ZQConjuge.FindField('nomecida') <> nil) and
     (AForm.ZQConjuge.FindField('estcida') <> nil) and
     (AForm.ZQConjuge.FindField('comar_cidnome') <> nil) and
     (AForm.ZQConjuge.FindField('comar_cidest') <> nil) and
     (AForm.ZQConjuge.FindField('cart_cidnome') <> nil) and
     (AForm.ZQConjuge.FindField('cart_cidest') <> nil) then
    Exit;

  LWasActive := AForm.ZQConjuge.Active;
  if LWasActive then
    AForm.ZQConjuge.Close;
  try
    EnsureTextField('ZQConjugenomecida', 'nomecida', 100);
    EnsureTextField('ZQConjugeestcida', 'estcida', 2);
    EnsureTextField('ZQConjugecomar_cidnome', 'comar_cidnome', 100);
    EnsureTextField('ZQConjugecomar_cidest', 'comar_cidest', 2);
    EnsureTextField('ZQConjugecart_cidnome', 'cart_cidnome', 100);
    EnsureTextField('ZQConjugecart_cidest', 'cart_cidest', 2);
  finally
    if LWasActive then
      AForm.ZQConjuge.Open;
  end;
end;

constructor TFrmRelparticipante_ficha.Create(AOwner: TComponent);
begin
  { O DFM possui campos calculados ligados ao dataset local de conjuge.
    Crie o formulario dono do dataset antes de carregar o DFM do relatorio. }
  if Frm_Participante = nil then
    Frm_Participante := TFrm_Participante.Create(Application);
  inherited Create(AOwner);
  EnsureReportConjugeFields(Frm_Participante);
end;

procedure TFrmRelparticipante_ficha.RLReport1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  if Frm_Participante = nil then
    Frm_Participante := TFrm_Participante.Create(Application);

  RLReport1.DataSource := Frm_Participante.DS_Participantes;

  if DM_Tabelas.ZQCidade.Active then
    DM_Tabelas.ZQCidade.Close;
  DM_Tabelas.ZQCidade.Open;

  if Frm_Participante.ZQConjuge.Active then
    Frm_Participante.ZQConjuge.Close;
  Frm_Participante.ZQConjuge.Open;
  if Frm_Participante.ZQConjuge.Active then
    Frm_Participante.ZQConjuge.First;
  RLLabelConjugeSecao.Visible := Frm_Participante.ZQConjuge.Active and
    (Frm_Participante.ZQConjuge.RecordCount > 0);
end;

procedure TFrmRelparticipante_ficha.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
 RLLabel8.Caption:='Usuario:'+Frm_principal.xusuario.Caption;
end;

procedure TFrmRelparticipante_ficha.RLBand2BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
var
  SpouseExists: Boolean;
begin
  if Frm_Participante.ZQConjuge.Active then
    Frm_Participante.ZQConjuge.First;

  SpouseExists := Frm_Participante.ZQConjuge.Active and (Frm_Participante.ZQConjuge.RecordCount > 0);

  RLLabelRegime.Visible                := SpouseExists;
  RLDBRegime.Visible                   := SpouseExists;
 // RLLabelConjCidade.Visible            := SpouseExists;
//  RLDBConjCidade.Visible               := SpouseExists;
//  RLDBConjCidadeUF.Visible             := SpouseExists;
 // RLLabelConjCidadeBarra.Visible       := SpouseExists;
  RLLabelConjNome.Visible              := SpouseExists;
  RLDBConjNome.Visible                 := SpouseExists;
  RLLabelConjNaturalidade.Visible      := SpouseExists;
  RLDBConjNaturalidade.Visible         := SpouseExists;
  RLDBConjNaturalidadeUF.Visible       := SpouseExists;
  RLLabelConjNaturalidadeBarra.Visible := SpouseExists;
  RLLabelConjCPF.Visible               := SpouseExists;
  RLDBConjCPF.Visible                  := SpouseExists;
  RLLabelConjRG.Visible                := SpouseExists;
  RLDBConjRG.Visible                   := SpouseExists;

  { Estes componentes eram ligados a campos calculados no DFM. O texto e
    preenchido aqui para evitar que o Fortes tente resolver o campo antes do
    dataset local receber seus campos runtime. }
  RLDBConjCidade.DataField          := '';
  RLDBConjCidadeUF.DataField        := '';
  RLDBConjNaturalidade.DataField    := '';
  RLDBConjNaturalidadeUF.DataField  := '';
  if SpouseExists then
  begin
    RLDBConjCidade.Caption         := ReportConjugeLookup('comarca', 'nomecid');
    RLDBConjCidadeUF.Caption       := ReportConjugeLookup('comarca', 'estado');
    RLDBConjNaturalidade.Caption   := ReportConjugeLookup('natural', 'nomecid');
    RLDBConjNaturalidadeUF.Caption := ReportConjugeLookup('natural', 'estado');
  end
  else
  begin
    RLDBConjCidade.Caption         := '';
    RLDBConjCidadeUF.Caption       := '';
    RLDBConjNaturalidade.Caption   := '';
    RLDBConjNaturalidadeUF.Caption := '';
  end;
end;


procedure TFrmRelparticipante_ficha.AfterConstruction;
begin
  inherited AfterConstruction;
  RLReport1.BeforePrint := RLReport1BeforePrint;
  EnsureRuntimeFields(Self);
end;

initialization
  if Classes.GetClass('TRLReport') = nil then
    Classes.RegisterClass(TRLReport);
  if Classes.GetClass('TRLBand') = nil then
    Classes.RegisterClass(TRLBand);
  if Classes.GetClass('TRLLabel') = nil then
    Classes.RegisterClass(TRLLabel);
  if Classes.GetClass('TRLDraw') = nil then
    Classes.RegisterClass(TRLDraw);
  if Classes.GetClass('TRLDBText') = nil then
    Classes.RegisterClass(TRLDBText);
  if Classes.GetClass('TRLSystemInfo') = nil then
    Classes.RegisterClass(TRLSystemInfo);

end.
