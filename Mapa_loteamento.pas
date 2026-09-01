unit Mapa_loteamento;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport;

type
  TFrm_Mapa_loteamento = class(TForm)
    RLReport1: TRLReport;
    RLImage1: TRLImage;
    RLDBImage1: TRLDBImage;
    RLLabel1: TRLLabel;
    RLDBText1: TRLDBText;
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_Mapa_loteamento: TFrm_Mapa_loteamento;

implementation

uses tabelas, uRuntimeFields;

{$R *.dfm}


procedure TFrm_Mapa_loteamento.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
