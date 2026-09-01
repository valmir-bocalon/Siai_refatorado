unit RelVenda02;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport;

type
  TFrm_RelVenda02 = class(TForm)
    RLReport1: TRLReport;
    RLBand1: TRLBand;
    RLDBImage1: TRLDBImage;
    RLDBText1: TRLDBText;
    RLLabel1: TRLLabel;
    RLLabel2: TRLLabel;
    RLLabel3: TRLLabel;
    RLLabel4: TRLLabel;
    RLSystemInfo1: TRLSystemInfo;
    RLSystemInfo2: TRLSystemInfo;
    RLLabel5: TRLLabel;
    RLLabel6: TRLLabel;
    RLLabel7: TRLLabel;
    RLSystemInfo3: TRLSystemInfo;
    RLLabel8: TRLLabel;
    RLSystemInfo4: TRLSystemInfo;
    RLBand2: TRLBand;
    RLDBText2: TRLDBText;
    RLDBImage2: TRLDBImage;
    RLSubDetail1: TRLSubDetail;
    RLBand3: TRLBand;
    RLDBText3: TRLDBText;
    RLLabel9: TRLLabel;
    RLLabel12: TRLLabel;
    RLLabel13: TRLLabel;
    RLLabel14: TRLLabel;
    RLLabel15: TRLLabel;
    RLDBText4: TRLDBText;
    RLDBText5: TRLDBText;
    RLDBText6: TRLDBText;
    RLDBText7: TRLDBText;
    RLSubDetail2: TRLSubDetail;
    RLBand4: TRLBand;
    RLDBText8: TRLDBText;
    RLBand5: TRLBand;
    RLLabel10: TRLLabel;
    RLDBText9: TRLDBText;
    RLLabel11: TRLLabel;
    RLLabel16: TRLLabel;
    RLLabel17: TRLLabel;
    RLDBText10: TRLDBText;
    RLSubDetail3: TRLSubDetail;
    RLBand6: TRLBand;
    RLLabel18: TRLLabel;
    RLBand7: TRLBand;
    RLDBText11: TRLDBText;
    RLDBText12: TRLDBText;
    RLDBText13: TRLDBText;
    RLDBText14: TRLDBText;
    RLBand9: TRLBand;
    RLLabel19: TRLLabel;
    RLLabel20: TRLLabel;
    RLLabel21: TRLLabel;
    RLLabel22: TRLLabel;
    RLDBResult3: TRLDBResult;
    RLSubDetail4: TRLSubDetail;
    RLDetailGrid1: TRLDetailGrid;
    RLDBText15: TRLDBText;
    RLDBText16: TRLDBText;
    RLDBText17: TRLDBText;
    RLDBText18: TRLDBText;
    RLBand8: TRLBand;
    RLLabel23: TRLLabel;
    RLBand10: TRLBand;
    RLDBResult1: TRLDBResult;
    RLDBText19: TRLDBText;
    RLDBText20: TRLDBText;
    RLLabel24: TRLLabel;
    RLLabel25: TRLLabel;
    RLLabel26: TRLLabel;
    RLLabel27: TRLLabel;
    RLLabel28: TRLLabel;
    RLDBResult2: TRLDBResult;
    RLLabel29: TRLLabel;
    RLDBText21: TRLDBText;
    RLLabel30: TRLLabel;
    RLLabel31: TRLLabel;
    RLDBText22: TRLDBText;
    RLDBText23: TRLDBText;
    RLSubDetail5: TRLSubDetail;
    RLBand11: TRLBand;
    RLDBText24: TRLDBText;
    RLDBText25: TRLDBText;
    RLBand12: TRLBand;
    RLLabel32: TRLLabel;
    RLLabel33: TRLLabel;
    RLLabel34: TRLLabel;
    RLDBText26: TRLDBText;
    RLLabel35: TRLLabel;
    RLLabel36: TRLLabel;
    RLLabel37: TRLLabel;
    RLLabel38: TRLLabel;
    RLDBText27: TRLDBText;
    RLDBText28: TRLDBText;
    RLDBText29: TRLDBText;
    RLDBResult4: TRLDBResult;
    RLDBResult5: TRLDBResult;
    RLDBResult6: TRLDBResult;
    RLDBResult7: TRLDBResult;
    RLLabel39: TRLLabel;
    RLLabel40: TRLLabel;
    RLLabel41: TRLLabel;
    RLLabel42: TRLLabel;
    RLDBResult8: TRLDBResult;
    RLDBResult9: TRLDBResult;
    RLDBResult10: TRLDBResult;
    RLDBText30: TRLDBText;
    RLDBText31: TRLDBText;
    RLDBText32: TRLDBText;
    RLDBText33: TRLDBText;
    procedure RLBand2BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_RelVenda02: TFrm_RelVenda02;

implementation

uses RelVenda, uRuntimeFields;
{$R *.dfm}

procedure TFrm_RelVenda02.RLBand2BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLReport1.NewPage;
end;


procedure TFrm_RelVenda02.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
