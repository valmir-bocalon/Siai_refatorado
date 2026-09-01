unit NumRemessa;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, XBanner, StdCtrls, XNum, DBGrids;

type
  TFrm_NumRemessa = class(TForm)
    XBanner13: TXBanner;
    Label1: TLabel;
    Label2: TLabel;
    XNENumRemessa: TXNumEdit;
    procedure XNENumRemessaKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_NumRemessa: TFrm_NumRemessa;

implementation
uses
  uRuntimeFields;

{$R *.dfm}

procedure TFrm_NumRemessa.XNENumRemessaKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 Then Close;
end;


procedure TFrm_NumRemessa.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
