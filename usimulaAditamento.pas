unit usimulaAditamento;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, XBanner, Vcl.StdCtrls, XNum, XLabel3D,
  Vcl.ExtCtrls, ACBrBase, ACBrEnterTab,math, Data.DB, Datasnap.DBClient,
  Vcl.Grids, Vcl.DBGrids, dxCore2, dxButton;

type
  Tfrmsimulaaditamento = class(TForm)
    XBanner10: TXBanner;
    XLabel3D1: TXLabel3D;
    XLabel3D2: TXLabel3D;
    XNumEdit1: TXNumEdit;
    XNumEdit2: TXNumEdit;
    tprice: TComboBox;
    XLabel3D3: TXLabel3D;
    XNumEdit3: TXNumEdit;
    XLabel3D4: TXLabel3D;
    XNumEdit4: TXNumEdit;
    XLabel3D5: TXLabel3D;
    Shape1: TShape;
    Shape2: TShape;
    XLabel3D6: TXLabel3D;
    XNumEdit5: TXNumEdit;
    XNumEdit6: TXNumEdit;
    XLabel3D7: TXLabel3D;
    XLabel3D8: TXLabel3D;
    XNumEdit7: TXNumEdit;
    ACBrEnterTab1: TACBrEnterTab;
    XLabel3D9: TXLabel3D;
    XNumEdit8: TXNumEdit;
    CdsAditamento: TClientDataSet;
    dsaditamento: TDataSource;
    CdsAditamentomes: TIntegerField;
    CdsAditamentovr_prestacao: TFloatField;
    CdsAditamentojuros: TFloatField;
    CdsAditamentoamortizacao: TFloatField;
    CdsAditamentosld_devedor: TFloatField;
    CdsAditamentoindice: TFloatField;
    DBGrid1: TDBGrid;
    dxButton5: TdxButton;
    dxButton1: TdxButton;
    Shape3: TShape;
    Shape4: TShape;
    procedure FormShow(Sender: TObject);
    procedure XNumEdit1Exit(Sender: TObject);
    procedure XNumEdit2Exit(Sender: TObject);
    procedure DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure dxButton1Click(Sender: TObject);
    procedure dxButton5Click(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  frmsimulaaditamento: Tfrmsimulaaditamento;

implementation

{$R *.dfm}

uses Venda, rel_SimulacaoAditamento,Funcoes, uRuntimeFields;

function CalcularParcela(PV: Double; i: Double; n: Integer): Double;
begin
  Result := PV * i * Power(1 + i, n) / (Power(1 + i, n) - 1);
end;

procedure Tfrmsimulaaditamento.DBGrid1DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  // zebrar dbgrid
  if not odd(CdsAditamento.RecNo) then
  begin
    if not (gdSelected in State) then
    begin
      DBGrid1.Canvas.Brush.Color := $00E9E9E9;
      DBGrid1.Canvas.FillRect(Rect);
    end;
  end;

  DBGrid1.DefaultDrawDataCell(Rect, DBGrid1.columns[datacol].field, State);
end;

procedure Tfrmsimulaaditamento.dxButton1Click(Sender: TObject);
begin
  close;
end;

procedure Tfrmsimulaaditamento.dxButton5Click(Sender: TObject);
begin
  try
    CdsAditamento.DisableControls;
    if frm_SimulacaoAditamento=nil then
       frm_SimulacaoAditamento:=Tfrm_SimulacaoAditamento.Create(self);
    frm_SimulacaoAditamento.RLLabel6.Caption :=tprice.Text;
    frm_SimulacaoAditamento.RLLabel13.Caption:=XNumEdit2.Text;
    frm_SimulacaoAditamento.RLReport1.PreviewModal;
  finally
    FreeAndNil(frm_SimulacaoAditamento);
    if not CdsAditamento.Active then
      CdsAditamento.Open;
    if CdsAditamento.Active then
      CdsAditamento.First;
    CdsAditamento.EnableControls;
  end;
end;

procedure Tfrmsimulaaditamento.FormShow(Sender: TObject);
begin
  xnumedit4.Value:=Frm_Venda.xdliquidado.Value;
  tprice.Text    :=Frm_Venda.tprice.Text;
  XNumEdit1.Value:=0;
  XNumEdit3.Value:=0;
  XNumEdit2.Value:=0;
  XNumEdit5.Value:=0;
  XNumEdit6.Value:=0;
  XNumEdit7.Value:=0;
  XNumEdit1.SetFocus;
end;

procedure Tfrmsimulaaditamento.XNumEdit1Exit(Sender: TObject);
begin
  XNumEdit3.Value:=XNumEdit4.Value-XNumEdit1.Value;
  tprice.SetFocus;
end;

procedure Tfrmsimulaaditamento.XNumEdit2Exit(Sender: TObject);
var
  PV,i,xindice,vrpresta,vrjuros,vramortizacao,vrsaldodev: Double;
  n,z: Integer;

begin
  PV := XNumEdit3.Value; // Valor da dívida
  i  := strtofloat(tprice.Text)/100; // Taxa de juros de 0,5% ao mês
  n  := floattostr(XNumEdit2.Value).ToInteger; // Número de prestações

  XNumEdit5.Value := CalcularParcela(PV, i, n);
  XNumEdit6.Value := XNumEdit5.Value * XNumEdit2.Value;
  XNumEdit7.Value := XNumEdit6.Value+XNumEdit1.Value;
  xindice         := (XNumEdit5.Value/XNumEdit3.Value);
  XNumEdit8.Value := (XNumEdit5.Value/XNumEdit3.Value)*1000;
  CdsAditamento.Close;
  CdsAditamento.CreateDataSet;
  CdsAditamento.Open;
  CdsAditamento.DisableControls;
  CdsAditamento.Append;
  CdsAditamentomes.Value          := 1;
  CdsAditamentovr_prestacao.Value := XNumEdit5.Value;
  CdsAditamentojuros.Value        := XNumEdit3.Value*i;
  CdsAditamentoamortizacao.Value  := XNumEdit5.Value-CdsAditamentojuros.Value;
  CdsAditamentosld_devedor.Value  := XNumEdit3.Value-CdsAditamentoamortizacao.Value;
  CdsAditamentoindice.Value       := xindice;
  CdsAditamento.Post;

  vrpresta     :=CdsAditamentovr_prestacao.Value;
  vrjuros      :=CdsAditamentojuros.Value;
  vramortizacao:=CdsAditamentoamortizacao.Value;
  vrsaldodev   :=CdsAditamentosld_devedor.Value;
  xindice      :=CdsAditamentoindice.Value;

  for z := 1 to n -1 do
  begin

    CdsAditamento.Append;
    CdsAditamentomes.Value          := z+1;
    CdsAditamentovr_prestacao.Value := vrpresta;
    CdsAditamentojuros.Value        := vrsaldodev*i;
    CdsAditamentoamortizacao.Value  := vrpresta-CdsAditamentojuros.Value;
    CdsAditamentosld_devedor.Value  := vrsaldodev-CdsAditamentoamortizacao.Value;
    CdsAditamentoindice.Value       := vrpresta/CdsAditamentosld_devedor.Value;
    CdsAditamento.Post;
    vrpresta                        :=CdsAditamentovr_prestacao.Value;
    vrjuros                         :=CdsAditamentojuros.Value;
    vramortizacao                   :=CdsAditamentoamortizacao.Value;
    vrsaldodev                      :=CdsAditamentosld_devedor.Value;
    xindice                         :=CdsAditamentoindice.Value;
  end;
  CdsAditamento.First;
  CdsAditamento.EnableControls;


end;


procedure Tfrmsimulaaditamento.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
