unit acerta_remessa_recebida;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset,  
  dxButton, Gauges, StdCtrls, Mask, XDate, XBanner, dxCore2;

type
  Tfrm_acerta_remessa = class(TForm)
    Gauge1: TGauge;
    BTFechar: TdxButton;
    dxButton1: TdxButton;
    ZQRemesRec: TZQuery;










    DS_RemesRec: TDataSource;
    GroupBox1: TGroupBox;
    XBanner1: TXBanner;
    Label2: TLabel;
    Label3: TLabel;
    XDEEntradaInicio: TXDateEdit;
    XDEEntradaFinal: TXDateEdit;
    XBanner2: TXBanner;
    procedure dxButton1Click(Sender: TObject);
    procedure BTFecharClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  frm_acerta_remessa: Tfrm_acerta_remessa;

implementation

uses tabelas, funcoes, uRuntimeFields;

{$R *.dfm}

procedure Tfrm_acerta_remessa.dxButton1Click(Sender: TObject);
begin
  Close;
end;

procedure Tfrm_acerta_remessa.BTFecharClick(Sender: TObject);
begin
  ZQRemesRec.Close;
  ZQRemesRec.SQL.Clear;
  ZQRemesRec.SQL.Add('delete from remessa_receb where dt_ocorrencia between :dt3 and :dt4');
  ZQRemesRec.ExecSQL;
  ZQRemesRec.ParamByName('dt3').AsDate:=strtodate(XDEEntradaInicio.DateText);
  ZQRemesRec.ParamByName('dt4').AsDate:=strtodate(XDEEntradaFinal.DateText);
  ZQRemesRec.ExecSQL;
  ZQRemesRec.close;
  Close;
end;

procedure Tfrm_acerta_remessa.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  frm_acerta_remessa:=nil;
  Action:=CaFree;
end;

procedure Tfrm_acerta_remessa.FormShow(Sender: TObject);
begin
  XDEEntradaInicio.setfocus;  
end;


procedure Tfrm_acerta_remessa.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(Tfrm_acerta_remessa, 'ZQRemesRec', False);
  RegisterRuntimeField(Tfrm_acerta_remessa, 'ZQRemesRec', 'ZQRemesRecidremessa_receb', 'idremessa_receb', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_acerta_remessa, 'ZQRemesRec', 'ZQRemesRecremessa', 'remessa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_acerta_remessa, 'ZQRemesRec', 'ZQRemesRecidrec', 'idrec', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_acerta_remessa, 'ZQRemesRec', 'ZQRemesRecretorno_motivo', 'retorno_motivo', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_acerta_remessa, 'ZQRemesRec', 'ZQRemesRecacao', 'acao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_acerta_remessa, 'ZQRemesRec', 'ZQRemesRecnossonumero', 'nossonumero', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_acerta_remessa, 'ZQRemesRec', 'ZQRemesRecocorrencia', 'ocorrencia', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_acerta_remessa, 'ZQRemesRec', 'ZQRemesRecdesc_motivo', 'desc_motivo', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_acerta_remessa, 'ZQRemesRec', 'ZQRemesReccredito', 'credito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrm_acerta_remessa, 'ZQRemesRec', 'ZQRemesRecdt_ocorrencia', 'dt_ocorrencia', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
