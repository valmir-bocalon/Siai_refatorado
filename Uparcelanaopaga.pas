unit Uparcelanaopaga;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, XBanner, Data.DB, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, Vcl.ComCtrls, Vcl.StdCtrls, Vcl.ExtCtrls;

type
  Tfrmnaopaga = class(TForm)
    ZQReceb_nao_pagas: TZQuery;
    DS_ZQReceb_nao_pagas: TDataSource;
    XBanner14: TXBanner;
    Button1: TButton;
    Button2: TButton;
    Mes: TMonthCalendar;
    ZQUltimodia: TZQuery;

    lprogress: TLabel;
    Timer1: TTimer;
    procedure Button2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  frmnaopaga: Tfrmnaopaga;
  totalreg:integer;

implementation

{$R *.dfm}

uses tabelas,funcoes, uRuntimeFields;

procedure Tfrmnaopaga.Button1Click(Sender: TObject);
var
DtMes:TdateTime;
dtini,dtultimo:string;

begin
  lprogress.Visible:=true;
  application.ProcessMessages;
  DtMes:=mes.Date;
  dtini:= copy(datetostr(dtMes),7,4)+'-'+copy(datetostr(dtMes),4,2)+'-01';

  ZQUltimodia.Close;
  ZQUltimodia.SQL.Clear;
  ZQUltimodia.SQL.Add('SELECT last_day('+quotedstr(dtini)+') as ultimo');
  ZQUltimodia.Open;
  dtUltimo:=copy(datetostr(ZQUltimodia.FieldByName('ultimo').AsDateTime),7,4)+'-'+copy(datetostr(ZQUltimodia.FieldByName('ultimo').AsDateTime),4,2)+'-'+copy(datetostr(ZQUltimodia.FieldByName('ultimo').AsDateTime),1,2);

  ZQReceb_nao_pagas.Close;
  ZQReceb_nao_pagas.SQL.Clear;
  ZQReceb_nao_pagas.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  ZQReceb_nao_pagas.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  ZQReceb_nao_pagas.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  ZQReceb_nao_pagas.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
  ZQReceb_nao_pagas.SQL.Add(' from recebimento where dt_vencimento between '+quotedstr(dtini)+' and last_day('+quotedstr(dtini)+') and saldo>0');
  ZQReceb_nao_pagas.open;
  totalreg:=ZQReceb_nao_pagas.RecordCount;

  ZQReceb_nao_pagas.Close;
  ZQReceb_nao_pagas.SQL.Clear;
  ZQReceb_nao_pagas.SQL.Add('Update recebimento set dt_nao_pagou_no_mes='+quotedstr(dtultimo)+' where dt_vencimento between '+quotedstr(dtini)+' and last_day('+quotedstr(dtini)+') and saldo>0');
  ZQReceb_nao_pagas.ExecSQL;
  DM_Tabelas.ZQRecebimento.Refresh;
  ZQReceb_nao_pagas.Close;
  lprogress.Visible:=false;
  application.ProcessMessages;
  showmessage('Processo Terminado!');
end;

procedure Tfrmnaopaga.Button2Click(Sender: TObject);
begin
  close;
end;

procedure Tfrmnaopaga.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  ZQReceb_nao_pagas.Close;
//  DM_Tabelas.ZQRecebimento.Refresh;
  frmnaopaga:=nil;
  Action:=CaFree;
end;


procedure Tfrmnaopaga.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(Tfrmnaopaga, 'ZQUltimodia', False);
  RegisterRuntimeField(Tfrmnaopaga, 'ZQUltimodia', 'ZQUltimodiaultimo', 'ultimo', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
