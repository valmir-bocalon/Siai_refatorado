unit integridade_baixa;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  Tfrm_baixa = class(TForm)
    Button1: TButton;
    Button2: TButton;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  frm_baixa: Tfrm_baixa;

implementation

uses tabelas, uRuntimeFields;

{$R *.dfm}

procedure Tfrm_baixa.Button1Click(Sender: TObject);
Var
  Varrefbaixa, varregistro : integer;
  varsqltexto : string;
  Varpago, percentbaixa, baixatot, varjuros, varx,
  varsomarepasse, varsomarepasseautomatico : Double;
  contador,varreg : Integer;
  varordem, varnumordem1, varnumordem2 : string;

begin
  DM_tabelas.ZQLoteamento.close;
  DM_tabelas.ZQLoteamento.SQL.Clear;
  DM_tabelas.ZQLoteamento.SQL.add('Select * from Loteamento order by refbaixa');
  DM_tabelas.ZQLoteamento.open;
  DM_tabelas.ZQLoteamento.first;
  while not DM_tabelas.ZQRecebimento.Eof do
  begin
    if DM_tabelas.ZQRecebimento.FieldByName('RefBaixa').Value>0 Then
    Begin
      DM_TAbelas.zqrecebbxTemp.open;
      Varrefbaixa := DM_tabelas.ZQRecebBxTemp.FieldByName('refer').Value;
      DM_TAbelas.zqrecebbxTemp.Close;

      DM_tabelas.ZQReBxHi.SQL.Clear;
      DM_tabelas.ZQReBxHi.SQL.Add('select * from RecBxHist where refer='+inttostr(Varrefbaixa));
      DM_tabelas.ZQReBxHi.Open;
      DM_tabelas.ZQReBxHi.First;
      varregistro := DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').Value;
      varsqltexto := DM_tabelas.ZQRecebimento.SQL.Text;
      DM_tabelas.ZQRecebimento.SQL.Clear;
      DM_tabelas.ZQRecebimento.SQL.Add('delete from recebimento where RefVinda = '+inttostr(Varrefbaixa));
      DM_tabelas.ZQRecebimento.ExecSQL;


      ZQReceb_Baixa.SQL.Clear;
      ZQReceb_Baixa.SQL.Add('delete from receb_baixa where refbaixa='+inttostr(Varrefbaixa));
      ZQReceb_Baixa.ExecSQL;


      ZQReceb_Baixa.Close;
      ZQReceb_Baixa.SQL.clear;
      ZQReceb_Baixa.SQL.Add(' Select * From receb_baixa');
      ZQReceb_Baixa.open;

      DM_tabelas.ZQReBxHi.Next;

      ZQReceb_Baixa.Close;
      DM_tabelas.ZQReBxHi.SQL.Clear;
      DM_tabelas.ZQReBxHi.SQL.Add('delete from RecBxHist where refer='+inttostr(Varrefbaixa));
      DM_tabelas.ZQReBxHi.ExecSQL;

///    DM_tabelas.ZQReBxHi.SQL.Clear;
//    DM_tabelas.ZQReBxHi.SQL.Add('delete from receb_baixa where refbaixa='+inttostr(Varrefbaixa));
  //  DM_tabelas.ZQReBxHi.ExecSQL;


      DM_tabelas.ZQReBxHi.SQL.Clear;
      DM_tabelas.ZQReBxHi.SQL.Add('select * from RecBxHist');
      DM_tabelas.ZQReBxHi.Open;
      DM_tabelas.ZQRecebimento.Locate('IdRecebimento',varregistro,[]);
    end;
    DM_tabelas.ZQRecebimento.Next;
  end;  
end;


procedure Tfrm_baixa.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
