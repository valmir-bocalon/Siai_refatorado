unit uSiaiCotacoes;

interface

uses System.Classes;

type
  TSiaiCotacoesThread = class(TThread)
  private
    FSuccess: Boolean;
    FUSD, FEUR, FBTC: string;
  protected
    function FetchResponse: string; virtual;
    procedure Execute; override;
  public
    constructor Create;
    property Success: Boolean read FSuccess;
    property USD: string read FUSD;
    property EUR: string read FEUR;
    property BTC: string read FBTC;
  end;

implementation

uses System.SysUtils, System.JSON, IdHTTP, IdSSLOpenSSL;

constructor TSiaiCotacoesThread.Create;
begin
  inherited Create(True);
  FreeOnTerminate := True;
end;

function TSiaiCotacoesThread.FetchResponse: string;
var
  LHTTP: TIdHTTP;
  LSSL: TIdSSLIOHandlerSocketOpenSSL;
begin
  LHTTP := TIdHTTP.Create(nil);
  try
    LSSL := TIdSSLIOHandlerSocketOpenSSL.Create(LHTTP);
    LSSL.SSLOptions.Method := sslvTLSv1_2;
    LSSL.SSLOptions.Mode := sslmClient;
    LHTTP.IOHandler := LSSL;
    LHTTP.ConnectTimeout := 3000;
    LHTTP.ReadTimeout := 3000;
    LHTTP.HandleRedirects := True;
    LHTTP.RedirectMaximum := 3;
    LHTTP.Request.UserAgent := 'Mozilla/5.0';
    Result := LHTTP.Get(
      'https://economia.awesomeapi.com.br/json/last/USD-BRL,EUR-BRL,BTC-BRL');
  finally
    LHTTP.Free;
  end;
end;

procedure TSiaiCotacoesThread.Execute;
var
  LJSON: TJSONValue;
  LResponse: string;

  function ReadBid(const ACode: string): string;
  var
    LCurrency, LBid: TJSONValue;
  begin
    Result := '';
    if not (LJSON is TJSONObject) then Exit;
    LCurrency := TJSONObject(LJSON).GetValue(ACode);
    if not (LCurrency is TJSONObject) then Exit;
    LBid := TJSONObject(LCurrency).GetValue('bid');
    if (LBid is TJSONString) or (LBid is TJSONNumber) then
      Result := LBid.Value;
  end;

begin
  FSuccess := False;
  try
    if Terminated then Exit;
    LResponse := FetchResponse;
    if Terminated then Exit;
    LJSON := TJSONObject.ParseJSONValue(LResponse);
    try
      FUSD := ReadBid('USDBRL');
      FEUR := ReadBid('EURBRL');
      FBTC := ReadBid('BTCBRL');
      FSuccess := (FUSD <> '') and (FEUR <> '') and (FBTC <> '') and
        not Terminated;
    finally
      LJSON.Free;
    end;
  except
    { Cotacoes sao opcionais; falhas de rede/SSL/JSON nao afetam o acesso. }
    FSuccess := False;
  end;
end;

end.
