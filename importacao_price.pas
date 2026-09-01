unit importacao_price;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ZConnection, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset,
  StdCtrls, ComCtrls;

type
  TForm1 = class(TForm)
    DS_Price: TDataSource;
    ZQprice: TZQuery;





    zconeccao: TZConnection;
    zm1: TZQuery;
    DS_zm1: TDataSource;


    ZA1: TZQuery;
    DS_za1: TDataSource;


    zm1_5: TZQuery;
    Data_zm1_5: TDataSource;
    Data_ZA1_5: TDataSource;
    ZA1_5: TZQuery;
    zm2: TZQuery;
    Data_zm2: TDataSource;
    Data_ZA2: TDataSource;
    ZA2: TZQuery;
    zm2_5: TZQuery;
    Data_zm2_5: TDataSource;
    ZA2_5: TZQuery;
    Data_ZA2_5: TDataSource;
    Button1: TButton;
    Button2: TButton;












    br: TProgressBar;
    Zm0625: TZQuery;
    DS_Zm0625: TDataSource;
    Zza0625: TZQuery;
    DSZza0625: TDataSource;






    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation
uses
  uRuntimeFields;

{$R *.dfm}

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  ZQprice.close;
  zm1.close;
  ZA1.close;
  zm1_5.close;
  ZA1_5.close;
  zm2.close;
  ZA2.close;
  zm2_5.close;
  ZA2_5.close;
  Zm0625.close;
  Zza0625.close;

end;

procedure TForm1.FormShow(Sender: TObject);
begin
  ZQprice.open;
  zm1.open;
  ZA1.open;
  zm1_5.open;
  ZA1_5.open;
  zm2.open;
  ZA2.open;
  zm2_5.open;
  ZA2_5.open;
  Zm0625.Open;
  Zza0625.Open;
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  ZQprice.First;
  while not ZQprice.eof do
  begin
    ZQprice.delete;
  end;

  br.Position:=0;
  br.Max:= Zm0625.RecordCount;
  Zm0625.First;
  while not Zm0625.eof do
  begin
    br.Position:=Zm0625.RecNo;
    ZQprice.insert;
    ZQprice.FieldByName('coeficiente_multiplicador').Value:=Zm0625.FieldByName('Valor').Value;
    ZQprice.FieldByName('Coeficiente_amortizador').Value:=Zza0625.FieldByName('valor').Value;
    ZQprice.FieldByName('Taxa').Value:=0.625;
    ZQprice.FieldByName('Mes').Value:=strtoint(Zm0625.FieldByName('Mes').Value);
    ZQprice.post;
    Zm0625.Next;
  end;

  br.Position:=0;
  br.Max:= zm1.RecordCount;
  zm1.First;
  while not zm1.eof do
  begin
    br.Position:=zm1.RecNo;
    ZQprice.insert;
    ZQprice.FieldByName('coeficiente_multiplicador').Value:=zm1.FieldByName('valor').Value;
    ZQprice.FieldByName('Coeficiente_amortizador').Value:=ZA1.FieldByName('valor').Value;
    ZQprice.FieldByName('Taxa').Value:=1.00;
    ZQprice.FieldByName('Mes').Value:=strtoint(zm1.FieldByName('mes').Value);
    ZQprice.post;
    zm1.Next;
  end;

  zm1_5.First;
  br.Position:=0;
  br.Max:= zm1.RecordCount;

  while not zm1_5.eof do
  begin
    br.Position:=zm1_5.RecNo;
    ZQprice.insert;
    ZQprice.FieldByName('coeficiente_multiplicador').Value:=zm1_5.FieldByName('valor').Value;
    ZQprice.FieldByName('Coeficiente_amortizador').Value:=ZA1_5.FieldByName('valor').Value;
    ZQprice.FieldByName('Taxa').Value:=1.50;
    ZQprice.FieldByName('Mes').Value:=strtoint(zm1_5.FieldByName('mes').Value);
    ZQprice.post;
    zm1_5.Next;
  end;


  zm2.First;
  br.Position:=0;
  br.Max:= zm2.RecordCount;

  while not zm2.eof do
  begin
    br.Position:=zm2.RecNo;
    ZQprice.insert;
    ZQprice.FieldByName('coeficiente_multiplicador').Value:=zm2.FieldByName('valor').Value;
    ZQprice.FieldByName('Coeficiente_amortizador').Value:=ZA2.FieldByName('valor').Value;
    ZQprice.FieldByName('Taxa').Value:=2.00;
    ZQprice.FieldByName('Mes').Value:=strtoint(zm2.FieldByName('mes').Value);
    ZQprice.post;
    zm2.Next;
  end;

  zm2_5.First;
  br.Position:=0;
  br.Max:= zm2_5.RecordCount;

  while not zm2_5.eof do
  begin
    br.Position:=zm2_5.RecNo;
    ZQprice.insert;
    ZQprice.FieldByName('coeficiente_multiplicador').Value:=zm2_5.FieldByName('valor').Value;
    ZQprice.FieldByName('Coeficiente_amortizador').Value:=ZA2_5.FieldByName('valor').Value;
    ZQprice.FieldByName('Taxa').Value:=2.50;
    ZQprice.FieldByName('Mes').Value:=strtoint(zm2_5.FieldByName('mes').Value);
    ZQprice.post;
    zm2_5.Next;
  end;
  br.Position:=0;
  close;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  close;
end;


procedure TForm1.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TForm1, 'ZQprice', False);
  RegisterRuntimeField(TForm1, 'ZQprice', 'ZQpriceidPrice', 'idPrice', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQprice', 'ZQpricecoeficiente_multiplicador', 'coeficiente_multiplicador', TFloatField, fkData, 0, 0, False, '', '##,###,##0.00000', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQprice', 'ZQpriceCoeficiente_amortizador', 'Coeficiente_amortizador', TFloatField, fkData, 0, 0, False, '', '##,###,##0.00000', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQprice', 'ZQpriceTaxa', 'Taxa', TFloatField, fkData, 0, 0, False, '', '##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQprice', 'ZQpriceMes', 'Mes', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TForm1, 'zm1', False);
  RegisterRuntimeField(TForm1, 'zm1', 'zm1mes', 'mes', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zm1', 'zm1valor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TForm1, 'ZA1', False);
  RegisterRuntimeField(TForm1, 'ZA1', 'ZA1mes', 'mes', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZA1', 'ZA1valor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TForm1, 'zm1_5', False);
  RegisterRuntimeField(TForm1, 'zm1_5', 'zm1_5mes', 'mes', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zm1_5', 'zm1_5valor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TForm1, 'ZA1_5', False);
  RegisterRuntimeField(TForm1, 'ZA1_5', 'ZA1_5mes', 'mes', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZA1_5', 'ZA1_5valor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TForm1, 'zm2', False);
  RegisterRuntimeField(TForm1, 'zm2', 'zm2mes', 'mes', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zm2', 'zm2valor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TForm1, 'ZA2', False);
  RegisterRuntimeField(TForm1, 'ZA2', 'ZA2mes', 'mes', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZA2', 'ZA2valor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TForm1, 'zm2_5', False);
  RegisterRuntimeField(TForm1, 'zm2_5', 'zm2_5mes', 'mes', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zm2_5', 'zm2_5valor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TForm1, 'ZA2_5', False);
  RegisterRuntimeField(TForm1, 'ZA2_5', 'ZA2_5mes', 'mes', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZA2_5', 'ZA2_5valor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TForm1, 'Zm0625', False);
  RegisterRuntimeField(TForm1, 'Zm0625', 'Zm0625Mes', 'Mes', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'Zm0625', 'Zm0625Valor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'Zm0625', 'Zm0625id', 'id', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TForm1, 'Zza0625', False);
  RegisterRuntimeField(TForm1, 'Zza0625', 'Zza0625mes', 'mes', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'Zza0625', 'Zza0625valor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'Zza0625', 'Zza0625id', 'id', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
