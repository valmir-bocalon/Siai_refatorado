unit HORLABEL2;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  quickrpt, Qrctrls, Db, DBTables, ExtCtrls, QRBarcode, CJVQRBarCode;

type
  THLabel = (h5160, h5161);

  TfrmHorLabels2 = class(TForm)
    QuickRep1: TQuickRep;
    QRBand1: TQRBand;
    QRMemo1: TQRMemo;
    QRMemo2: TQRMemo;
    QRMemo3: TQRMemo;
    procedure DetailBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure QuickRep1BeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    { Private declarations }
    FSkipRow: integer;
    LabelCnt: integer;
    Memos: array[0..2] of TQRMemo;
    code: array[0..2] of TCJVQRBarCode;
    procedure AddIfNotBlank(memo: TQRMemo; s: string);

    procedure AfterConstruction; override;
  public
    { Public declarations }
    LabelSize : THLabel;
    SkipRow: integer;
  end;

var
  frmHorLabels2: TfrmHorLabels2;

implementation

{$R *.DFM}
uses funcoes,horMain2, uRuntimeFields;

procedure TfrmHorLabels2.FormCreate(Sender: TObject);
begin
  EnsureRuntimeFields(Self);
  {
    Nºs temos uma ordem de TQRMemo e nós nomeamos cada elemento para o
    memo existentes no relatório.  Isto nos permite fixar a propridade do memo
    a tempo de desígnio e nos dá um modo fácil para referência
    os memo a runtime.
  }

  Memos[0] := qrmemo1;
  Memos[1] := qrmemo2;
  Memos[2] := qrmemo3;
//  code[0]  := bar1;
//  code[1]  := bar2;
//  code[2]  := bar3;

end;

procedure TfrmHorLabels2.AddIfNotBlank(memo: TQRMemo; s: string);
begin
  { Adiciona campos em branco}
  if Trim(s) <> '' then
    memo.Lines.Add(s);
end;

procedure TfrmHorLabels2.QuickRep1BeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
var
  nIdx: integer;
begin
  FSkipRow := SkipRow;
  if LabelSize = h5160 then
  begin
    memos[2].Enabled := true;
//    code[2].Enabled := true;
    LabelCnt := 2;
  end
  else begin
    memos[2].Enabled := FALSE;
//    code[2].Enabled := FALSE;    
    LabelCnt := 1;
  end;
  // mexe largura, esquerda da etiqueta
{  for nIdx := 0 to LabelCnt do
  begin
    with memos[nIdx] do
    begin
      if Enabled then begin
        // Fixe a largura baseado no tamanho da etiqueta escolhida
        case LabelSize of
          h5160: Size.Width := 2.60;
          h5161: Size.Width := 4;
        end;

        // Fixe o local baseado na largura
        if nIdx > 0 then
          Left := memos[nIdx-1].Left + memos[nIdx-1].Width
        else
          Left := 0;
      end;
    end
  end;}
end;

procedure TfrmHorLabels2.DetailBand1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
var
  nIdx : integer;
  Linhas: TArray<string>;
  i: Integer;
begin
  for nIdx := low(memos) to high(memos) do
    memos[nIdx].Lines.Clear;
//  for nIdx := low(memos) to high(memos) do
//    code[nIdx].Texto:='';

  if FSkipRow > 0 then
  begin
    dec(FSkipRow);
    quickrep1.DataSet.Prior;
  end
  else
  begin
    with quickrep1.DataSet do
    begin
      for nIdx := 0 to LabelCnt do
      begin
        { Acrescente texto ao memo.  Nºs não estamos usando Componentes de Dados do Quick Report
        porque nós estamos fazendo nosso próprio dataset navegação.}

      { Adicionando os campos no memo }
        if not EOF then
        begin
          { função para adiconar campos não em brancos no memo}
//          AddIfNotBlank(memos[nIdx], copy(Table1DESCRICAO.AsString,1,40));
          memos[nIdx].Lines.Add(''); // joga uma linha em branco fica mais para baixo

//          memos[nIdx].Lines.Add(uppercase(copy(frmMain2.CDS_etqnome_parte.Value,1,40)));
//          memos[nIdx].Lines.Add(uppercase(copy(frmMain2.CDS_etqendereco.Value,1,40)));
//          memos[nIdx].Lines.Add(uppercase(copy(frmMain2.CDS_etqbairro.Value,1,40))

          // NOME
          Linhas := QuebrarTexto(frmMain2.CDS_etqnome_parte.Value, 40, 1);  //40 e o tamnaho em caracteres e o 1 em quantas linhas
          for i := 0 to High(Linhas) do
            memos[nIdx].Lines.Add(Linhas[i]);

          // ENDEREÇO
          Linhas := QuebrarTexto(frmMain2.CDS_etqendereco.Value, 40, 2);
          for i := 0 to High(Linhas) do
            memos[nIdx].Lines.Add(Linhas[i]);

          // BAIRRO
          Linhas := QuebrarTexto(frmMain2.CDS_etqbairro.Value, 40, 1);
          for i := 0 to High(Linhas) do
            memos[nIdx].Lines.Add(Linhas[i]);

          if not empty(frmMain2.CDS_etqnomecidade.Value) then
             memos[nIdx].Lines.Add(uppercase(copy(frmMain2.CDS_etqnomecidade.Value,1,40))+' - '+uppercase(frmMain2.CDS_etqnomeestado.Value))
          else
             memos[nIdx].Lines.Add(uppercase(copy(frmMain2.CDS_etqnomecidade.Value,1,40)));   
          memos[nIdx].Lines.Add(frmMain2.CDS_etqcep.Value);
        end;
        {
          Se nós não estamos no Último label para esta Band e nós não estamos no
          fim do dataset, nós achamos o próximo registro.  Nºs não fazemos a busca
          no Último label o Relatório fará isto automaticamente e se nós
          fizermos, nós saltaremos todo 4ª registro.
        }
        if (nIdx < LabelCnt) and (not EOF) then
          Next
        else
          break;
      end;
    end;
  end;
end;


procedure TfrmHorLabels2.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
