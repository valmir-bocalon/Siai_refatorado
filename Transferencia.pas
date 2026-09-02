unit Transferencia;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,   dxButton, StdCtrls, Gauges, OleServer, ExcelXP, Grids,
  ComObj, IniFiles, JvExControls, JvSimPID, JvSimIndicator, Excel2000, dxCore2;

type
  TFrm_Transferencia = class(TForm)
    Gauge1: TGauge;
    Label1: TLabel;
    BTFechar: TdxButton;
    dxButton1: TdxButton;
    ExcelApplication1: TExcelApplication;
    SGPart: TStringGrid;
    Gauge2: TGauge;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    BarraVertical: TJvSimIndicator;
    procedure BTFecharClick(Sender: TObject);
    function XlsToStringGrid(AGrid: TStringGrid; AXLSFile: string): Boolean;
    procedure dxButton1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);

  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_Transferencia: TFrm_Transferencia;

implementation

uses tabelas, funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrm_Transferencia.BTFecharClick(Sender: TObject);
begin
  Close;
end;

function TFrm_Transferencia.XlsToStringGrid(AGrid: TStringGrid; AXLSFile: string): Boolean;
const
    xlCellTypeLastCell = $0000000B;
var
    XLApp, Sheet: OLEVariant;
    RangeMatrix: Variant;
    x, y, k, r: Integer;
begin
Result:=False;
//Cria Excel- OLE Object
XLApp:=CreateOleObject('Excel.Application');
try
    //Esconde Excel
    XLApp.Visible:=False;
    //Abre o Workbook
    XLApp.Workbooks.Open(AXLSFile);
    Sheet:=XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[1];
    Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;
    //Pegar o número da Última linha
    x:=XLApp.ActiveCell.Row;
    //Pegar o número da Última coluna
    y:=XLApp.ActiveCell.Column;
    //Seta Stringgrid linha e coluna
    AGrid.RowCount:=x;
    AGrid.ColCount:=y;
    Gauge1.MaxValue := x;
    //Associaca a variant WorkSheet com a variant do Delphi
    RangeMatrix:=XLApp.Range['A1', XLApp.Cells.Item[X, Y]].Value;
    //Cria o loop para listar os registros no TStringGrid
    k:=1;
    repeat
        Gauge1.Progress := k;
        for r:=1 to y do
            AGrid.Cells[(r - 1),(k - 1)]:=RangeMatrix[K, R];
        Inc(k,1);
    until k > x;
    RangeMatrix:=Unassigned;
finally
    //Fecha o Excel
    if not VarIsEmpty(XLApp) then
        begin
        XLApp.Quit;
        XLAPP:=Unassigned;
        Sheet:=Unassigned;
        Result:=True;
        end;
    end;
end;

procedure TFrm_Transferencia.dxButton1Click(Sender: TObject);
Var
  Varlinhas : integer;
  varind, varcid, varcep : string;
begin


  DM_Tabelas.ZqParticipante.SQL.Clear;
  DM_Tabelas.ZqParticipante.SQL.Add('delete from incorporador_loteamento');
  DM_Tabelas.ZqParticipante.ExecSQL;
  DM_Tabelas.ZqParticipante.SQL.Clear;
  DM_Tabelas.ZqParticipante.SQL.Add('delete from representa_participante');
  DM_Tabelas.ZqParticipante.ExecSQL;
  DM_Tabelas.ZqParticipante.SQL.Clear;
  DM_Tabelas.ZqParticipante.SQL.Add('delete from quadras');
  DM_Tabelas.ZqParticipante.ExecSQL;
  DM_Tabelas.ZqParticipante.SQL.Clear;
  DM_Tabelas.ZqParticipante.SQL.Add('delete from imovel');
  DM_Tabelas.ZqParticipante.ExecSQL;
  DM_Tabelas.ZqParticipante.SQL.Clear;
  DM_Tabelas.ZqParticipante.SQL.Add('delete from loteamento');
  DM_Tabelas.ZqParticipante.ExecSQL;
  DM_Tabelas.ZqParticipante.SQL.Clear;
  DM_Tabelas.ZqParticipante.SQL.Add('delete from Conta_Bancaria');
  DM_Tabelas.ZqParticipante.ExecSQL;
  DM_Tabelas.ZqParticipante.SQL.Clear;
  DM_Tabelas.ZqParticipante.SQL.Add('delete from Participante');
  DM_Tabelas.ZqParticipante.ExecSQL;
  DM_Tabelas.ZqParticipante.SQL.Clear;
  DM_Tabelas.ZqParticipante.SQL.Add('delete from cidade');
  DM_Tabelas.ZqParticipante.ExecSQL;
  DM_Tabelas.ZqParticipante.SQL.Clear;
  DM_Tabelas.ZqParticipante.SQL.Add('select * from Participante');
  DM_Tabelas.ZqParticipante.Open;
  BarraVertical.Maximum := 10;
  Label3.Caption := 'Guardando os imóveis....';
  BarraVertical.Value := 1;
  XlsToStringGrid(SGPart,ExtractFilePath( Application.ExeName )+'imovel.xls');
  Gauge2.MaxValue := SGPart.RowCount-1;
  for varlinhas:=1 to SGPart.RowCount-1 do Begin
    varind := SGPart.Cells[1,varlinhas];
    if not empty(varind) Then Begin
      Gauge2.Progress := varlinhas;
      Label4.Caption := inttostr(varlinhas);
      DM_Tabelas.ZQImovel.Insert;
      DM_Tabelas.ZQimovel.FieldByName('idimovel').AsLargeInt := varlinhas;
      DM_Tabelas.ZQimovel.FieldByName('loteamento_idloteamento').AsLargeInt := strtoint(SGPart.Cells[0,varlinhas]);
      DM_Tabelas.ZQimovel.FieldByName('quadra').AsString := SGPart.Cells[1,varlinhas];
      DM_Tabelas.ZQimovel.FieldByName('lote').AsString := SGPart.Cells[2,varlinhas];
      DM_Tabelas.ZQimovel.FieldByName('lado').AsString := SGPart.Cells[18,varlinhas];
      if not empty(SGPart.Cells[11,varlinhas]) Then
        DM_Tabelas.ZQimovel.FieldByName('area').AsFloat := strtofloat(BuscaTroca(SGPart.Cells[11,varlinhas],'.',''));
      DM_tabelas.ZQimovel.FieldByName('confrontofrente').AsString := UpperCase(SGPart.Cells[12,varlinhas]);
      DM_Tabelas.ZQimovel.FieldByName('confrontofundo').AsString := SGPart.Cells[13,varlinhas];
      DM_tabelas.ZQimovel.FieldByName('confrontodireita').AsString := SGPart.Cells[14,varlinhas];
      DM_Tabelas.ZQimovel.FieldByName('confrontoesquerda').AsString := SGPart.Cells[15,varlinhas];
      DM_Tabelas.ZQimovel.FieldByName('esquinanome').AsString := SGPart.Cells[20,varlinhas];
      DM_Tabelas.ZQimovel.FieldByName('valorvenal').AsFloat := strtofloat(SGPart.Cells[16,varlinhas]);
      DM_Tabelas.ZQimovel.FieldByName('Obervacaoloteamento').AsString := SGPart.Cells[3,varlinhas];
      DM_Tabelas.ZQimovel.FieldByName('cornomapa').AsString := UpperCase(SGPart.Cells[6,varlinhas]);
      if SGPart.Cells[4,varlinhas]='VERDADEIRO' then
        DM_Tabelas.ZQimovel.FieldByName('disponivel').AsString := 'SIM'
      else
        DM_Tabelas.ZQimovel.FieldByName('disponivel').AsString := 'NºO';
      DM_Tabelas.ZQimovel.FieldByName('medidafrente').AsString := SGPart.Cells[7,varlinhas];
      DM_Tabelas.ZQimovel.FieldByName('medidafundo').AsString := SGPart.Cells[8,varlinhas];
      DM_Tabelas.ZQimovel.FieldByName('medidadedireita').AsString := SGPart.Cells[9,varlinhas];
      DM_Tabelas.ZQimovel.FieldByName('medidaesquerda').AsString := SGPart.Cells[10,varlinhas];
      DM_Tabelas.ZQimovel.FieldByName('esquinalinha').AsString := SGPart.Cells[19,varlinhas];
      DM_Tabelas.ZQimovel.FieldByName('esquinacurva').AsString := SGPart.Cells[21,varlinhas];
      DM_Tabelas.zqimovel.Post;

    end;
  end;
  Label3.Caption := 'Guardando as quadras...';
  DM_Tabelas.ZQImovel.SQL.Clear;
  DM_Tabelas.ZQImovel.SQL.Add('select * from imovel group by loteamento_idloteamento, quadra');
  DM_Tabelas.ZQImovel.Open;
  Gauge2.MaxValue := DM_Tabelas.ZQImovel.RecordCount;
  while not DM_Tabelas.ZQImovel.Eof do begin
    Gauge2.Progress := DM_Tabelas.ZQImovel.RecNo;
    DM_Tabelas.ZQQuadras.Insert;
    DM_Tabelas.ZQquadras.FieldByName('loteamento_idLoteamento').AsInteger := DM_Tabelas.ZQimovel.FieldByName('loteamento_idloteamento').AsLargeInt;
    DM_Tabelas.ZQquadras.FieldByName('descricao').AsString := DM_Tabelas.ZQimovel.FieldByName('quadra').AsString;
    DM_Tabelas.ZQQuadras.Post;
    DM_tabelas.ZQImovel.Next;
  end;
  DM_Tabelas.ZQImovel.SQL.Clear;
  DM_Tabelas.ZQImovel.SQL.Add('select * from imovel ');
  DM_Tabelas.ZQImovel.Open;

  Label3.Caption := 'Guardando as participantes...';
  BarraVertical.Value := 2;
  XlsToStringGrid(SGPart,ExtractFilePath( Application.ExeName )+'participante.xls');
  Gauge2.MaxValue := SGPart.RowCount-1;
  for varlinhas:=1 to SGPart.RowCount-1 do Begin
    varind := SGPart.Cells[1,varlinhas];
    if not empty(varind) Then Begin
      Gauge2.Progress := varlinhas;
      Label4.Caption := inttostr(varlinhas);
      varcid := UpperCase(tiraacento(SGPart.Cells[12,varlinhas]));
      if not DM_Tabelas.ZQCidade.Locate('nomecid',varcid,[]) Then Begin
        DM_Tabelas.ZQCidade.Insert;
        DM_Tabelas.ZQCidade.FieldByName('nomecid').AsString := varcid;
        DM_Tabelas.ZQCidade.FieldByName('estado').AsString := SGPart.Cells[11,varlinhas];
        DM_Tabelas.ZQCidade.Post;
      end;
      DM_Tabelas.ZqParticipante.Insert;
      DM_Tabelas.ZQParticipante.FieldByName('idpaticipante').AsLargeInt := strtoint(SGPart.Cells[0,varlinhas]);
      DM_Tabelas.ZqParticipante.FieldByName('cadastrado').AsDateTime := strtodate(copy(SGPart.Cells[1,varlinhas],1,10));
      DM_Tabelas.ZQParticipante.FieldByName('nome_parte').AsString :=UpperCase(tiraacento(SGPart.Cells[2,varlinhas]));
      DM_Tabelas.ZqParticipante.FieldByName('nacionalidade').AsString := UpperCase(SGPart.Cells[4,varlinhas]);
      DM_Tabelas.ZqParticipante.FieldByName('ESTADOCIVIL').AsString := UpperCase(SGPart.Cells[5,varlinhas]);
      DM_Tabelas.ZqParticipante.FieldByName('profissao').AsString := UpperCase(SGPart.Cells[6,varlinhas]);
      DM_Tabelas.ZqParticipante.FieldByName('doc1').AsString := SGPart.Cells[7,varlinhas];
      DM_Tabelas.ZqParticipante.FieldByName('doc2').AsString := SGPart.Cells[8,varlinhas];
      DM_Tabelas.ZqParticipante.FieldByName('aniversario').AsDateTime := strtodate(copy(SGPart.Cells[9,varlinhas],1,10));
      DM_Tabelas.ZqParticipante.FieldByName('tipopessoa').AsString := copy(UpperCase(SGPart.Cells[10,varlinhas]),1,1);
      DM_Tabelas.ZqParticipante.FieldByName('localdetrab').AsString :=   UpperCase(SGPart.Cells[13,varlinhas]);
      DM_Tabelas.ZqParticipante.FieldByName('renda').AsFloat := strtofloat(SGPart.Cells[15,varlinhas]);
      DM_Tabelas.ZqParticipante.FieldByName('naturalidade').AsInteger := DM_Tabelas.ZQCidade.FieldByName('idcidade').AsLargeInt;
      if not Verif_doc(DM_Tabelas.ZqParticipante.FieldByName('doc1').AsString,False) Then
        DM_Tabelas.ZqParticipante.FieldByName('observacao').AsString := DM_Tabelas.ZqParticipante.FieldByName('observacao').AsString+chr(13)+'Documento inválido';
      DM_Tabelas.ZqParticipante.Post;
    end;
  end;
  Label3.Caption := 'Guardando os emails dos clientes...';
  BarraVertical.Value := 3;
  XlsToStringGrid(SGPart,ExtractFilePath( Application.ExeName )+'emailcli.xls');
  Gauge2.MaxValue := SGPart.RowCount-1;
  for varlinhas:=1 to SGPart.RowCount-1 do Begin
    varind := SGPart.Cells[1,varlinhas];
    if not empty(varind) Then Begin
      Gauge2.Progress := varlinhas;
      Label4.Caption := inttostr(varlinhas);
      if DM_Tabelas.ZqParticipante.Locate('idpaticipante',strtoint(SGPart.Cells[0,varlinhas]),[]) Then Begin
        DM_Tabelas.ZqParticipante.Edit;
        DM_Tabelas.ZqParticipante.FieldByName('email').AsString := SGPart.Cells[3,varlinhas];
        DM_Tabelas.ZQParticipante.Post;
      end;
    end;
  end;
  Label3.Caption := 'Guardando aos conjuges...';
  BarraVertical.Value := 4;
  XlsToStringGrid(SGPart,ExtractFilePath( Application.ExeName )+'conjuge.xls');
  Gauge2.MaxValue := SGPart.RowCount-1;
  for varlinhas:=1 to SGPart.RowCount-1 do Begin
    varind := SGPart.Cells[1,varlinhas];
    if not empty(varind) Then Begin
      Gauge2.Progress := varlinhas;
      Label4.Caption := inttostr(varlinhas);
      if DM_Tabelas.ZqParticipante.Locate('idpaticipante',strtoint(SGPart.Cells[1,varlinhas]),[]) Then Begin
        DM_Tabelas.ZQresponsavel.Insert;
        DM_Tabelas.ZQresponsavel.FieldByName('paticipante_idpaticipante').AsLargeInt := strtoint(SGPart.Cells[0,varlinhas]);
        DM_Tabelas.ZQresponsavel.FieldByName('documen1').AsString := DM_tabelas.ZqParticipante.FieldByName('doc1').AsString;
        DM_Tabelas.ZQresponsavel.FieldByName('documen2').AsString := DM_Tabelas.ZqParticipante.FieldByName('doc2').AsString;
        DM_Tabelas.ZQresponsavel.FieldByName('nasicmento').AsDateTime := DM_Tabelas.ZqParticipante.FieldByName('aniversario').AsDateTime;
        DM_Tabelas.ZQresponsavel.FieldByName('cid_natural').AsLargeInt := DM_tabelas.ZqParticipante.FieldByName('naturalidade').AsInteger;
        DM_Tabelas.ZQresponsavel.FieldByName('nomerepres').AsString := DM_Tabelas.ZQParticipante.FieldByName('nome_parte').AsString;
        DM_Tabelas.ZQresponsavel.FieldByName('renda').AsFloat := DM_Tabelas.ZqParticipante.FieldByName('renda').AsFloat;
        DM_Tabelas.ZQresponsavel.FieldByName('emailrepres').AsString := DM_Tabelas.ZqParticipante.FieldByName('email').AsString;
        DM_Tabelas.ZQresponsavel.FieldByName('regime').AsString := SGPart.Cells[2,varlinhas];
        DM_Tabelas.ZQresponsavel.FieldByName('livro').AsString := SGPart.Cells[7,varlinhas];
        DM_Tabelas.ZQresponsavel.FieldByName('folha').AsString := SGPart.Cells[8,varlinhas];
        if not empty(SGPart.Cells[9,varlinhas]) Then
          DM_Tabelas.ZQresponsavel.FieldByName('data').AsDateTime := strtodate(SGPart.Cells[9,varlinhas]);
        DM_Tabelas.ZQresponsavel.FieldByName('cartorio').AsString := SGPart.Cells[4,varlinhas];
        varcid := UpperCase(tiraacento(SGPart.Cells[6,varlinhas]));
        if not DM_Tabelas.ZQCidade.Locate('nomecid',varcid,[]) Then Begin
          DM_Tabelas.ZQCidade.Insert;
          DM_Tabelas.ZQCidade.FieldByName('nomecid').AsString := varcid;
          DM_Tabelas.ZQCidade.FieldByName('estado').AsString := SGPart.Cells[11,varlinhas];
          DM_Tabelas.ZQCidade.Post;
        end;
        DM_TAbelas.ZQresponsavel.FieldByName('cidade_cart').AsInteger := DM_Tabelas.ZQCidade.FieldByName('idcidade').AsLargeInt;
        DM_Tabelas.ZQresponsavel.FieldByName('comarca').AsInteger := DM_Tabelas.ZQCidade.FieldByName('idcidade').AsLargeInt;
        DM_Tabelas.ZQresponsavel.FieldByName('assina').AsString := 'S';
        DM_Tabelas.ZQresponsavel.Post;
        DM_Tabelas.ZQParticipante.Delete;
      end;
    end;
  end;
  Label3.Caption := 'Guardando os detalhes dos conjuges...';
  BarraVertical.Value := 5;
  XlsToStringGrid(SGPart,ExtractFilePath( Application.ExeName )+'conjuge_complemento.xls');
  Gauge2.MaxValue := SGPart.RowCount-1;
  for varlinhas:=1 to SGPart.RowCount-1 do Begin
    varind := SGPart.Cells[1,varlinhas];
    if not empty(varind) Then Begin
      Gauge2.Progress := varlinhas;
      Label4.Caption := inttostr(varlinhas);
      if DM_Tabelas.ZQresponsavel.Locate('paticipante_idpaticipante',strtoint(SGPart.Cells[10,varlinhas]),[]) Then Begin
        varcid := UpperCase(tiraacento(SGPart.Cells[9,varlinhas]));
        if not DM_Tabelas.ZQCidade.Locate('nomecid',varcid,[]) Then Begin
          DM_Tabelas.ZQCidade.Insert;
          DM_Tabelas.ZQCidade.FieldByName('nomecid').AsString := varcid;
          DM_Tabelas.ZQCidade.FieldByName('estado').AsString := SGPart.Cells[8,varlinhas];
          DM_Tabelas.ZQCidade.Post;
        end;
        DM_Tabelas.ZQresponsavel.Edit;
        DM_Tabelas.ZQresponsavel.FieldByName('cid_natural').AsLargeInt := DM_Tabelas.ZQCidade.FieldByName('idcidade').AsLargeInt;
        DM_Tabelas.ZQresponsavel.FieldByName('profiss').AsString := SGPart.Cells[3,varlinhas];
        DM_Tabelas.ZQresponsavel.Post;
      end;
    end;
  end;
  Label3.Caption := 'Guardando os telefones dos clientes...';
  BarraVertical.Value := 6;
  XlsToStringGrid(SGPart,ExtractFilePath( Application.ExeName )+'telcli.xls');
  Gauge2.MaxValue := SGPart.RowCount-1;
  for varlinhas:=1 to SGPart.RowCount-1 do Begin
    varind := SGPart.Cells[1,varlinhas];
    if not empty(varind) Then Begin
      Gauge2.Progress := varlinhas;
      Label4.Caption := inttostr(varlinhas);
      DM_Tabelas.ZqParticipante.Locate('idpaticipante',strtoint(SGPart.Cells[0,varlinhas]),[]);
      DM_Tabelas.ZqParticipante.Edit;
      if SGPart.Cells[2,varlinhas]='TELRESIDENCIAL' Then
        DM_tabelas.ZqParticipante.FieldByName('Fone1').AsString := '('+SGPart.Cells[3,varlinhas]+')'+SGPart.Cells[4,varlinhas]
      else if SGPart.Cells[2,varlinhas]='TELCELULAR' Then
        DM_tabelas.ZqParticipante.FieldByName('fone2').AsString := '('+SGPart.Cells[3,varlinhas]+')'+SGPart.Cells[4,varlinhas]
      else if SGPart.Cells[2,varlinhas]='TELCOMERCIAL' Then
        DM_tabelas.ZqParticipante.FieldByName('fone3').AsString := '('+SGPart.Cells[3,varlinhas]+')'+SGPart.Cells[4,varlinhas]
      else
        DM_tabelas.ZqParticipante.FieldByName('fone4').AsString := '('+SGPart.Cells[3,varlinhas]+')'+SGPart.Cells[4,varlinhas];
      DM_Tabelas.ZqParticipante.Post;
    end;
  end;
  Label3.Caption := 'Guardando endereços dos clientes...';
  BarraVertical.Value := 7;
  XlsToStringGrid(SGPart,ExtractFilePath( Application.ExeName )+'enderecocli.xls');
  Gauge2.MaxValue := SGPart.RowCount-1;
  for varlinhas:=1 to SGPart.RowCount-1 do Begin
    varind := SGPart.Cells[1,varlinhas];
    if not empty(varind) Then Begin
      Gauge2.Progress := varlinhas;
      Label4.Caption := inttostr(varlinhas);
      varcid := UpperCase(tiraacento(SGPart.Cells[9,varlinhas]));
      if not DM_Tabelas.ZQCidade.Locate('nomecid',varcid,[]) Then Begin
        DM_Tabelas.ZQCidade.Insert;
        DM_Tabelas.ZQCidade.FieldByName('nomecid').AsString := varcid;
        DM_Tabelas.ZQCidade.FieldByName('estado').AsString := SGPart.Cells[8,varlinhas];
        DM_Tabelas.ZQCidade.Post;
      end;
      varcep := tiramascara(SGPart.Cells[10,varlinhas]);
      varcep := copy(varcep,1,2)+'.'+copy(varcep,3,3)+'-'+copy(varcep,6,3);
      DM_Tabelas.ZqParticipante.Locate('idpaticipante',strtoint(SGPart.Cells[0,varlinhas]),[]);
      DM_Tabelas.ZqParticipante.Edit;
      DM_Tabelas.ZqParticipante.FieldByName('endereco').AsString := copy(SGPart.Cells[3,varlinhas],1,3)+'. '+SGPart.Cells[4,varlinhas]+', '+SGPart.Cells[5,varlinhas];
      DM_Tabelas.ZqParticipante.FieldByName('complemento').AsString := SGPart.Cells[6,varlinhas];
      DM_Tabelas.ZqParticipante.FieldByName('bairro').AsString := SGPart.Cells[7,varlinhas];
      DM_Tabelas.ZqParticipante.FieldByName('cep').AsString := varcep;
      DM_Tabelas.ZqParticipante.FieldByName('cidade').AsInteger := DM_Tabelas.ZQCidade.FieldByName('idcidade').AsLargeInt;
      DM_Tabelas.ZqParticipante.FieldByName('ende_cob').AsString := DM_Tabelas.ZqParticipante.FieldByName('endereco').AsString;
      DM_Tabelas.ZqParticipante.FieldByName('complemento_cob').AsString := DM_Tabelas.ZqParticipante.FieldByName('complemento').AsString;
      DM_Tabelas.ZqParticipante.FieldByName('bairro_cob').AsString := DM_Tabelas.ZqParticipante.FieldByName('bairro').AsString;
      DM_Tabelas.ZqParticipante.FieldByName('cep_cob').AsString := DM_Tabelas.ZqParticipante.FieldByName('cep').AsString;
      DM_Tabelas.ZqParticipante.FieldByName('cidade_cob').AsLargeInt := DM_Tabelas.ZqParticipante.FieldByName('cidade').AsInteger;

      DM_Tabelas.ZqParticipante.Post;
    end;
  end;
  Label3.Caption := 'Guardando as contas bancarias...';
  BarraVertical.Value := 8;
  XlsToStringGrid(SGPart,ExtractFilePath( Application.ExeName )+'contabanco.xls');
  Gauge2.MaxValue := SGPart.RowCount-1;
  for varlinhas:=1 to SGPart.RowCount-1 do Begin
    varind := SGPart.Cells[1,varlinhas];
    if not empty(varind) Then Begin
      Gauge2.Progress := varlinhas;
      Label4.Caption := inttostr(varlinhas);
      DM_Tabelas.ZQContaBancaria.Insert;
      DM_Tabelas.ZQContaBancaria.FieldByName('n_banco').AsString := SGPart.Cells[2,varlinhas];
      DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia').AsString := SGPart.Cells[7,varlinhas];
      DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString := SGPart.Cells[8,varlinhas];
      DM_Tabelas.ZQContaBancaria.FieldByName('n_conta').AsString := SGPart.Cells[9,varlinhas];
      DM_tabelas.ZQContaBancaria.FieldByName('n_conta_v').AsString := SGPart.Cells[10,varlinhas];
      DM_Tabelas.ZQContaBancaria.FieldByName('nomebanco').AsString := 'BRADESCO';
      DM_Tabelas.ZQContaBancaria.FieldByName('n_no_banco').AsString := SGPart.Cells[24,varlinhas];
      DM_Tabelas.ZQContaBancaria.FieldByName('ativa').AsString := 'SIM';
      DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat := strtofloat(SGPart.Cells[17,varlinhas]);
      DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat := strtofloat(SGPart.Cells[16,varlinhas]);
      DM_Tabelas.ZQContaBancaria.FieldByName('valoroupercent').AsString := 'P';
      DM_Tabelas.ZQContaBancaria.FieldByName('carteira').AsString := '009';
      DM_Tabelas.ZQContaBancaria.FieldByName('boleto').AsString := 'BRADESCO_400';
      DM_Tabelas.ZQContaBancaria.Post;
    end;
  end;
  Label3.Caption := 'Guardando os loteamentos...';
  BarraVertical.Value := 9;
  XlsToStringGrid(SGPart,ExtractFilePath( Application.ExeName )+'loteamento.xls');
  Gauge2.MaxValue := SGPart.RowCount-1;
  for varlinhas:=1 to SGPart.RowCount-1 do Begin
    varind := SGPart.Cells[1,varlinhas];
    if not empty(varind) Then Begin
      Gauge2.Progress := varlinhas;
      Label4.Caption := inttostr(varlinhas);
      varcid := UpperCase(tiraacento(SGPart.Cells[4,varlinhas]));
      DM_Tabelas.ZQCidade.Locate('nomecid',varcid,[]);
      DM_Tabelas.ZQLoteamento.Insert;
      DM_Tabelas.ZQLoteamento.FieldByName('idloteamento').AsLargeInt := strtoint(SGPart.Cells[0,varlinhas]);
      DM_Tabelas.ZQLoteamento.FieldByName('matriculaloteamento').AsString := SGPart.Cells[3,varlinhas];
      DM_Tabelas.ZQLoteamento.FieldByName('nomeloteamento').AsString := uppercase(SGPart.Cells[1,varlinhas]);
      DM_Tabelas.ZQLoteamento.FieldByName('apelido').AsString := UpperCase(SGPart.Cells[2,varlinhas]);
      DM_Tabelas.ZQLoteamento.FieldByName('cidade_idcidade').AsLargeInt := DM_Tabelas.ZQCidade.FieldByName('idcidade').AsLargeInt;
      DM_Tabelas.ZQLoteamento.FieldByName('testemunha1').AsString := UpperCase(SGPart.Cells[8,varlinhas]);
      DM_Tabelas.ZQLoteamento.FieldByName('testemunha2').AsString := UpperCase(SGPart.Cells[9,varlinhas]);
      DM_Tabelas.ZQLoteamento.Post;
    end;
  end;
  Label3.Caption := 'Guardando os incorporadores...';
  BarraVertical.Value := 10;
  XlsToStringGrid(SGPart,ExtractFilePath( Application.ExeName )+'incorporadores.xls');
  Gauge2.MaxValue := SGPart.RowCount-1;
  for varlinhas:=1 to SGPart.RowCount-1 do Begin
    varind := SGPart.Cells[1,varlinhas];
    if not empty(varind) Then Begin
      Gauge2.Progress := varlinhas;
      Label4.Caption := inttostr(varlinhas);
      if not DM_Tabelas.ZqParticipante.Locate('doc1',SGPart.Cells[1,varlinhas],[]) then Begin
        DM_Tabelas.ZqParticipante.Insert;
        DM_Tabelas.ZqParticipante.FieldByName('doc1').AsString := SGPart.Cells[1,varlinhas];
        DM_Tabelas.ZQParticipante.FieldByName('nome_parte').AsString := SGPart.Cells[0,varlinhas];
        DM_Tabelas.ZqParticipante.Post;
      end;
      DM_Tabelas.ZQincorp_loteame.Insert;
      DM_Tabelas.ZQincorp_loteame.FieldByName('incorporador_idincorporador').AsLargeInt := DM_Tabelas.ZQParticipante.FieldByName('idpaticipante').AsLargeInt;
      DM_Tabelas.ZQincorp_loteame.FieldByName('loteamento_idloteamento').AsLargeInt := strtoint(SGPart.Cells[2,varlinhas]);
      DM_Tabelas.ZQincorp_loteame.Post;
    end;
  end;
  Label3.Caption := 'Guardando as profissões...';
  BarraVertical.Value := 11;
  XlsToStringGrid(SGPart,ExtractFilePath( Application.ExeName )+'profissao.xls');
  Gauge2.MaxValue := SGPart.RowCount-1;
  for varlinhas:=1 to SGPart.RowCount-1 do Begin
    varind := SGPart.Cells[1,varlinhas];
    if not empty(varind) Then Begin
      Gauge2.Progress := varlinhas;
      Label4.Caption := inttostr(varlinhas);
      DM_Tabelas.ZQProfissao.Insert;
      DM_Tabelas.ZQProfissao.FieldByName('profissao').AsString := SGPart.Cells[0,varlinhas];
      DM_Tabelas.ZQProfissao.Post;
    end;
  end;

end;

procedure TFrm_Transferencia.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Frm_Transferencia:=nil;
  Action:=CaFree;
end;


procedure TFrm_Transferencia.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
