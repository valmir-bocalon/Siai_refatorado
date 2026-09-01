unit hormain2;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, Db, DBTables, XBanner,   dxButton,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, DBClient, Grids, DBGrids,
  Mask, ComCtrls, dxCore2;

type
  TfrmMain2 = class(TForm)
    SizeCbo: TComboBox;
    Label1: TLabel;
    Label6: TLabel;
    SkipEdt: TEdit;
    Label2: TLabel;
    SkipEdtn: TEdit;
    XBanner1: TXBanner;
    btnPreview: TdxButton;
    btnPrint: TdxButton;
    Button1: TdxButton;
    DS_Participante: TDataSource;
    ZqParticipante: TZQuery;







































    DS_etq: TDataSource;
    GBAdversa: TGroupBox;
    XBanner5: TXBanner;
    DBGrid2: TDBGrid;
    Label3: TLabel;
    Eparticipante: TMaskEdit;
    GroupBox1: TGroupBox;
    XBanner2: TXBanner;
    DBGrid1: TDBGrid;
    dxButton1: TdxButton;
    DS_Empreend: TDataSource;
    bar1: TProgressBar;
    DXBMarcEmpree: TdxButton;
    BtRelatorio: TdxButton;
    DataCDS_etq2: TDataSource;

    Label12: TLabel;
    dbinativo: TComboBox;
    Label13: TLabel;
    DBMES: TComboBox;
    Label11: TLabel;
    dbprofissao: TComboBox;
    Label5: TLabel;
    Button2: TButton;
    Label7: TLabel;
    bsc: TMaskEdit;
    Label4: TLabel;
    CB_marcar: TCheckBox;
    dxButton2: TdxButton;
    Timer1: TTimer;
    Label8: TLabel;
    CDS_etq: TClientDataSet;
    CDS_etqnome_parte: TStringField;
    CDS_etqdoc1: TStringField;
    CDS_etqdoc2: TStringField;
    CDS_etqendereco: TStringField;
    CDS_etqbairro: TStringField;
    CDS_etqcep: TStringField;
    CDS_etqende_cob: TStringField;
    CDS_etqbairro_cob: TStringField;
    CDS_etqcep_cob: TStringField;
    CDS_etqFone1: TStringField;
    CDS_etqfone2: TStringField;
    CDS_etqfone3: TStringField;
    CDS_etqaniversario: TDateField;
    CDS_etqESTADOCIVIL: TStringField;
    CDS_etqrenda: TFloatField;
    CDS_etqobservacao: TMemoField;
    CDS_etqfone4: TStringField;
    CDS_etqcomplemento: TStringField;
    CDS_etqcomplemento_cob: TStringField;
    CDS_etqcadastrado: TDateField;
    CDS_etqNome_Firma: TStringField;
    CDS_etqnomecidade: TStringField;
    CDS_etqnomeestado: TStringField;
    CDS_etqnomecidadecob: TStringField;
    CDS_etqnomeestadocob: TStringField;
    CDS_etqnomecidadenatural: TStringField;
    CDS_etqnomeestadonatural: TStringField;
    CDS_etqInativo: TStringField;
    CDS_etqprofissao: TStringField;
    CDS_etqMes: TStringField;
    CDS_etqMarcar: TBooleanField;
    CDS_etq2: TClientDataSet;
    CDS_etq2nome_parte: TStringField;
    CDS_etq2doc1: TStringField;
    CDS_etq2doc2: TStringField;
    CDS_etq2endereco: TStringField;
    CDS_etq2bairro: TStringField;
    CDS_etq2cep: TStringField;
    CDS_etq2ende_cob: TStringField;
    CDS_etq2bairro_cob: TStringField;
    CDS_etq2cep_cob: TStringField;
    CDS_etq2Fone1: TStringField;
    CDS_etq2fone2: TStringField;
    CDS_etq2fone3: TStringField;
    CDS_etq2aniversario: TDateField;
    CDS_etq2ESTADOCIVIL: TStringField;
    CDS_etq2renda: TFloatField;
    CDS_etq2observacao: TMemoField;
    CDS_etq2fone4: TStringField;
    CDS_etq2complemento: TStringField;
    CDS_etq2complemento_cob: TStringField;
    CDS_etq2cadastrado: TDateField;
    CDS_etq2Nome_Firma: TStringField;
    CDS_etq2nomecidade: TStringField;
    CDS_etq2nomeestado: TStringField;
    CDS_etq2nomecidadecob: TStringField;
    CDS_etq2nomeestadocob: TStringField;
    CDS_etq2nomecidadenatural: TStringField;
    CDS_etq2nomeestadonatural: TStringField;
    CDS_etq2Inativo: TStringField;
    CDS_etq2profissao: TStringField;
    CDS_etq2Mes: TStringField;
    CDS_etq2Marcar: TBooleanField;
    CDSEmpreend: TClientDataSet;
    CDSEmpreendidpaticipante: TLargeintField;
    CDSEmpreendnome_parte: TStringField;
    CDSEmpreenddoc1: TStringField;
    CDSEmpreenddoc2: TStringField;
    CDSEmpreendendereco: TStringField;
    CDSEmpreendbairro: TStringField;
    CDSEmpreendcidade: TIntegerField;
    CDSEmpreendcep: TStringField;
    CDSEmpreendende_cob: TStringField;
    CDSEmpreendbairro_cob: TStringField;
    CDSEmpreendcidade_cob: TLargeintField;
    CDSEmpreendcep_cob: TStringField;
    CDSEmpreendFone1: TStringField;
    CDSEmpreendfone2: TStringField;
    CDSEmpreendfone3: TStringField;
    CDSEmpreendtipopessoa: TStringField;
    CDSEmpreendaniversario: TDateField;
    CDSEmpreendnacionalidade: TStringField;
    CDSEmpreendemail: TStringField;
    CDSEmpreendprofissao: TStringField;
    CDSEmpreendfone4: TStringField;
    CDSEmpreendInativo: TStringField;
    CDSEmpreendmarcar: TIntegerField;
    CDSEmpreendnomecidade: TStringField;
    CDSEmpreendnomeestado: TStringField;
    CDSEmpreendnaturalidade: TIntegerField;
    CDSEmpreendlocaldetrab: TStringField;
    CDSEmpreendESTADOCIVIL: TStringField;
    CDSEmpreendrenda: TFloatField;
    CDSEmpreendobservacao: TMemoField;
    CDSEmpreendcodpaginc: TIntegerField;
    CDSEmpreendcodrecinc: TIntegerField;
    CDSEmpreendcodpagBx: TIntegerField;
    CDSEmpreendcodrecBx: TIntegerField;
    CDSEmpreendcomplemento: TStringField;
    CDSEmpreendcomplemento_cob: TStringField;
    CDSEmpreendcadastrado: TDateField;
    CDSEmpreendNome_Firma: TStringField;
    CDSEmpreendnomecidadecob: TStringField;
    CDSEmpreendnomeestadocob: TStringField;
    ProgressBar1: TProgressBar;
    Label9: TLabel;
    bsc0: TMaskEdit;
    Panel1: TPanel;
    procedure btnPreviewClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure SkipEdtExit(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure SkipEdtnExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure EparticipanteChange(Sender: TObject);
    procedure DBGrid2DblClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure dxButton1Click(Sender: TObject);
    procedure EparticipanteKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGrid1KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DXBMarcEmpreeClick(Sender: TObject);
    procedure DBGrid2DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure mudamarcaEmpreend;      
    procedure DBGrid2KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGrid2MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure BtRelatorioClick(Sender: TObject);
    procedure dbprofissaoDropDown(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure DBGrid1CellClick(Column: TColumn);
    procedure DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGrid1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure bscChange(Sender: TObject);
    procedure CB_marcarClick(Sender: TObject);
    procedure bscEnter(Sender: TObject);
    procedure dxButton2Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure bsc0Enter(Sender: TObject);
    procedure bsc0Change(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  frmMain2: TfrmMain2;

implementation

uses horLabel2,tabelas, principal,funcoes, uRuntimeFields;

{$R *.DFM}

procedure TfrmMain2.btnPreviewClick(Sender: TObject);
begin
  try
   CDS_etq.Filtered:=false;
   CDS_etq.Filter:='marcar=''TRUE''';
   CDS_etq.Filtered:=true;
   CDS_etq.IndexName:='nome';
   CDS_etq.DisableControls;
   Panel1.Visible:=true;
   frmMain2.FormStyle:=fsNormal;
   if frmHorLabels2=nil then
      frmHorLabels2:=TfrmHorLabels2.Create(Application);
   with TfrmHorLabels2.Create(Application) do
   begin
      case SizeCbo.ItemIndex of
         0: LabelSize := h5160;
         1: LabelSize := h5161;
      end;
      SkipRow := StrToInt(SkipEdt.Text);

//      QuickRep1.OnPreview:=Frm_principal.Prev;

      if sender = btnPreview then
         QuickRep1.PreviewModal
      else
         QuickRep1.Print;
      Free;
   end;
  finally
    frmHorLabels2:=nil;
    CDS_etq.Filtered:=false;
    CDS_etq.EnableControls;
    Panel1.Visible:=false;
    frmMain2.FormStyle:=fsStayOnTop;
  end;

end;

procedure TfrmMain2.FormCreate(Sender: TObject);
begin
  EnsureRuntimeFields(Self);
  SizeCbo.ItemIndex := 0;
end;

procedure TfrmMain2.SkipEdtExit(Sender: TObject);
  function IsInteger(const s: string): boolean;
  var
    E: Integer;
    NotUsed: integer;
  begin
    Val(S, NotUsed, E);
    result := E = 0;
  end;
begin
  if not IsInteger(SkipEdt.Text) then
    SkipEdt.Text := '0';
end;

procedure TfrmMain2.Button1Click(Sender: TObject);
begin
  close;
end;

procedure TfrmMain2.SkipEdtnExit(Sender: TObject);
var
ne:integer;
begin
  ne:=1;
 { Adicionando os campos no memo em branco}

 if strtoint(alltrim(SkipEdtn.text))>0 then
 begin
   SkipEdt.Text:='0';
   while ne<=strtoint(alltrim(SkipEdtn.text)) do
   begin
     CDS_etq.Insert;
     CDS_etqnome_parte.Value:='';
     CDS_etq.post;
     inc(ne);
    end;
 end;
 SkipEdtn.Text:='0';
end;

procedure TfrmMain2.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  zqParticipante.close;
  CDS_etq.close;
  CDSEmpreend.close;
  frmMain2:=nil;
  action:=CaFree;
end;

procedure TfrmMain2.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key=#13 then
  begin
    Key:=#0;
    Perform(Wm_NextDlgCtl,0,0);
  end;
end;

procedure TfrmMain2.EparticipanteChange(Sender: TObject);
begin
  ZqParticipante.close;
  zqParticipante.SQL.Clear;
  ZqParticipante.SQL.Add('Select  idpaticipante,nome_parte,doc1,doc2,endereco,bairro,cidade,cep,ende_cob,bairro_cob,cidade_cob,cep_cob,Fone1,fone2,fone3,tipopessoa,aniversario,nacionalidade,');
  ZqParticipante.SQL.Add('        email,naturalidade,localdetrab,profissao,ESTADOCIVIL,renda,observacao,codpaginc,codrecinc,codpagBx,codrecBx,fone4,complemento,complemento_cob,cadastrado,Nome_Firma,');
  ZqParticipante.SQL.Add('        Inativo,sexo,cc,cex,jb,judicial,xquadra,xlote,plantao,historico ');
  zqParticipante.SQL.Add(' from participante');
  zqParticipante.SQL.Add(' where nome_parte like '+quotedstr('%'+Eparticipante.Text+'%'));
  zqParticipante.SQL.Add(' or endereco like '+quotedstr('%'+Eparticipante.Text+'%'));
  zqParticipante.SQL.Add(' or bairro like '+quotedstr('%'+Eparticipante.Text+'%'));
  zqParticipante.SQL.Add(' or cidade like '+quotedstr('%'+Eparticipante.Text));
  zqParticipante.SQL.Add(' order by nome_parte');
  zqParticipante.Open;

end;

procedure TfrmMain2.DBGrid2DblClick(Sender: TObject);
begin
{  CDS_etq.Insert;
  CDS_etqnome_parte.Value:=ZqParticipantenome_parte.Value;
  CDS_etqende_cob.Value:=ZqParticipanteende_cob.Value;
  CDS_etqbairro_cob.Value:=ZqParticipantebairro_cob.Value;
  CDS_etqcep_cob.Value:=ZqParticipantecep_cob.Value;
  CDS_etqnomecidadecob.Value:=ZqParticipantenomecidadecob.Value;
  CDS_etqnomeestadocob.Value:=ZqParticipantenomeestadocob.Value;
  CDS_etq.post;
  Eparticipante.clear;}
end;

procedure TfrmMain2.FormShow(Sender: TObject);
begin
  ZqParticipante.close;
  ZqParticipante.open;
  CDS_etq.close;
  CDS_etq.CreateDataSet;
  ZqParticipante.First;
  CDSEmpreend.Close;
  CDSEmpreend.CreateDataSet;
end;

procedure TfrmMain2.dxButton1Click(Sender: TObject);
begin
  CDS_etq.First;
  while not CDS_etq.Eof do
  begin
    CDS_etq.Delete;
  end;  
end;

procedure TfrmMain2.EparticipanteKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = VK_ESCAPE then
     Eparticipante.clear;
end;

procedure TfrmMain2.DBGrid1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = VK_DELETE then
     CDS_etq.Delete;
end;

procedure TfrmMain2.DXBMarcEmpreeClick(Sender: TObject);
begin
  if DXBMarcEmpree.Caption = 'Desmarcar' Then Begin
    CDSEmpreend.First;
    CDSEmpreend.DisableControls;
    while not CDSEmpreend.Eof do begin
      CDSEmpreend.Edit;
      CDSEmpreendmarcar.Value := 1;
      CDSEmpreend.Post;
      CDSEmpreend.Next;
    end;
    DXBMarcEmpree.Caption := 'Marcar';
  end
  Else
  Begin
    CDSEmpreend.First;
    CDSEmpreend.DisableControls;
    while not CDSEmpreend.Eof do begin
      CDSEmpreend.Edit;
      CDSEmpreendmarcar.Value := 0;
      CDSEmpreend.Post;
      CDSEmpreend.Next;
    end;
    DXBMarcEmpree.Caption := 'Desmarcar';
  End;
  CDSEmpreend.EnableControls;
end;

procedure TfrmMain2.DBGrid2DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  DBGrid2.DefaultDrawDataCell(Rect, DBGrid2.columns[datacol].field, State);
  if Column.Field = CDSEmpreendmarcar then begin
    DBGrid2.Canvas.FillRect(Rect);
    if CDSEmpreendmarcar.Value = 0 then
      DM_Tabelas.ImageList1.Draw(DBGrid2.Canvas, Rect.Left + 15,Rect.Top + 1, 0)
    else
      DM_Tabelas.ImageList1.Draw(DBGrid2.Canvas, Rect.Left + 15,Rect.Top + 1, 1);
  end;
end;

procedure TfrmMain2.mudamarcaEmpreend;
Begin
  if DBGrid2.SelectedField.FieldName = 'marcar' Then Begin
    CDSEmpreend.Edit;
    if CDSEmpreendmarcar.Value = 0 Then
      CDSEmpreendmarcar.Value := 1
    else
      CDSEmpreendmarcar.Value := 0;
    CDSEmpreend.Post;
  end;
End;

procedure TfrmMain2.DBGrid2KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = 13 Then
    mudamarcaEmpreend;
end;

procedure TfrmMain2.DBGrid2MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  mudamarcaEmpreend;
end;

procedure TfrmMain2.BtRelatorioClick(Sender: TObject);
var
  mes, nom: string;
  MapFields: array of record
    Src, Dst: TField;
  end;
  i: Integer;

  procedure MontarMapaCampos(Origem, Destino: TDataSet);
  var
    j: Integer;
  begin
    SetLength(MapFields, 0);

    for j := 0 to Origem.FieldCount - 1 do
      if Destino.FindField(Origem.Fields[j].FieldName) <> nil then
      begin
        SetLength(MapFields, Length(MapFields) + 1);
        MapFields[High(MapFields)].Src := Origem.Fields[j];
        MapFields[High(MapFields)].Dst :=
          Destino.FieldByName(Origem.Fields[j].FieldName);
      end;
  end;

  procedure CopiarRegistro;
  var
    k: Integer;
  begin
    for k := 0 to High(MapFields) do
      MapFields[k].Dst.Value := MapFields[k].Src.Value;
  end;

begin
  CDSEmpreend.DisableControls;
  CDS_etq2.DisableControls;
  CDS_etq.DisableControls;

  try
    // 🔹 FILTRO
    CDSEmpreend.Filtered := False;
    CDSEmpreend.Filter := 'marcar=0';
    CDSEmpreend.Filtered := True;

    CDS_etq2.Close;
    CDS_etq2.CreateDataSet;
    CDS_etq2.Open;

    if CDSEmpreend.RecordCount = 0 then
    begin
      mensagem('Nenhum Cliente Selecionado!');
      Exit;
    end;

    CDSEmpreend.First;

    ProgressBar1.Visible := True;
    ProgressBar1.Max := CDSEmpreend.RecordCount;
    ProgressBar1.Position := 0;

    // =========================
    // 🔥 LOOP 1 (geração CDS_etq2)
    // =========================
    while not CDSEmpreend.Eof do
    begin
      if (CDSEmpreend.RecNo mod 50 = 0) then
        ProgressBar1.Position := CDSEmpreend.RecNo;

      CDS_etq2.Append;

      CDS_etq2nome_parte.Value := UpperCase(CDSEmpreendnome_parte.Value);

      if not Empty(CDSEmpreendende_cob.Value) then
        CDS_etq2endereco.Value := UpperCase(CDSEmpreendende_cob.Value)
      else
        CDS_etq2endereco.Value := UpperCase(CDSEmpreendendereco.Value);

      if not Empty(CDSEmpreendbairro_cob.Value) then
        CDS_etq2bairro.Value := UpperCase(CDSEmpreendbairro_cob.Value)
      else
        CDS_etq2bairro.Value := UpperCase(CDSEmpreendbairro.Value);

      if not Empty(CDSEmpreendcep_cob.Value) then
        CDS_etq2cep.Value := UpperCase(CDSEmpreendcep_cob.Value)
      else
        CDS_etq2cep.Value := UpperCase(CDSEmpreendcep.Value);

      if not Empty(CDSEmpreendnomecidadecob.Value) then
        CDS_etq2nomecidade.Value := UpperCase(CDSEmpreendnomecidadecob.Value)
      else
        CDS_etq2nomecidade.Value := UpperCase(CDSEmpreendnomecidade.Value);

      if not Empty(CDSEmpreendnomeestadocob.Value) then
        CDS_etq2nomeestado.Value := UpperCase(CDSEmpreendnomeestadocob.Value)
      else
        CDS_etq2nomeestado.Value := UpperCase(CDSEmpreendnomeestado.Value);

      CDS_etq2Inativo.Value := UpperCase(CDSEmpreendInativo.Value);
      CDS_etq2profissao.Value := UpperCase(CDSEmpreendprofissao.Value);

      if CDSEmpreendaniversario.Value > 0 then
      begin
        CDS_etq2aniversario.Value := CDSEmpreendaniversario.Value;
        mes := FormatDateTime('mm', CDSEmpreendaniversario.Value);
        CDS_etq2mes.Value := mes;
      end;

      CDS_etq2.Post;
      CDSEmpreend.Next;
    end;

    ProgressBar1.Visible := False;

    // 🔥 Ordena para remoção de duplicados mais eficiente
    CDS_etq2.IndexFieldNames := 'nome_parte';

    // =========================
    // 🔥 LOOP 2 (remove duplicados)
    // =========================
    CDS_etq.Close;
    CDS_etq.CreateDataSet;
    CDS_etq.Open;

    CDS_etq2.First;

    bar1.Max := CDS_etq2.RecordCount;
    bar1.Position := 0;

    // 🔥 cria mapa UMA vez
    MontarMapaCampos(CDS_etq2, CDS_etq);

    nom := CDS_etq2nome_parte.Value;

    CDS_etq.Append;
    CopiarRegistro;
    CDS_etqmarcar.AsBoolean := True;
    CDS_etq.Post;

    CDS_etq2.Next;

    while not CDS_etq2.Eof do
    begin
      if (CDS_etq2.RecNo mod 50 = 0) then
        bar1.Position := CDS_etq2.RecNo;

      if nom <> CDS_etq2nome_parte.Value then
      begin
        CDS_etq.Append;
        CopiarRegistro;
        CDS_etqmarcar.AsBoolean := True;
        CDS_etq.Post;

        nom := CDS_etq2nome_parte.Value;
      end;

      CDS_etq2.Next;
    end;

    CDS_etq2.Close;

    CDS_etq.First;
    bar1.Position := 0;

    Label4.Caption := IntToStr(CDS_etq.RecordCount) + ' Clientes Selecionados';

    CDSEmpreend.Filtered := False;
    CDSEmpreend.First;

  finally
    CDSEmpreend.EnableControls;
    CDS_etq2.EnableControls;
    CDS_etq.EnableControls;
  end;
end;
procedure TfrmMain2.dbprofissaoDropDown(Sender: TObject);
begin
  dbprofissao.Clear;
  DM_Tabelas.ZQProfissao.Open;
  DM_Tabelas.ZQProfissao.First;
  DM_Tabelas.ZQProfissao.DisableControls;
  ProgressBar1.Visible:=true;
  ProgressBar1.Position:=0;
  ProgressBar1.Max:=DM_Tabelas.ZQProfissao.RecordCount;
  while not DM_Tabelas.ZQProfissao.Eof do
  begin
    ProgressBar1.Position:=DM_Tabelas.ZQProfissao.RecNo;
    dbprofissao.Items.Add(uppercase(DM_Tabelas.ZQProfissao.FieldByName('profissao').AsString));
    DM_Tabelas.ZQProfissao.Next;
  end;
  DM_Tabelas.ZQProfissao.close;
  DM_Tabelas.ZQProfissao.EnableControls;
  ProgressBar1.Position:=0;
  ProgressBar1.Visible:=false;
end;

procedure TfrmMain2.Button2Click(Sender: TObject);
var
filtro:string;
begin
  CDS_etq.Filtered:=false;
  filtro:='';
  if not empty(dbinativo.Text) then
     filtro:=' and inativo='+quotedstr(dbinativo.Text);
  if not empty(dbprofissao.Text) then
     filtro:=filtro+' and profissao='+quotedstr(dbprofissao.Text);
  if not empty(DBMES.Text) then
     filtro:=filtro+' and Mes='+quotedstr(DBMES.Text);
  CDS_etq.Filter:='1=1 '+filtro;
  CDS_etq.Filtered:=true;
  Label4.Caption:=inttostr(CDS_etq.RecordCount)+' Clientes Selecionados';  
end;

procedure TfrmMain2.DBGrid1CellClick(Column: TColumn);
begin
    CDS_etq.Edit;
    if CDS_etqmarcar.AsBoolean = true then
      CDS_etqmarcar.AsBoolean := false
    else
      CDS_etqmarcar.AsBoolean := true;
    CDS_etq.Post
end;

procedure TfrmMain2.DBGrid1DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
var
  Check: Integer;
  R: TRect;
begin
  if Column.FieldName = 'Marcar' then //COLOQUE O NOME DA COLUNA, NO MEU CASO 'MARCAR'
  begin
    DBGrid1.Canvas.FillRect(Rect);
    Check := 0;
    if CDS_etqmarcar.AsBoolean = true then
      Check := DFCS_CHECKED
    else
      Check := 0;
    R:=Rect;
    InflateRect(R,-2,-2); {Diminue o tamanho do CheckBox}
    DrawFrameControl(DBGrid1.Canvas.Handle,R,DFC_BUTTON, DFCS_BUTTONCHECK or Check);
  end;
end;

procedure TfrmMain2.DBGrid1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  //bloqueia a tecla ctrl+del
  If ((Shift = [ssCtrl]) and (key = vk_delete)) then Abort;
end;

procedure TfrmMain2.bsc0Change(Sender: TObject);
begin
   CDSEmpreend.Locate('nome_parte',bsc0.Text,[loPartialkey]);
end;

procedure TfrmMain2.bsc0Enter(Sender: TObject);
begin
  bsc0.Clear;
end;

procedure TfrmMain2.bscChange(Sender: TObject);
begin
   CDS_etq.Locate('nome_parte',bsc.Text,[loPartialkey]);
end;

procedure TfrmMain2.CB_marcarClick(Sender: TObject);
begin
  if CB_marcar.Checked then
  begin
    CDS_etq.First;
    ProgressBar1.Visible:=true;
    ProgressBar1.Position:=0;
    ProgressBar1.Max:=CDS_etq.RecordCount;
    CDS_etq.DisableControls;
    while not CDS_etq.Eof do
    begin
      ProgressBar1.Position:=CDS_Etq.RecNo;
      CDS_etq.Edit;
      CDS_etqmarcar.AsBoolean := false;
      CDS_etq.Post;
      CDS_etq.Next;
    end;
    CDS_etq.First;
    ProgressBar1.Position:=0;
    ProgressBar1.Visible:=false;
    CDS_etq.EnableControls;
  end
  else
  begin
    CDS_etq.First;
    ProgressBar1.Visible:=true;
    ProgressBar1.Position:=0;
    ProgressBar1.Max:=CDS_etq.RecordCount;
    CDS_etq.DisableControls;

    while not CDS_etq.Eof do
    begin
      ProgressBar1.Position:=CDS_Etq.RecNo;
      CDS_etq.Edit;
      CDS_etqmarcar.AsBoolean := true;
      CDS_etq.Post;
      CDS_etq.Next;
    end;
    CDS_etq.First;
    ProgressBar1.Position:=0;
    ProgressBar1.Visible:=false;
    CDS_etq.EnableControls;

  end;

end;

procedure TfrmMain2.bscEnter(Sender: TObject);
begin
  bsc.Clear;
end;

procedure TfrmMain2.dxButton2Click(Sender: TObject);
begin
  bar1.Max := ZqParticipante.RecordCount;
  bar1.Position := 0;

  ZqParticipante.DisableControls;
  CDSEmpreend.DisableControls;

  CDSEmpreend.LogChanges := False; // evita overhead de delta

  try
    ZqParticipante.First;

    while not ZqParticipante.Eof do
    begin
      // Atualiza barra a cada X registros (melhora MUITO)
      if (ZqParticipante.RecNo mod 150 = 0) then
        bar1.Position := ZqParticipante.RecNo;

      CDSEmpreend.Append;

      CDSEmpreendmarcar.Value        := 1;
      CDSEmpreendnome_parte.Value    := ZQParticipante.FieldByName('nome_parte').AsString;
      CDSEmpreendende_cob.Value      := ZqParticipante.FieldByName('ende_cob').AsString;
      CDSEmpreendbairro_cob.Value    := ZqParticipante.FieldByName('bairro_cob').AsString;
      CDSEmpreendcep_cob.Value       := ZqParticipante.FieldByName('cep_cob').AsString;
      CDSEmpreendnomecidadecob.Value := ZqParticipante.FieldByName('nomecidadecob').AsString;
      CDSEmpreendnomeestadocob.Value := ZqParticipante.FieldByName('nomeestadocob').AsString;
      CDSEmpreendendereco.Value      := ZqParticipante.FieldByName('endereco').AsString;
      CDSEmpreendbairro.Value        := ZqParticipante.FieldByName('bairro').AsString;
      CDSEmpreendcep.Value           := ZqParticipante.FieldByName('cep').AsString;
      CDSEmpreendnomecidade.Value    := ZqParticipante.FieldByName('nomecidade').AsString;
      CDSEmpreendnomeestado.Value    := ZqParticipante.FieldByName('nomeestado').AsString;
      CDSEmpreendInativo.Value       := ZqParticipante.FieldByName('Inativo').AsString;
      CDSEmpreendaniversario.Value   := ZqParticipante.FieldByName('aniversario').AsDateTime;
      CDSEmpreendprofissao.Value     := ZqParticipante.FieldByName('profissao').AsString;

      CDSEmpreend.Post;

      ZqParticipante.Next;
    end;

    // Atualiza barra no final
    bar1.Position := bar1.Max;

  finally
    CDSEmpreend.EnableControls;
    ZqParticipante.EnableControls;
    CDSEmpreend.LogChanges := True;

    ZqParticipante.Close;

    bar1.Position := 0;
    Label8.Visible := False;
    Timer1.Enabled := False;
    CDSEmpreend.First;
  end;
end;

procedure TfrmMain2.Timer1Timer(Sender: TObject);
begin
  Label8.Visible:=true;
  application.ProcessMessages;
  dxButton2.OnClick(self);
  Timer1.Enabled:=false;
end;


procedure TfrmMain2.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TfrmMain2, 'ZqParticipante', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipanteidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantenome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantedoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantedoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipanteendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantebairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantecidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantecep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '00\.000\-999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipanteende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantebairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantecidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantecep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '00\.000\-999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipanteFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '\(99\)9999\-9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantefone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '\(99\)9999\-9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantefone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '\(99\)9999\-9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantetipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipanteaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantenacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipanteemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantenomecidade', 'nomecidade', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cidade', 'DM_Tabelas.ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantenomeestado', 'nomeestado', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade', 'DM_Tabelas.ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantenomecidadecob', 'nomecidadecob', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cidade_cob', 'DM_Tabelas.ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantenomeestadocob', 'nomeestadocob', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade_cob', 'DM_Tabelas.ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantenaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantelocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipanteprofissao', 'profissao', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantenomecidadenatural', 'nomecidadenatural', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'naturalidade', 'DM_Tabelas.ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantenomeestadonatural', 'nomeestadonatural', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'naturalidade', 'DM_Tabelas.ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipanteESTADOCIVIL', 'ESTADOCIVIL', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipanterenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantecodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantecodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantecodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantecodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantefone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '\(00\)0000\-0000;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantecomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantecomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipantecadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipanteNome_Firma', 'Nome_Firma', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain2, 'ZqParticipante', 'ZqParticipanteInativo', 'Inativo', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
