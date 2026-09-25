unit hormain;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, Db, DBTables, XBanner,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, DBClient, Grids, DBGrids,
  Mask, ComCtrls;

type
  TfrmMain = class(TForm)
    SizeCbo: TComboBox;
    Label1: TLabel;
    Label6: TLabel;
    SkipEdt: TEdit;
    Label2: TLabel;
    SkipEdtn: TEdit;
    XBanner1: TXBanner;
    btnPreview: TdxButtonArround;
    btnPrint: TdxButtonArround;
    Button1: TdxButtonArround;
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
    dxButton1: TdxButtonArround;
    DS_Empreend: TDataSource;
    bar1: TProgressBar;
    DXBMarcEmpree: TdxButtonArround;
    ZQRecebimento: TZQuery;
    DS_Recebimento: TDataSource;
    BtRelatorio: TdxButtonArround;
    ZQcomprador: TZQuery;
    DS_ZQcomprador: TDataSource;
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
    CDSEmpreendnomeempre: TStringField;
    CDSEmpreendcodempre: TIntegerField;
    CDSEmpreendmarca: TIntegerField;
    CDSEmpreendlogo: TBlobField;
    CDSEmpreendApelido: TStringField;
    ProgressBar1: TProgressBar;
    RGSelecione: TRadioGroup;
    Label8: TLabel;
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
    procedure bsc0Enter(Sender: TObject);
    procedure bsc0Change(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;


var
  frmMain: TfrmMain;

implementation

uses horLabel,tabelas, principal,funcoes,relparticipante2, uRuntimeFields;

{$R *.DFM}


procedure TfrmMain.btnPreviewClick(Sender: TObject);
begin
  if RGSelecione.ItemIndex=0 then
  begin
   try
     CDS_etq.Filtered:=false;
     CDS_etq.Filter:='marcar=''TRUE''';
     CDS_etq.Filtered:=true;
     CDS_etq.IndexName:='nome';
     CDS_etq.DisableControls;
     Panel1.Visible:=true;
     frmMain.FormStyle:=fsNormal;
     if frmHorLabels=nil then
        frmHorLabels:=TfrmHorLabels.Create(Application);
     with TfrmHorLabels.Create(Application) do
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
     frmHorLabels:=nil;
     CDS_etq.Filtered:=false;
     CDS_etq.EnableControls;
     Panel1.Visible:=false;
     frmMain.FormStyle:=fsStayOntop;
   end;
  end
  else
  begin
     CDS_etq.Filtered:=false;
     CDS_etq.Filter:='marcar=''TRUE''';
     CDS_etq.Filtered:=true;
     CDS_etq.IndexName:='nome';
     CDS_etq.DisableControls;
     Panel1.Visible:=true;
     try
      try
        frmMain.FormStyle:=fsNormal;
        if FrmRelparticipante2=nil then
           FrmRelparticipante2:=TFrmRelparticipante2.create(Application);
        FrmRelparticipante2.RLReport1.Previewmodal;
        FreeAndNil(FrmRelparticipante2);
      finally
        frmMain.FormStyle:=fsStayOnTop;

      end;
     finally
       CDS_etq.EnableControls;
       Panel1.Visible:=false;
     end;

  end;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  SizeCbo.ItemIndex := 0;
end;

procedure TfrmMain.SkipEdtExit(Sender: TObject);
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

procedure TfrmMain.Button1Click(Sender: TObject);
begin
  close;
end;

procedure TfrmMain.SkipEdtnExit(Sender: TObject);
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

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  zqParticipante.close;
  CDS_etq.close;
  ZQRecebimento.close;
  CDSEmpreend.close;
  frmMain:=nil;
  action:=CaFree;
end;

procedure TfrmMain.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key=#13 then
  begin
    Key:=#0;
    Perform(Wm_NextDlgCtl,0,0);
  end;
end;

procedure TfrmMain.EparticipanteChange(Sender: TObject);
begin
  ZqParticipante.close;
  zqParticipante.SQL.Clear;
  ZqParticipante.SQL.Add('Select  idpaticipante,nome_parte,doc1,doc2,endereco,bairro,cidade,cep,ende_cob,bairro_cob,cidade_cob,cep_cob,Fone1,fone2,fone3,tipopessoa,aniversario,nacionalidade,');
  ZqParticipante.SQL.Add('        email,naturalidade,localdetrab,profissao,ESTADOCIVIL,renda,observacao,codpaginc,codrecinc,codpagBx,codrecBx,fone4,complemento,complemento_cob,cadastrado,Nome_Firma,');
  ZqParticipante.SQL.Add('        Inativo,sexo,cc,cex,jb,judicial,xquadra,xlote,plantao,historico');
  zqParticipante.SQL.Add(' from participante');
  zqParticipante.SQL.Add(' where nome_parte like '+quotedstr('%'+Eparticipante.Text+'%'));
  zqParticipante.SQL.Add(' or endereco like '+quotedstr('%'+Eparticipante.Text+'%'));
  zqParticipante.SQL.Add(' or bairro like '+quotedstr('%'+Eparticipante.Text+'%'));
  zqParticipante.SQL.Add(' or cidade like '+quotedstr('%'+Eparticipante.Text));
  zqParticipante.SQL.Add(' order by nome_parte');
  zqParticipante.Open;

end;

procedure TfrmMain.DBGrid2DblClick(Sender: TObject);
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

procedure TfrmMain.FormShow(Sender: TObject);
begin
  ZqParticipante.close;
  ZqParticipante.open;
  CDS_etq.close;
  CDS_etq.CreateDataSet;
  DM_Tabelas.ZQLoteamento.open;
  DM_Tabelas.ZQLoteamento.First;
  CDSEmpreend.Close;
  CDSEmpreend.CreateDataSet;
  bar1.Max:=DM_Tabelas.ZQLoteamento.RecordCount;
  bar1.Position:=0;
  DM_Tabelas.ZQLoteamento.DisableControls;
  while not DM_Tabelas.ZQLoteamento.Eof do begin
    bar1.Position:=DM_Tabelas.ZQLoteamento.RecNo;
    CDSEmpreend.Insert;
    CDSEmpreendmarca.Value := 1;
    CDSEmpreendcodempre.Value := DM_Tabelas.ZQLoteamento.FieldByName('idloteamento').AsLargeInt;
    CDSEmpreendnomeempre.Value := DM_TAbelas.ZQLoteamento.FieldByName('nomeloteamento').AsString;
    CDSEmpreendlogo.Assign(DM_Tabelas.ZQLoteamento.FieldByName('logotipo'));
    CDSEmpreendApelido.Value := DM_Tabelas.ZQLoteamento.FieldByName('apelido').AsString;
    CDSEmpreend.Post;
    DM_Tabelas.ZQLoteamento.Next;
  end;
  DM_Tabelas.ZQLoteamento.EnableControls;
  DM_Tabelas.ZQLoteamento.close;
  bar1.Position:=0;
  ZQRecebimento.close;
end;

procedure TfrmMain.dxButton1Click(Sender: TObject);
begin
  CDS_etq.First;
  while not CDS_etq.Eof do
  begin
    CDS_etq.Delete;
  end;  
end;

procedure TfrmMain.EparticipanteKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = VK_ESCAPE then
     Eparticipante.clear;
end;

procedure TfrmMain.DBGrid1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = VK_DELETE then
     CDS_etq.Delete;
end;

procedure TfrmMain.DXBMarcEmpreeClick(Sender: TObject);
begin
  if DXBMarcEmpree.Caption = 'Desmarcar' Then Begin
    CDSEmpreend.First;
    CDSEmpreend.DisableControls;
    while not CDSEmpreend.Eof do begin
      CDSEmpreend.Edit;
      CDSEmpreendmarca.Value := 1;
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
      CDSEmpreendmarca.Value := 0;
      CDSEmpreend.Post;
      CDSEmpreend.Next;
    end;
    DXBMarcEmpree.Caption := 'Desmarcar';
  End;
  CDSEmpreend.EnableControls;
end;

procedure TfrmMain.DBGrid2DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  DBGrid2.DefaultDrawDataCell(Rect, DBGrid2.columns[datacol].field, State);
  if Column.Field = CDSEmpreendmarca then begin
    DBGrid2.Canvas.FillRect(Rect);
    if CDSEmpreendmarca.Value = 0 then
      DM_Tabelas.ImageList1.Draw(DBGrid2.Canvas, Rect.Left + 15,Rect.Top + 1, 0)
    else
      DM_Tabelas.ImageList1.Draw(DBGrid2.Canvas, Rect.Left + 15,Rect.Top + 1, 1);
  end;
end;

procedure TfrmMain.mudamarcaEmpreend;
Begin
  if DBGrid2.SelectedField.FieldName = 'marca' Then Begin
    CDSEmpreend.Edit;
    if CDSEmpreendmarca.Value = 0 Then
      CDSEmpreendmarca.Value := 1
    else
      CDSEmpreendmarca.Value := 0;
    CDSEmpreend.Post;
  end;
End;

procedure TfrmMain.DBGrid2KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = 13 Then
    mudamarcaEmpreend;
end;

procedure TfrmMain.DBGrid2MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  mudamarcaEmpreend;
end;

procedure TfrmMain.BtRelatorioClick(Sender: TObject);
var
  varlote: string;
  ZQRelatorio: TZQuery;
  ListaNomes: TStringList;
  mes: string;
  NomeCidadeLookup, EstadoLookup: string;
begin
  Screen.Cursor := crHourGlass;

  // [destaque] garante CDS abertos
  if not CDS_etq2.Active then
  begin
    CDS_etq2.Close;
    CDS_etq2.CreateDataSet;
    CDS_etq2.Open;
  end
  else
    CDS_etq2.EmptyDataSet;

  if not CDS_etq.Active then
  begin
    CDS_etq.Close;
    CDS_etq.CreateDataSet;
    CDS_etq.Open;
  end
  else
    CDS_etq.EmptyDataSet;

  try
    // [nota] monta lista de loteamentos
    varlote := '';
    CDSEmpreend.First;
    while not CDSEmpreend.Eof do
    begin
      if CDSEmpreendmarca.Value = 0 then
        varlote := varlote + QuotedStr(CDSEmpreendcodempre.AsString) + ',';

      CDSEmpreend.Next;
    end;

    if varlote = '' then
    begin
      mensagem('Nenhum Cliente Selecionado!');
      Exit;
    end;

    Delete(varlote, Length(varlote), 1);

    // [destaque] cria query dinâmica
    ZQRelatorio := TZQuery.Create(nil);
    ListaNomes := TStringList.Create;

    try
      ZQRelatorio.Connection := DM_Tabelas.zconeccao;

      ZQRelatorio.SQL.Text :=
        'SELECT DISTINCT '+
        'p.idpaticipante, '+
        'UPPER(p.nome_parte) AS nome_parte, '+

        'UPPER(CASE WHEN p.ende_cob <> '''' THEN p.ende_cob ELSE p.endereco END) AS endereco,'+
        'UPPER(CASE WHEN p.bairro_cob <> '''' THEN p.bairro_cob ELSE p.bairro END) AS bairro,'+
        'UPPER(CASE WHEN p.cep_cob <> '''' THEN p.cep_cob ELSE p.cep END) AS cep,'+

        'CASE WHEN p.cidade_cob IS NOT NULL THEN p.cidade_cob ELSE p.cidade END AS idcidade_lookup,'+

        'p.fone1, p.fone2, p.fone3,'+
        'UPPER(p.profissao) AS profissao,'+
        'UPPER(p.inativo) AS inativo,'+

        'p.aniversario,'+
        'DATE_FORMAT(p.aniversario, ''%m'') AS mes '+

        'FROM recebimento r '+
        'JOIN comprador c ON c.venda_idvenda = r.venda_idvenda '+
        'JOIN participante p ON p.idpaticipante = c.paticipante_idpaticipante '+


        'WHERE r.idloteamento IN ('+varlote+')';

      ZQRelatorio.Open;

      // [destaque] ProgressBar REAL
      ProgressBar1.Visible := True;
      ProgressBar1.Position := 0;
      ProgressBar1.Max := ZQRelatorio.RecordCount;

      ZQRelatorio.First;
      while not ZQRelatorio.Eof do
      begin
        ProgressBar1.Position := ZQRelatorio.RecNo;

        // [nota] CDS_etq2 (TODOS)
        NomeCidadeLookup := '';
        EstadoLookup := '';
        if not DM_Tabelas.ZQCidade.Active then
          DM_Tabelas.ZQCidade.Open;
        if not ZQRelatorio.FieldByName('idcidade_lookup').IsNull then
          if DM_Tabelas.ZQCidade.Locate('idcidade',
             ZQRelatorio.FieldByName('idcidade_lookup').AsLargeInt, []) then
          begin
            NomeCidadeLookup := UpperCase(DM_Tabelas.ZQCidade.FieldByName('nomecid').AsString);
            EstadoLookup := UpperCase(DM_Tabelas.ZQCidade.FieldByName('estado').AsString);
          end;
        CDS_etq2.Append;
        CDS_etq2nome_parte.Value    := ZQRelatorio.FieldByName('nome_parte').AsString;
        CDS_etq2endereco.Value      := ZQRelatorio.FieldByName('endereco').AsString;
        CDS_etq2bairro.Value        := ZQRelatorio.FieldByName('bairro').AsString;
        CDS_etq2cep.Value           := ZQRelatorio.FieldByName('cep').AsString;
        CDS_etq2nomecidade.Value    := NomeCidadeLookup;
        CDS_etq2nomeestado.Value    := EstadoLookup;
        CDS_etq2fone1.Value         := ZQRelatorio.FieldByName('fone1').AsString;
        CDS_etq2fone2.Value         := ZQRelatorio.FieldByName('fone2').AsString;
        CDS_etq2fone3.Value         := ZQRelatorio.FieldByName('fone3').AsString;
        CDS_etq2profissao.Value     := ZQRelatorio.FieldByName('profissao').AsString;
        if not ZQRelatorio.FieldByName('inativo').IsNull then
          CDS_etq2Inativo.Value     := ZQRelatorio.FieldByName('inativo').AsString
        else
          CDS_etq2Inativo.Value     := 'NÃO';

        if not ZQRelatorio.FieldByName('aniversario').IsNull then
        begin
          CDS_etq2aniversario.Value := ZQRelatorio.FieldByName('aniversario').AsDateTime;
          mes                       := ZQRelatorio.FieldByName('mes').AsString;
          CDS_etq2Mes.Value         := mes;
        end;

        CDS_etq2.Post;

        // [destaque] evita duplicados (rápido)
        if ListaNomes.IndexOf(ZQRelatorio.FieldByName('nome_parte').AsString) = -1 then
        begin
          ListaNomes.Add(ZQRelatorio.FieldByName('nome_parte').AsString);

          CDS_etq.Append;
          CDS_etqnome_parte.Value    := ZQRelatorio.FieldByName('nome_parte').AsString;
          CDS_etqendereco.Value      := ZQRelatorio.FieldByName('endereco').AsString;
          CDS_etqbairro.Value        := ZQRelatorio.FieldByName('bairro').AsString;
          CDS_etqcep.Value           := ZQRelatorio.FieldByName('cep').AsString;
          CDS_etqnomecidade.Value    := NomeCidadeLookup;
          CDS_etqnomeestado.Value    := EstadoLookup;
          CDS_etqfone1.Value         := ZQRelatorio.FieldByName('fone1').AsString;
          CDS_etqfone2.Value         := ZQRelatorio.FieldByName('fone2').AsString;
          CDS_etqfone3.Value         := ZQRelatorio.FieldByName('fone3').AsString;
          CDS_etqprofissao.Value     := ZQRelatorio.FieldByName('profissao').AsString;
          if not ZQRelatorio.FieldByName('inativo').IsNull then
            CDS_etqInativo.Value     := ZQRelatorio.FieldByName('inativo').AsString
          else
            CDS_etqInativo.Value     := 'NÃO';

          if not ZQRelatorio.FieldByName('aniversario').IsNull then
          begin
            CDS_etqaniversario.Value := ZQRelatorio.FieldByName('aniversario').AsDateTime;
            CDS_etqMes.Value         := ZQRelatorio.FieldByName('mes').AsString;
          end;

          CDS_etqMarcar.AsBoolean    := True;
          CDS_etq.Post;
        end;

        ZQRelatorio.Next;
      end;

      ProgressBar1.Visible := False;
      ProgressBar1.Position := 0;

      Label4.Caption := IntToStr(CDS_etq.RecordCount) + ' Clientes Selecionados';

    finally
      ListaNomes.Free;
      ZQRelatorio.Free;
    end;

  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmMain.dbprofissaoDropDown(Sender: TObject);
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

procedure TfrmMain.Button2Click(Sender: TObject);
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

procedure TfrmMain.DBGrid1CellClick(Column: TColumn);
begin
    CDS_etq.Edit;
    if CDS_etqmarcar.AsBoolean = true then
      CDS_etqmarcar.AsBoolean := false
    else
      CDS_etqmarcar.AsBoolean := true;
    CDS_etq.Post
end;

procedure TfrmMain.DBGrid1DrawColumnCell(Sender: TObject;
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

procedure TfrmMain.DBGrid1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  //bloqueia a tecla ctrl+del
  If ((Shift = [ssCtrl]) and (key = vk_delete)) then Abort;
end;

procedure TfrmMain.bsc0Change(Sender: TObject);
begin
  CDSEmpreend.Locate('Apelido',bsc0.Text,[loPartialkey]);
end;

procedure TfrmMain.bsc0Enter(Sender: TObject);
begin
  bsc0.Clear;
end;

procedure TfrmMain.bscChange(Sender: TObject);
begin
   CDS_etq.Locate('nome_parte',bsc.Text,[loPartialkey]);
end;

procedure TfrmMain.CB_marcarClick(Sender: TObject);
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

procedure TfrmMain.bscEnter(Sender: TObject);
begin
  bsc.Clear;
end;


procedure TfrmMain.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TfrmMain, 'ZqParticipante', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipanteidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantenome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantedoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantedoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipanteendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantebairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantecidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantecep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '00\.000\-999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipanteende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantebairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantecidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantecep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '00\.000\-999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipanteFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '\(99\)9999\-9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantefone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '\(99\)9999\-9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantefone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '\(99\)9999\-9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantetipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipanteaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantenacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipanteemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantenomecidade', 'nomecidade', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cidade', 'DM_Tabelas.ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantenomeestado', 'nomeestado', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade', 'DM_Tabelas.ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantenomecidadecob', 'nomecidadecob', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cidade_cob', 'DM_Tabelas.ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantenomeestadocob', 'nomeestadocob', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade_cob', 'DM_Tabelas.ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantenaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantelocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipanteprofissao', 'profissao', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantenomecidadenatural', 'nomecidadenatural', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'naturalidade', 'DM_Tabelas.ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantenomeestadonatural', 'nomeestadonatural', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'naturalidade', 'DM_Tabelas.ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipanteESTADOCIVIL', 'ESTADOCIVIL', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipanterenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantecodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantecodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantecodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantecodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantefone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '\(00\)0000\-0000;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantecomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantecomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipantecadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipanteNome_Firma', 'Nome_Firma', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZqParticipante', 'ZqParticipanteInativo', 'Inativo', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TfrmMain, 'ZQRecebimento', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentosaldocalc', 'saldocalc', TFloatField, fkCalculated, 0, 0, False, '', '#,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQRecebimento', 'ZQRecebimentoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TfrmMain, 'ZQcomprador', False);
  RegisterRuntimeField(TfrmMain, 'ZQcomprador', 'ZQcompradoridcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQcomprador', 'ZQcompradorpaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQcomprador', 'ZQcompradorvenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQcomprador', 'ZQcompradorpromissario', 'promissario', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQcomprador', 'ZQcompradorpercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmMain, 'ZQcomprador', 'ZQcompradorMarcar', 'Marcar', TSmallintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
