unit Configuracoes;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, TabNotBk, XBanner, StdCtrls, DBGrids,
  DBCtrls, XNum, ExtCtrls,DB;


type
  TFrm_Configuracoes = class(TForm)
    Label19: TLabel;
    XBanner14: TXBanner;
    TabbedNotebook1: TTabbedNotebook;
    Label1: TLabel;
    XBanner1: TXBanner;
    BtFechar: TdxButtonArround;
    EVendaPlano: TEdit;
    Label2: TLabel;
    ERecebePlano: TEdit;
    Label3: TLabel;
    taxa_price: TComboBox;
    Label4: TLabel;
    bxmanual: TComboBox;
    Label5: TLabel;
    Xmulta: TXNumEdit;
    Label6: TLabel;
    Xmora: TXNumEdit;
    Label7: TLabel;
    xreajuste: TComboBox;
    Label8: TLabel;
    Cmesmodia: TComboBox;
    Resolucao: TRadioGroup;
    procedure BtFecharClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure EVendaPlanoExit(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure ERecebePlanoExit(Sender: TObject);
    procedure taxa_priceExit(Sender: TObject);
    procedure bxmanualExit(Sender: TObject);
    procedure XmultaExit(Sender: TObject);
    procedure XmoraExit(Sender: TObject);
    procedure xreajusteExit(Sender: TObject);
    procedure ResolucaoExit(Sender: TObject);
    procedure CmesmodiaExit(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_Configuracoes: TFrm_Configuracoes;

implementation

uses tabelas, funcoes, principal, uRuntimeFields;

{$R *.dfm}


//resolução de tela
procedure PTela( Sender: Tobject );
  procedure PAlterar( Width, Height : Integer);
  var sDispMode: DevMode;
  begin
    EnumDisplaySettings(Nil,0,sDispMode);

    sDispMode.dmPelsWidth := Width;
    sDispMode.dmPelsHeight := Height;
    sDispMode.dmFields := DM_PELSWIDTH+DM_PELSHEIGHT;
    try
      ChangeDisplaySettings(sDispMode,0);
    except
      mensagem('Não é possível alterar configurações de vídeo.')
    end;

  end;

begin

  if ((Screen.Width <> OldWidth ) or (Screen.Height <> OldHeight )) and (Oldwidth <> 0) then
    Palterar(OldWidth, OldHeight);

  // Aqui testo se a resolução é 800 porque quero utilizar pro meu software resolução 1280 X 768.
  If ((Screen.Width <> 1280 ) or ( Screen.Height <> 768)) then
  begin
{    if (messageDlg('Para uma boa utilização recomenda-se resolução de video de 1280 X 768. ' +
                    #13#13 + ' Deseja Alterar suas configurações de vídeo agora ?',
       mtConfirmation, [mbYes, mbNo], 0)=mrYes) then}
    begin
      // Vairavel que pega o valor original de inicialização ex: 1280 X 720... etc...
      OldWidth := GetSystemMetrics(SM_CXSCREEN);
      OldHeight := GetSystemMetrics(SM_CYSCREEN);
      // Resolução para qual quero utilizar enquanto o programa estiver aberto.
      if DM_Tabelas<>nil then
      begin
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '5120x2160'  then
            PAlterar(5120, 2160); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '3840x2160'  then
           PAlterar(1366, 768); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '2560x1440'  then
            PAlterar(2560, 1440); // 1366 X 768
        if  DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString=  '2040x1080'  then
            PAlterar(2040, 1080); // 1366 X 768
        if  DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString=  '2040x858' then
            PAlterar(2040, 858); // 1366 X 768
        if  DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '1998x1080' then
            PAlterar(1998, 1080); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '1920x1080' then
            PAlterar(1920, 1080); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '1680x1050'  then
            PAlterar(1680, 1050); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString='1600x900' then
           PAlterar(1600, 900); // 1600 X 900
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString='1366x768' then
           PAlterar(1366, 768); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString='1360x768' then
           PAlterar(1360, 768); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString='1280x1024' then
           PAlterar(1280,1024); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString='1280x960' then
           PAlterar(1280, 960); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString='1280x800' then
           PAlterar(1280, 800); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString='1280x768' then
           PAlterar(1280, 768); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString='1280x720' then
           PAlterar(1280, 720); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString='1280x600' then
           PAlterar(1280, 600); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString='1024x768' then
           PAlterar(1024, 600); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString='800x600' then
           PAlterar(800, 600); // 1366 X 768
      end;
    end;
  end;

end;


procedure TFrm_Configuracoes.BtFecharClick(Sender: TObject);
begin
  //chama resolução de tela
  PTela(Sender);
  Close;
end;

procedure TFrm_Configuracoes.FormShow(Sender: TObject);
begin
  DM_Tabelas.ZQConfiguracoes.Edit;
  if DM_Tabelas.ZQConfiguracoes.FieldByName('venda_planodecontas').AsInteger>0 Then Begin
    (AchaPlanoDeConta(300,30,'S',DM_Tabelas.ZQConfiguracoes.FieldByName('venda_planodecontas').Text,'S'));
     EVendaPlano.Text := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').Text+' | '+DM_tabelas.ZQAchaPlanoDeContas.FieldByName('mascara').AsString+' | '+DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('descricao').AsString;
  end;
  if DM_Tabelas.ZQConfiguracoes.FieldByName('recebe_planodecontas').AsInteger>0 Then Begin
    (AchaPlanoDeConta(300,30,'S',DM_Tabelas.ZQConfiguracoes.FieldByName('recebe_planodecontas').Text,'S'));
    ERecebePlano.Text := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').Text+' | '+DM_tabelas.ZQAchaPlanoDeContas.FieldByName('mascara').AsString+' | '+DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('descricao').AsString;
  end;
  if not empty(DM_Tabelas.ZQConfiguracoes.FieldByName('price').AsString) Then Begin
    taxa_price.Text := DM_Tabelas.ZQConfiguracoes.FieldByName('price').AsString;
  end;
  xreajuste.Text:=DM_Tabelas.ZQConfiguracoes.FieldByName('Tipo_reajuste').AsString;
  bxmanual.Text:=DM_Tabelas.ZQConfiguracoes.FieldByName('Baixa').AsString;
  Xmulta.Value :=DM_Tabelas.ZQConfiguracoes.FieldByName('Multa').AsFloat;
  Xmora.Value :=DM_Tabelas.ZQConfiguracoes.FieldByName('Mora').AsFloat;
  Cmesmodia.Text := DM_Tabelas.ZQConfiguracoes.FieldByName('mesmodiames').AsString;
  EVendaPlano.SetFocus;

  if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '5120x2160'  then
      resolucao.ItemIndex:=0;
  if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '3840x2160'  then
      resolucao.ItemIndex:=1;
  if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '2560x1440'  then
      resolucao.ItemIndex:=2;
  if  DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString=  '2040x1080'  then
      resolucao.ItemIndex:=3;
  if  DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString=  '20400x858' then
      resolucao.ItemIndex:=4;
  if  DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '1998x1080' then
      resolucao.ItemIndex:=5;
  if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '1920x1080' then
      resolucao.ItemIndex:=6;
  if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '1680x1050'  then
      resolucao.ItemIndex:=7;
  if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '1600x900' then
      resolucao.ItemIndex:=8;
  if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '1440x900' then
      resolucao.ItemIndex:=9;
  if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '1366x768'  then
      resolucao.ItemIndex:=10;
  if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '1360x768'  then
      resolucao.ItemIndex:=11;
  if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '1280x1024'  then
      resolucao.ItemIndex:=12;
  if  DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString=  '1280x960'  then
      resolucao.ItemIndex:=13;
  if  DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString=  '1280x800' then
      resolucao.ItemIndex:=14;
  if  DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '1280x768' then
      resolucao.ItemIndex:=15;
  if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '1280x720' then
      resolucao.ItemIndex:=16;
  if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '1280x600'  then
      resolucao.ItemIndex:=17;
  if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '1024x768' then
      resolucao.ItemIndex:=18;
  if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '800x600' then
      resolucao.ItemIndex:=19;

end;

procedure TFrm_Configuracoes.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  if not Verif_senha('Configurações','Salvar Alterações','') then exit;
  DM_Tabelas.ZQConfiguracoes.Post;
end;

procedure TFrm_Configuracoes.EVendaPlanoExit(Sender: TObject);
begin
  if (not empty(EVendaPlano.Text)) and (AchaPlanoDeConta(350,30,'S',EVendaPlano.Text,'C')) Then Begin
    EVendaPlano.Text := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').Text+' | '+DM_tabelas.ZQAchaPlanoDeContas.FieldByName('mascara').AsString+' | '+DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('descricao').AsString;

    DM_Tabelas.ZQConfiguracoes.FieldByName('venda_planodecontas').AsInteger :=  DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').AsLargeInt;
  end
  else
    DM_Tabelas.ZQConfiguracoes.FieldByName('venda_planodecontas').AsInteger := 0;
end;

procedure TFrm_Configuracoes.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrm_Configuracoes.ERecebePlanoExit(Sender: TObject);
begin
  if (not empty(ERecebePlano.Text)) and (AchaPlanoDeConta(350,30,'S',ERecebePlano.Text,'C')) Then Begin
    ERecebePlano.Text := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').Text+' | '+DM_tabelas.ZQAchaPlanoDeContas.FieldByName('mascara').AsString+' | '+DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('descricao').AsString;

    DM_Tabelas.ZQConfiguracoes.FieldByName('recebe_planodecontas').AsInteger :=  DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').AsLargeInt;
  end
  else
    DM_Tabelas.ZQConfiguracoes.FieldByName('recebe_planodecontas').AsInteger := 0;

end;

procedure TFrm_Configuracoes.taxa_priceExit(Sender: TObject);
begin
  DM_Tabelas.ZQConfiguracoes.FieldByName('price').AsString:=taxa_price.Text;
end;

procedure TFrm_Configuracoes.bxmanualExit(Sender: TObject);
begin
  DM_Tabelas.ZQConfiguracoes.FieldByName('Baixa').AsString:=bxmanual.Text;
end;

procedure TFrm_Configuracoes.CmesmodiaExit(Sender: TObject);
begin
  DM_Tabelas.ZQConfiguracoes.FieldByName('mesmodiames').AsString:=Cmesmodia.Text;
end;

procedure TFrm_Configuracoes.XmultaExit(Sender: TObject);
begin
  DM_Tabelas.ZQConfiguracoes.FieldByName('Multa').AsFloat:=Xmulta.Value;
end;

procedure TFrm_Configuracoes.XmoraExit(Sender: TObject);
begin
  DM_Tabelas.ZQConfiguracoes.FieldByName('Mora').AsFloat:=Xmora.Value;
end;

procedure TFrm_Configuracoes.xreajusteExit(Sender: TObject);
begin
  DM_Tabelas.ZQConfiguracoes.FieldByName('Tipo_reajuste').AsString:=xreajuste.Text;
end;

procedure TFrm_Configuracoes.ResolucaoExit(Sender: TObject);
begin
  if DM_Tabelas.ZQConfiguracoes.State in [DsInsert, DsEdit] then
  begin
    if resolucao.ItemIndex=0 then
       DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString:= '5120x2160';
    if resolucao.ItemIndex=1 then
       DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString:= '3840x2160';
    if resolucao.ItemIndex=2 then
       DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString:= '2560x1440';
    if resolucao.ItemIndex=3 then
       DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString:=  '2040x1080';
    if resolucao.ItemIndex=4 then
       DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString:=  '20400x858';
    if resolucao.ItemIndex=5 then
       DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString:= '1998x1080';
    if resolucao.ItemIndex=6 then
       DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString:= '1920x1080';
    if resolucao.ItemIndex=7 then
       DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString:= '1680x1050';
    if resolucao.ItemIndex=8 then
       DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString:= '1600x900';
    if resolucao.ItemIndex=9 then
       DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString:= '1440x900';
    if resolucao.ItemIndex=10 then
       DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString:= '1366x768';
       if resolucao.ItemIndex=11 then
       DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString:= '1360x768';
    if resolucao.ItemIndex=12  then
       DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString:= '1280x1024';
    if  resolucao.ItemIndex=13  then
        DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString:=  '1280x960';
    if  resolucao.ItemIndex=14 then
        DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString:=  '1280x800';
    if  resolucao.ItemIndex=15 then
        DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString:= '1280x768';
    if  resolucao.ItemIndex=16 then
        DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString:= '1280x720';
    if  resolucao.ItemIndex=17 then
        DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString:= '1280x600';
    if  resolucao.ItemIndex=18 then
        DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString:= '1024x768';
    if resolucao.ItemIndex=19 then
        DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString:= '800x600';

  end;
end;


procedure TFrm_Configuracoes.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
