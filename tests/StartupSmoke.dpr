program StartupSmoke;

{$APPTYPE CONSOLE}

uses
  uSiaiPerformance in 'uSiaiPerformance.pas',
  uSiaiCotacoes in 'uSiaiCotacoes.pas',
  System.SysUtils, System.Classes, System.SyncObjs, Winapi.Windows,
  Vcl.Forms, Vcl.ExtCtrls, Vcl.Controls,
  principal in 'principal.pas',
  Apresentacao in 'Apresentacao.pas',
  tabelas in 'tabelas.pas',
  funcoes in 'funcoes.pas',
  AchaImovel in 'AchaImovel.pas',
  Achavendas in 'Achavendas.pas',
  Finanmsg in 'Finanmsg.pas',
  PerguntaSIMNAO in 'PerguntaSIMNAO.pas',
  Pergunta in 'Pergunta.pas',
  AchaParticipante in 'AchaParticipante.pas',
  AchaCidade in 'AchaCidade.pas',
  AchaCorretor in 'AchaCorretor.pas',
  AchaPlanoDeContas in 'AchaPlanoDeContas.pas',
  Acha_Contabancaria in 'Acha_Contabancaria.pas';

type
  TTestCotacoes = class(TSiaiCotacoesThread)
  public
    Payload: string;
    FailRequest: Boolean;
    Entered, ResumeRequest: TEvent;
    constructor Create;
    destructor Destroy; override;
    function FetchResponse: string; override;
  end;

  TObserver = class
  public
    Called: Boolean;
    OnMainThread: Boolean;
    procedure Completed(Sender: TObject);
  end;

  TDialogDriver = class
  private
    Timer: TTimer;
    Started: Cardinal;
    Handled: Boolean;
    Failure: string;
    procedure Tick(Sender: TObject);
  public
    Kind: Integer;
    Prompt, Initial, Answer: string;
    CancelInput: Boolean;
    constructor Create;
    destructor Destroy; override;
    procedure Arm;
    procedure Verify;
  end;

var
  Passed: Integer;

procedure Check(ACondition: Boolean; const AMessage: string);
begin
  if not ACondition then raise Exception.Create(AMessage);
end;

constructor TDialogDriver.Create;
begin
  inherited;
  Timer := TTimer.Create(nil);
  Timer.Enabled := False;
  Timer.Interval := 10;
  Timer.OnTimer := Tick;
end;

destructor TDialogDriver.Destroy;
begin
  Timer.Free;
  inherited;
end;

procedure TDialogDriver.Arm;
begin
  Handled := False;
  Failure := '';
  Started := GetTickCount;
  Timer.Enabled := True;
end;

procedure TDialogDriver.Tick(Sender: TObject);
var
  LForm: TForm;
  LKey: Char;
begin
  LForm := Screen.ActiveForm;
  if (LForm = nil) or not LForm.Visible then Exit;
  Timer.Enabled := False;
  try
    Check(GetTickCount - Started < 5000, 'Dialogo demorou demais');
    case Kind of
      0:
        begin
          Check(LForm = FormMensagem, 'Mensagem nao abriu o form esperado');
          Check(FormMensagem.Frase.Caption = Prompt, 'Texto da mensagem incorreto');
          FormMensagem.BtnOkClick(nil);
        end;
      1:
        begin
          Check(LForm = FrmPerguntaSIMNAO, 'SIMNAO nao abriu o form esperado');
          Check(FrmPerguntaSIMNAO.Label1.Caption = Prompt, 'Texto SIMNAO incorreto');
          Check(FrmPerguntaSIMNAO.Label2.Caption = Initial, 'Opcao inicial incorreta');
          if Answer = 'SIM' then FrmPerguntaSIMNAO.DXBSIMClick(nil)
          else FrmPerguntaSIMNAO.DXBNAOClick(nil);
        end;
      2:
        begin
          Check(LForm = FrmPergunta, 'Pergunta nao abriu o form esperado');
          Check(FrmPergunta.Label5.Caption = Prompt, 'Texto da pergunta incorreto');
          Check(FrmPergunta.XEdit1.Text = Initial, 'Resposta anterior nao foi substituida');
          if CancelInput then LKey := #27
          else
          begin
            FrmPergunta.XEdit1.Text := Answer;
            LKey := #13;
          end;
          FrmPergunta.XEdit1KeyPress(nil, LKey);
        end;
    end;
    Handled := True;
  except
    on E: Exception do
    begin
      Failure := E.Message;
      LForm.ModalResult := mrCancel;
    end;
  end;
end;

procedure TDialogDriver.Verify;
begin
  Timer.Enabled := False;
  Check(Failure = '', Failure);
  Check(Handled, 'Dialogo nao foi exercitado');
  Inc(Passed);
end;

procedure TestLazyDialogs;
var
  Driver: TDialogDriver;
  FirstInstance: TForm;
begin
  Check((FormMensagem = nil) and (FrmPerguntaSIMNAO = nil) and
    (FrmPergunta = nil), 'Dialogos criados antes do primeiro uso');
  Driver := TDialogDriver.Create;
  try
    Driver.Kind := 0;
    Driver.Prompt := 'Mensagem do primeiro uso';
    Driver.Arm;
    Mensagem(Driver.Prompt);
    Driver.Verify;
    FirstInstance := FormMensagem;
    Check(FormMensagem.Owner = Application, 'Owner da mensagem incorreto');
    Writeln('PASS: mensagem criada sob demanda e fechada por OK');
    Driver.Prompt := 'Mensagem da reabertura';
    Driver.Arm;
    Mensagem(Driver.Prompt);
    Driver.Verify;
    Check(FormMensagem = FirstInstance, 'Mensagem nao reutilizou a instancia');
    Writeln('PASS: mensagem reaberta com novo texto e mesma instancia');

    Driver.Kind := 1;
    Driver.Prompt := 'Confirmacao de teste';
    Driver.Initial := 'SIM';
    Driver.Answer := 'NAO';
    Driver.Arm;
    Check(not SIMNAO(Driver.Prompt, Driver.Initial), 'NAO nao retornou False');
    Driver.Verify;
    FirstInstance := FrmPerguntaSIMNAO;
    Writeln('PASS: SIMNAO criado sob demanda e resposta NAO preservada');
    Driver.Initial := 'NAO';
    Driver.Answer := 'SIM';
    Driver.Arm;
    Check(SIMNAO(Driver.Prompt, Driver.Initial), 'SIM nao retornou True');
    Driver.Verify;
    Check(FrmPerguntaSIMNAO = FirstInstance, 'SIMNAO nao reutilizou a instancia');
    Writeln('PASS: SIMNAO reaberto e resposta SIM preservada');

    Driver.Kind := 2;
    Driver.Prompt := 'Pergunta de teste';
    Driver.Initial := 'Inicial';
    Driver.Answer := 'Resposta preenchida';
    Driver.Arm;
    Check(funcoes.pergunta(Driver.Prompt, Driver.Initial) = Driver.Answer,
      'Resposta de texto nao sobreviveu ao fechamento');
    Driver.Verify;
    FirstInstance := FrmPergunta;
    Writeln('PASS: pergunta criada sob demanda e resposta lida apos Enter');
    Driver.Initial := 'Novo valor inicial';
    Driver.CancelInput := True;
    Driver.Arm;
    Check(funcoes.pergunta(Driver.Prompt, Driver.Initial) = Driver.Initial,
      'Escape alterou o comportamento anterior da pergunta');
    Driver.Verify;
    Check(FrmPergunta = FirstInstance, 'Pergunta nao reutilizou a instancia');
    Writeln('PASS: pergunta reaberta e fechada com Escape');
    Check(DM_Tabelas = nil, 'Dialogos nao devem criar o DataModule');
  finally
    Driver.Free;
    FreeAndNil(FormMensagem);
    FreeAndNil(FrmPerguntaSIMNAO);
    FreeAndNil(FrmPergunta);
  end;
end;

procedure TestSearchFormConstruction;
const
  FormClasses: array[0..4] of TFormClass = (TFrm_AchaParticipante,
    TFrm_AchaCidade, TFrm_AchaCorretor, TFrm_AchaPlanoDeContas,
    TFrm_Acha_Contabancaria);
var
  FormClass: TFormClass;
  Form: TForm;
  PreviousCount: Integer;
begin
  Check((Frm_AchaParticipante = nil) and (Frm_AchaCidade = nil) and
    (Frm_AchaCorretor = nil) and (Frm_AchaPlanoDeContas = nil) and
    (Frm_Acha_Contabancaria = nil), 'Pesquisa criada durante a inicializacao');
  for FormClass in FormClasses do
  begin
    PreviousCount := Application.ComponentCount;
    Form := FormClass.Create(Application);
    try
      Check(Form.Owner = Application, 'Owner da pesquisa incorreto');
      Check(DM_Tabelas = nil, 'Construcao da pesquisa criou DataModule');
      Check(not Form.Visible, 'Construcao da pesquisa exibiu tela antecipadamente');
      Writeln('PASS: construcao isolada do DFM ', FormClass.ClassName);
      Inc(Passed);
    finally
      Form.Free;
    end;
    Check(Application.ComponentCount = PreviousCount, 'Pesquisa nao foi liberada no teste');
  end;
end;

constructor TTestCotacoes.Create;
begin
  inherited Create;
  FreeOnTerminate := False;
  Entered := TEvent.Create(nil, True, False, '');
  ResumeRequest := TEvent.Create(nil, True, False, '');
end;

destructor TTestCotacoes.Destroy;
begin
  // Todos os testes liberam a requisicao e aguardam a thread antes de destruir.
  ResumeRequest.Free;
  Entered.Free;
  inherited;
end;

function TTestCotacoes.FetchResponse: string;
begin
  Entered.SetEvent;
  if ResumeRequest.WaitFor(5000) <> wrSignaled then
    raise Exception.Create('Timeout do teste');
  if FailRequest then raise Exception.Create('Falha de rede simulada');
  Result := Payload;
end;

procedure TObserver.Completed(Sender: TObject);
begin
  Called := True;
  OnMainThread := GetCurrentThreadId = MainThreadID;
end;

procedure TestResponse(const AName, APayload: string; AExpected, AFail: Boolean);
var
  Worker: TTestCotacoes;
  Observer: TObserver;
begin
  Worker := TTestCotacoes.Create;
  Observer := TObserver.Create;
  try
    Worker.Payload := APayload;
    Worker.FailRequest := AFail;
    Worker.OnTerminate := Observer.Completed;
    Worker.Start;
    Check(Worker.Entered.WaitFor(5000) = wrSignaled, 'Thread nao iniciou');
    Check(not Observer.Called, 'Callback ocorreu com requisicao pendente');
    // O fluxo principal chegou aqui enquanto a requisicao permanece bloqueada.
    Worker.ResumeRequest.SetEvent;
    Worker.WaitFor;
    Check(Observer.Called and Observer.OnMainThread, 'Callback fora da thread principal');
    Check(Worker.Success = AExpected, 'Resultado incorreto: ' + AName);
    if AExpected then
      Check((Worker.USD = '5.10') and (Worker.EUR = '6.20') and
        (Worker.BTC = '300000'), 'Valores de cotacao incorretos');
    Inc(Passed);
    Writeln('PASS: ', AName);
  finally
    Worker.ResumeRequest.SetEvent;
    Worker.WaitFor;
    Worker.Free;
    Observer.Free;
  end;
end;

procedure TestDetach;
var
  Worker: TTestCotacoes;
  Observer: TObserver;
begin
  Worker := TTestCotacoes.Create;
  Observer := TObserver.Create;
  try
    Worker.OnTerminate := Observer.Completed;
    Worker.Start;
    Check(Worker.Entered.WaitFor(5000) = wrSignaled, 'Thread nao iniciou');
    Worker.Terminate;
    Worker.OnTerminate := nil;
    FreeAndNil(Observer);
    Worker.ResumeRequest.SetEvent;
    Worker.WaitFor;
    Check(not Worker.Success, 'Cancelamento nao respeitado');
    Inc(Passed);
    Writeln('PASS: fechar destinatario durante requisicao pendente');
  finally
    Worker.ResumeRequest.SetEvent;
    Worker.WaitFor;
    Worker.Free;
    Observer.Free;
  end;
end;

procedure TestFormsWithoutDatabase;
begin
  Application.Initialize;
  Application.CreateForm(TFrm_principal, Frm_principal);
  try
    Check(DM_Tabelas = nil, 'Teste nao deve conectar ao banco');
    Check((Frm_Funcoes = nil) and (Frm_AchaImovel = nil) and
      (Frm_Achavendas = nil), 'Formulario removido foi criado');
    Frm_principal.Timer2Timer(nil);
    Check(Pos('Usada:', Frm_principal.StatusBar1.Panels[1].Text) > 0,
      'Indicador de memoria deixou de funcionar');
    FrmApresentacao := TFrmApresentacao.Create(nil);
    try
      FrmApresentacao.Timer1.Enabled := False;
      Check(not FrmApresentacao.Timer1.Enabled, 'Timer da apresentacao ativo');
      Check(Application.MainForm = Frm_principal, 'Apresentacao substituiu MainForm');
    finally
      FreeAndNil(FrmApresentacao);
    end;
    Check(DM_Tabelas = nil, 'Construcao tentou criar DataModule');
    Inc(Passed);
    Writeln('PASS: principal, apresentacao e indicador sem banco ou forms removidos');
  finally
    FreeAndNil(Frm_principal);
  end;
end;

begin
  try
    TestResponse('JSON valido',
      '{"USDBRL":{"bid":"5.10"},"EURBRL":{"bid":"6.20"},"BTCBRL":{"bid":"300000"}}', True, False);
    TestResponse('JSON numerico',
      '{"USDBRL":{"bid":"5.10"},"EURBRL":{"bid":"6.20"},"BTCBRL":{"bid":300000}}', True, False);
    TestResponse('JSON invalido', 'nao e JSON', False, False);
    TestResponse('JSON nulo', 'null', False, False);
    TestResponse('Moeda ausente', '{"USDBRL":{"bid":"5.10"}}', False, False);
    TestResponse('Resposta vazia', '', False, False);
    TestResponse('Falha HTTP/SSL simulada', '', False, True);
    TestDetach;
    TestFormsWithoutDatabase;
    TestLazyDialogs;
    TestSearchFormConstruction;
    PerformanceCheckpoint('Smoke tests concluidos');
    FlushPerformanceLog;
    Writeln('TOTAL PASS: ', Passed);
  except
    on E: Exception do
    begin
      Writeln('FAIL: ', E.ClassName, ': ', E.Message);
      ExitCode := 1;
    end;
  end;
end.
