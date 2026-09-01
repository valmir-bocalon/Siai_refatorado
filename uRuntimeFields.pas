unit uRuntimeFields;

interface

uses
  System.Classes,
  System.SysUtils,
  System.Generics.Collections,
  Data.DB,
  System.TypInfo,
  System.Variants,
  Vcl.Forms;

procedure RegisterRuntimeDataSet(AOwnerClass: TClass; const ADataSetName: string; AOpenAfterSetup: Boolean);
procedure RegisterRuntimeField(
  AOwnerClass: TClass; const ADataSetName, AComponentName, AFieldName: string;
  AFieldClass: TFieldClass;
  AFieldKind: TFieldKind;
  ASize, APrecision: Integer;
  ARequired: Boolean;
  const ADisplayLabel, ADisplayFormat, AEditFormat, AEditMask: string;
  ADisplayWidth: Integer;
  const AKeyFields, ALookupDataSetName, ALookupKeyFields, ALookupResultField: string;
  ALookup: Boolean);
procedure EnsureRuntimeFields(AOwner: TComponent);

implementation

type
  TFieldSetting = class
  public
    FieldName: string;
    KeyFields: string;
    LookupDataSetName: string;
    LookupKeyFields: string;
    LookupResultField: string;
    DisplayFormat: string;
    EditFormat: string;
    DisplayLabel: string;
    EditMask: string;
    DisplayWidth: Integer;
    HasDisplayWidth: Boolean;
    Lookup: Boolean;
  end;

  TFieldSpec = class
  public
    ComponentName: string;
    FieldName: string;
    FieldClass: TFieldClass;
    FieldKind: TFieldKind;
    Size: Integer;
    Precision: Integer;
    Required: Boolean;
    DisplayLabel: string;
    DisplayFormat: string;
    EditFormat: string;
    EditMask: string;
    DisplayWidth: Integer;
    KeyFields: string;
    LookupDataSetName: string;
    LookupKeyFields: string;
    LookupResultField: string;
    Lookup: Boolean;
  end;

  TDataSetSpec = class
  public
    OwnerClass: TClass;
    DataSetName: string;
    OpenAfterSetup: Boolean;
    Fields: TObjectList<TFieldSpec>;
    constructor Create;
    destructor Destroy; override;
  end;

  TDataSetFieldSettingsHook = class(TComponent)
  private
    FDataSet: TDataSet;
    FPrevAfterOpen: TDataSetNotifyEvent;
    FPrevOnCalcFields: TDataSetNotifyEvent;
    FSettings: TObjectList<TFieldSetting>;
    procedure DataSetAfterOpen(DataSet: TDataSet);
    procedure DataSetCalcFields(DataSet: TDataSet);
    function FindOrAddSetting(const AFieldName: string): TFieldSetting;
    procedure ApplySetting(ASetting: TFieldSetting);
    procedure ApplyAll;
    procedure ApplyCalculatedLookups;
  public
    constructor Create(ADataSet: TDataSet); reintroduce;
    destructor Destroy; override;
  end;

  TRuntimeFieldsMarker = class(TComponent);

var
  GDataSets: TObjectList<TDataSetSpec>;

function FindDataSetSpec(AOwnerClass: TClass; const ADataSetName: string): TDataSetSpec;
var
  LSpec: TDataSetSpec;
begin
  Result := nil;
  for LSpec in GDataSets do
    if (LSpec.OwnerClass = AOwnerClass) and
       SameText(LSpec.DataSetName, ADataSetName) then
      Exit(LSpec);
end;

function FindFieldSpec(ASpec: TDataSetSpec; const AFieldName: string): TFieldSpec;
var
  LField: TFieldSpec;
begin
  Result := nil;
  if ASpec = nil then
    Exit;
  for LField in ASpec.Fields do
    if SameText(LField.FieldName, AFieldName) then
      Exit(LField);
end;

function LastComponentName(const AQualifiedName: string): string;
var
  LPos: Integer;
begin
  LPos := LastDelimiter('.', AQualifiedName);
  if LPos > 0 then
    Result := Copy(AQualifiedName, LPos + 1, MaxInt)
  else
    Result := AQualifiedName;
end;

function FindComponentRecursive(AOwner: TComponent; const AName: string): TComponent;
var
  I: Integer;
  LComponent: TComponent;
begin
  Result := nil;
  if AOwner = nil then
    Exit;
  if SameText(AOwner.Name, AName) then
    Exit(AOwner);
  for I := 0 to AOwner.ComponentCount - 1 do
  begin
    LComponent := AOwner.Components[I];
    if SameText(LComponent.Name, AName) then
      Exit(LComponent);
    Result := FindComponentRecursive(LComponent, AName);
    if Result <> nil then
      Exit;
  end;
end;

function ResolveLookupDataSet(AOwner: TComponent; const AQualifiedName: string): TDataSet;
var
  LComponent: TComponent;
  I: Integer;
begin
  Result := nil;
  if AQualifiedName = '' then
    Exit;
  LComponent := FindComponentRecursive(AOwner, LastComponentName(AQualifiedName));
  if (LComponent = nil) and (Application <> nil) then
  begin
    for I := 0 to Application.ComponentCount - 1 do
    begin
      LComponent := FindComponentRecursive(Application.Components[I], LastComponentName(AQualifiedName));
      if LComponent <> nil then
        Break;
    end;
  end;
  if LComponent is TDataSet then
    Result := TDataSet(LComponent);
end;

function GetFieldClassForDataType(ADataType: TFieldType; ADefaultClass: TFieldClass): TFieldClass;
begin
  Result := ADefaultClass;
  case ADataType of
    ftSmallint, ftShortint: Result := TSmallintField;
    ftInteger: Result := TIntegerField;
    ftWord: Result := TWordField;
    ftLongWord: Result := TLongWordField;
    ftLargeint: Result := TLargeintField;
    ftAutoInc: Result := TAutoIncField;
    ftBoolean: Result := TBooleanField;
    ftFloat: Result := TFloatField;
    ftCurrency: Result := TCurrencyField;
    ftBCD: Result := TBCDField;
    ftFMTBcd: Result := TFMTBCDField;
    ftDate: Result := TDateField;
    ftTime: Result := TTimeField;
    ftDateTime: Result := TDateTimeField;
    ftTimeStamp: Result := TSQLTimeStampField;
    ftString, ftFixedChar: Result := TStringField;
    ftWideString, ftFixedWideChar: Result := TWideStringField;
    ftMemo: Result := TMemoField;
    ftWideMemo: Result := TWideMemoField;
    ftBlob, ftOraBlob, ftGraphic: Result := TBlobField;
  end;
end;

function ResolveRuntimeFieldClass(ADataSet: TDataSet; const AFieldName: string; ARequestedClass: TFieldClass): TFieldClass;
var
  I: Integer;
  LFieldDef: TFieldDef;
begin
  Result := ARequestedClass;
  if (ADataSet = nil) or (ARequestedClass = nil) then
    Exit;
  LFieldDef := nil;
  for I := 0 to ADataSet.FieldDefs.Count - 1 do
    if SameText(ADataSet.FieldDefs[I].Name, AFieldName) then
    begin
      LFieldDef := ADataSet.FieldDefs[I];
      Break;
    end;
  if LFieldDef <> nil then
    Result := GetFieldClassForDataType(LFieldDef.DataType, ARequestedClass);
end;

procedure TrySetOrdProp(AInstance: TObject; const APropName: string; AValue: NativeInt);
begin
  if IsPublishedProp(AInstance, APropName) then
    SetOrdProp(AInstance, APropName, AValue);
end;

procedure TrySetStrProp(AInstance: TObject; const APropName, AValue: string);
begin
  if (AValue <> '') and IsPublishedProp(AInstance, APropName) then
    SetStrProp(AInstance, APropName, AValue);
end;

procedure ClearCalculatedField(AField: TField);
begin
  if AField = nil then
    Exit;
  { Campos calculados de texto sao lidos em varios pontos antigos por meio
    de .Value. Para eles, uma ausencia de lookup deve ser texto vazio, nao
    Variant Null; campos numericos e de data continuam semanticamente nulos. }
  if AField.DataType in [ftString, ftFixedChar, ftWideString, ftFixedWideChar,
    ftMemo, ftWideMemo] then
    AField.AsString := ''
  else
    AField.Clear;
end;

constructor TDataSetSpec.Create;
begin
  inherited Create;
  Fields := TObjectList<TFieldSpec>.Create(True);
end;

destructor TDataSetSpec.Destroy;
begin
  Fields.Free;
  inherited Destroy;
end;

constructor TDataSetFieldSettingsHook.Create(ADataSet: TDataSet);
begin
  inherited Create(ADataSet);
  FDataSet := ADataSet;
  FSettings := TObjectList<TFieldSetting>.Create(True);
  FPrevAfterOpen := ADataSet.AfterOpen;
  FPrevOnCalcFields := ADataSet.OnCalcFields;
  ADataSet.AfterOpen := DataSetAfterOpen;
  ADataSet.OnCalcFields := DataSetCalcFields;
end;

destructor TDataSetFieldSettingsHook.Destroy;
begin
  FSettings.Free;
  inherited Destroy;
end;

function TDataSetFieldSettingsHook.FindOrAddSetting(const AFieldName: string): TFieldSetting;
var
  LSetting: TFieldSetting;
begin
  for LSetting in FSettings do
    if SameText(LSetting.FieldName, AFieldName) then
      Exit(LSetting);
  Result := TFieldSetting.Create;
  Result.FieldName := AFieldName;
  FSettings.Add(Result);
end;

procedure TDataSetFieldSettingsHook.DataSetAfterOpen(DataSet: TDataSet);
begin
  if Assigned(FPrevAfterOpen) then
    FPrevAfterOpen(DataSet);
  ApplyAll;
end;

procedure TDataSetFieldSettingsHook.DataSetCalcFields(DataSet: TDataSet);
begin
  if Assigned(FPrevOnCalcFields) then
    FPrevOnCalcFields(DataSet);
  ApplyCalculatedLookups;
end;

procedure TDataSetFieldSettingsHook.ApplySetting(ASetting: TFieldSetting);
var
  LField: TField;
begin
  if (FDataSet = nil) or (ASetting = nil) then
    Exit;
  LField := FDataSet.FindField(ASetting.FieldName);
  if LField = nil then
    Exit;
  TrySetStrProp(LField, 'DisplayFormat', ASetting.DisplayFormat);
  TrySetStrProp(LField, 'EditFormat', ASetting.EditFormat);
  TrySetStrProp(LField, 'DisplayLabel', ASetting.DisplayLabel);
  TrySetStrProp(LField, 'EditMask', ASetting.EditMask);
  if ASetting.HasDisplayWidth and IsPublishedProp(LField, 'DisplayWidth') then
    SetOrdProp(LField, 'DisplayWidth', ASetting.DisplayWidth);
end;

procedure TDataSetFieldSettingsHook.ApplyAll;
var
  LSetting: TFieldSetting;
begin
  for LSetting in FSettings do
    ApplySetting(LSetting);
end;

procedure TDataSetFieldSettingsHook.ApplyCalculatedLookups;
var
  LSetting: TFieldSetting;
  LField: TField;
  LKeyField: TField;
  LLookupDataSet: TDataSet;
  LKeyValue: Variant;
  LKeyFields: TArray<string>;
  LLookupKeyFields: TArray<string>;
  LHasNullKey: Boolean;
  LLookupFieldsAvailable: Boolean;
  LLookupValue: Variant;
  I: Integer;
begin
  if (FDataSet = nil) or (not FDataSet.Active) or FDataSet.IsEmpty then
    Exit;
  if FDataSet.Eof and not FDataSet.Bof then
    Exit;
  for LSetting in FSettings do
  begin
    if not LSetting.Lookup then
      Continue;
    LLookupDataSet := ResolveLookupDataSet(FDataSet.Owner, LSetting.LookupDataSetName);
    if (LSetting.KeyFields = '') or (LSetting.LookupKeyFields = '') or
       (LSetting.LookupResultField = '') or (LLookupDataSet = nil) or
       (not LLookupDataSet.Active) then
      Continue;
    LField := FDataSet.FindField(LSetting.FieldName);
    if LField = nil then
      Continue;
    LKeyFields := LSetting.KeyFields.Split([',']);
    LHasNullKey := False;
    if Length(LKeyFields) = 1 then
    begin
      LKeyField := FDataSet.FindField(Trim(LKeyFields[0]));
      if (LKeyField = nil) or LKeyField.IsNull then
        LHasNullKey := True
      else
        LKeyValue := LKeyField.Value;
    end
    else
    begin
      LKeyValue := VarArrayCreate([0, Length(LKeyFields) - 1], varVariant);
      for I := 0 to High(LKeyFields) do
      begin
        LKeyField := FDataSet.FindField(Trim(LKeyFields[I]));
        if (LKeyField = nil) or LKeyField.IsNull then
        begin
          LHasNullKey := True;
          Break;
        end;
        LKeyValue[I] := LKeyField.Value;
      end;
    end;
    if LHasNullKey then
    begin
      ClearCalculatedField(LField);
      Continue;
    end;
    { Alguns forms possuem datasets locais com o mesmo nome, mas com uma
      consulta diferente. Valide o lookup antes de chamar Lookup para que um
      campo opcional ausente nao interrompa a abertura do sistema. }
    LLookupFieldsAvailable := LLookupDataSet.FindField(LSetting.LookupResultField) <> nil;
    if LLookupFieldsAvailable then
    begin
      LLookupKeyFields := LSetting.LookupKeyFields.Split([',']);
      for I := 0 to High(LLookupKeyFields) do
        if (Trim(LLookupKeyFields[I]) <> '') and
           (LLookupDataSet.FindField(Trim(LLookupKeyFields[I])) = nil) then
        begin
          LLookupFieldsAvailable := False;
          Break;
        end;
    end;
    if not LLookupFieldsAvailable then
    begin
      ClearCalculatedField(LField);
      Continue;
    end;
    LLookupValue := LLookupDataSet.Lookup(LSetting.LookupKeyFields, LKeyValue, LSetting.LookupResultField);
    if VarIsNull(LLookupValue) or VarIsEmpty(LLookupValue) then
      ClearCalculatedField(LField)
    else
      LField.Value := LLookupValue;
  end;
end;

function GetDataSetHook(ADataSet: TDataSet): TDataSetFieldSettingsHook;
var
  I: Integer;
begin
  Result := nil;
  if ADataSet = nil then
    Exit;
  for I := 0 to ADataSet.ComponentCount - 1 do
    if ADataSet.Components[I] is TDataSetFieldSettingsHook then
      Exit(TDataSetFieldSettingsHook(ADataSet.Components[I]));
end;

function EnsureDataSetHook(ADataSet: TDataSet): TDataSetFieldSettingsHook;
begin
  Result := GetDataSetHook(ADataSet);
  if Result = nil then
    Result := TDataSetFieldSettingsHook.Create(ADataSet);
end;

procedure AddRuntimeField(AOwner: TComponent; ADataSet: TDataSet; AFieldSpec: TFieldSpec);
var
  LField: TField;
  LFieldClass: TFieldClass;
  LSetting: TFieldSetting;
  LLookup: Boolean;
  LLookupDataSet: TDataSet;
begin
  if (AOwner = nil) or (ADataSet = nil) or (AFieldSpec = nil) then
    Exit;
  LLookup := AFieldSpec.Lookup or (AFieldSpec.FieldKind = fkLookup);
  LLookupDataSet := ResolveLookupDataSet(AOwner, AFieldSpec.LookupDataSetName);
  LSetting := EnsureDataSetHook(ADataSet).FindOrAddSetting(AFieldSpec.FieldName);
  LSetting.KeyFields := AFieldSpec.KeyFields;
  LSetting.LookupDataSetName := AFieldSpec.LookupDataSetName;
  LSetting.LookupKeyFields := AFieldSpec.LookupKeyFields;
  LSetting.LookupResultField := AFieldSpec.LookupResultField;
  LSetting.Lookup := LLookup;
  LSetting.DisplayLabel := AFieldSpec.DisplayLabel;
  LSetting.DisplayFormat := AFieldSpec.DisplayFormat;
  LSetting.EditFormat := AFieldSpec.EditFormat;
  LSetting.EditMask := AFieldSpec.EditMask;
  LSetting.DisplayWidth := AFieldSpec.DisplayWidth;
  LSetting.HasDisplayWidth := AFieldSpec.DisplayWidth > 0;
  LField := ADataSet.FindField(AFieldSpec.FieldName);
  if LField = nil then
  begin
    LFieldClass := ResolveRuntimeFieldClass(ADataSet, AFieldSpec.FieldName, AFieldSpec.FieldClass);
    LField := LFieldClass.Create(AOwner);
    try
      if (AFieldSpec.ComponentName <> '') and (AOwner.FindComponent(AFieldSpec.ComponentName) = nil) then
        LField.Name := AFieldSpec.ComponentName;
      if LLookup then
        LField.FieldKind := fkCalculated
      else
        LField.FieldKind := AFieldSpec.FieldKind;
      LField.FieldName := AFieldSpec.FieldName;
      LField.Required := AFieldSpec.Required;
      if AFieldSpec.Size > 0 then
        TrySetOrdProp(LField, 'Size', AFieldSpec.Size);
      if AFieldSpec.Precision > 0 then
        TrySetOrdProp(LField, 'Precision', AFieldSpec.Precision);
      if AFieldSpec.DisplayWidth > 0 then
        LField.DisplayWidth := AFieldSpec.DisplayWidth;
      TrySetStrProp(LField, 'DisplayLabel', AFieldSpec.DisplayLabel);
      TrySetStrProp(LField, 'DisplayFormat', AFieldSpec.DisplayFormat);
      TrySetStrProp(LField, 'EditFormat', AFieldSpec.EditFormat);
      TrySetStrProp(LField, 'EditMask', AFieldSpec.EditMask);
      if LLookup then
      begin
        TrySetStrProp(LField, 'KeyFields', AFieldSpec.KeyFields);
        TrySetStrProp(LField, 'LookupKeyFields', AFieldSpec.LookupKeyFields);
        TrySetStrProp(LField, 'LookupResultField', AFieldSpec.LookupResultField);
        if (LLookupDataSet <> nil) and IsPublishedProp(LField, 'LookupDataSet') then
          SetObjectProp(LField, 'LookupDataSet', LLookupDataSet);
        if IsPublishedProp(LField, 'Lookup') then
          SetOrdProp(LField, 'Lookup', 1);
      end;
      LField.DataSet := ADataSet;
    except
      LField.Free;
      raise;
    end;
  end;
end;

procedure RebuildDataSet(AOwner: TComponent; ADataSet: TDataSet; ASpec: TDataSetSpec);
var
  LWasActive: Boolean;
  LFieldSpec: TFieldSpec;
begin
  LWasActive := ADataSet.Active;
  if LWasActive then
    ADataSet.Close;
  try
    try
      ADataSet.FieldDefs.Update;
    except
      { The requested field classes remain the fallback when metadata is unavailable. }
    end;
    for LFieldSpec in ASpec.Fields do
      AddRuntimeField(AOwner, ADataSet, LFieldSpec);
    { A abertura ocorre somente depois que todos os datasets receberam seus
      campos. Isso evita que um lookup seja avaliado antes do dataset alvo
      estar completamente configurado. }
  except
    if ADataSet.Active then
      ADataSet.Close;
    raise;
  end;
end;

function HasRuntimeFieldsMarker(AOwner: TComponent): Boolean;
var
  I: Integer;
begin
  Result := False;
  for I := 0 to AOwner.ComponentCount - 1 do
    if AOwner.Components[I] is TRuntimeFieldsMarker then
      Exit(True);
end;

procedure CollectDataSets(AOwner: TComponent; AList: TObjectList<TDataSet>);
var
  I: Integer;
  LComponent: TComponent;
begin
  if AOwner = nil then
    Exit;
  for I := 0 to AOwner.ComponentCount - 1 do
  begin
    LComponent := AOwner.Components[I];
    if LComponent is TDataSet then
      AList.Add(TDataSet(LComponent));
    CollectDataSets(LComponent, AList);
  end;
end;

procedure RegisterRuntimeDataSet(AOwnerClass: TClass; const ADataSetName: string; AOpenAfterSetup: Boolean);
var
  LSpec: TDataSetSpec;
begin
  LSpec := FindDataSetSpec(AOwnerClass, ADataSetName);
  if LSpec = nil then
  begin
    LSpec := TDataSetSpec.Create;
    LSpec.OwnerClass := AOwnerClass;
    LSpec.DataSetName := ADataSetName;
    GDataSets.Add(LSpec);
  end;
  LSpec.OpenAfterSetup := AOpenAfterSetup;
end;

procedure RegisterRuntimeField(
  AOwnerClass: TClass; const ADataSetName, AComponentName, AFieldName: string;
  AFieldClass: TFieldClass;
  AFieldKind: TFieldKind;
  ASize, APrecision: Integer;
  ARequired: Boolean;
  const ADisplayLabel, ADisplayFormat, AEditFormat, AEditMask: string;
  ADisplayWidth: Integer;
  const AKeyFields, ALookupDataSetName, ALookupKeyFields, ALookupResultField: string;
  ALookup: Boolean);
var
  LDataSet: TDataSetSpec;
  LField: TFieldSpec;
begin
  LDataSet := FindDataSetSpec(AOwnerClass, ADataSetName);
  if LDataSet = nil then
  begin
    RegisterRuntimeDataSet(AOwnerClass, ADataSetName, False);
    LDataSet := FindDataSetSpec(AOwnerClass, ADataSetName);
  end;
  LField := FindFieldSpec(LDataSet, AFieldName);
  if LField = nil then
  begin
    LField := TFieldSpec.Create;
    LDataSet.Fields.Add(LField);
  end;
  LField.ComponentName := AComponentName;
  LField.FieldName := AFieldName;
  LField.FieldClass := AFieldClass;
  LField.FieldKind := AFieldKind;
  LField.Size := ASize;
  LField.Precision := APrecision;
  LField.Required := ARequired;
  LField.DisplayLabel := ADisplayLabel;
  LField.DisplayFormat := ADisplayFormat;
  LField.EditFormat := AEditFormat;
  LField.EditMask := AEditMask;
  LField.DisplayWidth := ADisplayWidth;
  LField.KeyFields := AKeyFields;
  LField.LookupDataSetName := ALookupDataSetName;
  LField.LookupKeyFields := ALookupKeyFields;
  LField.LookupResultField := ALookupResultField;
  LField.Lookup := ALookup;
end;

procedure EnsureRuntimeFields(AOwner: TComponent);
var
  LDataSets: TObjectList<TDataSet>;
  LPreviouslyActive: TList<TDataSet>;
  LDataSet: TDataSet;
  LSpec: TDataSetSpec;
  LMarker: TRuntimeFieldsMarker;
begin
  if (AOwner = nil) or HasRuntimeFieldsMarker(AOwner) then
    Exit;
  LDataSets := TObjectList<TDataSet>.Create(False);
  LPreviouslyActive := TList<TDataSet>.Create;
  try
    CollectDataSets(AOwner, LDataSets);
    for LDataSet in LDataSets do
    begin
      LSpec := FindDataSetSpec(AOwner.ClassType, LDataSet.Name);
      if LSpec <> nil then
      begin
        if LDataSet.Active then
          LPreviouslyActive.Add(LDataSet);
        RebuildDataSet(AOwner, LDataSet, LSpec);
      end;
    end;
    for LDataSet in LDataSets do
    begin
      LSpec := FindDataSetSpec(AOwner.ClassType, LDataSet.Name);
      if (LSpec <> nil) and (LSpec.OpenAfterSetup or
         LPreviouslyActive.Contains(LDataSet)) and not LDataSet.Active then
        LDataSet.Open;
    end;
    LMarker := TRuntimeFieldsMarker.Create(AOwner);
  finally
    LPreviouslyActive.Free;
    LDataSets.Free;
  end;
end;

initialization
  GDataSets := TObjectList<TDataSetSpec>.Create(True);

finalization
  GDataSets.Free;

end.
