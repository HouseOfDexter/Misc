unit LoggerU;
{To get a detailed Stack Trace from Unhandled Exceptions,
 in Delphi select->Project/Options/Linker/Map file/Detailed}

interface

uses
  Classes, SysUtils, Windows, mtCheck, forms, JclDebug, JclHookExcept, StrUtils,
  ADODB, DB, Variants, TypInfo, DateUtils;


const
  cSlash = '/';
  cSpace = ' ';
  cTab = '  ';
  cLF = #13 + #10;  //added carraige return so Notepad would display correctly
  cCR = #13;
  cLogExt= '.log';
  cStar= '******************************************************************************';
  cNewLog = 'Newlog';
  cQuotedFldType = [ftString, ftDate, ftTime, ftDateTime, ftMemo,  ftFixedChar, ftWideString, ftFixedWideChar, ftWideMemo];


type
  TLogEvent = procedure(aLog: TStream; aLevel: integer; aHeader, aName, aValue: string) of object;

  ILogger = Interface(IInterface)
  ['{E18DC5DF-4B34-4BE3-B58B-B844A2BCA53A}']
    procedure SaveLog(aLog: Boolean = True);
    procedure AddLog(aValue: string); overload;
    procedure AddLog(aName, aValue: string; aLevel: integer= 0); overload;
    procedure AddLog(aHeader: string; aIsHeader: boolean; aAddDate: Boolean = True); overload;
    procedure AddLog(aStrings: TStrings); overload;
    procedure AddLog(aQry: TAdoQuery); overload;
    procedure AddLog(aQry: TAdoQuery; aName: string); overload;
    procedure AddLog(aTypeInfo: PTypeInfo; aIndex: integer); overload;
    procedure AddLogStrings(const aNames: array of string; const aStrings: array of string);
    procedure AddLog(aQry: TAdoQuery; aLoadFields: boolean); overload;
    procedure AddLog(const aFields: array of TField; aLogAllRecords: boolean=False); overload;
    procedure SetUpQueryLogging(aParent: TComponent);
    procedure SetUpConnectionLogging(aParent: TComponent);
    procedure AddLogLvl(aHeader, aName, aValue: string; aLevel: integer; aIsHeader, aAddDate: boolean; aError: boolean);
    procedure AddLogStar;
    procedure AddLogLine;
    procedure AddLog(aChar: Char); overload;
    procedure AddFileExists(aFileName: string);
    procedure AddError(aHandler, aClassName, aErrorMessage: string); overload;
    procedure AddError(aErrorList: TStrings); overload;
    function GetDeepDebug: boolean;
    procedure SetDeepDebug(const Value: Boolean);
    function GetLogDir: string;
    procedure SetLogDir(const Value: string);
    procedure SetMaxSize(const Value: Integer);
    function GetRefCount: integer;

    procedure ResetAppName;
    procedure SetLogDestroy(const Value: Boolean);
    procedure SetNewLog(const Value: boolean);
    function GetLogLevel: integer;
    procedure SetLogLevel(const Value: integer);

    property LogDir: string read GetLogDir write SetLogDir;
    property DeepDebug: Boolean read GetDeepDebug write SetDeepDebug;
    property MaxSize: Integer write SetMaxSize;  //1 = 1MByte
    property ReferenceCount: Integer read GetRefCount;
    property LogDestroy: Boolean write SetLogDestroy;
    property NewLog: boolean write SetNewLog;
    property LogLevel: integer read GetLogLevel write SetLogLevel;
  end;

  IReport = Interface(ILogger)
  ['{89CF0F2F-24BA-41DB-95C5-BD2155F2AB38}']
  End;

  TLogger = class(TInterfacedObject, ILogger)
  private
    FLogStream: TStream;
    FErrorStream: TStream;
    FOnAddLog: TLogEvent;
    FAppName: string;
    FLogDir: string;
    FAppDir: string;
    FDeepDebug: Boolean;
    FSaveAppName: string;
    FLogDestroy: boolean;
    FNewLog: boolean;
    FInitial: boolean;
    FDif: TDateTime;
    FLogLevel: integer;
    procedure setAppName(const Value: string);
    procedure SaveStream(aLog: Boolean; aStream: TStream);
    procedure CheckForErrorStream;
    class var _Instance: TLogger;
    class var _Application: TApplication;
    procedure SetMaxSize(const Value: Integer);
    function GetRefCount: Integer;
    procedure SetNewLog(const Value: boolean);
    procedure SetLogLevel(const Value: integer);    
    function GetDeepDebug: Boolean;
    procedure BeforeQueryOpen(DataSet: TDataSet);
    procedure AfterQueryOpen(DataSet: TDataSet);
    procedure AfterQueryClose(DataSet: TDataSet);    
    procedure WillExecute(Connection: TADOConnection;
  var CommandText: WideString; var CursorType: TCursorType;
  var LockType: TADOLockType; var CommandType: TCommandType;
  var ExecuteOptions: TExecuteOptions; var EventStatus: TEventStatus;
  const Command: _Command; const Recordset: _Recordset);
    procedure ExecuteComplete(Connection: TADOConnection;
  RecordsAffected: Integer; const Error: Error; var EventStatus: TEventStatus;
  const Command: _Command; const Recordset: _Recordset);
    function GetLogLevel: integer;
  protected
    property AppName: string read FAppName write setAppName;
    property AppDir: string read FAppDir write FAppDir;

    procedure SetDeepDebug(const Value: Boolean);
    procedure SetLogDir(const Value: string);
    function GetLogDir: string;
    procedure SetLogDestroy(const Value: Boolean);
    constructor CreateNew(aAppName: string; aSetInstance: boolean);
    property SaveAppName: string read FSaveAppName write FSaveAppName;
  public
    class function BooleanAsString(aBoolean: Boolean): string;
    procedure SaveLog(aLog: Boolean = False);
    procedure AddLog(aValue: string); overload;
    procedure AddLog(aName, aValue: string; aLevel: integer= 0); overload;
    procedure AddLog(aHeader: string; aIsHeader: boolean; aAddDate: Boolean = True); overload;
    procedure AddLog(aChar: Char); overload;
    procedure AddLog(aQry: TAdoQuery); overload;
    procedure AddLogStrings(const aNames: array of string; const aStrings: array of string); overload;
    procedure AddLog(aStrings: TStrings); overload;
    procedure AddLog(aQry: TAdoQuery; aName: string); overload;
    procedure AddLog(aQry: TAdoQuery; aLoadFields: boolean); overload;
    procedure AddLog(const aFields: array of TField; aLogAllRecords: boolean= False); overload;
    procedure AddLog(aTypeInfo: PTypeInfo; aIndex: integer); overload;    
    procedure AddLogLine;
    procedure AddFileExists(aFileName: string);
    procedure AddLogLvl(aHeader, aName, aValue: string; aLevel: integer; aIsHeader, aAddDate: boolean; aError: Boolean = False);
    procedure SetUpQueryLogging(aParent: TComponent);
    procedure SetUpConnectionLogging(aParent: TComponent);    

    procedure AddError(aHandler, aClassName, aErrorMessage: string); overload;
    procedure AddError(aErrorList: TStrings); overload;
    procedure AddLogStar;

    procedure ResetAppName;

    property LogDir: string read GetLogDir write SetLogDir;
    property DeepDebug: Boolean read GetDeepDebug write SetDeepDebug;
    property MaxSize: Integer write SetMaxSize;

    property OnAddLog: TLogEvent read FOnAddLog write FOnAddLog;

    property ReferenceCount: Integer read GetRefCount;

    property LogDestroy: Boolean write SetLogDestroy;
    property NewLog: boolean write SetNewLog;
    property LogLevel: integer read GetLogLevel write SetLogLevel;    

    class function Logger(aApplication: TApplication): ILogger; overload;
    class function Logger(aApplication: TApplication; aDll: string): ILogger; overload;
    class function Logger(aFileName: string; aNew: boolean): ILogger; overload;
    constructor Create; reintroduce;deprecated;
    destructor Destroy; override;
  end;

  TReport = class(TLogger, IReport)
    class function Report(aApplication: TApplication; aReportName: string = ''): IReport;
  end;

implementation

{ TLogger }

uses
  SyncObjs;

var
  Lock: TCriticalSection;

procedure TLogger.AddError(aHandler, aClassName, aErrorMessage: string);
begin
  CheckForErrorStream;
  AddLog('!');
  AddLogLvl(aHandler, '', '', 0, True, True, True);
  AddLogLvl('', 'ClassName', aClassName, 0, False, False, True);
  AddLogLvl('', 'ErrorMessage', aErrorMessage, 0, False, False, True);
  AddLogLvl(aHandler, '', '', 0, False, True, True);
  AddLog('!');
  SaveLog(False);
end;

procedure TLogger.AddError(aErrorList: TStrings);
var
  a_Index: integer;
begin
  AddLogLvl('Unhandled Exception', '', '', 0, True, True, True);
  CheckForErrorStream;
  for a_Index := 0 to aErrorList.Count - 1 do
    if (a_Index = 0) or (a_Index = aErrorList.Count - 1) then
      AddLogLvl(aErrorList.Strings[a_Index], '', '' , 0, True, False, True)
    else
      AddLogLvl('', '', aErrorList.Strings[a_Index] , 0, False, False, True);
  AddLogLvl('Unhandled Exception', '', '', 0, False, True, True);
  SaveLog(True);
end;

procedure TLogger.AddFileExists(aFileName: string);
begin
  AddLog('FileName', aFileName);
  AddLog('FileExists', BooleanAsString(FileExists(aFileName)));
end;

procedure TLogger.AddLog(aHeader: string; aIsHeader, aAddDate: Boolean);
begin
  AddLogLvl(aHeader, '', '', 0, aIsHeader, aAddDate)
end;

procedure TLogger.AddLog(aName, aValue: string; aLevel: integer);
begin
  if aLevel = 0 then
    AddLogLvl('', aName, aValue, LogLevel, False, False)
  else
    AddLogLvl('', aName, aValue, aLevel, False, False);
end;

procedure TLogger.AddLog(aValue: string);
begin
  AddLogLine;
  AddLogLvl('', '', aValue, 0, False, False);
end;

procedure TLogger.AddLogLine;
begin
  AddLogLvl('', '', ' ', 0, False, False);
end;

procedure TLogger.AddLogLvl(aHeader, aName, aValue: string; aLevel: integer;
  aIsHeader, aAddDate: boolean; aError: boolean);
var
  a_Level: string;
  a_Index: integer;
  a_Stream: TStream;
begin
  if aError then
    a_Stream := FErrorStream
  else
    a_Stream := FLogStream;

  if Assigned(a_Stream) then
  begin
    try
      Lock.Acquire;
      if aHeader <> ''  then
      begin
        if aAddDate then
          aHeader  := aHeader + '=' + DateTimeToStr(Now);
        if not aIsHeader then
          aHeader := cSlash + aHeader;
      end;
      if Assigned(FOnAddLog) then
        FOnAddLog(a_Stream, aLevel, aHeader, aName, aValue);

      a_Level := '';
      if (aHeader <> '') or (aValue <> '') then

      for a_Index := 1 to aLevel do
        a_Level := a_Level + cTab;
      if aHeader <> '' then
      begin
        aHeader := a_Level + '[' + aHeader + ']' + cLF;
        if aIsHeader then
          aHeader := cLF + aHeader
      end;
      //a_Level := a_Level + cTab;
      if aName <> '' then
        aName := aName + '=';
      if aValue <> '' then
        aValue := a_Level + aName + aValue + cLF;
      if aHeader <> '' then
        a_Stream.Write(Pointer(aHeader)^,  length(aHeader));
      if aValue <> '' then
        a_Stream.Write(Pointer(aValue)^,  length(aValue));
    finally
      Lock.Release;
    end;
  end;
end;

procedure TLogger.BeforeQueryOpen(DataSet: TDataSet);
begin
  if (Assigned(Dataset)) and (DataSet is TAdoQuery) then
  begin
    AddLogStar;
    AddLog(TAdoQuery(DataSet));
  end;
end;

class function TLogger.BooleanAsString(aBoolean: Boolean): string;
begin
  if aBoolean then
    Result := 'True'
  else
    Result := 'False';
end;

constructor TLogger.Create;
begin
  Raise Exception.Create('Do not call the constructor see Logger');
end;

constructor TLogger.CreateNew(aAppName: string; aSetInstance: boolean);
begin
  if aSetInstance then
    _Instance := Self;
  FInitial := True;
  FLogStream := TMemoryStream.Create;
  AppName := aAppName;
end;

destructor TLogger.Destroy;
begin
  if FLogDestroy then
  begin
    AddLog('Destroy');
    AddLog('Destroy', False);
  end;
  SaveLog(false);
  FLogStream.Free;
  if Assigned(FErrorStream) then
    FErrorStream.Free;
  _Instance := nil;    
  inherited;
end;

procedure TLogger.ExecuteComplete(Connection: TADOConnection;
  RecordsAffected: Integer; const Error: Error; var EventStatus: TEventStatus;
  const Command: _Command; const Recordset: _Recordset);
var
  a_Index: integer;
begin
  AddLog('AdoConnection ExecuteComplete', True);
  AddLog('Execution In MilliSeconds', IntToStr(MilliSecondsBetween(Time, FDif)));
  AddLog('Execution In Seconds', IntToStr(SecondsBetween (Time, FDif)));
  AddLog('Execution In Minutes', IntToStr(MinutesBetween (Time, FDif)));
  AddLog('CommandText', Command.CommandText);
  if Assigned(Command) then
  begin
    AddLog('Param Count', IntToStr(Command.Parameters.Count));
    for a_Index := 0 to Command.Parameters.Count - 1 do
    begin
      AddLog(Command.Parameters.Item[a_Index].Name, VarToWideStr(Command.Parameters.Item[a_Index].Value));
    end;
    AddLog('CommandType', GetEnumName(TypeInfo(TCommandType),Integer(Command.CommandType)));
  end;
  AddLog('EventStatus', GetEnumName(TypeInfo(TEventStatus),Integer(EventStatus)));
  if Assigned(RecordSet) then
  begin
    AddLog('CursorType', GetEnumName(TypeInfo(TCursorType),Integer(Recordset.CursorType)));
    AddLog('LockType',  GetEnumName(TypeInfo(TADOLockType),Integer(Recordset.LockType)));
  end;
  AddLog('RecordsAffected',  IntToStr(RecordsAffected));
  AddLog('AdoConnection ExecuteComplete', False);
end;

procedure TLogger.AddLogStar;
begin
  AddLogLvl(cStar, '', '', 0, True, False, False);
end;

procedure TLogger.AfterQueryClose(DataSet: TDataSet);
begin
  if (Assigned(Dataset)) and (DataSet is TAdoQuery) then
  begin
    AddLog('Query Closing', Dataset.Name);
  end;
end;

procedure TLogger.AfterQueryOpen(DataSet: TDataSet);
begin
  if (Assigned(Dataset)) and (DataSet is TAdoQuery) then
  begin
    if Assigned(TADOQuery(DataSet).Connection) then
      AddLog('Connection', TADOQuery(DataSet).Connection.Name)
    else
      AddLog('ConnectionString','WARNING: '+ Dataset.Name + ' using ConnectionString');
    AddLog('Query Opening', Dataset.Name);
    AddLog('RecordCount', IntToStr(TAdoQuery(DataSet).RecordCount ));
    AddLog('_');
    SaveLog;
  end;
end;

procedure TLogger.CheckForErrorStream;
begin
  if Assigned(_Application) and not Assigned(FErrorStream) then
  begin
    FErrorStream := TMemoryStream.Create;
  end;
end;

procedure TLogger.SaveStream(aLog: Boolean; aStream: TStream);
var
  a_UseExisting: Boolean;
  a_FileStream: TFileStream;
  a_LogFile: string;
begin
  if Assigned(aStream) and (aStream.Size > 0) then
  begin
    Lock.Acquire;
    try
      a_LogFile := LogDir + AppName + cLogExt;
      a_UseExisting := FileExists(a_LogFile) and (not (FNewLog and FInitial));
      FInitial := False;
      if aLog then
      begin
        AddLog('LogFile', a_LogFile);
        AddLog('Use Existing File', BooleanAsString(a_UseExisting));
      end;
      if a_UseExisting then
      begin
        a_FileStream := TFileStream.Create(a_LogFile, fmOpenReadWrite);
        a_FileStream.Position := a_FileStream.Size;
      end
      else
      begin
        a_FileStream := TFileStream.Create(a_LogFile, fmCreate);
      end;
      try
        a_FileStream.CopyFrom(aStream, 0);
        TMemoryStream(aStream).Clear;
      finally
        a_FileStream.Free;
      end;
    finally
      Lock.Release;
    end;
  end;
end;

function TLogger.GetDeepDebug: Boolean;
begin
  Result := FDeepDebug;
end;

function TLogger.GetLogDir: string;
begin
  if not DirectoryExists(FLogDir) then
    Result := AppDir
  else
    Result := FLogDir;
end;

function TLogger.GetLogLevel: integer;
begin
  Result := FLogLevel;
end;

function TLogger.GetRefCount: Integer;
begin
  Result := Self.RefCount;
end;

class function TLogger.Logger(aFileName: string;  aNew: boolean): ILogger;
begin
  Result := TLogger.CreateNew(aFileName, false);
end;

class function TLogger.Logger(aApplication: TApplication;
  aDll: string): ILogger;
begin
  Lock.Acquire;
  if not Assigned(_Application) then
    _Application := aApplication;
  try
    if not Assigned(_Instance) then
    begin
      Result := TLogger.CreateNew(aDll, true);
    end else
      Result := _Instance;
    _Instance.SaveAppName := _Instance.AppName;
    _Instance.AppName := aDLL;
  finally
    Lock.Release;
  end;
end;

procedure TLogger.ResetAppName;
begin
  AppName := SaveAppName;
end;

class function TLogger.Logger(aApplication: TApplication): ILogger;
begin
  Lock.Acquire;
  if not Assigned(_Application) then
    _Application := aApplication;
  try
    if not Assigned(_Instance) then
    begin
      Result := TLogger.CreateNew(_Application.ExeName, true);
    end else
      Result := _Instance;
  finally
    Lock.Release;
  end;
end;

procedure TLogger.SaveLog(aLog: Boolean);
begin
  try
    if aLog then
      AddLog('SaveLog');
    if FDeepDebug then
    begin
      SaveStream(aLog, FLogStream);
    end;
    SaveStream(aLog, FErrorStream);
    if aLog then
      AddLog('SaveLog', False);
  except
  //we swallow any exceptions
  end;
end;

procedure TLogger.setAppName(const Value: string);
var
  a_Value: string;
  a_Len: integer;
begin
  a_Value := Trim(Value);

  FAppName := ExtractFileName(a_Value);
  FAppDir :=  ExtractFilePath(a_Value);
  a_Len := Length(FAppName) -4;
  TmtCheck.IsGreaterThan(a_Len, 0, 'Invalid AppName length, No extension for exe', 'AppName');
  setLength(FAppName, a_Len); // get rid of .exe
end;

procedure TLogger.setDeepDebug(const Value: Boolean);
begin
  FDeepDebug := Value;
end;

procedure TLogger.SetLogDestroy(const Value: Boolean);
begin
  FLogDestroy := Value;
end;

procedure TLogger.setLogDir(const Value: string);
begin
  FLogDir := IncludeTrailingPathDelimiter(Trim(Value));
  if not DirectoryExists(FLogDir) then
    FLogDir := AppDir;
end;

procedure TLogger.SetLogLevel(const Value: integer);
begin
  FLogLevel := Value;
end;

procedure TLogger.SetMaxSize(const Value: Integer);
begin
//temp not done finish
end;

procedure TLogger.SetNewLog(const Value: boolean);
begin
  FNewLog := Value;
end;

procedure TLogger.SetUpConnectionLogging(aParent: TComponent);
var
  a_Index: integer;
begin
  for a_Index := 0 to aParent.ComponentCount - 1 do
    if aParent.Components[a_Index] is TAdoConnection then
    begin
      TAdoConnection(aParent.Components[a_Index]).OnWillExecute :=  WillExecute;
      TAdoConnection(aParent.Components[a_Index]).OnExecuteComplete :=  ExecuteComplete;
    end;
end;

procedure TLogger.SetUpQueryLogging(aParent: TComponent);
var
  a_Index: integer;
begin
  for a_Index := 0 to aParent.ComponentCount - 1 do
    if aParent.Components[a_Index] is TAdoQuery then
    begin
      TAdoQuery(aParent.Components[a_Index]).BeforeOpen :=  BeforeQueryOpen;
      TAdoQuery(aParent.Components[a_Index]).AfterOpen :=  AfterQueryOpen;
      TAdoQuery(aParent.Components[a_Index]).AfterClose :=  AfterQueryClose;
    end;
end;

procedure TLogger.WillExecute(Connection: TADOConnection;
  var CommandText: WideString; var CursorType: TCursorType;
  var LockType: TADOLockType; var CommandType: TCommandType;
  var ExecuteOptions: TExecuteOptions; var EventStatus: TEventStatus;
  const Command: _Command; const Recordset: _Recordset);
begin
  AddLog('Connection WillExecute', True);
  AddLog('Connection Name', Connection.Name);
  AddLog('CommandText', CommandText);
  AddLog('CommandType', GetEnumName(TypeInfo(TCommandType),Integer(CommandType)));
  AddLog('EventStatus', GetEnumName(TypeInfo(TEventStatus),Integer(EventStatus)));
  AddLog('CursorType', GetEnumName(TypeInfo(TCursorType),Integer(CursorType)));
  AddLog('Connection WillExecute', False);
  FDif :=  Time;
end;

procedure HookGlobalException(ExceptObj: TObject; ExceptAddr: Pointer; OSException: Boolean);
var
  a_List: TStringList;
  a_Error: string;
begin
  if Assigned(TLogger._Instance) then
  begin
    a_List := TStringList.Create;
    try
      a_List.Add(cStar);
      a_Error := Exception(ExceptObj).Message;
      a_List.Add(Format('{ Exception - %s }', [a_Error]));
      JclLastExceptStackListToStrings(a_List, False, True, True, False);
      a_List.Add(cStar);
      // save the error with stack log to file
      TLogger._Instance.AddError(a_List);
    finally
      a_List.Free;
      Raise Exception.Create(a_Error);
    end;
  end;
end;

procedure TLogger.AddLog(aChar: Char);
begin
  AddLogLvl(DupeString(aChar, 78), '', '', 0, True, False, False);
end;

procedure TLogger.AddLog(aQry: TAdoQuery; aName: string);
var
  a_File: TStrings;
  a_Index, a_Params: integer;
  a_ParamName, a_ParamName1,  a_Value: string;
  a_Variant: variant;
  a_Null: boolean;
  a_Field: TField;
  a_LevelStr: string;
begin
  a_LevelStr := '';
  for a_Index := 1 to LogLevel do
    a_LevelStr := a_LevelStr + cTab;

  a_File := TStringList.Create;
  try
    a_File.Assign(aQry.SQL);
    for a_Params := 0 to aQry.Parameters.Count - 1 do
    begin
      a_ParamName := aQry.Parameters.Items[a_Params].Name;
      a_Variant := aQry.Parameters.Items[a_Params].Value; //
      a_Null := False;
      if VarIsNull(a_Variant) then  //if we have a null parameter...
      begin//then check if the query has a Datasource and get it's value from the field that matches the ParamName
        if Assigned(aQry.DataSource) then
        begin
          a_Field := aQry.DataSource.DataSet.FindField(a_ParamName);
          if Assigned(a_Field) then
            a_Value := a_Field.AsString
          else
            a_Value := 'Unknown Value';            
        end
        else
        begin
          a_Value := 'Null';
          a_Null := True;
        end;
      end else
      begin
        a_Value := VarToStr(aQry.Parameters.Items[a_Params].Value);
      end;
      if (aQry.Parameters.Items[a_Params].DataType in cQuotedFldType) and not a_Null then
          a_Value := QuotedStr(a_Value);
      a_ParamName1 := '@' + a_ParamName;
      a_ParamName := ':'+ a_ParamName;
      for a_Index := a_File.Count - 1 downto 0 do
      begin
        a_File[a_Index] := a_File[a_Index];
        a_File[a_Index] := StringReplace(a_File[a_Index],a_ParamName, a_Value, []);
        if a_File[a_Index] = '' then
           a_File[a_Index] := StringReplace(a_File[a_Index],a_ParamName1, a_Value, []);
        if a_File[a_Index] = '' then
          a_File.Delete(a_Index);
      end;
    end;

    AddLog('--'+aName, 'SQL', 0);
    AddLogLine;
    AddLog(a_File);
    AddLogLine;
//    AddLogLvl('', a_File.Text)
  finally
    a_File.Free;
  end;
end;

procedure TLogger.AddLog(aStrings: TStrings);
var
  a_Index: integer;
begin
  for a_Index := 0 to aStrings.Count - 1 do
    AddLog('', aStrings[a_Index]);
end;

procedure TLogger.AddLog(aQry: TAdoQuery);
begin
  AddLog(aQry, aQry.Name);
end;

procedure TLogger.AddLog(aQry: TAdoQuery; aLoadFields: boolean);
var
  a_Index: integer;
begin
  if aLoadFields then  
    for a_Index := 0 to aQry.FieldCount - 1 do
      AddLog('--FieldName', aQry.Fields[a_Index].Name);
  AddLog(aQry);
end;

procedure TLogger.AddLog(const aFields: array of TField; aLogAllRecords: boolean);
var
  a_BM: string;
  procedure _AddLog(const aFields: array of TField);
  var
    a_Index: integer;
  begin
    for a_Index := Low(aFields) to High(aFields) do
      if aFields[a_Index].DataType in [ftSmallint, ftInteger] then
        AddLog(aFields[a_Index].FieldName, IntToStr(aFields[a_Index].AsInteger))
      else
        AddLog(aFields[a_Index].FieldName, aFields[a_Index].AsString);
  end;
begin
  if aLogAllRecords then
  begin
    a_BM:= aFields[0].DataSet.BookMark;
    aFields[0].DataSet.First;
    While not aFields[0].DataSet.Eof do
    begin
      _AddLog(aFields);
      aFields[0].DataSet.Next;
    end;
    aFields[0].DataSet.BookMark := a_BM;
  end else
    _AddLog(aFields);
end;

procedure TLogger.AddLogStrings(const aNames: array of string; const aStrings: array of string);
var
  a_Index: integer;
begin
  for a_Index := Low(aStrings) to High(aStrings) do
    AddLog(aNames[a_Index] , aStrings[a_Index]);
end;

procedure TLogger.AddLog(aTypeInfo: PTypeInfo; aIndex: integer);
begin
  AddLog(GetEnumName(aTypeInfo, aIndex),IntToStr(aIndex));
end;

{ TReport }

class function TReport.Report(aApplication: TApplication; aReportName: string = ''): IReport;
begin
  Lock.Acquire;
  if not Assigned(_Application) then
    _Application := aApplication;
  try
    if aReportName = '' then
      Result := CreateNew(_Application.ExeName, False)
    else
      Result := CreateNew(aReportName, False);
  finally
    Lock.Release;
  end;
end;

initialization
  Lock := TCriticalSection.Create;
  Include(JclStackTrackingOptions, stTraceAllExceptions);
  Include(JclStackTrackingOptions, stRawMode);

  // Initialize Exception tracking
  JclStartExceptionTracking;

  JclAddExceptNotifier(HookGlobalException, npFirstChain);
  JclHookExceptions;

finalization
  JclUnhookExceptions;
  JclStopExceptionTracking;
  Lock.Free;

end.
