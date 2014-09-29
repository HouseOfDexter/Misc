unit SQLBuilderU;

interface

uses
 Classes, SysUtils, Controls, WideStrings, mtCheck, UtilsU, db, LoggerU;

type
  TDBServerType = (dtUnknown, dtSQLServer, dtDB2, dtMySQL);
  TSQLType = (sqSelect, sqInsert, sqInsertSelect, sqUpdate, sqDelete, sqProc);

  TSQLTokenType = (ttLineBefore, ttSelect, ttJoin, ttWhere);

  TComparisonType = (ctEqual, ctNotEqual, ctLesser, ctLesserEqual,
     ctGreater, ctGreaterEqual, ctLike, ctLikeAll, ctLikeEnd);

  TJoinType = (jtNone, jtFrom, jtInner, jtLeftOuter, jtRightOuter, jtWhere, jtUpdate, jtSelect);
  TWhereType = (wtStart, wtBegin, wtAnd, wtOr, wtExists, wtNotExists, wtEnd, wtToken);
  TWhereTypes = set of TWhereType;

  TSchemaList = class;
  TSQLTokenList = class;

  TSQLBuilderType =(btWide, btString);

  TSQLBuilder = class(TObject)
  private
    FDBServerType: TDBServerType;
    FDataBaseName: string;
    FSQLType: TSQLType;
    FIsReadOnly: boolean;
    FSQL: TWideStrings;
    FSQLString: TStrings;
    FColumns: TStrings;
    FActualValues: TStrings;
    FSelectTopCount: integer;
    FSchemaList: TSchemaList;
    FSQLTokenList: TSQLTokenList;
    FSelectDistinct: boolean;
    FOrderBy: string;
    function getDBServerType: TDBServerType;
    procedure setDBServerType(const Value: TDBServerType);
    function getSQL: TWideStrings;
   
    procedure BuildColumns(aColumns, aSource, aDestColumns: TStrings; aQuote: boolean = false; aGroup: boolean = false);
    function getSelectTopCount: integer;
    function getSelectTopCountStr: string;
    function getFetchCountStr: string;
    function getReadOnlyEndStr: string;
    function GetSQLString: TStrings;
  protected
    function GetSchema(aColumn: string): string;
    procedure BuildSQL(aType: TSQLBuilderType = btWide);
    procedure AddSQLToken(aSelect: string; aBuild: boolean; aUseSchema: boolean); overload;
    procedure AddSQLToken(aLine: string; aIsLine: boolean; aBuild: boolean; aUseSchema: boolean); overload;
    procedure AddSQLToken(aJoinType: TJoinType; aTableName, aTableAlias, aOnJoin: string; aBuild: boolean; aUseSchema: boolean); overload;
    procedure AddSQLToken(aTableName, aTableAlias, aColumn, aWhere: string; aWhereType: TWhereType; aValue: Variant;
     aComparisonType: TComparisonType; aFieldType: TFieldType; aBuild: boolean; aUseSchema: boolean); overload;
    procedure AddSQLToken(aTableName, aTableAlias, aColumn, aWhere: string; aWhereTypes: TWhereTypes; aValue: Variant;
     aComparisonType: TComparisonType; aFieldType: TFieldType; aBuild: boolean; aUseSchema: boolean); overload;

    procedure ClearSQLTokens;
    procedure BuildBefore(aBefore: TStrings);
    procedure BuildSelectHeader(aHeader: TStrings);
    procedure BuildInsertHeader(aHeader: TStrings);
    procedure BuildUpdateHeader(aHeader: TStrings);
    procedure BuildDeleteHeader(aHeader: TStrings);
    procedure BuildInsertColumns(aColumns: TStrings);
    procedure BuildJoin(aJoin: TStrings);
    procedure BuildValues(aValues: TStrings);
    procedure BuildSelectQry(aSelect: TStrings);

    procedure BuildWhere(aWhere: TStrings);
  public
     constructor Create;
     destructor Destroy; override;

     function AsString: string;
     procedure ClearBuilder;
     procedure ClearSchemaInfo;

     procedure AddSchema(aDBServerType: TDBServerType; aSchemaName, aTable: string); overload;
     procedure AddSchema(aDBServerType: TDBServerType; aSchemaName: string; const aTables: Array of string); overload;

     property SelectDistinct: boolean read FSelectDistinct write FSelectDistinct;
     property SelectTopCount: integer read getSelectTopCount write FSelectTopCount;
     procedure AddColumn(aColumn: string; aFieldType: TFieldType= ftString; aBuild: boolean = true); overload;
     procedure AddColumn(aColumn: string; aFieldType: TFieldType; aValue: Variant; aBuild: boolean = true); overload;
     procedure AddValue(aValue: variant; aFieldType: TFieldType); 

     procedure AddColumns(const aColumns: Array of string);overload;
     procedure AddColumns(const aColumns: Array of string; const aFieldTypes: Array of TFieldType);overload;
     procedure AddColumns(const aColumns: Array of string; const aFieldTypes: Array of TFieldType; const aValues: Array of Variant);overload;
     procedure AddFrom(aTableName, aTableAlias: string; aUseSchema: boolean = True);
     procedure AddInnerJoin(aTableName, aTableAlias, aOn: string; aBuild: boolean = True; aUseSchema: boolean = True);
     procedure AddLeftOuterJoin(aTableName, aTableAlias, aOn: string; aBuild: boolean = True; aUseSchema: boolean = True);
     procedure AddRightOuterJoin(aTableName, aTableAlias, aOn: string; aBuild: boolean = True; aUseSchema: boolean = True);
     procedure AddUpdate(aTableName: string);
     procedure AddInsert(aTableName: string);     

     procedure AddOrderBy(aColumns: string);

     procedure NeedSQLToken(aTableName: string; aBuild: boolean = True);
     procedure NeedWhereToken(aTableName: string; aBuild: boolean = True);

     procedure BeginGroup;
     procedure EndGroup;
     procedure AddWhere(aWhere: string; aWhereType: TWhereType = wtToken;aBuild: boolean = true);overload;
     procedure AddWhere(aWhere: string; aWhereTypes: TWhereTypes;aBuild: boolean = true);overload;
     procedure AddWhere(aColumn: string; aColFieldType: TFieldType; aValue: Variant; aComparisonType: TComparisonType; aWhereType: TWhereType; aBuild: boolean = true); overload;
     procedure AddWhere(aSQL: TSQLBuilder; aWhereType: TWhereType;aBuild: boolean = true); overload;
     procedure AddWhere(aSQL: TSQLBuilder; aWhereTypes: TWhereTypes;aBuild: boolean = true); overload;
     procedure AddSelect(aSQL: TSQLBuilder; aBuild: boolean = true; aUseSchema: boolean= true); overload;
     procedure AddSelect(aSQL: string; aBuild: boolean = true; aUseSchema: boolean = True); overload;
     procedure AddLine(aLine: string; aBuild: boolean = true; aUseSchema: boolean = True); overload;     
     function GroupWhere(aClear: boolean): string;
     procedure Assign(aSource: TSQLBuilder);

     property DBServerType: TDBServerType read getDBServerType write setDBServerType;
     property DataBaseName: string read FDataBaseName write FDataBaseName;
     property IsReadOnly: boolean read FIsReadOnly write FIsReadOnly;
     property SQLType: TSQLType read FSQLType write FSQLType;

     property SQL: TWideStrings read GetSQL;
     property SQLString: TStrings read GetSQLString;
  end;

  TSchema = class(TObject)
  private
    FSchema: string;
    FTables: TStrings;
  private
    FDBServerType: TDBServerType;
  public
    constructor Create;
    destructor Destroy; override;
    property DBServerType: TDBServerType read FDBServerType write FDBServerType;
    property Schema: string read FSchema write FSchema;
    procedure AddTable(aTable: string);
    procedure AddTables(const aTables: array of string);overload;
    procedure AddTables(aTables: TStrings); overload;
    procedure ClearSchema;
    function HasTable(aTable: string): boolean;
  end;

  TSchemaList = class(TObject)
  private
    FList: TList;
    function findSchema(aDBServerType: TDBServerType; aSchema: string): TSchema;
    function addSchemaInfo(aDBServerType: TDBServerType; aSchemaName: string): TSchema;
  public
    constructor Create;
    destructor Destroy; override;
    procedure AddSchema(aDBServerType: TDBServerType; aSchemaName, aTable: string); overload;
    procedure AddSchema(aDBServerType: TDBServerType; aSchemaName: string; const aTables: Array of string); overload;

    function GetSchemaName(aDBServerType: TDBServerType; aTable: string = 'DEFAULT'): string;
    procedure ClearSchemas;
  end;

  TSQLToken = class(TObject)
  private
    FTableName: string;
    FTableAlias: string;
    FBuild: boolean;
    FSQLTokenType: TSQLTokenType;
    FIsReadOnly: boolean;
    FDBServerType: TDBServerType;
    FUseSchema: boolean;
  protected

  public
    procedure ClearToken; virtual;
    function AsString(aDBServerType: TDBServerType; aSchemaList: TSchemaList): string; virtual;
    property TableName: string read FTableName write FTableName;
    property TableAlias: string read FTableAlias write FTableAlias;
    property Build: boolean read FBuild write FBuild;
    property SQLTokenType: TSQLTokenType read FSQLTokenType write FSQLTokenType;
    property IsReadOnly: boolean read FIsReadOnly write FIsReadOnly;
    property DBServerType: TDBServerType read FDBServerType write FDBServerType;
    property UseSchema: boolean read FUseSchema write FUseSchema;
  end;

  TSQLBeforeToken = class(TSQLToken)
  private
    FLine: string;
  public
    Constructor Create;
    function AsString(aDBServerType: TDBServerType; aSchemaList: TSchemaList): string; override;
    property Line: string read FLine write FLine;
  end;

  TSQLSelectToken = class(TSQLToken)
  private
    FSelect: string;
  public
    Constructor Create;
    function AsString(aDBServerType: TDBServerType; aSchemaList: TSchemaList): string; override;
    property Select: string read FSelect write FSelect;
  end;


  TSQLJoinToken = class(TSQLToken)
  private
    FOnJoin: string;
    FJoinType: TJoinType;
    FSQLType: TSQLType;
    function getTableLockStr: string;
  public
    Constructor Create;
    function AsString(aDBServerType: TDBServerType; aSchemaList: TSchemaList): string; override;
    property OnJoin: string read FOnJoin write FOnJoin;
    property JoinType: TJoinType read FJoinType write FJoinType;
    property SQLType: TSQLType read FSQLType write FSQLType;
  end;

  TSQLWhereToken = class(TSQLToken)
  private
    FWhereTypes: TWhereTypes;
    FWhere: string;
    FComparisonType: TComparisonType;
    FValue: Variant;
    FColumn: string;
    FFieldType: TFieldType;
    function getQuoteField: boolean;
  published
    Constructor Create;
    function AsString(aDBServerType: TDBServerType; aSchemaList: TSchemaList): string; override;
    procedure AddWhereType(aWhereType: TWhereType);
    procedure AddWhereTypes(aWhereTypes: TWhereTypes);    
    property WhereTypes: TWhereTypes read FWhereTypes write FWhereTypes;
    property Where: string read FWhere write FWhere;
    property ComparisonType: TComparisonType read FComparisonType write FComparisonType;
    property Column: string read FColumn write FColumn;
    property Value: Variant read FValue write FValue;
    property QuoteField: boolean read getQuoteField;
    property FieldType: TFieldType read FFieldType write FFieldType;
  end;


  TSQLTokenList = class(TObject)
  private
    FList: TList;
    FIndex: integer;

    function findSQLToken(aTableName: string; aTokenType: TSQLTokenType): TSQLToken;
  public
    constructor Create;
    destructor Destroy; override;
    procedure ClearTokens;
    procedure AddSQLToken(aSQLToken: TSQLToken);
    function BuildToken(aDBServerType: TDBServerType; aSchemaList: TSchemaList;
      aTokenType: TSQLTokenType; aTableName: string; aBuild: boolean): string;
    function FirstToken(aTokenType: TSQLTokenType; aBuild: boolean = True): TSQLToken;
    function LastToken(aTokenType: TSQLTokenType; aBuild: boolean = True): TSQLToken;
    function PrevToken(aTokenType: TSQLTokenType; aBuild: boolean = True): TSQLToken;
    function NextToken(aTokenType: TSQLTokenType; aBuild: boolean = True): TSQLToken;
    function GetToken(aTokenType: TSQLTokenType; aIndex: integer; aBuild: boolean = True): TSQLToken;
  end;

implementation


{ TSQLBuilder }

procedure TSQLBuilder.AddColumn(aColumn: string; aFieldType: TFieldType; aBuild: boolean);
begin
  if aBuild then
    FColumns.AddObject(aColumn, pointer(aFieldType));
end;

procedure TSQLBuilder.AddColumns(const aColumns: array of string; const aFieldTypes: array of TFieldType);
var
  a_Index: integer;
begin
  for a_Index := low(aColumns) to high(aColumns) do
    AddColumn(aColumns[a_Index], aFieldTypes[a_Index]);
end;

procedure TSQLBuilder.AddColumns(const aColumns: array of string);
var
  a_Index: integer;
begin
  for a_Index := low(aColumns) to high(aColumns) do
    AddColumn(aColumns[a_Index]);
end;

procedure TSQLBuilder.AddColumn(aColumn: string; aFieldType: TFieldType;
  aValue: Variant; aBuild: boolean);
begin
  if aBuild then
  begin
    AddColumn(aColumn, aFieldType);
    AddValue(aValue, aFieldType);
  end;
end;

procedure TSQLBuilder.AddColumns(const aColumns: array of string;
  const aFieldTypes: array of TFieldType; const aValues: array of Variant);
var
  a_Index: integer;
begin
  for a_Index := low(aColumns) to high(aColumns) do
  begin
    AddColumn(aColumns[a_Index], aFieldTypes[a_Index]);
    AddValue(aValues[a_Index], aFieldTypes[a_Index]);
  end;
end;

procedure TSQLBuilder.AddFrom(aTableName, aTableAlias: string; aUseSchema: boolean);
begin
  AddSQLToken(jtFrom, aTableName, aTableAlias, '', True, aUseSchema);
end;

procedure TSQLBuilder.AddInnerJoin(aTableName, aTableAlias, aOn: string; aBuild: boolean; aUseSchema: boolean);
begin
  AddSQLToken(jtInner, aTableName, aTableAlias, aOn, aBuild, aUseSchema);
end;

procedure TSQLBuilder.AddInsert(aTableName: string);
begin
  AddUpdate(aTableName);
end;

procedure TSQLBuilder.AddLeftOuterJoin(aTableName, aTableAlias, aOn: string; aBuild: boolean; aUseSchema: boolean);
begin
  AddSQLToken(jtLeftOuter, aTableName, aTableAlias, aOn, aBuild, aUseSchema);
end;

procedure TSQLBuilder.AddLine(aLine: string; aBuild, aUseSchema: boolean);
begin
  if aBuild then
  begin
    AddSQLToken(aLine, True, True, aUseSchema);
  end;
end;

procedure TSQLBuilder.AddOrderBy(aColumns: string);
begin
  FOrderBy := aColumns;
end;

procedure TSQLBuilder.AddRightOuterJoin(aTableName, aTableAlias, aOn: string;
  aBuild: boolean; aUseSchema: boolean);
begin
  AddSQLToken(jtRightOuter, aTableName, aTableAlias, aOn, aBuild, aUseSchema);
end;

procedure TSQLBuilder.AddSchema(aDBServerType: TDBServerType; aSchemaName,
  aTable: string);
begin
  FSchemaList.AddSchema(aDBServerType,aSchemaName, aTable);
end;

procedure TSQLBuilder.AddSchema(aDBServerType: TDBServerType;
  aSchemaName: string; const aTables: array of string);
begin
  FSchemaList.AddSchema(aDBServerType,aSchemaName, aTables);
end;

procedure TSQLBuilder.AddSelect(aSQL: string; aBuild: boolean; aUseSchema: boolean);
begin
  if aBuild then
  begin
    AddSQLToken(aSQL, True, aUseSchema);
  end;
end;

procedure TSQLBuilder.AddSQLToken(aTableName, aTableAlias, aColumn,
  aWhere: string; aWhereTypes: TWhereTypes; aValue: Variant;
  aComparisonType: TComparisonType; aFieldType: TFieldType; aBuild,
  aUseSchema: boolean);
var
  a_SQL: TSQLWhereToken;
begin
  if aBuild then
  begin
    a_SQL := TSQLWhereToken.Create;
    a_SQL.TableName := aTableName;
    a_SQL.TableAlias := aTableAlias;
    a_SQL.Column := aColumn;
    a_SQL.Where := aWhere;
    a_SQL.AddWhereTypes(aWhereTypes);
    a_SQL.ComparisonType := aComparisonType;
    a_SQL.Build := aBuild;
    a_SQL.FieldType := aFieldType;
    a_SQL.IsReadOnly := IsReadOnly;
    a_SQL.DBServerType := DBServerType;
    a_SQL.UseSchema := aUseSchema;
    FSQLTokenList.AddSQLToken(a_SQL);
  end;
end;

procedure TSQLBuilder.AddSQLToken(aLine: string; aIsLine, aBuild,
  aUseSchema: boolean);
var
  a_SQL: TSQLBeforeToken;
begin
  if aBuild and aIsLine then
  begin
    a_SQL := TSQLBeforeToken.Create;
    a_SQL.Line := aLine;
    a_SQL.UseSchema := aUseSchema;
    a_SQL.DBServerType := DBServerType;
    FSQLTokenList.AddSQLToken(a_SQL);
  end;
end;

procedure TSQLBuilder.AddSelect(aSQL: TSQLBuilder; aBuild: boolean; aUseSchema: boolean);
begin
  if aBuild then
    AddSelect(aSQL.AsString, aBuild, aUseSchema);
end;

procedure TSQLBuilder.AddSQLToken(aTableName, aTableAlias,
  aColumn, aWhere: string; aWhereType: TWhereType; aValue: Variant; aComparisonType: TComparisonType; aFieldType: TFieldType; aBuild: boolean; aUseSchema: boolean);
var
  a_SQL: TSQLWhereToken;
begin
  if aBuild then
  begin
    a_SQL := TSQLWhereToken.Create;
    a_SQL.TableName := aTableName;
    a_SQL.TableAlias := aTableAlias;
    a_SQL.Column := aColumn;
    a_SQL.Where := aWhere;
    a_SQL.AddWhereType(aWhereType);
    a_SQL.ComparisonType := aComparisonType;
    a_SQL.Build := aBuild;
    a_SQL.FieldType := aFieldType;
    a_SQL.IsReadOnly := IsReadOnly;
    a_SQL.DBServerType := DBServerType;
    a_SQL.UseSchema := aUseSchema;
    FSQLTokenList.AddSQLToken(a_SQL);
  end;
end;

procedure TSQLBuilder.AddSQLToken(aSelect: string; aBuild, aUseSchema: boolean);
var
  a_SQL: TSQLSelectToken;
begin
  if aBuild then
  begin
    a_SQL := TSQLSelectToken.Create;
    a_SQL.Select := aSelect;
    a_SQL.UseSchema := aUseSchema;
    a_SQL.DBServerType := DBServerType;
    FSQLTokenList.AddSQLToken(a_SQL);
  end;
end;

procedure TSQLBuilder.AddUpdate(aTableName: string);
begin
  AddSQLToken(jtUpdate, aTableName, '', '', True, True);
end;

procedure TSQLBuilder.AddValue(aValue: variant; aFieldType: TFieldType);
begin
  FActualValues.AddObject(aValue, pointer(aFieldType))
end;

procedure TSQLBuilder.AddWhere(aSQL: TSQLBuilder; aWhereType: TWhereType; aBuild: boolean);
begin
  if aBuild then
    AddWhere('(' + aSQL.AsString + ')', aWhereType);
end;

procedure TSQLBuilder.AddWhere(aWhere: string; aWhereTypes: TWhereTypes;
  aBuild: boolean);
begin
  if aBuild then
    AddSQLToken('', '', '',aWhere, aWhereTypes ,varEmpty, ctEqual, ftString, True, True);
end;

procedure TSQLBuilder.AddWhere(aWhere: string; aWhereType: TWhereType; aBuild: boolean);
begin
  if aBuild then
    AddSQLToken('', '', '',aWhere, aWhereType ,varEmpty, ctEqual, ftString, True, True);
end;

procedure TSQLBuilder.AddWhere(aColumn: string; aColFieldType: TFieldType;
  aValue: Variant; aComparisonType: TComparisonType; aWhereType: TWhereType; aBuild: boolean);
var
  a_Comparison, a_Value, a_Where: string;
begin
  if aBuild then
  begin
    case aComparisonType of
      ctEqual: a_Comparison := '=' ;
      ctNotEqual: a_Comparison := '<>';
      ctLesser: a_Comparison := '<';
      ctLesserEqual: a_Comparison := '<=';
      ctGreater: a_Comparison := '>';
      ctGreaterEqual: a_Comparison := '>=';
      ctLike..ctLikeEnd: a_Comparison := ' Like ';
    end;
    case aComparisonType of
      ctLike: a_Value := QuotedStr('%' + aValue);
      ctLikeAll: a_Value := QuotedStr('%' + aValue + '%') ;
      ctLikeEnd: a_Value := QuotedStr(aValue + '%')
      else
        if aColFieldType in cQuotedFldType  then
         a_Value := QuotedStr(aValue)
        else
         a_Value := aValue;
    end;

    a_Where := aColumn + a_Comparison + a_Value;
    AddSQLToken('', '', '',a_Where,aWhereType  ,varEmpty, ctEqual, ftString, True, aBuild);
  end;
//  AddSQLToken(jtWhere , '', aColumn,aWhere ,aValue, aComparisonType, aFieldType);
end;

procedure TSQLBuilder.AddSQLToken(aJoinType: TJoinType; aTableName, aTableAlias,
  aOnJoin: string; aBuild: boolean; aUseSchema: boolean);
var
  a_SQL: TSQLJoinToken;
begin
  if aBuild then
  begin
    a_SQL := TSQLJoinToken.Create;

    a_SQL.JoinType := aJoinType;
    a_SQL.TableName := aTableName;
    a_SQL.TableAlias := aTableAlias;
    a_SQL.OnJoin := aOnJoin;
    a_SQL.Build := aBuild;
    a_SQL.DBServerType := DBServerType;
    a_SQL.IsReadOnly := IsReadOnly;
    a_SQL.UseSchema := aUseSchema;
    a_SQL.SQLType := SQLType;
    FSQLTokenList.AddSQLToken(a_SQL);
  end;
end;
{
procedure TSQLBuilder.AddWhere(aTable, aTableAlias, aColumn: string; aValue: Variant;
  aComparisonType: TComparisonType; aFieldType: TFieldType; aBuild: boolean);
begin
  AddSQLToken(aTable, aTableAlias, aColumn,'' ,aValue, aComparisonType, aFieldType, aBuild);
end;

procedure TSQLBuilder.AddWhere(aWhere: string; aWhereType: TWhereType = wtAnd);
begin
  AddSQLToken('', '', '',aWhere ,varEmpty, ctEqual, ftString, True);
end;

procedure TSQLBuilder.AddWhere(aTable, aTableAlias, aColumn, aWhere: string;
  aValue: Variant; aComparisonType: TComparisonType; aFieldType: TFieldType;
  aBuild: boolean);
begin
  AddSQLToken(aTable, aTableAlias, aColumn,aWhere ,aValue, aComparisonType, aFieldType, aBuild);
end;
}
procedure TSQLBuilder.Assign(aSource: TSQLBuilder);
begin
{     property DBServerType: TDBServerType read getDBServerType write setDBServerType;
     property DataBaseName: string read FDataBaseName write FDataBaseName;
     property IsReadOnly: boolean read FIsReadOnly write FIsReadOnly;
     property SQLType: TSQLType read FSQLType write FSQLType;
     property WhereType: TWhereType read getWhereType write FWheretype;}
  Self.DBServerType := aSource.DBServerType;
  Self.DataBaseName := aSource.DataBaseName;
end;


procedure TSQLBuilder.BuildBefore(aBefore: TStrings);
var
  a_Token: TSQLJoinToken;
begin
  a_Token := TSQLJoinToken(FSQLTokenList.FirstToken(ttLineBefore));
  if Assigned(a_Token) then
  begin
    aBefore.Add(a_Token.AsString(DBServerType, FSchemaList));
    While Assigned(a_Token) do
    begin
      a_Token := TSQLJoinToken(FSQLTokenList.NextToken(ttLineBefore));
      if Assigned(a_Token) then
        aBefore.Add(a_Token.AsString(DBServerType, FSchemaList));
    end;
  end;
end;

procedure TSQLBuilder.BuildColumns(aColumns, aSource, aDestColumns: TStrings; aQuote: boolean; aGroup: boolean);
var
  a_Index: Integer;
  a_Value, a_Column: string;
begin
  for a_Index := 0 to aSource.Count - 1 do
  begin
    a_Value := aSource[a_Index];
    if aQuote then
      if TFieldType(aColumns.Objects[a_Index]) in cQuotedFldType then
        a_Value := QuotedStr(a_Value);
    if aGroup and (a_Index = 0) then
      a_Column := '(' + a_Value
    else  
    if a_Index = 0 then
      a_Column := a_Value
    else
      a_Column := a_Column + ',' + a_Value;
    if aGroup and (a_Index = aSource.Count -1) then
      a_Column := a_Column + ')';

    if (Length(a_Column) > 78) then
    begin
      a_Column := '  ' + a_Column;
      aDestColumns.Add(a_Column);
      a_Column := '';
    end;
  end;
  if a_Column <> '' then
  begin
    a_Column := '  ' + a_Column;
    aDestColumns.Add(a_Column);
  end;
end;

function TSQLBuilder.AsString: string;
begin
  Result := SQL.Text;
end;

procedure TSQLBuilder.ClearSQLTokens;
begin
  FSQLTokenList.ClearTokens;
end;

procedure TSQLBuilder.BeginGroup;
begin
  AddWhere('(', wtBegin);
end;

procedure TSQLBuilder.BuildDeleteHeader(aHeader: TStrings);
begin
  aHeader.Add('DELETE');
end;

procedure TSQLBuilder.BuildInsertColumns(aColumns: TStrings);
begin
  if (aColumns.Count > 0) then
  begin
    BuildColumns(FColumns, FColumns, aColumns, False, True);
  end;
end;

procedure TSQLBuilder.BuildInsertHeader(aHeader: TStrings);
begin
  aHeader.Add('INSERT INTO');
  BuildJoin(aHeader);
  BuildInsertColumns(aHeader);
end;

procedure TSQLBuilder.BuildJoin(aJoin: TStrings);
var
  a_Token: TSQLJoinToken;
begin
  a_Token := TSQLJoinToken(FSQLTokenList.FirstToken(ttJoin));
  if Assigned(a_Token) then
  begin
    aJoin.Add(a_Token.AsString(DBServerType, FSchemaList));
    While Assigned(a_Token) do
    begin
      a_Token := TSQLJoinToken(FSQLTokenList.NextToken(ttJoin));
      if Assigned(a_Token) then
        aJoin.Add(a_Token.AsString(DBServerType, FSchemaList));
    end;
  end;
end;

procedure TSQLBuilder.BuildSelectHeader(aHeader: TStrings);
var
  a_SelectCount: string;
begin
  a_SelectCount := getSelectTopCountStr;
  if SelectDistinct then
    aHeader.Add('SELECT DISTINCT' + a_SelectCount)
  else
     aHeader.Add('SELECT' + a_SelectCount);

  if FColumns.Count > 0 then
  begin
    BuildColumns(FColumns, FColumns, aHeader);
  end else
    aHeader.Add('*');
end;

procedure TSQLBuilder.BuildSelectQry(aSelect: TStrings);
var
  a_Token: TSQLJoinToken;
begin
  a_Token := TSQLJoinToken(FSQLTokenList.FirstToken(ttSelect));
  if Assigned(a_Token) then
  begin
    aSelect.Add(a_Token.AsString(DBServerType, FSchemaList));
    While Assigned(a_Token) do
    begin
      a_Token := TSQLJoinToken(FSQLTokenList.NextToken(ttSelect));
      if Assigned(a_Token) then
        aSelect.Add(a_Token.AsString(DBServerType, FSchemaList));
    end;
  end;
end;

procedure TSQLBuilder.BuildSQL(aType: TSQLBuilderType);
var
  a_Sql: TStrings;
begin
  a_Sql := TStringList.Create;
  try
    BuildBefore(a_Sql);
    if FSQLType = sqSelect then
    begin
      BuildSelectHeader(a_Sql);
      BuildJoin(a_Sql);
      BuildWhere(a_Sql);
    end else
    if FSQLType = sqInsert then
    begin
      BuildInsertHeader(a_Sql);
      BuildValues(a_Sql);
      BuildWhere(a_Sql);
    end else
    if FSQLType = sqInsertSelect then
    begin
      BuildInsertHeader(a_Sql);
      BuildSelectQry(a_Sql);
    end else
    if FSQLType = sqUpdate then
    begin
      BuildUpdateHeader(a_Sql);
      BuildJoin(a_Sql);
      BuildValues(a_Sql);
      BuildWhere(a_Sql);
    end else
    if FSQLType = sqDelete then
    begin
      BuildDeleteHeader(a_Sql);
      BuildJoin(a_Sql);
      BuildWhere(a_Sql);
    end;
  FSQL.AddStrings(a_SQL);
  FSQLString.AddStrings(a_SQL);
  finally
    a_SQL.Free;
  end;
end;

procedure TSQLBuilder.BuildUpdateHeader(aHeader: TStrings);
begin
  aHeader.Add('UPDATE');
end;

procedure TSQLBuilder.BuildValues(aValues: TStrings);
var
  a_Index: integer;
  a_Column, a_Value: string;
begin
  if SQLType = sqInsert then
  begin
    aValues.Add('VALUES');
    BuildColumns(FColumns, FActualValues, aValues, True, True);
  end else
  if SQLType = sqUpdate then
  begin
    for a_Index := 0 to FActualValues.Count - 1 do
    begin
      a_Value := FActualValues[a_Index];
      if TFieldType(FActualValues.Objects[a_Index]) in cQuotedFldType then
        a_Value := QuotedStr(a_Value);
      a_Column := FColumns[a_Index] + '=' + a_Value;
      if a_Index =  0 then
        aValues.Add('SET ' + a_Column)
      else
        aValues.Add(',' + a_Column);
    end;  
  end;
end;

procedure TSQLBuilder.BuildWhere(aWhere: TStrings);
var
  a_Token: TSQLWhereToken;
  a_OrderBy, a_ReadOnly, a_Fetch: string;
begin
  a_Token := TSQLWhereToken(FSQLTokenList.FirstToken(ttWhere));

  if Assigned(a_Token) then
  begin
    a_Token.AddWhereType(wtStart);
    aWhere.Add(a_Token.AsString(DBServerType, FSchemaList));

    While Assigned(a_Token) do
    begin
      a_Token := TSQLWhereToken(FSQLTokenList.NextToken(ttWhere));
      if Assigned(a_Token) then
      begin
//        a_Token.AddWhereType(WhereType);
        aWhere.Add(a_Token.AsString(DBServerType, FSchemaList));
      end;
    end;
  end;
  a_OrderBy := FOrderBy;
  if a_OrderBy <> '' then
    aWhere.Add('ORDER BY ' + a_OrderBy);
  a_Fetch := getFetchCountStr;
  if a_Fetch <> '' then
    aWhere.Add(a_Fetch);
  a_ReadOnly := getReadOnlyEndStr;
  if a_ReadOnly <> '' then
    aWhere.Add(a_ReadOnly);
end;

procedure TSQLBuilder.NeedSQLToken(aTableName: string; aBuild: boolean);
begin
{Your not required to call NeedSQLToken...use this if your building
dynamic SQL for search...you will only get the Token if aBuild is true for your Table}
  FSQLTokenList.BuildToken(DBServerType, FSchemaList, ttJoin, aTableName, aBuild);
end;

procedure TSQLBuilder.NeedWhereToken(aTableName: string; aBuild: boolean);
begin
{Your not required to call NeedWhereToken...use this if your building
dynamic SQL for search...you will only get the Token if aBuild is true for your Table}
  FSQLTokenList.BuildToken(DBServerType, FSchemaList, ttWhere, aTableName, aBuild);
end;

procedure Clear(aStrings: TStrings);
begin
  if Assigned(aStrings) then
    aStrings.Clear;
end;

procedure TSQLBuilder.ClearBuilder;
begin
  FSQL.Clear;

  Clear(FSQLString);
  Clear(FColumns);
//  FSchemaList.ClearSchemas;
  Clear(FActualValues);
  FSQLTokenList.ClearTokens;
  FOrderBy := '';
  FSelectTopCount := 0;
  FSelectDistinct := false;
end;

procedure TSQLBuilder.ClearSchemaInfo;
begin
  FSchemaList.ClearSchemas;
end;

constructor TSQLBuilder.Create;
begin
  FSQL := TWideStringList.Create;
  FSQLString := TStringList.Create;
  FSchemaList := TSchemaList.Create;
  FColumns:= TStringList.Create;


  FActualValues:= TStringList.Create;
  FSQLTokenList := TSQLTokenList.Create;
end;

destructor TSQLBuilder.Destroy;
begin
  FSQL.Free;
  FSQLString.Free;
  FSchemaList.Free;
  FColumns.Free;
  FActualValues.Free;
  FSQLTokenList.Free;
  inherited;
end;


procedure TSQLBuilder.EndGroup;
begin
  AddWhere(')', wtEnd);
end;

function TSQLBuilder.getDBServerType: TDBServerType;
begin
  Result := FDBServerType;
  TmtCheck.IsFalse(FDBServerType = dtUnknown, 'DBServerType must be set before building SQL', 'TSQLBuilder.getDBServerType');
end;

function TSQLBuilder.getFetchCountStr: string;
begin
  if (FDBServerType = dtDB2) and (SelectTopCount > 0) then
    Result := 'FETCH FIRST ' + IntToStr(SelectTopCount) + ' ROWS ONLY';
end;

function TSQLBuilder.getReadOnlyEndStr: string;
begin
 if (FDBServerType = dtDB2) and IsReadOnly then
    Result := 'FOR READ ONLY WITH UR';
end;

function TSQLBuilder.GetSchema(aColumn: string): string;
begin
  Result := FSchemaList.GetSchemaName(DBServerType, aColumn);
end;

function TSQLBuilder.getSelectTopCount: integer;
begin
  Result := FSelectTopCount;
end;

function TSQLBuilder.getSelectTopCountStr: string;
begin
  if (FDBServerType = dtSQLServer) and (SelectTopCount > 0) then
    Result := ' TOP ' + IntToStr(SelectTopCount)
  else
    Result := '';  
end;

function TSQLBuilder.getSQL: TWideStrings;
begin
  BuildSQL;
  Result := FSQL;
end;

function TSQLBuilder.GetSQLString: TStrings;
begin
  BuildSQL(btString);
  Result := FSQLString;
end;

function TSQLBuilder.GroupWhere(aClear: boolean): string;
begin

end;

procedure TSQLBuilder.setDBServerType(const Value: TDBServerType);
begin
  FDBServerType := Value;
  FSchemaList.ClearSchemas;
  if FDBServerType = dtSQLServer then
    FSchemaList.AddSchema(dtSQLServer, 'DBO', 'DEFAULT' )
  else
    FSchemaList.AddSchema(dtDB2, 'DB2HARMI', 'DEFAULT' )
end;



procedure TSQLBuilder.AddWhere(aSQL: TSQLBuilder; aWhereTypes: TWhereTypes;
  aBuild: boolean);
begin
  if aBuild then
    AddWhere('(' + aSQL.AsString + ')', aWhereTypes);
end;

{ TSchemaList }

procedure TSchemaList.AddSchema(aDBServerType: TDBServerType; aSchemaName,
  aTable: string);
var
  a_Schema: TSchema;
begin
  a_Schema := addSchemaInfo(aDBServerType, aSchemaName);
  a_Schema.AddTable(aTable);
end;

procedure TSchemaList.AddSchema(aDBServerType: TDBServerType;
  aSchemaName: string; const aTables: array of string);
var
  a_Schema: TSchema;
begin
  a_Schema := addSchemaInfo(aDBServerType, aSchemaName);
  a_Schema.AddTables(aTables);
end;


function TSchemaList.addSchemaInfo(aDBServerType: TDBServerType; aSchemaName: string): TSchema;
begin
  Result := findSchema(aDBServerType, aSchemaName);
  if Result = nil then
  begin
    Result := TSchema.Create;
    Result.DBServerType := aDBServerType;
    Result.Schema := aSchemaName;
    FList.Add(Result);
  end;
end;

procedure TSchemaList.ClearSchemas;
var
  a_Index: integer;
begin
  for a_Index := 0 to FList.Count - 1 do
    TSchema(FList[a_Index]).Free;
  FList.Clear;  
end;

constructor TSchemaList.Create;
begin
  inherited;
  FList := TList.Create;
end;

destructor TSchemaList.Destroy;
begin
  ClearSchemas;
  FList.Free;
  inherited;
end;

function TSchemaList.findSchema(aDBServerType: TDBServerType; aSchema: string): TSchema;
var
  a_Index: integer;
begin
  Result := nil;
  for a_Index := 0 to FList.Count - 1 do
    if (TSchema(FList[a_Index]).DBServerType = aDBServerType) and (TSchema(FList[a_Index]).Schema = aSchema) then
    begin
      Result := TSchema(FList[a_Index]);
      break;
    end;
end;

function TSchemaList.GetSchemaName(aDBServerType: TDBServerType; aTable: string): string;
var
  a_Index: integer;
begin
  Result := '';
  for a_Index := 0 to FList.Count - 1 do
    if (TSchema(FList[a_Index]).DBServerType = aDBServerType) and (TSchema(FList[a_Index]).HasTable(aTable)) then
    begin
      Result := TSchema(FList[a_Index]).Schema;
      break;
    end;
  if Result = '' then
    Result := GetSchemaName(aDBServerType, 'DEFAULT');
end;

{ TSchema }

procedure TSchema.AddTable(aTable: string);
begin
  FTables.Add(aTable);
end;

procedure TSchema.AddTables(const aTables: array of string);
var
  a_Index: integer;
begin
  for a_Index := low(aTables) to high(aTables) do
    AddTable(aTables[a_Index]);
end;

procedure TSchema.AddTables(aTables: TStrings);
var
  a_Index: integer;
begin
  for a_Index := 0 to aTables.Count -1 do
    AddTable(aTables[a_Index]);
end;

procedure TSchema.ClearSchema;
begin
  FSchema := '';
  FTables.Clear;
end;

constructor TSchema.Create;
begin
  inherited;
  FTables := TStringList.Create;
end;

destructor TSchema.Destroy;
begin
  FTables.Free;
  inherited;
end;

function TSchema.HasTable(aTable: string): boolean;
begin
  Result := FTables.IndexOf(aTable) <> -1;
end;

{ TSQLToken }

function TSQLToken.AsString(aDBServerType: TDBServerType;
  aSchemaList: TSchemaList): string;
begin
  Result := '';
end;


procedure TSQLToken.ClearToken;
begin
  FTableName := '';
  FTableAlias := '';
end;

{ TSQLTokenList }

procedure TSQLTokenList.AddSQLToken(aSQLToken: TSQLToken);
begin
  FList.Add(aSQLToken);
end;

function TSQLTokenList.BuildToken(aDBServerType: TDBServerType; aSchemaList: TSchemaList;
   aTokenType: TSQLTokenType; aTableName: string; aBuild: boolean): string;
var
  a_SQL: TSQLToken;
begin
  if aBuild then
  begin
    a_SQL := findSQLToken(aTableName, aTokenType);
    if Assigned(a_SQL) then
    begin
      a_SQL.Build := True;
      Result :=  a_SQL.AsString(aDBServerType, aSchemaList);
    end;
  end;
end;

procedure TSQLTokenList.ClearTokens;
var
  a_Index: integer;
begin
  for a_Index := 0 to FList.Count - 1 do
    TSQLToken(FList[a_Index]).Free;
  FList.Clear;
end;

constructor TSQLTokenList.Create;
begin
  inherited;
  FList := TList.Create;
end;

destructor TSQLTokenList.Destroy;
begin
  ClearTokens;
  FList.Free;
  inherited;
end;

function TSQLTokenList.findSQLToken(aTableName: string;aTokenType: TSQLTokenType): TSQLToken;
var
  a_Index: Integer;
begin
  Result := nil;
  for a_Index := 0 to FList.Count - 1 do
    if (aTableName = TSQLToken(FList[a_Index]).TableName)  and
       (aTokenType = TSQLToken(FList[a_Index]).SQLTokenType) then
    begin
      Result :=  TSQLToken(FList[a_Index]);
      break;
    end;
end;

function TSQLTokenList.FirstToken(aTokenType: TSQLTokenType; aBuild: boolean): TSQLToken;
begin
  FIndex := 0;
  Result := GetToken(aTokenType, FIndex);
  if (Result = nil) and (FIndex <=  FList.Count -1) then
    Result := NextToken(aTokenType);
end;

function TSQLTokenList.GetToken(aTokenType: TSQLTokenType; aIndex: integer; aBuild: boolean): TSQLToken;
var
  a_Token: TSQLToken;
begin
  Result := nil;
  FIndex := aIndex;
  if aIndex <=  FList.Count -1 then
  begin
    a_Token := FList[aIndex];
    if a_Token.SQLTokenType = aTokenType then
      Result := a_Token;
  end;
end;

function TSQLTokenList.LastToken(aTokenType: TSQLTokenType; aBuild: boolean): TSQLToken;
begin
  FIndex := FList.Count -1;
  Result := GetToken(aTokenType, FIndex);
  if (Result = nil) and (FIndex <=  FList.Count -1) then
    Result := PrevToken(aTokenType);
end;

function TSQLTokenList.NextToken(aTokenType: TSQLTokenType; aBuild: boolean): TSQLToken;
begin
  inc(FIndex);
  Result := GetToken(aTokenType, FIndex);
  if (Result = nil) and (FIndex <=  FList.Count -1) then
    Result := NextToken(aTokenType);
end;

function TSQLTokenList.PrevToken(aTokenType: TSQLTokenType; aBuild: boolean): TSQLToken;
begin
  dec(FIndex);
  Result := GetToken(aTokenType, FIndex);
  if (Result = nil) and (FIndex <=  FList.Count -1) then
    Result := PrevToken(aTokenType);
end;

{ TSQLWhereToken }

procedure TSQLWhereToken.AddWhereType(aWhereType: TWhereType);
begin
//  FWhereTypes := FWhereTypes- [wtAnd, wtOr, wtExists, wtNotExists];
  include(FWhereTypes, aWhereType);
end;

procedure TSQLWhereToken.AddWhereTypes(aWhereTypes: TWhereTypes);
begin
  FWhereTypes := FWhereTypes + aWhereTypes;
end;

function TSQLWhereToken.AsString(aDBServerType: TDBServerType;
  aSchemaList: TSchemaList): string;
var
  a_Temp: string;
  a_WT: TWhereTypes;
begin
  Result := '';
  a_Temp := '';
  if (wtStart in WhereTypes) then
    a_Temp := 'WHERE '
  else
  if (wtAnd in WhereTypes) or (wtBegin in WhereTypes) then
    a_Temp := ' AND '
  else
  if wtOr in WhereTypes then
    a_Temp := ' OR ';

  if wtExists in WhereTypes then
    a_Temp := a_Temp + 'EXISTS'
  else
  if wtNotExists in WhereTypes then
    a_Temp := a_Temp + 'NOT EXISTS';
  a_WT :=  [wtBegin, wtToken, wtEnd] * WhereTypes;
  if (a_WT = [wtBegin]) or (a_WT = [wtEnd]) then
    Result := a_Temp + FWhere
  else
  begin
    Result := a_Temp + '(' + FWhere + ')';
  end;
end;

constructor TSQLWhereToken.Create;
begin
  inherited Create;
  FSQLTokenType := ttWhere;
end;

function TSQLWhereToken.getQuoteField: boolean;
begin
  Result := FFieldType in cQuotedFldType;
end;

{ TSQLJoinToken }

function TSQLJoinToken.AsString(aDBServerType: TDBServerType;
  aSchemaList: TSchemaList): string;
var
  a_Schema, a_Lock: string;
begin
  Result := '';
  a_Schema := '';
  a_Lock := '';
  if JoinType = jtFrom then
    Result := 'FROM '
  else
  if JoinType = jtInner then
    Result := 'JOIN '
  else
  if JoinType = jtLeftOuter then
    Result := 'LEFT OUTER JOIN '
  else
    Result := 'RIGHT OUTER JOIN ';
  if UseSchema then
    a_Schema := aSchemaList.GetSchemaName(aDBServerType, TableName);
  if a_Schema <> '' then
  begin
    a_Schema := a_Schema + '.';
    a_Lock := getTableLockStr;
  end;

  if JoinType = jtUpdate then
    Result := a_Schema + TableName
  else
    Result := Result + a_Schema + TableName + ' ' + TableAlias + a_Lock;
  Result := Trim(Result);
  if JoinType in [jtInner, jtLeftOuter, jtRightOuter] then
    Result := Result + ' ON ' +  OnJoin;
end;

constructor TSQLJoinToken.Create;
begin
  inherited Create;
  FSQLTokenType := ttJoin;
end;


function TSQLJoinToken.getTableLockStr: string;
begin
  if (DBServerType = dtSQLServer) and (SQLType in [sqSelect] ) then
    Result := ' WITH (NOLOCK)'
  else
    Result := '';  
end;

{ TSQLSelectToken }

function TSQLSelectToken.AsString(aDBServerType: TDBServerType;
  aSchemaList: TSchemaList): string;
begin
  Result := Select;
end;

constructor TSQLSelectToken.Create;
begin
  inherited Create;
  FSQLTokenType := ttSelect;
end;

{ TSQLBeforeToken }

function TSQLBeforeToken.AsString(aDBServerType: TDBServerType;
  aSchemaList: TSchemaList): string;
begin
  Result := Line;
end;

constructor TSQLBeforeToken.Create;
begin
  inherited Create;
  FSQLTokenType := ttLineBefore;
end;

end.
