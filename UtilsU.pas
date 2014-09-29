unit UtilsU;
{written by Rick Peterson}
interface


uses
{$IFDEF DEBUGMM4}
  FastMM4,
{$ENDIF}
Classes, SysUtils, Windows, ActiveX, ShlObj, ShellAPI, DB, mtCheck, ImgList,
  WideStrings, Controls, Forms, Menus, wwdblook, StdCtrls, AdoDB, Dialogs,dbCtrls,
  CommCtrl, LoggerU, Variants, ComCtrls;

type
  TQueryType = (qtUpdate, qtInsert, qtDelete, qtSelect, qtSelectEdit);

  TStringEvent = procedure(var aString: string) of object;
  TQueryEvent = procedure(aQuery: TDataset) of object;
  TQueryEvent2 = procedure(aQuery: TDataset; aQueryType: TQueryType) of object;
  TBuildQueryEvent = procedure(aQuery: TDataset; aSQL: TStrings) of object;
  TBuildSQl = procedure(aSQL: TStrings);
  TControlEvent = procedure(aControl: TControl) of object;
  THandleEvent = procedure(aControl: TControl; aControlEvent: TControlEvent; var aHandled: boolean) of object;
  TBoolType = (btDefault, btTF, btPos, btNeg);

  TArrayVariant = Array of Variant;
  TArrayField = Array of TField;
  TArrayString = Array of string;
  TAIntegers = Array of Integer;  

  TDataType = (dtBoolean, dtAlpha, dtAlphaNumeric, dtNumeric, dtDate, dtMulti,
               dtMix, dtFix, dtFixedMulti, dtRuleOverall, dtParent, dtClass, dtDoc,
               dtAlphaNo, dtAlphaNumericNo, dtNumericNo, dtDateNo, dtMixNo, dtFixNo,
               dtFixedMultiNo);

function ExtractFileNameWOExt(const FileName: string): string;
function GetKnownDir(aCSIDL: integer):string;
function GetTempDir: string;
function GetApplicationDir: string;



function SHGetFolderLocation (hwndOwnder: THandle; nFolder: Integer; hToken: THandle; dwReserved: DWORD; ppidl:
  PItemIdList): HRESULT; stdcall; external 'shell32.dll' name 'SHGetFolderLocation';

function NPI_Checker(aNPI: string; aRaise: boolean = false): boolean;
function LuhnAlgorithm(aValue: string; aPrefix: string; aRaise: boolean = false): boolean;

function MemoryUsed: cardinal;

function CheckForEmpty(aDataSource: TDataSource): string;
{CheckValues: aValue are strings to check if empty, If empty return comma delimeted aConst that has the
same index.  Use for knowing if missing required data, the return string if empty
will let you  know all required info is there, if not empty display in
statusbar. i.e.  LastName, FirstName, Address1}
function CheckValues(const aValue, aConst: Array of string): string;

{CheckFieldValues:  Returns String value of TField that has null value, empty string,
or 0.  Similar to CheckValues...but uses TFields instead of Array of strings, note
is safer because doesn't have 2 arrays that must be in sync.}
function CheckFieldValues(const aFields: Array of TField): string;

function CheckValue(aControl: TControl; const aClasses: Array of TControlClass): boolean;

procedure AddValue(aSQL: TWidestrings; const aConst, aValue: string;
       aQuote: boolean = True); overload;

procedure AddValue(aSQL: TWidestrings; const aConst: string;
       aCheckValue: boolean; aValue: string;  aQuote: boolean = True); overload;

procedure AddValue(var aSQL: string; const aConst, aValue: string;
       aQuote: boolean = True); overload;

procedure AddValue(var aSQL: string; const aConst: string;
       aCheckValue: boolean; aValue: string;  aQuote: boolean = True); overload;


procedure EnableComponents(aWinControl: TWinControl; aEnable: boolean);
procedure EnableMenu(aMenu: TMainMenu; aEnable: boolean);overload;
procedure EnableMenu(aMenu: TPopupMenu; aEnable: boolean);overload;

procedure CallNotifyEvent(aSender: TObject; aEvent: TNotifyEvent);

function LoadQry(aQuery: TAdoQuery; aBuildQryEvent: TQueryEvent; aRun: boolean = True): boolean; overload;
function LoadQry(aQuery: TAdoQuery; aQueryType: TQueryType; aBuildQryEvent: TQueryEvent2; aRun: boolean = True): boolean; overload;

procedure WalkEdits(aControl: TWinControl; const aClasses: Array of TControlClass;
    aControlEvent: TControlEvent; aHandleEvent: THandleEvent;
    var aHandled: boolean; aBreak: boolean);

function FindForm(aControl: TControl): TForm; overload;
function FindForm(aComponent: TComponent): TForm; overload;
function FindForm(aName: string): TForm; overload;

function ValidateChar(var aKey: char;aValueType, aDataValue: string): boolean; overload;
function ValidateChar(var aKey: char;aDataType: TDataType; aDataValue: string): boolean; overload;
procedure SetBooleanType(aBoolType: TBoolType);
function BooleanToStr(aValue: boolean; aBooleanType: TBoolType = btDefault): string;
function StrToBoolean(aValue: string): boolean;



procedure  ConvertToHighColor(aImageList: TImageList);

function IsUpCase(aChar: char): boolean;
function OccurrencesOfChar(const aString: string; const aChar: char): integer;
function CapitalizeFirstLetter(Const aString: String; aStripNonAlpha: boolean; Const aIgnore: Array of String): String;
function AorAn(Const aString: string): string;
procedure ConvertStringToArray(Const aString: string; Const aDelimeter : char; var aArray: Array of String);

procedure ModifiedFields(aDataset: TDataset; var aValues: TArrayField; aOriginal: TArrayVariant);
procedure ModifiedFieldsWthIdFld(aDataset: TDataset; var aValues: TArrayField; aOriginal: TArrayVariant);
procedure StoreValues(aDataset: TDataset; var aValues: TArrayVariant);
procedure UpdateDataset(aWork: TAdoQuery;const  aFields: TArrayField; aTableName, aWhere: string; aLogger: ILogger = nil); overload;
procedure UpdateDataset(aWork: TAdoQuery;const  aFields: Array of TField; aTableName, aWhere: string; aLogger: ILogger = nil); overload;
procedure UpdateDataset(aWork: TAdoQuery;const  aFields: TArrayField; aIdField: TField; aTableName, aWhere: string; aLogger: ILogger = nil); overload;
procedure UpdateDataset(aWork: TAdoQuery;const  aFields: Array of TField; aIdField: TField; aTableName, aWhere: string; aLogger: ILogger = nil); overload;


procedure UpdateDataset(aWork: TAdoQuery;const  aDataset: TDataset; aTableName, aIdendityField, aWhere: string; aLogger: ILogger = nil); overload;

function InsertDataset(aWork: TAdoQuery;const  aFields: Array of TField; aTableName, aWhere: string; aLogger: ILogger = nil): integer; overload;
function InsertDataset(aWork: TAdoQuery;const  aDataset: TDataset; aTableName, aIdendityField, aWhere: string; aLogger: ILogger = nil): integer; overload;


procedure DeleteDataset(aWork: TAdoQuery;aTableName, aWhere: string; aLogger: ILogger = nil);
procedure LoadSQL(aQuery: TAdoQuery; aSQLDir: string;aOwner: TComponent = nil); overload;
procedure LoadSQL(aQuery: TAdoQuery; aFileName, aSQLDir: string); overload;

procedure StringToArray(var aArray: TArrayString; const aString: string; aDelimeter: char); overload;
procedure StringToArray(var aArray: TArrayString; const aString: TStrings); overload;
procedure StringToArray(aArray: TStrings; const aString: string; aDelimeter: char); overload;
procedure ArrayToString(const aArray: TArrayString; var aString: string); overload;
procedure ArrayToString(const aArray: Array of String; var aString: string); overload;
procedure ArrayToString(const aArray: Array of String; aString: TStrings); overload;
procedure ArrayToString(const aArray: TStrings; var aString: string); overload;

function CombineArray(aArray: array of Integer; aArray2: TAIntegers; aUnique: boolean = True): TAIntegers; overload;
function CombineArray(aArray, aArray2: TAIntegers; aUnique: boolean = True): TAIntegers; overload;
function ConvertArray(aArray: array of Integer): TAIntegers;

function GetYN(aCheckBox: TCheckBox): char; overload;
procedure SetYN(aCheckBox: TCheckBox; aValue: char); overload;

function GetChecked(aCheckBox: TCheckBox): char; overload;
procedure SetChecked(aCheckBox:TCheckBox; aValue: char);

function GetChecked(aField: TField): string; overload;
function GetYN(aField: TField; aUseQuote: boolean): string; overload;
function GetIntStr(aField: TField): string;

procedure SetComboBox(aComboBox: TComboBox; aValue: string); overload;
procedure SetComboBox(aComboBox: TComboBox; aID: integer); overload;

function GetValue(aComponent: TComponent): string; overload;
function GetValueAsInteger(aComponent: TComponent): integer;
procedure SetComponent(aComponent: TComponent; aValue: string; aUpperCase: boolean = False); overload;
procedure SetComponent(aComponent: TComponent; aValue: TDateTime); overload;

function GetValue(aField: TField; aNull: boolean = False; aQuote: boolean = False): string; overload;
function GetKeyValue(aField: TField; aNull: boolean): string;

function StrToChar(const S: string; Default: Char = ' '): Char;


const
  crlf = #10#13;
  cNumberSet = ['0'..'9'];
  cSpecialChars = [#1..#31];//backspace, LineFeed, Carriage Return
  cAlpha = ['a'..'z', 'A'..'Z', ' ', '!', '@', '#', '%', '^', '&', '*', '(' ,
   ')', '_', '-', '+',  '=', '{', '}', '[', ']', ':', ';', ',', '.', '<', '>', '?', '/'];
  cExtra = ['/', '-'];
  cVowel = ['a','e', 'i', 'o', 'u', 'A', 'E', 'I', 'O', 'U'];
{
  TFieldType = (ftUnknown, ftString, ftSmallint, ftInteger, ftWord, // 0..4
    ftBoolean, ftFloat, ftCurrency, ftBCD, ftDate, ftTime, ftDateTime, // 5..11
    ftBytes, ftVarBytes, ftAutoInc, ftBlob, ftMemo, ftGraphic, ftFmtMemo, // 12..18
    ftParadoxOle, ftDBaseOle, ftTypedBinary, ftCursor, ftFixedChar, ftWideString, // 19..24
    ftLargeint, ftADT, ftArray, ftReference, ftDataSet, ftOraBlob, ftOraClob, // 25..31
    ftVariant, ftInterface, ftIDispatch, ftGuid, ftTimeStamp, ftFMTBcd, // 32..37
    ftFixedWideChar, ftWideMemo, ftOraTimeStamp, ftOraInterval); // 38..41
}
  cNumberType = [ftSmallInt, ftInteger, ftWord, ftLargeint];
  cTrue = 'TRUE';
  cFalse = 'FALSE';


implementation

var
  uTrueFalse : Array of String;
  uTF: TBoolType = btPos;

type
  THackedControl = class(TControl);

const
  cASCII_VAL_OF_CHAR_0 = 48;

function MemoryUsed: cardinal;
var
    st: TMemoryManagerState;
    sb: TSmallBlockTypeState;
begin
    GetMemoryManagerState(st);
    result := st.TotalAllocatedMediumBlockSize + st.TotalAllocatedLargeBlockSize;
    for sb in st.SmallBlockTypeStates do begin
        result := result + sb.UseableBlockSize * sb.AllocatedBlockCount;
    end;
end;

procedure ArrayToString(const aArray: TArrayString; var aString: string);
var
  a_Index: integer;
begin
  aString := '';
  for a_Index := low(aArray) to high(aArray) do
    if aString = '' then
      aString := aArray[a_Index]
    else
      aString := aString + ' ' + aArray[a_Index];
end;

procedure ArrayToString(const aArray: Array of String; var aString: string);
var
  a_Index: integer;
begin
  aString := '';
  for a_Index := low(aArray) to high(aArray) do
    if aString = '' then
      aString := aArray[a_Index]
    else
      aString := aString + ' ' + aArray[a_Index];
end;

procedure ArrayToString(const aArray: Array of String; aString: TStrings); overload;
var
  a_Index: integer;
begin
  for a_Index := low(aArray) to high(aArray) do
    aString.Add(aArray[a_Index]);
end;

procedure ArrayToString(const aArray: TStrings; var aString: string); overload;
var
  a_Index: integer;
begin
  for a_Index := 0 to aArray.Count -1 do
    if aString = '' then
      aString := aArray[a_Index]
    else
      aString := aString + ' ' + aArray[a_Index];
end;


procedure StringToArray(var aArray: TArrayString; const aString: string; aDelimeter: char);
var
  a_Index: integer;
  a_Strings: TStrings;
begin
  a_Strings := TStringList.Create;
  try
    a_Strings.Delimiter := aDelimeter;
    a_Strings.DelimitedText := aString;
    SetLength(aArray, a_Strings.Count);
    for a_Index := 0 to a_Strings.Count - 1 do
      aArray[a_Index] := a_Strings.Strings[a_Index];
  finally
    a_Strings.Free;
  end;
end;

procedure StringToArray(var aArray: TArrayString; const aString: TStrings);
var
  a_Index: integer;
begin
  SetLength(aArray, aString.Count);
  for a_Index := 0 to aString.Count - 1 do
    aArray[a_Index] := aString.Strings[a_Index];
end;

procedure StringToArray(aArray: TStrings; const aString: string; aDelimeter: char); overload;
var
  a_Strings: TStrings;
begin
  a_Strings := TStringList.Create;
  try
    a_Strings.Delimiter := aDelimeter;
    a_Strings.DelimitedText := aString;
    aArray.Assign(a_Strings);
  finally
    a_Strings.Free;
  end;
end;

function CombineArray(aArray: array of Integer; aArray2: TAIntegers; aUnique: boolean): TAIntegers;
begin
  Result := ConvertArray(aArray);
  Result := CombineArray(Result, aArray2, aUnique);
end;

function CombineArray(aArray, aArray2: TAIntegers; aUnique: boolean): TAIntegers;
var
  a_Index, a_Index2, a_AIndex: integer;
  a_found: boolean;
begin
  a_AIndex := 0;
  SetLength(Result, Length(aArray)+Length(aArray2));
  for a_Index := low(aArray) to high(aArray) do
  begin
    Result[a_AIndex] := aArray[a_Index];
    inc(a_AIndex);
  end;

  for a_Index2 := low(aArray2) to high(aArray2) do
  begin
    a_found := False;
    Result[a_AIndex] := aArray2[a_Index2];{If found we will overwrite previous value}
    if aUnique then  {If Array must be unique...look for existing value in other Array}
    begin
      for a_Index := low(aArray) to high(aArray) do
      begin
        if aArray[a_Index] = aArray2[a_Index2] then
        begin
          a_found := True;
          break;
        end;
      end;
    end;
    if (not a_found) then
      inc(a_AIndex);
  end;
  SetLength(Result, a_AIndex);  
end;

function ConvertArray(aArray: array of Integer): TAIntegers;
var
  a_Index: integer;
begin
  SetLength(Result, Length(aArray));
  for a_Index := Low(aArray) to High(aArray) do
    Result[a_Index] := aArray[a_Index];
end;

function _TF(aBoolean: Boolean): string;
begin
  if uTF = btTF then
  begin
    if aBoolean then
      Result := cTrue
    else
      Result := cFalse;
  end else
  if uTF = btPos then
  begin
    if aBoolean then
      Result := '1'
    else
      Result := '0';
  end else
  begin
    if aBoolean then
      Result := '-1'
    else
      Result := '0';
  end;
end;

procedure SetBooleanType(aBoolType: TBoolType);
begin
  if (aBoolType <> uTF) then
  begin
    if aBoolType = btDefault then
      uTF := btTF
    else
      uTF := aBoolType;
    SetLength(uTrueFalse, 0);
  end;
end;

procedure _VerifyBoolStrArray;
begin
  if (Length(uTrueFalse) = 0) then
  begin
    SetLength(uTrueFalse, 2);
    uTrueFalse[0] := _TF(False);
    uTrueFalse[1] := _TF(True);
  end;
end;

function BooleanToStr(aValue: boolean; aBooleanType: TBoolType): string;
var
  a_BT: TBoolType;
begin
  a_BT := uTF;
  try
    if aBooleanType <> btDefault then
      SetBooleanType(aBooleanType);
    _VerifyBoolStrArray;
    Result := uTrueFalse[Integer(aValue)];
  finally
    if aBooleanType <> btDefault then
      SetBooleanType(a_BT);
  end;
end;

function StrToBoolean(aValue: string): boolean;
begin
  Result := (aValue = cTrue) or (aValue = '1') or (aValue = '-1');
end;

function DoubleValueAndSumDigits(aValue: char; aDoubleSum: boolean): integer;
begin
{Note the algorithim to Double works like this:  You double the value and add the
digits together, example 6 is doubled to 12.  You add  1 + 2 to get 3.
 1->0+2=2;2->0+4=4;3->0+6=6;4->0+8=8;5->1+0=1;6->1+2=3;7->1+4=5;8->1+6=7;9->1+8=9
 0-4 = double...5-9 = (((n-4)x2)-1) example (((8-4)x2)-1)= ((4 x 2) -1)= 7
 0-4 will be even(0-8) and 5-9 will be odd(1-9)
 }
 //Note the function expects aValue to be a single digit.
 Result := -1;
 if aValue <> '' then
 begin
   Result := Integer(aValue) - cASCII_VAL_OF_CHAR_0;
   if aDoubleSum then
   begin
     if (Result >= 0) and (Result <= 4) then
       Result := Result shl 1//Note shl 1 is the same as multiplying by 2
     else
       Result := (((Result - 4) shl 1) -1);
   end;
 end;
end;

function LuhnAlgorithm(aValue: string; aPrefix: string; aRaise: boolean): boolean;
var
  a_Sum: Int64;
  a_CharIndex, a_Index: integer;
  a_Char: Char;
{
    Counting from the check digit, which is the rightmost, and moving left, double the value of every second digit.
    Sum the digits of the products (e.g., 10: 1 + 0 = 1, 14: 1 + 4 = 5) together with the undoubled digits from the original number.
    If the total modulo 10 is equal to 0 (if the total ends in zero) then the number is valid according to the Luhn formula; else it is not valid.
    80840 = 24
    8+0(0 doubled)+8+8(4 doubled)+0
}
//The function expects only numeric values...trim should be called before calling this
begin
  Result := False;
  a_Index := 0;
  a_Sum := 0;
  aValue := aPrefix + aValue;
  for a_CharIndex := Length(aValue) downto 1 do
  begin
    inc(a_Index);
    a_Char := aValue[a_CharIndex];
    if not (a_Char in cNumberSet)  then
    begin
      if aRaise then
        Raise Exception.Create('Invalid character passed to LuhnAlgorithm, only valid chars are 0 to 9');
      exit;
    end;
    a_Sum := a_Sum + DoubleValueAndSumDigits(a_Char, a_Index mod 2 = 0);  //we double every other character
  end;
  Result := a_Sum mod 10 = 0;
end;

function LuhnAlgorithmCheckDigit(aValue: string; aPrefix: string; aRaise: boolean): integer;
var
  a_Sum: Int64;
  a_CharIndex, a_Index: integer;
  a_Char: Char;
{   Don't pass the check digit...this will return the CheckDigit...
    Counting before the check digit, check digit is the rightmost, and moving left,
    double the value of every second digit starting with the first.
    Sum the digits of the products (e.g., 10: 1 + 0 = 1, 14: 1 + 4 = 5) together with the undoubled digits from the original number.
    If the total modulo 10 is equal to 0 (if the total ends in zero) then the number is valid according to the Luhn formula; else it is not valid.
    80840 = 24
    8+0(0 doubled)+8+8(4 doubled)+0
}
//The function expects only numeric values...trim should be called before calling this
begin
  Result := 0;
  a_Index := 1;
  a_Sum := 0;
  aValue := aPrefix + aValue;
  for a_CharIndex := Length(aValue) downto 1 do
  begin
    inc(a_Index);
    a_Char := aValue[a_CharIndex];
    if not (a_Char in cNumberSet)  then
    begin
      if aRaise then
        Raise Exception.Create('Invalid character passed to LuhnAlgorithm, only valid chars are 0 to 9');
      exit;
    end;
    a_Sum := a_Sum + DoubleValueAndSumDigits(a_Char, a_Index mod 2 = 0);  //we double every other character
  end;
  Result := 10 - (a_Sum mod 10);  //this is the check digit
//  Result := a_Sum mod 10 = 0;
end;


function NPI_Checker(aNPI: string; aRaise: boolean): boolean;
begin
  if Length(aNPI) = 10 then
    Result := LuhnAlgorithm(aNPI, '80840')//constant of 24
  else
    Result := False;
  if aRaise and not Result then
    Raise Exception.Create('Invalid NPI value:' + aNPI);
end;

function ExtractFileNameWOExt(const FileName: string): string;
var
  a_Index: Integer;
begin
  Result := ExtractFileName(FileName);
  a_Index := LastDelimiter('.', Result);
  if (a_Index > 0) and (Result[a_Index] = '.') then
    Result := Copy(Result, 1, a_Index -1);
end;

function GetTempDir: string;
var
  a_temp: array[0..MAX_PATH] of Char;
begin
  GetTempPath(MAX_PATH, @a_temp);
  Result := StrPas(a_temp);
  Result := IncludeTrailingPathDelimiter(Result);
end;

function GetApplicationDir: string;
begin
  Result := IncludeTrailingPathDelimiter(ExtractFileDir(Application.ExeName));
end;

function GetKnownDir(aCSIDL: integer): string;
const
  NTFS_MAX_PATH = 32767;
var
  a_PIDL: PItemIDList; {uses: ShlObj}
  a_Path: PChar;
begin
  Result := '';
  a_PIDL := nil;
  if Succeeded(SHGetFolderLocation(0, aCSIDL, 0, 0, @a_PIDL)) then
  begin
    if a_PIDL <> nil then
    begin
      try
      GetMem(a_Path, (NTFS_MAX_PATH + 1) * 2);
      try
        if SHGetPathFromIDList(a_PIDL, PChar(a_Path)) then
        begin
          Result := a_Path;
          Result := IncludeTrailingPathDelimiter(Result);
        end;
      finally
        FreeMem(a_Path);
      end;
      finally
        CoTaskMemFree(a_PIDL);
      end;
    end;
  end;
end;

function HandleCheckValue(aClass: TClass; const aClasses: Array of TControlClass):integer;
var
  a_Index: integer;
begin
  Result := -1;
  for a_Index := Ord(Low(aClasses)) to Ord(High(aClasses))do
    if (Assigned(aClasses[a_Index])) then
    begin
      if (aClass = aClasses[a_Index]) then
      begin
        Result := a_Index;
        break;
      end;
    end else
      Break;
end;

function CheckValue(aControl: TControl; const aClasses: Array of TControlClass): boolean;
begin
  if HandleCheckValue(aControl.ClassType, aClasses) >= 0 then
  begin
    Result := THackedControl(aControl).Text <> '';
    if not Result then
    begin
{special handling for TDBLookupComboBox, It's Text property does not inherit from
TControl...TDBLookupComboBox uses property Text: string read FText;}
       if aControl is TDBLookupComboBox  then
        Result := TDBLookupComboBox(aControl).Text <> '';
    end;
  end
  else
    Result := False;
end;

function CheckValues(const aValue, aConst: Array of string): string;
var
  a_Index: integer;
  a_Const: String;
begin
  Result := '';
  TmtCheck.IsTrue(High(aValue)= High(aConst), 'CheckValues', 'Programmer Error, Value and Const should have the same amount of elements');
  for a_Index := Low(aValue) to High(aValue) do
  begin
    if (aValue[a_Index] = '') then
    begin
      a_Const := aConst[a_Index];
      if Result = '' then
        Result := a_Const
      else
        Result := Result + ', ' +  a_Const;
    end;
  end;
end;

function CheckFieldValues(const aFields: Array of TField): string;
var
  a_Index: integer;
  a_Const: String;
  a_Field: TField;
begin
  Result := '';
  for a_Index := Low(aFields) to High(aFields) do
  begin
    a_Field := aFields[a_Index];
    if Assigned(a_Field) and (a_Field.IsNull or (a_Field.AsString = '' )
    or ((a_Field.DataType in cNumberType) and (a_Field.AsInteger = 0))) then
    begin
      a_Const := a_Field.FieldName;
      if Result = '' then
        Result := a_Const
      else
        Result := Result + ', ' +  a_Const;
    end else
    begin
      if not Assigned(a_Field) then      
        Result := 'Not Found Index:' + IntToStr(a_Index);
    end;
  end;
end;


function CheckForEmpty(aDataSource: TDataSource): string;
var
  a_DS: TDataSet;
  a_Index, a_FldIndex: integer;
  a_Field: TField;
begin
{To use this you have to at design time add the Fields for the Dataset, then
set the Required Property to True for all fields that can not be nullable
...can't get this to work in DB2...not pulling the info for the tables to
know if not nullable 
}
  Result := '';
  TmtCheck.IsAssigned(aDataSource,'CheckForEmpty', 'Invalid Datasource passed to CheckForEmpty');
  a_DS := aDataSource.DataSet;
  TmtCheck.IsAssigned(a_DS,'CheckForEmpty', 'No DataSet property set for the passed DataSource');
  TmtCheck.IsTrue(aDataSource.Dataset.Active,'CheckForEmpty', 'DataSet property is not to Active for the passed DataSource');
  a_DS.FieldDefs.Updated := False;
  a_DS.FieldDefs.Update;
  for a_Index := 0 to a_DS.FieldDefs.Count -1 do
  begin//faRequired means that the value can not be null
    if a_DS.FieldDefs.Items[a_Index].Required then
    begin
      for a_FldIndex := 0 to a_DS.FieldCount - 1 do
      begin
        if a_DS.FieldDefs.Items[a_Index].FieldNo = a_DS.Fields[a_FldIndex].FieldNo then
        begin
          a_Field := a_DS.Fields[a_FldIndex];
          if Assigned(a_Field) and ((a_Field.IsNull) or (a_Field.AsString = '')) then
            if Result = '' then
              Result := a_Field.FieldName
            else
              Result := Result + ',' + a_Field.Fieldname;
          end;
      end;
    end;
  end;
end;

procedure AddValue(aSQL: TWidestrings; const aConst, aValue: string;
  aQuote: boolean);(*defaults aQuote = True
if aConst has like or like% it will change aValue,
             for like it will be aValue + '%'
             for like% it will be '%' + aValue + '%'
             for both of these likes the aValue will be automatically quoted *)

{This is used to add SQL to SQL property where you need to check the value first
before adding it

example1 of Old Way:         if TRIM(editLastName.Text) <> '' THEN
          Add ('AND pat.LastName like ' + QuotedStr(Trim(editLastName.Text) + '%'));
example1 New Way:  AddValue(qryWork.SQL, 'AND pat.LastName like', editLastName.Text);

If TRIM(editLastName.Text) <> '' then Both of them should create SQL Code that
looks the same(will assume editLastName.Text = Smith)
AND pat.LastName like 'Smith%'
********************************************************************************
example1.1 of Old Way: if TRIM(editHistoryNotes.Text) <> '' then
    Add ('AND a.pkAuditID in (select fkAuditID from DTLOutreach..AuditHistory with(nolock) where Notes like ' + QuotedStr('%' + TRIM(editHistoryNotes.Text) + '%') + ')');

example1.1 New Way:  AddValue(qryWork.SQL, 'AND a.pkAuditID in (select fkAuditID from DTLOutreach..AuditHistory with(nolock) where Notes like%', editHistoryNotes.Text);

If TRIM(editHistoryNotes.Text) <> '' then Both of them should create SQL Code that
looks the same(will assume editHistoryNotes.Text = Hello World)
AND a.pkAuditID in (select fkAuditID from DTLOutreach..AuditHistory with(nolock) where Notes like '%Hello World%'
********************************************************************************
example2 of Old Way:  if TRIM(dtpReceivedFrom.Text) <> '' then
  Add ('AND CAST(CONVERT(varchar, a.AuditReceivedDate, 101) as DateTime) >= ' + QuotedStr(dtpReceivedFrom.Text));
example2 New Way:  AddValue(qryWork.SQL, 'AND CAST(CONVERT(varchar, a.AuditReceivedDate, 101) as DateTime) >=', dtpReceivedFrom.Text);

If TRIM(dtpReceivedFrom.Text) <> '' then Both of them should create SQL Code that
looks the same(will assume dtpReceivedFrom.Text = 1/1/12)
AND CAST(CONVERT(varchar, a.AuditReceivedDate, 101) as DateTime) >= '1/1/12'
********************************************************************************
example3 of Old Way:  if TRIM(editRxNo.Text) <> '' THEN
          Add ('AND o.RxNo = ' + editRxNo.Text);
example3 New Way:  AddValue(qryWork.SQL, 'AND o.RxNo =', editRxNo.Text, False);

If TRIM(editRxNo.Text) <> '' then Both of them should create SQL Code that
looks the same(will assume editRxNo.Text = 123456789)
AND o.RxNo = 123456789
}
var
  a_Const, a_Value: string;
  a_Like: boolean;
begin
  if Assigned(aSQL) then
  begin
    a_Value := trim(aValue);
    if a_Value <> '' then
    begin
      TmtCheck.IsTrue(POS(';', a_Value) = 0, 'AddValue', 'Invalid Character: ;');
      a_Const := UpperCase(Trim(aConst))+ ' ';
      a_Like := True;//we set this to true...because if we don't find a like we set it to false
      if POS(' LIKE ', a_Const) > 0 then
        a_Value := a_Value + '%'
      else
      if POS(' LIKE% ', a_Const) > 0 then
      begin
        a_Const := StringReplace(a_Const, 'LIKE%', 'LIKE ',[rfIgnoreCase]); 
        a_Value := '%' + a_Value + '%';
      end
      else
        a_Like := False;
      if a_Like or aQuote then
        a_Value := QuotedStr(a_Value)
      else
        a_Value := IntToStr(StrToInt(a_Value));//test for sql injection
      aSQL.Add(a_Const + a_Value);
    end;
  end;
end;

procedure AddValue(aSQL: TWidestrings; const aConst: string;
  aCheckValue: boolean; aValue: string;  aQuote: boolean);
{
example4 of Old Way:  if cbAllowedOver0.Checked then
          Add ('AND o.AmtAllowed > 0');
example4 New Way:  AddValue(qryWork.SQL, 'AND o.AmtAllowed >',cbAllowedOver0.Checked, '0', False);

If cbAllowedOver0.Checked then Both of them should create SQL Code that
looks the same
AND o.AmtAllowed > 0
********************************************************************************
example5 of Old Way:  if cbMyWorkOnly.Checked then
          Add ('AND a.AssignedTo = ' + QuotedStr(UserName));
example5 New Way:  AddValue(qryWork.SQL, 'AND a.AssignedTo =',cbMyWorkOnly.Checked, UserName);

If cbAllowedOver0.Checked then Both of them should create SQL Code that
looks the same(will assume UserName = cpeterso)
AND a.AssignedTo = 'cpeterso'
********************************************************************************
example6 of Old Way:  if cbPaidAudits.Checked then begin //added 1/2/2011 per ticket 12384
          Add ('AND c.ActualAllowed < ((c.TotalExpected + c.Copay) * .97)');
          Add ('AND c.ActualAllowed > 0');
          Add ('AND a.Status <> ''Closed''');
          Add ('AND o.AmtPaidDate > a.RequestDate');
        end;
example6 New Way:
  AddValue(qryWork.SQL, 'AND c.ActualAllowed < ((c.TotalExpected + c.Copay) * .97)',cbPaidAudits.Checked, '');
  AddValue(qryWork.SQL, 'AND c.ActualAllowed > 0',cbPaidAudits.Checked, '');
  AddValue(qryWork.SQL, 'AND a.Status <>',cbPaidAudits.Checked, 'Closed');
  AddValue(qryWork.SQL, ''AND o.AmtPaidDate > a.RequestDate',cbPaidAudits.Checked, '');
If cbPaidAudits.Checked then Both of them should create SQL Code that
looks the same(
AND c.ActualAllowed < ((c.TotalExpected + c.Copay) * .97)
AND c.ActualAllowed > 0',cbPaidAudits.Checked
AND a.Status <> 'Closed'
AND o.AmtPaidDate > a.RequestDate

}
var
  a_Const, a_Value: string;
  a_Like: boolean;
begin
  if Assigned(aSQL) then
  begin
    if aCheckValue then
    begin
      a_Value := trim(aValue);
      if a_Value <> '' then
      begin
        TmtCheck.IsTrue(POS(';', a_Value) = 0, 'AddValue', 'Invalid Character: ;');
        TmtCheck.IsTrue(POS('''', a_Value) = 0, 'AddValue', 'Invalid Character: ''''');
        a_Const := UpperCase(Trim(aConst))+ ' ';
        a_Like := True;//we set this to true...because if we don't find a like we set it to false
        if POS(' LIKE ', a_Const) > 0 then
          a_Value := a_Value + '%'
        else
        if POS(' LIKE% ', a_Const) > 0 then
        begin
          a_Const := StringReplace(a_Const, 'LIKE%', 'LIKE ',[rfIgnoreCase]);
          a_Value := '%' + a_Value + '%';
        end
        else
          a_Like := False;
        if a_Like or aQuote then
          a_Value := QuotedStr(a_Value)
        else
          a_Value := IntToStr(StrToInt(a_Value));
        aSQL.Add(a_Const + a_Value);
      end;
    end;
  end;
end;

procedure AddValue(var aSQL: string; const aConst, aValue: string;
  aQuote: boolean);(*defaults aQuote = True
if aConst has like or like% it will change aValue,
             for like it will be aValue + '%'
             for like% it will be '%' + aValue + '%'
             for both of these likes the aValue will be automatically quoted *)

{This is used to add SQL to SQL property where you need to check the value first
before adding it

example1 of Old Way:         if TRIM(editLastName.Text) <> '' THEN
          Add ('AND pat.LastName like ' + QuotedStr(Trim(editLastName.Text) + '%'));
example1 New Way:  AddValue(qryWork.SQL, 'AND pat.LastName like', editLastName.Text);

If TRIM(editLastName.Text) <> '' then Both of them should create SQL Code that
looks the same(will assume editLastName.Text = Smith)
AND pat.LastName like 'Smith%'
********************************************************************************
example1.1 of Old Way: if TRIM(editHistoryNotes.Text) <> '' then
    Add ('AND a.pkAuditID in (select fkAuditID from DTLOutreach..AuditHistory with(nolock) where Notes like ' + QuotedStr('%' + TRIM(editHistoryNotes.Text) + '%') + ')');

example1.1 New Way:  AddValue(qryWork.SQL, 'AND a.pkAuditID in (select fkAuditID from DTLOutreach..AuditHistory with(nolock) where Notes like%', editHistoryNotes.Text);

If TRIM(editHistoryNotes.Text) <> '' then Both of them should create SQL Code that
looks the same(will assume editHistoryNotes.Text = Hello World)
AND a.pkAuditID in (select fkAuditID from DTLOutreach..AuditHistory with(nolock) where Notes like '%Hello World%'
********************************************************************************
example2 of Old Way:  if TRIM(dtpReceivedFrom.Text) <> '' then
  Add ('AND CAST(CONVERT(varchar, a.AuditReceivedDate, 101) as DateTime) >= ' + QuotedStr(dtpReceivedFrom.Text));
example2 New Way:  AddValue(qryWork.SQL, 'AND CAST(CONVERT(varchar, a.AuditReceivedDate, 101) as DateTime) >=', dtpReceivedFrom.Text);

If TRIM(dtpReceivedFrom.Text) <> '' then Both of them should create SQL Code that
looks the same(will assume dtpReceivedFrom.Text = 1/1/12)
AND CAST(CONVERT(varchar, a.AuditReceivedDate, 101) as DateTime) >= '1/1/12'
********************************************************************************
example3 of Old Way:  if TRIM(editRxNo.Text) <> '' THEN
          Add ('AND o.RxNo = ' + editRxNo.Text);
example3 New Way:  AddValue(qryWork.SQL, 'AND o.RxNo =', editRxNo.Text, False);

If TRIM(editRxNo.Text) <> '' then Both of them should create SQL Code that
looks the same(will assume editRxNo.Text = 123456789)
AND o.RxNo = 123456789
}
var
  a_Const, a_Value: string;
  a_Like: boolean;
begin
  aSQL := '';
  a_Value := trim(aValue);
  if a_Value <> '' then
  begin
    TmtCheck.IsTrue(POS(';', a_Value) = 0, 'AddValue', 'Invalid Character: ;');
    a_Const := UpperCase(Trim(aConst))+ ' ';
    a_Like := True;//we set this to true...because if we don't find a like we set it to false
    if POS(' LIKE ', a_Const) > 0 then
      a_Value := a_Value + '%'
    else
    if POS(' LIKE% ', a_Const) > 0 then
    begin
      a_Const := StringReplace(a_Const, 'LIKE%', 'LIKE ',[rfIgnoreCase]);
      a_Value := '%' + a_Value + '%';
    end
    else
      a_Like := False;
    if a_Like or aQuote then
      a_Value := QuotedStr(a_Value)
    else
      a_Value := IntToStr(StrToInt(a_Value));//test for sql injection
    aSQL := a_Const + a_Value;
  end;
end;

procedure AddValue(var aSQL: string; const aConst: string;
  aCheckValue: boolean; aValue: string;  aQuote: boolean);
{
example4 of Old Way:  if cbAllowedOver0.Checked then
          Add ('AND o.AmtAllowed > 0');
example4 New Way:  AddValue(qryWork.SQL, 'AND o.AmtAllowed >',cbAllowedOver0.Checked, '0', False);

If cbAllowedOver0.Checked then Both of them should create SQL Code that
looks the same
AND o.AmtAllowed > 0
********************************************************************************
example5 of Old Way:  if cbMyWorkOnly.Checked then
          Add ('AND a.AssignedTo = ' + QuotedStr(UserName));
example5 New Way:  AddValue(qryWork.SQL, 'AND a.AssignedTo =',cbMyWorkOnly.Checked, UserName);

If cbAllowedOver0.Checked then Both of them should create SQL Code that
looks the same(will assume UserName = cpeterso)
AND a.AssignedTo = 'cpeterso'
********************************************************************************
example6 of Old Way:  if cbPaidAudits.Checked then begin //added 1/2/2011 per ticket 12384
          Add ('AND c.ActualAllowed < ((c.TotalExpected + c.Copay) * .97)');
          Add ('AND c.ActualAllowed > 0');
          Add ('AND a.Status <> ''Closed''');
          Add ('AND o.AmtPaidDate > a.RequestDate');
        end;
example6 New Way:
  AddValue(qryWork.SQL, 'AND c.ActualAllowed < ((c.TotalExpected + c.Copay) * .97)',cbPaidAudits.Checked, '');
  AddValue(qryWork.SQL, 'AND c.ActualAllowed > 0',cbPaidAudits.Checked, '');
  AddValue(qryWork.SQL, 'AND a.Status <>',cbPaidAudits.Checked, 'Closed');
  AddValue(qryWork.SQL, ''AND o.AmtPaidDate > a.RequestDate',cbPaidAudits.Checked, '');
If cbPaidAudits.Checked then Both of them should create SQL Code that
looks the same(
AND c.ActualAllowed < ((c.TotalExpected + c.Copay) * .97)
AND c.ActualAllowed > 0',cbPaidAudits.Checked
AND a.Status <> 'Closed'
AND o.AmtPaidDate > a.RequestDate

}
var
  a_Const, a_Value: string;
  a_Like: boolean;
begin
  if aCheckValue then
  begin
    aSQL := '';
    a_Value := trim(aValue);
    if a_Value <> '' then
    begin
      TmtCheck.IsTrue(POS(';', a_Value) = 0, 'AddValue', 'Invalid Character: ;');
      TmtCheck.IsTrue(POS('''', a_Value) = 0, 'AddValue', 'Invalid Character: ''''');
      a_Const := UpperCase(Trim(aConst))+ ' ';
      a_Like := True;//we set this to true...because if we don't find a like we set it to false
      if POS(' LIKE ', a_Const) > 0 then
        a_Value := a_Value + '%'
      else
      if POS(' LIKE% ', a_Const) > 0 then
      begin
        a_Const := StringReplace(a_Const, 'LIKE%', 'LIKE ',[rfIgnoreCase]);      
        a_Value := '%' + a_Value + '%'
      end
      else
        a_Like := False;
      if a_Like or aQuote then
        a_Value := QuotedStr(a_Value)
      else
        a_Value := IntToStr(StrToInt(a_Value));
      aSQL := a_Const + a_Value;
    end;
  end;
end;

procedure EnableMenu(aMenu: TMainMenu; aEnable: boolean);
var
  a_Index: integer;
begin
  for a_Index := 0 to aMenu.Items.Count -1 do
  begin
  {we loop through all the children}
    aMenu.Items[a_Index].Enabled := aEnable;
  end;
end;

procedure EnableMenu(aMenu: TPopupMenu; aEnable: boolean);
var
  a_Index: integer;
begin
  for a_Index := 0 to aMenu.Items.Count -1 do
  begin
  {we loop through all the children}
    aMenu.Items[a_Index].Enabled := aEnable;
  end;
end;


procedure EnableComponents(aWinControl: TWinControl; aEnable: boolean);
var
  a_Index: integer;
begin
  aWinControl.Enabled := aEnable;
  for a_Index := 0 to aWinControl.ControlCount -1 do
  begin
  {we loop through all the children}
    aWinControl.Controls[a_Index].Enabled := aEnable;
    if aWinControl.Controls[a_Index] is TWinControl then
{if its a WinControl...meaning it can be a Parent to other controls...we loop
 through them, and so and so on.}
      EnableComponents(TWinControl(aWinControl.Controls[a_Index]), aEnable);
  end;
end;

procedure CallNotifyEvent(aSender: TObject; aEvent: TNotifyEvent);
var
  a_Cursor: TCursor;
begin
  if Assigned(aEvent) then
  begin
    a_Cursor := Screen.Cursor;
    Screen.Cursor := crHourGlass;
    try
      Application.ProcessMessages;
      aEvent(aSender);
    finally
      Screen.Cursor := a_Cursor;
    end;
  end;
end;

function LoadQry(aQuery: TAdoQuery; aBuildQryEvent: TQueryEvent; aRun: boolean): boolean;
begin
  TmtCheck.isAssigned(aQuery, 'Query is unassigned', 'LoadForm');
  if Assigned(aBuildQryEvent) then
    aBuildQryEvent(aQuery);
  aQuery.DisableControls;
  try
    if aRun then
    begin
      aQuery.Active := True;
      Result := aQuery.RecordCount > 0;
    end else
    begin
      Result := aQuery.ExecSQL > 0;
    end;
  finally
    aQuery.EnableControls;
  end;
end;

function LoadQry(aQuery: TAdoQuery; aQueryType: TQueryType; aBuildQryEvent: TQueryEvent2; aRun: boolean = True): boolean;
begin
  TmtCheck.isAssigned(aQuery, 'Query is unassigned', 'LoadForm');
  if Assigned(aBuildQryEvent) then
    aBuildQryEvent(aQuery, aQueryType);
  aQuery.DisableControls;
  try
    if aRun then
    begin
      aQuery.Active := True;
      Result := aQuery.RecordCount > 0;
    end else
    begin
      Result := aQuery.ExecSQL > 0;
    end;
  finally
    aQuery.EnableControls;
  end;
end;

function HandleControl(aControl: TControl; const aClasses: Array of TControlClass;
      aControlEvent: TControlEvent; aHandleEvent: THandleEvent): boolean;
begin
  Result := False;
  if CheckValue(aControl, aClasses) then
  begin
    if Assigned(aControlEvent) then
      aControlEvent(aControl);
    Result := True;
  end;

  if Assigned(aHandleEvent) then
    aHandleEvent(aControl, aControlEvent, Result);
end;

procedure WalkEdits(aControl: TWinControl; const aClasses: Array of TControlClass;
    aControlEvent: TControlEvent; aHandleEvent: THandleEvent;
    var aHandled: boolean; aBreak: boolean);
var
  a_Index: integer;
begin
//aBreak is usually used to break out of the Walk if Control has a value
  if not (aHandled and aBreak) then
    for a_Index  := 0 to aControl.ControlCount - 1 do
    begin
      if aControl.Controls[a_Index] is TWinControl then
      begin
        if not (aHandled and aBreak) then  //aHandled = F and aBreak = F
          aHandled := HandleControl(aControl.Controls[a_Index],aClasses,  aControlEvent, aHandleEvent);

        if aHandled and aBreak then //aHandled = T and aBreak = T
          exit
        else  //(aHandled = T and aBreak = F) or (aHandled = F and aBreak = T)
          WalkEdits(TWinControl(aControl.Controls[a_Index]),aClasses,
                    aControlEvent, aHandleEvent, aHandled, aBreak);
      end else
      begin
        aHandled := HandleControl(aControl.Controls[a_Index],aClasses,  aControlEvent, aHandleEvent);
        if aHandled and aBreak then
          exit;
      end;
    end;
end;

function FindForm(aControl: TControl): TForm;
begin
  Result := nil;
  if aControl is TForm then
    Result := TForm(aControl)
  else
  if Assigned(aControl.Parent) then
  begin
    if aControl.Parent is TForm then
      Result := TForm(aControl.Parent)
    else
      Result := FindForm(aControl.Parent);
  end;
end;

function FindForm(aComponent: TComponent): TForm;
begin
  Result := nil;
  if Assigned(aComponent.Owner) then
  begin
    if (aComponent.Owner is TControl) then
      Result := FindForm(TControl(aComponent.Owner))
    else
      Result := FindForm(aComponent.Owner);
  end;
end;

function FindForm(aName: string): TForm; overload;
begin
  Result := TForm(Application.FindComponent(aName));    
end;

function ValidateChar(var aKey: char; aValueType,
  aDataValue: string): boolean;
begin
  Result := True;
  if not (aKey in cSpecialChars) then
    if aValueType = 'NUMERIC' then begin
      //festra.com
      if not (aKey in cNumberSet + [DecimalSeparator]) then begin
        ShowMessage('Invalid Key: ' + aKey + ' (Expecting Numeric)');
        aKey := #0;
        Result := False;
      end else begin
        if (aKey = DecimalSeparator) and (Pos(aKey, aDataValue) > 0) then begin
          ShowMessage('Invalid Key: ' + aKey + ' (Decimal Already Exists)');
          aKey := #0;
          Result := False;
        end;
      end;
    end else
    if aValueType = 'ALPHA' then begin
      if not (aKey in cAlpha) then begin
        ShowMessage('Invalid Key: ' + aKey + ' (Expecting Alpha)');
        aKey := #0;
        Result := False;
      end;
    end else
    if (aValueType = 'ALPHANUMERIC') or (aValueType = 'MIX') then begin
      //festra.com
      if not (aKey in cAlpha + cNumberSet + cExtra + [DecimalSeparator]) then begin
        ShowMessage('Invalid Key: ' + aKey + ' (Expecting AlphaNumeric)');
        aKey := #0;
        Result := False;
      end;
    end;
end;

function ValidateChar(var aKey: char;aDataType: TDataType; aDataValue: string): boolean;
begin
  Result := True;
  if not (aKey in cSpecialChars) then
    if aDataType = dtNumeric then begin
      //festra.com
      if not (aKey in cNumberSet + [DecimalSeparator]) then begin
        ShowMessage('Invalid Key: ' + aKey + ' (Expecting Numeric)');
        aKey := #0;
        Result := False;
      end else begin
        if (aKey = DecimalSeparator) and (Pos(aKey, aDataValue) > 0) then begin
          ShowMessage('Invalid Key: ' + aKey + ' (Decimal Already Exists)');
          aKey := #0;
          Result := False;
        end;
      end;
    end else
    if aDataType = dtAlpha then begin
      if not (aKey in cAlpha) then begin
        ShowMessage('Invalid Key: ' + aKey + ' (Expecting Alpha)');
        aKey := #0;
        Result := False;
      end;
    end else
    if aDataType in [dtMix, dtAlphaNumeric] then begin
      //festra.com
      if not (aKey in cAlpha + cNumberSet + cExtra + [DecimalSeparator]) then begin
        ShowMessage('Invalid Key: ' + aKey + ' (Expecting AlphaNumeric)');
        aKey := #0;
        Result := False;
      end;
    end;
end;


procedure ConvertToHighColor(aImageList: TImageList);
// To show smooth images we have to convert the image list from 16 colors to high color.
(*taken from VirtualTree*)
var
  IL: TImageList;

begin
  // Have to create a temporary copy of the given list, because the list is cleared on handle creation.
  IL := TImageList.Create(nil);
  IL.Assign(aImageList);

  with aImageList do
    Handle := ImageList_Create(Width, Height, ILC_COLOR16 or ILC_MASK, Count, AllocBy);
  aImageList.Assign(IL);
  IL.Free;
end;

function IsUpCase(aChar: char): boolean;
begin
  Result := aChar in ['A'..'Z'];
end;

function AorAn(Const aString: string): string;
begin
  if Length(aString) > 0 then
  begin
    if aString[1] in cVowel then
      Result := 'an'
    else
      Result := 'a';
  end;
end;

function OccurrencesOfChar(const aString: string; const aChar: char): integer;
var
  i: Integer;
begin
  result := 0;
  for i := 1 to Length(aString) do
    if aString[i] = aChar then
      inc(result);
end;

function CapitalizeFirstLetter(Const aString: String; aStripNonAlpha: boolean; Const aIgnore: Array of String): String;
var
  a_Flag, a_IsAlpha: boolean;

  a_Index, a_POS: Byte;
  a_work, a_temp, a_Ignore: STRING;
  a_Char: Char;
begin
  a_flag := True;
  a_work := AnsiLowerCase(aString);
  a_temp := '';
  for a_Index := 1 to length(a_work) do
  begin
    a_Char := a_work[a_Index];
    a_IsAlpha := a_Char in ['a'..'z'];
{if aStripNonAlpha and IsAlpha...we add chars
 if not aStripNonAlpha we always add chars
 if the a_Flag has been flipped(we've capitalized one alpha char)...we add the rest of the chars
 All others we don't add...
}
    if (aStripNonAlpha and (a_IsAlpha)) or (not aStripNonAlpha) or (not a_Flag) then
    begin
      if a_Flag and a_IsAlpha then//we only uppercase the first Alpha char we come across
      begin
        a_temp := a_temp + UpCase(a_Char);
        a_Flag := False;  //we flip this so we don't uppercase any more chars
      end
      else
        a_temp := a_temp + a_Char;
    end;
  end;
  for a_Index := Low(aIgnore) to High(aIgnore) do
  begin
    a_Ignore := aIgnore[a_Index];
    a_POS := POS(a_Ignore, a_temp);
    if a_POS > 0 then
      StringReplace(a_temp, LowerCase(a_Ignore), a_Ignore, [rfReplaceAll]);
  end;
  Result := a_temp;
end;

procedure ConvertStringToArray(Const aString: string; Const aDelimeter : char; var aArray: Array of String);
var
  a_Count, a_Index, a_AIndex, a_Start, a_Len, a_Max: integer;
begin
  a_AIndex := 0;
  a_Start := 1;
  a_Len := 0;
  a_Count := Ord(High(aArray));
  a_Max := length(aString);
  for a_Index := 1 to a_Max do
  begin
    if (aString[a_Index] = aDelimeter) or (a_Index = a_Max) then
    begin
      if (a_Index = a_Max) then
        inc(a_Len);
      if (a_Index > a_Start) and (a_Count >= a_AIndex ) then
      begin
        aArray[a_AIndex] := Copy(aString, a_Start, a_Len);
        inc(a_AIndex);
        a_Start := a_Index + 1;
      end;
      a_Len := 0;
    end else
      inc(a_Len);
  end;
end;

procedure ModifiedFieldsWthIdFld(aDataset: TDataset; var aValues: TArrayField; aOriginal: TArrayVariant);
{the first field must be your ID field}
var
  a_Cnt: integer;
begin
  ModifiedFields(aDataset, aValues, aOriginal);
  a_Cnt := High(aValues);
  if a_Cnt > 0 then
  begin
    SetLength(aValues, a_Cnt +1);
    aValues[a_Cnt] := aDataset.Fields[0];
  end;
end;

procedure ModifiedFields(aDataset: TDataset; var aValues: TArrayField; aOriginal: TArrayVariant);
var
  a_Index, a_Count: integer;
begin
{This will let you know what Fields have been modified when used in conjuction with
StoreValues.  It will return an Array of TField that have been modified.  You use it like:

  if global.qryClaimDetail.State in [dsEdit, dsInsert] then
  begin
    try
      ModifieldFields(global.qryClaimDetail, a_Fields, FClientDetail);
      UpdateClaimDetail(global.qryWork,  a_Fields , 'xOrders', 'Where OrderId =' + UpdatedOrderID, Logger);
    finally
      global.qryClaimDetail.Cancel;// Post;
    end;}
  a_Count := 0;
  SetLength(aValues, aDataset.FieldCount);
  for a_Index := Low(aOriginal) to High(aOriginal) do
  begin
    if not VarSameValue(aOriginal[a_Index], aDataset.Fields[a_Index].Value) then
    begin
      aValues[a_Count] := aDataset.Fields[a_Index];
      inc(a_Count);
    end;
  end;
  SetLength(aValues, a_Count);
end;

procedure StoreValues(aDataset: TDataset; var aValues: TArrayVariant);
var
  a_Index: integer;
begin
{Need to setup and event handler for query like:

procedure TformClaimDetails.FormCreate(Sender: TObject);
begin
  global.qryClaimDetail.BeforeEdit := BeforeEdit;
end;

This should be in called in the BeforeEdit event handler for the Dataset like:

procedure TformClaimDetails.BeforeEdit(DataSet: TDataSet);
begin
  StoreValues(Dataset, FClientDetail);
end;}
  SetLength(aValues, aDataset.FieldCount);
  for a_Index := 0 to aDataset.FieldCount - 1 do
    aValues[a_Index] := aDataset.Fields[a_Index].Value;
end;

procedure _AddValue(aSQL: TWideStrings; aField: TField; aStart: integer);
var
  a_Prefix, a_Value: string;
  a_DataType: TFieldType;
begin
  a_DataType := aField.DataType;
  if aField.IsNull then
    a_Value := 'Null'
  else
  begin
    a_Value := aField.AsString;
    if a_DataType in cQuotedFldType then
      a_Value := QuotedStr(a_Value)
    else
    if a_DataType = ftBoolean then
      if aField.Value then
        a_Value := '1'
      else
        a_Value := '0';
    end;
  if aSQL.Count = aStart then
    a_Prefix := ''
  else
    a_Prefix := ',';
  aSQL.Add(a_Prefix + a_Value);
end;


procedure _AddField(aSQL: TWideStrings; aField: TField; aStart: integer);
var
  a_Prefix, a_Value: string;
  a_DataType: TFieldType;
begin
  a_DataType := aField.DataType;
  if aField.IsNull then
    a_Value := 'Null'
  else
  begin
    a_Value := aField.AsString;
    if a_DataType in cQuotedFldType then
      a_Value := QuotedStr(a_Value)
    else
    if a_DataType = ftBoolean then
      if aField.Value then
        a_Value := '1'
      else
        a_Value := '0';
    end;
  if aSQL.Count = aStart then
    a_Prefix := ''
  else
    a_Prefix := ',';
  aSQL.Add(a_Prefix + aField.FieldName + '=' + a_Value);
end;

procedure _AddFieldName(aSQL: TWideStrings; aField: TField; aStart: integer);
var
  a_Prefix, a_Field : string;
begin
  a_Field := aField.FieldName;
  if aSQL.Count = aStart then
    a_Prefix := ''
  else
    a_Prefix := ',';
  aSQL.Add(a_Prefix + a_Field);
end;

procedure DeleteDataset(aWork: TAdoQuery; aTableName, aWhere: string; aLogger: ILogger = nil);
begin
  aWork.Active := False;
  aWork.Sql.Clear;
  aWork.SQL.Add('Delete from '+ aTableName);
  aWork.SQL.Add(aWhere + ';');//add to the tail.
  aLogger.AddLog(aWork);
  aWork.ExecSQL;
end;

function InsertDataset(aWork: TAdoQuery;const  aDataset: TDataset; aTableName, aIdendityField, aWhere: string; aLogger: ILogger = nil): integer; overload;
var
  a_Index, a_Start: integer;
begin
  aWork.Active := False;
  aWork.Sql.Clear;
  aWork.SQL.Add('Insert into '+ aTableName);
  aWork.SQL.Add('(');
  a_Start := aWork.Sql.Count;
  for a_Index := 0 to aDataset.FieldCount -1 do
    if aIdendityField <> aDataset.Fields[a_Index].FieldName then
    begin
      _AddFieldName(aWork.SQL, aDataset.Fields[a_Index], a_Start);
    end else
      inc(a_Start);
  aWork.SQL.Add(')');
  aWork.SQL.Add('Values');
  aWork.SQL.Add('(');
  a_Start := aWork.Sql.Count;
  for a_Index := 0 to aDataset.FieldCount -1 do
    if aIdendityField <> aDataset.Fields[a_Index].FieldName then
      _AddValue(aWork.SQL, aDataset.Fields[a_Index], a_Start);
  aWork.SQL.Add(')');
  aWork.SQL.Add(aWhere + ';');//add to the tail.
  aWork.SQL.Add('Select @@Identity as ID');
  aLogger.AddLog(aWork);
  aWork.Active := True;
  Result := aWork.FieldByName('ID').asInteger;
end;

function InsertDataset(aWork: TAdoQuery;const  aFields: Array of TField; aTableName, aWhere: string; aLogger: ILogger = nil): integer;
var
  a_Index, a_Start: integer;
begin
  Result := 0;
  if Length(aFields) > 0 then
  begin
    aWork.Active := False;
    aWork.Sql.Clear;
    aWork.SQL.Add('Insert into '+ aTableName);
    aWork.SQL.Add('(');
    a_Start := aWork.Sql.Count;
    for a_Index := Low(aFields) to High(aFields) do
      _AddFieldName(aWork.SQL, TField(aFields[a_Index]), a_Start);
    aWork.SQL.Add(')');
    aWork.SQL.Add('Values');
    aWork.SQL.Add('(');
    a_Start := aWork.Sql.Count;
    for a_Index := Low(aFields) to High(aFields) do
      _AddValue(aWork.SQL, TField(aFields[a_Index]), a_Start);
    aWork.SQL.Add(')');
    aWork.SQL.Add(aWhere + ';');//add to the tail.
    aWork.SQL.Add('Select @@Identity as ID');
    aLogger.AddLog(aWork);
    aWork.Active := True;
    Result := aWork.FieldByName('ID').asInteger;
  end;
end;

procedure UpdateDataset(aWork: TAdoQuery;const  aFields: TArrayField; aTableName, aWhere: string; aLogger: ILogger = nil);
begin
  UpdateDataset(aWork, aFields, nil, aTableName, aWhere, aLogger);
end;

procedure UpdateDataset(aWork: TAdoQuery;const  aFields: TArrayField; aIdField: TField; aTableName, aWhere: string; aLogger: ILogger = nil);
var
  a_Index, a_Start: integer;
begin
  if Length(aFields) > 0 then
  begin
    aWork.Active := False;
    aWork.Sql.Clear;
    aWork.SQL.Add('Update '+ aTableName);
    aWork.SQL.Add('Set ');
    a_Start := aWork.SQL.Count;
    for a_Index := Low(aFields) to High(aFields) do
      if not Assigned(aIdField) or (aIdField <> aFields[a_Index]) then
        _AddField(aWork.SQL, TField(aFields[a_Index]), a_Start);
    aWork.SQL.Add(aWhere);//add to the tail.
    aLogger.AddLog(aWork);
    aWork.ExecSQL;
  end;
end;

procedure UpdateDataset(aWork: TAdoQuery;const  aFields: Array of TField; aTableName, aWhere: string; aLogger: ILogger = nil);
begin
  UpdateDataset(aWork, aFields, nil, aTableName, aWhere, aLogger);
end;

procedure UpdateDataset(aWork: TAdoQuery;const  aFields: Array of TField; aIdField: TField; aTableName, aWhere: string; aLogger: ILogger = nil); overload;
var
  a_Index, a_Start: integer;
begin
  if Length(aFields) > 0 then
  begin
    aWork.Active := False;
    aWork.Sql.Clear;
    aWork.SQL.Add('Update '+ aTableName);
    aWork.SQL.Add('Set ');
    a_Start := aWork.SQL.Count;
    for a_Index := Low(aFields) to High(aFields) do
      if not Assigned(aIdField) or (aIdField <> aFields[a_Index]) then
      _AddField(aWork.SQL, TField(aFields[a_Index]), a_Start);
    aWork.SQL.Add(aWhere);//add to the tail.
    aLogger.AddLog(aWork);
    aWork.ExecSQL;
  end;
end;

procedure UpdateDataset(aWork: TAdoQuery;const  aDataset: TDataset; aTableName, aIdendityField,  aWhere: string; aLogger: ILogger = nil); overload;
var
  a_Index, a_Start: integer;
begin
  aWork.Active := False;
  aWork.Sql.Clear;
  aWork.SQL.Add('Update '+ aTableName);
  aWork.SQL.Add('Set ');
  a_Start := aWork.SQL.Count;
  for a_Index := 0 to aDataset.FieldCount -1 do
   if aIdendityField <> aDataset.Fields[a_Index].FieldName then
    _AddField(aWork.SQL, aDataset.Fields[a_Index], a_Start);
  aWork.SQL.Add(aWhere);//add to the tail.
  aLogger.AddLog(aWork);
  aWork.ExecSQL;
end;

procedure LoadSQL(aQuery: TAdoQuery; aFileName, aSQLDir: string);
var
  a_SQL: TStrings;
begin
  if DirectoryExists(aSQLDir) then
  begin
    aSQLDir := IncludeTrailingPathDelimiter(aSQLDir);
    aFileName := aSQLDir + aFileName;
    if FileExists(aFileName) then
    begin
      aQuery.Active := False;
      aQuery.SQL.Clear;
      a_SQL := TStringList.Create;
      try
        a_SQL.LoadFromFile(aFileName);
        aQuery.SQL.AddStrings(a_SQL);
      finally
        a_SQL.Free;
      end;
    end;
  end;
end;

procedure LoadSQL(aQuery: TAdoQuery; aSQLDir: string;aOwner: TComponent);
var
  a_Name: string;
begin
  if Assigned(aOwner) then
    a_Name := aOwner.Name + '_' + aQuery.Name
  else
    a_Name := aQuery.Name;
  LoadSQL(aQuery, a_Name, aSQLDir);  
end;

function GetChecked(aCheckBox: TCheckBox): char;
begin
  if aCheckBox.Checked then
    Result := '1'
  else
    Result := '0';
end;

procedure SetChecked(aCheckBox:TCheckBox; aValue: char);
begin
  aCheckBox.Checked := aValue = '1';
end;

function GetYN(aCheckBox: TCheckBox): char;
begin
  if aCheckBox.Checked then
    Result := 'Y'
  else
    Result := 'N';
end;

procedure SetYN(aCheckBox: TCheckBox; aValue: char);
begin
  aCheckBox.Checked := aValue = 'Y';
end;

procedure SetComboBox(aComboBox: TComboBox; aValue: string);
begin
  aComboBox.ItemIndex := aComboBox.Items.IndexOf(aValue);
end;

procedure SetComboBox(aComboBox: TComboBox; aId: integer);
var
  a_Index: integer;
begin
  for a_Index := 0 to aComboBox.Items.Count - 1 do
    if Integer(aComboBox.Items.Objects[a_Index]) = aId then
    begin
      aComboBox.ItemIndex := a_Index;
      break;
    end;
end;

function GetValue(aComponent: TComponent): string;
begin
  if aComponent is TEdit then
    Result :=  TEdit(aComponent).Text
  else
  if aComponent is TComboBox then
    Result:= TComboBox(aComponent).Text
  else
    Raise Exception.Create('Unknown Component in GetValue:' + aComponent.Name);
end;

function GetValueAsInteger(aComponent: TComponent): integer;
begin
  if aComponent is TEdit then
    Result :=  StrToIntDef(TEdit(aComponent).Text, 0)
  else
  if (aComponent is TComboBox) then
    if ( TComboBox(aComponent).ItemIndex <> -1) then
      Result:= Integer(TComboBox(aComponent).Items[TComboBox(aComponent).ItemIndex])
    else
      Result := 0
  else
    Raise Exception.Create('Unknown Component in GetValue:' + aComponent.Name);

end;

procedure SetComponent(aComponent: TComponent; aValue: string; aUpperCase: boolean);
begin
  if aUpperCase then
    aValue := UpperCase(aValue);
  if aComponent is TEdit then
    TEdit(aComponent).Text := Trim(aValue)
  else
    Raise Exception.Create('Unknown Component in SetComponent:' + aComponent.Name);
end;

procedure SetComponent(aComponent: TComponent; aValue: TDateTime);
begin
  if aComponent is TDateTimePicker then
    TDateTimePicker(aComponent).DateTime := aValue
  else
    Raise Exception.Create('Unknown Component in SetComponent:' + aComponent.Name);  
end;


function StrToChar(const S: string; Default: Char = ' '): Char;
begin
  if length(S) > 0 then
    Result := S[1]
  else
    Result := Default;
end;

function GetValue(aField: TField; aNull: boolean;aQuote: boolean): string;
begin
  if aNull and aField.DataSet.Active and (aField.IsNull or (aField.AsString = ''))  then
    Result := 'Null'
  else
  begin
    if aField.DataSet.Active then
      Result := aField.AsString
    else
      Result := '';
    if aQuote then
      Result := QuotedStr(Result);
  end;
end;

function GetKeyValue(aField: TField; aNull: boolean): string;
begin
  if aNull and aField.DataSet.Active and (aField.IsNull or (aField.AsInteger = 0))  then
    Result := 'Null'
  else
  begin
    if aField.DataSet.Active then
      Result := aField.AsString
    else
      Result := '0';
  end;
end;

function GetYN(aField: TField; aUseQuote: boolean): string;
begin
  if aField.Dataset.Active then
    Result := StrToChar(aField.AsString, 'N')
  else
    Result := 'N';
  if aUseQuote then
    Result := QuotedStr(Result);
end;

function GetChecked(aField: TField): string; overload;
begin
  if aField.Dataset.Active then
    Result := StrToChar(aField.AsString, '0')
  else
    Result := '0';
end;

function GetIntStr(aField: TField): string;
begin
  if aField.AsString = '' then
    Result := '0'
  else
    Result := IntToStr(aField.AsInteger);  
end;

(*
This was the original NPIChecker
{ This function implements the Luhn Formula for Modulus 10 "double-add-double" Check Digit Formula.
The Luhn check digit formula is calculated as follows:
  1. Double the value of alternate digits beginning with the rightmost digit.
  2. Add the individual digits of the products resulting from step 1 to the unaffected digits from the original number.
  3. Subtract the total obtained in step 2 from the next higher number ending in zero. This is the check digit.
     If the total obtained in step 2 is a number ending in zero, the check digit is zero.

  Example:
  Assume the 9-position identifier part of the NPI is 123456789. Using the Luhn formula on the identifier portion,
  the check digit is calculated as follows:
    NPI without check digit:
    1 2 3 4 5 6 7 8 9
    Step 1: Double the value of alternate digits, beginning with the rightmost digit.
    2 6 10 14 18
    Step 2: Add constant 24, to account for the 80840 prefix that would be present on a card issuer identifier,
            plus the individual digits of products of doubling, plus unaffected digits.
    24 + 2 + 2 + 6 + 4 + 1 + 0 + 6 + 1 + 4 + 8 + 1 + 8 = 67
    Step 3: Subtract from next higher number ending in zero.
    70 Ã¢â‚¬â€œ 67 = 3
    Check digit = 3
    NPI with check digit = 1234567893
}
{
function TGlobal.NPIChecker(NPI: String): Boolean;
const
  ASCII_VAL_OF_CHAR_0 = 48;
var
  tmp: String;
  checkDigit,i,len,sum: Integer;
begin
  len := Length(Trim(NPI));
  if len < 10 then
    Result := False
  else begin
    try
      // ensure all digits are numeric
      StrToInt(Trim(NPI));
      except begin
        Result := False;
        exit;
      end;
    end;
    sum := 0;

    i := 9;
    while i > 0  do begin
      //Step 1
      Str((Integer(NPI[i]) - ASCII_VAL_OF_CHAR_0) * 2,tmp);
      if Length(tmp) > 1 then begin
        sum := sum + (Integer(tmp[1]) - ASCII_VAL_OF_CHAR_0);
        sum := sum + (Integer(tmp[2]) - ASCII_VAL_OF_CHAR_0);
      end else
        sum := sum + (Integer(tmp[1]) - ASCII_VAL_OF_CHAR_0);

      i := i-1;
      if i > 0 then begin
        //Add unaffected digits - part of Step 2
        sum := sum + (Integer(NPI[i]) - ASCII_VAL_OF_CHAR_0);
        i := i-1;
      end;
    end;

    // Remainder of step 2
    sum := sum + 24;

    // Step 3
    checkDigit := 10 - (sum mod 10);
    if (checkDigit mod 10) = 0 then
      checkDigit := 0;

    if (Integer(NPI[len]) - ASCII_VAL_OF_CHAR_0) = checkDigit then
      Result := True
    else
      Result := False;
  end;
end;
}
*)

(*
see ShlObj.pas

  CSIDL_DESKTOP                       = $0000;
  {$EXTERNALSYM CSIDL_INTERNET}
  CSIDL_INTERNET                      = $0001;
  {$EXTERNALSYM CSIDL_PROGRAMS}
  CSIDL_PROGRAMS                      = $0002;
  {$EXTERNALSYM CSIDL_CONTROLS}
  CSIDL_CONTROLS                      = $0003;
  {$EXTERNALSYM CSIDL_PRINTERS}
  CSIDL_PRINTERS                      = $0004;
  {$EXTERNALSYM CSIDL_PERSONAL}
  CSIDL_PERSONAL                      = $0005;
  {$EXTERNALSYM CSIDL_FAVORITES}
  CSIDL_FAVORITES                     = $0006;
  {$EXTERNALSYM CSIDL_STARTUP}
  CSIDL_STARTUP                       = $0007;
  {$EXTERNALSYM CSIDL_RECENT}
  CSIDL_RECENT                        = $0008;
  {$EXTERNALSYM CSIDL_SENDTO}
  CSIDL_SENDTO                        = $0009;
  {$EXTERNALSYM CSIDL_BITBUCKET}
  CSIDL_BITBUCKET                     = $000a;
  {$EXTERNALSYM CSIDL_STARTMENU}
  CSIDL_STARTMENU                     = $000b;
  {$EXTERNALSYM CSIDL_DESKTOPDIRECTORY}
  CSIDL_DESKTOPDIRECTORY              = $0010;
  {$EXTERNALSYM CSIDL_DRIVES}
  CSIDL_DRIVES                        = $0011;
  {$EXTERNALSYM CSIDL_NETWORK}
  CSIDL_NETWORK                       = $0012;
  {$EXTERNALSYM CSIDL_NETHOOD}
  CSIDL_NETHOOD                       = $0013;
  {$EXTERNALSYM CSIDL_FONTS}
  CSIDL_FONTS                         = $0014;
  {$EXTERNALSYM CSIDL_TEMPLATES}
  CSIDL_TEMPLATES                     = $0015;
  {$EXTERNALSYM CSIDL_COMMON_STARTMENU}
  CSIDL_COMMON_STARTMENU              = $0016;
  {$EXTERNALSYM CSIDL_COMMON_PROGRAMS}
  CSIDL_COMMON_PROGRAMS               = $0017;
  {$EXTERNALSYM CSIDL_COMMON_STARTUP}
  CSIDL_COMMON_STARTUP                = $0018;
  {$EXTERNALSYM CSIDL_COMMON_DESKTOPDIRECTORY}
  CSIDL_COMMON_DESKTOPDIRECTORY       = $0019;
  {$EXTERNALSYM CSIDL_APPDATA}
  CSIDL_APPDATA                       = $001a;
  {$EXTERNALSYM CSIDL_PRINTHOOD}
  CSIDL_PRINTHOOD                     = $001b;
  {$EXTERNALSYM CSIDL_LOCAL_APPDATA}
  CSIDL_LOCAL_APPDATA                 = $001c;
  {$EXTERNALSYM CSIDL_ALTSTARTUP}
  CSIDL_ALTSTARTUP                    = $001d;         // DBCS
  {$EXTERNALSYM CSIDL_COMMON_ALTSTARTUP}
  CSIDL_COMMON_ALTSTARTUP             = $001e;         // DBCS
  {$EXTERNALSYM CSIDL_COMMON_FAVORITES}
  CSIDL_COMMON_FAVORITES              = $001f;
  {$EXTERNALSYM CSIDL_INTERNET_CACHE}
  CSIDL_INTERNET_CACHE                = $0020;
  {$EXTERNALSYM CSIDL_COOKIES}
  CSIDL_COOKIES                       = $0021;
  {$EXTERNALSYM CSIDL_HISTORY}
  CSIDL_HISTORY                       = $0022;
  {$EXTERNALSYM CSIDL_PROFILE}
  CSIDL_PROFILE                       = $0028; { USERPROFILE }
  {$EXTERNALSYM CSIDL_CONNECTIONS}
  CSIDL_CONNECTIONS                   = $0031; { Network and Dial-up Connections }
  {$EXTERNALSYM CSIDL_COMMON_MUSIC}
  CSIDL_COMMON_MUSIC                  = $0035; { All Users\My Music }
  {$EXTERNALSYM CSIDL_COMMON_PICTURES}
  CSIDL_COMMON_PICTURES               = $0036; { All Users\My Pictures }
  {$EXTERNALSYM CSIDL_COMMON_VIDEO}
  CSIDL_COMMON_VIDEO                  = $0037; { All Users\My Video }
  {$EXTERNALSYM CSIDL_CDBURN_AREA}
  CSIDL_CDBURN_AREA                   = $003b; { USERPROFILE\Local Settings\Application Data\Microsoft\CD Burning }
  {$EXTERNALSYM CSIDL_COMPUTERSNEARME}
  CSIDL_COMPUTERSNEARME               = $003d; { Computers Near Me (computered from Workgroup membership) }
  {$EXTERNALSYM CSIDL_PROFILES}
  CSIDL_PROFILES                      = $003e;

*)

end.
