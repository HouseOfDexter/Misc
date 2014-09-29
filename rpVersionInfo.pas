unit rpVersionInfo;
 //version 1.23 3/5/2009 rewritten and tested with Delphi 2006.
 //the original version was written in 6/03/98 and tested with Delphi 3.
 //minor changes - changed varibles to cardinals and minor bug fix
(*Written by Rick Peterson, this component is released to the public domain for
  any type of use, private or commercial.  Should you enhance the product
  please give me credit and send me a copy.  Also please report any bugs to me.
  Send all corespondence to houseofdexter@gmail.com.

  Thanks to Dietrich Raisin and  Ronald for the enhancements
  (http://www.experts-exchange.com/M_29098.html)*)

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  TypInfo;

type
{$M+}
(* Have you seen the $M+ before???This tells delphi to publish RTTI info for
   enumerated types.  Basically allowing your enumerated types to act as
   strings with GetEnumName *)
  TVersionType=(vtCompanyName, vtFileDescription, vtFileVersion, vtInternalName,
                vtLegalCopyright,vtLegalTradeMark, vtOriginalFileName,
                vtProductName, vtProductVersion, vtComments);
{$M-}
  TrpVersionInfo = class(TComponent)
(* This component will allow you to get Version Info from your program at
   RunTime *)
  private
    FVersionInfo : Array [0 .. ord(high(TVersionType))] of string;
    FAppName: string;
    function GetAppName: string;
    procedure SetAppName(const Value: string);
  protected
    function GetVersionInfo(index:integer): string;
    procedure SetVersionInfo;
  public
    constructor Create(AOwner: TComponent); override;
//NOTE: ApplicationName does not have to be set if this is in an application
//If it's in a DLL or console you are required to set it.
    property ApplicationName: string read GetAppName write SetAppName;

  published    
    property CompanyName: string index ord(vtCompanyName) read GetVersionInfo;
    property FileDescription: string index ord(vtFileDescription) read GetVersionInfo;
    property FileVersion: string index ord(vtFileVersion) read GetVersionInfo;
    property InternalName: string index ord(vtInternalName) read GetVersionInfo;
    property LegalCopyright: string index ord(vtLegalCopyright) read GetVersionInfo;
    property LegalTradeMark: string index ord(vtLegalTradeMark) read GetVersionInfo;
    property OriginalFileName: string index ord(vtOriginalFileName) read GetVersionInfo;
    property ProductName: string index ord(vtProductName) read GetVersionInfo;
    property ProductVersion: string index ord(vtProductVersion) read GetVersionInfo;
    property Comments: string index ord(vtComments) read GetVersionInfo;

(* Label1.Caption := VersionInfo1.FileVersion;  Simple as that.
   NOTE:  Most of the properties are READONLY so you can not set them with the
   Object Inspector, but you can see the Version Info for Delphi ;) *)

  end;

procedure Register;

implementation

constructor TrpVersionInfo.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  SetVersionInfo;
end;

function TrpVersionInfo.GetAppName: string;
begin
  if FAppName = '' then
    if Assigned(Application) then begin
      FAppName := Application.ExeName;
    end;
  Result := FAppName;
end;

procedure TrpVersionInfo.SetAppName(const Value: string);
begin
  if Value <> FAppName then
    FAppName := Value;
{If you are using a DLL you probably want to set this otherwise ignore it and
it will be set for you...If you don't know the name check out
 GetModuleFileName(hModule, lpFileName, nSize)}
end;

function TrpVersionInfo.GetVersionInfo(index: integer): string;
begin
  result := FVersionInfo[index];
end;


procedure TrpVersionInfo.SetVersionInfo;
type
   PLongInt= ^LongInt;
var
  sAppName,sVersionType : string;
  i: integer;
  cLenOfValue, cAppSize: Cardinal;
  pcBuf,pcValue: PChar;
begin
  pcValue :=  nil;
  sAppName := ApplicationName;
  cAppSize:= GetFileVersionInfoSize(PChar(sAppName),cAppSize);
  if cAppSize > 0 then  begin
  //moved the AllocMem outside the Try/Finally block where it should be
    pcBuf := AllocMem(cAppSize);
    try
      GetFileVersionInfo(PChar(sAppName),0,cAppSize,pcBuf);
      for i := 0 to Ord(High(TVersionType)) do  begin
        sVersionType := GetEnumName(TypeInfo(TVersionType),i);
        sVersionType := Copy(sVersionType,3,length(sVersionType));
        VerQueryValue(pcBuf,PChar('\VarFileInfo\Translation'),
              Pointer(pcValue),cLenOfValue);
        sVersionType:= IntToHex(LoWord(PLongInt(pcValue)^),4)+
                         IntToHex(HiWord(PLongInt(pcValue)^),4)+
                         '\'+sVersionType;
        if VerQueryValue(pcBuf,PChar('\StringFileInfo\'+
                              sVersionType), Pointer(pcValue),cLenOfValue) then
          FVersionInfo[i] := pcValue;
      end;
    finally
      FreeMem(pcBuf,cAppSize);
    end;
  end;
end;

procedure Register;
begin
  RegisterComponents('OmniSYS', [TrpVersionInfo]);
end;

end.
