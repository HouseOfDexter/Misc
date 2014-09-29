unit TImageHelperU;

interface

uses
  Classes, ExtCtrls, Windows, Graphics, Jpeg, SysUtils, DB, Dialogs;

const
  cGrayList: Array[0..4] of TColor = (clBlack, clDkGray, clMedGray, clLtGray, clWhite);
  cBWList: Array[0..2] of TColor = (clBlack, clMedGray, clWhite);  
  cColorList:  Array[0..15] of TColor =
    (clBlack, clMaroon,  clGreen,  clOlive, clNavy, clPurple,
     clTeal,  clGray,    clSilver, clRed,   clLime, clYellow,
     clBlue,  clFuchsia, clAqua,   clWhite);


 cJPEGstarts = 'FFD8';
 cBMPstarts = '424D';  //BM

type
  TUnitType = (utPixels, utInches, utCM); 
  TImageChanged = procedure(Sender: TObject; aImage: TImage) of object;
  TImageHelper = class helper for TImage
  private
    function GetTransparentColor: TColor;

  public
    procedure Combine(aPosition: TPoint; aGraphic: TGraphic; aLoadOriginalFirst: boolean=True);overload;
    procedure Combine(aOffsetOriginal, aPosition: TPoint; aGraphic: TGraphic; aLoadOriginalFirst: boolean=True);overload;

    procedure Resize(aAspectRation: integer; aBitMap: TBitmap = nil);

    procedure Select(aRect: TRect); overload;
    procedure Select(aStart, aEnd: TPoint); overload;
    procedure ConvertToGraphic(aGraphicClass: TGraphicClass);
    procedure ConvertNearestColorToColor(aColor: TColor;aColorList: Array of TColor);
    procedure ConvertNearestColor(aColorList: Array of TColor);

    function LoadFromDB(aField: TField): boolean;
    function  LoadFromDialog(aDialog: TOpenDialog): boolean;
    procedure SaveToDB(aField: TField);
    procedure ClearImage;

    procedure ConvertDPI(aDPI: Integer); 

    procedure ZoomImage(aPercent: integer);

    property TransparentColor: TColor read GetTransparentColor;

    class function NearestColor(const aColor:  TColor; const aColorList: Array of TColor): TColor;
    class procedure SetCanvasZoomPercent(aCanvas: TCanvas; AZoomPercent: Integer);
    class function IsJpeg(aStream: TStream): boolean;

  end;

implementation

uses Controls;

{ TImageHelper }
{ clBlack = TColor($000000);
  clMaroon = TColor($000080);
  clGreen = TColor($008000);
  clOlive = TColor($008080);
  clNavy = TColor($800000);
  clPurple = TColor($800080);
  clTeal = TColor($808000);
  clGray = TColor($808080);
  clSilver = TColor($C0C0C0);
  clRed = TColor($0000FF);
  clLime = TColor($00FF00);
  clYellow = TColor($00FFFF);
  clBlue = TColor($FF0000);
  clFuchsia = TColor($FF00FF);
  clAqua = TColor($FFFF00);
  clLtGray = TColor($C0C0C0);
  clDkGray = TColor($808080);
  clWhite = TColor($FFFFFF);

  clMedGray = TColor($A4A0A0);
  }

class procedure TImageHelper.SetCanvasZoomPercent(aCanvas: TCanvas; aZoomPercent: Integer);
{code from http://www.swissdelphicenter.ch/en/showcode.php?id=968}
begin
  SetMapMode(aCanvas.Handle, MM_ISOTROPIC); //this sets the Canvas to scale along an arbitrary scale
  SetWindowExtEx(aCanvas.Handle, 100, 100, nil);//this sets the scale...for the above
  SetViewportExtEx(aCanvas.Handle, aZoomPercent, aZoomPercent, nil);
end;

procedure TImageHelper.ZoomImage(aPercent: integer);
{Code by http://stackoverflow.com/users/243614/sertac-akyuz}
var
  a_x, a_y: integer;
begin
  if not ((aPercent = 100) or (aPercent = 0)) then
  begin
    ConvertToGraphic(TBitmap);
    SetCanvasZoomPercent(Self.Picture.Bitmap.Canvas, aPercent);
    a_x := (Self.Width * 50 div aPercent) - (Self.Picture.Bitmap.Width div 2);
    a_y := (Self.Height * 50 div aPercent) - (Self.Picture.Bitmap.Height div 2);
    Self.Canvas.Draw(a_x, a_y, Self.Picture.Bitmap);
    if (a_x > 0) or (a_y > 0)  then
    begin
      Self.Canvas.Brush.Color := clWhite;
      ExcludeClipRect(Self.Canvas.Handle, a_x, a_y, a_x + Self.Picture.Bitmap.Width,
                      a_y + Self.Picture.Bitmap.Height);
      Self.Canvas.FillRect(Self.Canvas.ClipRect);
    end;
  end;
end;

class function TImageHelper.NearestColor(const aColor:  TColor; const aColorList: Array of TColor):  TColor;
var
  a_DistanceSquared:  integer;
  a_Index:  integer;
  a_R1, a_R2:  integer;
  a_G1, a_G2:  integer;
  a_B1, a_B2:  integer;
  a_SmallestDistanceSquared:  integer;
begin
{Note this is not linear in the color perception space}
{Original Code from www.efg2.com}
  Result := clBlack;                       // Assume black is closest color
  a_SmallestDistanceSquared := 256*256*256;  // Any distance would be shorter

  a_R1 := GetRValue(aColor);
  a_G1 := GetGValue(aColor);
  a_B1 := GetBValue(aColor);

  for a_Index := Low(aColorList) to High(aColorList) do
  begin
    a_R2 := GetRValue(aColorList[a_Index]);
    a_G2 := GetGValue(aColorList[a_Index]);
    a_B2 := GetBValue(aColorList[a_Index]);

    a_DistanceSquared := Sqr(a_R1-a_R2) + Sqr(a_G1-a_G2) + Sqr(a_B1-a_B2);
    if  a_DistanceSquared < a_SmallestDistanceSquared then
    begin
      Result := aColorList[a_index];
      a_SmallestDistanceSquared := a_DistanceSquared;
    end
  end
end {NearestColor};


procedure TImageHelper.Resize(aAspectRation: integer; aBitMap: TBitmap);
{You should pass the original Bitmap that was loaded so that you don't get
loss of detail for multiple calls to resize.}
var
  a_Rect: TRect;
  a_Bitmap: TBitmap;
  a_Pt: TPoint;
begin
  ConvertToGraphic(TBitmap);
  a_Bitmap := TBitmap.Create;
  try
    if not Assigned(aBitmap) then
      aBitMap := Picture.Bitmap;
    a_Bitmap.Assign(aBitMap);
    a_Rect.Left := 0;
    a_Rect.Top := 0;
    a_Rect.Right := Round(a_Bitmap.Width * (aAspectRation / 100));
    a_Rect.Bottom := Round(a_Bitmap.Height * (aAspectRation / 100));
    Picture.Bitmap.Height := 0;
    Picture.Bitmap.Width := 0;
    Picture.Bitmap.Width := a_Rect.Right;
    Picture.Bitmap.Height := a_Rect.Bottom;
//    Picture.Bitmap.Canvas.StretchDraw(a_Rect,a_Bitmap);
    GetBrushOrgEx(Picture.Bitmap.Canvas.Handle, a_Pt);
    SetStretchBltMode(Picture.Bitmap.Canvas.Handle, HALFTONE);
    SetBrushOrgEx(Picture.Bitmap.Canvas.Handle, a_Pt.x, a_Pt.y, @a_Pt);
    StretchBlt(Picture.Bitmap.Canvas.Handle,0, 0, a_Rect.Right, a_Rect.Bottom,
    a_Bitmap.Canvas.Handle,
       0,0,a_Bitmap.Width,a_Bitmap.Height,SRCCOPY);
    Self.Left:= 0;
    Self.Top := 0;
    Width := a_Rect.Right;
    Height := a_Rect.Bottom;
  finally
    a_Bitmap.Free;
  end;
end;

procedure WriteJpegHeader(aStream: TStream; aDPI: integer);
const
  cBufferSize = 50;
  cDPI = 1; //inch
  cDPC = 2; //cm
var
  a_Buffer: string;
  a_Index: integer;
  a_Type: Byte;
  a_Xres, a_YRes: Word;
begin
  SetLength(a_Buffer, cBufferSize);
  aStream.Read(a_Buffer[1], cBufferSize);
  a_Index := POS('JFIF' + #$00, a_Buffer);
  if a_Index <> 0 then
  begin
    aStream.Seek(a_Index + 6, soFromBeginning);
    a_Type := cDPI;
    aStream.Write(a_Type, 1);
    a_Xres := Swap(aDPI);
    aStream.Write(a_Xres, 2);
    a_Yres := Swap(aDPI);
    aStream.Write(a_Yres, 2);
  end;
end;

procedure TImageHelper.ConvertDPI(aDPI: Integer);
var
  a_Jpeg: TJpegImage;
  a_Stream: TStream;
begin
  a_Stream := nil;
  a_Jpeg := nil;
  try
    a_Jpeg := TJpegImage.Create;
    a_Stream := TMemoryStream.Create;
    a_Jpeg.Assign(Picture.Graphic);
    a_Jpeg.SaveToStream(a_Stream);
    WriteJpegHeader(a_Stream, aDPI);
    ConvertToGraphic(TJpegImage);
    Picture.Graphic.LoadFromStream(a_Stream);
  finally
    a_Jpeg.Free;
    a_Stream.Free;
  end;
end;

procedure TImageHelper.ConvertNearestColor(aColorList: array of TColor);
var
  a_X, a_Y: integer;
  a_Color: TColor;
begin
  ConvertToGraphic(TBitmap);
  for a_X := 0 to Picture.Bitmap.Width do
     for a_Y := 0 to Picture.Bitmap.Height do
     begin
       a_Color := ColorToRGB(Picture.Bitmap.Canvas.Pixels[a_X, a_Y]);
       Picture.Bitmap.Canvas.Pixels[a_X, a_Y] := NearestColor(a_Color, aColorList);
     end;
end;

procedure TImageHelper.ConvertNearestColorToColor(aColor: TColor;
  aColorList: array of TColor);
var
  a_X, a_Y: integer;
  a_Color: TColor;
begin
  ConvertToGraphic(TBitmap);
  for a_X := 0 to Picture.Bitmap.Width do
     for a_Y := 0 to Picture.Bitmap.Height do
     begin
       a_Color := ColorToRGB(Picture.Bitmap.Canvas.Pixels[a_X, a_Y]);
       if aColor = NearestColor(a_Color, aColorList) then
         Picture.Bitmap.Canvas.Pixels[a_X, a_Y] := aColor;
     end;
end;

procedure TImageHelper.ConvertToGraphic(aGraphicClass: TGraphicClass);
var
  a_Graphic: TGraphic;
begin
{This function uses polymorphism to copy the correct graphic information from
the original graphic object, assigns the values to itself.  Picture.Bitmap
creates a Bitmap object for you.  Bitmap.Assign then copies the original graphic
information to your newly created Bitmap/Graphic.}
  if Assigned(Picture.Graphic) and (Picture.Graphic.ClassType <> aGraphicClass) then
  begin
    a_Graphic := aGraphicClass.Create;
    try
      {Note as soon as you call Picture.Bitmap you lose your original Picture.Graphic
see TPicture.ForceType...by creating a Graphic Object...we save all it's values
and then Set or assign them to our new Graphic object...
If you pass a TBitMap this will allow you to use commands like CopyRect...,
If you pass something like TJpegImage or some other TGraphicClass it will
allow you to save in the correct format when you call SaveToFile}
      a_Graphic.Assign(Picture.Graphic);
//      Picture.Graphic.Assign(a_Graphic);
      if aGraphicClass = TBitmap then
      begin
        Picture.Bitmap.Assign(a_Graphic);
//        Picture.Bitmap.TransparentColor := Picture.Bitmap.Canvas.Pixels[0, Picture.Bitmap.Height - 1];
      end else
        Picture.Graphic := a_Graphic;
    finally
      a_Graphic.Free;
    end;
  end;
end;

function TImageHelper.GetTransparentColor: TColor;
begin
  ConvertToGraphic(TBitmap);
  Result := Picture.Bitmap.TransparentColor;
end;

class function TImageHelper.IsJpeg(aStream: TStream): boolean;
var
 a_buffer: Word;
 a_hex: string[2];
 a_Pos: integer;
begin
  Result := False;
  a_Pos := aStream.Position;
  try
    aStream.Seek(0, soFromBeginning);
    while (not Result) and (aStream.Position + 1 < aStream.Size) do
    begin
      aStream.ReadBuffer(a_buffer, 1);
      a_hex := IntToHex(a_buffer, 2);
      if a_hex = 'FF' then begin
       aStream.ReadBuffer(a_buffer, 1);
       a_hex:=IntToHex(a_buffer, 2);
       if a_hex = 'D8' then
         Result := True
       else if a_hex = 'FF' then
         aStream.Position := aStream.Position-1;
       end;
    end;
  finally
    aStream.Position := a_POS;  //we go back to the streams origianl pos.
    aStream.Seek(0,soFromCurrent);
  end;
end;

function TImageHelper.LoadFromDB(aField: TField): boolean;
var
  a_Jpeg: TJpegImage;
  a_Blob: TStream;
begin
  Result := False;
  if aField.Dataset.Active and aField.IsBlob and not aField.IsNull then
  begin
    a_Blob := aField.Dataset.CreateBlobStream(aField, bmRead);
    try
      if IsJpeg(a_Blob) then
      begin
        a_Jpeg := TJPEGImage.Create;
        try
          a_Jpeg.LoadFromStream(a_Blob);
          Picture.Graphic := a_Jpeg;
        finally
          a_Jpeg.Free;
        end;
      end else
        Picture.Bitmap.LoadFromStream(a_Blob);
      Result := True;
    finally
      a_Blob.Free;
    end;
  end;
end;

function TImageHelper.LoadFromDialog(aDialog: TOpenDialog): boolean;
begin
  Result := False;
  if aDialog.Execute then
  begin
    if FileExists(aDialog.Filename) then
    begin
      Picture.LoadFromFile(aDialog.Filename);
      Result := True;
    end;
  end;
end;

procedure TImageHelper.Select(aRect: TRect);
{This copies the Rect value into a Bitmap and then assigns the Canvas back to
itself overriding the original Canvas.}
var
  a_Dest: TRect;
begin
  ConvertToGraphic(TBitmap);
  a_Dest.Left := 0;
  a_Dest.Top := 0;
  a_Dest.Right := aRect.Right - aRect.Left;
  a_Dest.Bottom := aRect.Bottom - aRect.Top;

  Picture.Bitmap.Canvas.CopyMode := cmSrcCopy;
//    Picture.Bitmap.Canvas.StretchDraw(FSelectRect,Picture.Graphic);
//  Top := 0;
//  Left := 0;
  Canvas.CopyRect(a_Dest, Picture.Bitmap.Canvas, aRect);
  Width := a_Dest.Right;
  Height := a_Dest.Bottom;
  Picture.Graphic.Width := Width;
  Picture.Graphic.Height := Height;
end;

procedure TImageHelper.SaveToDB(aField: TField);
var
  a_Blob: TStream;
begin
  if aField.IsBlob then
  begin
    if not (aField.DataSet.State in [dsEdit, dsInsert]) then
      aField.DataSet.Edit;
    a_Blob := aField.DataSet.CreateBlobStream(aField, bmWrite);
    try
      Picture.Graphic.SaveToStream(a_Blob);
    finally
      a_Blob.Free;
    end;
  end;
end;

procedure TImageHelper.Select(aStart, aEnd: TPoint);
begin
  Select(Classes.Rect(aStart, aEnd));
end;

procedure TImageHelper.Combine(aPosition: TPoint; aGraphic: TGraphic; aLoadOriginalFirst: boolean);
var
  a_OrgPosition: TPoint;
begin
  a_OrgPosition.X := 0;
  a_OrgPosition.Y := 0;
  Combine(a_OrgPosition, aPosition, aGraphic, aLoadOriginalFirst);
end;

procedure TImageHelper.ClearImage;
begin
  if Assigned(Picture.Graphic) then
  begin
    Picture.Graphic := nil;
  end;
end;

procedure TImageHelper.Combine(aOffsetOriginal, aPosition: TPoint;
  aGraphic: TGraphic; aLoadOriginalFirst: boolean);
var
  a_Bmp: TBitmap;
  a_OrigWidth, a_OrigHeight, a_NewWidth, a_NewHeight: integer;
begin
{Note this function combines 2 pictures into 1 picture, This is useful if one of
the pictures is transparent}
  a_Bmp := TBitmap.Create;
  try
    a_OrigWidth := Picture.Graphic.Width + aOffsetOriginal.X;
    a_OrigHeight := Picture.Graphic.Height + aOffsetOriginal.Y;
    a_NewWidth := aGraphic.Width + aPosition.X;
    a_NewHeight := aGraphic.Height + aPosition.Y;

    if a_OrigWidth > a_NewWidth then
      a_Bmp.Width := a_OrigWidth
    else
      a_Bmp.Width := a_NewWidth;

    if a_OrigHeight > a_NewHeight then
      a_Bmp.Height := a_OrigHeight
    else
      a_Bmp.Height := a_NewHeight;
    if aLoadOriginalFirst then
    begin
      a_Bmp.Canvas.Draw(aOffsetOriginal.X,aOffsetOriginal.Y, Picture.Graphic);
      a_Bmp.Canvas.Draw(aPosition.X, aPosition.Y, aGraphic);
    end else
    begin
      a_Bmp.Canvas.Draw(aPosition.X, aPosition.Y, aGraphic);
      a_Bmp.Canvas.Draw(aOffsetOriginal.X,aOffsetOriginal.Y, Picture.Graphic);
    end;
    Picture.Graphic := a_Bmp;
  finally
    a_Bmp.Free;
  end;
end;


end.
