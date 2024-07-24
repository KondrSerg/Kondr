//------------------------------------------
// Barcod-to-Bitmap
// Модуль кодировщика EAN-13
// - проверка на валидность
// - возможность поворота
// - возможность инверсии
// - передачи в буфер обмена
//
// Author: Бадло Сергей Григорьевич
// H-page: http://raxp.radioliga.com
// Cистемные требования: win32
//------------------------------------------

unit barcod_EKV;

interface

uses Windows, SysUtils, Graphics, clipbrd, classes;

 procedure InvertBitmap(bmp: TBitmap);                                  // + инверсия
 procedure Emboss(bmp: TBitmap; AMount: Integer);                       // - контур
 procedure RotateBitmap(Bitmap: TBitmap; Angle: Double; bg: TColor);    // + поворот
 procedure BlackWhite(bmp: TBitmap; gray: boolean);                     // - ч-б/серое
 //============================================================================
 function barcode(digits: shortstring; fontcolor: tcolor; fontsize,angle: smallint;
  anglecolor: tcolor; clipbrd, valid, invert: boolean): tbitmap; // кодировщик EAN-13


implementation

uses
  UShtrih_EKV;

//--- инверсия
type PByteArray = ^TByteArray;
     TByteArray = array[0..32767] of Byte;
procedure InvertBitmap(bmp: TBitmap);
var x, y: Integer;
    ByteArray: PByteArray;
begin
 bmp.PixelFormat:= pf24bit;
 for y:= 0 to bmp.Height - 1 do begin
  ByteArray:= bmp.ScanLine[y];
  for x:= 0 to bmp.Width * 3 - 1 do
   ByteArray[x] := 255 - ByteArray[x]
 end
end;
//---

//--- контур
procedure Emboss(bmp: TBitmap; AMount: Integer);
var x, y, i : integer;
    p1, p2: PByteArray;
begin
 bmp.PixelFormat:= pf24bit;
 for i:= 0 to AMount do begin
  for y:= 0 to bmp.Height-2 do begin
   p1:= bmp.ScanLine[y];
   p2:= bmp.ScanLine[y+1];
   for x:= 0 to bmp.Width do begin
    p1[x*3]  := (p1[x*3]+(p2[(x+3)*3] xor $FF)) shr 1;
    p1[x*3+1]:= (p1[x*3+1]+(p2[(x+3)*3+1] xor $FF)) shr 1;
    p1[x*3+2]:= (p1[x*3+1]+(p2[(x+3)*3+1] xor $FF)) shr 1
   end
  end
 end
end;
//---

//--- поворот
procedure RotateBitmap(Bitmap: TBitmap; Angle: Double; bg: TColor);
type TRGB = record
       B, G, R: Byte;
     end;
     pRGB = ^TRGB;
     pByteArray = ^TByteArray;
     TByteArray = array[0..32767] of Byte;
     TRectList = array [1..4] of TPoint;

var x, y, W, H, v1, v2: Integer;
    Dest, Src: pRGB;
    VertArray: array of pByteArray;
    Bmp: TBitmap;

  procedure SinCos(AngleRad: Double; var ASin, ACos: Double);
  begin
    ASin := Sin(AngleRad);
    ACos := Cos(AngleRad);
  end;

  function RotateRect(const Rect: TRect; const Center: TPoint; Angle: Double): TRectList;
  var DX, DY: Integer;
      SinAng, CosAng: Double;
    function RotPoint(PX, PY: Integer): TPoint;
    begin
      DX := PX - Center.x;
      DY := PY - Center.y;
      Result.x := Center.x + Round(DX * CosAng - DY * SinAng);
      Result.y := Center.y + Round(DX * SinAng + DY * CosAng);
    end;
  begin
    SinCos(Angle * (Pi / 180), SinAng, CosAng);
    Result[1] := RotPoint(Rect.Left, Rect.Top);
    Result[2] := RotPoint(Rect.Right, Rect.Top);
    Result[3] := RotPoint(Rect.Right, Rect.Bottom);
    Result[4] := RotPoint(Rect.Left, Rect.Bottom);
  end;

  function Min(A, B: Integer): Integer;
  begin
    if A < B then Result := A
             else Result := B;
  end;

  function Max(A, B: Integer): Integer;
  begin
    if A > B then Result := A
             else Result := B;
  end;

  function GetRLLimit(const RL: TRectList): TRect;
  begin
    Result.Left := Min(Min(RL[1].x, RL[2].x), Min(RL[3].x, RL[4].x));
    Result.Top := Min(Min(RL[1].y, RL[2].y), Min(RL[3].y, RL[4].y));
    Result.Right := Max(Max(RL[1].x, RL[2].x), Max(RL[3].x, RL[4].x));
    Result.Bottom := Max(Max(RL[1].y, RL[2].y), Max(RL[3].y, RL[4].y));
  end;

  procedure Rotate;
  var x, y, xr, yr, yp: Integer;
      ACos, ASin: Double;
      Lim: TRect;
  begin
    W := Bmp.Width;
    H := Bmp.Height;
    SinCos(-Angle * Pi/180, ASin, ACos);
    Lim := GetRLLimit(RotateRect(Rect(0, 0, Bmp.Width, Bmp.Height), Point(0, 0), Angle));
    Bitmap.Width := Lim.Right - Lim.Left;
    Bitmap.Height := Lim.Bottom - Lim.Top;
    Bitmap.Canvas.Brush.Color := bg;
    Bitmap.Canvas.FillRect(Rect(0, 0, Bitmap.Width, Bitmap.Height));
    for y := 0 to Bitmap.Height - 1 do begin
      Dest := Bitmap.ScanLine[y];
      yp := y + Lim.Top;
      for x := 0 to Bitmap.Width - 1 do begin
        xr := Round(((x + Lim.Left) * ACos) - (yp * ASin));
        yr := Round(((x + Lim.Left) * ASin) + (yp * ACos));
        if (xr > -1) and (xr < W) and (yr > -1) and (yr < H) then begin
          Src := Bmp.ScanLine[yr];
          Inc(Src, xr);
          Dest^ := Src^;
        end;
        Inc(Dest);
      end;
    end;
  end;

begin
  Bitmap.PixelFormat := pf24Bit;
  Bmp := TBitmap.Create;
  try
    Bmp.Assign(Bitmap);
    W := Bitmap.Width - 1;
    H := Bitmap.Height - 1;
    if Frac(Angle) <> 0.0
      then Rotate
      else
    case Trunc(Angle) of
      -360, 0, 360, 720: Exit;
      90, 270: begin
        Bitmap.Width := H + 1;
        Bitmap.Height := W + 1;
        SetLength(VertArray, H + 1);
        v1 := 0;
        v2 := 0;
        if Angle = 90.0 then v1 := H
                        else v2 := W;
        for y := 0 to H do VertArray[y] := Bmp.ScanLine[Abs(v1 - y)];
        for x := 0 to W do begin
          Dest := Bitmap.ScanLine[x];
          for y := 0 to H do begin
            v1 := Abs(v2 - x)*3;
            with Dest^ do begin
              B := VertArray[y, v1];
              G := VertArray[y, v1+1];
              R := VertArray[y, v1+2];
            end;
            Inc(Dest);
          end;
        end
      end;
      180: begin
        for y := 0 to H do begin
          Dest := Bitmap.ScanLine[y];
          Src := Bmp.ScanLine[H - y];
          Inc(Src, W);
          for x := 0 to W do begin
            Dest^ := Src^;
            Dec(Src);
            Inc(Dest);
          end;
        end;
      end;
      else Rotate;
    end;
  finally
    Bmp.Free;
  end;
end;
//---


//--- градации
type
 TRGB=record
 b,g,r : byte;
end;
type
  ARGB=array [0..1] of TRGB;
  PARGB=^ARGB;
procedure BlackWhite(bmp: tbitmap; gray: boolean);
var i,j,c: integer;
    p: PARGB;
begin
 bmp.PixelFormat:= pf24bit;
 for i:=0 to bmp.Height-1 do begin
  p:= bmp.ScanLine[i];
  for j:= 0 to bmp.Width-1 do begin
   c:= round(0.3*p[j].r+0.59*p[j].g+0.11*p[j].b);
   if not gray then //градации серого
    if c >= 128 then c:= 255 else c:= 0;
   p[j].r:= c;
   p[j].g:= c;
   p[j].b:= c
  end
 end
end;
//---



//=============================================================================
function barcode(digits: shortstring; fontcolor: tcolor; fontsize,angle: smallint;
 anglecolor: tcolor; clipbrd, valid, invert: boolean): tbitmap;
var temp: string;
    i,sum_n,sum_y,crc:integer;
    bmp : tbitmap;

    parity: array [1..9] of array[1..7] of integer;
    table : array[0..2] of array [0..9] of string;
begin
 // проверка кол-ва-
 if length(digits)<>13 then
 Begin
   FShtrih_EKV.Memo1.Lines.Add(digits+'  length(digits)<>13');
   exit;
 End;
 // проверка на валидность CRC ------------------------------------------------
 if valid then begin
  sum_y:= 0;
  sum_n:= 0;
  i:= length(digits);
  while i>0 do begin
   crc:= strtoint(digits[i]);
   if not Odd(i) then
    sum_y:= sum_y + crc         //четные
     else if i<= length(digits)-2 then
      sum_n:= sum_n + crc;      //нечетные
   i:= i-1
  end;
  sum_y:= (sum_y*3) + sum_n;
  crc  := (sum_y div 10)*10 + 10 - sum_y; // округляем до 10-ков и -
  FShtrih_EKV.Memo1.Lines.Add(digits+' '+IntToStr(crc));
  if (crc <> strtoint(digits[13]))  and (crc<>10) then
  Begin
  FShtrih_EKV.Memo1.Lines.Add(digits+' '+IntToStr(crc)+ ' Error');
   messagebox(0,'проверка валидности не пройдена!','Error code...',0) ;
  End;
 end;
 // ---------------------------------------------------------------------------

 parity[1][2]:= 0;
 parity[1][3]:= 0;
 parity[1][4]:= 1;
 parity[1][5]:= 0;
 parity[1][6]:= 1;
 parity[1][7]:= 1;
 parity[2][2]:= 0;
 parity[2][3]:= 0;
 parity[2][4]:= 1;
 parity[2][5]:= 1;
 parity[2][6]:= 0;
 parity[2][7]:= 1;
 parity[3][2]:= 0;
 parity[3][3]:= 0;
 parity[3][4]:= 1;
 parity[3][5]:= 1;
 parity[3][6]:= 1;
 parity[3][7]:= 0;
 parity[4][2]:= 0;
 parity[4][3]:= 1;
 parity[4][4]:= 0;
 parity[4][5]:= 0;
 parity[4][6]:= 1;
 parity[4][7]:= 1;
 parity[5][2]:= 0;
 parity[5][3]:= 1;
 parity[5][4]:= 1;
 parity[5][5]:= 0;
 parity[5][6]:= 0;
 parity[5][7]:= 1;
 parity[6][2]:= 0;
 parity[6][3]:= 1;
 parity[6][4]:= 1;
 parity[6][5]:= 1;
 parity[6][6]:= 0;
 parity[6][7]:= 0;
 parity[7][2]:= 0;
 parity[7][3]:= 1;
 parity[7][4]:= 0;
 parity[7][5]:= 1;
 parity[7][6]:= 0;
 parity[7][7]:= 1;
 parity[8][2]:= 0;
 parity[8][3]:= 1;
 parity[8][4]:= 0;
 parity[8][5]:= 1;
 parity[8][6]:= 1;
 parity[8][7]:= 0;
 parity[9][2]:= 0;
 parity[9][3]:= 1;
 parity[9][4]:= 1;
 parity[9][5]:= 0;
 parity[9][6]:= 1;
 parity[9][7]:= 0;

 table[0][0]:= '0001101';
 table[0][1]:= '0011001';
 table[0][2]:= '0010011';
 table[0][3]:= '0111101';
 table[0][4]:= '0100011';
 table[0][5]:= '0110001';
 table[0][6]:= '0101111';
 table[0][7]:= '0111011';
 table[0][8]:= '0110111';
 table[0][9]:= '0001011';
 table[1][0]:= '0100111';
 table[1][1]:= '0110011';
 table[1][2]:= '0011011';
 table[1][3]:= '0100001';
 table[1][4]:= '0011101';
 table[1][5]:= '0111001';
 table[1][6]:= '0000101';
 table[1][7]:= '0010001';
 table[1][8]:= '0001001';
 table[1][9]:= '0010111';
 table[2][0]:= '1110010';
 table[2][1]:= '1100110';
 table[2][2]:= '1101100';
 table[2][3]:= '1000010';
 table[2][4]:= '1011100';
 table[2][5]:= '1001110';
 table[2][6]:= '1010000';
 table[2][7]:= '1000100';
 table[2][8]:= '1001000';
 table[2][9]:= '1110100';
 //
 temp:= '101';
 for i:=2 to 7 do
  temp:= temp + table[parity[strtoint(digits[1])][i]][strtoint(digits[i])];
 temp:= temp + '01010';

 for i:= 8 to 13 do
  temp:= temp + table[2][strtoint(digits[i])];
 temp:= temp + '101';
 //
 bmp:= Tbitmap.Create;
 bmp.Width := 225;
 bmp.Height:= 55;
 bmp.Canvas.Pen.Color:= fontcolor;
 bmp.Transparent:=True;
 i:= 1;
 while (i<=190) do begin
  if (temp[(i+1)div 2] = '1') then
   if (i <= 6) or ((i >= 90)and(i <= 100)) or (i >= 184) then begin
    bmp.Canvas.MoveTo(i+14,3);
    bmp.Canvas.LineTo(i+14,51);
    bmp.Canvas.MoveTo(i+15,3);
    bmp.Canvas.LineTo(i+15,51)
   end else begin
   bmp.Canvas.MoveTo(i+14,3);
   bmp.Canvas.LineTo(i+14,38);
   bmp.Canvas.MoveTo(i+15,3);
   bmp.Canvas.LineTo(i+15,38)
  end;
  i:= i+2
 end;

 bmp.Canvas.Font.Name := 'Arial';
 bmp.Canvas.Font.Color:= fontcolor;
 bmp.Canvas.Font.Size := fontsize;
 bmp.Canvas.TextOut(4, 38, digits[1]);
 bmp.Canvas.TextOut(30, 38, digits[2] + ' ' + digits[3] + ' ' + digits[4] + ' ' + digits[5] + ' ' + digits[6] + ' ' + digits[7]);
 bmp.Canvas.TextOut(121, 38, digits[8] + ' ' + digits[9] + ' ' + digits[10] + ' ' + digits[11] + ' ' + digits[12] + ' ' + digits[13]);
 // поворот-
 //RotateBitmap(bmp, angle, anglecolor);
 // инверсия-
 //if invert then InvertBitmap(bmp);
 // в буфер обмена-
 if clipbrd then
 clipboard.assign(bmp);

 result:= bmp
end;


end.