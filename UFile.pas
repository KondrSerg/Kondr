unit UFile;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Menus, Vcl.StdCtrls, Vcl.ExtCtrls,
  Vcl.ComCtrls,shellapi, Vcl.Buttons;

type
  TFFile = class(TForm)
    ListView1: TListView;
    pnl1: TPanel;
    Edit3: TEdit;
    dlgSave1: TSaveDialog;
    dlgOpen2: TOpenDialog;
    pm4: TPopupMenu;
    N12: TMenuItem;
    N9: TMenuItem;
    N10: TMenuItem;
    N11: TMenuItem;
    Label1: TLabel;
    btn1: TButton;
    Memo1: TMemo;
    btn2: TSpeedButton;
    procedure FormCreate(Sender: TObject);
function AddFile(FileMask: string; FFileAttr:DWORD): Boolean;
  function SlashSep(Path, FName: string): string;

function FileTimeToDateTimestr(FileTime: TFileTime): string;
    procedure FormShow(Sender: TObject);
    procedure ListView1DblClick(Sender: TObject);
    procedure N9Click(Sender: TObject);
    procedure N12Click(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure N11Click(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FFile: TFFile;

implementation

{$R *.dfm}

uses Main;
function MyCopyFile( InFile,OutFile: String; From,Count: Longint ): Longint;
var
  InFS,OutFS: TFileStream;
begin
  InFS := TFileStream.Create( InFile, fmOpenRead );//создаем поток
  OutFS := TFileStream.Create( OutFile, fmCreate );//создаем поток
  InFS.Seek( From, soFromBeginning );//перемещаем указатель в From
  Result := OutFS.CopyFrom( InFS, Count );
  InFS.Free;//освобождаем
  OutFS.Free;//освобождаем
end;
procedure TFFile.FormCreate(Sender: TObject);
var

SysImageList: uint;
SFI: SHFILEINFO;

begin

//Создаем списки маленьких и больших иконок.

ListView1.LargeImages:=TImageList.Create(self);

ListView1.SmallImages:=TImageList.Create(self);

//Запрашиваем, большие иконки

SysImageList := SHGetFileInfo('', 0, SFI,

SizeOf(SHFILEINFO), SHGFI_SYSICONINDEX or SHGFI_LARGEICON);

if SysImageList <> 0 then begin

//Присваиваем системные иконки в ListViewl

ListView1.Largeimages.Handle := SysImageList;

ListView1.Largeimages.ShareImages := True;

end;

//Запрашиваем маленькие иконки

SysImageList := SHGetFileInfo('', 0, SFI, SizeOf (SHFILEINFO) ,

SHGFI_SYSICONINDEX or SHGFI_SMALLICON) ;

if SysImageList <> 0 then
begin

//Присваиваем маленькие системные иконки в ListViewl

ListView1.Smallimages.Handle := SysImageList;

ListView1.Smallimages.ShareImages := True;
end;
end;

procedure TFFile.FormShow(Sender: TObject);
begin
  edit3.Text:='V:\КТО\Развертки\ВРАН\*.*';
  AddFile(edit3.Text,faAnyFile);
end;
//+++++++++++++++
procedure CopyFiles(const FromFolder: string; const ToFolder: string);
var
  Fo: TSHFileOpStruct;
  buffer: array [0 .. 4096] of char;
  p: pchar;
begin
  FillChar(buffer, sizeof(buffer), #0);
  p := @buffer;
  StrECopy(p, pchar(FromFolder)); // директория, которую мы хотим скопировать
  FFile.Memo1.Lines.Add(FromFolder+'');
  FillChar(Fo, sizeof(Fo), #0);
  Fo.Wnd    := Application.Handle;
  Fo.wFunc :=FO_COPY;//FO_MOVE ;
  Fo.pFrom := @buffer;
  Fo.pTo := pchar(ToFolder); // куда будет скопирована директория
 FFile.Memo1.Lines.Add(ToFolder+'');
  Fo.fFlags := 0;
  if ((SHFileOperation(Fo) <> 0) or (Fo.fAnyOperationsAborted <> false)) then
  begin
    FFile.Memo1.Lines.Add(IntToStr(Fo.wFunc));
    FFile.Memo1.Lines.Add(BoolToStr(Fo.fAnyOperationsAborted));
    ShowMessage('File copy process cancelled') ;
  end
  else
  begin
      FFile.Memo1.Lines.Add(IntToStr(Fo.wFunc));
      FFile.Memo1.Lines.Add(BoolToStr(Fo.fAnyOperationsAborted));
    //ShowMessage('File copy Ok') ;
  end;
end;
//||||||||||||||||||||||
procedure TFFile.ListView1DblClick(Sender: TObject);
Var Res1:Integer;
S,S1,S3,Cap:string;
begin
  S:= edit3.Text;
  S1:= edit3.Text;
  Cap:=ListView1.Items.Item[ListView1.ItemIndex].Caption;
 Res1:=Pos('*.*',S);
 if Res1<>0  then
 Delete(S,Res1,3);
 S3:=S+Cap+'\*.*';
 if not AddFile(S3,faAnyFile) then
 begin
      CreateDir(S+'Архив');
      AddFile(S1,faAnyFile);
      CopyFiles(S+Cap,S+'Архив\'+FormatDateTIme('dd.mm.YYYY',Now)+Cap );

      ShellExecute(0,nil,PWideChar(S+Cap),nil,nil, SW_SHOWNORMAL);
      Exit;
 end;
 edit3.Text:=S3;
end;
procedure TFFile.N11Click(Sender: TObject);
Var  SaveF :file ;

begin
    AssignFile(SaveF, 'FileName.ini');
    FileExists('FileName.ini');
end;

procedure TFFile.N12Click(Sender: TObject);
Var S,S1:string;
Res1:Integer;
begin
 S:= edit3.Text;
 Res1:=Pos('*.*',S);
 if Res1<>0  then
 Delete(S,Res1,3);
 S1:=S+ ListView1.Items.Item[ListView1.ItemIndex].Caption;
  ShellExecute(0,nil,PWideChar(S1),nil,nil, SW_SHOWNORMAL);
end;

procedure TFFile.N9Click(Sender: TObject);
Var S,S1:string;
Res1:Integer;
begin
 S:= edit3.Text;
 Res1:=Pos('*.*',S);
 if Res1<>0  then
 Delete(S,Res1,3);
 S1:=S+ ''+ListView1.Items.Item[ListView1.ItemIndex].Caption;
 dlgSave1.FileName:=S1;
 if dlgSave1.Execute() then
 begin
   MyCopyFile(S1,dlgSave1.FileName,0,0);
 end;
 //CopyFilesToClipboard(S1);
end;


function TFFile.SlashSep(Path, FName: string): string;

begin

if Path[Length(Path)]='\'  then
Result := Path +'\' + FName else

Result := Path + FName;
end;

procedure TFFile.btn1Click(Sender: TObject);
begin
  edit3.Text:='V:\КТО\_Развертки new\СТАМ 18.05.2018\*.*';
  AddFile(edit3.Text,faAnyFile);
end;

function TFFile.FileTimeToDateTimestr(FileTime: TFileTime): string;
var

LocFTime: TFileTime; SysFTime: TSystemTime; Dt, Tm: TDateTime;
begin

FileTimeToLocalFileTime(FileTime, LocFTime);
FileTimeToSystemTime (LocFTime, SysFTime) ;
 try

with SysFTime do begin

Dt := EncodeDate(wYear, wMonth, wDay);

Tm := EncodeTime(wHour, wMinute, wSecond, wMilliseconds);

end;

Result := DateTimeToStr(Dt+Tm);
except

Result := '';
 end;

end;
function TFFile.AddFile(FileMask: string; FFileAttr: DWORD): Boolean;
var

Shinfo: SHFILEINFO; attributes: string;

FileName: string;

hFindFile: THandle;

SearchRec: TSearchRec;

function AttrStr(Attr: integer): string; begin Result := '';

if (FILE_ATTRIBUTE_directory and Attr) > 0 then Result := Result + '';

if (FILE_ATTRIBUTE_ARCHIVE and Attr) > 0 then Result := Result + 'A';

if (FILE_ATTRIBUTE_READONLY and Attr) > 0 then Result := Result + 'R';

if (FILE_ATTRIBUTE_HIDDEN and Attr) > 0 then Result := Result + 'H';

if (FILE_ATTRIBUTE_SYSTEM and Attr) > 0 then Result := Result + 'S';

end;

begin

ListView1.Items.BeginUpdate;
ListView1.Items.Clear;

Result := False;

hFindFile := FindFirst(FileMask, FFileAttr, SearchRec);

if hFindFile <> INVALID_HANDLE_VALUE then
try
repeat

with SearchRec.FindData do
begin

if (SearchRec.Name = '.') or (SearchRec.Name = '..') or
(SearchRec.Name = '') then
continue;

FileName := SlashSep(edit3.Text, SearchRec.Name) ;

SHGetFileInfo(PChar(FileName), 0, Shinfo, SizeOf(Shinfo),

SHGFI_TYPENAME or SHGFI_SYSICONINDEX) ;

Attributes := AttrStr(dwFileAttributes); //Добавляю новый элемент

with ListView1.Items.Add do
begin //Присваиваю его имя

Caption := SearchRec.Name;

//Присваиваю индекс из системного списка изображений

Imageindex := Shinfo.iIcon; //Присваиваю размер

Subitems.Add(IntToStr(SearchRec.Size));

Subitems.Add((Shinfo.szTypeName) ) ;

Subitems.Add(FileTimeToDateTimeStr(ftLastWriteTime)) ;

Subitems.Add(attributes) ;

Subitems.Add(edit3.Text + Filename) ;

if (FILE_ATTRIBUTE_DIRECTORY and dwFileAttributes) > 0 then

Subitems.Add ('dir') else

Subitems.Add('file');
end;

Result := true; end;

until (FindNext(SearchRec) <> 0); finally

FindClose(SearchRec);

end;

ListView1.Items.EndUpdate;
end;

end.
