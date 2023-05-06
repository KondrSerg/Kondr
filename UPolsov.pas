unit UPolsov;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Data.Win.ADODB, Vcl.Grids,
  Vcl.StdCtrls, Vcl.CheckLst, Vcl.Buttons;

type
  TFPolsov = class(TForm)
    chklst1: TCheckListBox;
    StringGrid1: TStringGrid;
    ADOQuery1: TADOQuery;
    ADOConnection1: TADOConnection;
    btn1: TSpeedButton;
    btn2: TSpeedButton;
    Memo2: TMemo;
    procedure FormShow(Sender: TObject);
    procedure chklst1ClickCheck(Sender: TObject);
    procedure StringGrid1DblClick(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure btn2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FPolsov: TFPolsov;

implementation

uses
  Main, UpolsRed;

{$R *.dfm}

procedure TFPolsov.btn1Click(Sender: TObject);
begin
  Form1.ExportGridtoExcel1(StringGrid1);
end;

procedure TFPolsov.btn2Click(Sender: TObject);
begin
    // chklst1.CheckAll(cbChecked);
     chklst1ClickCheck(Sender) ;
end;

procedure TFPolsov.chklst1ClickCheck(Sender: TObject);
var I,ind,y:Integer;
S,Serv:string;
begin
    S := ExtractFileDir(ParamStr(0));
    Memo2.Lines.Clear;
    Memo2.Lines.LoadFromFile(S + '\ConnKlap.ini');
    Serv:= Memo2.Lines[7] ;

    Y:=0;
    Form1.Clear_StringGrid(StringGrid1);
    StringGrid1.Cells[0,0]:= 'Фамилия';
    StringGrid1.Cells[1,0]:= 'Логин';
    StringGrid1.ColWidths[1]:=0;
    StringGrid1.Cells[2,0]:= 'Отдел';
    StringGrid1.Cells[3,0]:= 'Участок';
    StringGrid1.Cells[4,0]:= 'Руководитель';
    for ind :=0 to chklst1.Items.Count-1 do
    Begin
    try
     if chklst1.Checked[ind] then
     Begin
      ADOQuery1.Close;
      ADOQuery1.ConnectionString := 'Provider=ADsDSOObject;Encrypt Password=False;Data Source='+Serv+';'+
      'Mode=Read;Bind Flags=0;ADSI Flag=-2147483648';
      ADOQuery1.SQL.Clear;                 //Description
      ADOQuery1.SQL.Add('select samAccountName,name,CN,MAIL,title,Department,facsimileTelephoneNumber from ');
      ADOQuery1.SQL.Add(#39+'LDAP://'+Serv+#39);
      ADOQuery1.SQL.Add(' WHERE objectcategory = '+#39+'User'+#39+' AND memberof=''CN='+
      chklst1.Items.Strings[ind]+',CN=Users,DC=veza-miass,DC=local''' );
      ADOQuery1.SQL.Add(' order by CN');

      ADOQuery1.open;

      StringGrid1.RowCount:=StringGrid1.RowCount+ADOQuery1.RecordCount;
      for I:=0 to ADOQuery1.RecordCount-1 do
      begin
            StringGrid1.Cells[0,Y+1]:=ADOQuery1.FieldByName('CN').AsString;
            StringGrid1.Cells[1,Y+1]:=ADOQuery1.FieldByName('samAccountName').AsString;
            StringGrid1.Cells[2,Y+1]:=ADOQuery1.FieldByName('Department').AsString;
            StringGrid1.Cells[3,Y+1]:=ADOQuery1.FieldByName('title').AsString;
            StringGrid1.Cells[4,Y+1]:=ADOQuery1.FieldByName('facsimileTelephoneNumber').AsString;
            //StringGrid1.Cells[2,Y+1]:=ADOQuery1.FieldByName('').AsString;
            Inc(Y);
            ADOQuery1.Next;
      end;
      End;
    except on E: Exception do
      showmessage('Ошибка: '+e.Message);
    end;
    End;
end;

procedure TFPolsov.FormShow(Sender: TObject);
var I,ind,y:Integer;
begin
     chklst1.CheckAll(cbChecked);
     chklst1ClickCheck(Sender) ;
end;

procedure TFPolsov.StringGrid1DblClick(Sender: TObject);
begin
  FPolsRed.Caption:=StringGrid1.Cells[0,StringGrid1.Row];
  FPolsRed.Label4.Caption:=StringGrid1.Cells[1,StringGrid1.Row];
  FPolsRed.ShowModal;
end;

end.
