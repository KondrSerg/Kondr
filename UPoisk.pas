unit UPoisk;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.ComCtrls,UConnCeh,UConnKlap,UOsnova,
  Data.DB, Data.Win.ADODB, Vcl.Grids, Vcl.Buttons;


type
  TFPoisk = class(TForm)
    Memo1: TMemo;
    dtp1: TDateTimePicker;
    dtp2: TDateTimePicker;
    rg1: TRadioGroup;
    Edit1: TEdit;
    Button1: TButton;
    CBB1: TComboBox;
    Label1: TLabel;
    ADOQuery1: TADOQuery;
    ADOConnection1: TADOConnection;
    SG1: TStringGrid;
    btn1: TSpeedButton;
    procedure Button1Click(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure Edit1KeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
  UOsnova_Main:Osnova_Main ;
  Conn_Ceh: Connect_Miass_Ceh;
    Conn_Klap: Connect_Miass_Klap;
    { Public declarations }
  end;

var
  FPoisk: TFPoisk;

implementation
uses
Main;


{$R *.dfm}

procedure TFPoisk.btn1Click(Sender: TObject);
begin
  Form1.ExportGridtoExcel2(SG1);
end;

procedure TFPoisk.Button1Click(Sender: TObject);
var Mes,God,D,Dat,Dat1,SS1,S,S1,SS,Zak,Err:string;
Xl:Variant;
I,XX,J:Integer;
begin
    Conn_Klap := Connect_Miass_Klap.Create();
      God := FormatDateTime('yyyy', DTP1.DateTime);
  Mes := FormatDateTime('mmmm', DTP1.DateTime);


  Dat := FormatDateTime('mm.dd.yyyy', DTP1.DateTime);
  Dat1 := FormatDateTime('mm.dd.yyyy', DTP2.DateTime);
  ss1 := '  BETWEEN ' + #39 + Dat + #39 + ' And ' + #39 + Dat1 + #39;
  {XL := CreateOleObject('Excel.Application'); //
  XL.Application.EnableEvents := false;
  XL.Workbooks.Open(D + '\CKlapana\2013\�������.xlsx');
  XL.ActiveWorkBook.WorkSheets[2].Range['B' + IntToStr(1), 'B' + IntToStr(1)] := Dat2;
  XL.ActiveWorkBook.Sheets.Item[2].Activate;}
  //OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO
  ZAk := '';
  SS := '((������.������� LIKE ' + #39 + '%'+Edit1.Text+'%'+#39+') or (������.����������� LIKE ' + #39 +
   '%'+Edit1.Text+'%'+#39+') ) ';
  if not Conn_Klap.mkQuerySelect(
  'Select ������.�����������, ������.�������,������.�����������,Sum(������.����������*������.[��� �� ����������] ) As S ' +
  'FROM %s INNER JOIN ������ ON ������.Id�� = ������.Id��  Where (������.[������������] %s' +
  ') AND %s Group BY ������.�����������,������.�������,������.�����������', ['������', ss1, SS]) then
    exit;

  XX := 5;
        //�������� ��������� �������  idgp,zakaz,Nom
  Form1.SetConnectTemp(Form1.ADOQueryTemp);
    //�������� ��������� �������
  Form1.ADOQueryTemp := TADOQuery.Create(nil);
  Form1.ADOQueryTemp.Connection := Form1.ADOConnectionTemp;

  Form1.ADOQueryTemp.SQL.Text := 'CREATE TABLE #ClientToDBFL (����������� nvarchar(350),������� nvarchar(350),' + '����������� nvarchar(350),Kol float )';
  try
    Form1.ADOQueryTemp.ExecSQL;
  except
    Err := '1222';
  end;
  for J := 0 to Conn_Klap.AQuery.RecordCount - 1 do
  begin
        //SetConnectTemp(ADOQueryTemp);
    Form1.ADOQueryTemp.SQL.Text := 'INSERT INTO #ClientToDBFL ' + '(����������� nvarchar(350),�������,�����������,Kol) Values (' + #39 + Conn_Klap.AQuery.FieldByName('�����������').AsString + #39 + ',' + #39 + Conn_Klap.AQuery.FieldByName('�������').AsString + #39 + ',' + #39 + Conn_Klap.AQuery.FieldByName('�����������').AsString + #39 + ',' + #39 + Conn_Klap.AQuery.FieldByName('S').AsString + #39 + ')';
    try
     Form1. ADOQueryTemp.ExecSQL;
    except
      Err := '1222';
    end;

    Conn_Klap.AQuery.Next;
  end;
   //ppppppppppppppppppppppppppppppppppppppppppppppp
  //SS := '����������.������� LIKE ' + #39 + '���%' + #39;
 // SS := '(����������.������� LIKE ' + #39 + '���%'+#39+') ';
 SS := '((����������.������� LIKE ' + #39 + '%'+Edit1.Text+'%'+#39+') or (����������.����������� LIKE ' + #39 +
 '%'+Edit1.Text+'%'+#39+') ) ';
  if not Conn_Klap.mkQuerySelect('Select ����������.�����������,����������.�������,����������.�����������,Sum(����������.����������*����������.[��� �� ����������] ) As S ' +
  'FROM %s INNER JOIN ���������� ON ����������.Id�� = ����������.Id��  Where (����������.[������������] %s ' +
  ') AND (%s)  Group BY ����������.�����������,����������.�������,����������.�����������', ['����������', ss1, SS]) then
    exit;

  Inc(XX);
  for J := 0 to Conn_Klap.AQuery.RecordCount - 1 do
  begin
    Form1.ADOQueryTemp.SQL.Text := 'INSERT INTO #ClientToDBFL ' + '(�����������,�������,�����������,Kol) Values ('
     + #39 + Conn_Klap.AQuery.FieldByName('�����������').AsString + #39 + ','+ #39
      + Conn_Klap.AQuery.FieldByName('�������').AsString + #39 + ',' + #39
      + Conn_Klap.AQuery.FieldByName('�����������').AsString + #39 + ','
      + #39 + Conn_Klap.AQuery.FieldByName('S').AsString + #39 + ')';
    try
      Form1.ADOQueryTemp.ExecSQL;
    except
      Err := '1222';
    end;
    Conn_Klap.AQuery.Next;
  end;
  Form1.ADOQueryTemp.SQL.Text := ('Select �����������,�������,�����������,Sum(Kol ) As S ' +
  'FROM #ClientToDBFL  Group BY �����������,�������,�����������');
  Form1.ADOQueryTemp.Open;
  SG1.RowCount:=Form1.ADOQueryTemp.RecordCount+2;
  //Inc(XX);
  Form1.Clear_StringGrid(SG1);
  SG1.Cells[0,0]:='�\�';
  SG1.Cells[1,0]:='�������';
  SG1.Cells[2,0]:= '�����������';
  SG1.Cells[3,0]:='���';
  SG1.Cells[4,0]:='�����������';
  for J := 0 to Form1.ADOQueryTemp.RecordCount - 1 do
  begin
      SG1.Cells[0,J+1]:=IntToStr(J+1);
      SG1.Cells[1,J+1]:= Form1.ADOQueryTemp.FieldByName('�������').AsString;
      SG1.Cells[2,J+1]:= Form1.ADOQueryTemp.FieldByName('�����������').AsString;
      SG1.Cells[3,J+1]:= Form1.ADOQueryTemp.FieldByName('S').AsString;
      SG1.Cells[4,J+1]:= Form1.ADOQueryTemp.FieldByName('�����������').AsString;
      Form1.ADOQueryTemp.Next;
  end;
end;

procedure TFPoisk.Edit1KeyPress(Sender: TObject; var Key: Char);
begin
  if (Key = #13) then
    Button1Click(nil);
end;

end.
