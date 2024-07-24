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
  XL.Workbooks.Open(D + '\CKlapana\2013\Привода.xlsx');
  XL.ActiveWorkBook.WorkSheets[2].Range['B' + IntToStr(1), 'B' + IntToStr(1)] := Dat2;
  XL.ActiveWorkBook.Sheets.Item[2].Activate;}
  //OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO
  ZAk := '';
  SS := '((Специф.Элемент LIKE ' + #39 + '%'+Edit1.Text+'%'+#39+') or (Специф.Обозначение LIKE ' + #39 +
   '%'+Edit1.Text+'%'+#39+') ) ';
  if not Conn_Klap.mkQuerySelect(
  'Select Специф.ВидЭлемента, Специф.Элемент,Специф.Обозначение,Sum(Специф.Количество*Запуск.[Кол во запущенных] ) As S ' +
  'FROM %s INNER JOIN Запуск ON Специф.IdГП = Запуск.IdГП  Where (Запуск.[Планирование] %s' +
  ') AND %s Group BY Специф.ВидЭлемента,Специф.Элемент,Специф.Обозначение', ['Специф', ss1, SS]) then
    exit;

  XX := 5;
        //Создание временной таблицы  idgp,zakaz,Nom
  Form1.SetConnectTemp(Form1.ADOQueryTemp);
    //Создание временной таблицы
  Form1.ADOQueryTemp := TADOQuery.Create(nil);
  Form1.ADOQueryTemp.Connection := Form1.ADOConnectionTemp;

  Form1.ADOQueryTemp.SQL.Text := 'CREATE TABLE #ClientToDBFL (ВидЭлемента nvarchar(350),Элемент nvarchar(350),' + 'Обозначение nvarchar(350),Kol float )';
  try
    Form1.ADOQueryTemp.ExecSQL;
  except
    Err := '1222';
  end;
  for J := 0 to Conn_Klap.AQuery.RecordCount - 1 do
  begin
        //SetConnectTemp(ADOQueryTemp);
    Form1.ADOQueryTemp.SQL.Text := 'INSERT INTO #ClientToDBFL ' + '(ВидЭлемента nvarchar(350),Элемент,Обозначение,Kol) Values (' + #39 + Conn_Klap.AQuery.FieldByName('ВидЭлемента').AsString + #39 + ',' + #39 + Conn_Klap.AQuery.FieldByName('Элемент').AsString + #39 + ',' + #39 + Conn_Klap.AQuery.FieldByName('Обозначение').AsString + #39 + ',' + #39 + Conn_Klap.AQuery.FieldByName('S').AsString + #39 + ')';
    try
     Form1. ADOQueryTemp.ExecSQL;
    except
      Err := '1222';
    end;

    Conn_Klap.AQuery.Next;
  end;
   //ppppppppppppppppppppppppppppppppppppppppppppppp
  //SS := 'СпецифВозд.Элемент LIKE ' + #39 + 'Ось%' + #39;
 // SS := '(СпецифВозд.Элемент LIKE ' + #39 + 'Ось%'+#39+') ';
 SS := '((СпецифВозд.Элемент LIKE ' + #39 + '%'+Edit1.Text+'%'+#39+') or (СпецифВозд.Обозначение LIKE ' + #39 +
 '%'+Edit1.Text+'%'+#39+') ) ';
  if not Conn_Klap.mkQuerySelect('Select СпецифВозд.ВидЭлемента,СпецифВозд.Элемент,СпецифВозд.Обозначение,Sum(СпецифВозд.Количество*ЗапускВозд.[Кол во запущенных] ) As S ' +
  'FROM %s INNER JOIN ЗапускВозд ON СпецифВозд.IdГП = ЗапускВозд.IdГП  Where (ЗапускВозд.[Планирование] %s ' +
  ') AND (%s)  Group BY СпецифВозд.ВидЭлемента,СпецифВозд.Элемент,СпецифВозд.Обозначение', ['СпецифВозд', ss1, SS]) then
    exit;

  Inc(XX);
  for J := 0 to Conn_Klap.AQuery.RecordCount - 1 do
  begin
    Form1.ADOQueryTemp.SQL.Text := 'INSERT INTO #ClientToDBFL ' + '(ВидЭлемента,Элемент,Обозначение,Kol) Values ('
     + #39 + Conn_Klap.AQuery.FieldByName('ВидЭлемента').AsString + #39 + ','+ #39
      + Conn_Klap.AQuery.FieldByName('Элемент').AsString + #39 + ',' + #39
      + Conn_Klap.AQuery.FieldByName('Обозначение').AsString + #39 + ','
      + #39 + Conn_Klap.AQuery.FieldByName('S').AsString + #39 + ')';
    try
      Form1.ADOQueryTemp.ExecSQL;
    except
      Err := '1222';
    end;
    Conn_Klap.AQuery.Next;
  end;
  Form1.ADOQueryTemp.SQL.Text := ('Select ВидЭлемента,Элемент,Обозначение,Sum(Kol ) As S ' +
  'FROM #ClientToDBFL  Group BY ВидЭлемента,Элемент,Обозначение');
  Form1.ADOQueryTemp.Open;
  SG1.RowCount:=Form1.ADOQueryTemp.RecordCount+2;
  //Inc(XX);
  Form1.Clear_StringGrid(SG1);
  SG1.Cells[0,0]:='Н\п';
  SG1.Cells[1,0]:='Элемент';
  SG1.Cells[2,0]:= 'Обозначение';
  SG1.Cells[3,0]:='Кол';
  SG1.Cells[4,0]:='ВидЭлемента';
  for J := 0 to Form1.ADOQueryTemp.RecordCount - 1 do
  begin
      SG1.Cells[0,J+1]:=IntToStr(J+1);
      SG1.Cells[1,J+1]:= Form1.ADOQueryTemp.FieldByName('Элемент').AsString;
      SG1.Cells[2,J+1]:= Form1.ADOQueryTemp.FieldByName('Обозначение').AsString;
      SG1.Cells[3,J+1]:= Form1.ADOQueryTemp.FieldByName('S').AsString;
      SG1.Cells[4,J+1]:= Form1.ADOQueryTemp.FieldByName('ВидЭлемента').AsString;
      Form1.ADOQueryTemp.Next;
  end;
end;

procedure TFPoisk.Edit1KeyPress(Sender: TObject; var Key: Char);
begin
  if (Key = #13) then
    Button1Click(nil);
end;

end.
