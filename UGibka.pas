unit UGibka;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Grids, Vcl.ComCtrls,
  Vcl.ExtCtrls, Vcl.Menus;

type
  TFGibka = class(TForm)
    pnl1: TPanel;
    pnl2: TPanel;
    pgc1: TPageControl;
    ts1: TTabSheet;
    ts2: TTabSheet;
    StringGrid1: TStringGrid;
    pnl3: TPanel;
    Button1: TButton;
    Button2: TButton;
    Label1: TLabel;
    Label2: TLabel;
    StringGrid2: TStringGrid;
    pm1: TPopupMenu;
    N1: TMenuItem;
    ComboBox1: TComboBox;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure StringGrid2SelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure StringGrid1SelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure StringGrid2KeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FGibka: TFGibka;

implementation

uses
  Main;

{$R *.dfm}

procedure TFGibka.Button1Click(Sender: TObject);
var I:Integer;
begin
        Form1.Clear_StringGrid(StringGrid1);
        Form1.Clear_StringGrid(StringGrid2);
        StringGrid1.RowCount:=2;
        StringGrid2.RowCount:=2;
        StringGrid1.Cells[0,0]:='Норма';
        StringGrid1.Cells[1,0]:='Наладка';
        StringGrid1.Cells[2,0]:='Длина';
        StringGrid1.Cells[3,0]:='Толщина';
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select  * from %s'  ,
                ['Гибка']) then
                exit;
        StringGrid1.RowCount:=  Form1.ADOQuery1.RecordCount+1;
        for I := 0 to Form1.ADOQuery1.RecordCount-1   do
        begin
                StringGrid1.Cells[0,I+1]:=Form1.ADOQuery1.FieldByName('Норма').AsString;
                StringGrid1.Cells[1,I+1]:=Form1.ADOQuery1.FieldByName('Наладка').AsString;
                StringGrid1.Cells[2,I+1]:=Form1.ADOQuery1.FieldByName('Длина').AsString;
                StringGrid1.Cells[3,I+1]:=Form1.ADOQuery1.FieldByName('Толщина').AsString;
                Form1.ADOQuery1.Next;
        end;
        ///
        ///

        StringGrid2.Cells[0,0]:='Наладка';
        StringGrid2.Cells[1,0]:='Номер';
        StringGrid2.Cells[2,0]:='Длина';
        StringGrid2.Cells[3,0]:='Ширина';
        StringGrid2.Cells[4,0]:='Толщина';
        StringGrid2.Cells[5,0]:='Префикс';


        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select  * from %s'  ,
                ['ГибкаН']) then
                exit;
        StringGrid2.RowCount:=  Form1.ADOQuery1.RecordCount+1;
        for I := 0 to Form1.ADOQuery1.RecordCount-1   do
        begin
                StringGrid2.Cells[0,I+1]:=Form1.ADOQuery1.FieldByName('Наладка').AsString;
                StringGrid2.Cells[1,I+1]:=Form1.ADOQuery1.FieldByName('Номер').AsString;

                StringGrid2.Cells[2,I+1]:=Form1.ADOQuery1.FieldByName('Длина').AsString;
                StringGrid2.Cells[3,I+1]:=Form1.ADOQuery1.FieldByName('Ширина').AsString;
                StringGrid2.Cells[4,I+1]:=Form1.ADOQuery1.FieldByName('Толщина').AsString;
                StringGrid2.Cells[5,I+1]:=Form1.ADOQuery1.FieldByName('Префикс').AsString;
                Form1.ADOQuery1.Next;
        end;
end;

procedure TFGibka.Button2Click(Sender: TObject);
var N,N1,Dl,sh,t:string;
I:Integer;
begin
       for I := 0 to StringGrid1.RowCount-1 do
         begin
                N:= StringReplace(StringGrid1.Cells[0,I+1],',','.',[rfReplaceAll]);
                N1:=StringReplace(StringGrid1.Cells[1,I+1],',','.',[rfReplaceAll]);

                Dl:= StringReplace(StringGrid1.Cells[2,I+1],',','.',[rfReplaceAll]);
                t:=StringReplace(StringGrid1.Cells[3,I+1],',','.',[rfReplaceAll]);

                if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [Норма]='+#39+N+#39 +
                ',[Наладка]='+#39+N1+#39 +
                ' WHERE ([Длина]=' + #39 + Dl + #39 +
                ') AND ([Толщина]=' + #39 + t + #39 + ')', ['Гибка']) then
                Exit;
         end;
         //
         for I := 0 to StringGrid2.RowCount-1 do
         begin
                N1:=StringReplace(StringGrid2.Cells[0,I+1],',','.',[rfReplaceAll]);

                t:=StringReplace(StringGrid2.Cells[4,I+1],',','.',[rfReplaceAll]);

                Dl:= StringReplace(StringGrid2.Cells[2,I+1],',','.',[rfReplaceAll]);
                Sh:=StringReplace(StringGrid2.Cells[3,I+1],',','.',[rfReplaceAll]);
                if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [Наладка]='+#39+N1+#39 +
                ' WHERE ([Длина]=' + #39 + Dl + #39 +
                ') AND ([Ширина]=' + #39 + Sh+ #39 + ')'+
                ' AND ([Толщина]=' + #39 + t + #39 + ')'+
                ' AND ([Префикс]=' + #39 + StringGrid2.Cells[5,I+1] + #39 + ')'+
                ' AND ([Номер]=' + #39 + StringGrid2.Cells[1,I+1] + #39 + ')', ['ГибкаН']) then
                Exit;
         end;
end;

procedure TFGibka.N1Click(Sender: TObject);
begin
      StringGrid2.RowCount:=StringGrid2.RowCount+1;
      StringGrid2.Options:=StringGrid2.Options+[goEditing];
end;

procedure TFGibka.StringGrid1SelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
begin
  if (StringGrid1.Col=0) or (StringGrid1.Col=1) then
      StringGrid1.Options:=StringGrid1.Options+[goEditing]
  else
      StringGrid1.Options:=StringGrid1.Options-[goEditing];
end;

procedure TFGibka.StringGrid2KeyPress(Sender: TObject; var Key: Char);
var S:string;
L,i:Integer;
R:TRect;
begin
  {  if StringGrid2.Col=5 then
    begin
      S:=StringGrid2.Cells[StringGrid2.Col,StringGrid2.Row];
       L:=Length(S);
       if Key=#8 then
         Delete(S,L,1)
       else
       S:=S+Key;
       //StringGrid2.Cells[StringGrid2.Col,StringGrid2.Row]
       if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select  * from %s Where Обозначение LIKE '+#39+'%s'+#39  ,
                ['СпецифОбщая',S+'%']) then
                exit;
      ComboBox1.Visible:=False;
      ComboBox1.Text:='';
      if ((Form1.ADOQuery1.RecordCount<>0) ) then   //КЦКП
      begin
        {Ширина и положение ComboBox должно соответствовать ячейке StringGrid}
       { R := StringGrid2.CellRect(StringGrid2.Col, StringGrid2.Row);
        R.Left := R.Left + StringGrid2.Left;
        R.Right := R.Right + StringGrid2.Left;
        R.Top := R.Top + StringGrid2.Top;
        R.Bottom := R.Bottom + StringGrid2.Top;
        ComboBox1.Left := R.Left + 1;
        ComboBox1.Top := R.Top + 25;
        ComboBox1.Width :=200;// (R.Right + 1) - R.Left;
        ComboBox1.Height := (R.Bottom + 1) - R.Top; {Покажем combobox}
        {ComboBox1.Items.Clear;
        ComboBox1.Visible := True;
        for I := 0 to Form1.ADOQuery1.RecordCount-1 do
        begin
           ComboBox1.Items.Add(Form1.ADOQuery1.FieldByName('Обозначение').AsString);
           Form1.ADOQuery1.Next;
        end;
        ComboBox1.DropDownCount := 100;
        ComboBox1.DroppedDown:=True;
      end;
    end;  }
end;

procedure TFGibka.StringGrid2SelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
begin

    if (StringGrid2.Col=0)  then
      StringGrid2.Options:=StringGrid2.Options+[goEditing]
  else
      StringGrid2.Options:=StringGrid2.Options-[goEditing];
end;

end.
