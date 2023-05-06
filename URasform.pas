unit URasform;

interface

uses
        Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
        Forms,
        Dialogs, Grids, ExtCtrls, StdCtrls, ComCtrls, 
  AdvObj, BaseGrid, AdvGrid, DB, ADODB, AdvGridRTF;

type
        TFRasform = class(TForm)
                Panel1: TPanel;
                Panel2: TPanel;
                Panel3: TPanel;
                Button1: TButton;
                Button2: TButton;
                Label3: TLabel;
                Label4: TLabel;
                Memo1: TMemo;
                SG1: TStringGrid;
    SG2: TStringGrid;
    ComboBox1: TComboBox;
    Label1: TLabel;
    Button3: TButton;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    AdvGridRTFIO1: TAdvGridRTFIO;
    S1: TAdvStringGrid;
    S2: TAdvStringGrid;
    Button4: TButton;
    Button5: TButton;
                procedure FormShow(Sender: TObject);
                procedure Button2Click(Sender: TObject);
                procedure Button1Click(Sender: TObject);
                procedure SG1SelectCell(Sender: TObject; ACol, ARow: Integer;
                        var CanSelect: Boolean);
    procedure ComboBox1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
        private
                { Private declarations }
                Enabled_Ok1: Integer;
                Enabled_Ok2: Integer;
        public
                { Public declarations }
        end;

var
        FRasform: TFRasform;

implementation

uses Main;

{$R *.dfm}

procedure TFRasform.FormShow(Sender: TObject);
var
        I, Kol, Kol_Zap, Max: Integer;
        S,MSSQL_CONN_STR: string;
        Svar, Sbor, Svar_o, Sbor_o: Double;
begin
MSSQL_CONN_STR :=Form1.MSSQL_CONN_STR_Glob;
               { 'Provider=SQLOLEDB.1;Packet Size = 4096;Password=111;' +
                'Persist Security Info=True;User Id=testuser;' +
                'Initial Catalog=MIASSCEH'+
                ';Data Source=DINAMO\SQLEXPRESS';}
        ADOConnection1.ConnectionString:=MSSQL_CONN_STR;
        ADOConnection1.Connected:=True;
        Form1.Clear_StringGrid(SG1);
        Form1.Clear_StringGrid(SG2);
        ComboBox1.Clear;
        S := ExtractFileDir(ParamStr(0));
        SG1.ColCount := 8;
        SG1.Cells[0, 0] := '�';
        SG1.Cells[1, 0] := '����';
        SG1.Cells[2, 0] := FN_ZAK;
        SG1.Cells[3, 0] := '�����';
        SG1.Cells[4, 0] := FN_KOL;
        SG1.Cells[5, 0] := '���-�� ����������';
        //SG1.Cells[6, 0] := '��� ���';
        SG1.Cells[6, 0] := '����� ���';
        //SG1.Cells[8, 0] := '����';
        SG1.Cells[7, 0] := '�������';
        SG1.ColWidths[6] := 0;
        SG1.ColWidths[7] := 0;
        //SG1.ColWidths[8] := 0;
        //SG1.ColWidths[9] := 0;
        Form1.Clear_StringGrid(SG1);
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select * from %s Where ([��������� ������] IS NULL) AND (������ IS NULL) AND ([�] IS NULL) ',
                ['������']) then
                exit;
        SG1.RowCount := Form1.ADOQuery1.RecordCount + 1;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                Kol := Form1.ADOQuery1.FieldByName(FN_KOL).AsInteger;
                Kol_Zap := Form1.ADOQuery1.FieldByName(FN_KOL_ZAP).AsInteger;
                SG1.Cells[0, i + 1] := IntToStr(I + 1);
                //SG1.Cells[1,i+1]:=Form1.ADOQuery1.FieldByName(FN_PLAN_DATA).AsString;
                SG1.Cells[2, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_ZAK).AsString;
                SG1.Cells[3, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_NOM).AsString;
                SG1.Cells[4, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_KOL).AsString;
                SG1.Cells[5, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_KOL_ZAP).AsString;
                //SG1.Cells[5,i+1]:=Form1.ADOQuery1.FieldByName(FN_KOL_RAZ).AsString;
                SG1.Cells[6,i+1]:=Form1.ADOQuery1.FieldByName(FN_NOMER_RAZ).AsString;
                SG1.Cells[1, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_DAT).AsString;
                SG1.Cells[7, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_NAM).AsString;
                Form1.ADOQuery1.Next;

        end;
                if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select DISTINCT (�����) from %s Where ([��������� ������] IS NULL) AND (������ IS NULL) AND ([�] IS NULL) ',
                ['������']) then
                exit;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                ComboBox1.Items.Add(Form1.ADOQuery1.FieldByName(FN_NOM).AsString);
                Form1.ADOQuery1.Next;
        end;
end;

procedure TFRasform.Button2Click(Sender: TObject);
begin
        Close;
end;

procedure TFRasform.Button1Click(Sender: TObject);
var
        I, J, ID, Pos1,Pos2, Kol_Zap, Kol_Zap_R, Kol, Rec_Count, Kol_Ob, L, L1, Res:
        Integer;
        Svar, Sbor, Svar_o, Sbor_o: Double;
        Kol_Raz, Nomer_Raz,Nam,Zak: string;
        C:Char;
begin
        if ComboBox1.Text<>'' Then
        Begin
        if not Form1.mkQueryDelete(Form1.ADOQuery1, 'DELETE FROM %s Where (�����='
                + #39+ ComboBox1.Text + #39+' AND �������='+ #39+ '������ ���-4-01-800*500-1*�-���24-��-�-���'+ #39+' AND �����='+ #39+ '154056'+ #39,
                ['������']) then
                Exit;
        for i := 1 to Sg2.RowCount - 1 do
        begin
                Kol_Zap:=StrToInt(SG2.Cells[4,i]);
                Nomer_Raz:=SG2.Cells[5,i];
                Nam:=SG2.Cells[7,i];
                Zak:=SG2.Cells[1,i];
                Pos1:=Pos(',',Nomer_Raz);
                if Pos1=0 Then
                Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [�����]=' + #39
                        + '0' + #39 +',[����� ���]=' + #39+ '0' + #39 +',[��� �� ����������]=' + #39+ '0' + #39 +
                        ' WHERE (�������=' + #39 + Nam +
                        #39+') AND (�����='+ #39 + Zak +
                        #39+')', ['Klapana']) then
                        Exit;
                end
                Else
                Begin
                        Pos1:=Pos(ComboBox1.Text,Nomer_Raz);
                        L:=Length(ComboBox1.Text);
                        Delete(Nomer_Raz,Pos1,L+1) ;
                        L:=Length(Nomer_Raz);
                        C:=Nomer_Raz[l];
                        if C =',' Then
                        Delete(Nomer_Raz,L,1);
                        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                        'Select * from %s Where (�������=' + #39 + Nam +
                        #39+') AND (�����='+ #39 + Zak +
                        #39+')',
                        ['Klapana']) then
                        exit;
                        Kol_Zap_r:=StrToInt(Form1.ADOQuery2.FieldByName(FN_KOL_ZAP).AsString);
                        if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET {[����� ���]=' + #39
                        + Nomer_Raz + #39 +'),[��� �� ����������]='+#39+ IntToStr(Kol_Zap_R-Kol_Zap) + #39 +
                        ' WHERE (�������=' + #39 + Nam +
                        #39+') AND (�����='+ #39 + Zak +
                        #39+')', ['Klapana']) then
                        Exit;
                end;


        end;

        end;
        MessageDlg('��������� � ' + ComboBox1.Text + ' ���������������!',
                mtInformation,
                [mbOk], 0);
        Close;
end;

procedure TFRasform.SG1SelectCell(Sender: TObject; ACol, ARow: Integer;
        var CanSelect: Boolean);
begin
        Label4.Caption := SG1.Cells[3, ARow];
end;

procedure TFRasform.ComboBox1Click(Sender: TObject);
Var I:Integer;
begin
        if ComboBox1.Text<>'' Then
        Button1.Enabled:=True;
        SG2.ColCount := 10;
        SG2.Cells[0, 0] := '�';
        SG2.Cells[1, 0] := FN_ZAK;
        SG2.Cells[2, 0] := '�����';
        SG2.Cells[3, 0] := FN_KOL;
        SG2.Cells[4, 0] := '���-�� ����������';
        SG2.Cells[5, 0] := '����� ���';
        SG2.Cells[6, 0] := '����';
        SG2.Cells[7, 0] := '�������';

        Form1.Clear_StringGrid(SG2);
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select * from %s Where ([�����]=' + #39
                        + ComboBox1.Text + #39+') AND (������ IS NULL) AND ([�] IS NULL)',
                ['������']) then
                exit;
        SG2.RowCount := Form1.ADOQuery1.RecordCount + 1;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                //Kol := Form1.ADOQuery1.FieldByName(FN_KOL).AsInteger;
                //Kol_Zap := Form1.ADOQuery1.FieldByName(FN_KOL_ZAP).AsInteger;
                SG2.Cells[0, i + 1] := IntToStr(I + 1);
                //SG1.Cells[1,i+1]:=Form1.ADOQuery1.FieldByName(FN_PLAN_DATA).AsString;
                SG2.Cells[1, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_ZAK).AsString;
                SG2.Cells[2, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_NOM).AsString;
                SG2.Cells[3, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_KOL).AsString;
                SG2.Cells[4, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_KOL_ZAP).AsString;
                SG1.Cells[5,i+1]:=Form1.ADOQuery1.FieldByName(FN_KOL_RAZ).AsString;
                //SG1.Cells[6,i+1]:=Form1.ADOQuery1.FieldByName(FN_NOMER_RAZ).AsString;
                SG2.Cells[6, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_DAT).AsString;
                SG2.Cells[7, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_NAM).AsString;
                Form1.ADOQuery1.Next;

        end;
end;

procedure TFRasform.Button3Click(Sender: TObject);
Var StrDat1,StrDat2,Betwe:String;
I:Integer;
begin
           S1.Cells[0,0]:='�';
           S1.Cells[1,0]:='����';
           S1.Cells[2,0]:='�����';
           S1.Cells[3,0]:='���';
           S1.Cells[4,0]:='�����';
           S1.Cells[5,0]:='��������';
     StrDat1 := FormatDateTime('mm.dd.yyyy', DateTimePicker1.Date);
     StrDat2 := FormatDateTime('mm.dd.yyyy', DateTimePicker2.Date);
     Betwe := 'AND ( ��� BETWEEN ' + #39 + StrDat1 + #39 + ' AND ' +
                        #39 + StrDat2 + #39 + ') ';
     if not Form1.mkQuerySelect(Form1.ADOQuery2,
     'Select * from %s Where (������ IS NULL) AND ([' + FN_KOL_OTK+']=['+FN_KOL_ZAP+'])   '+ Betwe
     , ['������']) then
     exit;
     S1.RowCount:= Form1.ADOQuery2.RecordCount+1;
     For I:=0 To Form1.ADOQuery2.RecordCount-1 Do
     Begin
           S1.Cells[0,I+1]:=IntToStr(I+1);
           S1.Cells[1,I+1]:=Form1.ADOQuery2.FieldByName(FN_DAT).AsString;
           S1.Cells[2,I+1]:=Form1.ADOQuery2.FieldByName(FN_ZAK).AsString;
           S1.Cells[3,I+1]:=Form1.ADOQuery2.FieldByName(FN_KOL).AsString;
           S1.Cells[4,I+1]:=Form1.ADOQuery2.FieldByName(FN_NAM).AsString;
           S1.Cells[5,I+1]:=Form1.ADOQuery2.FieldByName('��������').AsString;
           Form1.ADOQuery2.Next;
     end;
     S1.SortByColumn(0);
     For I:=0 To  S1.RowCount-1 Do
             S1.AutoSizeRow(i,0);
     //ADOQuery1.Close;
     ADOQuery1.SQL.Clear;
     ADOQuery1.SQL.Add('Select * from �� Where ([������]  IS NULL) AND ([���� ������������] BETWEEN  '
     +#39+StrDat1+#39+' And '+#39+StrDat2+#39+')') ;
     ADOQuery1.Open;
     S1.Cells[0,0]:='�';
           S2.Cells[1,0]:='����';
           S2.Cells[2,0]:='�����';
           S2.Cells[3,0]:='� ����� ������';
           S2.Cells[4,0]:='�����';
           S2.Cells[5,0]:='��� �';
     S2.RowCount:= ADOQuery1.RecordCount+1;
     For I:=0 To ADOQuery1.RecordCount-1 Do
     Begin
           S2.Cells[0,I+1]:=IntToStr(I+1);
           S2.Cells[1,I+1]:=ADOQuery1.FieldByName('���� �������').AsString;
           S2.Cells[2,I+1]:=ADOQuery1.FieldByName('�����').AsString;
           S2.Cells[3,I+1]:=ADOQuery1.FieldByName('� ����� ������').AsString;
           S2.Cells[4,I+1]:=ADOQuery1.FieldByName('������������').AsString;
           S2.Cells[5,I+1]:=ADOQuery1.FieldByName('��������� �����').AsString;
           ADOQuery1.Next;
          // S2.Cells[6,I+1]:=ADOQuery1.FieldByName('').AsString;
     end;
     S2.SortByColumn(0);
             For I:=0 To  S2.RowCount-1 Do
             S2.AutoSizeRow(i,0);
end;

procedure TFRasform.Button4Click(Sender: TObject);
begin
        Form1.ExportGridtoExcel2(S1);
end;

procedure TFRasform.Button5Click(Sender: TObject);
begin
        Form1.ExportGridtoExcel2(S2);
end;

end.
