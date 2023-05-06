unit UNewNaklSTAM;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, AdvObj, BaseGrid, AdvGrid, ExtCtrls, StdCtrls,ComObj;

type
  TfNewNaklSTAM = class(TForm)
    pnl1: TPanel;
    SG211: TAdvStringGrid;
    SG1: TAdvStringGrid;
    btn1: TButton;
    Edit1: TEdit;
    Label1: TLabel;
    btn2: TButton;
    procedure FormShow(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure SG211DblClick(Sender: TObject);
    procedure btn2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fNewNaklSTAM: TfNewNaklSTAM;

implementation

uses Main;

{$R *.dfm}

procedure TfNewNaklSTAM.FormShow(Sender: TObject);
var
        I, Kol, J, y, Res, Nom1, Nom2,Summa: Integer;
        Nom_Poz, Zak, IDGP, Kol_Zap, Kol_Raz: string;
begin

        //Button1.Enabled := False;
        Nom2:=-1;
        Summa := 0;
        SG1.Enabled := False;
        SG211.ColCount:=7;
        SG211.Cells[0, 0] := '�����';
        SG211.Cells[1, 0] := '��������� ����';
        SG211.Cells[2, 0] := '��� �� ����������';
        SG211.Cells[3, 0] := '�����';
        SG211.Cells[4, 0] := 'ID';
        SG211.Cells[5, 0] := 'IDgp';
        SG211.Cells[6, 0] := '�\�';
        SG211.ColWidths[5] := 0;
        SG211.ColWidths[4] := 0;
        //AND ([�������� �����] IS NULL)  ( [�����]<>0)  AND
        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                'Select * from %s Where  ([������] IS NULL)  AND ( [�] IS NULL) Order By �����',
                ['����������']) then
                exit;

        Y := 1;
        j := 1;
        if  Form1.ADOQuery2.RecordCount=0 then
                SG211.RowCount :=2
        Else
                SG211.RowCount := Form1.ADOQuery2.RecordCount + 1;

        SG211.Row:= SG211.RowCount-2;
        if Form1.ADOQuery2.RecordCount <> 0 then
        begin
                for I := 0 to Form1.ADOQuery2.RecordCount do
                begin
                        Nom1 := Form1.ADOQuery2.FieldByName('�����').AsInteger;
                        if j > 2 then
                                Nom2 := StrToInt(SG211.Cells[3, j - 1]);
                        if Nom1 = Nom2 then
                        begin
                                Form1.ADOQuery2.Next;
                                Continue;
                        end;
                        SG211.Cells[0, j] :=
                                Form1.ADOQuery2.FieldByName('�����').AsString;
                        SG211.Cells[1, j] :=
                                Form1.ADOQuery2.FieldByName('��������� ���� ����������').AsString;
                        SG211.Cells[2, j] :=
                                Form1.ADOQuery2.FieldByName('��� �� ����������').AsString;
                        SG211.Cells[3, j] :=
                                Form1.ADOQuery2.FieldByName('�����').AsString;
                        SG211.Cells[4, j] :=
                                Form1.ADOQuery2.FieldByName('ID').AsString;
                        SG211.Cells[5, j] :=
                                Form1.ADOQuery2.FieldByName('ID��').AsString;
                        SG211.Cells[6, j] :=
                                Form1.ADOQuery2.FieldByName('�\� ������').AsString;
                        Inc(y);
                        Inc(j);
                        Form1.ADOQuery2.Next;
                end;
        end;
        //SG211.RowCount := Y + 2;
end;

procedure TfNewNaklSTAM.btn1Click(Sender: TObject);
Var
Briket,Tip,Tip1,Razmer:string;
Pos1,L,Res:Integer;
begin
        Briket:=Edit1.Text;
        Res:=Pos('-',Briket);
        Tip:=Copy(Briket,1,Res);
        Delete(Briket,1,Res);
        Res:=Pos('-',Briket);
        Tip:=Tip+Copy(Briket,1,Res-1);

    //������ ���-137.xlsx
                        //������ ���-137-�
                        //���� ���� SG1.Cells[6, I+1] := Briket; //���-137-�-0 �������� �����
        Briket:=Edit1.Text;
        Res:=Pos(' ',Briket);
        Delete(Briket,1,Res);
        Label1.Caption:=Tip;


end;

procedure TfNewNaklSTAM.SG211DblClick(Sender: TObject);
var
        I, Kol, Y, j, Pos1, x: Integer;
        Tip,Razmer,Nom_Poz, Zak, IDGP, Briket, Briket1, Briket2, Br, Br1, Br2, Kol_Zap,Nam_S:
        string;
begin
        SG1.Enabled:=True;
        //Clear_StringGrid1(SG1);
        SG1.Clear;
        SG1.ColCount := 12;
        SG1.Cells[0, 0] := '�����';
        SG1.Cells[1, 0] := '��� �� ����������';
        SG1.Cells[2, 0] := '�����';
        SG1.Cells[3, 0] := 'ID��';
        SG1.Cells[4, 0] := '�������';
        SG1.Cells[5, 0] := '���';
        SG1.Cells[6, 0] := '������';
        SG1.Cells[7, 0] := '������';
        SG1.Cells[8, 0] := '������';
        SG1.Cells[9, 0] := '������';
        SG1.Cells[10, 0] := '������';
        SG1.Cells[11, 0] := '�\�';
        Label1.Caption:=SG211.Cells[3,SG211.Row ];

        //  ([�����]= '+#39+Zak+#39+')AND
        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                'Select * from %s Where   (�����='
                + #39 + SG211.Cells[3,SG211.Row ] + #39 + ') Order By ����� ', ['����������']) then
                exit;
        SG1.RowCount:=1;
        if Form1.ADOQuery2.RecordCount <> 0 then
        begin
                SG1.RowCount := Form1.ADOQuery2.RecordCount + 2;
                SG1.FixedRows:=1;
                for i := 0 to Form1.ADOQuery2.RecordCount - 1 do
                begin
                        SG1.Cells[0, i + 1] :=
                                Form1.ADOQuery2.FieldByName('�����').AsString;
                        SG1.Cells[1, i + 1] :=
                                Form1.ADOQuery2.FieldByName('��� �� ����������').AsString;
                        SG1.Cells[2, i + 1] :=
                                Form1.ADOQuery2.FieldByName('�����').AsString;
                        SG1.Cells[3, i + 1] :=
                                Form1.ADOQuery2.FieldByName('ID��').AsString;
                        SG1.Cells[4, i + 1] :=
                                Form1.ADOQuery2.FieldByName('�������').AsString;
                        Nam_S:= Form1.ADOQuery2.FieldByName('�������').AsString;
                        Briket:= StringReplace(Nam_S, '���� ', '����',[rfReplaceAll, rfIgnoreCase]);
                        //������ ��������� ����610-35-�
                        Pos1:=Pos('����',Briket);
                        If Pos1<>0 then
                                Delete(Briket,1,Pos1-1);

                        Pos1:=Pos('-',Briket);
                        If Pos1<>0 then
                        begin
                                Tip:=Copy(Briket,1,Pos1-1);
                                Insert('-',Tip,5);
                        end;
                        SG1.Cells[5, I+1] := Tip;// ����-610 �������� �����
                        SG1.Cells[6, I+1] := Briket; //����610-35-� �������� �����
                        SG1.Cells[7, i + 1] :=
                                Form1.ADOQuery2.FieldByName('��������').AsString;
                        SG1.Cells[8, i + 1] :=
                                Form1.ADOQuery2.FieldByName('��������').AsString;
                        SG1.Cells[9, i + 1] :=
                                Form1.ADOQuery2.FieldByName('������').AsString;
                        SG1.Cells[10, i + 1] :=
                                Form1.ADOQuery2.FieldByName('������').AsString;
                        SG1.Cells[11, i + 1] :=
                                Form1.ADOQuery2.FieldByName('�\� ������').AsString;
                        Form1.ADOQuery2.Next;
                end;
        end;
end;

procedure TfNewNaklSTAM.btn2Click(Sender: TObject);
var I,J,Res,L,Res1,F:Integer;
XL:Variant;
God,mes,Dir,Nam,Nam_S,Zav_Nom,Briket,Tip,Naim1,S:string;
begin
        God := FormatDateTime('yyyy', Now);
        Mes := FormatDateTime('mmmm', Now);
        Dir :=Form1.Put_KTO+'\CKlapana\�������������\';
        CreateDir(Dir);

        Dir :=Form1.Put_KTO+'\CKlapana\�������������\' + God;
        CreateDir(Dir);

        Dir := Form1.Put_KTO+'\CKlapana\�������������\' + God + '\' + Mes+ '\';
        CreateDir(Dir);

        XL := CreateOleObject('Excel.Application');
        //CopyFile(PAnsiChar('\\192.168.7.1\Kto\CKlapana\2013\���� (����������������)\���������.xlsx'), PAnsiChar(Dir +'\� '+Label1.Caption+'.xlsx'), False);
        for i:=0 To SG1.RowCount-1 do
        begin
                Zav_Nom:=SG1.Cells[7, I+1]+'..'+SG1.Cells[8, I+1];
                Nam_S:=SG1.Cells[6, I+1];
                if Nam_S='' then
                  Continue;
                S:=SG1.Cells[5, I+1];
                Res:=Pos('����',Nam_S);
                if Res<>0 then //STAM
                begin
                        L:=Length(Nam_S);
                        Res1:=AnsiCompareStr('�',Nam_S[L]);
                        if Res1=0 then
                                Nam_S:=Nam_S+'-1';

                       // Res1:=AnsiCompareStr('�',Nam_S[L]);
                        //if Res1=0 then
                        S:= Nam_S ;
                        Nam_S:=StringReplace(S, '�1', '�-1',[rfReplaceAll, rfIgnoreCase]);
                        Naim1:= StringReplace(Nam_S, '���� ', '����',[rfReplaceAll, rfIgnoreCase]);
                        XL.Workbooks.Open(Form1.Put_KTO+'\CKlapana\2013\���� (����������������)\'+SG1.Cells[5, I+1]+'.xlsx');
                        //XL.Application.EnableEvents := false;
                        //���� ���� SG1.Cells[6, I+1] := Briket; //����610-35-� �������� �����
                        F:=0;
                        For j:=1 to 260 do
                        Begin
                               Nam:=Trim(XL.ActiveWorkBook.WorkSheets[j].Name);


                                Res:=AnsiCompareStr(Nam,Naim1);
                                if Res=0 then
                                Begin
                                        F:=1;
                                        XL.ActiveWorkBook.WorkSheets[j].Cells[1, 1] := '���������������� �' +SG1.Cells[2, i + 1];
                                        XL.ActiveWorkBook.WorkSheets[j].Cells[1, 6] := Zav_Nom;
                                        XL.ActiveWorkBook.WorkSheets[j].Cells[1, 3] := SG1.Cells[1, i + 1];
                                        XL.ActiveWorkBook.WorkSheets[j].Cells[1, 4] := SG1.Cells[0, i + 1];
                                        //XL.ActiveWorkBook.WorkSheets[j].Cells[8, 1] :='���������������� �' +SG1.Cells[2, i + 1];
                                        XL.ActiveWorkBook.WorkSheets[j].SaveAs(Dir+'\� '+Label1.Caption+' ('+Naim1+').xlsx');
                                        XL.Visible := True;
                                        XL.ActiveWorkBook.WorkSheets[j].Select;
                                        //XL.ActiveWorkBook.WorkSheets[j].Range('D7').Select;
                                        Break;
                                End;
                        end;
                        if F=0 then
                        Begin
                                XL.ActiveWorkbook.Close;
                                XL := UnAssigned;
                                MessageDlg('�� ������� ����� ���� � ��������� '+Naim1, mtError,
                                        [mbOk], 0);
                                {XL.ActiveWorkbook.Close;
                                XL.Application.Quit;
                                XL.Quit; }
                                Exit;
                        end;
                end;
                //
                Res:=Pos('������',Nam_S);
                if Res <>0 Then
                begin
                        Briket:=Nam_S;
                        Res:=Pos('-',Briket);
                        Tip:=Copy(Briket,1,Res);
                        Delete(Briket,1,Res);
                        Res:=Pos('-',Briket);
                        Tip:=Tip+Copy(Briket,1,Res-1);  //������ ���-137.xlsx

                        Briket:=Nam_S;
                        Res:=Pos(' ',Briket);
                        Delete(Briket,1,Res); //���-137-�-0 �������� �����

                        XL.Workbooks.Open(Form1.Put_KTO+'\CKlapana\2013\�������\'+Tip+'.xlsx');
                        XL.Application.EnableEvents := false;    //������ ���-137.xlsx
                        //������� ���-137-�
                        // Briket; //���-137-�-0 �������� �����
                        F:=0;
                        For j:=1 to 30 do
                        Begin

                                Nam:=XL.ActiveWorkBook.WorkSheets[j].Name;


                                Res:=AnsiCompareStr(Nam,Briket);
                                If Res<>0 then
                                Continue;
                                if Res=0 then
                                Begin
                                        F:=1;
                                        XL.ActiveWorkBook.WorkSheets[j].Cells[1, 1] := '���������������� �' +SG1.Cells[2, i + 1];
                                        XL.ActiveWorkBook.WorkSheets[j].Cells[1, 6] := Zav_Nom;
                                        XL.ActiveWorkBook.WorkSheets[j].Cells[1, 3] := SG1.Cells[1, i + 1];
                                        XL.ActiveWorkBook.WorkSheets[j].Cells[1, 4] := SG1.Cells[0, i + 1];
                                        XL.ActiveWorkBook.WorkSheets[j].SaveAs(Dir+'\� '+Label1.Caption+' ('+Briket+').xlsx');
                                        XL.Visible := True;
                                        XL.ActiveWorkBook.WorkSheets[j].Select;
                                        //XL.ActiveWorkBook.WorkSheets[j].Range('D7').Select;
                                        Break;
                                End;

                        end;
                        if F=0 then
                        Begin
                                        XL.ActiveWorkbook.Close;
                                        XL := UnAssigned;
                                        MessageDlg('�� ������� ����� ���� � ��������� '+Nam_S, mtError,
                                        [mbOk], 0);
                                        Exit;
                        end;
                end;
        end;
        //XL.Visible:=True;
        XL := UnAssigned;

end;

end.
