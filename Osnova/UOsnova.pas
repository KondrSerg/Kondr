unit UOsnova;

interface
uses SysUtils, Classes, Variants,Forms,
        ADODB, Controls,Dialogs,AdvGrid,Grids,Windows,
  UConnCeh, UConnKlap,ShellAPI,Math;
        Type
        Osnova_Main = class//(TForm)
        SG2:TStringGrid;
        SG3:TStringGrid;
        procedure Clear_StringGrid(StringGrid: TStringGrid);
        function ConvertDat1(StrDat: string): string;
        function Osnova( var Dat : string ;const tab1 : string;tab2 :string ;
                        Nom :string;fileName:String;Flag_Polz,Cvet:Integer) : Boolean;
                        //Cvet- �� �������� 0-�����, 1-������, 2-�������, 3-�����
        function btn1Click(fileName:String):Boolean;
        function btn11Click(Err2:String):Boolean;
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++
        //���� ���� ���
        //���� ������� ���������= ���� ��������� ������ ������ � ������������=1
        // ��������� ��������� ����������� ������� �� ���� ������ (������������)
        function Osnova2( var Dat : string ;const tab1 : string;tab2 :string ;
                        Nom :string;fileName:String;Flag_Polz:Integer) : Boolean;
        function btn2Click(fileName:String):Boolean;
        function btn22Click(Err2:String):Boolean;
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++
        public
                Column: array[0..550, 0..4] of string;
                StrToCol: array[0..100,0..6] of string;
                Smena: array[0..550, 0..4] of string;
                Temp: array[0..550, 0..6] of string;
                Conn_Ceh:Connect_Miass_Ceh;
                Conn_Klap:Connect_Miass_Klap;
                S1,Nomer_Glob:String;
                Dat_Glob,File_Name:String;
                Flag_Error:Integer;
                Flag_Polzov    //0 ��������   1 ���������
                ,Cvet1:Integer;
                SNCH,Nom_G:String;// ���������� �\� �� ������ ������ ��� ������ ����
                Leg:String;
        End;
implementation

uses
  Main;
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++ Check
//������� �����

procedure Osnova_Main.Clear_StringGrid(StringGrid: TStringGrid);
var
        i, j: integer;
begin
        for i := 0 to StringGrid.ColCount - 1 do
                for j := 1 to StringGrid.RowCount - 2 do
                begin
                        //StringGrid.Objects[i, j] := nil;
                        StringGrid.Cells[I, j] := '';

                end;

end;
function Osnova_Main.ConvertDat1(StrDat: string): string;
var
        G, M, D: string;
        PosM, PosD: integer;
begin
        if StrDat <> '' then
        begin
                PosD := AnsiPos('.', StrDat);
                D := Copy(StrDat, 1, PosD - 1);
                Delete(StrDat, 1, PosD);
                PosM := AnsiPos('.', StrDat);
                M := Copy(StrDat, 1, PosM - 1);
                Delete(StrDat, 1, PosM);
                G := M + '.' + D + '.' + StrDat;
                Result := G;
        end
        else
                Result := '';
end;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++
        //���� ���� ���
        //���� ������� ���������= ���� ��������� ������ ������ � ������������=1
        // ��������� ��������� ����������� ������� �� ���� ������ (������������)
function Osnova_Main.Osnova2( var Dat : string;  const tab1:string;tab2:string;
Nom:string;fileName:String;Flag_Polz:Integer ) : Boolean;
var a,Y,Glob,i,x,Res,T,Res1,L:Integer;
Stanok,Dat_S1,IDGP,Kol_Klap,Izdel,kol_Det,Elem,Obozn,Nomer,BZ,Str1,Str2,Dlit,Kod,
IDKO,Nam_S,ZAKAZ:string;
Dlit_D,Och_D:Double;
Res_Teh,Res_Stan,Memo12_Count,Res_Tab,Polzov,
KOL_RECORD_Tab1,KOL_RECORD_Tab2:Integer;
Memo12,Memo13:TStringList;
Memo12_text,Nam,God,Mes,Dir:String;
h: hwnd;
begin
    Result:=False;
    try
        T:=0;
        Nom_G:=Nom;
        Conn_Ceh := Connect_Miass_Ceh.Create();
        Conn_Klap:=Connect_Miass_Klap.Create();
        Memo12:=TStringList.Create;
        Memo12.Clear;
        Memo13:=TStringList.Create;
        Memo13.Clear;
        Polzov:= Flag_Polz;
        if not Conn_Klap.mkQuerySelect('SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME ='+#39+'�����������'+#39
        , []) then
        exit;
        Dat_Glob:= Dat;
        for I:=0 To Conn_Klap.AQuery.RecordCount-1 do
        begin
          Str1:=Conn_Klap.AQuery.FieldByName('COLUMN_NAME').AsString;
          Str2:=Conn_Klap.AQuery.FieldByName('COLUMN_NAME').AsString;
          Res:=Pos('�\�',str2);
          if Res<>0 then
          begin
               Delete(str2,1,3);
               Column[T,0]:=Str2;
               Column[T,1]:=Str2+' BIT';
               Column[T,2]:=IntToStr(T+7);
               Inc(T);
          end;
          Conn_Klap.AQuery.Next;
        End;
        //++++++++++++++++++++++++++++
        fillchar(Smena,sizeof(Smena),0);//������� �������
        if not Conn_Klap.mkQuerySelect('Select * From �����  ' , []) then //������
        exit;
      For I:=0 to Conn_Klap.AQuery.RecordCount-1 Do
      Begin
        Smena[i,0]:=Conn_Klap.AQuery.FieldByName('�������').AsString;
        Smena[i,1]:=Conn_Klap.AQuery.FieldByName('�����').AsString;
        Smena[i,2]:=Conn_Klap.AQuery.FieldByName('����').AsString;
        Conn_Klap.AQuery.Next;
      End;
          sg2 := TStringGrid.Create(Nil);
          Clear_StringGrid(sg2);
        //sg2.Clear;
        A:=1;
        sg2.ColCount:=11;
        sg2.Cells[0,0]:='�';
        sg2.Cells[1,0]:='�����';
        sg2.Cells[2,0]:='�������';
        sg2.Cells[3,0]:='��� ��';
        sg2.Cells[4,0]:='�����������';
        sg2.Cells[5,0]:='�������';
        sg2.Cells[6,0]:='�����';
        sg2.Cells[7,0]:='��';
        sg2.Cells[8,0]:='Id��';
        sg2.Cells[9,0]:='��������';
        sg2.Cells[10,0]:='Id��';
        for y := 0 to High(Column) do
        begin
                if Column[y,0]<>'' then
                Begin
                        Stanok:=Column[y,0];
                        SG2.ColCount:=SG2.ColCount+4;
                        sg2.Cells[9+A+1,0]:=Stanok;
                        sg2.Cells[9+A+2,0]:=Stanok+'��'; //�����������
                        sg2.Cells[9+A+3,0]:=Stanok+'��';//������������
                        sg2.Cells[9+A+4,0]:='�\�'+Stanok;//�\�
                        //sg2.Cells[9+A+5,0]:=Stanok+'����'; //����
                        A:=A+4;
                end;
        end;
        Glob:=1;//SG2.RowCount;
        //Dat_S1:=FormatDateTime('mm.dd.YYYY',Dtp1.DateTime);
        if Polzov<>4 then
        begin
          if not Conn_Klap.mkQuerySelect('Select * From %s Where (�����=%s) '+
          'AND (������ IS NULL) ' , [tab1,Nom]) then //������
          exit;
          KOL_RECORD_Tab1:= Conn_Klap.AQuery.RecordCount;
        end;
        ///����������� �� ����
        if Polzov=4 then
        begin
          if not Conn_Ceh.mkQuerySelect('Select * From %s Where (�����=%s) '+
          'AND (������ IS NULL) ' , [tab1,Nom]) then //������
          exit;
          KOL_RECORD_Tab1:= Conn_Ceh.AQuery.RecordCount;
        end;
        SG2.RowCount:= 2;
        SG2.FixedRows:=1;
        A:=1;
      For I:=0 to KOL_RECORD_Tab1-1 Do
      Begin
        if Polzov<>4 then
        begin
          IDGP:=Conn_Klap.AQuery.FieldByName('Id��').AsString;
          IDKO:=Conn_Klap.AQuery.FieldByName('Id��').AsString;
          Izdel:=Conn_Klap.AQuery.FieldByName('�������').AsString;
          Kol_Klap:=Conn_Klap.AQuery.FieldByName('��� �� ����������').AsString;
          Nomer:=Conn_Klap.AQuery.FieldByName('�����').AsString;
          Nomer_Glob:=Conn_Klap.AQuery.FieldByName('�����').AsString;
          BZ:=Conn_Klap.AQuery.FieldByName('��').AsString;
          ZAKAZ:=Conn_Klap.AQuery.FieldByName('�����').AsString;
        end;
        if Polzov=4 then
        begin
          IDGP:=Conn_Ceh.AQuery.FieldByName('Id��').AsString;
          IDKO:=Conn_Ceh.AQuery.FieldByName('Id��').AsString;
          Izdel:=Conn_Ceh.AQuery.FieldByName('�������').AsString;
          Kol_Klap:=Conn_Ceh.AQuery.FieldByName('��� ����������').AsString;
          Nomer:=Conn_Ceh.AQuery.FieldByName('�����').AsString;
          Nomer_Glob:=Conn_Ceh.AQuery.FieldByName('�����').AsString;
          BZ:=Conn_Ceh.AQuery.FieldByName('��').AsString;
          ZAKAZ:=Conn_Ceh.AQuery.FieldByName('�����').AsString;
        end;
        if (Tab2='') and (Polzov<>4) Then //Tab2 ������ ������ ������������ ��� � �����������  , ���� �� �������� ������� ����....
        Begin

                 if not Conn_Klap.mkQuerySelect2(
                'Select * from %s Where (�����������='+#39+'������'+#39+') '+
                'AND (([�������]='+#39+'������ ��������� '+Izdel+#39+') OR ([�������]='+#39+'������ '+Izdel+#39+'))'
                 , ['�����������']) then  //������
                 exit;
                 For X:=0 to Conn_Klap.AQuery_Del.RecordCount-1 Do
                Begin
                        Obozn:=Conn_Klap.AQuery_Del.FieldByName('�����������').AsString;
                        Elem:=Conn_Klap.AQuery_Del.FieldByName('�������').AsString;
                        Kol_Det:=Conn_Klap.AQuery_Del.FieldByName('����������').AsString;
                   if  Conn_Klap.AQuery_Del.RecordCount<>0 then
                   Begin
                        SG2.RowCount:=SG2.RowCount+ 1;
                        sg2.Cells[0,Glob+1]:=Kol_Klap;//IntToStr(Glob+1);

                        sg2.Cells[1,Glob+1]:=Conn_Klap.AQuery.FieldByName('�����').AsString;
                        sg2.Cells[2,Glob+1]:=Conn_Klap.AQuery.FieldByName('�������').AsString;

                        sg2.Cells[3,Glob+1]:=IntToStr(StrToInt(Kol_Klap)*StrToInt(Kol_Det));
                        sg2.Cells[4,Glob+1]:=Conn_Klap.AQuery_Del.FieldByName('�����������').AsString;
                        sg2.Cells[5,Glob+1]:=Conn_Klap.AQuery_Del.FieldByName('�������').AsString;
                        sg2.Cells[6,Glob+1]:=Nom;
                        sg2.Cells[7,Glob+1]:=BZ;
                        sg2.Cells[8,Glob+1]:=IDGP;
                        sg2.Cells[9,Glob+1]:=Conn_Klap.AQuery3.FieldByName('��������').AsString;
                        sg2.Cells[10,Glob+1]:=IDKO;
                        Res_Teh:=Pos('rue',sg2.Cells[9,Glob+1]);
                        if (Res_Teh=0)  Then //�������� FALSE
                        begin
                                  Memo12.Add('�������� FALSE ����� '+sg2.Cells[1,Glob+1]+' * '+sg2.Cells[2,Glob+1]+' * ������ '+sg2.Cells[4,Glob+1]);
                                  Memo13.Add(sg2.Cells[4,Glob+1]);
                        end;
                        A:=1;
                        for y := 0 to High(Column) do
                        begin
                                if Column[y,0]<>'' then
                                Begin
                                        Stanok:=Column[y,0]; //�������� ����

                                        sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery_Del.FieldByName(Stanok).AsString;
                                        Res_Stan:=Pos('rue',sg2.Cells[9+(A+1),Glob+1]);
                                        Inc(A);

                                        sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery_Del.FieldByName(Stanok+'��').AsString; //�����������
                                        Och_D:=StrToFloat(sg2.Cells[9+(A+1),Glob+1]);
                                        Inc(A);

                                        Dlit:= '0.1';//Conn_Klap.AQuery_Del.FieldByName(Stanok+'��').AsString;//������������
                                        sg2.Cells[9+(A+1),Glob+1]:='0.1';//Conn_Klap.AQuery_Del.FieldByName(Stanok+'��').AsString;//������������
                                        Dlit_D:=StrToFloat( sg2.Cells[9+(A+1),Glob+1]);//������������
                                        Inc(A);

                                        if (Res_Stan<>0) AND ((Och_D=0) OR (Dlit_D=0))  Then //������ TRUE
                                                                                             //� ����=0 ��� ���� =0 �����
                                        begin
                                                Memo12.Add('������ TRUE � ����=0 ��� ���� =0 ����� '+sg2.Cells[1,Glob+1]+' * '+sg2.Cells[2,Glob+1]+' * ������ '+sg2.Cells[4,Glob+1]);
                                                Memo13.Add(sg2.Cells[4,Glob+1]);
                                        end;

                                        sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery_Del.FieldByName('�\�'+Stanok).AsString;//�\�
                                        Inc(A);
                                        //sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery3.FieldByName(Stanok+'����').AsString; //����
                                        //Inc(A);
                                End;
                        end;
                        Inc(Glob);
                   end;
                   Conn_Klap.AQuery_Del.next;
                end;
        end
        Else
        Begin
           if (Polzov=4) then //���� � ������������ ID ������� 3 �������
           begin
              Res:=Pos('���������',tab2);
             if Res=0 then
             Begin
                //� Tab2 ��� ���������� ����� Id����
               if not Conn_Ceh.mkQuerySelect2(
                'SELECT ����������.Id,������ ,�������, ����������.kol '+
                'FROM ���������� INNER JOIN ���������� ON ����������.IdSb = ����������.Id '+
                'WHERE (����������."Id����"='+Tab2+') AND (����������.IdSb<>'+#39+'0'+#39+') UNION '+

                'SELECT we.Id , ������ ,  �������, we.kol '+
                'FROM ���������� INNER JOIN '+
	              '(SELECT �������.Id,������ , ����������� �������, �������.kol '+
                'from �������) '+
	              'we  ON ����������.IdDet = we.Id '+
                'WHERE (����������.Id����='+Tab2+') AND (����������.IdSb='+#39+'0'+#39+')'
                 , ['����������']) then  //������
                exit;
             End
             else
             begin
                if not Conn_Ceh.mkQuerySelect2(
                'Select * from %s Where (�����������='+#39+'������'+#39+
                ') AND ([Id��]='+#39+IDGP+#39+') '
                 , [Tab2]) then  //������
                exit;
             end;
             KOL_RECORD_Tab2:=Conn_Ceh.AQuery_Del.RecordCount;
           end
           else
           begin
                if not Conn_Klap.mkQuerySelect2(
                'Select * from %s Where (�����������='+#39+'������'+#39+
                ') AND ([Id��]='+#39+IDGP+#39+') AND ([Id��]='+#39+IDKO+#39+')'
                 , [Tab2]) then  //������
                exit;
                KOL_RECORD_Tab2:=Conn_Klap.AQuery_Del.RecordCount;
           end;

           For X:=0 to KOL_RECORD_Tab2-1 Do
           Begin
             if (Polzov<>4) then
             Begin
                Obozn:=Conn_Klap.AQuery_Del.FieldByName('�����������').AsString;
                Elem:=Conn_Klap.AQuery_Del.FieldByName('�������').AsString;
                Kol_Det:=Conn_Klap.AQuery_Del.FieldByName('����������').AsString; // AND (��������=1)
              end;
             if (Polzov=4) then
             Begin
               Res:=Pos('���������',tab2);
               if Res=0 then
               Begin
                Obozn:=Conn_Ceh.AQuery_Del.FieldByName('�������').AsString;
                Elem:=Conn_Ceh.AQuery_Del.FieldByName('������').AsString;
                Kol_Det:=Conn_Ceh.AQuery_Del.FieldByName('Kol').AsString; // AND (��������=1)
               end
               else
               begin
                 Obozn:=Conn_Ceh.AQuery_Del.FieldByName('�����������').AsString;
                Elem:=Conn_Ceh.AQuery_Del.FieldByName('�������').AsString;
                Kol_Det:=Conn_Ceh.AQuery_Del.FieldByName('����������').AsString;
               end;
              end;
                if not Conn_Klap.mkQuerySelect3(
                'Select * from %s Where (�����������='+#39+Obozn+#39+') '
                 , ['�����������']) then
                exit;
                if  Conn_Klap.AQuery3.RecordCount<>0 then
                Begin
                        SG2.RowCount:=SG2.RowCount+ 1;
                        sg2.Cells[0,Glob+1]:=Kol_Klap;//IntToStr(Glob+1);

                        sg2.Cells[1,Glob+1]:=ZAKAZ;//Conn_Klap.AQuery.FieldByName('�����').AsString;
                        sg2.Cells[2,Glob+1]:=Izdel;//Conn_Klap.AQuery.FieldByName('�������').AsString;
                        //�������� ���� ��������, ���� ���� �� ������ �� ����������
                        sg2.Cells[3,Glob+1]:=IntToStr(StrToInt(Kol_Klap)*StrToInt(Kol_Det));
                        sg2.Cells[4,Glob+1]:=Conn_Klap.AQuery3.FieldByName('�����������').AsString;
                        sg2.Cells[5,Glob+1]:=Conn_Klap.AQuery3.FieldByName('�������').AsString;
                        sg2.Cells[6,Glob+1]:=Nom;
                        sg2.Cells[7,Glob+1]:=BZ;
                        sg2.Cells[8,Glob+1]:=IDGP;
                        sg2.Cells[9,Glob+1]:=Conn_Klap.AQuery3.FieldByName('��������').AsString;
                        sg2.Cells[10,Glob+1]:=IDKO;
                        Res_Teh:=Pos('rue',sg2.Cells[9,Glob+1]);
                        if (Res_Teh=0)  Then //�������� FALSE
                        begin
                                  Memo12.Add('�������� FALSE ����� '+sg2.Cells[1,Glob+1]+' * '+sg2.Cells[2,Glob+1]+' * ������ '+sg2.Cells[4,Glob+1]);
                                  Memo13.Add(sg2.Cells[4,Glob+1]);
                        end;
                        A:=1;
                        for y := 0 to High(Column) do
                        begin
                                if Column[y,0]<>'' then
                                Begin
                                        Stanok:=Column[y,0]; //�������� ����

                                        sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery3.FieldByName(Stanok).AsString;
                                        Res_Stan:=Pos('rue',sg2.Cells[9+(A+1),Glob+1]);
                                        Inc(A);

                                        sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery3.FieldByName(Stanok+'��').AsString; //�����������
                                        Och_D:=StrToFloat(sg2.Cells[9+(A+1),Glob+1]);
                                        Inc(A);

                                        sg2.Cells[9+(A+1),Glob+1]:='0.1';//Conn_Klap.AQuery3.FieldByName(Stanok+'��').AsString;//������������
                                        Dlit_D:=StrToFloat( sg2.Cells[9+(A+1),Glob+1]);//������������
                                        Inc(A);

                                        if (Res_Stan<>0) AND ((Och_D=0) OR (Dlit_D=0))  Then //������ TRUE
                                                                                             //� ����=0 ��� ���� =0 �����
                                        begin
                                                Memo12.Add('������ TRUE � ����=0 ��� ���� =0 ����� '+sg2.Cells[1,Glob+1]+' * '+sg2.Cells[2,Glob+1]+' * ������ '+sg2.Cells[4,Glob+1]);
                                                Memo13.Add(sg2.Cells[4,Glob+1]);
                                        end;

                                        sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery3.FieldByName('�\�'+Stanok).AsString;//�\�
                                        Inc(A);
                                        //sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery3.FieldByName(Stanok+'����').AsString; //����
                                        //Inc(A);
                                End;
                        end;
                        Inc(Glob);
                end;
                if Polzov<>4 then
                  Conn_Klap.AQuery_Del.next;
                if Polzov=4 then
                  Conn_Ceh.AQuery_Del.next;
           end;
        end;
        if Polzov<>4 then
          Conn_Klap.AQuery.next;
        if Polzov=4 then  //����
          Conn_Ceh.AQuery.next;  //gulyakova-ei
      end;
      Memo12_Count:=Memo12.Count;
      Memo12_Text:= Memo12.Text;
      God := FormatDateTime('yyyy', Now);
      Mes := FormatDateTime('mmmm', Now);
      if Polzov=4 then  //����
      Begin
        Dir :=Form1.Put_KTO+'\CKlapana\VRAN)\' + God;
        CreateDir(Dir);

        Dir := Form1.Put_KTO+'\CKlapana\VRAN\' + God + '\' + Mes+'\';
        CreateDir(Dir);
      End;
      Dir :=Form1.Put_KTO+'\CKlapana\���������(�����)\' + God;
        CreateDir(Dir);

      Dir := Form1.Put_KTO+'\CKlapana\���������(�����)\' + God + '\' + Mes+'\';
        CreateDir(Dir);

      Dir :=Form1.Put_KTO+'\CKlapana\' + God;
        CreateDir(Dir);
      Dir := Form1.Put_KTO+'\CKlapana\' + God + '\' + Mes+'\';
        CreateDir(Dir);
      if (Polzov=0) AND (Memo12_Count=0)  Then //��������
      //��� ������ ���������� ,�������
      Begin
        Flag_Error:=0;
        Memo12.Destroy;
        Exit;
      end;
        if (Polzov=0) AND (Memo12_Count<>0) then  //�����   ��������
        Begin
                MessageDlg('�� ��� ������ ����������! ������ � ����� � ���������.', mtError,
                        [mbOk], 0);
                Res_Tab:=Pos('���',Tab2);
                if Res_Tab=0 Then           // PAnsiChar(P
                Begin
                        Memo12.SaveToFile(Form1.Put_KTO+'\CKlapana\���������(�����)\'+God+'\'+Mes+'\'+Nom+'.txt');
                        Nam :=  Form1.Put_KTO+'\CKlapana\���������(�����)\'+God+'\'+Mes+'\'+Nom+'.txt';
                End;
                if Res_Tab<>0 Then//���������
                Begin
                        Memo12.SaveToFile(Form1.Put_KTO+'\CKlapana\'+God+'\'+Mes+'\'+Nom+'.txt');
                        Nam := Form1.Put_KTO+'\CKlapana\'+God+'\'+Mes+'\'+Nom+'.txt';
                End;
                ShellExecute(0, nil, PChar(Nam), nil, nil, SW_SHOWNORMAL);
                File_Name:= Nam ;
                Flag_Error:=2;
                Memo12.Destroy;
                Result := False;
                Exit;
        End;
        if (Polzov<>0) AND (Memo12_Count<>0) then  //�����
        Begin
                MessageDlg('�� ��� ������ ����������! ������ � ����� � ���������.', mtError,
                        [mbOk], 0);
                Res_Tab:=Pos('���',Tab2);
                if (Res_Tab=0) AND (Polzov<>4) Then
                Begin
                        Memo12.SaveToFile(Form1.Put_KTO+'\CKlapana\���������(�����)\'+God+'\'+Mes+'\'+Nom+'.txt');
                        Nam :=  Form1.Put_KTO+'\CKlapana\���������(�����)\'+God+'\'+Mes+'\'+Nom+'.txt';
                End;
                if (Res_Tab<>0) AND (Polzov<>4) Then//���������
                Begin
                        Memo12.SaveToFile(Form1.Put_KTO+'\CKlapana\'+God+'\'+Mes+'\'+Nom+'.txt');
                        Nam :=  Form1.Put_KTO+'\CKlapana\'+God+'\'+Mes+'\'+Nom+'.txt';
                End;
                if  (Polzov=4) Then//VRAN
                Begin
                        Memo12.SaveToFile(Form1.Put_KTO+'\CKlapana\VRAN\'+God+'\'+Mes+'\'+Nom+'.txt');
                        Nam :=  Form1.Put_KTO+'\CKlapana\VRAN\'+God+'\'+Mes+'\'+Nom+'.txt';
                End;
                //ShellExecute(0, nil, PChar(Nam), nil, nil, SW_SHOWNORMAL);
                File_Name:= Nam ;
                Flag_Error:=2;
                Memo12.Destroy;
                Result := False;
                Exit;
        End;
      except

        exit;
    end;
    Memo12.Destroy;
    Memo13.Destroy;
    if not Conn_Klap.mkQueryDelete( 'DELETE FROM %s WHERE (�����='+Nom+') AND (�����='+#39+ZAKAZ+#39+')'
    , ['���������']) then
    exit;

    Result := true;
    btn2Click(fileName) ;
end;
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
function Osnova_Main.btn2Click(fileName:String):Boolean;
        var Dat_S1,IDGP,IDKO,Obozn,Detal,Str,Elem,Stanok,Str1,Str2,Izdel,Kol_Klap,Kol_Det,Zakaz,Nomer,Max:string;
        I,X,Y,A,Glob,Col,Col_Nom,
        Kol_Klap_I,Kol_Det_I:Integer;
        B,D:Boolean;                   // ,S1
        DEF_CatalogName,MSSQL_CONN_STR,S,Razmernost,Nom_Kol,Znach,Str_Stanok,Err,BZ:string;
        TempTableQ:TADOQuery;
        ADOConnection4:TADOConnection;
        Dat_S,Nedely,Kol,Dir,SQLServer,Dat_Z,Nowa_S,Dlit,Dlit_Ob,NCH,fileName1,Kod,Dat_Sbor:string;
        Dat, Dat1,Dat_Zag,nowa,Dat_Sbor_D:TDate;
        Ocheredn,Res_True:Integer;
        Dliteln,
        NC_D:Double;
        //StrToCol:array[0..100,0..5] of string;//[�������, ������������, ��� ������������, ����] ��� ����� ������
        Kol_Stan:Integer;
        IPs : TStringList;
        Res_Trum:Integer;
        Dl_Trum:Double;
        Memo11:TStringList;
begin
       SyStem.SysUtils.FormatSettings.DecimalSeparator :=('.');
       Memo11:=TStringList.Create;
       Memo11.Clear;
       Glob:=0;
       A:=0;
      //==================================
      S:='';
      Razmernost:='';
      A:=0;
      fileName1:=fileName;
      for y := 0 to High(Column) do
      begin
        if Column[y,0]<>'' then
        Begin
                Str_Stanok:=Str_Stanok+Column[y,0]+','+Column[y,0]+'��,'+Column[y,0]+'��,[�\�'+Column[y,0]+'],'+Column[y,0]+'����,'+Column[y,0]+'�����,';
                Nom_Kol:=IntToStr(StrToInt(Column[y,2])-2);
                Razmernost:=Razmernost+Column[y,0]+' BIT,'+Column[y,0]+'�� nvarchar(50),'+Column[y,0]+'�� nvarchar(50),[�\�'+Column[y,0]+'] nvarchar(50),'+Column[y,0]+'���� DateTime,'+Column[y,0]+'����� nvarchar(50),';  //DateTime����� bit,������� nvarchar(50)
        End;
      end;

      S:= SG2.Cells[7,0];
      //++++++++++++++++++++++++++++++++++++
      Dir:=ExtractFileDir(ParamStr(0));
      IPs := TStringList.Create;
      IPs.Clear;
      IPs.LoadFromFile(Dir+'\ConnKlap.ini');
        //Memo1.Lines.LoadFromFile(Dir+'\ConnKlap.ini');

        DEF_CatalogName :=IPs.Strings[0];
        SQLServer:=IPs.Strings[1];
        MSSQL_CONN_STR :=IPs.Strings[4];
       { 'Provider=SQLOLEDB.1;Packet Size = 4096;Password=111;' +
        'Persist Security Info=True;User Id=TestUser;' +
        'Initial Catalog='+ DEF_CatalogName +';Data Source=DINAMO\'+SQLServer; }
        ADOConnection4:=TADOConnection.Create(Nil);
        ADOConnection4.LoginPrompt:=False;
        ADOConnection4.Connected:=False;
        ADOConnection4.ConnectionString := MSSQL_CONN_STR;
        ADOConnection4.Connected:=True;
         //������� ������
        Y:=Length(Razmernost);
        Delete(Razmernost,Y,1);
        //�������� ��������� �������
        TempTableQ:=TADOQuery.Create(nil);
        TempTableQ.Connection:=ADOConnection4;
        //�������� ���� ��������, ���� ���� �� ������ �� ����������
        TempTableQ.SQL.Text:='CREATE TABLE #ClientToDBFL (����� nvarchar(100),������� nvarchar(100),����� nvarchar(100),'+
        '����������� nvarchar(100) ,������� nvarchar(100),���������� nvarchar(100),�� nvarchar(100),Id�� nvarchar(100),Id�� nvarchar(100),������� nvarchar(100),'+Razmernost+')';
        try
                TempTableQ.ExecSQL;
        except
                Err:='TempTableQ.ExecSQL;';
        end;
        //------------------------------------------------
      //������� ������
      Y:=Length(Str_Stanok);
      Delete(Str_Stanok,Y,1);

      //-------------------------------------------------
        for Y:=0 to SG2.RowCount-1 do
        begin
         if Y >0 then
         begin                        //  AND (SG2.Cells[2,Y]='')
                if (SG2.Cells[1,Y]<>'') then
                begin
                   //------------------------------------------------
                   S1:='';
                   Dliteln:=0;
                   Ocheredn:=0;
                   Col_Nom:=0;
                   A :=StrToInt(Column[0,2])+4;
                   i :=0;
                   Kol_Stan:=0;
                   fillchar(StrToCol,sizeof(StrToCol),0);//������� �������
                   for X := 0 to High(Column) do  //������ ������� �������
                   begin
                        if Column[X,2]<>'' then
                        Begin   //'����,������,������,[�\�����],
                                //'False','0','0','0',
                                i :=0;
                                Col_Nom:=StrToInt(Column[X,2])-1;
                               // S1:=S1+#39+SG2.Cells[A,Y]+#39+','; //True  0
                                StrToCol[X, I ]:=SG2.Cells[A,Y];
                                Inc(A);
                                Inc(i);
                                //
                               // S1:=S1+#39+SG2.Cells[A,Y]+#39+','; //Ocheredn 1
                                StrToCol[X,i]:=SG2.Cells[A,Y];
                                Inc(A);
                                Inc(i);
                                //
                               // S1:=S1+#39+SG2.Cells[A,Y]+#39+','; //Dliteln 2
                                StrToCol[X,i]:=SG2.Cells[A,Y];
                                Inc(A);
                                Inc(i);
                                //
                                //S1:=S1+#39+SG2.Cells[A,Y]+#39+','; //�\�  3
                                StrToCol[X,i]:=SG2.Cells[A,Y];
                                Inc(A);
                                Inc(i);
                                //
                                //S1:=S1+#39+'0'+#39+',';  // ����
                                StrToCol[X,i]:='0';
                                Inc(i);
                                //
                                StrToCol[X,i]:=Column[X,0];
                                Inc(i);
                                //
                        End;
                   end;
                   Kol_Klap:=SG2.Cells[0,Y];
                   if Kol_Klap='' Then
                   Kol_klap_I:=0
                   Else
                   Kol_Klap_I:=StrToInt(Kol_Klap);

                   ///()))))))))))))))))))))))(((((((((((((((((((((
                   Btn22Click(fileName1); ///()))))))))))))))))))))))(((((((((((((((((((((
                   ///()))))))))))))))))))))))(((((((((((((((((((((
                   //������� ������
                   i:=Length(S1);
                   Delete(S1,i,1);

                   Zakaz:=SG2.Cells[1,Y];
                   Obozn:=SG2.Cells[4,y];
                   Izdel:=SG2.Cells[2,y];
                   Kol_Det:=SG2.Cells[3,y];
                   Elem:=SG2.Cells[5,y];
                   Nomer:=SG2.Cells[6,y];
                   BZ:=SG2.Cells[7,y];
                   IDGP:=SG2.Cells[8,y];
                   IDKO:=SG2.Cells[10,y];
                   //���-�� ������� ��� �������� �� ��� ��������!!!!
                   TempTableQ.SQL.Text:='INSERT INTO #ClientToDBFL '+
                        '(�����,�������,�����������,�������,'+
                        '����������,�����,��,Id��, Id��,�������,'+Str_Stanok+') Values ('
                        +#39+Zakaz+#39+
                        ','+#39+Izdel+#39+
                        ','+#39+Obozn+#39+
                        ','+#39+Elem+#39+
                        ','+#39+Kol_Det+#39+
                        ','+Nomer+
                        ','+#39+BZ+#39+
                        ','+#39+IDGP+#39+ ','+#39+IDKO+#39+
                        ','+#39+Kol_Klap+#39+
                        ','+S1+
                        ')';

                        try
                        TempTableQ.ExecSQL;
                        except
                        Err:='1222';
                        end;
                end;

         end;
        end;
        //����������� �� TRUE
        S1:='';
        A:=1;
        Glob:=1;
        sg3 := TStringGrid.Create(Nil );
        Clear_StringGrid(sg3);
    SG3.ColCount:=17;
    sg3.Cells[0,0]:='�������';
    sg3.Cells[1,0]:='�����';
    sg3.Cells[2,0]:='�������';
    sg3.Cells[3,0]:='�����';

    sg3.Cells[4,0]:='�����������';
    sg3.Cells[5,0]:='�������';
    sg3.Cells[6,0]:='����������';
    sg3.Cells[7,0]:='�����������'; //�����������
    sg3.Cells[8,0]:='������������';//������������
    sg3.Cells[9,0]:='�\�';//�\�
    sg3.Cells[10,0]:='��';
    sg3.Cells[11,0]:='Id��';
    sg3.Cells[12,0]:='������';
    sg3.Cells[13,0]:='������������';
    sg3.Cells[14,0]:='���������';
    sg3.Cells[15,0]:='���� �����';
    sg3.Cells[16,0]:='Id��';
    //Memo12:=TStringList.Create;
    //Memo12.Clear;
        for Y:=0 to High(Column) do  //������ ������� �������
        begin
                if Column[y,0]<>'' then
                Begin
                        TempTableQ.SQL.Text:='SELECT *  FROM #ClientToDBFL'+
                        ' WHERE '+Column[y,0]+'='+#39+'True'+#39 ;
                        TempTableQ.Open;
                        //+++++++++++++++++++++++++++++++++++++++
                        TempTableQ.First;
                        if TempTableQ.RecordCount<>0 Then
                        Begin
                                sg3.RowCount:= sg3.RowCount+1 ;
                                sg3.Cells[0,A]:=Column[y,0];
                                Stanok:=Column[y,0]; //�������� ����
                                A:=A+1;

                          for I := 0 to TempTableQ.RecordCount - 1 do
                          begin
                                sg3.RowCount:= sg3.RowCount+1 ;
                                Str:=sg3.Cells[0,A];
                                sg3.Cells[0,A]:=TempTableQ.FieldByName('�������').AsString;//IntToStr(Glob);
                                sg3.Cells[1,A]:=TempTableQ.FieldByName('�����').AsString;
                                sg3.Cells[2,A]:=TempTableQ.FieldByName('�������').AsString;
                                sg3.Cells[3,A]:=TempTableQ.FieldByName('�����').AsString;

                                sg3.Cells[4,A]:=TempTableQ.FieldByName('�����������').AsString;
                                sg3.Cells[5,A]:=TempTableQ.FieldByName('�������').AsString;
                                Kol_Det:=TempTableQ.FieldByName('����������').AsString;
                                sg3.Cells[6,A]:= Kol_Det;
                                if Kol_Det='' Then
                                Kol_Det_I:=0
                                Else
                                Kol_Det_I:=StrToInt(Kol_Det);

                                sg3.Cells[7,A]:=TempTableQ.FieldByName(Stanok+'��').AsString; //�����������
                                Ocheredn:=StrToInt( sg3.Cells[7,A]);
                                sg3.Cells[8,A]:=TempTableQ.FieldByName(Stanok+'��').AsString;//������������
                                Dliteln:=StrToFloat(StringReplace(sg3.Cells[8,A],',','.',[rfReplaceAll]));
                                sg3.Cells[9,A]:=TempTableQ.FieldByName('�\�'+Stanok).AsString;//�\�
                                NCH:=StringReplace(sg3.Cells[9,A],',','.',[rfReplaceAll]);
                                if NCh='' Then
                                NC_D:=0
                                Else
                                NC_D:=StrToFloat(NCH);
                                sg3.Cells[10,A]:=TempTableQ.FieldByName('��').AsString;
                                sg3.Cells[11,A]:=TempTableQ.FieldByName('Id��').AsString;
                                sg3.Cells[12,A]:=Column[y,0];
                                sg3.Cells[13,A]:=Dat_Glob;

                                Dat_Z :=TempTableQ.FieldByName(Stanok+'����').AsString;
                                Dat_Zag:= StrToDate(Dat_Z);
                                Nowa_S:=FormatDateTime('dd.mm.YYYY',Now);
                                Nowa:=StrToDate(Nowa_s);
                                {if Dat_Zag<=Nowa then //���� ������ ��������� ������ �����������
                                begin
                                        MessageDlg(Nom_G+'  ���� ������ ��������� ������ �����������!' + #10#13 +
                                        '���������� �������� ���� ������������ �� ' + FloatToStr(Nowa-Dat_Zag)+' ����!' ,mtInformation, [mbOk], 0);
                                        Flag_Error:=1;
                                        Exit;
                                end; }
                                Dat_Sbor:=TempTableQ.FieldByName(Stanok+'����').AsString;
                                Dat_Sbor_D:=StrToDate(Dat_Sbor)+1;
                                sg3.Cells[13,A]:=FormatDateTime('dd.mm.YYYY',Dat_Sbor_D);
                                sg3.Cells[14,A]:=TempTableQ.FieldByName(Stanok+'����').AsString;
                                sg3.Cells[15,A]:=TempTableQ.FieldByName(Stanok+'�����').AsString;
                                sg3.Cells[16,A]:=TempTableQ.FieldByName('Id��').AsString;
                                Dlit:=StringReplace(sg3.Cells[8,A],',','.',[rfReplaceAll]);
                                Dlit_Ob:=StringReplace(sg3.Cells[15,A],',','.',[rfReplaceAll]);

                                //++++++++++++++++++++++++++++++++++++++++++++++++++
                                if sg3.Cells[14,A]='' then
                                        Dat_S1:='NULL'
                                else
                                        DaT_S1:=#39+ConvertDat1(SG3.Cells[14,A])+#39;
                                                        //   AND (��������=1)
                                if not Conn_Klap.mkQuerySelect3(
                                'Select * from %s Where (�����������='+#39+Obozn+#39+')'
                                , ['�����������']) then
                                exit;

                            if  (Conn_Klap.AQuery3.RecordCount<>0)  then
                            Begin
                                NCH:=StringReplace(FloatToStr(NC_D*Kol_Det_I),',','.',[rfReplaceAll]);
                                if not Conn_Klap.mkQueryInsert(
                                'INSERT INTO ��������� '+
                                '(�����,�������,�����,�����������,�������,'+
                                '����������,�����������,������������,��,��,Id��,������,������������,���������,���������������,�������,Id��) Values ('
                                +#39+sg3.Cells[1,A]+#39+
                                ','+#39+sg3.Cells[2,A]+#39+
                                ','+#39+sg3.Cells[3,A]+#39+
                                ','+#39+sg3.Cells[4,A]+#39+
                                ','+#39+sg3.Cells[5,A]+#39+
                                ','+#39+sg3.Cells[6,A]+#39+
                                ','+#39+sg3.Cells[7,A]+#39+
                                ','+#39+Dlit+#39+
                                ','+#39+NCH+#39+
                                ','+#39+sg3.Cells[10,A]+#39+
                                ','+#39+sg3.Cells[11,A]+#39+
                                ','+#39+sg3.Cells[12,A]+#39+
                                ','+#39+ConvertDat1(sg3.Cells[13,A])+#39+
                                ','+Dat_S1+
                                ','+#39+Dlit_Ob+#39+
                                ','+#39+sg3.Cells[0,A]+#39+
                                ','+#39+sg3.Cells[16,A]+#39+
                                ')',[] )then
                                Exit;
                                //+++++++++++++++++++++++++++++++++++++++++++
                                A:=A+1;
                                Glob:=Glob+1;
                            End;
                                TempTableQ.Next;

                          end;
                        end;

                end;
        end;
        TempTableQ.Close;
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        if not Conn_Klap.mkQuerySelect('Select ��������� From ��������� Where (�����='+
        #39+Nomer_Glob+#39+') Group By ���������',[]) Then
        Exit;
        For I:=0 To  Conn_Klap.AQuery.RecordCount-1 Do
        Begin
              Dat_S:= ConvertDat1(Conn_Klap.AQuery.FieldByName('���������').AsString);

              Memo11.Add('');
              Memo11.Add(Dat_S);
              if not Conn_Klap.mkQuerySelect3('Select ������, SUM(��) As SNH From ��������� Where (���������='+
                 #39+Dat_S+#39+') AND (������<>'+#39+'������'+#39+') AND (������<>'+#39+'Trumph'+#39+') Group by ������',[]) Then
                Exit;

              For Y:=0 To  Conn_Klap.AQuery3.RecordCount-1 Do
              Begin
                   Memo11.Add(Conn_Klap.AQuery3.FieldByName('������').AsString+'  �\�: '+Conn_Klap.AQuery3.FieldByName('SNH').AsString);
                   Inc(A);
                   Conn_Klap.AQuery3.Next;
              end;
              Conn_Klap.AQuery.Next;
        end;
        SNCH:=Memo11.Text;
        //================================================
        TempTableQ.SQL.Text:='DROP TABLE #ClientToDBFL';
        try
                TempTableQ.ExecSQL;
        except
        end;

   End;
//++++++++++++++++++++++++++++++++++++++++++++++++++
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function Osnova_Main.btn22Click(Err2:String):Boolean;
var //StrToCol:array[0..100,0..6] of string;
//[TRUE,�������, ������������, �/�, ����,Stanok] ��� ����� ������
I,X,J,y,A,Res:Integer;
Res_True,min_Col,Col,Col_1,
Col_Grid,Res1:Integer;
Izdel,Nowa_S:string;
Dat_Plan,Dat_Zag,Nowa:TDate;
Dat_P,Dat_Z:string;
Dlit,Dlit_Min,Dlit_Obsh,Dlit_Max,
Min,Och_Min,Och_Max,Dlit_ObshO,Ocher,Ocher_Pred:Double;
//SREZ1: array[0..10,0..100] of string;

SKPU :array[0..10,0..100] of string;

//Sort: TStringList;
Stanok,Stanok_Posl:String;
Kol_Smen,Max,F1,Den,Den_pred,Celoe,Res_Trum:Integer;
Chas,H1,H2,H3,Den1,Dlit_Trum:Double;
SREZ1:TAdvStringGrid;
//AD1:TAdvStringGrid;
Chet:Integer;
begin
     X:=0;
     J:=0;
     Dlit:=0;
     Dlit_Obsh:=0;
     Col_Grid:=97;
     Col_1:=0;
     Max:=0;
     SyStem.SysUtils.FormatSettings.DecimalSeparator :=('.');
     //Sort:=TStringList.Create;
     SREZ1 := TAdvStringGrid.Create(Nil );
     //AD1 := TAdvStringGrid.Create(Nil );
     fillchar(Temp,sizeof(Temp),0);//������� �������
     fillchar(SKPU,sizeof(SKPU),0);//������� �������
     Form1.Memo13.Clear;
     //++������� � ������������ ������ 1 ����
    { for I := 0 to High(StrToCol) do
     begin
        Res_True:=Pos('rue',StrToCol[i,0]);
        Res_Trum:=Pos('Trumph',StrToCol[i,5]);
        if (Res_Trum<>0) AND (Res_True<>0) then  //���
        begin
                Dlit_Min:=StrToFloat(StrToCol[i,2]);
                StrToCol[i,2]:=FloatToStr(Dlit_Min+1);
        end;
     end; }
     Dlit_Min:=0;
     //++++++++++++++++++++++++++++++++���� ����������� �� �����������
     for I := 0 to High(StrToCol) do
     begin
        Res_True:=Pos('rue',StrToCol[i,0]);
        if Res_True<>0 then
        begin
                min := StrToInt(StrToCol[i,1]); // ������ �� �������
                Och_Min:= StrToInt(StrToCol[i,1]);
                Col:=I;
                Dlit_Min:=StrToFloat(StrToCol[i,2]);
                Break;
        end;
     end;
     X:=0;
     A:=0;
     Y:=0;
     SREZ1.ColCount:=8;//AD   SREZ
     SREZ1.RowCount:=1;
     //Form1.SKPU.ColCount:=10; //AD1
     //Form1.SKPU.RowCount:=1;

     Dlit_Obsh:=0;
     for I := 0 to High(StrToCol) do
     begin
        Res_True:=Pos('rue',StrToCol[i,0]);
        if Res_True<>0 then
      begin
        Stanok_Posl:=StrToCol[i,5];
        For y:=0 To High(Smena) do
        Begin
          Stanok:=Smena[y,0];
          Res:=AnsiCompareStr(Stanok_Posl,Stanok);
          if Res=0 Then
          Begin
                Kol_Smen:=StrToInt(Smena[y,1]);
                Chas:=StrToFloat(Smena[y,2]);
                SREZ1.Cells[0,A]:=StrToCol[i,1]; //Ocher
                Ocher:= StrToInt(StrToCol[i,1]);

                SREZ1.Cells[1,A]:=Stanok_Posl;   //Stanok
                SREZ1.Cells[2,A]:=StrToCol[i,2];//Dlit
                Dlit_Min:=StrToFloat(StrToCol[i,2]);//*8;// 0,5*8=4 - ��� �����
                                                    //1*8=8 ������ �����
                Dlit_Obsh:=Dlit_Obsh+ (StrToFloat(StrToCol[i,2]));//� �����

                SREZ1.Cells[3,A]:=Smena[y,1];//Kol_Smen
                Kol_Smen:=StrToInt(Smena[y,1]);

                SREZ1.Cells[4,A]:=Smena[y,2];//Chas
                SREZ1.Cells[5,A]:='0';//Den
                SREZ1.Cells[6,A]:=StrToCol[I,3];//�\�
                SREZ1.Cells[7,A]:=StrToCol[I,0];//TRUE
                SREZ1.RowCount:=SREZ1.RowCount+1;
                Inc(A);
          end;
        End;
      End;
     End;
     F1:=0;
     SREZ1.Cells[5,0]:='1';
     //AD.Sort(0,sdAscending);//sdDescending �� ��������
     SREZ1.SortByColumn(0);
     Dat_Plan:=StrToDate(Dat_Glob);
     A:=0;
     For I:=0 To SREZ1.RowCount-1 Do
     Begin
       if SREZ1.Cells[0,i]<> '' Then
       Begin
           {Form1.SKPU.Cells[0,A]:=Form1.SREZ1.Cells[0,i]; //Ocher
           Form1.SKPU.Cells[1,A]:=Form1.SREZ1.Cells[1,i];   //Stanok
           Form1.SKPU.Cells[2,A]:=Form1.SREZ1.Cells[2,i];//Dlit
           Form1.SKPU.Cells[3,A]:=Form1.SREZ1.Cells[3,i];//Kol_Smen
           Form1.SKPU.Cells[4,A]:=Form1.SREZ1.Cells[4,i];//Chas
           Form1.SKPU.Cells[5,A]:='0';//Den
           Form1.SKPU.Cells[6,A]:=Form1.SREZ1.Cells[6,i];//�\�
           Form1.SKPU.Cells[7,A]:=Form1.SREZ1.Cells[7,i];//TRUE
           Form1.SKPU.RowCount:=Form1.SKPU.RowCount+1;}
           SKPU[0,a]:=SREZ1.Cells[0,i]; //Ocher
           SKPU[1,a]:=SREZ1.Cells[1,i];   //Stanok
           SKPU[2,a]:=SREZ1.Cells[2,i];//Dlit
           SKPU[3,a]:=SREZ1.Cells[3,i];//Kol_Smen
           SKPU[4,a]:=SREZ1.Cells[4,i];//Chas
           SKPU[5,a]:='0';//Den
           SKPU[6,a]:=SREZ1.Cells[6,i];//�\�
           SKPU[7,a]:=SREZ1.Cells[7,i];//TRUE
           Inc(A);
       end;
     end;

     //For I:=0 To Form1.SKPU.RowCount-1 Do
     For i:=0 To High(SKPU) do
     Begin
       //if Form1.SKPU.Cells[0,I]<> '' Then
       if SKPU[0,I]<> '' Then
       Begin

            Ocher:= StrToInt(SKPU[0,I]);//Ocher
            Stanok_Posl:=SKPU[1,i];   //Stanok
            Dlit_Min:=StrToFloat(SKPU[2,I]);//Dlit

            Kol_Smen:=StrToInt(SKPU[3,I]);//Kol_Smen
            Chas:=StrToFloat(SKPU[4,I]);
            Den:= StrToInt(SKPU[5,I]);
            Celoe:=Pos(',',SKPU[2,I]);
            if Celoe<>0 Then
            Begin
                if I=0 Then
                Begin
                        H1:=(1-Dlit_Min);
                        SKPU[5,0]:=FloatToStr(1);
                        Form1.Memo13.Lines.Add('1');
                End;
                if I=1 Then
                Begin
                 if (H1=0.5) Then //������� ����� �� ������� ��������
                 Begin
                   SKPU[5,1]:='1';
                   Form1.Memo13.Lines.Add('1');
                   H2:=1;
                  End;
                 if (H1=0) OR (H1<0) Then    //���   OR (H1<0)
                 Begin
                   SKPU[5,1]:='2';   // ��������� �� ������ ����
                   Form1.Memo13.Lines.Add('2');
                   H2:=2;
                  End;
                  H1:=(1-Dlit_Min);
                end;
                if I>1 Then //
                Begin
                 Chet:=I MOD 2;//������� �� �������
                 if (Chet<>0) AND (Dlit_Min=0.5) Then //�� ������
                 Begin
                    H1:=Ocher-Dlit_Min-H2;
                    Den1:= SimpleRoundTo(H1,0);
                    SKPU[5,I]:=(FloatToStr(Den1));
                    Form1.Memo13.Lines.Add(FloatToStr(Den1));
                 end;
                 if (Chet=0) AND (Dlit_Min=0.5) Then //������
                 Begin
                     H2:=Ocher-Dlit_Min-H1;
                     Den1:= SimpleRoundTo(H2,0);
                     SKPU[5,I]:=(FloatToStr(Den1));
                     Form1.Memo13.Lines.Add(FloatToStr(Den1));
                 end;
                end;
            End
            Else
            Begin
                  SKPU[5,I]:=(FloatToStr(Ocher));
                  Form1.Memo13.Lines.Add(FloatToStr(Ocher));
            end;
            {Den:= StrToInt(AD1.Cells[5,I]);
            Dat_Plan:=StrToDate(Dat_Glob);
            Dat_Z:=FormatDateTime('mm.dd.YYYY',Dat_Plan-Den);
            AD1.Cells[8,i]:=Dat_Z;  }
       end;
     end;
     A:= Form1.Memo13.Lines.Count-1 ;
     For i:=0 To High(SKPU) do
     Begin
       if Form1.Memo13.Lines.Strings[a]<> '' Then
       Begin
           SKPU[9,I]:=Form1.Memo13.Lines.Strings[a];
       end;
       A:=A-1;
     End;
     For i:=0 To High(SKPU) do
     Begin
       if SKPU[9,I]<> '' Then
       Begin
         Den:= StrToInt(SKPU[5,I]);//� ������ ������ ���_����- ��� ���� ������ ���������
         Dat_Plan:=StrToDate(Dat_Glob);  //
         Dat_Z:=FormatDateTime('mm.dd.YYYY',Dat_Plan+Den-1);//���� 1����
         SKPU[8,i]:=Dat_Z;

       end;
     End;

     for I := 0 to High(StrToCol) do
     begin
          if StrToCol[i,5]<>'' Then
          Begin
            For a:=0 To High(SKPU) do
            Begin
                Res:=AnsiCompareStr(StrToCol[i,5],SKPU[1,A]);
                if Res=0 Then
                Begin
                    StrToCol[i,4]:=FloatToStr(Dlit_Obsh);
                    StrToCol[i,6]:=SKPU[8,A];//���� ���������
                end;
            End;
          End;
     end;

     for I := 0 to High(StrToCol) do
     begin
        if StrToCol[i,5]<>'' then
        begin
        S1:=S1+#39+StrToCol[I,0]+#39+','; //True  0
        S1:=S1+#39+StrToCol[I,1]+#39+','; //����������� 1
        S1:=S1+#39+StrToCol[I,2]+#39+','; //�������  2
        S1:=S1+#39+StrToCol[I,3]+#39+','; //�\�  3
        if StrToCol[I,2]='' then
          S1:=S1+'NULL,';
        //
        if (StrToCol[I,2]<>'') then
        Begin
            S1:=S1+#39+StrToCol[I,6]+#39+',';
        End;
          //
          S1:=S1+#39+StrToCol[I,4]+#39+',';
          //���� ���������  5 ����� ������ �� ��������� ����
        end;
     end;
     FreeAndNil(Srez1);
     //FreeAndNil(SKPU);
end;

                                                //������                 //Tab2 ������������
function Osnova_Main.Osnova( var Dat : string;  const tab1:string;tab2:string;
Nom:string;fileName:String;Flag_Polz,Cvet:Integer ) : Boolean;

var a,Y,Glob,i,x,Res,T,Res1,L:Integer;
Stanok,Dat_S1,IDGP,Kol_Klap,Izdel,kol_Det,Elem,Obozn,Nomer,BZ,Str1,Str2,Dlit,Kod,
IDKO,Nam_S,ZAKAZ,Dl,Sh,DL_R,Sh_R,Mat,Diam,Mass,Kol_G,Sh1,Sh2,Ie:string;
Dlit_D,Och_D:Double;
Res_Teh,Res_Stan,Memo12_Count,Res_Tab,Polzov,Zavesa,//
KOL_RECORD_Tab1,KOL_RECORD_Tab2:Integer;
Memo12,Memo13:TStringList;
Memo12_text,Nam,God,Mes,Dir:String;
Leg:String;
h: hwnd;
begin
        Result:=False;
        Zavesa:=0;//�� ������
        SyStem.SysUtils.FormatSettings.DecimalSeparator :=('.');
    try
        T:=0;
        Nom_G:=Nom;
        Conn_Ceh := Connect_Miass_Ceh.Create();
        Conn_Klap:=Connect_Miass_Klap.Create();
        Memo12:=TStringList.Create;
        Memo12.Clear;
        Memo13:=TStringList.Create;
        Memo13.Clear;
        Polzov:= Flag_Polz;
        if Polzov=7 then
        Begin
           Zavesa:=1;//������
           Polzov:=6;
        End;
        CVet1:=Cvet;
        if not Conn_Klap.mkQuerySelect('SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME ='+#39+'�����������'+#39
        , []) then
        exit;
        Dat_Glob:= Dat;
        for I:=0 To Conn_Klap.AQuery.RecordCount-1 do
        begin
          Str1:=Conn_Klap.AQuery.FieldByName('COLUMN_NAME').AsString;
          Str2:=Conn_Klap.AQuery.FieldByName('COLUMN_NAME').AsString;
          Res:=Pos('�\�',str2);
          if Res<>0 then
          begin
               Delete(str2,1,3);
               Column[T,0]:=Str2;
               Column[T,1]:=Str2+' BIT';
               Column[T,2]:=IntToStr(T+7);
               Inc(T);
          end;
          Conn_Klap.AQuery.Next;
        End;
        //++++++++++++++++++++++++++++
        fillchar(Smena,sizeof(Smena),0);//������� �������
        if not Conn_Klap.mkQuerySelect('Select * From �����  ' , []) then //������
        exit;
      For I:=0 to Conn_Klap.AQuery.RecordCount-1 Do
      Begin
        Smena[i,0]:=Conn_Klap.AQuery.FieldByName('�������').AsString;
        Smena[i,1]:=Conn_Klap.AQuery.FieldByName('�����').AsString;
        Smena[i,2]:=Conn_Klap.AQuery.FieldByName('����').AsString;
        Conn_Klap.AQuery.Next;
      End;
          sg2 := TStringGrid.Create(Nil);
          Clear_StringGrid(sg2);
        //sg2.Clear;
        A:=1;
        sg2.ColCount:=23;
        sg2.Cells[0,0]:='�';
        sg2.Cells[1,0]:='�����';
        sg2.Cells[2,0]:='�������';
        sg2.Cells[3,0]:='��� ��';
        sg2.Cells[4,0]:='�����������';
        sg2.Cells[5,0]:='�������';
        sg2.Cells[6,0]:='�����';
        sg2.Cells[7,0]:='��';
        sg2.Cells[8,0]:='Id��';
        sg2.Cells[9,0]:='��������';
        sg2.Cells[10,0]:='Id��';
        sg2.Cells[11,0]:='������';
        //Dl,Sh,DL_R,Sh_R,Mat,Diam,Mass,Kol_G,Sh1,Sh2,Ie
        sg2.Cells[12,0]:='�����';
        sg2.Cells[13,0]:='������';
        sg2.Cells[14,0]:='���������';
        sg2.Cells[15,0]:='����������';
        sg2.Cells[16,0]:='�������';
        sg2.Cells[17,0]:='�����';
        sg2.Cells[18,0]:='��������';
        sg2.Cells[19,0]:='�����������1';
        sg2.Cells[20,0]:='�����������2';
        sg2.Cells[21,0]:='��';
        sg2.Cells[22,0]:='��������';
        for y := 0 to High(Column) do
        begin
                if Column[y,0]<>'' then
                Begin
                        Stanok:=Column[y,0];
                        SG2.ColCount:=SG2.ColCount+4;
                        sg2.Cells[10+A+1,0]:=Stanok;
                        sg2.Cells[10+A+2,0]:=Stanok+'��'; //�����������
                        sg2.Cells[10+A+3,0]:=Stanok+'��';//������������
                        sg2.Cells[10+A+4,0]:='�\�'+Stanok;//�\�
                        //sg2.Cells[9+A+5,0]:=Stanok+'����'; //����
                        A:=A+4;
                end;
        end;
        Glob:=1;//SG2.RowCount;
        //Dat_S1:=FormatDateTime('mm.dd.YYYY',Dtp1.DateTime);
        if Polzov<4 then
        begin
          Res1:=Pos('-2',Nom);
          Res:=Pos('-3',Nom);
          if (CVet1>10) and (Res1<>0) then
          begin
          if not Conn_Klap.mkQuerySelect('Select * From %s Where (ID��='+#39+'%s'+#39+') ', ['����',IntToStr(CVet1)]) then //������
          exit;
          end;
          if (CVet1>10) and (Res<>0) then
          begin
          if not Conn_Klap.mkQuerySelect('Select * From %s Where (ID��='+#39+'%s'+#39+') ', ['���',IntToStr(CVet1)]) then //������
          exit;
          end;
          if (Res=0) and (Res1=0) then
          if not Conn_Klap.mkQuerySelect('Select * From %s Where (�����=%s) '+
          'AND (������ IS NULL) ' , [tab1,Nom]) then //������
          exit;
          KOL_RECORD_Tab1:= Conn_Klap.AQuery.RecordCount;
        end;
        ///����������� �� ����
        if (Polzov=4) or (Polzov=6) then
        begin
          Res:=Pos('-1',Nom);
          if (CVet1>10) and (Res<>0) then
          begin
          if not Conn_Ceh.mkQuerySelect('Select * From %s Where (ID��='+#39+'%s'+#39+') ', ['������',IntToStr(CVet1)]) then //������
          exit;
          end

          else
          if not Conn_Ceh.mkQuerySelect('Select * From %s Where (�����=%s) '+
          'AND (������ IS NULL) ' , [tab1,Nom]) then //������
          exit;
          KOL_RECORD_Tab1:= Conn_Ceh.AQuery.RecordCount;
        end;
        ///�� �� ����    ������ ������ ������
        if Polzov=5 then
        begin
          if not Conn_Ceh.mkQuerySelect('Select * From %s Where (������=%s) '+
          'AND (������ IS NULL) AND (������=1) AND ��������='+#39+ConvertDat1(Dat)+#39 , [tab1,Nom]) then //��
          exit;
          KOL_RECORD_Tab1:= Conn_Ceh.AQuery.RecordCount;
        end;
        //
        if Polzov=55 then
        begin
        Res:=Pos('-55',Nom);
          if (CVet1>10) and (Res<>0) then
          begin
          if not Conn_Ceh.mkQuerySelect('Select * From %s Where (ID��='+#39+'%s'+#39+') AND (ID��='+#39+'%s'+#39+') '
          , ['����',IntToStr(CVet1),(fileName)]) then //CVet1=GP fileName  =KO
          exit;
          fileName:='';
          end

          else             //  AND (������=1)
          if not Conn_Ceh.mkQuerySelect('Select * From %s Where ([� ����� ������]=%s) '+
          'AND (������ IS NULL)  AND ��������='+#39+ConvertDat1(Dat)+#39 , [tab1,Nom]) then //��
          exit;
          KOL_RECORD_Tab1:= Conn_Ceh.AQuery.RecordCount;
        end;
        SG2.RowCount:= 2;
        SG2.FixedRows:=1;
        A:=1;
      For I:=0 to KOL_RECORD_Tab1-1 Do
      Begin
        if Polzov<4 then
        begin
          IDGP:=Conn_Klap.AQuery.FieldByName('Id��').AsString;
          IDKO:=Conn_Klap.AQuery.FieldByName('Id��').AsString;
          Izdel:=Conn_Klap.AQuery.FieldByName('�������').AsString;
          Kol_Klap:=Conn_Klap.AQuery.FieldByName('��� �� ����������').AsString;
          Nomer:=Conn_Klap.AQuery.FieldByName('�����').AsString;
          Nomer_Glob:=Conn_Klap.AQuery.FieldByName('�����').AsString;
          BZ:=Conn_Klap.AQuery.FieldByName('��').AsString;
          ZAKAZ:=Conn_Klap.AQuery.FieldByName('�����').AsString;
          Leg:=Conn_Klap.AQuery.FieldByName('������').AsString;

        end;
        if Polzov=4 then
        begin
          IDGP:=Conn_Ceh.AQuery.FieldByName('Id��').AsString;
          IDKO:=Conn_Ceh.AQuery.FieldByName('Id��').AsString;
          Izdel:=Conn_Ceh.AQuery.FieldByName('�������').AsString;
          Kol_Klap:=Conn_Ceh.AQuery.FieldByName('��� ����������').AsString;
          Nomer:=Conn_Ceh.AQuery.FieldByName('�����').AsString;
          Nomer_Glob:=Conn_Ceh.AQuery.FieldByName('�����').AsString;
          BZ:=Conn_Ceh.AQuery.FieldByName('��').AsString;
          ZAKAZ:=Conn_Ceh.AQuery.FieldByName('�����').AsString;
          //Leg:=Conn_Klap.AQuery.FieldByName('������').AsBoolean;
        end;
        if (Polzov=5)  then  //����
        begin
          IDGP:=Conn_Ceh.AQuery.FieldByName('Id��').AsString;
          IDKO:=Conn_Ceh.AQuery.FieldByName('Id��').AsString;
          Izdel:=Conn_Ceh.AQuery.FieldByName('������������').AsString;
          Kol_Klap:=Conn_Ceh.AQuery.FieldByName('���-��').AsString;
          Nomer:=Conn_Ceh.AQuery.FieldByName('������').AsString;
          Nomer_Glob:=Conn_Ceh.AQuery.FieldByName('������').AsString;
          BZ:=Conn_Ceh.AQuery.FieldByName('� ����� ������').AsString;
          ZAKAZ:=Conn_Ceh.AQuery.FieldByName('�����').AsString;
        end;
        if  (Polzov=55) then  //

        begin
          IDGP:=Conn_Ceh.AQuery.FieldByName('Id��').AsString;
          IDKO:=Conn_Ceh.AQuery.FieldByName('Id��').AsString;
          Izdel:=Conn_Ceh.AQuery.FieldByName('������������').AsString;
          Kol_Klap:=Conn_Ceh.AQuery.FieldByName('���-��').AsString;
          if Kol_Klap='0' then
          Kol_Klap:='1';
          Nomer:=Conn_Ceh.AQuery.FieldByName('� ����� ������').AsString;
          Nomer_Glob:=Conn_Ceh.AQuery.FieldByName('� ����� ������').AsString;
          BZ:=Conn_Ceh.AQuery.FieldByName('� ����� ������').AsString;
          ZAKAZ:=Conn_Ceh.AQuery.FieldByName('�����').AsString;
        end;

        if Polzov=6 then   //���
        begin
          IDGP:=Conn_Ceh.AQuery.FieldByName('Id��').AsString;
          IDKO:=Conn_Ceh.AQuery.FieldByName('Id��').AsString;
          res:=Pos('����',tab1);
          if Res=0 then
          Begin
            Izdel:=Conn_Ceh.AQuery.FieldByName('������������').AsString ;
            Kol_Klap:=Conn_Ceh.AQuery.FieldByName('���-��').AsString;
            BZ:=Conn_Ceh.AQuery.FieldByName('� ����� ������').AsString;
          End
          else
          Begin
            Izdel:=Conn_Ceh.AQuery.FieldByName('�������').AsString; //������
            Kol_Klap:=Conn_Ceh.AQuery.FieldByName('��� ����������').AsString;
            BZ:=Conn_Ceh.AQuery.FieldByName('��').AsString;
          End;
          Nomer:=Conn_Ceh.AQuery.FieldByName('�����').AsString;
          Nomer_Glob:=Conn_Ceh.AQuery.FieldByName('�����').AsString;
          ZAKAZ:=Conn_Ceh.AQuery.FieldByName('�����').AsString;
        end;
        if Tab2='' Then //Tab2 ������ ������ ������������ ��� � �����������  , ���� �� �������� ������� ����....
        Begin

                 if not Conn_Klap.mkQuerySelect2(
                'Select * from %s Where (�����������='+#39+'������'+#39+') '+
                'AND (([�������]='+#39+'������ ��������� '+Izdel+#39+') OR ([�������]='+#39+'������ '+Izdel+#39+'))'
                 , ['�����������']) then  //������
                 exit;
                 For X:=0 to Conn_Klap.AQuery_Del.RecordCount-1 Do
                Begin
                        Obozn:=Conn_Klap.AQuery_Del.FieldByName('�����������').AsString;
                        Elem:=Conn_Klap.AQuery_Del.FieldByName('�������').AsString;
                        Kol_Det:=Conn_Klap.AQuery_Del.FieldByName('����������').AsString;
                   if  Conn_Klap.AQuery_Del.RecordCount<>0 then
                   Begin
                        SG2.RowCount:=SG2.RowCount+ 1;
                        sg2.Cells[0,Glob+1]:=Kol_Klap;//IntToStr(Glob+1);

                        sg2.Cells[1,Glob+1]:=Conn_Klap.AQuery.FieldByName('�����').AsString;
                        sg2.Cells[2,Glob+1]:=Conn_Klap.AQuery.FieldByName('�������').AsString;

                        sg2.Cells[3,Glob+1]:=IntToStr(StrToInt(Kol_Klap)*StrToInt(Kol_Det));
                        sg2.Cells[4,Glob+1]:=Conn_Klap.AQuery_Del.FieldByName('�����������').AsString;
                        sg2.Cells[5,Glob+1]:=Conn_Klap.AQuery_Del.FieldByName('�������').AsString;
                        sg2.Cells[6,Glob+1]:=Nom;
                        sg2.Cells[7,Glob+1]:=BZ;
                        sg2.Cells[8,Glob+1]:=IDGP;
                        sg2.Cells[9,Glob+1]:=Conn_Klap.AQuery3.FieldByName('��������').AsString;
                        sg2.Cells[10,Glob+1]:=IDKO;
                        sg2.Cells[11,Glob+1]:=(Leg);

                        Res_Teh:=Pos('rue',sg2.Cells[9,Glob+1]);
                        if (Res_Teh=0)  Then //�������� FALSE
                        begin
                                  Memo12.Add('�������� FALSE ����� '+sg2.Cells[1,Glob+1]+' * '+sg2.Cells[2,Glob+1]+' * ������ '+sg2.Cells[4,Glob+1]);
                                  Memo13.Add(sg2.Cells[4,Glob+1]);
                        end;
                        A:=1;
                        for y := 0 to High(Column) do
                        begin
                                if Column[y,0]<>'' then
                                Begin
                                        Stanok:=Column[y,0]; //�������� ����

                                        sg2.Cells[10+(A+1),Glob+1]:=Conn_Klap.AQuery_Del.FieldByName(Stanok).AsString;
                                        Res_Stan:=Pos('rue',sg2.Cells[10+(A+1),Glob+1]);
                                        Inc(A);

                                        sg2.Cells[10+(A+1),Glob+1]:=Conn_Klap.AQuery_Del.FieldByName(Stanok+'��').AsString; //�����������
                                        Och_D:=StrToFloat(sg2.Cells[10+(A+1),Glob+1]);
                                        Inc(A);

                                        Dlit:= '0.1';//Conn_Klap.AQuery_Del.FieldByName(Stanok+'��').AsString;//������������
                                        sg2.Cells[10+(A+1),Glob+1]:='0.1';//Conn_Klap.AQuery_Del.FieldByName(Stanok+'��').AsString;//������������
                                        Dlit_D:=StrToFloat( sg2.Cells[10+(A+1),Glob+1]);//������������
                                        Inc(A);

                                        if (Res_Stan<>0) AND ((Och_D=0) OR (Dlit_D=0))  Then //������ TRUE
                                                                                             //� ����=0 ��� ���� =0 �����
                                        begin
                                                Memo12.Add('������ TRUE � ����=0 ��� ���� =0 ����� '+sg2.Cells[1,Glob+1]+' * '+sg2.Cells[2,Glob+1]+' * ������ '+sg2.Cells[4,Glob+1]);
                                                Memo13.Add(sg2.Cells[4,Glob+1]);
                                        end;

                                        sg2.Cells[10+(A+1),Glob+1]:=Conn_Klap.AQuery_Del.FieldByName('�\�'+Stanok).AsString;//�\�
                                        Inc(A);
                                        //sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery3.FieldByName(Stanok+'����').AsString; //����
                                        //Inc(A);
                                End;
                        end;
                        Inc(Glob);
                   end;
                   Conn_Klap.AQuery_Del.next;
                end;
        end
        Else
        Begin
           if (Polzov=4) then //���� � ������������ ID ������� 3 �������
           begin
                //� Tab2 ��� ���������� ����� Id����
               if not Conn_Ceh.mkQuerySelect2(
                'SELECT ����������.Id,������ ,�������, ����������.kol '+
                'FROM ���������� INNER JOIN ���������� ON ����������.IdSb = ����������.Id '+
                'WHERE (����������."Id����"='+#39+Tab2+#39+') AND (����������.IdSb<>'+#39+'0'+#39+') UNION '+

                'SELECT we.Id , ������ ,  �������, we.kol '+
                'FROM ���������� INNER JOIN '+
	              '(SELECT �������.Id,������ , ����������� �������, �������.kol '+
                'from �������) '+
	              'we  ON ����������.IdDet = we.Id '+
                'WHERE (����������.Id����='+#39+Tab2+#39+') AND (����������.IdSb='+#39+'0'+#39+')'
                 , ['����������']) then  //������
                exit;
                KOL_RECORD_Tab2:=Conn_Ceh.AQuery_Del.RecordCount;
           end;
           if (Polzov<4) then
           begin
                if not Conn_Klap.mkQuerySelect2(
                'Select * from %s Where (�����������='+#39+'������'+#39+
                ') AND ([Id��]='+#39+IDGP+#39+') AND (([Id��]='+#39+IDKO+#39+') OR ([Id��]='+#39+''+#39+'))'
                 , [Tab2]) then  //������
                exit;
                KOL_RECORD_Tab2:=Conn_Klap.AQuery_Del.RecordCount;
           end;
           if (Polzov=5) or (Polzov=6) then
           begin         //AND ([��������]<>'+#39+'99'+#39+') ��� ������ 99 �� ��� � ������� ������
                         //AND ([��������]<>'+#39+'55'+#39+') ��� ������ 55 �� ��� � ������� ������
                if not Conn_Ceh.mkQuerySelect2(
                'Select * from %s Where (�����������='+#39+'������'+#39+
                ') AND ([Id��]='+#39+IDGP+#39+') AND ([��������]<>'+#39+'99'+#39+
                ') AND ([��������]<>'+#39+'77'+#39+') AND ([��������]<>'+#39+'88'+#39+') AND ([��������]<>'+#39+'55'+#39+') AND (([Id��]='+#39+IDKO+
                #39+') or ([Id��]='+#39+#39+'))'
                 , [Tab2]) then  //������
                exit;
                KOL_RECORD_Tab2:=Conn_Ceh.AQuery_Del.RecordCount;
           end;
            if (Polzov=55) then
           begin
                if not Conn_Ceh.mkQuerySelect2(
                'Select * from %s Where (�����������='+#39+'������'+#39+
                ') AND ([Id��]='+#39+IDGP+#39+')  AND (([Id��]='+#39+IDKO+
                #39+') or ([Id��]='+#39+#39+'))'
                 , [Tab2]) then  //������
                exit;
                KOL_RECORD_Tab2:=Conn_Ceh.AQuery_Del.RecordCount;
           end;
           For X:=0 to KOL_RECORD_Tab2-1 Do
           Begin
             if (Polzov<4) then
             Begin
                Obozn:=Conn_Klap.AQuery_Del.FieldByName('�����������').AsString;
                Elem:=Conn_Klap.AQuery_Del.FieldByName('�������').AsString;
                Kol_Det:=Conn_Klap.AQuery_Del.FieldByName('����������').AsString; // AND (��������=1)
                                                //Dl,Sh,DL_R,Sh_R,Mat,Diam,Mass,Kol_G,Sh1,Sh2,Ie
                Dl :=Conn_Klap.AQuery_Del.FieldByName('�����').AsString;
                Sh :=Conn_Klap.AQuery_Del.FieldByName('������').AsString;
                DL_R  :=Conn_Klap.AQuery_Del.FieldByName('���������').AsString;
                Sh_R  :=Conn_Klap.AQuery_Del.FieldByName('����������').AsString;

                Diam  :=Conn_Klap.AQuery_Del.FieldByName('�������').AsString;
                Mass  :=Conn_Klap.AQuery_Del.FieldByName('�����').AsString;
                Kol_G :=Conn_Klap.AQuery_Del.FieldByName('��������').AsString;
                Sh1   :=Conn_Klap.AQuery_Del.FieldByName('�����������1').AsString;
                Sh2   :=Conn_Klap.AQuery_Del.FieldByName('�����������2').AsString;
                Ie    :=Conn_Klap.AQuery_Del.FieldByName('��').AsString;
                Mat   :=Conn_Klap.AQuery_Del.FieldByName('��������').AsString;

              end;
             if (Polzov=4) then   //VRAN
             Begin
                Obozn:=Conn_Ceh.AQuery_Del.FieldByName('�������').AsString;
                Elem:=Conn_Ceh.AQuery_Del.FieldByName('������').AsString;
                Kol_Det:=Conn_Ceh.AQuery_Del.FieldByName('Kol').AsString; // AND (��������=1)
              end;

              if (Polzov=5) or (Polzov=6) or (Polzov=55) then  //�/� +���
             Begin
                Obozn:=Conn_Ceh.AQuery_Del.FieldByName('�����������').AsString;
                Elem:=Conn_Ceh.AQuery_Del.FieldByName('�������').AsString;
                Kol_Det:=Conn_Ceh.AQuery_Del.FieldByName('����������').AsString; // AND (��������=1)
                Leg:=Conn_Ceh.AQuery_Del.FieldByName('��������').AsString; //���
                //Dl,Sh,DL_R,Sh_R,Mat,Diam,Mass,Kol_G,Sh1,Sh2,Ie
                Dl :=Conn_Ceh.AQuery_Del.FieldByName('�����').AsString;
                Sh :=Conn_Ceh.AQuery_Del.FieldByName('������').AsString;
                DL_R  :=Conn_Ceh.AQuery_Del.FieldByName('���������').AsString;
                Sh_R  :=Conn_Ceh.AQuery_Del.FieldByName('����������').AsString;

                Diam  :=Conn_Ceh.AQuery_Del.FieldByName('�������').AsString;
                Mass  :=Conn_Ceh.AQuery_Del.FieldByName('�����').AsString;
                Kol_G :=Conn_Ceh.AQuery_Del.FieldByName('��������').AsString;
                Sh1   :=Conn_Ceh.AQuery_Del.FieldByName('�����������1').AsString;
                Sh2   :=Conn_Ceh.AQuery_Del.FieldByName('�����������2').AsString;
                Ie    :=Conn_Ceh.AQuery_Del.FieldByName('��').AsString;
                Mat   :=Conn_Ceh.AQuery_Del.FieldByName('��������').AsString;
              end;

                if not Conn_Klap.mkQuerySelect3(
                'Select * from %s Where (�����������='+#39+Obozn+#39+') AND (����������=0)  '
                 , ['�����������']) then
                exit;
                if  Conn_Klap.AQuery3.RecordCount<>0 then
                Begin
                        SG2.RowCount:=SG2.RowCount+ 1;
                        sg2.Cells[0,Glob+1]:=Kol_Klap;//IntToStr(Glob+1);

                        sg2.Cells[1,Glob+1]:=ZAKAZ;//Conn_Klap.AQuery.FieldByName('�����').AsString;
                        sg2.Cells[2,Glob+1]:=Izdel;//Conn_Klap.AQuery.FieldByName('�������').AsString;
                        //�������� ���� ��������, ���� ���� �� ������ �� ����������
                        Conn_Klap.Check('Begin1', Kol_Klap+'  '+Kol_Det, 0);
                        sg2.Cells[3,Glob+1]:=FloatToStr(StrToInt(Kol_Klap)*StrToFloat(Kol_Det));
                        Conn_Klap.Check('End1', Kol_Klap+'  '+Kol_Det, 0);
                        sg2.Cells[4,Glob+1]:=Conn_Klap.AQuery3.FieldByName('�����������').AsString;
                        sg2.Cells[5,Glob+1]:=Conn_Klap.AQuery3.FieldByName('�������').AsString;
                        sg2.Cells[6,Glob+1]:=Nom;
                        sg2.Cells[7,Glob+1]:=BZ;
                        sg2.Cells[8,Glob+1]:=IDGP;
                        sg2.Cells[9,Glob+1]:=Conn_Klap.AQuery3.FieldByName('��������').AsString;
                        sg2.Cells[10,Glob+1]:=IDKO;
                        sg2.Cells[11,Glob+1]:=Leg;
                        sg2.Cells[12,Glob+1]:=Dl;
                        sg2.Cells[13,Glob+1]:=Sh;
                        sg2.Cells[14,Glob+1]:=DL_R;
                        sg2.Cells[15,Glob+1]:=Sh_R;
                        sg2.Cells[16,Glob+1]:=Diam;
                        sg2.Cells[17,Glob+1]:=Mass;
                        sg2.Cells[18,Glob+1]:=Kol_G;
                        sg2.Cells[19,Glob+1]:=Sh1;
                        sg2.Cells[20,Glob+1]:=Sh2;
                        sg2.Cells[21,Glob+1]:=Ie;
                        sg2.Cells[22,Glob+1]:=Mat;
                        Res_Teh:=Pos('rue',sg2.Cells[9,Glob+1]);
                        if (Res_Teh=0)  Then //�������� FALSE
                        begin
                                  Memo12.Add('�������� FALSE ����� '+sg2.Cells[1,Glob+1]+' * '+sg2.Cells[2,Glob+1]+' * ������ '+sg2.Cells[4,Glob+1]);
                                  Memo13.Add(sg2.Cells[4,Glob+1]);
                        end;
                        A:=13;
                        for y := 0 to High(Column) do
                        begin
                                if Column[y,0]<>'' then
                                Begin
                                        Stanok:=Column[y,0]; //�������� ����

                                        sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery3.FieldByName(Stanok).AsString;
                                        Res_Stan:=Pos('rue',sg2.Cells[9+(A+1),Glob+1]);
                                        Inc(A);

                                        sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery3.FieldByName(Stanok+'��').AsString; //�����������
                                        Och_D:=StrToFloat(sg2.Cells[9+(A+1),Glob+1]);
                                        Inc(A);

                                        sg2.Cells[9+(A+1),Glob+1]:='0.1';//Conn_Klap.AQuery3.FieldByName(Stanok+'��').AsString;//������������
                                        Dlit_D:=StrToFloat( sg2.Cells[9+(A+1),Glob+1]);//������������
                                        Inc(A);

                                        if (Res_Stan<>0) AND ((Och_D=0) OR (Dlit_D=0))  Then //������ TRUE
                                                                                             //� ����=0 ��� ���� =0 �����
                                        begin
                                                Memo12.Add('������ TRUE � ����=0 ��� ���� =0 ����� '+sg2.Cells[1,Glob+1]+' * '+sg2.Cells[2,Glob+1]+' * ������ '+sg2.Cells[4,Glob+1]);
                                                Memo13.Add(sg2.Cells[4,Glob+1]);
                                        end;

                                        sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery3.FieldByName('�\�'+Stanok).AsString;//�\�
                                        Inc(A);
                                        //sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery3.FieldByName(Stanok+'����').AsString; //����
                                        //Inc(A);
                                End;
                        end;
                        Inc(Glob);
                end;
                if Polzov<4 then
                  Conn_Klap.AQuery_Del.next;
                if (Polzov=4) or (Polzov=5) or (Polzov=55) or (Polzov=6) then
                  Conn_Ceh.AQuery_Del.next;
           end;
        end;
        if Polzov<4 then
          Conn_Klap.AQuery.next;
        if  (Polzov=4) or (Polzov=5) or (Polzov=55) or (Polzov=6) then
          Conn_Ceh.AQuery.next;  //gulyakova-ei
      end;
      Memo12_Count:=Memo12.Count;
      Memo12_Text:= Memo12.Text;
      God := FormatDateTime('yyyy', Now);
      Mes := FormatDateTime('mmmm', Now);
      if ( (Polzov=4) or (Polzov=5) or (Polzov=55) or (Polzov=6) ) AND (Memo12_Count<>0) then  //����� ����
      begin
        Dir :=Form1.Put_KTO+'\CKlapana\VRAN\' + God;
        CreateDir(Dir);

        Dir := Form1.Put_KTO+'\CKlapana\VRAN\' + God + '\' + Mes+'\';
        CreateDir(Dir);
      end;
      Dir :=Form1.Put_KTO+'\CKlapana\���������(�����)\' + God;
        CreateDir(Dir);

      Dir := Form1.Put_KTO+'\CKlapana\���������(�����)\' + God + '\' + Mes+'\';
        CreateDir(Dir);

      Dir :=Form1.Put_KTO+'\CKlapana\' + God;
        CreateDir(Dir);
      Dir := Form1.Put_KTO+'\CKlapana\' + God + '\' + Mes+'\';
        CreateDir(Dir);
      if (Polzov=0) AND (Memo12_Count=0)  Then //��������
      //��� ������ ���������� ,�������
      Begin
        Flag_Error:=0;
        Memo12.Destroy;
        Exit;
      end;
        if (Polzov=0) AND (Memo12_Count<>0) then  //�����   ��������
        Begin
                MessageDlg('�� ��� ������ ����������! ������ � ����� � ���������.', mtError,
                        [mbOk], 0);
                Res_Tab:=Pos('���',Tab2);
                if Res_Tab=0 Then           // PAnsiChar(P
                Begin
                        Memo12.SaveToFile(Form1.Put_KTO+'\CKlapana\���������(�����)\'+God+'\'+Mes+'\'+Nom+'.txt');
                        Nam :=  Form1.Put_KTO+'\CKlapana\���������(�����)\'+God+'\'+Mes+'\'+Nom+'.txt';
                End;
                if Res_Tab<>0 Then//���������
                Begin
                        Memo12.SaveToFile(Form1.Put_KTO+'\CKlapana\'+God+'\'+Mes+'\'+Nom+'.txt');
                        Nam := Form1.Put_KTO+'\CKlapana\'+God+'\'+Mes+'\'+Nom+'.txt';
                End;
                ShellExecute(0, nil, PChar(Nam), nil, nil, SW_SHOWNORMAL);
                File_Name:= Nam ;
                Flag_Error:=2;
                Memo12.Destroy;
                if fileName<>'' then
                begin
                    if not Conn_Klap.mkQueryDelete( 'DELETE FROM %s WHERE (�����='+Nom+') AND (������������='+#39+fileName+#39+')'
                                , ['���������']) then
                    exit;
                end
                Else
                begin
                    if not Conn_Klap.mkQueryDelete( 'DELETE FROM %s WHERE (�����='+Nom+') AND (������������='+#39+ConvertDat1(Dat_Glob)+#39+')'
                                , ['���������']) then
                    exit;
                end;
                Result := False;
                Exit;
        End;
        if ((Polzov<4) AND (Polzov<>0)) AND (Memo12_Count<>0) then  //�����
        Begin
                MessageDlg('�� ��� ������ ����������! ������ � ����� � ���������.', mtError,
                        [mbOk], 0);
                Res_Tab:=Pos('���',Tab2);
                if Res_Tab=0 Then
                Begin
                        Memo12.SaveToFile(Form1.Put_KTO+'\CKlapana\���������(�����)\'+God+'\'+Mes+'\'+Nom+'.txt');
                        Nam :=  Form1.Put_KTO+'\CKlapana\���������(�����)\'+God+'\'+Mes+'\'+Nom+'.txt';
                End;
                if Res_Tab<>0 Then//���������
                Begin
                        Memo12.SaveToFile(Form1.Put_KTO+'\CKlapana\'+God+'\'+Mes+'\'+Nom+'.txt');
                        Nam :=  Form1.Put_KTO+'\CKlapana\'+God+'\'+Mes+'\'+Nom+'.txt';
                End;
                ShellExecute(0, nil, PChar(Nam), nil, nil, SW_SHOWNORMAL);
                File_Name:= Nam ;
                Flag_Error:=2;
                Memo12.Destroy;
                if fileName<>'' then
                begin
                    if not Conn_Klap.mkQueryDelete( 'DELETE FROM %s WHERE (�����='+Nom+') AND (������������='+#39+fileName+#39+')'
                                , ['���������']) then
                    exit;
                end
                Else
                begin
                    if not Conn_Klap.mkQueryDelete( 'DELETE FROM %s WHERE (�����='+Nom+') AND (������������='+#39+ConvertDat1(Dat_Glob)+#39+')'
                                , ['���������']) then
                    exit;
                end;
                Result := False;
                Exit;
        End;

        if ( (Polzov=4) or (Polzov=5) or (Polzov=55) or (Polzov=6) ) AND (Memo12_Count<>0) then  //����� ����
        Begin
                MessageDlg('�� ��� ������ ����������! ������ � ����� � ���������.', mtError,
                        [mbOk], 0);

                Memo12.SaveToFile(Form1.Put_KTO+'\CKlapana\VRAN\'+God+'\'+Mes+'\'+Nom+'.txt');
                Nam :=  Form1.Put_KTO+'\CKlapana\VRAN\'+God+'\'+Mes+'\'+Nom+'.txt';

                ShellExecute(0, nil, PChar(Nam), nil, nil, SW_SHOWNORMAL);
                File_Name:= Nam ;
                Flag_Error:=2;
                Memo12.Destroy;
                if fileName<>'' then
                begin
                    if not Conn_Klap.mkQueryDelete( 'DELETE FROM %s WHERE (�����='+Nom+') AND (������������='+#39+fileName+#39+')'
                                , ['���������']) then
                    exit;
                end
                Else
                begin
                    if not Conn_Klap.mkQueryDelete( 'DELETE FROM %s WHERE (�����='+Nom+') AND (������������='+#39+ConvertDat1(Dat_Glob)+#39+')'
                                , ['���������']) then
                    exit;
                end;
                Result := False;
                Exit;
        End;
      except

        exit;
    end;
    Memo12.Destroy;
    Memo13.Destroy;
    if fileName<>'' then
    begin
       if not Conn_Klap.mkQueryDelete( 'DELETE FROM %s WHERE (�����='+Nom+') AND (������������='+#39+fileName+#39+')'
                                , ['���������']) then
    exit;
    end
    Else
    begin
       if not Conn_Klap.mkQueryDelete( 'DELETE FROM %s WHERE (�����='+Nom+') AND (������������='+#39+ConvertDat1(Dat_Glob)+#39+')'
                                , ['���������']) then
    exit;
    end;

    Result := true;
    btn1Click(fileName) ;
end;
function Osnova_Main.btn1Click(fileName:String):Boolean;
        var Dat_S1,IDGP,IDKO,Obozn,Detal,Str,Elem,Stanok,Str1,Str2,Izdel,Kol_Klap,Kol_Det,Zakaz,Nomer,Max:string;
        I,X,Y,A,Glob,Col,Col_Nom,
        Kol_Klap_I,Kol_Det_I,Res1,res2,Res:Integer;
        B,D:Boolean;                   // ,S1
        DEF_CatalogName,MSSQL_CONN_STR,S,Razmernost,Nom_Kol,Znach,Str_Stanok,Err,BZ:string;
        TempTableQ:TADOQuery;
        ADOConnection4:TADOConnection;
        Dat_S,Nedely,Kol,Dir,SQLServer,Dat_Z,Nowa_S,Dlit,Dlit_Ob,NCH,fileName1,Kod,Dl,Sh,DL_R,Sh_R,Mat,Diam,Mass,Kol_G,Sh1,Sh2,Ie:string;
        Dat, Dat1,Dat_Zag,nowa:TDate;
        Ocheredn,Res_True:Integer;
        Dliteln,
        NC_D:Double;
        //StrToCol:array[0..100,0..5] of string;//[�������, ������������, ��� ������������, ����] ��� ����� ������
        Kol_Stan:Integer;
        IPs : TStringList;
        Res_Trum,II:Integer;
        Dl_Trum,DD:Double;
        Memo11:TStringList;
        leg:String;
   begin
       SyStem.SysUtils.FormatSettings.DecimalSeparator :=('.');
       Memo11:=TStringList.Create;
       Memo11.Clear;
       Glob:=0;
       A:=0;
      //==================================
      S:='';
      Razmernost:='';
      A:=0;
      fileName1:=fileName;
      //Dat:=dtp1.Date;
      //Dat_S:=FormatDateTime('dd.mm.YYYY',Dat);
      //Dat_S1:=FormatDateTime('dd.mm.YYYY',Dat-1);
      for y := 0 to High(Column) do
      begin
        if Column[y,0]<>'' then
        Begin
                Str_Stanok:=Str_Stanok+Column[y,0]+','+Column[y,0]+'��,'+Column[y,0]+'��,[�\�'+Column[y,0]+'],'+Column[y,0]+'����,'+Column[y,0]+'�����,';
                Nom_Kol:=IntToStr(StrToInt(Column[y,2])-2);
                Razmernost:=Razmernost+Column[y,0]+' BIT,'+Column[y,0]+'�� nvarchar(50),'+Column[y,0]+'�� nvarchar(50),[�\�'+Column[y,0]+'] nvarchar(50),'+Column[y,0]+'���� DateTime,'+Column[y,0]+'����� nvarchar(50),';  //DateTime����� bit,������� nvarchar(50)
        End;
      end;
      //sg2.Group(4);
      //sg2.Group(5);
      //SG2.GroupSum(4);
     // SG2.ClearAll;
      S:= SG2.Cells[7,0];
      //++++++++++++++++++++++++++++++++++++
      Dir:=ExtractFileDir(ParamStr(0));
      IPs := TStringList.Create;
      IPs.Clear;
      IPs.LoadFromFile(Dir+'\ConnKlap.ini');
        //Memo1.Lines.LoadFromFile(Dir+'\ConnKlap.ini');

        DEF_CatalogName :=IPs.Strings[0];
        SQLServer:=IPs.Strings[1];
        MSSQL_CONN_STR :=IPs.Strings[4];
       { 'Provider=SQLOLEDB.1;Packet Size = 4096;Password=111;' +
        'Persist Security Info=True;User Id=TestUser;' +
        'Initial Catalog='+ DEF_CatalogName +';Data Source=DINAMO\'+SQLServer; }
        ADOConnection4:=TADOConnection.Create(Nil);
        ADOConnection4.LoginPrompt:=False;
        ADOConnection4.Connected:=False;
        ADOConnection4.ConnectionString := MSSQL_CONN_STR;
        ADOConnection4.Connected:=True;
         //������� ������
        Y:=Length(Razmernost);
        Delete(Razmernost,Y,1);
        //�������� ��������� �������
        TempTableQ:=TADOQuery.Create(nil);
        TempTableQ.Connection:=ADOConnection4;
        //�������� ���� ��������, ���� ���� �� ������ �� ����������
        TempTableQ.SQL.Text:='CREATE TABLE #ClientToDBFL (����� nvarchar(100),������� nvarchar(100),����� nvarchar(100),'+
        '����������� nvarchar(100) ,������� nvarchar(100),���������� nvarchar(100),�� nvarchar(100),'+
        'Id�� nvarchar(100),Id�� nvarchar(100),������� nvarchar(100),Leg nvarchar(100),'+
        '����� nvarchar(100),������ nvarchar(100),��������� nvarchar(100),���������� nvarchar(100),'+
        '������� nvarchar(100),����� nvarchar(100),�������� nvarchar(100),�����������1 nvarchar(100),'+
        '�����������2 nvarchar(100),�� nvarchar(100),�������� nvarchar(100),'+Razmernost+')';
                //Dl,Sh,DL_R,Sh_R,Mat,Diam,Mass,Kol_G,Sh1,Sh2,Ie
        {sg2.Cells[12,0]:='�����';
        sg2.Cells[13,0]:='������';
        sg2.Cells[14,0]:='���������';
        sg2.Cells[15,0]:='����������';

        sg2.Cells[16,0]:='�������';
        sg2.Cells[17,0]:='�����';
        sg2.Cells[18,0]:='��������';
        sg2.Cells[19,0]:='�����������1';
        sg2.Cells[20,0]:='�����������2';
        sg2.Cells[21,0]:='��';
        sg2.Cells[22,0]:='��������'; }
        try
                TempTableQ.ExecSQL;
        except
                Err:='TempTableQ.ExecSQL;';
        end;
        //------------------------------------------------
      //������� ������
      Y:=Length(Str_Stanok);
      Delete(Str_Stanok,Y,1);

      //-------------------------------------------------
        for Y:=0 to SG2.RowCount-1 do
        begin
         if Y >0 then
         begin                        //  AND (SG2.Cells[2,Y]='')
                if (SG2.Cells[1,Y]<>'') then
                begin
                   //------------------------------------------------
                   S1:='';
                   Dliteln:=0;
                   Ocheredn:=0;
                   Col_Nom:=0;
                   A :=StrToInt(Column[0,2])+16;
                   i :=0;
                   Kol_Stan:=0;
                   fillchar(StrToCol,sizeof(StrToCol),0);//������� �������
                   for X := 0 to High(Column) do  //������ ������� �������
                   begin
                        if Column[X,2]<>'' then
                        Begin   //'����,������,������,[�\�����],
                                //'False','0','0','0',
                                i :=0;
                                Col_Nom:=StrToInt(Column[X,2])-1;
                               // S1:=S1+#39+SG2.Cells[A,Y]+#39+','; //True  0
                                StrToCol[X, I ]:=SG2.Cells[A,Y];
                                Inc(A);
                                Inc(i);
                                //
                               // S1:=S1+#39+SG2.Cells[A,Y]+#39+','; //Ocheredn 1
                                StrToCol[X,i]:=SG2.Cells[A,Y];
                                Inc(A);
                                Inc(i);
                                //
                               // S1:=S1+#39+SG2.Cells[A,Y]+#39+','; //Dliteln 2
                                StrToCol[X,i]:=SG2.Cells[A,Y];
                                Inc(A);
                                Inc(i);
                                //
                                //S1:=S1+#39+SG2.Cells[A,Y]+#39+','; //�\�  3
                                StrToCol[X,i]:=SG2.Cells[A,Y];
                                Inc(A);
                                Inc(i);
                                //
                                //S1:=S1+#39+'0'+#39+',';  // ����
                                StrToCol[X,i]:='0';
                                Inc(i);
                                //
                                StrToCol[X,i]:=Column[X,0];
                                Inc(i);
                                //
                        End;
                   end;
                   Kol_Klap:=SG2.Cells[0,Y];
                   if Kol_Klap='' Then
                   Kol_klap_I:=0
                   Else
                   Kol_Klap_I:=StrToInt(Kol_Klap);

                   ///()))))))))))))))))))))))(((((((((((((((((((((
                   Btn11Click(fileName1); ///()))))))))))))))))))))))(((((((((((((((((((((
                   ///()))))))))))))))))))))))(((((((((((((((((((((
                   //������� ������
                   i:=Length(S1);
                   Delete(S1,i,1);

                   Zakaz:=SG2.Cells[1,Y];
                   Obozn:=SG2.Cells[4,y];
                   Izdel:=SG2.Cells[2,y];
                   res:=Pos('��-2510-1-3',Obozn);
                   if Res<>0 then
                      Kol_Det:='1';

                                                   {if not Conn_Klap.mkQuerySelect6('Select *  From %s Where (�����������='+#39+'%s'+#39+
                                ') ', ['���������',Obozn]) then //������
                                exit;   }
                                //�� 410.00.01.002-400 ������ �� �� 410.00.01.001-400  ������ 21.01.2019
                                {Res1:=Pos('���-', Izdel);
                                Res2:=Pos('-2*�-', Izdel);
                                res:=Pos('�� 410.00.01.002-',Obozn);
                                if (Res<>0) and (Res1<>0) AND (res2<>0) then
                                begin
                                  //StringReplace(Dlin_Raz_S,'.',',',[rfReplaceAll]);
                                  Res:=Pos('.002-',Obozn);
                                  if Res<>0 then
                                  begin
                                    Delete(Obozn,Res,5);
                                    S:=  '.001-';
                                    Insert(S,Obozn,Res);
                                  end;
                                end;
                                // �� 463.00.00.006 ����� �� �� 463.00.00.006-01   ������ 21.01.2019
                                res:=Pos('�� 463.00.00.006',Obozn);
                                if  (Res<>0) and (Res1<>0) AND (res2<>0) then
                                begin
                                     Obozn:='�� 463.00.00.006-01';
                                end;                  //Dl,Sh,DL_R,Sh_R,Mat,Diam,Mass,Kol_G,Sh1,Sh2,Ie
}
                   Kol_Det:=SG2.Cells[3,y];
                   Elem:=SG2.Cells[5,y];
                   Nomer:=SG2.Cells[6,y];
                   BZ:=SG2.Cells[7,y];
                   IDGP:=SG2.Cells[8,y];
                   IDKO:=SG2.Cells[10,y];
                   Leg:=(SG2.Cells[11,y]);

                   Dl :=SG2.Cells[12,y];
                   Sh :=SG2.Cells[13,y];
                   DL_R :=SG2.Cells[14,y];
                   Sh_R :=SG2.Cells[15,y];
                   Diam :=SG2.Cells[16,y];
                   Mass :=SG2.Cells[17,y];
                   Kol_G:=SG2.Cells[18,y];
                   Sh1  :=SG2.Cells[19,y];
                   Sh2  :=SG2.Cells[20,y];
                   Ie   :=SG2.Cells[21,y];
                   Mat  :=SG2.Cells[22,y];
                   //���-�� ������� ��� �������� �� ��� ��������!!!!

                   TempTableQ.SQL.Text:='INSERT INTO #ClientToDBFL '+
                        '(�����,�������,�����������,�������,'+
                        '����������,�����,��,Id��, Id��,�������,Leg,'+
                        '�����,������,���������,����������,�������,�����,��������,�����������1,�����������2,��,��������,'

                        +Str_Stanok+') Values ('
                        +#39+Zakaz+#39+
                        ','+#39+Izdel+#39+
                        ','+#39+Obozn+#39+
                        ','+#39+Elem+#39+
                        ','+#39+Kol_Det+#39+
                        ','+Nomer+
                        ','+#39+BZ+#39+
                        ','+#39+IDGP+#39+ ','+#39+IDKO+#39+
                        ','+#39+Kol_Klap+#39+
                        ','+#39+(Leg)+#39+

                        ','+#39+Dl+#39+
                        ','+#39+Sh+#39+
                        ','+#39+DL_R+#39+
                        ','+#39+Sh_R+#39+
                        ','+#39+Diam+#39+
                        ','+#39+Mass+#39+
                        ','+#39+Kol_G+#39+
                        ','+#39+Sh1+#39+
                        ','+#39+Sh2+#39+
                        ','+#39+Ie+#39+
                        ','+#39+Mat+#39+
                            ','+S1+
                        ')';

                        try
                        TempTableQ.ExecSQL;
                        except
                        Err:='1222';
                        end;
                end;

         end;
        end;
        //����������� �� TRUE
        S1:='';
        A:=1;
        Glob:=1;
        sg3 := TStringGrid.Create(Nil );
        Clear_StringGrid(sg3);
    SG3.ColCount:=29;
    sg3.Cells[0,0]:='�������';
    sg3.Cells[1,0]:='�����';
    sg3.Cells[2,0]:='�������';
    sg3.Cells[3,0]:='�����';

    sg3.Cells[4,0]:='�����������';
    sg3.Cells[5,0]:='�������';
    sg3.Cells[6,0]:='����������';
    sg3.Cells[7,0]:='�����������'; //�����������
    sg3.Cells[8,0]:='������������';//������������
    sg3.Cells[9,0]:='�\�';//�\�
    sg3.Cells[10,0]:='��';
    sg3.Cells[11,0]:='Id��';
    sg3.Cells[12,0]:='������';
    sg3.Cells[13,0]:='������������';
    sg3.Cells[14,0]:='���������';
    sg3.Cells[15,0]:='���� �����';
    sg3.Cells[16,0]:='Id��';
    sg3.Cells[17,0]:='Leg';
            //Dl,Sh,DL_R,Sh_R,Mat,Diam,Mass,Kol_G,Sh1,Sh2,Ie
        sg3.Cells[18,0]:='�����';
        sg3.Cells[19,0]:='������';
        sg3.Cells[20,0]:='���������';
        sg3.Cells[21,0]:='����������';
        sg3.Cells[22,0]:='�������';
        sg3.Cells[23,0]:='�����';
        sg3.Cells[24,0]:='��������';
        sg3.Cells[25,0]:='�����������1';
        sg3.Cells[26,0]:='�����������2';
        sg3.Cells[27,0]:='��';
        sg3.Cells[28,0]:='��������';
    //Memo12:=TStringList.Create;
    //Memo12.Clear;
        for Y:=0 to High(Column) do  //������ ������� �������
        begin
                if Column[y,0]<>'' then
                Begin
                        TempTableQ.SQL.Text:='SELECT *  FROM #ClientToDBFL'+
                        ' WHERE '+Column[y,0]+'='+#39+'True'+#39 ;
                        TempTableQ.Open;
                        //+++++++++++++++++++++++++++++++++++++++
                        TempTableQ.First;
                        if TempTableQ.RecordCount<>0 Then
                        Begin
                                sg3.RowCount:= sg3.RowCount+1 ;
                                sg3.Cells[0,A]:=Column[y,0];
                                Stanok:=Column[y,0]; //�������� ����
                                A:=A+1;

                          for I := 0 to TempTableQ.RecordCount - 1 do
                          begin
                                sg3.RowCount:= sg3.RowCount+1 ;
                                Str:=sg3.Cells[0,A];
                                sg3.Cells[0,A]:=TempTableQ.FieldByName('�������').AsString;//IntToStr(Glob);
                                sg3.Cells[1,A]:=TempTableQ.FieldByName('�����').AsString;
                                sg3.Cells[2,A]:=TempTableQ.FieldByName('�������').AsString;
                                sg3.Cells[3,A]:=TempTableQ.FieldByName('�����').AsString;

                                sg3.Cells[4,A]:=TempTableQ.FieldByName('�����������').AsString;
                                sg3.Cells[5,A]:=TempTableQ.FieldByName('�������').AsString;
                                Kol_Det:=TempTableQ.FieldByName('����������').AsString;
                                II:=Round(StrToFloat(Kol_Det));
                                Kol_Det:=IntToStr(II);
                                sg3.Cells[6,A]:= Kol_Det;
                                if Kol_Det='' Then
                                Kol_Det_I:=0
                                Else
                                Kol_Det_I:=StrToInt(Kol_Det);

                                sg3.Cells[7,A]:=TempTableQ.FieldByName(Stanok+'��').AsString; //�����������
                                Ocheredn:=StrToInt( sg3.Cells[7,A]);
                                sg3.Cells[8,A]:=TempTableQ.FieldByName(Stanok+'��').AsString;//������������
                                Dliteln:=StrToFloat( StringReplace(sg3.Cells[8,A],',','.',[rfReplaceAll]));
                                sg3.Cells[9,A]:=TempTableQ.FieldByName('�\�'+Stanok).AsString;//�\�
                                NCH:=StringReplace(sg3.Cells[9,A],',','.',[rfReplaceAll]);
                                if NCh='' Then
                                NC_D:=0
                                Else
                                NC_D:=StrToFloat(NCH);
                                sg3.Cells[10,A]:=TempTableQ.FieldByName('��').AsString;
                                sg3.Cells[11,A]:=TempTableQ.FieldByName('Id��').AsString;
                                sg3.Cells[12,A]:=Column[y,0];
                                sg3.Cells[13,A]:=Dat_Glob;

                                Dat_Z :=TempTableQ.FieldByName(Stanok+'����').AsString;
                                Dat_Zag:= StrToDate(Dat_Z);
                                Nowa_S:=FormatDateTime('dd.mm.YYYY',Now);
                                Nowa:=StrToDate(Nowa_s);
                                //28.05.2018 �������� ��������
                               { if Dat_Zag<=Nowa then //���� ������ ��������� ������ �����������
                                begin
                                        MessageDlg(Nom_G+'  ���� ������ ��������� ������ �����������!' + #10#13 +
                                        '���������� �������� ���� ������������ �� ' + FloatToStr(Nowa-Dat_Zag)+' ����!' ,mtInformation, [mbOk], 0);
                                        Flag_Error:=1;
                                        Exit;
                                end; }
                                sg3.Cells[14,A]:=TempTableQ.FieldByName(Stanok+'����').AsString;
                                sg3.Cells[15,A]:=TempTableQ.FieldByName(Stanok+'�����').AsString;
                                sg3.Cells[16,A]:=TempTableQ.FieldByName('Id��').AsString;
                                sg3.Cells[17,A]:=TempTableQ.FieldByName('Leg').AsString;
                                                //Dl,Sh,DL_R,Sh_R,Mat,Diam,Mass,Kol_G,Sh1,Sh2,Ie
                                sg3.Cells[18,A] :=TempTableQ.FieldByName('�����').AsString;
                                sg3.Cells[19,A] :=TempTableQ.FieldByName('������').AsString;
                                sg3.Cells[20,A]  :=TempTableQ.FieldByName('���������').AsString;
                                sg3.Cells[21,A]  :=TempTableQ.FieldByName('����������').AsString;

                                sg3.Cells[22,A]  :=TempTableQ.FieldByName('�������').AsString;
                                sg3.Cells[23,A]  :=TempTableQ.FieldByName('�����').AsString;
                                sg3.Cells[24,A] :=TempTableQ.FieldByName('��������').AsString;
                                sg3.Cells[25,A]   :=TempTableQ.FieldByName('�����������1').AsString;
                                sg3.Cells[26,A]   :=TempTableQ.FieldByName('�����������2').AsString;
                                sg3.Cells[27,A]    :=TempTableQ.FieldByName('��').AsString;
                                sg3.Cells[28,A]   :=TempTableQ.FieldByName('��������').AsString;
                                Dlit:=StringReplace(sg3.Cells[8,A],',','.',[rfReplaceAll]);
                                Dlit_Ob:=StringReplace(sg3.Cells[15,A],',','.',[rfReplaceAll]);

                                //++++++++++++++++++++++++++++++++++++++++++++++++++
                                if sg3.Cells[14,A]='' then
                                        Dat_S1:='NULL'
                                else
                                        DaT_S1:=#39+ConvertDat1(SG3.Cells[14,A])+#39;
                                                        //   AND (��������=1)
                                if not Conn_Klap.mkQuerySelect3(
                                'Select * from %s Where (�����������='+#39+Obozn+#39+')'
                                , ['�����������']) then
                                exit;
                                //�������� �� �����
                                {if not Conn_Klap.mkQuerySelect4(
                                'Select * from %s Where (�����='+#39+sg3.Cells[1,A]+#39+') AND '+
                                '(�������='+#39+sg3.Cells[2,A]+#39+') AND (�����='+#39+sg3.Cells[3,A]+#39+') AND '+
                                '(�����������='+#39+sg3.Cells[4,A]+#39+') AND (�������='+#39+sg3.Cells[5,A]+#39+') AND (Id��='+#39+sg3.Cells[11,A]+#39+')'+
                                ' AND (������='+#39+sg3.Cells[12,A]+#39+')'
                                , ['���������']) then
                                exit;
                                if  (Conn_Klap.AQuery4.RecordCount<>0) Then
                                Begin
                                  if not Conn_Klap.mkQueryDelete( 'DELETE FROM %s WHERE (�����='+#39+sg3.Cells[1,A]+#39+') AND '+
                                '(�������='+#39+sg3.Cells[2,A]+#39+') AND (�����='+#39+sg3.Cells[3,A]+#39+') AND '+
                                '(�����������='+#39+sg3.Cells[4,A]+#39+') AND (�������='+#39+sg3.Cells[5,A]+#39+') AND (Id��='+#39+sg3.Cells[11,A]+#39+')'+
                                ' AND (������='+#39+sg3.Cells[12,A]+#39+')'
                                , ['���������']) then
                                        exit;
                                end; }
                            if  (Conn_Klap.AQuery3.RecordCount<>0)  then
                            Begin
                                NCH:=StringReplace(FloatToStr(NC_D*Kol_Det_I),',','.',[rfReplaceAll]);

                                if not Conn_Klap.mkQueryInsert(
                                'INSERT INTO ��������� '+
                                '(�����,�������,�����,�����������,�������,'+
                                '����������,�����������,������������,��,��,Id��,������,������������,���������,���������������,�������,Id��,'+
                                '����,������,��������,�����,������,���������,����������,�������,�����,��������,�����������1,�����������2,��,��������) Values ('
                                +#39+sg3.Cells[1,A]+#39+
                                ','+#39+sg3.Cells[2,A]+#39+
                                ','+#39+sg3.Cells[3,A]+#39+
                                ','+#39+sg3.Cells[4,A]+#39+
                                ','+#39+sg3.Cells[5,A]+#39+
                                ','+#39+sg3.Cells[6,A]+#39+
                                ','+#39+sg3.Cells[7,A]+#39+
                                ','+#39+Dlit+#39+
                                ','+#39+NCH+#39+
                                ','+#39+sg3.Cells[10,A]+#39+
                                ','+#39+sg3.Cells[11,A]+#39+
                                ','+#39+sg3.Cells[12,A]+#39+
                                ','+#39+ConvertDat1(sg3.Cells[13,A])+#39+
                                ','+Dat_S1+
                                ','+#39+Dlit_Ob+#39+
                                ','+#39+sg3.Cells[0,A]+#39+
                                ','+#39+sg3.Cells[16,A]+#39+
                                ','+#39+IntToStr(Cvet1)+#39+
                                ','+#39+'0'+#39+
                                ','+#39+sg3.Cells[17,A]+#39+

                                ','+#39+sg3.Cells[18,A]+#39+
                                ','+#39+sg3.Cells[19,A]+#39+
                                ','+#39+sg3.Cells[20,A]+#39+
                                ','+#39+sg3.Cells[21,A]+#39+
                                ','+#39+sg3.Cells[22,A]+#39+
                                ','+#39+sg3.Cells[23,A]+#39+
                                ','+#39+sg3.Cells[24,A]+#39+
                                ','+#39+sg3.Cells[25,A]+#39+
                                ','+#39+sg3.Cells[26,A]+#39+
                                ','+#39+sg3.Cells[27,A]+#39+
                                ','+#39+sg3.Cells[28,A]+#39+
                                ')',[] )then
                                Exit;
                                //+++++++++++++++++++++++++++++++++++++++++++
                                A:=A+1;
                                Glob:=Glob+1;
                            End;
                                TempTableQ.Next;

                          end;
                        end;

                end;
        end;
        TempTableQ.Close;
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        if not Conn_Klap.mkQuerySelect('Select ��������� From ��������� Where (�����='+
        #39+Nomer_Glob+#39+') Group By ���������',[]) Then
        Exit;
        For I:=0 To  Conn_Klap.AQuery.RecordCount-1 Do
        Begin
              Dat_S:= ConvertDat1(Conn_Klap.AQuery.FieldByName('���������').AsString);

              Memo11.Add('');
              Memo11.Add(Dat_S);
              if not Conn_Klap.mkQuerySelect3('Select ������, SUM(��) As SNH From ��������� Where (���������='+
                 #39+Dat_S+#39+') AND (������<>'+#39+'������'+#39+') AND (������<>'+#39+'Trumph'+#39+') Group by ������',[]) Then
                Exit;

              For Y:=0 To  Conn_Klap.AQuery3.RecordCount-1 Do
              Begin
                   Memo11.Add(Conn_Klap.AQuery3.FieldByName('������').AsString+'  �\�: '+Conn_Klap.AQuery3.FieldByName('SNH').AsString);
                   Inc(A);
                   Conn_Klap.AQuery3.Next;
              end;
              Conn_Klap.AQuery.Next;
        end;
        SNCH:=Memo11.Text;
        //================================================
        TempTableQ.SQL.Text:='DROP TABLE #ClientToDBFL';
        try
                TempTableQ.ExecSQL;
        except
        end;

   End;
//++++++++++++++++++++++++++++++++++++++++++++++++++
Function Osnova_Main.btn11Click(Err2:String):Boolean;
var //StrToCol:array[0..100,0..6] of string;
//[TRUE,�������, ������������, �/�, ����,Stanok] ��� ����� ������
I,X,J,y,A,Res:Integer;
Res_True,min_Col,Col,Col_1,
Col_Grid,Res1:Integer;
Izdel,Nowa_S:string;
Dat_Plan,Dat_Zag,Nowa:TDate;
Dat_P,Dat_Z:string;
Dlit,Dlit_Min,Dlit_Obsh,Dlit_Max,
Min,Och_Min,Och_Max,Dlit_ObshO,Ocher,Ocher_Pred:Double;
//SREZ1: array[0..10,0..100] of string;

SKPU :array[0..10,0..100] of string;

//Sort: TStringList;
Stanok,Stanok_Posl:String;
Kol_Smen,Max,F1,Den,Den_pred,Celoe,Res_Trum:Integer;
Chas,H1,H2,H3,Den1,Dlit_Trum:Double;
SREZ1:TAdvStringGrid;
//AD1:TAdvStringGrid;
Chet:Integer;
begin
     X:=0;
     J:=0;
     Dlit:=0;
     Dlit_Obsh:=0;
     Col_Grid:=97;
     Col_1:=0;
     Max:=0;
     //Sort:=TStringList.Create;
     SREZ1 := TAdvStringGrid.Create(Nil );
     //AD1 := TAdvStringGrid.Create(Nil );
     fillchar(Temp,sizeof(Temp),0);//������� �������
     fillchar(SKPU,sizeof(SKPU),0);//������� �������
     Form1.Memo13.Clear;
     //++������� � ������������ ������ 1 ����
    { for I := 0 to High(StrToCol) do
     begin
        Res_True:=Pos('rue',StrToCol[i,0]);
        Res_Trum:=Pos('Trumph',StrToCol[i,5]);
        if (Res_Trum<>0) AND (Res_True<>0) then  //���
        begin
                Dlit_Min:=StrToFloat(StrToCol[i,2]);
                StrToCol[i,2]:=FloatToStr(Dlit_Min+1);
        end;
     end; }
     Dlit_Min:=0;
     //++++++++++++++++++++++++++++++++���� ����������� �� �����������
     for I := 0 to High(StrToCol) do
     begin
        Res_True:=Pos('rue',StrToCol[i,0]);
        if Res_True<>0 then
        begin
                min := StrToInt(StrToCol[i,1]); // ������ �� �������
                Och_Min:= StrToInt(StrToCol[i,1]);
                Col:=I;
                Dlit_Min:=StrToFloat(StrToCol[i,2]);
                Break;
        end;
     end;
     X:=0;
     A:=0;
     Y:=0;
     SREZ1.ColCount:=8;//AD   SREZ
     SREZ1.RowCount:=1;
     //Form1.SKPU.ColCount:=10; //AD1
     //Form1.SKPU.RowCount:=1;

     Dlit_Obsh:=0;
     for I := 0 to High(StrToCol) do
     begin
        Res_True:=Pos('rue',StrToCol[i,0]);
        if Res_True<>0 then
      begin
        Stanok_Posl:=StrToCol[i,5];
        For y:=0 To High(Smena) do
        Begin
          Stanok:=Smena[y,0];
          Res:=AnsiCompareStr(Stanok_Posl,Stanok);
          if Res=0 Then
          Begin
                Kol_Smen:=StrToInt(Smena[y,1]);
                Chas:=StrToFloat(Smena[y,2]);
                SREZ1.Cells[0,A]:=StrToCol[i,1]; //Ocher
                Ocher:= StrToInt(StrToCol[i,1]);

                SREZ1.Cells[1,A]:=Stanok_Posl;   //Stanok
                SREZ1.Cells[2,A]:=StrToCol[i,2];//Dlit
                Dlit_Min:=StrToFloat(StrToCol[i,2]);//*8;// 0,5*8=4 - ��� �����
                                                    //1*8=8 ������ �����
                Dlit_Obsh:=Dlit_Obsh+ (StrToFloat(StrToCol[i,2]));//� �����

                SREZ1.Cells[3,A]:=Smena[y,1];//Kol_Smen
                Kol_Smen:=StrToInt(Smena[y,1]);

                SREZ1.Cells[4,A]:=Smena[y,2];//Chas
                SREZ1.Cells[5,A]:='0';//Den
                SREZ1.Cells[6,A]:=StrToCol[I,3];//�\�
                SREZ1.Cells[7,A]:=StrToCol[I,0];//TRUE
                SREZ1.RowCount:=SREZ1.RowCount+1;
                Inc(A);
          end;
        End;
      End;
     End;
     F1:=0;
     SREZ1.Cells[5,0]:='1';
     //AD.Sort(0,sdAscending);//sdDescending �� ��������
     SREZ1.SortByColumn(0);
     Dat_Plan:=StrToDate(Dat_Glob);
     A:=0;
     For I:=0 To SREZ1.RowCount-1 Do
     Begin
       if SREZ1.Cells[0,i]<> '' Then
       Begin
           {Form1.SKPU.Cells[0,A]:=Form1.SREZ1.Cells[0,i]; //Ocher
           Form1.SKPU.Cells[1,A]:=Form1.SREZ1.Cells[1,i];   //Stanok
           Form1.SKPU.Cells[2,A]:=Form1.SREZ1.Cells[2,i];//Dlit
           Form1.SKPU.Cells[3,A]:=Form1.SREZ1.Cells[3,i];//Kol_Smen
           Form1.SKPU.Cells[4,A]:=Form1.SREZ1.Cells[4,i];//Chas
           Form1.SKPU.Cells[5,A]:='0';//Den
           Form1.SKPU.Cells[6,A]:=Form1.SREZ1.Cells[6,i];//�\�
           Form1.SKPU.Cells[7,A]:=Form1.SREZ1.Cells[7,i];//TRUE
           Form1.SKPU.RowCount:=Form1.SKPU.RowCount+1;}
           SKPU[0,a]:=SREZ1.Cells[0,i]; //Ocher
           SKPU[1,a]:=SREZ1.Cells[1,i];   //Stanok
           SKPU[2,a]:=SREZ1.Cells[2,i];//Dlit
           SKPU[3,a]:=SREZ1.Cells[3,i];//Kol_Smen
           SKPU[4,a]:=SREZ1.Cells[4,i];//Chas
           SKPU[5,a]:='0';//Den
           SKPU[6,a]:=SREZ1.Cells[6,i];//�\�
           SKPU[7,a]:=SREZ1.Cells[7,i];//TRUE
           Inc(A);
       end;
     end;

     //For I:=0 To Form1.SKPU.RowCount-1 Do
     For i:=0 To High(SKPU) do
     Begin
       //if Form1.SKPU.Cells[0,I]<> '' Then
       if SKPU[0,I]<> '' Then
       Begin

            Ocher:= StrToInt(SKPU[0,I]);//Ocher
            Stanok_Posl:=SKPU[1,i];   //Stanok
            Dlit_Min:=StrToFloat(SKPU[2,I]);//Dlit

            Kol_Smen:=StrToInt(SKPU[3,I]);//Kol_Smen
            Chas:=StrToFloat(SKPU[4,I]);
            Den:= StrToInt(SKPU[5,I]);
            Celoe:=Pos(',',SKPU[2,I]);
            if Celoe<>0 Then
            Begin
                if I=0 Then
                Begin
                        H1:=(1-Dlit_Min);
                        SKPU[5,0]:=FloatToStr(1);
                        Form1.Memo13.Lines.Add('1');
                End;
                if I=1 Then
                Begin
                 if (H1=0.5) Then //������� ����� �� ������� ��������
                 Begin
                   SKPU[5,1]:='1';
                   Form1.Memo13.Lines.Add('1');
                   H2:=1;
                  End;
                 if (H1=0) OR (H1<0) Then    //���   OR (H1<0)
                 Begin
                   SKPU[5,1]:='2';   // ��������� �� ������ ����
                   Form1.Memo13.Lines.Add('2');
                   H2:=2;
                  End;
                  H1:=(1-Dlit_Min);
                end;
                if I>1 Then //
                Begin
                 Chet:=I MOD 2;//������� �� �������
                 if (Chet<>0) AND (Dlit_Min=0.5) Then //�� ������
                 Begin
                    H1:=Ocher-Dlit_Min-H2;
                    Den1:= SimpleRoundTo(H1,0);
                    SKPU[5,I]:=(FloatToStr(Den1));
                    Form1.Memo13.Lines.Add(FloatToStr(Den1));
                 end;
                 if (Chet=0) AND (Dlit_Min=0.5) Then //������
                 Begin
                     H2:=Ocher-Dlit_Min-H1;
                     Den1:= SimpleRoundTo(H2,0);
                     SKPU[5,I]:=(FloatToStr(Den1));
                     Form1.Memo13.Lines.Add(FloatToStr(Den1));
                 end;
                end;
            End
            Else
            Begin
                  SKPU[5,I]:=(FloatToStr(Ocher));
                  Form1.Memo13.Lines.Add(FloatToStr(Ocher));
            end;
            {Den:= StrToInt(AD1.Cells[5,I]);
            Dat_Plan:=StrToDate(Dat_Glob);
            Dat_Z:=FormatDateTime('mm.dd.YYYY',Dat_Plan-Den);
            AD1.Cells[8,i]:=Dat_Z;  }
       end;
     end;
     A:= Form1.Memo13.Lines.Count-1 ;
     For i:=0 To High(SKPU) do
     Begin
       if Form1.Memo13.Lines.Strings[a]<> '' Then
       Begin
           SKPU[9,I]:=Form1.Memo13.Lines.Strings[a];
       end;
       A:=A-1;
     End;
     For i:=0 To High(SKPU) do
     Begin
       if SKPU[9,I]<> '' Then
       Begin
         Den:= StrToInt(SKPU[9,I]);
         Dat_Plan:=StrToDate(Dat_Glob);            //����� 18.07.2017
         //Dat_Z:=FormatDateTime('mm.dd.YYYY',Dat_Plan-Den-1);//����� 1����
         //�������, ��������� ����� �������
         Dat_Z:=FormatDateTime('mm.dd.YYYY',Dat_Plan);//����� 1����
         SKPU[8,i]:=Dat_Z;

       end;
     End;

     for I := 0 to High(StrToCol) do
     begin
          if StrToCol[i,5]<>'' Then
          Begin
            For a:=0 To High(SKPU) do
            Begin
                Res:=AnsiCompareStr(StrToCol[i,5],SKPU[1,A]);
                if Res=0 Then
                Begin
                    StrToCol[i,4]:=FloatToStr(Dlit_Obsh);
                    StrToCol[i,6]:=SKPU[8,A];//���� ���������
                end;
            End;
          End;
     end;

     for I := 0 to High(StrToCol) do
     begin
        if StrToCol[i,5]<>'' then
        begin
        S1:=S1+#39+StrToCol[I,0]+#39+','; //True  0
        S1:=S1+#39+StrToCol[I,1]+#39+','; //����������� 1
        S1:=S1+#39+StrToCol[I,2]+#39+','; //�������  2
        S1:=S1+#39+StrToCol[I,3]+#39+','; //�\�  3
        if StrToCol[I,2]='' then
          S1:=S1+'NULL,';
        //
        if (StrToCol[I,2]<>'') then
        Begin
            S1:=S1+#39+StrToCol[I,6]+#39+',';
        End;
        //
        S1:=S1+#39+StrToCol[I,4]+#39+',';
        //���� ���������  5 ����� ������ �� ��������� ����
        end;
     end;
     FreeAndNil(Srez1);
     //FreeAndNil(SKPU);
end;
end.
