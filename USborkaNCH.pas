unit USborkaNCH;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Grids, AdvObj,
  BaseGrid, AdvGrid, Vcl.ExtCtrls, Vcl.Menus, Vcl.ComCtrls,Vcl.Clipbrd;

type
  TFSborNCH = class(TForm)
    pm1: TPopupMenu;
    N1: TMenuItem;
    pgc1: TPageControl;
    ts1: TTabSheet;
    pnl2: TPanel;
    btn4: TButton;
    btn5: TButton;
    btn6: TButton;
    btndt1: TButtonedEdit;
    SG1: TAdvStringGrid;
    Memo2: TMemo;
    ts2: TTabSheet;
    Label1: TLabel;
    btn7: TButton;
    ComboBox1: TComboBox;
    N2: TMenuItem;
    N3: TMenuItem;
    ComboBox2: TComboBox;
    btn8: TButton;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    N4: TMenuItem;
    btn1: TButton;
    N5: TMenuItem;
    Label5: TLabel;
    procedure N1Click(Sender: TObject);
    procedure btndt1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure btn5Click(Sender: TObject);
    procedure btn4Click(Sender: TObject);
    procedure btn6Click(Sender: TObject);
    procedure btn7Click(Sender: TObject);
    procedure ComboBox1Select(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btn8Click(Sender: TObject);
    procedure ComboBox2Select(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;
const N=50;
var
  FSborNCH: TFSborNCH;
  CMB:array[1..n] of TComboBox;

implementation

uses
  Main, UnewTip;

{$R *.dfm}

procedure TFSborNCH.btn1Click(Sender: TObject);
var S,Kl,KL1,t,Tex:string;
I,ID,ind:Integer;
begin
    if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s '+
    '  Order by ID ',
    ['Измена']) then
    exit;
    SG1.RowCount:=  Form1.ADOQuery1.RecordCount+2;
    for I := 0 to Form1.ADOQuery1.RecordCount-1 do
    begin
        Kl:=Form1.ADOQuery1.FieldByName('П2').AsString;
        T:=StringReplace(Kl, '-1', 'X', [RFReplaceall]);
        KL:=StringReplace(T, ',', '-', [RFReplaceall]);
        SG1.Cells[0,i+1]:=Form1.ADOQuery1.FieldByName('ID').AsString;
        SG1.Cells[1,i+1]:=Form1.ADOQuery1.FieldByName('Клапан').AsString+' '+
        KL;
        SG1.Cells[2,i+1]:=Form1.ADOQuery1.FieldByName('Запрет').AsString;
        Form1.ADOQuery1.Next;
    end;
end;

procedure TFSborNCH.btn4Click(Sender: TObject);
var I:Integer;
begin
for i:= Low( CMB ) to High( CMB ) do
    begin
      if CMB[ i ]<>nil then
      begin
        CMB[i].Destroy;
        CMB[ i ] := nil;
      end;
    end;
    btndt1.Text:='';
    //Label1.Caption:='';
end;

procedure TFSborNCH.btn5Click(Sender: TObject);
var S,Kl,KL1,Tex,Kod:string;
I,ID,ind:Integer;
begin
    if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select MAX(ID) AS I from %s ',
  ['Измена']) then
  exit;
  ID := Form1.ADOQuery1.FieldByName('I').AsInteger;
  Tex:=ComboBox2.Text+' '+btndt1.Text;
    for i:= Low( CMB ) to High( CMB ) do
    begin
      if CMB[ i ]<>nil then
      begin
        ind:=CMB[i].ItemIndex;

        if ind=-1 then
        KL1:=Kl1+'-1,'
        Else
        KL1:=Kl1+CMB[i].Text+',';
        if ind=0 then
        ind:=-1;
        S:=S+IntToStr(ind)+',';
      end;
    end;
   //kl:=Label1.Caption;
    kl:=ComboBox2.Text;
    //
    Kod:=Label5.Caption;
    if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s '+
    ' Where (Клапан='+#39+KL+#39+') and (П1='+#39+S+#39+') ',
    ['Измена']) then
    exit;
    //
    if Form1.ADOQuery1.RecordCount<>0 then
    begin
       MessageDlg('Такое сочетание признаков уже есть в базе!',mtInformation,[mbOk],0);
        Exit;
    end;

    if not Form1.mkQueryInsert(Form1.ADOQuery2,
    'Insert Into %s ' +
    '(Клапан,ID,П1,П2,П3,П4) ' +
    'Values (%s,%s,%s,%s,%s,%s)',
    ['Измена',#39 + kl + #39,#39 +intToStr(ID+1)+ #39,#39 + S + #39,
    #39 + KL1 + #39,#39 + Tex + #39,#39 + Kod + #39  ])
    then
    exit;
end;

procedure TFSborNCH.btn6Click(Sender: TObject);
begin
  Close;
end;

procedure TFSborNCH.btn7Click(Sender: TObject);
var Res1,Res2,Res3,I,l,C,A,B:Integer;
S,S1,T,KL:string;
begin
     if CMB[1]<>nil then
         btn4.Click;
     if btndt1.Text='' then
       Exit;
     S:=btndt1.Text;
     Res1:=Pos(' ',S);
     if Res1<>0 then
     Delete(S,1,Res1);
     //++++++++++++++++++++++ Клапан лепестковый ТЮЛЬПАН-2-630*630-Н-0
     Res3:=Pos(' ',S);
     if Res3<>0 then
     Delete(S,1,Res3);
     //+++++++++++++++++++++++++
     Res2:=Pos('-',S);
     if (Res1<>0) and (Res2<>0) then
     begin
      KL:= Copy(S,1,Res2-1);
      if not Form1.mkQuerySelect1(Form1.ADOQuery2,
      'Select * from %s Where Обозначение='+#39+KL+#39 ,
      ['Тип']) then
      exit;
      if Form1.ADOQuery2.RecordCount=0  then
      begin
        MessageDlg('Заведите такой тип Клапана в базу!',mtInformation,[mbOk],0);
        Exit;
      end;
      Label5.Caption:=Form1.ADOQuery2.FieldByName('Код').AsString;
      Label1.Caption:=KL;
      ComboBox2.Text:=KL;
      Delete(S,1,Res2);
     end;
     btndt1.Text:=S;
     S1:=S;
     //________________________________
     c:=0;
    for i:=1 to length(s) do
    Begin
       if s[i] = '-' then
       begin
        inc(c);
        CMB[C]:=TComboBox.Create(FSborNCH);
        CMB[C].Parent:=FSborNCH;
        Res1:=Pos('-',S1);  //ГАБАРИТЫ со* надо разбить   Клапан ГЕРМИК-С -350*200- Н-1*LF230-S-1-УХЛ2

        T:= Copy(S1,1,Res1-1);
        //
          Res2:=Pos('*',T);
          if (Res2<>0) AND (Res2>3) then
          begin
            try
               A:=StrToInt(Copy(S1,1,Res2-1)) ;//350*200-
               CMB[C].Text:=IntToStr(A);
               CMB[C].Items.Add(IntToStr(A));
              Delete(S1,1,Res2);  //200-
              CMB[C].Name:='CMB'+inttostr(C);
              CMB[C].Width:=length(T)*7+25;
              if C=1 then
                CMB[C].Left:=Label1.Left+Label1.Width+10
              Else
                CMB[C].Left:=CMB[C-1].Left+CMB[C-1].Width+10;
              CMB[C].Top:=Label1.Top+20;
              inc(C);
              CMB[C]:=TComboBox.Create(FSborNCH);
              CMB[C].Parent:=FSborNCH;
              Res1:=Pos('-',S1);
              T:= Copy(S1,1,Res1-1);
            finally

            end;
          end;
        //
        CMB[C].Text:=T;
        CMB[C].Items.Add(T);
        Delete(S1,1,Res1);
        CMB[C].Name:='CMB'+inttostr(C);
        CMB[C].Width:=length(T)*7+25;
        //CMB[i].OnClick:=ButtonClick;
        if C=1 then
          CMB[C].Left:=Label1.Left+Label1.Width+10
        Else
          CMB[C].Left:=CMB[C-1].Left+CMB[C-1].Width+10;

        CMB[C].Top:=Label1.Top+20;
        //inc(c);
       end;
    End;
    inc(c);
    CMB[C]:=TComboBox.Create(FSborNCH);
    CMB[C].Parent:=FSborNCH;
    CMB[C].Text:=S1;
    CMB[C].Items.Add(S1);
    CMB[C].Name:='CMB'+inttostr(C);
    CMB[C].Width:=length(s1)*7+25;
    CMB[C].Left:=CMB[C-1].Left+CMB[C-1].Width+10;
    CMB[C].Top:=Label1.Top+20;

    for l := 1 to C do
    begin
      if not Form1.mkQuerySelect1(Form1.ADOQuery2,
      'Select * from %s Order By ID' ,
      ['Признаки']) then
      exit;
      for I := 0 to Form1.ADOQuery2.RecordCount-1 do
      begin
           //CMB[C].Items.Insert(I+2,Form1.ADOQuery2.FieldByName('Признак').AsString);
           CMB[l].Items.Add(Form1.ADOQuery2.FieldByName('Признак').AsString);
           Form1.ADOQuery2.Next;
      end;
    end;
    //==============================
   btn8.Click;

end;

procedure TFSborNCH.btn8Click(Sender: TObject);
var KL,T,S:string;
I:Integer;
begin
    Form1.Clear_StringGrid(SG1);
    if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s '+
    ' Where (Клапан='+#39+ComboBox2.Text+#39+') Order by ID ',
    ['Измена']) then
    exit;
    SG1.RowCount:=  Form1.ADOQuery1.RecordCount+2;
    for I := 0 to Form1.ADOQuery1.RecordCount-1 do
    begin
        Kl:=Form1.ADOQuery1.FieldByName('П2').AsString;
        T:=StringReplace(Kl, '-1', 'X', [RFReplaceall]);
        KL:=StringReplace(T, ',', '-', [RFReplaceall]);
        SG1.Cells[0,i+1]:=Form1.ADOQuery1.FieldByName('ID').AsString;
        SG1.Cells[1,i+1]:=Form1.ADOQuery1.FieldByName('Клапан').AsString+' '+
        KL;
        SG1.Cells[2,i+1]:=Form1.ADOQuery1.FieldByName('Запрет').AsString;
        Form1.ADOQuery1.Next;
    end;
end;

procedure TFSborNCH.btndt1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var Res1,Res2,Res3,I,l,C:Integer;
S,S1,T,KL:string;
begin
 if Key=13 then
 Begin
  btn7.Click;
 End;
end;

procedure TFSborNCH.ComboBox1Select(Sender: TObject);
var KL,T:string;
I:Integer;
begin
    //==============================
    Form1.Clear_StringGrid(SG1);
    if ComboBox1.ItemIndex=1 then
    if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s '+
    ' Where (Запрет='+#39+'1'+#39+') Order by ID ',
    ['Измена']) then
    exit;
    if ComboBox1.ItemIndex=0 then
    if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s '+
    '  Order by ID ',
    ['Измена']) then
    exit;
    SG1.RowCount:=  Form1.ADOQuery1.RecordCount+2;
    for I := 0 to Form1.ADOQuery1.RecordCount-1 do
    begin
        Kl:=Form1.ADOQuery1.FieldByName('П2').AsString;
        T:=StringReplace(Kl, '-1', 'X', [RFReplaceall]);
        KL:=StringReplace(T, ',', '-', [RFReplaceall]);
        SG1.Cells[0,i+1]:=Form1.ADOQuery1.FieldByName('ID').AsString;
        SG1.Cells[1,i+1]:=Form1.ADOQuery1.FieldByName('Клапан').AsString+' '+
        KL;
        SG1.Cells[2,i+1]:=Form1.ADOQuery1.FieldByName('Запрет').AsString;
        Form1.ADOQuery1.Next;
    end;
end;

procedure TFSborNCH.ComboBox2Select(Sender: TObject);
begin
  btn8.Click;
end;

procedure TFSborNCH.FormClose(Sender: TObject; var Action: TCloseAction);
var I:Integer;
begin
for i:= Low( CMB ) to High( CMB ) do
    begin
      if CMB[ i ]<>nil then
      begin
        CMB[i].Destroy;
        CMB[ i ] := nil;
      end;
    end;
    btndt1.Text:='';
end;

procedure TFSborNCH.FormShow(Sender: TObject);
var I:Integer;
begin
      if not Form1.mkQuerySelect1(Form1.ADOQuery2,
      'Select * from %s Order by ID' ,
      ['Тип']) then
      exit;
      for I := 0 to Form1.ADOQuery2.RecordCount-1 do
        begin
          ComboBox2.Items.Add(Form1.ADOQuery2.FieldByName('Обозначение').AsString);
          Form1.ADOQuery2.Next;
        end;
end;

procedure TFSborNCH.N1Click(Sender: TObject);
begin
  FNewTip.Show;
end;

procedure TFSborNCH.N2Click(Sender: TObject);
begin
    if not Form1.mkQueryUpdate(Form1.ADOQuery1,
    'UPDATE %s SET [Запрет]=' + #39 +'1' + #39 +
    ' WHERE ([ID]=' + #39 +SG1.Cells[0, SG1.Row] + #39 +')',
    ['Измена']) then
    Exit;
    ComboBox1.OnSelect(Sender);
end;

procedure TFSborNCH.N3Click(Sender: TObject);
begin
    if not Form1.mkQueryUpdate(Form1.ADOQuery1,
    'UPDATE %s SET [Запрет]=' + #39 +'0' + #39 +
    ' WHERE ([ID]=' + #39 +SG1.Cells[0, SG1.Row] + #39 +')',
    ['Измена']) then
    Exit;
    ComboBox1.OnSelect(Sender);
end;

procedure TFSborNCH.N4Click(Sender: TObject);
begin
  if not Form1.mkQueryDelete(Form1.ADOQuery1, 'DELETE FROM %s Where (Id= ' + #39 +
  SG1.Cells[0 , SG1.Row] + #39 + ') ',
  ['Измена']) then
      Exit;
  ComboBox1.OnSelect(Sender);
end;

procedure TFSborNCH.N5Click(Sender: TObject);
begin
    Clipboard.AsText := SG1.Cells[SG1.Col, SG1.Row];
end;

end.
