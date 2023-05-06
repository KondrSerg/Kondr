unit UNomer;

interface

uses
        Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
        Forms,
        Dialogs, StdCtrls, ComCtrls;

type
        TFNomer = class(TForm)
                Edit1: TEdit;
                Button1: TButton;
                Button2: TButton;
                DateTimePicker1: TDateTimePicker;
                Label3: TLabel;
                Label2: TLabel;
                Label5: TLabel;
                Label4: TLabel;
                ComboBox1: TComboBox;
                ComboBox2: TComboBox;
                Label1: TLabel;
                Label6: TLabel;
    Button3: TButton;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
                procedure Button2Click(Sender: TObject);
                procedure Button1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ComboBox1KeyPress(Sender: TObject; var Key: Char);
    procedure ComboBox2KeyPress(Sender: TObject; var Key: Char);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    function Log(Lin:String):Boolean;
        private
                { Private declarations }
        public
                { Public declarations }
                Summa: Integer;
        end;

var
        FNomer: TFNomer;

implementation

uses Main;

{$R *.dfm}
function TFNomer.Log(Lin:String):Boolean;
var
infile, outfile: TextFile;
num_lines, x: integer;
line: string;
begin
{assignfile(infile, 'C:\INFILE.TXT');
//assignfile(outfile, 'C:\OUTFILE.TXT');
reset(infile);  {ÔÂÂÏÂ˘‡ÂÏ ÛÍ‡Á‡ÚÂÎ¸}
{‚ Ì‡˜‡ÎÓ Ù‡ÈÎ‡.}
//rewrite(outfile);  {Ó˜Ë˘‡ÂÏ ÒÓ‰ÂÊËÏÓÂ Ù‡ÈÎ‡}
{readln(infile, num_lines);
for x:= 1 to num_lines do
begin
readln(infile, line);
writeln(infile, line);
end;
closefile(infile);
//closefile(outfile);}
end;
procedure TFNomer.Button2Click(Sender: TObject);
begin
        Close;
end;

procedure TFNomer.Button1Click(Sender: TObject);
var
        I, Res, Kol_ZAP, Kol: Integer;
        Str, Zak, Nam, Dat,BZ,IDGP,IDKO: string;
begin
        Nam := FNomer.Caption;
        Zak := Label2.Caption;
        BZ := Label7.Caption;
        IDGP:=Label8.Caption;
        IDKO:=Label9.Caption;
        if (DateTimePicker1.Visible = False) and (Edit1.Visible = True) then
        begin

                if Form1.Vozduh=1 Then
                Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label3.Caption + ']=' + #39 + Edit1.Text + #39 +
                        ' WHERE ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ')', ['Klapana']) then
                        Begin
                MessageDlg('Œ¯Ë·Í‡ Á‡ÔËÒË!', mtError,
                        [mbOk], 0);
                        Exit;
                        End;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1 , 'UPDATE %s SET [' +
                        Label3.Caption + ']=' + #39 + Edit1.Text + #39 +
                        ' WHERE ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ')', ['«‡ÔÛÒÍ']) then
                        Begin
                MessageDlg('Œ¯Ë·Í‡ Á‡ÔËÒË!', mtError,
                        [mbOk], 0);
                        Exit;
                        End;

                End;
                if Form1.Vozduh=0 Then
                Begin

                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label3.Caption + ']=' + #39 + Edit1.Text + #39 +
                        ' WHERE ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ') AND (¡«='+#39+BZ+#39+') AND (ID√œ='+#39+IDGP+#39+') AND (ID Œ='+#39+IDKO+#39+')',
                         ['KlapanaZap']) then
                        Exit;

                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label3.Caption + ']=' + #39 + Edit1.Text + #39 +
                        ' WHERE ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ') AND (¡«='+#39+BZ+#39+') AND (ID√œ='+#39+IDGP+#39+') AND (ID Œ='+#39+IDKO+#39+')',
                        ['«‡ÔÛÒÍ¬ÓÁ‰']) then
                        Exit;
                End;
{FTSP.Send( Label3.Caption +  ' = '+Edit1.Text+'  ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ')');  }
                Close;

        end;
        //====================================================
        if DateTimePicker1.Visible = True then
        begin
                if Form1.Vozduh=1 Then
                Begin
                Str := FormatDateTime('mm.dd.yyyy', DateTimePicker1.Date);
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label3.Caption + ']=' + #39 + Str + #39 +
                        ' WHERE ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ')', ['Klapana']) then
                        Exit;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label3.Caption + ']=' + #39 + Str + #39 +
                        ' WHERE ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ')', ['«‡ÔÛÒÍ']) then
                        Exit;
                if Form1.SG6.Col= I_FN_TEHNOLOG_GOTOV Then
                Form1.SG6.Cells[I_FN_TEHNOLOG_GOTOV ,Form1.SG6.Row ] := Str;
                end;
                if Form1.Vozduh=0 Then
                Begin
                Str := FormatDateTime('mm.dd.yyyy', DateTimePicker1.Date);
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label3.Caption + ']=' + #39 + Str + #39 +
                        ' WHERE ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ') AND (¡«='+#39+BZ+#39+') AND (ID√œ='+#39+IDGP+#39+') AND (ID Œ='+#39+IDKO+#39+')',
                        ['KlapanaZap']) then
                        Exit;

                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label3.Caption + ']=' + #39 + Str + #39 +
                        ' WHERE ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ') AND (¡«='+#39+BZ+#39+') AND (ID√œ='+#39+IDGP+#39+') AND (ID Œ='+#39+IDKO+#39+')',
                        ['«‡ÔÛÒÍ¬ÓÁ‰']) then
                        Exit;
                        {FTSP.Send( Label3.Caption +  ' = '+Str+'  ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ')'); }
                End;

                if Form1.Vozduh=2 Then   //—“¿Ã
                Begin
                Str := FormatDateTime('mm.dd.yyyy', DateTimePicker1.Date);
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label3.Caption + ']=' + #39 + Str + #39 +
                        ' WHERE  ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ') AND (¡«='+#39+BZ+#39+') AND (ID√œ='+#39+IDGP+#39+')',
                        ['—“¿Ã']) then
                        Exit;
                End;
                if Form1.Vozduh=12 Then   //Àﬁ 
                Begin
                Str := FormatDateTime('mm.dd.yyyy', DateTimePicker1.Date);
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label3.Caption + ']=' + #39 + Str + #39 +
                        ' WHERE  ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ') AND (¡«='+#39+BZ+#39+') AND (ID√œ='+#39+IDGP+#39+')',
                        ['Àﬁ ']) then
                        Exit;
                End;
        end;
        //========================================================
        if (DateTimePicker1.Visible = False) and (Edit1.Visible = False) then
        begin
                if Form1.Vozduh=1 Then
                Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label3.Caption + ']=' + #39 + ComboBox1.Text + #39 +
                        ',œË‚Ó‰=' + #39 +
                        ComboBox2.Text + #39 +
                        ' WHERE ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ')', ['Klapana']) then
                        Exit;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label3.Caption + ']=' + #39 + ComboBox1.Text + #39 +
                        ',œË‚Ó‰=' + #39 +
                        ComboBox2.Text + #39 +
                        ' WHERE ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ')', ['«‡ÔÛÒÍ']) then
                        Exit;

                end;
                if Form1.Vozduh=0 Then
                Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label3.Caption + ']=' + #39 + ComboBox1.Text + #39 +
                        ',œË‚Ó‰=' + #39 +
                        ComboBox2.Text + #39 +
                        ' WHERE ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ') AND (¡«='+#39+BZ+#39+') AND (ID√œ='+#39+IDGP+#39+') AND (ID Œ='+#39+IDKO+#39+')', ['KlapanaZap']) then
                        Exit;

                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label3.Caption + ']=' + #39 + ComboBox1.Text + #39 +
                        ',œË‚Ó‰=' + #39 +
                        ComboBox2.Text + #39 +
                        ' WHERE ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ') AND (¡«='+#39+BZ+#39+') AND (ID√œ='+#39+IDGP+#39+') AND (ID Œ='+#39+IDKO+#39+')', ['«‡ÔÛÒÍ¬ÓÁ‰']) then
                        Exit;
                        {FTSP.Send( Label3.Caption +  ' = '+ComboBox1.Text+' , œË‚Ó‰= '+ComboBox2.Text+'  ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ')');  }

                end;
                Close;

        end;
        //Form1.Button16.Click;
        Close;

end;

procedure TFNomer.Button3Click(Sender: TObject);
Var
 Nam,Zak,BZ,IDGP,IDKO:String;
begin
        Nam := FNomer.Caption;
        Zak := Label2.Caption;
        BZ := Label7.Caption;
        if DateTimePicker1.Visible = True then
        begin
                if Form1.Vozduh=1 Then
                Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label3.Caption + ']= NULL  WHERE ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ')', ['Klapana']) then
                        Exit;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label3.Caption + ']=NULL '+
                        ' WHERE ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ')', ['«‡ÔÛÒÍ']) then
                        Exit;
                End;
                if Form1.Vozduh=0 Then
                Begin
                  IDGP:=Label8.Caption;
                  IDKO:=Label9.Caption;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label3.Caption + ']= NULL '+
                        ' WHERE ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ') AND (¡«='+#39+BZ+#39+') AND (ID√œ='+#39+IDGP+#39+') AND (ID Œ='+#39+IDKO+#39+')', ['KlapanaZap']) then
                        Exit;

                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label3.Caption + ']= NULL '+
                        ' WHERE ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ') AND (¡«='+#39+BZ+#39+') AND (ID√œ='+#39+IDGP+#39+') AND (ID Œ='+#39+IDKO+#39+')', ['«‡ÔÛÒÍ¬ÓÁ‰']) then
                        Exit;
                        //FTSP.Send( Label3.Caption +  ' = NULL');
                End;

                if Form1.Vozduh=2 Then //—“¿Ã
                Begin
                  IDGP:=Label8.Caption;
                  IDKO:=Label9.Caption;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label3.Caption + ']= NULL '+
                        ' WHERE  ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ') AND (¡«='+#39+BZ+#39+') AND (ID√œ='+#39+IDGP+#39+')', ['—“¿Ã']) then
                        Exit;
                End;
               if Form1.Vozduh=12 Then //Àﬁ 
                Begin
                  IDGP:=Label8.Caption;
                  IDKO:=Label9.Caption;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label3.Caption + ']= NULL '+
                        ' WHERE  ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ') AND (¡«='+#39+BZ+#39+') AND (ID√œ='+#39+IDGP+#39+')', ['Àﬁ ']) then
                        Exit;
                End;
        end;
        Close;
end;

procedure TFNomer.FormShow(Sender: TObject);
begin
        if DateTimePicker1.Visible = True then
            Button3.Visible:=True
            Else
            Button3.Visible:=False;
end;

procedure TFNomer.ComboBox1KeyPress(Sender: TObject; var Key: Char);
begin
     If Key=#27 Then
     Close;
        if (Key = #13) then
                Button1Click(nil);
end;

procedure TFNomer.ComboBox2KeyPress(Sender: TObject; var Key: Char);
begin
     If Key=#27 Then
     Close;
        if (Key = #13) then
                Button1Click(nil);
end;

procedure TFNomer.FormKeyPress(Sender: TObject; var Key: Char);
begin
     If Key=#27 Then
     Close;
        if (Key = #13) then
                Button1Click(nil);
end;

end.
