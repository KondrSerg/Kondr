unit UServer;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, WinSock,
                uSockets, StdCtrls;

type
  TFServer = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Memo1: TMemo;
    Memo2: TMemo;
    ListBox1: TListBox;
    Edit1: TEdit;
    Button3: TButton;
    procedure Send(Str_Line: string);
    procedure Button1Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
  private
    ClientCounter: Integer;
    Server: TTCPServer;
    procedure ServerConnect(Client: TTCPConnection);
    procedure ServerDisconnect(Client: TTCPConnection);
    procedure ServerExecute(Thread: TTCPConnectionThread);
    procedure ServerError(Socket: TTCPSocket);
    function GetLocalIPs: TStringList;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FServer: TFServer;
type
  TClientData = class
    Name: String;
  end;
implementation

uses Main;

{$R *.dfm}
procedure TFServer.Send(Str_Line: string);
var
        II: Integer;
begin
        for Ii := 0 to ListBox1.Items.Count - 1 do
        Begin
               ListBox1.Selected[Ii] ;
                        with TTCPConnection(ListBox1.Items.Objects[Ii]) do
                        begin
                                with TClientData(Data) do
                                        Memo1.Lines.Add(Format('>  %s',
                                                [ Str_Line]));
                                WriteLn(Str_Line);
                        end;

                        end;



end;

function TFServer.GetLocalIPs: TStringList;
type
        TaPInAddr = array[0..10] of PInAddr;
        PaPInAddr = ^TaPInAddr;
var
        phe: PHostEnt;
        pptr: PaPInAddr;
        Buffer: array[0..63] of AnsiChar;
        I: Integer;
        GInitData: TWSAData;
        IPs: TStringList;
begin
        IPs := TStringList.Create;
        WSAStartup($101, GInitData);
        GetHostName(Buffer, SizeOf(Buffer));
        phe := GetHostByName(buffer);
        if phe = nil then
                IPs.Add('No IP found')
        else
        begin
                pPtr := PaPInAddr(phe^.h_addr_list);
                I := 0;
                while pPtr^[I] <> nil do
                begin
                        IPs.Add(inet_ntoa(pptr^[I]^));
                        Memo2.Lines.Add(inet_ntoa(pptr^[I]^));
                        Inc(I);
                end;
        end;
        WSACleanup;

        Result := IPs;
end;

procedure TFServer.ServerConnect(Client: TTCPConnection);
begin

end;

procedure TFServer.ServerDisconnect(Client: TTCPConnection);
begin
        with TClientData(Client.Data) do
        begin
                Memo1.Lines.Add(Format('%s disconnected', [Name]));
                ListBox1.Items.Delete(ListBox1.Items.IndexOf(Name));
                Free;
        end;
end;

procedure TFServer.ServerExecute(Thread: TTCPConnectionThread);
var
        Line: string;
begin
        Line := Thread.Connection.ReadLn;
        if Thread.Connection.Connected then
        begin
                Thread.Lock;
                with TClientData(Thread.Connection.Data) do
                try
                        Memo1.Lines.Add(Format('<  %s', [ Line]));
                        Send(Line);
                finally
                        Thread.Unlock;
                end;

        end;
end;

procedure TFServer.ServerError(Socket: TTCPSocket);
begin
        if Socket.LastError <> WSAECONNRESET then
                Memo1.Lines.Add(Format('Error #%d: %s', [Socket.LastError,
                        Socket.LastErrorMessage]));
end;
procedure TFServer.Button1Click(Sender: TObject);
begin
Server := TTCPServer.Create;
        with Server do
        begin
                OnConnect := ServerConnect;
                OnDisconnect := ServerDisconnect;
                OnExecute := ServerExecute;
                OnError := ServerError;
                GetLocalIPs;

                AddBinding(TTCPBinding.Create(Form1.Memo6.Lines.Strings[0]+':12345'));


        end;
        Server.Listen := not Server.Listen;
end;

procedure TFServer.FormDestroy(Sender: TObject);
begin

  //Server.Listen := False;
  //Server.Free;
end;

procedure TFServer.Button2Click(Sender: TObject);
begin
Send(Edit1.Text);
end;

procedure TFServer.Button3Click(Sender: TObject);
var I: Integer;
begin
  for I := ListBox1.Items.Count - 1 downto 0 do
  Begin
   ListBox1.Selected[I];
  with TTCPConnection(ListBox1.Items.Objects[I]) do
    Disconnect;
    End;
end;

end.
