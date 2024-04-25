program tping;

{$mode objfpc}{$H+}

uses
  Classes,
  SysUtils,
  CustApp,
  Windows,
  ActiveX,
  ComObj,
  DateUtils;

type
  TPingThread = class(TThread)
    Computer: string;
    procedure Execute; override;
  end;

var
  Retries: integer = 4;
  Buffer: integer = 64;
  Timeout: integer = 60;
  MaxThreads: integer = 100;
  MaskIP: string = '192.168.0.';
  str: string;
  t: TPingThread;
  i: integer;
  ActiveThreads: integer = 0;
  IpOffLine: integer = 0;
  IpOnLine: integer = 0;
  StartTime, StopTime: TDateTime;

  function WMIPing(const Address: string; const BufferSize, Timeout: word): integer;
  const
    WbemUser = '';
    WbemPassword = '';
    WbemComputer = 'localhost';
    wbemFlagForwardOnly = $00000020;
  var
    FSWbemLocator: olevariant;
    FWMIService: olevariant;
    FWbemObjectSet: olevariant;
    FWbemObject: olevariant;
    oEnum: ActiveX.IEnumvariant;
    FWbemQuery: string[250];
    _Nil: longword;
  begin
    Result := -1;
    CoInitialize(nil);
    try
      FSWbemLocator := ComObj.CreateOleObject('WbemScripting.SWbemLocator');
      FWMIService := FSWbemLocator.ConnectServer(WbemComputer,
        'root\CIMV2', WbemUser, WbemPassword);
      FWbemQuery := Format(
        'Select * From Win32_PingStatus Where Address=%s And BufferSize=%d And TimeOut=%d',
        [QuotedStr(Address), BufferSize, Timeout]);
      FWbemObjectSet := FWMIService.ExecQuery(FWbemQuery, 'WQL', wbemFlagForwardOnly);
      oEnum := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
      _Nil := 0;
      while oEnum.Next(1, FWbemObject, _Nil) = 0 do
      begin
        Result := longint(FWbemObject.Properties_.Item('StatusCode').Value);
        FWbemObject := Unassigned;
      end;
    finally
      CoUninitialize;
    end;
  end;

  procedure TPingThread.Execute;
  var
    i, k: integer;
  begin
    k := 0;
    for i := 1 to Retries do
      if WmiPing(Computer, Buffer, Timeout) = 0 then
        Inc(k);

    if k = 0 then
      Inc(IpOffLine)
    else
      Inc(IpOnLine);

    Writeln(Computer, ';', k, ';', Retries);
    Dec(ActiveThreads);
  end;

begin
  i := 0;

  if ParamCount = 0 then
  begin
    WriteLn('TPING - Threaded Ping');
    WriteLn('a software by stefan.arhip');
    WriteLn;
    Write('Max threads [', MaxThreads, ']=');
    ReadLn(str);
    if str <> '' then
      MaxThreads := StrToInt(str);
    Write('Retries [', Retries, ']=');
    ReadLn(str);
    if str <> '' then
      Retries := StrToInt(str);
    Write('Buffer size [', Buffer, ']=');
    ReadLn(str);
    if str <> '' then
      Buffer := StrToInt(str);
    Write('Timeout [', Timeout, ']=');
    ReadLn(str);
    if str <> '' then
      Timeout := StrToInt(str);
    Write('IP mask to scan [', MaskIP, ']=');
    ReadLn(str);
    if str <> '' then
      MaskIP := str;
  end;

  StartTime := Now();
  WriteLn('Start time: ', FormatDateTime('yyyymmdd-hhnnss', StartTime));
  WriteLn('IP;Success;Retries');

  while i < 255 do
    if ActiveThreads < MaxThreads then
    begin
      t := TPingThread.Create(True);
      Inc(ActiveThreads);
      Inc(i);
      t.Computer := Format(MaskIP + '%d', [i - 1]);
      t.FreeOnTerminate := True;
      t.Start;
    end;

  while ActiveThreads > 0 do ;

  StopTime := Now();
  WriteLn('End time: ', FormatDateTime('yyyymmdd-hhnnss', StopTime),
    ' ', Format('%.2f', [DateUtils.MilliSecondsBetween(StartTime, StopTime) / 1000]),
    ' seconds.');
  WriteLn('Offline = ', IpOffline);
  WriteLn('Online = ', IpOnline);
  //Readln;
end.
