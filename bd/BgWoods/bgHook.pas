unit bgHook;

interface
  uses
    Windows, Classes;

type

  TGameType = (BG1, TSC, SOA, TOB, IWD, HOW);
  TChrAttrType = (Str, Cri, Int, Wis, Dex, Con, Cha);
  LONG = Longword;

  PBgData = ^TBgData;
  TBgData = record
    GameType: BYTE;
    Points: BYTE;
    CriPoints: BYTE;
    CriStop: BYTE;
    Delay: LONG;
    Address: LONG;
  end;

  PChrAttrs = ^TChrAttrs;
  TChrAttrs = record
    Str: BYTE;
    Dex: BYTE;
    Con: BYTE;
    Int: BYTE;
    Wis: BYTE;
    Cha: BYTE;
    Cri: BYTE;
  end;

  TThreadBg = class(TThread)
  private
  { Private declarations }
  protected
    procedure Execute; override;
  end;

const
BUFF_SIZE = 1024;
CHRATTRS_COUNT = 7;
KEYS_MAXSIZE = 8;
BEGINPOS = 33554432;
ENDPOS = 201326592;
CMapFileName: PChar = 'BgWoods';
gClsName: PChar = 'ChitinClass';
NUMSTAR = 106;
NUMSLASH = 111;

Offsets: array[0..5] of Integer = (524, 524, 590, 590, 524, 524);
KeysBg1: array[0..5] of BYTE = ($1e, $36, $34, $0c, $17, $1e);
KeysTsc: array[0..5] of BYTE = ($1e, $36, $34, $0c, $17, $1e);
KeysSoa: array[0..7] of BYTE = ($43, $48, $41, $52, $42, $41, $53, $45);
KeysTob: array[0..7] of BYTE = ($43, $48, $41, $52, $42, $41, $53, $45);
KeysIwd: array[0..5] of BYTE = ($1e, $27, $29, $0c, $17, $1e);
KeysHow: array[0..5] of BYTE = ($1e, $27, $29, $0c, $17, $1e);

var
  hNextHookProc : HHook;
  bRun: BOOL;
  ThreadBg: TThreadBg;
  ptrBgData: PBgData;
  hMap: THandle;
  hProcess: THandle;
  procSavExit : Pointer;

function DisableHotKeyHook: BOOL; export;
function EnableHotKeyHook: BOOL; export;
function Search(pAttrs: PChrAttrs): TChrAttrs; stdcall; export;
function SetData(pData: PBgData): TBgData; stdcall; export;

function SetProcess: BOOL;
function KeyboardHook(nCode: Integer; wParam: WPARAM; lParam: LPARAM): LRESULT; stdcall;
procedure HotKeyHookExit; far;
procedure DllEntryPoint(dwReason:Dword);
procedure OpenShareData;
procedure CloseShareData;

implementation

procedure TThreadBg.Execute;
var
Buff : array[0..CHRATTRS_COUNT - 1] of BYTE;
PtScreen: TPoint;
lpNOBR: DWORD;
sum, i: Integer;
begin
  GetCursorPos(PtScreen);
  while bRun do
  begin
    if(ptrBgData^.Address > 0) then
    begin
      if ReadProcessMemory(hProcess, Ptr(ptrBgData^.Address), @Buff, CHRATTRS_COUNT, lpNOBR) then
      begin
        if(ptrBgData^.CriStop <> 0) and (Buff[Integer(Cri)] >= ptrBgData^.CriPoints)then
          break;
        sum := 0;
        for i := 0 to CHRATTRS_COUNT - 1 do
          sum := sum + Buff[i];
        if(sum >= ptrBgData^.Points + Buff[Integer(Cri)])then
          break;
      end;
    end;
    SetCursorPos(PtScreen.x, PtScreen.y);
    Mouse_Event(MOUSEEVENTF_LEFTDOWN, PtScreen.x, PtScreen.y, 0, 0);
    Mouse_Event(MOUSEEVENTF_LEFTUP, PtScreen.x, PtScreen.y, 0, 0);
    Sleep(ptrBgData^.Delay);
  end;
  bRun := FALSE;
end;

function DisableHotKeyHook:BOOL; export;
begin
  if hNextHookPRoc <> 0 then
  begin
    UnhookWindowshookEx(hNextHookProc);
    hNextHookProc := 0;
  end;

  Result := hNextHookProc = 0;
end;

function EnableHotKeyHook: BOOL; export;
begin
  Result := False;
  if hNextHookProc <> 0 then
    exit;
  hNextHookProc := SetWindowsHookEx(WH_KEYBOARD, KeyboardHook, hInstance, 0);

  Result := hNextHookProc <>0;
end;

function Search(pAttrs: PChrAttrs): TChrAttrs; stdcall; export;
var
  Count: LONG;
  i, j, szKeys, Offset: Integer;
  Buff: array[0..BUFF_SIZE + KEYS_MAXSIZE - 2] of BYTE;
  Keys: array[0..KEYS_MAXSIZE - 1] of BYTE;
  lpNOBR: DWORD;
begin
  with Result do
  begin
    Str := 0;
    Dex := 0;
    Con := 0;
    Int := 0;
    Wis := 0;
    Cha := 0;
    Cri := 0;
  end;
  if SetProcess then
  begin
    case ptrBgData^.GameType of
      BYTE(BG1):
      begin
        szKeys := Sizeof(KeysBg1) div Sizeof(BYTE);
        for i := 0 to szKeys - 1 do
          Keys[i] := KeysBg1[i];
      end;
      BYTE(TSC):
      begin
        szKeys := Sizeof(KeysTsc) div Sizeof(BYTE);
        for i := 0 to szKeys - 1 do
          Keys[i] := KeysTsc[i];
      end;
      BYTE(SOA):
      begin
        szKeys := Sizeof(KeysSoa) div Sizeof(BYTE);
        for i := 0 to szKeys - 1 do
          Keys[i] := KeysSoa[i];
      end;
      BYTE(TOB):
      begin
        szKeys := Sizeof(KeysTob) div Sizeof(BYTE);
        for i := 0 to szKeys - 1 do
          Keys[i] := KeysTob[i];
      end;
      BYTE(IWD):
      begin
        szKeys := Sizeof(KeysIwd) div Sizeof(BYTE);
        for i := 0 to szKeys - 1 do
          Keys[i] := KeysIwd[i];
      end;
      BYTE(HOW):
      begin
        szKeys := Sizeof(KeysHow) div Sizeof(BYTE);
        for i := 0 to szKeys - 1 do
          Keys[i] := KeysHow[i];
      end
      else
        exit;
    end;
    Offset := Offsets[ptrBgData^.GameType];
    if pAttrs.Str <> 0 then
    begin
      Keys[Integer(Str)] := pAttrs.Str;
      Keys[Integer(Cri)] := pAttrs.Cri;
      Keys[Integer(Dex)] := pAttrs.Dex;
      Keys[Integer(Con)] := pAttrs.Con;
      Keys[Integer(Int)] := pAttrs.Int;
      Keys[Integer(Wis)] := pAttrs.Wis;
      Keys[Integer(Cha)] := pAttrs.Cha;
      szKeys := 7;
      Offset := 0;
    end;
    Count := BEGINPOS;
    while Count < ENDPOS do
    begin
      if ReadProcessMemory(hProcess, Ptr(Count), @Buff, BUFF_SIZE + szKeys - 1, lpNOBR) then
      begin
        for i := 0 to BUFF_SIZE - 1 do
        begin
          for j := 0 to szKeys - 1 do
            if Buff[i + j] <> Keys[j] then
              break;

          if( j > szKeys - 1)then
          begin
            ptrBgData^.Address := Count + LONG(i + Offset);
            if ReadProcessMemory(hProcess, Ptr(ptrBgData^.Address), @Buff, CHRATTRS_COUNT, lpNOBR) then
            begin
              Result.Str := Buff[Integer(Str)];
              Result.Dex := Buff[Integer(Dex)];
              Result.Con := Buff[Integer(Con)];
              Result.Int := Buff[Integer(Int)];
              Result.Wis := Buff[Integer(Wis)];
              Result.Cha := Buff[Integer(Cha)];
              Result.Cri := Buff[Integer(Cri)];
            end;
            Count := ENDPOS;
            break;
          end;
        end;
      end;
      Count := Count + BUFF_SIZE;
    end;
    CloseHandle(hProcess);
    hProcess := 0;
  end;
end;

function SetData(pData: PBgData): TBgData; stdcall;
begin
  Result:= ptrBgData^;
  ptrBgData^ := pData^;
  ptrBgData^.Address := Result.Address;
end;

function SetProcess: BOOL;
var
  hGame: HWND;
  threadId, pId: DWORD;
begin
  Result := FALSE;

  if(hProcess <> 0)then
    CloseHandle(hProcess);

  hGame := FindWindow(gClsName, nil);
  if(hGame <> 0)then
  begin
    threadId := GetWindowThreadProcessId(hGame, @pId);
    if(threadId <> 0)then
    begin
      hProcess := OpenProcess(PROCESS_ALL_ACCESS, TRUE, pId);
      if(hProcess <> 0)then
        Result := TRUE;
    end;
  end;

end;

function KeyboardHook(nCode: Integer; wParam: WPARAM; lParam: LPARAM): LResult;
begin
  if (nCode > -1) then
  begin
    if (wParam = NUMSTAR) and (bRun = false) then
    begin
      SetProcess( );
      bRun := TRUE;
      ThreadBg := TThreadBg.Create(true);
      ThreadBg.FreeOnTerminate := true;
      ThreadBg.Resume;
      Result := 1;
    end
    else
    begin
      if(wParam = NUMSLASH) then
      begin
        bRun := FALSE;
        if(hProcess <> 0) then
          CloseHandle(hProcess);
        Result := 1;
      end
      else
        Result := 0;
    end;
  end
  else
    Result := CallNextHookEx(hNextHookProc, nCode, wParam, lParam);
end;

procedure OpenShareData;
begin
  hMap := CreateFileMapping(THandle($ffffffff), nil, PAGE_READWRITE, 0, SizeOf(TBgData), CMapFileName);
  if hMap <> 0 then
  begin
    ptrBgData := MapViewOfFile(hMap, FILE_MAP_ALL_ACCESS, 0, 0, Sizeof(TBgData));
    if(ptrBgData = nil) then
      CloseHandle(hMap);
  end;
end;

procedure CloseShareData;
begin
  if(hMap <> 0)then
  begin
    if(ptrBgData <> nil)then
      UnmapViewOfFile(Pointer(ptrBgData));
    hMap := 0;
  end;
end;

procedure DllEntryPoint(dwReason:Dword);
begin
  case dwReason of
    DLL_PROCESS_ATTACH:
      OpenShareData;
    DLL_PROCESS_DETACH:
      CloseShareData;
  end;
end;

procedure HotKeyHookExit;
begin
  if hNextHookProc <> 0 then
    DisableHotKeyHook;
  if hProcess <> 0 then
    CloseHandle(hProcess);

  ExitProc := procSavExit;
end;

end.

