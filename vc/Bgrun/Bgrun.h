
#define BUFF_SIZE	1024
#define BEGINPOS	33554432
#define ENDPOS	201326592
#define CHRATTRS_COUNT	7
#define NUMSTAR	106
#define NUMSLASH	111

/////////////////////////////////////////////////////////////////////////////
// BgData

typedef enum 
{
	BG1 = 0, 
	TSC = 1, 
	SOA = 2, 
	TOB = 3, 
	IWD = 4, 
	HOW = 5, 
}GameType;

typedef enum 
{ 
	Str = 0, 
	Cri = 1, 
	Int = 2, 
	Wis = 3, 
	Dex = 4, 
	Con = 5, 
	Cha = 6, 
}ChrAttrType; 

typedef struct
{
	BYTE Str;
	BYTE Dex;
	BYTE Con;
	BYTE Int;
	BYTE Wis;
	BYTE Cha;
	BYTE Cri;
}CChrAttrs;

typedef struct
{
	BYTE GameType;
	BYTE Points;
	BYTE CriPoints;
	BYTE CriStop;
	LONG Delay;
	LONG Address;
}CBgData;

static char CMapFileName[] = "BgWoods";
static char gClsName[] = "ChitinClass";
static int Offsets[] = {524, 524, 590, 590, 524, 524};
static BYTE KeysBg1[] = {0x1e, 0x36, 0x34, 0x0c, 0x17, 0x1e};
static BYTE KeysTsc[] = {0x1e, 0x36, 0x34, 0x0c, 0x17, 0x1e};
static BYTE KeysSoa[] = {0x43, 0x48, 0x41, 0x52, 0x42, 0x41, 0x53, 0x45};
static BYTE KeysTob[] = {0x43, 0x48, 0x41, 0x52, 0x42, 0x41, 0x53, 0x45};
static BYTE KeysIwd[] = {0x1e, 0x27, 0x29, 0x0c, 0x17, 0x1e};
static BYTE KeysHow[] = {0x1e, 0x27, 0x29, 0x0c, 0x17, 0x1e};

/////////////////////////////////////////////////////////////////////////////
// definitions for export functions

CBgData __stdcall SetData(const CBgData* pData);
CChrAttrs __stdcall Search(const CChrAttrs* pAttrs);
BOOL __stdcall EnableHotKeyHook(void);
BOOL __stdcall DisableHotKeyHook(void);
/////////////////////////////////////////////////////////////////////////////
//definitions for thread and keyboard hook functions 

DWORD _stdcall CheckChrAttrs(LPVOID lparam);
LRESULT CALLBACK KeyBoardHook(int nCode, WPARAM wParam, LPARAM lParam);

/////////////////////////////////////////////////////////////////////////////

BOOL bRun = FALSE;
HINSTANCE hInstance;
HANDLE hProcess = NULL;
HANDLE hMap = NULL;
HANDLE hThread = NULL;
HHOOK hNextHookProc = 0;
CBgData* pBgData = NULL;

void OpenShareData(void);
void CloseShareData(void);
BOOL SetProcess(void);