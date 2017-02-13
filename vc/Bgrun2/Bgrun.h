
// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the BGRUN_EXPORTS
// symbol defined on the command line. this symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// BGRUN_API functions as being imported from a DLL, wheras this DLL sees symbols
// defined with this macro as being exported.
#ifdef BGRUN_EXPORTS
#define BGRUN_API __declspec(dllexport)
#else
#define BGRUN_API __declspec(dllimport)
#endif

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

CBgData BGRUN_API SetData(const CBgData* pBgData);
CChrAttrs BGRUN_API Search(const CChrAttrs* pAttrs);
BOOL BGRUN_API EnableHotKeyHook(void);
BOOL BGRUN_API DisableHotKeyHook(void);

/////////////////////////////////////////////////////////////////////////////
//definitions for thread and keyboard hook functions 

DWORD _stdcall CheckChrAttrs(LPVOID lparam);
LRESULT CALLBACK KeyBoardHook(int nCode, WPARAM wParam, LPARAM lParam);

/////////////////////////////////////////////////////////////////////////////
// CBgrun class

class CBgrun {
public:
	//attributes
	HINSTANCE m_hInstance;
	BOOL m_bRun;
	//operations
	CBgrun(void);
	~CBgrun(void);
	CBgData SetData(const CBgData* pBgData);
	CBgData* GetData(void);
	CChrAttrs Search(const CChrAttrs* pAttrs);
	CChrAttrs* GetChrAttrs(void);
	void OpenShareData();
	void CloseShareData(void);
	BOOL EnableHotKeyHook(void);
	BOOL DisableHotKeyHook(void);
	HANDLE GetProcess(void);
	BOOL SetProcess(void);
	BOOL CloseProcess(void);
	HHOOK GetNextHookProc(void);
	BOOL StartCheck(void);
	// TODO: add your methods here.
private:
	//attributes
	HANDLE m_hThread;
	HHOOK	m_hNextHookProc;
	HANDLE m_hProcess;
	HANDLE m_hMap;
	CBgData* m_pBgData;
	CChrAttrs* m_pChrAttrs;
	//operations
};
