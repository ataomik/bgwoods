// Bgrun.cpp : Defines the entry point for the DLL application.
//

#include "stdafx.h"
#include "Bgrun.h"

CBgrun clsBgrun;

/////////////////////////////////////////////////////////////////////////////
//export functions
CBgData SetData(const CBgData* pBgData)
{
	return clsBgrun.SetData(pBgData);
}
CChrAttrs Search(const CChrAttrs* pAttrs)
{
	return clsBgrun.Search(pAttrs);
}

BOOL EnableHotKeyHook(void)
{
	return clsBgrun.EnableHotKeyHook();
}

BOOL DisableHotKeyHook(void)
{
	return clsBgrun.DisableHotKeyHook();
}

/////////////////////////////////////////////////////////////////////////////
//thread function

DWORD _stdcall CheckChrAttrs(LPVOID lparam)
{
	int i, sum;
	BYTE Buff[CHRATTRS_COUNT];
	DWORD NOBR;
	CBgrun* pBgrun = (CBgrun* )lparam;
	POINT PtScreen;

	GetCursorPos(&PtScreen);
	while(pBgrun->m_bRun)
	{
		if(ReadProcessMemory(pBgrun->GetProcess(), (LPCVOID)(pBgrun->GetData())->Address, (LPVOID)Buff, CHRATTRS_COUNT, &NOBR))
		{
			if((pBgrun->GetData())->CriStop && Buff[Cri] >= (pBgrun->GetData())->CriPoints)
			{
				pBgrun->m_bRun = false;
				return TRUE;
			}
			for(sum = 0, i = 0; i < CHRATTRS_COUNT; i ++)
				sum = sum + Buff[i];
			if(sum >= (pBgrun->GetData())->Points + Buff[Cri])
			{
				pBgrun->m_bRun = false;
				return TRUE;
			}
		}
		SetCursorPos(PtScreen.x, PtScreen.y);
		mouse_event(MOUSEEVENTF_LEFTDOWN, PtScreen.x, PtScreen.y, 0, 0);
		mouse_event(MOUSEEVENTF_LEFTUP, PtScreen.x, PtScreen.y, 0, 0);
		Sleep(pBgrun->GetData()->Delay);
	}

	return FALSE;
}

/////////////////////////////////////////////////////////////////////////////
//Keyboard hook function

LRESULT CALLBACK KeyBoardHook(int nCode, WPARAM wParam, LPARAM lParam)
{
	if(nCode > -1)
	{
		if(lParam & 0x80000000)
		{
			if(wParam == NUMSTAR && clsBgrun.m_bRun == FALSE)
			{
				clsBgrun.SetProcess();
				//if game doesn't exist, it still works. : -)
				clsBgrun.m_bRun = TRUE;
				clsBgrun.StartCheck();
				return 1;
			}
			else
			{
				if(wParam == NUMSLASH && clsBgrun.m_bRun)
				{
					clsBgrun.CloseProcess();
					clsBgrun.m_bRun = FALSE;
					return 1;
				}
				else
					return 0;
			}
		}
		else
			return 0;
	}
	else
	{
		CallNextHookEx(clsBgrun.GetNextHookProc(), nCode, wParam, lParam);
		return 0;
	}
}

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved)
{
    switch (ul_reason_for_call)
	{
		case DLL_PROCESS_ATTACH:
		case DLL_THREAD_ATTACH:
			clsBgrun.m_hInstance = (HINSTANCE)hModule;
			clsBgrun.OpenShareData();
			break;
		case DLL_THREAD_DETACH:
		case DLL_PROCESS_DETACH:
			clsBgrun.CloseShareData();
			break;
    }
    return TRUE;
}

// This is the constructor of a class that has been exported.
// see Bgrun.h for the class definition
CBgrun::CBgrun()
{ 
	m_hMap = NULL;
	m_hProcess = NULL;
	m_hNextHookProc = 0;
	m_pBgData = NULL;
	m_pChrAttrs = new CChrAttrs;
	return; 
}

CBgrun::~CBgrun()
{ 
	if(m_hProcess != NULL)
	{
		CloseHandle(m_hProcess);
		m_hProcess = NULL;
	}
	if(m_hNextHookProc != 0)
	{
		UnhookWindowsHookEx(m_hNextHookProc);
   		m_hNextHookProc = 0;
	}

	return; 
}

CBgData CBgrun::SetData(const CBgData* pBgData)
{
	CBgData Swap = *m_pBgData;
	*m_pBgData = *pBgData;
	m_pBgData->Address = Swap.Address;

	return Swap;
}

CBgData* CBgrun::GetData(void)
{
	return m_pBgData;
}

void CBgrun::OpenShareData()
{
	m_hMap = CreateFileMapping((HANDLE)0xFFFFFFFF, NULL, PAGE_READWRITE, 0, sizeof(CBgData), CMapFileName);
	if(m_hMap != NULL)
	{
		m_pBgData = (CBgData* )MapViewOfFile(m_hMap, FILE_MAP_ALL_ACCESS, 0, 0, sizeof(CBgData));
		if(m_pBgData == NULL)
		{
			CloseHandle(m_hMap);
			m_hMap = NULL;
		}
	}
}

void CBgrun::CloseShareData()
{
	if(m_hMap != NULL)
	{	
		if(m_pBgData != NULL)
			UnmapViewOfFile(m_pBgData);
		CloseHandle(m_hMap);
		m_hMap = NULL;
	}
}

BOOL CBgrun::EnableHotKeyHook()
{
	if(m_hNextHookProc == 0)
	{
		m_hNextHookProc = SetWindowsHookEx(WH_KEYBOARD, KeyBoardHook, m_hInstance, 0);
		return m_hNextHookProc != 0 ? TRUE : FALSE;
	}
	else
		return FALSE;
}

BOOL CBgrun::DisableHotKeyHook()
{
	if(m_hNextHookProc != 0)
	{
		UnhookWindowsHookEx(m_hNextHookProc);
		m_hNextHookProc = 0;
	}
	
	return (m_hNextHookProc == 0);
}

HANDLE CBgrun::GetProcess()
{
	return m_hProcess;
}

BOOL CBgrun::SetProcess()
{
	DWORD	threadId, pId;
	HWND	hGame;

	if(m_hProcess != 0)
		CloseHandle(m_hProcess);
	
	hGame = FindWindow((LPCTSTR)gClsName, NULL);
	if(hGame != NULL)
	{
		threadId = GetWindowThreadProcessId(hGame,&pId);
		if(threadId != 0)
		{
			m_hProcess = OpenProcess(PROCESS_ALL_ACCESS, TRUE, pId);
			if(m_hProcess != NULL)
				return TRUE;
		}
	}

	return FALSE;
}

BOOL CBgrun::CloseProcess()
{
	if(m_hProcess != NULL)
	{
		CloseHandle(m_hProcess);
		m_hProcess = NULL;
		return TRUE;
	}
	return FALSE;
}

CChrAttrs CBgrun::Search(const CChrAttrs* pAttrs)
{
	DWORD	threadId, pId;
	HWND	hGame;
	int i, j, szKeys, Offset;
	BYTE Keys[64], *pKeys;
	BYTE* Buff;
	DWORD NOBR;
	UINT Count;

	m_pChrAttrs->Cha = 0;
	m_pChrAttrs->Con = 0;
	m_pChrAttrs->Cri = 0;
	m_pChrAttrs->Dex = 0;
	m_pChrAttrs->Int = 0;
	m_pChrAttrs->Str = 0;
	m_pChrAttrs->Wis = 0;

	if(m_hProcess != NULL)
	{
		CloseHandle(m_hProcess);
		m_hProcess = NULL;
	}

	hGame = FindWindow(gClsName, NULL);
	if(hGame != NULL)
	{
		threadId = GetWindowThreadProcessId(hGame, &pId);
		if(threadId != 0)
		{
			m_hProcess = OpenProcess(PROCESS_ALL_ACCESS, TRUE, pId);
			if(m_hProcess != NULL)
			{
				switch(m_pBgData->GameType)
				{
					case BG1:
						pKeys = KeysBg1;
						szKeys = sizeof(KeysBg1) / sizeof(BYTE);
						break;
					case TSC:
						pKeys = KeysTsc;
						szKeys = sizeof(KeysTsc) / sizeof(BYTE);
						break;
					case SOA:
						pKeys = KeysSoa;
						szKeys = sizeof(KeysSoa) / sizeof(BYTE);
						break;
					case TOB:
						pKeys = KeysTob;
						szKeys = sizeof(KeysTob) / sizeof(BYTE);
						break;
					case IWD:
						pKeys = KeysIwd;
						szKeys = sizeof(KeysIwd) / sizeof(BYTE);
						break;
					case HOW:
						pKeys = KeysHow;
						szKeys = sizeof(KeysHow) / sizeof(BYTE);
						break;
				}
				Offset = Offsets[m_pBgData->GameType];
				if(pAttrs->Str)
				{
					Keys[Str] = pAttrs->Str;
					Keys[Cri] = pAttrs->Cri;
					Keys[Dex] = pAttrs->Dex;
					Keys[Con] = pAttrs->Con;
					Keys[Int] = pAttrs->Int;
					Keys[Wis] = pAttrs->Wis;
					Keys[Cha] = pAttrs->Cha;
					pKeys = Keys;
					szKeys = 7;
					Offset = 0;
				}
				Buff = new BYTE[BUFF_SIZE + szKeys - 1]; 
				if(Buff != NULL)
				{
					Count = BEGINPOS;
					while(Count < ENDPOS)
					{
						if(ReadProcessMemory(m_hProcess, (LPCVOID)Count, (LPVOID)Buff, (DWORD)BUFF_SIZE + szKeys - 1, &NOBR))
						{
							for(i = 0; i < BUFF_SIZE; i ++)
							{
								for(j = 0; j < szKeys; j ++)
									if(pKeys[j] != Buff[i + j])
										break;
								if(j > szKeys - 1)
								{
									m_pBgData->Address = Count + i + Offset;
									ReadProcessMemory(m_hProcess, (LPCVOID)m_pBgData->Address, (LPVOID)Buff, (DWORD)CHRATTRS_COUNT, &NOBR);
									m_pChrAttrs->Str = Buff[Str];
									m_pChrAttrs->Dex = Buff[Dex];
									m_pChrAttrs->Con = Buff[Con];
									m_pChrAttrs->Int = Buff[Int];
									m_pChrAttrs->Wis = Buff[Wis];
									m_pChrAttrs->Cha = Buff[Cha];
									m_pChrAttrs->Cri = Buff[Cri];
									Count = ENDPOS;
									break;
								}
							}
						}
						Count += BUFF_SIZE;
					}
					delete Buff;
				}
				CloseHandle(m_hProcess);
				m_hProcess = NULL;
			}
		}
	}

	return *m_pChrAttrs;
}

CChrAttrs* CBgrun::GetChrAttrs()
{
	return m_pChrAttrs;
}

HHOOK CBgrun::GetNextHookProc()
{
	return m_hNextHookProc;
}

BOOL CBgrun::StartCheck()
{
	DWORD ThreadId;

	m_hThread = CreateThread(NULL, 0, &CheckChrAttrs, (LPVOID)&clsBgrun, 0, &ThreadId);
	return (m_hThread != NULL);
}