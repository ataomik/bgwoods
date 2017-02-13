// Bgrun.cpp : Defines the entry point for the DLL application.
//

#include "stdafx.h"
#include "Bgrun.h"

/////////////////////////////////////////////////////////////////////////////
//export functions
CBgData __stdcall SetData(const CBgData* pData)
{
	CBgData Swap = *pBgData;
	*pBgData = *pData;
	pBgData->Address = Swap.Address;

	return Swap;
}

CChrAttrs __stdcall Search(const CChrAttrs* pAttrs)
{	
	int i, j, szKeys, Offset;
	BYTE Keys[64], *pKeys;
	BYTE* Buff;
	DWORD NOBR;
	UINT Count;
	CChrAttrs chrAttrs;

	chrAttrs.Cha = 0;
	chrAttrs.Con = 0;
	chrAttrs.Cri = 0;
	chrAttrs.Dex = 0;
	chrAttrs.Int = 0;
	chrAttrs.Str = 0;
	chrAttrs.Wis = 0;
	

	if(SetProcess())
	{
		switch(pBgData->GameType)
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
		Offset = Offsets[pBgData->GameType];
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
				if(ReadProcessMemory(hProcess, (LPCVOID)Count, (LPVOID)Buff, (DWORD)BUFF_SIZE + szKeys - 1, &NOBR))
				{
					for(i = 0; i < BUFF_SIZE; i ++)
					{
						for(j = 0; j < szKeys; j ++)
							if(pKeys[j] != Buff[i + j])
								break;
						if(j > szKeys - 1)
						{
							pBgData->Address = Count + i + Offset;
							ReadProcessMemory(hProcess, (LPCVOID)pBgData->Address, (LPVOID)Buff, (DWORD)CHRATTRS_COUNT, &NOBR);
							chrAttrs.Str = Buff[Str];
							chrAttrs.Dex = Buff[Dex];
							chrAttrs.Con = Buff[Con];
							chrAttrs.Int = Buff[Int];
							chrAttrs.Wis = Buff[Wis];
							chrAttrs.Cha = Buff[Cha];
							chrAttrs.Cri = Buff[Cri];
							Count = ENDPOS;
							break;
						}
					}
				}
				Count += BUFF_SIZE;
			}
			delete Buff;
		}
		CloseHandle(hProcess);
		hProcess = NULL;
	}

	return chrAttrs;
}

BOOL __stdcall EnableHotKeyHook()
{
	if(hNextHookProc == 0)
	{
		hNextHookProc = SetWindowsHookEx(WH_KEYBOARD, KeyBoardHook, hInstance, 0);
		return hNextHookProc != 0 ? TRUE : FALSE;
	}
	else
		return FALSE;
}

BOOL __stdcall DisableHotKeyHook()
{
	if(hNextHookProc != 0)
	{
		UnhookWindowsHookEx(hNextHookProc);
    hNextHookProc = 0;
	}
	
	return (hNextHookProc == 0);
}

/////////////////////////////////////////////////////////////////////////////
//thread function

DWORD _stdcall CheckChrAttrs(LPVOID lparam)
{
	int i, sum;
	BYTE Buff[CHRATTRS_COUNT];
	DWORD NOBR;
	POINT PtScreen;
	
	GetCursorPos(&PtScreen);
	while(bRun)
	{
		if(ReadProcessMemory(hProcess, (LPCVOID)pBgData->Address, (LPVOID)Buff, CHRATTRS_COUNT, &NOBR))
		{
			if(pBgData->CriStop && Buff[Cri] >= pBgData->CriPoints)
			{
				bRun = FALSE;
				return TRUE;
			}
			for(sum = 0, i = 0; i < CHRATTRS_COUNT; i ++)
				sum = sum + Buff[i];
			if(sum >= pBgData->Points + Buff[Cri])
			{
				bRun = FALSE;
				return TRUE;
			}
		}
		SetCursorPos(PtScreen.x, PtScreen.y);
		mouse_event(MOUSEEVENTF_LEFTDOWN, PtScreen.x, PtScreen.y, 0, 0);
		mouse_event(MOUSEEVENTF_LEFTUP, PtScreen.x, PtScreen.y, 0, 0);
		Sleep(pBgData->Delay);
	}

	return FALSE;
}

/////////////////////////////////////////////////////////////////////////////
//Keyboard hook function

LRESULT CALLBACK KeyBoardHook(int nCode, WPARAM wParam, LPARAM lParam)
{
	DWORD ThreadId;

	if(nCode > -1)
	{
		if(lParam & 0x80000000)
		{
			if(wParam == NUMSTAR && bRun == FALSE)
			{
				SetProcess();
				//if game doesn't exist, it still works. : -)
				bRun = TRUE;
				hThread = CreateThread(NULL, 0, &CheckChrAttrs, NULL, 0, &ThreadId);
				return 1;
			}
			else
			{
				if(wParam == NUMSLASH && bRun)
				{
					CloseHandle(hProcess);
					hProcess = NULL;
					bRun = FALSE;
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
		CallNextHookEx(hNextHookProc, nCode, wParam, lParam);
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
			hInstance = (HINSTANCE)hModule;
			OpenShareData();
			break;
		case DLL_THREAD_DETACH:
		case DLL_PROCESS_DETACH:
			CloseShareData();
			break;
    }
    return TRUE;
}

void OpenShareData()
{
	hMap = CreateFileMapping((HANDLE)0xFFFFFFFF, NULL, PAGE_READWRITE, 0, sizeof(CBgData), CMapFileName);
	if(hMap != NULL)
	{
		pBgData = (CBgData* )MapViewOfFile(hMap, FILE_MAP_ALL_ACCESS, 0, 0, sizeof(CBgData));
		if(pBgData == NULL)
		{
			CloseHandle(hMap);
			hMap = NULL;
		}
	}
}

void CloseShareData()
{
	if(hMap != NULL)
	{
		if(pBgData != NULL)
			UnmapViewOfFile(pBgData);
		CloseHandle(hMap);
		hMap = NULL;
	}
}

BOOL SetProcess()
{
	DWORD	threadId, pId;
	HWND	hGame;

	if(hProcess != NULL)
	{
		CloseHandle(hProcess);
		hProcess = NULL;
	}

	hGame = FindWindow((LPCTSTR)gClsName, NULL);
	if(hGame != NULL)
	{
		threadId = GetWindowThreadProcessId(hGame,&pId);
		if(threadId != 0)
		{
			hProcess = OpenProcess(PROCESS_ALL_ACCESS, TRUE, pId);
			if(hProcess != NULL)
				return TRUE;
		}
	}

	return FALSE;
}
