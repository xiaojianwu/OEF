// dllmain.cpp : 定义 DLL 应用程序的入口点。
#include "stdafx.h"


#define INITGUID // Init the GUIDS for the control...
#include "dsoframer.h"

HINSTANCE        v_hModule = NULL;   // DLL module handle
HANDLE           v_hPrivateHeap = NULL;   // Private Memory Heap
ULONG            v_cLocks = 0;      // Count of server locks
CRITICAL_SECTION v_csecThreadSynch;

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
					 )
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
		v_hModule = hModule;
		v_hPrivateHeap = HeapCreate(0, 0x1000, 0);
		InitializeCriticalSection(&v_csecThreadSynch);
		DisableThreadLibraryCalls(hModule);
		break;

	case DLL_PROCESS_DETACH:
		if (v_hPrivateHeap) HeapDestroy(v_hPrivateHeap);
		DeleteCriticalSection(&v_csecThreadSynch);
		break;
	}
	return TRUE;
}

