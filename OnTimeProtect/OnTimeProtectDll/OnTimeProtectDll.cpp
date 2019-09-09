#include <Windows.h>
#include "MinHook.h"
#include <string>

//链接lib
#if defined _M_X64
	#pragma comment(lib, "libMinHook.x64.lib")
#elif defined _M_IX86
	#pragma comment(lib, "libMinHook.x86.lib")
#endif

//MessageBoxW
typedef int(WINAPI *MESSAGEBOXW)(HWND , LPCSTR , LPCSTR , UINT );
MESSAGEBOXW fpMessageBoxW = NULL;

//MessageBoxExW
typedef int(WINAPI *MESSAGEBOXEXW)(HWND, LPCSTR, LPCSTR, UINT, WORD);
MESSAGEBOXEXW fpMessageBoxExW = NULL;

//CreateProcessW
typedef BOOL(WINAPI *CREATEPROCESSW)(LPCWSTR, LPWSTR, LPSECURITY_ATTRIBUTES, LPSECURITY_ATTRIBUTES,	BOOL , DWORD , LPVOID , LPCWSTR , LPSTARTUPINFOW , LPPROCESS_INFORMATION );
CREATEPROCESSW fpCreateProcessW = NULL;

/*
 * MessageBoxW替代函数
 */
int WINAPI DetourMessageBoxW(
	HWND hWnd, LPCSTR lpText, LPCSTR lpCaption, UINT uType
	)
{
	return 0;
}

/*
* MessageBoxExW替代函数
*/
int WINAPI DetourMessageBoxExW(
	HWND hWnd, LPCSTR lpText, LPCSTR lpCaption, UINT uType, WORD wLanguageId
	)
{
	return 0;
}

/*
 * LPCTSTR或者LPCWSTR转换成char*
 */
char* wtoc(LPCTSTR str)
{
	DWORD dwMinSize;
	//计算长度
	dwMinSize = WideCharToMultiByte(CP_ACP, NULL, str, -1, NULL, 0, NULL, FALSE); 
	char *return_Str = new char[dwMinSize];
	WideCharToMultiByte(CP_OEMCP, NULL, str, -1, return_Str, dwMinSize, NULL, FALSE);
	return return_Str;
}
/*
 * char*转换成LPCTSTR或者LPCWSTR
 */
wchar_t* ctow(const char *str)
{
	wchar_t* buffer;
	if (str)
	{
		size_t nu = strlen(str);
		size_t n = (size_t)MultiByteToWideChar(CP_ACP, 0, (const char *)str, int(nu), NULL, 0);
		buffer = 0;
		buffer = new wchar_t[n + 1];
		MultiByteToWideChar(CP_ACP, 0, (const char *)str, int(nu), buffer, int(n));
	}
	return buffer;
	delete buffer;
}

/*
 * CreateProcessW替代函数
 */
BOOL WINAPI DetourCreateProcessW(
	LPCWSTR lpApplicationName,
	LPWSTR lpCommandLine,
	LPSECURITY_ATTRIBUTES lpProcessAttributes,
	LPSECURITY_ATTRIBUTES lpThreadAttributes,
	BOOL bInheritHandles,
	DWORD dwCreationFlags,
	LPVOID lpEnvironment,
	LPCWSTR lpCurrentDirectory,
	LPSTARTUPINFOW lpStartupInfo,
	LPPROCESS_INFORMATION lpProcessInformation) {
	
		HWND hWnd = FindWindow(NULL, L"Ty2y杀毒软件");
		if (hWnd != NULL)
		{
			//启动的程序
			char *AppName = wtoc(lpApplicationName);

			COPYDATASTRUCT cpd;
			cpd.dwData = 0;
			cpd.cbData = strlen(AppName);
			cpd.lpData = AppName;
			LRESULT ret = SendMessage(hWnd, WM_COPYDATA, (WPARAM)GetCurrentProcessId(), (LPARAM)&cpd);

			if (ret == 915)
			{
				//可运行的标识，未检测到病毒
				return fpCreateProcessW(lpApplicationName,
					lpCommandLine,
					lpProcessAttributes,
					lpThreadAttributes,
					bInheritHandles,
					dwCreationFlags,
					lpEnvironment,
					lpCurrentDirectory,
					lpStartupInfo,
					lpProcessInformation);
			}else{
				return FALSE;
			}
		}else{
			//如果没有检测到接收窗口，则不拦截
			return fpCreateProcessW(lpApplicationName,
				lpCommandLine,
				lpProcessAttributes,
				lpThreadAttributes,
				bInheritHandles,
				dwCreationFlags,
				lpEnvironment,
				lpCurrentDirectory,
				lpStartupInfo,
				lpProcessInformation);
		}
}

/*
 * 开启API HOOK
 */
int StartHook()
{
	//MinHook初始化
	if (MH_Initialize() != MH_OK)
	{
		return 1;
	}
	
	/*
	//Hook MessageBoxW函数
	if (MH_CreateHook(&MessageBoxW, &DetourMessageBoxW, reinterpret_cast<LPVOID*>(&fpMessageBoxW)) != MH_OK)
	{
		return 1;
	}
	if (MH_EnableHook(&MessageBoxW) != MH_OK)
	{
		return 1;
	}

	//Hook MessageBoxExW函数
	if (MH_CreateHook(&MessageBoxExW, &DetourMessageBoxExW, reinterpret_cast<LPVOID*>(&fpMessageBoxExW)) != MH_OK)
	{
		return 1;
	}
	if (MH_EnableHook(&MessageBoxExW) != MH_OK)
	{
		return 1;
	}
	*/

	//Hook CreateProcessW函数
	if (MH_CreateHook(&CreateProcessW, &DetourCreateProcessW, reinterpret_cast<LPVOID*>(&fpCreateProcessW)) != MH_OK)
	{
		return 1;
	}
	if (MH_EnableHook(&CreateProcessW) != MH_OK)
	{
		return 1;
	}
	
	return 0;
}

/*
* 停止API HOOK
*/
int StopHook(void) {
	/*
	//停止Hook MessageBoxW
	if (MH_DisableHook(&MessageBoxW) != MH_OK)
	{
		return 1;
	}
	//停止Hook MessageBoxExW
	if (MH_DisableHook(&MessageBoxExW) != MH_OK)
	{
		return 1;
	}
	*/

	//停止Hook CreateProcessW
	if (MH_DisableHook(&CreateProcessW) != MH_OK)
	{
		return 1;
	}

	//卸载MinHook
	if (MH_Uninitialize() != MH_OK)
	{
		return 1;
	}
}

HHOOK hHook = 0;
HINSTANCE hMod = 0;

LRESULT CALLBACK HookProc(int nCode, WPARAM wParam, LPARAM lParam) {
	return(CallNextHookEx(hHook, nCode, wParam, lParam));
}

/*
 * 导出函数，开启API HOOK
 */
BOOL WINAPI OnTimeProtectON() {

	hHook = SetWindowsHookEx(WH_GETMESSAGE, (HOOKPROC)HookProc, hMod, 0);
	if (hHook)
	{
		return TRUE;
	}
	else {
		return FALSE;
	}
}

/*
 * 导出函数，停止API HOOK
 */
BOOL WINAPI OnTimeProtectOFF() {
	return(UnhookWindowsHookEx(hHook));
}

/*
 *	Dll入口函数
 */
BOOL APIENTRY DllMain(HANDLE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved)
{
	hMod = (HINSTANCE)hModule;

	//DLL加载事件
	if (ul_reason_for_call == DLL_PROCESS_ATTACH)
	{
		StartHook();
	}

	//Dll卸载事件
	if (ul_reason_for_call == DLL_PROCESS_DETACH)
	{
		StopHook();
	}

	return TRUE;
}