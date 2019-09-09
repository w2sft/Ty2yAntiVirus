
#include <Windows.h>
#include <direct.h>  
#include <stdio.h>  
#include <iostream>
#include <string>
#include <assert.h>

//修改链接设置，使运行时不出现命令行窗口
#pragma comment(linker, "/subsystem:\"windows\" /entry:\"mainCRTStartup\"")  

typedef BOOL(*OnTimeProtectON)(void);
typedef BOOL(*OnTimeProtectOFF)(void);

void ON(void){
	//加载dll
	#if defined _M_X64
		HINSTANCE hInstLibrary = LoadLibrary(TEXT("OnTimeProtectDll64.dll"));
	#elif defined _M_IX86
		HINSTANCE hInstLibrary = LoadLibrary(TEXT("OnTimeProtectDll32.dll"));
	#endif

	if (hInstLibrary == NULL)
	{
		MessageBox(NULL, L"Dll加载失败！", L"错误", MB_ICONERROR);
		FreeLibrary(hInstLibrary);
		return;
	}
	else {

		//获取API函数地址
		OnTimeProtectON MyOnTimeProtectON = (OnTimeProtectON)GetProcAddress(hInstLibrary, "OnTimeProtectON");
		if (MyOnTimeProtectON)
		{
			//启用API HOOK
			bool bHook = MyOnTimeProtectON();
			if (bHook = false)
			{
				MessageBox(NULL, L"OnTimeProtectON函数调用失败！", L"错误", MB_ICONERROR);
				return;
			}
		}
		else {
			MessageBox(NULL, L"获取OnTimeProtectOFF函数地址失败！", L"错误", MB_ICONERROR);
			return;
		}
	}
}

void OFF(void) {
	//加载dll
	#if defined _M_X64
		HINSTANCE hInstLibrary = LoadLibrary(TEXT("OnTimeProtectDll64.dll"));
	#elif defined _M_IX86
		HINSTANCE hInstLibrary = LoadLibrary(TEXT("OnTimeProtectDll32.dll"));
	#endif

	if (hInstLibrary == NULL)
	{
		MessageBox(NULL, L"Dll加载失败！", L"错误", MB_ICONERROR);
		FreeLibrary(hInstLibrary);
		return;
	}
	else {

		//获取API函数地址
		OnTimeProtectOFF MyOnTimeProtectOFF = (OnTimeProtectOFF)GetProcAddress(hInstLibrary, "OnTimeProtectOFF");
		if (MyOnTimeProtectOFF)
		{
			//停止API HOOK
			bool bHook = MyOnTimeProtectOFF();
			if (bHook = false)
			{
				MessageBox(NULL, L"OnTimeProtectOFF函数调用失败！", L"错误", MB_ICONERROR);
				return;
			}
		}
		else {
			MessageBox(NULL, L"获取OnTimeProtectOFF函数地址失败！", L"错误", MB_ICONERROR);
			return;
		}
	}
}

/*
 * 获取Ini文件内容
 */
char* getKeyValue(char *filename, char *section, char *key)
{
	char line[255];
	char sectname[255];
	char *skey, *valu;
	char seps[] = "=";
	bool flag = false;
	FILE *fp = fopen(filename, "r");
	assert(fp != NULL);

	memset(line, 0, 255);
	if (!strchr(section, '['))
	{
		strcpy(sectname, "[");
		strcat(sectname, section);
		strcat(sectname, "]");
	}
	else
	{
		strcpy(sectname, section);
	}


	while (fgets(line, 255, fp) != NULL)
	{
		//delete the  newline
		valu = strchr(line, '\n');
		*valu = 0;

		if (flag)
		{
			skey = strtok(line, seps);
			if (strcmp(skey, key) == 0)
			{
				//一定要关闭文件后再退出，否则会造成内存泄漏
				fclose(fp);
				return strtok(NULL, seps);
			}
		}
		else
		{
			if (strcmp(sectname, line) == 0)
			{
				flag = true;
			}

		}
	}
	//一定要关闭文件后再退出，否则会造成内存泄漏
	fclose(fp);
	return NULL;
}

/*
 * 获取防护值
 */
char* getShieldValue(void) {
	char szModuleFilePath[MAX_PATH];
	char SaveResult[MAX_PATH];
	//获得当前执行文件路径
	int n = GetModuleFileNameA(0, szModuleFilePath, MAX_PATH);
	//将最后一个"\\"后的字符置为0  
	szModuleFilePath[strrchr(szModuleFilePath, '\\') - szModuleFilePath + 1] = 0;
	strcpy(SaveResult, szModuleFilePath);
	//在当前路径后添加文件名
	strcat(SaveResult, "\\settings.ini");
	//获取值
	char * sShieldValue = getKeyValue(SaveResult, "Shield", "EnableShield");
	return sShieldValue;
}

HANDLE hMutex;
void main(void)
{
	hMutex = CreateMutex(NULL, FALSE, L"OnTimeProtectDll");
	if (GetLastError() == ERROR_ALREADY_EXISTS) {
		CloseHandle(hMutex);
		return;
	}

	//开启API HOOK
	ON();

	char* shieldValue;

	//循环
	while(1){
		//暂停0.1秒
		Sleep(100);

		shieldValue = getShieldValue();

		//配置文件中此项目为0表示关闭了防护功能，停止 API HOOK
		if (strcmp(shieldValue, "0") == 0) {

			//停止API HOOK
			OFF();

			Sleep(100);

			//退出程序
			return;
		}

		//主程序未启动或已经退出，停止API HOOK
		HWND hWnd = FindWindow(NULL, L"Ty2y杀毒软件");
		if (hWnd == NULL)
		{
			//停止API HOOK
			OFF();

			Sleep(100);

			//退出程序
			return;
		}
	}

	return;
}