#include <windows.h>

int main(int argc, char *argv[])
{
	DWORD dwReturn = 0;
	char *EnvStr = "Environment"; //This is what we want to refresh
	//Boardcast a message to inform windows to do an update
	SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, 0,  (LPARAM) EnvStr, SMTO_ABORTIFHUNG, 5000, &dwReturn);
    	return EXIT_SUCCESS; //Return good exit code
}
