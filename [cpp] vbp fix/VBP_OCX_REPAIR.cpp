#include "stdafx.h"
void CenterWindow(HWND hwndWindow)
{
	int nX, nY, nScreenWidth, nScreenHeight;
	RECT rectWindow;
	
	nScreenWidth = GetSystemMetrics(SM_CXSCREEN);
	nScreenHeight = GetSystemMetrics(SM_CYSCREEN);
	
	GetWindowRect(hwndWindow, &rectWindow);
	
	nX = (nScreenWidth - (rectWindow.right - rectWindow.left)) / 2;
	nY = (nScreenHeight - (rectWindow.bottom - rectWindow.top)) / 2;
	
	SetWindowPos(hwndWindow, 0, nX, nY, 0, 0, SWP_NOZORDER | SWP_NOSIZE);
}

void ErrorHandling()
{
	_getch();
	exit(EXIT_FAILURE);
}

void ConsoleCenterText(TCHAR* Str)
{
	
	HANDLE hConsole = GetStdHandle( STD_OUTPUT_HANDLE );

	CONSOLE_SCREEN_BUFFER_INFO ScreenBuffer = {0,};
	GetConsoleScreenBufferInfo(hConsole,&ScreenBuffer);
	
	COORD pos;
    pos.X = (ScreenBuffer.dwMaximumWindowSize.X-strlen(Str)) / 2;
    pos.Y = ScreenBuffer.dwCursorPosition.Y;
	
    SetConsoleCursorPosition( hConsole, pos );
	
    LPDWORD written = NULL;
    WriteConsole( hConsole, Str, strlen(Str), written, 0 );
}
#include <sys/STAT.H>
long GetFileSize(std::string filename)
{
    struct stat stat_buf;
    int rc = stat(filename.c_str(), &stat_buf);
    return rc == 0 ? stat_buf.st_size : -1;
}
int main(int argc, char* argv[])
{
	CenterWindow(GetConsoleWindow());

	setlocale(LC_ALL, "");
	SetConsoleTitle("Visual Basic 6.0 Project OCX Fix");

	ConsoleCenterText("Visual Basic 6.0 Project OCX Fix\n");
	ConsoleCenterText("Version 0.1\n");
	ConsoleCenterText("Made by mp662002@naver.com\n");
	std::cout << std::endl;

	SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE),FOREGROUND_GREEN | 0x88);
	ConsoleCenterText("If the results fail , please run regsvr32.exe.\n");
	SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE),7);

	//std::cout << ">>  Visual Basic 6.0 Project OCX Fix" << std::endl;
	//std::cout << ">>  Version 0.1" << std::endl;
	//std::cout << ">>  Made by mp662002@naver.com" << std::endl;
	//std::cout << ">>  If the results fail , please run regsvr32.exe." << std::endl;
	

	if(argc==1)
	{
		std::cout << "[Using]" << std::endl;
		std::cout << "  " << PathFindFileName(argv[0]) << " " << "*.VBP Path" << std::endl;
		std::cout << "  " << PathFindFileName(argv[0]) << " " << "C:\\VB6\\VB6.vbp" << std::endl;
		system("PAUSE");
		return EXIT_FAILURE;
	}

	// 디버그
	//char FilePath[] = "\"G:\\Users\\root\\Desktop\\[VB6] New Process\\Project1.vbp\"";
	char FilePath[MAX_PATH] = {0,};
	lstrcpyn(FilePath,argv[1],MAX_PATH);
	PathUnquoteSpaces(FilePath);


	if (!PathFileExists(FilePath))
	{
		std::cout << "This file is not exist." << std::endl;
		ErrorHandling();
	}

	if(!(stricmp(PathFindExtension(FilePath),".vbp")==0))
	{
		std::cout << "This file is not vbp" << std::endl;
		ErrorHandling();
	}
	DWORD dwFileSize = GetFileSize(FilePath) + 60;
	char* VBP = new char[dwFileSize];
	memset(VBP,0,dwFileSize);

	char str[300];	// 한 라인에 대한 버퍼
	fstream file(FilePath,ios::in | ios::out);

	 	while(!file.eof())
		{
			memset(str,0,300);
			file.getline(str,300);

			if (strlen(str))
				if(strncmp(str,"Object",6)==0)
				{
					// strtok 는 원본 문자열을 범하기 때문에 복사해서 사용한다.
					char Temp[300]={0,};
					strcpy(Temp,str);
					char* lp = Temp;
					char* CLSID = strtok(lp+7,"#");			// 전까지 백업
					char* Version = strtok(NULL,"#");
					char* Unknown = strtok(NULL,";");	// 그 뒤 문자열
					char* OCX = strtok(NULL," ");
					// ##############################################################################

					std::cout << "Using OCX : " << OCX << "  Fix ▶ ";

					std::string SubKey;
					SubKey = "TypeLib\\";
					SubKey += CLSID;
					

					HKEY hkey;
					DWORD Result = RegOpenKeyEx(HKEY_CLASSES_ROOT,SubKey.c_str(),NULL,KEY_READ,&hkey);
					char RegVersion[260]={0,}; DWORD dwSize = 260;
					RegEnumKey(hkey,0,RegVersion,dwSize);
					RegCloseKey(hkey);
		
					if(strlen(RegVersion))
					{
						std::string Resu;
						Resu = Temp;
						Resu += "#";
						Resu += RegVersion;
						Resu += "#";
						Resu += Unknown;
						Resu += ";";
						Resu += OCX-1;
						strcat(VBP,Resu.c_str());
						strcat(VBP,"\n");
						std::cout << "Success" << std::endl;
					}
					else
					{
						std::cout << "Fail" << std::endl;
						strcat(VBP,str);
						strcat(VBP,"\n");
					}
				}
				else
				{
					strcat(VBP,str);
					strcat(VBP,"\n");
				}
		}
	file.close();


	file.open(FilePath,ios::in | ios::out);
	file << VBP;
	file.close();
	std::cout << "It was successfully saved." << std::endl;

	delete[] VBP;
	system("PAUSE");
	return EXIT_SUCCESS;
}

