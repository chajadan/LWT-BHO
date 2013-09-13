#include "chajUtil.h"
#include <sstream>
#include "comdef.h" // _com_error
#include "Strsafe.h" // StringCbVPrintf

using namespace std;

HWND hMBParent = NULL;

void mb(tstring msg, HWND hParent)
{
	if (hParent == NULL)
		MessageBox(hMBParent, msg.c_str(), _T("Msg"), MB_OK);
	else
		MessageBox(hParent, msg.c_str(), _T("Msg"), MB_OK);
}
void mb(std::wstring msg, std::wstring cap)
{
	MessageBoxW(hMBParent, msg.c_str(), cap.c_str(), MB_OK);
}
void mb(char* msg, HWND hParent)
{
	mb(msg, "Msg", hParent);
}
void mb(char* msg, char* cap, HWND hParent)
{
	if (hParent == NULL)
		MessageBoxA(hMBParent, msg, cap, MB_OK);
	else
		MessageBoxA(hParent, msg, cap, MB_OK);
}
void PrintFormattedError(DWORD dwErr)
{
	_com_error error(dwErr);
	LPCTSTR errorText = error.ErrorMessage();
	mb(errorText);
}

#ifdef _DEBUG
bool _trace(TCHAR *format, ...)
{
   TCHAR buffer[5000];

   va_list argptr;
   va_start(argptr, format);
   StringCbVPrintf(buffer, 5000, format, argptr);
   //wvsprintf(buffer, format, argptr);
   va_end(argptr);

   OutputDebugString(buffer);

   return true;
}
#endif

std::string chaj::str::wstring_to_string(std::wstring in)
{
	std::wstring_convert<std::codecvt<wchar_t, char, mbstate_t>> wc;
	std::string out = wc.to_bytes(in);
	return out;
}

std::wstring chaj::str::wstring_tolower(const std::wstring& cwstr)
{
	std::wstring in(cwstr);
	transform(in.begin(), in.end(), in.begin(), towlower);
	return in;
}
std::wstring& chaj::str::wstring_tolower(std::wstring& in)
{
	transform(in.begin(), in.end(), in.begin(), towlower);
	return in;
}