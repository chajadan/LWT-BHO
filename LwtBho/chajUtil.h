#ifndef chajUtil_INCLUDED
#define chajUtil_INCLUDED

#define VC_EXTRALEAN
#define WIN32_LEAN_AND_MEAN

#include <string>
#include <algorithm>
#include <locale>
#include "windows.h"
#include "Unknwn.h"

#ifdef _UNICODE
	typedef std::wstring tstring;
#else
	typedef std::string wstring;
#endif

#ifdef _DEBUG
bool _trace(TCHAR *format, ...);
#define TRACE _trace
#else
#define TRACE __noop
#endif

extern HWND hMBParent;
void mb(tstring msg, HWND hParent = NULL);
void mb(char* msg, HWND hParent = NULL);
void mb(char* msg, char* cap, HWND hParent = NULL);
void mb(std::wstring msg, std::wstring cap);
void PrintFormattedError(DWORD dwErr);

namespace chaj
{
	namespace str
	{
		std::wstring& wstring_tolower(std::wstring& in);
		std::wstring wstring_tolower(const std::wstring& cwstr);

		template <class T>
		std::wstring to_wstring(const T& t)
		{
			std::wstringstream ss;
			ss << t;
			return ss.str();
		}
		std::string wstring_to_string(std::wstring in);

		template <class T>
		std::string to_string(const T& t)
		{
			std::stringstream ss;
			ss << t;
			return ss.str();
		}

	}

	namespace COM
	{
		template<typename IFrom, typename ITo>
		inline ITo* GetAlternateInterface(IFrom* pIn)
		{
			ITo* pOut = nullptr;
			pIn->QueryInterface(__uuidof(ITo), reinterpret_cast<void**>(&pOut));
			return pOut;
		}

		class SmartCOMRelease
		{
		public:
			SmartCOMRelease(IUnknown* pUnknown) : pUnk(pUnknown) {};
			~SmartCOMRelease() {pUnk->Release();};
		private:
			IUnknown* pUnk;
		};
	}
}

#endif