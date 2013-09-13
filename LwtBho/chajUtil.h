#ifndef chajUtil_INCLUDED
#define chajUtil_INCLUDED

#define VC_EXTRALEAN
#define WIN32_LEAN_AND_MEAN

#include <string>
#include "tstring.h"
#include <algorithm>
#include <locale>
#include "windows.h"
#include "Unknwn.h"

//#ifdef _UNICODE
//	typedef std::wstring tstring;
//#else
//	typedef std::string tstring;
//#endif

#ifdef _DEBUG
bool _trace(TCHAR *format, ...);
#define TRACE _trace
#else
#define TRACE __noop
#endif

extern HWND hMBParent;
void mb(std::tstring msg, HWND hParent = NULL);
void mb(char* msg, HWND hParent = NULL);
void mb(char* msg, char* cap, HWND hParent = NULL);
void mb(std::wstring msg, std::wstring cap);
void PrintFormattedError(DWORD dwErr);

namespace chaj
{
	namespace util
	{
		
		/*
			CatchSentinel accepts a series of inputs and watches for a sentinel value, i.e. a trigger input.
			There are two cases:
				1) (The default case): the class notices when a particular value happens to be received
					2) The class notices when the input is not the only input you expect
			
			The sentinel type requires an accessible copy constructor.

			Typical usage example:

			string s = "<some user input text>";
			
			CatchSentinel<bool> cs(true);
			for (int i = 0; i < s.size(); ++i)
				cs += (s[i] == '\');
			
			if (cs.Seen()) // we have encounter a backslash
				...handle case
		*/
		template<class T> class CatchSentinel
		{
		public:
			CatchSentinel(T tSentinel, bool bCatchException = false): tSentinel(tSentinel), tLastException(tSentinel), bCatchException(bCatchException), bCaught(false){}
			CatchSentinel& operator+=(T tEntry)
			{ 
				bool bTriggered = (tEntry == tSentinel) ^ bCatchException;
				bCaught |= bTriggered;
				if (bCatchException && bTriggered) tLastException = tEntry;
				return *this;
			}
			// Seen() indicates whether the monitored event has occurred
			bool Seen() {return bCaught;}
			// LastException returns the seed value unless exceptions are being monitored and there was one
			T LastException() {return bCatchException && !bCaught ? tSentinel : tLastException;}
		private:
			T tSentinel;
			T tLastException;
			bool bCatchException;
			bool bCaught;
		};
	}
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