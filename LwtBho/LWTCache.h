#ifndef LWTCache_INCLUDED
#define LWTCache_INCLUDED

#include <unordered_map>
#include <string>
#include "chajUtil.h"
#include "TermRecord.h"

CRITICAL_SECTION CS_AddCacheEntry;
typedef std::unordered_map<wstring,TermRecord>::const_iterator cache_cit;

class LWTCache
{
public:
	LWTCache(wstring wLgID) : _cur(0), _wLgID(wLgID)
	{
		if (!InitializeCriticalSectionAndSpinCount(&CS_AddCacheEntry, 0x00000400))
			TRACE(L"Cound not initialize critical section in LWTCache. 3850duhf");
	}
	~LWTCache()
	{
		DeleteCriticalSection(&CS_AddCacheEntry);
	}
	wstring getLgID() {return _wLgID;}
	std::unordered_map<std::wstring,TermRecord>::iterator find(const std::wstring& entry) {return _cache.find(entry);}
	std::unordered_map<std::wstring,TermRecord>::const_iterator cfind(const std::wstring& entry) {return _cache.find(entry);}
	std::unordered_map<std::wstring,TermRecord>::iterator end() {return _cache.end();}
	std::unordered_map<std::wstring,TermRecord>::const_iterator cend() {return _cache.cend();}
	void insert(std::unordered_map<std::wstring,TermRecord>::value_type newEntry)
	{
		EnterCriticalSection(&CS_AddCacheEntry);
		newEntry.second.uIdent = _cur++;
		_cache.insert(newEntry);
		LeaveCriticalSection(&CS_AddCacheEntry);
	}
	
private:
	std::unordered_map<std::wstring,TermRecord> _cache;
	unsigned int _cur;
	wstring _wLgID;
};

#endif