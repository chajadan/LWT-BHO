#ifndef TermRecord_INCLUDED
#define TermRecord_INCLUDED

#include <string>
#include "chajUtil.h"

using std::wstring;

struct TermRecord
{
	TermRecord(wstring wTheStatus, int nHasBeyond = -1) : wStatus(wTheStatus), nTermsBeyond(nHasBeyond) {}
	wstring wStatus;
	wstring wRomanization;
	wstring wTranslation;
	unsigned int uIdent;
	int nTermsBeyond; // -1 = unknown, 0 = no, 1 = yes
};

#endif