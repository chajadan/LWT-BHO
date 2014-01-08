#pragma once

#include <vector>
#include <tuple>
#define WIN32_LEAN_AND_MEAN
#include "Windows.h"
#include "ChuncVar.h"

class Chunc
{
	template <typename T>
	friend class ChuncVar;

public:
	Chunc();
protected:
	template <typename T>
	void ListenToChanges(ChuncVar<T>& var);
	template <typename T>
	void ListenToChanges(ChuncVar<T>& var, int key);
	virtual void HearChunkVarChange(int key, void* value) {}

	int mCurKey;
};