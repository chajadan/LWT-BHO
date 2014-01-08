#include "Chunc.h"

Chunc::Chunc() : mCurKey(0)
{
}

template <typename T>
void Chunc::ListenToChanges(ChuncVar<T>& var)
{
	++mCurKey;
	var->AddChangeListener(this, mCurKey);
}

template <typename T>
void Chunc::ListenToChanges(ChuncVar<T>& var, int key)
{
	var.AddChangeListener(this, key);
}