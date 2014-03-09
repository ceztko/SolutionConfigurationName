// Win32Project2.cpp : Defines the exported functions for the DLL application.
//

#include "stdafx.h"
#include "Win32Project2.h"


// This is an example of an exported variable
WIN32PROJECT2_API int nWin32Project2=0;

// This is an example of an exported function.
WIN32PROJECT2_API int fnWin32Project2(void)
{
	return 42;
}

// This is the constructor of a class that has been exported.
// see Win32Project2.h for the class definition
CWin32Project2::CWin32Project2()
{
	return;
}
