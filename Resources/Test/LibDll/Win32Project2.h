// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the WIN32PROJECT2_EXPORTS
// symbol defined on the command line. This symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// WIN32PROJECT2_API functions as being imported from a DLL, whereas this DLL sees symbols
// defined with this macro as being exported.
#ifdef WIN32PROJECT2_EXPORTS
#define WIN32PROJECT2_API __declspec(dllexport)
#else
#define WIN32PROJECT2_API __declspec(dllimport)
#endif

// This class is exported from the Win32Project2.dll
class WIN32PROJECT2_API CWin32Project2 {
public:
	CWin32Project2(void);
	// TODO: add your methods here.
};

extern WIN32PROJECT2_API int nWin32Project2;

WIN32PROJECT2_API int fnWin32Project2(void);
