// Calling CoInitializeEx repeated must be balanced with CoUninitialize.
//
// Output:
//
// Run Begin
// CoInitializeEx Begin
// 'IntroToComV3.exe' (Win32) : Loaded 'C:\Windows\SysWOW64\kernel.appcore.dll'.Cannot find or open the PDB file.
// 'IntroToComV3.exe' (Win32) : Loaded 'C:\Windows\SysWOW64\msvcrt.dll'.Cannot find or open the PDB file.
// CoInitializeEx End 0x00000000
// XMLTest Begin
// CoCreateInstance Begin
// 'IntroToComV3.exe' (Win32) : Loaded 'C:\Windows\SysWOW64\clbcatq.dll'.Cannot find or open the PDB file.
// 'IntroToComV3.exe' (Win32) : Loaded 'C:\Windows\SysWOW64\msxml3.dll'.Cannot find or open the PDB file.
// 'IntroToComV3.exe' (Win32) : Loaded 'C:\Windows\SysWOW64\bcrypt.dll'.Cannot find or open the PDB file.
// CoCreateInstance End 0x00000000
// CoFreeUnusedLibrariesEx Begin
// 'IntroToComV3.exe' (Win32) : Unloaded 'C:\Windows\SysWOW64\msxml3.dll'
// CoFreeUnusedLibrariesEx End
// XMLTest End 0x00000000
// CoUninitialize Begin
// CoUninitialize End
// Run End 0x00000000

#include "stdafx.h"

//----------------------------------------------------------------------

class CDbg
{
public:
	void out(const char* fmt, ...) const
	{
		va_list args;
		va_start(args, fmt);
		char text[1024] = {};
		vsprintf_s(text, fmt, args);
		va_end(args);
		OutputDebugStringA(text);
	}
	CDbg& operator << (const char* str) { OutputDebugStringA(str); return *this; }
	CDbg& operator << (HRESULT hr) { out("0x%-8.8x", hr); return *this; }
	CDbg& operator << (int v) { out("%d", v); return *this; }
	CDbg& operator << (VARIANT_BOOL b) { return CDbg::operator<<(b == VARIANT_TRUE ? "VARIANT_TRUE" : "VARIANT_FALSE"); }
	CDbg& operator << (bool b) { return CDbg::operator<<(b ? "true" : "false"); }
	CDbg& operator << (ULONG v) { out("%d", v); return *this; }
};

CDbg dbg;

//----------------------------------------------------------------------

HRESULT XMLTest()
{
	HRESULT hr = S_OK;

	dbg << "XMLTest::CoCreateInstance Begin\n";
	CComPtr<IUnknown> pIUnknown;
	hr = pIUnknown.CoCreateInstance(OLESTR("Microsoft.XMLDOM"), NULL, CLSCTX_INPROC_SERVER);
	dbg << "XMLTest::CoCreateInstance End " << hr << "\n";
	if (FAILED(hr)) return hr;

	dbg << "XMLTest::QueryInterface Begin\n";
	CComPtr<IXMLDOMDocument> pIXMLDOMDocument;
	hr = pIUnknown->QueryInterface(&pIXMLDOMDocument);
	dbg << "XMLTest::QueryInterface End " << hr << "\n";
	if (FAILED(hr)) return hr;

	dbg << "XMLTest::LoadXML Begin\n";
	CComBSTR bstrXML("<Hello><World></World></Hello>");
	VARIANT_BOOL ok = VARIANT_FALSE;
	//hr = pIXMLDOMDocument->loadXML(CComBSTR("<Hello><World></World></Hello>"), &ok);
	hr = pIXMLDOMDocument->loadXML(bstrXML, &ok);
	if (FAILED(hr)) return hr;
	dbg << "XMLTest::LoadXML End " << hr << " " << ok << "\n";

	pIXMLDOMDocument = NULL;
	pIUnknown = NULL;
	bstrXML.Empty();

	dbg << "XMLTest::CoFreeUnusedLibrariesEx Begin\n";
	CoFreeUnusedLibrariesEx(NULL, NULL);
	Sleep(1000);
	dbg << "XMLTest::CoFreeUnusedLibrariesEx End\n";

	return hr;
}


//----------------------------------------------------------------------

HRESULT XMLTest2()
{
	HRESULT hr = S_OK;

	dbg << "XMLTest2::CoCreateInstance Begin\n";
	CComPtr<IXMLDOMDocument> pIXMLDOMDocument;
	hr = pIXMLDOMDocument.CoCreateInstance(OLESTR("Microsoft.XMLDOM"), NULL, CLSCTX_ALL);
	dbg << "XMLTest2::CoCreateInstance End " << hr << "\n";
	if (FAILED(hr)) return hr;

	pIXMLDOMDocument = NULL;

	dbg << "XMLTest2::CoFreeUnusedLibrariesEx Begin\n";
	CoFreeUnusedLibrariesEx(NULL, NULL);
	Sleep(1000);
	dbg << "XMLTest2::CoFreeUnusedLibrariesEx End\n";

	return hr;
}

//----------------------------------------------------------------------

HRESULT XMLTest3()
{
	HRESULT hr = S_OK;
	ULONG ref = 0;

	dbg << "XMLTest3::CLSIDFromProgID Begin\n";
	CLSID clsid = { };
	hr = ::CLSIDFromProgID(OLESTR("Microsoft.XMLDOM"), &clsid);
	dbg << "XMLTest3::CLSIDFromProgID End " << hr << "\n";
	if (FAILED(hr)) return hr;

	dbg << "XMLTest3::CoCreateInstance Begin\n";
	IUnknown* pIUnknown = NULL;
	hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IUnknown, (void**) &pIUnknown);
	dbg << "XMLTest3::CoCreateInstance End " << hr << "\n";
	if (FAILED(hr)) return hr;

	dbg << "XMLTest3::QueryInterface Begin " << hr << "\n";
	IXMLDOMDocument* pIXMLDOMDocument = NULL;
	hr = pIUnknown->QueryInterface(IID_IXMLDOMDocument, (void**)& pIXMLDOMDocument);
	dbg << "XMLTest3::QueryInterface End " << hr << "\n";
	if (FAILED(hr))
	{
		pIUnknown->Release();
		return hr;
	}

	dbg << "XMLTest3::Release IXMLDOMDocument Begin\n";
	ref = pIXMLDOMDocument->Release();
	pIXMLDOMDocument = NULL;
	dbg << "XMLTest3::Release IXMLDOMDocument End " << ref << "\n";

	dbg << "XMLTest3::Release IUnknown Begin\n";
	ref = pIUnknown->Release();
	pIUnknown = NULL;
	dbg << "XMLTest3::Release IUnknown End " << ref << "\n";

	dbg << "XMLTest3::CoFreeUnusedLibrariesEx Begin\n";
	CoFreeUnusedLibrariesEx(NULL, NULL);
	Sleep(1000);
	dbg << "XMLTest3::CoFreeUnusedLibrariesEx End\n";

	return hr;
}

//----------------------------------------------------------------------

HRESULT Run()
{
	HRESULT hr = S_OK;

	dbg << "Run::CoInitializeEx Begin\n";
	hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	dbg << "Run::CoInitializeEx End " << hr << "\n";
	if (FAILED(hr)) return hr;

	dbg << "Run::XMLTest Begin\n";
	hr = XMLTest3();
	dbg << "Run::XMLTest End " << hr << "\n";

	dbg << "Run::CoFreeUnusedLibrariesEx Begin\n";
	CoFreeUnusedLibrariesEx(NULL, NULL);
	Sleep(1000);
	dbg << "Run::CoFreeUnusedLibrariesEx End\n";

	dbg << "Run::CoUninitialize Begin\n";
	CoUninitialize();
	Sleep(1000);
	dbg << "Run::CoUninitialize End\n";

	return hr;
}

//----------------------------------------------------------------------

int main()
{
	dbg << "Run Begin\n";
	HRESULT hr = Run();
	dbg << "Run End " << hr << "\n";

	//dbg << "XMLTest XMLTest\n";
	//HRESULT hr = XMLTest();
	//dbg << "XMLTest End " << hr << "\n";

	return 0;
}