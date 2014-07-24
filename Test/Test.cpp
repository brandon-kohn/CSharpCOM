// Test.cpp : Defines the entry point for the console application.
//
#include "stdafx.h"
#include <tchar.h>

#include <string>

#pragma warning(disable : 4278) //disable 'mscorlib.tlb' is already a macro warning.. it's benign.
#import <mscorlib.tlb>
#import <System.tlb>
#import "..\CSharpCOM\CSharpCOM.tlb" named_guids

#include <cassert>

int _tmain(int argc, TCHAR* argv[], TCHAR* envp[])
{    
    // Initializes the COM library on the current thread and identifies the
    // concurrency model as single-thread apartment (STA). 
    CoInitializeEx(0, COINIT_APARTMENTTHREADED);//COINIT_MULTITHREADED);

    // Create an instance of the component.
    // NOTE: Make sure that registration-free components are created with 
	// their CLSID and not the ProgId.

	// Consume the properties and the methods of the COM object.
	try
    {  
        using namespace CSharpCOM;
        ICSharpFunctorPtr pCSharp;
        HRESULT hr = pCSharp.CreateInstance(__uuidof(CSharpCOM::CSharpFunctor));
        assert(SUCCEEDED(hr));
        if (SUCCEEDED(hr))
        {
            IExampleClassPtr e1 = pCSharp->RunTheFunction(10);
            assert(e1->Value == 11);
            std::string name = e1->Name;
            assert(name == "Stack Overflow");
            IExampleClass2Ptr e2 = pCSharp->RunAnotherFunction("stackoverflow");
            std::string uc = e2->UpperString;
            assert(uc == "STACKOVERFLOW");
        }
    }
    catch (_com_error &err)
    {
        _tprintf(_T("The server throws the error: %s\n"), err.ErrorMessage());
        _tprintf(_T("Description: %s\n"), (LPCTSTR) err.Description());
    }
    
    // Uninitialize COM for this thread
    CoUninitialize();
    
	return 0;
}