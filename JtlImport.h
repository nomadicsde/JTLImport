
// JtlImport.h: Hauptheaderdatei f�r die PROJECT_NAME-Anwendung
//

#pragma once

#ifndef __AFXWIN_H__
	#error "'stdafx.h' vor dieser Datei f�r PCH einschlie�en"
#endif

#include "resource.h"		// Hauptsymbole


// CJtlImportApp:
// Siehe JtlImport.cpp f�r die Implementierung dieser Klasse
//

class CJtlImportApp : public CWinApp
{
public:
	CJtlImportApp();

// �berschreibungen
public:
	virtual BOOL InitInstance();

// Implementierung

	DECLARE_MESSAGE_MAP()
};

extern CJtlImportApp theApp;