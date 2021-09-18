
// JtlImport.h: Hauptheaderdatei für die PROJECT_NAME-Anwendung
//

#pragma once

#ifndef __AFXWIN_H__
	#error "'stdafx.h' vor dieser Datei für PCH einschließen"
#endif

#include "resource.h"		// Hauptsymbole


// CJtlImportApp:
// Siehe JtlImport.cpp für die Implementierung dieser Klasse
//

class CJtlImportApp : public CWinApp
{
public:
	CJtlImportApp();

// Überschreibungen
public:
	virtual BOOL InitInstance();

// Implementierung

	DECLARE_MESSAGE_MAP()
};

extern CJtlImportApp theApp;