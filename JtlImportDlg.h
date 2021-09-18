
// JtlImportDlg.h: Headerdatei
//

#pragma once


// CJtlImportDlg-Dialogfeld
class CJtlImportDlg : public CDialogEx
{
// Konstruktion
public:
	CJtlImportDlg(CWnd* pParent = NULL);	// Standardkonstruktor

// Dialogfelddaten
	enum { IDD = IDD_JTLIMPORT_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV-Unterstützung

public:
  void DoImport(LPCSTR lpszFilePath, LPCSTR lpsPath, LPCSTR lpszName);
  void DoImportAmazonSendung(LPCSTR lpszFilePath, LPCSTR lpsPath, LPCSTR lpszName);
  void DoImportAmazonRemission(LPCSTR lpszFilePath, LPCSTR lpsPath, LPCSTR lpszName);
  void DoExcelImport(LPCSTR lpszFilePath, LPCSTR lpsPath, LPCSTR lpszName);
  void AddItemToImport(LPCSTR lpszItem, int count, CString& csvLine);
  void GetLager(void);

// Implementierung
protected:
	HICON   m_hIcon;
	CString m_ExportLager;
	int     m_importLines;

	// Generierte Funktionen für die Meldungstabellen
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
  afx_msg void OnBnClickedOpen();
  afx_msg void OnBnClickedOpen2();
  CComboBox m_cmbLager;
  CComboBox m_cmbTyp;
  afx_msg void OnBnClickedExcelImport();
  afx_msg void OnBnClickedRemission();
};
