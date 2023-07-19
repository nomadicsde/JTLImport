
// JtlImportDlg.h: Headerdatei
//

#pragma once


#define AS_HEADER_FLD_NAME               { \
  "Marketplace ID",                        \
  "Merchant ID",                           \
  "Order Date",                            \
  "Transaction Type",                      \
  "Is Invoice Corrected",                  \
  "Order ID",                              \
  "Shipment Date",                         \
  "Shipment ID",                           \
  "Transaction ID",                        \
  "ASIN",                                  \
  "SKU",                                   \
  "Quantity",                              \
  "Tax Calculation Date",                  \
  "Tax Rate",                              \
  "Product Tax Code",                      \
  "Currency",                              \
  "Tax Type",                              \
  "Tax Calculation Reason Code",           \
  "Tax Reporting Scheme",                  \
  "Tax Collection Responsibility",         \
  "Tax Address Role",                      \
  "Jurisdiction Level",                    \
  "Jurisdiction Name",                     \
  "OUR_PRICE Tax Inclusive Selling Price", \
  "OUR_PRICE Tax Amount",                  \
  "OUR_PRICE Tax Exclusive Selling Price", \
  "OUR_PRICE Tax Inclusive Promo Amount",  \
  "OUR_PRICE Tax Amount Promo",            \
  "OUR_PRICE Tax Exclusive Promo Amount",  \
  "SHIPPING Tax Inclusive Selling Price",  \
  "SHIPPING Tax Amount",                   \
  "SHIPPING Tax Exclusive Selling Price",  \
  "SHIPPING Tax Inclusive Promo Amount",   \
  "SHIPPING Tax Amount Promo",             \
  "SHIPPING Tax Exclusive Promo Amount",   \
  "GIFTWRAP Tax Inclusive Selling Price",  \
  "GIFTWRAP Tax Amount",                   \
  "GIFTWRAP Tax Exclusive Selling Price",  \
  "GIFTWRAP Tax Inclusive Promo Amount",   \
  "GIFTWRAP Tax Amount Promo",             \
  "GIFTWRAP Tax Exclusive Promo Amount",   \
  "Seller Tax Registration",               \
  "Seller Tax Registration Jurisdiction",  \
  "Buyer Tax Registration",                \
  "Buyer Tax Registration Jurisdiction",   \
  "Buyer Tax Registration Type",           \
  "Buyer E Invoice Account Id",            \
  "Invoice Level Currency Code",           \
  "Invoice Level Exchange Rate",           \
  "Invoice Level Exchange Rate Date",      \
  "Converted Tax Amount",                  \
  "VAT Invoice Number",                    \
  "Invoice Url",                           \
  "Export Outside EU",                     \
  "Ship From City",                        \
  "Ship From State",                       \
  "Ship From Country",                     \
  "Ship From Postal Code",                 \
  "Ship From Tax Location Code",           \
  "Ship To City",                          \
  "Ship To State",                         \
  "Ship To Country",                       \
  "Ship To Postal Code",                   \
  "Ship To Location Code",                 \
  "Return Fc Country",                     \
  "Is Amazon Invoiced",                    \
  "Original VAT Invoice Number",           \
  "Invoice Correction Details",            \
  "SDI Invoice Delivery Status",           \
  "SDI Invoice Error Code",                \
  "SDI Invoice Error Description",         \
  "SDI Invoice Status Last Updated Date",  \
  "EInvoice URL"                           \
}

#define AS_INDEX_MARKETPLACE          0    // FR, DE, GB                   
#define AS_INDEX_DATUM                2    // Rechnungsdatum
#define AS_INDEX_TYP                  3    // SHIPMENT, RETURN, REFUND
#define AS_INDEX_BESTELLNR_EXT        5
#define AS_INDEX_SKU                 10    // Typ
#define AS_INDEX_ANZAHL              11    // Anzahl - Betrag bezieht sich auf die Gesamtsumme
#define AS_TAX_CALC_DATE		     12    // Steuer-Datum
#define AS_INDEX_UMST                13
#define AS_INDEX_WAEHRUNG            15    // EUR, GBP
#define AS_INDEX_TAX_REP_SCHEME      18    // VCS_EU_OSS
#define AS_INDEX_PREIS_BRUTTO        23
#define AS_INDEX_PREIS_UMST          24
#define AS_INDEX_PREIS_NETTO         25
#define AS_INDEX_VERSAND_BRUTTO      29    // Versandkosten
#define AS_INDEX_VERSAND_UMST        30
#define AS_INDEX_VERSAND_NETTO       31
#define AS_INDEX_VERSAND_AKT_BRUTTO  32    // Aktion Versand
#define AS_INDEX_VERSAND_AKT_UMST    33
#define AS_INDEX_VERSAND_AKT_NETTO   34
#define AS_INDEX_GESCHENK_BRUTTO     35    // Als Geschenk
#define AS_INDEX_GESCHENK_UMST       36
#define AS_INDEX_GESCHENK_NETTO      37
#define AS_INDEX_UMST_ID_KAEUFER     43
#define AS_INDEX_UMST_ID_LAND        44
#define AS_INDEX_UMST_BERECHNET_CUR  47    // Währung für berechnete Umsatzsteuer (solte EUR sein)
#define AS_INDEX_FAKTOR_WAEHRUNG     48    // Wenn hier ein Wert enthalten ist, ist dies der Faktor * 100 für die Umrechnung [Währung] -> [Umst Eur] = Faktor * [Betrag UmSt]
#define AS_INDEX_UMST_BERECHNET      50    // Berechneter Wert der Umsatzsteuer bei Ausland-Versand
#define AS_INDEX_EXPORT_NONEU        53    // true oder false
#define AS_INDEX_EMPF_STADT          59
#define AS_INDEX_EMPF_ISO            61
#define AS_INDEX_EMPF_PLZ            62

class CTaxRechnung
{
  public:
    CTaxRechnung();

  public:
    double m_dfBetragBrutto;
    double m_dfBetragNetto;
    double m_dfBetragUmst;

    CString m_szBetragBrutto;
    CString m_szBetragNetto;
    CString m_szBetragUmst;

    CString m_szDatumOrder;
    CString m_szDatumTax;
    CString m_szBestellnummer;

    CString m_szEmpfIso;
    CString m_szTaxIso;
    CString m_szUmstID;

    CString m_szUmsatzSteuerSatz;
    CString m_szTyp;

  public:
    void SetBetragString(void);
};


typedef CMap < CString, LPCSTR, int, int>                    CMapStringToInt;
typedef CMap <CString, LPCSTR, CTaxRechnung*, CTaxRechnung*> CMapTaxRechnung;

typedef int (COMPARE)(CString& szFirst, CString& szNext, int idISO, int idDate);

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
    void DoSort(CStringArray& arr, COMPARE comp, bool fOSS);
    void DoImport(LPCSTR lpszFilePath, LPCSTR lpsPath, LPCSTR lpszName);
	  void DoImportAmazonSendung(LPCSTR lpszFilePath, LPCSTR lpsPath, LPCSTR lpszName);
	  void DoImportAmazonRemission(LPCSTR lpszFilePath, LPCSTR lpsPath, LPCSTR lpszName);
	  void DoAmazonTaxReportOSS(LPCSTR lpszFilePath, LPCSTR lpsPath, LPCSTR lpszName);
	  void DoExcelImport(LPCSTR lpszFilePath, LPCSTR lpsPath, LPCSTR lpszName);
	  void AddItemToImport(LPCSTR lpszItem, int count, CString& csvLine);
    void AddMapToLines(CMapTaxRechnung& mapRechnung, CStringArray& arrLines, bool fOSS);
	  void GetLager(void);

	protected:
		CString GetDate(LPCSTR lpszData);
    void    MapRemoveAll(CMapTaxRechnung& map);

	// Implementierung
	protected:
		CMapStringToInt   m_monthMap;
		HICON		          m_hIcon;
		CString				    m_ExportLager;
		int               m_importLines;
    int               m_nCurLagerSel;

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
	  afx_msg void OnBnClickedOssTax();
	  afx_msg void OnCbnSelchangeComboLager();
};
