

// JtlImportDlg.cpp: Implementierungsdatei
//

#include "stdafx.h"
#include "JtlImport.h"
#include "JtlImportDlg.h"
#include "afxdialogex.h"
#include "CSVFile.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CJtlImportDlg-Dialogfeld

#define LINE_IMPORT_HEADER  _T("Artikel;Anzahl;Lager;Hinzu\r\n");

#define LAGER_NAME_FBA      _T("FBA: Nomadics")
#define LAGER_NAME_RUECK    _T("Rücklaufer")
#define LAGER_NAME_TROPLO   _T("Lager Troplowitz") 
#define LAGER_NAME_RESELLER _T("Reseller")
#define LAGER_NAME_MAASHOLM _T("Maasholm")

#define DEF_LAGER           LAGER_NAME_TROPLO

#define CMB_LAGER_TROPLO              0
#define CMB_LAGER_FBA                 1
#define CMB_LAGER_MAASHOLM            2
#define CMB_LAGER_MAAS_TO_TROPLO      3
#define CMB_LAGER_TROPLO_TO_MAAS      4
#define CMB_LAGER_RUEK                5
#define CMB_LAGER_RESELLER            6




CJtlImportDlg::CJtlImportDlg(CWnd* pParent /*=NULL*/)
	          :CDialogEx(CJtlImportDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);

    m_monthMap["Jan"] = 1;
    m_monthMap["Feb"] = 2;
    m_monthMap["Mar"] = 3;
    m_monthMap["Apr"] = 4;
    m_monthMap["May"] = 5;
    m_monthMap["Jun"] = 6;
    m_monthMap["Jul"] = 7;
    m_monthMap["Aug"] = 8;
    m_monthMap["Sep"] = 9;
    m_monthMap["Oct"] = 10;
    m_monthMap["Nov"] = 11;
    m_monthMap["Dec"] = 12;

    m_nCurLagerSel = -1;

}

void CJtlImportDlg::DoDataExchange(CDataExchange* pDX)
{
    CDialogEx::DoDataExchange(pDX);
    DDX_Control(pDX, IDC_COMBO_LAGER, m_cmbLager);
    DDX_Control(pDX, IDC_COMBO_TYPE, m_cmbTyp);
}

BEGIN_MESSAGE_MAP(CJtlImportDlg, CDialogEx)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
    ON_BN_CLICKED(IDC_OPEN, &CJtlImportDlg::OnBnClickedOpen)
    ON_BN_CLICKED(IDC_OPEN2, &CJtlImportDlg::OnBnClickedOpen2)
    ON_BN_CLICKED(IDC_EXCEL_IMPORT, &CJtlImportDlg::OnBnClickedExcelImport)
    ON_BN_CLICKED(IDC_REMISSION, &CJtlImportDlg::OnBnClickedRemission)
    ON_BN_CLICKED(IDC_OSS_TAX, &CJtlImportDlg::OnBnClickedOssTax)
  ON_CBN_SELCHANGE(IDC_COMBO_LAGER, &CJtlImportDlg::OnCbnSelchangeComboLager)
END_MESSAGE_MAP()


// CJtlImportDlg-Meldungshandler

BOOL CJtlImportDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Symbol für dieses Dialogfeld festlegen. Wird automatisch erledigt
	//  wenn das Hauptfenster der Anwendung kein Dialogfeld ist
	SetIcon(m_hIcon, TRUE);			// Großes Symbol verwenden
	SetIcon(m_hIcon, FALSE);		// Kleines Symbol verwenden
    
    m_cmbLager.SetCurSel(0);
    m_cmbTyp.SetCurSel(1);

	return TRUE;  // TRUE zurückgeben, wenn der Fokus nicht auf ein Steuerelement gesetzt wird
}

// Wenn Sie dem Dialogfeld eine Schaltfläche "Minimieren" hinzufügen, benötigen Sie
//  den nachstehenden Code, um das Symbol zu zeichnen. Für MFC-Anwendungen, die das 
//  Dokument/Ansicht-Modell verwenden, wird dies automatisch ausgeführt.

void CJtlImportDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // Gerätekontext zum Zeichnen

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Symbol in Clientrechteck zentrieren
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Symbol zeichnen
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// Die System ruft diese Funktion auf, um den Cursor abzufragen, der angezeigt wird, während der Benutzer
//  das minimierte Fenster mit der Maus zieht.
HCURSOR CJtlImportDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


#define GROESSE_SIZE_MIN 34
#define GROESSE_SIZE_MAX 46

typedef struct tagARTIKELCOUNT
{
  CString szName;                                  // JC Camel
  int     countSize[GROESSE_SIZE_MAX-GROESSE_SIZE_MIN+1];
} ARTIKELCOUNT;

typedef struct tagARTFIND
{
  CString szBezeichnung;                           // JC Camel
  int     groesse;
} ARTFIND;

// Zuordnung zur Map:
// DefName   -> Struktur
// JCxxCamel -> ARTIKELCOUNT
typedef CMap<CString, LPCSTR, ARTIKELCOUNT, ARTIKELCOUNT> CArtMap;
typedef CMap<CString, LPCSTR, ARTFIND, ARTFIND>           CArtFindMap;


typedef struct tagARTPAIR
{
  LPCSTR lpszBezeichnung;
  LPCSTR lspzFormat;
} ARTPAIR;

class CArtikel
{
  public:
    CArtikel()  {};

  public:
    void    AddType(LPCSTR lpszArtikelBezeichnung, LPCSTR lpszFormatArtikelNummer);
    int     AddArtikel(LPCSTR lpszArtikelNr,  int artikelMenge);
    CString GetFirstLine();
    CString GetArtikelLine(LPCSTR lpszArtikel);

  protected:
    CArtMap     m_mapArtikel;
    CArtFindMap m_mapFinder;
    

};

// Aufruf
// AddType("JC Camel", "jc%2dcamel");
void CArtikel::AddType(LPCSTR lpszArtikelBezeichnung, LPCSTR lpszFormatArtikelNummer)
{
  ARTIKELCOUNT artCount;
  ARTFIND      find;
  CString      szBezeichnung; 
  int          groessenCount;

  // Zunächst Finder anlegen 
  find.szBezeichnung = lpszArtikelBezeichnung;
  for (int groesse=GROESSE_SIZE_MIN; groesse<=GROESSE_SIZE_MAX; groesse++)
  {
     szBezeichnung.Format(lpszFormatArtikelNummer, groesse);
     find.groesse = groesse;
     m_mapFinder[szBezeichnung] = find;
  }

  // Nun die Artikel-Map
  groessenCount   = GROESSE_SIZE_MAX-GROESSE_SIZE_MIN+1;
  artCount.szName = lpszArtikelBezeichnung;
  for (int i=0; i<groessenCount; i++)
    artCount.countSize[i] = 0;
  
  m_mapArtikel[lpszArtikelBezeichnung]=artCount;

}

int CArtikel::AddArtikel(LPCSTR lpszArtikelNr,  int artikelMenge)
{
   // Suche nun nach dem Artkel und finde Nummer und Bezeichnung ...
  ARTIKELCOUNT artCount;
  ARTFIND      artFind;
  int          Index;
  int          menge = 0;
  
  if (m_mapFinder.Lookup(lpszArtikelNr, artFind) && m_mapArtikel.Lookup(artFind.szBezeichnung, artCount) && (Index = artFind.groesse - GROESSE_SIZE_MIN) >= 0)
  {
    artCount.countSize[Index]           +=artikelMenge;
    m_mapArtikel[artFind.szBezeichnung]  = artCount;
    menge                                = artikelMenge;
  }

  return menge;
}


CString CArtikel::GetFirstLine()
{
  CString szLine, szGroesse;
  szLine = "\"Artikel\"";
  for (int groesse=GROESSE_SIZE_MIN; groesse<=GROESSE_SIZE_MAX; groesse++)
  {
    szGroesse.Format(";\"%d\"", groesse);
    szLine += szGroesse;
  }
  szLine += "\r\n";
  return szLine;
}

CString CArtikel::GetArtikelLine(LPCSTR lpszArtikel)
{
  
  ARTIKELCOUNT artCount;
  CString      szLine, szCount;
  int          groessenCount;
  
  szLine = "";
  if (m_mapArtikel.Lookup(lpszArtikel, artCount))
  {
    szLine.Format("\"%s\"", lpszArtikel);
    groessenCount   = GROESSE_SIZE_MAX-GROESSE_SIZE_MIN+1;
    for (int i=0; i<groessenCount; i++)
    {
      szCount.Format(";\"%d\"", artCount.countSize[i]);
      szLine += szCount;
    }
    szLine += "\r\n";
  }
  return szLine;
}



#define FELDNAME_ARTIKEL     "Artikelnummer"
#define FELDNAME_MENGE       "Menge"
#define FELDNAME_VEWRFUEGBAR "Verfügbar"
#define FELDNAME_STAHLMANN   "Lagerbestand Stahlmann_Sachs"
#define FELDNAME_LG_TROPLO   "Lagerbestand Lager Troplowitz"
#define FELDNAME_MAASHOLM    "Lagerbestand Lager Maasholm"

// #define FELDNAME_LG_FBA      "Lagerbestand FBA: Nomadics"
#define FELDNAME_RESELLER    "Lagerbestand Reseller"
#define FELDNAME_LG_FBA       FELDNAME_STAHLMANN

#define ARTIKEL_NAME_ARR                                               \
{                                                                      \
  {"JC camel","jc%2dcamel"},                                           \
  {"JC sage","jc%2dsage"},                                             \
  {"JC cafe","jc%2dcafe"},                                             \
  {"JC black","jc%2dblack"},                                           \
  {"JC red","jc%2dred"},                                               \
  {"JC pumpkin","jc%2dpumpkin"},                                       \
  {"JC brick burgundy","jc%2dburgundy"},                               \
  {"San Juan Camel","sj%2dcamel"},                                     \
  {"JC camel mit Sohle","xxjc%2dcamel"},                               \
  {"JC black mit Sohle","xxjc%2dblack"},                               \
  {"Mountain Momma beige","mm-camel-%2d"},                             \
  {"Doppel Decker Platform JC camel","platform-camel-%2d"},            \
  {"Slip on camel","slip-camel-%2d"},                                  \
  {"Romano camel","rom%2dcamel"},                                      \
  {"Flower Flop","ff%2dtan-or"},                                       \
  {"Jester camel","jester%2dcamel"},                                   \
  {"Toe Joe camel","tj-camel-%2d"},                                    \
  {"Toe Joe navy", "tj-navy-%2d"},                                     \
  {"JC Disco Black","jc%2ddiscoblack"},                                \
  {"Slip on black","slip-black-%2d"},                                  \
  {"Slip on gray", "slip-gray-%2d"},                                   \
  {"Slip on denim", "slip-denim-%2d"},                                 \
  {"JC gray", "jc%2dgray"}                                             \
}


void CJtlImportDlg::AddItemToImport(LPCSTR lpszItem, int count, CString& csvLine)
{
    if (NULL == lpszItem)
    {
        m_importLines = 0;
        csvLine = LINE_IMPORT_HEADER;
    }
    else if (m_importLines >= 0)
    {
        CString szLine;
        szLine.Format("%s;%d;%s;Y\r\n", lpszItem, count, m_ExportLager);
        csvLine += szLine;
        m_importLines += 1;
    }
}




#define FIELD_NAME__ORDER_ID  _T("order-id")
#define FIELD_NAME__SKU       _T("sku")
#define FIELD_NAME_EINHEITEN  _T("shipped-quantity")

void CJtlImportDlg::DoImportAmazonRemission(LPCSTR lpszFilePath, LPCSTR lpsPath, LPCSTR lpszName)
{
    ARTPAIR      arrArtPair[] = ARTIKEL_NAME_ARR;
    CCSVFile     csv(lpszFilePath, ';');
    CStringArray arrFields;
    CString      szMenge, szNewCsv, szNewFilePath, szFileInfoPath, szType, szMissing, szJTLImport;
    CArtikel     artikel;
    bool         fHeaderFound, fSameOrder;
    int          i, pairCount, length;
    int          indexSKU, indexOrderId, indexCount, minIndexCount, totalMenge, TotalEinheiten, sollmenge, istmenge, missingtotal;
    CStringArray arrMissing;
    CString      orderID;

    m_importLines = -1;

    GetLager();

    // Typ: hier immer auf addieren ...
    m_cmbTyp.SetCurSel(0);

    pairCount = sizeof(arrArtPair) / sizeof(ARTPAIR);
    fHeaderFound    = false;
    indexSKU        = -1;
    indexCount      = -1;
    indexOrderId    = -1;
    TotalEinheiten  = 0;
    totalMenge      = 0;
    missingtotal    = 0;
    
    szType = "_AmazonRemission";

    // Initialisiere Header ...
    AddItemToImport(NULL, 0, szJTLImport);

    while (csv.ReadData(arrFields))
    {
        if (!fHeaderFound)
        {
            for (i = 0; i < arrFields.GetCount() && (indexSKU < 0 || indexCount < 0); i++)
            {
                if (indexOrderId < 0 && arrFields[i].Find(FIELD_NAME__ORDER_ID) >= 0)
                    indexOrderId = i;
                else if (indexSKU < 0 && arrFields[i].Find(FIELD_NAME__SKU) >= 0)
                    indexSKU = i;
                else if (indexCount < 0 && arrFields[i].Find(FIELD_NAME_EINHEITEN) >= 0)
                    indexCount = i;
            }
            fHeaderFound = indexOrderId >= 0 && indexCount >= 0 && indexSKU >= 0;
            if (fHeaderFound)
            {
                minIndexCount = max(indexCount, indexSKU) + 1;
                for (i = 0; i < pairCount; i++)
                    artikel.AddType(arrArtPair[i].lpszBezeichnung, arrArtPair[i].lspzFormat);
            }
            else
            {
                indexSKU = -1;
                indexCount = -1;
            }
        }
        else
        {
            if (orderID.IsEmpty())
                orderID = arrFields[indexOrderId];
            fSameOrder = orderID == arrFields[indexOrderId];
            if (fSameOrder && arrFields.GetCount() >= minIndexCount)
            {
                sollmenge = atoi(arrFields[indexCount]);
                TotalEinheiten += sollmenge;
                istmenge = artikel.AddArtikel(arrFields[indexSKU], sollmenge);
                AddItemToImport(arrFields[indexSKU], sollmenge, szJTLImport);
                totalMenge += istmenge;
                if (sollmenge != istmenge)
                {
                    szMissing.Format("- % 2d * %s", sollmenge, arrFields[indexSKU]);
                    arrMissing.Add(szMissing);
                    missingtotal += (sollmenge - istmenge);
                }

            }

        }

    }

    // Schreibe Import-Datei 
    if (m_importLines > 0)
    {

        szNewFilePath.Format("%s\\JTL_Import_%s_%s.csv", lpsPath, lpszName, szType);

        length = szJTLImport.GetLength();
        CFile fl;
        if (fl.Open(szNewFilePath, CFile::modeCreate | CFile::modeWrite))
        {
            fl.Write((LPCSTR)szJTLImport, length);
            fl.Close();
        }


    }

    if (totalMenge > 0)
    {
        // Wenn bis hier, dann schreiben ..
        szNewCsv = artikel.GetFirstLine();
        for (i = 0; i < pairCount; i++)
            szNewCsv += artikel.GetArtikelLine(arrArtPair[i].lpszBezeichnung);

        szNewFilePath.Format("%s\\%s_%s.csv", lpsPath, lpszName, szType);
        szFileInfoPath.Format("%s\\%s_%s.txt", lpsPath, lpszName, szType);

        length = szNewCsv.GetLength();
        CFile fl;
        if (fl.Open(szNewFilePath, CFile::modeCreate | CFile::modeWrite))
        {
            fl.Write((LPCSTR)szNewCsv, length);
            fl.Close();
        }

        CString szMsg, szInfo;
        // if (fTotalEinheiten)
        {
            if (totalMenge == TotalEinheiten)
                szInfo.Format("Vorgabe Gesamtmenge: %d\nGelese Einheiten: %d\nAlle Artikel wurden erfasst.", TotalEinheiten, totalMenge);
            else
            {
                CString szMissing;
                for (int i = 0; i < arrMissing.GetCount(); i++)
                {
                    szMissing += arrMissing[i];
                    szMissing += "\n";
                }
                szInfo.Format("Vorgabe Gesamtmenge: %d\r\nGelese Einheiten: %d\r\n\r\nACHTUNG\r\n%d Artikel wurden laut Exportdatei nicht erfasst.\r\n\r\nErmittelte Gesamtzahl der fehlenden Artikel: %d\r\nNicht erfasst:\r\n%s", TotalEinheiten, totalMenge, TotalEinheiten - totalMenge, missingtotal, szMissing);
            }

        }
        /*
        else
        {
            szMsg.Format("Insgesamt wurden %d Einheiten gesendet.", totalMenge);
        }
        */
        szMsg.Format("Die Datei %s wurde erzeugt.\nAls Info die Datei %s.\n\n%s", szNewFilePath, szFileInfoPath, szInfo);
        MessageBox(szMsg, "Info", MB_OK);

        length = szInfo.GetLength();
        if (fl.Open(szFileInfoPath, CFile::modeCreate | CFile::modeWrite))
        {
            fl.Write((LPCSTR)szInfo, length);
            fl.Close();
        }

    }
    else
        MessageBox("Die Datei enthält keine gültigen Remissions-Daten", "Info", MB_OK);

}



#define FIELD_NAME__AMAZON_SKU        _T("SKU")
#define FIELD_NAME_AMAZON_EINHEITEN   _T("Versendete Einheiten")
#define FIELD_EINHEITEN_TOTAL         _T("Einheiten insgesamt")

void CJtlImportDlg::DoImportAmazonSendung(LPCSTR lpszFilePath, LPCSTR lpsPath, LPCSTR lpszName)
{
    ARTPAIR      arrArtPair[] = ARTIKEL_NAME_ARR;
    CCSVFile     csv(lpszFilePath, '\t');
    CStringArray arrFields;
    CString      szMenge, szNewCsv, szNewFilePath, szType, szMissing, szJTLImport;
    CArtikel     artikel;
    bool         fHeaderFound, fTotalEinheiten;
    int          i, pairCount, length;
    int          indexSKU, indexCount, minIndexCount, totalMenge, TotalEinheiten, sollmenge, istmenge, missingtotal;
    CStringArray arrMissing;
    
    m_importLines   = -1;

    GetLager();
    
    // Typ: hier immer auf abziehen ...
    m_cmbTyp.SetCurSel(1);

    pairCount       = sizeof(arrArtPair) / sizeof(ARTPAIR);
    fHeaderFound    = false;
    fTotalEinheiten = false;
    indexSKU        = -1;
    indexCount      = -1;
    TotalEinheiten  =  0;
    totalMenge      =  0;
    missingtotal    =  0;
    szType          ="_AmazonVersand";

    // Initialisiere Header ...
    AddItemToImport(NULL, 0, szJTLImport);

    while (csv.ReadData(arrFields))
    {
        if (!fHeaderFound)
        {
            for (i = 0; i < arrFields.GetCount() && (indexSKU < 0 || indexCount < 0); i++)
            {
                if (indexSKU < 0 && arrFields[i].Find(FIELD_NAME__AMAZON_SKU) >= 0)
                   indexSKU = i;
                else if (indexCount < 0 && arrFields[i].Find(FIELD_NAME_AMAZON_EINHEITEN) >= 0)
                   indexCount = i;
                if (!fTotalEinheiten && (fTotalEinheiten = (arrFields[i].Find(FIELD_EINHEITEN_TOTAL) >= 0)) && i + 1 < arrFields.GetCount())
                    TotalEinheiten = atoi(arrFields[i + 1]);
            }
            fHeaderFound  = indexCount >= 0 && indexSKU >= 0;
            if (fHeaderFound)
            {
                minIndexCount = max(indexCount, indexSKU) + 1;
                for (i = 0; i < pairCount; i++)
                    artikel.AddType(arrArtPair[i].lpszBezeichnung, arrArtPair[i].lspzFormat);
            }
            else
            {
                indexSKU   = -1;
                indexCount = -1;
            }
        }
        else
        {
            if (arrFields.GetCount() >= minIndexCount)
            {
                sollmenge = atoi(arrFields[indexCount]);
                istmenge  = artikel.AddArtikel(arrFields[indexSKU], sollmenge);
                AddItemToImport(arrFields[indexSKU], -sollmenge, szJTLImport);
                totalMenge += istmenge;
                if (sollmenge != istmenge)
                {
                    szMissing.Format("- % 2d * %s", sollmenge, arrFields[indexSKU]);
                    arrMissing.Add(szMissing);
                    missingtotal += (sollmenge -istmenge);
                }

            }

        }

    }

    // Schreibe Import-Datei 
    if (m_importLines > 0)
    {
        
        szNewFilePath.Format("%s\\JTL_Import_%s_%s.csv", lpsPath, lpszName, szType);

        length = szJTLImport.GetLength();
        CFile fl;
        if (fl.Open(szNewFilePath, CFile::modeCreate | CFile::modeWrite))
        {
            fl.Write((LPCSTR)szJTLImport, length);
            fl.Close();
        }


    }

    if (totalMenge > 0)
    {
        // Wenn bis hier, dann schreiben ..
        szNewCsv = artikel.GetFirstLine();
        for (i = 0; i < pairCount; i++)
            szNewCsv += artikel.GetArtikelLine(arrArtPair[i].lpszBezeichnung);

        szNewFilePath.Format("%s\\%s_%s.csv", lpsPath, lpszName, szType);

        length = szNewCsv.GetLength();
        CFile fl;
        if (fl.Open(szNewFilePath, CFile::modeCreate | CFile::modeWrite))
        {
            fl.Write((LPCSTR)szNewCsv, length);
            fl.Close();
        }

        CString szMsg, szInfo;
        if (fTotalEinheiten)
        {
            if (totalMenge == TotalEinheiten)
                szInfo.Format("Vorgabe Gesamtmenge: %d\nGelese Einheiten: %d\nAlle Artikel wurden erfasst.", TotalEinheiten, totalMenge);
            else
            {
                CString szMissing;
                for (int i=0; i<arrMissing.GetCount(); i++)
                {
                    szMissing += arrMissing[i];
                    szMissing += "\n";
                }
                szInfo.Format("Vorgabe Gesamtmenge: %d\nGelese Einheiten: %d\n\nACHTUNG\n%d Artikel wurden laut Exportdatei nicht erfasst.\n\nErmittelte Gesamtzahl der fehlenden Artikel: %d\nNicht erfasst:\n%s", TotalEinheiten, totalMenge, TotalEinheiten - totalMenge, missingtotal, szMissing);
            }

        }
        else
        {
            szMsg.Format("Insgesamt wurden %d Einheiten gesendet.", totalMenge);
        }

        szMsg.Format("Die Datei %s wurde erzeugt.\n\n%s", szNewFilePath, szInfo);
        MessageBox(szMsg, "Info", MB_OK);

    }
    else
        MessageBox("Die Datei enthält keine gültigen Lagerversand-Daten", "Info", MB_OK);

}

// Lagerbestand Lager Troplowitz
// Lagerbestand FBA: Nomadics

bool LagerTroploUndFBA(CStringArray& arrFields, int &indexTroplo, int& indexFBA, int &indexMaasholm, int &indexVerfuegbar)
{
    indexTroplo     = -1;
    indexFBA        = -1;
    indexVerfuegbar = -1;
    indexMaasholm   = -1;

    for (int i = 0; i < arrFields.GetCount() && (indexTroplo < 0 || indexFBA < 0 || indexVerfuegbar < 0 || indexMaasholm < 0); i++)
    {
        if (indexTroplo < 0 && arrFields[i] == FELDNAME_LG_TROPLO)
            indexTroplo = i;
        else if (indexFBA < 0 && arrFields[i] == FELDNAME_LG_FBA)
            indexFBA = i;
        else if (indexMaasholm < 0 && arrFields[i] == FELDNAME_MAASHOLM)
          indexMaasholm = i;
        else if(indexVerfuegbar < 0 && arrFields[i] == FELDNAME_VEWRFUEGBAR)
            indexVerfuegbar = i;
    }
    return indexTroplo > 0 && indexFBA > 0 && indexVerfuegbar > 0;

}

void CJtlImportDlg::DoImport(LPCSTR lpszFilePath, LPCSTR lpsPath, LPCSTR lpszName)
{
   
  
    ARTPAIR      arrArtPair[] = ARTIKEL_NAME_ARR;
    CCSVFile     csv(lpszFilePath);
    CArtikel     artikel;
    CArtikel     artikelTroplowitz;
    CArtikel     artikelMaasholm;

    CStringArray arrFields;
    CString      szArtikel, szMenge, szNewCsv, szNewFilePath, szNewCsvTroplo, szNewCsvMaasholm, szNewFilePathTroplo, szNewFilePathMaasholm, szType;
    int          arrNrArtikel, arrNrCount;
    int          indexTroplo, indexFBA, indexVerfuegbar, indexMaasholm;
    int          i, count, pairCount, length, totalMenge = 0, lineCount = 0;
    bool         fVerkauf, fMitZweiLager, artikelSum;

    arrNrArtikel  = -1;
    arrNrCount    = -1;
    fVerkauf      = true;
    fMitZweiLager = false;


    pairCount    = sizeof(arrArtPair) / sizeof(ARTPAIR);
    
    while (csv.ReadData(arrFields))
    {
      if (0 == lineCount)
      {
        // Analysiere die erste Zeile ..
        for (i=0; i<arrFields.GetCount() && (arrNrCount < 0 || arrNrArtikel < 0); i++)
        {
          if (arrNrArtikel < 0 && arrFields[i] == FELDNAME_ARTIKEL)
            arrNrArtikel = i;
          else if (arrNrCount < 0 && arrFields[i]== FELDNAME_MENGE)
          {
            arrNrCount = i;
            szType     = "_Verkauf";
          }
          else if (arrNrCount < 0 && arrFields[i]== FELDNAME_VEWRFUEGBAR)
          {
            fMitZweiLager = LagerTroploUndFBA(arrFields, indexTroplo, indexFBA, indexMaasholm, indexVerfuegbar);
            arrNrCount    = i;
            szType        = "_Lager";
            fVerkauf      = false;
          }
        }
        if (arrNrCount < 0 || arrNrArtikel < 0)
        {
          CString szMsg;
          szMsg.Format("Dies ist nicht die korrekte Datei.\n\nEs wird eine csv-Datei erwwartet mit mindestens folgenden Feldern:\n%s\n%s\n\nDer Vorgang wird abgebrochen.", FELDNAME_ARTIKEL FELDNAME_MENGE);
          MessageBox(szMsg, "Achtung", MB_OK | MB_ICONEXCLAMATION);
          return;  
        }
        else
        {
          
          for (i=0; i<pairCount; i++)
          {
            artikel.AddType(arrArtPair[i].lpszBezeichnung, arrArtPair[i].lspzFormat);
            artikelTroplowitz.AddType(arrArtPair[i].lpszBezeichnung, arrArtPair[i].lspzFormat);
            artikelMaasholm.AddType(arrArtPair[i].lpszBezeichnung, arrArtPair[i].lspzFormat);
          }
        }

      }
      else
      {
        
        count     = arrFields.GetCount();
        if (fMitZweiLager)
        {
            if (arrNrCount < count && arrNrArtikel < count && arrFields[arrNrArtikel].GetLength() > 0)
            {
                artikelSum = 0;
                if (arrFields[indexTroplo].GetLength() > 0)
                    artikelSum += artikelTroplowitz.AddArtikel(arrFields[arrNrArtikel],  atoi(arrFields[indexTroplo]));
                if (arrFields[indexFBA].GetLength() > 0)
                    artikelSum += artikel.AddArtikel(arrFields[arrNrArtikel], atoi(arrFields[indexFBA]));
                if (indexMaasholm >= 0 && arrFields[indexMaasholm].GetLength() > 0)
                  artikelSum += artikelMaasholm.AddArtikel(arrFields[arrNrArtikel], atoi(arrFields[indexMaasholm]));
                
                totalMenge += artikelSum;
            }
        }
        else
        {
            if (arrNrCount < count && arrNrArtikel < count && arrFields[arrNrCount].GetLength() > 0 && arrFields[arrNrArtikel].GetLength() > 0)
                totalMenge += artikel.AddArtikel(arrFields[arrNrArtikel], atoi(arrFields[arrNrCount]));
        }
      }
      lineCount++;
    }

    // Wenn bis hier, dann schreiben ..
    szNewCsv = artikel.GetFirstLine();
    for (i=0; i<pairCount; i++)
      szNewCsv += artikel.GetArtikelLine(arrArtPair[i].lpszBezeichnung);
    
    if (fMitZweiLager)
    {
        szNewFilePath.Format("%s\\%s_FBA_%s.csv", lpsPath, lpszName, szType);
        szNewFilePathTroplo.Format("%s\\%s_Troplo_%s.csv", lpsPath, lpszName, szType);

        szNewCsvTroplo = artikelTroplowitz.GetFirstLine();
        for (i = 0; i < pairCount; i++)
            szNewCsvTroplo += artikelTroplowitz.GetArtikelLine(arrArtPair[i].lpszBezeichnung);
        
        if (indexMaasholm >= 0)
        {
          szNewFilePath.Format("%s\\%s_FBA_%s.csv", lpsPath, lpszName, szType);
          szNewFilePathMaasholm.Format("%s\\%s_Maasholm_%s.csv", lpsPath, lpszName, szType);

          szNewCsvMaasholm = artikelMaasholm.GetFirstLine();
          for (i = 0; i < pairCount; i++)
            szNewCsvMaasholm += artikelMaasholm.GetArtikelLine(arrArtPair[i].lpszBezeichnung);
        }
    }
    else
    {
        szNewFilePath.Format("%s\\%s_%s.csv", lpsPath, lpszName, szType);
    }


    length = szNewCsv.GetLength();
    CFile fl, flTroplo, flMaasholm;
    if (fl.Open(szNewFilePath, CFile::modeCreate | CFile::modeWrite))
    {
      fl.Write((LPCSTR)szNewCsv, length);
      fl.Close();
    }

    if (fMitZweiLager)
    {
        if (flTroplo.Open(szNewFilePathTroplo, CFile::modeCreate | CFile::modeWrite))
        {
            flTroplo.Write((LPCSTR)szNewCsvTroplo, szNewCsvTroplo.GetLength());
            flTroplo.Close();
        }
        
        if (indexMaasholm >= 0)
        {
          if (flMaasholm.Open(szNewFilePathMaasholm, CFile::modeCreate | CFile::modeWrite))
          {
            flMaasholm.Write((LPCSTR)szNewCsvMaasholm, szNewCsvMaasholm.GetLength());
            flMaasholm.Close();
          }
        }

    }


    CString szMsg;
    if (fVerkauf)
      szMsg.Format("Die Datei %s wurde erzeugt.\n\nInsgesamt wurden %d Artikel verkauft.", szNewFilePath, totalMenge);
    else
      szMsg.Format("Die Datei %s wurde erzeugt.\n\nInsgesamt sind %d Artikel verfügbar.", szNewFilePath, totalMenge);
    
    MessageBox(szMsg, "Info", MB_OK);

}
                             
void CJtlImportDlg::OnCbnSelchangeComboLager()
{
  m_nCurLagerSel = m_cmbLager.GetCurSel();
  switch (m_nCurLagerSel)
  {
  case CMB_LAGER_MAAS_TO_TROPLO: 
  case CMB_LAGER_TROPLO_TO_MAAS: 
    m_cmbTyp.EnableWindow(FALSE);
  break;
  default:
    m_cmbTyp.EnableWindow(TRUE);

  }
}

void CJtlImportDlg::GetLager(void)
{
    m_ExportLager = DEF_LAGER;
    switch (m_cmbLager.GetCurSel())
    {
        case CMB_LAGER_TROPLO:          m_ExportLager = LAGER_NAME_TROPLO;    break;
        case CMB_LAGER_FBA:             m_ExportLager = LAGER_NAME_FBA;       break;
        case CMB_LAGER_MAASHOLM:        m_ExportLager = LAGER_NAME_MAASHOLM;  break;
        case CMB_LAGER_MAAS_TO_TROPLO:  m_ExportLager = LAGER_NAME_MAASHOLM;  break;
        case CMB_LAGER_TROPLO_TO_MAAS:  m_ExportLager = LAGER_NAME_MAASHOLM;  break;
        case CMB_LAGER_RUEK:            m_ExportLager = LAGER_NAME_RUECK;     break;
        case CMB_LAGER_RESELLER:        m_ExportLager = LAGER_NAME_RESELLER;  break;
    }
}

#define EI_NAME_ARTIKEL _T("Artikel")

#define ARTIKEL_NAME_IMPORT_ARR                                        \
{                                                                      \
  {"JC camel","jc%2dcamel"},                                           \
  {"JC sage","jc%2dsage"},                                             \
  {"JC cafe","jc%2dcafe"},                                             \
  {"JC black","jc%2dblack"},                                           \
  {"JC red","jc%2dred"},                                               \
  {"JC pumpkin","jc%2dpumpkin"},                                       \
  {"JC brick burgundy","jc%2dburgundy"},                               \
  {"San Juan Camel","sj%2dcamel"},                                     \
  {"San Juan","sj%2dcamel"},                                           \
  {"JC camel mit Sohle","xxjc%2dcamel"},                               \
  {"JC black mit Sohle","xxjc%2dblack"},                               \
  {"Mountain Momma beige","mm-camel-%2d"},                             \
  {"Doppel Decker Platform JC camel","platform-camel-%2d"},            \
  {"Slip on camel","slip-camel-%2d"},                                  \
  {"Romano camel","rom%2dcamel"},                                      \
  {"Flower Flop","ff%2dtan-or"},                                       \
  {"Jester camel","jester%2dcamel"},                                   \
  {"Toe Joe camel","tj-camel-%2d"},                                    \
  {"Toe Joe navy", "tj-navy-%2d"},                                     \
  {"JC Disco Black","jc%2ddiscoblack"},                                \
  {"Slip on black","slip-black-%2d"},                                  \
  {"Slip on gray", "slip-gray-%2d"},                                   \
  {"Slip on denim", "slip-denim-%2d"},                                 \
  {"JC gray", "jc%2dgray"}                                             \
}


void CJtlImportDlg::DoExcelImport(LPCSTR lpszFilePath, LPCSTR lpsPath, LPCSTR lpszName)
{

    CStringArray arrFields;
    
    ARTPAIR      arrArtPair[] = ARTIKEL_NAME_IMPORT_ARR;
    CUIntArray   arrSize;
    CString      szFormat, szArtikelName, szLine, csvLine, szNewFilePath, szLagerFrom, szLagerTo;
    UINT         value;
    bool         fHasHeader;
    int          wert, faktor, length;
    int          i, indexArtikel, indexStartSize, indexBez, indArtikel, indexSize;
    int          pairCount = sizeof(arrArtPair) / sizeof(ARTPAIR);
    int          countTotal;
    
    CCSVFile     csv(lpszFilePath);

    faktor        = m_cmbTyp.GetCurSel() == 0 ? +1 : -1;
    csvLine       = LINE_IMPORT_HEADER;
    m_importLines = 0;
    fHasHeader    = false;
    indexArtikel  = -1;
    countTotal    =  0;
    
    GetLager();

    while (csv.ReadData(arrFields))
    {
        if (!fHasHeader)
        {
            for (i = 0; i < arrFields.GetCount(); i++)
            {
                if (arrFields[i] == EI_NAME_ARTIKEL)
                {
                    indexArtikel   = i;
                    indexStartSize = i + 1;
                    fHasHeader     = true;
                    continue;
                }
                if (indexArtikel >= 0)
                {
                    value = atoi(arrFields[i]);
                    arrSize.Add(value);
                }
            }
        }
        else
        {
            // Finde den Artikel ..
            if (indexArtikel < arrFields.GetCount())
            {
                indexBez = -1;
                for (indArtikel = 0; indexBez < 0 && indArtikel < pairCount; indArtikel++)
                {
                    arrFields[indexArtikel].TrimRight();
                    if (arrFields[indexArtikel] == arrArtPair[indArtikel].lpszBezeichnung)
                    {
                        szFormat = arrArtPair[indArtikel].lspzFormat;
                        indexBez = indArtikel;
                    }
                }
                for (indArtikel = indexStartSize; indexBez >= 0 && indArtikel < arrFields.GetCount(); indArtikel += 1)
                {
                    indexSize = indArtikel - indexStartSize;
                    if (indexSize < arrSize.GetSize())
                    {
                        value = arrSize[indexSize];
                        szArtikelName.Format(szFormat, value);
                        wert = atoi(arrFields[indArtikel]);
                        if (abs(wert) > 0)
                        {
                          if (m_nCurLagerSel != CMB_LAGER_MAAS_TO_TROPLO && m_nCurLagerSel != CMB_LAGER_TROPLO_TO_MAAS)
                          {
                            szLine.Format("%s;%d;%s;Y\r\n", szArtikelName, faktor * wert, m_ExportLager);
                            csvLine += szLine;
                            m_importLines += 1;
                          }
                          else
                          {
                            szLagerFrom = m_nCurLagerSel == CMB_LAGER_MAAS_TO_TROPLO ? LAGER_NAME_MAASHOLM : LAGER_NAME_TROPLO;
                            szLagerTo   = m_nCurLagerSel == CMB_LAGER_MAAS_TO_TROPLO ? LAGER_NAME_TROPLO   : LAGER_NAME_MAASHOLM;
                            
                            szLine.Format("%s;%d;%s;Y\r\n", szArtikelName, -1 * wert, szLagerFrom);
                            csvLine += szLine;
                            szLine.Format("%s;%d;%s;Y\r\n", szArtikelName, +1 * wert, szLagerTo);
                            csvLine += szLine;
                            m_importLines += 2;
                          }
                          countTotal += wert;
                        }
                    }
                }
            }
        }
    }

    if (m_importLines > 0)
    {
        if (m_nCurLagerSel != CMB_LAGER_MAAS_TO_TROPLO && m_nCurLagerSel != CMB_LAGER_TROPLO_TO_MAAS)
          szNewFilePath.Format("%s\\JTL_Import_%s_%s.csv", lpsPath, lpszName, "_Lager");
        else
          szNewFilePath.Format("%s\\JTL_Import_%s_nach_%s.csv", lpsPath, szLagerFrom, szLagerTo);

        length = csvLine.GetLength();
        CFile fl;
        if (fl.Open(szNewFilePath, CFile::modeCreate | CFile::modeWrite))
        {
            fl.Write((LPCSTR)csvLine, length);
            fl.Close();
        }

        CString szMsg;
        szMsg.Format("Es wurde die JTL-Importdatei %s\nmit %d Datensätzen erzeugt.\n\nDie Gesamtzahl der Artikel beträgt %d.", szNewFilePath, m_importLines, countTotal);
        MessageBox(szMsg, "Info", MB_OK);
    }

}

#define SET_AS_INDEX(id) int _##id = -1;
#define GET_AS_INDEX(id) _##id = mapHeaderIndex[headerNameArr[id]];

double GetPreis(LPCSTR szPreis)
{
    double pr = 0;
    if (szPreis)
    {
        CString preis(szPreis);
        if (preis.GetLength() > 0)
        {
            preis.Replace(',', '.');
            pr = atof(preis);
        }
    }
    return pr;
}

CString CJtlImportDlg::GetDate(LPCSTR lpszData)
{
    CString szDate = lpszData;
    int len = strlen(lpszData);
    if (len > 11 && lpszData[2] == '-' && lpszData[6] == '-')
    {
        int mon, t;

        char tag[3], monat[4], jahr[5];
        memcpy(tag, lpszData, 2);
        tag[2] = '\0';
        t = atoi(tag);
        memcpy(monat, lpszData + 3, 3);
        monat[3] = '\0';
        memcpy(jahr, lpszData + 7, 4);
        jahr[4] = '\0';

        if (m_monthMap.Lookup(monat, mon))
            szDate.Format("%02d.%02d.%s", t, mon, jahr);
    }
    return szDate;

}



int TaxGetDay(LPCSTR lpszDate)
{
  int days = 0;
  if (strlen(lpszDate) == 10 && lpszDate[2] == '.' && lpszDate[5] == '.')
  {
    int t, m, j;
    char tag[3], monat[4], jahr[5];
    
    memcpy(tag, lpszDate, 2);
    tag[2] = '\0';
    t = atoi(tag);
    memcpy(monat, lpszDate + 3, 3);
    monat[3] = '\0';
    m = atoi(monat);
    memcpy(jahr, lpszDate + 7, 4);
    jahr[4] = '\0';
    j = atoi(jahr);
    days = 366 * (j - 2000) + 31 * m + t;
  }
  return days;
}


int TaxCompare(CString& szFirst, CString& szNext, int idISO, int idDate)
{
  CStringArray arrFirst, arrNext;
  CCSVFile csv(';');
  int      cmp = 0;
  
  if (csv.DirectReadData(szFirst, arrFirst) && csv.DirectReadData(szNext, arrNext) && arrFirst.GetCount() == arrNext.GetCount())
  {
    cmp = arrFirst[idISO].Compare(arrNext[idISO]);
    if (0 == cmp)
    {
      int dayFirst = TaxGetDay(arrFirst[idDate]);
      int dayNext  = TaxGetDay(arrNext[idDate]);
      cmp = dayFirst - dayNext;
    }
  }
  return cmp;
}

CTaxRechnung::CTaxRechnung()
             :m_dfBetragBrutto(0)
             ,m_dfBetragNetto(0)
             ,m_dfBetragUmst(0)
{
}

void CTaxRechnung::SetBetragString(void)
{
  m_szBetragBrutto.Format("%02.2f", m_dfBetragBrutto);
  m_szBetragBrutto.Replace(".", ",");
  
  m_szBetragUmst.Format("%02.2f", m_dfBetragUmst);
  m_szBetragUmst.Replace(".", ",");

  m_szBetragNetto.Format("%02.2f", m_dfBetragNetto);
  m_szBetragNetto.Replace(".", ",");
}


void CJtlImportDlg::MapRemoveAll(CMapTaxRechnung& map)
{
  for (CMapTaxRechnung::CPair* pCurVal = map.PGetFirstAssoc(); pCurVal != NULL; pCurVal = map.PGetNextAssoc(pCurVal))
    delete pCurVal->value;
  map.RemoveAll();
}


void CJtlImportDlg::DoAmazonTaxReportOSS(LPCSTR lpszFilePath, LPCSTR lpsPath, LPCSTR lpszName)
{
  CMapTaxRechnung *pMapTaxRechnung;
  CMapTaxRechnung  mapOSSRechnung, mapOSSGutschrift;
  CMapTaxRechnung  mapEUVatRechnung, mapEUVatGutschrift;
  CTaxRechnung    *pTaxRechnung;
  CStringArray     arrFields, arrOSSLines, arrEUVatLines;
  CCSVFile         csv(lpszFilePath, ',');
  CString          szMessage, szError, szOSS, szTyp, szUmsteuer, szFaktor, szUmstID, curLine, szDatumOrder, szDatumTax,
                   szBetragPreis, szBetragUmst, szBetragNetto, szFileNameOSS, szFileNameUmst, szFile, szHeaderOSS, szHeaderEUVat;
  CFile            fl;

  double           dfBetragPreis, dfBetragUmst, dfBetragNetto, dfLandUmst, dfWaehrungFaktor;
  enum             dataType { typeReturn, typeRechnung, typeRefund, typeUndefined } type;
  bool             fHasError, fLineError, fOSS, fEUVat;
  int              i, countHeaderFields, countHeaderFldName, lineCount, index;
                   
  LPCSTR           headerNameArr[] = AS_HEADER_FLD_NAME;

  szFileNameOSS = lpsPath;
    
  szFile     = lpszName;
  szFile.Replace(".csv", "");
    
  szFileNameOSS.Format("%s\\%s_Amazon_OSS.csv", lpsPath, szFile);
  szFileNameUmst.Format("%s\\%s_Amazon_EU_VAT.csv", lpsPath, szFile);

  SET_AS_INDEX(AS_INDEX_MARKETPLACE)
  SET_AS_INDEX(AS_INDEX_DATUM)
  SET_AS_INDEX(AS_INDEX_TYP)
  SET_AS_INDEX(AS_INDEX_BESTELLNR_EXT)
  SET_AS_INDEX(AS_INDEX_SKU)
  SET_AS_INDEX(AS_INDEX_ANZAHL)
  SET_AS_INDEX(AS_TAX_CALC_DATE)
  SET_AS_INDEX(AS_INDEX_TAX_REP_SCHEME)
  SET_AS_INDEX(AS_INDEX_UMST)
  SET_AS_INDEX(AS_INDEX_WAEHRUNG)
  SET_AS_INDEX(AS_INDEX_PREIS_BRUTTO)
  SET_AS_INDEX(AS_INDEX_PREIS_UMST)
  SET_AS_INDEX(AS_INDEX_PREIS_NETTO)
  SET_AS_INDEX(AS_INDEX_VERSAND_BRUTTO)
  SET_AS_INDEX(AS_INDEX_VERSAND_UMST)
  SET_AS_INDEX(AS_INDEX_VERSAND_NETTO)
  SET_AS_INDEX(AS_INDEX_VERSAND_AKT_BRUTTO)
  SET_AS_INDEX(AS_INDEX_VERSAND_AKT_UMST)
  SET_AS_INDEX(AS_INDEX_VERSAND_AKT_NETTO)
  SET_AS_INDEX(AS_INDEX_GESCHENK_BRUTTO)
  SET_AS_INDEX(AS_INDEX_GESCHENK_UMST)
  SET_AS_INDEX(AS_INDEX_GESCHENK_NETTO)
  SET_AS_INDEX(AS_INDEX_UMST_ID_KAEUFER)
  SET_AS_INDEX(AS_INDEX_UMST_ID_LAND)
  SET_AS_INDEX(AS_INDEX_UMST_BERECHNET_CUR)
  SET_AS_INDEX(AS_INDEX_FAKTOR_WAEHRUNG)
  SET_AS_INDEX(AS_INDEX_UMST_BERECHNET)
  SET_AS_INDEX(AS_INDEX_EXPORT_NONEU)
  SET_AS_INDEX(AS_INDEX_EMPF_PLZ)
  SET_AS_INDEX(AS_INDEX_EMPF_STADT)
  SET_AS_INDEX(AS_INDEX_EMPF_ISO)

  CMap<CString, LPCSTR, int, int> mapHeaderIndex;

  countHeaderFldName = sizeof(headerNameArr) / sizeof(LPCSTR);
  lineCount  = 0;
  fLineError = false;
  szError    = "Unbekannter Fehler";
    
  szHeaderOSS   = "Bestelldatum;Rechnungsdatum;Bestellnummer;Betrag Brutto;Betrag Umst.;Betrag Netto;Umsatzsteuer;ISO Empfänger;Typ";
  szHeaderEUVat = "Bestelldatum;Rechnungsdatum;Bestellnummer;Betrag Brutto;Betrag Netto;ISO Empfänger;ISO Umst.;Umsatzsteuer-ID;Typ";


  while (!fLineError && csv.ReadData(arrFields))
  {

      fHasError = false;

      if (0 == lineCount++)
      {
          // Prüfe die erste Zeile
          countHeaderFields = arrFields.GetCount();
          for (i = 0; i < arrFields.GetCount(); i++)
              mapHeaderIndex[arrFields[i]] = i;

          for (i = 0; i < countHeaderFldName && !fHasError; i++)
              fHasError = !mapHeaderIndex.Lookup(headerNameArr[i], index);

          if (!fHasError)
          {
              GET_AS_INDEX(AS_INDEX_MARKETPLACE)
              GET_AS_INDEX(AS_INDEX_DATUM)
              GET_AS_INDEX(AS_INDEX_TYP)
              GET_AS_INDEX(AS_INDEX_BESTELLNR_EXT)
              GET_AS_INDEX(AS_INDEX_SKU)
              GET_AS_INDEX(AS_INDEX_ANZAHL)
              GET_AS_INDEX(AS_TAX_CALC_DATE)
              GET_AS_INDEX(AS_INDEX_TAX_REP_SCHEME)
              GET_AS_INDEX(AS_INDEX_UMST)
              GET_AS_INDEX(AS_INDEX_WAEHRUNG)
              GET_AS_INDEX(AS_INDEX_PREIS_BRUTTO)
              GET_AS_INDEX(AS_INDEX_PREIS_UMST)
              GET_AS_INDEX(AS_INDEX_PREIS_NETTO)
              GET_AS_INDEX(AS_INDEX_VERSAND_BRUTTO)
              GET_AS_INDEX(AS_INDEX_VERSAND_UMST)
              GET_AS_INDEX(AS_INDEX_VERSAND_NETTO)
              GET_AS_INDEX(AS_INDEX_VERSAND_AKT_BRUTTO)
              GET_AS_INDEX(AS_INDEX_VERSAND_AKT_UMST)
              GET_AS_INDEX(AS_INDEX_VERSAND_AKT_NETTO)
              GET_AS_INDEX(AS_INDEX_GESCHENK_BRUTTO)
              GET_AS_INDEX(AS_INDEX_GESCHENK_UMST)
              GET_AS_INDEX(AS_INDEX_GESCHENK_NETTO)
              GET_AS_INDEX(AS_INDEX_UMST_ID_KAEUFER)
              GET_AS_INDEX(AS_INDEX_UMST_ID_LAND)
              GET_AS_INDEX(AS_INDEX_UMST_BERECHNET_CUR)
              GET_AS_INDEX(AS_INDEX_FAKTOR_WAEHRUNG)
              GET_AS_INDEX(AS_INDEX_UMST_BERECHNET)
              GET_AS_INDEX(AS_INDEX_EXPORT_NONEU)
              GET_AS_INDEX(AS_INDEX_EMPF_PLZ)
              GET_AS_INDEX(AS_INDEX_EMPF_STADT)
              GET_AS_INDEX(AS_INDEX_EMPF_ISO)
          }
          else
              szError = "Anzahl der Feldnamen stimmt nicht mit der Vorgabe überein";

          fLineError = fHasError;
      }
      else
      {
          szOSS    = arrFields[_AS_INDEX_TAX_REP_SCHEME];
          szTyp    = arrFields[_AS_INDEX_TYP];
          szUmstID = arrFields[_AS_INDEX_UMST_ID_KAEUFER];
            
            
          type    = szTyp == "REFUND"   ? typeRefund   :            (
                    szTyp == "SHIPMENT" ? typeRechnung :             (
                    szTyp == "RETURN"   ? typeReturn : typeUndefined ));

          fOSS    = szOSS == "VCS_EU_OSS";

          szUmsteuer = arrFields[_AS_INDEX_UMST];
          dfLandUmst = 100.0 * atof(szUmsteuer);
          szUmsteuer.Format("%2.1f", dfLandUmst);

          dfWaehrungFaktor = 1;
          szFaktor         = arrFields[_AS_INDEX_FAKTOR_WAEHRUNG];
          if (szFaktor.GetLength() > 0)
          {
              double fak = atof(szFaktor);
              if (0 < fak)
                  dfWaehrungFaktor = (double)1.0 / fak;
          }

          fEUVat = (int)(100 * dfLandUmst + 0.05) == 0 && szUmstID.GetLength() > 0;

          szDatumOrder = GetDate(arrFields[_AS_INDEX_DATUM]);
          szDatumTax   = GetDate(arrFields[AS_TAX_CALC_DATE]);

          if (fOSS || fEUVat)
          {
              dfBetragPreis = GetPreis(arrFields[_AS_INDEX_PREIS_BRUTTO]) + GetPreis(arrFields[_AS_INDEX_VERSAND_BRUTTO]) + GetPreis(arrFields[_AS_INDEX_VERSAND_AKT_BRUTTO]) + GetPreis(arrFields[AS_INDEX_GESCHENK_BRUTTO]);
              dfBetragUmst  = GetPreis(arrFields[_AS_INDEX_PREIS_UMST])   + GetPreis(arrFields[_AS_INDEX_GESCHENK_UMST])  + GetPreis(arrFields[_AS_INDEX_VERSAND_UMST])       + GetPreis(arrFields[AS_INDEX_VERSAND_AKT_UMST]);
              dfBetragNetto = dfBetragPreis - dfBetragUmst;
                
              szBetragPreis.Format("%02.2f", dfBetragPreis);
              szBetragPreis.Replace(".", ",");

              szBetragUmst.Format("%02.2f", dfBetragUmst);
              szBetragUmst.Replace(".", ",");
                
              szBetragNetto.Format("%02.2f", dfBetragNetto);
              szBetragNetto.Replace(".", ",");

          }

          // Daten nur von Interesse, wenn OSS ...
          pMapTaxRechnung = NULL;
          if (fOSS)
          {
            
            if (typeRefund == type || typeReturn == type)
              pMapTaxRechnung = &mapOSSGutschrift;
            else
              pMapTaxRechnung = &mapOSSRechnung;
            
            // curLine.Format("\r\n%s;%s;%s;%s;%s;%s;%s;%s;%s", szDatumOrder, szDatumTax, arrFields[_AS_INDEX_BESTELLNR_EXT], szBetragPreis, szBetragUmst, szBetragNetto, arrFields[_AS_INDEX_UMST], arrFields[AS_INDEX_EMPF_ISO], szTyp);
            // arrOSSLines.Add(curLine);
          }
          else if (fEUVat)
          {

            if (typeRefund == type || typeReturn == type)
              pMapTaxRechnung = &mapEUVatGutschrift;
            else
              pMapTaxRechnung = &mapEUVatRechnung;

            // curLine.Format("\r\n%s;%s;%s;%s;%s;%s;%s;%s;%s", szDatumOrder, szDatumTax, arrFields[_AS_INDEX_BESTELLNR_EXT], szBetragPreis, szBetragNetto, arrFields[AS_INDEX_EMPF_ISO], arrFields[AS_INDEX_UMST_ID_LAND], arrFields[AS_INDEX_UMST_ID_KAEUFER], szTyp);
            // arrEUVatLines.Add(curLine);
          }

          if (pMapTaxRechnung)
          {
            if (!pMapTaxRechnung->Lookup(arrFields[_AS_INDEX_BESTELLNR_EXT], pTaxRechnung))
            {
              pTaxRechnung                       = new CTaxRechnung();
              
              pTaxRechnung->m_dfBetragBrutto     = dfBetragPreis;
              pTaxRechnung->m_dfBetragNetto      = dfBetragNetto;
              pTaxRechnung->m_dfBetragUmst       = dfBetragUmst;

              pTaxRechnung->m_szDatumOrder       = szDatumOrder;
              pTaxRechnung->m_szDatumTax         = szDatumTax;
              pTaxRechnung->m_szBestellnummer    = arrFields[_AS_INDEX_BESTELLNR_EXT];

              pTaxRechnung->m_szEmpfIso          = arrFields[AS_INDEX_EMPF_ISO];
              pTaxRechnung->m_szTaxIso           = arrFields[AS_INDEX_UMST_ID_LAND];
              pTaxRechnung->m_szUmstID           = arrFields[AS_INDEX_UMST_ID_KAEUFER];
              
              pTaxRechnung->m_szUmsatzSteuerSatz = arrFields[_AS_INDEX_UMST];

              pMapTaxRechnung->SetAt(arrFields[_AS_INDEX_BESTELLNR_EXT], pTaxRechnung);
              
              pTaxRechnung->m_szTyp              = szTyp;

            }
            else
            {
              pTaxRechnung->m_dfBetragBrutto    += dfBetragPreis;
              pTaxRechnung->m_dfBetragNetto     += dfBetragNetto;
              pTaxRechnung->m_dfBetragUmst      += dfBetragUmst;
            }

          }

      }

      if (fHasError)
      {
          szMessage.Format("Es trat folgender Fehler auf:\n%s\n", szError);
          MessageBox(szMessage, "Achtung", MB_OK);
      }
      else
      { 
        arrFields.RemoveAll();
      }
  }

  if (!fHasError)
  {
    arrOSSLines.RemoveAll();
    AddMapToLines(mapOSSRechnung,     arrOSSLines,   true);
    AddMapToLines(mapOSSGutschrift,   arrOSSLines,   true);
    
    arrEUVatLines.RemoveAll();
    AddMapToLines(mapEUVatRechnung,   arrEUVatLines, false);
    AddMapToLines(mapEUVatGutschrift, arrEUVatLines, false);
  }


  if (!fHasError)
  {
    CStringArray arrLine;
    CString      szFilenName, szCurLand, szLand;
    CCSVFile     csv(';');
    bool         fFileOpen;

    DoSort(arrOSSLines, TaxCompare, true);
    if (fl.Open(szFileNameOSS, CFile::modeCreate | CFile::modeWrite))
    {
        fl.Write((LPCSTR)szHeaderOSS, szHeaderOSS.GetLength());
        for (i = 0; i < arrOSSLines.GetCount(); i++)
            fl.Write((LPCSTR)arrOSSLines[i], arrOSSLines[i].GetLength());
        fl.Close();

        szMessage.Format("Die OSS-Datei\n%s\nwurde erzeugt.\n\nAnzahl der Datensätze: %d", szFileNameOSS, arrOSSLines.GetCount());
        MessageBox(szMessage, "Info", MB_OK);
    }

    DoSort(arrEUVatLines, TaxCompare, false);
    if (fl.Open(szFileNameUmst, CFile::modeCreate | CFile::modeWrite))
    {

        fl.Write((LPCSTR)szHeaderEUVat, szHeaderEUVat.GetLength());
        for (i = 0; i < arrEUVatLines.GetCount(); i++)
            fl.Write((LPCSTR)arrEUVatLines[i], arrEUVatLines[i].GetLength());
        fl.Close();

        szMessage.Format("Die EU-Vat-Datei\n%s\nwurde erzeugt.\n\nAnzahl der Datensätze: %d", szFileNameUmst, arrEUVatLines.GetCount());
        MessageBox(szMessage, "Info", MB_OK);
    }

    // Länderabhängig
    
    // OSS
    fFileOpen = false;
    szCurLand = "";
    for (i = 0; i < arrOSSLines.GetCount(); i++)
    {
      csv.DirectReadData(arrOSSLines[i], arrLine);
      szLand = arrLine[7];
      if (szLand != szCurLand)
      {
        szCurLand = szLand;
        if (fFileOpen)
          fl.Close();
        szFilenName.Format("%s\\%s_Amazon_OSS_%s.csv", lpsPath, szFile, szLand);
        fFileOpen = fl.Open(szFilenName, CFile::modeCreate | CFile::modeWrite);
        if (fFileOpen)
          fl.Write((LPCSTR)szHeaderOSS, szHeaderOSS.GetLength());
      }
      if (fFileOpen)
        fl.Write((LPCSTR)arrOSSLines[i], arrOSSLines[i].GetLength());
    }
    if (fFileOpen)
      fl.Close();


    // EU_VAT
    fFileOpen = false;
    szCurLand = "";
    for (i = 0; i < arrEUVatLines.GetCount(); i++)
    {
      csv.DirectReadData(arrEUVatLines[i], arrLine);
      szLand = arrLine[6];
      if (szLand != szCurLand)
      {
        szCurLand = szLand;
        if (fFileOpen)
          fl.Close();
        szFilenName.Format("%s\\%s_Amazon_EU_VAT_%s.csv", lpsPath, szFile, szLand);
        fFileOpen = fl.Open(szFilenName, CFile::modeCreate | CFile::modeWrite);
        if (fFileOpen)
          fl.Write((LPCSTR)szHeaderEUVat, szHeaderEUVat.GetLength());
      }
      if (fFileOpen)
        fl.Write((LPCSTR)arrEUVatLines[i], arrEUVatLines[i].GetLength());
    }
    if (fFileOpen)
      fl.Close();

  }

  MapRemoveAll(mapOSSRechnung);
  MapRemoveAll(mapOSSGutschrift);
  MapRemoveAll(mapEUVatRechnung);
  MapRemoveAll(mapEUVatGutschrift);




}


void CJtlImportDlg::AddMapToLines(CMapTaxRechnung& mapRechnung, CStringArray& arrLines, bool fOSS)
{

  CTaxRechnung *pTaxRechnung;
  CString       curLine;

  for (CMapTaxRechnung::CPair* pCurVal = mapRechnung.PGetFirstAssoc(); pCurVal != NULL; pCurVal = mapRechnung.PGetNextAssoc(pCurVal))
  {
    pTaxRechnung = pCurVal->value;
    pTaxRechnung->SetBetragString();
    if (fOSS)
      curLine.Format("\r\n%s;%s;%s;%s;%s;%s;%s;%s;%s",
        pTaxRechnung->m_szDatumOrder, pTaxRechnung->m_szDatumTax, pTaxRechnung->m_szBestellnummer, pTaxRechnung->m_szBetragBrutto, pTaxRechnung->m_szBetragUmst, pTaxRechnung->m_szBetragNetto,
        pTaxRechnung->m_szUmsatzSteuerSatz, pTaxRechnung->m_szEmpfIso, pTaxRechnung->m_szTyp);
    else
      curLine.Format("\r\n%s;%s;%s;%s;%s;%s;%s;%s;%s",
        pTaxRechnung->m_szDatumOrder, pTaxRechnung->m_szDatumTax, pTaxRechnung->m_szBestellnummer, pTaxRechnung->m_szBetragBrutto, pTaxRechnung->m_szBetragNetto,
        pTaxRechnung->m_szEmpfIso, pTaxRechnung->m_szTaxIso, pTaxRechnung->m_szUmstID, pTaxRechnung->m_szTyp);
    arrLines.Add(curLine);
  }
}



#define TC_INDEX_OSS_ISO   7
#define TC_INDEX_OSS_DATE  1

#define TC_INDEX_EUVAT_ISO   7
#define TC_INDEX_EUVAT_DATE  1

void CJtlImportDlg::DoSort(CStringArray &arr, COMPARE comp, bool fOSS)
{
  CString save;
  bool    fChange;
  int     i, nCount;

  int     indexISO  = fOSS ? TC_INDEX_OSS_ISO  : TC_INDEX_EUVAT_ISO;
  int     indexData = fOSS ? TC_INDEX_OSS_DATE : TC_INDEX_EUVAT_DATE;

  nCount = arr.GetCount();
  do
  {
    fChange = false;
    for (i = 0; i < nCount - 1; i++)
    {
      if (comp(arr[i], arr[i + 1], indexISO, indexData) > 0)
      {
        save       = arr[i];
        arr[i]     = arr[i + 1];
        arr[i + 1] = save;
        fChange    = true;
      }
    }
  } while (fChange);
}


void CJtlImportDlg::OnBnClickedOpen()
{
  
  CFileDialog dlg(TRUE);
  if (dlg.DoModal() == IDOK)
  {
    /*
    CString szMsg;
    szMsg = dlg.GetFolderPath();  // Pfad ohne \\
    szMsg += "\n";
    szMsg += dlg.GetFileName();
    szMsg += "\n";
    szMsg += dlg.GetFileTitle();  // Name ohne Endung
    szMsg += "\n";
    szMsg += dlg.GetPathName();
    szMsg += "\n";
    szMsg += dlg.GetFileExt();
    MessageBox(szMsg, "Info", MB_OK);
    */
    DoImport(dlg.GetPathName(), dlg.GetFolderPath(), dlg.GetFileTitle());
  }
}


void CJtlImportDlg::OnBnClickedOpen2()
{
    CFileDialog dlg(TRUE);
    if (dlg.DoModal() == IDOK)
        DoImportAmazonSendung(dlg.GetPathName(), dlg.GetFolderPath(), dlg.GetFileTitle());
}

void CJtlImportDlg::OnBnClickedExcelImport()
{
    CFileDialog dlg(TRUE);
    if (dlg.DoModal() == IDOK)
        DoExcelImport(dlg.GetPathName(), dlg.GetFolderPath(), dlg.GetFileTitle());
}


void CJtlImportDlg::OnBnClickedRemission()
{
    CFileDialog dlg(TRUE);
    if (dlg.DoModal() == IDOK)
        DoImportAmazonRemission(dlg.GetPathName(), dlg.GetFolderPath(), dlg.GetFileTitle());
}

void CJtlImportDlg::OnBnClickedOssTax()
{
    CFileDialog dlg(TRUE);
    if (dlg.DoModal() == IDOK)
        DoAmazonTaxReportOSS(dlg.GetPathName(), dlg.GetFolderPath(), dlg.GetFileTitle());
}



