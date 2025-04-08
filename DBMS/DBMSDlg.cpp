
// DBMSDlg.cpp : 구현 파일
//

#include "stdafx.h"
#include "DBMS.h"
#include "DBMSDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 응용 프로그램 정보에 사용되는 CAboutDlg 대화 상자입니다.

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 대화 상자 데이터입니다.
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 지원입니다.

// 구현입니다.
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);

}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CDBMSDlg 대화 상자



CDBMSDlg::CDBMSDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_DBMS_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CDBMSDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_DTP_START, m_dtpStartDate);
	DDX_Control(pDX, IDC_DTP_END, m_dtpEndDate);
}

BEGIN_MESSAGE_MAP(CDBMSDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BTN_LOAD, &CDBMSDlg::OnBnClickedBtnLoad)
	ON_BN_CLICKED(IDC_BTN_DELETE, &CDBMSDlg::OnBnClickedBtnDelete)
END_MESSAGE_MAP()


// CDBMSDlg 메시지 처리기

BOOL CDBMSDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 시스템 메뉴에 "정보..." 메뉴 항목을 추가합니다.

	// IDM_ABOUTBOX는 시스템 명령 범위에 있어야 합니다.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 이 대화 상자의 아이콘을 설정합니다.  응용 프로그램의 주 창이 대화 상자가 아닐 경우에는
	//  프레임워크가 이 작업을 자동으로 수행합니다.
	SetIcon(m_hIcon, TRUE);			// 큰 아이콘을 설정합니다.
	SetIcon(m_hIcon, FALSE);		// 작은 아이콘을 설정합니다.

	// TODO: 여기에 추가 초기화 작업을 추가합니다.

	m_Spread.SetParentWnd(this, IDC_FPSPREAD_DISPLAY_TABLE);
	m_Spread.Attach(ConvertTSpread(IDC_FPSPREAD_DISPLAY_TABLE));
	m_Spread.SetAppearanceStyle(0);

	m_strInitPath += _T("\\SysData\\");
	m_strConfigFileName = _T("Config.ini");

	m_nDBType = 1;

	LoadServerInfo();	
	m_testQuery.InitDB(m_strServerIp, m_strCatalog, m_strAccessID, m_strDBPassword, m_nDBType);

	COleDateTime dtStartDate, dtEndDate;
	dtEndDate = COleDateTime::GetCurrentTime() - COleDateTimeSpan(365, 0, 0, 0); // 현재 시간으로 맞추어 초기화 한다.
	m_dtpEndDate.SetTime(dtEndDate);

	dtStartDate = dtEndDate - COleDateTimeSpan(30, 0, 0, 0);
	m_dtpStartDate.SetTime(dtStartDate);


	return TRUE;  // 포커스를 컨트롤에 설정하지 않으면 TRUE를 반환합니다.
}

void CDBMSDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 대화 상자에 최소화 단추를 추가할 경우 아이콘을 그리려면
//  아래 코드가 필요합니다.  문서/뷰 모델을 사용하는 MFC 응용 프로그램의 경우에는
//  프레임워크에서 이 작업을 자동으로 수행합니다.

void CDBMSDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 그리기를 위한 디바이스 컨텍스트입니다.

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 클라이언트 사각형에서 아이콘을 가운데에 맞춥니다.
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 아이콘을 그립니다.
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// 사용자가 최소화된 창을 끄는 동안에 커서가 표시되도록 시스템에서
//  이 함수를 호출합니다.
HCURSOR CDBMSDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CDBMSDlg::LoadServerInfo()
{
	TCHAR szFilePath[MAX_PATH];
	CString strFilePath, strConfigFileName, strConfigFilePath;
	TCHAR szData[MAX_PATH];

	if (!GetModuleFileName(AfxGetApp()->m_hInstance, szFilePath, MAX_PATH))
	{
		CString strMsg;
		strMsg.Format(_T("Error code: %d\n"), GetLastError());
		AfxMessageBox(strMsg, MB_ICONSTOP);
		return;
	}
	else
	{
		strFilePath = szFilePath;
		ExtractLastWordcom(&strFilePath, '\\');
		ExtractLastWordcom(&strFilePath, '\\');
		ExtractLastWordcom(&strFilePath, '\\');
	}
	strConfigFileName = _T("Config.ini");

	strConfigFilePath = strFilePath + _T("\\SysData\\") + strConfigFileName;

	// ;0 mssql, 1mysql
	if (0 < ::GetPrivateProfileString(_T("SERVER INFORM"), _T("DBType"), NULL, szData, sizeof(szData), strConfigFilePath))
		m_nDBType = _ttoi(szData);
	else
		m_nDBType = 0;

	if (0 < ::GetPrivateProfileString(_T("SERVER INFORM"), _T("Server IP"), NULL, szData, sizeof(szData), strConfigFilePath))
		m_strServerIp = szData;
	else
		m_strServerIp = _T("100.100.101.10");

	if (0 < ::GetPrivateProfileString(_T("SERVER INFORM"), _T("DBNAME"), NULL, szData, sizeof(szData), strConfigFilePath))
		m_strCatalog = szData;
	else
		m_strCatalog = _T("GVISDB");
	

	if (0 < ::GetPrivateProfileString(_T("SERVER INFORM"), _T("USER"), NULL, szData, sizeof(szData), strConfigFilePath))
		m_strAccessID = szData;
	else
		m_strAccessID = _T("root");

	if (0 < ::GetPrivateProfileString(_T("SERVER INFORM"), _T("PASSWORD"), NULL, szData, sizeof(szData), strConfigFilePath))
		m_strDBPassword = szData;
	else
		m_strDBPassword = _T("My3309");

	InitSpread();
	
}


void CDBMSDlg::OnBnClickedBtnLoad()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	m_Spread.AutoAttach();

	GetLotList();
	
	SetSpread();
}

void CDBMSDlg::GetLotList()
{
	m_strLotList.RemoveAll();
	m_strRegistDate.RemoveAll();
	
	COleDateTime dtStartDate, dtEndDate;
	m_dtpStartDate.GetTime(dtStartDate);
	m_dtpEndDate.GetTime(dtEndDate);
	CString strStartDate = dtStartDate.Format(_T("%Y-%m-%d 00:00:00"));
	CString strEndDate = dtEndDate.Format(_T("%Y-%m-%d 23:59:59"));

	CString strQuery;
	strQuery.Format(_T("SELECT LOT_CODE, DT_REGIST FROM LOT ")
		_T("WHERE DT_REGIST BETWEEN '%s' AND '%s' ORDER BY DT_REGIST"), strStartDate, strEndDate);

	m_testQuery.GetList(strQuery,m_strLotList, m_strRegistDate);
}

void CDBMSDlg::InitSpread()
{	
	CDC* pDC = NULL;
	int nLogX = 0;
	int nTwipPerPixel = 0;
	CRect rect;
	long twpWidth = 0;
	CString strTemp;
	SetField();
	int nMaxCol = m_FieldIDList.GetSize();

	m_Spread.SetReDraw(FALSE);
	m_Spread.Reset();
	m_Spread.ClearRange(-1, -1, -1, -1, FALSE);
	m_Spread.SetMaxCols(nMaxCol);

	m_Spread.SetCol(-1);
	m_Spread.SetRow(-1);
	m_Spread.SetUserColAction(1);
	m_Spread.SetLock(TRUE);

	m_Spread.SetColHeadersShow(0);

	m_Spread.SetShadowColor(RGB(226, 233, 251), RGB(0, 0, 0), COLOR_3DSHADOW, COLOR_3DHILIGHT);

	SS_CELLTYPE stype;
	m_Spread.SetTypeEdit(&stype, SSS_ALIGN_VCENTER | SSS_ALIGN_CENTER, 128, SS_CHRSET_CHR, SS_CASE_NOCASE);
	m_Spread.SetCellType(SS_ALLCOLS, SS_ALLROWS, &stype);

	m_Spread.SetBorderRange(1, 1, nMaxCol, 1, SS_BORDERTYPE_LEFT, SS_BORDERSTYLE_SOLID, RGB_BLACK);
	m_Spread.SetBorderRange(1, 1, nMaxCol, 1, SS_BORDERTYPE_RIGHT, SS_BORDERSTYLE_SOLID, RGB_BLACK);
	m_Spread.SetBorderRange(1, 1, nMaxCol, 1, SS_BORDERTYPE_BOTTOM, SS_BORDERSTYLE_SOLID, RGB_BLACK);
	m_Spread.SetBorderRange(1, 1, nMaxCol, 1, SS_BORDERTYPE_TOP, SS_BORDERSTYLE_SOLID, RGB_BLACK);

	HFONT hfont = CreateFont(15, 0, 0, 0, 700, 0, 0, 0, 0, 0, 0, 0, 0, _T("Times New Roman"));
	m_Spread.SetFont(SS_ALLCOLS, SS_ALLROWS, hfont, TRUE);

	m_Spread.SetColorRange(1, 1, nMaxCol, 1, RGB(226, 233, 251), RGB(0, 0, 0));

	m_Spread.SetColWidth(1, 50);
	m_Spread.SetColWidth(2, 46);
	

	m_Spread.SetRowHeight(-1, 20);
	m_Spread.SetRowHeight(0, 25);

	//Title
	m_Spread.SetRow(1);
	for (int i = 1; i <= nMaxCol; i++)
	{
		m_Spread.SetCol(i);
		strTemp = m_FieldIDList.GetAt(i - 1);
		m_Spread.SetText(strTemp);
	}

	m_Spread.SetReDraw(TRUE);
}

void CDBMSDlg::SetField()
{
	m_FieldIDList.RemoveAll();

	m_FieldIDList.Add(_T("Lot List"));
	m_FieldIDList.Add(_T("Regist Date"));
}

void CDBMSDlg::SetSpread()
{
	m_Spread.AutoAttach();

	SS_CELLTYPE stype;
	m_Spread.SetTypeEdit(&stype, SSS_ALIGN_VCENTER | SSS_ALIGN_CENTER, 128, SS_CHRSET_CHR, SS_CASE_NOCASE);

	HFONT hfont = CreateFont(15, 0, 0, 0, 700, 0, 0, 0, 0, 0, 0, 0, 0, _T("Times New Roman"));
	m_Spread.SetFont(SS_ALLCOLS, SS_ALLROWS, hfont, TRUE);

	int nSize = m_strLotList.GetSize();
	CString strData;
	m_Spread.SetMaxRows(nSize + 1);
	for (int i = 0; i < nSize; i++)
	{
		m_Spread.SetCol(1);
		m_Spread.SetRow(i + 2);
		strData = m_strLotList.GetAt(i);
		m_Spread.SetText(strData);

		m_Spread.SetCol(2);
		m_Spread.SetRow(i + 2);
		strData = m_strRegistDate.GetAt(i);
		m_Spread.SetText(strData);
	}
}




void CDBMSDlg::OnBnClickedBtnDelete()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	CString strQuery, strLotName;
	int nSize = m_strLotList.GetSize();
	if (nSize > 0)
	{
		for (int i = 0; i < nSize; i++)
		{
			strLotName = m_strLotList.GetAt(i);
			strQuery.Format(_T("DELETE FROM LOT WHERE LOT_CODE = '%s'"), strLotName);
			m_testQuery.RIGHT(strQuery);
		}
		
	}

	
}
