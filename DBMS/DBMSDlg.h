
// DBMSDlg.h : 헤더 파일
//

#pragma once
#include "afxwin.h"
#include "afxdtctl.h"


// CDBMSDlg 대화 상자
class CDBMSDlg : public CDialogEx
{
// 생성입니다.
public:
	CDBMSDlg(CWnd* pParent = NULL);	// 표준 생성자입니다.

// 대화 상자 데이터입니다.
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DBMS_DIALOG };
#endif
public:
	CSpreadSheetEx m_Spread;		// CDataSource m_dataSource; 포함
	
	
	CTestQuery m_testQuery;			// CDataSource m_dataSource; 포함
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 지원입니다.
	
// 구현입니다.
protected:
	HICON m_hIcon;

	// 생성된 메시지 맵 함수
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	
	CString m_strInitPath;
	CString m_strConfigFileName;

	CString m_strUser;

	CString m_strServerIp;
	CString m_strCatalog;
	CString m_strAccessID;
	CString m_strDBPassword;

	int m_nDBType;

	void LoadServerInfo();	
	
	CDateTimeCtrl m_dtpStartDate;
	CDateTimeCtrl m_dtpEndDate;
	afx_msg void OnBnClickedBtnLoad();

	CStringArray m_strLotList;
	CStringArray m_strRegistDate;
	CDataSource m_dataSource;		// CDataSource m_dataSource; 포함

	void GetLotList();
	void InitSpread();
	void SetField();
	void SetSpread();
	CStringArray m_FieldIDList;
	afx_msg void OnBnClickedBtnDelete();
	
};
