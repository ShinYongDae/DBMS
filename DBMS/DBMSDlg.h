
// DBMSDlg.h : ��� ����
//

#pragma once
#include "afxwin.h"
#include "afxdtctl.h"


// CDBMSDlg ��ȭ ����
class CDBMSDlg : public CDialogEx
{
// �����Դϴ�.
public:
	CDBMSDlg(CWnd* pParent = NULL);	// ǥ�� �������Դϴ�.

// ��ȭ ���� �������Դϴ�.
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DBMS_DIALOG };
#endif
public:
	CSpreadSheetEx m_Spread;		// CDataSource m_dataSource; ����
	
	
	CTestQuery m_testQuery;			// CDataSource m_dataSource; ����
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV �����Դϴ�.
	
// �����Դϴ�.
protected:
	HICON m_hIcon;

	// ������ �޽��� �� �Լ�
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
	CDataSource m_dataSource;		// CDataSource m_dataSource; ����

	void GetLotList();
	void InitSpread();
	void SetField();
	void SetSpread();
	CStringArray m_FieldIDList;
	afx_msg void OnBnClickedBtnDelete();
	
};
