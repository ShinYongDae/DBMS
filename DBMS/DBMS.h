
// DBMS.h : PROJECT_NAME ���� ���α׷��� ���� �� ��� �����Դϴ�.
//

#pragma once

#ifndef __AFXWIN_H__
	#error "PCH�� ���� �� ������ �����ϱ� ���� 'stdafx.h'�� �����մϴ�."
#endif

#include "resource.h"		// �� ��ȣ�Դϴ�.


// CDBMSApp:
// �� Ŭ������ ������ ���ؼ��� DBMS.cpp�� �����Ͻʽÿ�.
//

class CDBMSApp : public CWinApp
{
public:
	CDBMSApp();

// �������Դϴ�.
public:
	virtual BOOL InitInstance();

// �����Դϴ�.

	DECLARE_MESSAGE_MAP()
};

extern CDBMSApp theApp;