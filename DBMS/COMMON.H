/*
//{{AFX_INCLUDES()
#include "msflexgrid.h"
//}}AFX_INCLUDES
*/
#ifndef COMMON_H
#define COMMON_H

typedef enum {DV_S,DV_F,DV_M} enMachineType;
typedef enum {TYPE_ORIGINAL=0,TYPE_VERIFIED,TYPE_MODIFIED} DATA_FILE_TYPE;   /// sch 2007.09.16
typedef enum {TYPE_DB=0,TYPE_FILE} DATA_SOURCE_TYPE;
typedef enum {LANGUAGE_LOCAL=0,LANGUAGE_KOREAN=1,LANGUAGE_ENGLISH=2,LANGUAGE_JAPANESE=3,LANGUAGE_CHINESE=4} LANGUAGE_TYPE;

#define PIECE_START_NUMBER 1
#define STRIP_START_NUMBER 1

#define UPLOAD_ERROR_NONE		0 
#define UPLOAD_ERROR_CONNECT	1
#define UPLOAD_ERROR_UPLOAD		2

#define MAX_DEFECT_TYPE         50
#define TOTAL_HOTKEY_NUM		82
#define DEFECT_TYPE				10
#define MAX_BAD_TYPE            1000
#define MAX_BADGROUP_TYPE       100


//Spread Sheet Cell Type
#define TYPE_DEFAULT	1
#define TYPE_COMBO		8
#define TYPE_CHECK		10
#define TYPE_BUTTON		7
#define TYPE_STATIC		5

#define ALIGN_LEFT		0
#define ALIGN_RIGHT		1
#define ALIGN_CENTER	2

#define TYPE_NORMAL				0
#define TYPE_READ_ONLY			1
#define TYPE_ROW_MODE			2
#define TYPE_SINGLE_SELECT		3
#define TYPE_MULTI_SELECT		4
#define TYPE_EXTENDED_SELECT	5

#define DELETE_NON					0
#define DELETE_MODEL				1
#define DELETE_MODEL_LAYER			2
#define DELETE_LOT					3
#define DELETE_LOT_LAYER			4
#define DELETE_LOT_LAYER_SERIAL		5
#define DELETE_DATE					10

CString ExtractFirstWordcom(CString *str, CString sep,BOOL bTrim=FALSE);
CString ExtractFirstWordcom(CString& str, CString sep,BOOL bTrim=FALSE);
CString ExtractLastWordcom(CString *str, char sep,BOOL bTrim=FALSE);
CString ExtractLastWordcom(CString& str, char sep,BOOL bTrim=FALSE);
void	Log(CString strMsg);
BOOL MakeDirectory(CString strPath);
BOOL DeleteDirectory(CString strDelDir);
BOOL DeleteDirectory2(CString strDelDir);
BOOL RenameDirectory(CString strFrom,CString strTo);
LONGLONG GetFolderCapacity(CString& strFolder);
BOOL CopyFolder2(CString &strFrom, CString &strTo);
BOOL CopyFolder(CString &strFrom, CString &strTo);
int ExCopyFile(CString strFrom, CString strTo);
BOOL FindFolder(CString strRootPath,CString strFolder, CString &strRetPath);
BOOL FindFile(CString strRootPath,CString strFile, CString &strRetPath);

// Unicode string 처리 함수
char* StringToChar(CString str);
void StringToChar(CString str, char* pCh);
char* StringToChar(CString str, int &nLen);


#ifdef UNICODE
    #define MessageBoxTimeout MessageBoxTimeoutW
#else
    #define MessageBoxTimeout MessageBoxTimeoutA
#endif

typedef int (__stdcall *MSGBOXAAPI)(IN HWND hWnd,IN LPCSTR lpText, IN LPCSTR lpCaption,IN UINT uType, IN WORD wLanguageId,IN DWORD dwMilliseconds);
typedef int (__stdcall *MSGBOXWAPI)(IN HWND hWnd,IN LPCWSTR lpText, IN LPCWSTR lpCaption,IN UINT uType, IN WORD wLanguageId,IN DWORD dwMilliseconds);



#endif // COMMON_H