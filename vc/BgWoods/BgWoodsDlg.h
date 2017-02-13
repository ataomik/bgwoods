// BgWoodsDlg.h : header file
//

#if !defined(AFX_BGWOODSDLG_H__D8C149B3_6793_4D48_813B_741C4F210B98__INCLUDED_)
#define AFX_BGWOODSDLG_H__D8C149B3_6793_4D48_813B_741C4F210B98__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


/////////////////////////////////////////////////////////////////////////////
// BgWoods

typedef enum 
{ 
	BG1 = 0, 
	TSC = 1, 
	SOA = 2,
	TOB = 3, 
	IWD = 4, 
	HOW = 5, 
}GameType;

typedef enum 
{ 
	Str = 0, 
	Cri = 1, 
	Int = 2, 
	Wis = 3, 
	Dex = 4, 
	Con = 5, 
	Cha = 6, 
}AttrType; 

typedef struct
{
	BYTE Str;
	BYTE Dex;
	BYTE Con;
	BYTE Int;
	BYTE Wis;
	BYTE Cha;
	BYTE Cri;
}CChrAttrs;

typedef struct
{
  BYTE GameType;
  BYTE Points;
  BYTE CriPoints;
  BYTE CriStop;
  LONG Delay; 
  LONG Address;
}CBgData;

/////////////////////////////////////////////////////////////////////////////
// CBgWoodsDlg dialog


class CBgWoodsDlg : public CDialog
{
// Construction
public:
	CBgWoodsDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	//{{AFX_DATA(CBgWoodsDlg)
	enum { IDD = IDD_BGWOODS_DIALOG };
	CStatic	m_stcFrame;
	CListCtrl	m_lstGt;
	CString	m_strTimes;
	CString	m_strCurPoints;
	CString	m_strPoints;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CBgWoodsDlg)
	public:
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;
	CStatusBarCtrl m_wndStatus;
	CImageList m_imglstMain;
	CBgData m_bgData;
	CChrAttrs m_chrAttrs;
	NOTIFYICONDATA	m_NID;
	BOOL m_bIsVisible;
	// Generated message map functions
	//{{AFX_MSG(CBgWoodsDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg void OnBitmapHelp();
	afx_msg void OnChangeEditPoints();
	afx_msg void OnItemchangedListGt(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnSearch();
	virtual void OnOK();
	afx_msg void OnShlIconNotify(WPARAM wParam, LPARAM lParam); 
	afx_msg void OnChangeEditCurPoints();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_BGWOODSDLG_H__D8C149B3_6793_4D48_813B_741C4F210B98__INCLUDED_)
