// BgWoods.h : main header file for the BGWOODS application
//

#if !defined(AFX_BGWOODS_H__14BFA03C_932E_459B_88D0_1C9AE6DE7E21__INCLUDED_)
#define AFX_BGWOODS_H__14BFA03C_932E_459B_88D0_1C9AE6DE7E21__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CBgWoodsApp:
// See BgWoods.cpp for the implementation of this class
//

class CBgWoodsApp : public CWinApp
{
public:
	CBgWoodsApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CBgWoodsApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CBgWoodsApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_BGWOODS_H__14BFA03C_932E_459B_88D0_1C9AE6DE7E21__INCLUDED_)
