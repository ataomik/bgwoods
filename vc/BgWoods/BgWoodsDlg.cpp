// BgWoodsDlg.cpp : implementation file
//

#include "stdafx.h"
#include "BgWoods.h"
#include "BgWoodsDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define WM_SHLICONNOTIFY  WM_USER + 8012

#define MAX_POINTS	108
#define MAX_FATAL	100
#define CLR_GRAY	0x808080
#define CLR_WHITE	0xffffff
#define CLR_BLACK	0x000000

#define GAMES_COUNT	6

static UINT uGMIconIds[] =
{
	IDI_ICON_BG1,
	IDI_ICON_TSC,
	IDI_ICON_SOA,
	IDI_ICON_TOB,
	IDI_ICON_IWD,
	IDI_ICON_HOW,
};

static CString strGMNames[] = 
{
	"Baldur's Gate I",
	"Baldur's Gate I: Tales of the Sword Coast",
	"Baldur's Gate II: Shadow of Amn",
	"Baldur's Gate II: Throne of Bhaal",
	"Icewind Dale",
	"IceWind Dale: How of Winter",
};

static CPoint ptIcoMain = CPoint(10, 10);
static CRect rctIcoMain = CRect(10, 10, 43, 43);

CBgData __declspec(dllimport) __stdcall SetData(const CBgData* pBgData);
CChrAttrs __declspec(dllimport) __stdcall Search(const CChrAttrs* pAttrs);
BOOL __declspec(dllimport) __stdcall EnableHotKeyHook(void);
BOOL __declspec(dllimport) __stdcall DisableHotKeyHook(void);
/////////////////////////////////////////////////////////////////////////////
// CBgWoodsDlg dialog

CBgWoodsDlg::CBgWoodsDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CBgWoodsDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CBgWoodsDlg)
	m_strTimes = _T("1");
	m_strCurPoints = _T("0");
	m_strPoints = _T("0");
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_bgData.Address = 0;
	m_bgData.Delay = 0;
	m_bgData.CriPoints = 0;
	m_bgData.CriStop = 0;
	m_bgData.GameType = SOA;
	m_bgData.Points = 0;
}

void CBgWoodsDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CBgWoodsDlg)
	DDX_Control(pDX, IDC_STATIC_FRAME, m_stcFrame);
	DDX_Control(pDX, IDC_LIST_GT, m_lstGt);
	DDX_CBString(pDX, IDC_COMBO_TIMES, m_strTimes);
	DDV_MaxChars(pDX, m_strTimes, 3);
	DDX_Text(pDX, IDC_EDIT_CUR_POINTS, m_strCurPoints);
	DDV_MaxChars(pDX, m_strCurPoints, 22);
	DDX_Text(pDX, IDC_EDIT_POINTS, m_strPoints);
	DDV_MaxChars(pDX, m_strPoints, 3);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CBgWoodsDlg, CDialog)
	//{{AFX_MSG_MAP(CBgWoodsDlg)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_WM_CTLCOLOR()
	ON_BN_CLICKED(IDC_BITMAP_HELP, OnBitmapHelp)
	ON_EN_CHANGE(IDC_EDIT_POINTS, OnChangeEditPoints)
	ON_NOTIFY(LVN_ITEMCHANGED, IDC_LIST_GT, OnItemchangedListGt)
	ON_BN_CLICKED(IDSEARCH, OnSearch)
	ON_MESSAGE(WM_SHLICONNOTIFY, OnShlIconNotify)
	ON_EN_CHANGE(IDC_EDIT_CUR_POINTS, OnChangeEditCurPoints)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBgWoodsDlg message handlers

BOOL CBgWoodsDlg::OnInitDialog()
{
	CRect rect;
	int i, stbWidths[] = {120, 600};
	
	CDialog::OnInitDialog();

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	// TODO: Add extra initialization here
	((CComboBox* )GetDlgItem(IDC_COMBO_TIMES))->SetCurSel(0);
	GetClientRect(&rect);
	rect.top = rect.bottom - 10;
	if(!m_wndStatus.Create(CBRS_BOTTOM, rect, this, ID_WND_STATUSBAR) ||
		!m_wndStatus.SetParts(sizeof(stbWidths) / sizeof(int), stbWidths))
	{
		TRACE0("Can't create statusbar.\n");
		return FALSE;
	}
	m_wndStatus.ShowWindow(SW_SHOW);
	
	if(!m_imglstMain.Create(32, 32, TRUE | ILC_COLOR8, 6, 6))
	{
		TRACE0("Can't create imagelist.\n");
		return FALSE;
	}
	m_imglstMain.SetBkColor(CLR_NONE);
	for(i = 0; i < sizeof(uGMIconIds) / sizeof(UINT); i ++)
		m_imglstMain.Add(AfxGetApp()->LoadIcon(uGMIconIds[i]));
	m_lstGt.SetBkColor(CLR_GRAY);
	m_lstGt.SetTextBkColor(CLR_GRAY);
	m_lstGt.SetTextColor(CLR_WHITE);
	m_lstGt.SetImageList(&m_imglstMain, LVSIL_NORMAL);
	for(i = 0; i < GAMES_COUNT; i ++)
		m_lstGt.InsertItem(i, strGMNames[i], i);
	
	m_bIsVisible = TRUE;
	return TRUE;  // return TRUE  unless you set the focus to a control
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CBgWoodsDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting
		
		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);
		
		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CBgWoodsDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

HBRUSH CBgWoodsDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor) 
{
	HBRUSH hbr = CDialog::OnCtlColor(pDC, pWnd, nCtlColor);
	CRect rect;

	switch(pWnd->GetDlgCtrlID())
	{
		case IDC_STATIC_FRAME:
			m_stcFrame.GetClientRect(&rect);
			pDC->FillSolidRect(rect, CLR_WHITE);
			m_imglstMain.Draw(m_stcFrame.GetDC(), m_bgData.GameType, ptIcoMain, ILD_NORMAL);
			break;
	}
	// TODO: Change any attributes of the DC here
	// TODO: Return a different brush if the default is not desired
	return hbr;
}

void CBgWoodsDlg::OnBitmapHelp() 
{
	// TODO: Add your control notification handler code here
	CFileFind Finder;

	if(Finder.FindFile("readme.txt", 0))
		ShellExecute(m_hWnd, "Open", "readme.txt", NULL, NULL, SW_NORMAL);
	Finder.Close();
}

void CBgWoodsDlg::OnChangeEditPoints() 
{
	// TODO: If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialog::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.
	
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	if(atoi(m_strPoints) > MAX_POINTS)
	{
		m_strPoints.Format("%d", MAX_POINTS);
		GetDlgItem(IDC_EDIT_POINTS)->SetWindowText(m_strPoints);
	}
}

void CBgWoodsDlg::OnItemchangedListGt(NMHDR* pNMHDR, LRESULT* pResult) 
{
	NM_LISTVIEW* pNMListView = (NM_LISTVIEW*)pNMHDR;
	// TODO: Add your control notification handler code here
	if(pNMListView->hdr.hwndFrom == m_lstGt.m_hWnd &&
		pNMListView->uNewState == (LVIS_SELECTED | LVIS_FOCUSED))
	{
		m_wndStatus.SetText("Current Game:", 0, 0);
		m_wndStatus.SetText(m_lstGt.GetItemText(pNMListView->iItem, pNMListView->iSubItem), 1, 0);
		m_stcFrame.GetDC()->FillSolidRect(rctIcoMain, CLR_WHITE);
		m_imglstMain.Draw(m_stcFrame.GetDC(), pNMListView->iItem, ptIcoMain, ILD_NORMAL);
		m_bgData.GameType = pNMListView->iItem;
	}

	*pResult = 0;
}

void CBgWoodsDlg::OnSearch() 
{
	// TODO: Add your control notification handler code here
	UINT i = 0;
	CString strAttrs;
	CChrAttrs Attrs;
	BYTE szKeys[64];

	UpdateData(TRUE);
	m_bgData.Delay = 1000 / atoi(m_strTimes);
	m_bgData.Points = atoi(m_strPoints);
	m_bgData.CriPoints = 0;
	m_bgData.CriStop = 0;
	SetData(&m_bgData);
	Attrs.Str = 0;
	if(m_strCurPoints.GetLength() > 1)
	{
		CString sToken=_T("");
		while (i < sizeof(szKeys)/sizeof(szKeys[0]) && AfxExtractSubString(sToken, m_strCurPoints, i,','))
		{
			szKeys[i] = atoi(sToken);
			i++;
		}
		if(i == 7)
		{
			Attrs.Str = szKeys[0];
			Attrs.Cri = szKeys[1];
			Attrs.Dex = szKeys[2];
			Attrs.Con = szKeys[3];
			Attrs.Int = szKeys[4];
			Attrs.Wis = szKeys[5];
			Attrs.Cha = szKeys[6];
		}
	}
	m_chrAttrs = Search(&Attrs);
	m_wndStatus.SetText("Search Result:", 0, 0);
	strAttrs.Format("%sStr: %d, Dex: %d, Con: %d, Int: %d, Wis: %d, Cha: %d", 
		Attrs.Str ? "* " : "", m_chrAttrs.Str, m_chrAttrs.Dex, m_chrAttrs.Con, m_chrAttrs.Int, m_chrAttrs.Wis, m_chrAttrs.Cha);
	m_wndStatus.SetText(strAttrs, 1, 0);
}

void CBgWoodsDlg::OnOK() 
{
	// TODO: Add extra validation here
	if(m_bIsVisible)
	{
		m_NID.cbSize = sizeof(m_NID);
		m_NID.hIcon = m_imglstMain.ExtractIcon(m_bgData.GameType);
		m_NID.hWnd = GetSafeHwnd();
		strcpy(m_NID.szTip, "Baldur's Woods...");
		m_NID.uCallbackMessage = WM_SHLICONNOTIFY;
		m_NID.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP; 
		m_NID.uID = ID_TASKICON;
		if(Shell_NotifyIcon(NIM_ADD, &m_NID))
		{
			UpdateData(TRUE);
			m_bgData.Delay = 1000 / atoi(m_strTimes);
			m_bgData.Points = atoi(m_strPoints);
			m_bgData.CriPoints = 0;
			m_bgData.CriStop = 0;
			if(EnableHotKeyHook())
			SetData(&m_bgData);
			ShowWindow(SW_HIDE);
			m_bIsVisible = FALSE;
		}
	}
	
}


void CBgWoodsDlg::OnShlIconNotify(WPARAM wParam, LPARAM lParam)
{
	if((UINT)lParam == WM_LBUTTONDOWN)
	{ 
		if(!m_bIsVisible)
		{
			ShowWindow(SW_NORMAL);
			m_bIsVisible = TRUE;
			Shell_NotifyIcon(NIM_DELETE, &m_NID);
			DisableHotKeyHook();
		}
	}
}

void CBgWoodsDlg::OnChangeEditCurPoints() 
{
	// TODO: If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialog::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.
	
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
}
