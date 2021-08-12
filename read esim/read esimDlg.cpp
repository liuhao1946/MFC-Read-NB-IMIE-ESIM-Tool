
// read esimDlg.cpp: 实现文件
//

#include "pch.h"
#include "framework.h"
#include "read esim.h"
#include "read esimDlg.h"
#include "afxdialogex.h"
#include "app_common.h"
#include "app_queue.h"
#include "winspool.h"

#include <windows.h>

#pragma comment(lib,"Winmm.lib")

#include <MMSystem.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#define RX_BUF_SIZE 1024

char excel_path[MAX_PATH];
CString excPathStr;
CString FileName;

char rx_buf[RX_BUF_SIZE+1];
uint32_t rx_timeout;
uint32_t rx_len;

uint8_t state_color;
uint8_t excel_cre_f;

char ble_mac[50];
char nb_esim[100];
char nb_imei[100];

WORD duser;

static uint32_t lst_idx;
static uint32_t excel_idx;
uint8_t read_step;
uint8_t read_syn_f;

#define SF_TIMER            1
#define SER_TIMER_NUM       1
#define EXCEL_CRE_TIMER_NUM 2

//队列定义,没有用到
#define Q_MAX 10

typedef struct
{
	char idx[10];
	char mac[30];
	char time[50];
	char nb_esim[50];
	char nb_imei[50];
}nb_dt_t;

typedef struct
{
	pos_t pos;
	nb_dt_t nb_dt[Q_MAX];
}d_q_t;

d_q_t d_q;

// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnCbnSelchangeCombo1();
//	afx_msg void OnTimer(UINT_PTR nIDEvent);
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
//	ON_CBN_SELCHANGE(IDC_COMBO1, &CAboutDlg::OnCbnSelchangeCombo1)
//ON_WM_TIMER()
END_MESSAGE_MAP()


// CreadesimDlg 对话框
CreadesimDlg::CreadesimDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_READESIM_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDI_ICON2);
}

void CreadesimDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EDIT1, db);
	DDX_Control(pDX, IDC_COMBO1, box);
	DDX_Control(pDX, IDC_COMBO2, box_baud);
	DDX_Control(pDX, IDC_STATIC11, com_st);
	DDX_Control(pDX, IDC_BUTTON2, but1);
	//  DDX_Control(pDX, IDC_LIST1, lst);
	DDX_Control(pDX, IDC_LIST4, mList);
	DDX_Control(pDX, IDC_STATIC6, mState);
	DDX_Control(pDX, IDC_COMBO3, mPrintList);
}

BEGIN_MESSAGE_MAP(CreadesimDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//ON_MESSAGE(WM_COMM_RXSTR, &CreadesimDlg::OnReceiveStr)
	ON_BN_CLICKED(IDC_BUTTON1, &CreadesimDlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, &CreadesimDlg::OnBnClickedButton2)
	ON_CBN_SELCHANGE(IDC_COMBO1, &CreadesimDlg::OnCbnSelchangeCombo1)
	ON_WM_CTLCOLOR()
	ON_WM_TIMER()
	ON_BN_CLICKED(IDC_BUTTON5, &CreadesimDlg::OnBnClickedButton5)
	ON_WM_CLOSE()
END_MESSAGE_MAP()


LRESULT CreadesimDlg::OnReceiveStr(WPARAM str, LPARAM commInfo)
{
	//struct serialPortInfo
	//{
	//	UINT portNr;//串口号
	//	DWORD bytesRead;//读取的字节数
	//}*pCommInfo;
	//pCommInfo = (serialPortInfo*)commInfo;

	//DWORD len;

	//len = pCommInfo->bytesRead;
	//if ((rx_len + len) > RX_BUF_SIZE)
	//{
	//	len = RX_BUF_SIZE - rx_len;
	//}
	//memcpy(rx_buf + rx_len, (char*)str, len);
	//rx_len += len;
	//rx_buf[rx_len] = 0;
	//rx_timeout = 3;

	return TRUE;
}

void CALLBACK My_MMTimerProc(UINT uID, UINT uMsg, DWORD dwUsers, DWORD dw1, DWORD dw2)
{
#ifndef SF_TIMER
	CreadesimApp* pApp = (CreadesimApp*)AfxGetApp();
	CreadesimDlg* pDlg = (CreadesimDlg*)pApp->m_pMainWnd;

	if (rx_timeout)
	{
		rx_timeout--;
		if (rx_timeout == 0)
		{
			char *p,*q;
			CString str;
			uint16_t len;

			if ((p = strstr(rx_buf, "MAC:")) && (q = strstr(p, "\r\n")))
			{
				read_suc_f = 0;
				read_step++;

				p = p + 4;
				len = q - p;

				if (len < 30)
				{
					memset(ble_mac, 0, len);
					memcpy(ble_mac, p, len);
					str = ble_mac;

					pDlg->db_str(_T("MAC:"));
					pDlg->db_str(str + _T("\r\n"));
				}
				else
				{
					pDlg->db_str(_T("MAC length error!\r\n"));
				}
			}
			else if ((p = strstr(rx_buf, "+QCCID:")) && (q = strstr(p, "\r\n")))
			{
				read_step++;
				p = p + 8;
				len = q - p;

				if (len < 50)
				{
					memset(nb_esim, 0, len);
					memcpy(nb_esim, p, len);

					str = nb_esim;
					pDlg->db_str(_T("esim:"));
					pDlg->db_str(str + _T("\r\n"));
				}
				else
				{
					pDlg->db_str(_T("esim length error!\r\n"));
				}
			}
			else if ((p = strstr(rx_buf, "+CGSN:")) && (q = strstr(p, "\r\n")))
			{
				read_step++;
				p = p + 7;
				len = q - p;

				if (len < 50)
				{
					memset(nb_imei, 0, len);
					memcpy(nb_imei, p, len);

					str = nb_imei;
					pDlg->db_str(_T("imei:"));
					pDlg->db_str(str + _T("\r\n"));
				}
				else
				{
					pDlg->db_str(_T("imei length error!\r\n"));
				}
			}
			else if (p = strstr(rx_buf, "+CEREG"))
			{
				//读取结束
				if (read_step == 3)
				{
					nb_dt_t dt;
					char temp[20];

					CString str_temp, time_str;
					CTime time = CTime::GetCurrentTime();
					time_str = time.Format("20%y-%m-%d-%H:%M:%S");

					read_suc_f = 1;

					pDlg->mList.InsertItem(lst_idx, _T(""));
					str_temp.Format(_T("%d"), lst_idx + 1);
					//序号
					pDlg->mList.SetItemText(lst_idx, 0, str_temp);
					//时间
					pDlg->mList.SetItemText(lst_idx,1, time_str);
					//MAC
					str_temp = ble_mac;
					pDlg->mList.SetItemText(lst_idx, 2, str_temp);
					//ESIM
					str_temp = nb_esim;
					pDlg->mList.SetItemText(lst_idx, 3, str_temp);
					//IMIE
					str_temp = nb_imei;
					pDlg->mList.SetItemText(lst_idx, 4, str_temp);
					//显示列表框中最新插入的行
					int nCount = mList.GetItemCount();
					if (nCount > 0)
					   mList.EnsureVisible(nCount - 1, FALSE);
					lst_idx++;

					app_get_cstring_unit(time_str, dt.time, 50);
					memset(temp, 0, 20);
					sprintf_s(temp, "%d", lst_idx);
					strcpy_s(dt.idx, temp);
					strcpy_s(dt.mac, ble_mac);

					strcat_s(nb_esim, "\n");
					strcpy_s(dt.nb_esim, nb_esim);
					
					strcat_s(nb_imei, "\n");
					strcpy_s(dt.nb_imei, nb_imei);

					app_enqueue(&d_q,&dt);

					pDlg->SetTimer(1, 100, 0);
					pDlg->db_str(_T("读取成功，准备写入excel！\r\n"));
				}
				else
				{
					pDlg->db_str(_T("读取错误！\r\n"));
					pDlg->db_print("read_step:%d\n", read_step);
				}
				read_step = 0;
			}

			rx_len = 0;
		}
	}
#endif
}

//CreadesimDlg 消息处理程序

BOOL CreadesimDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
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

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
	mStateFt.CreatePointFont(200, _T("宋体"));
	mState.SetFont(&mStateFt);

	//创建文件名字，文件创建放在定时器中
	state_color = 2;
	mState.SetWindowTextW(_T("正在创建excel文件，请等待......"));
	SetTimer(EXCEL_CRE_TIMER_NUM, 10, 0);

	CString timeStr;
	SYSTEMTIME sysTime;

	GetLocalTime(&sysTime);
	timeStr.Format(L"%4d.%2d.%2d-%2d.%2d ", sysTime.wYear, sysTime.wMonth, sysTime.wDay, sysTime.wHour, sysTime.wMinute);

	app_get_exe_path(excel_path,MAX_PATH);
	FileName = _T("nb") + timeStr + _T(".xls");
	excPathStr = excel_path;
	excPathStr += FileName;

	//读写excel前需要先初始化
	AfxOleInit();

	//设置列表框控件
	CRect rectL;
	int ser_wd;
	int time_wd;
	int mac_wd;
	int total_wd;
	int esim_imei_wd;
	mList.GetWindowRect(&rectL);

	total_wd = rectL.right - rectL.left;
	ser_wd = 120;
	time_wd = 200;
	mac_wd = 200;

	esim_imei_wd = (total_wd - ser_wd- time_wd- mac_wd) / 2;

	mList.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);
	mList.SetTextBkColor(RGB(224, 238, 238));
	mList.InsertColumn(0, _T("序号"), LVCFMT_LEFT, ser_wd);
	mList.InsertColumn(1, _T("时间"), LVCFMT_LEFT, time_wd);
	mList.InsertColumn(2, _T("MAC"), LVCFMT_LEFT, mac_wd);
	mList.InsertColumn(3, _T("NB esim(卡号)"), LVCFMT_LEFT, esim_imei_wd);
	mList.InsertColumn(4, _T("NB imei(模块号)"), LVCFMT_LEFT, esim_imei_wd);
	
	//找到全部串口
	vector<SerialPortInfo> m_portsList = CSerialPortInfo::availablePortInfos();
	TCHAR m_regKeyValue[255];
	std::string portInfo;

	portInfo.clear();
	for (int i = 0; i < m_portsList.size(); i++)
	{
#ifdef UNICODE
		int iLength;
		portInfo.append(m_portsList[i].portName)
				.append(" ")
				.append(m_portsList[i].description)
				.append("\n");

	    const char * _char = portInfo.c_str();
		CString test_str;

		test_str = _char;
		db_str(test_str + _T("\r\n"));

		iLength = MultiByteToWideChar(CP_ACP, 0, _char, strlen(_char) + 1, NULL, 0);
		MultiByteToWideChar(CP_ACP, 0, _char, strlen(_char) + 1, m_regKeyValue, iLength);
#else
		strcpy_s(m_regKeyValue, 255, m_portsList[i].portName.c_str());
#endif
		box.AddString(m_regKeyValue);

		portInfo.clear();
	}
	m_SerialPort.readReady.connect(this, &CreadesimDlg::OnReceive);
	
	CString baud_str = _T("");

	baud_str.Format(_T("%d"), 115200);
	box_baud.AddString(baud_str);
	baud_str.Format(_T("%d"), 9600);
	box_baud.AddString(baud_str);

	box.SetCurSel(0);
	box_baud.SetCurSel(0);

//不使用多媒体定时器(软件定时器可以满足要求)
#ifndef SF_TIMER
	app_queue_init(&d_q,10,sizeof(nb_dt_t));
	timeSetEvent(1, 1, (LPTIMECALLBACK)My_MMTimerProc, duser, TIME_PERIODIC);
#endif

	//查找打印机
	DWORD Flags = PRINTER_ENUM_FAVORITE | PRINTER_ENUM_LOCAL;
	DWORD cbBuf;
	DWORD pcReturned;

	DWORD Level = 2;
	TCHAR Name[500] = { 0 };

	::EnumPrinters(Flags,Name,Level,NULL,0,&cbBuf, &pcReturned); //需要多少内存 
	const LPPRINTER_INFO_2 pPrinterEnum = (LPPRINTER_INFO_2)LocalAlloc(LPTR, cbBuf + 4);
	EnumPrinters(Flags,Name,Level,(LPBYTE)pPrinterEnum,cbBuf,&cbBuf,&pcReturned);

	for (unsigned int i = 0; i < pcReturned; i++)
	{
		LPPRINTER_INFO_2 pInfo = &pPrinterEnum[i];
		mPrintList.AddString(pInfo->pPrinterName);
	}
	LocalFree(pPrinterEnum);
	mPrintList.SetCurSel(1);

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CreadesimDlg::OnReceive()
{
	char * str = NULL;
	str = new char[1024];
	//int buf_len = m_SerialPort.readAllData(str);

	int buf_len = m_SerialPort.readData(str, 1024);
	if ((rx_len + buf_len) > RX_BUF_SIZE)
	{
		buf_len = RX_BUF_SIZE - rx_len;
	}

	memcpy(rx_buf + rx_len, (char*)str, buf_len);
	rx_len += buf_len;
	rx_buf[rx_len] = 0;
	rx_timeout = 3;
}

void CreadesimDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CreadesimDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标显示。
HCURSOR CreadesimDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CreadesimDlg::OnBnClickedButton1()
{
	// TODO: 在此添加控件通知处理程序代码
	//SetTimer(1, 200, 0);
}

void CreadesimDlg::OnTimer(UINT_PTR nIDEvent)
{
	// TODO: 在此添加消息处理程序代码和/或调用默认值
#ifdef SF_TIMER
	if (nIDEvent == EXCEL_CRE_TIMER_NUM)
	{
		BOOL exl_suc = FALSE;

		KillTimer(EXCEL_CRE_TIMER_NUM);
		exl_suc = exc.InitExcel(excPathStr, TRUE);
		if (exl_suc == TRUE)
		{
			CString strSheetName;
		
			exc.OpenExcelFile(excPathStr);
			strSheetName = exc.GetSheetName(1);
			exc.LoadSheet(strSheetName);

			exc.SetCellString(1, 1, _T("序号"));
			exc.SetCellString(1, 2, _T("时间"));
			exc.SetCellString(1, 3, _T("MAC"));
			exc.SetCellString(1, 4, _T("ESIM"));
			exc.SetCellString(1, 5, _T("IMEI"));

			exc.SaveasXSLFile(excPathStr);
			exc.CloseExcelFile(FALSE);
		
			db_str(_T("excel创建成功！文件名称：") + FileName + _T("\r\n"));
			//文件创建成功
			excel_cre_f = 1;
			state_color = 1;
			excel_idx++;
			mState.SetWindowTextW(_T("文件创建成功！"));
			
			SetTimer(SER_TIMER_NUM,10,0);
		}
		else
		{
			MessageBox(_T("excel创建失败，请确保你的电脑安装了excel！"));
		}
	}

	if (nIDEvent == SER_TIMER_NUM)
	{
		if (rx_timeout)
		{
			rx_timeout--;
			if (rx_timeout == 0)
			{
				char *p, *q;
				CString str;
				uint16_t len;
				CString str_db;

				str_db = rx_buf;
				db_str(_T("RX:") + str_db);

				//防止重入
				KillTimer(EXCEL_CRE_TIMER_NUM);
				rx_len = 0;

				if ((p = strstr(rx_buf, "MAC:")) && (q = strstr(p, "\r\n")))
				{
					read_syn_f = 0;
					read_step = 1;

					p = p + 4;
					len = q - p;

					if (len < 30)
					{
						memset(ble_mac, 0, len+3);
						memcpy(ble_mac, p, len);
						str = ble_mac;
			
						db_str(_T("MAC:"));
						db_str(str + _T("\r\n"));
					}
					else
					{
						db_str(_T("MAC length error!\r\n"));
					}
				}
				else if ((p = strstr(rx_buf, "+QCCID:")) && (q = strstr(p, "\r\n")))
				{
					read_step++;
					p = p + 8;
					len = q - p;

					if (len < 50)
					{
						memset(nb_esim, 0, len+3);
						memcpy(nb_esim, p, len);

						str = nb_esim;
						db_str(_T("esim:"));
						db_str(str + _T("\r\n"));
					}
					else
					{
						db_str(_T("esim length error!\r\n"));
					}
				}
				else if ((p = strstr(rx_buf, "+CGSN:")) && (q = strstr(p, "\r\n")))
				{
					read_step++;
					p = p + 7;
					len = q - p;

					if (len < 50)
					{
						memset(nb_imei, 0, len+3);
						memcpy(nb_imei, p, len);

						str = nb_imei;
						db_str(_T("imei:"));
						db_str(str + _T("\r\n"));
					}
					else
					{
						db_str(_T("imei length error!\r\n"));
					}
				}
				else if ((p = strstr(rx_buf, "+CEREG")) && (read_syn_f == 0))
				{
					read_syn_f = 1;
					//读取结束
					if (read_step == 3)
					{
						//int list_num;
						uint8_t err_f = 0;
						CString str_temp, time_str;
						CTime time = CTime::GetCurrentTime();
						time_str = time.Format("20%y-%m-%d-%H:%M:%S");

						excel_idx++;
						//在结尾附带'\n'是因为如果不附带这个符号，excel中以数字展现这个字符串
						strcat_s(nb_esim, "\n");
						strcat_s(nb_imei, "\n");
						//如果列表框中的项数超过10000，清除列表框
						if (lst_idx > 10000)
						{
							mList.DeleteAllItems();
							lst_idx = 0;
						}
						////打开并加载excel的sheet1
						exc.OpenExcelFile(excPathStr);
						CString strSheetName = exc.GetSheetName(1);
						exc.LoadSheet(strSheetName);
						//查找重复
						if (excel_idx > 2)
						{
							CString  mac_str, imei_str, esim_str;
							CString  mac_str_rd,imei_str_rd,esim_str_rd;

							mac_str = ble_mac;
							imei_str = nb_imei;
							esim_str = nb_esim;
							for (uint32_t idx = 2; idx < excel_idx; idx++)
							{
								mac_str_rd = exc.GetCellString(idx, 3);
								esim_str_rd = exc.GetCellString(idx, 4);
								imei_str_rd = exc.GetCellString(idx, 5);

								if (mac_str_rd == mac_str && esim_str_rd == nb_esim && imei_str_rd == imei_str)
								{
									state_color = 2;
									mState.SetWindowTextW(_T("结果：MAC/ESIM/IMEI重复，请拿下机板后重新读取！"));
									db_str(_T("MAC/ESIM/IMEI重复!\r\n"));
									db_str(_T("重复MAC:") + mac_str + _T("\r\n"));
									db_str(_T("重复esim:") + esim_str_rd + _T("\r\n"));
									db_str(_T("重复imei:") + imei_str_rd + _T("\r\n"));
									err_f = 1;
									break;
								}
								else
								{
									err_f = 0;
								}
							}
						}
						//没有重复，写入文件并显示在列表框中///////
						if (err_f == 0)
						{
							mList.InsertItem(lst_idx, _T(""));
							str_temp.Format(_T("%d"), lst_idx + 1);
							//序号
							mList.SetItemText(lst_idx, 0, str_temp);
							exc.SetCellString(excel_idx, 1, str_temp);
							//时间
							mList.SetItemText(lst_idx, 1, time_str);
							exc.SetCellString(excel_idx, 2, time_str);
							//MAC
							str_temp = ble_mac;
							mList.SetItemText(lst_idx, 2, str_temp);
							exc.SetCellString(excel_idx, 3, str_temp);
							//ESIM
							str_temp = nb_esim;
							mList.SetItemText(lst_idx, 3, str_temp);
							exc.SetCellString(excel_idx, 4, str_temp);
							//IMIE
							str_temp = nb_imei;
							mList.SetItemText(lst_idx, 4, str_temp);
							exc.SetCellString(excel_idx, 5, str_temp);

							lst_idx++;

							db_str(_T("数据写入excel！\r\n"));
							state_color = 1;
							mState.SetWindowTextW(_T("结果：数据已写入excel！"));
						}

						exc.SaveasXSLFile(excPathStr);
						exc.CloseExcelFile(FALSE);
					}
					else
					{
						db_str(_T("读取错误！\r\n"));
						db_print("read_step:%d\r\n", read_step);
						state_color = 2;
						mState.SetWindowTextW(_T("结果：数据写入失败，请拿下机板后重新读取！"));
					}
					read_step = 0;
				}

				SetTimer(SER_TIMER_NUM, 10, 0);
			}
		}
#else
		if (app_queue_none(&d_q) != Q_NONE)
		{
			CString strSheetName;
			CString str;
			uint32_t row, col;
			nb_dt_t dt;

			exc.OpenExcelFile(excPathStr);
			strSheetName = exc.GetSheetName(1);
			exc.LoadSheet(strSheetName);
			col = exc.GetColumnCount();
			row = exc.GetRowCount();

			//跳过第一行
			row++;
			db_print("col=%d,row=%d\r\n", col, row);
			while (app_queue_none(&d_q) != Q_NONE)
			{
				nb_dt_t dt;
				
				app_dequeue(&d_q, &dt);
				//序号
				str = dt.idx;
				exc.SetCellString(row, 1, str);
				//时间
				str = dt.time;
				exc.SetCellString(row, 2, str);
				//MAC
				str = dt.mac;
				exc.SetCellString(row, 3, str);
				//imei
				str = dt.nb_imei;
				exc.SetCellString(row, 4, str);

				//esim
				str = dt.nb_esim;
				exc.SetCellString(row, 5, str);;

				row++;
			}
		
			exc.SaveasXSLFile(excPathStr);
			exc.CloseExcelFile(FALSE);
			db_str(_T("excel写入完成！"));
		}
		else
		{
			db_str(_T("timer stop\r\n"));
			KillTimer(1);
		}
#endif
	}

	CDialogEx::OnTimer(nIDEvent);
}

void CreadesimDlg::db_str(CString str)
{
//	int len;
	int nLength = db.GetWindowTextLength();

	if (nLength >= 59000)
	{
		db.SetWindowTextW(_T(""));
	}
	db.SetSel(nLength, nLength);
	db.ReplaceSel(str);
	db.LineScroll(db.GetLineCount());
}

void CreadesimDlg::db_print(char *pformat, ...)
{
	CString str;
	char temp[1024];

	va_list aptr;

	va_start(aptr, pformat);
	vsprintf_s(temp, pformat, aptr);
	va_end(aptr);

	str = temp;

	db_str(str);
}

void CreadesimDlg::db_hex(char *str_head, uint8_t *dt, uint16_t len, char *str_end)
{
	CString s_head,s_end;

	s_head = str_head;
	s_end = str_end;

	db_str(s_head);
	while (len--)
	{
		db_print("%02X ", *dt++);
	}
	db_str(s_end);
}

BOOL serail = FALSE;
uint8_t sel_pre, sel_cur;

void CreadesimDlg::OnBnClickedButton2()
{
	// TODO: 在此添加控件通知处理程序代码
	if (excel_cre_f != 1)
	{
		MessageBox(_T("打开失败，Excel文件未创建！"));
	}

	if (serail == FALSE)
	{
		int  SelBaudRate;
		CString temp, baud_str;
		string portName;

		box.GetWindowText(temp);///CString temp
		if (temp == _T(""))
		{
			MessageBox(_T("请选择串口！"), _T("提示："));
			return;
		}

		int idx,len,del_num;

		idx = temp.Find(_T(" "));
		len = temp.GetLength();
		del_num = len - idx;
		//提取字符串:COMxx
		temp.Delete(idx, del_num);
		
		db_str(_T("正在打开") + temp + _T("\r\n"));

		//波特率设置
		box_baud.GetWindowText(baud_str);
		SelBaudRate = _tstoi(baud_str);

#ifdef UNICODE
		portName = CW2A(temp.GetString());
#else
		portName = temp.GetBuffer();
#endif	
		temp.Delete(0, 3);
		sel_pre = sel_cur = _tstoi(temp);

		m_SerialPort.init(portName, SelBaudRate);
		m_SerialPort.open();

		if (m_SerialPort.isOpened())
		{
			serail = TRUE;

			db_str(_T("串口打开成功\r\n"));
			com_st.SetWindowText(__T("已打开"));
			but1.SetWindowTextW(_T("关闭"));
			box_baud.EnableWindow(FALSE);
		}
		else
		{
			box_baud.EnableWindow(TRUE);
			com_st.SetWindowText(__T("打开失败"));
		}
	}
	else
	{
		serail = FALSE;
		m_SerialPort.close();

		but1.SetWindowTextW(_T("打开"));
		com_st.SetWindowText(__T("已关闭"));
		box_baud.EnableWindow(TRUE);
	}
}

void CAboutDlg::OnCbnSelchangeCombo1()
{
	// TODO: 在此添加控件通知处理程序代码

}

void CreadesimDlg::OnCbnSelchangeCombo1()
{
	// TODO: 在此添加控件通知处理程序代码
	CString str = _T("");
	int idx, len, del_num;

	box.GetWindowText(str);

	idx = str.Find(_T(" "));
	len = str.GetLength();
	del_num = len - idx;
	//提取字符串:COMxx
	str.Delete(idx, del_num);

	if (str.GetLength() >= 4 && (str.Find(_T("COM")) != -1))
	{
		if (str.GetLength() == 4)
		{
			sel_cur = _ttoi(str.Mid(3));
			db_str(_T("COM") + str.Mid(3));
		}
		else
		{
			sel_cur = _ttoi(str.Mid(3, 2));
			db_str(_T("COM") + str.Mid(3, 2));
		}

		db_str(_T("被选中\r\n"));
	}
	else
	{
		return;
	}

	if (sel_cur != sel_pre)
	{
		if (serail == TRUE)
		{
			m_SerialPort.close();
			serail = FALSE;

			but1.SetWindowTextW(_T("打开"));
			com_st.SetWindowText(_T("已关闭"));
		}
	}

	sel_pre = sel_cur;
}


HBRUSH CreadesimDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CDialogEx::OnCtlColor(pDC, pWnd, nCtlColor);

	// TODO:  在此更改 DC 的任何特性
	switch (pWnd->GetDlgCtrlID())
	{
	  case IDC_STATIC11:
		if (serail == TRUE)
		{
			pDC->SetBkColor(RGB(0, 255, 0));
		}
		else
		{
			pDC->SetBkColor(RGB(255, 0, 0));
		}
		break;

	  case IDC_STATIC6:
		  if (state_color == 1)
		  {
			  pDC->SetBkColor(RGB(0, 255, 0));
		  }
		  else if (state_color == 2)
		  {
			  pDC->SetBkColor(RGB(255, 0, 0));
		  }
		  break;
	}

	// TODO:  如果默认的不是所需画笔，则返回另一个画笔
	return hbr;
}


void CreadesimDlg::OnBnClickedButton5()
{
	// TODO: 在此添加控件通知处理程序代码
	MessageBox(_T("不好意思，没打印机......"));
}


void CreadesimDlg::OnClose()
{
	// TODO: 在此添加消息处理程序代码和/或调用默认值
	int state;

	state = MessageBox(_T("要退出程序？"), _T("警告!"), MB_OKCANCEL| MB_ICONEXCLAMATION);
	if (state == IDOK)
	{
		exc.ReleaseExcel();
		m_SerialPort.close();
		CDialogEx::OnClose();
	}
}
