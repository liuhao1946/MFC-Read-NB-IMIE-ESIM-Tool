
// read esimDlg.h: 头文件
//

#pragma once

#include "IllusionExcelFile.h"
#include "stdint.h"
//#include "SerialPort.h"

#include "CSerialPort/SerialPort.h"
#include "CSerialPort/SerialPortInfo.h"

using namespace itas109;
using namespace std;


// CreadesimDlg 对话框
class CreadesimDlg : public CDialogEx, public has_slots<>
{
// 构造
public:
	CreadesimDlg(CWnd* pParent = nullptr);	// 标准构造函数

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_READESIM_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持
	void OnReceive();//About CSerialPort 

// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg LRESULT OnReceiveStr(WPARAM str, LPARAM commInfo);

	DECLARE_MESSAGE_MAP()
public:
	//debug
	afx_msg void db_str(CString);
	afx_msg void db_print(char *pformat, ...);
	afx_msg void db_hex(char *str_head, uint8_t *dt, uint16_t len, char *str_end);

	afx_msg void OnBnClickedButton1();

	IllusionExcelFile exc;

	CEdit db;
	afx_msg void OnBnClickedButton2();
	CComboBox box;
	CComboBox box_baud;
	afx_msg void OnBnClickedButton3();
	CStatic com_st;
	CButton but1;
	afx_msg void OnCbnSelchangeCombo1();
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
//	CListBox lst;
	CListCtrl mList;
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	CStatic mState;
	CComboBox mPrintList;
	CFont mStateFt;

	afx_msg void OnBnClickedButton5();
	afx_msg void OnClose();

public:
	CSerialPort m_SerialPort;//About CSerialPort
};
