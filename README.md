# MFC NB ESIM IMEI Read-Write Tool
***
# 使用说明
#### MFC软件版本：
[vs2017](https://github.com/liuhao1946/embedded-software-module/blob/master/MFC%E5%B8%B8%E7%94%A8%E6%93%8D%E4%BD%9C/vs2017.exe)

#### 目的：
1、用来读取并保存NB蓝牙手环中的每个蓝牙MCU的MAC地址对应的NB模块IMEI号以及配套使用的sim卡号。

2、该软件较为简单，但有嵌入式中常常用到的几个模块，比如串口模块、读写excel等。因此也可以做为***范例***使用。

#### 工作说明：
1、用一个串口模块（如CP2102之类的）监控MCU的RX（NB模块的TX）。

2、MCU发送AT，同步模块波特率。

3、MCU发送MAC，格式为MAC:xx xx xx xx xx xx'\r''\n'（每个x是一个字符，两个x之间用空格隔开，最后附带上回车换行）。

4、虽然向模块发送MAC是无效的，但是由于模块有回显功能，从而让软件通过模块的TX捕捉到MCU发送的MAC。

5、200ms后，MCU发送读取IMEI指令，模块产生应答后软件就可以捕捉到。

6、200ms后，MCU发送读取ESIM指令，模块产生应答后软件就可以捕捉到。

7、200ms后，MCU发送信号查询指令（AT+CERG?)，当然也可以是其他的，只要表示读取结束就可以。软件接收到“+CERG”后就认为MAC、IMEI、ESIM读取完成，软件就把这些数据一个是展现在列表框中，一个是写入excel中。

#### 软件效果图：
![软件效果图](https://github.com/liuhao1946/MFC-Read-NB-IMIE-ESIM-Tool/blob/master/image/%E8%BD%AF%E4%BB%B6%E8%BF%90%E8%A1%8C%E6%95%88%E6%9E%9C.png)

![写入excel的效果图](https://github.com/liuhao1946/MFC-Read-NB-IMIE-ESIM-Tool/blob/master/image/%E5%86%99%E5%85%A5excel%E7%9A%84%E6%95%88%E6%9E%9C.jpg)

# 程序说明
 - 串口模块
 
   该模块（软件版本V4.1.0.201010）使用了第三方的类实现，能够稳定的接收与发送串口信息。串口相关类放在工程中的CSerialPort文件夹下。
   
   相比早期版本，该版本的串口模块有如下优点：
   
   1、能展现完整的串口名称
   
   2、更稳定可靠的发送与接收串口数据
   
   目前该模块由itas109维护，github地址：[串口模块](https://github.com/itas109/CSerialPort)。
 
 - excel创建、读、写

   1、添加一些excel操作的系统文件，方法见[excel文件创建、读取、写入方法](https://github.com/liuhao1946/embedded-software-module/blob/master/MFC%E5%B8%B8%E7%94%A8%E6%93%8D%E4%BD%9C/excel%E5%88%9B%E5%BB%BA%E3%80%81%E8%AF%BB%E3%80%81%E5%86%99/%E4%BE%8B%E5%AD%90.md)
   
   2、在工程中加入IllusionExcelFile.cpp，IllusionExcelFile.h，所有excel相关的操作都由这个接口实现，具体操作参考read esimDlg.cpp。

 - 只允许同时运行一个程序的方法(代码在read esim.cpp中56~60行)
 ```c++
	HANDLE m_hMutex = CreateMutex(NULL, FALSE, _T("model"));
	DWORD nRet = GetLastError();
	if (nRet == ERROR_ALREADY_EXISTS)
	{
		AfxMessageBox(_T("exe已经打开!"));
		CloseHandle(m_hMutex);
		m_hMutex = NULL;
		return FALSE;
	}
 ```
  - MFC的.exe文件图标更新
  
    [exe文件图标更新方法](https://blog.csdn.net/txwtech/article/details/92980545)
    
    图标被替换成为们自己的了，但是exe文件的图标还是MFC的默认图标，怎么办？
    
    [刷新系统的应用程序图标](https://www.cnblogs.com/qintangtao/p/3397812.html)
