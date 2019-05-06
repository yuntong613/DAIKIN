#pragma once
#include "YOPCDevice.h"
#include "Log.h"

#include "resource.h"

class YSerialDevice : public YOPCDevice
{
public:
	enum { DEVICENAME = IDS_DEVICENAME };
	YSerialDevice(LPCSTR pszAppPath);
	virtual ~YSerialDevice(void);
	virtual bool SetDeviceItemValue(CBaseItem* pAppItem);
	virtual void OnUpdate();

	CString GetCommandHexStr(CString strAddr, CString strCmd, CByteArray& byteAll);
	bool CheckSum(CString szText);
	void Handle42Data(CString strCom, BYTE* cpData, int nLen);
	void Handle43Data(CString strCom, BYTE* cpData, int nLen);
	void Handle44Data(CString strCom, BYTE* cpData, int nLen);
	void Handle47Data(CString strCom, BYTE* cpData, int nLen);

	virtual void HandleData();
	virtual void Serialize(CArchive& ar);
	BOOL InitConfig(CString strFilePath);
	void Load(CArchive& ar);
	void LoadItems(CArchive& ar);
	
	int QueryOnce();

	virtual void BeginUpdateThread();
	virtual void EndUpdateThread();

	BYTE Hex2Bin(CString strHex);
	int HexStr2Bin(BYTE * cpData, CString strData);
	CString Bin2HexStr(BYTE* cpData, int nLen);
public:
	long y_lUpdateTimer;
public:
	CStringArray m_ComPortArray;
	int m_nBaudRate;
	int m_nParity;
	int m_nTimeOut;
public:
	int m_nUseLog;
	CLog m_Log;
	void OutPutLog(CString strMsg);
public:
	HANDLE m_hQueryThread;
	bool m_bStop;

protected:
	CString m_strConfigPath;
	CStringArray m_AddrArray;
};
