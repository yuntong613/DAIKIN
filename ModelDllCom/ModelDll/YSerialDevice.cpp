#include "StdAfx.h"
#include "common.h"
#include "YSerialDevice.h"
#include "IniFile.h"
#include "ItemBrowseDlg.h"
#include "YSerialItem.h"
#include <cstringt.h>
#include "ModelDll.h"
#include "OPCIniFile.h"

extern CModelDllApp theApp;

DWORD CALLBACK QuertyThread(LPVOID pParam)
{
	YSerialDevice* pDevice = (YSerialDevice*)pParam;
	while(!pDevice->m_bStop)
	{
		Sleep(1000);
		pDevice->QueryOnce();
	}
	return 0;
}


YSerialDevice::YSerialDevice(LPCSTR pszAppPath)
: m_nBaudRate(9600)
, m_bStop(true)
{
	m_nParity = 0;
	m_hQueryThread = INVALID_HANDLE_VALUE;
	y_lUpdateTimer = 0;
	m_nUseLog = 0;
	CString strConfigFile(pszAppPath);
	strConfigFile+= _T("\\ComFile.ini");
	if(!InitConfig(strConfigFile))
	{
		return;
	}

	CString strListItemsFile(pszAppPath);
	strListItemsFile += _T("\\ListItems.ini");
	COPCIniFile opcFile;
	if (!opcFile.Open(strListItemsFile,CFile::modeRead|CFile::typeText))
	{
		AfxMessageBox("Can't open INI file!");
		return;
	}
	CArchive ar(&opcFile,CArchive::load);
	Serialize(ar);
	opcFile.Close();
}

YSerialDevice::~YSerialDevice(void)
{
	POSITION pos = m_ItemsArray.GetStartPosition();
	YSerialItem* pItem = NULL;
	CString strItemName;
	while(pos){
		m_ItemsArray.GetNextAssoc(pos,strItemName,(CObject*&)pItem);
		if(pItem)
		{
			delete pItem;
			pItem = NULL;
		}
	}
	m_ItemsArray.RemoveAll();
}

void YSerialDevice::Serialize(CArchive& ar)
{
	if (ar.IsStoring()){
	}else{
		Load(ar);
	}
}

BOOL YSerialDevice::InitConfig(CString strFilePath)
{
	if(!PathFileExists(strFilePath))
		return FALSE;

	m_strConfigPath = strFilePath;

	CIniFile iniFile(m_strConfigPath);
	m_lRate = iniFile.GetUInt("param","UpdateRate",3000);
	m_nUseLog = iniFile.GetUInt("param","Log",0);

	iniFile.GetArray("ComInfo","ComPort",&m_ComPortArray);
	m_nBaudRate = iniFile.GetUInt("ComInfo","BaudRate",9600);
	m_nParity = iniFile.GetUInt("ComInfo","Parity",0);

	return TRUE;
}

void YSerialDevice::Load(CArchive& ar)
{
	LoadItems(ar);
}

void YSerialDevice::LoadItems(CArchive& ar)
{
	COPCIniFile* pIniFile = static_cast<COPCIniFile*>(ar.GetFile());
	YOPCItem* pItem  = NULL;
	int nItems = 0;
	CString strTmp("Item");
	CString strItemName;
	CString strItemDesc;
	CString strValueType;
	DWORD dwItemPId = 0L;
	strTmp+=CString(_T("List"));
	if(pIniFile->ReadNoSeqSection(strTmp))
	{
		nItems = pIniFile->GetItemsCount(strTmp,"Item");
		for (int i=0;i<nItems && !pIniFile->Endof();i++ )
		{
			try{
				if (pIniFile->ReadIniItem("Item",strTmp))
				{
					if (!pIniFile->ExtractSubValue(strTmp, strItemName, 1))
						continue;
					strValueType = strItemName.Right(1);

					if(!pIniFile->ExtractSubValue(strTmp,strItemDesc,2))strItemDesc = _T("Unknown");

					if (strValueType == "F")
					{
						pItem = new YFloatItem(dwItemPId, strItemName, strItemDesc);
					}
					else if (strValueType == "S")
					{
						pItem = new YShortItem(dwItemPId, strItemName, strItemDesc);
					}
					
					if(GetItemByName(strItemName))
						delete pItem;
					else 
						m_ItemsArray.SetAt(pItem->GetName(),(CObject*)pItem);

					dwItemPId++;
				}
			}
			catch(CItemException* e){
				if(pItem) delete pItem;
				e->Delete();
			}
		}
	}
}

void YSerialDevice::OnUpdate()
{
// 	y_lUpdateTimer--;
// 	if(y_lUpdateTimer>0)return;
// 	y_lUpdateTimer = m_lRate/1000;

}


CString YSerialDevice::GetCommandHexStr(CString strAddr, CString strCmd, CByteArray& byteAll)
{
	byteAll.RemoveAll();
	byteAll.Add(0x7E);
	CString strText = "20" + strAddr + "60" + strCmd + "0000";
	WORD lSum = 0;
	for (int i = 0; i < strText.GetLength(); i++)
	{
		char c = strText.GetAt(i);
		byteAll.Add(c);
		lSum += (int)c;
	}
	WORD dwRet = ~lSum + 1;
	CString strTmp;
	strTmp.Format("%04X", dwRet);

	for (int i = 0; i < strTmp.GetLength(); i++)
	{
		char c = strTmp.GetAt(i);
		byteAll.Add(c);
	}
	byteAll.Add(0x0D);

	CString strHex;
	for (int k = 0; k < byteAll.GetCount(); k++)
	{
		strTmp.Format("%02X ", byteAll.GetAt(k));
		strHex += strTmp;
	}
	strHex.TrimRight();
	return strHex;
}

bool YSerialDevice::CheckSum(CString szText)
{
	WORD lSum = 0;
	for (int i = 0; i < szText.GetLength()-4; i++)
	{
		char c = szText.GetAt(i);
		lSum += (int)c;
	}
	WORD dwRet = ~lSum + 1;
	CString strTmp;
	strTmp.Format("%04X", dwRet);
	return (strTmp == szText.Right(4));
}

void YSerialDevice::Handle42Data(CString strCom, BYTE* cpData, int nLen)
{
	CString strFun = "42";
	if ((nLen == 70) && (cpData[0] == 0x7E) && (cpData[nLen - 1] == 0x0D))
	{
		char szData[100] = { 0 };
		memcpy(szData, cpData + 1, nLen - 2);
		CString szText = szData;
		if (CheckSum(szText))
		{
			CString strVer = szText.Left(2);
			szText.Delete(0, 2);

			CString strAddr = szText.Left(2);
			szText.Delete(0, 2);

			CString strCID1 = szText.Left(2);
			szText.Delete(0, 2);

			CString strRTN = szText.Left(2);
			szText.Delete(0, 2);

			if (strRTN == "00")
			{
				CString strLenID = szText.Left(4);
				szText.Delete(0, 4);			
				if (strLenID == "9034")
				{
					szText.Delete(0, 30);//移除空格
					CString strTmp = szText.Left(4);
					CString strValue;
					strValue.Format("%.2f", strtol(strTmp, NULL, 16)/100.0);

					CString strItemName;
					strItemName.Format("%s-%s-%s-RAT-F", strCom, strAddr, strFun);
					YOPCItem* pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(strValue);

					int K = 0;
				}
			}
			else
			{
				OutPutLog("RTN异常代码 " + strRTN);
			}
		}
		else {
			OutPutLog("校验出错...");
		}
	}
}

void YSerialDevice::Handle43Data(CString strCom, BYTE* cpData, int nLen)
{
	CString strFun = "43";
	CString strItemName,strNameHead;
	YOPCItem* pItem = NULL;
	if ((nLen == 42) && (cpData[0] == 0x7E) && (cpData[nLen - 1] == 0x0D))
	{
		char szData[100] = { 0 };
		memcpy(szData, cpData + 1, nLen - 2);
		CString szText = szData;
		if (CheckSum(szText))
		{
			CString strVer = szText.Left(2);
			szText.Delete(0, 2);

			CString strAddr = szText.Left(2);
			szText.Delete(0, 2);

			CString strCID1 = szText.Left(2);
			szText.Delete(0, 2);

			CString strRTN = szText.Left(2);
			szText.Delete(0, 2);

			strNameHead.Format("%s-%s", strAddr, strFun);
			if (strRTN == "00")
			{
				CString strLenID = szText.Left(4);
				szText.Delete(0, 4);
				if (strLenID == "7018")
				{
					szText.Delete(0, 2);//移除FLAG

					CString strRunStatus = szText.Left(2);
					szText.Delete(0, 2);//移除运行状态 00 停止, 01 运转
					strItemName.Format("%s-%s-KTRS-S", strCom, strNameHead);
					pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(strRunStatus);

					CString strSelfVarNum = szText.Left(2);
					szText.Delete(0, 2);//移除自定义状态数量

					CString strJSZ = szText.Left(2);
					szText.Delete(0, 2);//移除加湿器状态 00 停止, 01 运转
					strItemName.Format("%s-%s-JSRRS-S", strCom, strNameHead);
					pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(strJSZ);


					CString strDJR = szText.Left(2); 
					szText.Delete(0, 2);//移除电加热状态 00 停止, 01 运转
					strItemName.Format("%s-%s-DJRRS-S", strCom, strNameHead);
					pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(strDJR);

					CString strFunRS = szText.Left(2);
					szText.Delete(0, 2);//移除风扇状态 00 停止, 01 运转
					strItemName.Format("%s-%s-FUNRS-S", strCom, strNameHead);
					pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(strFunRS);


					CString strYSJ = szText.Left(2);
					szText.Delete(0, 2);//移除压缩机状态 00 停止, 01 运转
					strItemName.Format("%s-%s-YSJRS-S", strCom, strNameHead);
					pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(strYSJ);

					CString strGLW = szText.Left(2);
					szText.Delete(0, 2);//移除过滤网状态 00 OFF, 01 ON
					strItemName.Format("%s-%s-GLWRS-S", strCom, strNameHead);
					pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(strGLW);

					CString strZYB = szText.Left(2);
					szText.Delete(0, 2);//移除注意报状态 00 OFF, 01 ON
					strItemName.Format("%s-%s-ZYBRS-S", strCom, strNameHead);
					pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(strZYB);

					CString strAlarm = szText.Left(2);
					szText.Delete(0, 2);//移除报警状态 00 OFF, 01 ON
					strItemName.Format("%s-%s-ALARM-S", strCom, strNameHead);
					pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(strAlarm);

					CString strError = szText.Left(2);
					szText.Delete(0, 2);//移除异常状态 00 OFF, 01 ON
					strItemName.Format("%s-%s-ERROR-S", strCom, strNameHead);
					pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(strError);

					CString strRunMode = szText.Left(2);
					szText.Delete(0, 2);//移除运转状态   00H 送风, 01H 制热, 02H 制冷
					strItemName.Format("%s-%s-MODE-S", strCom, strNameHead);
					pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(strRunMode);

					int K = 0;
				}
			}
			else
			{
				OutPutLog("RTN异常代码 " + strRTN);
			}
		}
		else {
			OutPutLog("校验出错...");
		}
	}
}

void YSerialDevice::Handle44Data(CString strCom, BYTE* cpData, int nLen)
{
	CString strFun = "44";
	if ((nLen == 46) && (cpData[0] == 0x7E) && (cpData[nLen - 1] == 0x0D))
	{
		char szData[100] = { 0 };
		memcpy(szData, cpData + 1, nLen - 2);
		CString szText = szData;
		if (CheckSum(szText))
		{
			CString strVer = szText.Left(2);
			szText.Delete(0, 2);

			CString strAddr = szText.Left(2);
			szText.Delete(0, 2);

			CString strCID1 = szText.Left(2);
			szText.Delete(0, 2);

			CString strRTN = szText.Left(2);
			szText.Delete(0, 2);

			if (strRTN == "00")
			{
				CString strLenID = szText.Left(4);
				szText.Delete(0, 4);
				if (strLenID == "301C")
				{
					szText.Delete(0, 2);//移除FLAG
					szText.Delete(0, 12);//移除空格

					CString strHFAalarm = szText.Left(2);
					szText.Delete(0, 2);//移除回风运行状态 00 正常， F0 故障 ， 01H 下限以下， 02H 上限以上 
					CString strItemName;
					strItemName.Format("%s-%s-%s-RATRS-S", strCom, strAddr, strFun);
					YOPCItem* pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(strHFAalarm);

				}
			}
			else
			{
				OutPutLog("RTN异常代码 " + strRTN);
			}
		}
		else {
			OutPutLog("校验出错...");
		}
	}
}

void YSerialDevice::Handle47Data(CString strCom, BYTE* cpData, int nLen)
{
	CString strFun = "47";
	CString strItemName, strNameHead;
	YOPCItem* pItem = NULL;
	if ((nLen == 44) && (cpData[0] == 0x7E) && (cpData[nLen - 1] == 0x0D))
	{
		char szData[100] = { 0 };
		memcpy(szData, cpData + 1, nLen - 2);
		CString szText = szData;
		if (CheckSum(szText))
		{
			CString strVer = szText.Left(2);
			szText.Delete(0, 2);

			CString strAddr = szText.Left(2);
			szText.Delete(0, 2);

			CString strCID1 = szText.Left(2);
			szText.Delete(0, 2);

			CString strRTN = szText.Left(2);
			szText.Delete(0, 2);

			strNameHead.Format("%s-%s", strAddr, strFun);
			if (strRTN == "00")
			{
				CString strLenID = szText.Left(4);
				szText.Delete(0, 4);
				if (strLenID == "501A")
				{
					//Step1
					CString strTmp = szText.Left(4);
					szText.Delete(0, 4);//移除开机温度
					
					CString strOnTmp;
					strOnTmp.Format("%.2f", strtol(strTmp, NULL, 16) / 100.0);
					strItemName.Format("%s-%s-TEMPON-F", strCom, strNameHead);
					pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(strOnTmp);

					//Step2
					strTmp = szText.Left(4);
					szText.Delete(0, 4);//移除关机温度

					CString strOffTmp;
					strOffTmp.Format("%.2f", strtol(strTmp, NULL, 16) / 100.0);
					strItemName.Format("%s-%s-TEMPOFF-F", strCom, strNameHead);
					pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(strOffTmp);

					//Step3
					strTmp = szText.Left(4);
					szText.Delete(0, 4);//移除回风温度上限

					CString strTmpHi;
					strTmpHi.Format("%.2f", strtol(strTmp, NULL, 16) / 100.0);
					strItemName.Format("%s-%s-RATHI-F", strCom, strNameHead);
					pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(strTmpHi);

					//Step4
					strTmp = szText.Left(4);
					szText.Delete(0, 4);//移除回风温度下限

					CString strTmpLo;
					strTmpLo.Format("%.2f", strtol(strTmp, NULL, 16) / 100.0);
					strItemName.Format("%s-%s-RATLO-F", strCom, strNameHead);
					pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(strTmpLo);

					//Step5
					strTmp = szText.Left(4);
					szText.Delete(0, 4);//移除回风湿度上限

					CString strHumidityHi;
					strHumidityHi.Format("%.2f", strtol(strTmp, NULL, 16) / 100.0);
					strItemName.Format("%s-%s-RAHHI-F", strCom, strNameHead);
					pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(strHumidityHi);

					//Step6
					strTmp = szText.Left(4);
					szText.Delete(0, 4);//移除回风温度下限

					CString strHumidityLo;
					strHumidityLo.Format("%.2f", strtol(strTmp, NULL, 16) / 100.0);
					strItemName.Format("%s-%s-RAHLO-F", strCom, strNameHead);
					pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(strHumidityLo);

				}
			}
			else
			{
				OutPutLog("RTN异常代码 " + strRTN);
			}
		}
		else {
			OutPutLog("校验出错...");
		}
	}
}

int YSerialDevice::QueryOnce()
{
	CString strAddr,strCom, strSec;
	CByteArray data;
	CString strHex;
	BYTE cpRecv[1024] = { 0 };
	DWORD dwRead = 0;
	int nComPort = 0;

	CIniFile iniFile(m_strConfigPath);
	for (int k = 0; k < m_ComPortArray.GetCount(); k++)
	{
		nComPort = atoi(m_ComPortArray.GetAt(k));	
		strSec.Format("COM%d", nComPort);
		iniFile.GetArray(strSec, "Addr", &m_AddrArray);

		if (m_Com.Open(nComPort, m_nBaudRate, m_nParity))
		{
			for (int i = 0; i < m_AddrArray.GetCount(); i++)
			{
				strAddr = m_AddrArray.GetAt(i);
				strCom.Format("C%d", nComPort);
				strHex = GetCommandHexStr(strAddr, "42", data);
				m_Com.Write(data.GetData(), data.GetCount());
				OutPutLog("[发送] " + strHex);
				dwRead = m_Com.Read(cpRecv, 1024, 1000);
				OutPutLog("[接收] " + Bin2HexStr(cpRecv, dwRead));
				Handle42Data(strCom, cpRecv, dwRead);


				data.RemoveAll();
				memset(cpRecv, 0, 1024);
				strHex = GetCommandHexStr(strAddr, "43", data);
				m_Com.Write(data.GetData(), data.GetCount());
				OutPutLog("[发送] " + strHex);
				dwRead = m_Com.Read(cpRecv, 1024, 1000);
				OutPutLog("[接收] " + Bin2HexStr(cpRecv, dwRead));
				Handle43Data(strCom, cpRecv, dwRead);


				data.RemoveAll();
				memset(cpRecv, 0, 1024);
				strHex = GetCommandHexStr(strAddr, "44", data);
				m_Com.Write(data.GetData(), data.GetCount());
				OutPutLog("[发送] " + strHex);
				dwRead = m_Com.Read(cpRecv, 1024, 1000);
				OutPutLog("[接收] " + Bin2HexStr(cpRecv, dwRead));
				Handle44Data(strCom, cpRecv, dwRead);

				data.RemoveAll();
				memset(cpRecv, 0, 1024);
				strHex = GetCommandHexStr(strAddr, "47", data);
				m_Com.Write(data.GetData(), data.GetCount());
				OutPutLog("[发送] " + strHex);
				dwRead = m_Com.Read(cpRecv, 1024, 1000);
				OutPutLog("[接收] " + Bin2HexStr(cpRecv, dwRead));
				Handle47Data(strCom, cpRecv, dwRead);
			}


			m_Com.Close();
		}
	}




	return 0;
}

void YSerialDevice::BeginUpdateThread()
{
	if(m_hQueryThread==INVALID_HANDLE_VALUE)
	{
		if(m_bStop == true)
		{
			m_bStop = false;
			m_hQueryThread = CreateThread(NULL,0,QuertyThread,this,0,NULL);
		}		
	}
}

void YSerialDevice::EndUpdateThread()
{
	if(!m_bStop)
	{
		m_bStop = true;
		DWORD dwRet = WaitForSingleObject(m_hQueryThread,3000);
		if(dwRet==WAIT_TIMEOUT)
			TerminateThread(m_hQueryThread,0);

		m_hQueryThread = INVALID_HANDLE_VALUE;
	}
}

BYTE YSerialDevice::Hex2Bin(CString strHex)
{
	int iDec = 0;
	if(strHex.GetLength() == 2){
		char cCh = strHex[0];
		if((cCh >= '0') && (cCh <= '9'))iDec = cCh - '0';
		else if((cCh >= 'A') && (cCh <= 'F'))iDec = cCh - 'A' + 10;
		else if((cCh >= 'a') && (cCh <= 'f'))iDec = cCh - 'a' + 10;
		else return 0;
		iDec *= 16;
		cCh = strHex[1];
		if((cCh >= '0') && (cCh <= '9'))iDec += cCh - '0';
		else if((cCh >= 'A') && (cCh <= 'F'))iDec += cCh - 'A' + 10;
		else if((cCh >= 'a') && (cCh <= 'f'))iDec += cCh - 'a' + 10;
		else return 0;
	}
	return iDec & 0xff;
}

int YSerialDevice::HexStr2Bin(BYTE * cpData,CString strData)
{
	CString strByte;
	for(int i=0;i<strData.GetLength();i+=2){
		strByte = strData.Mid(i,2);
		cpData[i/2] = Hex2Bin(strByte);
	}
	return strData.GetLength() / 2;
}

CString YSerialDevice::Bin2HexStr(BYTE* cpData,int nLen)
{
	CString strResult,strTemp;
	for (int i=0;i<nLen;i++)
	{
		strTemp.Format("%02X ",cpData[i]);
		strResult+=strTemp;
	}
	return strResult;
}

void YSerialDevice::HandleData()
{

	return;
}

bool YSerialDevice::SetDeviceItemValue(CBaseItem* pAppItem)
{
	return false;
}

void YSerialDevice::OutPutLog(CString strMsg)
{
	if(m_nUseLog)
		m_Log.Write(strMsg);
}
