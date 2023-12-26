#pragma once

#ifndef BIOSINFO_H
#define BIOSINFO_H

#pragma comment (lib, "Ole32.lib")
#pragma comment (lib, "User32.lib")
#pragma comment (lib, "wbemuuid.lib")

#include <string>
#include <wbemidl.h>
#include <Windows.h>

#include <vcclr.h>
#include <msclr/marshal_cppstd.h>

using namespace System;
using namespace System::Text;
using namespace System::Runtime::InteropServices;
using namespace msclr::interop;

class BiosInfo;
ref class BiosInfoOutput;

// class: 
// BiosInfo
//
// �����������:
// ������� � ������� WMI ��� 
// ���������� ������ ����� BIOS
class BiosInfo {
    // �������� ����� BiosInfoOutput ���������� 
    // ������ �� ��������� ����� ����� �����
    friend BiosInfoOutput;
private:
    // ����, �� ������ ��� ��������� ������ �� WMI
    HRESULT hRes;
    DWORD errorCode;

    IWbemLocator* pLoc;					// IWbemLocator-���������
    IWbemServices* pSvc;				// IWbemServices-��������c
    IEnumWbemClassObject* pEnumerator;	// ����� ��� ������� ������ �� WMI
    IWbemClassObject* pClsObj;			// ����� ��� ��������� ���������� ������ WMI
    wchar_t errBuff[256]{};             // ����� ��� ��������� ���� �������

    // �����, ���� ������� ��� ��������� ���������� ������ �� WMI (�++ ������)
    std::wstring result;

    // ���������� ��� BIOS:
    std::wstring caption;               // ����
    std::wstring biosVersion;           // �����
    std::wstring status;                // ����
    std::wstring releaseDate;           // ���� �������
    std::wstring SMBIOSPresent;         // �������� ���������� SMBIOS
    std::wstring SMBIOSVersion;         // ���������� ��� SMBIOS ����� 
    UINT32* pCharactDat;                // ����� �������������
    long numOfCharact;                  // ʳ������ ��������� �������������


    // ������ ��� �����䳿 � WMI
    BOOL ConnectToWMI();
    BOOL GetBiosInfo();
    std::wstring GetOutput();
    std::wstring WMIDateStringToDate(const std::wstring& wmiDate);
};

ref class BiosInfoOutput
{
private:
    // �����, ���� ������� ��� ��������� ���������� ������ �� WMI (.NET ������)
    String^ result;

public:
    Boolean ConnectToWMI();
    Boolean GetBiosInfo();
    String^ GetBiosCharacteristics();
    String^ GetBiosCaption();
    String^ GetBiosVersion();
    String^ GetBiosStatus();
    String^ GetBiosReleaseDate();
    String^ GetSMBIOSPresent();
    String^ GetSMBIOSVersion();

    // ����� ���� result
    String^ GetResult();
};

#endif // BIOSINFO_H