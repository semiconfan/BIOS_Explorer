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
    std::wstring output;

    // ���������� ��� BIOS:
    std::wstring caption;               // ����
    std::wstring biosVersion;           // �����
    std::wstring manufacturer;          // ��������
    std::wstring releaseDate;           // ���� �������
    std::wstring SMBIOSVersion;         // ���������� ��� SMBIOS �����

    // ����, �� ������ ��� ��������� ���������� ������ �� WMI
    UINT32* pCharactDat;                // ����� ������������� BIOS

    // ������ ��� �����䳿 � WMI
    BOOL ConnectToWMI();
    BOOL GetBiosCharacteristics();
    std::wstring GetOutput();
};

ref class BiosInfoOutput
{
public:
    String^ ConnectToWMI();
    String^ GetBiosCharacteristics();
};

#endif // BIOSINFO_H