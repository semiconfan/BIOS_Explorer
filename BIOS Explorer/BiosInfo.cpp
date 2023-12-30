#include "BiosInfo.h"

BiosInfo bInf;

std::wstring BiosInfo::WMIDateStringToDate(const std::wstring& wmiDate)
{
    // ��������� ���������, ��� ��������������� ��� ��������� ���� � ����
    struct tm tm = { 0 };

    // ����� ����� �� ����� ���������� (��, �����, ���� �� ��.)
    swscanf_s(wmiDate.c_str(), L"%4d%2d%2d%2d%2d%2d",
        &tm.tm_year, &tm.tm_mon, &tm.tm_mday, &tm.tm_hour,
        &tm.tm_min, &tm.tm_sec);

    // ����������� �������� ���� �� ����� ��� ������� wcsftime
    tm.tm_year -= 1900;
    tm.tm_mon--;

    // ������������ ��������� tm � ����� � ���������� ������ ����
    const unsigned int buffSize = 256;
    wchar_t buffer[buffSize];
    wcsftime(buffer, sizeof(buffer), L"%Y-%m-%d %H:%M:%S", &tm);
    buffer[buffSize - 1] = '\0';

    // ���������� ���������� ����� � ������ "����-��-�� ��:��:��"
    return std::wstring(buffer);
}

BOOL BiosInfo::ConnectToWMI()
{
    // ���� 1:***************************************************
    // ����������� COM.****************************************

    this->hRes = CoInitializeEx(0, COINIT_APARTMENTTHREADED);
    if (FAILED(this->hRes)) {
        // ��������� ���� ������� HRESULT
        this->errorCode = HRESULT_CODE(this->hRes);
        this->result = L"������� ����������� �������� COM. ";
        // ���������� ���� �������
        wsprintfW(errBuff, L"��� �������: 0x%X\n", errorCode);
        this->result += errBuff;
        return FALSE;                   // ϳ�������� ����������� �������
    }

    // ���� 2:***************************************************
    // ������������ ��������� ���� ������� COM ****************

    this->hRes = CoInitializeSecurity(
        NULL,
        -1,                             // �������������� COM
        NULL,                           // ������ ��������������
        NULL,                           // �������������
        RPC_C_AUTHN_LEVEL_DEFAULT,      // �������������� �� �������������
        RPC_C_IMP_LEVEL_IDENTIFY,       // ���������� �� �������������
        NULL,                           // ���������� ��� ��������������
        EOAC_NONE,                      // �������� ���������
        NULL                            // �������������
    );

    if (FAILED(this->hRes)) {
        this->errorCode = HRESULT_CODE(this->hRes);
        this->result = L"�� ������� ������������ �������. ";
        wsprintfW(errBuff, L"��� �������: 0x%X\n", errorCode);
        this->result += errBuff;

        CoUninitialize();
        return FALSE;
    }

    // ���� 3:***************************************************
    // ��������� ����������� �������� �� WMI ********************

    this->pLoc = NULL;

    this->hRes = CoCreateInstance(
        CLSID_WbemLocator,
        0,
        CLSCTX_INPROC_SERVER,
        IID_IWbemLocator,
        (LPVOID*)&this->pLoc);

    if (FAILED(this->hRes)) {
        this->errorCode = HRESULT_CODE(this->hRes);
        this->result = L"�� ������� �������� ��'��� IWbemLocator. ";
        wsprintfW(errBuff, L"��� �������: 0x%X\n", errorCode);
        this->result += errBuff;

        CoUninitialize();
        return FALSE;
    }

    // ���� 4:***************************************************
    // ϳ��������� �� WMI ����� IWbemLocator::ConnectServer �����

    this->pSvc = NULL;

    // ϳ��������� �� �������� ���� Root\Cimv2
    // ��������� ����������� � ��������� ��������� pSvc

    this->hRes = pLoc->ConnectServer(
        BSTR(L"ROOT\\CIMV2"),           // ���� �� ��'���� WMI
        NULL,                           // ��'� �����������. NULL = �������� ����������
        NULL,                           // ������ �����������. NULL = �������� ������
        0,
        NULL,
        0,
        0,
        &this->pSvc                     // �������� IWbemServices proxy
    );

    if (FAILED(this->hRes)) {
        this->errorCode = HRESULT_CODE(this->hRes);
        this->result = L"�� ������� �����������. ";
        wsprintfW(errBuff, L"��� �������: 0x%X\n", errorCode);
        this->result += errBuff;

        CoUninitialize();
        return FALSE;
    }

    this->result = L"ϳ�������� �� �������� ���� Root\Cimv2 ���������� WMI \n";

    // ���� 5:***************************************************
    // ������������ ���� ������� ��� WMI-�'������� ************

    this->hRes = CoSetProxyBlanket(
        this->pSvc,                     // ����� �� �����, ���� ������� ����������
        RPC_C_AUTHN_WINNT,              // RPC_C_AUTHN_xxx
        RPC_C_AUTHZ_NONE,               // RPC_C_AUTHZ_xxx
        NULL,                           // ������� ��'� �������
        RPC_C_AUTHN_LEVEL_CALL,         // RPC_C_AUTHN_LEVEL_xxx
        RPC_C_IMP_LEVEL_IMPERSONATE,    // RPC_C_IMP_LEVEL_xxx
        NULL,                           // ������������� �볺���
        EOAC_NONE                       // ��������� �����
    );

    if (FAILED(this->hRes)) {
        this->errorCode = HRESULT_CODE(this->hRes);
        this->result = L"�� ������� ������������ ��� ������� ��� WMI-�'�������. ";
        wsprintfW(errBuff, L"��� �������: 0x%X\n", errorCode);
        this->result += errBuff;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return FALSE;
    }

    return TRUE;
}

BOOL BiosInfo::GetBiosInfo()
{
    BOOL success = FALSE;

    // ���� 6: -------------------------------------------------
    // ������� ������ �� WMI -----------------------------------

    this->pEnumerator = NULL;
    this->hRes = pSvc->ExecQuery(
        BSTR(L"WQL"),
        BSTR(L"SELECT * FROM Win32_BIOS"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &pEnumerator);

    if (FAILED(this->hRes))
    {
        this->errorCode = HRESULT_CODE(this->hRes);

        this->result = L"����� ���������� �������.";
        wsprintfW(this->errBuff, L"��� �������: 0x%X\n", errorCode);
        this->result += this->errBuff;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return success;                 // �������� ����������� �������.
    }

    // ���� 7: -------------------------------------------------
    // ��������� ����� �� ������ �� ����� 6 --------------------

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;

    while (pEnumerator)
    {
        this->hRes = pEnumerator->Next(WBEM_INFINITE, 1,
            &pclsObj, &uReturn);

        if (0 == uReturn)
        {
            break;
        }

        VARIANT vtProp;

        VariantInit(&vtProp);

        // ��������� ����� BIOS
        this->hRes = pclsObj->Get(L"Caption", 0, &vtProp, 0, 0);
        this->caption = vtProp.bstrVal;
        VariantClear(&vtProp);

        // ��������� ������ ��������� ��� BIOS �������. 
        this->hRes = pclsObj->Get(L"BIOSVersion", 0, &vtProp, 0, 0);
        this->hRes = SafeArrayLock(vtProp.parray);
        if (SUCCEEDED(this->hRes))
        {
            // ��������� ��������� ��������� � �����
            BSTR bstrBuff;
            long lowerBound, upperBound;
            SafeArrayGetLBound(vtProp.parray, 1, &lowerBound);
            SafeArrayGetUBound(vtProp.parray, 1, &upperBound);
            long BSTRArrLen = upperBound - lowerBound + 1;
            for (long i = 0; i < BSTRArrLen; i++)
            {
                SafeArrayGetElement(vtProp.parray, &i, &bstrBuff);
                this->biosVersion += bstrBuff;
                this->biosVersion += ' ';
            }
            SysFreeString(bstrBuff);
        }
        SafeArrayUnlock(vtProp.parray);
        VariantClear(&vtProp);

        // ��������� ���� ������� BIOS
        this->hRes = pclsObj->Get(L"ReleaseDate", 0, &vtProp, 0, 0);
        this->releaseDate = this->WMIDateStringToDate(vtProp.bstrVal);
        VariantClear(&vtProp);

        // ��������� ����� BIOS
        this->hRes = pclsObj->Get(L"Status", 0, &vtProp, 0, 0);
        this->status = vtProp.bstrVal;
        VariantClear(&vtProp);

        // �������� ���������� SMBIOS
        this->hRes = pclsObj->Get(L"SMBIOSPresent", 0, &vtProp, 0, 0);
        if (vtProp.boolVal != 0)
            this->SMBIOSPresent = L"True";
        else
            this->SMBIOSPresent = L"False";
        VariantClear(&vtProp);

        // ��������� ���������� ��� SMBIOS �����
        this->hRes = pclsObj->Get(L"SMBIOSBIOSVersion", 0, &vtProp, 0, 0);
        this->SMBIOSVersion = vtProp.bstrVal;
        VariantClear(&vtProp);

        // ��������� �������� ���������� Bios Characteristics
        this->hRes = pclsObj->Get(L"BiosCharacteristics", 0, &vtProp, 0, 0);
        this->hRes = SafeArrayLock(vtProp.parray);
        if (SUCCEEDED(this->hRes))
        {
            // ��������� ��������� �� ��������������
            this->pCharactDat = static_cast<UINT32*>(vtProp.parray->pvData);

            // ��������� ��������� ������������� � �����
            long lowerBound, upperBound;
            SafeArrayGetLBound(vtProp.parray, 1, &lowerBound);
            SafeArrayGetUBound(vtProp.parray, 1, &upperBound);
            this->numOfCharact = upperBound - lowerBound + 1;
        }
        //SafeArrayUnlock(vtProp.parray);
        VariantClear(&vtProp);

        // ���������� TRUE, ���� ���� � ���� ��'��� ����
        // ������ ���������
        success = TRUE;
        pclsObj->Release();
    }
    // ��������� �������
    this->pSvc->Release();
    this->pLoc->Release();
    this->pEnumerator->Release();
    CoUninitialize();
    return success;
}

std::wstring BiosInfo::GetOutput()
{
    return this->result;
}

Boolean BiosInfoOutput::ConnectToWMI()
{
    Boolean isConnected = bInf.ConnectToWMI();
    this->result = marshal_as<String^>(bInf.GetOutput());

    return isConnected;
}

String^ BiosInfoOutput::GetBiosCharacteristics()
{
    String^ CLIOutput;

    for (int i = 0; i < bInf.numOfCharact; i++) {
        if (bInf.pCharactDat[i] <= 39) {
            switch (bInf.pCharactDat[i]) {
            case 0: CLIOutput += "00-Reserved \n"; break;
            case 1: CLIOutput += "01-Reserved \n"; break;
            case 2: CLIOutput += "02-Unknown \n"; break;
            case 3: CLIOutput += "03-BIOS Characteristics Not Supported \n"; break;
            case 4: CLIOutput += "04-ISA is supported \n"; break;
            case 5: CLIOutput += "05-MCA is supported \n"; break;
            case 6: CLIOutput += "06-EISA is supported \n"; break;
            case 7: CLIOutput += "07-PCI is supported \n"; break;
            case 8: CLIOutput += "08-PC Card (PCMCIA) is supported \n"; break;
            case 9: CLIOutput += "09-Plug and Play is supported \n"; break;
            case 10: CLIOutput += "10-APM is supported \n"; break;
            case 11: CLIOutput += "11-BIOS is Upgradable (Flash) \n"; break;
            case 12: CLIOutput += "12-BIOS shadowing is allowed \n"; break;
            case 13: CLIOutput += "13-VL-VESA is supported \n"; break;
            case 14: CLIOutput += "14-ESCD support is available \n"; break;
            case 15: CLIOutput += "15-Boot from CD is supported \n"; break;
            case 16: CLIOutput += "16-Selectable Boot is supported \n"; break;
            case 17: CLIOutput += "17-BIOS ROM is socketed \n"; break;
            case 18: CLIOutput += "18-Boot From PC Card (PCMCIA) is supported \n"; break;
            case 19: CLIOutput += "19-EDD (Enhanced Disk Drive) Specification is supported \n"; break;
            case 20: CLIOutput += "20-Int 13h - Japanese Floppy for NEC 9800 1.2mb (3.5, 1k Bytes/Sector, 360 RPM) is supported \n"; break;
            case 21: CLIOutput += "21-Int 13h - Japanese Floppy for Toshiba 1.2mb (3.5, 360 RPM) is supported \n"; break;
            case 22: CLIOutput += "22-Int 13h - 5.25 / 360 KB Floppy Services are supported \n"; break;
            case 23: CLIOutput += "23-Int 13h - 5.25 /1.2MB Floppy Services are supported \n"; break;
            case 24: CLIOutput += "24-Int 13h - 3.5 / 720 KB Floppy Services are supported \n"; break;
            case 25: CLIOutput += "25-Int 13h - 3.5 / 2.88 MB Floppy Services are supported \n"; break;
            case 26: CLIOutput += "26-Int 5h, Print Screen Service is supported \n"; break;
            case 27: CLIOutput += "27-Int 9h, 8042 Keyboard services are supported \n"; break;
            case 28: CLIOutput += "28-Int 14h, Serial Services are supported \n"; break;
            case 29: CLIOutput += "29-Int 17h, Printer services are supported \n"; break;
            case 30: CLIOutput += "30-Int 10h, CGA/Mono Video Services are supported \n"; break;
            case 31: CLIOutput += "31-NEC PC-98 \n"; break;
            case 32: CLIOutput += "32-ACPI supported \n"; break;
            case 33: CLIOutput += "33-USB Legacy is supported \n"; break;
            case 34: CLIOutput += "34-AGP is supported \n"; break;
            case 35: CLIOutput += "35-I2O boot is supported \n"; break;
            case 36: CLIOutput += "36-LS-120 boot is supported \n"; break;
            case 37: CLIOutput += "37-ATAPI ZIP Drive boot is supported \n"; break;
            case 38: CLIOutput += "38-1394 boot is supported \n"; break;
            case 39: CLIOutput += "39-Smart Battery supported \n"; break;
            default: CLIOutput += "";
            }
        }

        else if (bInf.pCharactDat[i] >= 40 && bInf.pCharactDat[i] <= 45) {
            CLIOutput += String::Format("{0}-Reserved for BIOS vendor \n", bInf.pCharactDat[i]);
        }

        else if (bInf.pCharactDat[i] >= 48 && bInf.pCharactDat[i] <= 63) {
            CLIOutput += String::Format("{0}-Reserved for system vendor \n", bInf.pCharactDat[i]);
        }

        else {
            CLIOutput += String::Format("{0}-Unknown Value \n", bInf.pCharactDat[i]);
        }
    }

    return CLIOutput;
}

String^ BiosInfoOutput::GetResult()
{
    return this->result;
}

Boolean BiosInfoOutput::GetBiosInfo()
{
    return bInf.GetBiosInfo();
}

String^ BiosInfoOutput::GetBiosCaption()
{
    return marshal_as<String^>(bInf.caption);
}

String^ BiosInfoOutput::GetBiosVersion()
{
    return marshal_as<String^>(bInf.biosVersion);
}

String^ BiosInfoOutput::GetBiosStatus()
{
    return marshal_as<String^>(bInf.status);
}

String^ BiosInfoOutput::GetBiosReleaseDate()
{
    return marshal_as<String^>(bInf.releaseDate);
}

String^ BiosInfoOutput::GetSMBIOSPresent()
{
    return marshal_as<String^>(bInf.SMBIOSPresent);
}

String^ BiosInfoOutput::GetSMBIOSVersion()
{
    return marshal_as<String^>(bInf.SMBIOSVersion);
}
