#include "BiosInfo.h"

BOOL BiosInfo::ConnectToWMI()
{
    // ���� 1:***************************************************
    // ����������� COM.****************************************

    this->hRes = CoInitializeEx(0, COINIT_APARTMENTTHREADED);
    if (FAILED(this->hRes)) {
        // ��������� ���� ������� HRESULT
        this->errorCode = HRESULT_CODE(this->hRes);
        this->output = L"������� ����������� �������� COM. ";
        // ���������� ���� �������
        wsprintfW(errBuff, L"��� �������: 0x%X\n", errorCode);
        this->output += errBuff;
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
        this->output = L"�� ������� ������������ �������. ";
        wsprintfW(errBuff, L"��� �������: 0x%X\n", errorCode);
        this->output += errBuff;

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
        this->output = L"�� ������� �������� ��'��� IWbemLocator. ";
        wsprintfW(errBuff, L"��� �������: 0x%X\n", errorCode);
        this->output += errBuff;

        CoUninitialize();
        return FALSE;
    }

    // ���� 4:***************************************************
    // ϳ��������� �� WMI ����� IWbemLocator::ConnectServer �����

    this->pSvc = NULL;

    // ϳ��������� �� �������� ���� Root\Cimv2
    // ��������� ����������� � ��������� ��������� pSvc

    this->hRes = pLoc->ConnectServer(
        BSTR(L"ROOT\\CIMV2"),        // ���� �� ��'���� WMI
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
        this->output = L"�� ������� �����������. ";
        wsprintfW(errBuff, L"��� �������: 0x%X\n", errorCode);
        this->output += errBuff;

        CoUninitialize();
        return FALSE;
    }

    this->output = L"ϳ�������� �� �������� ���� Root\Cimv2 ���������� WMI \n";

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
        this->output = L"�� ������� ������������ ��� ������� ��� WMI-�'�������. ";
        wsprintfW(errBuff, L"��� �������: 0x%X\n", errorCode);
        this->output += errBuff;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return FALSE;
    }

    return TRUE;
}

BOOL BiosInfo::GetTestData()
{
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

        this->output = L"Query for operating system name failed.";
        wsprintfW(this->errBuff, L"Erroe code: 0x%X\n", errorCode);
        this->output += this->errBuff;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return FALSE;
    }

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;

    while (pEnumerator)
    {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
            &pclsObj, &uReturn);

        if (0 == uReturn)
        {
            break;
        }

        VARIANT vtProp;

        VariantInit(&vtProp);

        hr = pclsObj->Get(L"Manufacturer", 0, &vtProp, 0, 0);
        this->output = L" Bios Manufacturer : ";
        this->output += vtProp.bstrVal;
        this->output += '\n';
        VariantClear(&vtProp);

        pclsObj->Release();
    }
    return TRUE;
}

std::wstring BiosInfo::GetOutput()
{
    return this->output;
}

String^ BiosInfoOutput::GetOutput()
{
    BiosInfo bInf;
    Boolean isConnected = bInf.ConnectToWMI();
    Boolean isGot = bInf.GetTestData();
    std::wstring res = bInf.GetOutput();

    // ������������ std::wstring �� System::String
    String^ CLIOutput = marshal_as<String^>(res);
   
    return CLIOutput;
}
