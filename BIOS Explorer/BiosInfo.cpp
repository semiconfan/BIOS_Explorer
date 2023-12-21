#include "BiosInfo.h"

BOOL BiosInfo::ConnectToWMI()
{
    // Крок 1:***************************************************
    // Ініціалізація COM.****************************************

    this->hRes = CoInitializeEx(0, COINIT_APARTMENTTHREADED);
    if (FAILED(this->hRes)) {
        // Отримання коду статусу HRESULT
        this->errorCode = HRESULT_CODE(this->hRes);
        this->output = L"Помилка ініціалізації бібліотеки COM. ";
        // Збереження коду помилки
        wsprintfW(errBuff, L"Код помилки: 0x%X\n", errorCode);
        this->output += errBuff;
        return FALSE;                   // Підрограма завершилася невдало
    }

    // Крок 2:***************************************************
    // Встановлення загальних рівнів безпеки COM ****************

    this->hRes = CoInitializeSecurity(
        NULL,
        -1,                             // Аутентифікація COM
        NULL,                           // Служби автентифікації
        NULL,                           // Зарезервовано
        RPC_C_AUTHN_LEVEL_DEFAULT,      // Автентифікація за замовчуванням
        RPC_C_IMP_LEVEL_IDENTIFY,       // Уособлення за замовчуванням
        NULL,                           // Інформація про автентифікацію
        EOAC_NONE,                      // Додаткові можливості
        NULL                            // Зарезервовано
    );

    if (FAILED(this->hRes)) {
        this->errorCode = HRESULT_CODE(this->hRes);
        this->output = L"Не вдалося ініціалізувати безпеку. ";
        wsprintfW(errBuff, L"Код помилки: 0x%X\n", errorCode);
        this->output += errBuff;

        CoUninitialize();
        return FALSE;
    }

    // Крок 3:***************************************************
    // Отримання початкового локатора до WMI ********************

    this->pLoc = NULL;

    this->hRes = CoCreateInstance(
        CLSID_WbemLocator,
        0,
        CLSCTX_INPROC_SERVER,
        IID_IWbemLocator,
        (LPVOID*)&this->pLoc);

    if (FAILED(this->hRes)) {
        this->errorCode = HRESULT_CODE(this->hRes);
        this->output = L"Не вдалося створити об'єкт IWbemLocator. ";
        wsprintfW(errBuff, L"Код помилки: 0x%X\n", errorCode);
        this->output += errBuff;

        CoUninitialize();
        return FALSE;
    }

    // Крок 4:***************************************************
    // Підключення до WMI через IWbemLocator::ConnectServer метод

    this->pSvc = NULL;

    // Підключення до простору імен Root\Cimv2
    // поточного користувача й одержання покажчика pSvc

    this->hRes = pLoc->ConnectServer(
        BSTR(L"ROOT\\CIMV2"),        // Шлях до об'єкта WMI
        NULL,                           // Ім'я користувача. NULL = поточний користувач
        NULL,                           // Пароль користувача. NULL = поточний пароль
        0,
        NULL,
        0,
        0,
        &this->pSvc                     // Покажчик IWbemServices proxy
    );

    if (FAILED(this->hRes)) {
        this->errorCode = HRESULT_CODE(this->hRes);
        this->output = L"Не вдалося підключитися. ";
        wsprintfW(errBuff, L"Код помилки: 0x%X\n", errorCode);
        this->output += errBuff;

        CoUninitialize();
        return FALSE;
    }

    this->output = L"Підключено по простору імен Root\Cimv2 репозиторію WMI \n";

    // Крок 5:***************************************************
    // Налаштування рівнів безпеки для WMI-з'єднання ************

    this->hRes = CoSetProxyBlanket(
        this->pSvc,                     // Вказує на проксі, який потрібно встановити
        RPC_C_AUTHN_WINNT,              // RPC_C_AUTHN_xxx
        RPC_C_AUTHZ_NONE,               // RPC_C_AUTHZ_xxx
        NULL,                           // Основне ім'я сервера
        RPC_C_AUTHN_LEVEL_CALL,         // RPC_C_AUTHN_LEVEL_xxx
        RPC_C_IMP_LEVEL_IMPERSONATE,    // RPC_C_IMP_LEVEL_xxx
        NULL,                           // Ідентифікація клієнта
        EOAC_NONE                       // Можливості проксі
    );

    if (FAILED(this->hRes)) {
        this->errorCode = HRESULT_CODE(this->hRes);
        this->output = L"Не вдалося налаштуванти рівні безпеки для WMI-з'єднання. ";
        wsprintfW(errBuff, L"Код помилки: 0x%X\n", errorCode);
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

    // Перетворення std::wstring на System::String
    String^ CLIOutput = marshal_as<String^>(res);
   
    return CLIOutput;
}
