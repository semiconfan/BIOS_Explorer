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
// Призначення:
// Взаємодія з класами WMI для 
// дослідження області даних BIOS
class BiosInfo {
    // Дозволяє класу BiosInfoOutput отримувати 
    // доступ до приватних членів цього класу
    friend BiosInfoOutput;
private:
    // Змінні, які потрібні для виконання запитів до WMI
    HRESULT hRes;
    DWORD errorCode;

    IWbemLocator* pLoc;					// IWbemLocator-інтерфейс
    IWbemServices* pSvc;				// IWbemServices-інтерфейc
    IEnumWbemClassObject* pEnumerator;	// Змінна для задання запиту до WMI
    IWbemClassObject* pClsObj;			// Змінна для вилучення результату запиту WMI
    wchar_t errBuff[256]{};             // Буфер для зберігання коду помилки

    // Рядок, який потрібен для зберігання результатів запитів до WMI (С++ формат)
    std::wstring output;

    // Інформація про BIOS:
    std::wstring caption;               // Опис
    std::wstring biosVersion;           // Версія
    std::wstring manufacturer;          // Виробник
    std::wstring releaseDate;           // Дата випуску
    std::wstring SMBIOSVersion;         // Інформація про SMBIOS версію

    // Змінні, які потрібні для зберігання результатів запитів до WMI
    UINT32* pCharactDat;                // Масив характеристик BIOS

    // Методи для взаємодії з WMI
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