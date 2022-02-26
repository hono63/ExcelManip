// ExcelManip.cpp : このファイルには 'main' 関数が含まれています。プログラム実行の開始と終了がそこで行われます。
//

#include <Ole2.h>
#include <atlstr.h>

#include <iostream>
#include <stdio.h>

/// <summary>
/// Excelを読む
/// 参考： https://docs.microsoft.com/ja-JP/previous-versions/office/troubleshoot/office-developer/automate-excel-from-c
/// </summary>
/// <param name="autoType"></param>
/// <param name="pvRsult"></param>
/// <param name="pDisp"></param>
/// <param name="ptName"></param>
/// <param name="..."></param>
/// <returns></returns>
HRESULT AutoWrap(int autoType, VARIANT* pvResult, IDispatch* pDisp, LPCTSTR tcName, int cArgs...) 
{
    va_list marker;
    va_start(marker, cArgs);

    if (!pDisp) {
        MessageBox(NULL, _T("NULL IDispatch passed to AutoWrap()"), _T("Error"), 0x10010);
        _exit(0);
    }

    // 変数たち
    DISPPARAMS dp = { NULL, NULL, 0, 0 };
    DISPID dispidNamed = DISPID_PROPERTYPUT;
    DISPID dispID;
    HRESULT hr;
    char szName[256];

    // (constじゃない) OLECHARへ変換
    OLECHAR oleName[256];
    _tcscpy_s<sizeof(oleName)/sizeof(OLECHAR)>(oleName, tcName);
    LPOLESTR ptName = oleName;
    // ANSIへ変換
    WideCharToMultiByte(CP_ACP, 0, ptName, -1, szName, sizeof(szName), NULL, NULL);

    // ptNameのDISPIDを取得
    hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, /*lcid*/1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        MessageBox(NULL, _T("Dispatch::GetIDsOfNames failed"), _T("AutoWrap()"), 0x10010);
        _exit(0);
        return hr;
    }

    // メモリ確保
    VARIANT* pArgs = new VARIANT[cArgs + 1];
    // 引数展開
    for (int i = 0; i < cArgs; i++) {
        pArgs[i] = va_arg(marker, VARIANT);
    }

    // Build DISPPARAMS
    dp.cArgs = cArgs;
    dp.rgvarg = pArgs;

    // 特殊ケース property-puts
    if (autoType & DISPATCH_PROPERTYPUT) {
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
    }

    // call
    hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, NULL, NULL);
    if (FAILED(hr)) {
        TCHAR buf[1024];
        _stprintf_s(buf, sizeof(buf)/sizeof(TCHAR), _T("IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx"), ptName, dispID, hr);
        MessageBox(NULL, buf, _T("AutoWrap()"), 0x10010);
        _exit(0);
        return hr;
    }
    // End variable-argument section...
    va_end(marker);

    delete[] pArgs;
    return hr;
}

int sample()
{
    // COM初期化
    HRESULT hr = CoInitialize(NULL);

    // サーバーのCLSID取得
    CLSID clsid;
    hr = CLSIDFromProgID(_T("Excel.Application"), &clsid);
    if (FAILED(hr)) {
        MessageBox(NULL, _T("CLSIDFromProgID() failed"), _T("Error"), 0x10010);
        return -1;
    }

    // サーバーstart & IDispatch取得
    IDispatch* pXlApp;
    hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&pXlApp);
    if (FAILED(hr)) {
        ::MessageBox(NULL, _T("Excel not registered properly"), _T("Error"), 0x10010);
        return -2;
    }

    // 可視化
    {
        VARIANT x;
        x.vt = VT_I4;
        x.lVal = 1;
        AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXlApp, _T("Visible"), 1, x);
    }

    // Workbookたち取得
    IDispatch* pXlBooks;
    {
        VARIANT result;
        VariantInit(&result);
        AutoWrap(DISPATCH_PROPERTYGET, &result, pXlApp, _T("Workbooks"), 0);
        pXlBooks = result.pdispVal;
    }

    // 新しいWorkbook
    IDispatch* pXlBook1;
    {
        VARIANT result;
        VariantInit(&result);
        AutoWrap(DISPATCH_PROPERTYGET, &result, pXlBooks, _T("Add"), 0);
        pXlBook1 = result.pdispVal;
    }

    // Create a 15x15 safearray of variants...
    VARIANT arr;
    arr.vt = VT_ARRAY | VT_VARIANT;
    {
        SAFEARRAYBOUND sab[2];
        sab[0].lLbound = 1; sab[0].cElements = 15;
        sab[1].lLbound = 1; sab[1].cElements = 15;
        arr.parray = SafeArrayCreate(VT_VARIANT, 2, sab);
    }

    // Fill safearray with some values...
    for (int i = 1; i <= 15; i++) {
        for (int j = 1; j <= 15; j++) {
            // Create entry value for (i,j)
            VARIANT tmp;
            tmp.vt = VT_I4;
            tmp.lVal = i * j;
            // Add to safearray...
            long indices[] = { i,j };
            hr = SafeArrayPutElement(arr.parray, indices, (void*)&tmp);
        }
    }

    // Get ActiveSheet object
    IDispatch* pXlSheet;
    {
        VARIANT result;
        VariantInit(&result);
        AutoWrap(DISPATCH_PROPERTYGET, &result, pXlApp, _T("ActiveSheet"), 0);
        pXlSheet = result.pdispVal;
    }

    // Get Range object for the Range A1:O15...
    IDispatch* pXlRange;
    {
        VARIANT parm;
        parm.vt = VT_BSTR;
        parm.bstrVal = ::SysAllocString(L"A1:O15");

        VARIANT result;
        VariantInit(&result);
        AutoWrap(DISPATCH_PROPERTYGET, &result, pXlSheet, _T("Range"), 1, parm);
        VariantClear(&parm);

        pXlRange = result.pdispVal;
    }

    // Set range with our safearray...
    AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXlRange, _T("Value"), 1, arr);

    // Wait for user...
    ::MessageBox(NULL, _T("All done."), _T("Notice"), 0x10000);

    // Set .Saved property of workbook to TRUE so we aren't prompted
    // to save when we tell Excel to quit...
    {
        VARIANT x;
        x.vt = VT_I4;
        x.lVal = 1;
        AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXlBook1, _T("Saved"), 1, x);
    }

    // Tell Excel to quit (i.e. App.Quit)
    AutoWrap(DISPATCH_METHOD, NULL, pXlApp, _T("Quit"), 0);

    // Release references...
    pXlRange->Release();
    pXlSheet->Release();
    pXlBook1->Release();
    pXlBooks->Release();
    pXlApp->Release();
    VariantClear(&arr);

    CoUninitialize();
}

#include <string>
using namespace std;

class CExcelManip {
private:
    IDispatch* mXlApp = NULL;
    IDispatch* mAllXlBooks = NULL;
    IDispatch* mXlBook = NULL;

public:
    CExcelManip() {
        HRESULT hr = CoInitialize(NULL);

        // サーバーのCLSID取得
        CLSID clsid;
        hr = CLSIDFromProgID(_T("Excel.Application"), &clsid);
        if (FAILED(hr)) {
            MessageBox(NULL, _T("CLSIDFromProgID() failed"), _T("Error"), 0x10010);
            exit(0);
        }

        // サーバーstart & IDispatch取得
        hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&mXlApp);
        if (FAILED(hr)) {
            ::MessageBox(NULL, _T("Excel not registered properly"), _T("Error"), 0x10010);
            exit(0);
        }

        // Workbookたち取得
        {
            VARIANT result;
            VariantInit(&result);
            AutoWrap(DISPATCH_PROPERTYGET, &result, mXlApp, _T("Workbooks"), 0);
            mAllXlBooks = result.pdispVal;
        }
    }

    ~CExcelManip() {
        // Excel終了
        if (mXlApp)
            AutoWrap(DISPATCH_METHOD, NULL, mXlApp, _T("Quit"), 0);

        // Release references...
        if (mAllXlBooks) mAllXlBooks->Release();
        if (mXlApp) mXlApp->Release();
        if (mXlBook) mXlBook->Release();
        
        CoUninitialize();
    }

    void Open(LPCTSTR filename) {
        VARIANT param;
        VARIANT result;
        param.vt = VT_BSTR;
        param.bstrVal = ::SysAllocString(filename);
        VariantInit(&result);
        AutoWrap(DISPATCH_PROPERTYGET, &result, mAllXlBooks, _T("Open"), 1, param);
        mXlBook = result.pdispVal;
    }

    void View() {
        // 可視化
        VARIANT x;
        x.vt = VT_I4;
        x.lVal = 1;
        AutoWrap(DISPATCH_PROPERTYPUT, NULL, mXlApp, _T("Visible"), 1, x);
    }
};

int main()
{
    std::cout << "Hello World!\n";

    //sample();

    CExcelManip manip;
    manip.Open(_T("C:\\Users\\nhodo\\Documents\\Sample.xlsx"));
    std::cout << "Good morning World!\n";
    int a = getchar();
    manip.View();
    std::cout << "Good night World!\n";
    a = getchar();

    return 0;
}

// プログラムの実行: Ctrl + F5 または [デバッグ] > [デバッグなしで開始] メニュー
// プログラムのデバッグ: F5 または [デバッグ] > [デバッグの開始] メニュー

// 作業を開始するためのヒント: 
//   1. ソリューション エクスプローラー ウィンドウを使用してファイルを追加/管理します 
//   2. チーム エクスプローラー ウィンドウを使用してソース管理に接続します
//   3. 出力ウィンドウを使用して、ビルド出力とその他のメッセージを表示します
//   4. エラー一覧ウィンドウを使用してエラーを表示します
//   5. [プロジェクト] > [新しい項目の追加] と移動して新しいコード ファイルを作成するか、[プロジェクト] > [既存の項目の追加] と移動して既存のコード ファイルをプロジェクトに追加します
//   6. 後ほどこのプロジェクトを再び開く場合、[ファイル] > [開く] > [プロジェクト] と移動して .sln ファイルを選択します
