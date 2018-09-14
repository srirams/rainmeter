#define _CRT_SECURE_NO_WARNINGS
#include <windows.h>
#include <uxtheme.h>
#include "DeskBand.h"
#include <string>

#define RECTWIDTH(x)   ((x).right - (x).left)
#define RECTHEIGHT(x)  ((x).bottom - (x).top)

extern long        g_cDllRef;
extern HINSTANCE    g_hInst;

extern CLSID CLSID_RainTaskbar;

static const WCHAR g_szRainTaskbarHostClass[] = L"RainBarHostClass";

EXTERN_C __declspec(dllimport) HWND CreateTaskbarWindow(HWND parent);
EXTERN_C __declspec(dllimport) void DestroyTaskbarWindow();

EXTERN_C IMAGE_DOS_HEADER __ImageBase;


CDeskBand::CDeskBand() :
    m_cRef(1), m_pSite(NULL), m_fHasFocus(FALSE), m_fIsDirty(FALSE), m_dwBandID(0), m_hwnd(NULL), m_hwndParent(NULL)
{
	InterlockedIncrement(&g_cDllRef);
}

CDeskBand::~CDeskBand()
{
//	FunctionLogger a(L"~CDeskBand");
	InterlockedDecrement(&g_cDllRef);
    if (m_pSite)
    {
        m_pSite->Release();
    }
}

//
// IUnknown
//
STDMETHODIMP CDeskBand::QueryInterface(REFIID riid, void **ppv)
{
//	FunctionLogger a(L"QueryInterface");
    HRESULT hr = S_OK;

    if (IsEqualIID(IID_IUnknown, riid)       ||
        IsEqualIID(IID_IOleWindow, riid)     ||
        IsEqualIID(IID_IDockingWindow, riid) ||
        IsEqualIID(IID_IDeskBand, riid)      ||
        IsEqualIID(IID_IDeskBand2, riid))
    {
        *ppv = static_cast<IOleWindow *>(this);
    }
    else if (IsEqualIID(IID_IPersist, riid) ||
             IsEqualIID(IID_IPersistStream, riid))
    {
        *ppv = static_cast<IPersist *>(this);
    }
    else if (IsEqualIID(IID_IObjectWithSite, riid))
    {
        *ppv = static_cast<IObjectWithSite *>(this);
    }
    else if (IsEqualIID(IID_IInputObject, riid))
    {
        *ppv = static_cast<IInputObject *>(this);
    }
    else
    {
//		log(L"No Interface");
        hr = E_NOINTERFACE;
        *ppv = NULL;
    }

    if (*ppv)
    {
        AddRef();
    }

    return hr;
}

STDMETHODIMP_(ULONG) CDeskBand::AddRef()
{
//	FunctionLogger a(L"AddRef");
    return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CDeskBand::Release()
{
//	FunctionLogger a(L"Release");
    ULONG cRef = InterlockedDecrement(&m_cRef);
    if (0 == cRef)
    {
		// TODO: This is crashing....
		//DestroyTaskbarWindow();
        delete this;
    }

    return cRef;
}

//
// IOleWindow
//
STDMETHODIMP CDeskBand::GetWindow(HWND *phwnd)
{
//	FunctionLogger a(L"GetWindow");
    *phwnd = m_hwnd;
    return S_OK;
}

STDMETHODIMP CDeskBand::ContextSensitiveHelp(BOOL)
{
//	FunctionLogger a(L"ContextSensitiveHelp");
    return E_NOTIMPL;
}

//
// IDockingWindow
//
STDMETHODIMP CDeskBand::ShowDW(BOOL fShow)
{
//	FunctionLogger a(L"ShowDW");
    if (m_hwnd)
    {
        ShowWindow(m_hwnd, fShow ? SW_SHOW : SW_HIDE);
    }
    return S_OK;
}

STDMETHODIMP CDeskBand::CloseDW(DWORD)
{
//	FunctionLogger a(L"CloseDW");
    if (m_hwnd)
    {
        ShowWindow(m_hwnd, SW_HIDE);
        DestroyWindow(m_hwnd);
        m_hwnd = NULL;
    }

    return S_OK;
}

STDMETHODIMP CDeskBand::ResizeBorderDW(const RECT *, IUnknown *, BOOL)
{
//	FunctionLogger a(L"ResizeBorderDW");
    return E_NOTIMPL;
}

//
// IDeskBand
//
STDMETHODIMP CDeskBand::GetBandInfo(DWORD dwBandID, DWORD, DESKBANDINFO *pdbi)
{
//	FunctionLogger a(L"GetBandInfo");
    HRESULT hr = E_INVALIDARG;

    if (pdbi)
    {
        m_dwBandID = dwBandID;

        if (pdbi->dwMask & DBIM_MINSIZE)
        {
			RECT rt;
			GetClientRect(m_hwnd, &rt);
			pdbi->ptMinSize.x = rt.right;
			pdbi->ptMinSize.y = rt.bottom;
        }

        if (pdbi->dwMask & DBIM_MAXSIZE)
        {
            pdbi->ptMaxSize.y = -1;
        }

        if (pdbi->dwMask & DBIM_INTEGRAL)
        {
            pdbi->ptIntegral.y = 1;
        }

        if (pdbi->dwMask & DBIM_ACTUAL)
        {
			RECT rt;
			GetClientRect(m_hwnd, &rt);
			pdbi->ptMinSize.x = rt.right;
			pdbi->ptMinSize.y = rt.bottom;
		}

        if (pdbi->dwMask & DBIM_TITLE)
        {
            // Don't show title by removing this flag.
            pdbi->dwMask &= ~DBIM_TITLE;
        }

        if (pdbi->dwMask & DBIM_MODEFLAGS)
        {
            pdbi->dwModeFlags = DBIMF_NORMAL | DBIMF_VARIABLEHEIGHT;
        }

        if (pdbi->dwMask & DBIM_BKCOLOR)
        {
            // Use the default background color by removing this flag.
            pdbi->dwMask &= ~DBIM_BKCOLOR;
        }

        hr = S_OK;
    }

    return hr;
}

//
// IDeskBand2
//
STDMETHODIMP CDeskBand::CanRenderComposited(BOOL *pfCanRenderComposited)
{
//	FunctionLogger a(L"CanRenderComposited");
    *pfCanRenderComposited = TRUE;

    return S_OK;
}

STDMETHODIMP CDeskBand::SetCompositionState(BOOL fCompositionEnabled)
{
//	FunctionLogger a(L"SetCompositionState");
    m_fCompositionEnabled = fCompositionEnabled;

    InvalidateRect(m_hwnd, NULL, TRUE);
    UpdateWindow(m_hwnd);

    return S_OK;
}

STDMETHODIMP CDeskBand::GetCompositionState(BOOL *pfCompositionEnabled)
{
//	FunctionLogger a(L"GetCompositionState");
    *pfCompositionEnabled = m_fCompositionEnabled;

    return S_OK;
}

//
// IPersist
//
STDMETHODIMP CDeskBand::GetClassID(CLSID *pclsid)
{
//	FunctionLogger a(L"GetClassID");
    *pclsid = CLSID_RainTaskbar;
    return S_OK;
}

//
// IPersistStream
//
STDMETHODIMP CDeskBand::IsDirty()
{
//	FunctionLogger a(L"IsDirty");
    return m_fIsDirty ? S_OK : S_FALSE;
}

STDMETHODIMP CDeskBand::Load(IStream * /*pStm*/)
{
//	FunctionLogger a(L"Load");
    return S_OK;
}

STDMETHODIMP CDeskBand::Save(IStream * /*pStm*/, BOOL fClearDirty)
{
//	FunctionLogger a(L"Save");
    if (fClearDirty)
    {
        m_fIsDirty = FALSE;
    }

    return S_OK;
}

STDMETHODIMP CDeskBand::GetSizeMax(ULARGE_INTEGER * /*pcbSize*/)
{
//	FunctionLogger a(L"GetSizeMax");
    return E_NOTIMPL;
}

//
// IObjectWithSite
//
STDMETHODIMP CDeskBand::SetSite(IUnknown *pUnkSite)
{
//	FunctionLogger a(L"SetSite");
    HRESULT hr = S_OK;

    m_hwndParent = NULL;

    if (m_pSite)
    {
        m_pSite->Release();
    }

    if (pUnkSite)
    {
        IOleWindow *pow;
        hr = pUnkSite->QueryInterface(IID_IOleWindow, reinterpret_cast<void **>(&pow));
        if (SUCCEEDED(hr))
        {
            hr = pow->GetWindow(&m_hwndParent);
            if (SUCCEEDED(hr))
            {
				wchar_t DllPath[MAX_PATH] = {0};
				GetModuleFileNameW((HINSTANCE)&__ImageBase, DllPath, _countof(DllPath));
				wchar_t* loc = wcsrchr(DllPath, L'\\');

				if (loc)
				{
					*loc = L'\0';
					SetDllDirectory(DllPath);
				}

				m_hwnd = CreateTaskbarWindow(m_hwndParent);

                if (!m_hwnd)
                {
                    hr = E_FAIL;
                }
            }

            pow->Release();
        }

        hr = pUnkSite->QueryInterface(IID_IInputObjectSite, reinterpret_cast<void **>(&m_pSite));
    }

    return hr;
}

STDMETHODIMP CDeskBand::GetSite(REFIID riid, void **ppv)
{
//	FunctionLogger a(L"GetSite");
    HRESULT hr = E_FAIL;

    if (m_pSite)
    {
        hr =  m_pSite->QueryInterface(riid, ppv);
    }
    else
    {
        *ppv = NULL;
    }

    return hr;
}

//
// IInputObject
//
STDMETHODIMP CDeskBand::UIActivateIO(BOOL fActivate, MSG *)
{
//	FunctionLogger a(L"UIActivateIO");
    if (fActivate)
    {
        SetFocus(m_hwnd);
    }

    return S_OK;
}

STDMETHODIMP CDeskBand::HasFocusIO()
{
//	FunctionLogger a(L"HasFocus");
    return m_fHasFocus ? S_OK : S_FALSE;
}

STDMETHODIMP CDeskBand::TranslateAcceleratorIO(MSG *)
{
//	FunctionLogger a(L"TranslateAcceleratorIO");
    return S_FALSE;
};

