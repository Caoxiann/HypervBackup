#include <stdafx.h>
#include <Windows.h>
#include <tchar.h>
#include <Shlwapi.h>
#include <vss.h>
#include <vswriter.h>
#include <vsbackup.h>
#include <vsprov.h>
#include <iostream>
#include <vsserror.h>


#define LOG_BUFFER_SIZE  (4096 * 2)
#define LogDebug  mrlog
#define LogInfo   mrlog
#define LogWarn   mrlog
#define LogError  mrlog

using namespace std;
void ReleaseInterface(IUnknown* unkn)
{
	if (unkn)
	{
		unkn->Release();
		unkn = NULL;
	}
}

#define CLEAR_VSS(pPrepare, pDoShadowCopy) \
	ReleaseInterface((pPrepare));	\
	ReleaseInterface((pDoShadowCopy));	\

#define CLEAR_META_DATA(pBackup, bFreeMetaData)	\
	if(bFreeMetaData)	\
		(*pBackup)->FreeWriterMetadata();	\

#define BUFFER_SIZE     (4096)

#pragma comment (lib, "ole32.lib")
#pragma comment (lib, "VssApi.lib")
#pragma comment (lib, "Advapi32.lib")

// Helper macros to print a GUID using printf-style formatting
#define WSTR_GUID_FMT  _T("{%.8x-%.4x-%.4x-%.2x%.2x-%.2x%.2x%.2x%.2x%.2x%.2x}")

#define GUID_PRINTF_ARG( X )                                \
    (X).Data1,                                              \
    (X).Data2,                                              \
    (X).Data3,                                              \
    (X).Data4[0], (X).Data4[1], (X).Data4[2], (X).Data4[3], \
    (X).Data4[4], (X).Data4[5], (X).Data4[6], (X).Data4[7]


void mrlog(const TCHAR* format, ...)
{
	TCHAR szLogBuf[LOG_BUFFER_SIZE] = { 0 };
	va_list arg_ptr;
	va_start(arg_ptr, format);
	_vsntprintf_s(szLogBuf, sizeof(szLogBuf) / sizeof(szLogBuf[0]), format, arg_ptr);
	va_end(arg_ptr);
	_tprintf(_T("%s\n"), szLogBuf);
}

BOOL CreateSnapshot(_In_ IVssBackupComponents* pBackup, _In_ const TCHAR* szVolumeName)
{
	if (!pBackup)
	{
		LogError(_T("[CreateSnapshot]Invalid param"));
	}

	HRESULT hResult = S_OK;
	BOOL bRetVal = TRUE;
	VSS_ID snapShotId = { 0 };
	IVssAsync* pPrepare = NULL;
	IVssAsync* pDoShadowCopy = NULL;
	VSS_SNAPSHOT_PROP snapshotProp = { 0 };

	hResult = pBackup->AddToSnapshotSet(const_cast<TCHAR*>(szVolumeName), GUID_NULL, &snapShotId);
	if (hResult != S_OK)
	{
		wcout << szVolumeName << endl;
		LogError(_T("AddToSnapshotSet failed, code=0x%x"), hResult);
		CLEAR_VSS(pPrepare, pDoShadowCopy);
		bRetVal = FALSE;
		return bRetVal;
	}
	hResult = pBackup->SetBackupState(false, false, VSS_BT_FULL);
	if (hResult != S_OK)
	{
		LogError(_T("SetBackupState failed, code=0x%x"), hResult);
		CLEAR_VSS(pPrepare, pDoShadowCopy);
		bRetVal = FALSE;
		return bRetVal;
	}

	hResult = pBackup->PrepareForBackup(&pPrepare);
	if (hResult != S_OK)
	{
		LogError(_T("PrepareForBackup failed, code=0x%x"), hResult);
		CLEAR_VSS(pPrepare, pDoShadowCopy);
		bRetVal = FALSE;
		return bRetVal;
	}

	LogInfo(_T("Prepare for backup"));
	hResult = pPrepare->Wait();
	if (hResult != S_OK)
	{
		LogError(_T("IVssAsync Wait failed, code=0x%x"), hResult);
		CLEAR_VSS(pPrepare, pDoShadowCopy);
		bRetVal = FALSE;
		return bRetVal;
	}

	hResult = pBackup->DoSnapshotSet(&pDoShadowCopy);
	if (hResult != S_OK)
	{
		LogError(_T("DoSnapShotSet failed, code=0x%x"), hResult);
		CLEAR_VSS(pPrepare, pDoShadowCopy);
		bRetVal = FALSE;
		return bRetVal;
	}

	LogInfo(_T("Taking snapshots..."));
	hResult = pDoShadowCopy->Wait();
	if (hResult != S_OK)
	{
		LogError(_T("IVssasync Wait failed, code=0x%x"), hResult);
		CLEAR_VSS(pPrepare, pDoShadowCopy);
		bRetVal = FALSE;
		return bRetVal;
	}

	LogInfo(_T("Get the snapshot device object from the properties..."));

	hResult = pBackup->GetSnapshotProperties(snapShotId, &snapshotProp);
	if (hResult != S_OK)
	{
		LogError(_T("GetSnapShotProperties failed, code=0x%x"), hResult);
		CLEAR_VSS(pPrepare, pDoShadowCopy);
		bRetVal = FALSE;
		return bRetVal;
	}

	LogDebug(_T(" Snapshot ID:") WSTR_GUID_FMT, GUID_PRINTF_ARG(snapshotProp.m_SnapshotId));
	LogDebug(_T(" Snapshot Set ID"), WSTR_GUID_FMT, GUID_PRINTF_ARG(snapshotProp.m_SnapshotSetId));
	LogDebug(_T(" Provider ID ") WSTR_GUID_FMT, GUID_PRINTF_ARG(snapshotProp.m_ProviderId));
	LogDebug(_T(" OriginalVolumeName: %ls"), snapshotProp.m_pwszOriginalVolumeName);

	if (snapshotProp.m_pwszExposedName)
	{
		LogDebug(_T(" ExposedName: %ls"), snapshotProp.m_pwszExposedName);
	}
	if (snapshotProp.m_pwszExposedPath)
	{
		LogDebug(_T(" ExposedPath: %ls"), snapshotProp.m_pwszExposedPath);
	}
	if (snapshotProp.m_pwszSnapshotDeviceObject)
	{
		LogDebug(_T(" DeviceObject: %ls"), snapshotProp.m_pwszSnapshotDeviceObject);
	}

	VssFreeSnapshotProperties(&snapshotProp);
	bRetVal = TRUE;
	return bRetVal;
}

BOOL CreateSnapshotSet(_Out_ IVssBackupComponents** pBackup, _Out_ VSS_ID* snapshotSetId)
{
	if (!pBackup || !snapshotSetId)
	{
		LogError(_T("[CreateSnapshotSet]Invalid param"));
		return FALSE;
	}
	
	IVssAsync* pAsync = NULL;
	HRESULT hResult = S_OK;
	BOOL bRetVal = TRUE;
	BOOL bFreeMetaData = FALSE;
	
	hResult = CoInitialize(NULL);
	if (hResult != S_OK)
	{
		LogError(_T("CoInitialize failed, code=0x%x"), hResult);
		return FALSE;
	}

	hResult = CreateVssBackupComponents(pBackup);
	if (hResult != S_OK)
	{
		LogError(_T("CreateVssBackupComponents failed, code=0x%x"), hResult);
		return FALSE;
	}

	hResult = (*pBackup)->InitializeForBackup();
	if (hResult != S_OK)
	{
		LogError(_T("InitializeForBackup failed, code=0x%x"), hResult);
		bRetVal = FALSE;
		CLEAR_META_DATA(pBackup, bFreeMetaData);
		return FALSE;
	}

	hResult = (*pBackup)->SetContext(VSS_CTX_BACKUP);
	if (hResult != S_OK)
	{
		LogError(_T("IVssBackupComponents SetContext failed, code=0x%x"), hResult);
		bRetVal = FALSE;
		CLEAR_META_DATA(pBackup, bFreeMetaData);
		return FALSE;
	}

	hResult = (*pBackup)->GatherWriterMetadata(&pAsync);
	bFreeMetaData = TRUE;
	if (hResult != S_OK)
	{
		LogError(_T("GatherWriterMetaData failed, code=0x%x"), hResult);
		bRetVal = FALSE;
		CLEAR_META_DATA(pBackup, bFreeMetaData);
		return FALSE;
	}

	hResult = pAsync->Wait();
	if (hResult != S_OK)
	{
		LogError(_T("IVssAsync Wait failed, code=0x%x"), hResult);
		bRetVal = FALSE;
		CLEAR_META_DATA(pBackup, bFreeMetaData);
		return FALSE;
	}

	hResult = (*pBackup)->StartSnapshotSet(snapshotSetId);
	if (hResult != S_OK)
	{
		LogError(_T("StartSnapshotSet failed, code=0x%x"), hResult);
		bRetVal = FALSE;
		CLEAR_META_DATA(pBackup, bFreeMetaData);
		return FALSE;
	}
	
	bRetVal = TRUE;
	CLEAR_META_DATA(pBackup, bFreeMetaData);
	return bRetVal;
}

BOOL VolumeShadow(_In_ const TCHAR* szVolumeName)
{
	if (!szVolumeName)
	{
		LogError(_T("[CopyVolume]Invalid param"));
		return FALSE;
	}

	VSS_ID snapshotSetID = { 0 };
	IVssBackupComponents* pBackup = NULL;
	BOOL bRetVal = TRUE;
	if(!CreateSnapshotSet(&pBackup, &snapshotSetID))
	{
		return FALSE;
	}

	if (!CreateSnapshot(pBackup, szVolumeName))
	{
		LogError(_T("CreateSnapshot failed"));
		bRetVal = FALSE;
	}

	ReleaseInterface(pBackup);
	return bRetVal;
}

int _tmain(int argc, const TCHAR* argv[])
{
	printf("argc:%d, argv:%s\n", argc, argv[1]);
	wcout << argv[1] << endl;
	if (argc != 2)
	{
		LogError(_T("Usage: %s, Usage %s volumeName. eg. %s C\\"), __FILE__, __FILE__);
		char t;
		std::cin >> t;
		return 1;
	}
	bool res = VolumeShadow(argv[1]);
	if (!res)
	{
		LogError(_T("Error happens when backup."));
	}
	else
	{
		LogInfo(_T("Backup success."));
	}
	char k;
	LogInfo(_T("Please enter any key."));
	std::cin >> k;
	return 0;
}