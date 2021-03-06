// HypervBackup.cpp: 定义控制台应用程序的入口点。
//

#include "stdafx.h"
#include <stdio.h>
#include <vss.h>
#include <vswriter.h>
#include <vsbackup.h>

typedef HRESULT(STDAPICALLTYPE* _CreateVssBackupComponentsInternal)(
	__out IVssBackupComponents **ppBackup
	);
typedef void (APIENTRY* _VssFreeSnapshotPropertiesInternal)(
	__in VSS_SNAPSHOT_PROP* pProp
	);

static _CreateVssBackupComponentsInternal CreateVssBackupComponentsInternal_I;
static _VssFreeSnapshotPropertiesInternal VssFreeSnapshotPropertiesInternal_I;


int main()
{
	HRESULT result;
	HMODULE vssapiBase;
	IVssBackupComponents* backupComponents;

	vssapiBase = LoadLibrary(L"vssapi.dll");

	if (vssapiBase)
	{
		CreateVssBackupComponentsInternal_I = (_CreateVssBackupComponentsInternal)GetProcAddress(vssapiBase, "CreateVssBackupComponentsInternal");
		VssFreeSnapshotPropertiesInternal_I = (_VssFreeSnapshotPropertiesInternal)GetProcAddress(vssapiBase, "CreateVssBackupComponentsInternal");
	}

	if (!CreateVssBackupComponentsInternal_I || !VssFreeSnapshotPropertiesInternal_I)
	{
		abort();
	}

	result = CreateVssBackupComponentsInternal_I(&backupComponents);
	if (!SUCCEEDED(result))
	{
		abort();
	}

	VSS_ID snapshotSetID;
	result = backupComponents->InitializeForBackup();
	if (!SUCCEEDED(result))
	{
		abort();
	}
	result = backupComponents->SetBackupState(FALSE, FALSE, VSS_BT_INCREMENTAL);
	if (!SUCCEEDED(result))
	{
		abort();
	}
	result = backupComponents->SetContext(VSS_CTX_FILE_SHARE_BACKUP);
	if (!SUCCEEDED(result))
	{
		abort();
	}

	backupComponents->StartSnapshotSet(&snapshotSetID);
	result = backupComponents->AddToSnapshotSet(const_cast<TCHAR *>(L"G:\Hyper-V\Virtual Hard Disks\\"), GUID_NULL, &snapshotSetID);
	if (!SUCCEEDED(result))
	{
		abort();
	}

	IVssAsync *async;
	result = backupComponents->DoSnapshotSet(&async);
	if (!SUCCEEDED(result))
	{
		abort();
	}
	result = async->Wait();
	async->Release();
	if (!SUCCEEDED(result))
	{
		abort();
	}

	VSS_SNAPSHOT_PROP prop;
	result = backupComponents->GetSnapshotProperties(snapshotSetID, &prop);
	VssFreeSnapshotPropertiesInternal_I(&prop);
	backupComponents->Release();
    return 0;
}
