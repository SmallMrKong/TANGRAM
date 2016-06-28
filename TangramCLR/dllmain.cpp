/********************************************************************************
*					Tangram Library - version 8.0								*
*********************************************************************************
* Copyright (C) 2002-2016 by Tangram Team.   All Rights Reserved.				*
*
* THIS SOURCE FILE IS THE PROPERTY OF TANGRAM TEAM AND IS NOT TO
* BE RE-DISTRIBUTED BY ANY MEANS WHATSOEVER WITHOUT THE EXPRESSED
* WRITTEN CONSENT OF TANGRAM TEAM.
*
* THIS SOURCE CODE CAN ONLY BE USED UNDER THE TERMS AND CONDITIONS
* OUTLINED IN THE GPL LICENSE AGREEMENT.TANGRAM TEAM
* GRANTS TO YOU (ONE SOFTWARE DEVELOPER) THE LIMITED RIGHT TO USE
* THIS SOFTWARE ON A SINGLE COMPUTER.
*
* CONTACT INFORMATION:
* mailto:sunhuizlz@yeah.net, mailto:sunhui@cloudaddin.com
* http://www.CloudAddin.com
*
*
********************************************************************************/

// dllmain.cpp : Implementation of DllMain.

#include "stdafx.h"
#include "resource.h"
#include "dllmain.h"
#include "TangramCoreEvents.cpp"

CCloudAddinCLRApp theApp;

#ifdef _WIN64
vector<CString> adsPropertys = {
	CString("DN"),
	CString("objectClass"),
	CString("distinguishedName"),
	CString("instanceType"),
	CString("whenCreated"),
	CString("whenChanged"),
	CString("subRefs"),
	CString("uSNCreated"),
	CString("uSNChanged"),
	CString("name"),
	CString("objectGUID"),
	CString("creationTime"),
	CString("forceLogoff"),
	CString("lockoutDuration"),
	CString("lockOutObservationWindow"),
	CString("lockoutThreshold"),
	CString("maxPwdAge"),
	CString("minPwdAge"),
	CString("minPwdLength"),
	CString("modifiedCountAtLastProm"),
	CString("nextRid"),
	CString("pwdProperties"),
	CString("pwdHistoryLength"),
	CString("objectSid"),
	CString("serverState"),
	CString("uASCompat"),
	CString("modifiedCount"),
	CString("auditingPolicy"),
	CString("nTMixedDomain"),
	CString("rIDManagerReference"),
	CString("fSMORoleOwner"),
	CString("systemFlags"),
	CString("wellKnownObjects"),
	CString("objectCategory"),
	CString("isCriticalSystemObject"),
	CString("gPLink"),
	CString("dSCorePropagationData"),
	CString("otherWellKnownObjects"),
	CString("masteredBy"),
	CString("ms-DS-MachineAccountQuota"),
	CString("msDS-Behavior-Version"),
	CString("msDS-PerUserTrustQuota"),
	CString("msDS-AllUsersTrustQuota"),
	CString("msDS-PerUserTrustTombstonesQuota"),
	CString("msDs-masteredBy"),
	CString("msDS-IsDomainFor"),
	CString("msDS-NcType"),
	CString("dc"),
	CString("cn"),
	CString("description"),
	CString("showInAdvancedViewOnly"),
	CString("ou"),
	CString("msDS-TombstoneQuotaFactor"),
	CString("displayName"),
	CString("flags"),
	CString("versionNumber"),
	CString("gPCFunctionalityVersion"),
	CString("gPCFileSysPath"),
	CString("gPCMachineExtensionNames"),
	CString("ipsecName"),
	CString("ipsecID"),
	CString("ipsecDataType"),
	CString("ipsecData"),
	CString("ipsecISAKMPReference"),
	CString("ipsecNFAReference"),
	CString("ipsecOwnersReference"),
	CString("ipsecNegotiationPolicyReference"),
	CString("ipsecFilterReference"),
	CString("iPSECNegotiationPolicyType"),
	CString("iPSECNegotiationPolicyAction"),
	CString("revision"),
	CString("memberOf"),
	CString("userAccountControl"),
	CString("badPwdCount"),
	CString("codePage"),
	CString("countryCode"),
	CString("badPasswordTime"),
	CString("lastLogoff"),
	CString("lastLogon"),
	CString("logonHours"),
	CString("pwdLastSet"),
	CString("primaryGroupID"),
	CString("adminCount"),
	CString("accountExpires"),
	CString("logonCount"),
	CString("sAMAccountName"),
	CString("sAMAccountType"),
	CString("lastLogonTimestamp"),
	CString("groupType"),
	CString("member"),
	CString("samDomainUpdates"),
	CString("localPolicyFlags"),
	CString("operatingSystem"),
	CString("operatingSystemVersion"),
	CString("serverReferenceBL"),
	CString("dNSHostName"),
	CString("rIDSetReferences"),
	CString("servicePrincipalName"),
	CString("msDS-SupportedEncryptionTypes"),
	CString("msDFSR-ComputerReferenceBL"),
	CString("rIDAvailablePool"),
	CString("rIDAllocationPool"),
	CString("rIDPreviousAllocationPool"),
	CString("rIDUsedPool"),
	CString("rIDNextRID"),
	CString("dnsRecord"),
	CString("msDFSR-Flags"),
	CString("msDFSR-ReplicationGroupType"),
	CString("msDFSR-FileFilter"),
	CString("msDFSR-DirectoryFilter"),
	CString("serverReference"),
	CString("msDFSR-ComputerReference"),
	CString("msDFSR-MemberReferenceBL"),
	CString("msDFSR-Version"),
	CString("msDFSR-ReplicationGroupGuid"),
	CString("msDFSR-MemberReference"),
	CString("msDFSR-RootPath"),
	CString("msDFSR-StagingPath"),
	CString("msDFSR-Enabled"),
	CString("msDFSR-Options"),
	CString("msDFSR-ContentSetGuid"),
	CString("msDFSR-ReadOnly"),
	CString("lastSetTime"),
	CString("priorSetTime"),
	CString("sn"),
	CString("givenName"),
	CString("userPrincipalName"),
};
#endif

CCloudAddinCLRApp::CCloudAddinCLRApp()
{
	ATLTRACE(_T("Loading CCloudAddinCLRApp :%p\n"), this);
	m_pTangram = nullptr;
	m_pCloudAddinCLREvent = nullptr;
	m_bHostObjCreated = false;
	m_nAppEndPointCount = 0;
	m_strAppEndpointsScript = _T("");
	InitializeCriticalSectionAndSpinCount(&m_csTaskRecycleSection, 0x00000400);
	InitializeCriticalSectionAndSpinCount(&m_csTaskListSection, 0x00000400);
}

CCloudAddinCLRApp::~CCloudAddinCLRApp()
{
	m_pCloudAddinCLREvent = nullptr;
	DeleteCriticalSection(&m_csTaskRecycleSection);
	DeleteCriticalSection(&m_csTaskListSection);
	ATLTRACE(_T("Release CCloudAddinCLRApp :%p\n"), this);
}

#include <wincrypt.h>

HCRYPTPROV hCryptProv;

BOOL CryptStartup()
{
	if (CryptAcquireContext(&hCryptProv, NULL, MS_ENHANCED_PROV, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) == 0) {
		if (GetLastError() == NTE_EXISTS) {
			if (CryptAcquireContext(&hCryptProv, NULL, MS_ENHANCED_PROV, PROV_RSA_FULL, 0) == 0)
				return FALSE;
		}
		else return FALSE;
	}
	return TRUE;
}

void CryptCleanup()
{
	if (hCryptProv) CryptReleaseContext(hCryptProv, 0);
	hCryptProv = 0;
}

typedef struct {
	unsigned char digest[16];
	unsigned long hHash;
} MD5Context;

void MD5Init(MD5Context *ctx)
{
	CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, (HCRYPTHASH *)&ctx->hHash);
}


void MD5Update(MD5Context *ctx, const unsigned char* buf, unsigned int len)
{
	CryptHashData(ctx->hHash, buf, len, 0);
}

void MD5Final(MD5Context *ctx)
{
	DWORD dwCount = 16;
	CryptGetHashParam(ctx->hHash, HP_HASHVAL, ctx->digest, &dwCount, 0);
	if (ctx->hHash) CryptDestroyHash(ctx->hHash);
	ctx->hHash = 0;
}

int CCloudAddinCLRApp::CalculateByteMD5(BYTE* pBuffer, int BufferSize, CString &MD5)
{
	int i, j;
	MD5Context md5Hash;

	unsigned char b;
	char c;
	if (!CryptStartup())
	{
		MessageBoxW(0, L"Could not start crypto library", L"MD5", MB_ICONERROR);
		return 0;
	}

	memset(&md5Hash, 0, sizeof(MD5Context));
	MD5Init(&md5Hash);
	MD5Update(&md5Hash, (const unsigned char *)pBuffer, BufferSize);
	MD5Final(&md5Hash);

	char *Value = new char[1024];
	int k = 0;
	for (i = 0; i < 16; i++)
	{
		b = md5Hash.digest[i];
		for (j = 4; j >= 0; j -= 4)
		{
			c = ((char)(b >> j) & 0x0F);
			if (c < 10) c += '0';
			else c = ('a' + (c - 10));
			Value[k] = c;
			k++;
		}
	}
	Value[k] = '\0';
	CryptCleanup();

	MD5 = CString(Value);

	delete Value;
	return 1;

}

CTangramNodeEvent::CTangramNodeEvent()
{
	m_pWndNode				= nullptr;
	m_pTangramNodeCLREvent	= nullptr;
}

CTangramNodeEvent::~CTangramNodeEvent()
{

}
