// CloudAddinRestObj.cpp : CRestObj µÄÊµÏÖ

#include "stdafx.h"
#include "RestObj.h"

#pragma once

#include <codecvt>
#include <sstream>
#include <iostream>
#include "cpprest/filestream.h"
#include <cpprest/http_client.h>
#include <cpprest/ws_client.h>
#include <cpprest/rawptrstream.h>
#include <cpprest\http_listener.h>
#include <cpprest/containerstream.h>
#include <cpprest/producerconsumerstream.h>

using namespace pplx;
using namespace concurrency::streams;
using namespace utility;
using namespace utility::conversions;
using namespace web;
using namespace web::http;
using namespace web::http::client;
using namespace web::experimental::web_sockets;
using namespace web::http::oauth2::experimental;

// CRestObj

CRestObj::CRestObj()
{
	m_nState				= 0;
	m_strKey				= _T("");
	m_hNotifyWnd			= NULL;
	m_pRestNotify			= nullptr;
	m_pCloudAddinUINode		= nullptr;
}

CRestObj::~CRestObj()
{
	if (m_strKey != _T(""))
	{
		::SendMessage(m_hNotifyWnd, WM_REMOVERESTKEY, (WPARAM)m_strKey.GetBuffer(), 0);
		m_strKey.ReleaseBuffer();
	}
	m_hNotifyWnd	= NULL;
	m_pRestNotify	= nullptr;
}

STDMETHODIMP CRestObj::AsyncRest(int nMethod, BSTR bstrURL, BSTR bstrData, LONGLONG hNotify)
{
	auto t = create_task([nMethod, bstrURL, bstrData, hNotify,this]()
	{
		string_t strURL = to_string_t(LPCTSTR(bstrURL));
		LPCTSTR pszData = LPCTSTR(bstrData);
		if (::PathFileExists(pszData) && nMethod)
		{
			HWND hWnd = m_hNotifyWnd;
			if (hNotify)
			{
				hWnd = (HWND)hNotify;
			}

			string_t strFile = to_string_t(LPCTSTR(bstrData));
			if (nMethod)
			{
				auto _task = file_stream<unsigned char>::open_istream(strFile).then([this, strURL](task<streams::basic_istream<unsigned char>> previousTask)
				{
					auto fileStream = previousTask.get();
					// Make HTTP request with the file stream as the body.
					http_client client(strURL);
					http_request request(methods::POST);
					if (m_mapHeaders.size())
					{
						for (auto it : m_mapHeaders)
						{
							string_t strHeader = to_string_t(LPCTSTR(it.first));
							CString s = it.second;
							s.Replace(_T("\r\n"), _T(""));
							string_t strValue = to_string_t(LPCTSTR(s));
							request.headers().add(strHeader, strValue);
						}
						m_mapHeaders.erase(m_mapHeaders.begin(), m_mapHeaders.end());
					}

					request.set_body(fileStream, U("application/octet-stream"));
					client.request(request).then([this, strURL, fileStream](task<http_response> previousTask)
					{
						fileStream.close();
						auto response = previousTask.get();
						streams::istream bodyStream = response.body();
						container_buffer<std::string> inStringBuffer;
						bodyStream.read_to_end(inStringBuffer).then([inStringBuffer,this](task<size_t> previousTask)
						{
							std::string &strRet = inStringBuffer.collection();
							//CoInitializeEx(NULL, COINIT_MULTITHREADED);
							std::wstring &strRet2 = to_utf16string(strRet);
							HWND hWnd = m_hNotifyWnd;
							if (::IsWindow(hWnd))
							{
								::SendMessage(hWnd, WM_UPLOADFILE, (WPARAM)strRet2.c_str(), 0);
							}
							CComBSTR bstrRet(strRet2.c_str());
							if (m_pRestNotify != nullptr)
							{
								m_pRestNotify->Notify(bstrRet);
							}
							delete this;
							//CoUninitialize();
						});
					});
				});
				try
				{
					_task.wait();
					return S_OK;
				}
				catch (const http_exception& e)
				{
					printf("Error http_exception:%s\n", e.what());
					delete this;
					return S_OK;
				}
				catch (const std::exception &e)
				{
					printf("Error exception:%s\n", e.what());
					delete this;
					return S_OK;
				}
			}
			return S_OK;
		}
		string_t strData = to_string_t(pszData);
		if (pszData == NULL)
			strData = U("");
		http_client client(strURL);
		web::http::method m = methods::GET;
		if (nMethod)
			m = methods::POST;
		http_request request(m);
		if (m_mapHeaders.size())
		{
			for (auto it : m_mapHeaders)
			{
				string_t strHeader = to_string_t(LPCTSTR(it.first));
				CString s = it.second;
				s.Replace(_T("\r\n"), _T(""));
				string_t strValue = to_string_t(LPCTSTR(s));
				request.headers().add(strHeader, strValue);
			}
			m_mapHeaders.erase(m_mapHeaders.begin(), m_mapHeaders.end());
		}
		if (strData != U(""))
		{
			request.set_body(strData, U("text/plain"));
		}

		auto _task = client.request(request).then([=](http_response response)
		{
			if (response.status_code() == status_codes::OK)
			{
				streams::istream bodyStream = response.body();
				container_buffer<std::string> inStringBuffer;
				bodyStream.read_to_end(inStringBuffer).then([=](task<size_t> previousTask)
				{
					CoInitializeEx(NULL, COINIT_MULTITHREADED);
					std::string &strRet = inStringBuffer.collection();
					std::wstring &strRet2 = to_utf16string(strRet);
					HWND hWnd = m_hNotifyWnd;
					if (hNotify)
					{
						hWnd = (HWND)hNotify;
					}
					if (::IsWindow(hWnd))
					{
						::SendMessage(hWnd, WM_UPLOADFILE, (WPARAM)strRet2.c_str(), 0);
					}
					CComBSTR bstrRet(strRet2.c_str());
					if (m_pRestNotify != nullptr)
					{
						m_pRestNotify->Notify(bstrRet);
					}
					CoUninitialize();
					delete this;
				});
			}
		});

		try
		{
			_task.wait();
		}
		catch (const http_exception& e)
		{
			printf("Error http_exception:%s\n", e.what());
		}
		catch (const std::exception &e)
		{
			printf("Error exception:%s\n", e.what());
		}
		return S_OK;
	});
	return S_OK;
}

STDMETHODIMP CRestObj::RestData(int nMethod, BSTR bstrServerURL, BSTR bstrRequest, BSTR bstrData, LONGLONG hNotify)
{
	auto t = create_task([nMethod, bstrServerURL, bstrRequest, bstrData, hNotify, this]()
	{
		string_t strURL = to_string_t(LPCTSTR(bstrServerURL));
		string_t strData = to_string_t(LPCTSTR(bstrData));
		string_t strRequest = to_string_t(LPCTSTR(bstrRequest));
		http_client client(strURL);
		web::http::method m = methods::GET;
		if (nMethod)
			m = methods::POST;
		http_request request(m);
		if (m_mapHeaders.size())
		{
			for (auto it : m_mapHeaders)
			{
				string_t strHeader = to_string_t(LPCTSTR(it.first));
				CString s = it.second;
				s.Replace(_T("\r\n"), _T(""));
				string_t strValue = to_string_t(LPCTSTR(s));
				request.headers().add(strHeader, strValue);
			}
			m_mapHeaders.erase(m_mapHeaders.begin(), m_mapHeaders.end());
		}
		if (strData != U(""))
		{
			request.set_body(strData, U("text/plain"));
		}

		auto _task = client.request(request).then([=](http_response response)
		{
			if (response.status_code() == status_codes::OK)
			{
				streams::istream bodyStream = response.body();
				container_buffer<std::string> inStringBuffer;
				try
				{
					bodyStream.read_to_end(inStringBuffer).then([=](task<size_t> previousTask)
					{
						std::string &strRet = inStringBuffer.collection();
						std::wstring &strRet2 = to_utf16string(strRet);
						if (hNotify)
						{
							HWND hWnd = (HWND)hNotify;
							if (::IsWindow(hWnd))
							{
								string_t strURL2 = strURL;
								strURL2 += strRequest;
								::SendMessage(hWnd, WM_UPLOADFILE, (WPARAM)strURL2.c_str(), (LPARAM)strRet2.c_str());
							}
						}
						BSTR bstrRet = CComBSTR(strRet2.c_str());
						if (m_pRestNotify != nullptr)
						{
							m_pRestNotify->Notify(bstrRet);
						}
						::SysFreeString(bstrRet);
						delete this;
					});
				}
				catch (const std::exception &e)
				{
					std::wostringstream ss;
					ss << e.what() << std::endl;
					std::wcout << ss.str();
					delete this;
				}
			}
		});
		try
		{
			_task.wait();
		}
		catch (const std::exception &e)
		{
			printf("Error exception:%s\n", e.what());
		}
		return S_OK;
	});
	return S_OK;
};

STDMETHODIMP CRestObj::UploadFile(VARIANT_BOOL bUpload, BSTR bstrServerURL, BSTR bstrRequest, BSTR bstrFilePath)
{
	auto t = create_task([bUpload, bstrServerURL, bstrRequest, bstrFilePath, this]()
	{
		string_t strURL = to_string_t(LPCTSTR(bstrServerURL));
		string_t strFile = to_string_t(LPCTSTR(bstrFilePath));
		string_t strRequest = to_string_t(LPCTSTR(bstrRequest));
		task<void> _task;
		
		if (bUpload)
		{
			_task = file_stream<unsigned char>::open_istream(strFile).then([this, strURL, strRequest](task<streams::basic_istream<unsigned char>> previousTask)
			{
				auto fileStream = previousTask.get();
				// Make HTTP request with the file stream as the body.
				http_client client(strURL);
				http_request request(methods::POST);
				if (m_mapHeaders.size())
				{
					for (auto it : m_mapHeaders)
					{
						string_t strHeader = to_string_t(LPCTSTR(it.first));
						CString s = it.second;
						s.Replace(_T("\r\n"), _T(""));
						string_t strValue = to_string_t(LPCTSTR(s));
						request.headers().add(strHeader, strValue);
					}
					m_mapHeaders.erase(m_mapHeaders.begin(), m_mapHeaders.end());
				}

				request.set_body(fileStream, U("application/octet-stream"));
				try
				{
					client.request(request).then([this, strURL, strRequest, fileStream](task<http_response> previousTask)
					{
						fileStream.close();
						try
						{
							auto response = previousTask.get();
							if (::IsWindow(m_hNotifyWnd))
							{
								::SendMessage(m_hNotifyWnd, WM_UPLOADFILE, (WPARAM)strURL.c_str(), (LPARAM)strRequest.c_str());
							}
							Notify(0x00002);
							CComPtr<IRestNotify> pCloudAddinRestNotify;
							this->get_CloudAddinRestNotify(&pCloudAddinRestNotify);
							if (pCloudAddinRestNotify != nullptr)
							{
								pCloudAddinRestNotify->Notify(CComBSTR(L"Success"));
							}
						}
						catch (const std::system_error& e)
						{
							printf("Error exception:%s\n", e.what());
							delete this;
							return pplx::task_from_result();
						}
						delete this;
						return pplx::task_from_result();
					});
				}
				catch (const http_exception& e)
				{
					std::wostringstream ss;
					ss << e.what() << std::endl;
					std::wcout << ss.str();
					// Return an empty task. 
					return pplx::task_from_result();
				}
				// Return an empty task. 
				return pplx::task_from_result();
			});
		}
		else
		{
			http_client client(strURL);
			web::http::method m = methods::POST;
			http_request request(m);
			if (m_mapHeaders.size())
			{
				for (auto it : m_mapHeaders)
				{
					string_t strHeader = to_string_t(LPCTSTR(it.first));
					CString s = it.second;
					s.Replace(_T("\r\n"), _T(""));
					string_t strValue = to_string_t(LPCTSTR(s));
					request.headers().add(strHeader, strValue);
				}
				m_mapHeaders.erase(m_mapHeaders.begin(), m_mapHeaders.end());
			}

			try
			{
				_task = client.request(request).then([=](http_response response)
				{
					if (response.status_code() == status_codes::OK)
					{
						int nLength = 0;
						web::http::http_headers::key_type strHeaderName = L"byteLength";
						auto it = response.headers().find(strHeaderName);
						if (it != response.headers().end())
						{
							string_t  val = it->second;
							std::wstring &strRet2 = to_utf16string(val);
							nLength = _wtoi64(strRet2.c_str());
						}
						else
						{
							strHeaderName = L"Content-Length";
							it = response.headers().find(strHeaderName);
							if (it != response.headers().end())
							{
								string_t  val = it->second;
								std::wstring &strRet2 = to_utf16string(val);
								nLength = _wtoi64(strRet2.c_str());
							}
						}
						streams::istream f = response.body();
						if (f.is_valid() && nLength)
						{
							try
							{
								auto fileBuffer = make_shared<streams::streambuf<uint8_t>>();
								task<streams::streambuf<uint8_t>> pTask = file_buffer<uint8_t>::open(strFile, std::ios::out);
								pTask.wait();
								*fileBuffer = pTask.get();
								task<size_t> pTask2 = f.read_to_end(*fileBuffer);
								pTask2.wait();
								fileBuffer->close().wait();
								size_t nSize = pTask2.get();
								if (nSize)
								{
									if (nLength == nSize)
									{
										if (::IsWindow(m_hNotifyWnd))
										{
											::SendMessage(m_hNotifyWnd, WM_UPLOADFILE, (WPARAM)strFile.c_str(), (LPARAM)strRequest.c_str());
										}
										//CComPtr<IRestNotify> pCloudAddinRestNotify;
										//this->get_CloudAddinRestNotify(&pCloudAddinRestNotify);
										if (m_pRestNotify != nullptr)
										{
											m_pRestNotify->Notify(CComBSTR(strFile.c_str()));
										}
										delete this;
									}
								};
							}
							catch (const std::exception &e)
							{
								std::wostringstream ss;
								ss << e.what() << std::endl;
								std::wcout << ss.str();
								return task_from_result();
							}
						}
					}
					//CoUninitialize();
					return task_from_result();
				});
			}
			catch (const http_exception& e)
			{
				std::wostringstream ss;
				ss << e.what() << std::endl;
				std::wcout << ss.str();
				// Return an empty task. 
				return pplx::task_from_result();
			}
		}
		try
		{
			_task.wait();
		}
		catch (const http_exception& e)
		{
			std::wostringstream ss;
			ss << e.what() << std::endl;
			std::wcout << ss.str();
			// Return an empty task. 
			return pplx::task_from_result();
		}
		catch (const std::exception &e)
		{
			printf("Error exception:%s\n", e.what());
			return pplx::task_from_result();
		}
		return pplx::task_from_result();
	});
	return S_OK;
}

STDMETHODIMP CRestObj::Notify(long nNotify)
{
	return S_OK;
}

STDMETHODIMP CRestObj::get_CloudAddinRestNotify(IRestNotify** pVal)
{
	if (m_pRestNotify)
	{
		*pVal = m_pRestNotify;
		(*pVal)->AddRef();
	}

	return S_OK;
}

STDMETHODIMP CRestObj::put_CloudAddinRestNotify(IRestNotify* newVal)
{
	m_pRestNotify = newVal;
	return S_OK;
}

STDMETHODIMP CRestObj::get_WndNode(IWndNode** pVal)
{
	if (m_pCloudAddinUINode)
	{
		*pVal = m_pCloudAddinUINode;
	}

	return S_OK;
}

STDMETHODIMP CRestObj::put_WndNode(IWndNode* newVal)
{
	m_pCloudAddinUINode = newVal;
	return S_OK;
}

STDMETHODIMP CRestObj::get_NotifyHandle(LONGLONG* pVal)
{
	*pVal = (LONGLONG)m_hNotifyWnd;
	return S_OK;
}

STDMETHODIMP CRestObj::put_NotifyHandle(LONGLONG newVal)
{
	m_hNotifyWnd = (HWND)newVal;
	return S_OK;
}

STDMETHODIMP CRestObj::get_Header(BSTR bstrHeaderName, BSTR* pVal)
{
	CString strHeader = OLE2T(bstrHeaderName);
	auto it = m_mapHeaders.find(strHeader);
	if (it != m_mapHeaders.end())
	{
		*pVal = it->second.AllocSysString();
	}
	return S_OK;
}

STDMETHODIMP CRestObj::put_Header(BSTR bstrHeaderName, BSTR newVal)
{
	CString strHeader = OLE2T(bstrHeaderName);
	string_t _strHeaderVal = to_string_t(newVal);
	auto it = m_mapHeaders.find(strHeader);
	if (it != m_mapHeaders.end())
		m_mapHeaders.erase(it);
	m_mapHeaders[strHeader] = OLE2T(newVal);
	return S_OK;
}

STDMETHODIMP CRestObj::ClearHeaders()
{
	m_mapHeaders.clear();
	return S_OK;
}

STDMETHODIMP CRestObj::get_RestKey(BSTR* pVal)
{
	*pVal = m_strKey.AllocSysString();
	return S_OK;
}

STDMETHODIMP CRestObj::put_RestKey(BSTR newVal)
{
	m_strKey = OLE2T(newVal);
	return S_OK;
}

STDMETHODIMP CRestObj::Clone(IRestObject* pTargetObj)
{
	if (pTargetObj != nullptr)
	{
		pTargetObj->put_NotifyHandle((LONGLONG)m_hNotifyWnd);
		if (m_mapHeaders.size())
		{
			for (auto it : m_mapHeaders)
			{
				pTargetObj->put_Header(it.first.AllocSysString(), it.second.AllocSysString());
			}
			m_mapHeaders.erase(m_mapHeaders.begin(), m_mapHeaders.end());
		}
	}
	return S_OK;
}

STDMETHODIMP CRestObj::get_State(int* pVal)
{
	*pVal = m_nState;
	return S_OK;
}

STDMETHODIMP CRestObj::put_State(int newVal)
{
	m_nState = newVal;
	return S_OK;
}
