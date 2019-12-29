Attribute VB_Name = "WinHttpEx"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WinHttpラップ・拡張モジュール
'* [詳  細] WinHttpのWrapperとしての機能を提供する他、Scriptingを使用した
'*          ユーティリティを提供する。
'*          ラップするWinHttpライブラリは以下のものとする。
'*              [name] Microsoft WinHTTP Services, version 5.1
'*              [dll] C:\Windows\System32\winhttpcom.dll
'* [参  考]
'*  <https://docs.microsoft.com/en-us/windows/win32/winhttp/winhttp-start-page>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* Enum定義
'******************************************************************************

'*-----------------------------------------------------------------------------
'* WinHttpRequestAutoLogonPolicy
'*
'*-----------------------------------------------------------------------------
Public Enum WinHttpRequestAutoLogonPolicy
    AutoLogonPolicy_Always = 0            'An authenticated log on, using the default credentials, is performed for all requests.
    AutoLogonPolicy_Never = 2             'Authentication is not used automatically.
    AutoLogonPolicy_OnlyIfBypassProxy = 1 'An authenticated log on, using the default credentials, is performed only for requests on the local intranet. The local intranet is considered to be any server on the proxy bypass list in the current proxy configuration.
End Enum

'*-----------------------------------------------------------------------------
'* WinHttpRequestOption
'*
'*-----------------------------------------------------------------------------
Public Enum WinHttpRequestOption
    WinHttpRequestOption_EnableCertificateRevocationCheck = 18  'Enables server certificate revocation checking during SSL negotiation. When the server presents a certificate, a check is performed to determine whether the certificate has been revoked by its issuer. If the certificate is indeed revoked, or the revocation check fails because the Certificate Revocation List (CRL) cannot be downloaded, the request fails; such revocation errors cannot be suppressed.
    WinHttpRequestOption_EnableHttp1_1 = 17                     'Sets or retrieves a boolean value that indicates whether HTTP/1.1 or HTTP/1.0 should be used. The default is TRUE, so that HTTP/1.1 is used by default.
    WinHttpRequestOption_EnableHttpsToHttpRedirects = 12        'Controls whether or not WinHTTP allows redirects. By default, all redirects are automatically followed, except those that transfer from a secure (https) URL to an non-secure (http) URL. Set this option to TRUE to enable HTTPS to HTTP redirects.
    WinHttpRequestOption_EnablePassportAuthentication = 13      'Enables or disables support for Passport authentication. By default, automatic support for Passport authentication is disabled; set this option to TRUE to enable Passport authentication support.
    WinHttpRequestOption_EnableRedirects = 6                    'Sets or retrieves a VARIANT that indicates whether requests are automatically redirected when the server specifies a new location for the resource. The default value of this option is VARIANT_TRUE to indicate that requests are automatically redirected.
    WinHttpRequestOption_EnableTracing = 10                     'Sets or retrieves a VARIANT that indicates whether tracing is currently enabled. For more information about the trace facility in Microsoft Windows HTTP Services (WinHTTP), see WinHTTP Trace Facility.
    WinHttpRequestOption_EscapePercentInURL = 3                 'Sets or retrieves a VARIANT that indicates whether percent characters in the URL string are converted to an escape sequence. The default value of this option is VARIANT_TRUE which specifies all unsafe American National Standards Institute (ANSI) characters except the percent symbol are converted to an escape sequence.
    WinHttpRequestOption_MaxAutomaticRedirects = 14             'ets or retrieves the maximum number of redirects that WinHTTP follows; the default is 10. This limit prevents unauthorized sites from making the WinHTTP client stall following a large number of redirects.
    WinHttpRequestOption_MaxResponseDrainSize = 16              'Sets or retrieves a bound on the amount of data that will be drained from responses in order to reuse a connection. The default is 1 MB.
    WinHttpRequestOption_MaxResponseHeaderSize = 15             'Sets or retrieves a bound set on the maximum size of the header portion of the server's response. This bound protects the client from a malicious server attempting to stall the client by sending a response with an infinite amount of header data. The default value is 64 KB.
    WinHttpRequestOption_RejectUserpwd = 19
    WinHttpRequestOption_RevertImpersonationOverSsl = 11        'Controls whether the WinHttpRequest object temporarily reverts client impersonation for the duration of the SSL certificate authentication operations. The default setting for the WinHttpRequest object is TRUE. Set this option to FALSE to keep impersonation while performing certificate authentication operations.
    WinHttpRequestOption_SecureProtocols = 9                    'Sets or retrieves a VARIANT that indicates which secure protocols can be used. This option selects the protocols acceptable to the client. The protocol is negotiated during the Secure Sockets Layer (SSL) handshake. The default value of this option is 0x0028, which indicates that SSL 2.0 or SSL 3.0 can be used. If this option is set to zero, the client and server are not able to determine an acceptable security protocol and the next Send results in an error.
    WinHttpRequestOption_SelectCertificate = 5                  'Sets a VARIANT that specifies the client certificate that is sent to a server for authentication. This option indicates the location, certificate store, and subject of a client certificate delimited with backslashes. For more information about selecting a client certificate, see SSL in WinHTTP.
    WinHttpRequestOption_SslErrorIgnoreFlags = 4                'Sets or retrieves a VARIANT that indicates which server certificate errors should be ignored. he default value of this option in Version 5.1 of WinHTTP is zero, which results in no errors being ignored. In earlier versions of WinHTTP, the default setting was 0x3300, which resulted in all server certificate errors being ignored by default.
    WinHttpRequestOption_URL = 1                                'Retrieves a VARIANT that contains the URL of the resource. This value is read-only; you cannot set the URL using this property. The URL cannot be read until the Open method is called. This option is useful for checking the URL after the Send method is finished to verify that any redirection occurred.
    WinHttpRequestOption_URLCodePage = 2                        'Sets or retrieves a VARIANT that identifies the code page for the URL string. The default value is the UTF-8 code page. The code page is used to convert the Unicode URL string, passed in the Open method, to a single-byte string representation.
    WinHttpRequestOption_UrlEscapeDisable = 7                   'Sets or retrieves a VARIANT that indicates whether unsafe characters in the path and query components of a URL are converted to escape sequences. The default value of this option is VARIANT_TRUE, which specifies that characters in the path and query are converted.
    WinHttpRequestOption_UrlEscapeDisableQuery = 8              'Sets or retrieves a VARIANT that indicates whether unsafe characters in the query component of the URL are converted to escape sequences. The default value of this option is VARIANT_TRUE, which specifies that characters in the query are converted.
    WinHttpRequestOption_UserAgentString = 0                    'Sets or retrieves a VARIANT that contains the user agent string.
End Enum

'*-----------------------------------------------------------------------------
'* WinHttpRequestSecureProtocols
'*
'*-----------------------------------------------------------------------------
Public Enum WinHttpRequestSecureProtocols
    SecureProtocol_ALL = 168
    SecureProtocol_SSL2 = 8
    SecureProtocol_SSL3 = 32
    SecureProtocol_TLS1 = 128
    SecureProtocol_TLS1_1 = 512
    SecureProtocol_TLS1_2 = 2048
End Enum

'*-----------------------------------------------------------------------------
'* WinHttpRequestSslErrorFlags
'*
'*-----------------------------------------------------------------------------
Public Enum WinHttpRequestSslErrorFlags
    SslErrorFlag_CertCNInvalid = 4096
    SslErrorFlag_CertDateInvalid = 8192
    SslErrorFlag_CertWrongUsage = 512
    SslErrorFlag_Ignore_All = 13056
    SslErrorFlag_UnknownCA = 256
End Enum

'*-----------------------------------------------------------------------------
'* HTTPREQUEST_PROXY_SETTING
'*
'*-----------------------------------------------------------------------------
Public Enum HTTPREQUEST_PROXY_SETTING
    HTTPREQUEST_PROXYSETTING_DEFAULT = 0   'Default proxy setting. Equivalent to HTTPREQUEST_PROXYSETTING_PRECONFIG.
    HTTPREQUEST_PROXYSETTING_PRECONFIG = 0 'ndicates that the proxy settings should be obtained from the registry. This assumes that Proxycfg.exe has been run. If Proxycfg.exe has not been run and HTTPREQUEST_PROXYSETTING_PRECONFIG is specified, then the behavior is equivalent to HTTPREQUEST_PROXYSETTING_DIRECT.
    HTTPREQUEST_PROXYSETTING_DIRECT = 1    'Indicates that all HTTP and HTTPS servers should be accessed directly. Use this command if there is no proxy server.
    HTTPREQUEST_PROXYSETTING_PROXY = 2     'When HTTPREQUEST_PROXYSETTING_PROXY is specified, varProxyServer should be set to a proxy server string and varBypassList should be set to a domain bypass list string. This proxy configuration applies only to the current instance of the WinHttpRequest object.
End Enum

'*-----------------------------------------------------------------------------
'* HTTPREQUEST_SETCREDENTIALS_FLAGS
'*
'*-----------------------------------------------------------------------------
Public Enum HTTPREQUEST_SETCREDENTIALS_FLAGS
    HTTPREQUEST_SETCREDENTIALS_FOR_SERVER = 0 'Credentials are passed to a server.
    HTTPREQUEST_SETCREDENTIALS_FOR_PROXY = 1  'Credentials are passed to a proxy.
End Enum

'******************************************************************************
'* メソッド定義
'******************************************************************************


