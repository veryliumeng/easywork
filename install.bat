:: Copyright 2014 The Chromium Authors. All rights reserved.
:: Use of this source code is governed by a BSD-style license that can be
:: found in the LICENSE file.
@echo off
powershell unblock-file -path easywork.ps1
:: Change HKCU to HKLM if you want to install globally.
:: %~dp0 is the directory containing this bat script and ends with a backslash.
REG ADD "HKCU\Software\Google\Chrome\NativeMessagingHosts\smartforce" /ve /t REG_SZ /d "%~dp0easywork.json" /f
REG ADD "HKCU\SOFTWARE\Mozilla\NativeMessagingHosts\smartforce" /ve /t REG_SZ /d "%~dp0mozilla.json" /f
echo.
echo.
PAUSE
