// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#include "targetver.h"

#include <stdio.h>
#include <tchar.h>


#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS      // some CString constructors will be explicit

//#include <windows.h>
#include <atlbase.h>
#include <atlstr.h>
#include <comdef.h>
#include <atlwin.h>

//#import "C:\Windows\SysWOW64\msxml3.dll" raw_interfaces_only, raw_native_types, named_guids

#pragma warning(push)
#pragma warning(disable : 4278)
#pragma warning(disable : 4192)
#pragma warning(disable : 4146)

#import "C:\Program Files (x86)\Common Files\ArcGIS\bin\ArcGISVersion.dll" raw_interfaces_only, raw_native_types, no_namespace, named_guids
#import "C:\Program Files (x86)\ArcGIS\Desktop10.5\com\esriSystem.olb" raw_interfaces_only, raw_native_types, no_namespace, named_guids exclude("OLE_COLOR", "OLE_HANDLE", "VARTYPE")  exclude("esriProductCode", "XMLSerializer")
#import "C:\Program Files (x86)\ArcGIS\Desktop10.5\com\esriSystemUI.olb" raw_interfaces_only, raw_native_types, no_namespace, named_guids exclude("IProgressDialog")
#import "C:\Program Files (x86)\ArcGIS\Desktop10.5\com\esriGeometry.olb" raw_interfaces_only, raw_native_types, no_namespace, named_guids rename("IGeometry", "esriIGeometry") rename("ISegment", "esriISegment")
#import "C:\Program Files (x86)\ArcGIS\Desktop10.5\com\esriDisplay.olb" raw_interfaces_only, raw_native_types, no_namespace, named_guids
#import "C:\Program Files (x86)\ArcGIS\Desktop10.5\com\esriOutput.olb" raw_interfaces_only, raw_native_types, no_namespace, named_guids
#import "C:\Program Files (x86)\ArcGIS\Desktop10.5\com\esriGeoDatabase.olb" raw_interfaces_only, raw_native_types, no_namespace, named_guids 
#import "C:\Program Files (x86)\ArcGIS\Desktop10.5\com\esriDataSourcesRaster.olb" raw_interfaces_only, raw_native_types, no_namespace, named_guids
#import "C:\Program Files (x86)\ArcGIS\Desktop10.5\com\esriCarto.olb" raw_interfaces_only, raw_native_types, no_namespace, named_guids exclude("UINT_PTR", "OLE_COLOR", "OLE_HANDLE", "VARTYPE") 
#import "C:\Program Files (x86)\ArcGIS\Desktop10.5\com\esriFramework.olb" raw_interfaces_only, raw_native_types, no_namespace, named_guids exclude("UINT_PTR")
#import "C:\Program Files (x86)\ArcGIS\Desktop10.5\com\esriGISClient.olb" raw_interfaces_only, raw_native_types, no_namespace, named_guids
#import "C:\Program Files (x86)\ArcGIS\Desktop10.5\com\esriEditor.olb" raw_interfaces_only, raw_native_types, no_namespace, named_guids
#import "C:\Program Files (x86)\ArcGIS\Desktop10.5\com\esriDataSourcesRaster.olb" raw_interfaces_only, raw_native_types, no_namespace, named_guids
#import "C:\Program Files (x86)\ArcGIS\Desktop10.5\com\esriGeoDatabaseUI.olb" raw_interfaces_only, raw_native_types, no_namespace, named_guids exclude("VersionManager", "CLSID_VersionManager")
#import "C:\Program Files (x86)\ArcGIS\Desktop10.5\com\esriGeoProcessing.olb" raw_interfaces_only, raw_native_types, no_namespace, named_guids
#import "C:\Program Files (x86)\ArcGIS\Desktop10.5\com\esriArcMapUI.olb" raw_interfaces_only, raw_native_types, no_namespace, named_guids
#import "C:\Program Files (x86)\ArcGIS\Desktop10.5\com\esriCatalog.olb" raw_interfaces_only, raw_native_types, no_namespace, named_guids
#import "C:\Program Files (x86)\ArcGIS\Desktop10.5\com\esriCatalogUI.olb" raw_interfaces_only, raw_native_types, no_namespace, named_guids
#import "C:\Program Files (x86)\ArcGIS\Desktop10.5\com\esriDataSourcesRasterUI.olb" raw_interfaces_only, raw_native_types, no_namespace, named_guids
#import "C:\Program Files (x86)\ArcGIS\Desktop10.5\com\esriDataSourcesGDB.olb" raw_interfaces_only, raw_native_types, no_namespace, named_guids
#import "C:\Program Files (x86)\ArcGIS\Desktop10.5\com\esriDataSourcesFile.olb" raw_interfaces_only, raw_native_types, no_namespace, named_guids
#import "C:\Program Files (x86)\ArcGIS\Desktop10.5\com\esriServer.olb" raw_interfaces_only, raw_native_types, no_namespace, named_guids


// TODO: reference additional headers your program requires here
