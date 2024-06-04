{\rtf1\ansi\ansicpg1252\cocoartf2709
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 ArialMT;}
{\colortbl;\red255\green255\blue255;\red0\green0\blue0;\red16\green60\blue192;}
{\*\expandedcolortbl;;\cssrgb\c0\c0\c0;\cssrgb\c6667\c33333\c80000;}
\margl1440\margr1440\vieww28600\viewh14980\viewkind0
\deftab720
\pard\pardeftab720\partightenfactor0

\f0\fs16 \cf0 \expnd0\expndtw0\kerning0
\outl0\strokewidth0 \strokec2 // Copyright 2021. Increase BV. All Rights Reserved.\
//\
// Created By: Tibbe van Asten\
// for Increase B.V.\
//\
// Created: 15-06-2021\
// Last update: \
//\
// ABOUT THE SCRIPT\
// This script will export the ad strength of all RSA's in your\
// account to a Google Sheet. The sheet will contain the campaign,\
// adgroup, ad id and ad strength.\
//\
////////////////////////////////////////////////////////////////////\
\
var config = \{\
  \
  LOG : true,\
  \
  // Make a copy of this script and copy the URL: {\field{\*\fldinst{HYPERLINK "https://docs.google.com/spreadsheets/d/1WvNSbaZi2dz3Uu74AIniKy5YLHctZ_6c6i0Gp2DQ8ns/copy"}}{\fldrslt \cf3 \ul \ulc3 \strokec3 https://docs.google.com/spreadsheets/d/1WvNSbaZi2dz3Uu74AIniKy5YLHctZ_6c6i0Gp2DQ8ns/copy}}\
  SPREADSHEET_URL : "https://docs.google.com/spreadsheets/d/1_Uy7QN741cJpvizcInJ__4ovtgtxZLRXY_5bGgssC34/edit#gid=0",\
  SHEET_NAME : "RSA"\
  \
\}\
\
////////////////////////////////////////////////////////////////////\
\
function main() \{\
  \
  if(config.SPREADSHEET_URL == "https://")\{\
    throw Error("Make a copy of the sheet and paste the URL in the config \\nhttps://docs.google.com/spreadsheets/d/1WvNSbaZi2dz3Uu74AIniKy5YLHctZ_6c6i0Gp2DQ8ns/copy");\
  \}  \
    \
  var ss = SpreadsheetApp.openByUrl(config.SPREADSHEET_URL);\
  var sheet = ss.getSheetByName(config.SHEET_NAME);\
  \
  var report = AdsApp.report(\
    "SELECT CampaignName, AdGroupName, Id, AdStrengthInfo " +\
    "FROM AD_PERFORMANCE_REPORT " +\
    "WHERE AdType = RESPONSIVE_SEARCH_AD " +\
    "AND CampaignStatus = ENABLED " +\
    "AND AdGroupStatus = ENABLED " +\
    "AND Status = ENABLED"\
  )\
  \
  // Export data and clean up sheet\
  sheet.clearContents();\
  report.exportToSheet(sheet);\
  sheet.autoResizeColumns(1, sheet.getLastColumn());\
  \
  if(sheet.getMaxColumns() - sheet.getLastColumn() != 0)\{\
    sheet.deleteColumns(sheet.getLastColumn() + 1, sheet.getMaxColumns() - sheet.getLastColumn());\
  \}\
  \
  if(config.LOG === true)\{\
    \
    var rows = report.rows();\
    while(rows.hasNext())\{\
      var row = rows.next();\
\
      Logger.log("Campaign: " + row["CampaignName"] + " - Adgroup: " + row["AdGroupName"] + " - " + row["AdStrengthInfo"]);\
    \} // rowIterator\
    \
  \}\
  \
  Logger.log("Export completed");\
  \
\} // function main()}